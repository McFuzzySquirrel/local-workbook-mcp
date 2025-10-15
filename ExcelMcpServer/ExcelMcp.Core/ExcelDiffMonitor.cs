using System;

using System.Collections.Generic;

using System.Data;

using System.IO;

using System.Threading;



namespace ExcelMcp.Core

{

    public class ExcelDiffMonitor : IDisposable

    {

        private readonly string _filePath;

        private readonly FileSystemWatcher _watcher;

        private Dictionary<string, List<DataTable>> _lastSnapshot;

        private readonly Timer _debounceTimer;

        private volatile bool _pending;



        public event Action<ExcelChangeEvent>? OnChange;



        public ExcelDiffMonitor(string filePath)

        {

            _filePath = filePath;

            _lastSnapshot = ExcelReader.LoadWorkbook(filePath);



            var dir = Path.GetDirectoryName(filePath) ?? ".";

            var fileName = Path.GetFileName(filePath);



            _watcher = new FileSystemWatcher(dir, fileName)

            {

                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName

            };



            _watcher.Changed += Watcher_Changed;

            _watcher.Renamed += Watcher_Changed;

            _watcher.Created += Watcher_Changed;

            _watcher.Deleted += Watcher_Changed;



            // debounce: wait 500ms after last event

            _debounceTimer = new Timer(OnDebounceTimer, null, Timeout.Infinite, Timeout.Infinite);

        }



        private void Watcher_Changed(object sender, FileSystemEventArgs e)

        {

            // mark pending and reset timer

            _pending = true;

            _debounceTimer.Change(500, Timeout.Infinite);

        }



        private void OnDebounceTimer(object? state)

        {

            if (!_pending) return;

            _pending = false;



            // Try a few times in case file is locked by Excel save

            for (int attempt = 0; attempt < 5; attempt++)

            {

                try

                {

                    var newSnapshot = ExcelReader.LoadWorkbook(_filePath);

                    var diffs = ExcelDiff.Compare(_lastSnapshot, newSnapshot);

                    foreach (var d in diffs)

                    {

                        var ev = new ExcelChangeEvent(d.Sheet, d.Table ?? string.Empty, d.ChangeType, d.Row, d.Column, d.OldValue, d.NewValue, DateTime.UtcNow);

                        OnChange?.Invoke(ev);

                    }

                    _lastSnapshot = newSnapshot;

                    break;

                }

                catch (IOException)
                {
                    int delayMs = 150 * (int)Math.Pow(2, attempt); // 150, 300, 600, 1200, 2400
                    Thread.Sleep(delayMs);
                    continue;
                }

                catch (Exception ex)

                {

                    // If parse fails, just log in host; do not throw here

                    Console.WriteLine($"ExcelDiffMonitor: error reading workbook: {ex.Message}");

                    break;

                }

            }

        }



        public void Start()

        {

            _watcher.EnableRaisingEvents = true;

        }



        public void Dispose()

        {

            _watcher.Dispose();

            _debounceTimer.Dispose();

        }

    }

}