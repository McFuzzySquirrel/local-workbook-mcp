using System;
using System.Collections.Generic;
using System.Data;

namespace ExcelMcp.Core
{
    // Thread-safe lazy cache + diff helper
    public static class WorkbookCache
    {
        private static readonly object _lock = new();
        private static Dictionary<string, List<DataTable>> _snapshot = new(StringComparer.OrdinalIgnoreCase);
        private static DateTime _lastLoadedUtc;
        private static string? _filePath;

        public static void Initialize(string filePath)
        {
            lock (_lock)
            {
                if (_filePath == null)
                {
                    _filePath = filePath;
                    _snapshot = ExcelReader.LoadWorkbook(filePath);
                    _lastLoadedUtc = DateTime.UtcNow;
                }
            }
        }

        public static (IReadOnlyList<ExcelDiff.TableDiff> diffs, DateTime timestampUtc) ReloadAndDiff()
        {
            if (_filePath == null) throw new InvalidOperationException("WorkbookCache not initialized");
            lock (_lock)
            {
                var newSnap = ExcelReader.LoadWorkbook(_filePath);
                var diffs = ExcelDiff.Compare(_snapshot, newSnap);
                _snapshot = newSnap;
                _lastLoadedUtc = DateTime.UtcNow;
                return (diffs, _lastLoadedUtc);
            }
        }

        public static DateTime LastLoadedUtc
        {
            get { lock (_lock) return _lastLoadedUtc; }
        }
    }
}
