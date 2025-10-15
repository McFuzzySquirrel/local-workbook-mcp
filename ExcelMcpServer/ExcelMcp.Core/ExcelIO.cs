using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelMcp.Core
{
    public static class ExcelReadHelpers
    {
        public static string[][] ReadSheetAsMatrix(string file, string sheetName)
        {
            if (!File.Exists(file)) throw new FileNotFoundException(file);
            using var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var wb = NPOI.SS.UserModel.WorkbookFactory.Create(fs);
            var sheet = wb.GetSheet(sheetName);
            if (sheet == null) throw new KeyNotFoundException($"Sheet {sheetName} not found");
            int first = sheet.FirstRowNum;
            int last = sheet.LastRowNum;
            var rows = new List<string[]>();
            for (int r = first; r <= last; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;
                int lastCell = row.LastCellNum;
                if (lastCell < 0) continue;
                var arr = new string[lastCell];
                for (int c = 0; c < lastCell; c++)
                {
                    var cell = row.GetCell(c);
                    arr[c] = cell?.ToString() ?? string.Empty;
                }
                rows.Add(arr);
            }
            (wb as IDisposable)?.Dispose();
            return rows.ToArray();
        }
    }

    public static class ExcelWriteHelpers
    {
        public static void WithRetry(Action action)
        {
            Exception? last = null;
            for (int i = 0; i < 5; i++)
            {
                try { action(); return; }
                catch (IOException ex) { last = ex; System.Threading.Thread.Sleep(150 * (int)Math.Pow(2, i)); }
            }
            throw last ?? new Exception("Unknown IO error");
        }

        private static void BackupFile(string path)
        {
            try
            {
                var bak = path + ".bak";
                File.Copy(path, bak, overwrite: true);
            }
            catch { /* ignore */ }
        }

        private static (int row, int col) ParseA1(string address)
        {
            if (string.IsNullOrWhiteSpace(address)) throw new ArgumentException("address");
            int i = 0; int col = 0;
            while (i < address.Length && char.IsLetter(address[i]))
            {
                col = col * 26 + (char.ToUpperInvariant(address[i]) - 'A' + 1);
                i++;
            }
            int row = 0;
            while (i < address.Length && char.IsDigit(address[i]))
            {
                row = row * 10 + (address[i] - '0');
                i++;
            }
            if (row <= 0 || col <= 0) throw new ArgumentException("Invalid A1 address");
            return (row - 1, col - 1);
        }

        public static void SetCellValue(string file, string sheetName, string address, string value)
        {
            if (!File.Exists(file)) CreateNewWorkbook(file, sheetName);
            BackupFile(file);
            byte[] bytes = File.ReadAllBytes(file);
            using var ms = new MemoryStream(bytes, writable: false);
            var wb = NPOI.SS.UserModel.WorkbookFactory.Create(ms);

            var sheet = wb.GetSheet(sheetName) ?? wb.CreateSheet(sheetName);
            var (rowIdx, colIdx) = ParseA1(address);
            var row = sheet.GetRow(rowIdx) ?? sheet.CreateRow(rowIdx);
            var cell = row.GetCell(colIdx) ?? row.CreateCell(colIdx);
            cell.SetCellValue(value ?? string.Empty);

            var tmp = file + ".tmp";
            using (var outFs = new FileStream(tmp, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                if (wb is NPOI.XSSF.UserModel.XSSFWorkbook xssf)
                    xssf.Write(outFs, true);
                else
                    wb.Write(outFs);
                outFs.Flush(true);
            }
            (wb as IDisposable)?.Dispose();
            File.Copy(tmp, file, overwrite: true);
            try { File.Delete(tmp); } catch { }
        }

        public static void AppendRow(string file, string sheetName, string[] values)
        {
            if (!File.Exists(file)) CreateNewWorkbook(file, sheetName);
            BackupFile(file);
            byte[] bytes = File.ReadAllBytes(file);
            using var ms = new MemoryStream(bytes, writable: false);
            var wb = NPOI.SS.UserModel.WorkbookFactory.Create(ms);
            var sheet = wb.GetSheet(sheetName) ?? wb.CreateSheet(sheetName);
            int newRowIdx = sheet.LastRowNum >= 0 ? sheet.LastRowNum + 1 : 0;
            var row = sheet.CreateRow(newRowIdx);
            for (int i = 0; i < values.Length; i++)
            {
                row.CreateCell(i).SetCellValue(values[i] ?? string.Empty);
            }

            var tmp = file + ".tmp";
            using (var outFs = new FileStream(tmp, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                if (wb is NPOI.XSSF.UserModel.XSSFWorkbook xssf)
                    xssf.Write(outFs, true);
                else
                    wb.Write(outFs);
                outFs.Flush(true);
            }
            (wb as IDisposable)?.Dispose();
            File.Copy(tmp, file, overwrite: true);
            try { File.Delete(tmp); } catch { }
        }

        private static void CreateNewWorkbook(string file, string initialSheet)
        {
            var ext = Path.GetExtension(file).ToLowerInvariant();
            Directory.CreateDirectory(Path.GetDirectoryName(file) ?? ".");
            if (ext == ".xls")
            {
                var wb = new NPOI.HSSF.UserModel.HSSFWorkbook();
                wb.CreateSheet(string.IsNullOrWhiteSpace(initialSheet) ? "Sheet1" : initialSheet);
                using var fs = new FileStream(file, FileMode.CreateNew, FileAccess.Write, FileShare.None);
                wb.Write(fs);
                (wb as IDisposable)?.Dispose();
            }
            else
            {
                var wb = new NPOI.XSSF.UserModel.XSSFWorkbook();
                wb.CreateSheet(string.IsNullOrWhiteSpace(initialSheet) ? "Sheet1" : initialSheet);
                using var fs = new FileStream(file, FileMode.CreateNew, FileAccess.Write, FileShare.None);
                wb.Write(fs);
                (wb as IDisposable)?.Dispose();
            }
        }
    }
}
