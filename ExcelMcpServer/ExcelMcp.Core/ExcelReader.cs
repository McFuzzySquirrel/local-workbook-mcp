using System;

using System.Collections.Generic;

using System.Data;

using System.IO;

using NPOI.SS.UserModel;

using NPOI.XSSF.UserModel; // .xlsx

using NPOI.HSSF.UserModel; // .xls



namespace ExcelMcp.Core

{

    // Loads workbook -> Dictionary<sheetName, List<DataTable>> (tables found by header rows)

    public static class ExcelReader

    {

        public static Dictionary<string, List<DataTable>> LoadWorkbook(string filePath)

        {

            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            IWorkbook workbook;

            var ext = Path.GetExtension(filePath).ToLowerInvariant();

            if (ext == ".xls")

                workbook = new HSSFWorkbook(stream);

            else // .xlsx / .xlsm

                workbook = new XSSFWorkbook(stream);



            var result = new Dictionary<string, List<DataTable>>(StringComparer.OrdinalIgnoreCase);



            for (int s = 0; s < workbook.NumberOfSheets; s++)

            {

                var sheet = workbook.GetSheetAt(s);

                var tables = ExtractTablesFromSheet(sheet);

                result[sheet.SheetName] = tables;

            }



            return result;

        }



        private static List<DataTable> ExtractTablesFromSheet(ISheet sheet)

        {

            var tables = new List<DataTable>();

            DataTable? current = null;

            int lastRow = sheet.LastRowNum;



            for (int r = 0; r <= lastRow; r++)

            {

                var row = sheet.GetRow(r);

                if (row == null)

                {

                    // break table if present

                    if (current != null && current.Rows.Count > 0) { tables.Add(current); current = null; }

                    continue;

                }



                // Heuristic: header row => mostly strings and non-empty

                if (IsHeaderRow(row))

                {

                    if (current != null && current.Rows.Count > 0) tables.Add(current);

                    current = new DataTable($"Table_{tables.Count + 1}");

                    int maxCol = row.LastCellNum;

                    for (int c = 0; c < maxCol; c++)

                    {

                        var cell = row.GetCell(c);

                        var colName = cell?.ToString() ?? $"Column{c+1}";

                        // avoid duplicate columnnames

                        var name = colName;

                        var idx = 1;

                        while (current.Columns.Contains(name))

                        {

                            name = $"{colName}_{idx++}";

                        }

                        current.Columns.Add(name);

                    }

                }

                else if (current != null)

                {

                    var dataRow = current.NewRow();

                    for (int c = 0; c < current.Columns.Count; c++)

                    {

                        var cell = row.GetCell(c);

                        dataRow[c] = cell?.ToString() ?? string.Empty;

                    }

                    // add if any data present

                    var hasData = false;

                    for (int i = 0; i < dataRow.ItemArray.Length; i++) if (dataRow[i] != null && !string.IsNullOrWhiteSpace(dataRow[i].ToString())) { hasData = true; break; }

                    if (hasData) current.Rows.Add(dataRow);

                }

            }



            if (current != null && current.Rows.Count > 0) tables.Add(current);



            return tables;

        }



        private static bool IsHeaderRow(IRow row)

        {

            // header heuristic: at least half of populated cells are strings and not numeric-only

            int populated = 0; int stringLike = 0;

            foreach (var cell in row.Cells)

            {

                if (cell == null) continue;

                var txt = cell.ToString();

                if (string.IsNullOrWhiteSpace(txt)) continue;

                populated++;

                if (cell.CellType == CellType.String) stringLike++;

                else

                {

                    // also consider as header if text content present (like dates are often headers? not usually)

                    if (!double.TryParse(txt, out _)) stringLike++;

                }

            }

            if (populated == 0) return false;

            return stringLike >= Math.Ceiling(populated / 2.0);

        }

    }

}