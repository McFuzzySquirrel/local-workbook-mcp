using System;

using System.Collections.Generic;

using System.Data;

using System.Linq;



namespace ExcelMcp.Core

{

    public static class ExcelDiff

    {

    public record TableDiff(string Sheet, string? Table, string ChangeType, int? Row, string? Column, string? OldValue, string? NewValue);



        public static List<TableDiff> Compare(

            Dictionary<string, List<DataTable>> oldData,

            Dictionary<string, List<DataTable>> newData)

        {

            var diffs = new List<TableDiff>();



            var allSheets = newData.Keys.Union(oldData.Keys, StringComparer.OrdinalIgnoreCase).Distinct(StringComparer.OrdinalIgnoreCase);



            foreach (var sheetName in allSheets)

            {

                var oldHas = oldData.ContainsKey(sheetName);

                var newHas = newData.ContainsKey(sheetName);



                if (!oldHas && newHas)

                {

                    diffs.Add(new TableDiff(sheetName, null, "SheetAdded", null, null, null, null));

                    continue;

                }

                if (oldHas && !newHas)

                {

                    diffs.Add(new TableDiff(sheetName, null, "SheetRemoved", null, null, null, null));

                    continue;

                }



                var oldTables = oldData[sheetName];

                var newTables = newData[sheetName];

                int maxT = Math.Max(oldTables.Count, newTables.Count);

                for (int t = 0; t < maxT; t++)

                {

                    if (t >= oldTables.Count)

                    {

                        diffs.Add(new TableDiff(sheetName, $"Table_{t+1}", "TableAdded", null, null, null, null));

                        continue;

                    }

                    if (t >= newTables.Count)

                    {

                        diffs.Add(new TableDiff(sheetName, $"Table_{t+1}", "TableRemoved", null, null, null, null));

                        continue;

                    }

                    var oldTable = oldTables[t];

                    var newTable = newTables[t];

                    diffs.AddRange(CompareTables(sheetName, oldTable, newTable));

                }

            }



            return diffs;

        }



        private static IEnumerable<TableDiff> CompareTables(string sheetName, DataTable oldTable, DataTable newTable)

        {

            var diffs = new List<TableDiff>();

            int maxRows = Math.Max(oldTable.Rows.Count, newTable.Rows.Count);

            int maxCols = Math.Max(oldTable.Columns.Count, newTable.Columns.Count);



            for (int r = 0; r < maxRows; r++)

            {

                if (r >= oldTable.Rows.Count)

                {

                    diffs.Add(new TableDiff(sheetName, oldTable.TableName, "RowAdded", r + 1, null, null, null));

                    continue;

                }

                if (r >= newTable.Rows.Count)

                {

                    diffs.Add(new TableDiff(sheetName, oldTable.TableName, "RowRemoved", r + 1, null, null, null));

                    continue;

                }



                for (int c = 0; c < maxCols; c++)

                {

                    string colName = c < oldTable.Columns.Count ? oldTable.Columns[c].ColumnName :

                                     (c < newTable.Columns.Count ? newTable.Columns[c].ColumnName : $"Column{c+1}");



                    var oldVal = c < oldTable.Columns.Count ? oldTable.Rows[r][c]?.ToString() : null;

                    var newVal = c < newTable.Columns.Count ? newTable.Rows[r][c]?.ToString() : null;



                    if (!string.Equals(oldVal, newVal, StringComparison.Ordinal))

                    {

                        diffs.Add(new TableDiff(sheetName, oldTable.TableName, "CellUpdated", r + 1, colName, oldVal, newVal));

                    }

                }

            }



            return diffs;

        }

    }

}