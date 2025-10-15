using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;
using ExcelMcp.Core;
using System.Data;
using System.Collections.Generic;

namespace ExcelMcp.Tests
{
    public class DiffTests
    {
        private static string CreateWorkbook(Action<XSSFWorkbook> build)
        {
            var wb = new XSSFWorkbook();
            build(wb);
            var tmp = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using var fs = new FileStream(tmp, FileMode.Create, FileAccess.Write);
            wb.Write(fs, false);
            return tmp;
        }

        [Fact]
        public void DetectsSheetAddition()
        {
            var f1 = CreateWorkbook(wb =>
            {
                var s1 = wb.CreateSheet("Sheet1");
                var h = s1.CreateRow(0); h.CreateCell(0).SetCellValue("Col");
                var r = s1.CreateRow(1); r.CreateCell(0).SetCellValue("A");
            });
            var f2 = CreateWorkbook(wb =>
            {
                var s1 = wb.CreateSheet("Sheet1");
                var h = s1.CreateRow(0); h.CreateCell(0).SetCellValue("Col");
                var r = s1.CreateRow(1); r.CreateCell(0).SetCellValue("A");
                var s2 = wb.CreateSheet("Sheet2");
                var h2 = s2.CreateRow(0); h2.CreateCell(0).SetCellValue("Col");
            });

            var oldData = ExcelReader.LoadWorkbook(f1);
            var newData = ExcelReader.LoadWorkbook(f2);
            var diffs = ExcelDiff.Compare(oldData, newData);
            Assert.Contains(diffs, d => d.ChangeType == "SheetAdded" && d.Sheet == "Sheet2");
        }

        [Fact]
        public void DetectsCellUpdateAndRowAdd()
        {
            // Build deterministic in-memory DataTables to focus on diff logic
            var oldTbl = new DataTable("Table_1");
            oldTbl.Columns.Add("Item");
            oldTbl.Columns.Add("Qty");
            var or = oldTbl.NewRow(); or[0] = "Apples"; or[1] = "10"; oldTbl.Rows.Add(or);

            var newTbl = new DataTable("Table_1");
            newTbl.Columns.Add("Item");
            newTbl.Columns.Add("Qty");
            var nr1 = newTbl.NewRow(); nr1[0] = "Apples"; nr1[1] = "11"; newTbl.Rows.Add(nr1);
            var nr2 = newTbl.NewRow(); nr2[0] = "Bananas"; nr2[1] = "5"; newTbl.Rows.Add(nr2);

            var oldData = new Dictionary<string, List<DataTable>>(StringComparer.OrdinalIgnoreCase)
            {
                ["Data"] = new List<DataTable> { oldTbl }
            };
            var newData = new Dictionary<string, List<DataTable>>(StringComparer.OrdinalIgnoreCase)
            {
                ["Data"] = new List<DataTable> { newTbl }
            };

            var diffs = ExcelDiff.Compare(oldData, newData);
            Assert.Contains(diffs, d => d.ChangeType == "CellUpdated" && d.OldValue == "10" && d.NewValue == "11");
            Assert.Contains(diffs, d => d.ChangeType == "RowAdded");
        }
    }
}
