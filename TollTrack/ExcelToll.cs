using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace TollTrack
{
    class ExcelToll
    {
        public static ExcelWorksheet Load(ExcelPackage package, string path)
        {
            try
            {
                package = new ExcelPackage(new FileInfo(path));
                // workSheet = package.Workbook.Worksheets.FirstOrDefault(w => w.Name.ToUpper() == "SHIPPED" || w.Name.ToUpper() == txtFormat.Text.ToUpper());
            }
            catch (Exception e)
            {
                // Log(e.Message);
                // return;
            }
            return package.Workbook.Worksheets.First();
        }

        // find cell with matching name in the worksheet
        // return null if no match
        public static ExcelRangeBase GetCell(ExcelWorksheet workSheet, string name)
        {
            var startRow = 1;
            var dataColumn = 1;
            foreach (var cell in workSheet.Cells)
            {
                var id = cell?.Value?.ToString()?.Replace("\n", "").ToUpper();
                if (id == name)
                {
                    startRow = cell.Start.Row;
                    dataColumn = cell.Start.Column;
                    return cell;
                }
            }
            return null;
        }

        // return column range from a cell with a matching name
        // return null if no match
        public static ExcelRange GetColumnRange(ExcelWorksheet workSheet, string name)
        {
            var cell = GetCell(workSheet, name);
            if (cell != null)
                return workSheet.Cells[cell.Start.Row + 1, cell.Start.Column, workSheet.Dimension.End.Row - 1, cell.Start.Column];
            return null;
        }

        // get list of values for a column
        public static List<string> GetColumn(ExcelWorksheet workSheet, string name)
        {
            var range = GetColumnRange(workSheet, name);
            if (range == null)
            {
                //Form1.Log(name + " column not found");
                return default;
            }
            List<string> items = new List<string>();
            foreach (var cell in range)
            {
                items.Add(cell.Value?.ToString() ?? "");
            }
            return items;
        }
    }
}
