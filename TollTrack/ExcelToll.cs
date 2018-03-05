using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace TollTrack
{
    public class ExcelToll
    {
        public static ExcelWorksheet Load(ref ExcelPackage package, string option1, string option2)
        {
            var ofd = new OpenFileDialog
            {
                Filter = @"Excel Files|*.xlsx;*.xlsm;*.xls;*.csv;",
                Title = @"Select Output File"
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return null;

            try
            {
                package = new ExcelPackage(new FileInfo(ofd.FileName));
                foreach(var workSheet in package.Workbook.Worksheets)
                {
                    if (workSheet.Name.ToUpper() == option1 || workSheet.Name.ToUpper() == option2)
                        return workSheet;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return null;
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
            var cell = GetCell(workSheet, name);
            if (cell == null)
            {
                return null;
            }

            var items = new List<string>();
            var row = cell.Start.Row + 1;
            var column = cell.Start.Column;
        
            // loop through column 
            // convert null to empty string
            for (int i = 0; i < workSheet.Dimension.End.Row; i++)
            {
                var item = workSheet.Cells[row + i, column];
                items.Add(item.Text ?? "");
            }
            return items;
        }
    }
}
