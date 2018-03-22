using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;

namespace TollTrack
{
    public class ExcelToll
    {
        public static ExcelWorksheet Load(ref ExcelPackage package,string fileName, string option1)
        {
            
            try
            {
                package = new ExcelPackage(new FileInfo(fileName));
                foreach(var workSheet in package.Workbook.Worksheets)
                {
                    if (workSheet.Name.ToUpper() == option1)
                        return workSheet;
                }
            }
            catch (Exception e)
            {
                //Log(e.Message);
                Console.WriteLine(e.Message);
            }
            return null;
        }

        public static List<ExcelWorksheet> LoadAll(ref ExcelPackage package)
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
                return package.Workbook.Worksheets.ToList();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return null;
        }

        // find all cells with name match
        public static List<ExcelRangeBase> GetCells(ExcelWorksheet workSheet, string name)
        {
            var cells = new List<ExcelRangeBase>();
            foreach (var cell in workSheet.Cells)
            {
                var id = cell?.Value?.ToString()?.Replace("\n", "").ToUpper();
                if (id == name)
                {
                    cells.Add(cell);
                }
            }
            return cells;
        }

        // find cell with matching name in the worksheet
        // return null if no match
        public static ExcelRangeBase GetCell(ExcelWorksheet workSheet, string name)
        {
            var cells = GetCells(workSheet, name);
            if (cells.Count > 0) return cells[0];
            return null;
        }

        // get colume from the id of the match found
        // util used in output
        public static int GetColumn(ExcelWorksheet workSheet, string name, int id)
        {
            var cells = GetCells(workSheet, name);
            if (cells.Count == 0 || id < 0 || id > cells.Count - 1)
            {
                return 0;
            }
            return cells[id]?.Start.Column ?? 0;
        }

        // return column range from a cell with a matching name
        // return null if no match
        public static ExcelRange GetColumnRange(ExcelWorksheet workSheet, string name)
        {
            var cell = GetCell(workSheet, name);
            if (cell != null)
                return workSheet.Cells[cell.Start.Row + 1, cell.Start.Column, workSheet.Cells.End.Row - 1, cell.Start.Column];
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
