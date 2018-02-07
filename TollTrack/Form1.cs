using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace TollTrack
{
    public partial class Form1 : Form
    {
        private string TollURL = @"https://online.toll.com.au/trackandtrace/";
        private SortedList<string,Tuple<string,DateTime>> consignmentIds = new SortedList<string,Tuple<string,DateTime>>() {{"AREW065066",new Tuple<string, DateTime>("Unknown",DateTime.MinValue)}}; // ID, Status
        public Form1()
        {
            InitializeComponent();
            //ExcelTest("test.xlsx");
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            webBrowser.Navigate(TollURL);
            ReadExcel();
        }

        private void ReadExcel()
        {
            var ofd = new OpenFileDialog
            {
                Filter = @"Excel Files|*.xlsx;*.xls;*.csv",
                Title = @"Select Input File"
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName));
            ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault();
            //"Con Note Number

            var startRow = 0;
            var dataColumn = 0;

            for (int rowIndex = 1; rowIndex < workSheet.Dimension.Rows; rowIndex++)
            {
                for (var colIndex = 1; colIndex < workSheet.Dimension.Columns; colIndex++)
                {
                    if (workSheet.Cells[rowIndex, colIndex]?.Value?.ToString()?.ToUpper() != "CON NOTE NUMBER") continue;
                    startRow = rowIndex + 1;
                    dataColumn = colIndex;
                    break;
                }

                if (dataColumn != 0) break;
            }

            if (dataColumn == 0)
            {
                MessageBox.Show(@"Could not find a cell with 'Con Note Number' in it");
                return;
            }

            for (int rowIndex = startRow; rowIndex < workSheet.Dimension.Rows; rowIndex++)
            {
                consignmentIds.Add(workSheet.Cells[rowIndex,dataColumn].Value.ToString(),default);
            }


        //private void ExcelTest(string filename)
        //{
        //    var workbook = LoadWorkbook(filename);
        //    var worksheet = workbook.ActiveSheet;
        //    worksheet.Cells[1,1] = "test";
        //    worksheet.Cells[2,1] = "space";
        //    worksheet.Cells[3,1] = "things";
        //    workbook.Close(true, Missing.Value, Missing.Value);
        //    excel?.Quit();
        //}

        private void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            var trackingIds = "";
            
            consignmentIds.Keys.ToList().ForEach(c=> trackingIds += $"{c}{Environment.NewLine}");
            //consignmentIds.ForEach(c=> trackingIds += $"{c}{Environment.NewLine}");

            var command = $"document.getElementById('connoteIds').innerText = '{trackingIds}'; $('.dijitButtonNode').click() ";

            webBrowser.Document?.ExecCommand(command,false,null); // populate text box where IDs are meant to be with some javascript
            webBrowser.Document.GetElementById("table where the results are") // then get the status and
                .GetAttribute("The status for each ID");
            // update the SortedList for each ID

            // write to Excel document
        }
    }
}
