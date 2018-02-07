using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;

namespace TollTrack
{
    public partial class Form1 : Form
    {
        private string TollURL = @"https://online.toll.com.au/trackandtrace/";
        //private SortedList<string,Tuple<string,DateTime>> consignmentIds = new SortedList<string,Tuple<string,DateTime>>() {{"AREW065066",("Unknown",DateTime.MinValue)}}; // ID, Status
        public Form1()
        {
            InitializeComponent();
            ExcelTest("test.xlsx");
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            webBrowser.Navigate(TollURL);
        }

        private void ExcelTest(string filename)
        {
            // open excel app
            var xl = new Excel.Application();
            if (xl == null)
            {
                MessageBox.Show("Excel is not installed");
                return;
            }
            xl.SheetsInNewWorkbook = 1;
            xl.Visible = false;

            // open workbook or create a new ones
            Excel.Workbook workbook;
            filename = Path.GetFullPath(filename);
            if (File.Exists(filename))
            {
                workbook = xl.Workbooks.Open(filename);
            }
            else
            {
                workbook = xl.Workbooks.Add(Missing.Value);
                workbook.SaveAs(filename);
            }

            var worksheet = workbook.Worksheets[1];
            worksheet.Cells[1,1] = "test";
            worksheet.Cells[2, 1] = "space";
            worksheet.Cells[3, 1] = "things";

            // save and quit
            workbook.Close(true, Missing.Value, Missing.Value);
            xl.Quit();
        }

        private void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            var trackingIds = "";
            
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
