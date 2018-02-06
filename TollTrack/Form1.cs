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

namespace TollTrack
{
    public partial class Form1 : Form
    {
        private string TollURL = @"https://online.toll.com.au/trackandtrace/";
        private SortedList<string,string> consignmentIds = new SortedList<string,string>() {{"AREW065066","Unknown"}}; // ID, Status
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            webBrowser.Navigate(TollURL);
        }

        private void ExcelTest()
        {
            var xl = new Excel.Application();
            if (xl == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }
            xl.Visible = true;

            var workbook = xl.Workbooks.Open("test");
            var worksheet = workbook.Worksheets[1];

            workbook.Save();
        }

        private void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            var trackingIds = "";
            var command = "";
            consignmentIds.ForEach(c=> trackingIds += $"{c}{Environment.NewLine}");

            webBrowser.Document?.ExecCommand(command,false,null);
            webBrowser.Document.GetElementById("table where the results are")
                .GetAttribute("The status for each ID");
            // update the SortedList for each ID
        }
    }
}
