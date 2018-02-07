﻿using System;
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
        private Excel.Application excel;
        public Form1()
        {
            InitializeComponent();
            ExcelTest("test.xlsx");
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            webBrowser.Navigate(TollURL);
        }

        private Excel.Workbook LoadWorkbook(string filename)
        {
            // open excel app once
            if (excel == null)
            {
                excel = new Excel.Application();
                if (excel == null)
                {
                    throw new Exception("Excel is not installed");
                }
                excel.SheetsInNewWorkbook = 1;
                excel.Visible = false;
            }

            // open workbook or create a new one
            Excel.Workbook workbook;
            filename = Path.GetFullPath(filename);
            if (File.Exists(filename))
            {
                workbook = excel.Workbooks.Open(filename);
            }
            else
            {
                workbook = excel.Workbooks.Add(Missing.Value);
                workbook.SaveAs(filename);
            }
            return workbook;
        }

        private void ExcelTest(string filename)
        {
            var workbook = LoadWorkbook(filename);
            var worksheet = workbook.ActiveSheet;
            worksheet.Cells[1,1] = "test";
            worksheet.Cells[2,1] = "space";
            worksheet.Cells[3,1] = "things";
            workbook.Close(true, Missing.Value, Missing.Value);
            excel?.Quit();
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
