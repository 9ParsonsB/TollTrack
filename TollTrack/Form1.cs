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
using CefSharp;
using CefSharp.WinForms;
using OfficeOpenXml;


namespace TollTrack
{
    public partial class Form1 : Form
    {
        private string TollURL = @"https://www.mytoll.com/";
        /// <summary>
        /// SortedList&lt;ConsignmentID,Tuple&lt;InvoiceID, DeliveryStatus, DeliveryDate&gt;&gt;
        /// </summary>
        private SortedList<string,Tuple<string, string,DateTime>> consignmentIds = new SortedList<string,Tuple<string,string,DateTime>>() {{"AREW065066",new Tuple<string, string, DateTime>("1210661","Unknown",DateTime.MinValue)}}; // ID, Status
        private ChromiumWebBrowser webBrowser;
        private const int maxPerRequest = 30;
        private int ConsignmentIdIndex = 0;

        public Form1()
        {
            InitializeComponent();
            webBrowser = new ChromiumWebBrowser(TollURL);
            webBrowser.Dock = DockStyle.Fill;
            this.Controls.Add(webBrowser);
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            webBrowser.LoadingStateChanged += WebBrowserOnLoadingStateChanged;
            webBrowser.GetBrowser().MainFrame.LoadUrl(TollURL);
        }

        private void WebBrowserOnLoadingStateChanged(object sender, LoadingStateChangedEventArgs loadingStateChangedEventArgs)
        {
            if (loadingStateChangedEventArgs.IsLoading) return;

            var trackingIds = "";
            
            consignmentIds.Keys.ToList().ForEach(c=> trackingIds += $"{c}{Environment.NewLine}");
            //consignmentIds.ForEach(c=> trackingIds += $"{c}{Environment.NewLine}");

            var command = $"document.getElementById('quickSearch').value = '{trackingIds}'; $('#search-shipment-btn').click() ";

            var task1 = webBrowser.GetBrowser().MainFrame.EvaluateScriptAsync(command).ContinueWith((task) =>
                {
                    
                }); // populate text box where IDs are meant to be with some javascript
             

             // then get the status and
            //.GetAttribute("The status for each ID");
            // update the SortedList for each ID

            // write to Excel document
            webBrowser.LoadingStateChanged -= WebBrowserOnLoadingStateChanged;
        }

        private void ReadExcel()
        {
            var ofd = new OpenFileDialog
            {
                Filter = @"Excel Files|*.xlsx;*.xlsm;*.xls;*.csv;",
                Title = @"Select Input File"
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName));
            ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault(w=>w.Name.ToUpper() == "SHIPPED");
            //"Con Note Number

            if (workSheet == null)
                return;

            var startRow = 0;
            var dataColumn = 0;

            for (int rowIndex = workSheet.Dimension.Start.Row; rowIndex < workSheet.Dimension.End.Row; rowIndex++)
            {
                for (var colIndex = workSheet.Dimension.Start.Column; colIndex < workSheet.Dimension.End.Column; colIndex++)
                {
                    if (workSheet.Cells[rowIndex, colIndex]?.Value?.ToString()?.ToUpper() != "CON NOTE NUMBER")
                        continue;
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

            for (int rowIndex = startRow; rowIndex < workSheet.Dimension.End.Row; rowIndex++)
            {
                var conId = workSheet.Cells[rowIndex, dataColumn]?.Value?.ToString() ?? "";
                if (conId.ToUpper() == "TRANSFER") continue;
                if (!consignmentIds.ContainsKey(conId) && !string.IsNullOrWhiteSpace(conId))
                    consignmentIds.Add(conId, default);
            }
        }

        // read, input to webpage and press go button
        private void btnSelect_Click(object sender, EventArgs e)
        {
            ReadExcel();

            var trackingIds = "";
            
            for (var i = ConsignmentIdIndex; ConsignmentIdIndex < consignmentIds.Keys.ToList().Count; i++)
            {
                var c = consignmentIds.Keys.ToList()[i];
                if (ConsignmentIdIndex >= maxPerRequest) break;
                trackingIds += i == ConsignmentIdIndex ? Environment.NewLine + $"{c}" : string.Empty;
                ConsignmentIdIndex++;
            }

            //consignmentIds.ForEach(c=> trackingIds += $"{c}{Environment.NewLine}");

            var command = $"document.getElementById('quickSearch').value = `{trackingIds.Substring(1)}`; $('#search-shipment-btn').click() ";

            var task1 = webBrowser.GetBrowser().MainFrame.EvaluateScriptAsync(command).ContinueWith((task) =>
            {
                if (task.IsCompleted && !task.IsCanceled && !task.IsFaulted &&
                    task.Status == TaskStatus.RanToCompletion)
                {
                    Console.WriteLine(@"ran JS");
                }
                else
                {
                    Console.WriteLine(@"JS Failed");
                }
            });
        }

        private void btnOut_Click(object sender, EventArgs e)
        {
            webBrowser.GetBrowser().MainFrame.EvaluateScriptAsync("document.getElementById('quickSearchTableResult').innerHTML").ContinueWith(
                x =>
                {
                    Console.WriteLine(x.Result.Result);
                });
        }


        private void OutputToExcel()
        {
            var ofd = new OpenFileDialog
            {
                Filter = @"Excel Files|*.xlsx;*.xlsm;*.xls;*.csv;",
                Title = @"Select Output File"
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return;
            ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName));
            
        }
    }
}
