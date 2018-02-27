using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Dynamic;
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
        public class Delivery
        {
            public string invoiceID;
            public string status;
            public DateTime date;

            public Delivery(string invoiceID, string status, DateTime date)
            {
                this.invoiceID = invoiceID;
                this.status = status;
                this.date = date;
            }
        }

        private string TollURL = @"https://online.toll.com.au/trackandtrace/";
        /// <summary>
        /// SortedList&lt;ConsignmentID,Tuple&lt;InvoiceID, DeliveryStatus, DeliveryDate&gt;&gt;
        /// </summary>
        private SortedList<string, Delivery> consignmentIds = new SortedList<string, Delivery>(){{"AREW065066", new Delivery("1210661","Unknown",DateTime.MinValue)}}; // ID, Status
        private ChromiumWebBrowser webBrowser;
        private const int maxPerRequest = 10;
        private int ConsignmentIdIndex = 2;
        private Timer doneTimer = new Timer();
        private bool loaded = false;
        delegate void LogCallback(string text);
        delegate void JSCallback(string result);

        public Form1()
        {
            InitializeComponent();
            webBrowser = new ChromiumWebBrowser(TollURL);
            webBrowser.Dock = DockStyle.Fill;
            Controls.Add(webBrowser);

            // wait for page to load then enable buttons
            webBrowser.LoadingStateChanged += (sender, args) =>
            {
                if (args.IsLoading == false)
                {
                    loaded = true;
                    Invoke(new Action(() => 
                    {
                        btnSelect.Enabled = true;
                        btnRun.Enabled = true;
                        btnOut.Enabled = true;
                        txtInfo.AppendText(Environment.NewLine + "Page loaded");
                    }));
                }
            };

            doneTimer.Interval = 1000;
            doneTimer.Enabled = false;
            doneTimer.Tick += DoneTimerOnTick;
        }

        private void DoneTimerOnTick(object sender, EventArgs eventArgs)
        {
            var command = @"(function () {
                return document.getElementById('TATMultiConnoteResultId') != null;
            })();";

            // check to see if our results are there
            RunJS(command, result => 
            {
                if (result.ToUpper() == "TRUE")
                {
                    doneTimer.Stop();
                    log("Found table");
                    GetDeliveries();
                }
            });
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            doneTimer.Start();
            log("Looking for table");
        }

        private void log(string str)
        {
            if (string.IsNullOrWhiteSpace(str)) return;
            if (txtInfo.InvokeRequired)
            {
                var d = new LogCallback(log);
                Invoke(d, str);
            }
            else
                txtInfo.AppendText(Environment.NewLine + str);
        }

        private void btnOut_Click(object sender, EventArgs e)
        {
            if (loaded)
            {
                OutputToExcel();
                /*webBrowser.GetBrowser().MainFrame.EvaluateScriptAsync("document.getElementById('quickSearchTableResult').innerHTML").ContinueWith(
                x =>{Console.WriteLine(x.Result.Result);});*/
            }
        }

        // read, input to webpage and press go button
        private void btnSelect_Click(object sender, EventArgs e)
        {
            if (loaded)
            { 
                ReadExcel();
                SearchForIDs();
            }
        }

        private void RunJS(string command, JSCallback cb = null)
        {
            // cannot run js before page is loaded
            if (!loaded)
            {
                log("Page is not loaded yet");
                return;
            }

            var task1 = webBrowser.GetBrowser().MainFrame.EvaluateScriptAsync(command).ContinueWith((task) =>
            {
                if (task.IsCompleted && !task.IsCanceled && !task.IsFaulted && task.Status == TaskStatus.RanToCompletion)
                {
                    log(@"Ran Javascript command");
                    cb?.Invoke(Convert.ToString(task.Result?.Result ?? string.Empty));
                }
                else
                {
                    log(@"Failed to run Javascript command on webpage");
                }
            });
        }

        // find cell with matching name in the worksheet
        // return null if no match
        private ExcelRangeBase GetCell(ExcelWorksheet workSheet, string name)
        {
            var startRow = 1;
            var dataColumn = 1;
            foreach (var cell in workSheet.Cells)
            {
                var id = cell?.Value?.ToString()?.ToUpper();
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
        private ExcelRange GetColumnRange(ExcelWorksheet workSheet, string name)
        {
            var cell = GetCell(workSheet, name);
            if (cell != null)
                return workSheet.Cells[cell.Start.Row, cell.Start.Column, workSheet.Dimension.End.Row-1, cell.Start.Column];
            return null;
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

            ExcelWorksheet workSheet;
            ExcelPackage package;
            try
            {
                package = new ExcelPackage(new FileInfo(ofd.FileName));
                workSheet = package.Workbook.Worksheets.FirstOrDefault(w => w.Name.ToUpper() == "SHIPPED");
            }
            catch (Exception e)
            {
                log(e.Message);
                return;
            }

            if (workSheet == null)
                return;

            // read conids
            var cell = GetCell(workSheet, "CON NOTE NUMBER");
            for (int rowIndex = cell.Start.Row; rowIndex < workSheet.Dimension.End.Row; rowIndex++)
            {
                var conId = workSheet.Cells[rowIndex, cell.Start.Column]?.Value?.ToString() ?? "";
                if (conId.ToUpper() == "TRANSFER") continue;
                if (!consignmentIds.ContainsKey(conId) && !string.IsNullOrWhiteSpace(conId))
                    consignmentIds.Add(conId, default);
            }
        }

        // limited to 10 conIds at a time
        // add ids to search bar and click button
        private void SearchForIDs()
        {
            var trackingIds = "";
            for (var i = ConsignmentIdIndex; ConsignmentIdIndex < consignmentIds.Keys.ToList().Count; i++)
            {
                var c = consignmentIds.Keys.ToList()[i];
                if (ConsignmentIdIndex >= maxPerRequest) break;
                trackingIds += i == ConsignmentIdIndex ? $"{c}" + Environment.NewLine : string.Empty;
                ConsignmentIdIndex++;
            }
            if (trackingIds.Length < 1) return;
            var command = $"document.getElementById('connoteIds').value = `{trackingIds}`; $('#tatSearchButton').click() ";
            RunJS(command);
        }

        // Store results from run webpage in consignmentIds
        private void GetDeliveries()
        {
            var command = @"(function () {
                var rows = $('.tatMultConRow');
                var ret = [];
                for (var i = 0; i<rows.length;i++) { ret.push({key: rows[i].children[0].children[0].innerText, value: new Date(rows[i].children[4].children[2].innerText).toISOString()}) };
                return JSON.stringify(ret,null,3);
            })();";
            log(Environment.NewLine + "Storing delivery results");
            RunJS(command, FormatOutput);
        }

        // deserialize json result and add to sorted list
        private void FormatOutput(string s)
        {
            var output = s.FromJson<TrackingResult>();
            output = output.Select(o => new TrackingResult {Key = o.Key, Value = o.Value.ToLocalTime()}).ToList();
            foreach(var i in output)
            {
                if (consignmentIds.ContainsKey(i.Key))
                {
                    consignmentIds[i.Key].date = i.Value;
                }
            }
            //TODO: store in global scope for output, then do next round of tracking IDs
            //TODO: when finished with getting TrackingResult of IDs then output them all to output doc
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
            ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault(w => w.Name.ToUpper() == "BNMA");
        
            if (workSheet == null)
                return;

            // ids to update
            var range = GetColumnRange(workSheet, "CONSIGNMENT \nREFERENCE");
            // var invoiceNo = GetColumnRange(workSheet, "Invoice #");

            // output column locations
            var dateCol = (GetCell(workSheet, "TEST")?.Start.Column ?? 0);
            var statusCol = (GetCell(workSheet, "DATE DELIVERED")?.Start.Column ?? 0);
   
            foreach (var cell in range)
            {
                // update matching id delivery date/status
                var conId = cell?.Value?.ToString() ?? "";        
                if (consignmentIds.ContainsKey(conId))
                {
                    var delivery = consignmentIds[conId];
                    if (delivery != null)
                    {
                        workSheet.Cells[cell.Start.Row, dateCol].Value = delivery.date;
                        workSheet.Cells[cell.Start.Row, dateCol + 1].Value = delivery.status;
                        log(Environment.NewLine + $"Updated Id {conId} date:{delivery.date} status:{delivery.status}");
                    }
                }
            }
            package.Save();
        }
    }
}
