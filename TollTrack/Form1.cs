using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using CefSharp;
using CefSharp.WinForms;
using OfficeOpenXml;
using Timer = System.Windows.Forms.Timer;

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

        // 10 ids in
        // search
        // wait for result
        // store result
        // repeat until all ids done

        private object SearchLock = new object();

        private string TollURL = @"https://online.toll.com.au/trackandtrace/";
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

            doneTimer.Interval = 2000;
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
                    Log("Found table");
                    GetDeliveries();
                }
            });
        }

        // read, input to webpage and press go button
        private void btnSelect_Click(object sender, EventArgs e)
        {
            if (loaded)
            {
                ReadExcel();
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            SearchForIDs();
        }

        private void btnOut_Click(object sender, EventArgs e)
        {
            if (loaded)
            {
                OutputToExcel();
            }
        }

        // add output to textbox
        private void Log(string str)
        {
            if (string.IsNullOrWhiteSpace(str)) return;
            if (txtInfo.InvokeRequired)
            {
                var d = new LogCallback(Log);
                Invoke(d, str);
            }
            else
                txtInfo.AppendText(Environment.NewLine + str);
        }

        private void RunJS(string command, JSCallback cb = null)
        {
            if (!loaded)
            {
                Log("Page is not loaded yet");
                return;
            }

            var task1 = webBrowser.GetBrowser().MainFrame.EvaluateScriptAsync(command).ContinueWith((task) =>
            {
                if (task.IsCompleted && !task.IsCanceled && !task.IsFaulted && task.Status == TaskStatus.RanToCompletion)
                {
                    Log(@"Ran Javascript command");
                    cb?.Invoke(Convert.ToString(task.Result?.Result ?? string.Empty));
                }
                else
                {
                    Log(@"Failed Javascript command");
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
                Log(e.Message);
                return;
            }

            if (workSheet == null)
                return;

            // read conids
            var cell = GetCell(workSheet, "CON NOTE NUMBER");
            for (int rowIndex = cell.Start.Row; rowIndex < workSheet.Dimension.End.Row; rowIndex++)
            {
                var conId = workSheet.Cells[rowIndex + 1, cell.Start.Column]?.Value?.ToString() ?? "";
                if (conId.ToUpper() == "TRANSFER") continue;
                if (!consignmentIds.ContainsKey(conId) && !string.IsNullOrWhiteSpace(conId))
                    consignmentIds.Add(conId, default);
            }
            Log("Done Loading input");
        }

        // limited to 10 conIds at a time
        // add ids to search bar and click button
        private bool SearchForIDs()
        {
            var trackingIds = "";
            int limit = ConsignmentIdIndex + maxPerRequest;

            for (var i = ConsignmentIdIndex; ConsignmentIdIndex < consignmentIds.Keys.ToList().Count; i++)
            {
                var c = consignmentIds.Keys.ToList()[i];
                if (i >= limit) break;
                trackingIds += i == ConsignmentIdIndex ? $"{c}" + Environment.NewLine : string.Empty;
                ConsignmentIdIndex++;
            }
            if (trackingIds.Length < 1) return false;

            var command = $@"document.getElementById('connoteIds').value = `{trackingIds}`; $('#tatSearchButton').click()";
            
            RunJS(command);
            doneTimer.Enabled = true;
            doneTimer.Start();
            return true;
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
            Log("Storing delivery results");
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
                    consignmentIds[i.Key] = new Delivery("", "", i.Value);
                }
            }
            Log($"Processing... {ConsignmentIdIndex}/{consignmentIds.Count} ({ConsignmentIdIndex/consignmentIds.Count * 100:F2}%)");
            //TODO: store in global scope for output, then do next round of tracking IDs
            //TODO: when finished with getting TrackingResult of IDs then output them all to output doc
            if (ConsignmentIdIndex < consignmentIds.Count)
            {
                var locked = !Monitor.TryEnter(SearchLock);
                if (!locked)
                {
                    Monitor.Enter(SearchLock, ref locked);
                    if (locked)
                    {
                        SearchForIDs();
                        Monitor.Exit(SearchLock);
                    }
                }
            }

            if (ConsignmentIdIndex >= consignmentIds.Count)
                Log("Done. Ready to output.");
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

            // ids and output cells
            var range = GetColumnRange(workSheet, "CONSIGNMENT REFERENCE");
            // var invoiceNo = GetColumnRange(workSheet, "Invoice #");
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
                        workSheet.Cells[cell.Start.Row, dateCol].Value = delivery.date.ToShortDateString();
                        workSheet.Cells[cell.Start.Row, dateCol + 1].Value = delivery.status;
                        Log($"Update {conId} date: {delivery.date} status: {delivery.status}");
                    }
                }
            }
            package.Save();
        }
    }
}
