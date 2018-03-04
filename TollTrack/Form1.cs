using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using CefSharp.WinForms;
using OfficeOpenXml;
using Timer = System.Windows.Forms.Timer;

namespace TollTrack
{
    public partial class Form1 : Form
    {
        public class Delivery
        {
            public string customerPO;
            public string invoiceID;
            public string conID;
            public string status;
            public DateTime date;

            public Delivery(string customerPO, string invoiceID, string conID, string status, DateTime date)
            {
                this.customerPO = customerPO;
                this.invoiceID = invoiceID;
                this.conID = conID;
                this.status = status;
                this.date = date;
            }
        }

        private string MyTollUrl = @"https://mytoll.com";
        private string TollURL = @"https://online.toll.com.au/trackandtrace/";

        private string NZCURL =
            "http://www.nzcouriers.co.nz/nzc/servlet/ITNG_TAndTServlet?page=1&VCCA=Enabled&Key_Type=BarCode&barcode_data="; //todo: concat the consignmentID to this stinrg and then open in CEF

        private List<Delivery> deliveries = new List<Delivery>();
        private ChromiumWebBrowser webBrowser;
        private const int maxPerRequest = 10;
        private int ConsignmentIdIndex = 0;
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
                        txtInfo.AppendText(Environment.NewLine + "Page loaded");
                    }));
                }
            };       
            doneTimer.Interval = 2000;
            doneTimer.Enabled = false;
            doneTimer.Tick += DoneTimerOnTick;
            Log("Loading page " + TollURL);
        }

        // wait for results from SearchForIDs then search for next batch
        private void DoneTimerOnTick(object sender, EventArgs eventArgs)
        {
            var command = @"(function () {
                return $('#loadingPopUpDialogId').css('display') === 'none';
            })();";

            var result = RunJS(command);
            if (result.ToUpper() == "TRUE")
            {
                Log("Result found!");
                GetDeliveries();
                if (!SearchForIDs())
                {
                    doneTimer.Stop();
                    btnOut.Enabled = true;
                }
            }
        }

        private void githubToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("https://github.com/9ParsonsB/TollTrack");
        }

        // read, input to webpage and press go button
        private void btnSelect_Click(object sender, EventArgs e)
        {
            if (loaded)
            {
                ReadExcel();
            }
        }

        // search for id batchs
        private void btnRun_Click(object sender, EventArgs e)
        {
            if (loaded)
            {
                processBar.Minimum = 0;
                processBar.Maximum = deliveries.Count;
                ConsignmentIdIndex = 0;
                SearchForIDs();
                doneTimer.Start();
            }
        }

        // output deliveries list to excel
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

        private string RunJS(string command, JSCallback cb = null)
        {
            if (!loaded)
            {
                Log("Page is not loaded yet");
                return "";
            }
            var task1 = webBrowser.GetBrowser().MainFrame.EvaluateScriptAsync(command);
            task1.Wait();
            var result = Convert.ToString(task1.Result?.Result ?? string.Empty);
            cb?.Invoke(result);
            return result;
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
                workSheet = package.Workbook.Worksheets.FirstOrDefault(w => w.Name.ToUpper() == "SHIPPED" || w.Name.ToUpper() == txtFormat.Text.ToUpper());
                if (workSheet is default)
                {
                    Log("shipped worksheet not found in " + ofd.FileName);
                    return;
                }
            }
            catch (Exception e)
            {
                Log(e.Message);
                return;
            }

            // package = null;
            // workSheet = ExcelToll.Load(package, ofd.FileName);
            deliveries.Clear();

            // adds all rows that are are missing the date delivered
            if (isUpdate.Checked)
            {
                // read columns into lists
                var invoiceIds = ExcelToll.GetColumn(workSheet, "INVOICE #");
                var customerPOs = ExcelToll.GetColumn(workSheet, "CUSTOMER PO #");
                var conIds = ExcelToll.GetColumn(workSheet, "CONSIGNMENT REFERENCE");
                var dates = ExcelToll.GetColumn(workSheet, "DATE DELIVERED");

                if (invoiceIds == null || customerPOs == null || conIds == null || dates == null)
                {
                    Log("Failed to find all columns with the correct names");
                    return;
                }

                for (int i = 0; i < conIds.Count; i++)
                {
                    if (dates[i] == "")
                    {
                        deliveries.Add(new Delivery(customerPOs[i], invoiceIds[i], conIds[i], "Unknown", new DateTime()));
                    }
                }
            }
            // adds all deliveries by default
            else
            {
                // read columns into lists
                var invoiceIds = ExcelToll.GetColumn(workSheet, "PACKSLIP");
                var customerPOs = ExcelToll.GetColumn(workSheet, "CUST REF");
                var conIds = ExcelToll.GetColumn(workSheet, "CON NOTE NUMBER");

                if (invoiceIds == null || customerPOs == null || conIds == null)
                {
                    Log("Failed to find all columns with the correct names");
                    return;
                }

                for (int i = 0; i < conIds.Count; i++)
                {
                    deliveries.Add(new Delivery(customerPOs[i], invoiceIds[i], conIds[i], "Unknown", new DateTime()));
                }
            }

            // remove blank entries(don't know when column ends so alot are blank)
            //Delivery empty = new Delivery("", "", "", "Unknown", new DateTime());
            //deliveries.RemoveAll(i => i == empty);

            // remove certain entries
            int num = 0;
            deliveries = deliveries.Distinct().ToList();
            deliveries.ForEach(i => i.invoiceID = i.invoiceID.Replace("GS", ""));
            deliveries.RemoveAll(i => i.invoiceID.ToUpper() == "SAMPLES" || i.invoiceID.ToUpper() == "REPLACEMENT");
            deliveries.RemoveAll(i => i.conID.ToUpper() == "TRANSFER" || int.TryParse(i.conID, out num));
            deliveries.RemoveAll(i => !i.conID.Contains("ARE"));
  
            Log($"Done Loading input {deliveries.Count}");
            btnRun.Enabled = true;
        }

        // limited to 10 conIds at a time
        // add ids to search bar and click button
        private bool SearchForIDs()
        {
            var trackingIds = "";
            int limit = ConsignmentIdIndex + maxPerRequest;
            for (var i = ConsignmentIdIndex; i < deliveries.Count; i++)
            {
                var delivery = deliveries[i];
                if (i >= limit) break;
                trackingIds += $"{delivery.conID}" + Environment.NewLine;
            }

            if (trackingIds.Length < 1)
                return false;

            Log("Searching for consignmment ids");
            var command = $@"document.getElementById('connoteIds').value = `{trackingIds}`; $('#tatSearchButton').click()";
            RunJS(command);
 
            return true;
        }

        // Store results from run webpage in consignmentIds
        private void GetDeliveries()
        {
            // magic date formating
            var commmand = @"(function () {
                var rows = $('.tatMultConRow');
                var ret = [];
                for (var i = 0; i < rows.length; i++)
                {
                    var splitTime = [];
                    console.log(i);
                    var dateString = rows[i].children[4].children[2].innerText;
                    var splitDateString = dateString.split(' ');
                    var justDatestring = splitDateString[1];
                    var splitDate = justDatestring.split('/');
                    if (splitDateString.length < 3)
                    {
                        splitTime = ['00','00'];
                        var hour = '00'
                    }
                    else
                    {
                        splitTime = splitDateString[2].split(':');

                        if (splitDateString[3].toUpperCase() == 'AM')
                            var hour = '0' + splitTime[0];
                        else
                            var hour = parseInt(splitTime[0]) + 12;
                    }


                    var date = new Date(splitDate[2], splitDate[1], splitDate[0], hour, splitTime[1]);
                    ret.push({ key: rows[i].children[0].children[0].innerText, value: date.toISOString()});  
                }   
                return JSON.stringify(ret,null,3);
            })();";


            //TODO: use this when on NZ Couriers website
            var NZCCommand = @"var ret

            var raw = $('#status-dark').find('tbody')[0].children[1].children[3].innerHTML

            var splitRaw = raw.split(' ')[1].split('/')

            var hour = raw.split(' ')[0].split(':')[0]
            var minute = raw.split(' ')[0].split(':')[1].substring(0,2) 
            var pm = raw.split(' ')[0].split(':')[1].substring(2).toUpperCase() === 'P.M.'

            if (pm)
            {
                hour = parseInt(hour) + 12
            }

            var date = new Date(splitRaw[2],splitRaw[1],splitRaw[0],hour,minute)

            var consignment = new URL(window.location.href).searchParams.get('barcode_data')

            ret.push({key: consignment, value: date.toIsoString()})";

            Log("Storing delivery results");
            RunJS(commmand, FormatOutput);
        }

        // deserialize json result and add to sorted list
        private void FormatOutput(string s)
        {
            var output = s.FromJson<TrackingResult>();
            if (output == null)
            {
                Log("Failed to deserialize Tracking result");
                return;
            }

            output = output.Select(o => new TrackingResult {Key = o.Key, Value = o.Value.ToLocalTime()}).ToList();
            for (int i = 0; i < output.Count(); i++)
            {
                deliveries[ConsignmentIdIndex + i].date = output[i].Value;
            }
            ConsignmentIdIndex = Math.Min(ConsignmentIdIndex + maxPerRequest, deliveries.Count());

            processBar.Increment(maxPerRequest);
            Log($"Processing... {ConsignmentIdIndex}/{deliveries.Count} ({((float)ConsignmentIdIndex / (float)deliveries.Count) * 100f:F2}%)");
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
            ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault(w => w.Name.ToUpper() == txtFormat.Text);

            if (workSheet is default)
            {
                Log(txtFormat.Text + " worksheet not found in " + ofd.FileName);
                return;
            }

            // customer po range to compare look for matches
            var range = ExcelToll.GetColumnRange(workSheet, "INVOICE #");

            // output column locations
            var customerPO = (ExcelToll.GetCell(workSheet, "CUSTOMER PO #")?.Start.Column ?? 0);
            var invoiceNO = (ExcelToll.GetCell(workSheet, "INVOICE #")?.Start.Column ?? 0);
            var conId = (ExcelToll.GetCell(workSheet, "CONSIGNMENT REFERENCE")?.Start.Column ?? 0);
            var date = (ExcelToll.GetCell(workSheet, "DATE DELIVERED")?.Start.Column ?? 0);

            // prevent crash if a column is missing
            if (customerPO == 0 || invoiceNO == 0 || conId == 0 || date == 0)
            {
                Log("Failed to find one of the columns");
                return;
            }

            var donelist = new List<Delivery>();
            int matches = 0;
            foreach (var cell in range)
            {
                // update data where id matches
                var id = cell?.Value?.ToString() ?? "";
                if (id == "")
                    continue;

                var delivery = deliveries.FirstOrDefault(i => i.invoiceID == id);
                if (delivery != null)
                {
                    donelist.Add(delivery);
                    matches++;
                    workSheet.Cells[cell.Start.Row, customerPO].Value = delivery.customerPO;
                    workSheet.Cells[cell.Start.Row, invoiceNO].Value = delivery.invoiceID;
                    workSheet.Cells[cell.Start.Row, conId].Value = delivery.conID;
                    workSheet.Cells[cell.Start.Row, date].Value = delivery.date.ToShortDateString();
                    // Log($"{matches}. customer:{delivery.customerPO} invoice:{delivery.invoiceID} conid:{delivery.conID} date:{delivery.date} status:{delivery.status}");
                }
            }
            Log($"{matches} matches updated");
            package.Save();

            // show details of deliveries not found in output for manual assignment
            if (donelist.Count == deliveries.Count) return;
            var notDone = deliveries.Where(d => !donelist.Contains(d)).ToList();
            var frm = new Form2(notDone);
            frm.ShowDialog();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
