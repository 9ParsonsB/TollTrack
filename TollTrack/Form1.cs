using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using CefSharp.WinForms;
using OfficeOpenXml;

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
            public string courier;
            public DateTime date;

            public Delivery(string customerPO, string invoiceID, string conID, string status, string courier, DateTime date)
            {
                this.customerPO = customerPO;
                this.invoiceID = invoiceID;
                this.conID = conID;
                this.status = status;
                this.date = date;
                this.courier = courier;
            }
        }

        private string MyTollUrl = @"https://mytoll.com";
        private string TollURL = @"https://online.toll.com.au/trackandtrace/";
        private string NZCURL = "http://www.nzcouriers.co.nz/nzc/servlet/ITNG_TAndTServlet?page=1&VCCA=Enabled&Key_Type=BarCode&barcode_data="; //todo: concat the consignmentID to this stinrg and then open in CEF
        private string PDTURL = "http://www.pbt.co.nz/default.aspx";
        private ChromiumWebBrowser webBrowser;
        private string CurrentURL;
        private List<Delivery> deliveries = new List<Delivery>();
        private List<List<Delivery>> testing;
        private const int maxPerRequest = 10;
        private int ConsignmentIdIndex = 0;
        private bool loaded = false;
        delegate void LogCallback(string text);
        delegate void JSCallback(string result);

        public Form1()
        {
            InitializeComponent();
            webBrowser = new ChromiumWebBrowser();
            webBrowser.Dock = DockStyle.Fill;
            Controls.Add(webBrowser);
            webBrowser.BringToFront();

            // wait for page to load then enable buttons
            webBrowser.LoadingStateChanged += (sender, args) =>
            {
                if (args.IsLoading == false)
                {
                    loaded = true;
                    CurrentURL = args.Browser.MainFrame.Url;
                    Invoke(new Action(() => 
                    {
                        btnSelect.Enabled = true;
                        txtInfo.AppendText(Environment.NewLine + "Page loaded " + CurrentURL);
                    }));
                }
            };       
        }

        private void githubToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("https://github.com/9ParsonsB/TollTrack");
        }

        // read, input to webpage and press go button
        private void btnSelect_Click(object sender, EventArgs e)
        { 
            ReadExcel();
            btnOut.Enabled = false;
        }

        // search for id batchs
        private void btnRun_Click(object sender, EventArgs e)
        {
            Thread thread = new Thread(() => ProcessData());
            thread.Start();
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
            ExcelPackage package = null;
            ExcelWorksheet workSheet = null;
            workSheet = ExcelToll.Load(ref package, "SHIPPED", txtFormat.Text);
            deliveries.Clear();
            if (workSheet == null)
            {
                Log("Failed to load worksheet ");
                return;
            }

            // adds all rows that are are missing the date delivered
            if (isUpdate.Checked)
            {
                // read columns into lists
                var invoiceIds = ExcelToll.GetColumn(workSheet, "INVOICE #");
                var customerPOs = ExcelToll.GetColumn(workSheet, "CUSTOMER PO #");
                var conIds = ExcelToll.GetColumn(workSheet, "CONSIGNMENT REFERENCE");
                var dates = ExcelToll.GetColumn(workSheet, "DATE DELIVERED");
                var courier = ExcelToll.GetColumn(workSheet, "COURIER");
                if (invoiceIds == null || customerPOs == null || conIds == null || dates == null || courier == null)
                {
                    Log("Failed to find all columns with the correct names");
                    return;
                }

                // row 94 failed in excel range
                // changed to default for loop
                for (int i = 0; i < conIds.Count; i++)
                {
                    if (dates[i] == "")
                    {
                        // Console.WriteLine($"{i + 3} {customerPOs[i]} {invoiceIds[i]} {conIds[i]} {courier[i]}");
                        deliveries.Add(new Delivery(customerPOs[i], invoiceIds[i], conIds[i], "Unknown", courier[i], new DateTime()));
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
                    deliveries.Add(new Delivery(customerPOs[i], invoiceIds[i], conIds[i], "Unknown", "Toll", new DateTime()));
                }
            }

            // remove certain entries
            int num = 0;
            deliveries = deliveries.Distinct().ToList();
            deliveries.ForEach(i => i.invoiceID = i.invoiceID.Replace("GS", ""));
            deliveries.RemoveAll(i => i.invoiceID.ToUpper() == "SAMPLES" || i.invoiceID.ToUpper() == "REPLACEMENT");
            deliveries.RemoveAll(i => i.conID.ToUpper() == "TRANSFER" || int.TryParse(i.conID, out num));
            // deliveries.RemoveAll(i => !i.conID.Contains("ARE"));

            // split into separate lists matched by courier
            // makes it easier to process
            testing = deliveries.GroupBy(i => i.courier).Select(grp => grp.ToList()).ToList();

            Log($"Done Loading input {deliveries.Count}");
            btnRun.Enabled = true;              
        }

        // process all data from input
        // run on a separate thread to prevent blocking
        public void ProcessData()
        {
            if (testing == null)
                return;

            // for each delivery group
            foreach (var list in testing)
            {
                Invoke(new Action(() =>
                {
                    processBar.Value = 0;
                    processBar.Minimum = 0;
                    processBar.Maximum = list.Count;
                    ConsignmentIdIndex = 0;
                }));

                // process based on courier
                var courier = list[0].courier;
                switch (courier)
                {
                    case "Toll":
                        ProcessToll(list);
                        break;
                    case "NZ COURIER ":
                        ProcessNZC(list);
                        break;
                    case "NZC":
                        ProcessNZC(list);
                        break;
                    case "PBT":
                        ProcessPBT(list);
                        break;
                }
            }
            Invoke(new Action(() =>
            {
                btnOut.Enabled = true;
            }));
        }

        public void LoadPage(string url)
        {
            // load toll url
            Log("Using page " + url);
            loaded = false;
            webBrowser.Load(url);
            while (!loaded) Thread.Sleep(500);
        }

        // 10 ids at a time(search button)
        public void ProcessToll(List<Delivery> data)
        {
            LoadPage(TollURL);
            int request = 10;
            for (int i = 0; i < data.Count; i += request)
            {
                // get next 10 ids string
                var batch = data.Skip(i).Take(10).ToList();
                string trackingIds = "";
                foreach(var delivery in batch)
                {             
                    trackingIds += $"{delivery.conID}" + Environment.NewLine;
                }

                // search for ids
                Log("Searching for consignmment ids");
                var search = $@"document.getElementById('connoteIds').value = `{trackingIds}`; $('#tatSearchButton').click()";
                RunJS(search);

                var command = @"(function () {
                    return $('#loadingPopUpDialogId').css('display') === 'none';
                })();";

                // wait for result
                while (true)
                {
                    var result = RunJS(command);
                    if (result.ToUpper() == "TRUE")
                    {
                        Log("Result found!");
                        GetDeliveries(data, "TOLL");
                        break;
                    }
                    Thread.Sleep(200);
                }
            }
            Log("Finished processing Toll");
        }

        // 1 id at a time(get request)
        public void ProcessNZC(List<Delivery> data)
        {
            Log("Using page " + NZCURL);
            for (int i = 0; i < data.Count; i++)
            {
                // load nzc url passing conId
                LoadPage(NZCURL + data[i].conID);
                Log("Result found!");
                GetDeliveries(data, "NZC");
            }
            Log("Finished processing NZC");
        }

        // 1 id at a time(search bar)
        public void ProcessPBT(List<Delivery> data)
        {
            LoadPage(PDTURL);
            for (int i = 0; i < data.Count; i++)
            {
                var search = $"document.getElementById('TicketNo').value = '{data[i].conID}'; checkFC();";
                RunJS(search);

                var command = @"(function () {
                    var chil = document.getElementsByTagName('table');
                    if (chil != null)
                    {
                        console.log(chil.length);
                        if (chil.length >= 10)
                        {
                            return true;
                        }
                    }
                    return false;
                })();";

                // wait for result
                while (true)
                {     
                    // if (CurrentURL == "http://www.pbt.co.nz/track/PBTresults_transport.cfm")
                    var result = RunJS(command);
                    if (result.ToUpper() == "TRUE")
                    {
                        Log("Result found");
                        GetDeliveries(data, "PBT");
                        // CurrentURL = "";
                        break;
                    }
                    Thread.Sleep(200);
                }
            }
            Log("Finished processing PDT");
        }

        // TODO: separate result commands
        // Store results from run webpage in consignmentIds
        private void GetDeliveries(List<Delivery> batch, string type) // nzc temp
        {
            // magic date formating
            var commmand = @"(function(){
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
            var NZCCommand = @"(function(){
                var ret = [];
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
                ret.push({key: consignment, value: date.toISOString()})
                return JSON.stringify(ret, null, 3);
            })();";

            var PBTCommand = @"(function(){
                var chil = document.getElementsByTagName('table')[10].children[0].children[0].children[0].children[1].children[0].children[0].children;
                var last = chil[chil.length - 2];
                var date = last.children[0].innerText;
                var time = last.children[1].innerText;
                var splitDate = date.split('/');
                var splitTime = time.split(':');
                var dateDate = new Date(splitDate[2],splitDate[1],splitDate[0],splitTime[0],splitTime[1]);
                var ret = [];
                ret.push(({key: ticketNo, value: dateDate.toISOString()}));
                return JSON.stringify(ret,null,3);
            })();";

            Log("Storing delivery results");
            if (type == "TOLL")
            {
                var result = RunJS(commmand);
                FormatOutput(result, batch, 10);
            }
            else if(type == "NZC")
            {
                var result = RunJS(NZCCommand);
                FormatOutput(result, batch, 1);
            }
            else if(type == "PBT")
            {
                var result = RunJS(PBTCommand);
                FormatOutput(result, batch, 1);
            }
        }

        // deserialize json result and add to list
        private void FormatOutput(string s, List<Delivery> batch, int increment)
        {
            var output = s.FromJson<TrackingResult>();
            if (output == null)
            {
                Log("Failed to deserialize Tracking result");
            }
            else
            {
                output = output.Select(o => new TrackingResult { Key = o.Key, Value = o.Value.ToLocalTime() }).ToList();
                for (int i = 0; i < output.Count; i++)
                {
                    batch[ConsignmentIdIndex + i].date = output[i].Value;
                }
            }
            ConsignmentIdIndex = Math.Min(ConsignmentIdIndex + increment, batch.Count);

            // output progress
            Invoke(new Action(() =>
            {
                processBar.Increment(increment);
            }));
            Log($"Processing... {ConsignmentIdIndex}/{batch.Count} ({((float)ConsignmentIdIndex / (float)batch.Count) * 100f:F2}%)");
        }

        private void OutputToExcel()
        {       
            ExcelPackage package = null;
            ExcelWorksheet workSheet = null;
            workSheet = ExcelToll.Load(ref package, txtFormat.Text, txtFormat.Text);
            if (workSheet == null)
            {
                Log("Failed to load worksheet ");
                return;
            }

            // invoice range to look for matches
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
                    if (cell.Start.Row == 108)
                    {
                        Console.WriteLine();
                    }

                    if (id == "NZ112001")
                    {
                        Console.WriteLine();
                    }

                    donelist.Add(delivery);
                    matches++;
                    workSheet.Cells[cell.Start.Row, customerPO].Value = delivery.customerPO;
                    workSheet.Cells[cell.Start.Row, invoiceNO].Value = delivery.invoiceID;
                    workSheet.Cells[cell.Start.Row, conId].Value = delivery.conID;
                    workSheet.Cells[cell.Start.Row, date].Value = delivery.date.ToShortDateString();
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
    }
}
