using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

// TODO TollTrackv2:
// cross platform open file dialog?
// build using mono/look into .NET standard.
namespace TollTrackV2
{
    /// <summary>
    /// Stores information on a delivery
    /// </summary>
    public class Delivery
    {
        public string conID;
        public string courier;
        public string customerPO;
        public DateTime date;
        public string invoiceID;
        public DateTime pickup;
        public string pieces;
        public string status;

        public Delivery(string customerPO, string invoiceID, string conID, string status, string courier,
            DateTime date, DateTime pickup = new DateTime(), string pieces = "")
        {
            this.customerPO = customerPO;
            this.invoiceID = invoiceID;
            this.conID = conID;
            this.status = status;
            this.date = date;
            this.courier = courier;
            this.pickup = pickup;
            this.pieces = pieces;
        }
    }

    /// <summary>
    /// Main process
    /// </summary>
    public partial class TollTrack
    {
        // general
        private ChromeDriver driver;
        private List<Delivery> deliveries = new List<Delivery>();
        private List<List<Delivery>> groupedDeliveries;
        private List<Task> tasks = new List<Task>();
        private const int maxPerRequest = 10;
        private int consignmentIdIndex;
        private string txtFormat = "BNMA";

        // urls
        // TollUrl deprecated(still use but redirects to MYTOLL
        private readonly string MYTOLLURL = @"https://mytoll.com/";
        private readonly string TollURL = @"https://online.toll.com.au/trackandtrace/";
        // todo: concat the consignmentID to this string and then open in CEF
        private readonly string NZCURL = "http://www.nzcouriers.co.nz/nzc/servlet/ITNG_TAndTServlet?page=1&VCCA=Enabled&Key_Type=BarCode&barcode_data=";
        private readonly string PBTURL = "http://www.pbt.co.nz/default.aspx";

        public TollTrack()
        {
            Console.WriteLine("Welcome to TollTrackV2!");
            Console.WriteLine("-----------------------");
            Console.WriteLine();

            // @"C:\Users\Chris\Documents\other\Ben\2018 Customer PO Register 6.xlsx";
            // @"C:\Users\Chris\Documents\other\Ben\2018 output.xlsx";
            Console.Write("Enter the file path to input: ");
            var input = Console.ReadLine();

            Console.Write("Enter the file path to output: ");
            var output = Console.ReadLine();

            ReadExcel(input);

            // headless = no window
            var options = new ChromeOptions();
            // options.AddArguments("headless");
            driver = new ChromeDriver(".");
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(60);
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(3600);
            driver.Navigate().GoToUrl(TollURL);
            ProcessData();
            driver.Quit();

            OutputExcel(output);

            Console.WriteLine("-----------------------");
            Console.WriteLine("Thanks for using TollTrackV2!");
        }

        private void githubToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("https://github.com/9ParsonsB/TollTrack");
        }

        // add output to textbox
        private void Log(string str)
        {
            if (string.IsNullOrWhiteSpace(str))
                return;
            Console.WriteLine(str);
        }

        private void ReadExcel(string filePath)
        {
            var isNZInput = false;
            var isAutoUpdate = false;
            ExcelPackage package = null;
            ExcelWorksheet workSheet = null;
            
            // TODO: cross platform file dialog
            workSheet = ExcelToll.Load(ref package, filePath, "SHIPPED");
            //deliveries.Clear();*/

            // only continue if excel file loaded
            if (package == null)
            {
                return;
            }

            workSheet = ExcelToll.Load(ref package, filePath, "SHIPPED") ??
                        ExcelToll.Load(ref package, filePath, "BNMA") ??
                        ExcelToll.Load(ref package, filePath, "BNM STATS") ??
                        ExcelToll.Load(ref package, filePath, "ABBOTTS STATS");
            
            if (workSheet == null)
            {
                // loads packages multiple times
                workSheet = ExcelToll.Load(ref package, filePath, "BNMA") ?? ExcelToll.Load(ref package, filePath, "BNM STATS")?? ExcelToll.Load(ref package, filePath, "ABBOTTS STATS");
                workSheet = ExcelToll.GetWorksheet(package, "BNMA") ?? ExcelToll.GetWorksheet(package, "BNM STATS") ??
                            ExcelToll.GetWorksheet(package, "ABBOTTS STATS");

                if (package.Workbook.Worksheets.Any(w => w.Name.ToUpper() == "BNMA"))
                {
                    // if there is a worksheet called BNMA /BNMNZ / BMA then it is reprocessing.
                    isAutoUpdate = true;
                    var work = package.Workbook.Worksheets.FirstOrDefault(w =>
                        string.Equals(w.Name, txtFormat, StringComparison.CurrentCultureIgnoreCase));
                    workSheet = work ?? package.Workbook.Worksheets.First();
                }
                else if (package.Workbook.Worksheets.Any(w =>
                    w.Name.ToUpper() == "BNM STATS" || w.Name.ToUpper() == "ABBOTTS STATS"))
                {
                    isNZInput = true;
                }
                else
                {
                    Log("Failed to load worksheet ");
                    return;
                }
            }

            if (package.Workbook.Worksheets.Any(w => w.Name.ToUpper() == "BNMA"))
            {
                // if there is a worksheet called BNMA /BNMNZ / BMA then it is reprocessing.
                isAutoUpdate = true;
                var work = package.Workbook.Worksheets.FirstOrDefault(w =>
                  string.Equals(w.Name, txtFormat, StringComparison.CurrentCultureIgnoreCase));
                workSheet = work ?? package.Workbook.Worksheets.First();
            }
            else if (package.Workbook.Worksheets.Any(w =>
                w.Name.ToUpper() == "BNM STATS" || w.Name.ToUpper() == "ABBOTTS STATS"))
            {
                isNZInput = true;
            }

            // adds all rows that are are missing the date delivered
            if (isAutoUpdate)
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
                for (var i = 0; i < conIds.Count; i++)
                {
                    if (i >= dates.Count) continue;
                    if (dates[i] == "")
                    {
                        if (string.IsNullOrWhiteSpace(conIds[i])) continue;
                        // Console.WriteLine($"{i + 3} {customerPOs[i]} {invoiceIds[i]} {conIds[i]} {courier[i]}");
                        deliveries.Add(new Delivery(customerPOs[i], invoiceIds[i], conIds[i], "Unknown", courier[i],
                            new DateTime()));
                    }
                }
            }
            else if (isNZInput)
            {
                workSheet = package.Workbook.Worksheets.First(w =>
                    w.Name.ToUpper() == "BNM STATS" || w.Name.ToUpper() == "ABBOTTS STATS");
                var consignmentIds = ExcelToll.GetColumn(workSheet, "Order #");
                var courier = ExcelToll.GetColumn(workSheet, "Carrier");
                var pickup = ExcelToll.GetColumn(workSheet, "First scan Date");
                var pieces = ExcelToll.GetColumn(workSheet, "No. of Cartons");

                if (consignmentIds == null || courier == null || pickup == null)
                {
                    Log("Failed to find all columns with the correct names");
                    return;
                }

                for (var i = 0; i < consignmentIds.Count; i++)
                {
                    if (string.IsNullOrWhiteSpace(consignmentIds[i])) continue;
                    deliveries.Add(new Delivery(string.Empty, "NZ" + consignmentIds[i].Substring(3), consignmentIds[i],
                        "Unknown", courier[i], new DateTime(),
                        DateTime.ParseExact(pickup[i], "d/MM/yyyy", new DateTimeFormatInfo()), pieces[i]));
                }
            }
            // adds all deliveries by default
            else
            {
                // read columns into lists
                var invoiceIds = ExcelToll.GetColumn(workSheet, "PACKSLIP");
                var customerPOs = ExcelToll.GetColumn(workSheet, "CUST REF");
                var conIds = ExcelToll.GetColumn(workSheet, "CON NOTE NUMBER");
                var pieces = ExcelToll.GetColumn(workSheet, "Cartons");
                var pickup = ExcelToll.GetColumn(workSheet, "Shipped Date");

                if (invoiceIds == null || customerPOs == null || conIds == null || pickup == null || pieces == null)
                {
                    Log("Failed to find all columns with the correct names");
                    return;
                }

                for (var i = 0; i < pickup.Count; i++)
                    if (string.IsNullOrWhiteSpace(pickup[i]))
                        pickup[i] = pickup[i - 1];

                for (var i = 0; i < conIds.Count; i++)
                {
                    if (string.IsNullOrWhiteSpace(conIds[i])) continue;
                    deliveries.Add(new Delivery(customerPOs[i],
                        invoiceIds[i],
                        conIds[i],
                        "Unknown",
                        "Toll",
                        new DateTime(),
                        DateTime.ParseExact(pickup[i],
                            "d/MM/yyyy",
                            new DateTimeFormatInfo()),
                        pieces[i])
                    );
                }
            }

            // remove certain entries
            var num = 0;
            deliveries = deliveries.GroupBy(d => d.conID).Select(group => group.First()).ToList();
            deliveries.ForEach(i => i.invoiceID = i.invoiceID.Replace("GS", ""));
            //deliveries.ForEach(i => i.invoiceID = i.invoiceID.Replace("NZ", ""));
            deliveries.RemoveAll(i =>
                i.invoiceID.ToUpper() == "SAMPLES" || i.invoiceID.ToUpper() == "PLES" ||
                i.invoiceID.ToUpper() == "REPLACEMENT");
            deliveries.RemoveAll(i => i.conID.ToUpper() == "TRANSFER" || int.TryParse(i.conID, out num));
            // deliveries.RemoveAll(i => !i.conID.Contains("ARE"));

            // split into separate lists matched by courier
            // makes it easier to process
            groupedDeliveries = deliveries.GroupBy(i => i.courier).Select(grp => grp.ToList()).ToList();

            Log($"Done Loading input {deliveries.Count}");
        }

        // process all data from input
        // run on a separate thread to prevent blocking
        public void ProcessData()
        {
            if (groupedDeliveries == null)
                return;

            // for each delivery group
            foreach (var list in groupedDeliveries)
            {
                // process based on courier
                var courier = list[0].courier;
                switch (courier.ToUpper())
                {
                    case "TOLL":
                        ProcessToll(list);
                        break;
                    case "NZ COURIER":
                    case "NZC":
                        ProcessNZC(list);
                        break;
                    case "PBT":
                        ProcessPBT(list);
                        break;
                }
            }
        }

        // 10 ids at a time(search button)
        public void ProcessToll(List<Delivery> data)
        {
            var command = @"(function () {
                    return $('#loadingPopUpDialogId').css('display') === 'none';
                })();";

            // magic date formating
            var Toll = @"(function(){
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

                    var date = new Date(splitDate[2], splitDate[1] -1, splitDate[0], hour, splitTime[1]);
                    ret.push({ key: rows[i].children[0].children[0].innerText, value: date.toISOString()});  
                }   
                return JSON.stringify(ret,null,3);
            })();";

            //LoadPage(TollURL);
            var request = 30;
            for (var i = 0; i < data.Count; i += request)
            {
                // get next set of ids string
                var batch = data.Skip(i).Take(request).ToList();
                var trackingIds = "";
                foreach (var delivery in batch)
                    trackingIds += $"{delivery.conID}" + Environment.NewLine;

                // search for ids
                Log("Searching for consignmment ids");
                var search =
                    $@"document.getElementById('quickSearch').value = `{trackingIds}`;";// $('#tatSearchButton').click()";

                driver.ExecuteScript(search);

                while (driver.FindElements(By.Id("quickSearchTableResult")).Count == 0)
                {
                
                }

                var result = driver.ExecuteScript(Toll);
                Console.WriteLine(result);
                // var list = driver.FindElements(By.Id("quickSearchTableResult"));
                // Console.WriteLine("FWUAHAHAHA");
                // GetDeliveries(Toll, data, request);

                Thread.Sleep(1000);

            }
            Log("Finished processing Toll");
        }

        // 1 id at a time(get request)
        public void ProcessNZC(List<Delivery> data)
        {
            Log("Using page " + NZCURL);

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

                var date = new Date(splitRaw[2],splitRaw[1] -1,splitRaw[0],hour,minute)
                var consignment = new URL(window.location.href).searchParams.get('barcode_data')
                ret.push({key: consignment, value: date.toISOString()})
                return JSON.stringify(ret, null, 3);
            })();";

            for (var i = 0; i < data.Count; i++)
            {
                // load nzc url passing conId
                // LoadPage(NZCURL + data[i].conID);
                Log("Result found!");
                GetDeliveries(NZCCommand, data, 1);
            }
            Log("Finished processing NZC");
        }

        // 1 id at a time(search bar)
        public void ProcessPBT(List<Delivery> data)
        {
            // finds table on main page before search
            // need to check if results found
            var command = @"(function () {
                    if (document.getElementsByTagName('frame').length > 0)
                        return document.getElementsByTagName('frame')[0].contentDocument.getElementsByTagName('table').length > 9;
                    return false;
                })();";

            var PBTCommand = @"(function(){
                var chil = document.getElementsByTagName('frame')[0].contentDocument.getElementsByTagName('table')[10].children[0].children[0].children[0].children[1].children[0].children[0].children;
                var last = chil[chil.length - 2];
                var date = last.children[0].innerText;
                var time = last.children[1].innerText;
                var splitDate = date.split('/');
                var splitTime = time.split(':');
                var dateDate = new Date(splitDate[2],splitDate[1] -1,splitDate[0],splitTime[0],splitTime[1]);
                var ret = [];
                ret.push(({key: ticketNo, value: dateDate.toISOString()}));
                return JSON.stringify(ret,null,3);
                })();";

            Log("Finished processing PDT");
        }

        // store results from run webpage
        private void GetDeliveries(string command, List<Delivery> batch, int count)
        {
            Log("Storing delivery results");
            //var result = RunJS(command);
            //FormatOutput(result, batch, count);
        }

        // deserialize json result and add to list
        private void FormatOutput(string s, List<Delivery> batch, int count)
        {
            var output = s.FromJson<TrackingResult>();
            if (output == null)
            {
                Log("Failed to deserialize Tracking result");
            }
            else
            {
                output = output.Select(o => new TrackingResult {Key = o.Key, Value = o.Value.ToLocalTime()}).ToList();
                for (var i = 0; i < output.Count; i++)
                    batch[consignmentIdIndex + i].date = output[i].Value;
                Log($"Found {output.Count} consignments");
            }

            consignmentIdIndex = Math.Min(consignmentIdIndex + count, batch.Count);
            foreach (var l in batch)
            {
                deliveries[deliveries.IndexOf(deliveries.First(d => d.conID == l.conID))] = l;
            }
            Log(
                $"Processing Toll: {batch.Count(b => b.date != default)}/{consignmentIdIndex}/{batch.Count} ({(float) consignmentIdIndex / (float) batch.Count * 100f:F2}%)");
        }

        private void OutputExcel(string filePath)
        {
            ExcelPackage package = null;
            ExcelWorksheet workSheet = null;

            workSheet = ExcelToll.Load(ref package, filePath, txtFormat);
            if (workSheet == null)
            {
                Log("Failed to load worksheet ");
                return;
            }

            // output column locations
            var customerPO = ExcelToll.GetCellColumn(workSheet, "CUSTOMER PO #", 0);
            var invoiceNO = ExcelToll.GetCellColumn(workSheet, "INVOICE #", 0);
            var conId = ExcelToll.GetCellColumn(workSheet, "CONSIGNMENT REFERENCE", 0);
            var date = ExcelToll.GetCellColumn(workSheet, "DATE DELIVERED", 0);
            var pickup = ExcelToll.GetCellColumn(workSheet, "Date of Pickup", 0);
            var pieces1 = ExcelToll.GetCellColumn(workSheet, "Pieces", 0);
            var pieces2 = ExcelToll.GetCellColumn(workSheet, "Pieces", 1);
            var anspec = ExcelToll.GetCellColumn(workSheet, "Anspec Date", 0);
            var courier = ExcelToll.GetCellColumn(workSheet, "Courier", 0);

            anspec = anspec == 0 ? ExcelToll.GetCellColumn(workSheet, "DBS Date", 0) : anspec;

            // prevent crash if a column is missing
            if (customerPO == 0 || invoiceNO == 0 || conId == 0 || date == 0 || pickup == 0 || anspec == 0)
            {
                Log("Failed to find one of the columns");
                return;
            }

            var donelist = new List<Delivery>();
            var matches = 0;
            FindMatchesByInvoiceID(ExcelToll.GetColumnRange(workSheet, "INVOICE #"), workSheet, anspec, pickup, ref matches, donelist, courier, pieces1, pieces2, customerPO, invoiceNO, conId, date);

            FindMatchesByCustomerPo( ExcelToll.GetColumnRange(workSheet, "CUSTOMER PO #") , workSheet, anspec, pickup, ref matches, donelist, courier, pieces1, pieces2, customerPO, invoiceNO, conId, date);

            if (donelist.Count == deliveries.Count)
                return;
            var notDone = deliveries.Where(d => !donelist.Contains(d)).ToList();
            deliveries = notDone;
            
            Log($"{matches} matches updated");
            package.Save();

            // show details of deliveries not found in output for manual assignment
            // only process the orders that could not be matched to speed up future processing       
        }

        private void FindMatchesByCustomerPo(ExcelRange range, ExcelWorksheet workSheet, int anspec, int pickup, ref int matches,
            List<Delivery> donelist, int courier, int pieces1, int pieces2, int customerPO, int invoiceNO, int conId, int date)
        {
            foreach (var cell in range)
            {
                // copy date to anspec date
                if (string.IsNullOrWhiteSpace(workSheet.Cells[cell.Start.Row, anspec].Text))
                {
                    if (!string.IsNullOrWhiteSpace(workSheet.Cells[cell.Start.Row, pickup].Text))
                    {
                        workSheet.Cells[cell.Start.Row, anspec].Value = workSheet.Cells[cell.Start.Row, pickup].Text;
                    }
                }

                // update data where id matches
                if (deliveries.Count == 0)
                    continue;
                var id = cell.Text ?? "";
                if (id == "")
                    continue;
                
                var delivery = deliveries.FirstOrDefault(i =>
                {
                    if (i.customerPO == id)
                        return true;
                    if (i.customerPO.Length > 1)
                        return i.customerPO.Substring(2) == id;
                    return false;
                });
                matches = WriteMatch(delivery, donelist, matches, workSheet, cell, courier, pieces1, pieces2, anspec, pickup,
                    customerPO, invoiceNO, conId, date);
            }
        }

        private void FindMatchesByInvoiceID(ExcelRange range, ExcelWorksheet workSheet, int anspec, int pickup, ref int matches,
            List<Delivery> donelist, int courier, int pieces1, int pieces2, int customerPO, int invoiceNO, int conId, int date)
        {
            foreach (var cell in range)
            {
                // copy date to anspec date
                if (string.IsNullOrWhiteSpace(workSheet.Cells[cell.Start.Row, anspec].Text))
                {
                    if (!string.IsNullOrWhiteSpace(workSheet.Cells[cell.Start.Row, pickup].Text))
                    {
                        workSheet.Cells[cell.Start.Row, anspec].Value = workSheet.Cells[cell.Start.Row, pickup].Text;
                    }
                }

                // update data where id matches
                if (deliveries.Count == 0)
                    continue;
                var id = cell.Text ?? "";
                if (id == "")
                    continue;

                var delivery = deliveries.FirstOrDefault(i => i.invoiceID == id || i.invoiceID.Substring(2) == id);
                matches = WriteMatch(delivery, donelist, matches, workSheet, cell, courier, pieces1, pieces2, anspec, pickup,
                    customerPO, invoiceNO, conId, date);
            }
        }

        private static int WriteMatch(Delivery delivery, List<Delivery> donelist, int matches, ExcelWorksheet workSheet,
            ExcelRangeBase cell, int courier, int pieces1, int pieces2, int anspec, int pickup, int customerPO, int invoiceNO,
            int conId, int date)
        {
            if (delivery != null)
            {
                // ignore dates that have not been updated
                if (delivery.date == DateTime.MinValue) return matches;

                var a = new[] {1, 2, 3};
                if (delivery.pieces == "")
                    a = null;

                var b = a?[0] + 1;

                // update matching delivery in spreadsheet
                donelist.Add(delivery);
                matches++;

                // write to cells for the row
                workSheet.Cells[cell.Start.Row, courier].Value = delivery.courier;
                workSheet.Cells[cell.Start.Row, pieces1].Value = delivery.pieces;
                workSheet.Cells[cell.Start.Row, pieces2].Value = delivery.pieces;
                workSheet.Cells[cell.Start.Row, anspec].Value =
                    delivery.pickup == default ? string.Empty : delivery.pickup.ToString("d");
                workSheet.Cells[cell.Start.Row, pickup].Value =
                    delivery.pickup == default ? string.Empty : delivery.pickup.ToString("d");
                workSheet.Cells[cell.Start.Row, customerPO].Value = delivery.customerPO;
                workSheet.Cells[cell.Start.Row, invoiceNO].Value = delivery.invoiceID;
                workSheet.Cells[cell.Start.Row, conId].Value = delivery.conID;
                workSheet.Cells[cell.Start.Row, date].Value = delivery.date.ToShortDateString();
            }
            return matches;
        }
    }
}
