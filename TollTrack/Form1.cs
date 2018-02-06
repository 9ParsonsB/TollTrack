using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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

        private void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            var trackingIds = "";
            
            consignmentIds.ForEach(c=> trackingIds += $"{c}{Environment.NewLine}");

            var command = $"document.getElementById('connoteIds').innerText = '{trackingIds}'";

            webBrowser.Document?.ExecCommand(command,false,null); // populate text box where IDs are meant to be with some javascript
            webBrowser.Document.GetElementById("table where the results are") // then get the status and
                .GetAttribute("The status for each ID");
            // update the SortedList for each ID

            // write to Excel document
        }
    }
}
