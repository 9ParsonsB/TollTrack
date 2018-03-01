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
    public partial class Form2 : Form
    {
        private List<Form1.Delivery> devliveries;

        public Form2(List<Form1.Delivery> deliveries) : base()
        {
            this.devliveries = deliveries;
        }

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            txtout.Text = "Invoice Id | Customer PO | Consignment | Date | Status";
            devliveries.ForEach(d =>
                {
                    txtout.Text +=
                        $"{d.invoiceID} | {d.customerPO} | {d.conID} | {d.date} | {d.status}{Environment.NewLine}";
                });
        }
    }
}
