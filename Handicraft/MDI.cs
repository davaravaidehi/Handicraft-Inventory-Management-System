using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Handicraft
{
    public partial class MDI : Form
    {
        public MDI()
        {
            InitializeComponent();
        }

        private void MDI_Load(object sender, EventArgs e)
        {

        }

        private void itemMasterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Item_master i1 = new Item_master();
            i1.Show();
        }

        private void customerMasterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Customer_master c1 = new Customer_master();
            c1.Show();
        }

        private void wORKERMASTERToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Worker_master w1 = new Worker_master();
            w1.Show();
        }

        private void purchaseEntryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Purchase_master p1 = new Purchase_master();
            p1.Show();
        }

        private void saleEntryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Sales_master s1 = new Sales_master();
            s1.Show();
        }

        private void pURCHASERETURNENTRYToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Purchase_return_master pr1 = new Purchase_return_master();
            pr1.Show();
        }

        private void sALESRETURNENTRYToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Sales_return_master sr1 = new Sales_return_master();
            sr1.Show();
        }

        private void stockDetailToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Stock_master sm = new Stock_master();
            sm.Show();
        }

        private void itemReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Item_report ir = new Item_report();
            ir.Show();
        }

        private void customerReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Customer_report cr = new Customer_report();
            cr.Show();
        }

        private void wORKERREPORTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Worker_report wr = new Worker_report();
            wr.Show();
        }

        private void purchaseEntryReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Purchase_report pr = new Purchase_report();
            pr.Show();
        }

        private void pURCHASERETURNREPORTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Purchase_return_report prm = new Purchase_return_report();
            prm.Show();
        }

        private void salesEntryReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Sales_report sr = new Sales_report();
            sr.Show();
        }

        private void sALESRETURNREPORTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Sales_return_report srm = new Sales_return_report();
            srm.Show();
        }

        private void sTOCKREPORTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Stock_report str = new Stock_report();
            str.Show();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do You Really Want To Exit??", "message", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        private void notepadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("Notepad");
        }

        private void calculatorToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("Calc");
        }

        private void wordpadToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("Wordpad");
        }

        private void mozilaFirefoxToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("Chrome");
        }

        private void logOutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
            login l1 = new login();
            l1.Show();
        }

        private void aboutUsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            About_us us = new About_us();
            us.Show();
        }

        private void changePasswordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            change_password cp = new change_password();
            cp.Show();
        }
    }
}
