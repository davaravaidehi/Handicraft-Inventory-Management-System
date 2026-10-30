using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Handicraft
{
    public partial class Purchase_report : Form
    {
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\@g_19\Handicraft\login.accdb");
        OleDbCommand cm;
       // OleDbDataReader dr;
        OleDbDataAdapter adp;
        DataSet ds = new DataSet();
        DataTable dt1;
        public Purchase_report()
        {
            InitializeComponent();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                id_com.Visible = true;
                nm_cmb.Visible = false;
            }
            
        }

        private void axCrystalReport1_Enter(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true || radioButton2.Checked == true || radioButton3.Checked == true)
            {
                axCrystalReport1.WindowShowRefreshBtn = true;
                string st;

                if (radioButton1.Checked)
                {
                    st = Application.StartupPath + "\\Report\\Purchase.rpt";
                    axCrystalReport1.SelectionFormula = "{purchase_master.pur_id}=" + id_com.Text + "";
                    axCrystalReport1.ReportFileName = st;
                }

                else if (radioButton2.Checked)
                {
                    st = Application.StartupPath + "\\Report\\Purchase.rpt";
                    axCrystalReport1.SelectionFormula = "{purchase_master.item_nm}='" + nm_cmb.Text + "'";
                    axCrystalReport1.ReportFileName = st;
                }
                else
                {
                    st = Application.StartupPath + "\\Report\\Purchase.rpt";
                    axCrystalReport1.SelectionFormula = "{purchase_master.pur_id}>0";
                    axCrystalReport1.ReportFileName = st;
                }
                axCrystalReport1.Connect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\\@g_19\\Handicraft\\login.accdb";
                axCrystalReport1.WindowState = Crystal.WindowStateConstants.crptMaximized;
                axCrystalReport1.WindowShowRefreshBtn = true;
                axCrystalReport1.Refresh();
                axCrystalReport1.Action = 1;
                nm_cmb.Visible = false;
                id_com.Visible = false;
            }

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

            if (radioButton2.Checked )
            {
                id_com.Visible = false;
                nm_cmb.Visible = true;
            }
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do You Really Want To Exit??", "message", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                this.Close();
            }
        }

        private void Purchase_report_Load(object sender, EventArgs e)
        {
            radioButton1.Checked = false;
            adp = new OleDbDataAdapter();
            cm = new OleDbCommand("select * from purchase_master", con);
            adp = new OleDbDataAdapter(cm);
            dt1 = new DataTable();
            adp.Fill(dt1);
            id_com.DataSource = dt1;
            id_com.DisplayMember = "purchase_master";
            id_com.ValueMember = "pur_id";

            nm_cmb.DataSource = dt1;
            nm_cmb.DisplayMember = "purchase_master";
            nm_cmb.ValueMember = "item_nm";
        }
    }
}
