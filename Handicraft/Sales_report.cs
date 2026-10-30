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
    public partial class Sales_report : Form
    {
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\@g_19\Handicraft\login.accdb");
        OleDbCommand cm;
        OleDbDataReader dr;
        OleDbDataAdapter adp = new OleDbDataAdapter();
        DataSet ds = new DataSet();
        public Sales_report()
        {
            InitializeComponent();
        }

        private void Sales_report_Load(object sender, EventArgs e)
        {
            label18.Text = DateTime.Now.ToShortDateString();
            label14.Text = DateTime.Now.ToShortTimeString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string st = Application.StartupPath + "\\Report\\Sales_report.rpt";
            axCrystalReport1.ReportFileName = st;
            if (radioButton1.Checked == true)
            {
                string str = Application.StartupPath + "\\Report\\Sales_report.rpt";
                axCrystalReport1.ReportFileName = str;
                axCrystalReport1.SelectionFormula = "{sales_master.sales_id}=" + id_com.SelectedItem;
            }
            else if (radioButton2.Checked == true)
            {
                string str = Application.StartupPath + "\\Report\\Sales_report.rpt";
                axCrystalReport1.ReportFileName = str;
                axCrystalReport1.SelectionFormula = "{sales_master.item_nm}='" + nm_cmb.SelectedItem + "'";
            }
            else
            {
                axCrystalReport1.SelectionFormula = "{sales_master.sales_id}>0";
            }
            axCrystalReport1.WindowState = Crystal.WindowStateConstants.crptMaximized;
            //  axCrystalReport1.RetrieveDataFiles();
            axCrystalReport1.WindowShowRefreshBtn = true;
            axCrystalReport1.Action = 1;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\@g_19\Handicraft\login.accdb");
            cm = new OleDbCommand();
            if (radioButton2.Checked == true)
            {

                nm_cmb.Items.Clear();
                nm_cmb.Visible = true;
                cm = new OleDbCommand("select item_nm from sales_master", con);
                adp = new OleDbDataAdapter(cm);
                ds = new DataSet();
                adp.Fill(ds);
                id_com.Visible = false;
                nm_cmb.Visible = true;
                con.Close();
                int a = Convert.ToInt32(ds.Tables[0].Rows.Count.ToString());
                for (int i = 0; i < a; i++)
                {
                    nm_cmb.Items.Add(ds.Tables[0].Rows[i][0].ToString());
                }
            }
            else
            {
                nm_cmb.Visible = false;
                nm_cmb.Text = "";
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                id_com.Visible = true;
                nm_cmb.Visible = false;
                con.Open();
                String s1 = "select sales_id from sales_master";
                cm = new OleDbCommand(s1, con);
                dr = cm.ExecuteReader();
                while (dr.Read())
                {
                    id_com.Items.Add(dr[0].ToString());

                }
                id_com.Visible = true;
                nm_cmb.Visible = false;
                con.Close();
            }
            if (radioButton1.Checked == false)
            {
                id_com.Items.Clear();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do You Really Want To Exit??", "message", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                this.Close();
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
