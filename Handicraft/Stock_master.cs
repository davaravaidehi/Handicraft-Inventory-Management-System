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
    public partial class Stock_master : Form
    {
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\@g_19\Handicraft\login.accdb");
        OleDbCommand cm;
        OleDbDataReader dr;
        
        public Stock_master()
        {
            InitializeComponent();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do You Really Want To Exit ??", "Message", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                this.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            id_com.Text = "";
            nm_cmb.Text = "";
        }

        private void button4_MouseHover(object sender, EventArgs e)
        {

        }
        public void display()
        {
             con.Open();
            String q = "select * from stock_master";
            cm = new OleDbCommand(q, con);
            var reader = cm.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(reader);
            dataGridView1.DataSource = dt;
            con.Close();
        }
       

        private void Stock_master_Load(object sender, EventArgs e)
        {
            display();
            label18.Text = DateTime.Now.ToShortDateString();
            label14.Text = DateTime.Now.ToShortTimeString();
           
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            radioButton3.Enabled = true;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                
                con.Open();
                String s1 = "select item_id from stock_master";
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

        private void id_com_SelectedIndexChanged(object sender, EventArgs e)
        {
            con.Open();
            String s = "select * from stock_master where item_id = " + id_com.Text + "";
            cm = new OleDbCommand(s, con);
            var reader = cm.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(reader);
            dataGridView1.DataSource = dt;
            con.Close();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
                
                con.Open();
                String s1 = "select item_nm from stock_master";
                cm = new OleDbCommand(s1, con);
                dr = cm.ExecuteReader();
                while (dr.Read())
                {
                    nm_cmb.Items.Add(dr[0].ToString());

                }
                id_com.Visible = false;
                nm_cmb.Visible = true;
                con.Close();
            }
            if (radioButton2.Checked == false)
            {
                nm_cmb.Items.Clear();
            }
        }

        private void nm_cmb_SelectedIndexChanged(object sender, EventArgs e)
        {
            con.Open();
            String s = "select * from stock_master where item_nm = '" + nm_cmb.Text + "'";
            cm = new OleDbCommand(s, con);
            var reader = cm.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(reader);
            dataGridView1.DataSource = dt;
            con.Close();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            id_com.Visible = false;
            nm_cmb.Visible = false;
            display();
        }
    }
}
