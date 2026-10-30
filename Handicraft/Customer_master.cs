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
using System.Text.RegularExpressions;

namespace Handicraft
{
    public partial class Customer_master : Form
    {
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\@g_19\Handicraft\login.accdb");
        OleDbCommand cm;
        OleDbDataReader dr;
        OleDbDataAdapter da;
        DataSet ds;
        DataTable dt = new DataTable();
        int index;
        public Customer_master()
        {
            InitializeComponent();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void btn_exit_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do You Really Want To Exit ??", "Message", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                this.Close();
            }
        }

        public void clear()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            richTextBox1.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            
        }

        public void enable()
        {
            textBox1.Enabled = true;
            textBox2.Enabled = true;
            textBox3.Enabled = true;
            richTextBox1.Enabled = true;
            textBox5.Enabled = true;
            textBox6.Enabled = true;
            textBox7.Enabled = true;
        }

        public void disable()
        {
            textBox1.Enabled = false;
            textBox2.Enabled = false;
            textBox3.Enabled = false;
            richTextBox1.Enabled = false;
            textBox5.Enabled = false;
            textBox6.Enabled = false;
            textBox7.Enabled = false;
        }

        public void display()
        {
             con.Open();
            String q = "select * from customer_master";
            cm = new OleDbCommand(q, con);
            var reader = cm.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(reader);
            dataGridView1.DataSource = dt;
            con.Close();
        }

        public void auto_inc()
        {
            da = new OleDbDataAdapter("select max(cust_id) from customer_master", con);
            ds = new DataSet();
            da.Fill(ds);

            int cno;

            if (ds.Tables[0].Rows[0][0].ToString().Equals(""))
            {
                cno = 101;
            }
            else
            {
                cno = Int32.Parse(ds.Tables[0].Rows[0][0].ToString());
                cno = cno + 1;
            }
            textBox1.Text = cno.ToString();
        }

        private void btn_cancle_Click(object sender, EventArgs e)
        {
            clear();
        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void btn_add_Click(object sender, EventArgs e)
        {
            enable();
            auto_inc();
            btn_save.Enabled = true;
            btn_cancle.Enabled = true;
        }

        private void btn_save_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("please Enter Cust_id", "Message", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            else
            {
                if (string.IsNullOrEmpty(textBox1.Text.Trim()))
                {
                    errorProvider1.SetError(textBox1, "Customer id is required.");
                    return;
                }
                else
                {
                    errorProvider1.SetError(textBox1, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox2.Text.Trim()))
                {
                    errorProvider2.SetError(textBox2, "Customer name is required.");
                    return;
                }
                else
                {
                    errorProvider2.SetError(textBox2, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox3.Text.Trim()))
                {
                    errorProvider3.SetError(textBox3, "Contact No is required.");
                    return;
                }
                else
                {
                    errorProvider3.SetError(textBox3, string.Empty);
                }
                if (string.IsNullOrEmpty(richTextBox1.Text.Trim()))
                {
                    errorProvider4.SetError(richTextBox1, "Customer Address is required.");
                    return;
                }
                else
                {
                    errorProvider4.SetError(richTextBox1, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox5.Text.Trim()))
                {
                    errorProvider5.SetError(textBox5, "Customer City is required.");
                    return;
                }
                else
                {
                    errorProvider5.SetError(textBox5, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox6.Text.Trim()))
                {
                    errorProvider6.SetError(textBox6, "Pincode is required.");
                    return;
                }
                else
                {
                    errorProvider6.SetError(textBox6, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox7.Text.Trim()))
                {
                    errorProvider7.SetError(textBox7, "Email is required.");
                    return;
                }
                else
                {
                    errorProvider7.SetError(textBox7, string.Empty);
                }
                if (MessageBox.Show("Do You Really Want To Add This Record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.Yes)
                {
                    textBox1.Enabled = true;
                   
                    con.Open();
                    cm = new OleDbCommand("insert into customer_master values('" + textBox1.Text + "','" + textBox2.Text + "','" + richTextBox1.Text + "','" + textBox5.Text + "','" + textBox3.Text + "','" + textBox6.Text + "','" + textBox7.Text + "')", con);
                    cm.Connection = con;
                    cm.ExecuteNonQuery();
                    con.Close();
                    MessageBox.Show("Record Has Been Successfully Saved", "Success", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    clear();
                    display();
                    disable();
                    btn_edit.Enabled = true;
                    btn_delete.Enabled = true;

                }
            }
            
           
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void Customer_master_Load(object sender, EventArgs e)
        {
            btn_add.Enabled = true;
            btn_save.Enabled = false;
            btn_edit.Enabled = false;
            btn_delete.Enabled = false;
            btn_cancle.Enabled = false;
            btn_exit.Enabled = true;
            label18.Text = DateTime.Now.ToShortDateString();
            label15.Text = DateTime.Now.ToShortTimeString();
            display();
        }

        private void btn_edit_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                if (MessageBox.Show("please enter cust_id", "confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.Yes)
                {
                    enable();
                }
            }
            else
            {
                int temp = 0;
                con.Open();
                OleDbCommand cm = new OleDbCommand("select cust_id from customer_master", con);
                dr = cm.ExecuteReader();
                while (dr.Read())
                {
                    if (textBox1.Text == dr.GetInt32(0).ToString())
                    {
                        temp = 1;
                        break;
                    }
                }
                con.Close();
                if (temp == 1)
                {
                    con.Open();
                    String s1 = "Update customer_master set cust_nm = '" + textBox2.Text + "',cust_add = '" + richTextBox1.Text + "',cust_city ='" + textBox5.Text + "',cont_no ='" + textBox3.Text + "',pincode = '" + textBox6.Text + "',email_id = '" + textBox7.Text + "' where cust_id = " + textBox1.Text + "";
                    cm = new OleDbCommand(s1, con);
                    cm.Connection = con;
                    cm.ExecuteNonQuery();
                    con.Close();
                    MessageBox.Show("Record Has Been Successfully update", "Success", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    clear();
                    disable();
                    display();
                }
                else
                {
                    MessageBox.Show("this is new id so can't updated.....plese save record", "Message", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void btn_delete_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {

                if (MessageBox.Show("Please enter id for delete record", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.Yes)
                {
                    textBox1.Enabled = true;
                    disable();
                }
            }

            else
            {

                con.Open();
                string s1 = "select cust_id from customer_master";
                int temp = 0;
                cm = new OleDbCommand(s1, con);
                dr = cm.ExecuteReader();
                while (dr.Read())
                {
                    if (textBox1.Text == dr.GetInt32(0).ToString())
                    {
                        temp = 1;
                        break;
                    }
                }
                if (temp == 1)
                {
                    string q = "delete from customer_master where cust_id=" + textBox1.Text.ToString() + "";
                    cm = new OleDbCommand(q, con);
                    cm.ExecuteNonQuery();
                    con.Close();
                    MessageBox.Show("Record Has Been Successfully deleted", "Success", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    disable();
                    clear();
                    display();
                }
                else
                {
                    MessageBox.Show("this is new id so can't delete.....plese save record", "Message", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            enable();
            btn_delete.Enabled = true;
            btn_edit.Enabled = true;
            btn_cancle.Enabled = true;
            textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            richTextBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            textBox5.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
            textBox6.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            textBox7.Text = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {

                con.Open();
                String s1 = "select cust_id from customer_master";
                cm = new OleDbCommand(s1, con);
                dr = cm.ExecuteReader();
                while (dr.Read())
                {
                    id_cmb.Items.Add(dr[0].ToString());

                }
                id_cmb.Visible = true;
                nm_cmb.Visible = false;
                con.Close();
            }
            if (radioButton1.Checked == false)
            {
                id_cmb.Items.Clear();
            }
        }

        private void id_cmb_SelectedIndexChanged(object sender, EventArgs e)
        {
            con.Open();
            String s = "select * from customer_master where cust_id = " + id_cmb.Text + "";
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
                String s1 = "select cust_nm from customer_master";
                cm = new OleDbCommand(s1, con);
                dr = cm.ExecuteReader();
                while (dr.Read())
                {
                    nm_cmb.Items.Add(dr[0].ToString());

                }
                con.Close();
                id_cmb.Visible = false;
                nm_cmb.Visible = true;
            }
            if (radioButton2.Checked == false)
            {
                nm_cmb.Items.Clear();
            }
        }

        private void nm_cmb_SelectedIndexChanged(object sender, EventArgs e)
        {
            con.Open();
            String s = "select * from customer_master where cust_nm = '" + nm_cmb.Text + "'";
            cm = new OleDbCommand(s, con);
            var reader = cm.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(reader);
            dataGridView1.DataSource = dt;
            con.Close();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            display();
        }

        public void load_data(int c)
        {
            cm = new OleDbCommand("select * from customer_master", con);
            da = new OleDbDataAdapter(cm);
            da.Fill(dt);

            textBox1.Text = dt.Rows[c][0].ToString();
            textBox2.Text = dt.Rows[c][1].ToString();
            richTextBox1.Text = dt.Rows[c][2].ToString();
            textBox5.Text = dt.Rows[c][3].ToString();
            textBox3.Text = dt.Rows[c][4].ToString();
            textBox6.Text = dt.Rows[c][5].ToString();
            textBox7.Text = dt.Rows[c][6].ToString();
        }

        private void first_Click(object sender, EventArgs e)
        {
            enable();
            index = 0;
            load_data(index);
        }

        private void next_Click(object sender, EventArgs e)
        {
            index += 1;
            load_data(index);
        }

        private void previous_Click(object sender, EventArgs e)
        {
            if (index > 0)
            {
                index -= 1;
                load_data(index);
            }
            else
            {
                MessageBox.Show("You are on First Record..");
            }
        }

        private void last_Click(object sender, EventArgs e)
        {
            index = dt.Rows.Count - 1;
            load_data(index);
        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void contact(object sender, EventArgs e)
        {
           
          if (!Regex.Match(textBox3.Text, @"^\d{10}$").Success)
            {
                MessageBox.Show("Enter Only 10 Digit Or Phone Number  ", "Message", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox3.Focus();
            }
        }

        private void email(object sender, EventArgs e)
        {
            if (!Regex.Match(textBox7.Text, @"^([\w\.\-]+)@([\w\-]+)((\.(\w){2,3})+)]*$").Success)
            {
                MessageBox.Show("Invalid Email ID ", "Message", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox7.Focus();
            }
        }

        private void pincode_leave(object sender, EventArgs e)
        {
            if (!Regex.Match(textBox6.Text, @"^\d{6}$").Success)
            {
                MessageBox.Show("Enter Only 6 Digit Pincode Number", "Message", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox6.Focus();
            }

        }
    }
}
