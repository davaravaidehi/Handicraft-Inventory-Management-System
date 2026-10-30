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
    public partial class Sales_return_master : Form
    {
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\@g_19\Handicraft\login.accdb");
        OleDbCommand cm;
        OleDbDataReader dr;
        OleDbDataAdapter da;
        DataSet ds;
        DataTable dt = new DataTable();
        int index;
        //string qq, tot, q2;
        public Sales_return_master()
        {
            InitializeComponent();
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
            comboBox1.Text = "";
            dateTimePicker1.Text = "";
            dateTimePicker2.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox7.Text = "";
            textBox6.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
        }

        public void enable()
        {
            textBox1.Enabled = true;
            comboBox1.Enabled = true;
            dateTimePicker1.Enabled = true;
            dateTimePicker2.Enabled = true;
            textBox2.Enabled = true;
            textBox3.Enabled = true;
            textBox4.Enabled = true;
            textBox5.Enabled = true;
            textBox7.Enabled = true;
            textBox6.Enabled = true;
            textBox8.Enabled = true;
            textBox9.Enabled = true;
        }

        public void disable()
        {
            textBox1.Enabled = false;
            comboBox1.Enabled = false;
            dateTimePicker1.Enabled = false;
            dateTimePicker2.Enabled = false;
            textBox2.Enabled = false;
            textBox3.Enabled = false;
            textBox4.Enabled = false;
            textBox5.Enabled = false;
            textBox7.Enabled = false;
            textBox6.Enabled = false;
            textBox8.Enabled = false;
            textBox9.Enabled = false;
        }

        public void display()
        {
             con.Open();
            String q = "select * from sales_return_master";
            cm = new OleDbCommand(q, con);
            var reader = cm.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(reader);
            dataGridView1.DataSource = dt;
            con.Close();
        }

        public void auto_inc()
        {
            da = new OleDbDataAdapter("select max(sales_ret_id) from sales_return_master", con);
            ds = new DataSet();
            da.Fill(ds);

            int cno;

            if (ds.Tables[0].Rows[0][0].ToString().Equals(""))
            {
                cno = 111;
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

        private void Sales_return_master_Load(object sender, EventArgs e)
        {
            btn_add.Enabled = true;
            btn_save.Enabled = false;
            btn_edit.Enabled = false;
            btn_delete.Enabled = false;
            btn_cancle.Enabled = false;
            btn_exit.Enabled = true;
            display();
            label23.Text = DateTime.Now.ToShortDateString();
            label22.Text = DateTime.Now.ToShortTimeString();
            con.Open();
            String s1 = "select sales_id from sales_master";
            cm = new OleDbCommand(s1, con);
            OleDbDataReader dr = cm.ExecuteReader();
            while (dr.Read())
            {
                comboBox1.Items.Add(dr[0].ToString());
            }
            con.Close();
        }

        private void btn_add_Click(object sender, EventArgs e)
        {
            btn_save.Enabled = true;
            btn_cancle.Enabled = true;
            enable();
            auto_inc();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            enable();
            btn_delete.Enabled = true;
            btn_edit.Enabled = true;
            btn_cancle.Enabled = true;
            textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            comboBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            dateTimePicker1.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            dateTimePicker2.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
            textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            textBox4.Text = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();
            textBox5.Text = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString();
            textBox6.Text = dataGridView1.Rows[e.RowIndex].Cells[8].Value.ToString();
            textBox7.Text = dataGridView1.Rows[e.RowIndex].Cells[9].Value.ToString();
            textBox8.Text = dataGridView1.Rows[e.RowIndex].Cells[10].Value.ToString();
            textBox9.Text = dataGridView1.Rows[e.RowIndex].Cells[11].Value.ToString();
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                con.Open();
                String s1 = "select sales_ret_id from sales_return_master";
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
            String s = "select * from sales_return_master where sales_ret_id = " + id_com.Text + "";
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
                String s1 = "select item_nm from sales_return_master";
                cm = new OleDbCommand(s1, con);
                dr = cm.ExecuteReader();
                while (dr.Read())
                {
                    nm_cmb.Items.Add(dr[0].ToString());

                }
                con.Close();
                id_com.Visible = false;
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
            String s = "select * from sales_return_master where item_nm = '" + nm_cmb.Text + "'";
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
       
        private void btn_save_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("please Enter Purchase return ID", "Message", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            else
            {
                if (string.IsNullOrEmpty(textBox1.Text.Trim()))
                {
                    errorProvider1.SetError(textBox1, "sales return ID is required.");
                    return;
                }
                else
                {
                    errorProvider1.SetError(textBox1, string.Empty);
                }
                if (string.IsNullOrEmpty(comboBox1.Text.Trim()))
                {
                    errorProvider2.SetError(comboBox1, "sales Id is required.");
                    return;
                }
                else
                {
                    errorProvider2.SetError(comboBox1, string.Empty);
                }
                if (string.IsNullOrEmpty(dateTimePicker1.Text.Trim()))
                {
                    errorProvider3.SetError(dateTimePicker1, "sales return date is required.");
                    return;
                }
                else
                {
                    errorProvider3.SetError(dateTimePicker1, string.Empty);
                }
                if (string.IsNullOrEmpty(dateTimePicker2.Text.Trim()))
                {
                    errorProvider4.SetError(dateTimePicker2, "sales date is required.");
                    return;
                }
                else
                {
                    errorProvider4.SetError(dateTimePicker2, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox2.Text.Trim()))
                {
                    errorProvider5.SetError(textBox2, "customer name is required.");
                    return;
                }
                else
                {
                    errorProvider5.SetError(textBox2, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox3.Text.Trim()))
                {
                    errorProvider6.SetError(textBox3, "Item name is required.");
                    return;
                }
                else
                {
                    errorProvider6.SetError(textBox3, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox4.Text.Trim()))
                {
                    errorProvider7.SetError(textBox4, "price is required.");
                    return;
                }
                else
                {
                    errorProvider7.SetError(textBox4, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox5.Text.Trim()))
                {
                    errorProvider8.SetError(textBox5, "qty is required.");
                    return;
                }
                else
                {
                    errorProvider8.SetError(textBox5, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox6.Text.Trim()))
                {
                    errorProvider9.SetError(textBox6, "return qty is required.");
                    return;
                }
                else
                {
                    errorProvider9.SetError(textBox6, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox7.Text.Trim()))
                {
                    errorProvider10.SetError(textBox7, "total is required.");
                    return;
                }
                else
                {
                    errorProvider10.SetError(textBox7, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox8.Text.Trim()))
                {
                    errorProvider11.SetError(textBox8, "discount is required.");
                    return;
                }
                else
                {
                    errorProvider11.SetError(textBox8, string.Empty);
                } 
                if (string.IsNullOrEmpty(textBox9.Text.Trim()))
                {
                    errorProvider12.SetError(textBox9, "total amount is required.");
                    return;
                }
                else
                {
                    errorProvider12.SetError(textBox9, string.Empty);
                }
                if (MessageBox.Show("Do You Really Want To Add This Record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.Yes)
                {
                    //qq = textBox7.Text;
                    textBox1.Enabled = true;

                    con.Open();
                    int q1 = 0, q2;
                    cm = new OleDbCommand("insert into sales_return_master(sales_ret_id,sales_id,sales_ret_date,sales_date,cust_nm,item_nm,price,qty,ret_qty,total,discount,total_amt) values('" + textBox1.Text + "','" + comboBox1.Text + "','" + dateTimePicker1.Text + "','" + dateTimePicker2.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "','" + textBox6.Text + "','" + textBox7.Text + "','" + textBox8.Text + "','" + textBox9.Text + "')", con);
                    cm.Connection = con;
                    cm.ExecuteNonQuery();

                    cm = new OleDbCommand("select qty from stock_master where item_nm='" + textBox3.Text + "'", con);
                    dr = cm.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            q1 = Convert.ToInt32(dr[0]);
                        }
                    }
                    q2 = q1 + (Convert.ToInt32(textBox6.Text));
                    cm = new OleDbCommand("update stock_master set qty='" + q2 + "' where item_nm='" + textBox3.Text + "'", con);
                    cm.ExecuteReader();

                  
                    MessageBox.Show("stock updated...");
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
        public void load_data(int c)
        {
            cm = new OleDbCommand("select * from sales_return_master", con);
            da = new OleDbDataAdapter(cm);
            da.Fill(dt);

            textBox1.Text = dt.Rows[c][0].ToString();
            comboBox1.Text = dt.Rows[c][1].ToString();
            dateTimePicker1.Text = dt.Rows[c][2].ToString();
            dateTimePicker2.Text = dt.Rows[c][3].ToString();
            textBox2.Text = dt.Rows[c][4].ToString();
            textBox3.Text = dt.Rows[c][5].ToString();
            textBox4.Text = dt.Rows[c][6].ToString();
            textBox5.Text = dt.Rows[c][7].ToString();
            textBox6.Text = dt.Rows[c][8].ToString();
            textBox7.Text = dt.Rows[c][9].ToString();
            textBox8.Text = dt.Rows[c][10].ToString();
            textBox9.Text = dt.Rows[c][11].ToString();
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

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void btn_edit_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                if (MessageBox.Show("please enter Sales Return Id", "confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.Yes)
                {
                    enable();
                }
            }
            else
            {
                int temp = 0;
                con.Open();
                OleDbCommand cm = new OleDbCommand("select sales_ret_id from sales_return_master", con);
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
                    String s1 = "update sales_return_master set sales_id='" + comboBox1.Text + "',sales_ret_date='" + dateTimePicker1.Text + "',sales_date='" + dateTimePicker2.Text + "',cust_nm='" + textBox2.Text + "',item_nm='" + textBox3.Text + "',price='" + textBox4.Text + "',qty='" + textBox5.Text + "',ret_qty='" + textBox6.Text + "',total='" + textBox7.Text + "',discount='" + textBox8.Text + "',total_amt='" + textBox9.Text + "' where sales_ret_id="+textBox1.Text+"";
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
                    MessageBox.Show("This is new id so can't updated.....plese save record", "Message", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                string s1 = "select sales_ret_id from sales_return_master";
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
                    string q = "delete from sales_return_master where sales_ret_id=" + textBox1.Text.ToString() + "";
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
                    MessageBox.Show("This is new id so can't delete.....plese save record", "Message", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            con.Open();

            string s1 = "select * from sales_master where sales_id=" + comboBox1.Text + "";
            cm = new OleDbCommand(s1, con);
            dr = cm.ExecuteReader();
            while (dr.Read())
            {
                String e1 = dr.GetString(1);
                String a = dr.GetString(3);
                String b = dr.GetString(5);
                String c = dr.GetString(7);
                String d = dr.GetString(8);
                String f = dr.GetString(10);

                dateTimePicker2.Text = e1;
                textBox2.Text = b;
                textBox3.Text = a;
                textBox4.Text = c;
                textBox5.Text = d;
                textBox8.Text = f;
                
            }

            con.Close();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox7_Leave(object sender, EventArgs e)
        {
            if (Convert.ToInt32(textBox6.Text) > Convert.ToInt32(textBox5.Text))
            {
                MessageBox.Show("Return Qty Is Maximum Than Purchase Qty ");
            }
            else
            {

                int a = Convert.ToInt32(textBox4.Text);
                int b = Convert.ToInt32(textBox6.Text);
                int amt = a * b;
                textBox7.Text = Convert.ToString(amt);
            }
        }

        private void textBox8_Leave(object sender, EventArgs e)
        {
            int a = Convert.ToInt32(textBox7.Text);
            int c = Convert.ToInt32(textBox8.Text);
            double tot = a;
            double di = (tot * c) / 100;
            double t_amt = tot - di;
            textBox9.Text = Convert.ToString(t_amt);
        }

        private void textBox9_Click(object sender, EventArgs e)
        {
            
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            
        }
    }
}
