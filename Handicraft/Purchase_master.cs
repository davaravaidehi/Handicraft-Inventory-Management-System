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
    public partial class Purchase_master : Form
    {
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\@g_19\Handicraft\login.accdb");
        OleDbCommand cm;
        OleDbDataReader dr;
        OleDbDataAdapter da;
        DataSet ds;
        int q1 = 0, q2;
        DataTable dt = new DataTable();
        int index;
        public Purchase_master()
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
            dateTimePicker1.Text = "";
            comboBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";

        }
        public void disable()
        {
            textBox1.Enabled = false;
            dateTimePicker1.Enabled = false;
            comboBox1.Enabled = false;
            textBox2.Enabled = false;
            textBox3.Enabled = false;
            textBox4.Enabled = false;
            textBox5.Enabled = false;
            textBox6.Enabled = false;
            textBox7.Enabled = false;

        }
        public void enable()
        {
            textBox1.Enabled = true;
            dateTimePicker1.Enabled = true;
            comboBox1.Enabled = true;
            textBox2.Enabled = true;
            textBox3.Enabled = true;
            textBox4.Enabled = true;
            textBox5.Enabled = true;
            textBox6.Enabled = true;
            textBox7.Enabled = true;
        }

        public void auto_inc()
        {
            da = new OleDbDataAdapter("select max(pur_id) from purchase_master", con);
            ds = new DataSet();
            da.Fill(ds);

            int cno;

            if (ds.Tables[0].Rows[0][0].ToString().Equals(""))
            {
                cno = 1001;
            }
            else
            {
                cno = Int32.Parse(ds.Tables[0].Rows[0][0].ToString());
                cno = cno + 1;
            }
            textBox1.Text = cno.ToString();
        }

        public void display()
        {
             con.Open();
            String q = "select * from purchase_master";
            cm = new OleDbCommand(q, con);
            var reader = cm.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(reader);
            dataGridView1.DataSource = dt;
            con.Close();
        }

        public void amtt()
        {
            double a = Convert.ToInt32(textBox3.Text);
            double b = Convert.ToInt32(textBox4.Text);
            double total = a * b;
            textBox5.Text = Convert.ToString(total);
        }
       
        public void totamt()
        {
            double a = Convert.ToInt32(textBox5.Text);
           
            double c = Convert.ToInt32(textBox6.Text);
           
            double dis = (a * c) / 100;
            double total_amt = a - dis;
            textBox7.Text = Convert.ToString(total_amt);
        }

        private void btn_cancle_Click(object sender, EventArgs e)
        {
            clear();
        }

        private void btn_save_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("please Enter pur_ID", "Message", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            else
            {
                if (string.IsNullOrEmpty(textBox1.Text.Trim()))
                {
                    errorProvider1.SetError(textBox1, "purchase ID is required.");
                    return;
                }
                else
                {
                    errorProvider1.SetError(textBox1, string.Empty);
                }
                if (string.IsNullOrEmpty(dateTimePicker1.Text.Trim()))
                {
                    errorProvider2.SetError(dateTimePicker1, "purchase Date is required.");
                    return;
                }
                else
                {
                    errorProvider2.SetError(dateTimePicker1, string.Empty);
                }
                if (string.IsNullOrEmpty(comboBox1.Text.Trim()))
                {
                    errorProvider3.SetError(comboBox1, "item ID is required.");
                    return;
                }
                else
                {
                    errorProvider3.SetError(comboBox1, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox2.Text.Trim()))
                {
                    errorProvider4.SetError(textBox2, "item name is required.");
                    return;
                }
                else
                {
                    errorProvider4.SetError(textBox2, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox3.Text.Trim()))
                {
                    errorProvider5.SetError(textBox3, "price is required.");
                    return;
                }
                else
                {
                    errorProvider5.SetError(textBox3, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox4.Text.Trim()))
                {
                    errorProvider6.SetError(textBox4, "qty is required.");
                    return;
                }
                else
                {
                    errorProvider6.SetError(textBox4, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox5.Text.Trim()))
                {
                    errorProvider7.SetError(textBox5, "total is required.");
                    return;
                }
                else
                {
                    errorProvider7.SetError(textBox5, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox6.Text.Trim()))
                {
                    errorProvider7.SetError(textBox6, "discount is required.");
                    return;
                }
                else
                {
                    errorProvider7.SetError(textBox6, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox7.Text.Trim()))
                {
                    errorProvider7.SetError(textBox7, "total amount is required.");
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
                    cm = new OleDbCommand("insert into purchase_master values('" + textBox1.Text + "','" + dateTimePicker1.Text +"','" +comboBox1.SelectedItem+"','"+ textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "','" + textBox6.Text + "','" + textBox7.Text + "')", con);
                    cm.Connection = con;
                    cm.ExecuteNonQuery();

                  
                    int temp = 0;

                    OleDbCommand cm1 = new OleDbCommand("select item_id from stock_master", con);
                     OleDbDataReader dr1 = cm1.ExecuteReader();

                    while (dr1.Read())
                    {
                        if (comboBox1.Text == dr1.GetInt32(0).ToString())
                        {
                            temp = 1;
                            break;
                        }
                    }

                    if (temp == 1)
                    {
                        if (comboBox1.Text != null)
                        {

                            cm1 = new OleDbCommand("select qty from stock_master where item_id=" + comboBox1.Text , con);
                             dr1 = cm1.ExecuteReader();
                            if (dr1.HasRows)
                            {
                                while (dr1.Read())
                                {
                                    q1 = Convert.ToInt32(dr1[0]);
                                }
                            }
                        }
                      
                        q2 = q1 + (Convert.ToInt32(textBox4.Text));
                        MessageBox.Show("updated qty is : " + q2);

                        cm1 = new OleDbCommand("update stock_master set qty='" + q2 + "' where item_id=" + comboBox1.Text , con);
                        cm1.ExecuteNonQuery();
                        MessageBox.Show("stock is updated..");
                        //  con.Close();
                       // MessageBox.Show("insert successfully");
                    }
                   else
                    {

                        cm1 = new OleDbCommand("insert into stock_master values('" + comboBox1.SelectedItem + "','" + textBox2.Text + "','" + textBox4.Text + "')", con);
                        cm1.Connection = con;
                        cm1.ExecuteNonQuery();

                       
                    }
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
       
        private void btn_add_Click(object sender, EventArgs e)
        {
            btn_save.Enabled = true;
            btn_cancle.Enabled = true;
            enable();
            auto_inc();
        
        }

        private void Purchase_master_Load(object sender, EventArgs e)
        {
            btn_add.Enabled = true;
            btn_save.Enabled = false;
            btn_edit.Enabled = false;
            btn_delete.Enabled = false;
            btn_cancle.Enabled = false;
            btn_exit.Enabled = true;
            label23.Text = DateTime.Now.ToShortDateString();
            label22.Text = DateTime.Now.ToShortTimeString();
            con.Open();
            String s1 = "select item_id,price from item_master";
            cm = new OleDbCommand(s1, con);
            OleDbDataReader dr = cm.ExecuteReader();
            while (dr.Read())
            {
                comboBox1.Items.Add(dr[0].ToString());
                
            }
            con.Close();
            display();
        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_Leave(object sender, EventArgs e)
        {
            amtt();
        }

        private void textBox6_Leave(object sender, EventArgs e)
        {
            totamt();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            enable();
            btn_delete.Enabled = true;
            btn_edit.Enabled = true;
            btn_cancle.Enabled = true; 
            textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            dateTimePicker1.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            comboBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
            textBox4.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            textBox5.Text = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();
            textBox6.Text = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString();
            textBox7.Text = dataGridView1.Rows[e.RowIndex].Cells[8].Value.ToString();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {

                con.Open();
                String s1 = "select pur_id from purchase_master";
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
            String s = "select * from purchase_master where pur_id = " + id_com.Text + "";
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
                String s1 = "select item_nm from purchase_master";
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
            String s = "select * from purchase_master where item_nm = '" + nm_cmb.Text + "'";
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
         
            con.Open();

            string s1 = "select * from item_master where item_id=" + comboBox1.Text + "";
            cm = new OleDbCommand(s1, con);
            dr = cm.ExecuteReader();
            while (dr.Read())
            {
                String a = dr.GetString(1);
                String b = dr.GetString(3);
                textBox2.Text = a;
                textBox3.Text = b;

            }

            con.Close();
        }
        public void load_data(int c)
        {
            cm = new OleDbCommand("select * from purchase_master", con);
            da = new OleDbDataAdapter(cm);
            da.Fill(dt);

            textBox1.Text = dt.Rows[c][0].ToString();
            dateTimePicker1.Text = dt.Rows[c][1].ToString();
            comboBox1.Text = dt.Rows[c][2].ToString();
            textBox2.Text = dt.Rows[c][3].ToString();
            textBox3.Text = dt.Rows[c][4].ToString();
            textBox4.Text = dt.Rows[c][5].ToString();
            textBox5.Text = dt.Rows[c][6].ToString();
            textBox6.Text = dt.Rows[c][7].ToString();
            textBox7.Text = dt.Rows[c][8].ToString();
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

        private void btn_edit_Click(object sender, EventArgs e)
        {
            
            if (textBox1.Text == "")
            {
                if (MessageBox.Show("please enterPurchase Return Id", "confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.Yes)
                {
                    enable();
                }
            }
            else
            {
                int temp = 0;
                con.Open();
                OleDbCommand cm = new OleDbCommand("select pur_id from purchase_master", con);
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
                    String s1 = "update purchase_master set pur_date='"+dateTimePicker1.Text+"',item_id='"+comboBox1.Text+"',item_nm='"+textBox2.Text+"',price='"+textBox3.Text+"',qty='"+textBox4.Text+"',total='"+textBox5.Text+"',discout='"+textBox6.Text+"',total_amt='"+textBox7.Text+"' where pur_id="+textBox1.Text+"";
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
                string s1 = "select pur_id from purchase_master";
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
                    string q = "delete from purchase_master where pur_id=" + textBox1.Text.ToString() + "";
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

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }
    }
}
