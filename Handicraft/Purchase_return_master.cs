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
    public partial class Purchase_return_master : Form
    {
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\@g_19\Handicraft\login.accdb");
        OleDbCommand cm;
        OleDbDataReader dr;
        OleDbDataAdapter da;
        DataSet ds;
        DataTable dt = new DataTable();
        int index;
       // string qt1, tot;
        public Purchase_return_master()
        {
            InitializeComponent();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

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
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
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
            textBox6.Enabled = true;
            textBox7.Enabled = true;
            textBox8.Enabled = true;

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
            textBox6.Enabled = false;
            textBox7.Enabled = false;
            textBox8.Enabled = false;

        }

        public void display()
        {
             con.Open();
            String q = "select * from purchase_return_master";
            cm = new OleDbCommand(q, con);
            var reader = cm.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(reader);
            dataGridView1.DataSource = dt;
            con.Close();
        }
        
        private void id_com_SelectedIndexChanged(object sender, EventArgs e)
        {
            con.Open();
            String s = "select * from purchase_return_master where pur_ret_id = " + id_com.Text + "";
            cm = new OleDbCommand(s, con);
            var reader = cm.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(reader);
            dataGridView1.DataSource = dt;
            con.Close();
        }

        private void nm_cmb_SelectedIndexChanged(object sender, EventArgs e)
        {
            con.Open();
            String s = "select * from purchase_return_master where item_nm = '" + nm_cmb.Text + "'";
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

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
               
                con.Open();
                String s1 = "select item_nm from purchase_return_master";
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

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                con.Open();
                String s1 = "select pur_ret_id from purchase_return_master";
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

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }


        public void auto_inc()
        {
            da = new OleDbDataAdapter("select max(pur_ret_id) from purchase_return_master", con);
            ds = new DataSet();
            da.Fill(ds);

            int cno;

            if (ds.Tables[0].Rows[0][0].ToString().Equals(""))
            {
                cno = 1101;
            }
            else
            {
                cno = Int32.Parse(ds.Tables[0].Rows[0][0].ToString());
                cno = cno + 1;
            }
            textBox1.Text = cno.ToString();
        }

        private void btn_exit_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do You Really Want To Exit ??", "Message", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                this.Close();
            }
        }

        private void btn_add_Click(object sender, EventArgs e)
        {
            btn_save.Enabled = true;
            btn_cancle.Enabled = true;
            enable();
            auto_inc();
            
            
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
                    errorProvider1.SetError(textBox1, "purchase return ID is required.");
                    return;
                }
                else
                {
                    errorProvider1.SetError(textBox1, string.Empty);
                }
                if (string.IsNullOrEmpty(comboBox1.Text.Trim()))
                {
                    errorProvider2.SetError(comboBox1, "purchase Id is required.");
                    return;
                }
                else
                {
                    errorProvider2.SetError(comboBox1, string.Empty);
                }
                if (string.IsNullOrEmpty(dateTimePicker1.Text.Trim()))
                {
                    errorProvider3.SetError(dateTimePicker1, "purchase return date is required.");
                    return;
                }
                else
                {
                    errorProvider3.SetError(dateTimePicker1, string.Empty);
                }
                if (string.IsNullOrEmpty(dateTimePicker2.Text.Trim()))
                {
                    errorProvider4.SetError(dateTimePicker2, "Purchase date is required.");
                    return;
                }
                else
                {
                    errorProvider4.SetError(dateTimePicker2, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox2.Text.Trim()))
                {
                    errorProvider5.SetError(textBox2, "Item name is required.");
                    return;
                }
                else
                {
                    errorProvider5.SetError(textBox2, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox3.Text.Trim()))
                {
                    errorProvider6.SetError(textBox3, "price is required.");
                    return;
                }
                else
                {
                    errorProvider6.SetError(textBox3, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox4.Text.Trim()))
                {
                    errorProvider7.SetError(textBox4, "qty is required.");
                    return;
                }
                else
                {
                    errorProvider7.SetError(textBox4, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox5.Text.Trim()))
                {
                    errorProvider8.SetError(textBox5, "return qty is required.");
                    return;
                }
                else
                {
                    errorProvider8.SetError(textBox5, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox6.Text.Trim()))
                {
                    errorProvider9.SetError(textBox6, "total is required.");
                    return;
                }
                else
                {
                    errorProvider9.SetError(textBox6, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox7.Text.Trim()))
                {
                    errorProvider10.SetError(textBox7, "discount is required.");
                    return;
                }
                else
                {
                    errorProvider10.SetError(textBox7, string.Empty);
                }
                if (string.IsNullOrEmpty(textBox8.Text.Trim()))
                {
                    errorProvider11.SetError(textBox8, "total amount is required.");
                    return;
                }
                else
                {
                    errorProvider11.SetError(textBox8, string.Empty);
                }
                if (MessageBox.Show("Do You Really Want To Add This Record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.Yes)
                {
                    textBox1.Enabled = true;
                    con.Open();
                    int q1 = 0, q2;
                    cm = new OleDbCommand("insert into purchase_return_master(pur_ret_id,pur_id,pur_ret_date,pur_date,item_nm,price,qty,ret_qty,total,discount,total_amt) values('" + textBox1.Text + "','" + comboBox1.Text + "','" + dateTimePicker1.Text + "','" + dateTimePicker2.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "','" + textBox6.Text + "','" + textBox7.Text + "','" + textBox8.Text + "')", con);
                    cm.Connection = con;
                    cm.ExecuteNonQuery();
                    /*cm = new OleDbCommand("select qty from stock_master where item_id=" + comboBox1.SelectedItem, con);
                    dr = cm.ExecuteReader();
                    qt1 = textBox5.Text;
                    int q1 = Convert.ToInt32(textBox4.Text);
                    tot = Convert.ToString(Convert.ToInt32(qt1) - Convert.ToInt32(q1));
                    cm = new OleDbCommand("update stock_master set qty='" + tot + "' where item_id=" + comboBox1.SelectedItem, con);
                    cm.ExecuteNonQuery();
                    MessageBox.Show("stock updated...");*/
                    cm = new OleDbCommand("select qty from stock_master where item_nm='" + textBox2.Text + "'", con);
                    dr = cm.ExecuteReader();
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            q1 = Convert.ToInt32(dr[0]);
                        }
                    }


                    q2 = q1 - (Convert.ToInt32(textBox5.Text));




                    cm = new OleDbCommand("update stock_master set qty='" + q2 + "' where item_nm='" + textBox2.Text + "'", con);
                    cm.ExecuteReader();
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
                OleDbCommand cm = new OleDbCommand("select pur_ret_id from purchase_return_master", con);
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
                    String s1 = "update purchase_return_master set pur_id='" + comboBox1.Text + "',pur_ret_date='" + dateTimePicker1.Text + "',pur_date='" + dateTimePicker2.Text + "',item_nm='" + textBox2.Text + "',price='" + textBox3.Text + "',qty='" + textBox4.Text + "',ret_qty='" + textBox5.Text + "',total='" + textBox6.Text + "',discount='" + textBox7.Text + "',total_amt='" + textBox8.Text + "' where pur_ret_id=" + textBox1.Text + "";
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
                string s1 = "select pur_ret_id from purchase_return_master";
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
                    string q = "delete from purchase_return_master where pur_ret_id=" + textBox1.Text.ToString() + "";
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

       

        private void btn_cancle_Click(object sender, EventArgs e)
        {
            clear();
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

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        public void load_data(int c)
        {
            cm = new OleDbCommand("select * from purchase_return_master", con);
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

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void last_Click(object sender, EventArgs e)
        {
            index = dt.Rows.Count - 1;
            load_data(index);
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void Purchase_return_master_Load(object sender, EventArgs e)
        {
            btn_add.Enabled = true;
            btn_save.Enabled = false;
            btn_edit.Enabled = false;
            btn_delete.Enabled = false;
            btn_cancle.Enabled = false;
            btn_exit.Enabled = true;
            label2.Text = DateTime.Now.ToShortDateString();
            label14.Text = DateTime.Now.ToShortTimeString();
            
            con.Open();
            String s1 = "select pur_id from purchase_master";
            cm = new OleDbCommand(s1, con);
            OleDbDataReader dr = cm.ExecuteReader();
            while (dr.Read())
            {
                comboBox1.Items.Add(dr[0].ToString());

            }
            con.Close();
            display();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            con.Open();
            string s1 = "select * from purchase_master where pur_id=" + comboBox1.Text + "";
            cm = new OleDbCommand(s1, con);
            dr = cm.ExecuteReader();
            while (dr.Read())
            {
                String a = dr.GetString(1);
                String b = dr.GetString(3);
                String c = dr.GetString(4);
                String d = dr.GetString(5);
                String h = dr.GetString(7);
                dateTimePicker2.Text = a;
                textBox2.Text = b;
                textBox3.Text = c;
                textBox4.Text = d;
                textBox7.Text = h;

            }
            con.Close();
            con.Open();
            string s2 = "select * from stock_master where item_nm='" + textBox2.Text + "'";
            cm = new OleDbCommand(s2, con);
            dr = cm.ExecuteReader();
            while (dr.Read())
            {
                textBox4.Text = dr.GetValue(2).ToString();
            }
            con.Close();
        }

        private void textBox5_Leave(object sender, EventArgs e)
        {
            if (Convert.ToInt32(textBox5.Text) > Convert.ToInt32(textBox4.Text))
            {
                MessageBox.Show("Return Qty Is Maximum Than Purchase Qty ");
            }
            else
            {

                int a = Convert.ToInt32(textBox3.Text);
                int b = Convert.ToInt32(textBox5.Text);
                int amt = a * b;
                textBox6.Text = Convert.ToString(amt);
            }
        }

        private void textBox7_Leave(object sender, EventArgs e)
        {
            int a = Convert.ToInt32(textBox6.Text);
            int c = Convert.ToInt32(textBox7.Text);
            double tot = a;
            double di = (tot * c) / 100;
            double t_amt = tot - di;
            textBox8.Text = Convert.ToString(t_amt);
        }
    }
}
