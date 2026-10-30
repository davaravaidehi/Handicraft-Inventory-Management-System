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
    public partial class change_password : Form
    {
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\@g_19\Handicraft\login.accdb");
        public change_password()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do You Really Want To Exit ??", "Message", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                this.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            con.Open();
            OleDbCommand cm = new OleDbCommand("select * from login where username = '"+textBox1.Text+"'",con);
            OleDbDataReader dr = cm.ExecuteReader();
            if (textBox1.Text == "'|| textbox3.Text == '" || textBox4.Text == "")
            {
                MessageBox.Show("Please enter password..");
            }
            else
            {
                if (dr.Read())
                {
                    if (textBox3.Text == textBox4.Text)
                    {
                        //con.Open();
                        OleDbCommand cm1 = new OleDbCommand("update login set [password] = '" + textBox3.Text + "' where [username] ='" + textBox1.Text + "'", con);
                        cm1.ExecuteNonQuery();
                        con.Close();
                       
                        if (MessageBox.Show("Password  Changed Successfully...", "Message", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                        {
                            this.Close();
                        }
                       
                    }
                    else
                    {
                        MessageBox.Show("Password Not Same");
                    }
                }
                else
                {
                    MessageBox.Show("Username is not correct..");
                }
            }
        }

        private void change_password_Load(object sender, EventArgs e)
        {

        }
    }
}
