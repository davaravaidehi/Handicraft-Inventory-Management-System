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
    public partial class login : Form
    {
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\@g_19\Handicraft\login.accdb");
        OleDbCommand cm;
        OleDbDataReader dr;
        public login()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            con.Open();
            cm = new OleDbCommand("select * from login where username='" + textBox1.Text + "' and password = '" + textBox2.Text + "'", con);
            dr = cm.ExecuteReader();
            if (dr.Read())
            {
                DialogResult res = MessageBox.Show("Login Successfully...","LOGIN", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (res == DialogResult.Yes)
                {
                    MDI m1 = new MDI();
                    m1.Show();
                    this.Close();
                   // con.Close();
                }
                else  
                {
                    login l1 = new login();
                    l1.Show();
                }
             }
                else
                {
                    MessageBox.Show("Invalid USERNAME or PASSWORD...", "INVALID", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                    textBox1.Text = "";
                    textBox2.Text = "";
                }
               
            
            con.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox2_MouseDown(object sender, MouseEventArgs e)
        {
            textBox2.PasswordChar = '\0';
            textBox2.UseSystemPasswordChar = false;
        }

        private void pictureBox2_MouseUp(object sender, MouseEventArgs e)
        {
            textBox2.PasswordChar = '.';
            textBox2.UseSystemPasswordChar = true;
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void login_Load(object sender, EventArgs e)
        {

        }
    }
}
