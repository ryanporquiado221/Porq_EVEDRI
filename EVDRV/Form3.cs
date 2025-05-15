using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EVDRV
{
    public partial class Form3: Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        Form4 form4;

        private void btnLogin_Click(object sender, EventArgs e)
        {
            
            Workbook book = new Workbook();
            book.LoadFromFile(path.pathfile); //Change the path to where is the excel locate.
            Worksheet sheet = book.Worksheets[0];

            int row = sheet.Rows.Length;
            bool log = false;

            for (int i = 2; i <= row; i++)
            {
                if (sheet.Range[i, 9].Value == txtUserName.Text && sheet.Range[i, 10].Value == txtPassword.Text)
                {
                    log = true;
                    Admin.Name = sheet.Range[i, 1].Value;
                    form4 = new Form4(Admin.Name);
                    form4.pictureBox1.ImageLocation = sheet.Range[i, 12].Value;
                    break;
                }
                else
                {
                    log = false;
                }
            }

            if (log == true)
            {
                MessageBox.Show("Successfully Log in", "Log in", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Logs.Log(Admin.Name, "Has been log in");
                this.Hide();
                form4.Show();
            }
            else if (log == false)
            {
                MessageBox.Show("Invalid Username or Password", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void choShowPsssword_CheckedChanged(object sender, EventArgs e)
        {
            if (choShowPsssword.Checked)
            {
                txtPassword.PasswordChar = '\0';
            }
            else
            {
                txtPassword.PasswordChar = '*';
            }
        }
    }
}
