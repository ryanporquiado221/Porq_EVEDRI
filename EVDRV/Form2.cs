using Spire.Xls;
using Spire.Xls.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace EVDRV
{
    public partial class Form2 : Form
    {
        Form5 form5;
        Form4 form4;

        public Form2()
        {
            InitializeComponent();
            form5 = new Form5();
            form4 = new Form4(Admin.Name);
            lblName.Text = Admin.Name;
        }

        public void LoadActiveData()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(path.pathfile); //Change the path to where is the excel locate.
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = new DataTable();

            int columnCount = sheet.Columns.Length;
            for (int col = 1; col <= columnCount; col++)
            {
                dt.Columns.Add(sheet.Range[1, col].Value);
            }
            
            for (int row = 2; row <= sheet.LastRow; row++)
            {
                string activedata = sheet.Range[row, 11].Value;

                if (activedata == "1")
                {
                    DataRow dr = dt.NewRow();

                    for (int col = 1; col <= columnCount; col++)
                    {
                        dr[col - 1] = sheet.Range[row, col].Value;
                    }

                    dt.Rows.Add(dr);
                }
            }

            dataGridView1.DataSource = dt;
        }

        public void UpdateDataToExcel(int ID, string name, string gender, string hobbies, string favcolor, string saying, string course, string username, string password, string status, string email, string profilepath, string age)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(path.pathfile); //Change the path to where is the excel locate.
            Worksheet sheet = book.Worksheets[0];

            int index = ID + 2;
            sheet.Range[index, 1].Value = name;
            sheet.Range[index, 2].Value = gender;
            sheet.Range[index, 3].Value = hobbies;
            sheet.Range[index, 4].Value = favcolor;
            sheet.Range[index, 5].Value = saying;
            sheet.Range[index, 6].Value = course;
            sheet.Range[index, 7].Value = email;
            sheet.Range[index, 8].Value = age;
            sheet.Range[index, 9].Value = username;
            sheet.Range[index, 10].Value = password;
            sheet.Range[index, 11].Value = status;
            sheet.Range[index, 12].Value = profilepath;

            book.SaveToFile(path.pathfile);
            Logs.Log(Admin.Name, $"Updating a data");
        }

        //public void GetDataFromFr1(string name, string gender, string hobbies, string favcolor, string saying, string username, string password)
        //{
        //    int index = dataGridView1.Rows.Add();
        //    dataGridView1.Rows[index].Cells[0].Value = name;
        //    dataGridView1.Rows[index].Cells[1].Value = gender;
        //    dataGridView1.Rows[index].Cells[2].Value = hobbies;
        //    dataGridView1.Rows[index].Cells[3].Value = favcolor;
        //    dataGridView1.Rows[index].Cells[4].Value = saying;
        //    dataGridView1.Rows[index].Cells[5].Value = username;
        //    dataGridView1.Rows[index].Cells[6].Value = password;
        //}   

        public void GetUpdatedDataFromFr1(int ID, string name, string gender, string hobbies, string favcolor, string saying, string course, string username, string password, string status, string email, string profilepath, string age)
        {
            int index = ID;
            dataGridView1.Rows[index].Cells[0].Value = name;
            dataGridView1.Rows[index].Cells[1].Value = gender;
            dataGridView1.Rows[index].Cells[2].Value = hobbies;
            dataGridView1.Rows[index].Cells[3].Value = favcolor;
            dataGridView1.Rows[index].Cells[4].Value = saying;
            dataGridView1.Rows[index].Cells[5].Value = course;
            dataGridView1.Rows[index].Cells[6].Value = email;
            dataGridView1.Rows[index].Cells[6].Value = age;
            dataGridView1.Rows[index].Cells[7].Value = username;
            dataGridView1.Rows[index].Cells[8].Value = password;
            dataGridView1.Rows[index].Cells[9].Value = status;
            dataGridView1.Rows[index].Cells[10].Value = profilepath;
        }


        public void GetInactiveData()
        {
            try
            {
                int i = dataGridView1.SelectedRows[0].Index;
                bool hasFound = false;
                string name = "";

                Workbook book = new Workbook();
                book.LoadFromFile(path.pathfile); //Change the path to where is the excel locate.
                Worksheet sheet = book.Worksheets[0];
                int row = sheet.Rows.Length;

                for (int r = 2; r <= row; r++)
                {
                    if (r == i + 2)
                    {
                        sheet.Range[r, 11].Value = "0";
                        hasFound = true;
                        name = sheet.Range[r, 1].Value;
                    }
                }

                book.SaveToFile(path.pathfile);

                if (hasFound == true)
                {
                    Logs.Log(Admin.Name, $"Remove {name} in active list");
                }

                LoadActiveData();
                form5.LoadInactiveData();
                DataSorting.datasorting();
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("No data can be deleted.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            DialogResult res = MessageBox.Show("Are you sure you want to Inactive this Student?", "Confimation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (res == DialogResult.Yes)
            {
                Form6 form6 = new Form6();
                GetInactiveData();
                form6.DisplayLogs();
            }
            else
            {
                //
            }
        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Form1 form1 = new Form1(this);

            int r = dataGridView1.CurrentCell.RowIndex;
            form1.lblID.Text = r.ToString();

            form1.txtName.Text = dataGridView1.Rows[r].Cells[0].Value.ToString();
            form1.txtSaying.Text = dataGridView1.Rows[r].Cells[4].Value.ToString();
            form1.cmbFavcolor.Text = dataGridView1.Rows[r].Cells[3].Value.ToString();
            form1.txtUserName.Text = dataGridView1.Rows[r].Cells[8].Value.ToString();
            form1.txtPassword.Text = dataGridView1.Rows[r].Cells[9].Value.ToString();
            form1.txtCourse.Text = dataGridView1.Rows[r].Cells[5].Value.ToString();
            form1.txtStatus.Text = dataGridView1.Rows[r].Cells[10].Value.ToString();
            form1.txtEmail.Text = dataGridView1.Rows[r].Cells[6].Value.ToString();
            form1.pictureBox1.ImageLocation = dataGridView1.Rows[r].Cells[11].Value.ToString();

            Workbook book = new Workbook();
            book.LoadFromFile(path.pathfile); //Change the path to where is the excel locate.
            Worksheet sheet = book.Worksheets[2];

            int rows = sheet.Rows.Length;

            DateTime parsedDate;

            for (int i = 2; i <= rows; i++)
            {
                if (dataGridView1.Rows[r].Cells[8].Value.ToString() == sheet.Range[i, 1].Value)
                {
                    if (DateTime.TryParse(sheet.Range[i, 2].Value, out parsedDate))
                    {
                        form1.dateTimePicker1.Value = parsedDate;
                    }
                }
            }


            switch (dataGridView1.Rows[r].Cells[1].Value)
            {
                case "Male":
                    form1.radMale.Checked = true;
                    break;
                default:
                    form1.radFemale.Checked = true;
                    break;
            }
            string cellValue = dataGridView1.Rows[r].Cells[2].Value?.ToString();
            if (!string.IsNullOrEmpty(cellValue))
            {
                string[] words = cellValue.Split(' '); // Split by space

                foreach (string word in words)
                {
                    if (word == "Basketball")
                    {
                        form1.chkBasketball.Checked = true;
                    }
                    if (word == "Volleyball")
                    {
                        form1.chkVolleyball.Checked = true;
                    }
                    if (word == "Online-Games")
                    {
                        form1.chkOG.Checked = true;
                    }
                    if (word == "Others.")
                    {
                        form1.chkOthers.Checked = true;
                    }
                }
            }

            this.Hide();
            form1.Show();
            form1.btnUpdate.Visible = true;
            form1.txtStatus.Enabled = false;
            form1.dateTimePicker1.Enabled = false;
            form1.btnDisplay.Visible = false;
            form1.btnAdd.Visible = false;
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string searchText = txtSearch.Text.Trim().ToLower();

            dataGridView1.ClearSelection();
            dataGridView1.CurrentCell = null;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    string firstName = row.Cells[0].Value?.ToString().ToLower();

                    row.Visible = string.IsNullOrEmpty(searchText) || (firstName != null && firstName.Contains(searchText));
                }
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            LoadActiveData();
            dataGridView1.ClearSelection();
        }

        private void btnAddData_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form1 form1 = new Form1(this);
            form1.Show();
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form4 form4 = new Form4(Admin.Name);
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            if (panel1.Visible == true)
            {
                panel1.Visible = false;
            }
            else if (panel1.Visible == false)
            {
                panel1.Visible = true;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            DateTime currentDateTime = DateTime.Now;
            dateTimePicker1.Value = currentDateTime;
            lblDate.Text = currentDateTime.ToString("MM/dd/yyyy hh:mm:ss tt");
        }
    }
}
