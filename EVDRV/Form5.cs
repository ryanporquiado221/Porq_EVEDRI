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
    public partial class Form5: Form
    {
        Form4 form4;
        

        public Form5()
        {
            InitializeComponent();
            LoadInactiveData();
            dataGridView1.ClearSelection(); 
            form4 = new Form4(Admin.Name);
            lblName.Text = Admin.Name;
        }

        public void LoadInactiveData()
        {
            Workbook book = new Workbook();
            /*book.LoadFromFile(path.pathfile);*/ //Change the path to where is the excel locate.
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = new DataTable();

            int columnCount = sheet.Columns.Length;
            for (int col = 1; col <= columnCount; col++)
            {
                // Get column header from first row
                dt.Columns.Add(sheet.Range[1, col].Value);
            }

            // Loop through rows starting from row 2 (assuming row 1 is header)
            for (int row = 2; row <= sheet.LastRow; row++)
            {
                string activedata = sheet.Range[row, 11].Value;

                if (activedata == "0")
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

        public void GetActiveData()
        {
            //try
            //{
                DialogResult res = MessageBox.Show("Are you sure you want to active this user?", "Convfirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    int i = dataGridView1.SelectedRows[0].Index;
                    Form2 form2 = new Form2();
                    bool hasFound = false;
                    string name = "";
                    int active = 0;
                    int count = 0;

                    Workbook book = new Workbook();
                    book.LoadFromFile(path.pathfile); //Change the path to where is the excel locate.
                    Worksheet sheet = book.Worksheets[0];
                    int row = sheet.Rows.Length;

                    for (int r = 2; r <= row; r++)
                    {
                        if (sheet.Range[r, 11].Value == "1")
                        {
                            active++;
                        }
                    }

                    for (int rw = 2 + active; rw <= row; rw++)
                    {
                        if (count == i)
                        {
                            sheet.Range[rw, 11].Value = "1";
                            hasFound = true;
                            name = sheet.Range[rw, 1].Value;
                        }
                        count++;
                    }

                    book.SaveToFile(path.pathfile);
                    if (hasFound == true)
                    {
                        Logs.Log(Admin.Name, $"Remove {name} in inactive list");
                    }

                    LoadInactiveData();
                    form2.LoadActiveData();
                    DataSorting.datasorting();
                }
            //}
            //catch(Exception ex)
            //{
            //    MessageBox.Show("No data can be Activate.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
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

        private void btnDelete_Click(object sender, EventArgs e)
        {
            GetActiveData();
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(path.pathfile); //Change the path to where is the excel locate.
            Worksheet sheet = book.Worksheets[0];
            book.SaveToFile(path.pathfile);
            this.Hide();
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
