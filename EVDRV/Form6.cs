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
    public partial class Form6 : Form
    {
        Workbook book = new Workbook();

        public Form6()
        {
            InitializeComponent();
            lblName.Text = Admin.Name;
        }

        public void DisplayLogs()
        {
            book.LoadFromFile(path.pathfile); //Change the path to where is the excel locate.
            Worksheet sheet = book.Worksheets[1];
            DataTable dt = new DataTable();
            dt = sheet.ExportDataTable();

            dataGridView1.DataSource = dt;
        }

        private void Form6_Load(object sender, EventArgs e)
        {
            DisplayLogs();
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form4 form4 = new Form4(Admin.Name);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            DateTime currentDateTime = DateTime.Now;
            dateTimePicker1.Value = currentDateTime;
            lblDate.Text = currentDateTime.ToString("MM/dd/yyyy hh:mm:ss tt");
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
    }
}
