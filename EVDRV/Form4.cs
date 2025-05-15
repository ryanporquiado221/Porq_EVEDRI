using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Xls;
using System.IO;
using System.Threading;
using System.Windows.Forms.DataVisualization.Charting;

namespace EVDRV
{
    public partial class Form4: Form
    {
        Workbook book = new Workbook();
        string name;
        FileSystemWatcher fileWatcher;
        
        public Form4(string Name)
        {
            
            InitializeComponent();
            name = Name;
            InitializeFileWatcher();
            DataSorting.datasorting();
            LoadPieChartForInActiveAndActive();
            LoadPieChartForMaleAndFemale();
            LoadBarChartForColors();
            LoadBarChartForHobbies();
            LoadBarChartForCourses();
            chart1.Titles.Add("Students");
            chart2.Titles.Add("Gender");
            chart3.Titles.Add("Colors");
            chart4.Titles.Add("Hobbies");
            chart5.Titles.Add("Courses");
            lblName.Text = name;
        }

        private void LoadPieChartForInActiveAndActive()
        {
            chart1.Series.Clear();

            Series series = new Series
            {
                Name = "Students",
                IsVisibleInLegend = true,
                ChartType = SeriesChartType.Pie,
                Font = new Font("Segoe UI", 9)
            };

            chart1.Series.Add(series);

            series.Points.AddXY($"Active\n{ShowCounts(11, "1")}", ShowCounts(11, "1"));
            series.Points.AddXY($"Inactive\n{ShowCounts(11, "0")}", ShowCounts(11, "0"));

            StylePieChart(chart1);
        }

        private void LoadPieChartForMaleAndFemale()
        {
            chart2.Series.Clear();

            Series series = new Series
            {
                Name = "Gender",
                IsVisibleInLegend = true,
                ChartType = SeriesChartType.Pie,
                Font = new Font("Segoe UI", 9)
            };

            chart2.Series.Add(series);
            series.Points.AddXY($"Male\n{ShowCounts(2, "Male")}", ShowCounts(2, "Male"));
            series.Points.AddXY($"Female\n{ShowCounts(2, "Female")}", ShowCounts(2, "Female"));

            StylePieChart(chart2);
        }

        private void LoadBarChartForColors()
        {
            chart3.Series.Clear();
            chart3.Legends.Clear();

            Series series = new Series
            {
                Name = "Colors",
                IsVisibleInLegend = false,
                ChartType = SeriesChartType.Bar
            };

            chart3.Series.Add(series);

            series.Points.AddXY("Blue", ShowCounts(4, "Blue"));
            series.Points.AddXY("Yellow", ShowCounts(4, "Yellow"));
            series.Points.AddXY("Black", ShowCounts(4, "Black"));
            series.Points.AddXY("White", ShowCounts(4, "White"));
            series.Points.AddXY("Pink", ShowCounts(4, "Pink"));
            series.Points.AddXY("Red", ShowCounts(4, "Red"));
            series.Points.AddXY("Orange", ShowCounts(4, "Orange"));
            series.Points.AddXY("Green", ShowCounts(4, "Green"));

            StyleBarChart(chart3);
        }

        private void LoadBarChartForCourses()
        {
            chart5.Series.Clear();
            chart5.Legends.Clear();

            Series series = new Series
            {
                Name = "Courses",
                IsVisibleInLegend = false,
                ChartType = SeriesChartType.Bar
            };

            chart5.Series.Add(series);
            series.Points.AddXY("BSIT", ShowCounts(6, "BSIT"));
            series.Points.AddXY("BSComEng", ShowCounts(6, "BSComEng"));
            series.Points.AddXY("BSCS", ShowCounts(6, "BSCS"));
            series.Points.AddXY("BSNursing", ShowCounts(6, "BSNursing"));

            StyleBarChart(chart5);
        }

        private void LoadBarChartForHobbies()
        {
            int basketball = 0, volleyball = 0, onlinegames = 0, others = 0;
            chart4.Series.Clear();
            book.LoadFromFile(path.pathfile);
            Worksheet sheet = book.Worksheets[0];

            for (int i = 2; i <= sheet.Rows.Length; i++)
            {
                string values = sheet.Range[i, 3].Value;
                string[] data = values.Split(' ');
                foreach (var hobby in data)
                {
                    if (hobby.Contains("Basketball")) basketball++;
                    if (hobby.Contains("Volleyball")) volleyball++;
                    if (hobby.Contains("Online-Games")) onlinegames++;
                    if (hobby.Contains("Others.")) others++;
                }
            }

            chart4.Legends.Clear();

            Series series = new Series
            {
                Name = "Hobbies",
                IsVisibleInLegend = false,
                ChartType = SeriesChartType.Bar
            };

            chart4.Series.Add(series);

            series.Points.AddXY("Basketball", basketball);
            series.Points.AddXY("Volleyball", volleyball);
            series.Points.AddXY("Online-Games", onlinegames);
            series.Points.AddXY("Others", others);

            StyleBarChart(chart4);
        }

        private void StylePieChart(System.Windows.Forms.DataVisualization.Charting.Chart chart)
        {

            chart.Series[0]["PieLabelStyle"] = "Inside";
            chart.Series[0]["PieStartAngle"] = "270";
            chart.Series[0].BorderWidth = 2;
            chart.Series[0].BorderColor = Color.Transparent;
            chart.Series[0].LabelForeColor = Color.White;

            chart.BackColor = /*Color.FromArgb(240, 240, 240);*/Color.Transparent;
            chart.ChartAreas[0].BackColor = Color.Transparent;


            // Custom pastel colors
            Color[] pastelPalette = new Color[]
            {
                Color.FromArgb(102, 194, 165),
                Color.FromArgb(252, 141, 98),
                Color.FromArgb(141, 160, 203),
                Color.FromArgb(231, 138, 195)
            };
            chart.Palette = ChartColorPalette.None;
            chart.PaletteCustomColors = pastelPalette;
        }

        private void StyleBarChart(System.Windows.Forms.DataVisualization.Charting.Chart chart)
        {
            var area = chart.ChartAreas[0];
            area.BackColor = Color.Transparent;
            area.AxisX.MajorGrid.Enabled = false;
            area.AxisY.MajorGrid.Enabled = false;
            area.AxisX.LabelStyle.Font = new Font("Segoe UI", 9);
            area.AxisY.LabelStyle.Font = new Font("Segoe UI", 9);
            area.AxisX.LabelStyle.Angle = -45;
            area.AxisX.Interval = 1;

            chart.Series[0].IsValueShownAsLabel = true;
            chart.Series[0].LabelForeColor = Color.Transparent;
            chart.Series[0].Font = new Font("Segoe UI", 9, FontStyle.Regular);
            chart.Series[0]["PointWidth"] = "0.6"; // narrower bars
            chart.Series[0]["DrawingStyle"] = "Emboss";

            // Optional pastel colors
            chart.Palette = ChartColorPalette.None;
            chart.PaletteCustomColors = new Color[]
            {
                Color.FromArgb(114, 147, 203),
                Color.FromArgb(225, 151, 76),
                Color.FromArgb(132, 186, 91),
                Color.FromArgb(211, 94, 96)
            };
        }

        private int ShowCounts(int c, string value)
        {
            book.LoadFromFile(path.pathfile); //Change the path to where is the excel locate.
            Worksheet sheet = book.Worksheets[0];
            int count = 0;
            int rows = sheet.Rows.Length;

            for (int i = 2; i <= rows; i++)
            {
                if (sheet.Range[i, c].Value == value)
                {
                    count++;
                }
            }
            return count;
        }

        private void Form4_Load(object sender, EventArgs e)
        {

        }

        private void InitializeFileWatcher()
        {
            try
            {
                string fullPath = path.pathfile;

                if (string.IsNullOrWhiteSpace(fullPath) || !File.Exists(fullPath))
                {
                    MessageBox.Show("The file path is not valid or the file does not exist: " + fullPath);
                    return;
                }

                fileWatcher = new FileSystemWatcher
                {
                    Path = Path.GetDirectoryName(fullPath),
                    Filter = Path.GetFileName(fullPath), 
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName
                };

                fileWatcher.Changed += FileSystemWatcher1_Changed;
                fileWatcher.EnableRaisingEvents = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error initializing file watcher: " + ex.Message);
            }
        }

        private void FileSystemWatcher1_Changed(object sender, FileSystemEventArgs e)
        {
            Thread.Sleep(500);

            if (this.IsHandleCreated)
            {
                this.Invoke(new Action(() =>
                {
                    try
                    {
                        fileWatcher.EnableRaisingEvents = false;

                        LoadPieChartForInActiveAndActive();
                        LoadPieChartForMaleAndFemale();
                        LoadBarChartForColors();
                        LoadBarChartForHobbies();
                        LoadBarChartForCourses();
                    }
                    finally
                    {
                        fileWatcher.EnableRaisingEvents = true;
                    }
                }));
            }

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            DateTime currentDateTime = DateTime.Now;
            dateTimePicker1.Value = currentDateTime;
            lblDate.Text = currentDateTime.ToString("MM/dd/yyyy hh:mm:ss tt");
        }

        private void btnActive_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
        }

        private void btnInactive_Click(object sender, EventArgs e)
        {
            Form5 form5 = new Form5();
            form5.Show();
        }

        private void btnLogs_Click(object sender, EventArgs e)
        {
            Form6 form6 = new Form6();
            form6.Show();
            form6.DisplayLogs();
        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            DialogResult res = MessageBox.Show("Are you sure you want to logout? ", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (res == DialogResult.Yes)
            {
                Logs.Log(name, "Has been log out");
                Form3 form3 = new Form3();
                this.Hide();
                form3.Show();
            }
        }

        private void btnAddStud_Click(object sender, EventArgs e)
        {
            Form2 form = new Form2();
            Form1 form1 = new Form1(form);
            form1.Show();
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {
            if (panel5.Visible == true)
            {
                panel5.Visible = false;
                panel3.Visible = true;
            }
            else
            {
                panel3.Visible = false;
                panel5.Visible = true;
            }
        }
    }
}
