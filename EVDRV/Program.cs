using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EVDRV
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Form2 form2 = new Form2();

            Form1 form1 = new Form1(form2);

            Form3 form3 = new Form3();

            Form4 form4 = new Form4(Admin.Name);

            Form6 form6 = new Form6();

            Form5 form5 = new Form5();

            Application.Run(form3);
        }
    }
}
