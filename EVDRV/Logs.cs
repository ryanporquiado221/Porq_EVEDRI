using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Spire.Xls;

namespace EVDRV
{
    class Logs
    {
        public static void Log(string name, string message)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(path.pathfile); //Change the path to where is the excel locate.
            Worksheet sheet = book.Worksheets[1];

            string date = DateTime.Now.ToString("MMMM/d/yyyy");
            string time = DateTime.Now.ToString("hh:mm:ss tt");

            int r = sheet.Rows.Length + 1;

            sheet.Range[r, 1].Value = name;
            sheet.Range[r, 2].Value = message;
            sheet.Range[r, 3].Value = date;
            sheet.Range[r, 4].Value = time;

            book.SaveToFile(path.pathfile);
        }
    }
}
