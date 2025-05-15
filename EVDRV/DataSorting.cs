using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EVDRV
{
    public class DataSorting
    {
        public static void datasorting()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(path.pathfile); //Change the path to where is the excel locate.
            Worksheet sheet = book.Worksheets[0];

            int rowCount = sheet.LastRow;
            int colCount = sheet.LastColumn;

            List<string[]> data = new List<string[]>();

            for (int row = 2; row <= rowCount; row++)  
            {
                string[] rowData = new string[colCount];
                for (int col = 1; col <= colCount; col++)
                {
                    rowData[col - 1] = sheet.Range[row, col].Value;
                }
                data.Add(rowData);
            }

            var sortedData = data.OrderByDescending(r => r[10]).ToList();

            for (int row = 2; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    sheet.Range[row, col].Value = sortedData[row - 2][col - 1];
                }
            }

            book.SaveToFile(path.pathfile);
        }
    }
}
