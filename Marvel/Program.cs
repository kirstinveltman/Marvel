using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Marvel
{
    class Program
    {
        static void Main(string[] args)
        {
            Application excelApp = new Application();

            // return an error if Excel is not installed
            if (excelApp == null)
            {
                Console.WriteLine("Excel is not installed!");
                return;
            }

            // Open the Excel workbook the application is referencing
            Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\derwy\Desktop\Marvel.xlsx");
            _Worksheet excelSheet = excelBook.Sheets[1];
            Range excelRange = excelSheet.UsedRange;

            // getting amount of rows and columns
            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;


            // get specific data from the Excel spreadsheet and print them to console
            for (int i = 1; i <= rows; i++)
            {
                string firstName = Convert.ToString((excelRange.Cells[i, 2] as Range).Text);
                string lastName = Convert.ToString((excelRange.Cells[i, 3] as Range).Text);
                string color = Convert.ToString((excelRange.Cells[i, 5] as Range).Text);

                Console.WriteLine(firstName + " " + lastName + "'s favorite color is: " + color);
            }

            // exiting Excel application and releasing the COM (which allows binary communication across machine boundaries)
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            Console.ReadLine();

        }
    }
}
