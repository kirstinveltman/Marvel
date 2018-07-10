using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

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
                string firstName = Convert.ToString((excelRange.Cells[i, 1] as Range).Text);
                string lastName = Convert.ToString((excelRange.Cells[i, 2] as Range).Text);
                int age = Convert.ToInt32((excelRange.Cells[i, 3] as Range).Text);
                string color = Convert.ToString((excelRange.Cells[i, 4] as Range).Text);

                // connection string and insert string query to db
                var connectionString = @"Data Source=RAVN\SQLEXPRESS;Initial Catalog=characters;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";

                string queryString = @"INSERT INTO charactersMarvel (FirstName, LastName, Age, FavoriteColor)
                                    VALUES (@FirstName, @LastName, @Age, @FavoriteColor)";

                // set up and execute data from each line above and insert into db
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlCommand command = new SqlCommand(queryString, connection);
                    command.Parameters.Add("@FirstName", SqlDbType.VarChar);
                    command.Parameters.Add("@LastName", SqlDbType.VarChar);
                    command.Parameters.Add("@Age", SqlDbType.Int);
                    command.Parameters.Add("@FavoriteColor", SqlDbType.VarChar);

                    command.Parameters["@FirstName"].Value = firstName;
                    command.Parameters["@LastName"].Value = lastName;
                    command.Parameters["@Age"].Value = age;
                    command.Parameters["@FavoriteColor"].Value = color;

                    connection.Open();
                    command.ExecuteNonQuery();
                    connection.Close();
                }
            }

            // exiting Excel application and releasing the COM (which allows binary communication across machine boundaries)
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            Console.WriteLine("Excel File Transfer Successful");
            Console.ReadLine();

        }

    }
}
