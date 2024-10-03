using System;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;

string connectionString = "Data Source=NANDKISHOR;Initial Catalog=PPMS_Solution;User ID=sa; Password=root;";
string tableName = "[dbo].[Config_Station]";
string excelFilePath = @"C:\temp\ExecelDownloadConfigStation.xlsx";

using (SqlConnection connection = new SqlConnection(connectionString))
{
    connection.Open();

    // Create a SQL command to select data from the table
    using (SqlCommand command = new SqlCommand($"SELECT * FROM {tableName}", connection))
    {
        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
        {
            DataTable dataTable = new DataTable();
            adapter.Fill(dataTable);

            // Set the license context to suppress the license exception
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            // Create a new Excel package and add a worksheet
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Clear the existing data in the worksheet
                worksheet.Cells.Clear();

                // Populate Excel worksheet with data from DataTable
                worksheet.Cells.LoadFromDataTable(dataTable, true);

                // Save the Excel package to a file
                package.SaveAs(new System.IO.FileInfo(excelFilePath));

                Console.WriteLine("Config Activity Downloaded in Excel Successfully.");
            }
        }
    }
}
