using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using MySql.Data.MySqlClient;
using MySqlImportExportExcelProject.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace MySqlImportExportExcelProject.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private string connectionString = "Server=server;Port=port;Database=dbname;Uid=user;Pwd=pwd;CharSet=utf8mb4;";

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }
        public void ImportDataFromExcelToMySql()
        {
            string excelFilePath = "C:/Users/rashe/source/repos/data.xlsx";

            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                using (var connection = new MySqlConnection(connectionString))
                {
                    connection.Open();

                    var worksheet = package.Workbook.Worksheets[0]; // Assuming you are working with the first worksheet.

                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;

                    for (int row = 2; row <= rowCount; row++) // Start from row 2 (assuming the first row is header).
                    {
                        var values = new object[colCount];
                        for (int col = 1; col <= colCount; col++)
                        {
                            values[col - 1] = worksheet.Cells[row, col].Value;
                        }

                        // Perform validation here
                        string name = values[1] as string; // Assuming the name is in the first column.
                        if (!string.IsNullOrEmpty(name))
                        {
                            // Check if the value already exists in the database
                            string selectSql = "SELECT COUNT(*) FROM area WHERE name = @name";
                            using (var selectCmd = new MySqlCommand(selectSql, connection))
                            {
                                selectCmd.Parameters.AddWithValue("@name", name);
                                int count = Convert.ToInt32(selectCmd.ExecuteScalar());

                                // If count is 0, the value doesn't exist in the database, so insert it
                                if (count == 0)
                                {
                                    // Construct your SQL INSERT statement and execute it.
                                    string insertSql = "INSERT INTO area (name) VALUES (@name)";
                                    using (var cmd = new MySqlCommand(insertSql, connection))
                                    {
                                        cmd.Parameters.AddWithValue("@name", name);
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Handle invalid data, log it, or perform any necessary action.
                            // For example, you can log an error message.
                            _logger.LogError("Invalid data found in Excel at row {Row}. Skipping.", row);
                        }
                    }

                    connection.Close();
                }
            }

            _logger.LogInformation("Data imported successfully.");
        }

        public void ExportDataFromMySqlToExcel()
        {
            string excelFilePath = "output_excel_file.xlsx";

            using (var package = new ExcelPackage())
            {
                using (var connection = new MySqlConnection(connectionString))
                {
                    connection.Open();
                    string query = "SELECT * FROM area"; // Replace YourTableName with your table name.

                    using (var adapter = new MySqlDataAdapter(query, connection))
                    {
                        var dataTable = new DataTable();
                        adapter.Fill(dataTable);

                        var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                        worksheet.Cells.LoadFromDataTable(dataTable, true);
                    }

                    connection.Close();
                }

                package.SaveAs(new FileInfo(excelFilePath));
            }

            Console.WriteLine("Data exported successfully.");
        }
        public IActionResult Index()
        {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // try
            // {
            //ImportDataFromExcelToMySql();
            //ExportDataFromMySqlToExcel();
            //}catch(Exception ex)
            //{
            // Console.WriteLine(""+ex.Message);
            //}
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
