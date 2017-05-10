using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.IO;
using OfficeOpenXml;

namespace AutomationUtils
{
    public class SQL2Excel
    {
        // Execute SQL and save results in a sheet
        // Support SQL Server
        public static void saveSQLResultsInExcel(string connectionString, string queryString, string excelFilePath, string sheetName)
        {
            var ExcelFile = new FileInfo(excelFilePath);
            using (ExcelPackage package = new ExcelPackage(ExcelFile))
            {

                ExcelWorksheet workSheet;

                if (null != package.Workbook.Worksheets[sheetName])
                {
                    workSheet = package.Workbook.Worksheets[sheetName];
                }
                else
                {
                    workSheet = package.Workbook.Worksheets.Add(sheetName);
                }

                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    connection.Open();
                    SqlCommand command = new SqlCommand(queryString, connection);                    
                    SqlDataReader reader = command.ExecuteReader();
                    command.Dispose();
                    reader.Close();
                    connection.Close();
                }

                package.Save();
              
            }
        }
    }
}
