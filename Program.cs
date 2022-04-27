using System;
using System.Data.SqlClient;
using System.IO;
using ExcelDataReader;

using OfficeOpenXml;
using Syncfusion.XlsIO;

namespace repos
{
    class Program
    {

        static void Main(string[] args)
        {
            //เชื่อมต่อ SQL Server
            // string connectionString = "Data Source=localhost\\SQLEXPRESS;Database=DB_Excel;Integrated Security=true";
            // SqlConnection conn = new SqlConnection(connectionString);
            // conn.Open();
            // Console.WriteLine("done");

            //อ่านไฟล์ Excel 
            using (var package = new ExcelPackage(new FileInfo("Tax.xlsx")))
            {
                var worksheet = package.Workbook.Worksheets["ภาษีบำรุงท้องถิ่น"];
                int totalRows = worksheet.Dimension.End.Row;

               
                //อ่าน row และบันทึกข้อมูลลง SQL Server
                for (int i = 2; i <= totalRows; i++)
                {
                    // string strSQL = "INSERT INTO TAX_INCOME (id,TAX_PAYMENT) "
                    //     + " VALUES  ("                   
                    //     + " '" + worksheet.Cells[i, 1].Text.ToString() + "',"                      
                    //     + " '" + worksheet.Cells[i, 2].Text + "' "
                    //     + ")";
                    // var objCmd = new SqlCommand(strSQL, conn);
                    // objCmd.ExecuteNonQuery();
                    Console.WriteLine(worksheet.Cells[i, 1].Text.ToString());
                          
                }
                 //conn.Close();
             }
        }
    }
}
