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
            string connectionString = "Data Source=localhost\\SQLEXPRESS;Database=DB_Excel;Integrated Security=true";
            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();
            Console.WriteLine("done");

            //อ่านไฟล์ Excel 
            using (var package = new ExcelPackage(new FileInfo("Tax.xlsx")))
            {
                var worksheet = package.Workbook.Worksheets["ภาษีบำรุงท้องที่"];
                int totalRows = worksheet.Dimension.End.Row;

                
                //อ่าน row และบันทึกข้อมูลลง SQL Server
                for (int i = 2; i <= totalRows; i++)
                {
                    //ปี
                    Console.WriteLine($"Cell A2 Value   : {worksheet.Cells[i,1].Text.Remove(0,4)}");
                    //เดือน
                    Console.WriteLine($"Cell A2 Value   : {worksheet.Cells[i,1].Text.Remove(3,3)}");

                string strSQL = "INSERT INTO TAX_INCOME (Tax_Payment,Tax_Year,Tax_Month) "
                    + " VALUES  ("                   
                    + " '" + worksheet.Cells[i, 2].Text+"',"                      
                    + " '" + worksheet.Cells[i, 1].Text.Remove(0,4) + "', "
                    + " '" + worksheet.Cells[i, 1].Text.Remove(3,3) + "' "
                    + ")";
                var objCmd = new SqlCommand(strSQL, conn);
                objCmd.ExecuteNonQuery();                        
                }
                conn.Close();
            }
        }
    }
}
