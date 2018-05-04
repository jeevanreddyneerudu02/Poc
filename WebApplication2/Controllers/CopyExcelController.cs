using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.IO;

namespace WebApplication2.Controllers
{
    public class CopyExcelController : ApiController
    {
        [HttpPost]
        public HttpResponseMessage CopyExcel([FromBody]string[] values)
        {
            Dictionary<string, string> dictionary = new Dictionary<string, string>();
            for(int i=0;i<values.Length;i++)
            {
                string[] tokens = values[i].Split(':');
                dictionary.Add(tokens[0],tokens[1]);
            }
           
            DataSet DataSet;

            try
            {
                OleDbConnection MyConnection;
                OleDbDataAdapter MyCommand;
                string filePath = (Directory.GetCurrentDirectory()) + "\\JeevanExcel6.xls";
                MyConnection = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=" + filePath);

                MyCommand = new OleDbDataAdapter("select * from [Sheet1$]", MyConnection);

                MyCommand.TableMappings.Add("Table", "TestTable");

                DataSet = new System.Data.DataSet();

                MyCommand.Fill(DataSet);

                MyConnection.Close();
                ExportDataSetToExcel(DataSet);

            }

            catch (Exception ex)
            {
                throw ex;
            }
            return Request.CreateResponse(HttpStatusCode.Created);
        }

        private void ExportDataSetToExcel(DataSet ds)
        {
            try
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = excelApp.Workbooks.Add(misValue);

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                foreach (DataTable table in ds.Tables)
                {

                    Excel.Worksheet excelWorkSheet = xlWorkBook.Sheets.Add();
                    excelWorkSheet.Name = table.TableName;


                    Dictionary<string, string> dictionary = new Dictionary<string, string>();

                    dictionary.Add("ID", "FirstColumn");
                    dictionary.Add("Name", "SecondColumn");


                    string[] selectedColumns = dictionary.Keys.ToArray();


                    DataTable dt = new DataView(table).ToTable(false, selectedColumns);
                    for (int i = 1; i < dt.Columns.Count + 1; i++)
                    {
                        if (dictionary.ContainsKey(dt.Columns[i - 1].ColumnName))
                        {
                            excelWorkSheet.Cells[1, i] = dictionary[dt.Columns[i - 1].ColumnName];
                        }
                    }
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        for (int k = 0; k < dt.Columns.Count; k++)
                        {
                            excelWorkSheet.Cells[j + 2, k + 1] = dt.Rows[j].ItemArray[k].ToString();
                        }
                    }
                }

                if (File.Exists(Directory.GetCurrentDirectory() + "\\Jeevan.xls"))
                {
                    File.Delete(Directory.GetCurrentDirectory() + "\\Jeevan.xls");
                }
                Console.WriteLine(AppDomain.CurrentDomain.BaseDirectory);
                xlWorkBook.SaveAs(System.IO.Directory.GetCurrentDirectory() + "\\Jeevan.xls");
                xlWorkBook.Save();
                xlWorkBook.Close();
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
