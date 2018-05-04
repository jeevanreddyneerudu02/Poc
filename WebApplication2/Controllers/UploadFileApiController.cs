using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using System.Data.OleDb;
using System.IO;
using System.Threading.Tasks;

namespace WebApplication2.Controllers
{
    public class UploadFileApiController : ApiController
    {
        #region Constants
        static string filePath;
        static string fileName;
        static string copiedFilePath;
        #endregion

        [HttpPost]
        public async Task<HttpResponseMessage> UploadJsonFile()
        {
            HttpResponseMessage response = new HttpResponseMessage();
            var httpRequest = HttpContext.Current.Request;
            if (httpRequest.Files.Count > 0)
            {
                foreach (string file in httpRequest.Files)
                {
                    var postedFile = httpRequest.Files[file];

                    if (File.Exists(HttpContext.Current.Server.MapPath("~/App_Data/" + postedFile.FileName)))
                    {
                        File.Delete(HttpContext.Current.Server.MapPath("~/App_Data/" + postedFile.FileName));
                    }

                    filePath = HttpContext.Current.Server.MapPath("~/App_Data/" + postedFile.FileName);

                    fileName = postedFile.FileName;

                    postedFile.SaveAs(filePath);
                }
            }
            return response;
        }

        public async Task<IHttpActionResult> Get()
        {
            var dataBytes = File.ReadAllBytes(copiedFilePath);
            var dataStream = new MemoryStream(dataBytes);
            return new DownLoadFileResult(dataStream, Request, fileName);
        }

        [HttpPost]
        public async Task<HttpResponseMessage> CopyExcel([FromBody]string[] values)
        {
            Dictionary<string, string> dictionary = new Dictionary<string, string>();
            for (int i = 0; i < values.Length; i++)
            {
                string[] tokens = values[i].Split(':');
                dictionary.Add(tokens[0], tokens[1]);
            }

            DataSet DataSet;

            try
            {
                OleDbConnection MyConnection;
                OleDbDataAdapter MyCommand;

                 MyConnection = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=" + filePath);
               // MyConnection = new OleDbConnection(" Provider = Microsoft.ACE.OLEDB.12.0;  Extended Properties = Excel 12.0 Xml;HDR=YES;Data Source=" + filePath);
             
                MyCommand = new OleDbDataAdapter("select * from [Sheet1$]", MyConnection);

                MyCommand.TableMappings.Add("Table", "TestTable");

                DataSet = new System.Data.DataSet();

                MyCommand.Fill(DataSet);

                MyConnection.Close();
                ExportDataSetToExcel(DataSet,dictionary);
            }

            catch (Exception ex)
            {
                throw ex;
            }
            return Request.CreateResponse(HttpStatusCode.Created);
        }

        private void ExportDataSetToExcel(DataSet ds, Dictionary<string, string> dictionary)
        {
            try
            {
                if (File.Exists(HttpContext.Current.Server.MapPath("~/Copy/" + fileName)))
                {
                    File.Delete(HttpContext.Current.Server.MapPath("~/Copy/" + fileName));
                }
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

             
                copiedFilePath = HttpContext.Current.Server.MapPath("~/Copy/" + fileName);
                xlWorkBook.SaveAs(HttpContext.Current.Server.MapPath("~/Copy/" + fileName));
                xlWorkBook.Save();
                xlWorkBook.Close();
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public class DownLoadFileResult : IHttpActionResult
        {
            MemoryStream fileStuff;
            string excelFileName;
            HttpRequestMessage httpRequestMessage;
            HttpResponseMessage httpResponseMessage;
            public DownLoadFileResult(MemoryStream data, HttpRequestMessage request, string filename)
            {
                fileStuff = data;
                httpRequestMessage = request;
                excelFileName = filename;
            }

            public System.Threading.Tasks.Task<HttpResponseMessage> ExecuteAsync(System.Threading.CancellationToken cancellationToken)
            {
                httpResponseMessage = httpRequestMessage.CreateResponse(HttpStatusCode.OK);
                httpResponseMessage.Content = new StreamContent(fileStuff);
                //httpResponseMessage.Content = new ByteArrayContent(bookStuff.ToArray());
                httpResponseMessage.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment");
                httpResponseMessage.Content.Headers.ContentDisposition.FileName = excelFileName;
                httpResponseMessage.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");

                return System.Threading.Tasks.Task.FromResult(httpResponseMessage);
            }
        }
    }
}
