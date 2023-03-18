using Dapper;
using ExcelImport.Models;
using ExcelImport.Services.BIInterface;
using OfficeOpenXml;
using System.Data.SqlClient;
using System.Transactions;

namespace ExcelImport.Services.BIService
{
    public class ExcelImportService:IExcelImport
    {
        private IConfiguration Configuration;
        public ExcelImportService(IConfiguration _configuration)
        {
            Configuration = _configuration;
        }
        public async Task<List<UploadOrderModel>> UploadExcel(IFormFile files)
        {
            var filename = Path.GetFileNameWithoutExtension(files.FileName);
            var list = new List<UploadOrderModel>();
            using (var stream = new MemoryStream())
            {
                await files.CopyToAsync(stream);
                #region Writing Excel data into Model
                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    var rowcount = worksheet.Dimension.Rows;
                    for (int row = 2; row <= rowcount; row++)
                    {
                        if (worksheet.Cells[row, 1].Value != null)
                        {
                            list.Add(new UploadOrderModel
                            {
                                orderdate = worksheet.Cells[row, 1].Value.ToString().Trim(),
                                region = worksheet.Cells[row, 2].Value.ToString().Trim(),
                                city = worksheet.Cells[row, 3].Value.ToString().Trim(),
                                category = worksheet.Cells[row, 4].Value.ToString().Trim(),
                                product = worksheet.Cells[row, 5].Value.ToString().Trim(),
                                quantity = (double)worksheet.Cells[row, 6].Value,
                                unitprice = (double)worksheet.Cells[row, 7].Value,
                                totalprice = (double)worksheet.Cells[row, 8].Value,
                                filename = files.FileName
                            });
                        }
                    }
                }
                #endregion
                string connString = this.Configuration.GetConnectionString("sqlconnection");/*Getting Connection String From Appsetting/ Replace Your Connection String Appsettings.json*/

                #region Inserting Data into Database Table
                using (var connection = new SqlConnection(connString))
                {
                    using (var scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
                    {

                        connection.Open();
                        try
                        {
                            connection.InsertBulk(list);
                            list = connection.Query<UploadOrderModel>("Select * from get_orderdetails('"+ files.FileName + "')").ToList();
                        }
                        catch (Exception ex)
                        {

                        }
                        connection.Close();
                        scope.Complete();
                    }
                }
                #endregion
            }
            return list;
        }
    }
}
