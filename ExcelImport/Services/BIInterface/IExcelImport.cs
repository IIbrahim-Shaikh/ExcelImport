using ExcelImport.Models;

namespace ExcelImport.Services.BIInterface
{
    public interface IExcelImport
    {
        public Task<List<UploadOrderModel>> UploadExcel(IFormFile file);
    }
}
