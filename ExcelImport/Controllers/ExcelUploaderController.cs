using ExcelImport.Models;
using Microsoft.AspNetCore.Mvc;
using ExcelImport.Services.BIInterface;

namespace ExcelImport.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelUploaderController : ControllerBase
    {
        private readonly ILogger<ExcelUploaderController> _logger;
        private readonly IExcelImport _excelImport;
        public ExcelUploaderController(ILogger<ExcelUploaderController> logger, IExcelImport excelImport)
        {
            _logger = logger;
            _excelImport = excelImport;
        }

        [HttpPost]
        [Route("UploadData")]
        /* If Url is other than localhost:4200 add it in program.cs cors policy*/
        public async Task<ResultModel> UploadExcel()
        {
            ResultModel result = new ResultModel();
            try
            {
                IFormFile files = Request.Form.Files[0];
                var list = _excelImport.UploadExcel(files);
                result.isSuccess = true;
                result.data = list.Result;
                result.message = "Successfully uploaded the file";
            }
            catch (Exception ex)
            {
                result.isSuccess = false;
                result.data = null;
                result.message = "Error in ExcelImport Controller: "+ex;
            }
            return result;
        }
       
    }
    
}
