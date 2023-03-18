using DapperPlus.Attributes;

namespace ExcelImport.Models
{
    [Table("orderdetails")]
    public class UploadOrderModel
    {
        public string product { get; set; }
        public string category { get; set; }
        public string region { get; set; }
        public string city { get; set; }
        public string filename { get; set; }
        public string orderdate { get; set; }
        public double unitprice { get; set; }
        public double totalprice { get; set; }
        public double quantity { get; set; }
    }
}
