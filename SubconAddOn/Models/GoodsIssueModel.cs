using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SubconAddOn.Models
{
    public class GoodsIssueModel
    {
        public DateTime DocDate { get; set; }
        public List<GoodsIssueLineModel> Lines { get; set; } = new List<GoodsIssueLineModel>();
    }

    public class GoodsIssueLineModel
    {
        public string ItemCode { get; set; }   // wajib
        public double Quantity { get; set; }   // wajib
        public string WarehouseCode { get; set; }   // wajib
        public string AccountCode { get; set; }   // opsional (non‑stock)
        public long PODocEntry { get; set; }   

    }
}
