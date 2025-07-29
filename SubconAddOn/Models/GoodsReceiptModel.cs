using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SubconAddOn.Models
{
    public class GoodsReceiptModel
    {
        public DateTime DocDate { get; set; }
        public int GRPODocEntry { get; set; }
        public List<GoodsReceiptLineModel> Lines { get; set; } = new List<GoodsReceiptLineModel>();
    }

    public class GoodsReceiptLineModel
    {
        public string ItemCode { get; set; }   // wajib
        public double Quantity { get; set; }   // wajib
        public string WarehouseCode { get; set; }   // wajib
        public string AccountCode { get; set; }   // opsional (non‑stock)
        public double UnitPrice { get; set; }
        public string PODocEntry { get; set; }
        public string GRPOLineNum { get; set; }

    }
}
