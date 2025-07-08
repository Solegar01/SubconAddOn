using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SAPbobsCOM;
using Dapper;
using System.Linq;
using System.Configuration;
using System.Data.SqlClient;

namespace SubconAddOn.Services
{
    /// <summary>
    /// DTO baris Goods Issue.
    /// </summary>
    public class GoodsIssueModel
    {
        public DateTime docDate { get; set; }
        public string comments { get; set; }
        public IEnumerable<GoodsIssueLineModel> lines { get; set; }
    }

    public class GoodsIssueLineModel
    {
        public string ItemCode { get; set; }   // wajib
        public double Quantity { get; set; }   // wajib
        public string WarehouseCode { get; set; }   // wajib
        public string AccountCode { get; set; }   // opsional (non‑stock)
        
    }

    public class GoodsReceiptModel
    {
        public DateTime docDate { get; set; }
        public string comments { get; set; }
        public IEnumerable<GoodsReceiptLineModel> lines { get; set; }
    }

    public class GoodsReceiptLineModel
    {
        public string ItemCode { get; set; }   // wajib
        public double Quantity { get; set; }   // wajib
        public string WarehouseCode { get; set; }   // wajib
        public string AccountCode { get; set; }   // opsional (non‑stock)
        
    }

    public static class InventoryService
    {
        private static readonly string _connStr =
        ConfigurationManager.ConnectionStrings["B1Connection"].ConnectionString;

        public static int CreateGoodsIssue(Company cmp, GoodsIssueModel model)
        {
            if (cmp == null || !cmp.Connected)
                throw new InvalidOperationException("DI Company belum terkoneksi.");

            if (model.lines == null || model.lines.Count() == 0)
                throw new ArgumentException("Lines kosong.", nameof(model.lines));

            Documents gi = null;
            bool ownTrans = false;

            try
            {
                if (!cmp.InTransaction)                  // buka transaksi jika belum ada
                {
                    cmp.StartTransaction();
                    ownTrans = true;
                }

                gi = (Documents)cmp.GetBusinessObject(BoObjectTypes.oInventoryGenExit);

                // ===== HEADER =====
                gi.DocDate = model.docDate;
                gi.TaxDate = gi.DocDate;
                gi.Comments = model.comments;

                // ===== LINES =====
                for (int i = 0; i < model.lines.Count(); i++)
                {
                    var l = model.lines.ElementAt(i);

                    gi.Lines.ItemCode = l.ItemCode;
                    gi.Lines.Quantity = l.Quantity;
                    gi.Lines.WarehouseCode = l.WarehouseCode;

                    if (!string.IsNullOrEmpty(l.AccountCode))
                        gi.Lines.AccountCode = l.AccountCode;

                    // Tambah baris berikutnya jika belum di baris terakhir
                    if (i < model.lines.Count() - 1)
                        gi.Lines.Add();
                }

                // ===== SIMPAN DOKUMEN =====
                if (gi.Add() != 0)
                {
                    cmp.GetLastError(out int ec, out string em);
                    throw new Exception($"Gagal Goods Issue [{ec}] {em}");
                }

                int docEntry = int.Parse(cmp.GetNewObjectKey());

                if (ownTrans)
                    cmp.EndTransaction(BoWfTransOpt.wf_Commit);

                return docEntry;
            }
            catch
            {
                if (ownTrans && cmp.InTransaction)
                    cmp.EndTransaction(BoWfTransOpt.wf_RollBack);
                throw;
            }
            finally
            {
                if (gi != null)
                    Marshal.ReleaseComObject(gi);
            }
        }

        public static int CreateGoodsReceipt(Company cmp, GoodsReceiptModel model)
        {
            if (cmp == null || !cmp.Connected)
                throw new InvalidOperationException("DI Company belum terkoneksi.");

            if (model.lines == null || model.lines.Count() == 0)
                throw new ArgumentException("Lines kosong.", nameof(model.lines));

            Documents gr = null;
            bool ownTrans = false;

            try
            {
                if (!cmp.InTransaction)                  // buka transaksi jika belum ada
                {
                    cmp.StartTransaction();
                    ownTrans = true;
                }

                gr = (Documents)cmp.GetBusinessObject(BoObjectTypes.oInventoryGenEntry);

                // ===== HEADER =====
                gr.DocDate = model.docDate;
                gr.TaxDate = gr.DocDate;
                gr.Comments = model.comments;

                // ===== LINES =====
                for (int i = 0; i < model.lines.Count(); i++)
                {
                    var l = model.lines.ElementAt(i);

                    gr.Lines.ItemCode = l.ItemCode;
                    gr.Lines.Quantity = l.Quantity;
                    gr.Lines.WarehouseCode = l.WarehouseCode;

                    if (!string.IsNullOrEmpty(l.AccountCode))
                        gr.Lines.AccountCode = l.AccountCode;

                    // Tambah baris berikutnya jika belum di baris terakhir
                    if (i < model.lines.Count() - 1)
                        gr.Lines.Add();
                }

                // ===== SIMPAN DOKUMEN =====
                if (gr.Add() != 0)
                {
                    cmp.GetLastError(out int ec, out string em);
                    throw new Exception($"Gagal Goods Receipt [{ec}] {em}");
                }

                int docEntry = int.Parse(cmp.GetNewObjectKey());

                if (ownTrans)
                    cmp.EndTransaction(BoWfTransOpt.wf_Commit);

                return docEntry;
            }
            catch
            {
                if (ownTrans && cmp.InTransaction)
                    cmp.EndTransaction(BoWfTransOpt.wf_RollBack);
                throw;
            }
            finally
            {
                if (gr != null)
                    Marshal.ReleaseComObject(gr);
            }
        }

        public static GoodsIssueModel GetGoodIssueByGRPO(int docEntry)
        {
            if (docEntry == 0)
                throw new ArgumentNullException(nameof(docEntry));

            var dataModel = new GoodsIssueModel();
            dataModel.docDate = DateTime.Now;
            dataModel.comments = "Auto Generated";
            
            const string sql = @"
            SELECT 
              t4.Code                          AS ItemCode,        -- Item anak (BOM)
              (t4.Quantity * t1.Quantity)      AS Quantity,        -- Qty kebutuhan aktual
              t1.WhsCode                       AS WarehouseCode,   -- Gudang GRPO
              (
                SELECT TOP 1 WipAcct 
                FROM OGAR 
                WHERE UDF1 = '4'
              )                                AS AccountCode      -- Akun WIP dari custom logic
            FROM OPDN  t0                      -- Header GRPO
            JOIN PDN1  t1 ON t1.DocEntry = t0.DocEntry      -- Baris GRPO
            JOIN OITM  t2 ON t2.ItemCode = t1.ItemCode      -- Item master
            JOIN OITT  t3 ON t3.Code     = t2.U_T2_BOM      -- BOM header (custom UDF: U_T2_BOM)
            JOIN ITT1  t4 ON t4.Father   = t3.Code          -- BOM child items
            WHERE t0.DocEntry = @DocEntry;";


            try
            {
                using (var cn = new SqlConnection(_connStr))
                {
                    cn.Open();
                    var result = cn.Query<GoodsIssueLineModel>(sql, new { DocEntry = docEntry }).ToList();
                    if (result != null)
                    {
                        dataModel.lines = result;
                    }
                    var grpoDocNum = cn.Query<string>("SELECT DocNum FROM OPDN WHERE DocEntry = @DocEntry", new { DocEntry = docEntry }).FirstOrDefault();

                    dataModel.comments = $"Auto Generated by GRPO - {grpoDocNum ?? ""}";
                }
                return dataModel;
            }
            catch (Exception ex)
            {
                throw new Exception("Error while retrieving WIP production orders: " + ex.Message, ex);
            }
        }

        public static GoodsReceiptModel GetGoodReceiptByGRPO(int docEntry)
        {
            if (docEntry == 0)
                throw new ArgumentNullException(nameof(docEntry));

            var dataModel = new GoodsReceiptModel();
            dataModel.docDate = DateTime.Now;

            const string sql = @"
            SELECT 
                t3.Code AS ItemCode,
			    t1.Quantity,
			    t3.ToWH AS WarehouseCode,
                (
                    SELECT TOP 1 WipAcct 
                    FROM OGAR 
                    WHERE UDF1 = '4'
                )                       AS AccountCode      -- Akun WIP dari custom logic
            FROM OPDN  t0                      -- Header GRPO
            JOIN PDN1  t1 ON t1.DocEntry = t0.DocEntry      -- Baris GRPO
            JOIN OITM  t2 ON t2.ItemCode = t1.ItemCode      -- Item master
            JOIN OITT  t3 ON t3.Code     = t2.U_T2_BOM      -- BOM header (custom UDF: U_T2_BOM)
            WHERE t0.DocEntry = @DocEntry;";


            try
            {
                using (var cn = new SqlConnection(_connStr))
                {
                    cn.Open();
                    var result = cn.Query<GoodsReceiptLineModel>(sql, new { DocEntry = docEntry }).ToList();
                    if (result != null)
                    {
                        dataModel.lines = result;
                    }
                    var grpoDocNum = cn.Query<string>("SELECT DocNum FROM OPDN WHERE DocEntry = @DocEntry", new { DocEntry = docEntry }).FirstOrDefault();

                    dataModel.comments = $"Auto Generated by GRPO - {grpoDocNum ?? ""}";
                }
                return dataModel;
            }
            catch (Exception ex)
            {
                throw new Exception("Error while retrieving WIP production orders: " + ex.Message, ex);
            }
        }
    }


}
