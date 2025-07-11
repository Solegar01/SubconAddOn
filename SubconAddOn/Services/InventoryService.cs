using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SAPbobsCOM;
using System.Linq;
using System.Configuration;
using System.Data.SqlClient;
using SubconAddOn.Models;

namespace SubconAddOn.Services
{
    public static class InventoryService
    {
        private static readonly Company oCompany = CompanyService.GetCompany();

        public static int CreateGoodsIssue(GoodsIssueModel model)
        {
            if (oCompany == null || !oCompany.Connected)
                throw new InvalidOperationException("DI Company belum terkoneksi.");

            if (model.Lines == null || model.Lines.Count() == 0)
                throw new ArgumentException("Lines kosong.", nameof(model.Lines));

            Documents gi = null;
            bool ownTrans = false;

            try
            {
                if (!oCompany.InTransaction)                  // buka transaksi jika belum ada
                {
                    oCompany.StartTransaction();
                    ownTrans = true;
                }

                gi = (Documents)oCompany.GetBusinessObject(BoObjectTypes.oInventoryGenExit);

                // ===== HEADER =====
                gi.DocDate = model.DocDate;
                gi.TaxDate = gi.DocDate;
                if (model.Lines.Any())
                {
                    var poDocEntry = model.Lines.ElementAt(0).PODocEntry;
                    gi.UserFields.Fields.Item("U_T2_Ref_PO").Value = poDocEntry;
                }

                // ===== LINES =====
                for (int i = 0; i < model.Lines.Count(); i++)
                {
                    var l = model.Lines.ElementAt(i);

                    gi.Lines.ItemCode = l.ItemCode;
                    gi.Lines.Quantity = l.Quantity;
                    gi.Lines.WarehouseCode = l.WarehouseCode;

                    if (!string.IsNullOrEmpty(l.AccountCode))
                        gi.Lines.AccountCode = l.AccountCode;

                    // Tambah baris berikutnya jika belum di baris terakhir
                    if (i < model.Lines.Count() - 1)
                        gi.Lines.Add();
                }

                // ===== SIMPAN DOKUMEN =====
                if (gi.Add() != 0)
                {
                    oCompany.GetLastError(out int ec, out string em);
                    throw new Exception($"Gagal Goods Issue [{ec}] {em}");
                }

                int docEntry = int.Parse(oCompany.GetNewObjectKey());

                if (ownTrans)
                    oCompany.EndTransaction(BoWfTransOpt.wf_Commit);

                return docEntry;
            }
            catch(Exception e)
            {
                if (ownTrans && oCompany.InTransaction)
                    oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                throw e;
            }
            finally
            {
                if (gi != null)
                    Marshal.ReleaseComObject(gi);
            }
        }

        public static int CreateGoodsReceipt(GoodsReceiptModel model)
        {
            if (oCompany == null || !oCompany.Connected)
                throw new InvalidOperationException("DI Company belum terkoneksi.");

            if (model.Lines == null || model.Lines.Count() == 0)
                throw new ArgumentException("Lines kosong.", nameof(model.Lines));

            Documents gr = null;
            bool ownTrans = false;

            try
            {
                if (!oCompany.InTransaction)                  // buka transaksi jika belum ada
                {
                    oCompany.StartTransaction();
                    ownTrans = true;
                }

                gr = (Documents)oCompany.GetBusinessObject(BoObjectTypes.oInventoryGenEntry);

                // ===== HEADER =====
                gr.DocDate = model.DocDate;
                gr.TaxDate = gr.DocDate;
                if (model.Lines.Any())
                {
                    var poDocEntry = model.Lines.ElementAt(0).PODocEntry;
                    gr.UserFields.Fields.Item("U_T2_Ref_PO").Value = poDocEntry;
                }

                // ===== LINES =====
                for (int i = 0; i < model.Lines.Count(); i++)
                {
                    var l = model.Lines.ElementAt(i);

                    gr.Lines.ItemCode = l.ItemCode;
                    gr.Lines.Quantity = l.Quantity;
                    gr.Lines.WarehouseCode = l.WarehouseCode;
                    gr.Lines.UnitPrice = l.UnitPrice;

                    if (!string.IsNullOrEmpty(l.AccountCode))
                        gr.Lines.AccountCode = l.AccountCode;

                    // Tambah baris berikutnya jika belum di baris terakhir
                    if (i < model.Lines.Count() - 1)
                        gr.Lines.Add();
                }

                // ===== SIMPAN DOKUMEN =====
                if (gr.Add() != 0)
                {
                    oCompany.GetLastError(out int ec, out string em);
                    throw new Exception($"Gagal Goods Receipt [{ec}] {em}");
                }

                int docEntry = int.Parse(oCompany.GetNewObjectKey());

                if (ownTrans)
                    oCompany.EndTransaction(BoWfTransOpt.wf_Commit);

                return docEntry;
            }
            catch(Exception e)
            {
                if (ownTrans && oCompany.InTransaction)
                    oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                throw e;
            }
            finally
            {
                if (gr != null)
                    Marshal.ReleaseComObject(gr);
            }
        }

        public static GoodsIssueModel GetGoodIssueByGRPO(int docEntry)
        {
            if (docEntry <= 0)
                throw new ArgumentException("Invalid GRPO DocEntry", nameof(docEntry));

            var dataModel = new GoodsIssueModel
            {
                DocDate = DateTime.Now,
                Lines = new List<GoodsIssueLineModel>()
            };

            string sql = $@"
                        SELECT 
                            t4.Code                          AS ItemCode,
                            (t4.Quantity * t1.Quantity)      AS Quantity,
                            t1.WhsCode                       AS WarehouseCode,
                            (
                                SELECT TOP 1 WipAcct 
                                FROM OGAR 
                                WHERE UDF1 = '4'
                            )                                AS AccountCode,
                        t1.BaseEntry                         AS PODocEntry
                        FROM OPDN t0
                        JOIN PDN1 t1 ON t1.DocEntry = t0.DocEntry
                        JOIN OITM t2 ON t2.ItemCode = t1.ItemCode
                        JOIN OITT t3 ON t3.Code = t2.U_T2_BOM
                        JOIN ITT1 t4 ON t4.Father = t3.Code
                        WHERE t0.DocEntry = {docEntry}";
            
            Recordset rs = null;
            Recordset rsDoc = null;

            try
            {
                // Ambil lines BOM untuk GRPO
                rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                while (!rs.EoF)
                {
                    var line = new GoodsIssueLineModel
                    {
                        ItemCode = rs.Fields.Item("ItemCode").Value.ToString(),
                        Quantity = Convert.ToDouble(rs.Fields.Item("Quantity").Value),
                        WarehouseCode = rs.Fields.Item("WarehouseCode").Value.ToString(),
                        AccountCode = rs.Fields.Item("AccountCode").Value.ToString(),
                        PODocEntry = Convert.ToInt64(rs.Fields.Item("PODocEntry").Value),
                    };

                    dataModel.Lines.Add(line);
                    rs.MoveNext();
                }

                if (dataModel.Lines.Count == 0)
                    throw new Exception("No BOM components found for this GRPO.");
                
                return dataModel;
            }
            catch (Exception ex)
            {
                throw new Exception("Error while retrieving BOM for GRPO: " + ex.Message, ex);
            }
            finally
            {
                if (rs != null) Marshal.ReleaseComObject(rs);
                if (rsDoc != null) Marshal.ReleaseComObject(rsDoc);
            }
        }
        
        public static GoodsReceiptModel GetGoodReceiptByGRPO(int grpoDocEntry, int giDocEntry)
        {
            if (grpoDocEntry <= 0)
                throw new ArgumentException("Invalid GRPO DocEntry", nameof(grpoDocEntry));
            if (giDocEntry <= 0)
                throw new ArgumentException("Invalid GI DocEntry", nameof(giDocEntry));

            var dataModel = new GoodsReceiptModel
            {
                DocDate = DateTime.Now,
                Lines = new List<GoodsReceiptLineModel>()
            };

            string sql = $@"
                SELECT 
                T3.Code AS ItemCode,
                T1.Quantity,
			                T3.ToWH AS WarehouseCode,
			                (
					                SELECT TOP 1 WipAcct 
					                FROM OGAR 
					                WHERE UDF1 = '4'
			                ) AS AccountCode,
			                (
					                (
							                SELECT SUM(ISNULL(_T1.LineTotal, 0))
							                FROM OIGE _t0
							                INNER JOIN IGE1 _t1 ON _T1.DocEntry = _t0.DocEntry
							                INNER JOIN ITT1 _t2 ON _T2.Code = _T1.ItemCode
							                WHERE _t0.DocEntry = {giDocEntry}      
								                AND _T2.Father = T3.Code
					                ) + ISNULL(T1.LineTotal, 0)
			                ) / NULLIF(T1.Quantity, 0) AS UnitPrice,
                T1.BaseEntry AS PODocEntry
                FROM OPDN T0
                INNER JOIN PDN1 T1 ON T1.DocEntry=T0.DocEntry
                INNER JOIN OITM T2 ON T2.ItemCode=T1.ItemCode
                INNER JOIN OITT T3 ON T3.Code=T2.U_T2_BOM
                WHERE T0.DocEntry={grpoDocEntry}";
            
            Recordset rs = null;
            Recordset rsDoc = null;

            try
            {
                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                while (!rs.EoF)
                {
                    double unitPrice = 0;
                    double.TryParse(rs.Fields.Item("UnitPrice").Value.ToString(), out unitPrice);
                    var line = new GoodsReceiptLineModel
                    {
                        ItemCode = rs.Fields.Item("ItemCode").Value.ToString(),
                        Quantity = Convert.ToDouble(rs.Fields.Item("Quantity").Value),
                        WarehouseCode = rs.Fields.Item("WarehouseCode").Value.ToString(),
                        AccountCode = rs.Fields.Item("AccountCode").Value.ToString(),
                        UnitPrice = unitPrice,
                        PODocEntry = Convert.ToInt64(rs.Fields.Item("PODocEntry").Value),
                    };

                    dataModel.Lines.Add(line);
                    rs.MoveNext();
                }

                if (dataModel.Lines.Count == 0)
                    throw new Exception("No data found for Goods Receipt.");
                
                return dataModel;
            }
            catch (Exception ex)
            {
                throw new Exception("Error while retrieving WIP production orders: " + ex.Message, ex);
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                if (rsDoc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsDoc);
            }
        }


    }
}
