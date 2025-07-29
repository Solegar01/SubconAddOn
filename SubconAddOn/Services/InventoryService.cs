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
                throw new InvalidOperationException("DI Company not connected.");

            if (model.Lines == null || model.Lines.Count() == 0)
                throw new ArgumentException("Empty lines.", nameof(model.Lines));

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
                string grpoEntry = model.GRPODocEntry.ToString();
                gi.UserFields.Fields.Item("U_T2_Ref_GRPO").Value = grpoEntry;
                if (model.Lines.Any())
                {
                    var poDocEntry = model.Lines.ElementAt(0).PODocEntry.ToString();
                    gi.UserFields.Fields.Item("U_T2_Ref_PO").Value = poDocEntry;
                }

                // ===== LINES =====
                for (int i = 0; i < model.Lines.Count(); i++)
                {
                    var l = model.Lines.ElementAt(i);

                    gi.Lines.ItemCode = l.ItemCode;
                    gi.Lines.Quantity = l.Quantity;
                    gi.Lines.WarehouseCode = l.WarehouseCode;
                    gi.Lines.UserFields.Fields.Item("U_T2_GRPO_LineNum").Value = l.GRPOLineNum;

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
                    throw new Exception($"Failed to Create Goods Issue [{ec}] {em}");
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
                throw new InvalidOperationException("DI Company not connected.");

            if (model.Lines == null || model.Lines.Count() == 0)
                throw new ArgumentException("Empty lines.", nameof(model.Lines));

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
                string grpoEntry = model.GRPODocEntry.ToString();
                gr.UserFields.Fields.Item("U_T2_Ref_GRPO").Value = grpoEntry;
                if (model.Lines.Any())
                {
                    var poDocEntry = model.Lines.ElementAt(0).PODocEntry.ToString();
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
                    gr.Lines.UserFields.Fields.Item("U_T2_GRPO_LineNum").Value = l.GRPOLineNum;

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
                    throw new Exception($"Failed to Create Goods Receipt [{ec}] {em}");
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
                GRPODocEntry = docEntry,
                Lines = new List<GoodsIssueLineModel>()
            };

            string sql = $@"
                        SELECT 
                            t4.Code                          AS ItemCode,
                            (t4.Quantity * t1.Quantity)      AS Quantity,
                            t4.Warehouse                     AS WarehouseCode,
                            (
                                SELECT TOP 1 WipAcct 
                                FROM OGAR 
                                WHERE UDF1 = '4' AND ISNULL(WipAcct,'') <> ''
                            )                                AS AccountCode,
                        t1.BaseEntry                         AS PODocEntry,
                        t1.LineNum                         AS GRPOLineNum
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
                        PODocEntry =rs.Fields.Item("PODocEntry").Value.ToString(),
                        GRPOLineNum = rs.Fields.Item("GRPOLineNum").Value.ToString(),
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
                GRPODocEntry = grpoDocEntry,
                Lines = new List<GoodsReceiptLineModel>()
            };

            string sql = $@"
            SELECT 
                T3.Code AS ItemCode,
                T1.Quantity,
                T1.WhsCode AS WarehouseCode,
                (
                    SELECT TOP 1 WipAcct 
                    FROM OGAR 
                    WHERE UDF1 = '4' AND ISNULL(WipAcct,'') <> '' 
                ) AS AccountCode,
                (
                    (
                        SELECT SUM(ISNULL(_T1.LineTotal, 0))   -- GI LineTotal is in local currency
                        FROM OIGE _T0
                        INNER JOIN IGE1 _T1 ON _T1.DocEntry = _T0.DocEntry
                        INNER JOIN ITT1 _T2 ON _T2.Code = _T1.ItemCode
                        WHERE _T0.DocEntry = {giDocEntry}
                            AND _T2.Father = T3.Code
                            AND _T1.U_T2_GRPO_LineNum = T1.LineNum
                    ) + ISNULL(T1.LineTotal, 0)                -- GRPO line total in local currency
                ) / NULLIF(T1.Quantity, 0) AS UnitPrice,
                T1.BaseEntry AS PODocEntry,
                T1.LineNum AS GRPOLineNum
            FROM OPDN T0
            INNER JOIN PDN1 T1 ON T1.DocEntry = T0.DocEntry
            INNER JOIN OITM T2 ON T2.ItemCode = T1.ItemCode
            INNER JOIN OITT T3 ON T3.Code = T2.U_T2_BOM
            WHERE T0.DocEntry = {grpoDocEntry}";


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
                        PODocEntry = rs.Fields.Item("PODocEntry").Value.ToString(),
                        GRPOLineNum = rs.Fields.Item("GRPOLineNum").Value.ToString(),
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

        public static void IsStockAvailable(string itemCode, double qty)
        {
            string sql = $@"
                SELECT T2.Code AS ItemCode, T2.Warehouse AS WhsCode, CASE WHEN ISNULL(T3.OnHand,0) >= ({qty} * ISNULL(T2.Quantity,0)) THEN 1 ELSE 0 END IsAvailable 
                FROM OITM T0
                INNER JOIN OITT T1 ON T0.U_T2_BOM=T1.Code
                INNER JOIN ITT1 T2 ON T2.Father = T1.Code
                INNER JOIN OITW T3 ON T3.ItemCode = T2.Code AND T2.Warehouse=T3.WhsCode
                WHERE T0.ItemCode='{itemCode}'";

            Recordset rs = null;
            Recordset rsDoc = null;

            try
            {
                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                while (!rs.EoF)
                {
                    if (rs.Fields.Item("IsAvailable").Value.ToString() == "0")
                    {
                        string itemBom = rs.Fields.Item("ItemCode").Value.ToString();
                        string whsBom = rs.Fields.Item("WhsCode").Value.ToString();
                        throw new Exception($"Stock for item {itemBom} in {whsBom} is not available ");
                    }
                    rs.MoveNext();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                if (rsDoc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsDoc);
            }
        }

        public static bool IsBom(string itemCode)
        {
            bool isAvailable = false;
            string sql = $@"
                SELECT CASE WHEN ISNULL(T0.U_T2_BOM,'') <> '' THEN 1 ELSE 0 END IsBom 
                FROM OITM T0
                WHERE T0.ItemCode='{itemCode}'";

            Recordset rs = null;
            Recordset rsDoc = null;

            try
            {
                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                if (!rs.EoF && rs.Fields.Item("IsBom").Value.ToString() == "1")
                    isAvailable = true;

                return isAvailable;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                if (rsDoc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsDoc);
            }
        }

        public static int CreateJESubcon(int grpoDocEntry)
        {
            int docEntry = 0;
            JournalEntries oJE = (JournalEntries)oCompany.GetBusinessObject(BoObjectTypes.oJournalEntries);
            Recordset rs = null;

            try
            {
                // Load GRPO
                SAPbobsCOM.Documents oGRPO = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
                if (!oGRPO.GetByKey(grpoDocEntry))
                    throw new Exception("GRPO not found");

                // Get local currency
                string localCurrency = "";
                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                rs.DoQuery("SELECT MainCurncy FROM OADM");
                if (!rs.EoF)
                    localCurrency = rs.Fields.Item(0).Value.ToString();

                bool isForeign = !oGRPO.DocCurrency.Equals(localCurrency, StringComparison.OrdinalIgnoreCase);

                // Get account setup
                string sql = @"
                SELECT TOP 1 T0.DfltExpn, T0.WipAcct, T0.WipVarAcct
                FROM OGAR T0
                WHERE T0.UDF1 = '4' AND T0.ItmsGrpCod = 112";
                rs.DoQuery(sql);

                if (rs.EoF)
                    throw new Exception("WIP/Expense account config not found.");

                string expenseAcct = rs.Fields.Item("DfltExpn").Value.ToString();
                string wipAcct = rs.Fields.Item("WipAcct").Value.ToString();
                string varAcct = rs.Fields.Item("WipVarAcct").Value.ToString();

                // Create JE Header
                oJE.ReferenceDate = DateTime.Today;
                oJE.TaxDate = DateTime.Today;
                oJE.DueDate = DateTime.Today;
                oJE.Memo = "GRPO Subcontract JE";

                // ===== Line 1: DEBIT (WIP Account) =====
                oJE.Lines.AccountCode = wipAcct;
                if (isForeign)
                {
                    oJE.Lines.FCCurrency = oGRPO.DocCurrency;
                    oJE.Lines.FCDebit = oGRPO.DocTotalFc;
                }
                else
                {
                    oJE.Lines.Debit = oGRPO.DocTotal;
                }
                oJE.Lines.LineMemo = "Subcon WIP - GRPO";
                oJE.Lines.Add();

                // ===== Line 2: CREDIT (Expense Account) =====
                oJE.Lines.AccountCode = expenseAcct;
                if (isForeign)
                {
                    oJE.Lines.FCCurrency = oGRPO.DocCurrency;
                    oJE.Lines.FCCredit = oGRPO.DocTotalFc;
                }
                else
                {
                    oJE.Lines.Credit = oGRPO.DocTotal;
                }
                oJE.Lines.LineMemo = "Subcon Expense - GRPO";

                // Add the JE
                int result = oJE.Add();

                if (result == 0)
                {
                    var docEntryStr = oCompany.GetNewObjectKey();
                    docEntry = Convert.ToInt32(docEntryStr);
                }
                else
                {
                    oCompany.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Failed to create JE. Error {errCode}: {errMsg}");
                }

                return docEntry;
            }
            catch (Exception ex)
            {
                throw new Exception("CreateJESubcon failed: " + ex.Message, ex);
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        public static bool LinkJEToGRPO(int grpoDocEntry, int jeDocEntry)
        {
            string sql = "SELECT T21.RefDocEntr AS TransId FROM OPDN T0 INNER JOIN PDN21 T21 ON T21.DocEntry = T0.DocEntry WHERE T0.DocEntry = '" + grpoDocEntry + "' AND T21.RefObjType = '30'";
            try
            {
                Recordset rs = null;
                int recCount = 0;
                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                if (!rs.EoF)
                {
                    recCount = rs.RecordCount;
                }
                // Get GRPO document
                SAPbobsCOM.Documents oGRPO = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);

                if (!oGRPO.GetByKey(grpoDocEntry))
                    throw new Exception($"GRPO with DocEntry {grpoDocEntry} not found.");

                // Add Journal Entry as referenced document
                oGRPO.DocumentReferences.Add();
                oGRPO.DocumentReferences.SetCurrentLine(recCount);
                oGRPO.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_JournalEntry; 
                oGRPO.DocumentReferences.ReferencedDocEntry = jeDocEntry;
                oGRPO.DocumentReferences.IssueDate = DateTime.Today;
                oGRPO.DocumentReferences.Remark = "Auto-linked JE";

                // Update GRPO
                int updateResult = oGRPO.Update();

                if (updateResult != 0)
                {
                    oCompany.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Failed to link JE to GRPO. Error {errCode}: {errMsg}");
                }
                return true;
            }
            catch (Exception )
            {
                throw;
            }
        }

        public static int CreateGoodsReceiptFromGI(int giDocEntry)
        {
            Documents oGI = (Documents)oCompany.GetBusinessObject(BoObjectTypes.oInventoryGenExit); // Goods Issue
            Documents oGR = (Documents)oCompany.GetBusinessObject(BoObjectTypes.oInventoryGenEntry); // Goods Receipt

            //GoodsReceiptModel grModel = new GoodsReceiptModel();

            if (!oGI.GetByKey(giDocEntry))
                throw new Exception("Goods Issue not found.");
            
            // Set header info
            oGR.DocDate = DateTime.Now;
            oGR.TaxDate = DateTime.Now;
            oGR.Comments = "Created from canceled GI : " + oGI.DocNum;
            oGR.UserFields.Fields.Item("U_T2_Ref_PO").Value = oGI.UserFields.Fields.Item("U_T2_Ref_PO").Value.ToString();

            // Loop through GI lines and copy to GR
            for (int i = 0; i < oGI.Lines.Count; i++)
            {
                oGI.Lines.SetCurrentLine(i);

                oGR.Lines.ItemCode = oGI.Lines.ItemCode;
                oGR.Lines.Quantity = oGI.Lines.Quantity;
                oGR.Lines.WarehouseCode = oGI.Lines.WarehouseCode;
                oGR.Lines.UnitPrice = oGI.Lines.UnitPrice;
                oGR.Lines.AccountCode = oGI.Lines.AccountCode;
                oGR.Lines.UserFields.Fields.Item("U_T2_GRPO_LineNum").Value = oGI.Lines.UserFields.Fields.Item("U_T2_GRPO_LineNum").Value.ToString();
                // Optional: Copy UDFs or serial/batch if needed

                if (i < oGI.Lines.Count - 1)
                    oGR.Lines.Add();
            }

            if (oGR.Add() != 0)
            {
                oCompany.GetLastError(out int errCode, out string errMsg);
                throw new Exception($"Failed to add Goods Receipt: {errMsg} (Code: {errCode})");
            }

            return int.Parse(oCompany.GetNewObjectKey()); // GR DocEntry
        }

        public static int CreateGoodsIssueFromGR(int grDocEntry)
        {
            Documents oGR = (Documents)oCompany.GetBusinessObject(BoObjectTypes.oInventoryGenEntry); // Goods Receipt
            Documents oGI = (Documents)oCompany.GetBusinessObject(BoObjectTypes.oInventoryGenExit);  // Goods Issue

            if (!oGR.GetByKey(grDocEntry))
                throw new Exception("Goods Receipt not found.");

            // Set GI header info
            oGI.DocDate = DateTime.Now;
            oGI.TaxDate = DateTime.Now;
            oGI.Comments = "Created from canceled GR : " + oGR.DocNum;
            oGI.UserFields.Fields.Item("U_T2_Ref_PO").Value = oGR.UserFields.Fields.Item("U_T2_Ref_PO").Value.ToString();

            // Loop through GR lines and copy to GI
            for (int i = 0; i < oGR.Lines.Count; i++)
            {
                oGR.Lines.SetCurrentLine(i);

                oGI.Lines.ItemCode = oGR.Lines.ItemCode;
                oGI.Lines.Quantity = oGR.Lines.Quantity;
                oGI.Lines.WarehouseCode = oGR.Lines.WarehouseCode;
                oGI.Lines.UnitPrice = oGR.Lines.UnitPrice;
                oGI.Lines.AccountCode = oGR.Lines.AccountCode;
                oGI.Lines.UserFields.Fields.Item("U_T2_GRPO_LineNum").Value = oGR.Lines.UserFields.Fields.Item("U_T2_GRPO_LineNum").Value.ToString();
                // Optional: Copy Batch/Serial, UDFs, CostingCode, etc.

                if (i < oGR.Lines.Count - 1)
                    oGI.Lines.Add();
            }

            if (oGI.Add() != 0)
            {
                oCompany.GetLastError(out int errCode, out string errMsg);
                throw new Exception($"Failed to create Goods Issue. Error {errCode}: {errMsg}");
            }

            return int.Parse(oCompany.GetNewObjectKey()); // GI DocEntry
        }

        public static int CancelJournalEntry(int originalJEDocEntry)
        {
            try
            {
                JournalEntries je = (JournalEntries)oCompany.GetBusinessObject(BoObjectTypes.oJournalEntries);
                if (je.GetByKey(originalJEDocEntry))
                {
                    int result = je.Cancel();
                    if (result != 0)
                    {
                        oCompany.GetLastError(out int errCode, out string errMsg);
                        throw new Exception($"Cancel failed. Error {errCode}: {errMsg}");
                    }
                }

                // Now get the cancellation JE DocEntry
                string sql = $@"
                SELECT TransId FROM OJDT
                WHERE StornoToTr = {originalJEDocEntry}
                  AND TransType = 30";

                Recordset rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                if (!rs.EoF)
                {
                    int cancelDocEntry = Convert.ToInt32(rs.Fields.Item("TransId").Value);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                    return cancelDocEntry;
                }
                else
                {
                    throw new Exception("Cancellation Journal Entry not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception in CancelJournalEntry: " + ex.Message);
                return -1;
            }
        }


        public static int CancelJEByGRPO(int originEntry,int docEntry)
        {
            int res = 0;
            string sql = "SELECT T21.RefDocEntr AS TransId FROM OPDN T0 INNER JOIN PDN21 T21 ON T21.DocEntry = T0.DocEntry WHERE T0.DocEntry = '" + originEntry + "' AND T21.RefObjType = '30'";

            Recordset rs = null;
            Recordset rsDoc = null;

            try
            {
                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                while (!rs.EoF)
                {
                    var transIdTemp = rs.Fields.Item("TransId").Value.ToString();
                    if (int.TryParse(transIdTemp, out int transId))
                    {
                        int reversalEntry = CancelJournalEntry(transId);
                        LinkJEToGRPO(docEntry, reversalEntry);
                    }
                    res++;
                    rs.MoveNext();
                }
                return res;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                if (rsDoc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsDoc);
            }
        }

        public static bool CancelGIGRByGRPO(int originEntry,int docEntry)
        {
            string sqlGI = "SELECT DocEntry FROM OIGE WHERE U_T2_REF_GRPO = " + originEntry;
            string sqlGR = "SELECT DocEntry FROM OIGN WHERE U_T2_REF_GRPO = " + originEntry;

            Recordset rs = null;
            Recordset rsDoc = null;

            var grJE = 0;
            var giCancelJE = 0;
            var grCancelJE = 0;

            try
            {
                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(sqlGI);

                if (!rs.EoF)
                {
                    var giEntry = rs.Fields.Item("DocEntry").Value.ToString();
                    
                    if (int.TryParse(giEntry, out int entry))
                    {
                        var grEntry = CreateGoodsReceiptFromGI(entry);
                        Documents oGR = (Documents)oCompany.GetBusinessObject(BoObjectTypes.oInventoryGenEntry); // Goods Receipt
                        if (oGR.GetByKey(grEntry))
                        {
                            grCancelJE = oGR.TransNum;
                        }
                    }
                }

                rs.DoQuery(sqlGR);

                if (!rs.EoF)
                {
                    var grEntry = rs.Fields.Item("DocEntry").Value.ToString();

                    if (int.TryParse(grEntry, out int entry))
                    {
                        Documents oGR = (Documents)oCompany.GetBusinessObject(BoObjectTypes.oInventoryGenEntry); // Goods Receipt
                        if (oGR.GetByKey(entry))
                        {
                            grJE = oGR.TransNum;
                        }

                        var giEntry  = CreateGoodsIssueFromGR(entry);
                        Documents oGI = (Documents)oCompany.GetBusinessObject(BoObjectTypes.oInventoryGenExit); // Goods Issue
                        if (oGI.GetByKey(giEntry))
                        {
                            giCancelJE = oGI.TransNum;
                        }

                        if (grJE != 0 && giCancelJE != 0)
                        {
                            int jeVar = CreateJEVariance(giCancelJE, grJE);
                            LinkJEToGRPO(docEntry, jeVar);
                        }
                    }
                }
                
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                if (rsDoc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsDoc);
            }
        }

        public static int CreateJEVariance(int giJEId, int grJEId)
        {
            int docEntry = 0;
            Recordset rs = null;
            Recordset rsDoc = null;
            string sql = @"
                SELECT TOP 1 T0.DfltExpn, T0.WipAcct, T0.WipVarAcct
                FROM OGAR T0
                WHERE T0.UDF1 = '4' AND T0.ItmsGrpCod = 112";
            JournalEntries oJE = (JournalEntries)oCompany.GetBusinessObject(BoObjectTypes.oJournalEntries);
            SAPbobsCOM.JournalEntries giJE = (SAPbobsCOM.JournalEntries)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
            SAPbobsCOM.JournalEntries grJE = (SAPbobsCOM.JournalEntries)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
            try
            {
                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                if (rs.EoF)
                    throw new Exception("WIP/Expense account config not found.");
                
                string wipAcct = rs.Fields.Item("WipAcct").Value.ToString();
                string varAcct = rs.Fields.Item("WipVarAcct").Value.ToString();

                if (giJE.GetByKey(giJEId) && grJE.GetByKey(grJEId))
                {
                    double debitGI = 0, creditGI = 0;
                    double debitGR = 0, creditGR = 0;

                    // --- JE 1 Loop ---
                    for (int i = 0; i < giJE.Lines.Count; i++)
                    {
                        giJE.Lines.SetCurrentLine(i);
                        debitGI += giJE.Lines.Debit;
                        creditGI += giJE.Lines.Credit;
                    }

                    // --- JE 2 Loop ---
                    for (int i = 0; i < grJE.Lines.Count; i++)
                    {
                        grJE.Lines.SetCurrentLine(i);
                        debitGR += grJE.Lines.Debit;
                        creditGR += grJE.Lines.Credit;
                    }

                    var diffDeb = (debitGI - debitGR);
                    double absDiff = Math.Abs(diffDeb);

                    // Create JE Header
                    oJE.ReferenceDate = DateTime.Today;
                    oJE.TaxDate = DateTime.Today;
                    oJE.DueDate = DateTime.Today;
                    oJE.Memo = $"Variance between GR JE {grJEId} and GI JE {giJEId}";


                    // WIP account line
                    oJE.Lines.AccountCode = wipAcct;
                    if (diffDeb > 0)
                    {
                        oJE.Lines.Debit = absDiff;
                    }
                    else
                    {
                        oJE.Lines.Credit = absDiff;
                    }
                    oJE.Lines.Add();

                    // Variance account line
                    oJE.Lines.AccountCode = varAcct;
                    if (diffDeb > 0)
                    {
                        oJE.Lines.Credit = absDiff;
                    }
                    else
                    {
                        oJE.Lines.Debit = absDiff;
                    }
                    oJE.Lines.Add();

                    // Add the JE
                    int result = oJE.Add();

                    if (result == 0)
                    {
                        var docEntryStr = oCompany.GetNewObjectKey();
                        docEntry = Convert.ToInt32(docEntryStr);
                    }
                    else
                    {
                        oCompany.GetLastError(out int errCode, out string errMsg);
                        throw new Exception($"Failed to create JE. Error {errCode}: {errMsg}");
                    }

                }
                else
                {
                    throw new Exception("One or both Journal Entries not found.");
                }
                return docEntry;

            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                if (rsDoc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsDoc);
            }
        }

        public static int GetOriginGRPOEntry(int cancelEntry)
        {
            int originEntry = 0;
            string sql = $@"
            SELECT [OriginalGRPO].DocEntry AS OriginalGRPO_DocEntry
            FROM OPDN AS [CancelGRPO]
            JOIN PDN1 AS [CancelLines] ON [CancelLines].DocEntry = [CancelGRPO].DocEntry
            JOIN OPDN AS [OriginalGRPO] ON [OriginalGRPO].DocEntry = [CancelLines].BaseEntry
            WHERE 
            --[CancelGRPO].Canceled = 'Y' AND 
            [CancelGRPO].DocEntry = {cancelEntry}";


            Recordset rs = null;
            Recordset rsDoc = null;

            try
            {
                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                if (!rs.EoF)
                {
                    string originEntryStr = rs.Fields.Item("OriginalGRPO_DocEntry").Value.ToString();
                    if (int.TryParse(originEntryStr, out int tempEntry))
                    {
                        originEntry = tempEntry;
                    }
                }

                return originEntry;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                if (rsDoc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsDoc);
            }
        }
    }
}
