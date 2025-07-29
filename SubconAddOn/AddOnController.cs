using SAPbouiCOM.Framework;
using SAPbobsCOM;
using System.Runtime.InteropServices;
using System;
using SubconAddOn.Services;
using SubconAddOn.Models;
using System.Globalization;
using System.Collections.Generic;

namespace SubconAddOn
{
    internal static class AddonController
    {
        private static SAPbouiCOM.ProgressBar _pb;
        private static bool _userCanceled = false;

        public static void Start()
        {
            RegisterAppEvents();
            Application.SBO_Application.StatusBar.SetText("Subcon (GI-GR auto generate) add‑on loaded.",
                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            // ➡️  Panggil job background, scheduler, atau listener menu di sini
            // Example: listen ke event menu Production Order Release
            Application.SBO_Application.MenuEvent += SBO_Application_MenuEvent;
            Application.SBO_Application.FormDataEvent += OnFormDataEvent;
            Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
            Application.SBO_Application.ProgressBarEvent += new SAPbouiCOM._IApplicationEvents_ProgressBarEventEventHandler(OnProgressBarEvent);
        }

        private static void RegisterAppEvents()
        {
            Application.SBO_Application.AppEvent += ev =>
            {
                if (ev == SAPbouiCOM.BoAppEventTypes.aet_ShutDown ||
                    ev == SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged ||
                    ev == SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition)
                {
                    CleanExit();
                }
            };
        }
        
        private static void OnProgressBarEvent(ref SAPbouiCOM.ProgressBarEvent pVal,out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.EventType == SAPbouiCOM.BoProgressBarEventTypes.pbet_ProgressBarStopped &&
                pVal.BeforeAction)
                _userCanceled = true;
        }


        private static void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (!pVal.BeforeAction && pVal.MenuUID == "1284")
            {
                try
                {
                    // Tunggu sejenak agar form selesai update
                    System.Threading.Thread.Sleep(500);

                    SAPbouiCOM.Form oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;

                    // Pastikan form adalah GRPO (Form Type 143)
                    if (oForm.TypeEx == "143")
                    {
                        // Cek apakah dokumen sudah dibatalkan
                        SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("OPDN");
                        string docStatus = ds.GetValue("CANCELED", 0).Trim();

                        if (docStatus == "Y")
                        {
                            string docEntryStr = ds.GetValue("DocEntry", 0).Trim();
                            int docEntry = int.Parse(docEntryStr);

                            // 👇 Panggil fungsi setelah berhasil cancel
                            SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Test after cancel.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Error after cancel GRPO: " + ex.Message);
                }
            }
        }


        private static void OnFormDataEvent(ref SAPbouiCOM.BusinessObjectInfo bi, out bool bubbleEvent)
        {
            bubbleEvent = true;
            // GRPO (FormType 143) selesai disimpan
            if (bi.FormTypeEx == "143" &&
                !bi.BeforeAction &&
                bi.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD &&
                bi.ActionSuccess)
            {
                bool success = false;
                bool isCancel = false;
                try
                {
                    Company oCompany = Services.CompanyService.GetCompany();
                    // ——— grab DocEntry ———
                    string raw = bi.ObjectKey;                                // <DocEntry>9</DocEntry>
                    int docEntry = int.Parse(
                        System.Text.RegularExpressions.Regex.Match(raw, @"<DocEntry>(\d+)</DocEntry>")
                             .Groups[1].Value);

                    // ——— load GRPO ———
                    var grpo = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(
                                  SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
                    if (!grpo.GetByKey(docEntry))
                        throw new Exception($"GRPO {docEntry} not found.");

                    // PO Type 4 & bukan cancel
                    string poType = grpo.UserFields.Fields.Item("U_T2_PO_TYPE").Value?.ToString()?.Trim();

                    if (poType == "4" && grpo.Cancelled == SAPbobsCOM.BoYesNoEnum.tYES)
                    {
                        isCancel = true;
                        int originEntry = InventoryService.GetOriginGRPOEntry(docEntry);
                        if (originEntry == 0) throw new Exception("Origin GRPO not found.");

                        _pb = Application.SBO_Application.StatusBar
                              .CreateProgressBar("Create cancellation document…", 2, false);
                        _userCanceled = false;
                        
                        // ── LANGKAH 1: Create reversal JE ──
                        _pb.Value = 1;
                        _pb.Text = "Creating cancellation JE GRPO Subcontract…";
                        if (_userCanceled) throw new OperationCanceledException();

                        int resCancelJE = InventoryService.CancelJEByGRPO(originEntry,docEntry);

                        // ── LANGKAH 2: Cancel GI & GR ──
                        _pb.Value = 2;
                        _pb.Text = "Creating cancellatin GI & GR…";
                        if (_userCanceled) throw new OperationCanceledException();
                        InventoryService.CancelGIGRByGRPO(originEntry, docEntry);
                    }
                    else
                    {
                        if (poType != "4" || grpo.Cancelled == SAPbobsCOM.BoYesNoEnum.tYES) return;

                        // ——— Progress Bar (4 step) ———
                        _pb = Application.SBO_Application.StatusBar
                              .CreateProgressBar("Auto‑generate GI & GR…", 4, false);
                        _userCanceled = false;

                        // ── LANGKAH 1: JE ──
                        _pb.Value = 1;
                        _pb.Text = "Creating JE Subcon…";
                        if (_userCanceled) throw new OperationCanceledException();

                        var entryJe = InventoryService.CreateJESubcon(docEntry);
                        if (entryJe == 0)
                            throw new Exception("JE Subcon fail to create.");

                        // ── LANGKAH 2: LINKED JE TO GRPO──
                        _pb.Value = 2;
                        _pb.Text = "Linked JE Subcon to GRPO…";
                        if (_userCanceled) throw new OperationCanceledException();

                        var resLinkJE = InventoryService.LinkJEToGRPO(docEntry, entryJe);
                        if (!resLinkJE)
                            throw new Exception("JE Subcon fail to link with GRPO.");

                        // ── LANGKAH 3: Goods Issue ──
                        _pb.Value = 3;
                        _pb.Text = "Creating Goods Issue…";
                        if (_userCanceled) throw new OperationCanceledException();

                        var resGi = InventoryService.GetGoodIssueByGRPO(docEntry);
                        int giDocEntry = InventoryService.CreateGoodsIssue(resGi);
                        if (resGi != null && giDocEntry == 0)
                            throw new Exception("Goods Issue fail to create.");

                        // ── LANGKAH 4: Goods Receipt ──
                        _pb.Value = 4;
                        _pb.Text = "Creating Goods Receipt…";
                        if (_userCanceled) throw new OperationCanceledException();

                        var resGr = InventoryService.GetGoodReceiptByGRPO(docEntry, giDocEntry);
                        if (resGr != null && InventoryService.CreateGoodsReceipt(resGr) == 0)
                            throw new Exception("Goods Receipt fail to create.");
                    }
                    
                    success = true;
                }
                catch (OperationCanceledException)
                {
                    new System.Threading.Timer(_ =>
                    {
                        Application.SBO_Application.StatusBar.SetText(
                        "Process cancelled by user.",
                        SAPbouiCOM.BoMessageTime.bmt_Long,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }, null, 1000, System.Threading.Timeout.Infinite);
                }
                catch (Exception ex)
                {
                    new System.Threading.Timer(_ =>
                    {
                        Application.SBO_Application.StatusBar.SetText(
                        "Error Auto GI/GR: " + ex.Message,
                        SAPbouiCOM.BoMessageTime.bmt_Long,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }, null, 1000, System.Threading.Timeout.Infinite);
                }
                finally
                {
                    // ——— always clean up ———
                    _pb?.Stop();
                    if (_pb != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_pb);
                    _pb = null;
                    _userCanceled = false;
                    if (success && !isCancel)
                    {
                        new System.Threading.Timer(_ =>
                        {
                            System.Threading.Thread.Sleep(500); // Delay agar SAP selesai menampilkan pesan
                            Application.SBO_Application.StatusBar.SetText(
                            "Auto GI & GR Success.",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }, null, 1000, System.Threading.Timeout.Infinite);
                    }
                    if (success && isCancel)
                    {
                        new System.Threading.Timer(_ =>
                        {
                            System.Threading.Thread.Sleep(500); // Delay agar SAP selesai menampilkan pesan
                            Application.SBO_Application.StatusBar.SetText(
                            "Successfully cancel GI & GR.",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }, null, 1000, System.Threading.Timeout.Infinite);
                    }
                }
            }
        }

        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.ProgressBar progress = null;

            if (pVal.FormTypeEx == "143" &&
                pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED &&
                pVal.BeforeAction &&
                pVal.ItemUID == "1") // "1" = Add/Update button
            {
                try
                {
                    progress = Application.SBO_Application.StatusBar.CreateProgressBar("Validating BOM stock...", 100, false);

                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                    SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("OPDN");
                    string poType = ds.GetValue("U_T2_PO_TYPE", 0).Trim();
                    string canceled = ds.GetValue("CANCELED", 0).Trim();

                    if (canceled == "N" && poType == "4")
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                        int rowCount = oMatrix.RowCount;

                        for (int i = 1; i <= rowCount; i++)
                        {
                            progress.Value = (int)((i / (double)rowCount) * 100);

                            string itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value;

                            if (InventoryService.IsBom(itemCode))
                            {
                                string quantityStr = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(i).Specific).Value;
                                if (double.TryParse(quantityStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double qty))
                                {
                                    InventoryService.IsStockAvailable(itemCode, qty);
                                }
                            }
                        }
                    }

                    progress.Stop();
                }
                catch (Exception ex)
                {
                    BubbleEvent = false;
                    if (progress != null) progress.Stop();
                    Application.SBO_Application.StatusBar.SetText(ex.Message,
                        SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
            }
        }


        private static void CleanExit()
        {
            Company oCompany = Services.CompanyService.GetCompany();
            if (oCompany != null && oCompany.Connected)
            {
                oCompany.Disconnect();
                Marshal.ReleaseComObject(oCompany);
            }
            Application.SBO_Application.StatusBar.SetText("Subcon (GI-GR auto generate) add‑on unloaded.",
                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            System.Environment.Exit(0);
        }
    }
}
