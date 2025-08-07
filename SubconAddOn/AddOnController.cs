using SAPbouiCOM.Framework;
using SAPbobsCOM;
using System.Runtime.InteropServices;
using System;
using SubconAddOn.Services;
using SubconAddOn.Models;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;
using System.Text.RegularExpressions;

namespace SubconAddOn
{
    internal static class AddonController
    {
        private static SAPbouiCOM.ProgressBar _pb;
        private static bool _userCanceled = false;
        private static Company oCompany = null;
        private static Mutex mutex = new Mutex(false, "Global\\SAPB1_GRPO_Mutex");
        private static bool mutexAcquired = false; // status lokal, karena FormDataEvent tidak tahu siapa yang ambil mutex
        private static bool _isCancelTrans = false;
        private static int _docEntry = 0;
        
        public static void Start()
        {
            RegisterAppEvents();
            Application.SBO_Application.StatusBar.SetText(
            "Subcon Add-on (GI-GR auto-generation) has been loaded.",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            // ➡️  Panggil job background, scheduler, atau listener menu di sini
            // Example: listen ke event menu Production Order Release
            Application.SBO_Application.MenuEvent += SBO_Application_MenuEvent;
            Application.SBO_Application.FormDataEvent += SBO_Application_FormDataEvent;
            Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
            Application.SBO_Application.ProgressBarEvent += new SAPbouiCOM._IApplicationEvents_ProgressBarEventEventHandler(OnProgressBarEvent);
            oCompany = Services.CompanyService.GetCompany();
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

        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.FormTypeEx == "143")
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && !pVal.BeforeAction)
                {
                    ReleaseMutexIfAcquired();
                }
            }
        }

        private static void HandleAddOrCancelPressed(string formUID, ref bool bubbleEvent)
        {
            try
            {
                var form = Application.SBO_Application.Forms.Item(formUID);
                var ds = form.DataSources.DBDataSources.Item("OPDN");

                string poType = ds.GetValue("U_T2_PO_TYPE", 0).Trim();
                string canceled = ds.GetValue("CANCELED", 0).Trim();

                if (!TryAcquireMutex(ref bubbleEvent)) return;

                if (canceled == "N" && poType == "4")
                {
                    ValidateBomStock(form);
                }
                else if (canceled == "C" && poType == "4")
                {
                    ValidateCancellationDocuments();
                }
                //throw new Exception("TEST");
            }
            catch (Exception ex)
            {
                bubbleEvent = false;
                ShowStatusDelayed(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                ReleaseMutexIfAcquired();
            }
            finally
            {
                CleanupProgressBar();
                _userCanceled = false;
            }
        }

        private static bool TryAcquireMutex(ref bool bubbleEvent)
        {
            try { mutex.ReleaseMutex(); mutexAcquired = false; } catch { }

            if (mutex.WaitOne(0))
            {
                mutexAcquired = true;
                return true;
            }
            else
            {
                CleanupProgressBar();
                ShowStatusDelayed("Another GRPO process is running. Please wait...", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                bubbleEvent = false;
                return false;
            }
        }

        private static void ValidateBomStock(SAPbouiCOM.Form form)
        {
            _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Validating BOM stock...", 100, false);
            _pb.Text = "Validating BOM stock...";

            var matrix = (SAPbouiCOM.Matrix)form.Items.Item("38").Specific;
            int rowCount = matrix.RowCount;

            for (int i = 1; i <= rowCount; i++)
            {
                _pb.Value = (int)((i / (double)rowCount) * 100);

                string itemCode = ((SAPbouiCOM.EditText)matrix.Columns.Item("1").Cells.Item(i).Specific).Value;
                if (InventoryService.IsBom(oCompany, itemCode))
                {
                    string qtyStr = ((SAPbouiCOM.EditText)matrix.Columns.Item("11").Cells.Item(i).Specific).Value;
                    if (double.TryParse(qtyStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double qty))
                    {
                        InventoryService.IsStockAvailable(oCompany, itemCode, qty);
                    }
                }
            }
        }

        private static void ValidateCancellationDocuments()
        {
            _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Validating cancellation document…", 3, false);
            _pb.Text = "Validating cancellation document…";

            if (_isCancelTrans && _docEntry != 0)
            {
                try
                {
                    if (oCompany == null) oCompany = Services.CompanyService.GetCompany();
                    if (!oCompany.InTransaction) oCompany.StartTransaction();

                    _pb.Text = "Validating Stock Cancellation...";
                    _pb.Value = 1;
                    InventoryService.IsStockAvailableCancel(oCompany, _docEntry);

                    _pb.Text = "Validating JE Cancellation...";
                    _pb.Value = 2;
                    InventoryService.CancelJEByGRPOTemp(oCompany, _docEntry);

                    _pb.Text = "Validating GI & GR Cancellation...";
                    _pb.Value = 3;
                    InventoryService.CancelGIGRByGRPOTemp(oCompany, _docEntry);
                }
                catch
                {
                    ReleaseMutexIfAcquired();
                    throw;
                }
                finally
                {
                    if (oCompany.InTransaction)
                        oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
            }
        }

        private static void CleanupProgressBar()
        {
            _pb?.Stop();
            if (_pb != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_pb);
            _pb = null;
        }

        private static void ReleaseMutexIfAcquired()
        {
            if (mutexAcquired)
            {
                try { mutex.ReleaseMutex(); } catch { }
                mutexAcquired = false;
            }
        }

        private static void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo bi, out bool bubbleEvent)
        {
            bubbleEvent = true;
            string formUID = Application.SBO_Application.Forms.ActiveForm.UniqueID;
            if (bi.FormTypeEx == "143" && bi.BeforeAction && bi.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            {
                HandleAddOrCancelPressed(formUID, ref bubbleEvent);
            }

            if (bi.FormTypeEx == "143" && !bi.BeforeAction && bi.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD && bi.ActionSuccess && mutexAcquired)
            {
                bool success = false;
                bool isCancel = false;
                int docEntry = 0;
                int docNum = 0;

                try
                {
                    EnsureCompanyConnected();

                    if (!oCompany.InTransaction)
                        oCompany.StartTransaction();

                    docEntry = ExtractDocEntry(bi.ObjectKey);
                    var grpo = LoadGRPO(docEntry);
                    docNum = grpo.DocNum;

                    string poType = grpo.UserFields.Fields.Item("U_T2_PO_TYPE").Value?.ToString()?.Trim();

                    if (poType == "4" && grpo.Cancelled == SAPbobsCOM.BoYesNoEnum.tYES)
                    {
                        isCancel = true;
                        ProcessCancellation(docEntry);
                    }
                    else if (poType == "4" && grpo.Cancelled == SAPbobsCOM.BoYesNoEnum.tNO)
                    {
                        ProcessAutoGeneration(docEntry);
                    }

                    success = true;
                    if (oCompany.InTransaction)
                        oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                }
                catch (OperationCanceledException)
                {
                    RollbackTransaction();
                    ShowStatusDelayed("Process cancelled by user.", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
                catch (Exception ex)
                {
                    RollbackTransaction();
                    HandleAutoGenError(ex, docEntry, docNum, isCancel);
                }
                finally
                {
                    CleanupProgressBar();

                    if (success)
                    {
                        var message = isCancel ? "GI and GR were successfully canceled." : "Auto-generation of GI and GR completed successfully.";
                        ShowStatusDelayed(message, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }

                    ReleaseMutexIfAcquired();
                    _docEntry = 0;
                    _isCancelTrans = false;
                }
            }
        }

        private static int ExtractDocEntry(string objectKey)
        {
            var match = Regex.Match(objectKey, @"<DocEntry>(\d+)</DocEntry>");
            return int.Parse(match.Groups[1].Value);
        }

        private static SAPbobsCOM.Documents LoadGRPO(int docEntry)
        {
            var grpo = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
            if (!grpo.GetByKey(docEntry))
                throw new Exception($"GRPO {docEntry} not found.");
            return grpo;
        }

        private static void ProcessCancellation(int docEntry)
        {
            int originEntry = InventoryService.GetOriginGRPOEntry(oCompany, docEntry);
            if (originEntry == 0)
                throw new Exception("Origin GRPO not found.");

            _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Create cancellation document…", 2, false);
            _pb.Value = 1;
            _pb.Text = "Creating cancellation JE GRPO Subcontract…";

            if (_userCanceled) throw new OperationCanceledException();

            InventoryService.DeleteAllRefGRPO(oCompany, docEntry);
            InventoryService.CancelJEByGRPO(oCompany, originEntry, docEntry);

            _pb.Value = 2;
            _pb.Text = "Creating cancellation GI & GR…";

            if (_userCanceled) throw new OperationCanceledException();
            InventoryService.CancelGIGRByGRPO(oCompany, originEntry, docEntry);
        }

        private static void ProcessAutoGeneration(int docEntry)
        {
            _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Auto‑generate GI & GR…", 4, false);

            // Step 1: JE
            _pb.Value = 1;
            _pb.Text = "Creating JE Subcon…";
            if (_userCanceled) throw new OperationCanceledException();
            int entryJe = InventoryService.CreateJESubcon(oCompany, docEntry);
            if (entryJe == 0) throw new Exception("JE Subcon fail to create.");

            // Step 2: Link JE
            _pb.Value = 2;
            _pb.Text = "Linking JE Subcon to GRPO…";
            if (_userCanceled) throw new OperationCanceledException();
            if (!InventoryService.LinkJEToGRPO(oCompany, docEntry, entryJe))
                throw new Exception("JE Subcon fail to link with GRPO.");

            // Step 3: Goods Issue
            _pb.Value = 3;
            _pb.Text = "Creating Goods Issue…";
            if (_userCanceled) throw new OperationCanceledException();
            var resGi = InventoryService.GetGoodIssueByGRPO(oCompany, docEntry);
            int giDocEntry = InventoryService.CreateGoodsIssue(oCompany, resGi);
            if (resGi != null && giDocEntry == 0)
                throw new Exception("Goods Issue fail to create.");
            if (!InventoryService.LinkGIToGRPO(oCompany, docEntry, giDocEntry))
                throw new Exception("Goods Issue fail to link with GRPO.");

            // Step 4: Goods Receipt
            _pb.Value = 4;
            _pb.Text = "Creating Goods Receipt…";
            if (_userCanceled) throw new OperationCanceledException();
            var resGr = InventoryService.GetGoodReceiptByGRPO(oCompany, docEntry, giDocEntry);
            int grDocEntry = InventoryService.CreateGoodsReceipt(oCompany, resGr);
            if (resGr != null && grDocEntry == 0)
                throw new Exception("Goods Receipt fail to create.");
            if (!InventoryService.LinkGRToGRPO(oCompany, docEntry, grDocEntry))
                throw new Exception("Goods Receipt fail to link with GRPO.");
        }

        private static void HandleAutoGenError(Exception ex, int docEntry, int docNum, bool isCancel)
        {
            ShowStatusDelayed("Auto-generation of GI and GR failed: " + ex.Message,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error);

            if (!isCancel)
            {
                InventoryService.CancelGoodsReceiptPO(oCompany, docEntry);
                Application.SBO_Application.MessageBox(
                    $"Auto-generation of GI and GR failed.\n\nError:\n{ex.Message}\n\nThe GRPO ({docNum}) has been canceled.",
                    1, "OK", "", "");
            }
            else
            {
                Application.SBO_Application.MessageBox(
                    "Auto-cancellation of GI and GR failed.\nPlease cancel them manually.",
                    1, "OK", "", "");
            }
        }

        private static void ShowStatusDelayed(string text, SAPbouiCOM.BoStatusBarMessageType type)
        {
            new System.Threading.Timer(_ =>
            {
                Application.SBO_Application.StatusBar.SetText(text, SAPbouiCOM.BoMessageTime.bmt_Long, type);
            }, null, 1000, System.Threading.Timeout.Infinite);
        }

        private static void EnsureCompanyConnected()
        {
            if (oCompany == null || !oCompany.Connected)
                oCompany = Services.CompanyService.GetCompany();
        }

        private static void RollbackTransaction()
        {
            if (oCompany?.InTransaction == true)
                oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
        }

        private static void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.BeforeAction && pVal.MenuUID == "1284")
            {
                try
                {
                    var form = Application.SBO_Application.Forms.ActiveForm;
                    if (form.TypeEx != "143") return;

                    var ds = form.DataSources.DBDataSources.Item("OPDN");
                    _docEntry = int.Parse(ds.GetValue("DocEntry", 0).Trim());
                    _isCancelTrans = true;
                }
                catch (Exception ex)
                {
                    BubbleEvent = false;
                    ShowStatusDelayed(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
            }
        }
        
        private static void CleanExit()
        {
            //oCompany = Services.CompanyService.GetCompany();
            if (oCompany != null && oCompany.Connected)
            {
                oCompany.Disconnect();
                Marshal.ReleaseComObject(oCompany);
            }
            Application.SBO_Application.StatusBar.SetText(
            "Subcon Add-on (GI-GR auto-generation) has been unloaded.",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            System.Environment.Exit(0);
        }
    }
}
