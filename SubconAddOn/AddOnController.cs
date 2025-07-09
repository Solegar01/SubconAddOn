using SAPbouiCOM.Framework;
using SAPbobsCOM;
using System.Runtime.InteropServices;
using System;
using SubconAddOn.Services;

namespace SubconAddOn
{
    internal static class AddonController
    {
        private static SAPbobsCOM.Company _diCmp;
        private static SAPbouiCOM.ProgressBar _pb;
        private static bool _userCanceled = false;


        public static void Start()
        {
            _diCmp = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

            RegisterAppEvents();
            Application.SBO_Application.StatusBar.SetText("Subcon (GI-GR auto generate) add‑on loaded.",
                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            // ➡️  Panggil job background, scheduler, atau listener menu di sini
            // Example: listen ke event menu Production Order Release
            Application.SBO_Application.MenuEvent += OnMenuEvent;
            Application.SBO_Application.FormDataEvent += OnFormDataEvent;
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


        private static void OnMenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool bubble)
        {
            bubble = true;

            if (!pVal.BeforeAction && pVal.MenuUID == "2306")   // contoh: Production Order
            {
                
                //Application.SBO_Application.StatusBar.SetText("GRPO menu clicked (headless add‑on).",
                //    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                // Tambahkan logika bisnis di sini.
            }
        }

        private static void OnFormDataEvent(ref SAPbouiCOM.BusinessObjectInfo bi,out bool bubbleEvent)
        {
            bubbleEvent = true;

            // GRPO (FormType 143) selesai disimpan
            if (bi.FormTypeEx == "143" &&
                !bi.BeforeAction &&
                bi.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD &&
                bi.ActionSuccess)
            {
                try
                {
                    // ——— grab DocEntry ———
                    string raw = bi.ObjectKey;                                // <DocEntry>9</DocEntry>
                    int docEntry = int.Parse(
                        System.Text.RegularExpressions.Regex.Match(raw, @"<DocEntry>(\d+)</DocEntry>")
                             .Groups[1].Value);

                    // ——— load GRPO ———
                    var grpo = (SAPbobsCOM.Documents)_diCmp.GetBusinessObject(
                                  SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
                    if (!grpo.GetByKey(docEntry))
                        throw new Exception($"GRPO {docEntry} not found.");

                    // PO Type 4 & bukan cancel
                    string poType = grpo.UserFields.Fields.Item("U_T2_PO_TYPE").Value?.ToString()?.Trim();
                    if (poType != "4" || grpo.Cancelled == SAPbobsCOM.BoYesNoEnum.tYES) return;

                    // ——— Progress Bar (2 step) ———
                    _pb = Application.SBO_Application.StatusBar
                          .CreateProgressBar("Auto‑generate GI & GR…", 2, true);
                    _userCanceled = false;

                    // ── LANGKAH 1: Goods Issue ──
                    _pb.Value = 1;
                    _pb.Text = "Creating Goods Issue…";
                    if (_userCanceled) throw new OperationCanceledException();

                    var resGi = InventoryService.GetGoodIssueByGRPO(docEntry);
                    if (resGi != null && InventoryService.CreateGoodsIssue(_diCmp, resGi) == 0)
                        throw new Exception("Goods Issue fail to create.");

                    // ── LANGKAH 2: Goods Receipt ──
                    _pb.Value = 2;
                    _pb.Text = "Creating Goods Receipt…";
                    if (_userCanceled) throw new OperationCanceledException();

                    var resGr = InventoryService.GetGoodReceiptByGRPO(docEntry);
                    if (resGr != null && InventoryService.CreateGoodsReceipt(_diCmp, resGr) == 0)
                        throw new Exception("Goods Receipt fail to create.");

                    Application.SBO_Application.StatusBar.SetText(
                        "Auto GI & GR Success.",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                catch (OperationCanceledException)
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Process cancelled by user.",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Error Auto GI/GR: " + ex.Message,
                        SAPbouiCOM.BoMessageTime.bmt_Long,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                finally
                {
                    // ——— always clean up ———
                    _pb?.Stop();
                    if (_pb != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_pb);
                    _pb = null;
                    _userCanceled = false;
                }
            }
        }
        
        private static void CleanExit()
        {
            if (_diCmp != null && _diCmp.Connected)
            {
                _diCmp.Disconnect();
                Marshal.ReleaseComObject(_diCmp);
            }
            Application.SBO_Application.StatusBar.SetText("Subcon (GI-GR auto generate) add‑on unloaded.",
                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            System.Environment.Exit(0);
        }
    }
}
