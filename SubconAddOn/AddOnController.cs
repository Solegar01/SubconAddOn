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

        public static void Start()
        {
            _diCmp = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

            RegisterAppEvents();
            Application.SBO_Application.StatusBar.SetText("Headless add‑on loaded.",
                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            // ➡️  Panggil job background, scheduler, atau listener menu di sini
            // Example: listen ke event menu Production Order Release
            Application.SBO_Application.MenuEvent += OnMenuEvent;
            Application.SBO_Application.FormDataEvent += OnFormDataEvent;
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

        private static void OnMenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool bubble)
        {
            bubble = true;
            if (!pVal.BeforeAction && pVal.MenuUID == "2306")   // contoh: Production Order
            {
                

                Application.SBO_Application.StatusBar.SetText("GRPO menu clicked (headless add‑on).",
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                // Tambahkan logika bisnis di sini.
            }
        }

        private static void OnFormDataEvent(ref SAPbouiCOM.BusinessObjectInfo businessObjectInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;

            if (businessObjectInfo.FormTypeEx == "143" &&  // GRPO
                businessObjectInfo.BeforeAction == false &&
                businessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD &&
                businessObjectInfo.ActionSuccess)
            {
                try
                {
                    string raw = businessObjectInfo.ObjectKey;

                    // Ambil isi <DocEntry>9</DocEntry>
                    string docEntryStr = System.Text.RegularExpressions.Regex
                        .Match(raw, @"<DocEntry>(\d+)</DocEntry>")
                        .Groups[1].Value;

                    int docEntry = int.Parse(docEntryStr);
                    // Ambil GRPO dari DI API
                    var grpo = (SAPbobsCOM.Documents)_diCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);

                    if (!grpo.GetByKey(docEntry))
                        throw new Exception($"GRPO with DocEntry {docEntry} not found.");

                    string poType = grpo.UserFields.Fields.Item("U_T2_PO_TYPE").Value?.ToString()?.Trim();

                    // PO Type = "4" && BUKAN GRPO cancel
                    if (poType == "4" && grpo.Cancelled == SAPbobsCOM.BoYesNoEnum.tNO)
                    {
                        // Panggil layanan untuk create Good Issue berdasarkan GRPO
                        var resGi = InventoryService.GetGoodIssueByGRPO(docEntry); // Anda bisa cek null/error kalau perlu
                        if (resGi != null)
                        {
                            int giSuccess = InventoryService.CreateGoodsIssue(_diCmp, resGi);
                            if (giSuccess == 0)
                            {
                                throw new Exception($"There is no GI created.");
                            }
                        }
                        var resGr = InventoryService.GetGoodReceiptByGRPO(docEntry);
                        if (resGr != null)
                        {
                            int grSuccess = InventoryService.CreateGoodsReceipt(_diCmp, resGr);
                            if (grSuccess == 0)
                            {
                                throw new Exception($"There is no GR created.");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.StatusBar.SetText("Error in GRPO after-save event: " + ex.Message,
                        SAPbouiCOM.BoMessageTime.bmt_Long,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
            Application.SBO_Application.StatusBar.SetText("Headless add‑on unloaded.",
                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            System.Environment.Exit(0);
        }
    }
}
