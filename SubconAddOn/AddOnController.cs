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
using SubconAddOn.Helpers;

namespace SubconAddOn
{
    internal static class AddonController
    {
        private static bool _isCancelTrans = false;
        private static int _docEntry = 0;
        private static SqlLockHelper _lockHelper;
        private const string LOCK_KEY = "GRPO_SUBCON";

        public static void Start()
        {
            RegisterAppEvents();
            Application.SBO_Application.StatusBar.SetText(
            "Subcontract Add-On has been loaded.",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            // ➡️  Panggil job background, scheduler, atau listener menu di sini
            // Example: listen ke event menu Production Order Release
            Application.SBO_Application.MenuEvent += SBO_Application_MenuEvent;
            Application.SBO_Application.FormDataEvent += SBO_Application_FormDataEvent;
            Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
            //oCompany = Services.CompanyService.GetCompany();
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

        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.FormTypeEx == "143")
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && !pVal.BeforeAction)
                {
                    if (_lockHelper != null)
                    {
                        try
                        {
                            _lockHelper.ReleaseLock();
                        }
                        catch (Exception ex)
                        {
                            Application.SBO_Application.StatusBar.SetText(
                                $"Error releasing lock: {ex.Message}",
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Error
                            );
                        }
                        finally
                        {
                            _lockHelper = null;
                        }
                    }
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
                
                try
                {
                    var oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                    
                    _lockHelper = new SqlLockHelper(oCompany);

                    bool gotLock = _lockHelper.AcquireLockAsync(LOCK_KEY, Environment.UserName).Result; // blocking here is fine
                    if (!gotLock)
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            "Another user is processing subcontract Goods Receipt PO. Please wait.",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Warning
                        );
                        bubbleEvent = false;
                    }
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.StatusBar.SetText(
                        $"Error acquiring lock: {ex.Message}",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Error
                    );
                    bubbleEvent = false;
                }

                if (canceled == "N" && poType == "4")
                {
                    ValidateBomStock(form);
                }
                else if (canceled == "C" && poType == "4")
                {
                    ValidateCancellationDocuments(form);
                }
                //throw new Exception("TEST");
            }
            catch (Exception ex)
            {
                bubbleEvent = false;
                ShowStatusDelayed(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                
                if (_lockHelper != null)
                {
                    try
                    {
                        _lockHelper.ReleaseLock();
                    }
                    catch (Exception exc)
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            $"Error releasing lock: {exc.Message}",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error
                        );
                    }
                    finally
                    {
                        _lockHelper = null;
                    }
                }
            }
        }
        
        private static void ValidateBomStock(SAPbouiCOM.Form oForm)
        {
            Company oCompany = Services.CompanyService.GetCompany();
            
            try
            {
                FormHelper.StartLoading(oForm, "Validating BOM stock...", 100, false);
                System.Threading.Thread.Sleep(1000); // Delay 2 seconds

                var matrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                int rowCount = matrix.RowCount;
                List<string> err = new List<string>();

                for (int i = 1; i <= rowCount; i++)
                {
                    var progVal = (int)((i / (double)rowCount) * 100);
                    FormHelper.SetTextValueLoading(oForm, progVal);

                    string itemCode = ((SAPbouiCOM.EditText)matrix.Columns.Item("1").Cells.Item(i).Specific).Value;
                    if (InventoryService.IsBom(oCompany, itemCode))
                    {
                        string qtyStr = ((SAPbouiCOM.EditText)matrix.Columns.Item("11").Cells.Item(i).Specific).Value;
                        if (double.TryParse(qtyStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double qty))
                        {
                            var errItem = InventoryService.IsStockAvailable(oCompany, itemCode, qty);
                            if (errItem.Any())
                            {
                                err.AddRange(errItem);
                            }
                        }
                    }
                }

                if (err.Any())
                {
                    string message;

                    if (err.Count == 1)
                    {
                        // Single item
                        message = $"Item not available in stock:\n{err[0]}";
                    }
                    else
                    {
                        // Multiple items
                        string itemsList = string.Join(Environment.NewLine, err);
                        message = $"These items are not available in stock:\n{itemsList}";
                    }

                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(message);
                    throw new Exception("An error occurred while validating BOM item stock.");
                }
            }catch(Exception e)
            {
                throw e;
            }
            finally
            {
                FormHelper.FinishLoading(oForm);
            }
        }

        private static void ValidateCancellationDocuments(SAPbouiCOM.Form oForm)
        {
            Company oCompany = null;
            try
            {
                if (_isCancelTrans && _docEntry != 0)
                {
                    FormHelper.StartLoading(oForm, "Validating cancellation document…", 3, false);
                    if (oCompany == null) oCompany = Services.CompanyService.GetCompany();
                    if (!oCompany.InTransaction) oCompany.StartTransaction();

                    FormHelper.SetTextValueLoading(oForm, 1, "Validating Stock Cancellation...");
                    var err = InventoryService.IsStockAvailableCancel(oCompany, _docEntry);
                    if (err.Any())
                    {
                        string message;

                        if (err.Count == 1)
                        {
                            // Single item
                            message = $"Item not available in stock:\n{err[0]}";
                        }
                        else
                        {
                            // Multiple items
                            string itemsList = string.Join(Environment.NewLine, err);
                            message = $"These items are not available in stock:\n{itemsList}";
                        }

                        SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(message);
                        throw new Exception("An error occurred while validating BOM item stock.");
                    }

                    FormHelper.SetTextValueLoading(oForm, 2, "Validating Jountry Entry Cancellation...");
                    InventoryService.CancelJEByGRPOTemp(oCompany, _docEntry);

                    FormHelper.SetTextValueLoading(oForm, 3, "Validating Goods Issue & Goods Receipt Cancellation...");
                    InventoryService.CancelGIGRByGRPOTemp(oCompany, _docEntry);
                }
            }
            catch
            {
                if (_lockHelper != null)
                {
                    try
                    {
                        _lockHelper.ReleaseLock();
                    }
                    catch (Exception ex)
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            $"Error releasing lock: {ex.Message}",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error
                        );
                    }
                    finally
                    {
                        _lockHelper = null;
                    }
                }
                throw;
            }
            finally
            {
                if (oCompany.InTransaction)
                    oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                FormHelper.FinishLoading(oForm);
            }
        }
        
        private static void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo bi, out bool bubbleEvent)
        {
            bubbleEvent = true;

            string type = bi.Type;
            string formUID = Application.SBO_Application.Forms.ActiveForm.UniqueID;
            if (bi.FormTypeEx == "143" && bi.BeforeAction && bi.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            {
                if (type == "20")
                {
                    HandleAddOrCancelPressed(formUID, ref bubbleEvent);
                }
            }

            if (bi.FormTypeEx == "143" && !bi.BeforeAction && bi.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD && bi.ActionSuccess)
            {
                if (type == "20")
                {
                    bool success = false;
                    bool isCancel = false;
                    bool isCreate = false;
                    int docEntry = 0;
                    int docNum = 0;
                    Company oCompany = Services.CompanyService.GetCompany();
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;

                    try
                    {
                        //EnsureCompanyConnected();

                        if (!oCompany.InTransaction)
                            oCompany.StartTransaction();

                        docEntry = ExtractDocEntry(bi.ObjectKey);
                        var grpo = LoadGRPO(oCompany, docEntry);
                        docNum = grpo.DocNum;

                        string poType = grpo.UserFields.Fields.Item("U_T2_PO_TYPE").Value?.ToString()?.Trim();

                        if (poType == "4" && grpo.Cancelled == SAPbobsCOM.BoYesNoEnum.tYES)
                        {
                            isCancel = true;
                            ProcessCancellation(oForm, oCompany, docEntry);
                        }
                        else if (poType == "4" && grpo.Cancelled == SAPbobsCOM.BoYesNoEnum.tNO)
                        {
                            isCreate = true;
                            ProcessAutoGeneration(oForm, oCompany, docEntry);
                        }

                        success = true;
                        if (oCompany.InTransaction)
                            oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                    }
                    catch (OperationCanceledException)
                    {
                        if (oCompany?.InTransaction == true)
                            oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                        ShowStatusDelayed("Process cancelled by user.", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }
                    catch (Exception ex)
                    {
                        if (oCompany?.InTransaction == true)
                            oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                        HandleAutoGenError(oCompany, ex, docEntry, docNum, isCancel);
                    }
                    finally
                    {
                        if (success && (isCreate || isCancel))
                        {
                            var message = (isCancel) ? "Goods Issue and Goods Receipt were successfully canceled." : "Auto-generation of Goods Issue and Goods Receipt completed successfully.";
                            ShowStatusDelayed(message, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }

                        if (_lockHelper != null)
                        {
                            try
                            {
                                _lockHelper.ReleaseLock();
                            }
                            catch (Exception ex)
                            {
                                Application.SBO_Application.StatusBar.SetText(
                                    $"Error releasing lock: {ex.Message}",
                                    SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                                );
                            }
                            finally
                            {
                                _lockHelper = null;
                            }
                        }
                        _docEntry = 0;
                        _isCancelTrans = false;
                    }
                }
            }
        }

        private static int ExtractDocEntry(string objectKey)
        {
            var match = Regex.Match(objectKey, @"<DocEntry>(\d+)</DocEntry>");
            return int.Parse(match.Groups[1].Value);
        }

        private static SAPbobsCOM.Documents LoadGRPO(Company oCompany, int docEntry)
        {
            var grpo = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
            if (!grpo.GetByKey(docEntry))
                throw new Exception($"Goods Receipt PO {docEntry} not found.");
            return grpo;
        }

        private static void ProcessCancellation(SAPbouiCOM.Form oForm, Company oCompany, int docEntry)
        {
            try
            {
                FormHelper.StartLoading(oForm, "Create cancellation document…", 2, false);
                FormHelper.SetTextValueLoading(oForm, 1, "Creating cancellation Journal Entry Goods Receipt PO Subcontract…");
                
                int originEntry = InventoryService.GetOriginGRPOEntry(oCompany, docEntry);
                if (originEntry == 0)
                    throw new Exception("Origin Goods Receipt PO not found.");
                
                InventoryService.DeleteAllRefGRPO(oCompany, docEntry);
                InventoryService.CancelJEByGRPO(oCompany, originEntry, docEntry);

                FormHelper.SetTextValueLoading(oForm, 2, "Creating cancellation Goods Issue & Goods Receipt…");
                
                InventoryService.CancelGIGRByGRPO(oCompany, originEntry, docEntry);
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                FormHelper.FinishLoading(oForm);
            }
        }

        private static void ProcessAutoGeneration(SAPbouiCOM.Form oForm, Company oCompany, int docEntry)
        {
            try
            {
                FormHelper.StartLoading(oForm, "Starting auto‑generate document…", 4, false);
                // Step 1: JE
                FormHelper.SetTextValueLoading(oForm, 1, "Creating Journal Entry Subcontract…");
                int entryJe = InventoryService.CreateJESubcon(oCompany, docEntry);
                if (entryJe == 0) throw new Exception("Journal Entry Subcon fail to create.");

                // Step 2: Link JE
                FormHelper.SetTextValueLoading(oForm, 2, "Linking Journal Entry Subcontract to Goods Receipt PO…");

                if (!InventoryService.LinkJEToGRPO(oCompany, docEntry, entryJe))
                    throw new Exception("Journal Entry Subcontract fail to link with Goods Receipt PO.");

                // Step 3: Goods Issue
                FormHelper.SetTextValueLoading(oForm, 3, "Creating Goods Issue…");

                var resGi = InventoryService.GetGoodIssueByGRPO(oCompany, docEntry);
                int giDocEntry = InventoryService.CreateGoodsIssue(oCompany, resGi);
                if (resGi != null && giDocEntry == 0)
                    throw new Exception("Goods Issue fail to create.");
                if (!InventoryService.LinkGIToGRPO(oCompany, docEntry, giDocEntry))
                    throw new Exception("Goods Issue fail to link with Goods Receipt PO.");

                // Step 4: Goods Receipt
                FormHelper.SetTextValueLoading(oForm, 4, "Creating Goods Receipt…");

                var resGr = InventoryService.GetGoodReceiptByGRPO(oCompany, docEntry, giDocEntry);
                int grDocEntry = InventoryService.CreateGoodsReceipt(oCompany, resGr);
                if (resGr != null && grDocEntry == 0)
                    throw new Exception("Goods Receipt fail to create.");
                if (!InventoryService.LinkGRToGRPO(oCompany, docEntry, grDocEntry))
                    throw new Exception("Goods Receipt fail to link with Goods Receipt PO.");
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                FormHelper.FinishLoading(oForm);
            }
        }

        private static void HandleAutoGenError(Company oCompany, Exception ex, int docEntry, int docNum, bool isCancel)
        {
            ShowStatusDelayed("Auto-generation of documents failed: " + ex.Message,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error);

            if (!isCancel)
            {
                InventoryService.CancelGoodsReceiptPO(oCompany, docEntry);
                Application.SBO_Application.MessageBox(
                    $"Auto-generation of documents failed.\n\nError:\n{ex.Message}\n\nThe Goods Receipt PO ({docNum}) has been canceled.",
                    1, "OK", "", "");
            }
            else
            {
                Application.SBO_Application.MessageBox(
                    "Auto-cancellation of documents failed.\nPlease cancel them manually.",
                    1, "OK", "", "");
            }
        }

        private static void ShowStatusDelayed(string text, SAPbouiCOM.BoStatusBarMessageType type)
        {
            new System.Threading.Timer(_ =>
            {
                Application.SBO_Application.StatusBar.SetText(text, SAPbouiCOM.BoMessageTime.bmt_Short, type);
            }, null, 500, System.Threading.Timeout.Infinite);
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
            Application.SBO_Application.StatusBar.SetText(
            "Subcontract Add-On has been unloaded.",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            System.Environment.Exit(0);
        }
    }
}
