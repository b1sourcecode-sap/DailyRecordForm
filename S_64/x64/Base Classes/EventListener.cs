using System;
using SAPbobsCOM;
using SAPbouiCOM;
using System.IO;
using System.Data;
using System.Collections;
using System.Windows.Forms;

namespace LGBOM_LaxmiMotors
{
    /// <summary>
    /// public sealed class EventListener
    /// </summary>
    public sealed class EventListener
    {
        #region FORM VARIABLES DECLARATION
        private static EventListener    _Listener;
        private SAPbouiCOM.Application  _Application;
        private SAPbobsCOM.Company      _Company;
        private Base                    _Object;
        private Hashtable               _Collection;
        private Hashtable               _LookCollection;
        private string                  _FormUID;
        #endregion

        #region CONSTRUCTOR

        private EventListener()
        {
            this.Connect();

            this._Application.MenuEvent       += new _IApplicationEvents_MenuEventEventHandler(SboApplication_MenuEvent);
            this._Application.ItemEvent       += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SboApplication_ItemEvent);
            this._Application.AppEvent        += new _IApplicationEvents_AppEventEventHandler(SboApplication_AppEvent);
            this._Application.RightClickEvent += new _IApplicationEvents_RightClickEventEventHandler(SboApplication_RightClickEvent);
            this._Application.FormDataEvent   += new _IApplicationEvents_FormDataEventEventHandler(SboApplication_FormDataEvent);
            
            

            this._Collection                      = new Hashtable(10, (float)0.5);
            this._LookCollection                  = new Hashtable(10, (float)0.5);
        }

        public static EventListener getEventListener()
        {
            if (_Listener == null)
                _Listener = new EventListener();


            return _Listener;
        }

        #endregion

        #region PROPERTIES

        public SAPbobsCOM.Company Company
        {
            get { return this._Company; }
        }

        public SAPbouiCOM.Application SBO_Application
        {
            get { return this._Application; }
        }

        public Hashtable Collection
        {
            get { return this._Collection; }
        }

        public Hashtable LookUpCollection
        {
            get { return this._LookCollection; }
        }

        #endregion

        #region COMPANY CONNECT

        private void Connect()
        {
            this.SetApplication();
            this.ConnectToCompany();
        }

        private void SetApplication()
        {
            SAPbouiCOM.SboGuiApi oSboGuiApi;
            string               oConnectionString;
            try
            {
                if (Environment.GetCommandLineArgs().Length > 0)
                {
                    //oSboGuiApi = new SAPbouiCOM.SboGuiApiClass();
                    oSboGuiApi = new SAPbouiCOM.SboGuiApi();
                    oConnectionString = Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
                    oSboGuiApi.Connect(oConnectionString);
                    this._Application = oSboGuiApi.GetApplication(-1);
                    //this._Company = new SAPbobsCOM.CompanyClass();
                    this._Company = new SAPbobsCOM.Company();
                }
                else
                {
                    throw new Exception("Connection string missing.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void ConnectToCompany()
        {
            string oCookie, oConnectionContext;
            try
            {
                oCookie = this._Company.GetContextCookie();
                oConnectionContext = this._Application.Company.GetConnectionContext(oCookie);
                this._Company.SetSboLoginContext(oConnectionContext);

                if (this._Company.Connect() != 0)
                {
                    _Application.StatusBar.SetText("Add-On is not Connected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception(_Company.GetLastErrorDescription());
                }
                else
                {
                    _Application.StatusBar.SetText("Add-On is Connecting...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

        #region EVENT HANDLERS

        #region MENU EVENT

        private void SboApplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (!pVal.BeforeAction)
                {

                    switch (pVal.MenuUID)
                    {

                        case Constants.User_Menus.Menu_MaterialManagement:
                            this._Object = new ClsMaterialManagement();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.MaterialManagement);
                            this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                            ((ClsMaterialManagement)_Object).FormDefault();
                            break;

                        case Constants.User_Menus.MENU_Production:
                            this._Object = new ClsProduction();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.Production);
                            this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                            ((ClsProduction)_Object).FormDefault();
                            break;

                        case Constants.User_Menus.MENU_MainAndEngg:
                            this._Object = new CLsMainEngg();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.MainAndEngg);
                            this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                            ((CLsMainEngg)_Object).FormDefault();
                            break;

                        case Constants.User_Menus.MENU_QualityControl:
                            this._Object = new ClsQualityControl();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.QualityControl);
                            this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                            ((ClsQualityControl)_Object).FormDefault();
                            break;

                        case Constants.User_Menus.MENU_QualityAssurance:
                            this._Object = new ClsQualityAssurance();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.QualityAssurance);
                            this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                            ((ClsQualityAssurance)_Object).FormDefault();
                            break;

                        case Constants.User_Menus.MENU_AdminAndOperation:
                            this._Object = new CLsAdmin();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.AdminAndOperation);
                            this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                            ((CLsAdmin)_Object).FormDefault();
                            break;
                            //call UDO from custom menu
                        case Constants.User_Menus.Menu_MaterialManagementMaster:
                            Utilities.Application.SBO_Application.Menus.Item("MaterialManagement").Activate();
                            break;
                        case Constants.User_Menus.MENU_ProductionMaster:
                            Utilities.Application.SBO_Application.Menus.Item("ProductionMaster").Activate();
                            break;
                        case Constants.User_Menus.MENU_QualityControlMaster:
                            Utilities.Application.SBO_Application.Menus.Item("QualityControl").Activate();
                            break;
                        case Constants.User_Menus.MENU_AdminAndOperationMaster:
                            Utilities.Application.SBO_Application.Menus.Item("AdminMaster").Activate();
                            break;
                        case Constants.User_Menus.MENU_MainAndEnggMaster:
                            Utilities.Application.SBO_Application.Menus.Item("Main&Engg").Activate();
                            break;
                        case Constants.User_Menus.MENU_QualityAssuranceMaster:
                            Utilities.Application.SBO_Application.Menus.Item("QualityAssurance").Activate();
                            break;

                        /* case Constants.User_Menus.MENU_LGBOM:
                             this._Object = new ClsLaxmiGroupBOM();
                             Utilities.LoadForm(ref this._Object, Constants.Forms.LGBOM);
                             this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                             ((ClsLaxmiGroupBOM)_Object).FormDefault();
                             break;

                         case Constants.User_Menus.MENU_ECO:
                             this._Object = new ClsLaxmiGroupECO();
                             Utilities.LoadForm(ref this._Object, Constants.Forms.ECO);
                             this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                             ((ClsLaxmiGroupECO)_Object).FormDefault();
                             break;*/
                        /* case Constants.User_Menus.MENU_LGProduction:
                              this._Object = new ClsLGProductionOrder();
                              Utilities.LoadForm(ref this._Object, Constants.Forms.ProductionPlanning);
                              this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                              ((ClsLGProductionOrder)_Object).FormDefault();
                              break;
                          case Constants.User_Menus.MENU_LGApproval:
                              this._Object = new ClsApprovalTemplate();
                              Utilities.LoadForm(ref this._Object, Constants.Forms.ApprovalTemplate);
                              this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                              ((ClsApprovalTemplate)_Object).FormDefault();
                              break;*/
                        // System Menu's
                        case Constants.System_Menus.mnu_ADD:
                        case Constants.System_Menus.mnu_FIND:
                        case Constants.System_Menus.mnu_FIRST:
                        case Constants.System_Menus.mnu_LAST:
                        case Constants.System_Menus.mnu_NEXT:
                        case Constants.System_Menus.mnu_PREVIOUS:
                        case Constants.System_Menus.mnu_REFRESH:
                        case Constants.System_Menus.mnu_ADD_ROW:
                        case Constants.System_Menus.mnu_DELETE_ROW:
                        case Constants.System_Menus.mnu_SALES_ORDER:
                        case Constants.System_Menus.mnu_OUTGOING_PAYMENTS:
                        case Constants.System_Menus.mnu_GL_ACCOUNT_DETERMINATION:
                        case Constants.System_Menus.mnu_Duplicate:

                            if (this._Collection.Contains(this._FormUID))
                            {
                                this._Object = (Base)this._Collection[this._FormUID];
                                ((Base)this._Object).Menu_Event(ref pVal, ref BubbleEvent);
                            }
                            break;
                    }
                }
                else
                {
                    switch (pVal.MenuUID)
                    {
                        //case Constants.System_Menus.mnu_GROSS_PROFIT:
                        //    BubbleEvent = false;
                        //    break;
                        case Constants.System_Menus.mnu_ROW_DETAILS:
                            BubbleEvent = false;
                            break;

                        case Constants.System_Menus.mnu_DELETE_ROW:
                        case Constants.System_Menus.mnu_Duplicate:
                            if (this._Collection.Contains(this._FormUID))
                            {
                                this._Object = (Base)this._Collection[this._FormUID];
                                ((Base)this._Object).Menu_Event(ref pVal, ref BubbleEvent);
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region ITEM EVENT

        private void SboApplication_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                this._FormUID = FormUID;
               
                //Utilities.HideSystemMessege(ref pVal);
                if (!pVal.BeforeAction && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD) //
                {
                    switch (pVal.FormType)
                    {
                        //case Constants.System_Forms.AR_INVOICE:
                        //    this._Object = new ClsARInvoice();
                        //    _Collection.Add(FormUID, this._Object);
                        //    ((ClsARInvoice)this._Object).Form = _Application.Forms.Item(FormUID);
                        //    ((ClsARInvoice)this._Object).FormDefault();
                        //    break;

                        //case Constants.System_Forms.AR_CREDITNOTE:
                        //    this._Object = new ClsARCreditNote();
                        //    _Collection.Add(FormUID, this._Object);
                        //    ((ClsARCreditNote)this._Object).Form = _Application.Forms.Item(FormUID);
                        //    break;

                        //case Constants.System_Forms.GOODS_RETURN:
                        //    this._Object = new ClsGoodsReturn();
                        //    _Collection.Add(FormUID, this._Object);
                        //    ((ClsGoodsReturn)this._Object).Form = _Application.Forms.Item(FormUID);
                        //    ((ClsGoodsReturn)this._Object).FormDefault();
                        //    break;

                        //case Constants.System_Forms.Production_ORDER:
                        //    this._Object = new ClsProductionOrder();
                        //    _Collection.Add(FormUID, this._Object);
                        //    ((ClsProductionOrder)this._Object).Form = _Application.Forms.Item(FormUID);
                        //    ((ClsProductionOrder)this._Object).FormDefault();
                        //    break;
                    }
                }

                if (this._Collection.Contains(FormUID))
                {
                    this._Object = (LGBOM_LaxmiMotors.Base)this._Collection[FormUID];

                    if (pVal.BeforeAction && this._Object.IsLookUpOpen)
                    {
                        BubbleEvent = false;
                        _Application.Forms.Item(this._Object.LookUpUID).Select();
                    }
                    else
                    {
                        this._Object.Item_Event(FormUID, ref pVal, ref BubbleEvent);
                    }

                    if (!pVal.BeforeAction && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                    {
                        if (_LookCollection.ContainsKey(FormUID))
                        {
                            this._Object = (LGBOM_LaxmiMotors.Base)_Collection[_LookCollection[FormUID]];
                            _LookCollection.Remove(FormUID);
                            this._Object.IsLookUpOpen = false;
                        }
                        this._Object.Dispose();
                        _Collection.Remove(FormUID);
                        System.GC.Collect();
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region RIGHT CLICK EVENT

        private void SboApplication_RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (this._Collection.Contains(eventInfo.FormUID))
            {
                this._Object = (LGBOM_LaxmiMotors.Base)this._Collection[eventInfo.FormUID];
                this._Object.RightClick_Event(ref eventInfo, ref BubbleEvent);
            }
        }

        #endregion

        #region APPLICATION EVENT
        private void SboApplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            try
            {
                switch (EventType)
                {
                    case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                        Utilities.LoadMenus(Constants.Menus.REMOVE_MENUS);
                        Utilities.ShowWarningMessage("Add-On is Disconnected Successfully ... ");
                        CloseApplication();
                        break;
                }
            }
            catch (Exception ex)
            {
                Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }
        #endregion

        #region FORM DATA EVENT
        private void SboApplication_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (BusinessObjectInfo.FormUID != null)
                {
                    if (this._Collection.Contains(BusinessObjectInfo.FormUID))
                    {
                        this._Object = (LGBOM_LaxmiMotors.Base)this._Collection[BusinessObjectInfo.FormUID];
                        this._Object.FormData_Event(ref BusinessObjectInfo, ref BubbleEvent);
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion
        #endregion

        #region CLOSE APPLICATION
        private void CloseApplication()
        {
            try
            {
                if (this._Application != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_Application);
                    _Application = null;
                }
                if (this._Company != null)
                {
                    if (this._Company.Connected)
                        this._Company.Disconnect();

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_Company);
                    _Company = null;
                }
                this._Collection = null;
                this._LookCollection = null;
                _Listener = null;
                System.Windows.Forms.Application.Exit();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                // clsUtilities.oApplication = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }
        #endregion
    }
}
