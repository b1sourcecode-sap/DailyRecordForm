using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sap.Data.Hana;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Xml;
using System.Collections;
using System.Threading;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Globalization;
using Sap.Data.Hana;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace LGBOM_LaxmiMotors
{
    class ClsMaterialManagement : Base
    {

        #region VARIABLE DECLARATION
        SAPbouiCOM.Item _ExistItem, _NewItem, _Items;
        private bool flag;
        private bool errorFlag;
        public SAPbouiCOM.Matrix _Matrix,_Matrix1,_Matrix2 ,_Matrix3;
        SAPbouiCOM.StaticText _StaticText;
        SAPbouiCOM.Button _Button;
        string _DocEntry = "";
        string _PRNum = string.Empty, _PRName = string.Empty, _RevNum = string.Empty, _Warehouse = string.Empty, _RevRmrks = string.Empty;
        double _Qty = 0;
        public string ErrMsg = string.Empty;
        public int ErrCode = 0;
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        private static HanaConnection _conn;
        private static HanaDataAdapter _da;
        //private static DataTable _dataTable;
        private static HanaParameter _parm1;
        private static HanaParameter _parm2;
        private static int _errors = 0;
        string strcon, _Sql;
        private SAPbouiCOM.EditText _EditText;
        private SAPbouiCOM.LinkedButton _Linked;
        private SAPbouiCOM.ComboBox _ComboBoxM;
        private SAPbouiCOM.ComboBox _ComboBoxC, _ComboBoxCRow;
        private SAPbouiCOM.ComboBox _ComboBoxB;
        private SAPbouiCOM.ComboBox _ComboBoxL;
        private SAPbouiCOM.ComboBox _ComboBoxH;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource = null, _ChildDBDataSource = null , _ChildDBDataSource1 = null, _ChildDBDataSource2 = null,
            _ChildDBDataSource3 = null;
        private SAPbouiCOM.ChooseFromListCollection _CFLS;
        private SAPbouiCOM.ChooseFromList _CFL;
        private SAPbouiCOM.ChooseFromListCreationParams _CFLC;
        private SAPbouiCOM.Conditions _Cons;
        private SAPbouiCOM.Condition _Con;
        private int RowID { get; set; }
        private string _Code { get; set; }
        private string _LineId { get; set; }
        private string _ActivityType = string.Empty;
        string _ItemCode = string.Empty, _ItemName = string.Empty;
        private SAPbobsCOM.Recordset _Rs, _RS1, _RS2, _RS3 ,_RSgetAttach ,RSgetEnty;
        SAPbouiCOM.ChooseFromListEvent _SysCFLEvent = null;
        SAPbouiCOM.DBDataSource _DBData = null;
        private System.Collections.Generic.List<string> ListOfBlockedID;
        internal string _ServiceID { get; set; }
        internal string _ActivityTypes { get; set; }
        public bool chkStatus = true;
        static int StageID = 0;
        #endregion

        //#region CONSTRUCTOR & DISTRUCTOR
        public ClsMaterialManagement()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsMaterialManagement()
        {
        }
        //#endregion

        //#region FormDataEvent
        //public override void FormData_Event(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool oBubbleEvent)
        //{
        //    try
        //    {
        //        switch (BusinessObjectInfo.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //                Form.Freeze(true);
        //                string itemcode = ((SAPbouiCOM.EditText)Form.Items.Item("txtPrNo").Specific).Value;
        //                string status = ((SAPbouiCOM.ComboBox)Form.Items.Item("cmbStatus").Specific).Selected.Value;
        //                Utilities.SerializedMartix(ref _Form, ref _Matrix);
        //                AddSequenceNumbersToMatrix(_Matrix);
        //                Route_Line_Set(_Matrix);
        //                Disable_Form_Item(itemcode,status);
        //                Form.Freeze(false);
        //                break;

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Utilities.ShowErrorMessage(ex.Message);
        //    }

        //}
        //#endregion

        #region ItemEvent
        public override void Item_Event(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
        {
            #region Before Action
            if (pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        //if (pVal.ColUID == "V_0" && Form.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)//&& pVal.CharPressed == 9
                        //{
                        //    _ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific;
                        //    if (_ComboBoxC.Selected.Value == "Resource")
                        //    {
                        //        _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                        //        DataSet ds = LoadCompanySetting();
                        //        string ServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
                        //        string UserID = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
                        //        string Password = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();
                        //        string CurrentSchema = ds.Tables[0].Rows[0]["Database"].ToString().Trim();
                        //        _conn = new HanaConnection("Server=" + ServerName + ";UserID=" + UserID + ";Password=" + Password + ";Current Schema=" + CurrentSchema + "");
                        //        this._Object = new ClsChooseFromList();
                        //        ((ClsChooseFromList)this._Object).FormTitle = "Resource List";
                        //        ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref oBubbleEvent, "select \"VisResCode\",\"ResName\",\"UnitOfMsr\" from \"ORSC\" ", strcon, true);

                        //    }
                        //    else if (_ComboBoxC.Selected.Value == "Route Stage")
                        //    {
                        //        _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                        //        DataSet ds = LoadCompanySetting();
                        //        string ServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
                        //        string UserID = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
                        //        string Password = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();
                        //        string CurrentSchema = ds.Tables[0].Rows[0]["Database"].ToString().Trim();
                        //        _conn = new HanaConnection("Server=" + ServerName + ";UserID=" + UserID + ";Password=" + Password + ";Current Schema=" + CurrentSchema + "");
                        //        this._Object = new ClsChooseFromList();
                        //        ((ClsChooseFromList)this._Object).FormTitle = "Route List";
                        //        ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref oBubbleEvent, "select \"Code\",\"Desc\" from \"ORST\" ", strcon, true);
                        //    }
                        //}
                        //if (pVal.ItemUID == "5" && pVal.CharPressed != 9 && pVal.InnerEvent == false)
                        //{
                        //    oBubbleEvent = false;
                        //    return;
                        //}
                        //if (pVal.ItemUID == "6" && pVal.CharPressed != 9 && pVal.InnerEvent == false)
                        //{
                        //    oBubbleEvent = false;
                        //    return;
                        //}

                        break;
                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
                        //if (pVal.ColUID == "V_0")
                        //{
                        //    _ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific;
                        //    //if (_ComboBoxC.Selected.Value == "Resource")
                        //    //{
                        //    //    //string ColVal = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific).Value.ToString();
                        //    //    SAPbouiCOM.Column _Column = (SAPbouiCOM.Column)_Matrix.Columns.Item("V_0");
                        //    //    SAPbouiCOM.LinkedButton _LinkedButton = ((SAPbouiCOM.LinkedButton)(_Column.ExtendedObject));
                        //    //    _LinkedButton.LinkedObjectType = "290";
                        //    //}
                        //    if (_ComboBoxC.Selected.Value == "Item")
                        //    {
                        //        //string ColVal = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific).Value.ToString();
                        //        SAPbouiCOM.Column _Column = (SAPbouiCOM.Column)_Matrix.Columns.Item("V_0");
                        //        SAPbouiCOM.LinkedButton _LinkedButton = ((SAPbouiCOM.LinkedButton)(_Column.ExtendedObject));
                        //        _LinkedButton.LinkedObjectType = "4";
                        //    }
                        //    //else
                        //    //{
                        //    //    //string ColVal = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific).Value.ToString();
                        //    //    SAPbouiCOM.Column _Column = (SAPbouiCOM.Column)_Matrix.Columns.Item("V_0");
                        //    //    SAPbouiCOM.LinkedButton _LinkedButton = ((SAPbouiCOM.LinkedButton)(_Column.ExtendedObject));
                        //    //    _LinkedButton.LinkedObjectType = "296";

                        //    //}
                        //}
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                        //if (pVal.ItemUID == "1" && (Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                        //{
                        //    _PRNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value;
                        //    _RevNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPPNo").Specific).Value;
                        //    //_MasterDBDataSource.SetValue("U_PPNum", 0, (Utilities.getMaxColumnValueNumWithItem("@PRODUCTIONPH", "U_PPNum", _PRNum)));
                        //    if (((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value == "")
                        //    {
                        //        Utilities.Message("Product No can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //        oBubbleEvent = false;
                        //        return;
                        //    }
                        //    else if (((SAPbouiCOM.EditText)_Form.Items.Item("txtVDate").Specific).Value == "")
                        //    {
                        //        Utilities.Message("Valid Until date can not left blank. ", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //        oBubbleEvent = false;
                        //        return;
                        //    }
                        //    else if (((SAPbouiCOM.EditText)_Form.Items.Item("txtQty").Specific).Value == "")
                        //    {
                        //        Utilities.Message("Quantity can not left blank. ", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //        oBubbleEvent = false;
                        //        return;
                        //    }
                        //    else if (((SAPbouiCOM.EditText)_Form.Items.Item("txtWhse").Specific).Value == "")
                        //    {
                        //        Utilities.Message("Warehouse can not left blank. ", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //        oBubbleEvent = false;
                        //        return;
                        //    }
                        //    if (((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbStatus").Specific).Value == "")
                        //    {
                        //        Utilities.Message("Status can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //        oBubbleEvent = false;
                        //        return;
                        //    }
                        //    //if (((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbPList").Specific).Value == "")
                        //    //{
                        //    //    Utilities.Message("Price List can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //    //    oBubbleEvent = false;
                        //    //    return;
                        //    //}
                        //    if (Validation() == false)
                        //    {
                        //        oBubbleEvent = false;
                        //        return;
                        //    }
                        //    errorFlag = false;
                        //    //AddPO();
                        //    if (errorFlag == true)
                        //    {
                        //        oBubbleEvent = false;
                        //        return;
                        //    }
                        //}
                        //if (pVal.ItemUID == "1" && Form.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                        //{
                        //    //Disable_Form_item();
                        //}
                        //if (pVal.ItemUID == "1" && (Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                        //{

                        //    string docCode = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value;
                        //    Utilities.ExecuteSQL(ref _RSUsr, "Select T0.\"USER_CODE\" " +
                        //  "From \"OUSR\" T0 where T0.\"U_StatusAuth\" ='YES'");
                        //    List<string> approvers = new List<string>();
                        //    while (!_RSUsr.EoF)
                        //    {
                        //        string code = _RSUsr.Fields.Item(0).Value.ToString().Trim();
                        //        if (!string.IsNullOrEmpty(code)) approvers.Add(code);
                        //        _RSUsr.MoveNext();
                        //    }

                        //    if (approvers.Count > 0)
                        //        SendInternalAlert(docCode, approvers.Count == 1 ? new List<string> { approvers[0] } : approvers);
                        //    else
                        //        Utilities.Application.SBO_Application.MessageBox("No approvers found.");

                        //    _PRNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value;
                        //    _RevNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPPNo").Specific).Value;
                        //    Utilities.ExecuteSQL(ref _RSProd, "Select T0.\"ItemCode\" " +
                        //   "From \"OWOR\" T0   Where T0.\"U_RevNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtPPNo").Specific).Value + "' " +
                        //   "AND T0.\"ItemCode\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value + "'");

                        //    if (_RSProd.RecordCount > 0 && ((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbStatus").Specific).Value == "Draft")
                        //    {
                        //        Utilities.Message("This Revision Alredy in Use, Can Not Change Status To Draft.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //        oBubbleEvent = false;
                        //        return;
                        //    }
                        //    if (((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value == "")
                        //    {
                        //        Utilities.Message("Product No can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //        oBubbleEvent = false;
                        //        return;
                        //    }
                        //    else if (((SAPbouiCOM.EditText)_Form.Items.Item("txtVDate").Specific).Value == "")
                        //    {
                        //        Utilities.Message("Valid Until date can not left blank. ", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //        oBubbleEvent = false;
                        //        return;
                        //    }
                        //    else if (((SAPbouiCOM.EditText)_Form.Items.Item("txtSDate").Specific).Value == "")
                        //    {
                        //        Utilities.Message("Start date can not left blank. ", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //        oBubbleEvent = false;
                        //        return;
                        //    }
                        //    else if (((SAPbouiCOM.EditText)_Form.Items.Item("txtQty").Specific).Value == "")
                        //    {
                        //        Utilities.Message("Quantity can not left blank. ", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //        oBubbleEvent = false;
                        //        return;
                        //    }
                        //    else if (((SAPbouiCOM.EditText)_Form.Items.Item("txtWhse").Specific).Value == "")
                        //    {
                        //        Utilities.Message("Warehouse can not left blank. ", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //        oBubbleEvent = false;
                        //        return;
                        //    }
                        //    if (((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbStatus").Specific).Value == "")
                        //    {
                        //        Utilities.Message("Status can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //        oBubbleEvent = false;
                        //        return;
                        //    }
                        //    //if (((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbPList").Specific).Value == "")
                        //    //{
                        //    //    Utilities.Message("Price List can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //    //    oBubbleEvent = false;
                        //    //    return;
                        //    //}
                        //    if (Validation() == false)
                        //    {
                        //        oBubbleEvent = false;
                        //        return;
                        //    }


                        //    errorFlag = false;
                        //    //AddUpdateProductionOrder();
                        //    //AddUpdateBOM();
                        //    if (errorFlag == true)
                        //    {
                        //        oBubbleEvent = false;
                        //        return;
                        //    }
                        //}
                        if (pVal.ItemUID == "25")
                        {
                            _Form.PaneLevel = 1;
                        }
                        else if (pVal.ItemUID == "26")
                        {
                            _Form.PaneLevel = 2;
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        //if (pVal.ItemUID == "txtPrNo")
                        //{
                        //    _ComboBoxH = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbType").Specific;
                        //    if (_ComboBoxH.Selected.Value == "Standard" || _ComboBoxH.Selected.Value == "Disassembly")
                        //    {
                        //        this.FilterItemType("");
                        //    }
                        //    else
                        //    {
                        //        FilterItemTypeSp("");
                        //    }

                        //}
                        //if (pVal.ItemUID == "27" && pVal.ColUID == "V_0")
                        //{
                        //    this.FilterItemLineType("");
                        //}
                        break;
                }
            }

            #endregion

            #region After Action
            else
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:

                        //if (pVal.ItemUID == "cmbType")
                        //{
                        //    _ComboBoxH = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbType").Specific;
                        //    if (_ComboBoxH.Selected.Value == "Special")
                        //    {
                        //        LoadProductionPC("Special", "");
                        //    }
                        //    if (_ComboBoxH.Selected.Value == "Standard")
                        //    {
                        //        var editTxt = (SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific;
                        //        string itemCode = editTxt.Value?.Trim() ?? string.Empty;
                        //        if (itemCode != "")
                        //            LoadProductionPC("Standard", itemCode);
                        //    }
                        //    if (_ComboBoxH.Selected.Value == "Disassembly")
                        //    {
                        //        var editTxt = (SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific;
                        //        string itemCode = editTxt.Value?.Trim() ?? string.Empty;
                        //        if (itemCode != "")
                        //            LoadProductionPC("Disassembly", itemCode);
                        //    }
                        //}

                        //if (pVal.ColUID == "V_8" && (Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
                        //{
                        //    _ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific;
                        //    if (_ComboBoxC.Selected.Value == "Route Stage")
                        //    {
                        //        _Matrix.CommonSetting.SetCellFontSize(pVal.Row, 1, 12);
                        //        _Matrix.CommonSetting.SetCellFontSize(pVal.Row, 2, 12);
                        //        _Matrix.CommonSetting.SetCellFontSize(pVal.Row, 3, 12);
                        //        _Matrix.CommonSetting.SetCellFontStyle(pVal.Row, 1, BoFontStyle.fs_Bold);
                        //        _Matrix.CommonSetting.SetCellFontStyle(pVal.Row, 2, BoFontStyle.fs_Bold);
                        //        _Matrix.CommonSetting.SetCellFontStyle(pVal.Row, 3, BoFontStyle.fs_Bold);
                        //        //SetCellForeColor(pVal.Row, 1, greenColor);
                        //        _Matrix.CommonSetting.SetCellEditable(pVal.Row, 4, false);
                        //        _Matrix.CommonSetting.SetCellEditable(pVal.Row, 6, false);
                        //        _Matrix.CommonSetting.SetCellEditable(pVal.Row, 7, false);
                        //    }

                        //}
                        //if (pVal.ItemUID == "cmbPList")
                        //{
                        //    _ComboBoxM = (SAPbouiCOM.ComboBox)Form.Items.Item("cmbPList").Specific;
                        //    _EditText = (SAPbouiCOM.EditText)Form.Items.Item("txtPNo").Specific;
                        //    _EditText.Value = _ComboBoxM.Selected.Description.ToString();
                        //}
                        break;
                    #region Item Pressed
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        #region Add Attachment Path 
                        if (pVal.ItemUID == "btnAttach" && Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            string filePath = ShowFileDialog();
                            if (!string.IsNullOrEmpty(filePath))
                            {
                                AttachmentResult result = SaveAttachmentToSAP(filePath);
                                if (result.AbsEntry != -1)
                                {
                                    string destPath = result.DestPath;
                                    int absEntry = result.AbsEntry;
                                    SAPbouiCOM.Form oForm = Utilities.Application.SBO_Application.Forms.ActiveForm;
                                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                    // Update the DB fields
                                    _Form.DataSources.DBDataSources.Item(0).SetValue("U_AttachPathF", 0, destPath);
                                    _Form.DataSources.DBDataSources.Item(0).SetValue("U_AttachEntry", 0, absEntry.ToString());
                                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    Utilities.Application.SBO_Application.StatusBar.SetText("Attachment added and UDO updated.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                }

                            }

                        }
                        #endregion

                        #region View Attachment Path
                        if (pVal.ItemUID == "btnview" && Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            string absEntryStr = _Form.DataSources.DBDataSources.Item(0).GetValue("U_AttachEntry", 0);

                            if (int.TryParse(absEntryStr, out int absEntry))
                            {
                                OpenAttachment(absEntry);
                            }
                            else
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText("Invalid AbsEntry to open attachment.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            }
                        }
                        #endregion

                        if (pVal.ItemUID== "btnCalcu" && Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            CalculateMatrixLogic();

                            Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

                            SAPbouiCOM.Item updateBtn = Form.Items.Item("1");
                            if (updateBtn != null && updateBtn.Enabled)
                            {
                                updateBtn.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }

                            System.Threading.Thread.Sleep(500);

                            SAPbouiCOM.Item okBtn = Form.Items.Item("2");
                            if (okBtn != null && okBtn.Enabled)
                            {
                                okBtn.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                        }


                        //if (pVal.ItemUID == "btnCrtPO" && Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        //{
                        //    AddUpdateProductionOrder();
                        //}
                        break;
                    #endregion

                    #region Set CFL Values
                    case BoEventTypes.et_CHOOSE_FROM_LIST:
                        //_SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                        //if (pVal.ItemUID == "txtPrNo" && _SysCFLEvent.SelectedObjects != null)
                        //{
                        //    // Set main DB fields
                        //    string itemCode = _SysCFLEvent.SelectedObjects.GetValue("ItemCode", 0).ToString().Trim();
                        //    _MasterDBDataSource.SetValue("U_PRNum", 0, itemCode);
                        //    _MasterDBDataSource.SetValue("U_PRName", 0, _SysCFLEvent.SelectedObjects.GetValue("ItemName", 0).ToString().Trim());
                        //    _MasterDBDataSource.SetValue("U_UOM", 0, _SysCFLEvent.SelectedObjects.GetValue("InvntryUom", 0).ToString().Trim());
                        //    //_MasterDBDataSource.SetValue("U_PPNum", 0, Utilities.getMaxColumnValueNumWithItem("@PRODUCTIONPH", "U_PPNum", itemCode));
                        //    _MasterDBDataSource.SetValue("U_PQty", 0, "1");
                        //    string whsCode = "";
                        //    SAPbobsCOM.Recordset oRSWhse = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //    string query = $"SELECT \"ToWH\" FROM \"OITT\" WHERE \"Code\" = '{itemCode}'";

                        //    oRSWhse.DoQuery(query);

                        //    if (!oRSWhse.EoF)
                        //    {
                        //        _MasterDBDataSource.SetValue("U_Whse", 0, oRSWhse.Fields.Item("ToWH").Value.ToString());
                        //    }
                        //    string orderType = ((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbType").Specific).Value;
                        //    _Form.Select();
                        //    LoadProductionPC(orderType, itemCode);
                        //    AddSequenceNumbersToMatrix(_Matrix);
                        //    //if(orderType =="Special")
                        //    //{
                        //    //    _Matrix.AddRow(1, -1);
                        //    //}
                        //    AddSequenceNumbersToMatrix(_Matrix);
                        //}
                        //if (pVal.ItemUID == "txtWhse" && _SysCFLEvent.SelectedObjects != null)
                        //{
                        //    _MasterDBDataSource.SetValue("U_Whse", 0, _SysCFLEvent.SelectedObjects.GetValue("WhsCode", 0).ToString().Trim());
                        //}
                        //if (pVal.ItemUID == "27" && pVal.ColUID == "V_0" && _SysCFLEvent.SelectedObjects != null)
                        //{

                        //    try
                        //    {
                        //        _ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific;
                        //        if (_ComboBoxC.Selected.Value == "Item")
                        //        {
                        //            _Matrix.FlushToDataSource();
                        //            _ChildDBDataSource.SetValue("U_ItemCode", _ChildDBDataSource.Offset, _SysCFLEvent.SelectedObjects.GetValue("ItemCode", 0).ToString().Trim());
                        //            _ChildDBDataSource.SetValue("U_ItemName", _ChildDBDataSource.Offset, _SysCFLEvent.SelectedObjects.GetValue("ItemName", 0).ToString().Trim());
                        //            _ChildDBDataSource.SetValue("U_UOM", _ChildDBDataSource.Offset, _SysCFLEvent.SelectedObjects.GetValue("InvntryUom", 0).ToString().Trim());
                        //            _ChildDBDataSource.SetValue("U_Quantity", _ChildDBDataSource.Offset, "1");
                        //            _ComboBoxB = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_9").Cells.Item(pVal.Row).Specific;
                        //            _Matrix.LoadFromDataSource();
                        //            if (_SysCFLEvent.SelectedObjects.GetValue("IssueMthd", 0).ToString().Trim() == "B")
                        //            {
                        //                _ComboBoxB.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        //            }
                        //            else
                        //            {
                        //                _ComboBoxB.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
                        //            }
                        //            _Matrix.FlushToDataSource();
                        //            _Matrix.LoadFromDataSource();
                        //            AddNewRow();
                        //            _ComboBoxCRow = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row + 1).Specific;
                        //            _ComboBoxCRow.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        //        }
                        //    }
                        //    catch
                        //    {
                        //        Utilities.Application.Company.GetLastError(out ErrCode, out ErrMsg);
                        //    }

                        //}
                        //if (pVal.ItemUID == "27" && pVal.ColUID == "V_3" && _SysCFLEvent.SelectedObjects != null)
                        //{
                        //    try
                        //    {
                        //        _Matrix.FlushToDataSource();
                        //        _ChildDBDataSource.Offset = pVal.Row - 1; // Set the correct row
                        //        _ChildDBDataSource.SetValue("U_Whse", _ChildDBDataSource.Offset, _SysCFLEvent.SelectedObjects.GetValue("WhsCode", 0).ToString().Trim());
                        //        _Matrix.LoadFromDataSource();
                        //        _Matrix.FlushToDataSource();
                        //        //int i = pVal.Row;
                        //        //dynamic cflArg1 = pVal;
                        //        //string UID1 = cflArg1.ChooseFromListUID;
                        //        //dynamic dt1 = cflArg1.SelectedObjects;
                        //        //_Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("27").Specific;
                        //        //((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(i).Specific).Value = dt1.GetValue("WhsCode", 0).ToString();
                        //    }

                        //    catch (Exception ex)
                        //    {
                        //        // Handle the exception, log it, or display an error message.
                        //    }
                        //    chkStatus = true;
                        //}
                        break;
                    #endregion

                    #region Key Down 
                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        if (pVal.ItemUID == "txtQty" && !pVal.BeforeAction)
                        {
                            double inputQty = 0;
                            double.TryParse(((SAPbouiCOM.EditText)_Form.Items.Item("txtQty").Specific).Value, out inputQty);

                            var childDS = _Form.DataSources.DBDataSources.Item("@PRODUCTIONPC");
                            for (int i = 0, len = childDS.Size; i < len; i++)
                            {
                                double baseQty = 0;
                                double.TryParse(childDS.GetValue("U_BaseQty", i).ToString(), out baseQty);

                                childDS.SetValue("U_PlannedQty", i, (baseQty * inputQty).ToString());
                            }
                            _Matrix.LoadFromDataSource();
                            // ((SAPbouiCOM.Grid)_Form.Items.Item("yourGridUID").Specific).LoadFromDataSource();
                        }
                        break;
                        #endregion
                }
            }

            #endregion
        }
        #endregion

        //#region MENU EVENT
        //public override void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool oBubbleEvent)
        //{
        //    #region BEFORE ACTION
        //    if (pVal.BeforeAction)
        //    {
        //        Form.Freeze(true);
        //        switch (pVal.MenuUID)
        //        {
        //            case Constants.System_Menus.mnu_Duplicate:
        //                {
        //                    _PRNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value;
        //                    //Disable_Form_Item();
        //                }
        //                break;
        //        }
        //        Form.Freeze(false);
        //    }
        //    #endregion

        //    #region AFTER ACTION
        //    else
        //    {
        //        try
        //        {
        //            Form.Freeze(true);
        //            switch (pVal.MenuUID)
        //            {
        //                case Constants.System_Menus.mnu_FIND:
        //                    {
        //                        _Form.Items.Item("txtPPNo").Enabled = true;

        //                    }
        //                    break;
        //                case Constants.System_Menus.mnu_ADD:
        //                    {
        //                        _Form.Items.Item("txtPPNo").Enabled = false;
        //                    }
        //                    break;
        //                case Constants.System_Menus.mnu_Duplicate:
        //                    {
        //                        //_MasterDBDataSource.SetValue("U_PPNum", 0, (Utilities.getMaxColumnValueNumWithItem("@PRODUCTIONPH", "U_PPNum", _PRNum)));
        //                        //_Form.Items.Item("txtPPNo").Enabled = false;
        //                        //_ComboBoxH.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //                        //AddSequenceNumbersToMatrix(_Matrix);
        //                    }
        //                    break;
        //                case Constants.System_Menus.mnu_FIRST:
        //                    _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("27").Specific;
        //                    AddSequenceNumbersToMatrix(_Matrix);
        //                    //Disable_Form_item();
        //                    break;

        //                case Constants.System_Menus.mnu_LAST:
        //                    _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("27").Specific;
        //                    AddSequenceNumbersToMatrix(_Matrix);
        //                    //Disable_Form_item();
        //                    break;

        //                case Constants.System_Menus.mnu_NEXT:
        //                    _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("27").Specific;
        //                    AddSequenceNumbersToMatrix(_Matrix);
        //                    //Disable_Form_item();

        //                    break;

        //                case Constants.System_Menus.mnu_PREVIOUS:
        //                    _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("27").Specific;
        //                    AddSequenceNumbersToMatrix(_Matrix);
        //                    //Disable_Form_item();
        //                    break;

        //            }
        //            if (pVal.MenuUID == "1292" && pVal.BeforeAction == false)
        //            {
        //                AddNewRow();
        //            }
        //        }

        //        catch (Exception ex)
        //        {
        //            Utilities.ShowErrorMessage(ex.Message);
        //        }
        //        finally { Form.Freeze(false); }
        //    }
        //    #endregion
        //}
        //#endregion

        //#region CHOOSE FROM LIST EVENT

        //public override void ChooseFromList_Event()
        //{
        //    try
        //    {
        //        _Form.Freeze(true);
        //        if (this._CFLEvent.ItemUID == "27" && _CFLEvent.ColUID == "V_0")
        //        {
        //            for (int i = 0; i < _CFLEvent.oCFL_SelectedResult.Rows.Count + _CFLEvent.Row - 1; i++)
        //            {
        //                _ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(_Matrix.VisualRowCount).Specific;
        //                _ComboBoxB = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_9").Cells.Item(_Matrix.VisualRowCount).Specific;
        //                if (_ComboBoxC.Selected.Value == "Resource")
        //                {
        //                    int a = i;
        //                    _Matrix.FlushToDataSource();
        //                    _ChildDBDataSource.SetValue("U_ItemCode", _ChildDBDataSource.Offset, _CFLEvent.oCFL_SelectedResult.Rows[i][0].ToString());
        //                    _ChildDBDataSource.SetValue("U_ItemName", _ChildDBDataSource.Offset, _CFLEvent.oCFL_SelectedResult.Rows[i][1].ToString());
        //                    _ChildDBDataSource.SetValue("U_UOM", _ChildDBDataSource.Offset, _CFLEvent.oCFL_SelectedResult.Rows[i][2].ToString());
        //                    _ChildDBDataSource.SetValue("U_Quantity", _ChildDBDataSource.Offset, "1");
        //                    _ComboBoxB.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //                    _Matrix.LoadFromDataSource();
        //                    AddNewRow();
        //                    a = i;
        //                    _ComboBoxCRow = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(_Matrix.VisualRowCount).Specific;
        //                    _ComboBoxCRow.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //                }
        //                else
        //                {
        //                    int a = i;
        //                    _Matrix.FlushToDataSource();
        //                    _ChildDBDataSource.SetValue("U_ItemCode", _ChildDBDataSource.Offset, _CFLEvent.oCFL_SelectedResult.Rows[i][0].ToString());
        //                    _ChildDBDataSource.SetValue("U_ItemName", _ChildDBDataSource.Offset, _CFLEvent.oCFL_SelectedResult.Rows[i][1].ToString());
        //                    _Matrix.LoadFromDataSource();
        //                    AddNewRow();
        //                    a = i;
        //                    _ComboBoxCRow = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(_Matrix.VisualRowCount).Specific;
        //                    _ComboBoxCRow.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        ex.ToString();
        //    }
        //    finally
        //    {
        //        _Form.Freeze(false);
        //    }
        //}

        //#endregion

        //#region  Conditional Reload data in Form

        //private void ConditionalReloadData()
        //{
        //    try
        //    {
        //        _Form.Freeze(true);
        //        _DBData.Query(null);
        //        _Matrix.LoadFromDataSource();
        //        _DBData.Clear();
        //        _Matrix.AddRow(1, _Matrix.RowCount);

        //    }
        //    catch (Exception ex)
        //    {
        //        Utilities.ShowErrorMessage(ex.Message);
        //    }
        //    finally { _Form.Freeze(false); }
        //}

        //#endregion

        #region FORM DEFAULT
        internal void FormDefault()
        {
            _Form.Freeze(true);

            try
            {
                string today = DateTime.Now.ToString("yyyyMMdd");
                string checkQuery = $@"SELECT COUNT(*) AS ""Count"" FROM ""@DLYWRKRPRTMMH"" WHERE ""U_DocDate"" = '{today}'";
                SAPbobsCOM.Recordset rsCheck = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Utilities.ExecuteSQL(ref rsCheck, checkQuery);

                int count = Convert.ToInt32(rsCheck.Fields.Item("Count").Value);

                if (count > 0)
                {
                    //Utilities.Application.SBO_Application.StatusBar.SetText("There is already one entry for today. Please update it by using Find.",
                    //SAPbouiCOM.BoMessageTime.bmt_Short,
                    //SAPbouiCOM.BoStatusBarMessageType.smt_Error); 
                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    ((SAPbouiCOM.EditText)_Form.Items.Item("Item_1").Specific).Value = today;
                    _Form.Items.Item("1").Click();
                    return;
                }

                // Ensure form is in ADD mode
                if (_Form.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                try
                {
                    SAPbouiCOM.ComboBox cmbStatus = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbStatus").Specific;
                    if (cmbStatus.ValidValues.Count > 0)
                    {
                        cmbStatus.Select("Open", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                }
                catch (Exception ex)
                {
                    Utilities.Application.SBO_Application.StatusBar.SetText("Failed to set default status: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }

                // Freeze the form while loading data
                _Form.Freeze(true);
                ((SAPbouiCOM.EditText)_Form.Items.Item("Item_1").Specific).Value = today;
                // Query for default matrix values
                string query = @"
                SELECT T0.""U_INPUTMATRIC"", T0.""U_AUTO"", T0.""U_ATTENDENCE"", T0.""U_TASKOVERVIEW""
                FROM ""@DLYWRKRPRTMMM"" T0
                ORDER BY T0.""U_SqNo""";

                SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Utilities.ExecuteSQL(ref oRS, query);

                // Get matrices
                SAPbouiCOM.Matrix matrix1 = (SAPbouiCOM.Matrix)_Form.Items.Item("Item_7").Specific;
                SAPbouiCOM.Matrix matrix2 = (SAPbouiCOM.Matrix)_Form.Items.Item("Item_8").Specific;
                SAPbouiCOM.Matrix matrix3 = (SAPbouiCOM.Matrix)_Form.Items.Item("Item_9").Specific;
                SAPbouiCOM.Matrix matrix4 = (SAPbouiCOM.Matrix)_Form.Items.Item("Item_10").Specific;

                // Get DBDataSources
                var ds1 = _Form.DataSources.DBDataSources.Item("@DLYWRKRPRTMMC1");
                var ds2 = _Form.DataSources.DBDataSources.Item("@DLYWRKRPRTMMC2");
                var ds3 = _Form.DataSources.DBDataSources.Item("@DLYWRKRPRTMMC3");
                var ds4 = _Form.DataSources.DBDataSources.Item("@DLYWRKRPRTMMC4");

                // Clear data
                matrix1.Clear(); matrix2.Clear(); matrix3.Clear(); matrix4.Clear();
                ds1.Clear(); ds2.Clear(); ds3.Clear(); ds4.Clear();

                int row = 0;
                while (!oRS.EoF)
                {
                    ds1.InsertRecord(row);
                    ds2.InsertRecord(row);
                    ds3.InsertRecord(row);
                    ds4.InsertRecord(row);
                    ds1.SetValue("U_Content", row, oRS.Fields.Item("U_INPUTMATRIC").Value.ToString());
                    ds2.SetValue("U_Content", row, oRS.Fields.Item("U_ATTENDENCE").Value.ToString());
                    ds3.SetValue("U_Content", row, oRS.Fields.Item("U_AUTO").Value.ToString());
                    ds4.SetValue("U_Content", row, oRS.Fields.Item("U_TASKOVERVIEW").Value.ToString());

                    row++;
                    oRS.MoveNext();
                }

                // Load data into matrices
                matrix1.LoadFromDataSource();
                matrix2.LoadFromDataSource();
                matrix3.LoadFromDataSource();
                matrix4.LoadFromDataSource();

                matrix1.Columns.Item("Col_0").Editable = false;
                matrix2.Columns.Item("Col_0").Editable = false;
                matrix3.Columns.Item("Col_0").Editable = false;
                matrix4.Columns.Item("Col_0").Editable = false;

                for (int i = 1; i <= matrix1.RowCount; i++)
                {
                    matrix1.SetCellWithoutValidation(i, "#", i.ToString());
                    matrix2.SetCellWithoutValidation(i, "#", i.ToString());
                    matrix3.SetCellWithoutValidation(i, "#", i.ToString());
                    matrix4.SetCellWithoutValidation(i, "#", i.ToString());
                }

                // Make # column read-only
                matrix1.Columns.Item("#").Editable = false;
                matrix2.Columns.Item("#").Editable = false;
                matrix3.Columns.Item("#").Editable = false;
                matrix4.Columns.Item("#").Editable = false;


                // Unfreeze the form
                _Form.Freeze(false);

            }
            finally
            {
                _Form.Freeze(false);
            }


        }
        #endregion
        //#region Add Sequence No
        //private void AddSequenceNumbersToMatrix(Matrix oMatrix)
        //{
        //    //oMatrix.Columns.Item("V_-1").Editable = true;
        //    oMatrix.Columns.Item("V_-1").Visible = true;
        //    for (int i = 1; i <= oMatrix.RowCount; i++)
        //    {
        //        oMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.Value = i.ToString();
        //    }
        //}
        //#endregion

        //#region Set New line
        //private void SetNewLine(string FormUID, int Row)
        //{
        //    _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("25").Specific;
        //    _ChildDBDataSource.Offset = Row - 1;
        //    AddSequenceNumbersToMatrix(_Matrix);
        //    _ChildDBDataSource.SetValue("U_Type", _ChildDBDataSource.Offset, Row.ToString());
        //    _ChildDBDataSource.SetValue("U_ItemCode", _ChildDBDataSource.Offset, "");
        //    _ChildDBDataSource.SetValue("U_ItemName", _ChildDBDataSource.Offset, "");
        //    _ChildDBDataSource.SetValue("U_Quantity", _ChildDBDataSource.Offset, "");
        //    _ChildDBDataSource.SetValue("U_UOM", _ChildDBDataSource.Offset, "");
        //    _ChildDBDataSource.SetValue("U_Whse", _ChildDBDataSource.Offset, "");
        //    _Matrix.SetLineData(Row);
        //    _Matrix.FlushToDataSource();
        //    Utilities.SerializedMartix(ref _Form, ref _Matrix);
        //}
        //#endregion(

        //#region FILTER Item Type
        //internal void FilterItemType(string CFLId)
        //{

        //    _Form.Freeze(true);
        //    try
        //    {
        //        _Cons = new Conditions();
        //        _Con = _Cons.Add();
        //        _Con.BracketOpenNum = 1;
        //        _Con.Alias = "U_BOMItem";
        //        _Con.Operation = BoConditionOperation.co_EQUAL;
        //        _Con.CondVal = "YES";
        //        _Con.BracketCloseNum = 1;
        //        _Form.ChooseFromLists.Item("CFL_Itm").SetConditions(_Cons);
        //    }


        //    catch (Exception ex)
        //    {
        //        Utilities.ShowErrorMessage(ex.Message);
        //    }
        //    finally
        //    {
        //        _Form.Freeze(false);
        //    }

        //}

        //internal void FilterItemTypeSp(string CFLId)
        //{

        //    _Form.Freeze(true);
        //    try
        //    {
        //        _Cons = new Conditions();
        //        _Con = _Cons.Add();
        //        _Con.BracketOpenNum = 1;
        //        _Con.Alias = "InvntItem";
        //        _Con.Operation = BoConditionOperation.co_EQUAL;
        //        _Con.CondVal = "Y";
        //        _Con.BracketCloseNum = 1;
        //        _Form.ChooseFromLists.Item("CFL_Itm").SetConditions(_Cons);
        //    }


        //    catch (Exception ex)
        //    {
        //        Utilities.ShowErrorMessage(ex.Message);
        //    }
        //    finally
        //    {
        //        _Form.Freeze(false);
        //    }

        //}


        //internal void FilterItemLineType(string CFLId)
        //{

        //    _Form.Freeze(true);
        //    try
        //    {
        //        _Cons = new Conditions();

        //        _Con = _Cons.Add();
        //        _Con.BracketOpenNum = 1;
        //        _Con.Alias = "ItemType";
        //        _Con.Operation = BoConditionOperation.co_NOT_EQUAL;
        //        _Con.CondVal = "F";
        //        _Con.BracketCloseNum = 1;
        //        _Form.ChooseFromLists.Item("CFL_LItm").SetConditions(_Cons);

        //    }


        //    catch (Exception ex)
        //    {
        //        Utilities.ShowErrorMessage(ex.Message);
        //    }
        //    finally
        //    {
        //        _Form.Freeze(false);
        //    }

        //}
        //#endregion

        //#region LOAD COMPANY SETTINGS
        //public static DataSet LoadCompanySetting()
        //{
        //    XmlDataDocument xmlDoc = new XmlDataDocument();
        //    DataSet dataSet = new DataSet();
        //    dataSet.ReadXml("AppConfig.xml");
        //    //dataSet.ReadXml("App.Config");
        //    return dataSet;

        //}
        //#endregion

        //#region VALIDATION FUNCTION

        //private bool Validation()
        //{
        //    try
        //    {


        //        bool skipRow = false;
        //        for (int i = 1; i <= _Matrix.VisualRowCount - 1; i++)
        //        {

        //            if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value.Trim() == "")
        //            {
        //                Utilities.Message("Type can not left blank", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                return false;
        //            }
        //            if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value.Trim() == "")
        //            {
        //                Utilities.Message("Child Item Code can not left blank", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                return false;
        //            }
        //            if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_7").Cells.Item(i).Specific).Value.Trim() == "")
        //            {
        //                Utilities.Message("Child Item Name can not left blank", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                return false;
        //            }

        //            if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(i).Specific).Value.Trim() == "" && ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value.Trim() != "Route Stage")
        //            {
        //                Utilities.Message("Issue Warehouse can not left blank", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                return false;
        //            }
        //            if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_9").Cells.Item(i).Specific).Value.Trim() == "" && ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value.Trim() != "Route Stage")
        //            {
        //                Utilities.Message("Issue Method can not left blank", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                return false;
        //            }
        //            Utilities.ExecuteSQL(ref _RSMgItm, "Select T0.\"ManSerNum\", T0.\"ManBtchNum\"  " +
        //                  "From \"OITM\" T0   Where T0.\"ItemCode\" = '" + ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value.Trim() + "' ");
        //            if (_RSMgItm.Fields.Item(0).Value == "Y" || _RSMgItm.Fields.Item(1).Value == "Y")
        //            {
        //                if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_9").Cells.Item(i).Specific).Value.Trim() == "BackFlush")
        //                {
        //                    string itemCode = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value.Trim();
        //                    Utilities.Message("Issue Method can not be BackFlush for Serial/Batch method , Item Code: " + itemCode, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                    return false;
        //                }
        //            }
        //        }
        //        return true;
        //    }
        //    catch
        //    {
        //        throw;
        //    }
        //}

        //#endregion

        //#region Add New Row
        //private void AddNewRow()
        //{
        //    SAPbouiCOM.DBDataSource dbData = default(SAPbouiCOM.DBDataSource);
        //    dbData = (SAPbouiCOM.DBDataSource)_Form.DataSources.DBDataSources.Item(1);
        //    int selRow = _Matrix.GetNextSelectedRow();
        //    dbData.Clear();
        //    _Matrix.AddRow(1, selRow - 1);
        //    AddSequenceNumbersToMatrix(_Matrix);
        //}
        //#endregion

        //#region Create BOM
        //private void AddUpdateBOM()
        //{
        //    if (((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbStatus").Specific).Value == "Active")
        //    {
        //        try
        //        {
        //            _PRNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value;
        //            _PRName = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNam").Specific).Value;
        //            _Qty = Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("txtQty").Specific).Value);
        //            _Warehouse = ((SAPbouiCOM.EditText)_Form.Items.Item("txtWhse").Specific).Value;
        //            _RevNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPPNo").Specific).Value;
        //            _RevRmrks = ((SAPbouiCOM.EditText)_Form.Items.Item("txtVRmrks").Specific).Value;
        //            int _PriceList = Convert.ToInt32(((SAPbouiCOM.EditText)_Form.Items.Item("txtPNo").Specific).Value);
        //            Utilities.ExecuteSQL(ref _RSBom, "Select T0.\"Code\" " +
        //               "From \"OITT\" T0   Where  T0.\"Code\" = '" + _PRNum + "'");
        //            // Create a BOM object
        //            SAPbobsCOM.ProductTrees _BOM = (SAPbobsCOM.ProductTrees)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees);
        //            StageID = 0;
        //            if (_Matrix.RowCount > 0)
        //            {
        //                if (_RSBom.RecordCount > 0)
        //                {
        //                    _BOM.TreeCode = _PRNum;
        //                    _BOM.ProductDescription = _PRName;
        //                    _BOM.Quantity = _Qty;
        //                    _BOM.Warehouse = _Warehouse;
        //                    _BOM.TreeType = SAPbobsCOM.BoItemTreeTypes.iProductionTree;
        //                    _BOM.UserFields.Fields.Item("U_PPNum").Value = _RevNum;
        //                    _BOM.UserFields.Fields.Item("U_Remarks").Value = _RevRmrks;
        //                    _BOM.PriceList = _PriceList;

        //                    for (int i = 1; i <= _Matrix.RowCount; i++)
        //                    {
        //                        if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value != "")
        //                        {
        //                            if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Route Stage")
        //                            {
        //                                StageID = (StageID + 1);
        //                                _BOM.Stages.StageEntry = StageID;
        //                                _BOM.Stages.Name = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_7").Cells.Item(i).Specific).Value;
        //                                _BOM.Stages.SequenceNumber = StageID;
        //                                _BOM.Stages.Add();
        //                            }
        //                            else if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Resource" || ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Item")
        //                            {
        //                                if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Item")
        //                                {
        //                                    _BOM.Items.ItemType = ProductionItemType.pit_Item;

        //                                }
        //                                else if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Resource")
        //                                {
        //                                    _BOM.Items.ItemType = ProductionItemType.pit_Resource;
        //                                }
        //                                _BOM.Items.ItemCode = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value;
        //                                _BOM.Items.Quantity = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_6").Cells.Item(i).Specific).Value);
        //                                _BOM.Items.Warehouse = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(i).Specific).Value;
        //                                if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_9").Cells.Item(i).Specific).Value == "BackFlush")
        //                                {
        //                                    _BOM.Items.IssueMethod = BoIssueMethod.im_Backflush;
        //                                }
        //                                else
        //                                {
        //                                    _BOM.Items.IssueMethod = BoIssueMethod.im_Manual;
        //                                }
        //                                _BOM.Items.PriceList = _PriceList;
        //                                _BOM.Items.StageID = StageID;
        //                                _BOM.Items.Add();
        //                            }


        //                        }
        //                    }
        //                    int Result = _BOM.Update();
        //                    if (Result == 0)
        //                    {
        //                        Utilities.ShowWarningMessage("Bills of Materials Updated successfully.");
        //                        errorFlag = false;
        //                    }
        //                    else
        //                    {
        //                        Utilities.Application.Company.GetLastError(out ErrCode, out ErrMsg);
        //                        Utilities.ShowErrorMessage(ErrMsg);
        //                        errorFlag = true;
        //                        return;
        //                    }

        //                }

        //                else
        //                {
        //                    _BOM.TreeCode = _PRNum;
        //                    _BOM.ProductDescription = _PRName;
        //                    _BOM.Quantity = _Qty;
        //                    _BOM.Warehouse = _Warehouse;
        //                    _BOM.TreeType = SAPbobsCOM.BoItemTreeTypes.iProductionTree;
        //                    _BOM.UserFields.Fields.Item("U_PPNum").Value = _RevNum;
        //                    _BOM.UserFields.Fields.Item("U_Remarks").Value = _RevRmrks;
        //                    _BOM.PriceList = _PriceList;
        //                    for (int i = 1; i <= _Matrix.RowCount; i++)
        //                    {
        //                        if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value != "")
        //                        {
        //                            if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Route Stage")
        //                            {
        //                                StageID = (StageID + 1);
        //                                _BOM.Stages.StageEntry = StageID;
        //                                _BOM.Stages.SequenceNumber = StageID;
        //                                _BOM.Stages.Name = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_7").Cells.Item(i).Specific).Value;
        //                                _BOM.Stages.Add();

        //                            }

        //                            else if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Resource" || ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Item")
        //                            {
        //                                if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Item")
        //                                {
        //                                    _BOM.Items.ItemType = ProductionItemType.pit_Item;

        //                                }
        //                                else if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Resource")
        //                                {
        //                                    _BOM.Items.ItemType = ProductionItemType.pit_Resource;
        //                                }
        //                                _BOM.Items.ItemCode = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value;
        //                                _BOM.Items.Quantity = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_6").Cells.Item(i).Specific).Value);
        //                                _BOM.Items.Warehouse = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(i).Specific).Value;
        //                                if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_9").Cells.Item(i).Specific).Value == "BackFlush")
        //                                {
        //                                    _BOM.Items.IssueMethod = BoIssueMethod.im_Backflush;
        //                                }
        //                                else
        //                                {
        //                                    _BOM.Items.IssueMethod = BoIssueMethod.im_Manual;
        //                                }
        //                                _BOM.Items.StageID = StageID;
        //                                _BOM.Items.PriceList = _PriceList;
        //                                _BOM.Items.Add();
        //                            }


        //                        }
        //                    }
        //                    int Result = _BOM.Add();
        //                    if (Result == 0)
        //                    {
        //                        Utilities.ShowWarningMessage("Bills of Materials Created successfully.");
        //                        errorFlag = false;
        //                    }
        //                    else
        //                    {
        //                        Utilities.Application.Company.GetLastError(out ErrCode, out ErrMsg);
        //                        Utilities.ShowErrorMessage(ErrMsg);
        //                        errorFlag = true;
        //                        return;
        //                    }

        //                }
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine($"Error: {ex.Message}");
        //            errorFlag = true;
        //            return;
        //        }
        //        finally
        //        {
        //            // Disconnect from the SAP Business One company database
        //            //Utilities.Application.Company.Disconnect();
        //        }



        //    }
        //}
        //#endregion

        //#region Create or Update Production Order
        //private void AddUpdateProductionOrder()
        //{
        //    try
        //    {
        //        string type = ((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbType").Specific).Value;
        //        string docNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPPNo").Specific).Value;
        //        string itemCode = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value;
        //        string warehouse = ((SAPbouiCOM.EditText)_Form.Items.Item("txtWhse").Specific).Value;
        //        double plannedQty = Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("txtQty").Specific).Value);
        //        DateTime dueDate = DateTime.ParseExact(((SAPbouiCOM.EditText)_Form.Items.Item("txtVDate").Specific).Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
        //        string docEntry = _Form.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0);
        //        // Check if production order already exists
        //        Utilities.ExecuteSQL(ref _RSBom, $"SELECT \"DocEntry\" FROM \"OWOR\" WHERE \"ItemCode\" = '{itemCode}' AND \"U_RevNum\" = '{docEntry}'");
        //        bool isUpdate = _RSBom.RecordCount > 0;
        //        SAPbobsCOM.ProductionOrders prodOrder = (SAPbobsCOM.ProductionOrders)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);
        //        if (!isUpdate)
        //        {
        //            prodOrder.ItemNo = itemCode;
        //            if(type=="Standard")
        //            {
        //                prodOrder.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard;
        //            }
        //           else if(type == "Special")
        //            {
        //                prodOrder.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial;
        //            }
        //            else
        //            {
        //                prodOrder.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotDisassembly;
        //            }
        //            prodOrder.PlannedQuantity = plannedQty;
        //            prodOrder.PostingDate = DateTime.Now;
        //            prodOrder.DueDate = dueDate;
        //            prodOrder.Warehouse = warehouse;
        //            prodOrder.UserFields.Fields.Item("U_RevNum").Value = docEntry; 
        //            prodOrder.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned;

        //            for (int i = 1; i <= _Matrix.RowCount; i++)
        //            {
        //                string lineItem = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value;
        //                if (!string.IsNullOrEmpty(lineItem))
        //                {
        //                    prodOrder.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Item;
        //                    prodOrder.Lines.ItemNo = lineItem;
        //                    prodOrder.Lines.PlannedQuantity = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_12").Cells.Item(i).Specific).Value);
        //                    prodOrder.Lines.Warehouse = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(i).Specific).Value;
        //                    prodOrder.Lines.Add();
        //                }
        //            }

        //            int result = prodOrder.Add();

        //            if (result == 0)
        //            {
        //                int newDocEntry = Convert.ToInt32(Utilities.Application.Company.GetNewObjectKey());
        //                _Form.DataSources.DBDataSources.Item(0).SetValue("U_PRODEntry", 0, newDocEntry.ToString());
        //                if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
        //                _Form.Items.Item("1").Click(); 
        //                Utilities.ShowWarningMessage("Production Order Created successfully.");
        //                _Form.Items.Item("btnCrtPO").Enabled = false;
        //                errorFlag = false;
        //            }
        //            else
        //            {
        //                Utilities.Application.Company.GetLastError(out ErrCode, out ErrMsg);
        //                Utilities.ShowErrorMessage(ErrMsg);
        //                errorFlag = true;
        //            }
        //        }
        //        else
        //        {
        //            Utilities.ShowErrorMessage("Production Order already exists for this production planning document.");
        //            errorFlag = true;
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        Utilities.ShowErrorMessage($"Error: {ex.Message}");
        //        errorFlag = true;
        //    }
        //}
        //#endregion

        //#region EnableFromMenu
        //private void EnableFromMenu()
        //{
        //    _Form.EnableMenu("1288", true);
        //    _Form.EnableMenu("1289", true);
        //    _Form.EnableMenu("1290", true);
        //    _Form.EnableMenu("1291", true);
        //    _Form.EnableMenu("1292", true);
        //    _Form.EnableMenu("1287", true);
        //    _Form.EnableMenu("779", true);
        //}
        //#endregion

        //#region Disable Form item on Load
        //private void Disable_Form_Item( string itemcode,string status)
        //{

        //    if (status == "Approved" && ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrdEnt").Specific).Value== "")
        //    {
        //        _Form.Items.Item("btnCrtPO").Enabled = true;
        //        _Form.Items.Item("cmbType").Enabled = false;
        //        _Form.Items.Item("txtPrNo").Enabled = false;
        //        _Form.Items.Item("txtQty").Enabled = false;
        //        _Form.Items.Item("txtWhse").Enabled = false;
        //        _Form.Items.Item("txtPPNo").Enabled = false;
        //        _Form.Items.Item("txtSDate").Enabled = false;
        //        _Form.Items.Item("txtVDate").Enabled = false;
        //        _Form.Items.Item("txtVRmrks").Enabled = false;
        //        _Form.Items.Item("cmbPList").Enabled = false;
        //        //_Form.Items.Item("btnCrtPO").Enabled = false;
        //        _Form.Items.Item("27").Enabled = false;
        //       // _Form.Items.Item("1").Enabled = false;
        //        _Form.Items.Item("2").Enabled = false;
        //    }
        //    if(((SAPbouiCOM.EditText)_Form.Items.Item("txtPrdEnt").Specific).Value!="")
        //    {
        //        _Form.Items.Item("cmbType").Enabled = false;
        //        _Form.Items.Item("txtPrNo").Enabled = false;
        //        _Form.Items.Item("txtQty").Enabled = false;
        //        _Form.Items.Item("txtWhse").Enabled = false;
        //        _Form.Items.Item("txtPPNo").Enabled = false;
        //        _Form.Items.Item("txtSDate").Enabled = false;
        //        _Form.Items.Item("txtVDate").Enabled = false;
        //        _Form.Items.Item("txtVRmrks").Enabled = false;
        //        _Form.Items.Item("cmbPList").Enabled = false;
        //        _Form.Items.Item("btnCrtPO").Enabled = false;
        //        _Form.Items.Item("27").Enabled = false;
        //        _Form.Items.Item("1").Enabled = false;
        //    }
        //    else if(status!="Approved")
        //    {
        //        _Form.Items.Item("cmbType").Enabled = true;
        //        _Form.Items.Item("txtPrNo").Enabled = true;
        //        _Form.Items.Item("txtQty").Enabled = true;
        //        _Form.Items.Item("txtWhse").Enabled = true;
        //        _Form.Items.Item("txtPPNo").Enabled = true;
        //        _Form.Items.Item("txtSDate").Enabled = true;
        //        _Form.Items.Item("txtVDate").Enabled = true;
        //        _Form.Items.Item("txtVRmrks").Enabled = true;
        //        _Form.Items.Item("cmbPList").Enabled = true;
        //       //_Form.Items.Item("btnCrtPO").Enabled = false;
        //        _Form.Items.Item("27").Enabled = true;
        //        _Form.Items.Item("1").Enabled = true;
        //        _Form.Items.Item("2").Enabled = true;
        //    }
        //}
        //#endregion

        //#region Enable Form item on Load
        //private void Enable_Form_item()
        //{

        //    _Form.Items.Item("txtQty").Enabled = true;
        //    _Form.Items.Item("txtWhse").Enabled = true;
        //    _Form.Items.Item("txtPPNo").Enabled = true;
        //    _Form.Items.Item("txtSDate").Enabled = true;
        //    _Form.Items.Item("txtVDate").Enabled = true;
        //    _Form.Items.Item("cmbStatus").Enabled = true;
        //    _Form.Items.Item("txtVRmrks").Enabled = true;
        //    _Form.Items.Item("cmbPList").Enabled = true;
        //    _Form.Items.Item("27").Enabled = true;
        //}
        //#endregion

        //#region Route line setup
        //private void Route_Line_Set(Matrix _Matrix)
        //{
        //    for (int i = 1; i <= _Matrix.RowCount; i++)
        //    {

        //        if (_Matrix.Columns.Item("V_8").Cells.Item(i).Specific.Value == "Route Stage")
        //        {
        //            _Matrix.CommonSetting.SetCellFontSize(i, 1, 10);
        //            _Matrix.CommonSetting.SetCellFontSize(i, 2, 10);
        //            _Matrix.CommonSetting.SetCellFontSize(i, 3, 10);
        //            _Matrix.CommonSetting.SetCellFontStyle(i, 1, BoFontStyle.fs_Bold);
        //            _Matrix.CommonSetting.SetCellFontStyle(i, 2, BoFontStyle.fs_Bold);
        //            _Matrix.CommonSetting.SetCellFontStyle(i, 3, BoFontStyle.fs_Bold);
        //            //SetCellForeColor(pVal.Row, 1, greenColor);
        //            _Matrix.CommonSetting.SetCellEditable(i, 4, false);
        //            _Matrix.CommonSetting.SetCellEditable(i, 6, false);
        //            _Matrix.CommonSetting.SetCellEditable(i, 7, false);
        //        }
        //    }

        //}
        //#endregion

        //#region Load Matrix Based on Type
        //private void LoadProductionPC(string orderType, string itemCode)
        //{
        //    _Form.Freeze(true);

        //    var childDS = _Form.DataSources.DBDataSources.Item("@PRODUCTIONPC");
        //    while (childDS.Size > 0)
        //        childDS.RemoveRecord(0);

        //    if (orderType == "Standard" || orderType == "Disassembly")
        //    {
        //        string query = $@"
        //        SELECT T1.""Code"", T2.""ItemName"", T1.""Quantity"", T2.""InvntryUom"",T0.""Qauntity"",T1.""Warehouse"",T1.""IssueMthd""
        //        FROM ""OITT"" T0 
        //        INNER JOIN ""ITT1"" T1 ON T0.""Code"" = T1.""Father""
        //        INNER JOIN ""OITM"" T2 ON T1.""Code"" = T2.""ItemCode""
        //        WHERE T0.""Code"" = '{itemCode.Replace("'", "''")}'";
        //        Utilities.ExecuteSQL(ref _RSBomP, query);
        //         while (!_RSBomP.EoF)
        //            {
        //                childDS.InsertRecord(childDS.Size);
        //                int dsIndex = childDS.Size - 1;

        //                childDS.SetValue("U_Type", dsIndex, "Item");
        //                childDS.SetValue("U_ItemCode", dsIndex, _RSBomP.Fields.Item("Code").Value.ToString().Trim());
        //                childDS.SetValue("U_ItemName", dsIndex, _RSBomP.Fields.Item("ItemName").Value.ToString().Trim());
        //                //double qty = Convert.ToDouble(_RSBomP.Fields.Item("Quantity").Value.ToString());
        //                childDS.SetValue("U_Quantity", dsIndex, _RSBomP.Fields.Item("Quantity").Value.ToString());
        //                childDS.SetValue("U_BaseQty", dsIndex, (_RSBomP.Fields.Item("Quantity").Value / _RSBomP.Fields.Item("Qauntity").Value));
        //                childDS.SetValue("U_PlannedQty", dsIndex, (_RSBomP.Fields.Item("Quantity").Value / _RSBomP.Fields.Item("Qauntity").Value));
        //                childDS.SetValue("U_UOM", dsIndex, _RSBomP.Fields.Item("InvntryUom").Value.ToString().Trim());
        //                childDS.SetValue("U_Whse", dsIndex, _RSBomP.Fields.Item("Warehouse").Value.ToString().Trim());
        //                if (_RSBomP.Fields.Item("IssueMthd").Value.ToString().Trim() == "B")
        //                {
        //                    childDS.SetValue("U_IssueMethod", dsIndex, "BackFlush");
        //                }
        //                else
        //                {
        //                    childDS.SetValue("U_IssueMethod", dsIndex, "Manual");
        //                }
        //                _RSBomP.MoveNext();
        //            }
        //        _Matrix.LoadFromDataSource();
        //    }
        //    else 
        //    {
        //        //_Form.Freeze(true);
        //        var childDSs = _Form.DataSources.DBDataSources.Item("@PRODUCTIONPC");
        //        while (childDSs.Size > 0)
        //            childDSs.RemoveRecord(0);
        //        childDSs.InsertRecord(0);
        //        _Matrix.LoadFromDataSource();
        //        _Matrix.AddRow(1, -1);
        //        //_Form.Freeze(false);
        //    }

        //    _Form.Freeze(false);

        //}

        //#endregion

        //#region Send Alert
        //void SendInternalAlert(string docCode, List<string> approverUserCodes)
        //{
        //    SAPbobsCOM.Messages alert = (SAPbobsCOM.Messages)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages);

        //    alert.Subject = "Approval Required: UDO Document";
        //    alert.MessageText = $"A new document (Code: {docCode}) has been submitted and requires your approval.";

        //    for (int i = 0; i < approverUserCodes.Count; i++)
        //    {
        //        alert.Recipients.Add();
        //        alert.Recipients.SetCurrentLine(i);
        //        alert.Recipients.UserCode = approverUserCodes[i];
        //        //alert.Recipients.NameTo = approverUserCodes[i]; 
        //        alert.Recipients.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
        //    }

        //    int ret = alert.Add();
        //    if (ret != 0)
        //    {
        //        string errMsg;
        //        int errCode;
        //        Utilities.Application.Company.GetLastError(out errCode, out errMsg);
        //        Utilities.Application.SBO_Application.MessageBox("Error sending alert: " + errMsg);
        //    }
        //    else
        //    {
        //        Utilities.Application.SBO_Application.MessageBox("Approval request sent to: " + string.Join(", ", approverUserCodes));
        //    }
        //}
        //#endregion
        private string ShowFileDialog()
        {
            string selectedFile = null;
            Thread t = new Thread(() =>
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Title = "Select a file";
                dialog.Filter = "All files (*.*)|*.*";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFile = dialog.FileName;
                }
            });

            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();

            return selectedFile;
           

        }
        private AttachmentResult SaveAttachmentToSAP(string sourceFile)
        {
            AttachmentResult result = new AttachmentResult { AbsEntry = -1, DestPath = "" };

            try
            {
                string fileName = Path.GetFileNameWithoutExtension(sourceFile);
                string extension = Path.GetExtension(sourceFile).Replace(".", "");
                string attachPath = GetSAPAttachmentPath();
                string destPath = Path.Combine(attachPath, Path.GetFileName(sourceFile));
                result.DestPath = destPath;

                // Copy the file to SAP Attachments folder
                File.Copy(sourceFile, destPath, true);

                // Create SAP Attachment object
                SAPbobsCOM.Attachments2 oAttach = (SAPbobsCOM.Attachments2)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.oAttachments2);
                oAttach.Lines.SourcePath = attachPath;
                oAttach.Lines.FileName = fileName;
                oAttach.Lines.FileExtension = extension;

                if (oAttach.Add() == 0)
                {
                    result.AbsEntry = int.Parse(Utilities.Application.Company.GetNewObjectKey());
                }
                else
                {
                    Utilities.Application.Company.GetLastError(out int errCode, out string errMsg);
                    Utilities.Application.SBO_Application.StatusBar.SetText($"Attachment failed: {errMsg}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }

                Marshal.ReleaseComObject(oAttach);
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText($"Error in SaveAttachmentToSAP: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }

            return result;
        }
        private string GetSAPAttachmentPath()
        {
           
            
            Utilities.ExecuteSQL(ref _RSgetAttach, "SELECT \"AttachPath\" FROM \"OADP\"");
            string path = _RSgetAttach.Fields.Item(0).Value.ToString();
            Marshal.ReleaseComObject(_RSgetAttach);
            return path;
        }
        private void OpenAttachment(int absEntry)
        {
            SAPbobsCOM.Attachments2 attach = (SAPbobsCOM.Attachments2)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.oAttachments2);
            if (attach.GetByKey(absEntry))
            {
                string fullPath = Path.Combine(attach.Lines.SourcePath, attach.Lines.FileName + "." + attach.Lines.FileExtension);
                System.Diagnostics.Process.Start(fullPath);
            }
            Marshal.ReleaseComObject(attach);
        }
        public class AttachmentResult
        {
            public int AbsEntry { get; set; }
            public string DestPath { get; set; }
        }
        private void CalculateMatrixLogic()
        {
            SAPbouiCOM.Matrix matrix1 = (SAPbouiCOM.Matrix)_Form.Items.Item("Item_7").Specific;
            SAPbouiCOM.Matrix matrix3 = (SAPbouiCOM.Matrix)_Form.Items.Item("Item_9").Specific;
            SAPbouiCOM.Matrix matrix2 = (SAPbouiCOM.Matrix)_Form.Items.Item("Item_8").Specific;

            var ds1 = _Form.DataSources.DBDataSources.Item("@DLYWRKRPRTMMC1");
            var ds2 = _Form.DataSources.DBDataSources.Item("@DLYWRKRPRTMMC2");
            var ds3 = _Form.DataSources.DBDataSources.Item("@DLYWRKRPRTMMC3");

            double total = 0;
            string v1 = ds1.GetValue("U_Input", 1).Trim();
            string v2 = ds1.GetValue("U_Input", 2).Trim();
            string v3 = ds1.GetValue("U_Input", 3).Trim();
            string v4 = ds1.GetValue("U_Input", 4).Trim();
            string v5 = ds1.GetValue("U_Input", 5).Trim();
            string v6 = ds1.GetValue("U_Input", 6).Trim();
            string v8 = ds1.GetValue("U_Input", 8).Trim();
            string v14 = ds1.GetValue("U_Input", 14).Trim();

            string A1 = ds2.GetValue("U_Input", 0).Trim();
            string A3 = ds2.GetValue("U_Input", 1).Trim();
            string A2 = ds2.GetValue("U_Input", 2).Trim();
            //string v14 = ds1.GetValue("U_Input", 14).Trim();
            if (double.TryParse(v1, out double value))
            {
              total = (Convert.ToDouble(v2)/ Convert.ToDouble(v1) * Convert.ToDouble(v3));
              ds3.SetValue("U_Input", 0, Convert.ToString(total));
            }
            else
            {
                ds3.SetValue("U_Input", 0, "N/A");
            }
            if (double.TryParse(v4, out double value1))
            {
                total = (Convert.ToDouble(v4) / Convert.ToDouble(v5) * Convert.ToDouble(v6));
                ds3.SetValue("U_Input", 1, Convert.ToString(total));
            }
            else
            {
                ds3.SetValue("U_Input", 1, "N/A");
            }
            if (double.TryParse(v8, out double value2))
            {
                total = (Convert.ToDouble(v8) / 6 * 100);
                ds3.SetValue("U_Input", 2, Convert.ToString(total));
            }
            else
            {
                ds3.SetValue("U_Input", 2, "N/A");
            }
            if (double.TryParse(v14, out double value3))
            {
                total = (Convert.ToDouble(v14) / 6 * 100);
                ds3.SetValue("U_Input", 3, Convert.ToString(total));
            }
            else
            {
                ds3.SetValue("U_Input", 3, "N/A");
            }

            if (double.TryParse(A3, out double value5))
            {
                total = (Convert.ToDouble(A1) + Convert.ToDouble(A3));
                ds2.SetValue("U_Input", 2, Convert.ToString(total));
                A2 = total.ToString();
            }
            if (double.TryParse(A1, out double value4))
            {
                total = (Convert.ToDouble(A1) / Convert.ToDouble(A2) * 100);
                ds3.SetValue("U_Input", 4, Convert.ToString(total));
            }
            else
            {
                ds3.SetValue("U_Input", 4, "N/A");
            }

            

            matrix3.LoadFromDataSource();
        }
        private void CalculateMatrixLogic1()
        {
            SAPbouiCOM.Matrix matrix3 = (SAPbouiCOM.Matrix)_Form.Items.Item("Item_9").Specific;

            var ds1 = _Form.DataSources.DBDataSources.Item("@DLYWRKRPRTMMC1");
            var ds2 = _Form.DataSources.DBDataSources.Item("@DLYWRKRPRTMMC2");
            var ds3 = _Form.DataSources.DBDataSources.Item("@DLYWRKRPRTMMC3");

            try
            {
                double v1 = GetVal(ds1, 1);
                double v2 = GetVal(ds1, 2);
                double v3 = GetVal(ds1, 3);
                double v4 = GetVal(ds1, 4);
                double v5 = GetVal(ds1, 5);
                double v6 = GetVal(ds1, 6);
                double v8 = GetVal(ds1, 8);
                double v14 = GetVal(ds1, 14);

                double a1 = GetVal(ds2, 0);
                double a2 = GetVal(ds2, 2);

                // Set calculated values into ds3
                ds3.SetValue("U_Input", 0, v1 != 0 ? ((v2 / v1) * v3).ToString("0.00") : "N/A");
                ds3.SetValue("U_Input", 1, v5 != 0 ? ((v4 / v5) * v6).ToString("0.00") : "N/A");
                ds3.SetValue("U_Input", 2, ((v8 / 6) * 100).ToString("0.00"));
                ds3.SetValue("U_Input", 3, ((v14 / 6) * 100).ToString("0.00"));
                ds3.SetValue("U_Input", 4, a2 != 0 ? ((a1 / a2) * 100).ToString("0.00") : "N/A");
            }
            catch
            {
                // Optional: handle or log error
            }

            matrix3.LoadFromDataSource();
        }
        // Helper function
        private double GetVal(SAPbouiCOM.DBDataSource ds, int row)
        {
            string val = ds.GetValue("U_Input", row).Trim();
            return double.TryParse(val, out double result) ? result : 0;
        }
    }
}
