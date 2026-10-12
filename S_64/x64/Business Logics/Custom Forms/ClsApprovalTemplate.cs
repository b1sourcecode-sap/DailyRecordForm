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

namespace LGBOM_LaxmiMotors
{
    class ClsApprovalTemplate : Base
    {
        //#region VARIABLE DECLARATION

        //SAPbouiCOM.Item _ExistItem, _NewItem, _Items;
        //private bool flag;
        //private bool errorFlag;
        //public SAPbouiCOM.Matrix _Matrix;
        //SAPbouiCOM.StaticText _StaticText;
        //SAPbouiCOM.Button _Button;
        //string _DocEntry = "";
        //string _PRNum = string.Empty, _PRName = string.Empty, _RevNum = string.Empty, _Warehouse = string.Empty, _RevRmrks = string.Empty;
        //double _Qty = 0;
        //public string ErrMsg = string.Empty;
        //public int ErrCode = 0;
        //CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        //string ReportPath = string.Empty;
        //private static HanaConnection _conn;
        //private static HanaDataAdapter _da;
        ////private static DataTable _dataTable;
        //private static HanaParameter _parm1;
        //private static HanaParameter _parm2;
        //private static int _errors = 0;
        //string strcon, _Sql;
        //private SAPbouiCOM.EditText _EditText;
        //private SAPbouiCOM.LinkedButton _Linked;
        //private SAPbouiCOM.ComboBox _ComboBoxM;
        //private SAPbouiCOM.ComboBox _ComboBoxC, _ComboBoxCRow;
        //private SAPbouiCOM.ComboBox _ComboBoxB;
        //private SAPbouiCOM.ComboBox _ComboBoxL;
        //private SAPbouiCOM.ComboBox _ComboBoxH;
        //private SAPbouiCOM.DBDataSource _MasterDBDataSource = null, _ChildDBDataSource = null;
        //private SAPbouiCOM.ChooseFromListCollection _CFLS;
        //private SAPbouiCOM.ChooseFromList _CFL;
        //private SAPbouiCOM.ChooseFromListCreationParams _CFLC;
        //private SAPbouiCOM.Conditions _Cons;
        //private SAPbouiCOM.Condition _Con;
        //private int RowID { get; set; }
        //private string _Code { get; set; }
        //private string _LineId { get; set; }
        //private string _ActivityType = string.Empty;
        //string _ItemCode = string.Empty, _ItemName = string.Empty;
        //private SAPbobsCOM.Recordset _Rs, _RS, _rs, _RSD, _RSBom, _RSProd, _RSRoute, _RSUsr, _RSMgItm, _RSBomP;
        //SAPbouiCOM.ChooseFromListEvent _SysCFLEvent = null;
        //SAPbouiCOM.DBDataSource _DBData = null;
        //private System.Collections.Generic.List<string> ListOfBlockedID;
        //internal string _ServiceID { get; set; }
        //internal string _ActivityTypes { get; set; }
        //public bool chkStatus = true;
        //static int StageID = 0;
        //#endregion

        //#region CONSTRUCTOR & DISTRUCTOR
        //public ClsApprovalTemplate()
        //    : base()
        //{
        //    //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
        //    //
        //    // TODO: Add constructor logic here
        //    //
        //}
        //~ClsApprovalTemplate()
        //{
        //}
        //#endregion
        //#region ItemEvent
        //public override void Item_Event(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
        //{
        //    #region Before Action
        //    if (pVal.BeforeAction)
        //    {
        //        switch (pVal.EventType)
        //        {
                   
        //        }
        //    }
        //    #endregion

        //    #region After Action
        //    else
        //    {
        //        switch (pVal.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                       
        //                if (pVal.ItemUID == "btnApprove")
        //                {
        //                    ProcessSelections("Approved");
        //                }
        //                if (pVal.ItemUID == "btnReject")
        //                {
        //                    ProcessSelections("Rejected");
        //                }
        //                if (pVal.ItemUID == "btnPending")
        //                {
        //                    ProcessSelections("Pending");
        //                }
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:

        //                //var matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("matrix").Specific;
        //                //SAPbouiCOM.Column column = matrix.Columns.Item("colDoc");
        //                //SAPbouiCOM.LinkedButton linkedButton = (SAPbouiCOM.LinkedButton)column.ExtendedObject;
        //                //linkedButton.LinkedObject = (BoLinkedObject)SAPbouiCOM.BoUDOObjectType.udo_Document;
        //                //linkedButton.LinkedObjectType = "ProductionP"; 

        //                if (pVal.EventType == BoEventTypes.et_MATRIX_LINK_PRESSED)
        //                {
        //                    if (pVal.ColUID == "colDoc")
        //                    {
        //                        SAPbouiCOM.Form form = Utilities.Application.SBO_Application.Forms.Item(pVal.FormUID);
        //                        SAPbouiCOM.Matrix matrix = (SAPbouiCOM.Matrix)form.Items.Item("matrix").Specific;

        //                        string docEntry = ((SAPbouiCOM.EditText)matrix.Columns.Item("colDoc").Cells.Item(pVal.Row).Specific).Value;

        //                        if (!string.IsNullOrEmpty(docEntry))
        //                        {
        //                            try
        //                            {
        //                                Utilities.Application.SBO_Application.OpenForm(BoFormObjectEnum.fo_UserDefinedObject, "ProductionP", docEntry);
        //                                form.Close(); 
        //                            }
        //                            catch (Exception ex)
        //                            {
        //                                Utilities.Application.SBO_Application.StatusBar.SetText($"Error opening UDO form: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //                            }
        //                        }
        //                    }
        //                }


        //                break;
        //        }
        //    }

        //    #endregion
        //}
        //#endregion

        //#region FORM DEFAULT
        //internal void FormDefault()
        //{
        //    string userCode = Utilities.Application.SBO_Application.Company.UserName;
        //    Utilities.ExecuteSQL(ref _RSUsr, $"SELECT \"U_StatusAuth\" FROM \"OUSR\" WHERE \"USER_CODE\" = '{userCode.Replace("'", "''")}'");


        //    // Step 4: Check authorization
        //    if (!_RSUsr.EoF && _RSUsr.Fields.Item("U_StatusAuth").Value.ToString() == "YES")
        //    {
        //        LoadApprovalData();
        //    }
        //    else
        //    {
        //        // Show error message
        //        Utilities.Application.SBO_Application.MessageBox("You are not an approver.", 1, "OK", "", "");
        //    }

        //    // Cleanup
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(_RSUsr);
        //    _RSUsr = null;
          
        //}

        //private void LoadApprovalData()
        //{
        //    var matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("matrix").Specific;

        //    // Step 1: Prepare DataTable (with safe check)
        //    SAPbouiCOM.DataTable dt;

        //    try
        //    {
        //        dt = _Form.DataSources.DataTables.Item("dtDisplay");
        //        dt.Clear();  // reuse if already exists
        //    }
        //    catch
        //    {
        //        // create if doesn't exist
        //        dt = _Form.DataSources.DataTables.Add("dtDisplay");
        //        dt.Columns.Add("U_Selected", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 1);
        //        dt.Columns.Add("U_DocEntry", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
        //        dt.Columns.Add("U_PRNum", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
        //        dt.Columns.Add("U_Status", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
        //        dt.Columns.Add("U_Remarks", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254);
        //        dt.Columns.Add("U_CretedBy", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
        //        dt.Columns.Add("U_ARemarks", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254);
        //    }

        //    // Step 2: Bind Matrix Columns to DataTable
        //    matrix.Columns.Item("colDoc").DataBind.Bind("dtDisplay", "U_DocEntry");
        //    matrix.Columns.Item("colCode").DataBind.Bind("dtDisplay", "U_PRNum");
        //    matrix.Columns.Item("colStatus").DataBind.Bind("dtDisplay", "U_Status");
        //    matrix.Columns.Item("colRmrks").DataBind.Bind("dtDisplay", "U_Remarks");
        //    matrix.Columns.Item("colCretBy").DataBind.Bind("dtDisplay", "U_CretedBy");
        //    matrix.Columns.Item("colAPRmrks").DataBind.Bind("dtDisplay", "U_ARemarks");

        //    // Step 3: Query Data from UDO Table
        //    string sql = @"
        //    SELECT  ""DocEntry"",""U_PRNum"", ""U_Status"", ""U_Remarks"", ""U_CretedBy"",""U_ARemarks""
        //    FROM ""@PRODUCTIONPH""
        //    WHERE ""U_Status"" = 'Draft'";

        //    var rs = (SAPbobsCOM.Recordset)Utilities.Application.Company
        //                 .GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    rs.DoQuery(sql);

        //    // Step 4: Load Data into DataTable
        //    while (!rs.EoF)
        //    {
        //        int row = dt.Rows.Count;
        //        dt.Rows.Add();
        //        dt.SetValue("U_DocEntry", row, rs.Fields.Item("DocEntry").Value.ToString());
        //        dt.SetValue("U_PRNum", row, rs.Fields.Item("U_PRNum").Value.ToString());
        //        dt.SetValue("U_Status", row, rs.Fields.Item("U_Status").Value.ToString());
        //        dt.SetValue("U_Remarks", row, rs.Fields.Item("U_Remarks").Value.ToString());
        //        dt.SetValue("U_CretedBy", row, rs.Fields.Item("U_CretedBy").Value.ToString());
        //        dt.SetValue("U_ARemarks", row, rs.Fields.Item("U_ARemarks").Value.ToString());

        //        rs.MoveNext();
        //    }

        //    // Step 5: Load Matrix from DataTable
        //    matrix.LoadFromDataSource();
        //}

        //#endregion
        //private void ProcessSelections(string newStatus)
        //{
        //    var matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("matrix").Specific;
        //    var dt = _Form.DataSources.DataTables.Item("dtDisplay");

        //    // 1) Get the CompanyService and GeneralService for your UDO
        //    var companyService = Utilities.Application.Company.GetCompanyService();
        //    var svc = (SAPbobsCOM.GeneralService)companyService.GetGeneralService("ProductionP");

        //    // 2) Retrieve the correct GeneralDataParams interface
        //    SAPbobsCOM.GeneralDataParams generalParams;
        //    try
        //    {
        //        generalParams = (SAPbobsCOM.GeneralDataParams)svc.GetDataInterface(
        //            SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
        //    }
        //    catch
        //    {
        //        // Fallback if enum member differs by DI API version
        //        generalParams = (SAPbobsCOM.GeneralDataParams)svc.GetDataInterface(
        //            SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralCollectionParams);
        //    }

        //    // 3) Iterate selected rows and update UDO
        //        for (int i = 1; i <= matrix.RowCount; i++)
        //        {
        //            var chkColumn = matrix.Columns.Item("chkSelect");
        //            var cell = chkColumn.Cells.Item(i).Specific;

        //            bool isSelected = false;
        //            if (chkColumn.Type == SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
        //            {
        //                var chk = (SAPbouiCOM.CheckBox)cell;
        //                isSelected = chk.Checked;
        //            }
        //            else if (chkColumn.Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        //            {
        //                var combo = (SAPbouiCOM.ComboBox)cell;
        //                isSelected = combo.Selected.Value == "Y";
        //            }

        //        if (isSelected)
        //        {
                   
        //            var matrixCol = matrix.Columns.Item("colDoc");
        //            var cellSpec = matrixCol.Cells.Item(i).Specific;

                   
        //            var txt = (SAPbouiCOM.EditText)cellSpec;
        //            string docEntryStr = txt.Value?.Trim();

                   
        //            if (!int.TryParse(docEntryStr, out int docEntry))
        //            {
        //                Utilities.Application.SBO_Application.MessageBox($"Row {i}: Invalid DocEntry value '{docEntryStr}'");
        //                continue;
        //            }
        //            generalParams.SetProperty("DocEntry", docEntry);
        //            var data = svc.GetByParams(generalParams);
        //            data.SetProperty("U_Status", newStatus);
        //            var remarksCell = (SAPbouiCOM.EditText)matrix.Columns.Item("colAPRmrks").Cells.Item(i).Specific;
        //            string remarksValue = remarksCell.Value?.Trim() ?? "";
        //            data.SetProperty("U_ARemarks", remarksValue);
        //            svc.Update(data);
        //        }

        //    }
        //    Utilities.Application.SBO_Application.MessageBox($"{newStatus} successfully applied.");
        //    _Form.Refresh();
        //}



    }
}

