using System;
using SAPbobsCOM;
using SAPbouiCOM;
using System.IO;
using System.Data;
using System.Collections;
using System.Windows.Forms;
using System.IO.Ports;
using System.Diagnostics;
using System.Text;

namespace Production_Planning_LMC
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
        private SerialPort _scannerPort;
        string FirstScan = string.Empty;
        string SecondScan = string.Empty;
        int scancount = 0;
        private SAPbobsCOM.Recordset _RSWorkId, _RSRoute, _RSWorkExecDetails, _RSWorkOrderDetails, _RSD, _RSBom, _RSPP;
        string route = string.Empty;
        string routeId = string.Empty;
        string WorkId = string.Empty;
        DateTime LDurationTime;
        private bool _isSystemClear = false;
        string productionOrderEntry = string.Empty;
        public SAPbouiCOM.Matrix _Matrix;
        public SAPbouiCOM.Form oForm;
        private static SAPbouiCOM.Form oScanForm;
        private static SAPbouiCOM.EditText oScanTxt;
        private string _scanBuffer = "";
        private string routeDsp, productName, model;
        System.Windows.Forms.Timer focusTimer = new System.Windows.Forms.Timer();
        private SAPbouiCOM.Form _oForm;
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
                    //OpenUDOScanForm();
                    CreateHiddenListenerForm();
                    StartFocusTimer(oForm);
                    ForceScanFocus();
                    //OpenTrimAssemblyStatusBoard();
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
                        //case Constants.User_Menus.MENU_LGProduction:
                        //     this._Object = new ClsLGProductionOrder();
                        //     Utilities.LoadForm(ref this._Object, Constants.Forms.ProductionPlanning);
                        //     this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                        //     ((ClsLGProductionOrder)_Object).FormDefault();
                        //     break;
                        // case Constants.User_Menus.MENU_LGApproval:
                        //     this._Object = new ClsApprovalTemplate();
                        //     Utilities.LoadForm(ref this._Object, Constants.Forms.ApprovalTemplate);
                        //     this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                        //     ((ClsApprovalTemplate)_Object).FormDefault();
                        //     break;
                        // case Constants.User_Menus.MENU_ABPC:
                        //     this._Object = new ClsApprovedBPCreation();
                        //     Utilities.LoadForm(ref this._Object, Constants.Forms.ApprovedBPCreation);
                        //     this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                        //     ((ClsApprovedBPCreation)_Object).FormDefault();
                        //     break;
                        //case Constants.User_Menus.MENU_ABPCApproval:
                        //    this._Object = new ClsApprovalTemplate();
                        //    Utilities.LoadForm(ref this._Object, Constants.Forms.ABPTemplate);
                        //    this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                        //    ((ClsApprovalTemplate)_Object).FormDefault();
                        //    break;
                        case Constants.User_Menus.MENU_EngineChasis:
                            this._Object = new ProductionPlanningLMC();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.EngineChasisMapping);
                            this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                            ((ProductionPlanningLMC)_Object).FormDefault();
                            break;
                        case Constants.User_Menus.MENU_WorkOrderDetailss:
                            this._Object = new ClsWorkOrderDetails();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.WorkOrderDetailss);
                            this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                            ((ClsWorkOrderDetails)_Object).FormDefault();
                            break;

                        case Constants.User_Menus.MENU_JobOrderExecutions:
                            this._Object = new ClsJobOrderExecution();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.JobOrderExecutions);
                            this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                            ((ClsJobOrderExecution)_Object).FormDefault();
                            break;
                        case Constants.User_Menus.MENU_LOTMASTER:
                            Utilities.Application.SBO_Application.Menus.Item("LOTNOMASTER").Activate();
                            break;
                        case Constants.User_Menus.MENU_OCNMASTER:
                            Utilities.Application.SBO_Application.Menus.Item("47644").Activate();
                            break;
                        case Constants.User_Menus.MENU_PRODUCTIONORDER:
                            Utilities.Application.SBO_Application.Menus.Item("4369").Activate();
                            break;
                        case Constants.User_Menus.MENU_INVENTORYTRANSFERREQ:
                            Utilities.Application.SBO_Application.Menus.Item("3088").Activate();
                            break;
                        case Constants.User_Menus.MENU_INVENTORYTRANSFER:
                            Utilities.Application.SBO_Application.Menus.Item("3080").Activate();
                            break;
                        case Constants.User_Menus.MENU_ISSUEPROD:
                            Utilities.Application.SBO_Application.Menus.Item("4371").Activate();
                            break;
                        case Constants.User_Menus.MENU_RECEIPTPROD:
                            Utilities.Application.SBO_Application.Menus.Item("4370").Activate();
                            break;


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

                        case Constants.System_Forms.Production_ORDER:
                            this._Object = new ClsProductionOrderr();
                            _Collection.Add(FormUID, this._Object);
                            ((ClsProductionOrderr)this._Object).Form = _Application.Forms.Item(FormUID);
                            ((ClsProductionOrderr)this._Object).FormDefault();
                            break;
                    }
                }

                if (this._Collection.Contains(FormUID))
                {
                    this._Object = (Production_Planning_LMC.Base)this._Collection[FormUID];

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
                            this._Object = (Production_Planning_LMC.Base)_Collection[_LookCollection[FormUID]];
                            _LookCollection.Remove(FormUID);
                            this._Object.IsLookUpOpen = false;
                        }
                        this._Object.Dispose();
                        _Collection.Remove(FormUID);
                        System.GC.Collect();
                    }
                }

                if (FormUID == "frmBarScan" && pVal.ItemUID == "txtScan" && !pVal.BeforeAction)
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
                    {
                        ForceScanFocus(); 
                    }
                }

                /////////////////////////
                if (oScanTxt == null) return;

                if (FormUID == "frmBarScan" && pVal.ItemUID == "txtScan" && !pVal.BeforeAction)
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN)
                    {
                        int key = pVal.CharPressed;

                        // 2️⃣ Ignore System Keys (Shift, Ctrl, etc which are -1 or 0)
                        if (key <= 0) return;

                        // 3️⃣ Check for ENTER (13)
                        if (key == 13 || key == 10)
                        {
                            try
                            {
                                // ✅ Process the accumulated buffer, NOT oScanTxt.Value
                                string finalBarcode = _scanBuffer.Trim();

                                if (!string.IsNullOrEmpty(finalBarcode))
                                {
                                    // --- YOUR LOGIC HERE ---
                                    Utilities.Application.SBO_Application.StatusBar.SetText(
                                        "Processing: " + finalBarcode,
                                        SAPbouiCOM.BoMessageTime.bmt_Short,
                                        SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                                    // Clear Buffer & TextBox after success
                                    // _scanBuffer = "";
                                    try { oScanTxt.Value = ""; } catch { }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Error: " + ex.Message);
                                //  _scanBuffer = ""; // Reset on error
                            }

                            // Focus ready for next scan
                            ForceScanFocus();
                        }
                        // 4️⃣ Check for Backspace (8) - Optional but good for manual correction
                        else if (key == 8)
                        {
                            if (_scanBuffer.Length > 0)
                                _scanBuffer = _scanBuffer.Substring(0, _scanBuffer.Length - 1);
                        }
                        // 5️⃣ Accumulate Characters (If not Enter)
                        else
                        {
                            // Convert ASCII code to Character and add to buffer
                            _scanBuffer += (char)key;

                        }
                    }
                }
                //////////////////////

                #region Scanning Chassis No and its respective Actions
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && pVal.BeforeAction == false && pVal.InnerEvent == true)
                {
                    SAPbouiCOM.Form _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                    if (pVal.FormTypeEx == "SCAN_BAR" && pVal.ItemUID == "txtScan" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && pVal.BeforeAction == false)
                    {

                        SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)_Form.Items.Item("txtScan").Specific;
                        //SAPbouiCOM.Item addBtn = _Form.Items.Item("1"); // The Add Button (UI ID 1)

                        string scannedValue = _scanBuffer;

                        if (string.IsNullOrEmpty(scannedValue))
                        {
                            return;
                        }

                        // --- FIRST SCAN ---
                        if (scancount == 0)
                        {
                            // 1. Process the First Scan String
                            //string first = scannedValue.ToUpper();
                            //if (first.Length > 1)
                            //{
                            //    first = first.Substring(0, first.Length - 1) + "*" + first.Substring(first.Length - 1);
                            //}

                            string first = scannedValue.ToUpper();

                            if (first.Length > 1)
                            {
                                int secondLastIndex = first.Length - 2;

                                // Case 1: second last character is '8' → replace with '*'
                                if (first.Length == 19)
                                {
                                    if (first[secondLastIndex] == '8')
                                    {
                                        first = first.Substring(0, secondLastIndex) + "*" + first.Substring(secondLastIndex + 1);
                                    }
                                }
                                else if (first.Length == 18) {
                                // Case 2: no '*' exists → insert '*' before last character
                                if (!first.Contains("*"))
                                    {
                                        first = first.Substring(0, first.Length - 1) + "*" + first.Substring(first.Length - 1);
                                    }
                                }
                            }


                            FirstScan = first;
                            scancount = 1;
                            OpenScanInBrowser(first);
                            // 2. Clear the field visually so it is ready for Second Scan
                            _scanBuffer = "";

                            // IMPORTANT: Stop the event here. 
                            // This keeps the focus in Item "14" and prevents SAP from checking OITM table.
                            BubbleEvent = false;
                            ForceScanFocus();

                            // Show a small status message so user knows Scan 1 was accepted
                            // SBO_Application.MessageBox("First scan accepted. Please scan secondary code...");
                            //CloseSystemPopupNo();

                        }
                        // --- SECOND SCAN ---
                        else if (scancount == 1)
                        {
                            //
                            SecondScan = scannedValue.ToUpper().Replace(" ", "_");
                            
                            //SecondScan = scannedValue.ToUpper();
                            SAPbobsCOM.Recordset _RSWorkId = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string query = $@"SELECT ""U_FirstBarcode"", ""U_LastBarcode"", ""U_JobPause"", ""U_JobResume""  
                                             FROM ""@WRKORDRDTLSC"" WHERE ""U_ChasisNo"" = '{FirstScan}'";
                            _RSWorkId.DoQuery(query);
                            string matchedColumn = "";
                            if (_RSWorkId.RecordCount > 0)
                            {
                                while (!_RSWorkId.EoF)  // loop rows
                                {
                                    for (int i = 0; i < _RSWorkId.Fields.Count; i++)
                                    {
                                        var field = _RSWorkId.Fields.Item(i);
                                        string fieldValue = field.Value != null ? field.Value.ToString().Trim() : "";

                                        if (!string.IsNullOrEmpty(fieldValue) && fieldValue.Equals(SecondScan, StringComparison.OrdinalIgnoreCase))
                                        {
                                            matchedColumn = field.Name;
                                            break;
                                        }

                                    }
                                    if (!string.IsNullOrEmpty(matchedColumn))
                                        break; // stop searching if match is found

                                    _RSWorkId.MoveNext(); // go to next row
                                }
                                if (string.IsNullOrEmpty(matchedColumn))
                                {
                                    ShowBrowserMessage("❌", "INVALID ACTION", FirstScan, "", "", "", DateTime.Now.ToString("HH:mm:ss"), "#ff5252", "ACTION: Barcode not recognized for this chassis. Scan valid action.");
                                    //SBO_Application.MessageBox("Invalid Barcode! This code is not associated with the scanned Chassis actions.");
                                    _scanBuffer = "";
                                    scancount = 0; // Reset to start
                                    BubbleEvent = false;
                                    return;
                                }


                                SAPbobsCOM.Recordset _RSRoute = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                string query_route = $@"SELECT T0.""LineId"",T0.""U_RouteDsp"",T0.""U_RouteId"", T0.""U_WorkId"", T1.""U_PrdOrdEnt"", T0.""U_RouteIdNum"",T1.""U_ProductName"",T1.""U_Model""
                                                     FROM ""@WRKORDRDTLSC"" T0 INNER JOIN ""@WRKORDRDTLSH"" T1 on T0.""DocEntry"" = T1.""DocEntry"" 
                                                     WHERE T0.""U_ChasisNo"" = '{FirstScan}' and ""{matchedColumn}""='{SecondScan}'";
                                _RSRoute.DoQuery(query_route);

                                if (_RSRoute.RecordCount > 0)
                                {
                                   
                                     route = _RSRoute.Fields.Item("U_RouteDsp").Value;
                                    //routeId = _RSRoute.Fields.Item("U_RouteIdNum").Value;
                                    //routeId = _RSRoute.Fields.Item("U_RouteId").Value;
                                    string rawRouteId = _RSRoute.Fields.Item("U_RouteId").Value.ToString();

                                    // Check if it's not empty to avoid errors
                                    if (!string.IsNullOrEmpty(rawRouteId))
                                    {
                                        // Split by '-' and take the first item (Index 0), then Trim whitespace
                                        // Result: "1"
                                        routeId = rawRouteId.Split('-')[0].Trim();
                                    }
                                    else
                                    {
                                        routeId = "";
                                    }
                                    WorkId = _RSRoute.Fields.Item("U_WorkId").Value;
                                    productionOrderEntry = _RSRoute.Fields.Item("U_PrdOrdEnt").Value.ToString();
                                     routeDsp = _RSRoute.Fields.Item("U_RouteDsp").Value.ToString();
                                    productName = _RSRoute.Fields.Item("U_ProductName").Value.ToString();
                                    model = _RSRoute.Fields.Item("U_Model").Value.ToString();

                                    int currentLineId = Convert.ToInt32(_RSRoute.Fields.Item("U_RouteIdNum").Value);
                                    if (Convert.ToInt32(currentLineId) > 1)
                                    {
                                        int previousLineId = currentLineId - 1;

                                        SAPbobsCOM.Recordset _RSCheckPrev = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                        // Query to check if the previous line has a value in U_IssueForProdEntry
                                        // We use IFNULL/NULLIF to handle both NULLs and Empty strings as Empty
                                        string query_prev = $@"
                                        SELECT IFNULL(NULLIF(T0.""U_IssueForProdEntry"", ''), '') AS ""PrevIssue"", ""U_RouteDsp""
                                        FROM ""@WRKORDRDTLSC"" T0
                                        INNER JOIN  ""@WRKORDRDTLSH"" T1 on T0.""DocEntry"" = T1.""DocEntry"" 
                                        WHERE T0.""U_ChasisNo"" = '{FirstScan}'
                                          AND T0.""U_RouteIdNum"" = {previousLineId} ";

                                        _RSCheckPrev.DoQuery(query_prev);


                                        //issue check baseEntry equals production docentry
                                        SAPbobsCOM.Recordset _RSCheckIssue = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        if (_RSCheckPrev.Fields.Item("PrevIssue").Value.ToString().Trim() != "" )
                                        {
                                           
                                            string issueBaseEntry = $@"
                                        SELECT ""BaseEntry""
                                        FROM IGE1
                                        WHERE ""DocEntry"" = '{_RSCheckPrev.Fields.Item("PrevIssue").Value.ToString().Trim()}'";
                                            _RSCheckIssue.DoQuery(issueBaseEntry);
                                        }

                                        if (_RSCheckPrev.RecordCount > 0)
                                        {
                                            string prevIssueEntry = _RSCheckPrev.Fields.Item("PrevIssue").Value.ToString().Trim();

                                            // 3. Block logic if previous entry is empty
                                            if (string.IsNullOrEmpty(prevIssueEntry))
                                            {
                                                ShowBrowserMessage("⚠️", "SEQUENCE ERROR", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#ffea00", $"ACTION: Previous step [{_RSCheckPrev.Fields.Item("U_RouteDsp").Value}] must be completed first.");
                                               // SBO_Application.MessageBox(
                                                //    $"Sequence Error: The previous operation ({_RSCheckPrev.Fields.Item("U_RouteDsp").Value}) is not completed yet.");

                                                _scanBuffer = "";
                                                scancount = 0;
                                                BubbleEvent = false;

                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(_RSCheckPrev);
                                                return;
                                            }
                                        }
                                        if(_RSCheckIssue.RecordCount > 0)
                                        {
                                            string BaseEntry = _RSCheckIssue.Fields.Item("BaseEntry").Value.ToString().Trim();
                                            string prodDocEntry = _RSRoute.Fields.Item("U_PrdOrdEnt").Value.ToString().Trim();
                                            if (BaseEntry != prodDocEntry)
                                            {
                                                ShowBrowserMessage("❌", "INVALID ISSUE FOR PRODUCTION", FirstScan, "", "", "", DateTime.Now.ToString("HH:mm:ss"), "#ff5252", "ACTION: Invalid Issue for Production.");
                                               

                                                _scanBuffer = "";
                                                scancount = 0;
                                                BubbleEvent = false;

                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(_RSCheckIssue);
                                                return;
                                            }
                                        }
                                    }
                                    //OpenScanInBrowser(SecondScan,FirstScan,routeDsp,productName,model);
                                    //OpenTrimAssemblyStatusBoard();
                                }
                                string query_Execution = $@"SELECT ""DocEntry"",""U_JobID"",""U_JobDesc"",""U_ProdOrdNo"",""U_WorkOrdNo"",
                                                   ""U_FGCode"",""U_FGDesc"",""U_EngineNo"",""U_ChassisNo"",""U_BatchNo"",
                                                   ""U_Qty"",""U_Status"",""U_StartDate"",""U_StartTime"",""U_EndDate"",
                                                   ""U_EndTime"",""U_TotBrkTime"",""U_TotActTime"",""U_Operator""
                                            FROM ""@WJOBEXEH""
                                            WHERE ""U_JobID"" = '{WorkId}'";

                                SAPbobsCOM.Recordset _RSWorkExecDetails = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                _RSWorkExecDetails.DoQuery(query_Execution);
                                if (_RSWorkExecDetails.RecordCount > 0)
                                {
                                    string currentStatus = _RSWorkExecDetails.Fields.Item("U_Status").Value.ToString();

                                    // --- Load UDO Service ---
                                    SAPbobsCOM.CompanyService oCompanyService = Utilities.Application.Company.GetCompanyService();
                                    SAPbobsCOM.GeneralService oGeneralService = oCompanyService.GetGeneralService("WJOBEXE");
                                    SAPbobsCOM.GeneralDataParams oParams =
                                        (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(
                                            SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);

                                    oParams.SetProperty("DocEntry", _RSWorkExecDetails.Fields.Item("DocEntry").Value.ToString());
                                    SAPbobsCOM.GeneralData oGeneralData = oGeneralService.GetByParams(oParams);

                                    SAPbobsCOM.GeneralDataCollection oChildren = oGeneralData.Child("WJOBEXEC");
                                    SAPbobsCOM.GeneralData oChild;

                                    DateTime now = DateTime.Now;

                                    #region When current status is "Start" (Can add a pause or Stop)
                                    if (currentStatus == "Start")
                                    {
                                        if (matchedColumn == "U_JobPause")
                                        {
                                            oGeneralData.SetProperty("U_Status", "Pause");
                                            // Create Pause Row
                                            oChild = oChildren.Add();
                                            oChild.SetProperty("U_BrkStartDt", now);
                                            oChild.SetProperty("U_BrkStartTm", now);

                                            //SBO_Application.StatusBar.SetText("Job Paused",
                                            //    SAPbouiCOM.BoMessageTime.bmt_Short,
                                            //    SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                            //SBO_Application.MessageBox("Job Paused!");
                                            //ShowBrowserMessage("⏸️", "JOB PAUSED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#ffea00", "ACTION: Job is suspended. Scan RESUME to continue.");
                                             ShowBrowserMessage("✅", "JOB PAUSED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#ffea00", "ACTION: Job is suspended. Scan RESUME to continue.");
                                        }

                                        else if(matchedColumn == "U_LastBarcode")
                                        {
                                            SAPbouiCOM.ProgressBar pb = null;
                                            try
                                            {
                                                pb = SBO_Application.StatusBar.CreateProgressBar("Creating Production Issue........", 1, false);
                                                int goodIssue = Create_IssueFor_Production(_RSRoute.Fields.Item("U_PrdOrdEnt").Value.ToString(), routeId, WorkId, FirstScan);
                                                //int goodIssue = 790;

                                                if (goodIssue > 0)
                                                {
                                                    // Only update Status and Times if Goods Issue succeeded
                                                    oGeneralData.SetProperty("U_Status", "Stop");
                                                    oGeneralData.SetProperty("U_EndDate", now);
                                                    oGeneralData.SetProperty("U_EndTime", now);
                                                    //oGeneralData.SetProperty("U_TotBrkTime", new DateTime(1899, 12, 30)); // 00:00:00 No breaks
                                                    oGeneralData.SetProperty("U_TotBrkTime", "00:00");
                                                    oGeneralData.SetProperty("U_IForProdEntry", goodIssue.ToString());

                                                    // Calculate Actual Time Consumed
                                                    //double totalActMinutes = 0;
                                                    //try
                                                    //{
                                                    //    object hStartDt = oGeneralData.GetProperty("U_StartDate");
                                                    //    object hStartTm = oGeneralData.GetProperty("U_StartTime");

                                                    //    object cEndDt = oGeneralData.GetProperty("U_EndDate");
                                                    //    object cEndTm = oGeneralData.GetProperty("U_EndTime");

                                                    //    DateTime dtHeaderStart = GetFullDateTime(hStartDt, hStartTm);
                                                    //    DateTime dtLastChildEnd = GetFullDateTime(cEndDt, cEndTm);

                                                    //    if (dtHeaderStart != DateTime.MinValue)
                                                    //        totalActMinutes = (dtLastChildEnd - dtHeaderStart).TotalMinutes;
                                                    //}
                                                    //catch { }

                                                    //oGeneralData.SetProperty("U_TotActTime", new DateTime(1899, 12, 30).AddMinutes(totalActMinutes < 0 ? 0 : totalActMinutes));
                                                    double totalActMinutes = 0;
                                                    string durationString = "";
                                                    try
                                                    {
                                                        DateTime dtHeaderStart = GetFullDateTime(oGeneralData.GetProperty("U_StartDate"), oGeneralData.GetProperty("U_StartTime"));
                                                        DateTime dtLastChildEnd = GetFullDateTime(oGeneralData.GetProperty("U_EndDate"), oGeneralData.GetProperty("U_EndTime"));

                                                        double totalElapsedMinutes = (dtLastChildEnd - dtHeaderStart).TotalMinutes;
                                                        totalActMinutes = totalElapsedMinutes;

                                                        int hourss = (int)(totalActMinutes / 60);
                                                        int minss = (int)(totalActMinutes % 60);
                                                        durationString = $"{hourss:D2}:{minss:D2}";
                                                    }
                                                    catch { }


                                                    //if alpha
                                                    oGeneralData.SetProperty("U_TotActTime", durationString);

                                                    // NOTE: Visually in SAP, this will show "01:16", but the underlying value 
                                                    // will contain the extra day (1899-12-31) which is important for the next step.
                                                    //oGeneralData.SetProperty("U_TotActTime", new DateTime(1899, 12, 30).AddMinutes(totalActMinutes));


                                                    // Update External Tables 
                                                    try
                                                    {
                                                        //object startDate = oGeneralData.GetProperty("U_StartDate");
                                                        //object startTime = oGeneralData.GetProperty("U_StartTime");
                                                        //object endDate = oGeneralData.GetProperty("U_EndDate");
                                                        //object endTime = oGeneralData.GetProperty("U_EndTime");
                                                        //object totalBrkTime = oGeneralData.GetProperty("U_TotBrkTime");
                                                        //object actualTimeConsumed = oGeneralData.GetProperty("U_TotActTime");

                                                        DateTime dtStartDate = (DateTime)oGeneralData.GetProperty("U_StartDate");
                                                        DateTime dtStartTime = (DateTime)oGeneralData.GetProperty("U_StartTime");
                                                        DateTime dtEndDate = (DateTime)oGeneralData.GetProperty("U_EndDate");
                                                        DateTime dtEndTime = (DateTime)oGeneralData.GetProperty("U_EndTime");
                                                        //DateTime dtBrkTime = (DateTime)oGeneralData.GetProperty("U_TotBrkTime");
                                                        //DateTime dtActTime = (DateTime)oGeneralData.GetProperty("U_TotActTime");
                                                        string strBrkTime = Convert.ToString(oGeneralData.GetProperty("U_TotBrkTime"));
                                                        string strActTime = Convert.ToString(oGeneralData.GetProperty("U_TotActTime"));

                                                        string sqlStartDate = dtStartDate.ToString("yyyyMMdd");
                                                        string sqlEndDate = dtEndDate.ToString("yyyyMMdd");
                                                        string sqlStartTime = dtStartTime.ToString("HHmm");
                                                        string sqlEndTime = dtEndTime.ToString("HHmm");
                                                        //string sqlBrkTime = dtBrkTime.ToString("HHmm");
                                                        //string sqlActTime = dtActTime.ToString("HHmm");
                                                        string sqlBrkTime = strBrkTime.Replace(":", "").PadLeft(4, '0');
                                                        string sqlActTime = strActTime.Replace(":", "").PadLeft(4, '0');


                                                        /// for exceeding 24 hours 
                                                        DateTime dtHeaderStart = GetFullDateTime(oGeneralData.GetProperty("U_StartDate"), oGeneralData.GetProperty("U_StartTime"));
                                                        DateTime dtLastChildEnd = GetFullDateTime(oGeneralData.GetProperty("U_EndDate"), oGeneralData.GetProperty("U_EndTime"));

                                                        double totalElapsedMinutes = (dtLastChildEnd - dtHeaderStart).TotalMinutes;
                                                        totalActMinutes = totalElapsedMinutes;

                                                        int totalHrs = (int)(totalActMinutes / 60);
                                                        int totalMns = (int)(totalActMinutes % 60);

                                                        // Use a string format for the SQL update
                                                        // This will result in "2516" instead of "0116"
                                                        string manualTimeStr = $"{totalHrs:D2}:{totalMns:D2}";
                                                        /// end for exceeding 24 hours


                                                        SAPbobsCOM.Recordset rsUpdate = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                                        //string updateQuery = $@"UPDATE ""@WRKORDRDTLSC"" 
                                                        //    SET ""U_IssueForProdEntry"" = '{goodIssue}' , ""U_Status"" = 'Closed', ""U_StartDate"" = '{sqlStartDate}',
                                                        //    ""U_StartTime"" = '{sqlStartTime}', ""U_EndDate"" = '{sqlEndDate}', ""U_EndTime"" = '{sqlEndTime}', 
                                                        //    ""U_TotalBrkDwnT"" = '{sqlBrkTime}', ""U_ActualTimeCon"" = '{sqlActTime}'
                                                        //    WHERE ""U_WorkId"" = '{WorkId}'";
                                                        //rsUpdate.DoQuery(updateQuery);

                                                        //for 24 hours exceed
                                                        string updateQuery = $@"UPDATE ""@WRKORDRDTLSC"" 
                                                        SET ""U_IssueForProdEntry"" = '{goodIssue}' , ""U_Status"" = 'Closed', ""U_StartDate"" = '{sqlStartDate}',
                                                        ""U_StartTime"" = '{sqlStartTime}', ""U_EndDate"" = '{sqlEndDate}', ""U_EndTime"" = '{sqlEndTime}', 
                                                        ""U_TotalBrkDwnT"" = '{strBrkTime}', ""U_ActualTimeCon"" = '{manualTimeStr}'
                                                        WHERE ""U_WorkId"" = '{WorkId}'";
                                                        rsUpdate.DoQuery(updateQuery);
                                                        /////////

                                                        string chasisNo = oGeneralData.GetProperty("U_ChassisNo");

                                                        // Check if other processes are still pending
                                                        SAPbobsCOM.Recordset rsSelect = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                        string selectQuery = $@"SELECT * FROM ""@WRKORDRDTLSC"" 
                                                    WHERE ""U_ChasisNo"" = '{chasisNo}' and ""U_IssueForProdEntry"" IS NULL ";
                                                        rsSelect = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                        rsSelect.DoQuery(selectQuery);

                                                        if (rsSelect.RecordCount > 0)
                                                        {
                                                            SBO_Application.StatusBar.SetText("Goods Issue Created, but Engine/Chassis status not completed (other processes pending).",
                                                                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                                            //SBO_Application.MessageBox("Goods Issue Created, but Engine/Chassis status not completed (other processes pending).");
                                                        }
                                                        else
                                                        {
                                                            try
                                                            {
                                                                int totalMinutes = 0;
                                                                string sumQuery = $@"SELECT ""U_ActualTimeCon"" FROM ""@WRKORDRDTLSC"" WHERE ""U_ChasisNo"" = '{chasisNo}'";

                                                                rsSelect.DoQuery(sumQuery);

                                                                while (!rsSelect.EoF)
                                                                {
                                                                    string timeStr = rsSelect.Fields.Item("U_ActualTimeCon").Value.ToString().Replace(":", "").Trim();

                                                                    if (!string.IsNullOrEmpty(timeStr))
                                                                    {
                                                                        // Pad to ensure we have at least HHmm (e.g., "15" -> "0015", "130" -> "0130")
                                                                        timeStr = timeStr.PadLeft(4, '0');

                                                                        // Take the last 4 digits only to be safe
                                                                        if (timeStr.Length > 4)
                                                                            timeStr = timeStr.Substring(timeStr.Length - 4);

                                                                        if (int.TryParse(timeStr.Substring(0, 2), out int hours) &&
                                                                            int.TryParse(timeStr.Substring(2, 2), out int mins))
                                                                        {
                                                                            totalMinutes += (hours * 60) + mins;
                                                                        }
                                                                    }
                                                                    rsSelect.MoveNext();
                                                                }

                                                                // Convert total minutes back to HHmm format
                                                                int finalHours = totalMinutes / 60;
                                                                int finalMins = totalMinutes % 60;
                                                                string finalTotalTime = $"{finalHours:00}:{finalMins:00}";

                                                                //Update 
                                                                string updateEngineChasis = $@"UPDATE ""@ENGCHASISMMC"" 
                                                        SET ""U_Status"" = 'Completed' 
                                                        WHERE ""U_ProdOrdNo"" = '{productionOrderEntry}' and ""U_ChasisNo"" = '{chasisNo}'";
                                                                rsUpdate.DoQuery(updateEngineChasis);

                                                                string updateProductionOrder = $@"UPDATE ""@ENGCHASPO"" SET ""U_Status"" = 'Completed', ""U_TotalTime"" = '{finalTotalTime}'  WHERE ""U_ProdDocEntry"" = '{productionOrderEntry}' and ""U_ChasisNo"" = '{chasisNo}'";
                                                                rsUpdate.DoQuery(updateProductionOrder);


                                                                //SBO_Application.StatusBar.SetText("Status updated successfully in Engine Chasis Mapping Master and Production Order.",
                                                                //    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                                                //SBO_Application.MessageBox("Status updated successfully in Engine Chasis Mapping Master and Production Order.");
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                SBO_Application.StatusBar.SetText("Error updating Engine/Chasis: " + ex.Message,
                                                                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                                                //SBO_Application.MessageBox("Error updating Engine/Chasis: " + ex.Message);
                                                            }
                                                        }
                                                        //SBO_Application.StatusBar.SetText("Job Stopped Successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                                        //SBO_Application.MessageBox("Job Stopped Successfully.");
                                                        ShowBrowserMessage("✅", "JOB FINISHED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#00e676", "ACTION: Operation closed. Issue for Production Created.");

                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        SBO_Application.StatusBar.SetText("Job Stopped, but failed to update details: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                                        //SBO_Application.MessageBox("Job Stopped, but failed to update details: " + ex.Message);
                                                    }
                                                }
                                                else
                                                {
                                                    // Goods Issue Failed
                                                    //SBO_Application.StatusBar.SetText("Failed to create Issue for Production. Job Status NOT changed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                                    SBO_Application.MessageBox("Failed to create Issue for Production. Job Status NOT changed.");
                                                    return;
                                                }
                                            }
                                            catch(Exception ex) { }
                                            finally
                                            {
                                                if (pb != null)
                                                {
                                                    pb.Stop();
                                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(pb);
                                                    pb = null;
                                                    scancount = 0;
                                                    _scanBuffer = "";
                                                    BubbleEvent = false;
                                                }
                                            }
                                        }

                                        else
                                        {
                                            ShowBrowserMessage("❌", "INVALID ACTION", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#ff5252", "ACTION: Job is currently RUNNING. Only PAUSE or STOP allowed.");
                                            //SBO_Application.MessageBox($"Invalid Action! Current status is 'Start'. You can only PAUSE or STOP.");
                                            _scanBuffer = "";
                                            scancount = 0; // Reset
                                            BubbleEvent = false;
                                            return;
                                        }

                                    }
                                    #endregion

                                    #region When current status is "Pause" (User can scan Resume BarCode)
                                    else if (currentStatus == "Pause")
                                    {
                                        if (matchedColumn == "U_JobResume")
                                        {
                                            oGeneralData.SetProperty("U_Status", "Resume");

                                            // We only need to update the LAST row (the one currently open)
                                            if (oChildren.Count > 0)
                                            {
                                                // Get the last row added
                                                oChild = oChildren.Item(oChildren.Count - 1);

                                                DateTime manualTime = new DateTime(2025, 12, 08, 10, 40, 00);
                                                oChild.SetProperty("U_BrkEndDt", now);
                                                oChild.SetProperty("U_BrkEndTm", now);

                                                // Calculate duration for THIS specific break
                                                string sDt = Convert.ToString(oChild.GetProperty("U_BrkStartDt"));
                                                string sTm = Convert.ToString(oChild.GetProperty("U_BrkStartTm"));

                                                //DateTime startDT;
                                                //if (TryGetDateTime(sDt, sTm, out startDT))
                                                //{
                                                //    double minutes = (now - startDT).TotalMinutes;

                                                //    // Store line duration
                                                //    TimeSpan ts = TimeSpan.FromMinutes(minutes);
                                                //    DateTime lineDuration = new DateTime(1899, 12, 30).Add(ts);
                                                //    oChild.SetProperty("U_BrkTime", lineDuration);
                                                //}
                                                if (sDt != null && sTm != null)
                                                {
                                                    DateTime fullStart = GetFullDateTime(sDt, sTm);
                                                    DateTime fullEnd = now; // Since we just set it to manualTime

                                                    // This subtraction (fullEnd - fullStart) automatically handles 
                                                    // date changes (e.g. Dec 7 to Dec 8)
                                                    double minutes = (fullEnd - fullStart).TotalMinutes;

                                                    if (minutes < 0) minutes = 0; // Safety

                                                    //TimeSpan ts = TimeSpan.FromMinutes(minutes);
                                                    //DateTime lineDuration = new DateTime(1899, 12, 30).Add(ts);
                                                    //oChild.SetProperty("U_BrkTime", lineDuration);

                                                    TimeSpan ts = TimeSpan.FromMinutes(minutes);
                                                    string breakTimeStr = string.Format("{0:00}:{1:00}", (int)ts.TotalHours, ts.Minutes);
                                                    oChild.SetProperty("U_BrkTime", breakTimeStr);
                                                }
                                                //oForm = Utilities.Application.SBO_Application.Forms.Item("frmWJobEx");
                                                //_Matrix = (SAPbouiCOM.Matrix)oForm.Items.Item("mtBreak").Specific;
                                                //((SAPbouiCOM.EditTextColumn)_Matrix.Columns.Item("colBrkTm")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
                                            }

                                            //SBO_Application.StatusBar.SetText("Job Resumed", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                            //SBO_Application.MessageBox("Job Resumed");
                                            //ShowBrowserMessage("▶️", "JOB RESUMED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#00e676", "ACTION: Assembly restarted. Continue work.");
                                            ShowBrowserMessage("✅", "JOB RESUMED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#00e676", "ACTION: Assembly restarted. Continue work.");
                                        }
                                        else
                                        {
                                            ShowBrowserMessage("❌", "RESUME REQUIRED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#ff5252", "ACTION: Job is PAUSED. You must scan RESUME to proceed.");
                                            //SBO_Application.MessageBox($"Invalid Action! Current status is 'Pause'. You must scan RESUME to continue.");
                                            scancount = 0;
                                            _scanBuffer = "";
                                            BubbleEvent = false;
                                            return;
                                        }
                                    }
                                    #endregion

                                    #region When Current Status is "Resumed" then two scenario's (Pause and Stop))
                                    else if (currentStatus == "Resume")
                                    {
                                        #region  Pause scanned(add a new line)   
                                        if (matchedColumn == "U_JobPause")
                                        {
                                            // Change Status to Pause
                                            oGeneralData.SetProperty("U_Status", "Pause");

                                            // Safety Check: Ensure the PREVIOUS break row is actually closed.
                                            if (oChildren.Count > 0)
                                            {
                                                SAPbobsCOM.GeneralData oLastChild = oChildren.Item(oChildren.Count - 1);
                                                string lastEndDt = Convert.ToString(oLastChild.GetProperty("U_BrkEndDt"));

                                                // If the last row has no End Date, close it now before starting a new one
                                                if (string.IsNullOrEmpty(lastEndDt))
                                                {
                                                    
                                                    oLastChild.SetProperty("U_BrkEndDt", now);
                                                    oLastChild.SetProperty("U_BrkEndTm", now);
                                                    // Calculate time for this "forgotten" break here if needed
                                                }
                                            }

                                            // ADD A NEW ROW for the new Pause
                                            oChild = oChildren.Add(); 
                                            oChild.SetProperty("U_BrkStartDt", now);
                                            oChild.SetProperty("U_BrkStartTm", now);

                                            //SBO_Application.StatusBar.SetText("Job Paused again. New break line added.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                            //SBO_Application.MessageBox("Job Paused again. New break line added.");
                                            //ShowBrowserMessage("✅ Success", "Job Paused Sucessfully", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#28a745");
                                            ShowBrowserMessage("✅", "JOB PAUSED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#ff5252", "ACTION: Job is PAUSED. You must scan RESUME to proceed.");
                                        }
                                        #endregion

                                        #region Final Stop BarCode Scanned After Resume Status (Calculation of all Breaks and creation of goods issue)
                                        else if (matchedColumn == "U_LastBarcode")
                                        {
                                            SAPbouiCOM.ProgressBar pb = null;
                                            try
                                            {
                                                pb = SBO_Application.StatusBar.CreateProgressBar("Creating Production Issue........", 1, false);

                                                int goodIssue = Create_IssueFor_Production(_RSRoute.Fields.Item("U_PrdOrdEnt").Value.ToString(), routeId, WorkId, FirstScan);
                                                //int goodIssue = 786;

                                                // Only proceed if Goods Issue was successful
                                                if (goodIssue > 0)
                                                {
                                                    oGeneralData.SetProperty("U_Status", "Stop");
                                                    oGeneralData.SetProperty("U_EndDate", now);
                                                    oGeneralData.SetProperty("U_EndTime", now);
                                                    oGeneralData.SetProperty("U_IForProdEntry", goodIssue.ToString());

                                                    // -------------------------------------------------------------
                                                    // CALCULATE TOTAL BREAK TIME (SUM OF ALL LINES)
                                                    // -------------------------------------------------------------
                                                    //double totalBreakMinutes = 0;

                                                    //for (int i = 0; i < oChildren.Count; i++)
                                                    //{
                                                    //    oChild = oChildren.Item(i);
                                                    //    object oSDt = oChild.GetProperty("U_BrkStartDt");
                                                    //    object oSTm = oChild.GetProperty("U_BrkStartTm");
                                                    //    object oEDt = oChild.GetProperty("U_BrkEndDt");
                                                    //    object oETm = oChild.GetProperty("U_BrkEndTm");

                                                    //    if (oSDt != null && oSTm != null && oEDt != null && oETm != null)
                                                    //    {
                                                    //        try
                                                    //        {
                                                    //            DateTime rowStart = GetFullDateTime(oSDt, oSTm);
                                                    //            DateTime rowEnd = GetFullDateTime(oEDt, oETm);

                                                    //            double rowMinutes = (rowEnd - rowStart).TotalMinutes;

                                                    //            if (rowMinutes > 0)
                                                    //            {
                                                    //                // Update individual line total
                                                    //                TimeSpan tsRow = TimeSpan.FromMinutes(rowMinutes);
                                                    //                oChild.SetProperty("U_BrkTime", new DateTime(1899, 12, 30).Add(tsRow));

                                                    //                // Add to Grand Total
                                                    //                totalBreakMinutes += rowMinutes;
                                                    //            }
                                                    //        }
                                                    //        catch(Exception ex) 
                                                    //    { 

                                                    //    }
                                                    //    }
                                                    //}

                                                    //// Set the Grand Total in the Header (Sum of all breaks)
                                                    //oGeneralData.SetProperty("U_TotBrkTime", new DateTime(1899, 12, 30).AddMinutes(totalBreakMinutes));

                                                    double totalBreakMinutes = 0;

                                                    for (int i = 0; i < oChildren.Count; i++)
                                                    {
                                                        oChild = oChildren.Item(i);
                                                        object oSDt = oChild.GetProperty("U_BrkStartDt");
                                                        object oSTm = oChild.GetProperty("U_BrkStartTm");
                                                        object oEDt = oChild.GetProperty("U_BrkEndDt");
                                                        object oETm = oChild.GetProperty("U_BrkEndTm");

                                                        if (oSDt != null && oSTm != null && oEDt != null && oETm != null)
                                                        {
                                                            try
                                                            {
                                                                DateTime rowStart = GetFullDateTime(oSDt, oSTm);
                                                                DateTime rowEnd = GetFullDateTime(oEDt, oETm);
                                                                double rowMinutes = (rowEnd - rowStart).TotalMinutes;

                                                                if (rowMinutes > 0)
                                                                {
                                                                    // Format individual row break time as HH:mm (supports > 24h)
                                                                    int rHrs = (int)(rowMinutes / 60);
                                                                    int rMns = (int)(rowMinutes % 60);
                                                                    string rowDuration = $"{rHrs:D2}:{rMns:D2}";

                                                                    // If U_BrkTime is also Alpha, set as string. If it's still Time, keep old logic.
                                                                    oChild.SetProperty("U_BrkTime", rowDuration);

                                                                    totalBreakMinutes += rowMinutes;
                                                                }
                                                            }
                                                            catch (Exception ex) { }
                                                        }
                                                    }

                                                    // Set the Grand Total Break Time in the Header as a String
                                                    int bHrs = (int)(totalBreakMinutes / 60);
                                                    int bMns = (int)(totalBreakMinutes % 60);
                                                    string totalBrkDurationString = $"{bHrs:D2}:{bMns:D2}";
                                                    oGeneralData.SetProperty("U_TotBrkTime", totalBrkDurationString);

                                                    try
                                                    {
                                                        //double totalActMinutes = 0;
                                                        //double totalTime = 0;
                                                        //if (oChildren.Count > 0)
                                                        //{
                                                        //    object hStartDt = oGeneralData.GetProperty("U_StartDate");
                                                        //    object hStartTm = oGeneralData.GetProperty("U_StartTime");

                                                        //    SAPbobsCOM.GeneralData oLastChildForAct = oChildren.Item(oChildren.Count - 1);
                                                        //    object cEndDt = oGeneralData.GetProperty("U_EndDate");
                                                        //    object cEndTm = oGeneralData.GetProperty("U_EndTime");

                                                        //    DateTime dtHeaderStart = GetFullDateTime(hStartDt, hStartTm);
                                                        //    DateTime dtLastChildEnd = GetFullDateTime(cEndDt, cEndTm);

                                                        //    if (dtHeaderStart != DateTime.MinValue && dtLastChildEnd != DateTime.MinValue)
                                                        //        totalTime = (dtLastChildEnd - dtHeaderStart).TotalMinutes;
                                                        //        totalActMinutes = totalTime - totalBreakMinutes;
                                                        //}
                                                        //oGeneralData.SetProperty("U_TotActTime", new DateTime(1899, 12, 30).AddMinutes(totalActMinutes < 0 ? 0 : totalActMinutes));

                                                        DateTime dtHeaderStart = GetFullDateTime(oGeneralData.GetProperty("U_StartDate"), oGeneralData.GetProperty("U_StartTime"));
                                                        DateTime dtLastChildEnd = GetFullDateTime(oGeneralData.GetProperty("U_EndDate"), oGeneralData.GetProperty("U_EndTime"));

                                                        double totalElapsedMinutes = (dtLastChildEnd - dtHeaderStart).TotalMinutes;
                                                        double totalActMinutes = totalElapsedMinutes - totalBreakMinutes;

                                                        int hours = (int)(totalActMinutes / 60);
                                                        int mins = (int)(totalActMinutes % 60);
                                                        string durationString = $"{hours:D2}:{mins:D2}";


                                                        //if alpha
                                                        //oGeneralData.SetProperty("U_TotActTime", durationString);
                                                        oGeneralData.SetProperty("U_TotActTime", durationString);

                                                        // NOTE: Visually in SAP, this will show "01:16", but the underlying value 
                                                        // will contain the extra day (1899-12-31) which is important for the next step.
                                                        //oGeneralData.SetProperty("U_TotActTime", new DateTime(1899, 12, 30).AddMinutes(totalActMinutes));
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                    }


                                                    // Update External Tables
                                                    try
                                                    {

                                                        DateTime dtStartDate = (DateTime)oGeneralData.GetProperty("U_StartDate");
                                                        DateTime dtStartTime = (DateTime)oGeneralData.GetProperty("U_StartTime");
                                                        DateTime dtEndDate = (DateTime)oGeneralData.GetProperty("U_EndDate");
                                                        DateTime dtEndTime = (DateTime)oGeneralData.GetProperty("U_EndTime");
                                                        //DateTime dtBrkTime = (DateTime)oGeneralData.GetProperty("U_TotBrkTime");
                                                        //DateTime dtActTime = (DateTime)oGeneralData.GetProperty("U_TotActTime");
                                                        string strBrkTime = Convert.ToString(oGeneralData.GetProperty("U_TotBrkTime"));
                                                        string strActTime = Convert.ToString(oGeneralData.GetProperty("U_TotActTime"));


                                                        string sqlStartDate = dtStartDate.ToString("yyyyMMdd");
                                                        string sqlEndDate = dtEndDate.ToString("yyyyMMdd");
                                                        string sqlStartTime = dtStartTime.ToString("HHmm");
                                                        string sqlEndTime = dtEndTime.ToString("HHmm");
                                                        //string sqlBrkTime = dtBrkTime.ToString("HHmm");
                                                        //string sqlActTime = dtActTime.ToString("HHmm");
                                                        string sqlBrkTime = strBrkTime.Replace(":", "").PadLeft(4, '0');
                                                        string sqlActTime = strActTime.Replace(":", "").PadLeft(4, '0');

                                                        /// for exceeding 24 hours 
                                                        DateTime dtHeaderStart = GetFullDateTime(oGeneralData.GetProperty("U_StartDate"), oGeneralData.GetProperty("U_StartTime"));
                                                        DateTime dtLastChildEnd = GetFullDateTime(oGeneralData.GetProperty("U_EndDate"), oGeneralData.GetProperty("U_EndTime"));

                                                        double totalElapsedMinutes = (dtLastChildEnd - dtHeaderStart).TotalMinutes;
                                                        double totalActMinutes = totalElapsedMinutes - totalBreakMinutes;

                                                        int totalHrs = (int)(totalActMinutes / 60);
                                                        int totalMns = (int)(totalActMinutes % 60);

                                                        // Use a string format for the SQL update
                                                        // This will result in "2516" instead of "0116"
                                                        string manualTimeStr = $"{totalHrs:D2}:{totalMns:D2}";
                                                        /// end for exceeding 24 hours


                                                        SAPbobsCOM.Recordset rsUpdate = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                                        //string updateQuery = $@"UPDATE ""@WRKORDRDTLSC"" 
                                                        //    SET ""U_IssueForProdEntry"" = '{goodIssue}' , ""U_Status"" = 'Closed', ""U_StartDate"" = '{sqlStartDate}',
                                                        //    ""U_StartTime"" = '{sqlStartTime}', ""U_EndDate"" = '{sqlEndDate}', ""U_EndTime"" = '{sqlEndTime}', 
                                                        //    ""U_TotalBrkDwnT"" = '{sqlBrkTime}', ""U_ActualTimeCon"" = '{sqlActTime}'
                                                        //    WHERE ""U_WorkId"" = '{WorkId}'";
                                                        //rsUpdate.DoQuery(updateQuery);

                                                        //for 24 hours exceed
                                                        string updateQuery = $@"UPDATE ""@WRKORDRDTLSC"" 
                                                        SET ""U_IssueForProdEntry"" = '{goodIssue}' , ""U_Status"" = 'Closed', ""U_StartDate"" = '{sqlStartDate}',
                                                        ""U_StartTime"" = '{sqlStartTime}', ""U_EndDate"" = '{sqlEndDate}', ""U_EndTime"" = '{sqlEndTime}', 
                                                        ""U_TotalBrkDwnT"" = '{totalBrkDurationString}', ""U_ActualTimeCon"" = '{manualTimeStr}'
                                                        WHERE ""U_WorkId"" = '{WorkId}'";
                                                        rsUpdate.DoQuery(updateQuery);
                                                        /////////

                                                        string chasisNo = oGeneralData.GetProperty("U_ChassisNo");
                                                        string selectQuery = $@"SELECT * FROM ""@WRKORDRDTLSC"" WHERE ""U_ChasisNo"" = '{chasisNo}' and ""U_IssueForProdEntry"" IS NULL ";
                                                        SAPbobsCOM.Recordset rsSelect = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                        rsSelect.DoQuery(selectQuery);

                                                        if (rsSelect.RecordCount > 0)
                                                        {
                                                            SBO_Application.StatusBar.SetText("All processes must be completed before updating Engine/Chasis.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                                        }
                                                        else
                                                        {
                                                            try
                                                            {
                                                                int totalMinutes = 0;
                                                                string sumQuery = $@"SELECT ""U_ActualTimeCon"" FROM ""@WRKORDRDTLSC"" WHERE ""U_ChasisNo"" = '{chasisNo}'";

                                                                rsSelect.DoQuery(sumQuery);

                                                                while (!rsSelect.EoF)
                                                                {
                                                                    string timeStr = rsSelect.Fields.Item("U_ActualTimeCon").Value.ToString().Replace(":", "").Trim();

                                                                    if (!string.IsNullOrEmpty(timeStr))
                                                                    {
                                                                        // Pad to ensure we have at least HHmm (e.g., "15" -> "0015", "130" -> "0130")
                                                                        timeStr = timeStr.PadLeft(4, '0');

                                                                        // Take the last 4 digits only to be safe
                                                                        if (timeStr.Length > 4)
                                                                            timeStr = timeStr.Substring(timeStr.Length - 4);

                                                                        if (int.TryParse(timeStr.Substring(0, 2), out int hours) &&
                                                                            int.TryParse(timeStr.Substring(2, 2), out int mins))
                                                                        {
                                                                            totalMinutes += (hours * 60) + mins;
                                                                        }
                                                                    }
                                                                    rsSelect.MoveNext();
                                                                }

                                                                // Convert total minutes back to HHmm format
                                                                int finalHours = totalMinutes / 60;
                                                                int finalMins = totalMinutes % 60;
                                                                string finalTotalTime = $"{finalHours:00}:{finalMins:00}";

                                                                //Update 
                                                                string updateEngineChasis = $@"UPDATE ""@ENGCHASISMMC"" 
                                                        SET ""U_Status"" = 'Completed' 
                                                        WHERE ""U_ProdOrdNo"" = '{productionOrderEntry}' and ""U_ChasisNo"" = '{chasisNo}'";
                                                                rsUpdate.DoQuery(updateEngineChasis);

                                                                string updateProductionOrder = $@"UPDATE ""@ENGCHASPO"" SET ""U_Status"" = 'Completed', ""U_TotalTime"" = '{finalTotalTime}'  WHERE ""U_ProdDocEntry"" = '{productionOrderEntry}' and ""U_ChasisNo"" = '{chasisNo}'";
                                                                rsUpdate.DoQuery(updateProductionOrder);


                                                                //SBO_Application.StatusBar.SetText("Status updated successfully in Engine Chasis Mapping Master and Production Order.",
                                                                //    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                                                //SBO_Application.MessageBox("Status updated successfully in Engine Chasis Mapping Master and Production Order.");
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                SBO_Application.StatusBar.SetText("Error updating Engine/Chasis: " + ex.Message,
                                                                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                                                //SBO_Application.MessageBox("Error updating Engine/Chasis: " + ex.Message);
                                                            }
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        SBO_Application.StatusBar.SetText("Job Stopped, but failed to update details: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                                    }
                                                    ShowBrowserMessage("✅", "JOB FINISHED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#00e676", "ACTION: Operation closed. Issue for Production Created.");
                                                }
                                                else
                                                {
                                                    //SBO_Application.StatusBar.SetText("Failed to create Issue for Production. Job NOT stopped.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                                    SBO_Application.MessageBox("Failed to create Issue for Production. Job NOT stopped.");
                                                    return; // Exit
                                                }
                                            }
                                            catch(Exception ex)
                                            {
                                                SBO_Application.MessageBox("Critical Error: " + ex.Message);
                                            }
                                            finally
                                            {
                                                if (pb != null)
                                                {
                                                    pb.Stop();
                                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(pb);
                                                    pb = null;
                                                    scancount = 0;
                                                    _scanBuffer = "";
                                                    BubbleEvent = false;
                                                }
                                                //pb.Stop();
                                                //ShowBrowserMessage("✅", "JOB FINISHED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#00e676", "ACTION: Operation closed. Issue for Production Created.");
                                            }
                                           
                                        }
                                        #endregion

                                        else
                                        {
                                            ShowBrowserMessage("❌", "PAUSE/ STOP REQUIRED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#ff5252", "INVALID ACTION: You can only PAUSE or STOP.");
                                            //SBO_Application.MessageBox($"Invalid Action! Current status is 'Resume'. You can only PAUSE or STOP.");
                                            scancount = 0;
                                            _scanBuffer = "";
                                            BubbleEvent = false;
                                            return; // Scanned something else while in Resume
                                        }
                                    }
                                    #endregion

                                    #region When Current Status is "Stop" dont let to do anything
                                    else if (currentStatus == "Stop")
                                    {
                                        //SBO_Application.StatusBar.SetText("Job already stopped!", SAPbouiCOM.BoMessageTime.bmt_Short,SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        //SBO_Application.MessageBox("Job already stopped!");
                                        ShowBrowserMessage("🔔", "ALREADY COMPLETED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#17a2b8", "ACTION: This stage is closed. No further actions allowed.");
                                    }
                                    #endregion

                                    #region Try to reopen the Job Execution Form
                                    //// --- Final Commit ---
                                    oGeneralService.Update(oGeneralData);

                                    //string docEntry = _RSWorkExecDetails.Fields.Item("DocEntry").Value.ToString();
                                    //string uniqueFormID = "WJOBEXE_" + docEntry;

                                    //try
                                    //{
                                    //    bool formFound = false;

                                    //    // Check if this specific Job Form is ALREADY open
                                    //    try
                                    //    {
                                    //        // Try to grab the form by its Unique ID
                                    //        SAPbouiCOM.Form existingForm = SBO_Application.Forms.Item(uniqueFormID);
                                    //        existingForm.Select(); // Bring it to the front

                                    //        // Refresh to show the new status 
                                    //        existingForm.Refresh();

                                    //        formFound = true;
                                    //    }
                                    //    catch
                                    //    {
                                    //        formFound = false;
                                    //    }

                                    //    // If the form is NOT open, we create it and load the data
                                    //    if (!formFound)
                                    //    {
                                    //        SAPbouiCOM.FormCreationParams cp = (SAPbouiCOM.FormCreationParams)SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);

                                    //        cp.FormType = "WJOBEXE";  
                                    //        cp.ObjectType = "WJOBEXE"; 
                                    //        cp.UniqueID = uniqueFormID;

                                    //        SAPbouiCOM.Form oForm = SBO_Application.Forms.AddEx(cp);

                                    //        // Safely switch to Find Mode and Load the Record
                                    //        oForm.Freeze(true);
                                    //        try
                                    //        {
                                    //            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                                    //            // IMPORTANT: "0_U_E" is usually the default UI ID for the DocEntry/DocNum field in a UDO.
                                    //            // We use the DocEntry from the DB (docEntry variable) to find the record exact match.
                                    //            SAPbouiCOM.EditText txtFind = (SAPbouiCOM.EditText)oForm.Items.Item("0_U_E").Specific;
                                    //            txtFind.Value = docEntry;

                                    //            // Click the Find Button (Item "1")
                                    //            oForm.Items.Item("1").Click();
                                    //        }
                                    //        finally
                                    //        {
                                    //            oForm.Freeze(false);
                                    //        }
                                    //    }
                                    //}
                                    //catch (Exception ex)
                                    //{
                                    //    SBO_Application.StatusBar.SetText(
                                    //        "Job updated, but error opening form: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning
                                    //    );
                                    //}
                                    #endregion
                                }
                                else
                                {
                                    if (matchedColumn == "U_FirstBarcode")
                                    {
                                        // 1. Initialize Company Service and General Service
                                        // Note: oCompany must be your valid SAP Company object
                                        SAPbobsCOM.CompanyService oCompanyService = Utilities.Application.Company.GetCompanyService();
                                        SAPbobsCOM.GeneralService oGeneralService;
                                        SAPbobsCOM.GeneralData oGeneralData;
                                        SAPbobsCOM.GeneralDataParams oGeneralParams;

                                        try
                                        {
                                            string query_WorkOrderD = $@"SELECT T0.""U_ProdOrdNo"",T0.""U_ProductCode"",T0.""U_AdvancedSBNo"",T0.""U_ProductName"",T0.""U_AdvancedSBNo"",T0.""U_DocNo"",
                                                             T1.""U_EngineNo"",T1.""U_ChasisNo"",T1.""U_BatchNo"",T1.""U_WorkId"",T1.""U_RouteDsp""
                                                            from ""@WRKORDRDTLSH""  T0 
                                                            inner join ""@WRKORDRDTLSC""  T1 on T0.""DocEntry"" = T1.""DocEntry""  
                                                            WHERE T1.""U_WorkId"" = '{WorkId}'";

                                            SAPbobsCOM.Recordset _RSWorkOrderDetails = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                            _RSWorkOrderDetails.DoQuery(query_WorkOrderD);
                                            oGeneralService = oCompanyService.GetGeneralService("WJOBEXE");
                                            oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                                            // --- Populate UDO fields directly from Recordset values ---
                                            oGeneralData.SetProperty("U_JobID", WorkId);
                                            oGeneralData.SetProperty("U_DocNo", Utilities.getMaxColumnValueNum("@WJOBEXEH", "U_DocNo"));
                                            oGeneralData.SetProperty("U_JobDesc", _RSWorkOrderDetails.Fields.Item("U_RouteDsp").Value.ToString());
                                            oGeneralData.SetProperty("U_ProdOrdNo", _RSWorkOrderDetails.Fields.Item("U_ProdOrdNo").Value.ToString());
                                            oGeneralData.SetProperty("U_WorkOrdNo", _RSWorkOrderDetails.Fields.Item("U_DocNo").Value.ToString());
                                            oGeneralData.SetProperty("U_FGCode", _RSWorkOrderDetails.Fields.Item("U_ProductCode").Value.ToString());
                                            oGeneralData.SetProperty("U_FGDesc", _RSWorkOrderDetails.Fields.Item("U_ProductName").Value.ToString());
                                            oGeneralData.SetProperty("U_EngineNo", _RSWorkOrderDetails.Fields.Item("U_EngineNo").Value.ToString());
                                            oGeneralData.SetProperty("U_ChassisNo", _RSWorkOrderDetails.Fields.Item("U_ChasisNo").Value.ToString());
                                            oGeneralData.SetProperty("U_BatchNo", _RSWorkOrderDetails.Fields.Item("U_BatchNo").Value.ToString());
                                            oGeneralData.SetProperty("U_Status", "Start");

                                            // Auto Start Date & Time
                                            oGeneralData.SetProperty("U_StartDate", DateTime.Now);
                                            oGeneralData.SetProperty("U_StartTime", DateTime.Now);
                                            oGeneralParams = oGeneralService.Add(oGeneralData);
                                            string newDocEntry = oGeneralParams.GetProperty("DocEntry").ToString();
                                            // Operator from RS (or current user)
                                            //oGeneralData.SetProperty("U_Operator", _RSWorkExecDetails.Fields.Item("U_Operator").Value.ToString());

                                            // Update Work Order Details Document with recently created DocEntry
                                            string prodOrdNo = _RSWorkOrderDetails.Fields.Item("U_ProdOrdNo").Value.ToString();
                                            string chasisNo = _RSWorkOrderDetails.Fields.Item("U_ChasisNo").Value.ToString();
                                            if (!string.IsNullOrEmpty(newDocEntry))
                                            {
                                                try
                                                {
                                                    SAPbobsCOM.Recordset rsUpdateRef = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                    string updateRefQuery = $@"UPDATE ""@WRKORDRDTLSC"" 
                                                       SET ""U_JobExeNo"" = {newDocEntry} , ""U_Status"" = 'WIP'
                                                       WHERE ""U_WorkId"" = '{WorkId}'";

                                                    rsUpdateRef.DoQuery(updateRefQuery);

                                                    SAPbobsCOM.Recordset rsUpdate = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                    string updateQuery = $@"UPDATE ""@ENGCHASPO"" 
                                                       SET  ""U_Status"" = 'Started'
                                                       WHERE ""U_ProdDocEntry"" = '{prodOrdNo}' and ""U_ChasisNo"" = '{chasisNo}' ";

                                                    rsUpdate.DoQuery(updateQuery);
                                                }
                                                catch (Exception exUpdate)
                                                {
                                                    SBO_Application.StatusBar.SetText("Job created, but failed to link DocEntry in Work Order Details: " + exUpdate.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                                }
                                            }

                                        }
                                        catch (Exception ex)
                                        {
                                            // Handle Errors
                                            string errorMsg = ex.Message;
                                            // System.Windows.Forms.MessageBox.Show("Error creating Job: " + errorMsg);
                                        }
                                        finally
                                        {
                                            // Clean up COM objects if necessary (though GC handles them usually)
                                            //if (oGeneralData != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralData);
                                            //if (oGeneralService != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralService);
                                        }
                                        //ShowBrowserMessage("🚀", "JOB STARTED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#00e676", "ACTION: Assembly initiated. Work in progress.");
                                        ShowBrowserMessage("✅", "JOB STARTED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#00e676", "ACTION: Assembly initiated. Work in progress.");
                                    }
                                    else
                                    {
                                        string msg = "";
                                        if (string.IsNullOrEmpty(matchedColumn))
                                            msg = "Scanned barcode is not valid for this Chassis.";
                                        else
                                            msg = $"Cannot perform '{matchedColumn}' action. The Job has not been Started yet. Please scan the Start Barcode.";

                                        //SBO_Application.StatusBar.SetText(msg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                        //SBO_Application.MessageBox(msg);
                                        ShowBrowserMessage("❌", "START REQUIRED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#ff5252", "ACTION: Scan the START barcode to initiate this operation.");

                                        _scanBuffer = "";
                                        scancount = 0; 
                                        BubbleEvent = false;
                                    }
                                }
                            }
                            else
                            {
                                //SBO_Application.MessageBox("First Scan (Chassis) not found in Database.");
                                ShowBrowserMessage("❌", "DATA NOT FOUND", FirstScan, "", "", "", DateTime.Now.ToString("HH:mm:ss"), "#ff5252", "ACTION: Chassis mapping missing in Production Planning.");
                                scancount = 0;
                                _scanBuffer = "";
                                BubbleEvent = false;
                            }
                            scancount = 0;
                            _scanBuffer = "";
                            //addBtn.Enabled = true;
                            BubbleEvent = false;
                        }
                    }
                }
                #endregion

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
                this._Object = (Production_Planning_LMC.Base)this._Collection[eventInfo.FormUID];
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
                        this._Object = (Production_Planning_LMC.Base)this._Collection[BusinessObjectInfo.FormUID];
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


        #region full Date and Time for calculating total break down time
        private DateTime GetFullDateTime(object oDate, object oTime)
        {
            // 1. Convert Objects to DateTime
            // SAP "Date" fields usually return "MM/dd/yyyy 00:00:00"
            // SAP "Time" fields usually return "1899/12/30 HH:mm:ss"

            DateTime dtDate = Convert.ToDateTime(oDate);
            DateTime dtTime = Convert.ToDateTime(oTime);

            // 2. Combine them: Take Year/Month/Day from Date, and Hour/Min/Sec from Time
            DateTime fullDateTime = new DateTime(
                dtDate.Year,
                dtDate.Month,
                dtDate.Day,
                dtTime.Hour,
                dtTime.Minute,
                dtTime.Second
            );

            return fullDateTime;
        }
        #endregion

        #region Create Issue For Production 
        private int Create_IssueFor_Production(string ProdEntry, string route, string workId, string chassisNo)
        {
            int newDocEntry = 0;
            //SAPbouiCOM.ProgressBar oProgBar = null;

            try
            {
                if (string.IsNullOrWhiteSpace(ProdEntry))
                    throw new Exception("Production Order missing.");

                //oProgBar = SBO_Application.StatusBar.CreateProgressBar("Initializing Auto Issue...", 2, false);

                int prodOrderDocEntry = Convert.ToInt32(ProdEntry);

                // Fetch BOM (WOR1) for the given Production Order + Stage
                SAPbobsCOM.Recordset rsBOM = (SAPbobsCOM.Recordset)
                    Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string bomQuery = $@"
                    SELECT  
                        T0.""ItemCode"",
                        T0.""PlannedQty"",
                        T0.""IssueType"",
                        T0.""wareHouse"",
                        T0.""BaseQty"" AS ""UnitReq"",
                        T0.""LineNum"", T2.""ManSerNum"", T1.""PlannedQty"" ""headerQty""
                    FROM WOR1 T0 
                    INNER JOIN OWOR T1 ON T0.""DocEntry"" = T1.""DocEntry""
                    INNER JOIN OITM T2 on T0.""ItemCode"" = T2.""ItemCode""
                    WHERE T0.""DocEntry"" = '{prodOrderDocEntry}'
                     AND T0.""StageId"" = '{route}' and T2.""InvntItem"" = 'Y' ";////'{prodOrderDocEntry}' AND T0.""StageId"" = '1' {route} ";

                rsBOM.DoQuery(bomQuery);

                //  Fetch Chassis Mapping from UDT
                SAPbobsCOM.Recordset rsMap = (SAPbobsCOM.Recordset)
                    Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string mapQuery = $@"
                SELECT * 
                FROM ""@ENGCHASISMMC""
                WHERE ""U_ChasisNo"" = '{chassisNo}'";

                rsMap.DoQuery(mapQuery);

                if (rsMap.EoF)
                    throw new Exception("Mapping not found for Chassis No " + chassisNo);

                // Create Issue for Production (Inventory Gen Exit)
                SAPbobsCOM.Documents oIssue = (SAPbobsCOM.Documents)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);

                oIssue.DocDate = DateTime.Now;
                oIssue.Comments = $"Auto Issue — STOP Trigger | Work: {workId}, Route: {route}";
                oIssue.UserFields.Fields.Item("U_ProdEntry").Value = prodOrderDocEntry;

                //oProgBar.Text = "Preparing Lines... Please wait.";

                //  Loop through BOM items
                while (!rsBOM.EoF)
                {
                    string itemCode = rsBOM.Fields.Item("ItemCode").Value.ToString();
                    double plannedQty = Convert.ToDouble(rsBOM.Fields.Item("PlannedQty").Value);
                    double unitReq = Convert.ToDouble(rsBOM.Fields.Item("UnitReq").Value);
                    string issueMethod = rsBOM.Fields.Item("IssueType").Value.ToString();
                    string whs = rsBOM.Fields.Item("wareHouse").Value.ToString();
                    int lineNum = Convert.ToInt32(rsBOM.Fields.Item("LineNum").Value);
                    string isSerial = rsBOM.Fields.Item("ManSerNum").Value.ToString();
                    double headerPlannedQty = Convert.ToDouble(rsBOM.Fields.Item("headerQty").Value);

                    // Link Line to Production Order
                    oIssue.Lines.BaseEntry = prodOrderDocEntry;
                    oIssue.Lines.BaseLine = lineNum;
                    oIssue.Lines.BaseType = 202; // Production Order

                    //oIssue.Lines.ItemCode = itemCode;
                    //oIssue.Lines.WarehouseCode = whs;

                    double issueQty = 0;

                    // SERIAL-MANAGED ITEMS (Manual Issue Method)
                    if (isSerial == "Y")
                    {
                        // Check Serial Type
                        SAPbobsCOM.Recordset rsSer = (SAPbobsCOM.Recordset)
                            Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                        string serialQuery = $@"
                        SELECT ""U_SerialType"" FROM ""OITM"" WHERE ""ItemCode"" = '{itemCode}'";

                        rsSer.DoQuery(serialQuery);

                        string serialType = rsSer.Fields.Item("U_SerialType").Value.ToString();
                        string mappedKeyValue = "";

                        // Map according to serial type
                        switch (serialType)
                        {
                            case "ENGINE": mappedKeyValue = rsMap.Fields.Item("U_EngineNo").Value.ToString(); break;
                            case "CHASSIS": mappedKeyValue = rsMap.Fields.Item("U_ChasisNo").Value.ToString(); break;
                            case "TRANSMISSION": mappedKeyValue = rsMap.Fields.Item("U_TransNo").Value.ToString(); break;
                            case "KEY": mappedKeyValue = rsMap.Fields.Item("U_SetKey").Value.ToString(); break;
                            default:
                                throw new Exception($"SerialType '{serialType}' not mapped for item {itemCode}");
                        }

                        // Always 1 
                        issueQty = 1;
                        oIssue.Lines.Quantity = issueQty;
                        // Attach serial
                        oIssue.Lines.SerialNumbers.InternalSerialNumber = mappedKeyValue;
                        oIssue.Lines.SerialNumbers.Quantity = issueQty;
                        oIssue.Lines.SerialNumbers.Add();
                    }
                    else
                    {
                        issueQty = plannedQty/headerPlannedQty;
                        oIssue.Lines.Quantity = issueQty;
                    }

                    oIssue.Lines.Add();
                    rsBOM.MoveNext();
                }
                //oProgBar.Value = 1; // Move bar to 50%
                //oProgBar.Text = "Posting Issue for Production... Please Wait.";

                // Add the Issue Document
                int ret = oIssue.Add();
                if (ret != 0)
                {
                    Utilities.Application.Company.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Issue creation failed: {errMsg}");
                }
                else
                {
                    //oProgBar.Value = 2;
                    string key = Utilities.Application.Company.GetNewObjectKey();
                    if (int.TryParse(key, out newDocEntry))
                    {
                        SBO_Application.StatusBar.SetText($"Auto Issue created (DocEntry: {newDocEntry})", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                }

                // SBO_Application.StatusBar.SetText("Auto Issue for Production created successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                //SBO_Application.MessageBox("Auto Issue for Production created successfully.");
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return 0;
            }
            //finally
            //{
            //    if (oProgBar != null)
            //    {
            //        oProgBar.Stop();
            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar);
            //        oProgBar = null;
            //    }
            //}
            return newDocEntry;
        }
        #endregion

        #region Open Form for Scanning
        private void OpenUDOScanForm()
        {
            // 1. The Menu ID you provided
            string menuId = "2050";
            string formTypeToCheck = "139";

            try
            {
                SAPbouiCOM.Forms oForms = ((SAPbouiCOM.Application)this._Application).Forms;
                bool found = false;
                if (!found)
                {
                    ((SAPbouiCOM.Application)this._Application).ActivateMenuItem(menuId);
                    SAPbouiCOM.Form newForm = ((SAPbouiCOM.Application)this._Application).Forms.ActiveForm;
                    if (newForm.TypeEx == formTypeToCheck)
                    {
                        try
                        {
                            newForm.Items.Item("14").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            ((SAPbouiCOM.EditText)newForm.Items.Item("14").Specific).Active = true;
                        }
                        catch (Exception ex)
                        {

                        }
                    }

                }
            }
            catch (Exception ex)
            {
                ((SAPbouiCOM.Application)this._Application).StatusBar.SetText(
                    "Error opening form via Menu: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        //private void OpenScanInBrowser(string scan)
        //{
        //    string html = $"<html><body style='font-family:Arial;font-size:22px;'>Scanned Code:<br><b>{scan}</b></body></html>";
        //    string url = "data:text/html," + Uri.EscapeDataString(html);

        //    Process.Start(new ProcessStartInfo
        //    {
        //        FileName = url,
        //        UseShellExecute = true
        //    });
        //}

        private void OpenScanInBrowser(string scan, string chasisNo = null, string station = null, string modelCode = null, string modelName = null)
                    {
                        bool showDetails =
                            !string.IsNullOrEmpty(chasisNo) ||
                            !string.IsNullOrEmpty(station) ||
                            !string.IsNullOrEmpty(modelCode) ||
                            !string.IsNullOrEmpty(modelName);

                        string labelText = showDetails
                            ? "Second scan taken for break down step:"
                            : "Scan Chasis No:";

                        string createdDate = DateTime.Now.ToString("dd-MMM-yyyy");
                        string createdTime = DateTime.Now.ToString("HH:mm:ss");

                        string detailsHtml = "";

                        if (showDetails)
                        {
                            detailsHtml = $@"
                    <div class='details'>
                        <div><span>Chasis No</span>{chasisNo}</div>
                        <div><span>Model Name</span>{modelName}</div>
                        <div><span>Model Code</span>{modelCode}</div>
                        <div><span>Station</span>{station}</div>
                        <div><span>Created Date</span>{createdDate}</div>
                        <div><span>Created Time</span>{createdTime}</div>
                    </div>";
                        }

                        string html = $@"
            <!DOCTYPE html>
            <html>
            <head>
            <meta charset='utf-8'>
            <title>Scan Result</title>

            <style>
            body {{
                margin: 0;
                height: 100vh;
                font-family: Segoe UI, Arial;
                background: radial-gradient(circle at top, #2196f3, #0d47a1);
                display: flex;
                align-items: center;
                justify-content: center;
            }}

            .container {{
                width: 900px;
                max-height: 90vh; /* ensures container never overflows viewport */
                background: white;
                border-radius: 24px;
                padding: 20px 35px;
                box-shadow: 0 25px 60px rgba(0,0,0,.35);
                text-align: center;
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: flex-start;
                overflow: auto; /* allows scrolling if needed */
            }}

            .brand {{
                display: flex;
                flex-direction: column;
                align-items: center;
                margin-bottom: 15px;
            }}

            .logo svg {{
                height: 60px; /* slightly smaller to save space */
                fill: #0d47a1;
            }}

            .car-banner {{
                width: 105%;
                height: 100px;
                border-radius: 16px;
                margin-top: 10px;
                background: linear-gradient(135deg,#0d47a1,#42a5f5);
                display: flex;
                align-items: center;
                justify-content: center;
                color: white;
                font-size: 28px;
                font-weight: 700;
                letter-spacing: 2px;
            }}

            .label {{
                margin - top: 20px;
                font-size: 36px;
                font-weight: 700;
                color: #0d47a1;
            }}

            .scan {{
                margin - top: 10px;
                font-size: 50px;
                font-weight: 900;
                color: #0d47a1;
                background: linear-gradient(90deg,#bbdefb,#e3f2fd);
                padding: 10px 20px;
                border-radius: 16px;
                display: inline-block;
                word-break: break-all;
            }}

            .details {{
                margin - top: 20px;
                background: #e3f2fd;
                border-radius: 18px;
                padding: 15px 20px;
                text-align: left;
                font-size: 20px; /* slightly smaller font */
                width: 100%;
                box-sizing: border-box;
            }}

            .details div {{
                display: flex;
                justify-content: space-between;
                padding: 5px 0;
            }}

            .details span {{
                font - weight: 700;
                color: #1565c0;
            }}

            .countdown {{
                margin - top: 15px;
                font-size: 22px;
                font-weight: 700;
                color: #0d47a1;
            }};
            </style>

            <script>
            let sec = 5;
            setInterval(() => {{
                document.getElementById('cd').innerText =
                    'Closing in ' + sec + ' sec';
                sec--;
                if (sec < 0) window.close();
            }}, 500);
            </script>

            </head>

            <body>
            <div class='container'>

                <div class='brand'>
                    <div class='car-banner'>
                        LAXMI MOTOR CORPS PRODUCTION LINE
                    </div>
                </div>

                <div class='label'>{labelText}</div>
                <div class='scan'>{scan}</div>

                {detailsHtml}

                <div id='cd' class='countdown'>Closing in 5 sec</div>

            </div>
            </body>
            </html>";

                        string filePath = Path.Combine(Path.GetTempPath(), "ScanResult.html");
                        File.WriteAllText(filePath, html);

                        Process.Start(new ProcessStartInfo
                        {
                            FileName = filePath,
                            UseShellExecute = true
                        });
                    }

        private void CreateHiddenListenerForm()
        {
            string formUID = "frmBarScan";
            string itemUID = "txtScan";

            try
            {

                try
                {
                    SAPbouiCOM.Forms oForms = this._Application.Forms;
                    oForms.Item(formUID).Close();
                }
                catch { }

                SAPbouiCOM.FormCreationParams cp = (SAPbouiCOM.FormCreationParams)this._Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);

                cp.UniqueID = formUID;
                cp.FormType = "SCAN_BAR";
                cp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_FixedNoTitle;

                // 3. Add Form
                oScanForm = this._Application.Forms.AddEx(cp);
                oScanForm.AutoManaged = false;
                // 4. Set Properties (Off-screen)
                oScanForm.Left = 100000;
                oScanForm.Top = -100;
                oScanForm.ClientHeight = 1;
                oScanForm.ClientWidth = 1;
                oScanForm.Height = 1;
                oScanForm.Width = 1;

                oScanForm.Visible = true; // Must be true to receive keys

                // 5. Add EditText
                SAPbouiCOM.Item oItem = oScanForm.Items.Add(itemUID, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem.Left = 0;
                oItem.Top = 0;
                oItem.Width = 1;
                oItem.Height = 1;
                oScanTxt = (SAPbouiCOM.EditText)oItem.Specific;
                // Success Log
                this._Application.StatusBar.SetText("Scanner Ready!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error creating scan form: " + ex.Message);
            }
        }

        private static void ForceScanFocus()
        {
            try
            {
                if (oScanForm != null)
                {
                    oScanForm.Select();
                    oScanForm.Items.Item("txtScan").Click();
                }
            }
            catch { }
        }

        private void ShowBrowserMessage(string icon, string subTitle, string chassis, string model, string modelName, string line, string time, string color = "#28a745", string nextAction = "")
        {
            try
            {
                string filePath = Path.Combine(Path.GetTempPath(), "ScannedJobStatus.html");

                // 1. Detect Error/Warning Keywords
                string stUpper = subTitle.ToUpper();
                bool isError = stUpper.Contains("INVALID ACTION") ||
                               stUpper.Contains("SEQUENCE ERROR") ||
                               stUpper.Contains("PAUSE/ STOP REQUIRED") ||
                               stUpper.Contains("ALREADY COMPLETED") ||
                               stUpper.Contains("START REQUIRED") ||
                               stUpper.Contains("RESUME REQUIRED") ||
                               stUpper.Contains("DATA NOT FOUND") ||
                               stUpper.Contains("INVALID ISSUE FOR PRODUCTION");

                StringBuilder sidebarRows = new StringBuilder();
                string filterStatus = "";

                // 2. Identify current scanned status for the "Instant Injection" row
                if (stUpper.Contains("PAUSE")) filterStatus = "Pause";
                else if (stUpper.Contains("RESUME")) filterStatus = "Resume";
                else if (stUpper.Contains("FINISHED") || stUpper.Contains("STOP")) filterStatus = "Stop";
                else filterStatus = "Start";

                // 3. Status Board Logic
                if (!isError)
                {
                    // --- FIX: MANUALLY ADD THE CURRENT SCAN AT THE TOP ---
                    // This ensures the current scan shows up immediately even if DB update is lagging.
                    if (filterStatus != "Stop")
                    {
                        sidebarRows.Append($@"
                <div class='row' style='background-color: #fff9c4;'>
                    <div class='cell' style='font-weight:bold; color: #0d47a1;'>{chassis}</div>
                    <div class='cell status {filterStatus.ToLower()}'>{filterStatus}</div>
                </div>");
                    }

                    SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    // Exclude the current chassis from the query to prevent duplicate rows (since we injected it above)
                    string query = $@"
            SELECT ""U_ChassisNo"", ""U_Status""
            FROM ""@WJOBEXEH""
            WHERE ""U_JobDesc"" = '{line}' 
              AND ""U_Status"" <> 'Stop' 
              AND ""U_ChassisNo"" <> '{chassis}'
            ORDER BY ""U_StartTime"" DESC";

                    rs.DoQuery(query);
                    while (!rs.EoF)
                    {
                        string rowStatus = rs.Fields.Item("U_Status").Value.ToString();
                        sidebarRows.Append($@"
                <div class='row'>
                    <div class='cell' style='font-weight:bold;'>{rs.Fields.Item("U_ChassisNo").Value}</div>
                    <div class='cell status {rowStatus.ToLower()}'>{rowStatus}</div>
                </div>");
                        rs.MoveNext();
                    }
                }

                // Formatting Card Details (Left Side)
                string GetRow(string lbl, string val)
                {
                    if (string.IsNullOrEmpty(val) || val.ToUpper() == "N/A" || string.IsNullOrWhiteSpace(val)) return "";
                    return $@"<div class='card-row'><span>{lbl}</span> {val}</div>";
                }

                string detailsHtml = GetRow("Chassis No", chassis) + GetRow("Model Code", model) +
                                     GetRow("Model Name", modelName) + GetRow("Station/Line", line) + GetRow("Log Time", time);

                string actionHtml = string.IsNullOrEmpty(nextAction) ? "" :
                    $@"<div class='action-box'>
                <div class='action-title'>INSTRUCTION</div>
                <div class='action-text'>{nextAction}</div>
               </div>";

                // Generate Dynamic HTML with Height Matching
                string html = $@"
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset='utf-8'>
            <style>
                body {{ 
                    margin: 0; padding: 0; font-family: 'Segoe UI', Arial; 
                    background: #f1f3f5; height: 100vh; 
                    display: flex; align-items: center; justify-content: center; 
                    overflow: hidden; 
                }}
                .wrapper {{ 
                    display: flex; gap: 60px; align-items: stretch; 
                    justify-content: center; width: auto; max-width: 98%;
                }}
                .main-container {{ 
                    background: white; border-radius: 24px; padding: 30px; 
                    box-shadow: 0 15px 40px rgba(0,0,0,0.12); text-align: center; 
                    width: {(isError ? "900px" : "750px")}; 
                    flex-shrink: 0; border: 1px solid #e0e0e0;
                    display: flex; flex-direction: column; justify-content: center;
                }}
                .status-icon {{ font-size: 60px; color: {color}; margin-bottom: 5px; }}
                .sub-title {{ font-size: 28px; font-weight: 800; color: #333; text-transform: uppercase; margin-bottom: 15px; }}
                .details {{ background: #e3f2fd; border-radius: 15px; padding: 15px 20px; text-align: left; font-size: 20px; border-left: 10px solid {color}; }}
                .card-row {{ display: flex; justify-content: space-between; padding: 6px 0; border-bottom: 1px solid rgba(0,0,0,0.05); }}
                .card-row span {{ font-weight: 700; color: #1565c0; }}
                .action-box {{ margin-top: 20px; background: #fffde7; border: 2px dashed #ffb300; border-radius: 12px; padding: 15px; text-align: left; }}
                .action-title {{ font-size: 13px; font-weight: 800; color: #f57f17; margin-bottom: 3px; }}
                .action-text {{ font-size: 18px; font-weight: 700; color: #333; }}

                .status-board {{ 
                    display: {(isError ? "none" : "flex")};
                    flex-direction: column; background: white; border-radius: 20px; 
                    width: 310px; box-shadow: 0 10px 30px rgba(0,0,0,0.1); 
                    overflow: hidden; flex-shrink: 0; border: 1px solid #dee2e6;
                }}
                header {{ background: #0d47a1; color: white; padding: 12px; font-size: 13px; text-align: center; font-weight: bold; text-transform: uppercase; }}
                .grid-header, .row {{ display: flex; width: 100%; border-bottom: 1px solid #eee; }}
                .grid-header {{ background: #f8f9fa; font-weight: bold; color: #495057; font-size: 12px; }}
                .cell {{ flex: 1; padding: 10px; text-align: center; font-size: 13px; }}
                .scroll-box {{ flex: 1; overflow-y: auto; background: #fafafa; }}
                
                .status.start {{ color: #28a745; font-weight: bold; }}
                .status.pause {{ color: #fd7e14; font-weight: bold; }}
                .status.resume {{ color: #007bff; font-weight: bold; }}
                .status.stop {{ color: #dc3545; font-weight: bold; }}

                .footer-text {{ margin-top: 15px; font-size: 13px; font-weight: 600; color: #6c757d; }}
                .countdown {{ margin-top: 5px; font-size: 14px; font-weight: 700; color: #dc3545; }}
            </style>
            <script>
                let seconds = 15;
                setInterval(() => {{
                    seconds--;
                    document.getElementById('cd').innerText = 'Auto-closing in ' + seconds + ' seconds';
                    if (seconds <= 0) window.close();
                }}, 1000);
            </script>
        </head>
        <body>
            <div class='wrapper'>
                <div class='main-container'>
                    <div class='status-icon'>{icon}</div>
                    <div class='sub-title'>{subTitle}</div>
                    <div class='details'>{detailsHtml}</div>
                    {actionHtml}
                    <div class='footer-text'>LAXMI MOTOR CORPS</div>
                    <div id='cd' class='countdown'>Auto-closing in 15 seconds</div>
                </div>

                <div class='status-board'>
                    <header>{line} - ACTIVE JOBS</header>
                    <div class='grid-header'>
                        <div class='cell'>Chassis No</div>
                        <div class='cell'>Status</div>
                    </div>
                    <div class='scroll-box'>
                        {sidebarRows}
                    </div>
                </div>
            </div>
        </body>
        </html>";

                File.WriteAllText(filePath, html);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = filePath, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.SetStatusBarMessage("Browser Error: " + ex.Message);
            }
        }

        //private void ShowBrowserMessage(string icon, string subTitle, string chassis, string model, string modelName, string line, string time, string color = "#28a745", string nextAction = "")
        //{
        //    try
        //    {
        //        string filePath = Path.Combine(Path.GetTempPath(), "ScannedJobStatus.html");

        //        // Helper to generate a row only if value is not empty
        //        string GetRow(string label, string value)
        //        {
        //            if (string.IsNullOrEmpty(value) || value.ToUpper() == "N/A") return "";
        //            return $@"<div><span>{label}</span> {value}</div>";
        //        }

        //        // Build the dynamic list of details
        //        string detailsHtml = "";
        //        detailsHtml += GetRow("Chassis No", chassis);
        //        detailsHtml += GetRow("Model Code", model);
        //        detailsHtml += GetRow("Model Name", modelName);
        //        detailsHtml += GetRow("Station/Line", line);
        //        detailsHtml += GetRow("Log Time", time);

        //        // Build the instruction box only if nextAction is provided
        //        // Styled to match your light theme (using a very light version of the status color)
        //        string actionHtml = string.IsNullOrEmpty(nextAction) ? "" :
        //            $@"<div class='action-box'>
        //        <div class='action-title'>NEXT ACTION</div>
        //        <div class='action-text'>{nextAction}</div>
        //       </div>";

        //        string html = $@"
        //<!DOCTYPE html>
        //<html>
        //<head>
        //    <meta charset='utf-8'>
        //    <title>Scanned Job Status</title>
        //    <style>
        //        body {{
        //            margin: 0;
        //            height: 100vh;
        //            font-family: 'Segoe UI', Arial;
        //            background: #f4f7f6; /* Your preferred background */
        //            display: flex;
        //            align-items: center;
        //            justify-content: center;
        //            overflow: hidden;
        //        }}

        //        .container {{
        //            width: 850px;
        //            background: white;
        //            border-radius: 24px;
        //            padding: 30px 40px;
        //            box-shadow: 0 25px 60px rgba(0,0,0,.35);
        //            text-align: center;
        //            display: flex;
        //            flex-direction: column;
        //            align-items: center;
        //        }}

        //        .status-icon {{
        //            font-size: 55px; /* Slightly larger for visibility */
        //            font-weight: 900;
        //            color: {color};
        //            margin-bottom: 5px;
        //        }}

        //        .sub-title {{
        //            font-size: 36px;
        //            font-weight: 700;
        //            color: #333;
        //            text-transform: uppercase;
        //            margin-bottom: 25px;
        //        }}

        //        .details {{
        //            background: #e3f2fd; /* Your preferred light blue */
        //            border-radius: 18px;
        //            padding: 20px 30px;
        //            text-align: left;
        //            font-size: 24px;
        //            width: 100%;
        //            box-sizing: border-box;
        //            border-left: 10px solid {color};
        //        }}

        //        .details div {{
        //            display: flex;
        //            justify-content: space-between;
        //            padding: 10px 0;
        //            border-bottom: 1px solid rgba(0,0,0,0.05);
        //        }}

        //        .details div:last-child {{ border-bottom: none; }}

        //        .details span {{
        //            font-weight: 700;
        //            color: #1565c0; /* Your preferred blue value color */
        //        }}

        //        /* Instruction/Action Box */
        //        .action-box {{
        //            margin-top: 20px;
        //            background: #fff;
        //            border: 2px dashed {color};
        //            border-radius: 15px;
        //            padding: 15px 25px;
        //            width: 100%;
        //            box-sizing: border-box;
        //            text-align: left;
        //        }}
        //        .action-title {{ font-size: 16px; font-weight: 800; color: {color}; margin-bottom: 5px; }}
        //        .action-text {{ font-size: 24px; font-weight: 700; color: #333; }}

        //        .footer-text {{
        //            margin-top: 25px;
        //            font-size: 16px;
        //            font-weight: 600;
        //            color: #0d47a1;
        //            opacity: 0.8;
        //        }}

        //        .countdown {{
        //            margin-top: 10px;
        //            font-size: 18px;
        //            font-weight: 700;
        //            color: #d32f2f;
        //        }}
        //    </style>
        //    <script>
        //        let seconds = 10;
        //        function startCountdown() {{
        //            const cd = document.getElementById('cd');
        //            const timer = setInterval(() => {{
        //                seconds--;
        //                cd.innerText = 'Auto-closing in ' + seconds + ' seconds';
        //                if (seconds <= 0) {{
        //                    clearInterval(timer);
        //                    window.close();
        //                }}
        //            }}, 1000);
        //        }}
        //    </script>
        //</head>

        //<body onload='startCountdown()'>
        //    <div class='container'>
        //        <div class='status-icon'>{icon}</div>
        //        <div class='sub-title'>{subTitle}</div>

        //        <div class='details'>
        //            {detailsHtml}
        //        </div>

        //        {actionHtml}

        //        <div class='footer-text'>LAXMI MOTOR CORPS - PRODUCTION LINE</div>
        //        <div id='cd' class='countdown'>Auto-closing in 10 seconds</div>
        //    </div>
        //</body>
        //</html>";

        //        // 3. Save and Open
        //        File.WriteAllText(filePath, html);

        //        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        //        {
        //            FileName = filePath,
        //            UseShellExecute = true
        //        });
        //    }
        //    catch (Exception ex)
        //    {
        //        Utilities.Application.SBO_Application.SetStatusBarMessage("Browser Error: " + ex.Message);
        //    }
        //}
        private void OpenTrimAssemblyStatusBoard()
        {
            SAPbobsCOM.Recordset rs = null;

            try
            {
                if (_Company == null || !_Company.Connected)
                    return;

                rs = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(
                    SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = @"
            SELECT 
                ""U_JobID"",
                ""U_JobDesc"",
                ""U_FGDesc"",
                ""U_ChassisNo"",
                ""U_Status"",
                ""U_StartDate"",
                ""U_StartTime""
            FROM ""@WJOBEXEH"" T0
            INNER JOIN ""@WJOBEXEC"" T1
                ON T0.""DocEntry"" = T1.""DocEntry""
            WHERE 
                
                ""U_Status"" <> 'Stop'
            ORDER BY ""U_StartDate"", ""U_StartTime""";

                rs.DoQuery(query);

                StringBuilder rows = new StringBuilder();

                while (!rs.EoF)
                {
                    string startDateStr = "";
                    if (rs.Fields.Item("U_StartDate").Value != null)
                    {
                        DateTime d = Convert.ToDateTime(
                            rs.Fields.Item("U_StartDate").Value);
                        startDateStr = d.ToString("dd-MMM-yyyy");
                    }

                    string startTimeStr = "";
                    if (rs.Fields.Item("U_StartTime").Value != null)
                    {
                        int t = Convert.ToInt32(
                            rs.Fields.Item("U_StartTime").Value);

                        int h = t / 100;
                        int m = t % 100;

                        startTimeStr = DateTime.Today
                            .AddHours(h)
                            .AddMinutes(m)
                            .ToString("hh:mm tt");
                    }

                    string status =
                        rs.Fields.Item("U_Status").Value.ToString();

                    rows.Append($@"
                        <div class='row'>
                            <div class='cell'>{rs.Fields.Item("U_ChassisNo").Value}</div>
                            <div class='cell'>{rs.Fields.Item("U_FGDesc").Value}</div>
                            <div class='cell'>{rs.Fields.Item("U_JobDesc").Value}</div>
                            <div class='cell'>{rs.Fields.Item("U_JobID").Value}</div>
                            <div class='cell'>{startDateStr}</div>
                            <div class='cell'>{startTimeStr}</div>
                            <div class='cell status {status.ToLower()}'>{status}</div>
                        </div>");

                                            rs.MoveNext();
                                        }

                                        string html = $@"<!DOCTYPE html>
                        <html>
                        <head>
                        <meta charset='utf-8'>
                        <title>Trim Assembly Status</title>

                        <style>
                        body {{
                            margin: 0;
                            font-family: Segoe UI, Arial;
                            background: #f5f7fa; /* smoke white */
                        }}

                        header {{
                            background: #cfe8ff; /* light blue */
                            color: #0f172a;
                            padding: 16px;
                            font-size: 30px;
                            text-align: center;
                            font-weight: bold;
                        }}

                        .grid-header, .row {{
                            display: flex;
                            width: 100%;
                            box-sizing: border-box;
                        }}

                        .grid-header {{
                            background: #e6f2ff;
                            font-weight: bold;
                            border-bottom: 2px solid #b6d7ff;
                        }}

                        .cell {{
                            flex: 1;
                            padding: 10px;
                            text-align: center;
                            border-right: 1px solid #b6d7ff;
                            white-space: nowrap;
                            overflow: hidden;
                            text-overflow: ellipsis;
                        }}

                        .cell:last-child {{
                            border-right: none;
                        }}

                        .row {{
                            background: #ffffff;
                            border-bottom: 1px solid #dbeafe;
                        }}

                        .row:nth-child(even) {{
                            background: #f0f8ff;
                        }}

                        .scroll-box {{
                            height: calc(100vh - 110px);
                            overflow: hidden;
                        }}

                        .status.start {{
                            color: #16a34a;
                            font-weight: bold;
                        }}

                        .status.pause {{
                            color: #f59e0b;
                            font-weight: bold;
                        }}

                        .status.stop {{
                            color: #dc2626;
                            font-weight: bold;
                        }}
                        </style>

                        <script>
                        function autoScroll() {{
                            const box = document.getElementById('scrollBox');
                            if (!box) return;

                            if (box.scrollHeight <= box.clientHeight)
                                return;

                            let pos = box.scrollHeight;

                            setInterval(() => {{
                                box.scrollTop = pos;
                                pos -= 1;
                                if (pos <= 0)
                                    pos = box.scrollHeight;
                            }}, 35);
                        }}

                        window.onload = autoScroll;

                        setTimeout(() => {{
                            location.reload();
                        }}, 20000);
                        </script>
                        </head>

                        <body>

                        <header>TRIM ASSEMBLY LINE – LIVE STATUS</header>

                        <div class='grid-header'>
                            <div class='cell'>Chassis</div>
                            <div class='cell'>Product</div>
                            <div class='cell'>Operation</div>
                            <div class='cell'>Job ID</div>
                            <div class='cell'>Start Date</div>
                            <div class='cell'>Start Time</div>
                            <div class='cell'>Status</div>
                        </div>

                        <div id='scrollBox' class='scroll-box'>
                            {rows}
                        </div>

                        </body>
                        </html>";

                string filePath = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    "TrimAssemblyStatusBoard.html");

                System.IO.File.WriteAllText(filePath, html);
                System.Diagnostics.Process.Start(filePath);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
            }
            finally
            {
                if (rs != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        private void StartFocusTimer(SAPbouiCOM.Form form)
        {
            _oForm = form; // Assign form reference
            focusTimer = new System.Windows.Forms.Timer();
            focusTimer.Interval = 200; // 0.2 sec
            focusTimer.Tick += FocusTimer_Tick;
            focusTimer.Start();
        }

        private void FocusTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (_oForm != null && _oForm.Visible)
                {
                    _oForm.Items.Item("txtScan").Specific.Focus();
                }
                else
                {
                    // Stop timer if form is closed
                    focusTimer.Stop();
                }
            }
            catch
            {
                focusTimer.Stop();
            }
        }


    }
}
