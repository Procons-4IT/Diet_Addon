Imports SAPbobsCOM

Public Class clsCustomerProgram

    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private InvForConsumedItems, count As Integer
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines_0 As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines_1 As SAPbouiCOM.DBDataSource
    Public intSelectedMatrixrow As Integer = 0
    Private RowtoDelete As Integer
    Private MatrixId As String
    Private oRecordSet As SAPbobsCOM.Recordset
    Private dtValidFrom, dtValidTo As Date
    Private strQuery As String
    Private oDTDocument As SAPbouiCOM.DataTable
    Private oDTProgram As SAPbouiCOM.DataTable
    Dim oGrid As SAPbouiCOM.Grid
    Dim oCombo As SAPbouiCOM.ComboBox

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Public Sub LoadForm()
        Try
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_OCPM) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_OCPM, frm_Z_OCPM)
            oForm = oApplication.SBO_Application.Forms.Item(strUID)
            oForm.Freeze(True)
            initialize(oForm)
            loadCombo(oForm)
            oForm.DataSources.DataTables.Add("Documents")
            oForm.DataSources.DataTables.Add("ProgramDL")
            addChooseFromListConditions(oForm)
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            'oForm.EnableMenu(mnu_ADD, True)
            oForm.EnableMenu(mnu_FIND, True)
            oForm.Items.Item("37").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            loadDocuments(oForm)
            oForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

    Public Sub LoadForm(ByVal strCardCode As String, ByVal strCardName As String, ByVal strDiscount As String, ByVal strDisRemarks As String, ByVal strCurrency As String)
        Try
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_OCPM) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_OCPM, frm_Z_OCPM)
            oForm = oApplication.SBO_Application.Forms.Item(strUID)
            oForm.Freeze(True)
            initialize(oForm)
            loadCombo(oForm)
            oForm.DataSources.DataTables.Add("Documents")
            oForm.DataSources.DataTables.Add("ProgramDL")
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPM")
            addChooseFromListConditions(oForm)
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            oForm.EnableMenu(mnu_ADD, True)
            oForm.EnableMenu(mnu_FIND, True)
            oDBDataSource.SetValue("U_CardCode", 0, strCardCode)
            oDBDataSource.SetValue("U_CardName", 0, strCardName)
            oDBDataSource.SetValue("U_Discount", 0, strDiscount)
            oDBDataSource.SetValue("U_Remarks", 0, strDisRemarks)
            oDBDataSource.SetValue("U_VenCur", 0, strCurrency)
            If strCurrency = "##" Then
                oDBDataSource.SetValue("U_CurSour", 0, "C")
            End If
            DefaultCurrency(oForm, strCurrency)
            loadDocuments(oForm)
            oForm.Items.Item("37").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

    Public Sub LoadForm(ByVal strDocEntry As String)
        Try
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_OCPM) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_OCPM, frm_Z_OCPM)
            oForm = oApplication.SBO_Application.Forms.Item(strUID)
            oForm.Freeze(True)
            addChooseFromListConditions(oForm)
            initialize(oForm)
            loadCombo(oForm)
            oForm.DataSources.DataTables.Add("Documents")
            oForm.DataSources.DataTables.Add("ProgramDL")
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            oForm.EnableMenu(mnu_ADD, True)
            oForm.EnableMenu(mnu_FIND, True)
            oForm.Freeze(False)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            oForm.Items.Item("10").Specific.value = strDocEntry
            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            'oForm.PaneLevel = 1
            oForm.Items.Item("37").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            loadDocuments(oForm)
            oForm.EnableMenu(mnu_FIND, False)
            'oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Z_OCPM Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or _
                                                           oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    oForm.Freeze(True)
                                    If Validation(oForm) = False Then
                                        oForm.Freeze(False)
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
                                            oForm.Freeze(False)
                                            BubbleEvent = False
                                            Exit Sub
                                        Else

                                            'If oApplication.SBO_Application.MessageBox("Please check service items information...?", , "Yes", "No") = 2 Then
                                            '    oForm.Freeze(False)
                                            '    BubbleEvent = False
                                            '    Exit Sub
                                            'End If

                                            'Newly Added for Validate for the Service Qty & Applied Date.
                                            'Dim blnService As Double = validation_Service(oForm)
                                            'If Not blnService Then
                                            '    If oApplication.SBO_Application.MessageBox("No of Service Qty/Applied Date not matched with No Days & Program To Date Continue...?", , "Yes", "No") = 2 Then
                                            '        oForm.Freeze(False)
                                            '        BubbleEvent = False
                                            '        Exit Sub
                                            '    End If
                                            'End If

                                        End If
                                        'Else
                                        'addInvoiceServiceItemsAutomatically(oForm)
                                    End If
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "_31" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    If oApplication.SBO_Application.MessageBox("Sure you want to Create Invoice for Customer Program...Continue...?", , "Yes", "No") = 1 Then
                                        If oApplication.Utilities.AddInvoiceDocument(oForm) Then
                                            oApplication.SBO_Application.MessageBox("Invoice Document Created Successfully...")
                                            oApplication.SBO_Application.Menus.Item(mnu_ADD).Activate()
                                        End If
                                    End If
                                ElseIf pVal.ItemUID = "2_" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    Dim strReport As String = System.Configuration.ConfigurationManager.AppSettings(oForm.TypeEx)
                                    oApplication.Utilities.PrintUDO(strReport, oDBDataSource.GetValue("DocNum", 0).ToString())
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                If (pVal.ItemUID = "3" Or pVal.ItemUID = "40") And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Row > 0 Then
                                    MatrixId = pVal.ItemUID
                                    intSelectedMatrixrow = pVal.Row
                                    Dim strCardCode As String = CType(oForm.Items.Item("6").Specific, SAPbouiCOM.EditText).Value
                                    Dim strFromDate As String = CType(oForm.Items.Item("12").Specific, SAPbouiCOM.EditText).Value
                                    Dim strNoofDays As String = CType(oForm.Items.Item("14").Specific, SAPbouiCOM.EditText).Value

                                    If strCardCode.Trim().Length = 0 Or strFromDate.Trim().Length = 0 Or strFromDate.Trim().Length = 0 Then
                                        oApplication.Utilities.Message("Fill All Header Information...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                ElseIf pVal.ItemUID = "3" And pVal.ColUID = "V_7" Then
                                    BubbleEvent = False
                                    Exit Sub
                                ElseIf pVal.ItemUID = "3" Or pVal.ItemUID = "40" Then
                                    intSelectedMatrixrow = pVal.Row
                                    If pVal.ItemUID = "3" And pVal.Row > 0 Then
                                        oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM6")
                                        If CInt(IIf(oDBDataSourceLines_0.GetValue("U_OrdDays", pVal.Row - 1).Trim() = "", 0, oDBDataSourceLines_0.GetValue("U_OrdDays", pVal.Row - 1).Trim())) > 0 Then
                                            oApplication.Utilities.Message("Program Row Already Converted to Order...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        End If
                                    ElseIf pVal.ItemUID = "40" And pVal.Row > 0 Then
                                        oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@Z_CPM7")
                                        If oDBDataSourceLines_1.GetValue("U_InvCreated", pVal.Row - 1).Trim() = "Y" Then
                                            oApplication.Utilities.Message("Service Item Already Converted to Invoice...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    ElseIf pVal.ItemUID = "3" And pVal.ColUID = "V_0" And pVal.Row > 0 Then
                                        'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then

                                        'End If
                                    End If
                                ElseIf pVal.ItemUID = "37" Then
                                    'oForm.Items.Item("34").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    'oForm.PaneLevel = 1
                                    oForm.Settings.MatrixUID = "3"
                                ElseIf pVal.ItemUID = "38" Then
                                    'oForm.Items.Item("34").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    'oForm.PaneLevel = 2
                                    oForm.Settings.MatrixUID = "40"
                                ElseIf pVal.ItemUID = "39" Then
                                    'oForm.Items.Item("34").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    'oForm.PaneLevel = 3
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    'removeInvoiceSericeItemsAutomatically(oForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or _
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPM")
                                    oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM6")
                                    If pVal.ItemUID = "14" Or pVal.ItemUID = "15" Then
                                        Dim intNoofDays As Integer = CInt(IIf(oDBDataSource.GetValue("U_NoOfDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_NoOfDays", 0)))
                                        Dim intFreeDays As Integer = CInt(IIf(oDBDataSource.GetValue("U_FreeDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_FreeDays", 0)))
                                        Dim intOrderDays As Integer = CInt(IIf(oDBDataSource.GetValue("U_OrdDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_OrdDays", 0)))
                                        If (intNoofDays + intFreeDays) < intOrderDays Then
                                            oApplication.Utilities.Message("Already Order Document Processed Cannot...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    ElseIf pVal.ItemUID = "3" And pVal.ColUID = "V_1" And pVal.Row > 0 Then
                                        oMatrix = oForm.Items.Item("3").Specific
                                        oMatrix.FlushToDataSource()
                                        oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPM")
                                        oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM6")
                                        Dim intNoofDays As Integer = CInt(IIf(oDBDataSourceLines_0.GetValue("U_NoOfDays", pVal.Row - 1).ToString() = "", 0, oDBDataSourceLines_0.GetValue("U_NoOfDays", pVal.Row - 1)))
                                        Dim intOrderDays As Integer = CInt(IIf(oDBDataSourceLines_0.GetValue("U_OrdDays", pVal.Row - 1).ToString() = "", 0, oDBDataSourceLines_0.GetValue("U_OrdDays", pVal.Row - 1)))
                                        If intNoofDays < intOrderDays Then
                                            oApplication.Utilities.Message("Already Order Document Processed Cannot...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                            'Else
                                            '    Dim intBOrder As Integer = 0
                                            '    Dim intTOrderDays As Integer = 0
                                            '    For index As Integer = 0 To pVal.Row - 1
                                            '        intBOrder += CInt(IIf(oDBDataSourceLines_0.GetValue("U_OrdDays", index).ToString() = "", _
                                            '                              0, oDBDataSourceLines_0.GetValue("U_OrdDays", index)))
                                            '    Next
                                            '    intTOrderDays = CInt(IIf(oDBDataSource.GetValue("U_OrdDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_OrdDays", 0)))
                                            '    If (intTOrderDays - intBOrder) >= (intBOrder) Then
                                            '        oApplication.Utilities.Message("Cannot Increase No of Days Please Add New Row to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            '        BubbleEvent = False
                                            '        Exit Sub
                                            '    End If
                                        End If
                                    ElseIf pVal.ItemUID = "3" And pVal.ColUID = "V_0" And pVal.Row > 0 Then
                                        oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                        Dim strFDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", pVal.Row)
                                        If strFDate <> "" Then
                                            If Not oApplication.Utilities.validateDate(oForm, strFDate) Then
                                                oApplication.Utilities.Message("Program From Date Should be Greater than Or Equal Current Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf pVal.ItemUID = "40" And pVal.ColUID = "V_0_0" And pVal.Row > 0 Then
                                        oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                        Dim strFDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0_0", pVal.Row)
                                        If strFDate <> "" Then
                                            If Not oApplication.Utilities.validateDate(oForm, strFDate) Then
                                                oApplication.Utilities.Message("Applied Date Should be Greater than Or Equal Current Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPM")
                                    If pVal.ItemUID = "12" Then
                                        Dim strFDate As String = oDBDataSource.GetValue("U_PFromDate", 0).ToString()
                                        If strFDate <> "" Then
                                            If Not oApplication.Utilities.validateDate(oForm, strFDate) Then
                                                oApplication.Utilities.Message("Program From Date Should be Greater than Or Equal Current Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf pVal.ItemUID = "3" And pVal.ColUID = "V_0" And pVal.Row > 0 Then
                                        oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                        Dim strFDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", pVal.Row)
                                        If strFDate <> "" Then
                                            If Not oApplication.Utilities.validateDate(oForm, strFDate) Then
                                                oApplication.Utilities.Message("Program From Date Should be Greater than Or Equal Current Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf pVal.ItemUID = "40" And pVal.ColUID = "V_0_0" Then
                                        oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                        Dim strFDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0_0", pVal.Row)
                                        If strFDate <> "" Then
                                            If Not oApplication.Utilities.validateDate(oForm, strFDate) Then
                                                oApplication.Utilities.Message("Applied Date Should be Greater than Or Equal Current Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "14"
                                        'oApplication.SBO_Application.ActivateMenuItem(mnu_ADD_ROW)
                                    Case "15"
                                        'oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                                    Case "1"
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Action_Success Then
                                            initialize(oForm)
                                        End If
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                    If pVal.ItemUID = "3" And (pVal.ColUID = "V_3" Or pVal.ColUID = "V_5_") And pVal.Row > 0 Then
                                        oMatrix = oForm.Items.Item("3").Specific
                                        oForm.Freeze(True)
                                        oMatrix.FlushToDataSource()
                                        oMatrix.LoadFromDataSource()
                                        oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM6")
                                        oDBDataSourceLines_0.SetValue("U_Discount", pVal.Row - 1, "0")
                                        oMatrix.LoadFromDataSource()
                                        oMatrix.FlushToDataSource()
                                        calculatePriceAfterDis(oForm, pVal.Row)
                                        addDayRowsDynamically(oForm)
                                        calculate_Document_Values(oForm)
                                        oForm.Freeze(False)
                                    ElseIf pVal.ItemUID = "45" Then
                                        oForm.Freeze(True)
                                        strQuery = "SELECT MainCurncy FROM OADM "
                                        Dim strLCurrency As String = oApplication.Utilities.getRecordSetValueString(strQuery, "MainCurncy")
                                        If strLCurrency <> oDBDataSource.GetValue("U_DocCur", 0).Trim() Then
                                            strQuery = "Select Rate FROM ORTT "
                                            strQuery += " Where Convert(VarChar,RateDate,112) = '" & DateTime.Now.ToString("yyyyMMdd") & "'"
                                            strQuery += " AND Currency ='" + oDBDataSource.GetValue("U_DocCur", 0).Trim() + "'"
                                            Dim dblRate As Double = oApplication.Utilities.getRecordSetValue(strQuery, "Rate")
                                            If dblRate = 0 Then
                                                oDBDataSource.SetValue("U_DocCur", 0, strLCurrency)
                                                oApplication.SBO_Application.ActivateMenuItem("3333")
                                            Else
                                                'Calculate Logic
                                                oDBDataSource.SetValue("U_DocRate", 0, dblRate.ToString())
                                                oMatrix = oForm.Items.Item("3").Specific
                                                For index As Integer = 1 To oMatrix.VisualRowCount
                                                    calculatePriceAfterDis(oForm, index)
                                                Next
                                                oMatrix = oForm.Items.Item("40").Specific
                                                For index As Integer = 1 To oMatrix.VisualRowCount
                                                    calculatePriceAfterDis_S(oForm, index)
                                                Next
                                                calculate_Document_Values(oForm)
                                            End If
                                        Else
                                            'Calculate Logic
                                            oMatrix = oForm.Items.Item("3").Specific
                                            For index As Integer = 1 To oMatrix.VisualRowCount
                                                calculatePriceAfterDis(oForm, index)
                                            Next
                                            oMatrix = oForm.Items.Item("40").Specific
                                            For index As Integer = 1 To oMatrix.VisualRowCount
                                                calculatePriceAfterDis_S(oForm, index)
                                            Next
                                            calculate_Document_Values(oForm)
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                    If pVal.ItemUID = "12" Or pVal.ItemUID = "14" Or pVal.ItemUID = "15" Then
                                        oForm.Freeze(True)
                                        calculateRNoofDays(oForm)
                                        addDayRowsDynamically(oForm)
                                        calculate_Document_Values(oForm)
                                        oForm.Freeze(False)
                                    ElseIf pVal.ItemUID = "3" And (pVal.ColUID = "V_0" Or pVal.ColUID = "V_1" Or pVal.ColUID = "V_2" Or pVal.ColUID = "V_4" Or pVal.ColUID = "V_5") And pVal.Row > 0 Then
                                        oForm.Freeze(True)
                                        If pVal.ItemUID = "3" And (pVal.ColUID = "V_0" Or pVal.ColUID = "V_1") Then
                                            calculateRNoofDays(oForm, pVal.Row)
                                        End If
                                        calculatePriceAfterDis(oForm, pVal.Row)
                                        calculate_Document_Values(oForm)
                                        addDayRowsDynamically(oForm)
                                        oForm.Freeze(False)
                                    ElseIf pVal.ItemUID = "40" And (pVal.ColUID = "V_2" Or pVal.ColUID = "V_3" Or pVal.ColUID = "V_4" Or pVal.ColUID = "V_5_") And pVal.Row > 0 Then
                                        oForm.Freeze(True)
                                        calculatePriceAfterDis_S(oForm, pVal.Row)
                                        calculate_Document_Values(oForm)
                                        oForm.Freeze(False)
                                    ElseIf pVal.ItemUID = "17" Then
                                        oForm.Freeze(True)
                                        calculate_Document_Values(oForm)
                                        oForm.Freeze(False)
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                'If pVal.ItemUID = "3" And pVal.ColUID = "V_-1" And pVal.Row > 0 Then
                                '    oMatrix = oForm.Items.Item("3").Specific
                                '    Dim strProgDt As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", pVal.Row)
                                '    Dim strRef As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_11", pVal.Row)
                                '    Dim strCardCode As String = oApplication.Utilities.getEditTextvalue(oForm, "6")
                                '    If strProgDt.Length > 0 Then
                                '        Dim objInvoiceServiceItem As clsInvoiceServiceItem
                                '        objInvoiceServiceItem = New clsInvoiceServiceItem
                                '        If strRef.Length = 0 Then
                                '            strRef = oApplication.Utilities.AddServiceItemDocument(oForm)
                                '            oApplication.Utilities.SetMatrixValues(oMatrix, "V_11", pVal.Row, strRef)
                                '            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                '        End If
                                '        objInvoiceServiceItem.LoadForm(strRef, strCardCode)
                                '    End If
                                'End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPM")
                                    oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM6")
                                    oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@Z_CPM7")
                                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                    Dim oDataTable As SAPbouiCOM.DataTable
                                    Try
                                        oCFLEvento = pVal
                                        oDataTable = oCFLEvento.SelectedObjects

                                        If IsNothing(oDataTable) Then
                                            Exit Sub
                                        End If

                                        If pVal.ItemUID = "7" Then
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 0 Then
                                                oDBDataSource.SetValue("U_CardCode", 0, oDataTable.GetValue("CardCode", 0))
                                                oDBDataSource.SetValue("U_CardName", 0, oDataTable.GetValue("CardName", 0))
                                                oDBDataSource.SetValue("U_Discount", 0, oDataTable.GetValue("Discount", 0))
                                                oDBDataSource.SetValue("U_VenCur", 0, oDataTable.GetValue("Currency", 0))
                                                If oDataTable.GetValue("Currency", 0) = "##" Then
                                                    oDBDataSource.SetValue("U_CurSour", 0, "C")
                                                End If
                                                DefaultCurrency(oForm, oDataTable.GetValue("Currency", 0))
                                                Dim strDiscount As String = oApplication.Utilities.getRecordSetValueString("Select U_DisRemarks From [@Z_OCPR] Where U_CardCode = '" & oDataTable.GetValue("CardCode", 0).Trim() & "'", "U_DisRemarks")
                                                oDBDataSource.SetValue("U_Remarks", 0, strDiscount)
                                            End If
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf pVal.ItemUID = "11" Then
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 0 Then
                                                oDBDataSource.SetValue("U_PrgCode", 0, oDataTable.GetValue("ItemCode", 0))
                                                oDBDataSource.SetValue("U_PrgName", 0, oDataTable.GetValue("ItemName", 0))
                                            End If
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf pVal.ItemUID = "11_" Then
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 0 Then
                                                oDBDataSource.SetValue("U_PrgCode", 0, oDataTable.GetValue("ItemCode", 0))
                                                oDBDataSource.SetValue("U_PrgName", 0, oDataTable.GetValue("ItemName", 0))
                                            End If
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf (pVal.ItemUID = "40" And (pVal.ColUID = "V_0" Or pVal.ColUID = "V_1")) Then
                                            oMatrix = oForm.Items.Item("40").Specific
                                            oMatrix.LoadFromDataSource()
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 1 Then
                                                intAddRows -= 1
                                                oMatrix.AddRow(intAddRows, pVal.Row - 1)
                                            End If
                                            oMatrix.FlushToDataSource()

                                            For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                oDBDataSourceLines_1.SetValue("LineId", pVal.Row + index - 1, (pVal.Row + index).ToString())
                                                oDBDataSourceLines_1.SetValue("U_ItemCode", pVal.Row + index - 1, oDataTable.GetValue("ItemCode", index))
                                                oDBDataSourceLines_1.SetValue("U_ItemName", pVal.Row + index - 1, oDataTable.GetValue("ItemName", index))
                                                oDBDataSourceLines_1.SetValue("U_TaxCode", pVal.Row + index - 1, oDataTable.GetValue("VatGourpSa", index))
                                                Dim strNoofDays As String = oDBDataSource.GetValue("U_NoofDays", 0)
                                                oDBDataSourceLines_1.SetValue("U_Quantity", pVal.Row + index - 1, strNoofDays)
                                                oDBDataSourceLines_1.SetValue("U_Date", pVal.Row + index - 1, oDBDataSource.GetValue("U_PToDate", 0))

                                                Dim dblItemPrice, dblBasePrice As Double
                                                Dim strICurrency As String = String.Empty
                                                Dim strLCurrency As String = String.Empty
                                                strLCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
                                                oApplication.Utilities.GetCustItemPrice(oApplication.Utilities.getEditTextvalue(oForm, "6"), _
                                                                                        oDataTable.GetValue("ItemCode", index), _
                                                                                        System.DateTime.Now.Date, dblBasePrice, strICurrency)
                                                oDBDataSourceLines_1.SetValue("U_Currency", pVal.Row + index - 1, strICurrency)
                                                oDBDataSourceLines_1.SetValue("U_IPrice", pVal.Row + index - 1, dblBasePrice)

                                                If strICurrency = strLCurrency Then
                                                    If strICurrency = oDBDataSource.GetValue("U_DocCur", 0).Trim Then
                                                        oDBDataSourceLines_1.SetValue("U_Price", pVal.Row + index - 1, dblBasePrice)
                                                        oDBDataSourceLines_1.SetValue("U_LineTotal", pVal.Row + index - 1, dblBasePrice)
                                                    Else
                                                        getPrice(oDBDataSource.GetValue("U_DocCur", 0).Trim, strICurrency, dblBasePrice, dblItemPrice)
                                                        oDBDataSourceLines_1.SetValue("U_Price", pVal.Row + index - 1, dblItemPrice)
                                                        oDBDataSourceLines_1.SetValue("U_LineTotal", pVal.Row + index - 1, dblItemPrice)
                                                    End If
                                                Else
                                                    getPrice(oDBDataSource.GetValue("U_DocCur", 0).Trim, strICurrency, dblBasePrice, dblItemPrice)
                                                    oDBDataSourceLines_1.SetValue("U_Price", pVal.Row + index - 1, dblItemPrice)
                                                    oDBDataSourceLines_1.SetValue("U_LineTotal", pVal.Row + index - 1, dblItemPrice)
                                                End If
                                            Next

                                            oMatrix.LoadFromDataSource()
                                            oMatrix.FlushToDataSource()
                                            calculate_Document_Values(oForm)
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf (pVal.ItemUID = "40" And (pVal.ColUID = "V_4")) Then
                                            oMatrix = oForm.Items.Item("40").Specific
                                            oMatrix.LoadFromDataSource()
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 1 Then
                                                intAddRows -= 1
                                                oMatrix.AddRow(intAddRows, pVal.Row - 1)
                                            End If
                                            oMatrix.FlushToDataSource()
                                            oMatrix.LoadFromDataSource()
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                    Catch

                                    End Try
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                If oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Or oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    Try
                                        reDrawForm(oForm)
                                    Catch ex As Exception

                                    End Try
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.BeforeAction
                Case True
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPM")
                    oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM6")
                    oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@Z_CPM7")
                    oMatrix = oForm.Items.Item("3").Specific
                    'For index As Integer = 1 To oMatrix.VisualRowCount
                    '    If oApplication.Utilities.getMatrixValues(oMatrix, "V_7", index).ToString() = "N" Then
                    '        oMatrix.CommonSetting.SetRowEditable(index, True)
                    '    Else
                    '        oMatrix.CommonSetting.SetRowEditable(index, False)
                    '    End If
                    'Next
                    oMatrix.Columns.Item("V_-1").Editable = False
                    oMatrix.Columns.Item("V_2").Editable = False
                    oMatrix.Columns.Item("V_6").Editable = False
                    oMatrix.Columns.Item("V_8").Editable = False
                    oMatrix.Columns.Item("V_9").Editable = False
                    oMatrix.Columns.Item("V_10").Editable = False


                    Select Case pVal.MenuUID
                        Case mnu_CLOSECP
                            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPM")
                            If Not oDBDataSource.GetValue("DocEntry", 0).ToString = "" Then
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    Dim intRemDays As Integer = CInt(IIf(oDBDataSource.GetValue("U_RemDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_RemDays", 0)))
                                    Dim strPaidSt As String = oDBDataSource.GetValue("U_PaidSta", 0).ToString()
                                    If intRemDays <> 0 Then
                                        oApplication.Utilities.Message("No Remaining Days Should be 0 for the Document to Close...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    ElseIf intRemDays = 0 And strPaidSt = "O" Then
                                        oApplication.Utilities.Message("Paid Status of the Program Should be Paid for Program Close...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        Dim _retVal As Integer = oApplication.SBO_Application.MessageBox("Sure you wanted to Close Customer Program...?", 2, "Yes", "No", "")
                                        If _retVal = 2 Then
                                            Exit Sub
                                        End If
                                        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        Dim strDocEntry As String = oDBDataSource.GetValue("DocEntry", 0).Trim().ToString
                                        strQuery = " Update [@Z_OCPM] SET U_DocStatus = 'C' "
                                        strQuery += " Where DocEntry = '" + strDocEntry + "'"
                                        oRecordSet.DoQuery(strQuery)

                                        strQuery = "PROCON_UPDATEONOFFSTATUS_u"
                                        oRecordSet.DoQuery(strQuery)

                                        oApplication.SBO_Application.MessageBox("Program Registration Closed Successfully...")
                                        oApplication.SBO_Application.Menus.Item(mnu_ADD).Activate()
                                    End If
                                End If
                            End If
                        Case mnu_CANCELCP
                            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPM")
                            If Not oDBDataSource.GetValue("DocEntry", 0).ToString = "" Then
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oForm.Items.Item("34").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    Dim strCancelRemarks As String = oDBDataSource.GetValue("U_CRemarks", 0).ToString()

                                    If strCancelRemarks.Trim().Length = 0 Then
                                        oApplication.Utilities.Message("Please specify cancel remarks to cancel program...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        oForm.Items.Item("46").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If


                                    Dim intOrderDays As Integer = CInt(IIf(oDBDataSource.GetValue("U_OrdDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_OrdDays", 0)))
                                    Dim intDelDays As Integer = CInt(IIf(oDBDataSource.GetValue("U_DelDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_DelDays", 0)))
                                    Dim intInvDays As Integer = CInt(IIf(oDBDataSource.GetValue("U_InvDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_InvDays", 0)))
                                    Dim intRemDays As Integer = CInt(IIf(oDBDataSource.GetValue("U_RemDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_RemDays", 0)))


                                    Dim intNoofDays As Integer = CInt(IIf(oDBDataSource.GetValue("U_NoOfDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_NoOfDays", 0)))
                                    Dim intFreeDays As Integer = CInt(IIf(oDBDataSource.GetValue("U_FreeDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_FreeDays", 0)))
                                    Dim strPaidSt As String = oDBDataSource.GetValue("U_PaidSta", 0).ToString()

                                    'If (intOrderDays - intDelDays) > 0 Then
                                    '    oApplication.Utilities.Message("Open Order Document Exist Cannot Cancel...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    '    BubbleEvent = False
                                    '    Exit Sub
                                    'Else

                                    If intOrderDays > 0 And intDelDays <= 0 Then

                                        Dim _retVal As Integer = oApplication.SBO_Application.MessageBox("Sure you wanted to Cancel Customer Program...?", 2, "Yes", "No", "")
                                        If _retVal = 2 Then
                                            Exit Sub
                                        End If

                                        Dim strDocEntry As String = oDBDataSource.GetValue("DocEntry", 0).Trim().ToString

                                        'Close Open SalesOrder related to Program.
                                        oApplication.Utilities.closeOpenOrdersFromProgramifCancelled(strDocEntry)
                                        oApplication.SBO_Application.MessageBox("Open Sale Order Closed Successfully...")

                                        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        strQuery = " Update [@Z_OCPM] SET U_Cancel = 'Y',U_DocStatus = 'L',U_RemDays = 0, "
                                        strQuery += " U_CRemarks = '" & strCancelRemarks & "'"
                                        strQuery += " Where DocEntry = '" + strDocEntry + "'"
                                        oRecordSet.DoQuery(strQuery)

                                        strQuery = "PROCON_UPDATEONOFFSTATUS_u"
                                        oRecordSet.DoQuery(strQuery)

                                        oApplication.SBO_Application.MessageBox("Program Registration Canceled Successfully...")
                                        oApplication.SBO_Application.Menus.Item(mnu_ADD).Activate()

                                    Else


                                        If (intDelDays - intInvDays) > 0 Then
                                            oApplication.Utilities.Message("Open Delivery Document Exist Cannot Cancel...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        ElseIf intRemDays = 0 Then
                                            oApplication.Utilities.Message("No Remaining Days Exist for the Document to Cancel...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        Else

                                            If (intNoofDays + intFreeDays) > intRemDays Then
                                                If strPaidSt = "O" Then
                                                    oApplication.Utilities.Message("Paid Status of the Program Should be Paid for Program Cancel...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    BubbleEvent = False
                                                    Exit Sub
                                                End If
                                            End If

                                            Dim _retVal As Integer = oApplication.SBO_Application.MessageBox("Sure you wanted to Cancel Customer Program...?", 2, "Yes", "No", "")
                                            If _retVal = 2 Then
                                                Exit Sub
                                            End If

                                            Dim strDocEntry As String = oDBDataSource.GetValue("DocEntry", 0).Trim().ToString

                                            'Close Open SalesOrder related to Program.
                                            oApplication.Utilities.closeOpenOrdersFromProgramifCancelled(strDocEntry)
                                            oApplication.SBO_Application.MessageBox("Open Sale Order Closed Successfully...")

                                            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            strQuery = " Update [@Z_OCPM] SET U_Cancel = 'Y',U_DocStatus = 'L',U_RemDays = 0, "
                                            strQuery += " U_CRemarks = '" & strCancelRemarks & "'"
                                            strQuery += " Where DocEntry = '" + strDocEntry + "'"
                                            oRecordSet.DoQuery(strQuery)

                                            strQuery = "PROCON_UPDATEONOFFSTATUS_u"
                                            oRecordSet.DoQuery(strQuery)

                                            oApplication.SBO_Application.MessageBox("Program Registration Canceled Successfully...")
                                            oApplication.SBO_Application.Menus.Item(mnu_ADD).Activate()

                                        End If

                                    End If


                                    
                                End If
                            End If
                        Case mnu_DELETE_ROW
                            If oForm.PaneLevel = 1 Then
                                oMatrix = oForm.Items.Item("3").Specific
                                oMatrix.FlushToDataSource()
                                oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM6")
                                If intSelectedMatrixrow > 0 Then
                                    Dim intOrderDays As Integer = CInt(IIf(oDBDataSourceLines_0.GetValue("U_OrdDays", intSelectedMatrixrow - 1).ToString() = "", 0, oDBDataSourceLines_0.GetValue("U_OrdDays", intSelectedMatrixrow - 1)))
                                    If intOrderDays > 0 Then
                                        oApplication.Utilities.Message("Order Created for the Selected Row....Cannot Remove...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            ElseIf oForm.PaneLevel = 2 Then
                                oMatrix = oForm.Items.Item("40").Specific
                                oMatrix.FlushToDataSource()
                                oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM7")
                                For index As Integer = 1 To oMatrix.VisualRowCount
                                    If intSelectedMatrixrow <= oMatrix.VisualRowCount Then
                                        If intSelectedMatrixrow > 0 Then
                                            If oApplication.Utilities.getMatrixValues(oMatrix, "V_8", intSelectedMatrixrow).ToString() = "Y" Then
                                                oApplication.Utilities.Message("Invoice Created for the Service....Cannot Remove Row ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        Case mnu_ADD
                            For index As Integer = 1 To oMatrix.VisualRowCount
                                oMatrix.CommonSetting.SetRowEditable(index, True)
                            Next
                            oMatrix.Columns.Item("V_-1").Editable = False
                            oMatrix.Columns.Item("V_2").Editable = False
                            oMatrix.Columns.Item("V_6").Editable = False
                            oMatrix.Columns.Item("V_8").Editable = False
                            oMatrix.Columns.Item("V_9").Editable = False
                            oMatrix.Columns.Item("V_10").Editable = False
                        Case mnu_ADD_ROW
                            'For index As Integer = 1 To oMatrix.VisualRowCount
                            '    If oApplication.Utilities.getMatrixValues(oMatrix, "V_7", index).ToString() = "N" Then
                            '        oMatrix.CommonSetting.SetRowEditable(index, True)
                            '    Else
                            '        oMatrix.CommonSetting.SetRowEditable(index, False)
                            '    End If
                            'Next
                            oMatrix.Columns.Item("V_-1").Editable = False
                            oMatrix.Columns.Item("V_2").Editable = False
                            oMatrix.Columns.Item("V_6").Editable = False
                            oMatrix.Columns.Item("V_8").Editable = False
                            oMatrix.Columns.Item("V_9").Editable = False
                            oMatrix.Columns.Item("V_10").Editable = False
                    End Select
                Case False
                    Select Case pVal.MenuUID
                        Case mnu_Z_OCPM
                            LoadForm()
                        Case mnu_ADD
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            initialize(oForm)
                        Case mnu_ADD_ROW
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            AddRow(oForm)
                            addDayRowsDynamically(oForm)
                            'Case mnu_ADD_ROW
                            oMatrix = oForm.Items.Item("3").Specific
                            'For index As Integer = 1 To oMatrix.VisualRowCount
                            '    If oApplication.Utilities.getMatrixValues(oMatrix, "V_7", index).ToString() = "N" Then
                            '        oMatrix.CommonSetting.SetRowEditable(index, True)
                            '    Else
                            '        oMatrix.CommonSetting.SetRowEditable(index, False)
                            '    End If
                            'Next
                            oMatrix.Columns.Item("V_-1").Editable = False
                            oMatrix.Columns.Item("V_2").Editable = False
                            oMatrix.Columns.Item("V_6").Editable = False
                            oMatrix.Columns.Item("V_8").Editable = False
                            oMatrix.Columns.Item("V_9").Editable = False
                            oMatrix.Columns.Item("V_10").Editable = False
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        Case mnu_DELETE_ROW
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            RefereshDeleteRow(oForm)
                            addDayRowsDynamically(oForm)
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End Select
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Data Events"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            If oForm.TypeEx = frm_Z_OCPM Then
                Select Case BusinessObjectInfo.BeforeAction
                    Case True

                    Case False
                        Select Case BusinessObjectInfo.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, _
                            SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                If BusinessObjectInfo.ActionSuccess Then
                                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                                    Dim oXmlDoc As System.Xml.XmlDocument = New Xml.XmlDocument()
                                    oXmlDoc.LoadXml(BusinessObjectInfo.ObjectKey)
                                    Dim strDocEntry As String = oXmlDoc.SelectSingleNode("/Customer_ProgramParams/DocEntry").InnerText
                                    oApplication.Utilities.updateCustomerProgramInProgram(strDocEntry)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                                Try
                                    oForm.Freeze(True)
                                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPM")
                                    oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM6")
                                    oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@Z_CPM7")
                                    If oForm.DataSources.DataTables.Count = 0 Then
                                        oForm.DataSources.DataTables.Add("Documents")
                                        oForm.DataSources.DataTables.Add("ProgramDL")
                                    End If

                                    If (oDBDataSource.GetValue("U_DocStatus", oDBDataSource.Offset) = "L" _
                                        Or oDBDataSource.GetValue("U_DocStatus", oDBDataSource.Offset) = "C") Then

                                        oForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'From Date
                                        oForm.Items.Item("13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'To Date
                                        oForm.Items.Item("14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'No of Days
                                        oForm.Items.Item("15").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Free Days
                                        oForm.Items.Item("17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Discount
                                        oForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Paid Status
                                        oForm.Items.Item("3").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Matrix
                                        oForm.Items.Item("40").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Matrix
                                        If oDBDataSource.GetValue("U_DocStatus", oDBDataSource.Offset).Trim() = "L" Then
                                            oForm.Items.Item("_31").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Invoice Create Button
                                        End If
                                        oForm.Items.Item("45").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Document Currency
                                        oForm.Items.Item("46").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Cancel Remarks
                                        oForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Document No.
                                        oForm.Items.Item("50").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Sequence Status.
                                    Else

                                        oForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'From Date
                                        oForm.Items.Item("13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'To Date
                                        oForm.Items.Item("14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'No of Days
                                        oForm.Items.Item("15").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Free Days
                                        oForm.Items.Item("17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True) 'Discount
                                        oForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True) 'Paid Status
                                        oForm.Items.Item("3").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True) 'Matrix
                                        oForm.Items.Item("40").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True) 'Matrix
                                        oForm.Items.Item("_31").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True) 'Invoice Create Button

                                        Dim strInvDays As String = oDBDataSource.GetValue("U_InvDays", oDBDataSource.Offset)
                                        If CInt(IIf(strInvDays = "", 0, strInvDays)) > 0 Then
                                            oForm.Items.Item("45").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Document Currency
                                        Else
                                            oForm.Items.Item("45").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True) 'Document Currency
                                        End If
                                        oForm.Items.Item("46").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True) 'Cancel Remarks

                                        oMatrix = oForm.Items.Item("3").Specific
                                        For index As Integer = 1 To oMatrix.VisualRowCount
                                            If CType(oMatrix.Columns.Item("V_8").Cells.Item(index).Specific, SAPbouiCOM.EditText).Value <> "" Then
                                                If CType(oMatrix.Columns.Item("V_8").Cells.Item(index).Specific, SAPbouiCOM.EditText).Value <> "0" Then
                                                    oMatrix.CommonSetting.SetCellEditable(index, 1, False)
                                                    oMatrix.CommonSetting.SetCellEditable(index, 4, False)
                                                    oMatrix.CommonSetting.SetCellEditable(index, 5, False)
                                                    oMatrix.CommonSetting.SetCellEditable(index, 6, False)
                                                    oMatrix.CommonSetting.SetCellEditable(index, 7, False)
                                                    oMatrix.CommonSetting.SetCellEditable(index, 8, False)
                                                    'oMatrix.CommonSetting.SetCellEditable(index, 9, False)
                                                Else
                                                    oMatrix.CommonSetting.SetCellEditable(index, 1, True)
                                                    oMatrix.CommonSetting.SetCellEditable(index, 4, True)
                                                    oMatrix.CommonSetting.SetCellEditable(index, 5, True)
                                                    oMatrix.CommonSetting.SetCellEditable(index, 6, True)
                                                    oMatrix.CommonSetting.SetCellEditable(index, 7, True)
                                                    oMatrix.CommonSetting.SetCellEditable(index, 8, True)
                                                    'oMatrix.CommonSetting.SetCellEditable(index, 9, True)
                                                End If
                                            Else
                                                oMatrix.CommonSetting.SetCellEditable(index, 1, True)
                                                oMatrix.CommonSetting.SetCellEditable(index, 4, True)
                                                oMatrix.CommonSetting.SetCellEditable(index, 5, True)
                                                oMatrix.CommonSetting.SetCellEditable(index, 6, True)
                                                oMatrix.CommonSetting.SetCellEditable(index, 7, True)
                                                oMatrix.CommonSetting.SetCellEditable(index, 8, True)
                                                'oMatrix.CommonSetting.SetCellEditable(index, 9, True)
                                            End If
                                            'Dim strNoofDays As String = CType(oMatrix.Columns.Item("V_1").Cells.Item(index).Specific, SAPbouiCOM.EditText).Value
                                            Dim strOrdDays As String = CType(oMatrix.Columns.Item("V_8").Cells.Item(index).Specific, SAPbouiCOM.EditText).Value
                                            'If CInt(IIf(strNoofDays = "", 0, strNoofDays)) = CInt(IIf(strInvDays = "", 0, strInvDays)) Then
                                            '    oMatrix.CommonSetting.SetCellEditable(index, 2, False)
                                            'Else
                                            '    oMatrix.CommonSetting.SetCellEditable(index, 2, True)
                                            'End If
                                            If index > 1 Then
                                                If CInt(IIf(strOrdDays = "", 0, strOrdDays)) > 0 Then
                                                    oMatrix.CommonSetting.SetCellEditable(index - 1, 2, False)
                                                Else
                                                    oMatrix.CommonSetting.SetCellEditable(index - 1, 2, True)
                                                End If
                                            End If
                                        Next

                                        oMatrix.Columns.Item("V_-1").Editable = False
                                        oMatrix.Columns.Item("V_2").Editable = False
                                        oMatrix.Columns.Item("V_6").Editable = False
                                        oMatrix.Columns.Item("V_8").Editable = False
                                        oMatrix.Columns.Item("V_9").Editable = False
                                        oMatrix.Columns.Item("V_10").Editable = False
                                        oForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                    End If
                                    loadDocuments(oForm)
                                    oForm.Items.Item("34").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oForm.Freeze(False)
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                End Try
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Right Click Event"

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
        If oForm.TypeEx = frm_Z_OCPM Then
            Dim oMenuItem As SAPbouiCOM.MenuItem
            Dim oMenus As SAPbouiCOM.Menus
            oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPM")
            oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM6")
            oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@Z_CPM7")

            If (eventInfo.BeforeAction = True) Then
                Try

                    'If oForm.PaneLevel = 1 And eventInfo.ItemUID = "3" Then
                    '    oMatrix = oForm.Items.Item("3").Specific
                    '    oMatrix.FlushToDataSource()
                    '    oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM6")
                    '    If eventInfo.Row > 0 Then
                    '        Dim intOrderDays As Integer = CInt(IIf(oDBDataSourceLines_0.GetValue("U_OrdDays", eventInfo.Row - 1).ToString() = "", 0, oDBDataSourceLines_0.GetValue("U_OrdDays", eventInfo.Row - 1)))
                    '        If intOrderDays > 0 Then
                    '            'oApplication.Utilities.Message("Order Created for the Selected Row....Cannot Remove...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            BubbleEvent = False
                    '            Exit Sub
                    '        End If
                    '    End If
                    'ElseIf oForm.PaneLevel = 2 And eventInfo.ItemUID = "40" Then
                    '    oMatrix = oForm.Items.Item("40").Specific
                    '    For index As Integer = 1 To oMatrix.VisualRowCount
                    '        If eventInfo.Row <= oMatrix.VisualRowCount Then
                    '            If eventInfo.Row > 0 Then
                    '                If oApplication.Utilities.getMatrixValues(oMatrix, "V_8", eventInfo.Row).ToString() = "Y" Then
                    '                    'oApplication.Utilities.Message("Invoice Created for the Service....Cannot Remove Row ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '                    BubbleEvent = False
                    '                    Exit Sub
                    '                End If
                    '            End If
                    '        End If
                    '    Next
                    'End If

                    If Not oMenuItem.SubMenus.Exists(mnu_CANCELCP) And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                        oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage.UniqueID = mnu_CANCELCP
                        oCreationPackage.String = "Cancel Program"
                        If oDBDataSource.GetValue("U_Cancel", oDBDataSource.Offset).Trim() = "Y" Then
                            oCreationPackage.Enabled = False
                        End If
                        oMenus = oMenuItem.SubMenus
                        oMenus.AddEx(oCreationPackage)
                    End If

                    If Not oMenuItem.SubMenus.Exists(mnu_CLOSECP) And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                        oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage.UniqueID = mnu_CLOSECP
                        oCreationPackage.String = "Close Program"
                        oCreationPackage.Enabled = True
                        oMenus = oMenuItem.SubMenus
                        oMenus.AddEx(oCreationPackage)
                    End If

                    oMenuItem.SubMenus.Item(mnu_CANCEL).Enabled = False
                    oMenuItem.SubMenus.Item(mnu_CLOSE).Enabled = False
                    If eventInfo.ItemUID = "3" Or eventInfo.ItemUID = "40" Then
                        oMenuItem.SubMenus.Item(mnu_ADD_ROW).Enabled = True
                        oMenuItem.SubMenus.Item(mnu_DELETE_ROW).Enabled = True
                    Else
                        'oMenuItem.SubMenus.Item(mnu_ADD_ROW).Enabled = False
                        'oMenuItem.SubMenus.Item(mnu_DELETE_ROW).Enabled = False
                    End If

                Catch ex As Exception
                    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try
            Else
                If oMenuItem.SubMenus.Exists(mnu_GenerateSO) Then
                    oMenuItem.SubMenus.RemoveEx(mnu_GenerateSO)
                End If
                If oMenuItem.SubMenus.Exists(mnu_ViewSO) Then
                    oMenuItem.SubMenus.RemoveEx(mnu_ViewSO)
                End If
                If oMenuItem.SubMenus.Exists(mnu_CANCELCP) Then
                    oMenuItem.SubMenus.RemoveEx(mnu_CANCELCP)
                End If
                If oMenuItem.SubMenus.Exists(mnu_CLOSECP) Then
                    oMenuItem.SubMenus.RemoveEx(mnu_CLOSECP)
                End If
                oMenuItem.SubMenus.Item(mnu_CANCEL).Enabled = True
                oMenuItem.SubMenus.Item(mnu_CLOSE).Enabled = True
            End If
        End If
    End Sub

#End Region

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPM")
            oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM6")
            oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@Z_CPM7")

            oMatrix = oForm.Items.Item("3").Specific
            oMatrix.LoadFromDataSource()
            oMatrix.AddRow(1, -1)
            oMatrix.FlushToDataSource()
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            oMatrix = oForm.Items.Item("40").Specific
            oMatrix.LoadFromDataSource()
            oMatrix.AddRow(1, -1)
            oMatrix.FlushToDataSource()
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single



            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select IsNull(MAX(DocEntry),0) +1 From [@Z_OCPM]")
            If Not oRecordSet.EoF Then
                oApplication.Utilities.setEditText(oForm, "9", oRecordSet.Fields.Item(0).Value.ToString())
                oForm.Items.Item("8").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                oApplication.Utilities.setEditText(oForm, "8", "t")
                oApplication.SBO_Application.SendKeys("{TAB}")
            End If
            oForm.Items.Item("34").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("9").Enabled = False

            MatrixId = "3"

            oForm.Items.Item("7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Customer 
            oForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Doc Date
            oForm.Items.Item("11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Program
            oForm.Items.Item("11_").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Program
            oForm.Items.Item("_31").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True) 'Invoice Button
            'oForm.Items.Item("2_").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            oForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim blnStatus As Boolean = False
        Try
            oMatrix = oForm.Items.Item("3").Specific
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPM")
            oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM6")
            oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@Z_CPM7")

            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If oApplication.Utilities.getEditTextvalue(aForm, "6") = "" Then
                oApplication.Utilities.Message("Enter Customer Code...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "7") = "" Then
                oApplication.Utilities.Message("Enter Customer Name...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "11") = "" Then
                oApplication.Utilities.Message("Enter Program...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "12") = "" Then
                oApplication.Utilities.Message("Enter Program Start Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "13") = "" Then
                oApplication.Utilities.Message("Enter Program End Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Dim strFromDate As String = CType(oForm.Items.Item("12").Specific, SAPbouiCOM.EditText).Value
            Dim strToDate As String = CType(oForm.Items.Item("13").Specific, SAPbouiCOM.EditText).Value
            Dim strNoofDays As String = CType(oForm.Items.Item("14").Specific, SAPbouiCOM.EditText).Value
            Dim strFreeDays As String = CType(oForm.Items.Item("15").Specific, SAPbouiCOM.EditText).Value
            Dim strDocEntry As String = CType(oForm.Items.Item("10").Specific, SAPbouiCOM.EditText).Value
            Dim strDocStatus As String = oDBDataSource.GetValue("U_DocStatus", 0)
            Dim strTransRef As String = oDBDataSource.GetValue("U_TrnRef", 0)

            'If strDocStatus = "L" Or (strTransRef <> "0" And strTransRef.Length > 0) Then
            '    Return True
            'End If

            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                If Not oApplication.Utilities.validateDate(oForm, strFromDate, 0) Then
                    oApplication.Utilities.Message("Program From Date Should be Greater than Or Equal Current Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If

            If CInt(strFromDate) > CInt(strToDate) Then
                oApplication.Utilities.Message("From Date Should be Less than or Equal to To Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Dim blnRowAdded As Boolean = False
            Dim blnBrkDateEx As Boolean = False
            Dim dblNoofDays As Double = 0
            Dim dblFreeDays As Double = 0
            Dim dblInvoiceDays As Double = 0

            If oMatrix.RowCount = 0 Then
                oApplication.Utilities.Message("Enter Invoice Break Up to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                If 1 = 1 Then
                    For index As Integer = 1 To oMatrix.VisualRowCount

                        If oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index).ToString() <> "" Then
                            blnRowAdded = True
                        End If
                        If oApplication.Utilities.getMatrixValues(oMatrix, "V_3", index).ToString() = "P" Then
                            dblNoofDays += CDbl(IIf(oApplication.Utilities.getMatrixValues(oMatrix, "V_1", index).ToString() = "", 0, oApplication.Utilities.getMatrixValues(oMatrix, "V_1", index).ToString()))
                        End If
                        If oApplication.Utilities.getMatrixValues(oMatrix, "V_3", index).ToString() = "F" Then
                            dblFreeDays += CDbl(IIf(oApplication.Utilities.getMatrixValues(oMatrix, "V_1", index).ToString() = "", 0, oApplication.Utilities.getMatrixValues(oMatrix, "V_1", index).ToString()))
                        End If

                        Dim dblRowQty As Double
                        dblRowQty = (IIf(oApplication.Utilities.getMatrixValues(oMatrix, "V_1", index).ToString() = "", 0, oApplication.Utilities.getMatrixValues(oMatrix, "V_1", index).ToString()))

                        Dim strFDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index).ToString()
                        'If Not oApplication.Utilities.validateDate(oForm, strFDate, 0) Then
                        '    oApplication.Utilities.Message("Program From Date Should be Greater than Or Equal Current Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    Return False
                        'End If

                        Dim strTDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_2", index).ToString()
                        If strFDate <> "" And strTDate <> "" And dblRowQty > 0 Then
                            If (CInt(strFromDate) > CInt(strFDate)) Or (CInt(strFDate) > CInt(strToDate)) Then
                                blnBrkDateEx = True
                            ElseIf (CInt(strFromDate) > CInt(strTDate)) Or (CInt(strTDate) > CInt(strToDate)) Then
                                blnBrkDateEx = True
                            End If
                        End If

                        If oApplication.Utilities.getMatrixValues(oMatrix, "V_10", index).ToString() <> "" Then
                            dblInvoiceDays += CDbl(IIf(oApplication.Utilities.getMatrixValues(oMatrix, "V_10", index).ToString() = "", 0, oApplication.Utilities.getMatrixValues(oMatrix, "V_10", index).ToString()))
                        End If

                    Next
                End If
            End If

            If Not blnRowAdded Then
                oApplication.Utilities.Message("Enter Invoice Break Up to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If blnBrkDateEx Then
                oApplication.Utilities.Message("Invoice Break Up Days Should be between Program From & To Defined in Header...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If (dblNoofDays > CDbl(IIf(strNoofDays = "", "0", strNoofDays))) Or (dblNoofDays < CDbl(IIf(strNoofDays = "", "0", strNoofDays))) Then
                oApplication.Utilities.Message("Invoice Break Up No of Days Equal to No of Program Days...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If (dblFreeDays > CDbl(IIf(strFreeDays = "", "0", strFreeDays))) Or (dblFreeDays < CDbl(IIf(strFreeDays = "", "0", strFreeDays))) Then
                oApplication.Utilities.Message("Invoice Break Up No of Days Should Equal to No of Free Days...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Dim dtFromdate As Date = oApplication.Utilities.GetDateTimeValue(strFromDate)
            Dim dtToDate As Date = oApplication.Utilities.GetDateTimeValue(strToDate)

            strQuery = "Select DocEntry from [@Z_OCPM] where U_CardCode='" & oApplication.Utilities.getEditTextvalue(aForm, "6") & "'" & _
                " And '" & dtFromdate.ToString("yyyy-MM-dd") & "' between U_PFromDate and U_PToDate And IsNull(U_Cancel,'N') = 'N' " & _
                " And ISNULL(U_Transfer,'N') = 'N' " & _
                    " And ISNULL(U_DocStatus,'O') = 'O' "

            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                strQuery &= " And DocEntry <> '" & strDocEntry & "'"
            End If

            oTest.DoQuery(strQuery)
            If oTest.RecordCount > 0 Then
                oApplication.Utilities.Message("Program From date is overlapped with another program for selected customer", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            strQuery = "Select DocEntry from [@Z_OCPM] where U_CardCode= '" & oApplication.Utilities.getEditTextvalue(aForm, "6") & "'" & _
                " And '" & dtToDate.ToString("yyyy-MM-dd") & "' between U_PFromDate and U_PToDate And IsNull(U_Cancel,'N') = 'N' " & _
                " And ISNULL(U_Transfer,'N') = 'N' " & _
            " And ISNULL(U_DocStatus,'O') = 'O' "
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                strQuery &= " And DocEntry <> '" & strDocEntry & "'"
            End If
            oTest.DoQuery(strQuery)
            If oTest.RecordCount > 0 Then
                oApplication.Utilities.Message("Program End date is overlapped with another program for selected customer", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If (CDbl(IIf(strNoofDays = "", "0", strNoofDays)) + CDbl(IIf(strFreeDays = "", "0", strFreeDays))) < dblInvoiceDays Then
                oApplication.Utilities.Message("Already Invoice Created For : " & dblInvoiceDays & "  Days...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            oMatrix = oForm.Items.Item("3").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                If oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index).ToString() <> "" Then
                    Dim strTaxCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_5_", index).ToString()
                    If strTaxCode.Length = 0 Then
                        oApplication.Utilities.Message("Select the TaxCode Invoice Break Up Row No : " & index.ToString & "", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            Next

            oMatrix = oForm.Items.Item("40").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim strItemCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index).ToString()
                Dim strDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0_0", index).ToString()
                Dim strTaxCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_5_", index).ToString()

                'If Not oApplication.Utilities.validateDate(oForm, strDate, 0) Then
                '    oApplication.Utilities.Message("Applied From Date Should be Greater than Or Equal Current Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Return False
                'End If

                If strItemCode.Length > 0 Then
                    If strDate.Length = 0 Then
                        oApplication.Utilities.Message("Select the Applied Date in Service Items Row No : " & index.ToString & "", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf strTaxCode.Length = 0 Then
                        oApplication.Utilities.Message("Select the TaxCode Service Items Up Row No : " & index.ToString & "", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If

            Next

            'validation_Service(oForm)

            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Private Function validation_Service(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim blnStatus As Boolean = False
        Try
            oMatrix = oForm.Items.Item("3").Specific
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPM")
            oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@Z_CPM7")

            Dim strToDate As String = CType(oForm.Items.Item("13").Specific, SAPbouiCOM.EditText).Value
            Dim strNoofDays As String = CType(oForm.Items.Item("14").Specific, SAPbouiCOM.EditText).Value
          
            oMatrix = oForm.Items.Item("40").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim dblNoofDay, dblQty As Double
                Dim strItemCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index).ToString()
                Dim strDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0_0", index).ToString()
                Dim strTaxCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_5_", index).ToString()
                Dim strQuantity As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_2", index).ToString()
                Dim strInvoiceCreated As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_8", index).ToString()
                Double.TryParse(strNoofDays, dblNoofDay)
                Double.TryParse(strQuantity, dblQty)

                If strItemCode.Length > 0 Then
                    If strInvoiceCreated = "N" Then
                        If dblNoofDay <> dblQty Then
                            'Return False
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", index, strNoofDays)
                        End If
                        If strDate <> strToDate Then
                            'Return False
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_0_0", index, strToDate)
                        End If
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    'Private Sub addInvoiceServiceItemsAutomatically(ByVal oForm As SAPbouiCOM.Form)
    '    Try
    '        oMatrix = oForm.Items.Item("3").Specific
    '        Dim blnAutoAdded As Boolean = False
    '        For index As Integer = 1 To oMatrix.RowCount
    '            If oApplication.Utilities.getMatrixValues(oMatrix, "V_2", index).ToString() = "" Then
    '                blnAutoAdded = True
    '            End If
    '        Next
    '        If blnAutoAdded Then
    '            For index As Integer = 1 To oMatrix.RowCount
    '                Dim strMenu As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
    '                Dim strRef As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_2", index)
    '                If strMenu.Length > 0 Then
    '                    If strRef.Length = 0 Then
    '                        strRef = oApplication.Utilities.AddServiceItemDocument(oForm)
    '                        oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", index, strRef)
    '                    End If
    '                End If
    '            Next
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    'Private Sub removeInvoiceSericeItemsAutomatically(ByVal oForm As SAPbouiCOM.Form)
    '    Try
    '        oMatrix = oForm.Items.Item("3").Specific
    '        Dim blnAutoRemove As Boolean = False
    '        For index As Integer = 1 To oMatrix.RowCount
    '            If oApplication.Utilities.getMatrixValues(oMatrix, "V_2", index).ToString() <> "" Then
    '                blnAutoRemove = True
    '            End If
    '        Next
    '        If blnAutoRemove Then
    '            For index As Integer = 1 To oMatrix.RowCount
    '                Dim strRef As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_2", index)
    '                If strRef.Length = 0 Then
    '                    oApplication.Utilities.RemoveServiceItemDocument(oForm, strRef)
    '                End If
    '            Next
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    Private Sub calculateRNoofDays(ByVal oForm As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPM")

            Dim strNoofDays As String = CType(oForm.Items.Item("14").Specific, SAPbouiCOM.EditText).Value
            Dim strFreeDays As String = CType(oForm.Items.Item("15").Specific, SAPbouiCOM.EditText).Value
            Dim strDelDays As String = CInt(IIf(oDBDataSource.GetValue("U_DelDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_DelDays", 0)))
            Dim strCardCode As String = CType(oForm.Items.Item("6").Specific, SAPbouiCOM.EditText).Value
            Dim strProgFdt As String = CType(oForm.Items.Item("12").Specific, SAPbouiCOM.EditText).Value

            If strNoofDays.Length > 0 Or strFreeDays.Length > 0 Then
                Dim strTotalDays As String = (CInt(IIf(strNoofDays = "", "0", strNoofDays)) + CInt(IIf(strFreeDays = "", "0", strFreeDays))).ToString
                Dim strPrgToDate As String = oApplication.Utilities.getProgramToDate(oForm, strCardCode, strProgFdt, strTotalDays)
                CType(oForm.Items.Item("13").Specific, SAPbouiCOM.EditText).Value = strPrgToDate
                CType(oForm.Items.Item("16").Specific, SAPbouiCOM.EditText).Value = (CInt(IIf(strTotalDays = "", "0", strTotalDays)) - CInt(IIf(strDelDays = "", "0", strDelDays))).ToString
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub calculateRNoofDays(ByVal oForm As SAPbouiCOM.Form, ByVal iRow As Integer)
        Try
            oMatrix = oForm.Items.Item("3").Specific

            Dim strCardCode As String = CType(oForm.Items.Item("6").Specific, SAPbouiCOM.EditText).Value
            Dim strProgFdt As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", iRow).ToString()
            Dim strNoofDays As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", iRow).ToString()

            If strNoofDays.Length > 0 Then
                Dim strTotalDays As String = CInt(IIf(strNoofDays = "", "0", strNoofDays))
                Dim strPrgToDate As String = oApplication.Utilities.getProgramToDate(oForm, strCardCode, strProgFdt, strTotalDays)
                If strTotalDays = "0" Then
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", iRow, strProgFdt)
                Else
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", iRow, strPrgToDate)
                End If
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Select Case aForm.PaneLevel
                Case "1"
                    oMatrix = aForm.Items.Item("3").Specific
                    oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM6")
                    If oMatrix.RowCount <= 0 Then
                        oMatrix.AddRow()
                    Else
                        oMatrix.AddRow(1, oMatrix.RowCount + 1)
                        oMatrix.ClearRowData(oMatrix.RowCount)
                        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If
                    End If
                    oMatrix.FlushToDataSource()
                    For count = 1 To oDBDataSourceLines_0.Size
                        oDBDataSourceLines_0.SetValue("LineId", count - 1, count)
                    Next
                    oMatrix.LoadFromDataSource()
                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    AssignLineNo(aForm)
                Case "2"
                    oMatrix = aForm.Items.Item("40").Specific
                    oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@Z_CPM7")
                    If oMatrix.RowCount <= 0 Then
                        oMatrix.AddRow()
                    Else
                        oMatrix.AddRow(1, oMatrix.RowCount + 1)
                        oMatrix.ClearRowData(oMatrix.RowCount)
                        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If
                    End If
                    oMatrix.FlushToDataSource()
                    For count = 1 To oDBDataSourceLines_1.Size
                        oDBDataSourceLines_1.SetValue("LineId", count - 1, count)
                    Next
                    oMatrix.LoadFromDataSource()
                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    AssignLineNo(aForm)
            End Select
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub deleterow(ByVal aForm As SAPbouiCOM.Form)
        Try
            Select Case aForm.PaneLevel
                Case "1"
                    oMatrix = aForm.Items.Item("3").Specific
                    oDBDataSourceLines_0 = aForm.DataSources.DBDataSources.Item("@Z_CPM6")
                Case "2"
                    oMatrix = aForm.Items.Item("40").Specific
                    oDBDataSourceLines_0 = aForm.DataSources.DBDataSources.Item("@Z_CPM7")
            End Select
            oMatrix.FlushToDataSource()
            For introw As Integer = 1 To oMatrix.RowCount
                If oMatrix.IsRowSelected(introw) Then
                    oMatrix.DeleteRow(introw)
                    oDBDataSourceLines_0.RemoveRecord(introw - 1)
                    oMatrix.FlushToDataSource()
                    For count As Integer = 1 To oDBDataSourceLines_0.Size
                        oDBDataSourceLines_0.SetValue("LineId", count - 1, count)
                    Next
                    Select Case aForm.PaneLevel
                        Case "1"
                            oMatrix = aForm.Items.Item("3").Specific
                            oDBDataSourceLines_0 = aForm.DataSources.DBDataSources.Item("@Z_CPM6")
                            AssignLineNo(aForm)
                        Case "2"
                            oMatrix = aForm.Items.Item("40").Specific
                            oDBDataSourceLines_0 = aForm.DataSources.DBDataSources.Item("@Z_CPM7")
                            AssignLineNo(aForm)
                    End Select
                    oMatrix.LoadFromDataSource()
                    If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Select Case aForm.PaneLevel
                Case "1"
                    oMatrix = aForm.Items.Item("3").Specific
                    oDBDataSourceLines_0 = aForm.DataSources.DBDataSources.Item("@Z_CPM6")
                Case "2"
                    oMatrix = aForm.Items.Item("40").Specific
                    oDBDataSourceLines_0 = aForm.DataSources.DBDataSources.Item("@Z_CPM7")
            End Select
            oMatrix = aForm.Items.Item("3").Specific
            oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM6")
            oMatrix.FlushToDataSource()
            For count = 1 To oDBDataSourceLines_0.Size - 1
                oDBDataSourceLines_0.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            Select Case aForm.PaneLevel
                Case "1"
                    oMatrix = aForm.Items.Item("3").Specific
                    oDBDataSourceLines_0 = aForm.DataSources.DBDataSources.Item("@Z_CPM6")
                Case "2"
                    oMatrix = aForm.Items.Item("40").Specific
                    oDBDataSourceLines_0 = aForm.DataSources.DBDataSources.Item("@Z_CPM7")
            End Select

            Me.RowtoDelete = intSelectedMatrixrow
            If Me.RowtoDelete - 1 >= 0 Then
                oDBDataSourceLines_0.RemoveRecord(Me.RowtoDelete - 1)
                oMatrix.LoadFromDataSource()
                oMatrix.FlushToDataSource()
                For count = 1 To oDBDataSourceLines_0.Size - 1
                    oDBDataSourceLines_0.SetValue("LineId", count - 1, count)
                Next
                oMatrix.LoadFromDataSource()
            End If
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub addChooseFromListConditions(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList

            oCFLs = oForm.ChooseFromLists

            'oCFL = oCFLs.Item("CFL_1")
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "CardType"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "C"
            'oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_1")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.BracketOpenNum = 2
            For i As Integer = 0 To 2

                If i = 1 Then
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCon = oCons.Add()
                    oCon.BracketOpenNum = 1
                End If

                If i = 0 Then
                    oCon.[Alias] = "CardType"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "C"
                ElseIf i = 1 Then
                    oCon.[Alias] = "validFor"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 2 Then

                    'Modified by Madhu for DIET PHASE II On 20150710.
                    'oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    'oRecordSet.DoQuery("Select Series From NNM1 Where ObjectCode = '2' And SeriesName Like 'CR%'")
                    'If oRecordSet.RecordCount > 0 Then
                    '    While Not oRecordSet.EoF
                    '        oCon.[Alias] = "Series"
                    '        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    '        oCon.CondVal = oRecordSet.Fields.Item(0).Value.ToString()
                    '        oRecordSet.MoveNext()
                    '    End While
                    'End If

                    oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery("Select U_Prefix From [@Z_OFCI] Where U_Type = 'C' And U_Active = 'Y' ")
                    If oRecordSet.RecordCount > 0 Then
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                        oCon = oCons.Add()
                        oCon.BracketOpenNum = 2
                        Dim intConCount As Integer = 0

                        While Not oRecordSet.EoF

                            If intConCount > 0 Then
                                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                                oCon = oCons.Add()
                                oCon.BracketOpenNum = 1
                            End If

                            oCon.[Alias] = "CardCode"
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_START
                            oCon.CondVal = oRecordSet.Fields.Item(0).Value.ToString()
                            oCon.BracketCloseNum = 1
                            oRecordSet.MoveNext()
                            intConCount += 1

                        End While
                        oCon.BracketCloseNum = 2
                    End If

                End If

                If i = 0 Then
                    oCon.BracketCloseNum = 2
                ElseIf i = 1 Then
                    oCon.BracketCloseNum = 1
                End If

            Next

            oCFL.SetConditions(oCons)

            'MessageBox.Show(oCFL.GetConditions().GetAsXML())
            'strQuery = "Select ItmsGrpCod From OITB Where U_Program = 'Y' "
            'oRecordSet.DoQuery(strQuery)
            'If Not oRecordSet.EoF Then
            '    Dim strIG As String = oRecordSet.Fields.Item(0).Value
            '    oCFL = oCFLs.Item("CFL_2")
            '    oCons = oCFL.GetConditions()
            '    oCon = oCons.Add()
            '    oCon.[Alias] = "ItmsGrpCod"
            '    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '    oCon.CondVal = strIG
            '    oCFL.SetConditions(oCons)
            'End If

            strQuery = "Select ItmsGrpCod From OITB Where U_Program = 'Y' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oCFL = oCFLs.Item("CFL_2")
                oCons = oCFL.GetConditions()
                oCon = oCons.Add()
                oCon.BracketOpenNum = 2
                Dim intConCount As Integer = 0
                While Not oRecordSet.EoF
                    Dim strIG As String = oRecordSet.Fields.Item(0).Value
                    If intConCount > 0 Then
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                        oCon = oCons.Add()
                        oCon.BracketOpenNum = 1
                    End If
                    oCon.[Alias] = "ItmsGrpCod"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = strIG

                    oRecordSet.MoveNext()
                    If Not oRecordSet.EoF Then
                        oCon.BracketCloseNum = 1
                    End If

                    intConCount += 1
                End While
                oCon.BracketCloseNum = 2
                oCFL.SetConditions(oCons)
            End If

            strQuery = "Select ItmsGrpCod From OITB Where U_Program = 'Y' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oCFL = oCFLs.Item("CFL_2_1")
                oCons = oCFL.GetConditions()
                oCon = oCons.Add()
                oCon.BracketOpenNum = 2
                Dim intConCount As Integer = 0
                While Not oRecordSet.EoF
                    Dim strIG As String = oRecordSet.Fields.Item(0).Value
                    If intConCount > 0 Then
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                        oCon = oCons.Add()
                        oCon.BracketOpenNum = 1
                    End If
                    oCon.[Alias] = "ItmsGrpCod"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = strIG

                    oRecordSet.MoveNext()
                    If Not oRecordSet.EoF Then
                        oCon.BracketCloseNum = 1
                    End If

                    intConCount += 1
                End While
                oCon.BracketCloseNum = 2
                oCFL.SetConditions(oCons)
            End If

            strQuery = "Select ItmsGrpCod From OITB Where U_Service = 'Y' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oCFL = oCFLs.Item("CFL_4")
                oCons = oCFL.GetConditions()
                oCon = oCons.Add()
                oCon.BracketOpenNum = 2
                Dim intConCount As Integer = 0

                While Not oRecordSet.EoF
                    Dim strIG As String = oRecordSet.Fields.Item(0).Value
                    If intConCount > 0 Then
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                        oCon = oCons.Add()
                        oCon.BracketOpenNum = 1
                    End If
                    oCon.[Alias] = "ItmsGrpCod"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = strIG

                    oRecordSet.MoveNext()
                    If Not oRecordSet.EoF Then
                        oCon.BracketCloseNum = 1
                    End If

                    intConCount += 1
                End While

                oCon.BracketCloseNum = 2
                oCFL.SetConditions(oCons)
            End If

            'oCFL = oCFLs.Item("CFL_4")
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.BracketOpenNum = 2
            'For i As Integer = 0 To 4
            '    If i > 0 And i < 4 Then
            '        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            '        oCon = oCons.Add()
            '        oCon.BracketOpenNum = 1
            '    End If
            '    If i = 0 Then
            '        oCon.[Alias] = "InvntItem"
            '        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '        oCon.CondVal = "N"
            '    ElseIf i = 1 Then
            '        oCon.[Alias] = "SellItem"
            '        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '        oCon.CondVal = "Y"
            '    ElseIf i = 2 Then
            '        oCon.[Alias] = "validFor"
            '        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '        oCon.CondVal = "Y"
            '    ElseIf i = 3 Then
            '        oCon.[Alias] = "PrchseItem"
            '        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '        oCon.CondVal = "N"
            '    ElseIf i = 4 Then
            '        oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            '        oRecordSet.DoQuery("Select ItmsGrpCod From OITB Where U_Service = 'Y' ")
            '        If oRecordSet.RecordCount > 0 Then
            '            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            '            oCon = oCons.Add()
            '            oCon.BracketOpenNum = 2
            '            Dim intConCount As Integer = 0

            '            While Not oRecordSet.EoF

            '                If intConCount > 0 Then
            '                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
            '                    oCon = oCons.Add()
            '                    oCon.BracketOpenNum = 1
            '                End If

            '                oCon.[Alias] = "ItmsGrpCod"
            '                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '                oCon.CondVal = oRecordSet.Fields.Item(0).Value.ToString()
            '                oCon.BracketCloseNum = 1
            '                oRecordSet.MoveNext()
            '                intConCount += 1

            '            End While
            '            oCon.BracketCloseNum = 2
            '        End If
            '    End If
            '    If i = 0 Then
            '        oCon.BracketCloseNum = 2
            '    ElseIf i > 0 And i < 4 Then
            '        oCon.BracketCloseNum = 1
            '    End If
            'Next
            'oCFL.SetConditions(oCons)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub reDrawForm(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)

            oForm.Items.Item("36").Width = oForm.Width - 25
            oForm.Items.Item("36").Height = oForm.Items.Item("3").Height + 10

            oForm.Freeze(False)
        Catch ex As Exception
            'oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Private Sub calculatePriceAfterDis(ByVal oForm As SAPbouiCOM.Form, ByVal intRow As Integer)
        Try
            oMatrix = oForm.Items.Item("3").Specific
            oMatrix.FlushToDataSource()
            ' oMatrix.LoadFromDataSource()
            oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM6")
            Dim dblPrice, dblDiscount, dblPriceAfterDis, dblNoofDays, dblInvDays, dblBasePrice, dblItemPrice As Double
            Dim strType As String = oDBDataSourceLines_0.GetValue("U_PaidType", intRow - 1)
            Dim strICurrency = String.Empty, strLCurrency As String = String.Empty
            Dim strCardCode As String = CType(oForm.Items.Item("6").Specific, SAPbouiCOM.EditText).Value
            Dim strProgram As String = CType(oForm.Items.Item("11").Specific, SAPbouiCOM.EditText).Value
            dblNoofDays = IIf(oDBDataSourceLines_0.GetValue("U_NoofDays", intRow - 1).ToString() = "", 0, oDBDataSourceLines_0.GetValue("U_NoofDays", intRow - 1))
            dblInvDays = IIf(oDBDataSourceLines_0.GetValue("U_InvDays", intRow - 1).ToString() = "", 0, oDBDataSourceLines_0.GetValue("U_InvDays", intRow - 1))

            dblPrice = IIf(oDBDataSourceLines_0.GetValue("U_Price", intRow - 1).ToString() = "", 0, oDBDataSourceLines_0.GetValue("U_Price", intRow - 1).ToString())
            dblDiscount = IIf(oDBDataSourceLines_0.GetValue("U_Discount", intRow - 1).ToString() = "", 0, oDBDataSourceLines_0.GetValue("U_Discount", intRow - 1).ToString())
            strLCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency

            If strProgram.Length > 0 Then
                If dblInvDays = 0 Then
                    oApplication.Utilities.GetCustItemPrice(strCardCode, strProgram, System.DateTime.Now.Date, dblBasePrice, strICurrency)
                Else
                    dblBasePrice = IIf(oDBDataSourceLines_0.GetValue("U_IPrice", intRow - 1).ToString() = "", 0, oDBDataSourceLines_0.GetValue("U_IPrice", intRow - 1))
                    strICurrency = IIf(oDBDataSourceLines_0.GetValue("U_Currency", intRow - 1).ToString() = "", 0, oDBDataSourceLines_0.GetValue("U_Currency", intRow - 1))
                End If
            End If

            oDBDataSourceLines_0.SetValue("U_Currency", intRow - 1, strICurrency)
            oDBDataSourceLines_0.SetValue("U_IPrice", intRow - 1, dblBasePrice)

            If strICurrency = strLCurrency Then
                If strICurrency = oDBDataSource.GetValue("U_DocCur", 0).Trim Then
                    If dblPrice = 0 Then
                        oDBDataSourceLines_0.SetValue("U_Price", intRow - 1, dblBasePrice)
                    End If
                Else
                    getPrice(oDBDataSource.GetValue("U_DocCur", 0).Trim, strICurrency, dblBasePrice, dblItemPrice)
                    If dblPrice = 0 Then
                        oDBDataSourceLines_0.SetValue("U_Price", intRow - 1, dblItemPrice)
                    End If
                End If
            Else
                getPrice(oDBDataSource.GetValue("U_DocCur", 0).Trim, strICurrency, dblBasePrice, dblItemPrice)
                If dblPrice = 0 Then
                    oDBDataSourceLines_0.SetValue("U_Price", intRow - 1, dblItemPrice)
                End If
            End If

            dblItemPrice = oDBDataSourceLines_0.GetValue("U_Price", intRow - 1)
            dblDiscount = oDBDataSourceLines_0.GetValue("U_Discount", intRow - 1)
            If strType = "P" Then
                dblPriceAfterDis = dblItemPrice - (dblItemPrice * (dblDiscount / 100)) '(IIf(dblPrice = 0, dblItemPrice, dblPrice) - ((IIf(dblPrice = 0, dblItemPrice, dblPrice) * dblDiscount) / 100))
                oDBDataSourceLines_0.SetValue("U_LineTotal", intRow - 1, dblPriceAfterDis * dblNoofDays)
                oDBDataSourceLines_0.SetValue("U_IsIReq", intRow - 1, "Y")
            ElseIf strType = "F" Then
                oDBDataSourceLines_0.SetValue("U_Discount", intRow - 1, "100")
                oDBDataSourceLines_0.SetValue("U_LineTotal", intRow - 1, "0")
                oDBDataSourceLines_0.SetValue("U_IsIReq", intRow - 1, "N")
            End If

            oMatrix.LoadFromDataSource()
            oMatrix.FlushToDataSource()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub loadDocuments(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim strQuery As String = String.Empty

            Dim strDocEntry As String = CType(oForm.Items.Item("10").Specific, SAPbouiCOM.EditText).Value

            strQuery = " Select Distinct '1-Order' As 'Document Type',DocNum "
            strQuery += " ,Min(T1.U_DelDate) As 'From Date',Max(T1.U_DelDate) 'To Date' "
            strQuery += " ,Case  When DocStatus = 'O' Then 'Open' When DocStatus = 'C' Then 'Closed'   End As 'DocStatus', "
            strQuery += " Case  When U_CanFrom = 'M' Then 'Modified' When U_CanFrom = 'E' Then 'Exclude'  "
            strQuery += " When U_CanFrom = 'R' Then 'Remove' When U_CanFrom = 'S' Then 'Suspend' End As 'Action' "
            strQuery += " From ORDR T0 JOIN RDR1 T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " Where T1.U_ProgramID = '" + strDocEntry + "'"
            strQuery += " Group By DocNum,DocStatus,U_CanFrom  "
            strQuery += " Union All "
            strQuery += " Select Distinct '2-Delivery' As 'Document Type',DocNum "
            strQuery += " ,Min(T1.U_DelDate) As 'From Date',Max(T1.U_DelDate) 'To Date' "
            strQuery += ",Case  When DocStatus = 'O' Then 'Open' When DocStatus = 'C' Then 'Closed'   End As 'DocStatus', "
            strQuery += " (Case  When U_CanFrom = 'M' Then 'Modified' When U_CanFrom = 'E' Then 'Exclude'  "
            strQuery += " When U_CanFrom = 'R' Then 'Remove' When U_CanFrom = 'S' Then 'Suspend' End) "
            'strQuery += " + (Case  When ISNULL(U_InvNo,'') <> '' Then 'Invoice No-->' + Convert(VarChar,U_InvNo) When  ISNULL(U_InvNo,'') = '' Then '' End) "
            strQuery += " As 'Action' "
            strQuery += " From ODLN T0 JOIN DLN1 T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " Where T1.U_ProgramID = '" + strDocEntry + "'"
            strQuery += " Group By DocNum,DocStatus,U_CanFrom  "
            strQuery += " Union All "
            strQuery += " Select Distinct '3-Invoice' As 'Document Type',DocNum "
            strQuery += " ,(T1.U_FDate) As 'From Date',(T1.U_EDate) 'To Date' "
            strQuery += " ,Case  When DocStatus = 'O' Then 'Open' When DocStatus = 'C' Then 'Closed'  End As 'DocStatus','' As 'Action' "
            strQuery += " From OINV T0 JOIN INV1 T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " Where T1.U_ProgramID = '" + strDocEntry + "'"

            oDTDocument = oForm.DataSources.DataTables.Item("Documents")
            oDTDocument.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("41").Specific
            oGrid.DataTable = oDTDocument
            oGrid.Columns.Item("DocNum").TitleObject.Caption = "Document No."
            oGrid.Columns.Item("DocStatus").TitleObject.Caption = "Document Status."
            oGrid.Columns.Item("Action").TitleObject.Caption = "Action."
            oGrid.CollapseLevel = 1
            oApplication.Utilities.assignLineNo(oGrid, oForm)

            strQuery = " Select "
            strQuery += " T3.U_PrgDate, "
            strQuery += " (Case WHEN ISNULL(T3.U_AppStatus,'I') = 'I' THEN 'INCLUDE' WHEN ISNULL(T3.U_AppStatus,'I') = 'E' THEN 'EXCLUDE' END) As U_AppStatus  , "
            strQuery += " (Case WHEN ISNULL(T3.U_ONOFFSTA,'O') = 'O' THEN 'ON' WHEN ISNULL(T3.U_ONOFFSTA,'O') = 'F' THEN 'OFF' END) As U_ONOFFSTA  "
            strQuery += " From [@Z_OCPM] T0  JOIN OITM T1 On T0.U_PrgCode = T1.ItemCode  "
            strQuery += " JOIN [@Z_CPM1] T3 On T3.DocEntry = T0.DocEntry And T3.U_PrgDate Is Not Null  "
            strQuery += " And T3.U_PrgDate Between T0.U_PFromDate And T0.U_PToDate  "
            strQuery += " Where T0.DocEntry = '" + strDocEntry + "'"
            strQuery += " Order By T3.U_PrgDate "

            oDTProgram = oForm.DataSources.DataTables.Item("ProgramDL")
            oDTProgram.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("48").Specific
            oGrid.DataTable = oDTProgram
            oGrid.Columns.Item("U_PrgDate").TitleObject.Caption = "Program Date"
            oGrid.Columns.Item("U_AppStatus").TitleObject.Caption = "Include/Exclude Status."
            oGrid.Columns.Item("U_ONOFFSTA").TitleObject.Caption = "On/Off Status"
            oApplication.Utilities.assignLineNo(oGrid, oForm)

            For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(index, (index + 1).ToString())
                If oGrid.DataTable.GetValue("U_AppStatus", index).ToString() = "INCLUDE" And oGrid.DataTable.GetValue("U_ONOFFSTA", index).ToString() = "ON" Then
                    oGrid.CommonSetting.SetCellBackColor(index + 1, 1, RGB(0, 255, 0))
                Else
                    oGrid.CommonSetting.SetCellBackColor(index + 1, 1, RGB(255, 255, 255))
                End If
            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub calculatePriceAfterDis_S(ByVal oForm As SAPbouiCOM.Form, ByVal intRow As Integer)
        Try
            oMatrix = oForm.Items.Item("40").Specific
            oMatrix.FlushToDataSource()
            oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@Z_CPM7")
            Dim dblPrice, dblDiscount, dblPriceAfterDis, dblNoofDays As Double
            Dim strICurrency = String.Empty, strLCurrency As String = String.Empty
            Dim dblItemPrice, dblBasePrice As Double
            Dim strCardCode As String = CType(oForm.Items.Item("6").Specific, SAPbouiCOM.EditText).Value
            Dim strServiceItem As String = oDBDataSourceLines_1.GetValue("U_ItemCode", intRow - 1).ToString()
            Dim strIsCreated As String = oDBDataSourceLines_1.GetValue("U_InvCreated", intRow - 1).ToString()

            dblNoofDays = IIf(oDBDataSourceLines_1.GetValue("U_Quantity", intRow - 1).ToString() = "", 0, oDBDataSourceLines_1.GetValue("U_Quantity", intRow - 1))
            dblPrice = IIf(oDBDataSourceLines_1.GetValue("U_Price", intRow - 1).ToString() = "", 0, oDBDataSourceLines_1.GetValue("U_Price", intRow - 1).ToString())
            strLCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency

            If strServiceItem.Length > 0 Then
                If strIsCreated = "Y" Then
                    Exit Sub
                End If
                If strIsCreated = "N" Then
                    oApplication.Utilities.GetCustItemPrice(strCardCode, strServiceItem, System.DateTime.Now.Date, dblBasePrice, strICurrency)
                Else
                    dblBasePrice = IIf(oDBDataSourceLines_1.GetValue("U_IPrice", intRow - 1).ToString() = "", 0, oDBDataSourceLines_1.GetValue("U_IPrice", intRow - 1))
                    strICurrency = IIf(oDBDataSourceLines_1.GetValue("U_Currency", intRow - 1).ToString() = "", 0, oDBDataSourceLines_1.GetValue("U_Currency", intRow - 1))
                End If
            End If

            oDBDataSourceLines_1.SetValue("U_Currency", intRow - 1, strICurrency)
            oDBDataSourceLines_1.SetValue("U_IPrice", intRow - 1, dblBasePrice)

            If strICurrency = strLCurrency Then
                If strICurrency = oDBDataSource.GetValue("U_DocCur", 0).Trim Then
                    If dblPrice = 0 Then
                        oDBDataSourceLines_1.SetValue("U_Price", intRow - 1, dblBasePrice)
                    End If
                Else
                    getPrice(oDBDataSource.GetValue("U_DocCur", 0).Trim, strICurrency, dblBasePrice, dblItemPrice)
                    If dblPrice = 0 Then
                        oDBDataSourceLines_1.SetValue("U_Price", intRow - 1, dblBasePrice)
                    End If
                End If
            Else
                getPrice(oDBDataSource.GetValue("U_DocCur", 0).Trim, strICurrency, dblBasePrice, dblItemPrice)
                If dblPrice = 0 Then
                    oDBDataSourceLines_1.SetValue("U_Price", intRow - 1, dblBasePrice)
                End If
            End If

            oMatrix.LoadFromDataSource()
            oMatrix.FlushToDataSource()

            dblItemPrice = oDBDataSourceLines_1.GetValue("U_Price", intRow - 1)
            dblDiscount = oDBDataSourceLines_1.GetValue("U_Discount", intRow - 1)
            dblPriceAfterDis = dblItemPrice - (dblItemPrice * (dblDiscount / 100)) '(dblPrice - ((dblPrice * dblDiscount) / 100))
            oDBDataSourceLines_1.SetValue("U_LineTotal", intRow - 1, dblPriceAfterDis * dblNoofDays)
            oMatrix.LoadFromDataSource()
            oMatrix.FlushToDataSource()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub calculate_Document_Values(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim strTaxCode As String

            Dim dblTBD, dblDiscount, dblTaxRate, dblDisAmt, dblTaxAmount, dblDocTotal As Double
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPM")
            oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM6")
            oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@Z_CPM7")

            oMatrix = oForm.Items.Item("3").Specific
            oMatrix.FlushToDataSource()
            oMatrix.LoadFromDataSource()

            oMatrix = oForm.Items.Item("40").Specific
            oMatrix.FlushToDataSource()
            oMatrix.LoadFromDataSource()

            dblDiscount = CDbl(IIf(oDBDataSource.GetValue("U_Discount", 0) = "", 0, oDBDataSource.GetValue("U_Discount", 0)))
            For IntRow As Integer = 0 To oDBDataSourceLines_0.Size - 1
                dblTBD += CDbl(IIf(oDBDataSourceLines_0.GetValue("U_LineTotal", IntRow) = "", 0, oDBDataSourceLines_0.GetValue("U_LineTotal", IntRow)))
                strTaxCode = oDBDataSourceLines_0.GetValue("U_TaxCode", IntRow).Trim
                strQuery = " Select Rate From OVTG Where Code = '" & strTaxCode & "'"
                dblTaxRate = oApplication.Utilities.getRecordSetValue(strQuery, "Rate")
                Dim dblLineTotal As Double = CDbl(IIf(oDBDataSourceLines_0.GetValue("U_LineTotal", IntRow) = "", 0, oDBDataSourceLines_0.GetValue("U_LineTotal", IntRow)))
                dblTaxAmount += (dblTaxRate / 100) * ((dblLineTotal) - (dblLineTotal * (dblDiscount / 100)))
                'dblTaxAmount += (dblTaxRate / 100) * ((dblLineTotal))
            Next

            For IntRow As Integer = 0 To oDBDataSourceLines_1.Size - 1
                dblTBD += CDbl(IIf(oDBDataSourceLines_1.GetValue("U_LineTotal", IntRow) = "", 0, oDBDataSourceLines_1.GetValue("U_LineTotal", IntRow)))
                strTaxCode = oDBDataSourceLines_1.GetValue("U_TaxCode", 0).Trim
                strQuery = " Select Rate From OVTG Where Code = '" & strTaxCode & "'"
                dblTaxRate = oApplication.Utilities.getRecordSetValue(strQuery, "Rate")
                Dim dblLineTotal As Double = CDbl(IIf(oDBDataSourceLines_1.GetValue("U_LineTotal", IntRow) = "", 0, oDBDataSourceLines_1.GetValue("U_LineTotal", IntRow)))
                dblTaxAmount += (dblTaxRate / 100) * ((dblLineTotal) - (dblLineTotal * (dblDiscount / 100)))
                'dblTaxAmount += (dblTaxRate / 100) * ((dblLineTotal))
            Next

            'oMatrix.LoadFromDataSource()

            oDBDataSource.SetValue("U_TBDisc", 0, dblTBD)
            dblDisAmt = dblTBD * (dblDiscount / 100)
            oDBDataSource.SetValue("U_DisAmount", 0, dblDisAmt)
            oDBDataSource.SetValue("U_TaxAmount", 0, dblTaxAmount)
            dblDocTotal = (dblTBD - (dblTBD * (dblDiscount / 100)) + dblTaxAmount)
            oDBDataSource.SetValue("U_DocTotal", 0, dblDocTotal)

            oForm.Update()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub loadCombo(ByVal oForm As SAPbouiCOM.Form)
        Try

            oMatrix = oForm.Items.Item("3").Specific
            oCombo = oMatrix.Columns.Item("V_5_").Cells.Item(oMatrix.RowCount).Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select Code,Name From OVTG Where Category = 'O' And Inactive = 'N' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("Code").Value, oRecordSet.Fields.Item("Name").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            oMatrix = oForm.Items.Item("40").Specific
            oCombo = oMatrix.Columns.Item("V_5_").Cells.Item(oMatrix.RowCount).Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select Code,Name From OVTG Where Category = 'O' And Inactive = 'N' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("Code").Value, oRecordSet.Fields.Item("Name").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            oMatrix = oForm.Items.Item("3").Specific
            oCombo = oMatrix.Columns.Item("V_12").Cells.Item(oMatrix.RowCount).Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select CurrCode,CurrName from OCRN  "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("CurrCode").Value, oRecordSet.Fields.Item("CurrName").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            oMatrix = oForm.Items.Item("40").Specific
            oCombo = oMatrix.Columns.Item("V_12").Cells.Item(oMatrix.RowCount).Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select CurrCode,CurrName from OCRN  "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("CurrCode").Value, oRecordSet.Fields.Item("CurrName").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            oCombo = oForm.Items.Item("45").Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select CurrCode,CurrName from OCRN "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("CurrCode").Value, oRecordSet.Fields.Item("CurrName").Value)
                    oRecordSet.MoveNext()
                End While
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub addDayRowsDynamically(ByVal oForm As SAPbouiCOM.Form)
        Try
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPM")
                oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM6")
                oMatrix = oForm.Items.Item("3").Specific
                oMatrix.FlushToDataSource()
                oMatrix.LoadFromDataSource()
                Dim strFromDate As String
                Dim strToDate As String
                Dim intNoofDays, intFreeDays As Integer
                Dim strProgram As String = oDBDataSource.GetValue("U_PrgCode", 0)
                intNoofDays = CInt(IIf(oDBDataSource.GetValue("U_NoOfDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_NoOfDays", 0)))
                intFreeDays = CInt(IIf(oDBDataSource.GetValue("U_FreeDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_FreeDays", 0)))
                strQuery = "Select VatGourpSa From OITM  Where ItemCode = '" & strProgram & "'"
                Dim strTax As String = oApplication.Utilities.getRecordSetValueString(strQuery, "VatGourpSa")

                strFromDate = oDBDataSource.GetValue("U_PFromDate", 0).ToString()
                Dim strCardCode As String = CType(oForm.Items.Item("6").Specific, SAPbouiCOM.EditText).Value
                Dim strProgFdt As String = CType(oForm.Items.Item("12").Specific, SAPbouiCOM.EditText).Value
                Dim strMProgFdt As String = String.Empty
                Dim intMNoofDays, intMFreeDays As Integer
                For intRow As Integer = 0 To oDBDataSourceLines_0.Size - 1
                    If intRow = 0 Then
                        strMProgFdt = oDBDataSourceLines_0.GetValue("U_Fdate", intRow).ToString()
                    End If
                    If intRow >= 0 Then
                        If oDBDataSourceLines_0.GetValue("U_PaidType", intRow).ToString() = "P" Then
                            intMNoofDays += CInt(IIf(oDBDataSourceLines_0.GetValue("U_NoofDays", intRow).ToString() = "", _
                                                              0, oDBDataSourceLines_0.GetValue("U_NoofDays", intRow).ToString()))
                        ElseIf oDBDataSourceLines_0.GetValue("U_PaidType", intRow).ToString() = "F" Then
                            intMFreeDays += CInt(IIf(oDBDataSourceLines_0.GetValue("U_NoofDays", intRow).ToString() = "", _
                                                              0, oDBDataSourceLines_0.GetValue("U_NoofDays", intRow).ToString()))
                        End If
                    End If
                Next

                'intNoofDays <> intMNoofDays Or intFreeDays <> intMFreeDays
                If intNoofDays <> intMNoofDays Or intFreeDays <> intMFreeDays Or strFromDate <> strMProgFdt Then
                    oMatrix.Clear()
                    oMatrix.FlushToDataSource()
                    If intNoofDays > 0 Then
                        oMatrix.AddRow(1, -1)
                        oMatrix.FlushToDataSource()
                        oDBDataSourceLines_0.SetValue("LineId", oMatrix.RowCount - 1, oMatrix.RowCount.ToString())
                        oDBDataSourceLines_0.SetValue("U_Fdate", oMatrix.RowCount - 1, strFromDate)
                        oDBDataSourceLines_0.SetValue("U_NoofDays", oMatrix.RowCount - 1, intNoofDays)
                        Dim strPrgToDate As String = oApplication.Utilities.getProgramToDate(oForm, strCardCode, strProgFdt, intNoofDays.ToString)
                        oDBDataSourceLines_0.SetValue("U_Edate", oMatrix.RowCount - 1, strPrgToDate)
                        oDBDataSourceLines_0.SetValue("U_PaidType", oMatrix.RowCount - 1, "P")
                        oDBDataSourceLines_0.SetValue("U_IsIReq", oMatrix.RowCount - 1, "Y")
                        oDBDataSourceLines_0.SetValue("U_TaxCode", oMatrix.RowCount - 1, strTax)
                        oMatrix.LoadFromDataSource()
                        calculatePriceAfterDis(oForm, oMatrix.RowCount)
                        oMatrix.LoadFromDataSource()
                    End If
                    If intFreeDays > 0 Then
                        Dim strPRFromDate As String
                        If oMatrix.RowCount > 0 Then
                            strPRFromDate = oDBDataSourceLines_0.GetValue("U_Edate", oMatrix.RowCount - 1).ToString()
                            oMatrix.AddRow(1, oMatrix.RowCount)
                            oMatrix.FlushToDataSource()
                            oDBDataSourceLines_0.SetValue("LineId", oMatrix.RowCount - 1, oMatrix.RowCount.ToString())
                            Dim dtFromDate As DateTime = strPRFromDate.Substring(0, 4) + "-" + strPRFromDate.Substring(4, 2) + "-" + strPRFromDate.Substring(6, 2)
                            oDBDataSourceLines_0.SetValue("U_Fdate", oMatrix.RowCount - 1, dtFromDate.AddDays(1).ToString("yyyyMMdd"))
                            oDBDataSourceLines_0.SetValue("U_NoofDays", oMatrix.RowCount - 1, intFreeDays)
                            Dim strPrgToDate As String = oApplication.Utilities.getProgramToDate(oForm, strCardCode, dtFromDate.AddDays(1).ToString("yyyyMMdd"), intFreeDays.ToString)
                            oDBDataSourceLines_0.SetValue("U_Edate", oMatrix.RowCount - 1, strPrgToDate)
                            oDBDataSourceLines_0.SetValue("U_PaidType", oMatrix.RowCount - 1, "F")
                            oDBDataSourceLines_0.SetValue("U_IsIReq", oMatrix.RowCount - 1, "F")
                            oDBDataSourceLines_0.SetValue("U_TaxCode", oMatrix.RowCount - 1, strTax)
                            oMatrix.LoadFromDataSource()
                            calculatePriceAfterDis(oForm, oMatrix.RowCount)
                            oMatrix.LoadFromDataSource()
                        Else
                            oMatrix.AddRow(1, -1)
                            oMatrix.FlushToDataSource()
                            strPRFromDate = oDBDataSource.GetValue("U_PFromDate", 0).ToString()
                            Dim dtFromDate As DateTime = strPRFromDate.Substring(0, 4) + "-" + strPRFromDate.Substring(4, 2) + "-" + strPRFromDate.Substring(6, 2)
                            oDBDataSourceLines_0.SetValue("LineId", oMatrix.RowCount - 1, oMatrix.RowCount.ToString())
                            oDBDataSourceLines_0.SetValue("U_Fdate", oMatrix.RowCount - 1, dtFromDate.AddDays(0).ToString("yyyyMMdd"))
                            oDBDataSourceLines_0.SetValue("U_NoofDays", oMatrix.RowCount - 1, intFreeDays)
                            Dim strPrgToDate As String = oApplication.Utilities.getProgramToDate(oForm, strCardCode, dtFromDate.AddDays(0).ToString("yyyyMMdd"), intFreeDays.ToString)
                            oDBDataSourceLines_0.SetValue("U_Edate", oMatrix.RowCount - 1, strPrgToDate)
                            oDBDataSourceLines_0.SetValue("U_PaidType", oMatrix.RowCount - 1, "F")
                            oDBDataSourceLines_0.SetValue("U_IsIReq", oMatrix.RowCount - 1, "Y")
                            oDBDataSourceLines_0.SetValue("U_TaxCode", oMatrix.RowCount - 1, strTax)
                            oMatrix.LoadFromDataSource()
                            calculatePriceAfterDis(oForm, oMatrix.RowCount)
                            oMatrix.LoadFromDataSource()
                        End If
                    End If
                Else
                    oMatrix.LoadFromDataSource()
                    oMatrix.FlushToDataSource()
                End If

                oMatrix = oForm.Items.Item("40").Specific
                For intRow As Integer = 0 To oDBDataSourceLines_1.Size - 1
                    If oDBDataSourceLines_1.GetValue("U_InvCreated", intRow) = "" Or oDBDataSourceLines_1.GetValue("U_InvCreated", intRow).Trim() = "N" Then
                        oDBDataSourceLines_1.SetValue("U_Quantity", intRow, oDBDataSource.GetValue("U_NoofDays", 0))
                        oDBDataSourceLines_1.SetValue("U_Date", intRow, oDBDataSource.GetValue("U_PToDate", 0))
                    End If
                Next
                oMatrix.LoadFromDataSource()
                oMatrix.FlushToDataSource()

            ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                oDBDataSourceLines_0 = oForm.DataSources.DBDataSources.Item("@Z_CPM6")
                oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@Z_CPM7")

                oMatrix = oForm.Items.Item("3").Specific
                oMatrix.FlushToDataSource()
                Dim intNoofDays, intFreeDays, intDelDays As Integer
                Dim strFromDate As String = ""
                Dim strToDate As String = ""
                intNoofDays = CInt(IIf(oDBDataSource.GetValue("U_NoOfDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_NoOfDays", 0)))
                intFreeDays = CInt(IIf(oDBDataSource.GetValue("U_FreeDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_FreeDays", 0)))
                intDelDays = CInt(IIf(oDBDataSource.GetValue("U_DelDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_DelDays", 0)))
                Dim intMNoofDays, intMFreeDays As Integer
                strFromDate = oDBDataSource.GetValue("U_PFromDate", 0).ToString()

                For intRow As Integer = 0 To oDBDataSourceLines_0.Size - 1

                    If intRow = 0 Then
                        strFromDate = oDBDataSourceLines_0.GetValue("U_Fdate", intRow).ToString()
                    End If
                    strToDate = oDBDataSourceLines_0.GetValue("U_Edate", intRow).ToString()

                    If intRow >= 0 Then
                        If oDBDataSourceLines_0.GetValue("U_PaidType", intRow).ToString() = "P" Then
                            intMNoofDays += CInt(IIf(oDBDataSourceLines_0.GetValue("U_NoofDays", intRow).ToString() = "", _
                                                              0, oDBDataSourceLines_0.GetValue("U_NoofDays", intRow).ToString()))
                        ElseIf oDBDataSourceLines_0.GetValue("U_PaidType", intRow).ToString() = "F" Then
                            intMFreeDays += CInt(IIf(oDBDataSourceLines_0.GetValue("U_NoofDays", intRow).ToString() = "", _
                                                              0, oDBDataSourceLines_0.GetValue("U_NoofDays", intRow).ToString()))
                        End If
                    End If

                    If intRow > 0 Then
                        Dim strTDate As String = oDBDataSourceLines_0.GetValue("U_Edate", intRow - 1).ToString()
                        Dim strFDate As String = oDBDataSourceLines_0.GetValue("U_Fdate", intRow).ToString()
                        Dim strTotalDays As String = oDBDataSourceLines_0.GetValue("U_NoOfDays", intRow).ToString()

                        Dim dtToDate As DateTime = CDate(strTDate.Substring(0, 4) + "-" + strTDate.Substring(4, 2) + "-" + strTDate.Substring(6, 2))
                        If strFDate <> dtToDate.AddDays(1).ToString("yyyyMMdd") And strFDate <> "" Then
                            'oDBDataSourceLines_0.SetValue("U_Fdate", intRow, dtToDate.AddDays(1).ToString("yyyyMMdd"))
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", intRow + 1, dtToDate.AddDays(1).ToString("yyyyMMdd"))
                            Dim strPrgToDate As String = oApplication.Utilities.getProgramToDate(oForm, oDBDataSource.GetValue("U_CardCode", 0).Trim(), dtToDate.AddDays(1).ToString("yyyyMMdd"), IIf(strTotalDays = "", "1", strTotalDays).ToString)
                            'oDBDataSourceLines_0.SetValue("U_Edate", intRow, strPrgToDate)
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", intRow + 1, strPrgToDate)
                            calculatePriceAfterDis(oForm, intRow + 1)
                            oMatrix.FlushToDataSource()
                        End If
                    Else
                        Dim strTotalDays As String = oDBDataSourceLines_0.GetValue("U_NoOfDays", intRow).ToString()
                        Dim strDelDays As String = oDBDataSourceLines_0.GetValue("U_DelDays", intRow).ToString()
                        Dim strFDate As String = oDBDataSourceLines_0.GetValue("U_Fdate", intRow).ToString()
                        Dim strPGRef As String = oDBDataSourceLines_0.GetValue("DocEntry", intRow).ToString()
                        If CInt(IIf(strDelDays = "", 0, strDelDays)) > 0 Then
                            strTotalDays = (CInt(IIf(strTotalDays = "", 0, strTotalDays)) - CInt(IIf(strDelDays = "", 0, strDelDays))).ToString
                            If CInt(strTotalDays) > 0 Then

                                strQuery = " Select Convert(Varchar(8),Max(T0.U_DelDate+1),112) As 'FD' From DLN1 T0 "
                                strQuery += " JOIN [@Z_CPM6] T1 On T0.U_ProgramID = T1.DocEntry "
                                strQuery += " And T0.U_DelDate BetWeen T1.U_Fdate And T1.U_Edate "
                                strQuery += " And (T0.LineStatus = 'O' Or (T0.LineStatus = 'C' And T0.TargetType = '-1')) "
                                strQuery += " JOIN ODLN T2 On T0.DocEntry = T2.DocEntry And T2.CANCELED = 'N' "
                                strQuery += " Where T0.U_ProgramID = '" & strPGRef & "' "
                                strQuery += " And T1.LineId = '" & oDBDataSourceLines_0.GetValue("LineId", intRow).ToString() & "'"
                                Dim strPFromDt As String = oApplication.Utilities.getRecordSetValueString(strQuery, "FD")

                                If strPFromDt <> "" Then
                                    strFDate = strPFromDt
                                End If

                                Dim strPrgToDate As String = oApplication.Utilities.getProgramToDate(oForm, oDBDataSource.GetValue("U_CardCode", 0).Trim(), _
                                                                   strFDate, IIf(strTotalDays = "", "1", strTotalDays).ToString)
                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", intRow + 1, strPrgToDate)
                                calculatePriceAfterDis(oForm, intRow + 1)
                                oMatrix.FlushToDataSource()

                            End If
                        End If
                    End If


                    Dim strTotalDays1 As String = oDBDataSourceLines_0.GetValue("U_NoOfDays", intRow).ToString()
                    If intRow = 0 Then
                        If CInt(IIf(strTotalDays1 = "", 0, strTotalDays1)) > 0 Then
                            strToDate = oDBDataSourceLines_0.GetValue("U_Edate", intRow).ToString()
                        End If
                    Else
                        If CInt(IIf(strTotalDays1 = "", 0, strTotalDays1)) = 0 Then
                            strToDate = oDBDataSourceLines_0.GetValue("U_Edate", intRow - 1).ToString()
                        Else
                            strToDate = oDBDataSourceLines_0.GetValue("U_Edate", intRow).ToString()
                        End If
                    End If
                Next

                oMatrix.LoadFromDataSource()


                If intNoofDays <> intMNoofDays Or intFreeDays <> intMFreeDays Then
                    Dim strPRFromDate As String = oDBDataSource.GetValue("U_PFromDate", 0).ToString()
                    Dim strCardCode As String = CType(oForm.Items.Item("6").Specific, SAPbouiCOM.EditText).Value
                    Dim strProgFdt As String = CType(oForm.Items.Item("12").Specific, SAPbouiCOM.EditText).Value
                    Dim dtFromDate As DateTime = strPRFromDate.Substring(0, 4) + "-" + strPRFromDate.Substring(4, 2) + "-" + strPRFromDate.Substring(6, 2)

                    Dim intDelDays1 As Integer = CInt(IIf(oDBDataSource.GetValue("U_DelDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_DelDays", 0).ToString))

                    oDBDataSource.SetValue("U_NoOfDays", 0, intMNoofDays)
                    oDBDataSource.SetValue("U_FreeDays", 0, intMFreeDays)
                    If intDelDays1 > 0 Then
                        strQuery = " Select Convert(Varchar(8),Max(T0.U_DelDate+1),112) As 'FD' From DLN1 T0 "
                        strQuery += " JOIN ODLN T2 On T0.DocEntry = T2.DocEntry And T2.CANCELED = 'N' "
                        strQuery += " Where (T0.LineStatus = 'O' Or (T0.LineStatus = 'C' And T0.TargetType = '-1')) "
                        strQuery += " And T0.U_ProgramID = '" & oDBDataSource.GetValue("DocEntry", 0).ToString() & "' "
                        strPRFromDate = oApplication.Utilities.getRecordSetValueString(strQuery, "FD")
                        If strPRFromDate <> "" Then
                            dtFromDate = CDate(strPRFromDate.Substring(0, 4) + "-" + strPRFromDate.Substring(4, 2) + "-" + strPRFromDate.Substring(6, 2))
                        End If
                    End If
                    Dim strPrgToDate As String = oApplication.Utilities.getProgramToDate(oForm, strCardCode, dtFromDate.AddDays(0).ToString("yyyyMMdd"), _
                                                                                         ((intMNoofDays + intMFreeDays) - intDelDays1).ToString)
                    oDBDataSource.SetValue("U_PToDate", 0, (strPrgToDate).ToString)
                    oDBDataSource.SetValue("U_RemDays", 0, ((intMNoofDays + intMFreeDays) - (intDelDays)).ToString)
                    'Dim intRemDays As Integer = ((intMNoofDays + intMFreeDays) - (intDelDays))
                    'If oDBDataSource.GetValue("U_PaidSta", 0).Trim() = "P" And intRemDays > 0 Then
                    '    oDBDataSource.SetValue("U_PaidSta", 0, "O")
                    'End If
                End If

                If oDBDataSource.GetValue("U_PFromDate", 0).ToString() <> strFromDate Then
                    oDBDataSource.SetValue("U_PFromDate", 0, strFromDate)
                    Dim strTotalDays As String = (intNoofDays + intFreeDays).ToString
                    Dim strPRFromDate As String = oDBDataSource.GetValue("U_PFromDate", 0).ToString()
                    Dim dtFromDate As DateTime = strPRFromDate.Substring(0, 4) + "-" + strPRFromDate.Substring(4, 2) + "-" + strPRFromDate.Substring(6, 2)
                    Dim intDelDays1 As Integer = CInt(IIf(oDBDataSource.GetValue("U_DelDays", 0).ToString() = "", 0, oDBDataSource.GetValue("U_DelDays", 0).ToString))
                    If intDelDays1 > 0 Then
                        strQuery = " Select Convert(Varchar(8),Max(T0.U_DelDate+1),112) As 'FD' From DLN1 T0 "
                        strQuery += " JOIN ODLN T2 On T0.DocEntry = T2.DocEntry And T2.CANCELED = 'N' "
                        strQuery += " Where (T0.LineStatus = 'O' Or (T0.LineStatus = 'C' And T0.TargetType = '-1')) "
                        strQuery += " And T0.U_ProgramID = '" & oDBDataSource.GetValue("DocEntry", 0).ToString() & "' "
                        strPRFromDate = oApplication.Utilities.getRecordSetValueString(strQuery, "FD")
                        If strPRFromDate <> "" Then
                            dtFromDate = CDate(strPRFromDate.Substring(0, 4) + "-" + strPRFromDate.Substring(4, 2) + "-" + strPRFromDate.Substring(6, 2))
                        End If
                    End If
                    Dim strPrgToDate As String = oApplication.Utilities.getProgramToDate(oForm, CType(oForm.Items.Item("6").Specific, SAPbouiCOM.EditText).Value, strFromDate, strTotalDays)
                    oDBDataSource.SetValue("U_PToDate", 0, (strPrgToDate).ToString)
                End If

                If oDBDataSource.GetValue("U_PToDate", 0).ToString() <> strToDate Then
                    If strToDate <> "" Then
                        oDBDataSource.SetValue("U_PToDate", 0, strToDate)
                    End If
                End If

                oMatrix = oForm.Items.Item("40").Specific
                For intRow As Integer = 0 To oDBDataSourceLines_1.Size - 1
                    If oDBDataSourceLines_1.GetValue("U_InvCreated", intRow) = "" Or oDBDataSourceLines_1.GetValue("U_InvCreated", intRow).Trim() = "N" Then
                        oDBDataSourceLines_1.SetValue("U_Quantity", intRow, oDBDataSource.GetValue("U_NoofDays", 0))
                        oDBDataSourceLines_1.SetValue("U_Date", intRow, oDBDataSource.GetValue("U_PToDate", 0))
                    End If
                Next
                oMatrix.LoadFromDataSource()
                oMatrix.FlushToDataSource()

            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub DefaultCurrency(ByVal oForm As SAPbouiCOM.Form, strCurr As String)
        Try
            Dim blnCheck As Boolean = False
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPM")
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim strLC As String
            Dim strSC As String
            Dim strDocCurr As String
            oRecordSet = DirectCast(oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            Dim strQry As String = "Select MainCurncy,SysCurrncy FROM OADM"
            oRecordSet.DoQuery(strQry)
            If Not oRecordSet.EoF Then
                strLC = oRecordSet.Fields.Item("MainCurncy").Value.ToString()
                strSC = oRecordSet.Fields.Item("SysCurrncy").Value.ToString()
                If oRecordSet.Fields.Item("MainCurncy").Value.ToString() = strCurr.ToString() Then
                    If DirectCast(oForm.Items.Item("_45").Specific, SAPbouiCOM.ComboBox).Selected.Value <> "L" Then
                        strDocCurr = oRecordSet.Fields.Item("MainCurncy").Value.ToString()
                        oDBDataSource.SetValue("U_CurSour", oDBDataSource.Offset, "L")
                        oDBDataSource.SetValue("U_DocCur", oDBDataSource.Offset, strDocCurr)
                        oDBDataSource.SetValue("U_DocRate", oDBDataSource.Offset, GetCurrencyRate(oForm, strCurr).ToString())
                        blnCheck = True
                        oForm.Items.Item("45").Visible = False
                        oForm.Items.Item("45_").Visible = False
                    Else
                        oDBDataSource.SetValue("U_DocCur", oDBDataSource.Offset, strCurr)
                        blnCheck = True
                        Dim dblRate As Double = GetCurrencyRate(oForm, strCurr).ToString()
                        oDBDataSource.SetValue("U_DocRate", oDBDataSource.Offset, dblRate)
                        oForm.Items.Item("45").Visible = False
                        oForm.Items.Item("45_").Visible = False
                    End If
                End If
                If oRecordSet.Fields.Item("SysCurrncy").Value.ToString() = strCurr.ToString() AndAlso blnCheck <> True Then
                    If DirectCast(oForm.Items.Item("_45").Specific, SAPbouiCOM.ComboBox).Selected.Value <> "S" Then
                        oDBDataSource.SetValue("U_CurSour", oDBDataSource.Offset, "S")
                        strDocCurr = oRecordSet.Fields.Item("SysCurrncy").Value.ToString()
                        oDBDataSource.SetValue("U_DocCur", oDBDataSource.Offset, strDocCurr)
                        oDBDataSource.SetValue("U_DocRate", oDBDataSource.Offset, GetCurrencyRate(oForm, strCurr).ToString())
                        blnCheck = True
                        oForm.Items.Item("45").Visible = False
                        oForm.Items.Item("45_").Visible = False
                    End If
                End If
                If strCurr.ToString() = "##" Then
                    oDBDataSource.SetValue("U_CurSour", oDBDataSource.Offset, "C")
                    oDBDataSource.SetValue("U_DocCur", oDBDataSource.Offset, strSC)
                    strDocCurr = "##"
                    blnCheck = True
                    oForm.Items.Item("45").Visible = True
                    oDBDataSource.SetValue("U_DocRate", oDBDataSource.Offset, GetCurrencyRate(oForm, strSC).ToString())
                End If
                If blnCheck <> True Then
                    oDBDataSource.SetValue("U_CurSour", oDBDataSource.Offset, "C")
                    oDBDataSource.SetValue("U_DocCur", oDBDataSource.Offset, strCurr)
                    oDBDataSource.SetValue("U_DocRate", oDBDataSource.Offset, GetCurrencyRate(oForm, strCurr).ToString())
                    oForm.Items.Item("45").Visible = True
                    oForm.Items.Item("45_").Enabled = False
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function GetCurrencyRate(ByVal oForm As SAPbouiCOM.Form, strCurr As String) As Double
        Dim dblRate As Double = 1
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = DirectCast(oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            Dim strQry As String = " Select RATE FROM ORTT Where Convert(VarChar,RateDate,112) "
            strQry += " = '" & System.DateTime.Now.ToString("yyyyMMdd") & "'"
            strQry += " AND Currency='" & strCurr & "'"
            oRecordSet.DoQuery(strQry)
            If Not oRecordSet.EoF Then
                dblRate = Convert.ToDouble(oRecordSet.Fields.Item("Rate").Value.ToString())
                Return dblRate
            Else
                Return dblRate
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return dblRate
    End Function

    Private Sub getPrice(ByVal strDCurrency As String, strICurrency As String, ByVal dblBasePrice As String, ByRef dblPrice As Double)
        Try
            Dim oExRecordSet As SAPbobsCOM.Recordset
            Dim dblRExRate, dblAExRate As Double

            oExRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim dblLocalCurrency As String = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency

            If strDCurrency <> dblLocalCurrency Then
                strQuery = "Select Rate From ORTT Where Currency = '" + strDCurrency + "' And Convert(VarChar(8),RateDate,112) = Convert(VarChar(8),GetDate(),112)"
                oExRecordSet.DoQuery(strQuery)
                If Not oExRecordSet.EoF Then
                    dblRExRate = oExRecordSet.Fields.Item("Rate").Value
                    If strICurrency = dblLocalCurrency Then
                        dblPrice = (dblBasePrice / dblRExRate)
                    Else
                        strQuery = "Select isnull(Rate,1) 'Rate' From ORTT Where Currency = '" + strICurrency + "' And Convert(VarChar(8),RateDate,112) = Convert(VarChar(8),GetDate(),112)"
                        oExRecordSet.DoQuery(strQuery)
                        If Not oExRecordSet.EoF Then
                            dblAExRate = oExRecordSet.Fields.Item("Rate").Value
                            dblPrice = ((dblBasePrice * dblAExRate) / dblRExRate)
                        End If
                    End If
                End If
            ElseIf strDCurrency = dblLocalCurrency Then
                strQuery = "Select Rate From ORTT Where Currency = '" + strICurrency + "' And Convert(VarChar(8),RateDate,112) = Convert(VarChar(8),GetDate(),112)"
                oExRecordSet.DoQuery(strQuery)
                If Not oExRecordSet.EoF Then
                    dblAExRate = oExRecordSet.Fields.Item("Rate").Value
                    dblPrice = (dblBasePrice * dblAExRate)
                End If
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class

'Public Class clsInvoiceServiceItem
'    Inherits clsBase

'    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
'    Private oMatrix As SAPbouiCOM.Matrix
'    Private oEditText As SAPbouiCOM.EditText
'    Private InvForConsumedItems, count As Integer
'    Dim oDBDataSource As SAPbouiCOM.DBDataSource
'    Dim oDBDataSourceLines As SAPbouiCOM.DBDataSource
'    Public intSelectedMatrixrow As Integer = 0
'    Private RowtoDelete As Integer
'    Private MatrixId As String
'    Private oRecordSet As SAPbobsCOM.Recordset
'    Private dtValidFrom, dtValidTo As Date
'    Private strQuery As String

'#Region "Initialization"

'    Public Sub New()
'        MyBase.New()
'        InvForConsumedItems = 0
'    End Sub

'    Public Sub LoadForm()
'        Try
'            oForm = oApplication.Utilities.LoadForm(xml_Z_OISI, frm_Z_OISI)
'            oForm = oApplication.SBO_Application.Forms.ActiveForm()
'            oForm.Freeze(True)
'            initialize(oForm)
'            oForm.EnableMenu(mnu_ADD_ROW, True)
'            oForm.EnableMenu(mnu_DELETE_ROW, True)
'            oForm.Freeze(False)
'        Catch ex As Exception
'            Throw ex
'        Finally
'            oForm.Freeze(False)
'        End Try
'    End Sub

'    Public Sub LoadForm(ByVal strRef As String, ByVal strCardCode As String)
'        Try
'            oForm = oApplication.Utilities.LoadForm(xml_Z_OISI, frm_Z_OISI)
'            oForm = oApplication.SBO_Application.Forms.ActiveForm()
'            oForm.Freeze(True)
'            initialize(oForm)
'            oForm.EnableMenu(mnu_ADD_ROW, True)
'            oForm.EnableMenu(mnu_DELETE_ROW, True)
'            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
'            oForm.Items.Item("6").Specific.value = strRef
'            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
'            oForm.Items.Item("11").Specific.value = strCardCode
'            oForm.Freeze(False)
'        Catch ex As Exception
'            Throw ex
'        Finally
'            oForm.Freeze(False)
'        End Try
'    End Sub

'#End Region

'#Region "Item Event"
'    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
'        Try
'            If pVal.FormTypeEx = frm_Z_OISI Then
'                Select Case pVal.BeforeAction
'                    Case True
'                        Select Case pVal.EventType
'                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
'                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
'                                Select Case pVal.ItemUID
'                                    Case "1"
'                                        If (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
'                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
'                                                If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
'                                                    BubbleEvent = False
'                                                    Exit Sub
'                                                Else
'                                                    If validation(oForm) = False Then
'                                                        BubbleEvent = False
'                                                        Exit Sub
'                                                    End If
'                                                End If
'                                            End If
'                                        End If
'                                End Select
'                        End Select
'                    Case False
'                        Select Case pVal.EventType
'                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
'                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
'                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
'                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
'                                Select Case pVal.ItemUID
'                                    Case "14"
'                                        oApplication.SBO_Application.ActivateMenuItem(mnu_ADD_ROW)
'                                    Case "15"
'                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
'                                End Select
'                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
'                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
'                                    If pVal.ItemUID = "3" And (pVal.ColUID = "V_2" Or pVal.ColUID = "V_3" Or pVal.ColUID = "V_4") And pVal.Row > 0 Then
'                                        oForm.Freeze(True)
'                                        calculatePriceAfterDis(oForm, pVal.Row)
'                                        oForm.Freeze(False)
'                                    End If
'                                End If
'                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
'                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
'                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OISI")
'                                oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_ISI1")
'                                oMatrix = oForm.Items.Item("3").Specific
'                                oMatrix.FlushToDataSource()
'                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
'                                Dim oDataTable As SAPbouiCOM.DataTable
'                                Try
'                                    oCFLEvento = pVal
'                                    oDataTable = oCFLEvento.SelectedObjects

'                                    If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
'                                        If (pVal.ItemUID = "3" And (pVal.ColUID = "V_0" Or pVal.ColUID = "V_1")) Then
'                                            oMatrix = oForm.Items.Item("3").Specific
'                                            oMatrix.LoadFromDataSource()
'                                            Dim intAddRows As Integer = oDataTable.Rows.Count
'                                            If intAddRows > 1 Then
'                                                intAddRows -= 1
'                                                oMatrix.AddRow(intAddRows, pVal.Row - 1)
'                                            End If
'                                            oMatrix.FlushToDataSource()
'                                            For index As Integer = 0 To oDataTable.Rows.Count - 1
'                                                oDBDataSourceLines.SetValue("LineId", pVal.Row + index - 1, (pVal.Row + index).ToString())
'                                                oDBDataSourceLines.SetValue("U_ItemCode", pVal.Row + index - 1, oDataTable.GetValue("ItemCode", index))
'                                                oDBDataSourceLines.SetValue("U_ItemName", pVal.Row + index - 1, oDataTable.GetValue("ItemName", index))
'                                                oDBDataSourceLines.SetValue("U_Quantity", pVal.Row + index - 1, "1")
'                                                Dim dblItemPrice As Double = oApplication.Utilities.GetCustItemPrice(oApplication.Utilities.getEditTextvalue(oForm, "11"), oDataTable.GetValue("ItemCode", index), System.DateTime.Now.Date)
'                                                oDBDataSourceLines.SetValue("U_Price", pVal.Row + index - 1, dblItemPrice)
'                                                oDBDataSourceLines.SetValue("U_LineTotal", pVal.Row + index - 1, dblItemPrice)
'                                            Next
'                                            oMatrix.LoadFromDataSource()
'                                            oMatrix.FlushToDataSource()
'                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
'                                        ElseIf (pVal.ItemUID = "3" And (pVal.ColUID = "V_4")) Then
'                                            oMatrix = oForm.Items.Item("3").Specific
'                                            oMatrix.LoadFromDataSource()
'                                            Dim intAddRows As Integer = oDataTable.Rows.Count
'                                            If intAddRows > 1 Then
'                                                intAddRows -= 1
'                                                oMatrix.AddRow(intAddRows, pVal.Row - 1)
'                                            End If
'                                            oMatrix.FlushToDataSource()
'                                            oMatrix.LoadFromDataSource()
'                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
'                                        End If
'                                    End If
'                                Catch ex As Exception

'                                End Try
'                        End Select
'                End Select
'            End If
'        Catch ex As Exception
'            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
'        End Try
'    End Sub
'#End Region

'#Region "Menu Event"
'    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
'        Try
'            oForm = oApplication.SBO_Application.Forms.ActiveForm()
'            If oForm.TypeEx = frm_Z_OISI Then
'                Select Case pVal.MenuUID
'                    Case mnu_ADD_ROW
'                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
'                        If pVal.BeforeAction = False Then
'                            AddRow(oForm)
'                        End If
'                    Case mnu_DELETE_ROW
'                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
'                        If pVal.BeforeAction = False Then
'                            RefereshDeleteRow(oForm)
'                        End If
'                End Select
'            End If
'        Catch ex As Exception
'            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
'            oForm.Freeze(False)
'        End Try
'    End Sub
'#End Region

'#Region "Data Events"

'    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
'        Try
'            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
'            If oForm.TypeEx = frm_Z_OISI Then
'                Select Case BusinessObjectInfo.BeforeAction
'                    Case True

'                    Case False
'                        Select Case BusinessObjectInfo.EventType
'                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
'                                oMatrix = oForm.Items.Item("3").Specific
'                                For index As Integer = 1 To oMatrix.VisualRowCount
'                                    If CType(oMatrix.Columns.Item("V_8").Cells.Item(index).Specific, SAPbouiCOM.ComboBox).Value = "Y" Then
'                                        oMatrix.CommonSetting.SetRowEditable(index, False)
'                                    Else
'                                        oMatrix.CommonSetting.SetRowEditable(index, True)
'                                    End If
'                                Next
'                        End Select
'                End Select
'            End If
'        Catch ex As Exception
'            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
'        End Try
'    End Sub

'#End Region

'#Region "Function"

'    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
'        Try
'            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OISI")
'            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_ISI1")

'            oMatrix = oForm.Items.Item("3").Specific
'            oMatrix.LoadFromDataSource()
'            oMatrix.AddRow(1, -1)
'            oMatrix.FlushToDataSource()

'            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
'            oRecordSet.DoQuery("Select IsNull(MAX(DocEntry),1) From [@Z_OISI]")
'            If Not oRecordSet.EoF Then
'                oDBDataSource.SetValue("DocNum", 0, oRecordSet.Fields.Item(0).Value.ToString())
'            End If

'            MatrixId = "3"
'        Catch ex As Exception
'            Throw ex
'        End Try
'    End Sub

'    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
'        Try
'            aForm.Freeze(True)
'            oMatrix = aForm.Items.Item("3").Specific
'            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_ISI1")
'            oMatrix.FlushToDataSource()
'            For count = 1 To oDBDataSourceLines.Size
'                oDBDataSourceLines.SetValue("LineId", count - 1, count)
'            Next
'            oMatrix.LoadFromDataSource()
'            aForm.Freeze(False)
'        Catch ex As Exception
'            aForm.Freeze(False)
'            Throw ex
'        End Try
'    End Sub

'    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
'        Try
'            aForm.Freeze(True)
'            Select Case aForm.PaneLevel
'                Case "0"
'                    oMatrix = aForm.Items.Item("3").Specific
'                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_ISI1")
'                    If oMatrix.RowCount <= 0 Then
'                        oMatrix.AddRow()
'                    Else
'                        oMatrix.AddRow(1, oMatrix.RowCount + 1)
'                        oMatrix.ClearRowData(oMatrix.RowCount)
'                        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
'                            aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
'                        End If
'                    End If
'                    oMatrix.FlushToDataSource()
'                    For count = 1 To oDBDataSourceLines.Size
'                        oDBDataSourceLines.SetValue("LineId", count - 1, count)
'                    Next
'                    oMatrix.LoadFromDataSource()
'                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
'                    AssignLineNo(aForm)
'            End Select
'            aForm.Freeze(False)
'        Catch ex As Exception
'            aForm.Freeze(False)
'            Throw ex
'        End Try
'    End Sub

'    Private Sub deleterow(ByVal aForm As SAPbouiCOM.Form)
'        Try
'            Select Case aForm.PaneLevel
'                Case "0"
'                    oMatrix = aForm.Items.Item("11").Specific
'                    oDBDataSourceLines = aForm.DataSources.DBDataSources.Item("@Z_ISI1")
'            End Select
'            oMatrix.FlushToDataSource()
'            For introw As Integer = 1 To oMatrix.RowCount
'                If oMatrix.IsRowSelected(introw) Then
'                    oMatrix.DeleteRow(introw)
'                    oDBDataSourceLines.RemoveRecord(introw - 1)
'                    oMatrix.FlushToDataSource()
'                    For count As Integer = 1 To oDBDataSourceLines.Size
'                        oDBDataSourceLines.SetValue("LineId", count - 1, count)
'                    Next
'                    Select Case aForm.PaneLevel
'                        Case "0"
'                            oMatrix = aForm.Items.Item("3").Specific
'                            oDBDataSourceLines = aForm.DataSources.DBDataSources.Item("@Z_ISI1")
'                            AssignLineNo(aForm)
'                    End Select
'                    oMatrix.LoadFromDataSource()
'                    If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
'                        aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
'                    End If
'                    Exit Sub
'                End If
'            Next
'        Catch ex As Exception
'            aForm.Freeze(False)
'            Throw ex
'        End Try
'    End Sub

'    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
'        Try

'            oMatrix = aForm.Items.Item("3").Specific
'            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OISI")
'            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_ISI1")

'            If Me.MatrixId = "3" Then
'                oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_ISI1")
'            End If

'            Me.RowtoDelete = intSelectedMatrixrow
'            oDBDataSourceLines.RemoveRecord(Me.RowtoDelete - 1)
'            oMatrix.LoadFromDataSource()
'            oMatrix.FlushToDataSource()

'            For count = 1 To oDBDataSourceLines.Size - 1
'                oDBDataSourceLines.SetValue("LineId", count - 1, count)
'            Next
'            oMatrix.LoadFromDataSource()

'        Catch ex As Exception
'            aForm.Freeze(False)
'            Throw ex
'        End Try
'    End Sub

'    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
'        Try
'            oMatrix = oForm.Items.Item("3").Specific
'            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OISI")
'            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_ISI1")

'            Dim oTest As SAPbobsCOM.Recordset
'            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

'            Return True
'        Catch ex As Exception
'            aForm.Freeze(False)
'            Throw ex
'        End Try
'    End Function

'    Private Sub EnableControls(ByVal oForm As SAPbouiCOM.Form, ByVal blnEnable As Boolean, ByVal strReference As String)
'        Try
'            Dim strMenu As String = CType(oForm.Items.Item("10").Specific, SAPbouiCOM.EditText).Value
'            Dim oRecordSet As SAPbobsCOM.Recordset
'            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
'            Dim strQuery As String = "Select U_InvCreated From [@Z_CPM6] Where U_Reference = '" + strReference + "'"
'            oRecordSet.DoQuery(strQuery)
'            If Not oRecordSet.EoF Then
'                If oRecordSet.Fields.Item(0).Value = "N" Then
'                    oForm.Items.Item("3").Enabled = False
'                    oForm.Items.Item("14").Enabled = False
'                    oForm.Items.Item("15").Enabled = False
'                End If
'            End If
'        Catch ex As Exception
'            Throw ex
'        End Try
'    End Sub

'    Private Sub calculatePriceAfterDis(ByVal oForm As SAPbouiCOM.Form, ByVal intRow As Integer)
'        Try
'            oMatrix = oForm.Items.Item("3").Specific
'            oMatrix.FlushToDataSource()
'            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_ISI1")
'            Dim dblPrice, dblDiscount, dblPriceAfterDis, dblNoofDays As Double

'            Dim strCardCode As String = CType(oForm.Items.Item("11").Specific, SAPbouiCOM.EditText).Value
'            Dim strServiceItem As String = oDBDataSourceLines.GetValue("U_ItemCode", intRow - 1).ToString()
'            dblNoofDays = IIf(oDBDataSourceLines.GetValue("U_Quantity", intRow - 1).ToString() = "", 0, oDBDataSourceLines.GetValue("U_Quantity", intRow - 1))
'            dblPrice = IIf(oDBDataSourceLines.GetValue("U_Price", intRow - 1).ToString() = "", 0, oDBDataSourceLines.GetValue("U_Price", intRow - 1).ToString())

'            If dblPrice = 0 Then
'                Dim dblItemPrice As Double = oApplication.Utilities.GetCustItemPrice(strCardCode, strServiceItem, System.DateTime.Now.Date)
'                oDBDataSourceLines.SetValue("U_Price", intRow - 1, dblItemPrice)
'                oMatrix.LoadFromDataSource()
'                oMatrix.FlushToDataSource()
'            End If

'            dblDiscount = oDBDataSourceLines.GetValue("U_Discount", intRow - 1)
'            dblPriceAfterDis = dblNoofDays * (dblPrice - ((dblPrice * dblDiscount) / 100))
'            oDBDataSourceLines.SetValue("U_LineTotal", intRow - 1, dblPriceAfterDis)
'            oMatrix.LoadFromDataSource()
'        Catch ex As Exception
'            Throw ex
'        End Try
'    End Sub

'#End Region

'End Class
