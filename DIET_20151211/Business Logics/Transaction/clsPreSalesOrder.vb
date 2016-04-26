Imports SAPbobsCOM

Public Class clsPreSalesOrder
    Inherits clsBase

    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private objForm As SAPbouiCOM.Form
    Private oEditText As SAPbouiCOM.EditText
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines As SAPbouiCOM.DBDataSource
    Private oMode As SAPbouiCOM.BoFormMode
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Private oRecordSet As SAPbobsCOM.Recordset
    Private oMatrix As SAPbouiCOM.Matrix
    Private MatrixID As String = String.Empty
    Dim strQuery As String = String.Empty

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Sub LoadForm()
        Try
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_OPSL) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_OPSL, frm_Z_OPSL)
            oForm = oApplication.SBO_Application.Forms.Item(strUID)
            oForm.DataBrowser.BrowseBy = "14"
            initialize(oForm)
            oForm.DataSources.UserDataSources.Add("PrgName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            Dim oEditText As SAPbouiCOM.EditText
            oEditText = oForm.Items.Item("8__").Specific
            oEditText.DataBind.SetBound(True, "", "PrgName")
            addChooseFromListConditions(oForm)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm.Freeze(False)
        End Try
    End Sub

    Public Sub LoadForm(ByVal strCardCode As String, ByVal strCardName As String, ByVal strFromDate As String, ByVal strToDate As String, ByVal strProgramID As String)
        Try
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_OPSL) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_OPSL, frm_Z_OPSL)
            oForm = oApplication.SBO_Application.Forms.Item(strUID)
            oForm.DataBrowser.BrowseBy = "14"
            initialize(oForm)
            oForm.DataSources.UserDataSources.Add("PrgName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            Dim oEditText As SAPbouiCOM.EditText
            oEditText = oForm.Items.Item("8__").Specific
            oEditText.DataBind.SetBound(True, "", "PrgName")
            addChooseFromListConditions(oForm)

            oForm.Freeze(True)
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPSL")
            oDBDataSource.SetValue("U_CardCode", 0, strCardCode)
            oDBDataSource.SetValue("U_CardName", 0, strCardName)
            oDBDataSource.SetValue("U_FromDate", 0, strFromDate)
            oDBDataSource.SetValue("U_ProgramID", 0, strProgramID)

            'Overwriting the Program From Date if Selected Date greater than Program ID.
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            strQuery = "Select Convert(VarChar(8),U_PFromDate,112) As 'FromDate' From [@Z_OCPM] Where DocEntry = '" & strProgramID & "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                If CInt(strFromDate) < CInt(oRecordSet.Fields.Item(0).Value.ToString()) Then
                    strFromDate = CInt(oRecordSet.Fields.Item(0).Value.ToString())
                End If
            End If
            oDBDataSource.SetValue("U_FromDate", 0, strFromDate)

            'Overwriting the Program To Date if Selected Date greater than Program ID.
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            strQuery = "Select Convert(VarChar(8),U_PToDate,112) As 'TillDate' From [@Z_OCPM] Where DocEntry = '" & strProgramID & "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                If CInt(strToDate) > CInt(oRecordSet.Fields.Item(0).Value.ToString()) Then
                    strToDate = CInt(oRecordSet.Fields.Item(0).Value.ToString())
                End If
            End If
            oDBDataSource.SetValue("U_TillDate", 0, strToDate)

            oDBDataSource.SetValue("U_Type", 0, "P")
            GetProgramDetails(oForm, strProgramID)
            calculateNoofDays(oForm)

            'oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            'strQuery = " Select T0.DocEntry,ISNULL(T0.U_InvRef,T2.U_InvRef) As 'U_InvRef',T1.DocNum,T0.U_TrnRef From [@Z_OCPM] T0 "
            'strQuery += " LEFT OUTER JOIN [@Z_CPM6] T2 On T0.DocEntry = T2.DocEntry "
            'strQuery += " LEFT OUTER JOIN OINV T1 ON ISNULL(T0.U_InvRef,T2.U_InvRef) = T1.DocEntry "
            'strQuery += " Where U_CardCode = '" + strCardCode + "'"
            'strQuery += " And T0.U_RemDays > 0 "
            'strQuery += " And ISNULL(T0.U_Cancel,'N') = 'N' "
            'strQuery += " Order By U_PFromDate "
            'oRecordSet.DoQuery(strQuery)
            'If Not oRecordSet.EoF Then
            '    If oRecordSet.Fields.Item("U_InvRef").Value.ToString().Length > 0 Then
            '        changeControlBasedOnType(oForm, "I")
            '        oDBDataSource.SetValue("U_Type", 0, "I")
            '        oDBDataSource.SetValue("U_InvoiceNo", 0, oRecordSet.Fields.Item("DocNum").Value.ToString())
            '        oDBDataSource.SetValue("U_InvoiceRef", 0, oRecordSet.Fields.Item("U_InvRef").Value.ToString())
            '        oForm.Update()
            '        GetProgramInvoiceDetails(oForm)
            '    ElseIf oRecordSet.Fields.Item("U_TrnRef").Value.ToString().Length > 0 Then
            '        changeControlBasedOnType(oForm, "T")
            '        oDBDataSource.SetValue("U_Type", 0, "T")
            '        oDBDataSource.SetValue("U_TranNo", 0, oRecordSet.Fields.Item("U_TrnRef").Value.ToString())
            '        oForm.Update()
            '        GetProgramTransferDetails(oForm, oRecordSet.Fields.Item("U_TrnRef").Value.ToString())
            '    Else
            '        oDBDataSource.SetValue("U_Type", 0, "P")
            '        GetProgramDetails(oForm, strProgramID)
            '    End If
            '    calculateNoofDays(oForm)
            'End If

            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm.Freeze(False)
        End Try
    End Sub

    Public Sub LoadForm(ByVal strDocEntry As String)
        Try
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_OPSL) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oForm = oApplication.Utilities.LoadForm(xml_Z_OPSL, frm_Z_OPSL)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.DataBrowser.BrowseBy = "14"
            oForm.Freeze(True)
            initialize(oForm)
            oForm.DataSources.UserDataSources.Add("PrgName", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 100)
            Dim oEditText As SAPbouiCOM.EditText
            oEditText = oForm.Items.Item("8__").Specific
            oEditText.DataBind.SetBound(True, "", "PrgName")
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            oForm.EnableMenu(mnu_ADD, True)
            oForm.Freeze(False)
            addChooseFromListConditions(oForm)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            oForm.Items.Item("14").Specific.value = strDocEntry
            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.EnableMenu(mnu_FIND, False)
            oForm.Items.Item("13").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("9").Enabled = False
            oForm.Items.Item("10").Enabled = False
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm.Freeze(False)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Z_OPSL Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or _
                                                           oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    oForm.Freeze(True)
                                    If validation(oForm) = False Then
                                        oForm.Freeze(False)
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        If Not validate_DaysFood(oForm) Then
                                            If oApplication.SBO_Application.MessageBox("All Foods Not Selected for Specified dates...Continue...?", , "Yes", "No") = 1 Then
                                                If Not validate_Custom(oForm) Then
                                                    'If oApplication.SBO_Application.MessageBox("Select Food Other than Regular Do you want to Proceed?", , "Yes", "No") = 2 Then
                                                    '    oForm.Freeze(False)
                                                    '    BubbleEvent = False
                                                    '    Exit Sub
                                                    'Else
                                                    If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
                                                        oForm.Freeze(False)
                                                        BubbleEvent = False
                                                        Exit Sub
                                                    End If
                                                    'End If
                                                Else
                                                    If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
                                                        oForm.Freeze(False)
                                                        BubbleEvent = False
                                                        Exit Sub
                                                    End If
                                                End If
                                            Else
                                                oForm.Freeze(False)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        Else

                                            If Not validate_Custom(oForm) Then
                                                'If oApplication.SBO_Application.MessageBox("Select Food Other than Regular Do you want to Proceed?", , "Yes", "No") = 2 Then
                                                '    oForm.Freeze(False)
                                                '    BubbleEvent = False
                                                '    Exit Sub
                                                'Else
                                                If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
                                                    oForm.Freeze(False)
                                                    BubbleEvent = False
                                                    Exit Sub
                                                End If
                                                'End If
                                            Else
                                                If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
                                                    oForm.Freeze(False)
                                                    BubbleEvent = False
                                                    Exit Sub
                                                End If
                                            End If

                                        End If
                                    End If
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "18" Then
                                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPSL")
                                    Dim strFromDate As String = oDBDataSource.GetValue("U_FromDate", 0).Trim()
                                    Dim strToDate As String = oDBDataSource.GetValue("U_TillDate", 0).Trim()

                                    If strFromDate <> "" And strToDate <> "" Then
                                        If CInt(strToDate) < CInt(strFromDate) Then
                                            oApplication.Utilities.Message("To Date Should be Greater than or Equal to From Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If

                                    'Calculate no of Days When load Food. Add by Madhu On 20150510
                                    calculateNoofDays(oForm)

                                    'Dim strFromDate As String = oDBDataSource.GetValue("U_FromDate", 0).Trim()
                                    'Dim strToDate As String = oDBDataSource.GetValue("U_TillDate", 0).Trim()
                                    Dim strCardCode As String = oDBDataSource.GetValue("U_CardCode", 0).Trim()
                                    Dim strType As String = oDBDataSource.GetValue("U_Type", 0).Trim()
                                    Dim strProgram As String = oDBDataSource.GetValue("U_Program", 0).Trim()
                                    Dim strProgramID As String = oDBDataSource.GetValue("U_ProgramID", 0).Trim()
                                    Dim strNoofDays As String = oDBDataSource.GetValue("U_NoOfDays", 0).Trim()
                                    Dim strRNoofDays As String = oDBDataSource.GetValue("U_RNoOfDays", 0).Trim()

                                    If CInt(IIf(strNoofDays = "", 0, strNoofDays)) > CInt(IIf(strRNoofDays = "", 0, strRNoofDays)) Then
                                        oApplication.Utilities.Message("Remaining No of Days Lesser then Open No of Days..", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If

                                    Dim strDocRef As String = String.Empty
                                    If strType = "I" Then
                                        strDocRef = oDBDataSource.GetValue("U_InvoiceRef", 0).Trim()
                                    ElseIf strType = "T" Then
                                        strDocRef = oDBDataSource.GetValue("U_TranNo", 0).Trim()
                                    ElseIf strType = "P" Then
                                        strDocRef = oDBDataSource.GetValue("U_ProgramID", 0).Trim()
                                    End If

                                    If Not validation_PreSalesDate(oForm) Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If

                                    If strFromDate.Trim().Length = 0 Or strToDate.Trim().Length = 0 Then
                                        oApplication.Utilities.Message("Enter From & To Date...to Proceed..", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        oApplication.Utilities.updateCustomerProgram(oForm, strCardCode.Trim(), strFromDate, strToDate, _
                                                                                     strType.Trim(), strDocRef.Trim())
                                        Dim oLoadMenu As clsFoodMenu
                                        oLoadMenu = New clsFoodMenu()
                                        oLoadMenu.LoadForm(oForm.UniqueID, strFromDate, strToDate, strCardCode.Trim(), strType.Trim(), _
                                                           strDocRef.Trim(), strProgram.ToString(), strProgramID)

                                    End If
                                ElseIf pVal.ItemUID = "35" Or pVal.ItemUID = "_35" Then
                                    Dim oOption As SAPbouiCOM.OptionBtn
                                    oOption = oForm.Items.Item(pVal.ItemUID).Specific
                                    oForm.Freeze(True)
                                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPSL")
                                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PSL1")
                                    If oOption.Selected Then
                                        If pVal.ItemUID = "_35" Then
                                            oForm.Items.Item("14").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            oForm.Items.Item("_37").Visible = False
                                            oForm.Items.Item("37").Visible = False

                                            oForm.Items.Item("_6").Visible = True
                                            oForm.Items.Item("6").Visible = True

                                            oDBDataSource.SetValue("U_TranNo", 0, "")
                                            oDBDataSource.SetValue("U_ProgramID", 0, "")
                                            oDBDataSource.SetValue("U_FromDate", 0, "")
                                            oDBDataSource.SetValue("U_TillDate", 0, "")
                                            oDBDataSource.SetValue("U_NoOfDays", 0, "")
                                            oForm.DataSources.UserDataSources.Item("PrgName").ValueEx = ""
                                        Else
                                            oForm.Items.Item("14").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            oForm.Items.Item("_37").Visible = True
                                            oForm.Items.Item("37").Visible = True

                                            oForm.Items.Item("_6").Visible = False
                                            oForm.Items.Item("6").Visible = False
                                            oDBDataSource.SetValue("U_InvoiceNo", 0, "")
                                            oDBDataSource.SetValue("U_InvoiceRef", 0, "")
                                            oDBDataSource.SetValue("U_ProgramID", 0, "")
                                            oDBDataSource.SetValue("U_FromDate", 0, "")
                                            oDBDataSource.SetValue("U_TillDate", 0, "")
                                            oDBDataSource.SetValue("U_NoOfDays", 0, "")
                                            oForm.DataSources.UserDataSources.Item("PrgName").ValueEx = ""

                                        End If
                                    End If
                                    oForm.Freeze(False)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "6" Then
                                    filterInvoiceChooseFromList(oForm, "CFL_2")
                                ElseIf pVal.ItemUID = "37" Then
                                    filterTransferChooseFromList(oForm, "CFL_3")
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "9" Or pVal.ItemUID = "10") And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPSL")
                                    Dim strType As String = oDBDataSource.GetValue("U_Type", 0).Trim()
                                    Dim strProgramID As String = oDBDataSource.GetValue("U_ProgramID", 0).Trim()
                                    Dim strFromDate As String = oDBDataSource.GetValue("U_FromDate", 0).Trim()
                                    Dim strToDate As String = oDBDataSource.GetValue("U_TillDate", 0).Trim()
                                    oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    If strType = "I" And strFromDate.Length > 0 Then
                                        strQuery = "Select * From [@Z_OCPM] Where '" + strFromDate + "' "
                                        strQuery += " BETWEEN Convert(VarChar(8),U_PFromDate,112) AND Convert(VarChar(8),U_PToDate,112) "
                                        strQuery += " AND DocEntry ='" + strProgramID + "'"
                                        oRecordSet.DoQuery(strQuery)
                                        If oRecordSet.EoF Then
                                            oDBDataSource.SetValue("U_FromDate", 0, "")
                                            oDBDataSource.SetValue("U_TillDate", 0, "")
                                            oApplication.Utilities.Message("From Date should be greater then Program From Date..", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    ElseIf strType = "P" And (strFromDate.Length > 0 Or strToDate.Length > 0) Then
                                        If strFromDate.Length > 0 Then
                                            strQuery = "Select * From [@Z_OCPM] Where '" & strFromDate & "' "
                                            strQuery += " BETWEEN Convert(VarChar(8),U_PFromDate,112) AND Convert(VarChar(8),U_PToDate,112) "
                                            strQuery += " AND DocEntry ='" + strProgramID + "'"
                                            oRecordSet.DoQuery(strQuery)
                                            If oRecordSet.EoF Then
                                                'oDBDataSource.SetValue("U_FromDate", 0, "")
                                                'oDBDataSource.SetValue("U_TillDate", 0, "")
                                                oApplication.Utilities.Message("From Date should be greater then Program From Date..", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If

                                        If strToDate.Length > 0 Then
                                            strQuery = "Select * From [@Z_OCPM] Where '" & strToDate & "' "
                                            strQuery += " BETWEEN Convert(VarChar(8),U_PFromDate,112) AND Convert(VarChar(8),U_PToDate,112) "
                                            strQuery += " AND DocEntry ='" + strProgramID + "'"
                                            oRecordSet.DoQuery(strQuery)
                                            If oRecordSet.EoF Then
                                                'oDBDataSource.SetValue("U_FromDate", 0, "")
                                                'oDBDataSource.SetValue("U_TillDate", 0, "")
                                                oApplication.Utilities.Message("From Date should be Lesser then Program To Date..", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                                    Case "1"
                                        If pVal.Action_Success Then
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                                initialize(oForm)
                                            End If
                                        End If
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "9" Or pVal.ItemUID = "10" Then
                                    calculateNoofDays(oForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPSL")
                                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PSL1")
                                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                    Dim oDataTable As SAPbouiCOM.DataTable
                                    Try
                                        oCFLEvento = pVal
                                        oDataTable = oCFLEvento.SelectedObjects

                                        If IsNothing(oDataTable) Then
                                            Exit Sub
                                        End If

                                        If pVal.ItemUID = "4" Then
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 0 Then
                                                oDBDataSource.SetValue("U_CardCode", 0, oDataTable.GetValue("CardCode", 0))
                                                oDBDataSource.SetValue("U_CardName", 0, oDataTable.GetValue("CardName", 0))
                                            End If
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf pVal.ItemUID = "5" Then
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 0 Then
                                                oDBDataSource.SetValue("U_CardCode", 0, oDataTable.GetValue("CardCode", 0))
                                                oDBDataSource.SetValue("U_CardName", 0, oDataTable.GetValue("CardName", 0))

                                                oDBDataSource.SetValue("U_InvoiceNo", 0, "")
                                                oDBDataSource.SetValue("U_InvoiceRef", 0, "")
                                                oDBDataSource.SetValue("U_TranNo", 0, "")
                                                oDBDataSource.SetValue("U_ProgramID", 0, "")
                                                oDBDataSource.SetValue("U_FromDate", 0, "")
                                                oDBDataSource.SetValue("U_TillDate", 0, "")
                                                oDBDataSource.SetValue("U_NoOfDays", 0, "")
                                                oForm.DataSources.UserDataSources.Item("PrgName").ValueEx = ""

                                                oMatrix = oForm.Items.Item("3").Specific
                                                oMatrix.Clear()
                                                oMatrix.FlushToDataSource()
                                                oMatrix.LoadFromDataSource()

                                                oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                                'Madhu Modified this Query for Phase II Requirement 20150708.
                                                'strQuery = "Select T0.DocEntry,U_InvRef,T1.DocNum,U_TrnRef From [@Z_OCPM] T0 "
                                                'strQuery += " LEFT OUTER JOIN OINV T1 ON T0.U_InvRef = T1.DocEntry "
                                                'strQuery += " Where U_CardCode = '" + oDataTable.GetValue("CardCode", 0).ToString().Trim() + "'"
                                                'strQuery += " And U_RemDays > 0 "
                                                'strQuery += " And ISNULL(T0.U_Cancel,'N') = 'N' "
                                                'strQuery += " Order By U_PFromDate "

                                                strQuery = " Select T0.DocEntry,ISNULL(T0.U_InvRef,T2.U_InvRef) As 'U_InvRef',T1.DocNum,T0.U_TrnRef From [@Z_OCPM] T0 "
                                                strQuery += " LEFT OUTER JOIN [@Z_CPM6] T2 On T0.DocEntry = T2.DocEntry "
                                                strQuery += " LEFT OUTER JOIN OINV T1 ON ISNULL(T0.U_InvRef,T2.U_InvRef) = T1.DocEntry "
                                                strQuery += " Where U_CardCode = '" + oDataTable.GetValue("CardCode", 0).ToString().Trim() + "'"
                                                strQuery += " And T0.U_RemDays > 0 "
                                                strQuery += " And ISNULL(T0.U_Cancel,'N') = 'N' "
                                                strQuery += " Order By U_PFromDate "
                                                oRecordSet.DoQuery(strQuery)
                                                If Not oRecordSet.EoF Then
                                                    If oRecordSet.Fields.Item("U_InvRef").Value.ToString().Length > 0 Then
                                                        changeControlBasedOnType(oForm, "I")
                                                        oDBDataSource.SetValue("U_Type", 0, "I")
                                                        'CType(oForm.Items.Item("_35").Specific, SAPbouiCOM.OptionBtn).Selected = True
                                                        'oDBDataSource.SetValue("U_InvoiceNo", 0, oRecordSet.Fields.Item("DocNum").Value.ToString())
                                                        oDBDataSource.SetValue("U_InvoiceNo", 0, oRecordSet.Fields.Item("DocNum").Value.ToString())
                                                        oDBDataSource.SetValue("U_InvoiceRef", 0, oRecordSet.Fields.Item("U_InvRef").Value.ToString())
                                                        oForm.Update()
                                                        GetProgramInvoiceDetails(oForm)
                                                    ElseIf oRecordSet.Fields.Item("U_TrnRef").Value.ToString().Length > 0 Then
                                                        changeControlBasedOnType(oForm, "T")
                                                        'CType(oForm.Items.Item("35").Specific, SAPbouiCOM.OptionBtn).Selected = True
                                                        oDBDataSource.SetValue("U_Type", 0, "T")
                                                        oDBDataSource.SetValue("U_TranNo", 0, oRecordSet.Fields.Item("U_TrnRef").Value.ToString())
                                                        oForm.Update()
                                                        GetProgramTransferDetails(oForm, oRecordSet.Fields.Item("U_TrnRef").Value.ToString())
                                                    End If
                                                End If
                                            End If
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf pVal.ItemUID = "6" Then
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 0 Then
                                                oDBDataSource.SetValue("U_InvoiceNo", 0, oDataTable.GetValue("DocNum", 0))
                                                oDBDataSource.SetValue("U_InvoiceRef", 0, oDataTable.GetValue("DocEntry", 0))
                                                oForm.Update()
                                                GetProgramInvoiceDetails(oForm)
                                            End If
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf pVal.ItemUID = "37" Then
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 0 Then
                                                oDBDataSource.SetValue("U_TranNo", 0, oDataTable.GetValue("DocEntry", 0))
                                                oForm.Update()
                                                GetProgramTransferDetails(oForm, oDataTable.GetValue("DocEntry", 0))
                                            End If
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                    Catch ex As Exception
                                        oApplication.Log.Trace_DIET_AddOn_Error(ex)
                                        Throw ex
                                        'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
                                    End Try
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                If oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Or oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    Try
                                        reDrawForm(oForm)
                                    Catch ex As Exception
                                        'oApplication.Log.Trace_DIET_AddOn_Error(ex)
                                    End Try
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.BeforeAction
                Case True
                    Select Case pVal.MenuUID

                        Case mnu_GenerateSO
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            oDBDataSource = oForm.DataSources.DBDataSources.Item(0)
                            If Not oDBDataSource.GetValue("DocEntry", 0).ToString = "" Then
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                                    Dim _retVal As Integer = oApplication.SBO_Application.MessageBox("Sure you wanted to Create Sales Order...?", 2, "Yes", "No", "")
                                    If _retVal = 2 Then
                                        Exit Sub
                                    End If
                                    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim strDocEntry As String = CType(oForm.Items.Item("14").Specific, SAPbouiCOM.EditText).Value
                                    strQuery = "Select U_SalesO From [@Z_OPSL] Where DocEntry ='" + strDocEntry + "'"
                                    oRecordSet.DoQuery(strQuery)
                                    If Not oRecordSet.EoF Then
                                        If oRecordSet.Fields.Item(0).Value.ToString().Length = 0 Then
                                            If (oApplication.Utilities.AddOrder(oForm, strDocEntry)) Then
                                                oApplication.SBO_Application.MessageBox("Sales Order Created Successfully...")
                                                oApplication.SBO_Application.Menus.Item(mnu_ADD).Activate()
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Case mnu_ViewSO
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            oForm.Items.Item("32").Visible = True
                            oForm.Items.Item("32").Click(SAPbouiCOM.BoCellClickType.ct_Linked)
                            oForm.Items.Item("32").Visible = False
                    End Select
                Case False
                    Select Case pVal.MenuUID
                        Case mnu_ADD
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            initialize(oForm)
                            oForm.Items.Item("3").Enabled = True
                            oForm.Items.Item("18").Enabled = True
                            oForm.Items.Item("12").Enabled = True
                            oForm.DataSources.UserDataSources.Item("PrgName").ValueEx = ""
                        Case mnu_FIND
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            oForm.Items.Item("3").Enabled = False
                            oForm.Items.Item("18").Enabled = True
                            oForm.Items.Item("12").Enabled = True
                            oForm.DataSources.UserDataSources.Item("PrgName").ValueEx = ""
                        Case mnu_Z_OPSL
                            LoadForm()
                            'LoadForm("CR0001", "DIALA WALID  TABBARA", "20150702", "20150702")
                    End Select
            End Select
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Data Event"

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            Select Case BusinessObjectInfo.BeforeAction
                Case True
                Case False
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                            If BusinessObjectInfo.ActionSuccess Then
                                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                                Dim oXmlDoc As System.Xml.XmlDocument = New Xml.XmlDocument()
                                oXmlDoc.LoadXml(BusinessObjectInfo.ObjectKey)
                                Dim strDocEntry As String = oXmlDoc.SelectSingleNode("/Pre_Sales_OrderParams/DocEntry").InnerText
                                If (oApplication.Utilities.AddOrder(oForm, strDocEntry)) Then
                                    oApplication.SBO_Application.MessageBox("Sales Order Created Successfully...")
                                Else
                                    oApplication.SBO_Application.MessageBox(oApplication.Company.GetLastErrorDescription(), 1, "OK", "", "")
                                End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            If BusinessObjectInfo.ActionSuccess Then
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPSL")
                                oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PSL1")

                                Dim strType As String = oDBDataSource.GetValue("U_Type", 0).Trim()
                                oForm.Freeze(True)
                                If strType = "I" Then
                                    oForm.Items.Item("14").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oForm.Items.Item("_37").Visible = False
                                    oForm.Items.Item("37").Visible = False

                                    oForm.Items.Item("_6").Visible = True
                                    oForm.Items.Item("6").Visible = True
                                Else
                                    oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oForm.Items.Item("_6").Visible = False
                                    oForm.Items.Item("_6").Visible = False

                                    oForm.Items.Item("_37").Visible = True
                                    oForm.Items.Item("37").Visible = True
                                End If

                                If (oDBDataSource.GetValue("Status", oDBDataSource.Offset) = "L") Then
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                Else
                                    Dim strOrder As String = CType(oForm.Items.Item("31").Specific, SAPbouiCOM.EditText).Value
                                    If strOrder.Length > 0 Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                    Else
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                    End If
                                End If

                                'Just to Show Program Name While Old Record in Screen. On 20150510
                                Dim oRecordSet As SAPbobsCOM.Recordset
                                oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                Dim strProCode As String = oDBDataSource.GetValue("U_Program", 0).Trim()
                                oRecordSet.DoQuery("Select ItemName From OITM Where ItemCode = '" + strProCode + "'")
                                If Not oRecordSet.EoF Then
                                    oForm.DataSources.UserDataSources.Item("PrgName").ValueEx = oRecordSet.Fields.Item(0).Value.ToString()
                                End If

                                oForm.Freeze(False)
                            End If
                    End Select
            End Select
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

#End Region

#Region "Right Click Event"

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
        oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPSL")
        oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PSL1")

        If oForm.TypeEx = frm_Z_OPSL Then
            Dim oMenuItem As SAPbouiCOM.MenuItem
            Dim oMenus As SAPbouiCOM.Menus
            oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data
            If (eventInfo.BeforeAction = True) Then
                Try
                    oDBDataSource = oForm.DataSources.DBDataSources.Item(0)

                    If Not oMenuItem.SubMenus.Exists(mnu_GenerateSO) And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                        oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage.UniqueID = mnu_GenerateSO
                        oCreationPackage.String = "Generate Sales Order"
                        If oDBDataSource.GetValue("U_SalesO", oDBDataSource.Offset).Trim().Length > 0 Then
                            oCreationPackage.Enabled = False
                        End If
                        oMenus = oMenuItem.SubMenus
                        oMenus.AddEx(oCreationPackage)
                    End If

                    If Not oMenuItem.SubMenus.Exists(mnu_ViewSO) And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE) Then
                        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                        oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage.UniqueID = mnu_ViewSO
                        oCreationPackage.String = "View Sales Order"
                        If oDBDataSource.GetValue("U_SalesO", oDBDataSource.Offset).Trim().Length > 0 Then
                            oCreationPackage.Enabled = True
                        End If
                        oMenus = oMenuItem.SubMenus
                        oMenus.AddEx(oCreationPackage)
                    End If
                    oMenuItem.SubMenus.Item(mnu_CANCEL).Enabled = False
                    oMenuItem.SubMenus.Item(mnu_CLOSE).Enabled = False
                Catch ex As Exception
                    oApplication.Log.Trace_DIET_AddOn_Error(ex)
                    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try
            Else
                If oMenuItem.SubMenus.Exists(mnu_GenerateSO) Then
                    oMenuItem.SubMenus.RemoveEx(mnu_GenerateSO)
                End If
                If oMenuItem.SubMenus.Exists(mnu_ViewSO) Then
                    oMenuItem.SubMenus.RemoveEx(mnu_ViewSO)
                End If
                oMenuItem.SubMenus.Item(mnu_CANCEL).Enabled = True
                oMenuItem.SubMenus.Item(mnu_CLOSE).Enabled = True
            End If
        End If
    End Sub

#End Region

#Region "Function"

    Private Sub calculateNoofDays(ByVal oForm As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPSL")
            Dim strCardCode As String = oDBDataSource.GetValue("U_CardCode", 0).Trim()
            Dim strFromDate As String = oDBDataSource.GetValue("U_FromDate", 0).Trim()
            Dim strToDate As String = oDBDataSource.GetValue("U_TillDate", 0).Trim()
            Dim strType As String = oDBDataSource.GetValue("U_Type", 0).Trim()
            Dim strDocRef As String = String.Empty
            If strType = "I" Then
                strDocRef = oDBDataSource.GetValue("U_InvoiceRef", 0).Trim()
            ElseIf strType = "T" Then
                strDocRef = oDBDataSource.GetValue("U_TranNo", 0).Trim()
            ElseIf strType = "P" Then
                strDocRef = oDBDataSource.GetValue("U_ProgramID", 0).Trim()
            End If
            If strFromDate.Length > 0 And strToDate.Length > 0 Then
                Dim intNoofDays As Integer = oApplication.Utilities.getDateDiff_PreSales(oForm, strCardCode.Trim(), _
                                                                                         strFromDate, strToDate, _
                                                                                         strType, strDocRef)
                Dim intReminDays As Integer = CInt(oDBDataSource.GetValue("U_RNoOfDays", 0).Trim())

                If intNoofDays > intReminDays Then
                    oDBDataSource.SetValue("U_NoOfDays", 0, intReminDays)
                Else
                    oDBDataSource.SetValue("U_NoOfDays", 0, intNoofDays)
                End If

                oForm.Update()
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.PaneLevel = 1
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select IsNull(MAX(DocEntry),0) +1 From [@Z_OPSL]")
            If Not oRecordSet.EoF Then
                oApplication.Utilities.setEditText(oForm, "13", oRecordSet.Fields.Item(0).Value.ToString())
                oForm.Items.Item("12").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                'oApplication.Utilities.setEditText(oForm, "12", "t")
                'oApplication.SBO_Application.SendKeys("{TAB}")
            End If

            Dim oOption As SAPbouiCOM.OptionBtn
            oOption = oForm.Items.Item("_35").Specific
            oOption.Selected = True

            oForm.Items.Item("14").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            MatrixID = "3"
            oForm.Update()
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim _retVal As Boolean = True
        Try
            Dim strCardCode, strCardName, strInvoiceNo, strTransferNo, strRNoDays As String
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPSL")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PSL1")

            strCardCode = oDBDataSource.GetValue("U_CardCode", 0)
            strCardName = oDBDataSource.GetValue("U_CardName", 0)
            strInvoiceNo = oDBDataSource.GetValue("U_InvoiceNo", 0)
            strTransferNo = oDBDataSource.GetValue("U_TranNo", 0)
            strRNoDays = oDBDataSource.GetValue("U_RNoOfDays", 0)

            Dim strFrDt As String = oDBDataSource.GetValue("U_FromDate", 0).Trim()
            Dim strToDt As String = oDBDataSource.GetValue("U_TillDate", 0).Trim()
            Dim strNoofDays As String = oDBDataSource.GetValue("U_NoOfDays", 0).Trim()

            If strCardCode = "" Then
                oApplication.Utilities.Message("Select Customer Code ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strCardName = "" Then
                oApplication.Utilities.Message("Select Customer Name ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strFrDt = "" Then
                oApplication.Utilities.Message("Select From Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strToDt = "" Then
                oApplication.Utilities.Message("Select Till Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                'ElseIf (strInvoiceNo.Trim() = "" And strTransferNo.Trim() = "") Then
                '    oApplication.Utilities.Message("Select Invoice No / Transfer No For the Customer to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Return False
            ElseIf strFrDt.Length > 0 And strToDt.Length > 0 And CInt(strFrDt) > CInt(strToDt) Then
                oApplication.Utilities.Message("From Date Should be Lesser than or Equal Menu Till Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strNoofDays <> "" And strRNoDays <> "" And CInt(strNoofDays) > CInt(strRNoDays) Then
                oApplication.Utilities.Message("Selected No of Days Should be Less than or Equal to Remaining Days ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If Not validation_PreSalesDate(oForm) Then
                If oApplication.SBO_Application.MessageBox("Overlapping of Pre Sales Order Dates Found. Do you want to Proceed?", , "Yes", "No") = 2 Then
                    Return False
                    Exit Function
                End If
            End If

            oMatrix = oForm.Items.Item("3").Specific
            If oMatrix.RowCount = 0 Then
                oApplication.Utilities.Message("Cannot add document with out selecting any Food...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            Else
                For index As Integer = 0 To oDBDataSourceLines.Size - 1
                    Dim strPrgDate As String = oDBDataSourceLines.GetValue("U_DelDate", index).ToString
                    If strPrgDate <> "U_DelDate" Then
                        strQuery = "Select T0.DocEntry From RDR1 T0 "
                        strQuery += " Where T0.BaseCard = '" + oDBDataSource.GetValue("U_CardCode", 0).Trim() + "'"
                        strQuery += " AND '" + strPrgDate + "' Between Convert(VarChar(8),T0.U_DelDate,112) And Convert(VarChar(8),T0.U_DelDate,112)  "
                        strQuery += " And ((T0.LineStatus = 'O') OR (T0.LineStatus = 'C' And T0.TargetType <> '-1')) "
                        oRecordSet.DoQuery(strQuery)
                        If Not oRecordSet.EoF Then
                            oApplication.Utilities.Message("Cannot add document Other Order Exists for Program Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                            Exit Function
                        End If
                    End If
                Next
            End If

            Return True
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
        Return _retVal
    End Function

    Private Function validation_PreSalesDate(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim _retVal As Boolean = True
        Try
            Dim strCardCode, strCardName, strInvoiceNo, strTransferNo, strType, strProgramID, strDocEntry As String
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            strCardCode = oDBDataSource.GetValue("U_CardCode", 0).Trim()
            strCardName = oDBDataSource.GetValue("U_CardName", 0).Trim()
            strInvoiceNo = oDBDataSource.GetValue("U_InvoiceNo", 0).Trim()
            strTransferNo = oDBDataSource.GetValue("U_TranNo", 0).Trim()
            strType = oDBDataSource.GetValue("U_Type", 0).Trim()
            strProgramID = oDBDataSource.GetValue("U_ProgramID", 0).Trim()
            strDocEntry = oDBDataSource.GetValue("DocEntry", 0).Trim()

            Dim strFrDt As String = oDBDataSource.GetValue("U_FromDate", 0).Trim()
            Dim strToDt As String = oDBDataSource.GetValue("U_TillDate", 0).Trim()
            Dim strNoofDays As String = oDBDataSource.GetValue("U_NoOfDays", 0).Trim()

            If strType = "I" Then
                'From Date
                'strQuery = "Select Convert(VarChar(8),U_TillDate+1,112) As U_FromDate,Max(U_TillDate) From [@Z_OPSL] "
                'strQuery += " Where U_ProgramID =  '" + strProgramID + "' "
                'strQuery += " And U_CardCode = '" + oDBDataSource.GetValue("U_CardCode", 0).Trim() + "'"
                'strQuery += " And DocEntry <> '" + strDocEntry + "'"
                'strQuery += " And U_Type = 'I'"
                'strQuery += " Group By U_TillDate "
                'strQuery += " Order By U_TillDate Desc "

                strQuery = "Select DocEntry From [@Z_OPSL] "
                strQuery += " Where U_ProgramID =  '" + strProgramID + "' "
                strQuery += " And U_CardCode = '" + oDBDataSource.GetValue("U_CardCode", 0).Trim() + "'"
                strQuery += " And DocEntry <> '" + strDocEntry + "'"
                strQuery += " And U_Type = 'I'"
                strQuery += " AND '" + strFrDt + "' Between Convert(VarChar(8),U_FromDate,112) And Convert(VarChar(8),U_TillDate,112)  "
                strQuery += " Group By DocEntry "
                strQuery += " Order By DocEntry Desc "
                oRecordSet.DoQuery(strQuery)
                If oRecordSet.EoF Then

                Else
                    If oApplication.SBO_Application.MessageBox("Overlapping of Pre Sales Order From Date Found. Do you want to Proceed?", , "Yes", "No") = 2 Then
                        Return False
                        Exit Function
                    End If
                End If
            ElseIf strType = "T" Then
                'To Date
                'strQuery = "Select Convert(VarChar(8),U_TillDate+1,112) As U_FromDate,Max(U_TillDate) From [@Z_OPSL] "
                'strQuery += " Where U_ProgramID =  '" + strProgramID + "' "
                'strQuery += " And U_CardCode = '" + oDBDataSource.GetValue("U_CardCode", 0).Trim() + "'"
                'strQuery += " And DocEntry <> '" + strDocEntry + "'"
                'strQuery += " And U_Type = 'T'"
                'strQuery += " Group By U_TillDate "
                'strQuery += " Order By U_TillDate Desc "

                strQuery = "Select DocEntry From [@Z_OPSL] "
                strQuery += " Where U_ProgramID =  '" + strProgramID + "' "
                strQuery += " And U_CardCode = '" + oDBDataSource.GetValue("U_CardCode", 0).Trim() + "'"
                strQuery += " And DocEntry <> '" + strDocEntry + "'"
                strQuery += " And U_Type = 'T'"
                strQuery += " AND '" + strToDt + "' Between Convert(VarChar(8),U_FromDate,112) And Convert(VarChar(8),U_TillDate,112)  "
                strQuery += " Group By DocEntry "
                strQuery += " Order By DocEntry Desc "
                oRecordSet.DoQuery(strQuery)
                If oRecordSet.EoF Then

                Else
                    If oApplication.SBO_Application.MessageBox("Overlapping of Pre Sales Order To Date Found. Do you want to Proceed?", , "Yes", "No") = 2 Then
                        Return False
                        Exit Function
                    End If
                End If
            ElseIf strType = "P" Then
                'To Date
                'strQuery = "Select Convert(VarChar(8),U_TillDate+1,112) As U_FromDate,Max(U_TillDate) From [@Z_OPSL] "
                'strQuery += " Where U_ProgramID =  '" + strProgramID + "' "
                'strQuery += " And U_CardCode = '" + oDBDataSource.GetValue("U_CardCode", 0).Trim() + "'"
                'strQuery += " And DocEntry <> '" + strDocEntry + "'"
                'strQuery += " And U_Type = 'T'"
                'strQuery += " Group By U_TillDate "
                'strQuery += " Order By U_TillDate Desc "

                strQuery = "Select DocEntry From [@Z_OPSL] "
                strQuery += " Where U_ProgramID =  '" + strProgramID + "' "
                strQuery += " And U_CardCode = '" + oDBDataSource.GetValue("U_CardCode", 0).Trim() + "'"
                strQuery += " And DocEntry <> '" + strDocEntry + "'"
                strQuery += " And U_Type = 'P'"
                strQuery += " AND '" + strToDt + "' Between Convert(VarChar(8),U_FromDate,112) And Convert(VarChar(8),U_TillDate,112)  "
                strQuery += " Group By DocEntry "
                strQuery += " Order By DocEntry Desc "
                oRecordSet.DoQuery(strQuery)
                If oRecordSet.EoF Then

                Else
                    If oApplication.SBO_Application.MessageBox("Overlapping of Pre Sales Order To Date Found. Do you want to Proceed?", , "Yes", "No") = 2 Then
                        Return False
                        Exit Function
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
        Return _retVal
    End Function

    Private Function validate_DaysFood(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Dim _retVal As Boolean = True
        Try
            oMatrix = oForm.Items.Item("3").Specific
            Dim strNoofDays As String = oDBDataSource.GetValue("U_NoOfDays", 0).Trim()
            If (CInt(strNoofDays) * 6) <> oMatrix.VisualRowCount Then
                oApplication.Utilities.Message("Select All Food for all dates to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False
            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Private Function validate_Custom(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Dim _retVal As Boolean = True
        Try
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PSL1")
            oMatrix = oForm.Items.Item("3").Specific
            For intRow As Integer = 0 To oDBDataSourceLines.Size - 1
                If oDBDataSourceLines.GetValue("U_SFood", intRow).ToString() = "C" Then
                    _retVal = False
                    Exit For
                End If
            Next
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Private Sub addChooseFromListConditions(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList

            oCFLs = oForm.ChooseFromLists

            oCFL = oCFLs.Item("CFL_1")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)

            'oCFL = oCFLs.Item("CFL_4")
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "CardType"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "C"
            'oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_4")
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

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub GetProgramInvoiceDetails(ByVal oForm As SAPbouiCOM.Form)
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet, oRecordSet1 As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet1 = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            'Madhu Modified this Query for Phase II Requirement on 20150708
            'strQuery = "Select DocEntry,U_PrgCode,Convert(VarChar(8),U_PFromDate,112) As 'FD',U_PrgName From [@Z_OCPM] Where U_InvRef = '" + oForm.Items.Item("7").Specific.value + "'"
            strQuery = " Select T0.DocEntry,T0.U_PrgCode,Convert(VarChar(8),T0.U_PFromDate,112) As 'FD',T0.U_PrgName From [@Z_OCPM] T0 "
            strQuery += " LEFT OUTER JOIN [@Z_CPM6] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " Where ISNULL(T0.U_InvRef,T1.U_InvRef) = '" + oForm.Items.Item("7").Specific.value + "' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then

                oDBDataSource.SetValue("U_ProgramID", 0, oRecordSet.Fields.Item(0).Value.ToString())
                oDBDataSource.SetValue("U_Program", 0, oRecordSet.Fields.Item(1).Value.ToString())
                oForm.DataSources.UserDataSources.Item("PrgName").ValueEx = oRecordSet.Fields.Item("U_PrgName").Value.ToString()

                'Madhu Modified this Query for Phase II Requirement on 20150708
                'strQuery = "Select U_RemDays,U_Transfer from [@Z_OCPM] "
                'strQuery += " Where U_InvRef = '" + oForm.Items.Item("7").Specific.value + "'"
                strQuery = "Select T0.U_RemDays,T0.U_Transfer from [@Z_OCPM] T0 "
                strQuery += " LEFT OUTER JOIN [@Z_CPM6] T1 On T0.DocEntry = T1.DocEntry "
                strQuery += " Where ISNULL(T0.U_InvRef,T1.U_InvRef) = '" + oForm.Items.Item("7").Specific.value + "'"
                oRecordSet1.DoQuery(strQuery)
                If Not oRecordSet1.EoF Then
                    If oRecordSet1.Fields.Item("U_Transfer").Value.ToString() = "Y" Then
                        Throw New Exception("Program Already Transfered...to Other Customer...")
                    ElseIf CInt(oRecordSet1.Fields.Item("U_RemDays").Value.ToString()) = 0 Then
                        Throw New Exception("Remaining No of Days Already Zero Cannot Proceed...")
                    End If
                    oDBDataSource.SetValue("U_RNoOfDays", 0, CInt(oRecordSet1.Fields.Item("U_RemDays").Value.ToString()))
                End If

                'From Date
                strQuery = "Select Convert(VarChar(8),U_TillDate+1,112) As U_FromDate,Max(U_TillDate) From [@Z_OPSL] Where U_ProgramID =  '" + oRecordSet.Fields.Item(0).Value.ToString() + "' "
                strQuery += " And U_CardCode = '" + oDBDataSource.GetValue("U_CardCode", 0).Trim() + "'"
                strQuery += " And U_Type = 'I'"
                strQuery += " Group By U_TillDate "
                strQuery += " Order By U_TillDate Desc "
                oRecordSet1.DoQuery(strQuery)
                If oRecordSet1.EoF Then
                    Dim oFromDate As String = oRecordSet.Fields.Item("FD").Value.ToString()
                    oDBDataSource.SetValue("U_FromDate", 0, oFromDate)
                Else
                    Dim oFromDate As String = oRecordSet1.Fields.Item(0).Value.ToString()
                    oDBDataSource.SetValue("U_FromDate", 0, oFromDate)
                End If
                oForm.Update()

                'To Date
                'strQuery = "Select Convert(VarChar(8),U_TillDate,112) As U_TillDate,Max(U_TillDate) From [@Z_OPSL] Where U_ProgramID =  '" + oRecordSet.Fields.Item(0).Value.ToString() + "' Group By U_TillDate  "
                'oRecordSet1.DoQuery(strQuery)
                'If oRecordSet1.EoF Then
                '    oDBDataSource.SetValue("U_TillDate", 0, oRecordSet.Fields.Item(2).Value.ToString())
                '    oDBDataSource.SetValue("U_NoOfDays", 0, CInt(oRecordSet.Fields.Item(3).Value) + 1)
                'Else
                '    Dim oToDate As String = oRecordSet1.Fields.Item(0).Value.ToString()
                '    oDBDataSource.SetValue("U_TillDate", 0, oToDate)
                '    Dim intNoDays As Integer = CInt(oDBDataSource.GetValue("U_TillDate", 0)) - _
                '        CInt(oDBDataSource.GetValue("U_FromDate", 0))
                '    oDBDataSource.SetValue("U_NoOfDays", 0, intNoDays + 1)
                'End If

                'oDBDataSource.SetValue("U_TillDate", 0, oRecordSet.Fields.Item(2).Value.ToString())
                'Dim strCardCode As String = oDBDataSource.GetValue("U_CardCode", 0)
                'Dim strFromDate As String = oDBDataSource.GetValue("U_FromDate", 0)
                'Dim strToDate As String = oDBDataSource.GetValue("U_TillDate", 0)
                'If strFromDate.Length > 0 And strToDate > 0 Then
                '    'Dim intNoofDays As Integer = oApplication.Utilities.getDateDiff(oForm, strCardCode.Trim(), strFromDate, strToDate)
                '    'oDBDataSource.SetValue("U_NoOfDays", 0, intNoofDays)
                '    'oForm.Update()
                'End If

            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub GetProgramTransferDetails(ByVal oForm As SAPbouiCOM.Form, strTransNo As String)
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet, oRecordSet1 As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet1 = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            strQuery = "Select DocEntry,U_PrgCode,Convert(VarChar(8),U_PFromDate,112) As 'FD',U_PrgName From [@Z_OCPM] Where U_TrnRef = '" + strTransNo + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then

                oDBDataSource.SetValue("U_ProgramID", 0, oRecordSet.Fields.Item(0).Value.ToString())
                oDBDataSource.SetValue("U_Program", 0, oRecordSet.Fields.Item(1).Value.ToString())
                oForm.DataSources.UserDataSources.Item("PrgName").ValueEx = oRecordSet.Fields.Item("U_PrgName").Value.ToString()

                strQuery = "Select U_RemDays,U_Transfer from [@Z_OCPM] "
                strQuery += " Where U_TrnRef = '" + strTransNo + "'"
                oRecordSet1.DoQuery(strQuery)
                If Not oRecordSet1.EoF Then
                    If oRecordSet1.Fields.Item("U_Transfer").Value.ToString() = "Y" Then
                        Throw New Exception("Program Already Transfered...to Other Customer...")
                    ElseIf CInt(oRecordSet1.Fields.Item("U_RemDays").Value.ToString()) = 0 Then
                        Throw New Exception("Remaining No of Days Already Zero Cannot Proceed...")
                    End If
                    oDBDataSource.SetValue("U_RNoOfDays", 0, CInt(oRecordSet1.Fields.Item("U_RemDays").Value.ToString()))
                End If

                'From Date
                strQuery = " Select Convert(VarChar(8),U_TillDate+1,112) As U_FromDate,Max(U_TillDate) From [@Z_OPSL] Where U_ProgramID =  '" + oRecordSet.Fields.Item(0).Value.ToString() + "' "
                strQuery += " And U_CardCode = '" + oDBDataSource.GetValue("U_CardCode", 0).Trim() + "'"
                strQuery += " And U_Type = 'T' "
                strQuery += " Group By U_TillDate "
                strQuery += " Order By U_TillDate Desc "
                oRecordSet1.DoQuery(strQuery)
                If oRecordSet1.EoF Then
                    Dim oFromDate As String = oRecordSet.Fields.Item("FD").Value.ToString()
                    oDBDataSource.SetValue("U_FromDate", 0, oFromDate)
                Else
                    Dim oFromDate As String = oRecordSet1.Fields.Item(0).Value.ToString()
                    oDBDataSource.SetValue("U_FromDate", 0, oFromDate)
                End If

                'oDBDataSource.SetValue("U_TillDate", 0, oRecordSet.Fields.Item(2).Value.ToString())
                'Dim strCardCode As String = oDBDataSource.GetValue("U_CardCode", 0)
                'Dim strFromDate As String = oDBDataSource.GetValue("U_FromDate", 0)
                'Dim strToDate As String = oDBDataSource.GetValue("U_TillDate", 0)
                'If strFromDate.Length > 0 And strToDate > 0 Then
                '    Dim intNoofDays As Integer = oApplication.Utilities.getDateDiff(oForm, strCardCode.Trim(), strFromDate, strToDate)
                '    oDBDataSource.SetValue("U_NoOfDays", 0, intNoofDays)
                '    oForm.Update()
                'End If

            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub GetProgramDetails(ByVal oForm As SAPbouiCOM.Form, strProgramNo As String)
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet, oRecordSet1 As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet1 = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            strQuery = "Select DocEntry,U_PrgCode,Convert(VarChar(8),U_PFromDate,112) As 'FD',U_PrgName From [@Z_OCPM] Where DocEntry = '" + strProgramNo + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then

                oDBDataSource.SetValue("U_ProgramID", 0, oRecordSet.Fields.Item(0).Value.ToString())
                oDBDataSource.SetValue("U_Program", 0, oRecordSet.Fields.Item(1).Value.ToString())
                oForm.DataSources.UserDataSources.Item("PrgName").ValueEx = oRecordSet.Fields.Item("U_PrgName").Value.ToString()

                strQuery = "Select U_RemDays,U_Transfer from [@Z_OCPM] "
                strQuery += " Where DocEntry = '" + strProgramNo + "'"
                oRecordSet1.DoQuery(strQuery)
                If Not oRecordSet1.EoF Then
                    If oRecordSet1.Fields.Item("U_Transfer").Value.ToString() = "Y" Then
                        Throw New Exception("Program Already Transfered...to Other Customer...")
                    ElseIf CInt(oRecordSet1.Fields.Item("U_RemDays").Value.ToString()) = 0 Then
                        Throw New Exception("Remaining No of Days Already Zero Cannot Proceed...")
                    End If
                    oDBDataSource.SetValue("U_RNoOfDays", 0, CInt(oRecordSet1.Fields.Item("U_RemDays").Value.ToString()))
                End If

                'From Date
                'strQuery = " Select Convert(VarChar(8),U_TillDate+1,112) As U_FromDate,Max(U_TillDate) From [@Z_OPSL] Where U_ProgramID =  '" + oRecordSet.Fields.Item(0).Value.ToString() + "' "
                'strQuery += " And U_CardCode = '" + oDBDataSource.GetValue("U_CardCode", 0).Trim() + "'"
                'strQuery += " And U_Type = 'P' "
                'strQuery += " Group By U_TillDate "
                'strQuery += " Order By U_TillDate Desc "
                'oRecordSet1.DoQuery(strQuery)
                'If oRecordSet1.EoF Then
                '    Dim oFromDate As String = oRecordSet.Fields.Item("FD").Value.ToString()
                '    oDBDataSource.SetValue("U_FromDate", 0, oFromDate)
                'Else
                '    Dim oFromDate As String = oRecordSet1.Fields.Item(0).Value.ToString()
                '    oDBDataSource.SetValue("U_FromDate", 0, oFromDate)
                'End If

            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub filterInvoiceChooseFromList(ByVal oForm As SAPbouiCOM.Form, ByVal strCFLID As String)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList

            oCFLs = oForm.ChooseFromLists

            'oCFL = oCFLs.Item(strCFLID)

            'oCons = oCFL.GetConditions()

            'If oCons.Count = 0 Then
            '    oCon = oCons.Add()
            'Else
            '    oCon = oCons.Item(0)
            'End If

            'oCon.Alias = "CardCode"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = oForm.Items.Item("4").Specific.value
            'oCFL.SetConditions(oCons)


            oCFL = oCFLs.Item(strCFLID)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.BracketOpenNum = 2

            For i As Integer = 0 To 1
                If i > 0 Then
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCon = oCons.Add()
                    oCon.BracketOpenNum = 1
                End If
                If i = 0 Then
                    oCon.[Alias] = "CardCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oForm.Items.Item("4").Specific.value
                ElseIf i = 1 Then
                    oCon.[Alias] = "CANCELED"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "N"
                ElseIf i = 2 Then
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
                End If
                If i + 1 = 1 Then
                    oCon.BracketCloseNum = 2
                Else
                    oCon.BracketCloseNum = 1
                End If
            Next
            oCFL.SetConditions(oCons)

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub filterTransferChooseFromList(ByVal oForm As SAPbouiCOM.Form, ByVal strCFLID As String)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList

            oCFLs = oForm.ChooseFromLists
            oCFL = oCFLs.Item(strCFLID)

            oCons = oCFL.GetConditions()

            If oCons.Count = 0 Then
                oCon = oCons.Add()
            Else
                oCon = oCons.Item(0)
            End If

            oCon.Alias = "U_TCardCode"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = oForm.Items.Item("4").Specific.value
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub reDrawForm(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)

            oForm.Items.Item("34").Width = oForm.Width - 30
            oForm.Items.Item("34").Height = oForm.Items.Item("3").Height + 10

            oForm.Freeze(False)
        Catch ex As Exception
            'oApplication.Log.Trace_DIET_AddOn_Error(ex)
            'oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Private Sub changeControlBasedOnType(ByVal oForm As SAPbouiCOM.Form, ByVal strType As String)
        Try
            If strType = "I" Then
                'oForm.Items.Item("14").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                oForm.Items.Item("_37").Visible = False
                oForm.Items.Item("37").Visible = False

                oForm.Items.Item("_6").Visible = True
                oForm.Items.Item("6").Visible = True
            Else
                'oForm.Items.Item("14").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                oForm.Items.Item("_37").Visible = True
                oForm.Items.Item("37").Visible = True

                oForm.Items.Item("_6").Visible = False
                oForm.Items.Item("6").Visible = False
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

#End Region

End Class

Public Class clsFoodMenu
    Inherits clsBase

    Private objForm As SAPbouiCOM.Form
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim oGrid As SAPbouiCOM.Grid
    Dim strQuery As String = String.Empty
    Dim oEditTextCol As SAPbouiCOM.EditTextColumn

    Public Sub LoadForm(ByVal objParentFormID As String, ByVal strFromDate As String, ByVal strTodate As String, _
                        ByVal strCardCode As String, ByVal strType As String, ByVal strRef As String, ByVal strProgram As String, ByVal strProgramID As String)
        Try

            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_OPSL_2, frm_Z_OPSL_2)
            oForm = oApplication.SBO_Application.Forms.Item(strUID)

            CType(oForm.Items.Item("23").Specific, SAPbouiCOM.EditText).Value = objParentFormID ' Ref Form
            CType(oForm.Items.Item("22").Specific, SAPbouiCOM.EditText).Value = strFromDate ' From 
            CType(oForm.Items.Item("22_").Specific, SAPbouiCOM.EditText).Value = strTodate ' Till Date
            CType(oForm.Items.Item("31").Specific, SAPbouiCOM.EditText).Value = strCardCode ' CardCode
            CType(oForm.Items.Item("32").Specific, SAPbouiCOM.EditText).Value = strProgram 'Program Code
            CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value = strType 'I/T/P
            CType(oForm.Items.Item("34").Specific, SAPbouiCOM.EditText).Value = strRef 'Invoice/Transfer Ref/Program ID
            CType(oForm.Items.Item("36").Specific, SAPbouiCOM.EditText).Value = strProgramID 'Program ID
            CType(oForm.Items.Item("38").Specific, SAPbouiCOM.EditText).Value = System.DateTime.Now.ToString("yyyyMMddhhmmss") 'Session Instance

            addChooseFromListConditions(oForm)
            initialize(oForm, strProgram)

            'oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm.Freeze(False)
        End Try
    End Sub

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form, ByVal strProgram As String)
        Try
            oForm.PaneLevel = 7
            oForm.Items.Item("_8").TextStyle = 5
            'oForm.Items.Item("_9").TextStyle = 5
            oForm.Items.Item("_10").TextStyle = 5
            'oForm.Items.Item("_11").TextStyle = 5
            oForm.Items.Item("_12").TextStyle = 5
            'oForm.Items.Item("_13").TextStyle = 5
            oForm.Items.Item("_14").TextStyle = 5
            'oForm.Items.Item("_15").TextStyle = 5
            oForm.Items.Item("_16").TextStyle = 5
            'oForm.Items.Item("_17").TextStyle = 5
            oForm.Items.Item("_18").TextStyle = 5
            'oForm.Items.Item("_19").TextStyle = 5
            oForm.Items.Item("_24").TextStyle = 5
            oForm.Items.Item("_25").TextStyle = 5
            oForm.Items.Item("_26").TextStyle = 5
            oForm.Items.Item("_27").TextStyle = 5
            oForm.Items.Item("_28").TextStyle = 5
            oForm.Items.Item("_29").TextStyle = 5
            oForm.Items.Item("37").TextStyle = 5

            oForm.DataSources.DataTables.Add("Dt_BF_R")
            'oForm.DataSources.DataTables.Add("Dt_BF_A")
            oForm.DataSources.DataTables.Add("Dt_BF_C")
            oForm.DataSources.DataTables.Add("Dt_Lunch_R")
            'oForm.DataSources.DataTables.Add("Dt_Lunch_A")
            oForm.DataSources.DataTables.Add("Dt_Lunch_C")
            oForm.DataSources.DataTables.Add("Dt_LunchS_R")
            'oForm.DataSources.DataTables.Add("Dt_LunchS_A")
            oForm.DataSources.DataTables.Add("Dt_LunchS_C")
            oForm.DataSources.DataTables.Add("Dt_Snack_R")
            'oForm.DataSources.DataTables.Add("Dt_Snack_A")
            oForm.DataSources.DataTables.Add("Dt_Snack_C")
            oForm.DataSources.DataTables.Add("Dt_Dinner_R")
            'oForm.DataSources.DataTables.Add("Dt_Dinner_A")
            oForm.DataSources.DataTables.Add("Dt_Dinner_C")
            oForm.DataSources.DataTables.Add("Dt_DinnerS_R")
            'oForm.DataSources.DataTables.Add("Dt_DinnerS_A")
            oForm.DataSources.DataTables.Add("Dt_DinnerS_C")
            oForm.DataSources.DataTables.Add("Dt_ProgramDates")

            fillProgramDate(oForm)
            'fillMenuBasedOnDate(oForm)

            enableSession(oForm, strProgram)

            oForm.Update()
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Z_OPSL_2 Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "20" Then
                                    ' oForm.Freeze(True)
                                    If Not validationRegular(oForm, "8") Then
                                        If oApplication.SBO_Application.MessageBox("Cannot Select More the One Regular Food...Continue...?", , "Yes", "No") = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    ElseIf Not validationRegular(oForm, "10") Then
                                        If oApplication.SBO_Application.MessageBox("Cannot Select More the One Regular Food...Continue...?", , "Yes", "No") = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    ElseIf Not validationRegular(oForm, "12") Then
                                        If oApplication.SBO_Application.MessageBox("Cannot Select More the One Regular Food...Continue...?", , "Yes", "No") = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    ElseIf Not validationRegular(oForm, "14") Then
                                        If oApplication.SBO_Application.MessageBox("Cannot Select More the One Regular Food...Continue...?", , "Yes", "No") = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    ElseIf Not validationRegular(oForm, "16") Then
                                        If oApplication.SBO_Application.MessageBox("Cannot Select More the One Regular Food...Continue...?", , "Yes", "No") = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    ElseIf Not validationRegular(oForm, "18") Then
                                        If oApplication.SBO_Application.MessageBox("Cannot Select More the One Regular Food...Continue...?", , "Yes", "No") = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If

                                    If Not validationSessions(oForm) Then
                                        'If oApplication.SBO_Application.MessageBox("Food Not Selected for Some of the Sessions...Continue...?", , "Yes", "No") = 2 Then
                                        If 1 = 2 Then
                                            oForm.Freeze(False)
                                            BubbleEvent = False
                                            Exit Sub
                                        Else
                                            oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "8", "BF", "R") 'Break Fast - Regular
                                            oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "24", "BF", "C") 'Break Fast Custom
                                            oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "10", "LN", "R") 'Lunch- Regular
                                            oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "25", "LN", "C") 'Lunch Custom
                                            oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "12", "LS", "R") 'Lunch Side- Regular
                                            oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "26", "LS", "C") 'Lunch Side Custom
                                            oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "14", "SK", "R") 'Snack - Regular
                                            oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "27", "SK", "C") 'Snack Custom
                                            oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "16", "DI", "R") 'Dinner- Regular
                                            oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "28", "DI", "C") 'Dinner Custom
                                            oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "18", "DS", "R") 'Dinner Side- Regular
                                            oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "29", "DS", "C") 'Dinner Side Custom
                                            oApplication.SBO_Application.MessageBox("Food Saved Successfully...")
                                        End If
                                    Else
                                        oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "8", "BF", "R") 'Break Fast - Regular
                                        oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "24", "BF", "C") 'Break Fast Custom
                                        oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "10", "LN", "R") 'Lunch- Regular
                                        oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "25", "LN", "C") 'Lunch Custom
                                        oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "12", "LS", "R") 'Lunch Side- Regular
                                        oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "26", "LS", "C") 'Lunch Side Custom
                                        oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "14", "SK", "R") 'Snack - Regular
                                        oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "27", "SK", "C") 'Snack Custom
                                        oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "16", "DI", "R") 'Dinner- Regular
                                        oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "28", "DI", "C") 'Dinner Custom
                                        oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "18", "DS", "R") 'Dinner Side- Regular
                                        oApplication.Utilities.UpdateCustomerFoodMenu(oForm, "29", "DS", "C") 'Dinner Side Custom
                                        oApplication.SBO_Application.MessageBox("Food Saved Successfully...")
                                    End If
                                    'oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "20_" Then
                                    If Not validation(oForm) Then

                                        'Newly Added for Not adding All Session for Date Selected. '20150510
                                        If Not validationSessions(oForm) Then
                                            If oApplication.SBO_Application.MessageBox("Food Not Selected for Some of the Sessions...Continue...?", , "Yes", "No") = 2 Then
                                                oForm.Freeze(False)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If

                                        If oApplication.SBO_Application.MessageBox("Food Not Selected for Some of the dates...Continue...?", , "Yes", "No") = 2 Then
                                            oForm.Freeze(False)
                                            BubbleEvent = False
                                            Exit Sub
                                        Else
                                            Dim oBaseForm As SAPbouiCOM.Form = Nothing
                                            For index As Integer = 0 To oApplication.SBO_Application.Forms.Count
                                                If oApplication.SBO_Application.Forms.Item(index).UniqueID = CType(oForm.Items.Item("23").Specific, SAPbouiCOM.EditText).Value Then
                                                    oBaseForm = oApplication.SBO_Application.Forms.Item(index)
                                                    Exit For
                                                End If
                                            Next
                                            If Not IsNothing(oBaseForm) Then
                                                Dim strQuery As String = String.Empty
                                                Dim oRecordSet As Recordset
                                                oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                                Dim oBDBDataSource As SAPbouiCOM.DBDataSource
                                                Dim oBaseMatrix As SAPbouiCOM.Matrix
                                                oBDBDataSource = oBaseForm.DataSources.DBDataSources.Item("@Z_PSL1")
                                                oBaseMatrix = oBaseForm.Items.Item("3").Specific
                                                oBaseMatrix.Clear()
                                                oBaseMatrix.FlushToDataSource()
                                                Dim intMatrixRow As Integer = 0

                                                Dim strCardCode As String = CType(oForm.Items.Item("31").Specific, SAPbouiCOM.EditText).Value
                                                Dim strProgramID As String = CType(oForm.Items.Item("36").Specific, SAPbouiCOM.EditText).Value
                                                Dim strProFromDate As String = CType(oForm.Items.Item("22").Specific, SAPbouiCOM.EditText).Value
                                                Dim strProToDate As String = CType(oForm.Items.Item("22_").Specific, SAPbouiCOM.EditText).Value
                                                Dim strSession As String = CType(oForm.Items.Item("38").Specific, SAPbouiCOM.EditText).Value

                                                strQuery = "Select T0.U_ItemCode,T1.ItemName,Convert(VarChar(8),U_PrgDate,112) As U_PrgDate,T0.U_Quantity,T0.U_Dislike,T0.U_Medical,T0.U_FType,T0.U_SFood,T0.U_Remarks "
                                                strQuery += " ,(Case WHEN U_FType = 'BF' THEN '1' WHEN U_FType = 'LN' THEN '2' WHEN U_FType = 'LS' THEN '3' WHEN U_FType = 'SK' THEN '4' WHEN U_FType = 'DI' THEN '5' WHEN U_FType = 'DS' THEN '6' END) As 'SR'  "
                                                strQuery += " From [@Z_OFSL] T0 "
                                                strQuery += " JOIN OITM T1 On T0.U_ItemCode = T1.ItemCode "
                                                strQuery += " Where T0.U_ProgramID = '" + strProgramID + "'"
                                                strQuery += " And T0.U_Session = '" + strSession + "'"
                                                strQuery += " And Convert(VarChar(8),T0.U_PrgDate,112) >= '" + strProFromDate + "'"
                                                strQuery += " And Convert(VarChar(8),T0.U_PrgDate,112) <= '" + strProToDate + "'"
                                                strQuery += " And T0.U_Select = 'Y'"
                                                strQuery += " Order By T0.U_PrgDate "
                                                strQuery += " ,(Case WHEN U_FType = 'BF' THEN '1' WHEN U_FType = 'LN' THEN '2' WHEN U_FType = 'LS' THEN '3' WHEN U_FType = 'SK' THEN '4' WHEN U_FType = 'DI' THEN '5' WHEN U_FType = 'DS' THEN '6' END)  "
                                                oRecordSet.DoQuery(strQuery)
                                                If Not oRecordSet.EoF Then
                                                    While Not oRecordSet.EoF
                                                        oBaseMatrix.AddRow(1, oBaseMatrix.RowCount)
                                                        oBaseMatrix.FlushToDataSource()
                                                        oBDBDataSource.SetValue("LineId", intMatrixRow, (intMatrixRow + 1).ToString())

                                                        oBDBDataSource.SetValue("U_DelDate", intMatrixRow, oRecordSet.Fields.Item("U_PrgDate").Value.ToString())
                                                        oBDBDataSource.SetValue("U_FType", intMatrixRow, oRecordSet.Fields.Item("U_FType").Value.ToString())
                                                        oBDBDataSource.SetValue("U_ItemCode", intMatrixRow, oRecordSet.Fields.Item("U_ItemCode").Value.ToString())
                                                        oBDBDataSource.SetValue("U_ItemName", intMatrixRow, oRecordSet.Fields.Item("ItemName").Value.ToString())
                                                        oBDBDataSource.SetValue("U_Quantity", intMatrixRow, CDbl(oRecordSet.Fields.Item("U_Quantity").Value.ToString()))
                                                        oBDBDataSource.SetValue("U_UnitPrice", intMatrixRow, "0")
                                                        oBDBDataSource.SetValue("U_SFood", intMatrixRow, oRecordSet.Fields.Item("U_SFood").Value.ToString())
                                                        oBDBDataSource.SetValue("U_Remarks", intMatrixRow, oRecordSet.Fields.Item("U_Remarks").Value.ToString())
                                                        oBDBDataSource.SetValue("U_Dislike", intMatrixRow, oRecordSet.Fields.Item("U_Dislike").Value.ToString())
                                                        oBDBDataSource.SetValue("U_Medical", intMatrixRow, oRecordSet.Fields.Item("U_Medical").Value.ToString())
                                                        intMatrixRow += 1
                                                        oBaseMatrix.LoadFromDataSource()
                                                        oBaseMatrix.FlushToDataSource()

                                                        oRecordSet.MoveNext()
                                                    End While
                                                End If
                                                oBaseMatrix.LoadFromDataSource()
                                                oBaseMatrix.FlushToDataSource()
                                                If oBaseForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oBaseForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                oBaseForm.Update()
                                                oForm.Close()
                                                Exit Sub
                                            End If
                                        End If
                                    Else

                                        'Newly Added for Not adding All Session for Date Selected. '20150510
                                        If Not validationSessions(oForm) Then
                                            If oApplication.SBO_Application.MessageBox("Food Not Selected for Some of the Sessions...Continue...?", , "Yes", "No") = 2 Then
                                                oForm.Freeze(False)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If

                                        If oApplication.SBO_Application.MessageBox("Sure Want to Load All Selected to Food to PreSales...Continue...?", , "Yes", "No") = 2 Then
                                            oForm.Freeze(False)
                                            BubbleEvent = False
                                            Exit Sub
                                        Else
                                            Dim oBaseForm As SAPbouiCOM.Form = Nothing
                                            For index As Integer = 0 To oApplication.SBO_Application.Forms.Count
                                                If oApplication.SBO_Application.Forms.Item(index).UniqueID = CType(oForm.Items.Item("23").Specific, SAPbouiCOM.EditText).Value Then
                                                    oBaseForm = oApplication.SBO_Application.Forms.Item(index)
                                                    Exit For
                                                End If
                                            Next
                                            If Not IsNothing(oBaseForm) Then
                                                Dim strQuery As String = String.Empty
                                                Dim oRecordSet As Recordset
                                                oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                                Dim oBDBDataSource As SAPbouiCOM.DBDataSource
                                                Dim oBaseMatrix As SAPbouiCOM.Matrix
                                                oBDBDataSource = oBaseForm.DataSources.DBDataSources.Item("@Z_PSL1")
                                                oBaseMatrix = oBaseForm.Items.Item("3").Specific
                                                oBaseMatrix.Clear()
                                                oBaseMatrix.FlushToDataSource()
                                                Dim intMatrixRow As Integer = 0

                                                Dim strCardCode As String = CType(oForm.Items.Item("31").Specific, SAPbouiCOM.EditText).Value
                                                Dim strProgramID As String = CType(oForm.Items.Item("36").Specific, SAPbouiCOM.EditText).Value
                                                Dim strProFromDate As String = CType(oForm.Items.Item("22").Specific, SAPbouiCOM.EditText).Value
                                                Dim strProToDate As String = CType(oForm.Items.Item("22_").Specific, SAPbouiCOM.EditText).Value
                                                Dim strSession As String = CType(oForm.Items.Item("38").Specific, SAPbouiCOM.EditText).Value

                                                strQuery = "Select T0.U_ItemCode,T1.ItemName,Convert(VarChar(8),U_PrgDate,112) As U_PrgDate,T0.U_Quantity,T0.U_Dislike,T0.U_Medical,T0.U_FType,T0.U_SFood,T0.U_Remarks "
                                                strQuery += " ,(Case WHEN U_FType = 'BF' THEN '1' WHEN U_FType = 'LN' THEN '2' WHEN U_FType = 'LS' THEN '3' WHEN U_FType = 'SK' THEN '4' WHEN U_FType = 'DI' THEN '5' WHEN U_FType = 'DS' THEN '6' END) As 'SR'  "
                                                strQuery += " From [@Z_OFSL] T0 "
                                                strQuery += " JOIN OITM T1 On T0.U_ItemCode = T1.ItemCode "
                                                strQuery += " Where T0.U_ProgramID = '" + strProgramID + "'"
                                                strQuery += " And T0.U_Session = '" + strSession + "'"
                                                strQuery += " And Convert(VarChar(8),T0.U_PrgDate,112) >= '" + strProFromDate + "'"
                                                strQuery += " And Convert(VarChar(8),T0.U_PrgDate,112) <= '" + strProToDate + "'"
                                                strQuery += " AND T0.U_Select = 'Y'"
                                                strQuery += " Order By T0.U_PrgDate "
                                                strQuery += " ,(Case WHEN U_FType = 'BF' THEN '1' WHEN U_FType = 'LN' THEN '2' WHEN U_FType = 'LS' THEN '3' WHEN U_FType = 'SK' THEN '4' WHEN U_FType = 'DI' THEN '5' WHEN U_FType = 'DS' THEN '6' END)  "
                                                oRecordSet.DoQuery(strQuery)
                                                If Not oRecordSet.EoF Then
                                                    While Not oRecordSet.EoF
                                                        oBaseMatrix.AddRow(1, oBaseMatrix.RowCount)
                                                        oBaseMatrix.FlushToDataSource()
                                                        oBDBDataSource.SetValue("LineId", intMatrixRow, (intMatrixRow + 1).ToString())

                                                        oBDBDataSource.SetValue("U_DelDate", intMatrixRow, oRecordSet.Fields.Item("U_PrgDate").Value.ToString())
                                                        oBDBDataSource.SetValue("U_FType", intMatrixRow, oRecordSet.Fields.Item("U_FType").Value.ToString())
                                                        oBDBDataSource.SetValue("U_ItemCode", intMatrixRow, oRecordSet.Fields.Item("U_ItemCode").Value.ToString())
                                                        oBDBDataSource.SetValue("U_ItemName", intMatrixRow, oRecordSet.Fields.Item("ItemName").Value.ToString())
                                                        oBDBDataSource.SetValue("U_Quantity", intMatrixRow, CDbl(oRecordSet.Fields.Item("U_Quantity").Value.ToString()))
                                                        oBDBDataSource.SetValue("U_UnitPrice", intMatrixRow, "0")
                                                        oBDBDataSource.SetValue("U_SFood", intMatrixRow, oRecordSet.Fields.Item("U_SFood").Value.ToString())
                                                        oBDBDataSource.SetValue("U_Remarks", intMatrixRow, oRecordSet.Fields.Item("U_Remarks").Value.ToString())
                                                        oBDBDataSource.SetValue("U_Dislike", intMatrixRow, oRecordSet.Fields.Item("U_Dislike").Value.ToString())
                                                        oBDBDataSource.SetValue("U_Medical", intMatrixRow, oRecordSet.Fields.Item("U_Medical").Value.ToString())
                                                        intMatrixRow += 1
                                                        oBaseMatrix.LoadFromDataSource()
                                                        oBaseMatrix.FlushToDataSource()

                                                        oRecordSet.MoveNext()
                                                    End While
                                                End If
                                                oBaseMatrix.LoadFromDataSource()
                                                oBaseMatrix.FlushToDataSource()
                                                If oBaseForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oBaseForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                oBaseForm.Update()
                                                oForm.Close()
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                ElseIf pVal.ItemUID = "21" Then
                                    oForm.Close()
                                    'ElseIf pVal.ItemUID = "33" And pVal.Row > -1 Then
                                    '    Dim strFrmDate As String = getSelectedPrgDate(oForm, pVal.ItemUID)
                                    '    CType(oForm.Items.Item("35").Specific, SAPbouiCOM.EditText).Value = strFrmDate
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                If (pVal.ItemUID = "1" Or pVal.ItemUID = "2" Or pVal.ItemUID = "3" _
                                         Or pVal.ItemUID = "4" Or pVal.ItemUID = "5" Or pVal.ItemUID = "6") Then
                                    oForm.Freeze(True)
                                    If oForm.Items.Item(pVal.ItemUID).Enabled Then
                                        changePane(oForm, pVal.ItemUID)
                                    End If
                                    oForm.Freeze(False)
                                    'oForm.Items.Item("30").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                ElseIf pVal.ColUID = "Select" And pVal.Row > -1 Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    Dim strSelected As String = oGrid.DataTable.GetValue("Select", pVal.Row)
                                    If strSelected = "Y" Then
                                        '   oGrid.CommonSetting.SetRowFontColor(pVal.Row + 1, RGB(0, 255, 0))
                                        oGrid.CommonSetting.SetRowBackColor(pVal.Row + 1, RGB(0, 255, 0))
                                    Else
                                        'oGrid.CommonSetting.SetRowFontColor(pVal.Row + 1, RGB(0, 0, 0))
                                        oGrid.CommonSetting.SetRowBackColor(pVal.Row + 1, RGB(255, 255, 255))
                                    End If
                                ElseIf pVal.ItemUID = "33" And pVal.ColUID = "RowsHeader" And pVal.Row > -1 Then

                                    Dim strFrmDate As String = getSelectedPrgDate(oForm, pVal.ItemUID)
                                    If strFrmDate <> "00010101" Then
                                        '  oForm.PaneLevel = 7
                                        CType(oForm.Items.Item("35").Specific, SAPbouiCOM.EditText).Value = strFrmDate
                                        CType(oForm.Items.Item("_8").Specific, SAPbouiCOM.StaticText).Caption = "Loading Menu details...."
                                        oForm.Freeze(True)
                                        oApplication.Utilities.Message("Please wait Food Definition is loading for selected date....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        fillMenuBasedOnDate(oForm)
                                        oApplication.Utilities.Message("Please Continue to select Food Definition  for selected date....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        'oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        'oForm.PaneLevel = 1
                                        Dim strID As String = enableSession(oForm, CType(oForm.Items.Item("32").Specific, SAPbouiCOM.EditText).Value)
                                        changePane(oForm, IIf(strID = "", "1", strID))
                                        oForm.Refresh()
                                        oForm.Freeze(False)
                                        CType(oForm.Items.Item("_8").Specific, SAPbouiCOM.StaticText).Caption = "Regular"
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                    Dim oDataTable As SAPbouiCOM.DataTable
                                    Try
                                        oCFLEvento = pVal
                                        oDataTable = oCFLEvento.SelectedObjects
                                        If (pVal.ItemUID = "24" Or pVal.ItemUID = "25" Or pVal.ItemUID = "26" _
                                            Or pVal.ItemUID = "27" Or pVal.ItemUID = "28" Or pVal.ItemUID = "29") _
                                            And pVal.ColUID = "U_ItemName" And pVal.Row > -1 Then
                                            oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                            If Not IsNothing(oDataTable) Then
                                                Dim intAddRows As Integer = oDataTable.Rows.Count
                                                If intAddRows > 0 Then
                                                    oGrid.DataTable.SetValue("U_ItemCode", pVal.Row, oDataTable.GetValue("ItemCode", 0).ToString())
                                                    oGrid.DataTable.SetValue("U_ItemName", pVal.Row, oDataTable.GetValue("ItemName", 0).ToString())
                                                    'Madhu Commented for Phase II On 20150710'
                                                    Dim dblQty As Double
                                                    dblQty = getQuantityBasedonCaloriesRatio(oForm, pVal.ItemUID)
                                                    If dblQty = 0 Then
                                                        dblQty = getQuantityBasedonCalories(oForm, pVal.ItemUID)
                                                    End If
                                                    oGrid.DataTable.SetValue("Qty", pVal.Row, dblQty.ToString())
                                                    oGrid.DataTable.SetValue("Select", pVal.Row, "Y")

                                                    Dim strDisLike As String = String.Empty
                                                    Dim strMedical As String = String.Empty
                                                    Dim strCardCode As String = CType(oForm.Items.Item("31").Specific, SAPbouiCOM.EditText).Value
                                                    Dim strItemCode As String = oDataTable.GetValue("ItemCode", 0).ToString()

                                                    If (oApplication.Utilities.hasBOM(strItemCode)) Then
                                                        strDisLike = oApplication.Utilities.GetDisLikeItem(strCardCode, strItemCode)
                                                        strMedical = oApplication.Utilities.GetMedicalItem(strCardCode, strItemCode)
                                                        oApplication.Utilities.get_ChildItems(strCardCode, strItemCode, strDisLike, strMedical)
                                                    Else
                                                        strDisLike = oApplication.Utilities.GetDisLikeItem(strCardCode, strItemCode)
                                                        strMedical = oApplication.Utilities.GetMedicalItem(strCardCode, strItemCode)
                                                    End If

                                                    If strDisLike.Trim().Length > 0 Then
                                                        oGrid.DataTable.SetValue("U_Dislike", pVal.Row, strDisLike)
                                                    End If
                                                    If strMedical.Trim().Length > 0 Then
                                                        oGrid.DataTable.SetValue("U_Medical", pVal.Row, strMedical)
                                                    End If

                                                    If pVal.Row = oGrid.DataTable.Rows.Count - 1 Then
                                                        oGrid.DataTable.Rows.Add(1)
                                                        fillHeader(oForm, pVal.ItemUID)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Catch ex As Exception
                                        oApplication.Log.Trace_DIET_AddOn_Error(ex)
                                        Throw ex
                                        'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
                                    End Try
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                If oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Or oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    Try
                                        reDrawForm(oForm)
                                    Catch ex As Exception
                                        'oApplication.Log.Trace_DIET_AddOn_Error(ex)
                                    End Try
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Private Sub reDrawForm(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)

            oForm.Items.Item("7").Width = oForm.Width - 180
            oForm.Items.Item("7").Height = oForm.Height - 110

            'Regular Lab
            oForm.Items.Item("_8").Top = oForm.Items.Item("7").Top + 5
            oForm.Items.Item("_10").Top = oForm.Items.Item("_8").Top
            oForm.Items.Item("_12").Top = oForm.Items.Item("_8").Top
            oForm.Items.Item("_14").Top = oForm.Items.Item("_8").Top
            oForm.Items.Item("_16").Top = oForm.Items.Item("_8").Top
            oForm.Items.Item("_18").Top = oForm.Items.Item("_8").Top

            'Regular BF Grid - Top
            oForm.Items.Item("8").Top = oForm.Items.Item("_8").Top + oForm.Items.Item("_8").Height + 5
            oForm.Items.Item("10").Top = oForm.Items.Item("8").Top
            oForm.Items.Item("12").Top = oForm.Items.Item("8").Top
            oForm.Items.Item("14").Top = oForm.Items.Item("8").Top
            oForm.Items.Item("16").Top = oForm.Items.Item("8").Top
            oForm.Items.Item("18").Top = oForm.Items.Item("8").Top

            'Regular BF Grid - Top
            oForm.Items.Item("8").Width = oForm.Items.Item("7").Width - 30
            oForm.Items.Item("10").Width = oForm.Items.Item("8").Width
            oForm.Items.Item("12").Width = oForm.Items.Item("8").Width
            oForm.Items.Item("14").Width = oForm.Items.Item("8").Width
            oForm.Items.Item("16").Width = oForm.Items.Item("8").Width
            oForm.Items.Item("18").Width = oForm.Items.Item("8").Width

            'Regular BF Grid - Height
            oForm.Items.Item("8").Height = (oForm.Items.Item("7").Height / 2) - 40
            oForm.Items.Item("10").Height = oForm.Items.Item("8").Height
            oForm.Items.Item("12").Height = oForm.Items.Item("8").Height
            oForm.Items.Item("14").Height = oForm.Items.Item("8").Height
            oForm.Items.Item("16").Height = oForm.Items.Item("8").Height
            oForm.Items.Item("18").Height = oForm.Items.Item("8").Height

            ''Alternative Lab
            'oForm.Items.Item("_9").Top = oForm.Items.Item("8").Top + oForm.Items.Item("8").Height + 5
            'oForm.Items.Item("_11").Top = oForm.Items.Item("_9").Top
            'oForm.Items.Item("_13").Top = oForm.Items.Item("_9").Top
            'oForm.Items.Item("_15").Top = oForm.Items.Item("_9").Top
            'oForm.Items.Item("_17").Top = oForm.Items.Item("_9").Top
            'oForm.Items.Item("_19").Top = oForm.Items.Item("_9").Top

            ''Regular BF Grid - Top
            'oForm.Items.Item("9").Top = oForm.Items.Item("_9").Top + oForm.Items.Item("_9").Height + 5
            'oForm.Items.Item("11").Top = oForm.Items.Item("9").Top
            'oForm.Items.Item("13").Top = oForm.Items.Item("9").Top
            'oForm.Items.Item("15").Top = oForm.Items.Item("9").Top
            'oForm.Items.Item("17").Top = oForm.Items.Item("9").Top
            'oForm.Items.Item("19").Top = oForm.Items.Item("9").Top

            ''Regular BF Grid - Width
            'oForm.Items.Item("9").Width = oForm.Items.Item("8").Width
            'oForm.Items.Item("11").Width = oForm.Items.Item("8").Width
            'oForm.Items.Item("13").Width = oForm.Items.Item("8").Width
            'oForm.Items.Item("15").Width = oForm.Items.Item("8").Width
            'oForm.Items.Item("17").Width = oForm.Items.Item("8").Width
            'oForm.Items.Item("19").Width = oForm.Items.Item("8").Width

            ''Regular BF Grid - Height
            'oForm.Items.Item("9").Height = oForm.Items.Item("8").Height
            'oForm.Items.Item("11").Height = oForm.Items.Item("8").Height
            'oForm.Items.Item("13").Height = oForm.Items.Item("8").Height
            'oForm.Items.Item("15").Height = oForm.Items.Item("8").Height
            'oForm.Items.Item("17").Height = oForm.Items.Item("8").Height
            'oForm.Items.Item("19").Height = oForm.Items.Item("8").Height


            'Custom Lab
            oForm.Items.Item("_24").Top = oForm.Items.Item("8").Top + oForm.Items.Item("8").Height + 5
            oForm.Items.Item("_25").Top = oForm.Items.Item("_24").Top
            oForm.Items.Item("_26").Top = oForm.Items.Item("_24").Top
            oForm.Items.Item("_27").Top = oForm.Items.Item("_24").Top
            oForm.Items.Item("_28").Top = oForm.Items.Item("_24").Top
            oForm.Items.Item("_29").Top = oForm.Items.Item("_24").Top

            'Regular BF Grid - Top
            oForm.Items.Item("24").Top = oForm.Items.Item("_24").Top + oForm.Items.Item("_24").Height + 5
            oForm.Items.Item("25").Top = oForm.Items.Item("24").Top
            oForm.Items.Item("26").Top = oForm.Items.Item("24").Top
            oForm.Items.Item("27").Top = oForm.Items.Item("24").Top
            oForm.Items.Item("28").Top = oForm.Items.Item("24").Top
            oForm.Items.Item("29").Top = oForm.Items.Item("24").Top

            'Regular BF Grid - Width
            oForm.Items.Item("24").Width = oForm.Items.Item("8").Width
            oForm.Items.Item("25").Width = oForm.Items.Item("8").Width
            oForm.Items.Item("26").Width = oForm.Items.Item("8").Width
            oForm.Items.Item("27").Width = oForm.Items.Item("8").Width
            oForm.Items.Item("28").Width = oForm.Items.Item("8").Width
            oForm.Items.Item("29").Width = oForm.Items.Item("8").Width

            'Regular BF Grid - Height
            oForm.Items.Item("24").Height = oForm.Items.Item("8").Height
            oForm.Items.Item("25").Height = oForm.Items.Item("8").Height
            oForm.Items.Item("26").Height = oForm.Items.Item("8").Height
            oForm.Items.Item("27").Height = oForm.Items.Item("8").Height
            oForm.Items.Item("28").Height = oForm.Items.Item("8").Height
            oForm.Items.Item("29").Height = oForm.Items.Item("8").Height

            oForm.Items.Item("33").Height = oForm.Height - 70
            oForm.Items.Item("33").Top = 5
            oForm.Items.Item("33").Width = 110
            oGrid = oForm.Items.Item("33").Specific
            oGrid.RowHeaders.Width = 20

            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            'oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    'Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
    '    Dim _retVal As Boolean = True
    '    Try
    '        Dim strBF As String = String.Empty
    '        Dim dblBFQty As Double

    '        getFoodValue(oForm, "8", strBF, dblBFQty)
    '        If strBF.Trim().Length = 0 Then
    '            'getFoodValue(oForm, "9", strBF, dblBFQty)
    '            If strBF.Trim().Length = 0 Then
    '                getFoodValue(oForm, "24", strBF, dblBFQty)
    '            End If
    '        End If

    '        Dim strLunch As String = String.Empty
    '        Dim dblLQty As Double
    '        getFoodValue(oForm, "10", strLunch, dblLQty)
    '        If strLunch.Trim().Length = 0 Then
    '            'getFoodValue(oForm, "11", strLunch, dblLQty)
    '            If strLunch.Trim().Length = 0 Then
    '                getFoodValue(oForm, "25", strLunch, dblLQty)
    '            End If
    '        End If

    '        Dim strLunchS As String = String.Empty
    '        Dim dblLSQty As Double
    '        getFoodValue(oForm, "12", strLunchS, dblLSQty)
    '        If strLunchS.Trim().Length = 0 Then
    '            'getFoodValue(oForm, "13", strLunchS, dblLSQty)
    '            If strLunchS.Trim().Length = 0 Then
    '                getFoodValue(oForm, "26", strLunchS, dblLSQty)
    '            End If
    '        End If

    '        Dim strSnack As String = String.Empty
    '        Dim dblSnack As Double
    '        getFoodValue(oForm, "14", strSnack, dblSnack)
    '        If strSnack.Trim().Length = 0 Then
    '            'getFoodValue(oForm, "15", strSnack, dblSnack)
    '            If strSnack.Trim().Length = 0 Then
    '                getFoodValue(oForm, "27", strSnack, dblSnack)
    '            End If
    '        End If

    '        Dim strDinner As String = String.Empty
    '        Dim dblDinner As Double
    '        getFoodValue(oForm, "16", strDinner, dblDinner)
    '        If strDinner.Trim().Length = 0 Then
    '            'getFoodValue(oForm, "17", strDinner, dblDinner)
    '            If strDinner.Trim().Length = 0 Then
    '                getFoodValue(oForm, "28", strDinner, dblDinner)
    '            End If
    '        End If

    '        Dim strDinnerS As String = String.Empty
    '        Dim dblDinnerS As Double
    '        getFoodValue(oForm, "18", strDinnerS, dblDinnerS)
    '        If strDinnerS.Trim().Length = 0 Then
    '            'getFoodValue(oForm, "19", strDinnerS, dblDinnerS)
    '            If strDinnerS.Trim().Length = 0 Then
    '                getFoodValue(oForm, "29", strDinnerS, dblDinnerS)
    '            End If
    '        End If

    '        If strBF.Trim.Length = 0 Then
    '            oApplication.Utilities.Message("Select Break Fast To Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '            Return False
    '        ElseIf strLunch.Trim.Length = 0 Then
    '            oApplication.Utilities.Message("Select Lunch To Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '            Return False
    '        ElseIf strLunchS.Trim.Length = 0 Then
    '            oApplication.Utilities.Message("Select Lunch(Side) To Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '            Return False
    '        ElseIf strSnack.Trim.Length = 0 Then
    '            oApplication.Utilities.Message("Select Snack To Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '            Return False
    '        ElseIf strDinner.Trim.Length = 0 Then
    '            oApplication.Utilities.Message("Select Dinner To Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '            Return False
    '        ElseIf strDinnerS.Trim.Length = 0 Then
    '            oApplication.Utilities.Message("Select Dinner(Side) To Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '            Return False
    '        End If

    '        Return _retVal
    '    Catch ex As Exception 
    'oApplication.Log.Trace_DIET_AddOn_Error(ex)

    '    End Try
    'End Function

    Private Sub fillMenuBasedOnDate(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oDataTable As SAPbouiCOM.DataTable
            Dim strPrgDate As String = oForm.Items.Item("35").Specific.value.ToString().Trim()
            Dim strCardCode As String = oForm.Items.Item("31").Specific.value.ToString().Trim()
            Dim strProgram As String = oForm.Items.Item("32").Specific.value.ToString().Trim()
            Dim strProgramID As String = oForm.Items.Item("36").Specific.value.ToString().Trim()
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)


            'Break Fast - Regular
            oGrid = oForm.Items.Item("8").Specific
            oDataTable = oForm.DataSources.DataTables.Item("Dt_BF_R")
            Dim dblQty As Double
            dblQty = getQuantityBasedonCaloriesRatio(oForm, "8")
            strQuery = " Select (Select Case When T0.U_ItemCode = T2.U_ItemCode Then 'Y' ELSE 'N' END) As 'Select',"
            strQuery += " T0.U_ItemCode,U_ItemName, "
            strQuery += " (Select ISNULL( "
            strQuery += " (ISNULL(T2.U_Quantity, "
            strQuery += " (Select ISNULL(" & dblQty & ",U_BFactor) From [@Z_OCAJ] T0 "
            strQuery += " JOIN [@Z_OCPR] T1 On T0.U_Calories "
            strQuery += " = T1.U_CPAdj Where T1.U_CardCode = '" + strCardCode + "')) "
            strQuery += " ),1)) As 'Qty', "
            strQuery += " T2.U_Dislike,T2.U_Medical, "
            strQuery += " T2.U_Remarks As 'Remarks' "
            strQuery += " From [@Z_MED1] T0 JOIN [@Z_OMED] T1 On T1.DocEntry = T0.DocEntry  "

            strQuery += " And T1.U_CatType = 'I' "

            strQuery += " LEFT OUTER JOIN [@Z_OFSL] T2 ON T2.U_ItemCode = T0.U_ItemCode  "
            strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = Convert(VarChar(8),T2.U_PrgDate,112)  "
            strQuery += " And T2.U_CardCode = '" + strCardCode + "' "
            strQuery += " And T2.U_FType = 'BF' "
            strQuery += " And T2.U_SFood = 'R'  "
            strQuery += " And T2.U_Select = 'Y' "
            strQuery += " And T2.U_ProgramID = '" + strProgramID + "'"
            strQuery += " Where T1.U_MenuType = 'R' "
            strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = '" + strPrgDate + "'"
            strQuery += " And T1.U_PrgCode = '" + strProgram + "'"
            strQuery += " And T0.U_ItemCode Is Not Null "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oDataTable.ExecuteQuery(strQuery)
                oGrid.DataTable = oDataTable
            Else
                strQuery = " Select (Select Case When T0.U_ItemCode = T2.U_ItemCode Then 'Y' ELSE 'N' END) As 'Select',"
                strQuery += " T0.U_ItemCode,U_ItemName, "
                strQuery += " (Select ISNULL( "
                strQuery += " (ISNULL(T2.U_Quantity, "
                strQuery += " (Select ISNULL(" & dblQty & ",U_BFactor) From [@Z_OCAJ] T0 "
                strQuery += " JOIN [@Z_OCPR] T1 On T0.U_Calories "
                strQuery += " = T1.U_CPAdj Where T1.U_CardCode = '" + strCardCode + "')) "
                strQuery += " ),1)) As 'Qty', "
                strQuery += " T2.U_Dislike,T2.U_Medical, "
                strQuery += " T2.U_Remarks As 'Remarks' "
                strQuery += " From [@Z_MED1] T0 JOIN [@Z_OMED] T1 On T1.DocEntry = T0.DocEntry  "

                strQuery += " And T1.U_CatType = 'G' "

                strQuery += " LEFT OUTER JOIN [@Z_OFSL] T2 ON T2.U_ItemCode = T0.U_ItemCode  "
                strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = Convert(VarChar(8),T2.U_PrgDate,112)  "
                strQuery += " And T2.U_CardCode = '" + strCardCode + "' "
                strQuery += " And T2.U_FType = 'BF' "
                strQuery += " And T2.U_SFood = 'R'  "
                strQuery += " And T2.U_Select = 'Y' "
                strQuery += " And T2.U_ProgramID = '" + strProgramID + "'"

                strQuery += " JOIN OITB T3 On T3.ItmsGrpCod = T1.U_GrpCode "
                strQuery += " JOIN OITM T4 On T4.ItmsGrpCod = T3.ItmsGrpCod And T4.ItmsGrpCod = T1.U_GrpCode "

                strQuery += " Where T1.U_MenuType = 'R' "
                strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = '" + strPrgDate + "'"
                strQuery += " And T4.ItemCode = '" + strProgram + "'"
                strQuery += " And T0.U_ItemCode Is Not Null "
                oDataTable.ExecuteQuery(strQuery)
                oGrid.DataTable = oDataTable
            End If


            'Break Fast - Custom
            oGrid = oForm.Items.Item("24").Specific
            oDataTable = oForm.DataSources.DataTables.Item("Dt_BF_C")
            strQuery = " Select (Select Case When T2.U_ItemCode = T1.ItemCode Then 'Y' ELSE 'N' END) As 'Select', "
            strQuery += " T2.U_ItemCode,T1.ItemName As U_ItemName, "
            strQuery += " T2.U_Quantity As 'Qty', "
            strQuery += " T2.U_Dislike,T2.U_Medical, "
            strQuery += " T2.U_Remarks As 'Remarks' "
            strQuery += " From  [@Z_OFSL] T2 JOIN OITM T1  ON T2.U_ItemCode = T1.ItemCode  "
            strQuery += " And T2.U_CardCode = '" + strCardCode + "' "
            strQuery += " And T2.U_FType = 'BF' "
            strQuery += " And T2.U_SFood = 'C'  "
            strQuery += " And T2.U_Select = 'Y' "
            'strQuery += " And T2.U_PrgCode = '" + strProgram + "'"
            strQuery += " And T2.U_ProgramID = '" + strProgramID + "'"
            strQuery += " And Convert(VarChar(8),T2.U_PrgDate,112) = '" + strPrgDate + "'"
            oDataTable.ExecuteQuery(strQuery)
            oGrid.DataTable = oDataTable

            formatGrid(oForm, "8")
            formatGrid_Custom(oForm, "24", "CFL_1")

            'Lunch - Regular
            oGrid = oForm.Items.Item("10").Specific
            dblQty = getQuantityBasedonCaloriesRatio(oForm, "10")
            oDataTable = oForm.DataSources.DataTables.Item("Dt_Lunch_R")
            strQuery = " Select (Select Case When T0.U_ItemCode = T2.U_ItemCode Then 'Y' ELSE 'N' END) As 'Select', "
            strQuery += " T0.U_ItemCode,U_ItemName, "
            strQuery += " (Select ISNULL( "
            strQuery += " (ISNULL(T2.U_Quantity, "
            strQuery += " (Select ISNULL(" & dblQty & ",U_LFactor) From [@Z_OCAJ] T0 "
            strQuery += " JOIN [@Z_OCPR] T1 On T0.U_Calories "
            strQuery += " = T1.U_CPAdj Where T1.U_CardCode = '" + strCardCode + "')) "
            strQuery += " ),1)) As 'Qty', "
            strQuery += " T2.U_Dislike,T2.U_Medical, "
            strQuery += " T2.U_Remarks As 'Remarks' "
            strQuery += " From [@Z_MED2] T0 JOIN [@Z_OMED] T1 On T1.DocEntry = T0.DocEntry  "

            strQuery += " And T1.U_CatType = 'I' "

            strQuery += " LEFT OUTER JOIN [@Z_OFSL] T2 ON T2.U_ItemCode = T0.U_ItemCode  "
            strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = Convert(VarChar(8),T2.U_PrgDate,112)  "
            strQuery += " And T2.U_CardCode = '" + strCardCode + "' "
            strQuery += " And T2.U_FType = 'LN' "
            strQuery += " And T2.U_SFood = 'R'  "
            strQuery += " And T2.U_Select = 'Y' "
            strQuery += " And T2.U_ProgramID = '" + strProgramID + "'"

            strQuery += " Where T1.U_MenuType = 'R' "
            strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = '" + strPrgDate + "'"
            strQuery += " And T1.U_PrgCode = '" + strProgram + "'"
            strQuery += " And T0.U_ItemCode Is Not Null "

            If Not oRecordSet.EoF Then
                oDataTable.ExecuteQuery(strQuery)
                oGrid.DataTable = oDataTable
            Else

                strQuery = " Select (Select Case When T0.U_ItemCode = T2.U_ItemCode Then 'Y' ELSE 'N' END) As 'Select', "
                strQuery += " T0.U_ItemCode,U_ItemName, "
                strQuery += " (Select ISNULL( "
                strQuery += " (ISNULL(T2.U_Quantity, "
                strQuery += " (Select ISNULL(" & dblQty & ",U_LFactor) From [@Z_OCAJ] T0 "
                strQuery += " JOIN [@Z_OCPR] T1 On T0.U_Calories "
                strQuery += " = T1.U_CPAdj Where T1.U_CardCode = '" + strCardCode + "')) "
                strQuery += " ),1)) As 'Qty', "
                strQuery += " T2.U_Dislike,T2.U_Medical, "
                strQuery += " T2.U_Remarks As 'Remarks' "
                strQuery += " From [@Z_MED2] T0 JOIN [@Z_OMED] T1 On T1.DocEntry = T0.DocEntry  "

                strQuery += " And T1.U_CatType = 'G' "

                strQuery += " LEFT OUTER JOIN [@Z_OFSL] T2 ON T2.U_ItemCode = T0.U_ItemCode  "
                strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = Convert(VarChar(8),T2.U_PrgDate,112)  "
                strQuery += " And T2.U_CardCode = '" + strCardCode + "' "
                strQuery += " And T2.U_FType = 'LN' "
                strQuery += " And T2.U_SFood = 'R'  "
                strQuery += " And T2.U_Select = 'Y' "
                strQuery += " And T2.U_ProgramID = '" + strProgramID + "'"

                strQuery += " JOIN OITB T3 On T3.ItmsGrpCod = T1.U_GrpCode "
                strQuery += " JOIN OITM T4 On T4.ItmsGrpCod = T3.ItmsGrpCod And T4.ItmsGrpCod = T1.U_GrpCode "

                strQuery += " Where T1.U_MenuType = 'R' "
                strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = '" + strPrgDate + "'"
                strQuery += " And T4.ItemCode = '" + strProgram + "'"
                strQuery += " And T0.U_ItemCode Is Not Null "

                oDataTable.ExecuteQuery(strQuery)
                oGrid.DataTable = oDataTable
            End If

            'Lunch - Custom
            oGrid = oForm.Items.Item("25").Specific
            oDataTable = oForm.DataSources.DataTables.Item("Dt_Lunch_C")
            strQuery = " Select (Select Case When T2.U_ItemCode = T1.ItemCode Then 'Y' ELSE 'N' END) As 'Select', "
            strQuery += " T2.U_ItemCode,T1.ItemName As U_ItemName, "
            strQuery += " T2.U_Quantity As 'Qty', "
            strQuery += " T2.U_Dislike,T2.U_Medical, "
            strQuery += " T2.U_Remarks As 'Remarks' "
            strQuery += " From  [@Z_OFSL] T2 JOIN OITM T1  ON T2.U_ItemCode = T1.ItemCode  "
            strQuery += " And T2.U_CardCode = '" + strCardCode + "' "
            strQuery += " And T2.U_FType = 'LN' "
            strQuery += " And T2.U_SFood = 'C'  "
            strQuery += " And T2.U_Select = 'Y' "
            strQuery += " And T2.U_ProgramID = '" + strProgramID + "'"
            strQuery += " And Convert(VarChar(8),T2.U_PrgDate,112) = '" + strPrgDate + "'"
            oDataTable.ExecuteQuery(strQuery)
            oGrid.DataTable = oDataTable

            formatGrid(oForm, "10")
            formatGrid_Custom(oForm, "25", "CFL_2")

            'Lunch(Side) - Regular
            oGrid = oForm.Items.Item("12").Specific
            dblQty = getQuantityBasedonCaloriesRatio(oForm, "12")
            oDataTable = oForm.DataSources.DataTables.Item("Dt_LunchS_R")


            strQuery = " Select (Select Case When T0.U_ItemCode = T2.U_ItemCode Then 'Y' ELSE 'N' END) As 'Select', "
            strQuery += " T0.U_ItemCode,U_ItemName, "
            strQuery += " (Select ISNULL( "
            strQuery += " (ISNULL(T2.U_Quantity, "
            strQuery += " (Select ISNULL(" & dblQty & ",U_LSFactor) From [@Z_OCAJ] T0 "
            strQuery += " JOIN [@Z_OCPR] T1 On T0.U_Calories "
            strQuery += " = T1.U_CPAdj Where T1.U_CardCode = '" + strCardCode + "')) "
            strQuery += " ),1)) As 'Qty', "
            strQuery += " T2.U_Dislike,T2.U_Medical, "
            strQuery += " T2.U_Remarks As 'Remarks' "
            strQuery += " From [@Z_MED3] T0 JOIN [@Z_OMED] T1 On T1.DocEntry = T0.DocEntry  "

            strQuery += " And T1.U_CatType = 'I' "

            strQuery += " LEFT OUTER JOIN [@Z_OFSL] T2 ON T2.U_ItemCode = T0.U_ItemCode  "
            strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = Convert(VarChar(8),T2.U_PrgDate,112)  "
            strQuery += " And T2.U_CardCode = '" + strCardCode + "' "
            strQuery += " And T2.U_FType = 'LS' "
            strQuery += " And T2.U_SFood = 'R'  "
            strQuery += " And T2.U_Select = 'Y' "
            strQuery += " And T2.U_ProgramID = '" + strProgramID + "'"
            strQuery += " Where T1.U_MenuType = 'R' "
            strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = '" + strPrgDate + "'"
            strQuery += " And T1.U_PrgCode = '" + strProgram + "'"
            strQuery += " And T0.U_ItemCode Is Not Null "

            If Not oRecordSet.EoF Then
                oDataTable.ExecuteQuery(strQuery)
                oGrid.DataTable = oDataTable
            Else

                strQuery = " Select (Select Case When T0.U_ItemCode = T2.U_ItemCode Then 'Y' ELSE 'N' END) As 'Select', "
                strQuery += " T0.U_ItemCode,U_ItemName, "
                strQuery += " (Select ISNULL( "
                strQuery += " (ISNULL(T2.U_Quantity, "
                strQuery += " (Select ISNULL(" & dblQty & ",U_LSFactor) From [@Z_OCAJ] T0 "
                strQuery += " JOIN [@Z_OCPR] T1 On T0.U_Calories "
                strQuery += " = T1.U_CPAdj Where T1.U_CardCode = '" + strCardCode + "')) "
                strQuery += " ),1)) As 'Qty', "
                strQuery += " T2.U_Dislike,T2.U_Medical, "
                strQuery += " T2.U_Remarks As 'Remarks' "
                strQuery += " From [@Z_MED3] T0 JOIN [@Z_OMED] T1 On T1.DocEntry = T0.DocEntry  "

                strQuery += " And T1.U_CatType = 'G' "

                strQuery += " LEFT OUTER JOIN [@Z_OFSL] T2 ON T2.U_ItemCode = T0.U_ItemCode  "
                strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = Convert(VarChar(8),T2.U_PrgDate,112)  "
                strQuery += " And T2.U_CardCode = '" + strCardCode + "' "
                strQuery += " And T2.U_FType = 'LS' "
                strQuery += " And T2.U_SFood = 'R'  "
                strQuery += " And T2.U_Select = 'Y' "
                strQuery += " And T2.U_ProgramID = '" + strProgramID + "'"

                strQuery += " JOIN OITB T3 On T3.ItmsGrpCod = T1.U_GrpCode "
                strQuery += " JOIN OITM T4 On T4.ItmsGrpCod = T3.ItmsGrpCod And T4.ItmsGrpCod = T1.U_GrpCode "

                strQuery += " Where T1.U_MenuType = 'R' "
                strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = '" + strPrgDate + "'"
                strQuery += " And T4.ItemCode = '" + strProgram + "'"
                strQuery += " And T0.U_ItemCode Is Not Null "

                oDataTable.ExecuteQuery(strQuery)
                oGrid.DataTable = oDataTable

            End If

            'Lunch(Side) - Custom
            oGrid = oForm.Items.Item("26").Specific
            oDataTable = oForm.DataSources.DataTables.Item("Dt_LunchS_C")
            strQuery = " Select (Select Case When T2.U_ItemCode = T1.ItemCode Then 'Y' ELSE 'N' END) As 'Select' , "
            strQuery += " T2.U_ItemCode,T1.ItemName As U_ItemName, "
            strQuery += " T2.U_Quantity As 'Qty', "
            strQuery += " T2.U_Dislike,T2.U_Medical, "
            strQuery += " T2.U_Remarks As 'Remarks' "
            strQuery += " From  [@Z_OFSL] T2 JOIN OITM T1  ON T2.U_ItemCode = T1.ItemCode  "
            strQuery += " And T2.U_CardCode = '" + strCardCode + "' "
            strQuery += " And T2.U_FType = 'LS' "
            strQuery += " And T2.U_SFood = 'C'  "
            strQuery += " And T2.U_Select = 'Y' "
            strQuery += " And T2.U_ProgramID = '" + strProgramID + "'"
            strQuery += " And Convert(VarChar(8),T2.U_PrgDate,112) = '" + strPrgDate + "'"
            oDataTable.ExecuteQuery(strQuery)
            oGrid.DataTable = oDataTable

            formatGrid(oForm, "12")
            formatGrid_Custom(oForm, "26", "CFL_3")

            'Snack - Regular
            oGrid = oForm.Items.Item("14").Specific
            dblQty = getQuantityBasedonCaloriesRatio(oForm, "14")
            oDataTable = oForm.DataSources.DataTables.Item("Dt_Snack_R")

            strQuery = " Select (Select Case When T0.U_ItemCode = T2.U_ItemCode Then 'Y' ELSE 'N' END) As 'Select', "
            strQuery += " T0.U_ItemCode,U_ItemName, "
            strQuery += " (Select ISNULL( "
            strQuery += " (ISNULL(T2.U_Quantity, "
            strQuery += " (Select ISNULL(" & dblQty & ",U_SFactor) From [@Z_OCAJ] T0 "
            strQuery += " JOIN [@Z_OCPR] T1 On T0.U_Calories "
            strQuery += " = T1.U_CPAdj Where T1.U_CardCode = '" + strCardCode + "')) "
            strQuery += " ),1)) As 'Qty', "
            strQuery += " T2.U_Dislike,T2.U_Medical, "
            strQuery += " T2.U_Remarks As 'Remarks' "
            strQuery += " From [@Z_MED4] T0 JOIN [@Z_OMED] T1 On T1.DocEntry = T0.DocEntry  "

            strQuery += " And T1.U_CatType = 'I' "

            strQuery += " LEFT OUTER JOIN [@Z_OFSL] T2 ON T2.U_ItemCode = T0.U_ItemCode  "
            strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = Convert(VarChar(8),T2.U_PrgDate,112)  "
            strQuery += " And T2.U_CardCode = '" + strCardCode + "' "
            strQuery += " And T2.U_FType = 'SK' "
            strQuery += " And T2.U_SFood = 'R'  "
            strQuery += " And T2.U_Select = 'Y' "
            strQuery += " And T2.U_ProgramID = '" + strProgramID + "'"
            strQuery += " Where T1.U_MenuType = 'R' "
            strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = '" + strPrgDate + "'"
            strQuery += " And T1.U_PrgCode = '" + strProgram + "'"
            strQuery += " And T0.U_ItemCode Is Not Null "

            If Not oRecordSet.EoF Then
                oDataTable.ExecuteQuery(strQuery)
                oGrid.DataTable = oDataTable
            Else

                strQuery = " Select (Select Case When T0.U_ItemCode = T2.U_ItemCode Then 'Y' ELSE 'N' END) As 'Select', "
                strQuery += " T0.U_ItemCode,U_ItemName, "
                strQuery += " (Select ISNULL( "
                strQuery += " (ISNULL(T2.U_Quantity, "
                strQuery += " (Select ISNULL(" & dblQty & ",U_SFactor) From [@Z_OCAJ] T0 "
                strQuery += " JOIN [@Z_OCPR] T1 On T0.U_Calories "
                strQuery += " = T1.U_CPAdj Where T1.U_CardCode = '" + strCardCode + "')) "
                strQuery += " ),1)) As 'Qty', "
                strQuery += " T2.U_Dislike,T2.U_Medical, "
                strQuery += " T2.U_Remarks As 'Remarks' "
                strQuery += " From [@Z_MED4] T0 JOIN [@Z_OMED] T1 On T1.DocEntry = T0.DocEntry  "

                strQuery += " And T1.U_CatType = 'G' "

                strQuery += " LEFT OUTER JOIN [@Z_OFSL] T2 ON T2.U_ItemCode = T0.U_ItemCode  "
                strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = Convert(VarChar(8),T2.U_PrgDate,112)  "
                strQuery += " And T2.U_CardCode = '" + strCardCode + "' "
                strQuery += " And T2.U_FType = 'SK' "
                strQuery += " And T2.U_SFood = 'R'  "
                strQuery += " And T2.U_Select = 'Y' "
                strQuery += " And T2.U_ProgramID = '" + strProgramID + "'"

                strQuery += " JOIN OITB T3 On T3.ItmsGrpCod = T1.U_GrpCode "
                strQuery += " JOIN OITM T4 On T4.ItmsGrpCod = T3.ItmsGrpCod And T4.ItmsGrpCod = T1.U_GrpCode "

                strQuery += " Where T1.U_MenuType = 'R' "
                strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = '" + strPrgDate + "'"
                strQuery += " And T4.ItemCode = '" + strProgram + "'"
                strQuery += " And T0.U_ItemCode Is Not Null "

                oDataTable.ExecuteQuery(strQuery)
                oGrid.DataTable = oDataTable

            End If

            'Snack - Custom
            oGrid = oForm.Items.Item("27").Specific
            oDataTable = oForm.DataSources.DataTables.Item("Dt_Snack_C")
            strQuery = " Select (Select Case When T2.U_ItemCode = T1.ItemCode Then 'Y' ELSE 'N' END) As 'Select', "
            strQuery += " T2.U_ItemCode,T1.ItemName As U_ItemName, "
            strQuery += " T2.U_Quantity As 'Qty', "
            strQuery += " T2.U_Dislike,T2.U_Medical, "
            strQuery += " T2.U_Remarks As 'Remarks' "
            strQuery += " From  [@Z_OFSL] T2 JOIN OITM T1  ON T2.U_ItemCode = T1.ItemCode  "
            strQuery += " And T2.U_CardCode = '" + strCardCode + "' "
            strQuery += " And T2.U_FType = 'SK' "
            strQuery += " And T2.U_SFood = 'C'  "
            strQuery += " And T2.U_Select = 'Y' "
            strQuery += " And T2.U_ProgramID = '" + strProgramID + "'"
            strQuery += " And Convert(VarChar(8),T2.U_PrgDate,112) = '" + strPrgDate + "'"
            oDataTable.ExecuteQuery(strQuery)
            oGrid.DataTable = oDataTable

            formatGrid(oForm, "14")
            formatGrid_Custom(oForm, "27", "CFL_4")


            'Dinner - Regular
            oGrid = oForm.Items.Item("16").Specific
            dblQty = getQuantityBasedonCaloriesRatio(oForm, "16")
            oDataTable = oForm.DataSources.DataTables.Item("Dt_Dinner_R")

            strQuery = " Select (Select Case When T0.U_ItemCode = T2.U_ItemCode Then 'Y' ELSE 'N' END) As 'Select', "
            strQuery += " T0.U_ItemCode,U_ItemName, "
            strQuery += " (Select ISNULL( "
            strQuery += " (ISNULL(T2.U_Quantity, "
            strQuery += " (Select ISNULL(" & dblQty & ",U_DFactor) From [@Z_OCAJ] T0 "
            strQuery += " JOIN [@Z_OCPR] T1 On T0.U_Calories "
            strQuery += " = T1.U_CPAdj Where T1.U_CardCode = '" + strCardCode + "')) "
            strQuery += " ),1)) As 'Qty', "
            strQuery += " T2.U_Dislike,T2.U_Medical, "
            strQuery += " T2.U_Remarks As 'Remarks' "
            strQuery += " From [@Z_MED5] T0 JOIN [@Z_OMED] T1 On T1.DocEntry = T0.DocEntry  "

            strQuery += " And T1.U_CatType = 'I' "

            strQuery += " LEFT OUTER JOIN [@Z_OFSL] T2 ON T2.U_ItemCode = T0.U_ItemCode  "
            strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = Convert(VarChar(8),T2.U_PrgDate,112)  "
            strQuery += " And T2.U_CardCode = '" + strCardCode + "' "
            strQuery += " And T2.U_FType = 'DI' "
            strQuery += " And T2.U_SFood = 'R'  "
            strQuery += " And T2.U_Select = 'Y' "
            strQuery += " And T2.U_ProgramID = '" + strProgramID + "'"
            strQuery += " Where T1.U_MenuType = 'R' "
            strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = '" + strPrgDate + "'"
            strQuery += " And T1.U_PrgCode = '" + strProgram + "'"
            strQuery += " And T0.U_ItemCode Is Not Null "
            If Not oRecordSet.EoF Then
                oDataTable.ExecuteQuery(strQuery)
                oGrid.DataTable = oDataTable
            Else
                strQuery = " Select (Select Case When T0.U_ItemCode = T2.U_ItemCode Then 'Y' ELSE 'N' END) As 'Select', "
                strQuery += " T0.U_ItemCode,U_ItemName, "
                strQuery += " (Select ISNULL( "
                strQuery += " (ISNULL(T2.U_Quantity, "
                strQuery += " (Select ISNULL(" & dblQty & ",U_DFactor) From [@Z_OCAJ] T0 "
                strQuery += " JOIN [@Z_OCPR] T1 On T0.U_Calories "
                strQuery += " = T1.U_CPAdj Where T1.U_CardCode = '" + strCardCode + "')) "
                strQuery += " ),1)) As 'Qty', "
                strQuery += " T2.U_Dislike,T2.U_Medical, "
                strQuery += " T2.U_Remarks As 'Remarks' "
                strQuery += " From [@Z_MED5] T0 JOIN [@Z_OMED] T1 On T1.DocEntry = T0.DocEntry  "

                strQuery += " And T1.U_CatType = 'G' "

                strQuery += " LEFT OUTER JOIN [@Z_OFSL] T2 ON T2.U_ItemCode = T0.U_ItemCode  "
                strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = Convert(VarChar(8),T2.U_PrgDate,112)  "
                strQuery += " And T2.U_CardCode = '" + strCardCode + "' "
                strQuery += " And T2.U_FType = 'DI' "
                strQuery += " And T2.U_SFood = 'R'  "
                strQuery += " And T2.U_Select = 'Y' "
                strQuery += " And T2.U_ProgramID = '" + strProgramID + "'"

                strQuery += " JOIN OITB T3 On T3.ItmsGrpCod = T1.U_GrpCode "
                strQuery += " JOIN OITM T4 On T4.ItmsGrpCod = T3.ItmsGrpCod And T4.ItmsGrpCod = T1.U_GrpCode "

                strQuery += " Where T1.U_MenuType = 'R' "
                strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = '" + strPrgDate + "'"
                strQuery += " And T4.ItemCode = '" + strProgram + "'"
                strQuery += " And T0.U_ItemCode Is Not Null "

                oDataTable.ExecuteQuery(strQuery)
                oGrid.DataTable = oDataTable
            End If


            'Dinner - Custom
            oGrid = oForm.Items.Item("28").Specific
            oDataTable = oForm.DataSources.DataTables.Item("Dt_Dinner_C")
            strQuery = " Select (Select Case When T2.U_ItemCode = T1.ItemCode Then 'Y' ELSE 'N' END) As 'Select', "
            strQuery += " T2.U_ItemCode,T1.ItemName As U_ItemName, "
            strQuery += " T2.U_Quantity As 'Qty', "
            strQuery += " T2.U_Dislike,T2.U_Medical, "
            strQuery += " T2.U_Remarks As 'Remarks' "
            strQuery += " From  [@Z_OFSL] T2 JOIN OITM T1  ON T2.U_ItemCode = T1.ItemCode  "
            strQuery += " And T2.U_CardCode = '" + strCardCode + "' "
            strQuery += " And T2.U_FType = 'DI' "
            strQuery += " And T2.U_SFood = 'C'  "
            strQuery += " And T2.U_Select = 'Y' "
            strQuery += " And T2.U_ProgramID = '" + strProgramID + "'"
            strQuery += " And Convert(VarChar(8),T2.U_PrgDate,112) = '" + strPrgDate + "'"
            oDataTable.ExecuteQuery(strQuery)
            oGrid.DataTable = oDataTable

            formatGrid(oForm, "16")
            formatGrid_Custom(oForm, "28", "CFL_5")

            'Dinner(Side) - Regular
            oGrid = oForm.Items.Item("18").Specific
            dblQty = getQuantityBasedonCaloriesRatio(oForm, "18")
            oDataTable = oForm.DataSources.DataTables.Item("Dt_DinnerS_R")
            strQuery = " Select (Select Case When T0.U_ItemCode = T2.U_ItemCode Then 'Y' ELSE 'N' END) As 'Select', "
            strQuery += " T0.U_ItemCode,U_ItemName, "
            strQuery += " (Select ISNULL( "
            strQuery += " (ISNULL(T2.U_Quantity, "
            strQuery += " (Select ISNULL(" & dblQty & ",U_DSFactor) From [@Z_OCAJ] T0 "
            strQuery += " JOIN [@Z_OCPR] T1 On T0.U_Calories "
            strQuery += " = T1.U_CPAdj Where T1.U_CardCode = '" + strCardCode + "')) "
            strQuery += " ),1)) As 'Qty', "
            strQuery += " T2.U_Dislike,T2.U_Medical, "
            strQuery += " T2.U_Remarks As 'Remarks' "
            strQuery += " From [@Z_MED6] T0 JOIN [@Z_OMED] T1 On T1.DocEntry = T0.DocEntry  "

            strQuery += " And T1.U_CatType = 'I' "

            strQuery += " LEFT OUTER JOIN [@Z_OFSL] T2 ON T2.U_ItemCode = T0.U_ItemCode  "
            strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = Convert(VarChar(8),T2.U_PrgDate,112)  "
            strQuery += " And T2.U_CardCode = '" + strCardCode + "' "
            strQuery += " And T2.U_FType = 'DS' "
            strQuery += " And T2.U_SFood = 'R'  "
            strQuery += " And T2.U_Select = 'Y' "
            strQuery += " And T2.U_ProgramID = '" + strProgramID + "'"
            strQuery += " Where T1.U_MenuType = 'R' "
            strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = '" + strPrgDate + "'"
            strQuery += " And T1.U_PrgCode = '" + strProgram + "'"
            strQuery += " And T0.U_ItemCode Is Not Null "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oDataTable.ExecuteQuery(strQuery)
                oGrid.DataTable = oDataTable

            Else

                strQuery = " Select (Select Case When T0.U_ItemCode = T2.U_ItemCode Then 'Y' ELSE 'N' END) As 'Select', "
                strQuery += " T0.U_ItemCode,U_ItemName, "
                strQuery += " (Select ISNULL( "
                strQuery += " (ISNULL(T2.U_Quantity, "
                strQuery += " (Select ISNULL(" & dblQty & ",U_DSFactor) From [@Z_OCAJ] T0 "
                strQuery += " JOIN [@Z_OCPR] T1 On T0.U_Calories "
                strQuery += " = T1.U_CPAdj Where T1.U_CardCode = '" + strCardCode + "')) "
                strQuery += " ),1)) As 'Qty', "
                strQuery += " T2.U_Dislike,T2.U_Medical, "
                strQuery += " T2.U_Remarks As 'Remarks' "
                strQuery += " From [@Z_MED6] T0 JOIN [@Z_OMED] T1 On T1.DocEntry = T0.DocEntry  "

                strQuery += " And T1.U_CatType = 'G' "

                strQuery += " LEFT OUTER JOIN [@Z_OFSL] T2 ON T2.U_ItemCode = T0.U_ItemCode  "
                strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = Convert(VarChar(8),T2.U_PrgDate,112)  "
                strQuery += " And T2.U_CardCode = '" + strCardCode + "' "
                strQuery += " And T2.U_FType = 'DS' "
                strQuery += " And T2.U_SFood = 'R'  "
                strQuery += " And T2.U_Select = 'Y' "
                strQuery += " And T2.U_ProgramID = '" + strProgramID + "'"

                strQuery += " JOIN OITB T3 On T3.ItmsGrpCod = T1.U_GrpCode "
                strQuery += " JOIN OITM T4 On T4.ItmsGrpCod = T3.ItmsGrpCod And T4.ItmsGrpCod = T1.U_GrpCode "

                strQuery += " Where T1.U_MenuType = 'R' "
                strQuery += " And Convert(VarChar(8),T1.U_MenuDate,112) = '" + strPrgDate + "'"
                strQuery += " And T4.ItemCode = '" + strProgram + "'"
                strQuery += " And T0.U_ItemCode Is Not Null "

                oDataTable.ExecuteQuery(strQuery)
                oGrid.DataTable = oDataTable
            End If

            'Dinner(Side) - Custom
            oGrid = oForm.Items.Item("29").Specific
            oDataTable = oForm.DataSources.DataTables.Item("Dt_DinnerS_C")
            strQuery = " Select (Select Case When T2.U_ItemCode = T1.ItemCode Then 'Y' ELSE 'N' END) As 'Select', "
            strQuery += " T2.U_ItemCode,T1.ItemName As U_ItemName, "
            strQuery += " T2.U_Quantity As 'Qty', "
            strQuery += " T2.U_Dislike,T2.U_Medical, "
            strQuery += " T2.U_Remarks As 'Remarks' "
            strQuery += " From  [@Z_OFSL] T2 JOIN OITM T1  ON T2.U_ItemCode = T1.ItemCode  "
            strQuery += " And T2.U_CardCode = '" + strCardCode + "' "
            strQuery += " And T2.U_FType = 'DS' "
            strQuery += " And T2.U_SFood = 'C'  "
            strQuery += " And T2.U_Select = 'Y' "
            strQuery += " And T2.U_ProgramID = '" + strProgramID + "'"
            strQuery += " And Convert(VarChar(8),T2.U_PrgDate,112) = '" + strPrgDate + "'"
            oDataTable.ExecuteQuery(strQuery)
            oGrid.DataTable = oDataTable

            formatGrid(oForm, "18")
            formatGrid_Custom(oForm, "29", "CFL_6")

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub formatGrid(ByVal oForm As SAPbouiCOM.Form, ByVal strID As String)
        Try
            oForm.Freeze(True)
            oGrid = oForm.Items.Item(strID).Specific
            fillHeader(oForm, strID)
            formatAll(oForm, oGrid, strID)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm.Freeze(False)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub formatGrid_Custom(ByVal oForm As SAPbouiCOM.Form, ByVal strID As String, ByVal strCFLID As String)
        Try
            oForm.Freeze(True)
            oGrid = oForm.Items.Item(strID).Specific
            fillHeader(oForm, strID)
            formatAll_Custom(oForm, oGrid, strID, strCFLID)
            If oGrid.DataTable.GetValue("U_ItemCode", oGrid.DataTable.Rows.Count - 1).ToString() <> "" Then
                oGrid.DataTable.Rows.Add(1)
                oGrid.CommonSetting.SetRowBackColor(oGrid.DataTable.Rows.Count, RGB(255, 255, 255))
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm.Freeze(False)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub fillHeader(ByVal aForm As SAPbouiCOM.Form, ByVal strGridID As String)
        Try
            aForm.Freeze(True)
            oGrid = aForm.Items.Item(strGridID).Specific
            oGrid.RowHeaders.TitleObject.Caption = "#"
            Dim strCardCode As String = oForm.Items.Item("31").Specific.value.ToString().Trim()
            Dim strDislike As String = String.Empty
            Dim strMedical As String = String.Empty
            For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(index, (index + 1).ToString())

                'To Fill Dislike & Medical...In Regular Foods.
                '========
                Dim strItemCode As String = oGrid.DataTable.GetValue("U_ItemCode", index).ToString()
                If (oApplication.Utilities.hasBOM(strItemCode)) Then
                    strDislike = oApplication.Utilities.GetDisLikeItem(strCardCode, strItemCode)
                    strMedical = oApplication.Utilities.GetMedicalItem(strCardCode, strItemCode)
                    oApplication.Utilities.get_ChildItems(strCardCode, strItemCode, strDislike, strMedical)
                Else
                    strDislike = oApplication.Utilities.GetDisLikeItem(strCardCode, strItemCode)
                    strMedical = oApplication.Utilities.GetMedicalItem(strCardCode, strItemCode)
                End If
                If strDislike.Trim().Length > 0 Then
                    oGrid.DataTable.SetValue("U_Dislike", index, strDislike)
                End If
                If strMedical.Trim().Length > 0 Then
                    oGrid.DataTable.SetValue("U_Medical", index, strMedical)
                End If
                '========

                If oGrid.DataTable.GetValue("Select", index).ToString() = "Y" Then
                    oGrid.CommonSetting.SetRowBackColor(index + 1, RGB(0, 255, 0))
                Else
                    oGrid.CommonSetting.SetRowBackColor(index + 1, RGB(255, 255, 255))
                End If
            Next
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            aForm.Freeze(False)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Function getItemName(ByVal strItemCode As String) As String
        Dim _retVal As String = String.Empty
        Try
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select ItemName From OITM Where ItemCode = '" + strItemCode + "'")
            If Not oRecordSet.EoF Then
                _retVal = oRecordSet.Fields.Item(0).Value
            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Private Sub formatAll(ByVal oForm As SAPbouiCOM.Form, ByVal oGrid As SAPbouiCOM.Grid, ByVal strID As String)
        Try
            oGrid.Columns.Item("U_ItemCode").TitleObject.Caption = "Food Code"
            oGrid.Columns.Item("U_ItemName").TitleObject.Caption = "Food Name"
            oGrid.Columns.Item("Select").TitleObject.Caption = "Select"
            oGrid.Columns.Item("Qty").TitleObject.Caption = "Quantity"
            oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("U_ItemCode").Editable = False
            oGrid.Columns.Item("U_ItemName").Editable = False
            'oGrid.Columns.Item("U_Dislike").Visible = False
            'oGrid.Columns.Item("U_Medical").Visible = False
            oGrid.Columns.Item("U_Dislike").TitleObject.Caption = "Dislike"
            oGrid.Columns.Item("U_Medical").TitleObject.Caption = "Medical"
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub formatAll_Custom(ByVal oForm As SAPbouiCOM.Form, ByVal oGrid As SAPbouiCOM.Grid, ByVal strID As String, ByVal strCFLID As String)
        Try
            oGrid.Columns.Item("U_ItemCode").TitleObject.Caption = "Food Code"
            oGrid.Columns.Item("U_ItemName").TitleObject.Caption = "Food Name"
            oGrid.Columns.Item("Select").TitleObject.Caption = "Select"
            oGrid.Columns.Item("Qty").TitleObject.Caption = "Quantity"
            oGrid.Columns.Item("U_ItemCode").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditTextCol = oGrid.Columns.Item("U_ItemName")
            oEditTextCol.ChooseFromListUID = strCFLID
            oEditTextCol.ChooseFromListAlias = "ItemName"
            oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("U_ItemCode").Visible = True
            oGrid.Columns.Item("U_ItemCode").Editable = False
            oGrid.Columns.Item("U_ItemName").Editable = True
            'oGrid.Columns.Item("U_Dislike").Visible = False
            'oGrid.Columns.Item("U_Medical").Visible = False
            oGrid.Columns.Item("U_Dislike").TitleObject.Caption = "Dislike"
            oGrid.Columns.Item("U_Medical").TitleObject.Caption = "Medical"
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Function changePane(ByVal oForm As SAPbouiCOM.Form, ByVal strItem As String) As String
        Try
            Dim _retVal As String = String.Empty
            Select Case strItem
                Case "1"
                    oForm.PaneLevel = 1
                Case "2"
                    oForm.PaneLevel = 2
                Case "3"
                    oForm.PaneLevel = 3
                Case "4"
                    oForm.PaneLevel = 4
                Case "5"
                    oForm.PaneLevel = 5
                Case "6"
                    oForm.PaneLevel = 6
            End Select
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Private Sub getFoodValue(ByVal aForm As SAPbouiCOM.Form, ByVal strGridID As String, _
                             ByRef strItemCode As String, ByRef dblQty As Double, Optional ByRef strRemarks As String = "")
        Try
            Dim blnSelected As Boolean = False
            oGrid = aForm.Items.Item(strGridID).Specific
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue("Select", intRow) = "Y" Then
                    strItemCode = oGrid.DataTable.GetValue("U_ItemCode", intRow)
                    dblQty = CDbl(oGrid.DataTable.GetValue("Qty", intRow))
                    strRemarks = oGrid.DataTable.GetValue("Remarks", intRow)
                    Exit For
                End If
            Next
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub addChooseFromListConditions(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oCFLs = oForm.ChooseFromLists

            oCFL = oCFLs.Item("CFL_1")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.BracketOpenNum = 2
            For i As Integer = 0 To 5
                If i > 0 And i < 4 Then
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCon = oCons.Add()
                    oCon.BracketOpenNum = 1
                End If
                If i = 0 Then
                    oCon.[Alias] = "InvntItem"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 1 Then
                    oCon.[Alias] = "SellItem"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 2 Then
                    oCon.[Alias] = "validFor"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 3 Then
                    oCon.[Alias] = "U_ISFOOD"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 4 Then
                    oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery("Select U_Prefix From [@Z_OFCI] Where U_Type = 'I' And U_Active = 'Y' ")
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

                            oCon.[Alias] = "ItemCode"
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_START
                            oCon.CondVal = oRecordSet.Fields.Item(0).Value.ToString()
                            oCon.BracketCloseNum = 1
                            oRecordSet.MoveNext()
                            intConCount += 1

                        End While
                        oCon.BracketCloseNum = 2
                    End If
                ElseIf i = 5 Then
                    oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery("Select ItmsGrpCod From OITB Where U_FinishedFood = 'Y' ")
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

                            oCon.[Alias] = "ItmsGrpCod"
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
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
                ElseIf i > 0 And i < 4 Then
                    oCon.BracketCloseNum = 1
                End If
            Next
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_2")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.BracketOpenNum = 2
            For i As Integer = 0 To 5
                If i > 0 And i < 4 Then
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCon = oCons.Add()
                    oCon.BracketOpenNum = 1
                End If
                If i = 0 Then
                    oCon.[Alias] = "InvntItem"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 1 Then
                    oCon.[Alias] = "SellItem"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 2 Then
                    oCon.[Alias] = "validFor"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 3 Then
                    oCon.[Alias] = "U_ISFOOD"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 4 Then
                    oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery("Select U_Prefix From [@Z_OFCI] Where U_Type = 'I' And U_Active = 'Y' ")
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

                            oCon.[Alias] = "ItemCode"
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_START
                            oCon.CondVal = oRecordSet.Fields.Item(0).Value.ToString()
                            oCon.BracketCloseNum = 1
                            oRecordSet.MoveNext()
                            intConCount += 1

                        End While
                        oCon.BracketCloseNum = 2
                    End If
                ElseIf i = 5 Then
                    oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery("Select ItmsGrpCod From OITB Where U_FinishedFood = 'Y' ")
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

                            oCon.[Alias] = "ItmsGrpCod"
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
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
                ElseIf i > 0 And i < 4 Then
                    oCon.BracketCloseNum = 1
                End If
            Next
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_3")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.BracketOpenNum = 2
            For i As Integer = 0 To 5
                If i > 0 And i < 4 Then
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCon = oCons.Add()
                    oCon.BracketOpenNum = 1
                End If
                If i = 0 Then
                    oCon.[Alias] = "InvntItem"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 1 Then
                    oCon.[Alias] = "SellItem"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 2 Then
                    oCon.[Alias] = "validFor"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 3 Then
                    oCon.[Alias] = "U_ISFOOD"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 4 Then
                    oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery("Select U_Prefix From [@Z_OFCI] Where U_Type = 'I' And U_Active = 'Y' ")
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

                            oCon.[Alias] = "ItemCode"
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_START
                            oCon.CondVal = oRecordSet.Fields.Item(0).Value.ToString()
                            oCon.BracketCloseNum = 1
                            oRecordSet.MoveNext()
                            intConCount += 1

                        End While
                        oCon.BracketCloseNum = 2
                    End If
                ElseIf i = 5 Then
                    oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery("Select ItmsGrpCod From OITB Where U_FinishedFood = 'Y' ")
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

                            oCon.[Alias] = "ItmsGrpCod"
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
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
                ElseIf i > 0 And i < 4 Then
                    oCon.BracketCloseNum = 1
                End If
            Next
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_4")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.BracketOpenNum = 2
            For i As Integer = 0 To 5
                If i > 0 And i < 4 Then
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCon = oCons.Add()
                    oCon.BracketOpenNum = 1
                End If
                If i = 0 Then
                    oCon.[Alias] = "InvntItem"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 1 Then
                    oCon.[Alias] = "SellItem"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 2 Then
                    oCon.[Alias] = "validFor"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 3 Then
                    oCon.[Alias] = "U_ISFOOD"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 4 Then
                    oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery("Select U_Prefix From [@Z_OFCI] Where U_Type = 'I' And U_Active = 'Y' ")
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

                            oCon.[Alias] = "ItemCode"
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_START
                            oCon.CondVal = oRecordSet.Fields.Item(0).Value.ToString()
                            oCon.BracketCloseNum = 1
                            oRecordSet.MoveNext()
                            intConCount += 1

                        End While
                        oCon.BracketCloseNum = 2
                    End If
                ElseIf i = 5 Then
                    oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery("Select ItmsGrpCod From OITB Where U_FinishedFood = 'Y' ")
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

                            oCon.[Alias] = "ItmsGrpCod"
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
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
                ElseIf i > 0 And i < 4 Then
                    oCon.BracketCloseNum = 1
                End If
            Next
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_5")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.BracketOpenNum = 2
            For i As Integer = 0 To 5
                If i > 0 And i < 4 Then
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCon = oCons.Add()
                    oCon.BracketOpenNum = 1
                End If
                If i = 0 Then
                    oCon.[Alias] = "InvntItem"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 1 Then
                    oCon.[Alias] = "SellItem"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 2 Then
                    oCon.[Alias] = "validFor"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 3 Then
                    oCon.[Alias] = "U_ISFOOD"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 4 Then
                    oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery("Select U_Prefix From [@Z_OFCI] Where U_Type = 'I' And U_Active = 'Y' ")
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

                            oCon.[Alias] = "ItemCode"
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_START
                            oCon.CondVal = oRecordSet.Fields.Item(0).Value.ToString()
                            oCon.BracketCloseNum = 1
                            oRecordSet.MoveNext()
                            intConCount += 1

                        End While
                        oCon.BracketCloseNum = 2
                    End If
                ElseIf i = 5 Then
                    oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery("Select ItmsGrpCod From OITB Where U_FinishedFood = 'Y' ")
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

                            oCon.[Alias] = "ItmsGrpCod"
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
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
                ElseIf i > 0 And i < 4 Then
                    oCon.BracketCloseNum = 1
                End If
            Next
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_6")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.BracketOpenNum = 2
            For i As Integer = 0 To 5
                If i > 0 And i < 4 Then
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCon = oCons.Add()
                    oCon.BracketOpenNum = 1
                End If
                If i = 0 Then
                    oCon.[Alias] = "InvntItem"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 1 Then
                    oCon.[Alias] = "SellItem"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 2 Then
                    oCon.[Alias] = "validFor"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 3 Then
                    oCon.[Alias] = "U_ISFOOD"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                ElseIf i = 4 Then
                    oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery("Select U_Prefix From [@Z_OFCI] Where U_Type = 'I' And U_Active = 'Y' ")
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

                            oCon.[Alias] = "ItemCode"
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_START
                            oCon.CondVal = oRecordSet.Fields.Item(0).Value.ToString()
                            oCon.BracketCloseNum = 1
                            oRecordSet.MoveNext()
                            intConCount += 1

                        End While
                        oCon.BracketCloseNum = 2
                    End If
                ElseIf i = 5 Then
                    oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery("Select ItmsGrpCod From OITB Where U_FinishedFood = 'Y' ")
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

                            oCon.[Alias] = "ItmsGrpCod"
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
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
                ElseIf i > 0 And i < 4 Then
                    oCon.BracketCloseNum = 1
                End If
            Next
            oCFL.SetConditions(oCons)

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Function getQuantityBasedonCalories(ByVal oForm As SAPbouiCOM.Form, ByVal strItem As String) As Double
        Try
            Dim _retVal As Double = 0
            Dim strCardCode As String = CType(oForm.Items.Item("31").Specific, SAPbouiCOM.EditText).Value
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            Select Case strItem
                Case "24"
                    strQuery = " Select ISNULL(U_BFactor,0) From [@Z_OCAJ] T0 JOIN [@Z_OCPR] T1 On T0.U_Calories = T1.U_CPAdj "
                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
                    oRecordSet.DoQuery(strQuery)
                    If Not oRecordSet.EoF Then
                        _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
                    Else
                        _retVal = 1
                    End If
                Case "25"
                    strQuery = " Select ISNULL(U_LFactor,0) From [@Z_OCAJ] T0 JOIN [@Z_OCPR] T1 On T0.U_Calories = T1.U_CPAdj "
                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
                    oRecordSet.DoQuery(strQuery)
                    If Not oRecordSet.EoF Then
                        _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
                    Else
                        _retVal = 1
                    End If
                Case "26"
                    strQuery = " Select ISNULL(U_LSFactor,0) From [@Z_OCAJ] T0 JOIN [@Z_OCPR] T1 On T0.U_Calories = T1.U_CPAdj "
                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
                    oRecordSet.DoQuery(strQuery)
                    If Not oRecordSet.EoF Then
                        _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
                    Else
                        _retVal = 1
                    End If
                Case "27"
                    strQuery = " Select ISNULL(U_SFactor,0) From [@Z_OCAJ] T0 JOIN [@Z_OCPR] T1 On T0.U_Calories = T1.U_CPAdj "
                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
                    oRecordSet.DoQuery(strQuery)
                    If Not oRecordSet.EoF Then
                        _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
                    Else
                        _retVal = 1
                    End If
                Case "28"
                    strQuery = " Select ISNULL(U_DFactor,0) From [@Z_OCAJ] T0 JOIN [@Z_OCPR] T1 On T0.U_Calories = T1.U_CPAdj "
                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
                    oRecordSet.DoQuery(strQuery)
                    If Not oRecordSet.EoF Then
                        _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
                    Else
                        _retVal = 1
                    End If
                Case "29"
                    strQuery = " Select ISNULL(U_DSFactor,0) From [@Z_OCAJ] T0 JOIN [@Z_OCPR] T1 On T0.U_Calories = T1.U_CPAdj "
                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
                    oRecordSet.DoQuery(strQuery)
                    If Not oRecordSet.EoF Then
                        _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
                    Else
                        _retVal = 1
                    End If
            End Select
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Private Sub fillProgramDate(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oDataTable As SAPbouiCOM.DataTable

            oGrid = oForm.Items.Item("33").Specific
            Dim strFrmDate As String = oForm.Items.Item("22").Specific.value
            Dim strToDate As String = oForm.Items.Item("22_").Specific.value
            Dim strCardCode As String = oForm.Items.Item("31").Specific.value
            Dim strType As String = oForm.Items.Item("30").Specific.value
            Dim strRef As String = oForm.Items.Item("34").Specific.value

            'Item
            oDataTable = oForm.DataSources.DataTables.Item("Dt_ProgramDates")

            'Madhu Commented this for Phase II Requirement On 20150708
            'strQuery = " Select U_PrgDate From [@Z_CPM1] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry "
            'strQuery += " Where Convert(VarChar(8),U_PrgDate,112) >= '" + strFrmDate + "'"
            'strQuery += " And Convert(VarChar(8),U_PrgDate,112) <= '" + strToDate + "'"
            'strQuery += " And T1.U_CardCode = '" + strCardCode + "'"
            'strQuery += " And U_AppStatus = 'I' "
            'strQuery += " AND T1.U_RemDays > 0  "

            strQuery = " Select T0.U_PrgDate From [@Z_CPM1] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry "
            'strQuery += " LEFT OUTER JOIN [@Z_CPM6] T2 On T1.DocEntry = T2.DocEntry "
            strQuery += " JOIN [@Z_OCPR] T3 On T1.U_CardCode = T3.U_CardCode  "
            strQuery += " Where Convert(VarChar(8),T0.U_PrgDate,112) >= '" & strFrmDate & "'"
            strQuery += " And Convert(VarChar(8),T0.U_PrgDate,112) <= '" & strToDate & "'"
            strQuery += " And T1.U_CardCode = '" + strCardCode + "'"
            strQuery += " And T0.U_AppStatus = 'I' "
            strQuery += " And ISNULL(T0.U_ONOFFSTA,'O') = 'O' "
            strQuery += " AND T1.U_RemDays > 0  "
            If strType.Trim() = "I" Then
                'strQuery += " And T1.U_Type = 'I'  "
                strQuery += " And ISNULL(T1.U_InvRef,T2.U_InvRef) = '" & strRef & "'"
            ElseIf strType.Trim() = "T" Then
                'strQuery += " And T1.U_Type = 'I'  "
                strQuery += " And T1.U_TrnRef = '" & strRef & "'"
            ElseIf strType.Trim() = "P" Then
                strQuery += " And T1.DocEntry = '" & strRef & "'"
            End If
            strQuery += " And T0.U_PrgDate Not In (Select Distinct U_DelDate From RDR1 Where U_ProgramID = '" & strRef & "' "
            strQuery += " And (LineStatus = 'O' Or (LineStatus = 'C' And TargetType <> '-1')) "
            strQuery += " ) "
            strQuery += "  And  "
            strQuery += "  (  "
            strQuery += "  (T0.U_PrgDate < T3.U_SuFrDt And ISNULL(T3.U_SuToDt,'') = '')  "
            strQuery += "  OR  "
            strQuery += "  (1 = 1)  "
            strQuery += "  ) "

            oDataTable.ExecuteQuery(strQuery)
            oGrid.DataTable = oDataTable
            oGrid.Columns.Item("U_PrgDate").Editable = False
            oGrid.Columns.Item("U_PrgDate").TitleObject.Caption = "Program Date"
            oGrid.RowHeaders.TitleObject.Caption = "#"
            For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(index, (index + 1).ToString())
            Next
            'Dim strFromDate As String = oGrid.DataTable.GetValue("U_PrgDate", 0).ToString()
            'CType(oForm.Items.Item("35").Specific, SAPbouiCOM.EditText).Value = strFromDate
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Function getSelectedPrgDate(ByVal aForm As SAPbouiCOM.Form, ByVal strGridID As String) As String
        Try
            Dim _retVal As String = String.Empty
            oGrid = aForm.Items.Item(strGridID).Specific
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.Rows.IsSelected(intRow) Then
                    _retVal = CDate(oGrid.DataTable.GetValue(0, intRow)).ToString("yyyyMMdd")
                    CType(oForm.Items.Item("35").Specific, SAPbouiCOM.EditText).Value = _retVal
                End If
            Next
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Private Function validation(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Dim _retVal As Boolean = True
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim strProgramID As String = CType(oForm.Items.Item("36").Specific, SAPbouiCOM.EditText).Value
            oGrid = oForm.Items.Item("33").Specific
            For intRow As Integer = 0 To oGrid.Rows.Count - 1
                Dim strPrgDate As String = CDate(oGrid.DataTable.GetValue("U_PrgDate", intRow)).ToString("yyyyMMdd")
                strQuery = "Select Code From [@Z_OFSL] Where U_ProgramID = '" + strProgramID + "'"
                strQuery += " And Convert(VarChar(8),U_PrgDate,112) = '" + strPrgDate + "'"
                oRecordSet.DoQuery(strQuery)
                If oRecordSet.EoF Then
                    _retVal = False
                    Exit For
                ElseIf oRecordSet.RecordCount = 0 Then
                    _retVal = False
                    Exit For
                End If
            Next
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Private Function validationSessions(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Dim _retVal As Boolean = True
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim strProgramID As String = CType(oForm.Items.Item("36").Specific, SAPbouiCOM.EditText).Value
            oGrid = oForm.Items.Item("33").Specific
            For intRow As Integer = 0 To oGrid.Rows.Count - 1
                Dim strPrgDate As String = CDate(oGrid.DataTable.GetValue(0, intRow)).ToString("yyyyMMdd")
                strQuery = "Select Count(U_FType) As 'Sessions',U_FType From [@Z_OFSL] Where U_ProgramID = '" + strProgramID + "'"
                strQuery += " And Convert(VarChar(8),U_PrgDate,112) = '" + strPrgDate + "'"
                strQuery += "  Group By U_FType,U_Select Having ISNULL(U_Select ,'N') = 'Y' "
                oRecordSet.DoQuery(strQuery)
                If oRecordSet.EoF Then
                    _retVal = False
                    Exit For
                ElseIf oRecordSet.RecordCount < 6 Then
                    _retVal = False
                    Exit For
                End If
            Next
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Private Function validationRegular(ByVal oForm As SAPbouiCOM.Form, ByVal strGridID As String) As Boolean
        Dim _retVal As Boolean = True
        Try
            oGrid = oForm.Items.Item(strGridID).Specific
            Dim IsSelectedAlready As Boolean = False
            For intRow As Integer = 0 To oGrid.Rows.Count - 1
                If (oGrid.DataTable.GetValue("Select", intRow)) = "Y" Then
                    If Not IsSelectedAlready Then
                        IsSelectedAlready = True
                    Else
                        _retVal = False
                        Exit For
                    End If
                End If
            Next
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function enableSession(ByVal aForm As SAPbouiCOM.Form, ByVal strProgram As String) As String
        Dim _retVal As String = String.Empty
        Try
            oForm.Freeze(True)
            oForm.Items.Item("1").Enabled = True
            oForm.Items.Item("2").Enabled = True
            oForm.Items.Item("3").Enabled = True
            oForm.Items.Item("4").Enabled = True
            oForm.Items.Item("5").Enabled = True
            oForm.Items.Item("6").Enabled = True
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            strQuery = "Select "
            strQuery += " ISNULL(U_BF,'N') As 'U_BF',ISNULL(U_LN,'N') As 'U_LN',ISNULL(U_LS,'N') As 'U_LS', "
            strQuery += " ISNULL(U_SK,'N') As 'U_SK',ISNULL(U_DN,'N') As 'U_DN',ISNULL(U_DS,'N') As 'U_DS' "
            strQuery += " From OITM Where ItemCode = '" & strProgram & "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then

                oForm.Items.Item("1").Enabled = IIf(oRecordSet.Fields.Item("U_BF").Value = "Y", True, False)
                oForm.Items.Item("2").Enabled = IIf(oRecordSet.Fields.Item("U_LN").Value = "Y", True, False)
                oForm.Items.Item("3").Enabled = IIf(oRecordSet.Fields.Item("U_LS").Value = "Y", True, False)
                oForm.Items.Item("4").Enabled = IIf(oRecordSet.Fields.Item("U_SK").Value = "Y", True, False)
                oForm.Items.Item("5").Enabled = IIf(oRecordSet.Fields.Item("U_DN").Value = "Y", True, False)
                oForm.Items.Item("6").Enabled = IIf(oRecordSet.Fields.Item("U_DS").Value = "Y", True, False)

                Dim strNotSelFol As String = IIf(Not oForm.Items.Item("1").Enabled, _
                                                 IIf(Not oForm.Items.Item("2").Enabled, _
                                                     IIf(Not oForm.Items.Item("3").Enabled, _
                                                         IIf(Not oForm.Items.Item("4").Enabled, _
                                                             IIf(Not oForm.Items.Item("5").Enabled, _
                                                                 IIf(Not oForm.Items.Item("6").Enabled, _
                                                                     "",
                                                                     oForm.Items.Item("6").UniqueID),
                                                                 oForm.Items.Item("5").UniqueID),
                                                             oForm.Items.Item("4").UniqueID),
                                                         oForm.Items.Item("3").UniqueID), _
                                                     oForm.Items.Item("2").UniqueID) _
                                                 , oForm.Items.Item("1").UniqueID)


                If strNotSelFol = "" Then
                    oForm.Items.Item("8").Enabled = False
                    oForm.Items.Item("24").Enabled = False
                Else
                    oForm.Items.Item(strNotSelFol).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                End If

                _retVal = strNotSelFol
                oForm.Freeze(False)
                Return _retVal
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm.Freeze(False)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        Finally
            oForm.Freeze(False)
        End Try
        Return _retVal
    End Function

    Private Function getQuantityBasedonCaloriesRatio(ByVal oForm As SAPbouiCOM.Form, ByVal strItem As String) As Double
        Try
            Dim _retVal As Double = 0
            Dim strCardCode As String = CType(oForm.Items.Item("31").Specific, SAPbouiCOM.EditText).Value
            Dim strPrgDate As String = CType(oForm.Items.Item("35").Specific, SAPbouiCOM.EditText).Value
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            Select Case strItem
                Case "24", "8"
                    strQuery = " Select TOP 1 U_Ratio From [@Z_CPR7] T0 "
                    strQuery += " JOIN [@Z_OCPR] T1 ON T0.DocEntry = T1.DocEntry  "
                    strQuery += " JOIN [@Z_OCRT] T2 On T0.U_BF = T2.U_Code "
                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "' And T2.U_FType = 'BF' "
                    strQuery += " And Convert(VarChar(8),T0.U_PrgDate,112) <= '" + strPrgDate + "'"
                    strQuery += " Order By T0.U_PrgDate DESC "
                    oRecordSet.DoQuery(strQuery)
                    If Not oRecordSet.EoF Then
                        _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
                    Else
                        _retVal = 1
                    End If
                Case "25", "10"
                    strQuery = " Select TOP 1 U_Ratio From [@Z_CPR7] T0 "
                    strQuery += " JOIN [@Z_OCPR] T1 ON T0.DocEntry = T1.DocEntry  "
                    strQuery += " JOIN [@Z_OCRT] T2 On T0.U_LN = T2.U_Code "
                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "' And T2.U_FType = 'LN' "
                    strQuery += " And Convert(VarChar(8),T0.U_PrgDate,112) <= '" + strPrgDate + "'"
                    strQuery += " Order By T0.U_PrgDate DESC "
                    oRecordSet.DoQuery(strQuery)
                    If Not oRecordSet.EoF Then
                        _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
                    Else
                        _retVal = 1
                    End If
                Case "26", "12"
                    strQuery = " Select TOP 1 U_Ratio From [@Z_CPR7] T0 "
                    strQuery += " JOIN [@Z_OCPR] T1 ON T0.DocEntry = T1.DocEntry  "
                    strQuery += " JOIN [@Z_OCRT] T2 On T0.U_LS = T2.U_Code "
                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "' And T2.U_FType = 'LS' "
                    strQuery += " And Convert(VarChar(8),T0.U_PrgDate,112) <= '" + strPrgDate + "'"
                    strQuery += " Order By T0.U_PrgDate DESC "
                    oRecordSet.DoQuery(strQuery)
                    If Not oRecordSet.EoF Then
                        _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
                    Else
                        _retVal = 1
                    End If
                Case "27", "14"
                    strQuery = " Select TOP 1 U_Ratio From [@Z_CPR7] T0 "
                    strQuery += " JOIN [@Z_OCPR] T1 ON T0.DocEntry = T1.DocEntry  "
                    strQuery += " JOIN [@Z_OCRT] T2 On T0.U_SK = T2.U_Code "
                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "' And T2.U_FType = 'SK' "
                    strQuery += " And Convert(VarChar(8),T0.U_PrgDate,112) <= '" + strPrgDate + "'"
                    strQuery += " Order By T0.U_PrgDate DESC "
                    oRecordSet.DoQuery(strQuery)
                    If Not oRecordSet.EoF Then
                        _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
                    Else
                        _retVal = 1
                    End If
                Case "28", "16"
                    strQuery = " Select TOP 1 U_Ratio From [@Z_CPR7] T0 "
                    strQuery += " JOIN [@Z_OCPR] T1 ON T0.DocEntry = T1.DocEntry  "
                    strQuery += " JOIN [@Z_OCRT] T2 On T0.U_DI = T2.U_Code "
                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "' And T2.U_FType = 'DI' "
                    strQuery += " And Convert(VarChar(8),T0.U_PrgDate,112) <= '" + strPrgDate + "'"
                    strQuery += " Order By T0.U_PrgDate DESC "
                    oRecordSet.DoQuery(strQuery)
                    If Not oRecordSet.EoF Then
                        _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
                    Else
                        _retVal = 1
                    End If
                Case "29", "18"
                    strQuery = " Select TOP 1 U_Ratio From [@Z_CPR7] T0 "
                    strQuery += " JOIN [@Z_OCPR] T1 ON T0.DocEntry = T1.DocEntry  "
                    strQuery += " JOIN [@Z_OCRT] T2 On T0.U_DS = T2.U_Code "
                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "' And T2.U_FType = 'DS' "
                    strQuery += " And Convert(VarChar(8),T0.U_PrgDate,112) <= '" + strPrgDate + "'"
                    strQuery += " Order By T0.U_PrgDate DESC "
                    oRecordSet.DoQuery(strQuery)
                    If Not oRecordSet.EoF Then
                        _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
                    Else
                        _retVal = 1
                    End If
            End Select
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

End Class

'Public Class clsSelectFood
'    Inherits clsBase

'    Private objForm As SAPbouiCOM.Form
'    Dim oRecordSet As SAPbobsCOM.Recordset
'    Dim strQuery As String = String.Empty
'    Dim oGrid As SAPbouiCOM.Grid
'    Private oDt_ProgramFood As SAPbouiCOM.DataTable
'    Private objBaseFormID As String
'    Public oDT_Quantity As SAPbouiCOM.DataTable
'    Public oDT_SType As SAPbouiCOM.DataTable
'    Public oDT_SRemarks As SAPbouiCOM.DataTable

'    Public Sub LoadForm(ByVal objParentFormID As String, ByVal strFromDate As String, ByVal strTodate As String, _
'                        ByVal strCardCode As String, ByVal strType As String, ByVal strRef As String, ByVal strProgram As String)
'        Try
'            oForm = oApplication.Utilities.LoadForm(xml_Z_OPSL_1, frm_Z_OPSL_1)
'            oForm = oApplication.SBO_Application.Forms.ActiveForm()
'            CType(oForm.Items.Item("6").Specific, SAPbouiCOM.EditText).Value = strFromDate
'            CType(oForm.Items.Item("7").Specific, SAPbouiCOM.EditText).Value = strTodate
'            CType(oForm.Items.Item("8").Specific, SAPbouiCOM.EditText).Value = objParentFormID
'            CType(oForm.Items.Item("10").Specific, SAPbouiCOM.EditText).Value = strCardCode
'            CType(oForm.Items.Item("11").Specific, SAPbouiCOM.EditText).Value = strType
'            CType(oForm.Items.Item("12").Specific, SAPbouiCOM.EditText).Value = strRef
'            CType(oForm.Items.Item("14").Specific, SAPbouiCOM.EditText).Value = strProgram
'            initialize(oForm)
'        Catch ex As Exception 
' oApplication.Log.Trace_DIET_AddOn_Error(ex)
'            oForm.Freeze(False)
'        End Try
'    End Sub

'#Region "Item Event"

'    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
'        Try
'            If pVal.FormTypeEx = frm_Z_OPSL_1 Then
'                Select Case pVal.BeforeAction
'                    Case True
'                        Select Case pVal.EventType
'                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
'                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
'                                If pVal.ItemUID = "_2" Then
'                                    oForm.Freeze(True)
'                                    If validation(oForm) = False Then
'                                        If oApplication.SBO_Application.MessageBox("Food Not Selected for Some of the dates...Continue...?", , "Yes", "No") = 2 Then
'                                            oForm.Freeze(False)
'                                            BubbleEvent = False
'                                            Exit Sub
'                                        Else
'                                            Dim oBaseForm As SAPbouiCOM.Form = Nothing
'                                            For index As Integer = 0 To oApplication.SBO_Application.Forms.Count
'                                                If oApplication.SBO_Application.Forms.Item(index).UniqueID = CType(oForm.Items.Item("8").Specific, SAPbouiCOM.EditText).Value Then
'                                                    oBaseForm = oApplication.SBO_Application.Forms.Item(index)
'                                                    Exit For
'                                                End If
'                                            Next
'                                            If Not IsNothing(oBaseForm) Then
'                                                Dim oBDBDataSource As SAPbouiCOM.DBDataSource
'                                                Dim oBaseMatrix As SAPbouiCOM.Matrix
'                                                oBDBDataSource = oBaseForm.DataSources.DBDataSources.Item("@Z_PSL1")
'                                                oBaseMatrix = oBaseForm.Items.Item("3").Specific
'                                                oBaseMatrix.Clear()
'                                                oBaseMatrix.FlushToDataSource()
'                                                oGrid = oForm.Items.Item("1").Specific
'                                                Dim intMatrixRow As Integer = 0
'                                                oDT_Quantity = oForm.DataSources.DataTables.Item("Dt_Quantity")
'                                                oDT_SType = oForm.DataSources.DataTables.Item("Dt_SType")
'                                                oDT_SRemarks = oForm.DataSources.DataTables.Item("Dt_SRemarks")
'                                                For intRow As Integer = 0 To oGrid.Rows.Count - 1
'                                                    Dim strPrgDate As String = CDate(oGrid.DataTable.GetValue(0, intRow)).ToString("yyyyMMdd")
'                                                    For intCol As Integer = 1 To oGrid.Columns.Count - 1
'                                                        Dim strColName As String = oGrid.DataTable.Columns.Item(intCol).Name
'                                                        Dim strFoodType As String = getFoodType(oForm, strColName)
'                                                        Dim strItem As String = oGrid.DataTable.GetValue(oGrid.DataTable.Columns.Item(intCol).Name, intRow)
'                                                        Dim dblQty As Double = CDbl(oDT_Quantity.GetValue(oGrid.DataTable.Columns.Item(intCol).Name, intRow))
'                                                        Dim strSType As String = (oDT_SType.GetValue(oGrid.DataTable.Columns.Item(intCol).Name, intRow))
'                                                        Dim strSRemarks As String = (oDT_SRemarks.GetValue(oGrid.DataTable.Columns.Item(intCol).Name, intRow))

'                                                        If strItem.Trim().Length > 0 Then
'                                                            Dim strName As String = oApplication.Utilities.GetItemName(strItem)
'                                                            oBaseMatrix.AddRow(1, oBaseMatrix.RowCount)
'                                                            oBaseMatrix.FlushToDataSource()
'                                                            oBDBDataSource.SetValue("LineId", intMatrixRow, (intMatrixRow + 1).ToString())
'                                                            oBDBDataSource.SetValue("U_DelDate", intMatrixRow, strPrgDate)
'                                                            oBDBDataSource.SetValue("U_FType", intMatrixRow, strFoodType)
'                                                            oBDBDataSource.SetValue("U_ItemCode", intMatrixRow, strItem)
'                                                            oBDBDataSource.SetValue("U_ItemName", intMatrixRow, strName)
'                                                            'Dim dblQty As Double = getQuantityBasedonCalories(oForm, strColName) * 1
'                                                            oBDBDataSource.SetValue("U_Quantity", intMatrixRow, dblQty)
'                                                            oBDBDataSource.SetValue("U_UnitPrice", intMatrixRow, "0")
'                                                            oBDBDataSource.SetValue("U_SFood", intMatrixRow, strSType)
'                                                            oBDBDataSource.SetValue("U_Remarks", intMatrixRow, strSRemarks)
'                                                            intMatrixRow += 1
'                                                            oBaseMatrix.LoadFromDataSource()
'                                                            oBaseMatrix.FlushToDataSource()
'                                                        End If
'                                                    Next
'                                                Next
'                                                oBaseMatrix.LoadFromDataSource()
'                                                oBaseMatrix.FlushToDataSource()

'                                                'Check Dislike & Medical item
'                                                For index As Integer = 0 To oBDBDataSource.Size - 1
'                                                    Dim strItem As String = oBDBDataSource.GetValue("U_ItemCode", index).Trim()
'                                                    Dim strCardCode As String = oForm.Items.Item("10").Specific.value.Trim()
'                                                    Dim strDisLike As String = String.Empty
'                                                    Dim strMedical As String = String.Empty

'                                                    If (oApplication.Utilities.hasBOM(strItem)) Then
'                                                        strDisLike = oApplication.Utilities.GetDisLikeItem(strCardCode, strItem)
'                                                        strMedical = oApplication.Utilities.GetMedicalItem(strCardCode, strItem)
'                                                        oApplication.Utilities.get_ChildItems(strCardCode, strItem, strDisLike, strMedical)
'                                                        If strDisLike.Length > 0 Then
'                                                            oBDBDataSource.SetValue("U_Dislike", index, strDisLike)
'                                                        End If
'                                                        If strMedical.Length > 0 Then
'                                                            oBDBDataSource.SetValue("U_Medical", index, strMedical)
'                                                        End If
'                                                    Else
'                                                        strDisLike = oApplication.Utilities.GetDisLikeItem(strCardCode, strItem)
'                                                        strMedical = oApplication.Utilities.GetMedicalItem(strCardCode, strItem)
'                                                        If strDisLike.Length > 0 Then
'                                                            oBDBDataSource.SetValue("U_Dislike", index, strDisLike)
'                                                        End If
'                                                        If strMedical.Length > 0 Then
'                                                            oBDBDataSource.SetValue("U_Medical", index, strMedical)
'                                                        End If
'                                                    End If

'                                                Next
'                                                oBaseMatrix.LoadFromDataSource()
'                                                oBaseMatrix.FlushToDataSource()
'                                                If oBaseForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oBaseForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
'                                                oBaseForm.Update()
'                                                updateFoodPreference(oForm, "1")
'                                                oForm.Close()
'                                                Exit Sub
'                                            End If
'                                        End If
'                                    Else
'                                        Dim oBaseForm As SAPbouiCOM.Form = Nothing
'                                        For index As Integer = 0 To oApplication.SBO_Application.Forms.Count
'                                            If oApplication.SBO_Application.Forms.Item(index).UniqueID = CType(oForm.Items.Item("8").Specific, SAPbouiCOM.EditText).Value Then
'                                                oBaseForm = oApplication.SBO_Application.Forms.Item(index)
'                                                Exit For
'                                            End If
'                                        Next
'                                        If Not IsNothing(oBaseForm) Then
'                                            Dim oBDBDataSource As SAPbouiCOM.DBDataSource
'                                            Dim oBaseMatrix As SAPbouiCOM.Matrix
'                                            oBDBDataSource = oBaseForm.DataSources.DBDataSources.Item("@Z_PSL1")
'                                            oBaseMatrix = oBaseForm.Items.Item("3").Specific
'                                            oBaseMatrix.Clear()
'                                            oBaseMatrix.FlushToDataSource()
'                                            oGrid = oForm.Items.Item("1").Specific
'                                            Dim intMatrixRow As Integer = 0
'                                            oDT_Quantity = oForm.DataSources.DataTables.Item("Dt_Quantity")
'                                            oDT_SType = oForm.DataSources.DataTables.Item("Dt_SType")
'                                            oDT_SRemarks = oForm.DataSources.DataTables.Item("Dt_SRemarks")

'                                            For intRow As Integer = 0 To oGrid.Rows.Count - 1
'                                                Dim strPrgDate As String = CDate(oGrid.DataTable.GetValue(0, intRow)).ToString("yyyyMMdd")
'                                                For intCol As Integer = 1 To oGrid.Columns.Count - 1
'                                                    Dim strColName As String = oGrid.DataTable.Columns.Item(intCol).Name
'                                                    Dim strFoodType As String = getFoodType(oForm, strColName)
'                                                    Dim strItem As String = oGrid.DataTable.GetValue(oGrid.DataTable.Columns.Item(intCol).Name, intRow)
'                                                    Dim dblQty As Double = CDbl(oDT_Quantity.GetValue(oGrid.DataTable.Columns.Item(intCol).Name, intRow))
'                                                    Dim strSType As String = (oDT_SType.GetValue(oGrid.DataTable.Columns.Item(intCol).Name, intRow))
'                                                    Dim strSRemarks As String = (oDT_SRemarks.GetValue(oGrid.DataTable.Columns.Item(intCol).Name, intRow))
'                                                    If strItem.Trim().Length > 0 Then
'                                                        Dim strName As String = oApplication.Utilities.GetItemName(strItem)
'                                                        oBaseMatrix.AddRow(1, oBaseMatrix.RowCount)
'                                                        oBaseMatrix.FlushToDataSource()
'                                                        oBDBDataSource.SetValue("LineId", intMatrixRow, (intMatrixRow + 1).ToString())
'                                                        oBDBDataSource.SetValue("U_DelDate", intMatrixRow, strPrgDate)
'                                                        oBDBDataSource.SetValue("U_FType", intMatrixRow, strFoodType)
'                                                        oBDBDataSource.SetValue("U_ItemCode", intMatrixRow, strItem)
'                                                        oBDBDataSource.SetValue("U_ItemName", intMatrixRow, strName)
'                                                        ' Dim dblQty As Double = getQuantityBasedonCalories(oForm, strColName) * 1
'                                                        oBDBDataSource.SetValue("U_Quantity", intMatrixRow, dblQty)
'                                                        oBDBDataSource.SetValue("U_UnitPrice", intMatrixRow, "0")
'                                                        oBDBDataSource.SetValue("U_SFood", intMatrixRow, strSType)
'                                                        oBDBDataSource.SetValue("U_Remarks", intMatrixRow, strSRemarks)
'                                                        intMatrixRow += 1
'                                                        oBaseMatrix.LoadFromDataSource()
'                                                        oBaseMatrix.FlushToDataSource()
'                                                    End If
'                                                Next
'                                            Next

'                                            oBaseMatrix.LoadFromDataSource()
'                                            oBaseMatrix.FlushToDataSource()

'                                            'Check Dislike & Medical item
'                                            For index As Integer = 0 To oBDBDataSource.Size - 1
'                                                Dim strItem As String = oBDBDataSource.GetValue("U_ItemCode", index).Trim()
'                                                Dim strCardCode As String = oForm.Items.Item("10").Specific.value.Trim()
'                                                Dim strDisLike As String = String.Empty
'                                                Dim strMedical As String = String.Empty

'                                                If (oApplication.Utilities.hasBOM(strItem)) Then
'                                                    strDisLike = oApplication.Utilities.GetDisLikeItem(strCardCode, strItem)
'                                                    strMedical = oApplication.Utilities.GetMedicalItem(strCardCode, strItem)
'                                                    oApplication.Utilities.get_ChildItems(strCardCode, strItem, strDisLike, strMedical)
'                                                    If strDisLike.Length > 0 Then
'                                                        oBDBDataSource.SetValue("U_Dislike", index, strDisLike)
'                                                    End If
'                                                    If strMedical.Length > 0 Then
'                                                        oBDBDataSource.SetValue("U_Medical", index, strMedical)
'                                                    End If
'                                                Else
'                                                    strDisLike = oApplication.Utilities.GetDisLikeItem(strCardCode, strItem)
'                                                    strMedical = oApplication.Utilities.GetMedicalItem(strCardCode, strItem)
'                                                    If strDisLike.Length > 0 Then
'                                                        oBDBDataSource.SetValue("U_Dislike", index, strDisLike)
'                                                    End If
'                                                    If strMedical.Length > 0 Then
'                                                        oBDBDataSource.SetValue("U_Medical", index, strMedical)
'                                                    End If
'                                                End If

'                                            Next
'                                            oBaseMatrix.LoadFromDataSource()
'                                            oBaseMatrix.FlushToDataSource()
'                                            If oBaseForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oBaseForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
'                                            oBaseForm.Update()
'                                            updateFoodPreference(oForm, "1")
'                                            oForm.Close()
'                                            Exit Sub
'                                        End If
'                                    End If
'                                    oForm.Freeze(False)
'                                ElseIf pVal.ItemUID = "5" Then
'                                    Dim strFrmDate As String = getSelectedPrgDate(oForm, "1")
'                                    Dim strCardCode As String = CType(oForm.Items.Item("10").Specific, SAPbouiCOM.EditText).Value
'                                    Dim strProgram As String = CType(oForm.Items.Item("14").Specific, SAPbouiCOM.EditText).Value
'                                    If strFrmDate.Length > 0 Then
'                                        Dim oFoodMenu As clsFoodMenu
'                                        oFoodMenu = New clsFoodMenu()
'                                        'oFoodMenu.LoadForm(oForm.UniqueID, strFrmDate, strCardCode, strProgram)
'                                    Else
'                                        oApplication.Utilities.Message("No row selected", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
'                                        BubbleEvent = False
'                                    End If
'                                End If
'                        End Select
'                    Case False
'                        Select Case pVal.EventType
'                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
'                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
'                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
'                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
'                        End Select
'                End Select
'            End If
'        Catch ex As Exception 
'oApplication.Log.Trace_DIET_AddOn_Error(ex)
'            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
'            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
'            oForm.Freeze(False)
'        End Try
'    End Sub

'#End Region

'    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
'        Try
'            oForm.PaneLevel = 1
'            oForm.Items.Item("4").TextStyle = 5
'            oForm.DataSources.DataTables.Add("Dt_ProgramFood")
'            oForm.DataSources.DataTables.Add("Dt_Quantity")
'            oForm.DataSources.DataTables.Add("Dt_SType")
'            oForm.DataSources.DataTables.Add("Dt_SRemarks")

'            oDt_ProgramFood = oForm.DataSources.DataTables.Item("Dt_ProgramFood")
'            strQuery = " Select U_PrgDate,U_BF,U_Lunch,U_LunchS,U_Snack,U_Dinner,U_DinnerS From [@Z_CPM1] Where 1 = 2 "
'            oDt_ProgramFood.ExecuteQuery(strQuery)
'            oGrid = oForm.Items.Item("1").Specific
'            oGrid.DataTable = oDt_ProgramFood

'            oDT_Quantity = oForm.DataSources.DataTables.Item("Dt_Quantity")
'            strQuery = " Select U_PrgDate,U_BF,U_Lunch,U_LunchS,U_Snack,U_Dinner,U_DinnerS From [@Z_CPM4] Where 1 = 2 "
'            oDT_Quantity.ExecuteQuery(strQuery)

'            oDT_SType = oForm.DataSources.DataTables.Item("Dt_SType")
'            strQuery = " Select U_PrgDate,U_BF,U_Lunch,U_LunchS,U_Snack,U_Dinner,U_DinnerS From [@Z_CPM5] Where 1 = 2 "
'            oDT_SType.ExecuteQuery(strQuery)

'            oDT_SRemarks = oForm.DataSources.DataTables.Item("Dt_SRemarks")
'            strQuery = " Select U_PrgDate,Convert(VarChar(254),U_BF) As U_BF ,Convert(VarChar(254),U_Lunch) As U_Lunch,Convert(VarChar(254),U_LunchS) As U_LunchS,Convert(VarChar(254),U_Snack) As U_Snack,Convert(VarChar(254),U_Dinner) As U_Dinner,Convert(VarChar(254),U_DinnerS) As U_DinnerS From [@Z_CPM4] Where 1 = 2 "
'            oDT_SRemarks.ExecuteQuery(strQuery)

'            fillProgramDate(oForm)

'            oForm.Update()
'        Catch ex As Exception 
' oApplication.Log.Trace_DIET_AddOn_Error(ex)
'            Throw ex 
''oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
'        End Try
'    End Sub

'    Private Sub fillProgramDate(ByVal oForm As SAPbouiCOM.Form)
'        Try
'            Dim oDataTable As SAPbouiCOM.DataTable
'            Dim oDataTable1 As SAPbouiCOM.DataTable
'            Dim oDataTable2 As SAPbouiCOM.DataTable
'            Dim oDataTable3 As SAPbouiCOM.DataTable

'            oGrid = oForm.Items.Item("1").Specific
'            Dim strFrmDate As String = oForm.Items.Item("6").Specific.value
'            Dim strToDate As String = oForm.Items.Item("7").Specific.value
'            Dim strCardCode As String = oForm.Items.Item("10").Specific.value
'            Dim strType As String = oForm.Items.Item("11").Specific.value
'            Dim strRef As String = oForm.Items.Item("12").Specific.value

'            'Item
'            oDataTable = oForm.DataSources.DataTables.Item("Dt_ProgramFood")
'            strQuery = " Select U_PrgDate,U_BF,U_Lunch,U_LunchS,U_Snack,U_Dinner,U_DinnerS From [@Z_CPM1] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry "
'            strQuery += " Where Convert(VarChar(8),U_PrgDate,112) >= '" + strFrmDate + "'"
'            strQuery += " And Convert(VarChar(8),U_PrgDate,112) <= '" + strToDate + "'"
'            strQuery += " And T1.U_CardCode = '" + strCardCode + "'"
'            strQuery += " And U_AppStatus = 'I' "
'            strQuery += " AND T1.U_RemDays > 0  "
'            If strType.Trim() = "I" Then
'                'strQuery += " And T1.U_Type = 'I'  "
'                strQuery += " And T1.U_InvRef = '" + strRef + "'"
'            Else
'                'strQuery += " And T1.U_Type = 'I'  "
'                strQuery += " And T1.U_TrnRef = '" + strRef + "'"
'            End If
'            oDataTable.ExecuteQuery(strQuery)
'            oGrid.DataTable = oDataTable

'            'Quantity
'            oDataTable1 = oForm.DataSources.DataTables.Item("Dt_Quantity")
'            strQuery = " Select T0.U_PrgDate,ISNULL(T0.U_BF,0) As U_BF,ISNULL(T0.U_Lunch,0) As U_Lunch , ISNULL(T0.U_LunchS,0) As U_LunchS ,ISNULL(T0.U_Snack,0) As U_Snack ,ISNULL(T0.U_Dinner,0) As U_Dinner, ISNULL(T0.U_DinnerS,0) As U_DinnerS From [@Z_CPM4] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry "
'            strQuery += " JOIN [@Z_CPM1] T2 On Convert(VarChar(8),T0.U_PrgDate,112) = Convert(VarChar(8),T2.U_PrgDate,112)  "
'            strQuery += " Where Convert(VarChar(8),T0.U_PrgDate,112) >= '" + strFrmDate + "'"
'            strQuery += " And Convert(VarChar(8),T0.U_PrgDate,112) <= '" + strToDate + "'"
'            strQuery += " And T1.U_CardCode = '" + strCardCode + "'"
'            strQuery += " And T2.U_AppStatus = 'I' "
'            strQuery += " AND T1.U_RemDays > 0  "
'            If strType.Trim() = "I" Then
'                strQuery += " And T1.U_InvRef = '" + strRef + "'"
'            Else
'                strQuery += " And T1.U_TrnRef = '" + strRef + "'"
'            End If
'            oDataTable1.ExecuteQuery(strQuery)

'            'Selection Food Type
'            oDataTable2 = oForm.DataSources.DataTables.Item("Dt_SType")
'            strQuery = " Select T0.U_PrgDate,T0.U_BF,T0.U_Lunch,T0.U_LunchS,T0.U_Snack,T0.U_Dinner,T0.U_DinnerS From [@Z_CPM5] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry "
'            strQuery += " JOIN [@Z_CPM1] T2 On Convert(VarChar(8),T0.U_PrgDate,112) = Convert(VarChar(8),T2.U_PrgDate,112)  "
'            strQuery += " Where Convert(VarChar(8),T0.U_PrgDate,112) >= '" + strFrmDate + "'"
'            strQuery += " And Convert(VarChar(8),T0.U_PrgDate,112) <= '" + strToDate + "'"
'            strQuery += " And T1.U_CardCode = '" + strCardCode + "'"
'            strQuery += " And T2.U_AppStatus = 'I' "
'            strQuery += " AND T1.U_RemDays > 0  "
'            If strType.Trim() = "I" Then
'                strQuery += " And T1.U_InvRef = '" + strRef + "'"
'            Else
'                strQuery += " And T1.U_TrnRef = '" + strRef + "'"
'            End If
'            oDataTable2.ExecuteQuery(strQuery)

'            'Remarks
'            oDataTable3 = oForm.DataSources.DataTables.Item("Dt_SRemarks")
'            strQuery = " Select T0.U_PrgDate,Convert(VarChar(254),T0.U_BF) As U_BF ,Convert(VarChar(254),T0.U_Lunch) As U_Lunch,Convert(VarChar(254),T0.U_LunchS) As U_LunchS,Convert(VarChar(254),T0.U_Snack) As U_Snack,Convert(VarChar(254),T0.U_Dinner) As U_Dinner,Convert(VarChar(254),T0.U_DinnerS) As U_DinnerS From [@Z_CPM5] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry "
'            strQuery += " JOIN [@Z_CPM1] T2 On Convert(VarChar(8),T0.U_PrgDate,112) = Convert(VarChar(8),T2.U_PrgDate,112)  "
'            strQuery += " Where Convert(VarChar(8),T0.U_PrgDate,112) >= '" + strFrmDate + "'"
'            strQuery += " And Convert(VarChar(8),T0.U_PrgDate,112) <= '" + strToDate + "'"
'            strQuery += " And T1.U_CardCode = '" + strCardCode + "'"
'            strQuery += " And T2.U_AppStatus = 'I' "
'            strQuery += " AND T1.U_RemDays > 0  "
'            If strType.Trim() = "I" Then
'                strQuery += " And T1.U_InvRef = '" + strRef + "'"
'            Else
'                strQuery += " And T1.U_TrnRef = '" + strRef + "'"
'            End If
'            oDataTable3.ExecuteQuery(strQuery)

'            formatGrid(oForm)
'        Catch ex As Exception 
'.Log.Trace_DIET_AddOn_Error(ex)
'            Throw ex 
''oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
'        End Try
'    End Sub

'    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
'        Dim _retVal As Boolean = True
'        Try
'            oGrid = oForm.Items.Item("1").Specific
'            For intRow As Integer = 0 To oGrid.Rows.Count - 1
'                Dim strPrgDate As String = CDate(oGrid.DataTable.GetValue(0, intRow)).ToString("yyyyMMdd")
'                For intCol As Integer = 1 To oGrid.Columns.Count - 1
'                    Dim strValue As String = oGrid.DataTable.GetValue(oGrid.DataTable.Columns.Item(intCol).Name, intRow)
'                    If strValue.Trim().Length = 0 Then
'                        _retVal = False
'                        oApplication.Utilities.Message("Food Not Selected for Program Date : " + CDate(oGrid.DataTable.GetValue(0, intRow)).ToString("dd-MM-yyyy"), SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
'                        Return _retVal
'                    End If
'                Next
'            Next
'            Return _retVal
'        Catch ex As Exception 
'oApplication.Log.Trace_DIET_AddOn_Error(ex)
'            Throw ex 
''oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
'        End Try
'    End Function

'    Private Sub formatGrid(ByVal oForm As SAPbouiCOM.Form)
'        Try
'            oForm.Freeze(True)
'            oGrid = oForm.Items.Item("1").Specific
'            fillHeader(oForm, "1")
'            formatAll(oForm, oGrid, "1")
'            oForm.Freeze(False)
'        Catch ex As Exception 
'.Log.Trace_DIET_AddOn_Error(ex)
'            oForm.Freeze(False)
'            Throw ex 
''oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
'        End Try
'    End Sub

'    Private Sub formatAll(ByVal oForm As SAPbouiCOM.Form, ByVal oGrid As SAPbouiCOM.Grid, ByVal strID As String)
'        Try
'            oGrid.Columns.Item("U_PrgDate").TitleObject.Caption = "Program Date"
'            oGrid.Columns.Item("U_BF").TitleObject.Caption = "Break Fast"
'            oGrid.Columns.Item("U_Lunch").TitleObject.Caption = "Lunch"
'            oGrid.Columns.Item("U_LunchS").TitleObject.Caption = "Lunch-Side"
'            oGrid.Columns.Item("U_Snack").TitleObject.Caption = "Snack"
'            oGrid.Columns.Item("U_Dinner").TitleObject.Caption = "Dinner"
'            oGrid.Columns.Item("U_DinnerS").TitleObject.Caption = "Dinner-Side"
'        Catch ex As Exception 
'.Log.Trace_DIET_AddOn_Error(ex)
'            Throw ex 
' 'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
'        End Try
'    End Sub

'    Private Sub fillHeader(ByVal aForm As SAPbouiCOM.Form, ByVal strGridID As String)
'        Try
'            aForm.Freeze(True)
'            oGrid = aForm.Items.Item(strGridID).Specific
'            oGrid.RowHeaders.TitleObject.Caption = "#"
'            For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
'                oGrid.RowHeaders.SetText(index, (index + 1).ToString())
'            Next
'            aForm.Freeze(False)
'        Catch ex As Exception 
'.Log.Trace_DIET_AddOn_Error(ex)
'            aForm.Freeze(False)
'            Throw ex 
''oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
'        End Try
'    End Sub

'    Private Function getSelectedPrgDate(ByVal aForm As SAPbouiCOM.Form, ByVal strGridID As String) As String
'        Try
'            Dim _retVal As String = String.Empty
'            Dim blnSelected As Boolean = False
'            oGrid = aForm.Items.Item(strGridID).Specific
'            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
'                If oGrid.Rows.IsSelected(intRow) Then
'                    _retVal = CDate(oGrid.DataTable.GetValue(0, intRow)).ToString("yyyyMMdd")
'                    blnSelected = True
'                    CType(oForm.Items.Item("9").Specific, SAPbouiCOM.EditText).Value = intRow
'                End If
'            Next
'            If Not blnSelected Then
'                oApplication.Utilities.Message("No row selected", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
'            End If
'            Return _retVal
'        Catch ex As Exception 
'.Log.Trace_DIET_AddOn_Error(ex)
'            Throw ex 
''oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
'        End Try
'    End Function

'    Private Function getFoodType(ByVal oForm As SAPbouiCOM.Form, ByVal strColType As String) As String
'        Try
'            Dim _retVal As String = String.Empty
'            Select Case strColType
'                Case "U_BF"
'                    _retVal = "BF"
'                Case "U_Lunch"
'                    _retVal = "LN"
'                Case "U_LunchS"
'                    _retVal = "LS"
'                Case "U_Snack"
'                    _retVal = "SK"
'                Case "U_Dinner"
'                    _retVal = "DI"
'                Case "U_DinnerS"
'                    _retVal = "DS"
'            End Select
'            Return _retVal
'        Catch ex As Exception 
'.Log.Trace_DIET_AddOn_Error(ex)
'            Throw ex 
''oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
'        End Try
'    End Function

'    Private Function getQuantityBasedonCalories(ByVal oForm As SAPbouiCOM.Form, ByVal strColType As String) As Double
'        Try
'            Dim _retVal As Double = 0
'            Dim strCardCode As String = CType(oForm.Items.Item("10").Specific, SAPbouiCOM.EditText).Value
'            Dim oRecordSet As SAPbobsCOM.Recordset
'            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
'            Select Case strColType
'                Case "U_BF"
'                    strQuery = " Select ISNULL(U_BFactor,0) From [@Z_OCAJ] T0 JOIN [@Z_OCPR] T1 On T0.U_Calories = T1.U_CPAdj "
'                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
'                    oRecordSet.DoQuery(strQuery)
'                    If Not oRecordSet.EoF Then
'                        _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
'                    Else
'                        _retVal = 1
'                    End If
'                Case "U_Lunch"
'                    strQuery = " Select ISNULL(U_LFactor,0) From [@Z_OCAJ] T0 JOIN [@Z_OCPR] T1 On T0.U_Calories = T1.U_CPAdj "
'                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
'                    oRecordSet.DoQuery(strQuery)
'                    If Not oRecordSet.EoF Then
'                        _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
'                    Else
'                        _retVal = 1
'                    End If
'                Case "U_LunchS"
'                    strQuery = " Select ISNULL(U_LSFactor,0) From [@Z_OCAJ] T0 JOIN [@Z_OCPR] T1 On T0.U_Calories = T1.U_CPAdj "
'                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
'                    oRecordSet.DoQuery(strQuery)
'                    If Not oRecordSet.EoF Then
'                        _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
'                    Else
'                        _retVal = 1
'                    End If
'                Case "U_Snack"
'                    strQuery = " Select ISNULL(U_SFactor,0) From [@Z_OCAJ] T0 JOIN [@Z_OCPR] T1 On T0.U_Calories = T1.U_CPAdj "
'                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
'                    oRecordSet.DoQuery(strQuery)
'                    If Not oRecordSet.EoF Then
'                        _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
'                    Else
'                        _retVal = 1
'                    End If
'                Case "U_Dinner"
'                    strQuery = " Select ISNULL(U_DFactor,0) From [@Z_OCAJ] T0 JOIN [@Z_OCPR] T1 On T0.U_Calories = T1.U_CPAdj "
'                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
'                    oRecordSet.DoQuery(strQuery)
'                    If Not oRecordSet.EoF Then
'                        _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
'                    Else
'                        _retVal = 1
'                    End If
'                Case "U_DinnerS"
'                    strQuery = " Select ISNULL(U_DSFactor,0) From [@Z_OCAJ] T0 JOIN [@Z_OCPR] T1 On T0.U_Calories = T1.U_CPAdj "
'                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
'                    oRecordSet.DoQuery(strQuery)
'                    If Not oRecordSet.EoF Then
'                        _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
'                    Else
'                        _retVal = 1
'                    End If
'            End Select
'            Return _retVal
'        Catch ex As Exception 
'.Log.Trace_DIET_AddOn_Error(ex)
'            Throw ex 
''oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
'        End Try
'    End Function

'    Private Sub updateFoodPreference(ByVal aForm As SAPbouiCOM.Form, ByVal strGridID As String)
'        Try
'            oGrid = aForm.Items.Item(strGridID).Specific
'            For intRow As Integer = 0 To oGrid.Rows.Count - 1
'                Dim strPrgDate As String = CDate(oGrid.DataTable.GetValue(0, intRow)).ToString("yyyyMMdd")
'                Dim strCardCode As String = CType(oForm.Items.Item("10").Specific, SAPbouiCOM.EditText).Value
'                Dim strType As String = CType(oForm.Items.Item("11").Specific, SAPbouiCOM.EditText).Value
'                Dim strRef As String = CType(oForm.Items.Item("12").Specific, SAPbouiCOM.EditText).Value

'                Dim strBreakFast As String = oGrid.DataTable.GetValue(1, intRow).ToString()
'                Dim strLunch As String = oGrid.DataTable.GetValue(2, intRow).ToString()
'                Dim strLunchS As String = oGrid.DataTable.GetValue(3, intRow).ToString()
'                Dim strSnack As String = oGrid.DataTable.GetValue(4, intRow).ToString()
'                Dim strDinner As String = oGrid.DataTable.GetValue(5, intRow).ToString()
'                Dim strDinnerS As String = oGrid.DataTable.GetValue(6, intRow).ToString()
'                If strBreakFast.Trim.Length > 0 Or strLunch.Trim.Length > 0 Or strLunchS.Trim.Length > 0 _
'                    Or strSnack.Trim.Length > 0 Or strDinner.Trim.Length > 0 Or strDinnerS.Trim.Length > 0 Then
'                    oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
'                    strQuery = "Update T0 Set "
'                    strQuery += " U_BF = '" + strBreakFast + "'"
'                    strQuery += " ,U_Lunch = '" + strLunch + "'"
'                    strQuery += " ,U_LunchS  = '" + strLunchS + "'"
'                    strQuery += " ,U_Snack  = '" + strSnack + "'"
'                    strQuery += " ,U_Dinner  = '" + strDinner + "'"
'                    strQuery += " ,U_DinnerS  = '" + strDinnerS + "'"
'                    strQuery += " From [@Z_CPM1] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry "
'                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
'                    strQuery += " And Convert(VarChar(8),U_PrgDate,112) = '" + strPrgDate + "'"
'                    If strType.Trim() = "I" Then
'                        'strQuery += " And T1.U_Type = 'I'  "
'                        strQuery += " And T1.U_InvRef = '" + strRef + "'"
'                    Else
'                        'strQuery += " And T1.U_Type = 'I'  "
'                        strQuery += " And T1.U_TrnRef = '" + strRef + "'"
'                    End If
'                    oRecordSet.DoQuery(strQuery)
'                End If
'            Next

'            oDT_Quantity = oForm.DataSources.DataTables.Item("Dt_Quantity")
'            For intRow As Integer = 0 To oDT_Quantity.Rows.Count - 1
'                Dim strPrgDate As String = CDate(oDT_Quantity.GetValue(0, intRow)).ToString("yyyyMMdd")
'                Dim strCardCode As String = CType(oForm.Items.Item("10").Specific, SAPbouiCOM.EditText).Value
'                Dim strType As String = CType(oForm.Items.Item("11").Specific, SAPbouiCOM.EditText).Value
'                Dim strRef As String = CType(oForm.Items.Item("12").Specific, SAPbouiCOM.EditText).Value

'                Dim dblBreakFast As Double = CDbl(oDT_Quantity.GetValue(1, intRow).ToString())
'                Dim dblLunch As Double = CDbl(oDT_Quantity.GetValue(2, intRow).ToString())
'                Dim dblLunchS As Double = CDbl(oDT_Quantity.GetValue(3, intRow).ToString())
'                Dim dblSnack As Double = CDbl(oDT_Quantity.GetValue(4, intRow).ToString())
'                Dim dblDinner As Double = CDbl(oDT_Quantity.GetValue(5, intRow).ToString())
'                Dim dblDinnerS As Double = CDbl(oDT_Quantity.GetValue(6, intRow).ToString())
'                If dblBreakFast > 0 Or dblLunch > 0 Or dblLunchS > 0 _
'                    Or dblSnack > 0 Or dblDinner > 0 Or dblDinnerS > 0 Then
'                    oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
'                    strQuery = "Update T0 Set "
'                    strQuery += " U_BF = " + dblBreakFast.ToString() + ""
'                    strQuery += " ,U_Lunch = " + dblLunch.ToString() + ""
'                    strQuery += " ,U_LunchS  = " + dblLunchS.ToString() + ""
'                    strQuery += " ,U_Snack  = " + dblSnack.ToString() + ""
'                    strQuery += " ,U_Dinner  = " + dblDinner.ToString() + ""
'                    strQuery += " ,U_DinnerS  = " + dblDinnerS.ToString() + ""
'                    strQuery += " From [@Z_CPM4] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry "
'                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
'                    strQuery += " And Convert(VarChar(8),U_PrgDate,112) = '" + strPrgDate + "'"
'                    If strType.Trim() = "I" Then
'                        strQuery += " And T1.U_InvRef = '" + strRef + "'"
'                    Else
'                        strQuery += " And T1.U_TrnRef = '" + strRef + "'"
'                    End If
'                    oRecordSet.DoQuery(strQuery)
'                End If
'            Next

'            oDT_SType = oForm.DataSources.DataTables.Item("Dt_SType")
'            For intRow As Integer = 0 To oDT_SType.Rows.Count - 1
'                Dim strPrgDate As String = CDate(oDT_SType.GetValue(0, intRow)).ToString("yyyyMMdd")
'                Dim strCardCode As String = CType(oForm.Items.Item("10").Specific, SAPbouiCOM.EditText).Value
'                Dim strType As String = CType(oForm.Items.Item("11").Specific, SAPbouiCOM.EditText).Value
'                Dim strRef As String = CType(oForm.Items.Item("12").Specific, SAPbouiCOM.EditText).Value

'                Dim dblBreakFast As String = oDT_SType.GetValue(1, intRow).ToString()
'                Dim dblLunch As String = oDT_SType.GetValue(2, intRow).ToString()
'                Dim dblLunchS As String = oDT_SType.GetValue(3, intRow).ToString()
'                Dim dblSnack As String = oDT_SType.GetValue(4, intRow).ToString()
'                Dim dblDinner As String = oDT_SType.GetValue(5, intRow).ToString()
'                Dim dblDinnerS As String = oDT_SType.GetValue(6, intRow).ToString()

'                If dblBreakFast <> "" Or dblLunch <> "" Or dblLunchS <> "" _
'                    Or dblSnack <> "" Or dblDinner <> "" Or dblDinnerS <> "" Then
'                    oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
'                    strQuery = "Update T0 Set "
'                    strQuery += " U_BF = '" + dblBreakFast + "'"
'                    strQuery += " ,U_Lunch = '" + dblLunch + "'"
'                    strQuery += " ,U_LunchS  = '" + dblLunchS + "'"
'                    strQuery += " ,U_Snack  = '" + dblSnack + "'"
'                    strQuery += " ,U_Dinner  = '" + dblDinner + "'"
'                    strQuery += " ,U_DinnerS  = '" + dblDinnerS + "'"
'                    strQuery += " From [@Z_CPM5] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry "
'                    strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
'                    strQuery += " And Convert(VarChar(8),U_PrgDate,112) = '" + strPrgDate + "'"
'                    If strType.Trim() = "I" Then
'                        strQuery += " And T1.U_InvRef = '" + strRef + "'"
'                    Else
'                        strQuery += " And T1.U_TrnRef = '" + strRef + "'"
'                    End If
'                    oRecordSet.DoQuery(strQuery)
'                End If
'            Next

'        Catch ex As Exception 
'.Log.Trace_DIET_AddOn_Error(ex)
'            Throw ex 
''oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
'        End Try
'    End Sub

'End Class