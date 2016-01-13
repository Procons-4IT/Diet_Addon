Imports SAPbobsCOM

Public Class clsProgramTransfer

    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private InvForConsumedItems, count As Integer
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines As SAPbouiCOM.DBDataSource
    Public intSelectedMatrixrow As Integer = 0
    Private RowtoDelete As Integer
    Private MatrixId As String
    Private oRecordSet As SAPbobsCOM.Recordset
    Private dtValidFrom, dtValidTo As Date
    Private strQuery As String

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Public Sub LoadForm()
        Try
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_OPGT) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_OPGT, frm_Z_OPGT)
            oForm = oApplication.SBO_Application.Forms.Item(strUID)
            oForm.Freeze(True)
            oForm.DataBrowser.BrowseBy = "10"
            initialize(oForm)
            addChooseFromListConditions(oForm)
            'oForm.EnableMenu(mnu_ADD_ROW, False)
            'oForm.EnableMenu(mnu_DELETE_ROW, False)
            oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

    Public Sub LoadForm(ByVal strDocEntry As String)
        Try
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_OPGT) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_OPGT, frm_Z_OPGT)
            oForm = oApplication.SBO_Application.Forms.Item(strUID)
            oForm.Freeze(True)
            initialize(oForm)
            addChooseFromListConditions(oForm)
            oForm.DataBrowser.BrowseBy = "10"
            'oForm.EnableMenu(mnu_ADD_ROW, False)
            'oForm.EnableMenu(mnu_DELETE_ROW, False)
            oForm.Freeze(False)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            oForm.Items.Item("10").Specific.value = strDocEntry
            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Z_OPGT Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        Else
                                            If validation(oForm) = False Then
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPGT")
                                If pVal.ItemUID = "3" Then
                                    intSelectedMatrixrow = pVal.Row
                                    If (oDBDataSource.GetValue("U_Code", 0).ToString() = "") Then
                                        BubbleEvent = False
                                        oApplication.SBO_Application.SetStatusBarMessage("Enter Dislike Code to Proceed...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "8" Then
                                    Dim strCardCode As String = oForm.Items.Item("4").Specific.value
                                    filterProgramChooseFromList(oForm, "CFL_5", strCardCode)
                                ElseIf pVal.ItemUID = "34" Then
                                    Dim strCardCode As String = oForm.Items.Item("4").Specific.value
                                    filterProgramChooseFromList(oForm, "CFL_6", strCardCode)
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
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_ADD_ROW)
                                    Case "15"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPGT")
                                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PGT1")
                                    oMatrix = oForm.Items.Item("3").Specific
                                    oMatrix.FlushToDataSource()
                                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                    Dim oDataTable As SAPbouiCOM.DataTable
                                    Try
                                        oCFLEvento = pVal
                                        oDataTable = oCFLEvento.SelectedObjects

                                        If pVal.ItemUID = "4" Then
                                            If IsNothing(oDataTable) Then
                                                Exit Sub
                                            End If
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 0 Then
                                                oDBDataSource.SetValue("U_CardCode", 0, oDataTable.GetValue("CardCode", 0))
                                                oDBDataSource.SetValue("U_CardName", 0, oDataTable.GetValue("CardName", 0))
                                                If oDBDataSource.GetValue("U_TrnType", oDBDataSource.Offset).ToString().Trim() = "P" Then
                                                    oDBDataSource.SetValue("U_TCardCode", 0, oDataTable.GetValue("CardCode", 0))
                                                    oDBDataSource.SetValue("U_TCardName", 0, oDataTable.GetValue("CardName", 0))
                                                Else
                                                    oDBDataSource.SetValue("U_TCardCode", 0, "")
                                                    oDBDataSource.SetValue("U_TCardName", 0, "")
                                                End If
                                            End If
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            'GetProgramDetails(oForm)
                                        ElseIf pVal.ItemUID = "5" Then
                                            If IsNothing(oDataTable) Then
                                                Exit Sub
                                            End If
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 0 Then
                                                oDBDataSource.SetValue("U_CardCode", 0, oDataTable.GetValue("CardCode", 0))
                                                oDBDataSource.SetValue("U_CardName", 0, oDataTable.GetValue("CardName", 0))
                                            End If
                                            If oDBDataSource.GetValue("U_TrnType", oDBDataSource.Offset).ToString().Trim() = "P" Then
                                                oDBDataSource.SetValue("U_TCardCode", 0, oDataTable.GetValue("CardCode", 0))
                                                oDBDataSource.SetValue("U_TCardName", 0, oDataTable.GetValue("CardName", 0))
                                            Else
                                                oDBDataSource.SetValue("U_TCardCode", 0, "")
                                                oDBDataSource.SetValue("U_TCardName", 0, "")
                                            End If
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            'GetProgramDetails(oForm)
                                        ElseIf pVal.ItemUID = "6" Then
                                            If IsNothing(oDataTable) Then
                                                Exit Sub
                                            End If
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 0 Then
                                                oDBDataSource.SetValue("U_TCardCode", 0, oDataTable.GetValue("CardCode", 0))
                                                oDBDataSource.SetValue("U_TCardName", 0, oDataTable.GetValue("CardName", 0))
                                            End If
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf pVal.ItemUID = "7" Then
                                            If IsNothing(oDataTable) Then
                                                Exit Sub
                                            End If
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 0 Then
                                                oDBDataSource.SetValue("U_TCardCode", 0, oDataTable.GetValue("CardCode", 0))
                                                oDBDataSource.SetValue("U_TCardName", 0, oDataTable.GetValue("CardName", 0))
                                            End If
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf pVal.ItemUID = "8" Then
                                            If IsNothing(oDataTable) Then
                                                Exit Sub
                                            End If
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 0 Then
                                                oDBDataSource.SetValue("U_ProgramID", 0, oDataTable.GetValue("DocEntry", 0))
                                                GetFProgramDetails(oForm, oDataTable.GetValue("DocEntry", 0))
                                            End If
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf pVal.ItemUID = "34" Then
                                            If IsNothing(oDataTable) Then
                                                Exit Sub
                                            End If
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 0 Then
                                                oDBDataSource.SetValue("U_TProgramID", 0, oDataTable.GetValue("DocEntry", 0))
                                                GetTProgramDetails(oForm, oDataTable.GetValue("DocEntry", 0))
                                            End If
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf pVal.ItemUID = "40" Then
                                            If IsNothing(oDataTable) Then
                                                Exit Sub
                                            End If
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 0 Then
                                                oDBDataSource.SetValue("U_TPrgCode", 0, oDataTable.GetValue("ItemCode", 0))
                                                oDBDataSource.SetValue("U_TPrgName", 0, oDataTable.GetValue("ItemName", 0))
                                                'GetTProgramDetails(oForm, oDataTable.GetValue("DocEntry", 0))
                                            End If
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                    Catch ex As Exception
                                        oApplication.Log.Trace_DIET_AddOn_Error(ex)

                                    End Try
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Z_OPGT
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then

                    End If
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        AddRow(oForm)
                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        RefereshDeleteRow(oForm)
                    End If
                Case mnu_ADD
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        initialize(oForm)
                        EnableControls(oForm, True)
                    End If
                Case mnu_FIND
                    If pVal.BeforeAction = False Then

                    End If
            End Select
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Data Events"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            If oForm.TypeEx = frm_Z_OPGT Then
                Select Case BusinessObjectInfo.BeforeAction
                    Case True

                    Case False
                        Select Case BusinessObjectInfo.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                EnableControls(oForm, False)
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If BusinessObjectInfo.ActionSuccess Then
                                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                                    Dim oXmlDoc As System.Xml.XmlDocument = New Xml.XmlDocument()
                                    oXmlDoc.LoadXml(BusinessObjectInfo.ObjectKey)
                                    Dim DocEntry As String = oXmlDoc.SelectSingleNode("/Program_TransferParams/DocEntry").InnerText
                                    oApplication.Company.StartTransaction()
                                    Try
                                        If oApplication.Utilities.updateProgramDocument(DocEntry) Then
                                            If oApplication.Utilities.AddTransferProgram(oForm, DocEntry) Then
                                                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                            Else
                                                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                            End If
                                        End If
                                    Catch ex As Exception
                                        oApplication.Log.Trace_DIET_AddOn_Error(ex)
                                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                        oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End Try
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            If oForm.TypeEx = frm_Z_OPGT Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data
                If (eventInfo.BeforeAction = True) Then
                    'oMenuItem.SubMenus.Item(mnu_ADD_ROW).Enabled = False
                    'oMenuItem.SubMenus.Item(mnu_DELETE_ROW).Enabled = False
                    'oMenuItem.SubMenus.Item(mnu_CANCEL).Enabled = False
                    'oMenuItem.SubMenus.Item(mnu_CLOSE).Enabled = False
                Else
                    'oMenuItem.SubMenus.Item(mnu_ADD_ROW).Enabled = True
                    'oMenuItem.SubMenus.Item(mnu_DELETE_ROW).Enabled = True
                    'oMenuItem.SubMenus.Item(mnu_CANCEL).Enabled = True
                    'oMenuItem.SubMenus.Item(mnu_CLOSE).Enabled = True
                End If
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPGT")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PGT1")

            oMatrix = oForm.Items.Item("3").Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select IsNull(MAX(DocEntry),0) +1 From [@Z_OPGT]")
            If Not oRecordSet.EoF Then
                oApplication.Utilities.setEditText(oForm, "9", oRecordSet.Fields.Item(0).Value.ToString())
                oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                'oApplication.Utilities.setEditText(oForm, "11", "t")
                'oApplication.SBO_Application.SendKeys("{TAB}")
            End If

            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            MatrixId = "3"
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

#Region "Methods"
    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("3").Specific
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PGT1")
            oMatrix.FlushToDataSource()
            For count = 1 To oDBDataSourceLines.Size
                oDBDataSourceLines.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            aForm.Freeze(False)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub
#End Region

    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Select Case aForm.PaneLevel
                Case "0"
                    oMatrix = aForm.Items.Item("3").Specific
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PGT1")
                    If oMatrix.RowCount <= 0 Then
                        oMatrix.AddRow()
                    Else
                        If oApplication.Utilities.getMatrixValues(oMatrix, "V_0", oMatrix.RowCount) <> "" Then


                            oMatrix.AddRow(1, oMatrix.RowCount + 1)
                            oMatrix.ClearRowData(oMatrix.RowCount)
                            If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            End If
                        End If
                    End If
                    oMatrix.FlushToDataSource()
                    For count = 1 To oDBDataSourceLines.Size
                        oDBDataSourceLines.SetValue("LineId", count - 1, count)
                    Next
                    oMatrix.LoadFromDataSource()
                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    AssignLineNo(aForm)
            End Select
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            aForm.Freeze(False)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub deleterow(ByVal aForm As SAPbouiCOM.Form)
        Try
            Select Case aForm.PaneLevel
                Case "0"
                    oMatrix = aForm.Items.Item("11").Specific
                    oDBDataSourceLines = aForm.DataSources.DBDataSources.Item("@Z_PGT1")
            End Select
            oMatrix.FlushToDataSource()
            For introw As Integer = 1 To oMatrix.RowCount
                If oMatrix.IsRowSelected(introw) Then
                    oMatrix.DeleteRow(introw)
                    oDBDataSourceLines.RemoveRecord(introw - 1)
                    oMatrix.FlushToDataSource()
                    For count As Integer = 1 To oDBDataSourceLines.Size
                        oDBDataSourceLines.SetValue("LineId", count - 1, count)
                    Next
                    Select Case aForm.PaneLevel
                        Case "0"
                            oMatrix = aForm.Items.Item("3").Specific
                            oDBDataSourceLines = aForm.DataSources.DBDataSources.Item("@Z_PGT1")
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
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            aForm.Freeze(False)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

#Region "Delete Row"
    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            oMatrix = aForm.Items.Item("3").Specific
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPGT")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PGT1")

            If Me.MatrixId = "3" Then
                oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PGT1")
            End If

            Me.RowtoDelete = intSelectedMatrixrow
            oDBDataSourceLines.RemoveRecord(Me.RowtoDelete - 1)
            oMatrix.LoadFromDataSource()
            oMatrix.FlushToDataSource()
            For count = 0 To oDBDataSourceLines.Size - 1
                oDBDataSourceLines.SetValue("LineId", count, count + 1)
            Next
            oMatrix.LoadFromDataSource()
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            aForm.Freeze(False)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub
#End Region

#Region "Validations"
    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            oMatrix = oForm.Items.Item("3").Specific
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPGT")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PGT1")
            Dim strTCardCode As String = oDBDataSource.GetValue("U_TCardCode", 0).Trim()
            Dim strFrDt As String = oDBDataSource.GetValue("U_PFromDate", 0).Trim()
            Dim strToDt As String = oDBDataSource.GetValue("U_PToDate", 0).Trim()
            Dim strToType As String = oDBDataSource.GetValue("U_TrnType", 0).Trim()
            Dim strProgramID As String = oDBDataSource.GetValue("U_ProgramID", 0).Trim()

            Dim intNoDays As Integer = CInt(IIf(oDBDataSource.GetValue("U_NoOfDays", 0).Trim() = "", 0, oDBDataSource.GetValue("U_NoOfDays", 0).Trim()))
            Dim intTNoDays As Integer = CInt(IIf(oDBDataSource.GetValue("U_TNoOfDays", 0).Trim() = "", 0, oDBDataSource.GetValue("U_TNoOfDays", 0).Trim()))
            Dim intOPDays As Integer = CInt(IIf(strToDt.Trim() = "", 0, strToDt.Trim())) - CInt(IIf(strFrDt.Trim() = "", 0, strFrDt.Trim())) + 1

            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If oApplication.Utilities.getEditTextvalue(aForm, "4") = "" Then
                oApplication.Utilities.Message("Enter From Customer...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "6") = "" Then
                oApplication.Utilities.Message("Enter To Customer...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "8") = "" Then
                oApplication.Utilities.Message("Enter Program ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf intNoDays <= 0 Then
                oApplication.Utilities.Message("No of Day Should be Greater than Zero ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "4") = oApplication.Utilities.getEditTextvalue(aForm, "6") And strToType = "C" Then
                oApplication.Utilities.Message("To Customer Should be different from Parent Customer  ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "16") = "" Then
                oApplication.Utilities.Message("Enter Program From Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                'ElseIf oApplication.Utilities.getEditTextvalue(aForm, "17") = "" Then
                '    oApplication.Utilities.Message("Enter Program To ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Return False
                'ElseIf CInt(strFrDt) > CInt(strToDt) Then
                '    oApplication.Utilities.Message("Program From Date Should be Greater than or Equal Program To Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Return False
                'ElseIf intOPDays < intNoDays Then
                '    oApplication.Utilities.Message("Difference in Program Days Should be Greater than No of Days ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "4") <> oApplication.Utilities.getEditTextvalue(aForm, "6") And strToType = "P" Then
                oApplication.Utilities.Message("To Customer Should be Equal from Parent Customer ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "_18") = oApplication.Utilities.getEditTextvalue(aForm, "39") And strToType = "P" Then
                oApplication.Utilities.Message("Cannot have Same From/To Program for Transfer to Program Type...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "39") = "" And strToType = "P" Then
                oApplication.Utilities.Message("Select To Program to Proceed ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            strQuery = " Select ISNULL(U_Cancel,'N') As Cancel From [@Z_OCPM] "
            strQuery += " Where DocEntry = '" + strProgramID + "'"
            If Not oRecordSet.EoF Then
                If oRecordSet.Fields.Item(0).Value.ToString() = "Y" Then
                    oApplication.Utilities.Message("Program Selected Already Cancelled...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If

            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            strQuery = " Select (ISNULL(U_OrdDays,0) - ISNULL(U_DelDays,0)) From [@Z_OCPM] "
            strQuery += " Where DocEntry = '" + strProgramID + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                Dim intOpenOrder As Integer = CInt(oRecordSet.Fields.Item(0).Value)
                If intOpenOrder > 0 Then
                    oApplication.Utilities.Message("Open Order Document Exist for Selected Program ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If


            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select DocEntry from [@Z_OCPM] where U_CardCode='" & oApplication.Utilities.getEditTextvalue(aForm, "6") & "'" & _
                " And '" & strFrDt & "' between Convert(VarChar(8),U_PFromDate,112) and Convert(VarChar(8),U_PToDate,112) And IsNull(U_Cancel,'N') = 'N' " & _
                " And ISNULL(U_Transfer,'N') = 'N' " & _
                    " And ISNULL(U_DocStatus,'O') = 'O' " & _
                    " And DocEntry <> '" & strProgramID & "'"
            oTest.DoQuery(strQuery)
            If oTest.RecordCount > 0 Then
                oApplication.Utilities.Message("Program From date is overlapped with another program for selected customer...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Dim strPrgToDate As String = oApplication.Utilities.getProgramToDate(oForm, strTCardCode, strFrDt _
                                                                                 , (intTNoDays).ToString)
            strQuery = "Select DocEntry from [@Z_OCPM] where U_CardCode='" & oApplication.Utilities.getEditTextvalue(aForm, "6") & "'" & _
               " And '" & strPrgToDate & "' between Convert(VarChar(8),U_PFromDate,112) and Convert(VarChar(8),U_PToDate,112) And IsNull(U_Cancel,'N') = 'N' " & _
               " And ISNULL(U_Transfer,'N') = 'N' " & _
                   " And ISNULL(U_DocStatus,'O') = 'O' " & _
            " And DocEntry <> '" & strProgramID & "'"
            oTest.DoQuery(strQuery)
            If oTest.RecordCount > 0 Then
                oApplication.Utilities.Message("Program To date is overlapped with another program for selected customer...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            aForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#End Region

#Region "Disable Controls"

    Private Sub EnableControls(ByVal oForm As SAPbouiCOM.Form, ByVal blnEnable As Boolean)
        Try
            oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("6").Enabled = blnEnable
            oForm.Items.Item("7").Enabled = blnEnable

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

#End Region

    Private Sub GetProgramDetails(ByVal oForm As SAPbouiCOM.Form)
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet, oRecordSet1 As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet1 = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            strQuery = "Select DocEntry,DocNum,U_RemDays,U_PrgCode From [@Z_OCPM] Where U_CardCode = '" + oForm.Items.Item("4").Specific.value + "' And U_RemDays > 0 "
            oRecordSet.DoQuery(strQuery)

            If Not oRecordSet.EoF Then
                oDBDataSource.SetValue("U_ProgramID", 0, oRecordSet.Fields.Item(0).Value.ToString())
                oDBDataSource.SetValue("U_NoOfDays", 0, oRecordSet.Fields.Item(2).Value.ToString())
                oDBDataSource.SetValue("U_PrgCode", 0, oRecordSet.Fields.Item(3).Value.ToString())

                'strQuery = "Select DocEntry,LineID,Convert(VarChar(8),U_PrgDate,112) As U_PrgDate From [@Z_CPM1] Where DocEntry = '" + oRecordSet.Fields.Item(0).Value.ToString() + "' And U_Status = 'O' "
                'oRecordSet1.DoQuery(strQuery)
                'If Not oRecordSet1.EoF Then
                '    oMatrix.Clear()
                '    oMatrix.FlushToDataSource()
                '    Dim intRow As Integer = 0
                '    While Not oRecordSet1.EoF
                '        oMatrix.AddRow(1, oMatrix.RowCount)
                '        oMatrix.FlushToDataSource()
                '        oDBDataSourceLines.SetValue("U_PrgDate", intRow, oRecordSet1.Fields.Item("U_PrgDate").Value)
                '        oDBDataSourceLines.SetValue("U_PrgNo", intRow, oRecordSet1.Fields.Item("DocEntry").Value)
                '        oDBDataSourceLines.SetValue("U_PrgLine", intRow, oRecordSet1.Fields.Item("LineID").Value)
                '        oMatrix.LoadFromDataSource()
                '        oMatrix.FlushToDataSource()
                '        intRow += 1
                '        oRecordSet1.MoveNext()
                '    End While
                '    oMatrix.LoadFromDataSource()
                '    oMatrix.FlushToDataSource()
                'End If
            End If

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub filterProgramChooseFromList(ByVal oForm As SAPbouiCOM.Form, ByVal strCFLID As String, ByVal strCardCode As String)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oCFLs = oForm.ChooseFromLists
            oCFL = oCFLs.Item(strCFLID)
            oCons = oCFL.GetConditions()
            If oCons.Count > 0 Then

                oCons.Item(0).Alias = "U_CardCode"
                oCons.Item(0).Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCons.Item(0).CondVal = strCardCode

                oCons.Item(1).[Alias] = "U_RemDays"
                oCons.Item(1).Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_THAN
                oCons.Item(1).CondVal = "0"

                oCFL.SetConditions(oCons)
            Else
                oCon = oCons.Add()
                oCon.BracketOpenNum = 2
                oCon.[Alias] = "U_CardCode"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = strCardCode
                oCon.BracketCloseNum = 1
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                oCon = oCons.Add()
                oCon.BracketOpenNum = 1
                oCon.[Alias] = "U_RemDays"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_THAN
                oCon.CondVal = "0"
                oCon.BracketCloseNum = 2
                oCFL.SetConditions(oCons)
            End If
            'MessageBox.Show(oCFL.GetConditions().GetAsXML())
            'strQuery = "Select DocEntry From [@Z_OCPM] Where U_CardCode = '" & strCardCode & "' AND U_RemDays > 0 "
            'oRecordSet.DoQuery(strQuery)
            'If Not oRecordSet.EoF Then
            '    oCFL = oCFLs.Item(strCFLID)
            '    oCons = oCFL.GetConditions()

            '    oCon = oCons.Add()
            '    oCon.BracketOpenNum = 2
            '    Dim intConCount As Integer = 0
            '    While Not oRecordSet.EoF
            '        Dim strDE As String = oRecordSet.Fields.Item(0).Value
            '        If intConCount > 0 Then
            '            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
            '            oCon = oCons.Add()
            '            oCon.BracketOpenNum = 1
            '        End If
            '        oCon.[Alias] = "DocEntry"
            '        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '        oCon.CondVal = strDE

            '        oRecordSet.MoveNext()
            '        If Not oRecordSet.EoF Then
            '            oCon.BracketCloseNum = 1
            '        End If

            '        intConCount += 1
            '    End While

            '    oCon.BracketCloseNum = 2
            '    oCFL.SetConditions(oCons)
            'End If

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub GetFProgramDetails(ByVal oForm As SAPbouiCOM.Form, ByVal strFProgramID As String)
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet, oRecordSet1 As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet1 = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            strQuery = "Select DocEntry,DocNum,U_RemDays,U_PrgCode,U_PrgName From [@Z_OCPM] "
            strQuery += " Where U_CardCode = '" + oForm.Items.Item("4").Specific.value + "' And U_RemDays > 0 "
            strQuery += " AND DocEntry ='" + strFProgramID + "'"
            oRecordSet.DoQuery(strQuery)

            If Not oRecordSet.EoF Then
                oDBDataSource.SetValue("U_ProgramID", 0, oRecordSet.Fields.Item(0).Value.ToString())
                oDBDataSource.SetValue("U_NoOfDays", 0, oRecordSet.Fields.Item(2).Value.ToString())
                oDBDataSource.SetValue("U_PrgCode", 0, oRecordSet.Fields.Item(3).Value.ToString())
                oDBDataSource.SetValue("U_PrgName", 0, oRecordSet.Fields.Item(4).Value.ToString())
                oDBDataSource.SetValue("U_TNoOfDays", 0, oRecordSet.Fields.Item(2).Value.ToString())
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub GetTProgramDetails(ByVal oForm As SAPbouiCOM.Form, ByVal strTProgramID As String)
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet, oRecordSet1 As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet1 = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            strQuery = "Select DocEntry,DocNum,U_RemDays,U_PrgCode From [@Z_OCPM] "
            strQuery += " Where U_CardCode = '" + oForm.Items.Item("4").Specific.value + "' And U_RemDays > 0 "
            strQuery += " AND DocEntry ='" + strTProgramID + "'"
            oRecordSet.DoQuery(strQuery)

            If Not oRecordSet.EoF Then
                oDBDataSource.SetValue("U_TProgramID", 0, oRecordSet.Fields.Item(0).Value.ToString())
                oDBDataSource.SetValue("U_TNoOfDays", 0, oRecordSet.Fields.Item(2).Value.ToString())
                oDBDataSource.SetValue("U_TPrgCode", 0, oRecordSet.Fields.Item(3).Value.ToString())
            End If
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

            oCFLs = oForm.ChooseFromLists

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

            oCFL = oCFLs.Item("CFL_2")
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

            oCFL = oCFLs.Item("CFL_3")
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


            'strQuery = "Select ItmsGrpCod From OITB Where U_Program = 'Y' "
            'oRecordSet.DoQuery(strQuery)
            'If Not oRecordSet.EoF Then
            '    Dim strIG As String = oRecordSet.Fields.Item(0).Value
            '    oCFL = oCFLs.Item("CFL_7")
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
                oCFL = oCFLs.Item("CFL_7")
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

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

#End Region

End Class
