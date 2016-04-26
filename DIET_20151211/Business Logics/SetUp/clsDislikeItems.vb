Public Class clsDislikeItem

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
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_ODLK) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_ODLK, frm_Z_ODLK)
            oForm = oApplication.SBO_Application.Forms.Item(strUID)
            addChooseFromListConditions(oForm)
            oForm.Freeze(True)
            oForm.DataBrowser.BrowseBy = "16"
            initialize(oForm)
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            EnableControls(oForm, True)
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
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_ODLK) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_ODLK, frm_Z_ODLK)
            oForm = oApplication.SBO_Application.Forms.Item(strUID)
            oForm.Freeze(True)
            initialize(oForm)
            oForm.DataBrowser.BrowseBy = "16"
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            oForm.Freeze(False)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            oForm.Items.Item("16").Specific.value = strDocEntry
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
            If pVal.FormTypeEx = frm_Z_ODLK Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        If validation(oForm) = False Then
                                            BubbleEvent = False
                                            Exit Sub
                                        Else
                                            If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_ODLK")
                                If pVal.ItemUID = "3" Then
                                    intSelectedMatrixrow = pVal.Row
                                    If (oDBDataSource.GetValue("U_Code", 0).ToString() = "") Then
                                        BubbleEvent = False
                                        oApplication.SBO_Application.SetStatusBarMessage("Enter Dislike Code to Proceed...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
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
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_ADD_ROW)
                                    Case "15"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                                    Case "1"
                                        If pVal.Action_Success Then
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                                initialize(oForm)
                                            End If
                                        End If

                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_ODLK")
                                oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_DLK1")
                                oMatrix = oForm.Items.Item("3").Specific
                                oMatrix.FlushToDataSource()
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects

                                    If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If (pVal.ItemUID = "3" And (pVal.ColUID = "V_0" Or pVal.ColUID = "V_1")) Then
                                            oMatrix = oForm.Items.Item("3").Specific
                                            oMatrix.LoadFromDataSource()
                                            If Not IsNothing(oDataTable) Then
                                                Dim intAddRows As Integer = oDataTable.Rows.Count
                                                If intAddRows > 1 Then
                                                    intAddRows -= 1
                                                    oMatrix.AddRow(intAddRows, pVal.Row - 1)
                                                End If
                                                oMatrix.FlushToDataSource()
                                                For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                    oDBDataSourceLines.SetValue("LineId", pVal.Row + index - 1, (pVal.Row + index).ToString())
                                                    oDBDataSourceLines.SetValue("U_ItemCode", pVal.Row + index - 1, oDataTable.GetValue("ItemCode", index))
                                                    oDBDataSourceLines.SetValue("U_ItemName", pVal.Row + index - 1, oDataTable.GetValue("ItemName", index))
                                                Next
                                                oMatrix.LoadFromDataSource()
                                                oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            End If
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                    End If
                                Catch ex As Exception
                                    oApplication.Log.Trace_DIET_AddOn_Error(ex)

                                End Try
                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_ODLK")
                                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_DLK1")
                                    oMatrix = oForm.Items.Item("3").Specific
                                    oMatrix.FlushToDataSource()
                                    If (pVal.ItemUID = "3" And (pVal.ColUID = "V_0") And pVal.Row > 0) Then
                                        Dim strItemCode As String = oDBDataSourceLines.GetValue("U_ItemCode", pVal.Row - 1)
                                        Dim strItemName As String = oDBDataSourceLines.GetValue("U_ItemName", pVal.Row - 1)
                                        If strItemCode.Trim().Length = 0 And strItemName.Trim().Length > 0 Then
                                            oDBDataSourceLines.SetValue("U_ItemName", pVal.Row - 1, "")
                                            oMatrix.LoadFromDataSource()
                                            oMatrix.FlushToDataSource()
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                    End If
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
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Z_ODLK
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then

                    End If
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        AddRow(oForm)
                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        RefereshDeleteRow(oForm)
                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                Case mnu_ADD
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        initialize(oForm)
                        EnableControls(oForm, True)
                        oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oForm.Update()
                        oForm.Refresh()
                    End If
                Case mnu_FIND
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        EnableControls(oForm, True)
                        oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oForm.Update()
                        oForm.Refresh()
                    End If
                Case mnu_Remove
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = True Then
                        If oApplication.SBO_Application.MessageBox("This action will delete the current document. Are you sure you want to proceed?", , "Yes", "No") = 2 Then
                            BubbleEvent = False
                            Exit Sub
                        ElseIf oApplication.Utilities.ValidateRemoveSetup(oApplication.Utilities.getEditTextvalue(oForm, "6"), "DisLike") = False Then
                            oApplication.Utilities.Message("Dislike Code : " & oApplication.Utilities.getEditTextvalue(oForm, "6") & " already mapped to Transacton / Customer.", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
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
        If oForm.TypeEx = frm_Z_ODLK Then
            Dim oMenuItem As SAPbouiCOM.MenuItem
            oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data
            If (eventInfo.BeforeAction = True) Then
                Try
                    If oMenuItem.SubMenus.Exists(mnu_Remove) Then
                        oMenuItem.SubMenus.Item(mnu_Remove).String = "Delete Document"
                    End If
                Catch ex As Exception
                    oApplication.Log.Trace_DIET_AddOn_Error(ex)
                    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try
            End If
        End If
    End Sub

#End Region

#Region "Data Events"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            If oForm.TypeEx = frm_Z_ODLK Then
                Select Case BusinessObjectInfo.BeforeAction
                    Case True

                    Case False
                        Select Case BusinessObjectInfo.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                EnableControls(oForm, False)
                                oMatrix = oForm.Items.Item("3").Specific
                                oMatrix.AutoResizeColumns()
                                oMatrix.Columns.Item("V_2").Width = 300
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_ODLK")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_DLK1")

            oMatrix = oForm.Items.Item("3").Specific
            oMatrix.LoadFromDataSource()
            oMatrix.AddRow(1, -1)
            oMatrix.FlushToDataSource()

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select IsNull(MAX(DocEntry),0) +1 From [@Z_ODLK]")
            If Not oRecordSet.EoF Then
                oApplication.Utilities.setEditText(oForm, "13", oRecordSet.Fields.Item(0).Value.ToString())
                oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                oApplication.Utilities.setEditText(oForm, "10", "t")
                oApplication.SBO_Application.SendKeys("{TAB}")
            End If

            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oMatrix.AutoResizeColumns()
            MatrixId = "3"
            'oForm.ActiveItem = "6"
            oMatrix.ClearRowData(oMatrix.RowCount)
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
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_DLK1")
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
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_DLK1")
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
                    oDBDataSourceLines = aForm.DataSources.DBDataSources.Item("@Z_DLK1")
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
                            oDBDataSourceLines = aForm.DataSources.DBDataSources.Item("@Z_DLK1")
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
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_ODLK")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_DLK1")

            If Me.MatrixId = "3" Then
                oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_DLK1")
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
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_ODLK")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_DLK1")

            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oApplication.Utilities.getEditTextvalue(aForm, "6") = "" Then
                oApplication.Utilities.Message("Enter Code...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "7") = "" Then
                oApplication.Utilities.Message("Enter Name...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Dim blnRowExists As Boolean = False
            For index As Integer = 0 To oDBDataSourceLines.Size - 1
                If oDBDataSourceLines.GetValue("U_ItemCode", index) <> "" Then
                    blnRowExists = True
                End If
            Next
            If Not blnRowExists Then
                oApplication.Utilities.Message("Add Dislike Item in Matrix to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            'For index As Integer = 1 To oMatrix.VisualRowCount
            '    Dim strCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
            '    Dim strCode2 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", index)
            '    If strCode2 <> "" And strCode = "" Then
            '        oApplication.Utilities.Message("Item Code missing...in Item : " + strCode2.ToString(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        Return False
            '    End If
            '    For intRow As Integer = 1 To oMatrix.VisualRowCount
            '        If index <> intRow Then
            '            Dim strCode1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
            '            If strCode = strCode1 Then
            '                oApplication.Utilities.Message("Item Code Already Exist...in Item : " + strCode1.ToString(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '                Return False
            '            End If
            '        End If
            '    Next
            'Next

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select 1 As 'Return',DocEntry From [@Z_ODLK]"
            strQuery += " Where "
            strQuery += " U_Code = '" + oDBDataSource.GetValue("U_Code", 0).Trim() + "' And DocEntry <> '" + oDBDataSource.GetValue("DocEntry", 0).ToString() + "'"
            oRecordSet.DoQuery(strQuery)

            If Not oRecordSet.EoF Then
                oApplication.Utilities.Message("Dislike Code Already Exist...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            aForm.Freeze(False)
            Return False
        End Try
    End Function
#End Region

#Region "Disable Controls"

    Private Sub EnableControls(ByVal oForm As SAPbouiCOM.Form, ByVal blnEnable As Boolean)
        Try
            oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("6").Enabled = blnEnable
            oForm.Items.Item("7").Enabled = True
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

#End Region

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
            oCon.Alias = "InvntItem"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_2")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "InvntItem"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
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

            oForm.Items.Item("17").Width = oForm.Width - 30
            oForm.Items.Item("17").Height = oForm.Items.Item("3").Height + 10

            oForm.Freeze(False)
        Catch ex As Exception
            'oApplication.Log.Trace_DIET_AddOn_Error(ex)
            'oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub



#End Region

End Class
