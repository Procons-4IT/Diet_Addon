Imports SAPbobsCOM

Public Class clsMenuDefinition
    Inherits clsBase

    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private InvForConsumedItems, count As Integer
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines1 As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines2 As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines3 As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines4 As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines5 As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines6 As SAPbouiCOM.DBDataSource
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
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_OMED) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_OMED, frm_Z_OMED)
            oForm = oApplication.SBO_Application.Forms.Item(strUID)
            oForm.Freeze(True)
            initialize(oForm)
            addChooseFromListConditions(oForm)
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            'oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            oForm.Items.Item("6").Enabled = False
            oForm.Items.Item("7").Enabled = True
            oForm.Items.Item("7_").Enabled = True
            oForm.Items.Item("_21").Enabled = True
            oForm.DataSources.UserDataSources.Add("PicSource", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 1000)
            oForm.Items.Item("17").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
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
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_OMED) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_OMED, frm_Z_OMED)
            oForm = oApplication.SBO_Application.Forms.Item(strUID)
            oForm.Freeze(True)
            initialize(oForm)
            oForm.DataBrowser.BrowseBy = "6"
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            oForm.Freeze(False)
            'oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            oForm.Items.Item("16").Specific.value = strDocEntry
            oForm.DataSources.UserDataSources.Add("PicSource", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 1000)
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
            If pVal.FormTypeEx = frm_Z_OMED Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And _
                                    (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        If validation(oForm) Then
                                            If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        Else
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                ElseIf pVal.ItemUID = "18" Or pVal.ItemUID = "_18" Then
                                    Dim oOption As SAPbouiCOM.OptionBtn
                                    oOption = oForm.Items.Item(pVal.ItemUID).Specific
                                    oForm.Freeze(True)
                                    If oOption.Selected Then
                                        If pVal.ItemUID = "_18" Then
                                            oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            oForm.Items.Item("20").Visible = False
                                            oForm.Items.Item("21").Visible = False
                                            oForm.Items.Item("22").Visible = False
                                            oForm.Items.Item("23").Visible = False

                                            oForm.Items.Item("_20").Visible = True
                                            oForm.Items.Item("_21").Visible = True
                                            oDBDataSource.SetValue("U_FromDate", 0, "")
                                            oDBDataSource.SetValue("U_ToDate", 0, "")
                                        Else
                                            oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            oForm.Items.Item("_20").Visible = False
                                            oForm.Items.Item("_21").Visible = False

                                            oForm.Items.Item("20").Visible = True
                                            oForm.Items.Item("21").Visible = True
                                            oForm.Items.Item("22").Visible = True
                                            oForm.Items.Item("23").Visible = True
                                            oDBDataSource.SetValue("U_MenuDate", 0, "")
                                        End If
                                    End If
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "47" Or pVal.ItemUID = "47_" Then
                                    Dim oOption As SAPbouiCOM.OptionBtn
                                    oOption = oForm.Items.Item(pVal.ItemUID).Specific
                                    oForm.Freeze(True)
                                    If oOption.Selected Then
                                        If pVal.ItemUID = "47" Then
                                            oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            oForm.Items.Item("4_").Visible = False
                                            oForm.Items.Item("5_").Visible = False
                                            oForm.Items.Item("6_").Visible = False
                                            oForm.Items.Item("7_").Visible = False


                                            oForm.Items.Item("4").Visible = True
                                            oForm.Items.Item("5").Visible = True
                                            oForm.Items.Item("6").Visible = True
                                            oForm.Items.Item("7").Visible = True
                                            oForm.Items.Item("7").Enabled = True

                                            oDBDataSource.SetValue("U_GrpCode", 0, "")
                                            oDBDataSource.SetValue("U_GrpName", 0, "")

                                            'Folders
                                            oForm.Items.Item("17").Enabled = True
                                            oForm.Items.Item("24").Enabled = True
                                            oForm.Items.Item("25").Enabled = True
                                            oForm.Items.Item("26").Enabled = True
                                            oForm.Items.Item("27").Enabled = True
                                            oForm.Items.Item("28").Enabled = True
                                            oForm.Items.Item("17").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            oForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        Else
                                            oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                                            oForm.Items.Item("4_").Visible = True
                                            oForm.Items.Item("5_").Visible = True
                                            oForm.Items.Item("6_").Visible = True
                                            oForm.Items.Item("7_").Visible = True
                                            oForm.Items.Item("7_").Enabled = True

                                            oForm.Items.Item("4").Visible = False
                                            oForm.Items.Item("5").Visible = False
                                            oForm.Items.Item("6").Visible = False
                                            oForm.Items.Item("7").Visible = False

                                            oDBDataSource.SetValue("U_PrgCode", 0, "")
                                            oDBDataSource.SetValue("U_PrgName", 0, "")

                                            'Folders
                                            oForm.Items.Item("17").Enabled = True
                                            oForm.Items.Item("24").Enabled = True
                                            oForm.Items.Item("25").Enabled = True
                                            oForm.Items.Item("26").Enabled = True
                                            oForm.Items.Item("27").Enabled = True
                                            oForm.Items.Item("28").Enabled = True
                                            oForm.Items.Item("17").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            oForm.Items.Item("7_").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        End If
                                    End If
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "35" Then 'Browse
                                    oApplication.Utilities.OpenFileDialogBox(oForm, "36", "37")
                                    If copyFile(oForm) Then
                                        loadImage(oForm)
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                alldataSource(oForm)
                                If (pVal.ItemUID = "3" Or pVal.ItemUID = "29" Or pVal.ItemUID = "30" _
                                                 Or pVal.ItemUID = "31" Or pVal.ItemUID = "32" Or pVal.ItemUID = "33") Then
                                    intSelectedMatrixrow = pVal.Row
                                    If (oDBDataSource.GetValue("U_PrgCode", 0).ToString() = "") Then
                                        BubbleEvent = False
                                        oApplication.SBO_Application.SetStatusBarMessage("Enter Program Code to Proceed...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    End If
                                ElseIf (pVal.ItemUID = "17" Or pVal.ItemUID = "28" Or pVal.ItemUID = "24" _
                                         Or pVal.ItemUID = "25" Or pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "38") Then
                                    oForm.Freeze(True)
                                    If oForm.Items.Item(pVal.ItemUID).Enabled Then
                                        changePane(oForm, pVal.ItemUID)
                                        Dim strItem As String = getMatrixItem(oForm)
                                        If strItem <> "" Then
                                            oMatrix = oForm.Items.Item(strItem).Specific
                                            If oMatrix.RowCount > 0 Then
                                                oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            End If
                                        End If
                                    End If
                                    oForm.Freeze(False)
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
                                                oForm.DataSources.UserDataSources.Item("PicSource").ValueEx = ""
                                            End If
                                        End If
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                alldataSource(oForm)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If pVal.ItemUID = "6" Then
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If IsNothing(oDataTable) Then
                                                Exit Sub
                                            End If
                                            If intAddRows > 0 Then
                                                oDBDataSource.SetValue("U_PrgCode", 0, oDataTable.GetValue("ItemCode", 0))
                                                oDBDataSource.SetValue("U_PrgName", 0, oDataTable.GetValue("ItemName", 0))
                                            End If
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf pVal.ItemUID = "7" Then
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If IsNothing(oDataTable) Then
                                                Exit Sub
                                            End If
                                            If intAddRows > 0 Then
                                                oDBDataSource.SetValue("U_PrgName", 0, oDataTable.GetValue("ItemName", 0))
                                                oDBDataSource.SetValue("U_PrgCode", 0, oDataTable.GetValue("ItemCode", 0))
                                                oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                'oApplication.SBO_Application.SendKeys("{TAB}")
                                                enableSession(oForm) 'Need to Enable Code merge in ItemGroup & ItemMaster
                                            End If
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf pVal.ItemUID = "7_" Then
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If IsNothing(oDataTable) Then
                                                Exit Sub
                                            End If
                                            If intAddRows > 0 Then
                                                oDBDataSource.SetValue("U_GrpName", 0, oDataTable.GetValue("ItmsGrpNam", 0))
                                                oDBDataSource.SetValue("U_GrpCode", 0, oDataTable.GetValue("ItmsGrpCod", 0))
                                                'oApplication.SBO_Application.SendKeys("{TAB}")
                                                enableSession(oForm) 'Need to Enable Code merge in ItemGroup & ItemMaster
                                            End If
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf ((pVal.ItemUID = "3" Or pVal.ItemUID = "29" Or pVal.ItemUID = "30" _
                                                 Or pVal.ItemUID = "31" Or pVal.ItemUID = "32" Or pVal.ItemUID = "33") _
                                                  And (pVal.ColUID = "V_0" Or pVal.ColUID = "V_1")) Then
                                            oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                            oMatrix.FlushToDataSource()
                                            oMatrix.LoadFromDataSource()
                                            If IsNothing(oDataTable) Then
                                                Exit Sub
                                            End If
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 1 Then
                                                intAddRows -= 1
                                                oMatrix.AddRow(intAddRows, pVal.Row - 1)
                                            End If
                                            oMatrix.FlushToDataSource()
                                            Select Case pVal.ItemUID
                                                Case "3"
                                                    For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                        oDBDataSourceLines1.SetValue("LineId", pVal.Row + index - 1, (pVal.Row + index).ToString())
                                                        oDBDataSourceLines1.SetValue("U_ItemCode", pVal.Row + index - 1, oDataTable.GetValue("ItemCode", index))
                                                        oDBDataSourceLines1.SetValue("U_ItemName", pVal.Row + index - 1, oDataTable.GetValue("ItemName", index))
                                                    Next
                                                Case "29"
                                                    For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                        oDBDataSourceLines2.SetValue("LineId", pVal.Row + index - 1, (pVal.Row + index).ToString())
                                                        oDBDataSourceLines2.SetValue("U_ItemCode", pVal.Row + index - 1, oDataTable.GetValue("ItemCode", index))
                                                        oDBDataSourceLines2.SetValue("U_ItemName", pVal.Row + index - 1, oDataTable.GetValue("ItemName", index))
                                                    Next
                                                Case "30"
                                                    For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                        oDBDataSourceLines3.SetValue("LineId", pVal.Row + index - 1, (pVal.Row + index).ToString())
                                                        oDBDataSourceLines3.SetValue("U_ItemCode", pVal.Row + index - 1, oDataTable.GetValue("ItemCode", index))
                                                        oDBDataSourceLines3.SetValue("U_ItemName", pVal.Row + index - 1, oDataTable.GetValue("ItemName", index))
                                                    Next
                                                Case "31"
                                                    For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                        oDBDataSourceLines4.SetValue("LineId", pVal.Row + index - 1, (pVal.Row + index).ToString())
                                                        oDBDataSourceLines4.SetValue("U_ItemCode", pVal.Row + index - 1, oDataTable.GetValue("ItemCode", index))
                                                        oDBDataSourceLines4.SetValue("U_ItemName", pVal.Row + index - 1, oDataTable.GetValue("ItemName", index))
                                                    Next
                                                Case "32"
                                                    For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                        oDBDataSourceLines5.SetValue("LineId", pVal.Row + index - 1, (pVal.Row + index).ToString())
                                                        oDBDataSourceLines5.SetValue("U_ItemCode", pVal.Row + index - 1, oDataTable.GetValue("ItemCode", index))
                                                        oDBDataSourceLines5.SetValue("U_ItemName", pVal.Row + index - 1, oDataTable.GetValue("ItemName", index))
                                                    Next
                                                Case "33"
                                                    For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                        oDBDataSourceLines6.SetValue("LineId", pVal.Row + index - 1, (pVal.Row + index).ToString())
                                                        oDBDataSourceLines6.SetValue("U_ItemCode", pVal.Row + index - 1, oDataTable.GetValue("ItemCode", index))
                                                        oDBDataSourceLines6.SetValue("U_ItemName", pVal.Row + index - 1, oDataTable.GetValue("ItemName", index))
                                                    Next
                                            End Select
                                            oMatrix.LoadFromDataSource()
                                            oMatrix.FlushToDataSource()
                                            oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            Dim strValues As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", oMatrix.RowCount)
                                            If strValues <> "" Then
                                                oMatrix.AddRow(1, oMatrix.RowCount)
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_-1", oMatrix.RowCount, oMatrix.RowCount.ToString)
                                                oMatrix.ClearRowData(oMatrix.RowCount)
                                                oMatrix.FlushToDataSource()
                                            Else
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_-1", oMatrix.RowCount, oMatrix.RowCount.ToString)
                                                oMatrix.ClearRowData(oMatrix.RowCount)
                                                oMatrix.FlushToDataSource()
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
                                    If ((pVal.ItemUID = "3" Or pVal.ItemUID = "29" Or pVal.ItemUID = "30" _
                                                 Or pVal.ItemUID = "31" Or pVal.ItemUID = "32" Or pVal.ItemUID = "33") _
                                                And (pVal.ColUID = "V_0") And pVal.Row > 0) Then

                                        alldataSource(oForm)
                                        oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                        oMatrix.FlushToDataSource()
                                        oMatrix.LoadFromDataSource()
                                        oMatrix.FlushToDataSource()
                                        Select Case pVal.ItemUID
                                            Case "3"
                                                Dim strItemCode As String = oDBDataSourceLines1.GetValue("U_ItemCode", pVal.Row - 1)
                                                Dim strItemName As String = oDBDataSourceLines1.GetValue("U_ItemName", pVal.Row - 1)
                                                If strItemCode.Trim().Length = 0 And strItemName.Trim().Length > 0 Then
                                                    oDBDataSourceLines1.SetValue("U_ItemName", pVal.Row - 1, "")
                                                    oMatrix.LoadFromDataSource()
                                                    oMatrix.FlushToDataSource()
                                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                End If
                                            Case "29"
                                                Dim strItemCode As String = oDBDataSourceLines2.GetValue("U_ItemCode", pVal.Row - 1)
                                                Dim strItemName As String = oDBDataSourceLines2.GetValue("U_ItemName", pVal.Row - 1)
                                                If strItemCode.Trim().Length = 0 And strItemName.Trim().Length > 0 Then
                                                    oDBDataSourceLines2.SetValue("U_ItemName", pVal.Row - 1, "")
                                                    oMatrix.LoadFromDataSource()
                                                    oMatrix.FlushToDataSource()
                                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                End If
                                            Case "30"
                                                Dim strItemCode As String = oDBDataSourceLines3.GetValue("U_ItemCode", pVal.Row - 1)
                                                Dim strItemName As String = oDBDataSourceLines3.GetValue("U_ItemName", pVal.Row - 1)
                                                If strItemCode.Trim().Length = 0 And strItemName.Trim().Length > 0 Then
                                                    oDBDataSourceLines3.SetValue("U_ItemName", pVal.Row - 1, "")
                                                    oMatrix.LoadFromDataSource()
                                                    oMatrix.FlushToDataSource()
                                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                End If
                                            Case "31"
                                                Dim strItemCode As String = oDBDataSourceLines4.GetValue("U_ItemCode", pVal.Row - 1)
                                                Dim strItemName As String = oDBDataSourceLines4.GetValue("U_ItemName", pVal.Row - 1)
                                                If strItemCode.Trim().Length = 0 And strItemName.Trim().Length > 0 Then
                                                    oDBDataSourceLines4.SetValue("U_ItemName", pVal.Row - 1, "")
                                                    oMatrix.LoadFromDataSource()
                                                    oMatrix.FlushToDataSource()
                                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                End If
                                            Case "32"
                                                Dim strItemCode As String = oDBDataSourceLines5.GetValue("U_ItemCode", pVal.Row - 1)
                                                Dim strItemName As String = oDBDataSourceLines5.GetValue("U_ItemName", pVal.Row - 1)
                                                If strItemCode.Trim().Length = 0 And strItemName.Trim().Length > 0 Then
                                                    oDBDataSourceLines5.SetValue("U_ItemName", pVal.Row - 1, "")
                                                    oMatrix.LoadFromDataSource()
                                                    oMatrix.FlushToDataSource()
                                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                End If
                                            Case "33"
                                                Dim strItemCode As String = oDBDataSourceLines6.GetValue("U_ItemCode", pVal.Row - 1)
                                                Dim strItemName As String = oDBDataSourceLines6.GetValue("U_ItemName", pVal.Row - 1)
                                                If strItemCode.Trim().Length = 0 And strItemName.Trim().Length > 0 Then
                                                    oDBDataSourceLines6.SetValue("U_ItemName", pVal.Row - 1, "")
                                                    oMatrix.LoadFromDataSource()
                                                    oMatrix.FlushToDataSource()
                                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                End If
                                        End Select
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                If pVal.CharPressed = 9 Then
                                    If pVal.ItemUID = "7" Or pVal.ItemUID = "7_" Then
                                        'enableSession(oForm)
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                If oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Or oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    Try
                                        reDrawForm(oForm)
                                    Catch ex As Exception
                                        ' oApplication.Log.Trace_DIET_AddOn_Error(ex)
                                    End Try
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED
                                If pVal.ItemUID = "21" Then
                                    oForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
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
                Case mnu_Z_OMED
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oForm.Items.Item("17").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        Dim strItem As String = getMatrixItem(oForm)
                        AddRow(oForm, strItem)
                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        Dim strItem As String = getMatrixItem(oForm)
                        RefereshDeleteRow(oForm, strItem)
                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                Case mnu_ADD
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        initialize(oForm)
                        EnableControls(oForm, True)
                        oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oForm.Items.Item("6").Enabled = False
                        oForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oForm.Items.Item("20").Visible = False
                        oForm.Items.Item("21").Visible = False
                        oForm.Items.Item("22").Visible = False
                        oForm.Items.Item("23").Visible = False

                        oForm.Items.Item("_20").Visible = True
                        oForm.Items.Item("_21").Visible = True

                        'Folders
                        oForm.Items.Item("18").Enabled = True
                        oForm.Items.Item("24").Enabled = True
                        oForm.Items.Item("25").Enabled = True
                        oForm.Items.Item("26").Enabled = True
                        oForm.Items.Item("27").Enabled = True
                        oForm.Items.Item("28").Enabled = True


                        oDBDataSource.SetValue("U_FromDate", 0, "")
                        oDBDataSource.SetValue("U_ToDate", 0, "")
                        oForm.DataSources.UserDataSources.Item("PicSource").ValueEx = ""
                        oForm.Items.Item("17").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oForm.Items.Item("7_").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oForm.Update()
                        oForm.Refresh()
                    End If
                Case mnu_FIND
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        EnableControls(oForm, True)
                        oForm.Items.Item("6").Enabled = True
                        oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oForm.Items.Item("17").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oForm.Update()
                        oForm.Refresh()
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
            If oForm.TypeEx = frm_Z_OMED Then
                Select Case BusinessObjectInfo.BeforeAction
                    Case True

                    Case False
                        Select Case BusinessObjectInfo.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                EnableControls(oForm, False)
                                alldataSource(oForm)
                                Dim strType As String = oDBDataSource.GetValue("U_MenuType", 0).Trim()
                                oForm.Freeze(True)
                                If strType = "R" Then
                                    oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oForm.Items.Item("20").Visible = False
                                    oForm.Items.Item("21").Visible = False
                                    oForm.Items.Item("22").Visible = False
                                    oForm.Items.Item("23").Visible = False

                                    oForm.Items.Item("_20").Visible = True
                                    oForm.Items.Item("_21").Visible = True
                                Else
                                    oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oForm.Items.Item("_20").Visible = False
                                    oForm.Items.Item("_21").Visible = False

                                    oForm.Items.Item("20").Visible = True
                                    oForm.Items.Item("21").Visible = True
                                    oForm.Items.Item("22").Visible = True
                                    oForm.Items.Item("23").Visible = True
                                    oDBDataSource.SetValue("U_MenuDate", 0, "")
                                End If
                                Dim strCatType As String = oDBDataSource.GetValue("U_CatType", 0).Trim()
                                Dim strProgram As String = String.Empty
                                If strCatType = "I" Then

                                    strProgram = oDBDataSource.GetValue("U_PrgCode", 0).Trim()
                                    oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oForm.Items.Item("4_").Visible = False
                                    oForm.Items.Item("5_").Visible = False
                                    oForm.Items.Item("6_").Visible = False
                                    oForm.Items.Item("7_").Visible = False

                                    oForm.Items.Item("4").Visible = True
                                    oForm.Items.Item("5").Visible = True
                                    oForm.Items.Item("6").Visible = True
                                    oForm.Items.Item("7").Visible = True

                                Else

                                    strProgram = oDBDataSource.GetValue("U_GrpCode", 0).Trim()
                                    oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                                    oForm.Items.Item("4").Visible = False
                                    oForm.Items.Item("5").Visible = False
                                    oForm.Items.Item("6").Visible = False
                                    oForm.Items.Item("7").Visible = False

                                    oForm.Items.Item("4_").Visible = True
                                    oForm.Items.Item("5_").Visible = True
                                    oForm.Items.Item("6_").Visible = True
                                    oForm.Items.Item("7_").Visible = True
                                End If
                                enableSession(oForm) 'Need to Enable Code merge in ItemGroup & ItemMaster
                                loadImage(oForm)
                                oForm.Freeze(False)
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
            oForm.PaneLevel = 1
            alldataSource(oForm)

            addblankRow(oForm, "3")
            addblankRow(oForm, "29")
            addblankRow(oForm, "30")
            addblankRow(oForm, "31")
            addblankRow(oForm, "32")
            addblankRow(oForm, "33")

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select IsNull(MAX(DocEntry),0) +1 From [@Z_OMED]")
            If Not oRecordSet.EoF Then
                oApplication.Utilities.setEditText(oForm, "13", oRecordSet.Fields.Item(0).Value.ToString())
                'oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                'oApplication.Utilities.setEditText(oForm, "10", "t")
                'oApplication.SBO_Application.SendKeys("{TAB}")
            End If

            Dim oOption As SAPbouiCOM.OptionBtn
            oOption = oForm.Items.Item("_18").Specific
            oOption.Selected = True

            oOption = oForm.Items.Item("47_").Specific
            oOption.Selected = True

            oForm.Items.Item("20").Visible = False
            oForm.Items.Item("21").Visible = False
            oForm.Items.Item("22").Visible = False
            oForm.Items.Item("23").Visible = False

            oForm.Items.Item("4_").Visible = False
            oForm.Items.Item("5_").Visible = False
            oForm.Items.Item("6_").Visible = False
            Try
                oForm.Items.Item("7_").Visible = False
            Catch ex As Exception
                'oApplication.Log.Trace_DIET_AddOn_Error(ex)
            End Try
            oForm.Items.Item("_20").Visible = True
            oForm.Items.Item("_21").Visible = True

            oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("47_").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            MatrixId = "3"
            oForm.Update()
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

#Region "Methods"

    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form, ByVal strItem As String)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item(strItem).Specific
            alldataSource(oForm)
            oMatrix.FlushToDataSource()

            Select Case strItem
                Case "3"
                    For count = 1 To oDBDataSourceLines1.Size
                        oDBDataSourceLines1.SetValue("LineId", count - 1, count)
                    Next
                Case "29"
                    For count = 1 To oDBDataSourceLines2.Size
                        oDBDataSourceLines2.SetValue("LineId", count - 1, count)
                    Next
                Case "30"
                    For count = 1 To oDBDataSourceLines3.Size
                        oDBDataSourceLines3.SetValue("LineId", count - 1, count)
                    Next
                Case "31"
                    For count = 1 To oDBDataSourceLines4.Size
                        oDBDataSourceLines4.SetValue("LineId", count - 1, count)
                    Next
                Case "32"
                    For count = 1 To oDBDataSourceLines5.Size
                        oDBDataSourceLines5.SetValue("LineId", count - 1, count)
                    Next
                Case "33"
                    For count = 1 To oDBDataSourceLines6.Size
                        oDBDataSourceLines6.SetValue("LineId", count - 1, count)
                    Next
            End Select

            oMatrix.LoadFromDataSource()
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            aForm.Freeze(False)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub alldataSource(ByVal aForm As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OMED")
            oDBDataSourceLines1 = oForm.DataSources.DBDataSources.Item("@Z_MED1")
            oDBDataSourceLines2 = oForm.DataSources.DBDataSources.Item("@Z_MED2")
            oDBDataSourceLines3 = oForm.DataSources.DBDataSources.Item("@Z_MED3")
            oDBDataSourceLines4 = oForm.DataSources.DBDataSources.Item("@Z_MED4")
            oDBDataSourceLines5 = oForm.DataSources.DBDataSources.Item("@Z_MED5")
            oDBDataSourceLines6 = oForm.DataSources.DBDataSources.Item("@Z_MED6")
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

#End Region

    Private Sub addblankRow(ByVal oForm As SAPbouiCOM.Form, ByVal strItem As String)
        Try
            oMatrix = oForm.Items.Item(strItem).Specific
            oMatrix.LoadFromDataSource()
            oMatrix.AddRow(1, -1)
            oMatrix.FlushToDataSource()
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oMatrix.ClearRowData(oMatrix.RowCount)
            AssignLineNo(oForm, strItem)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form, ByVal strItem As String)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item(strItem).Specific

            Select Case aForm.PaneLevel
                Case "0", "1"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED1")
                Case "2"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED2")
                Case "3"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED3")
                Case "4"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED4")
                Case "5"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED5")
                Case "6"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED6")
            End Select

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
            oMatrix.FlushToDataSource()
            AssignLineNo(aForm, strItem)
            oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            aForm.Freeze(False)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub deleterow(ByVal aForm As SAPbouiCOM.Form, ByVal strItem As String)
        Try
            oMatrix = aForm.Items.Item(strItem).Specific
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
                        Case "0", "1"
                            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED1")
                        Case "2"
                            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED2")
                        Case "3"
                            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED3")
                        Case "4"
                            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED4")
                        Case "5"
                            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED5")
                        Case "6"
                            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED6")
                    End Select
                    AssignLineNo(aForm, strItem)
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

    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form, ByVal strItem As String)
        Try
            oMatrix = aForm.Items.Item(strItem).Specific
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OMED")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED1")

            Select Case aForm.PaneLevel
                Case "0", "1"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED1")
                Case "2"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED2")
                Case "3"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED3")
                Case "4"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED4")
                Case "5"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED5")
                Case "6"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_MED6")
            End Select

            Me.RowtoDelete = intSelectedMatrixrow
            oDBDataSourceLines.RemoveRecord(Me.RowtoDelete - 1)
            oMatrix.LoadFromDataSource()
            oMatrix.FlushToDataSource()
            For count = 0 To oDBDataSourceLines.Size - 1
                oDBDataSourceLines.SetValue("LineId", count, count + 1)
            Next
            oMatrix.LoadFromDataSource()
            oMatrix.FlushToDataSource()
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
            alldataSource(oForm)

            Dim strProgram As String = oDBDataSource.GetValue("U_PrgCode", 0).Trim()
            Dim strProgramN As String = oDBDataSource.GetValue("U_PrgName", 0).Trim()
            Dim strGroup As String = oDBDataSource.GetValue("U_GrpCode", 0).Trim()
            Dim strGroupN As String = oDBDataSource.GetValue("U_GrpName", 0).Trim()
            Dim strType As String = oDBDataSource.GetValue("U_MenuType", 0).Trim()
            Dim strMeDt As String = oDBDataSource.GetValue("U_MenuDate", 0).Trim()
            Dim strFrDt As String = oDBDataSource.GetValue("U_FromDate", 0).Trim()
            Dim strToDt As String = oDBDataSource.GetValue("U_ToDate", 0).Trim()
            Dim strCatType As String = oDBDataSource.GetValue("U_CatType", 0).Trim()

            If strProgram.Trim().Length = 0 And strCatType = "I" Then
                oApplication.Utilities.Message("Enter Program Code...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strProgramN.Trim().Length = 0 And strCatType = "I" Then
                oApplication.Utilities.Message("Enter Program Name...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strGroup.Trim().Length = 0 And strCatType = "G" Then
                oApplication.Utilities.Message("Enter Group Code...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strGroupN.Trim().Length = 0 And strCatType = "G" Then
                oApplication.Utilities.Message("Enter Group Name...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strType = "R" Then

                If strMeDt.Trim().Length = 0 Then
                    oApplication.Utilities.Message("Enter Menu Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strQuery = "Select 1 As 'Return',DocEntry From [@Z_OMED]"
                strQuery += " Where "
                If strCatType = "I" Then
                    strQuery += " U_PrgCode = '" + oDBDataSource.GetValue("U_PrgCode", 0).Trim() + "' "
                Else
                    strQuery += " U_GrpCode = '" + oDBDataSource.GetValue("U_GrpCode", 0).Trim() + "' "
                End If
                strQuery += " And DocEntry <> '" + oDBDataSource.GetValue("DocEntry", 0).ToString() + "'"
                strQuery += " And U_MenuType = '" + strType + "'"
                strQuery += " And U_MenuDate = '" + oDBDataSource.GetValue("U_MenuDate", 0).Trim() + "'"
                oRecordSet.DoQuery(strQuery)
                If Not oRecordSet.EoF Then
                    oApplication.Utilities.Message("Menu Already defined for Regular Menu Type...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

            ElseIf strType = "A" Then

                If strFrDt.Trim().Length = 0 Then
                    oApplication.Utilities.Message("Enter Menu From Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                ElseIf strToDt.Trim().Length = 0 Then
                    oApplication.Utilities.Message("Enter Menu To Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                ElseIf CInt(strFrDt) > CInt(strToDt) Then
                    oApplication.Utilities.Message("Menu From Date Should be Lesser than or Equal Menu To Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                strQuery = "Select 1 As 'Return',DocEntry From [@Z_OMED]"
                strQuery += " Where "
                strQuery += " (('" + oDBDataSource.GetValue("U_FromDate", 0).ToString() + "' Between Convert(VarChar(12),U_FromDate,112) And Convert(VarChar(12),U_ToDate,112)) "
                strQuery += " OR "
                strQuery += " ('" + oDBDataSource.GetValue("U_ToDate", 0).ToString() + "' Between Convert(VarChar(12),U_FromDate,112) And Convert(VarChar(12),U_ToDate,112))) "
                strQuery += " And DocEntry <> '" + oDBDataSource.GetValue("DocEntry", 0).ToString() + "'"
                oRecordSet.DoQuery(strQuery)
                If Not oRecordSet.EoF Then
                    oApplication.Utilities.Message("Menu Already defined for Alternative Menu Type...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

            End If

            Dim blnBreak As Boolean = False
            For index As Integer = 0 To oDBDataSourceLines1.Size - 1
                If oDBDataSourceLines1.GetValue("U_ItemCode", index) <> "" Then
                    blnBreak = True
                    Exit For
                End If
            Next

            Dim blnLunch As Boolean = False
            For index As Integer = 0 To oDBDataSourceLines2.Size - 1
                If oDBDataSourceLines2.GetValue("U_ItemCode", index) <> "" Then
                    blnLunch = True
                    Exit For
                End If
            Next

            Dim blnLunchSide As Boolean = False
            For index As Integer = 0 To oDBDataSourceLines3.Size - 1
                If oDBDataSourceLines3.GetValue("U_ItemCode", index) <> "" Then
                    blnLunchSide = True
                    Exit For
                End If
            Next

            Dim blnSnack As Boolean = False
            For index As Integer = 0 To oDBDataSourceLines4.Size - 1
                If oDBDataSourceLines4.GetValue("U_ItemCode", index) <> "" Then
                    blnSnack = True
                    Exit For
                End If
            Next

            Dim blnDinner As Boolean = False
            For index As Integer = 0 To oDBDataSourceLines5.Size - 1
                If oDBDataSourceLines5.GetValue("U_ItemCode", index) <> "" Then
                    blnDinner = True
                    Exit For
                End If
            Next

            Dim blnDinnerSide As Boolean = False
            For index As Integer = 0 To oDBDataSourceLines6.Size - 1
                If oDBDataSourceLines6.GetValue("U_ItemCode", index) <> "" Then
                    blnDinnerSide = True
                    Exit For
                End If
            Next

            If Not blnBreak And Not blnLunch And Not blnLunchSide And Not blnSnack And Not blnDinner And Not blnDinnerSide Then
                oApplication.Utilities.Message("At least one tab should be filled", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            'If Not blnBreak Then
            '    oApplication.Utilities.Message("Add Item in Break to Proceed", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If


            'If Not blnLunch Then
            '    oApplication.Utilities.Message("Add Item in Lunch to Proceed", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If


            'If Not blnLunchSide Then
            '    oApplication.Utilities.Message("Add Item in Lunch(Side) to Proceed", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

            'If Not blnSnack Then
            '    oApplication.Utilities.Message("Add Item in Snack to Proceed", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

            'If Not blnDinner Then
            '    oApplication.Utilities.Message("Add Item in Dinner to Proceed", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

            'If Not blnDinnerSide Then
            '    oApplication.Utilities.Message("Add Item in Dinner(Side) to Proceed", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

            oMatrix = oForm.Items.Item("3").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim strCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    If index <> intRow Then
                        Dim strCode1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
                        If strCode = strCode1 Then
                            oApplication.Utilities.Message("Item Code Already Exist...For Breakfast - Item : " + strCode1.ToString(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next
            Next

            oMatrix = oForm.Items.Item("29").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim strCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    If index <> intRow Then
                        Dim strCode1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
                        If strCode = strCode1 Then
                            oApplication.Utilities.Message("Item Code Already Exist...For Lunch - Item : " + strCode1.ToString(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next
            Next

            oMatrix = oForm.Items.Item("30").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim strCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    If index <> intRow Then
                        Dim strCode1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
                        If strCode = strCode1 Then
                            oApplication.Utilities.Message("Item Code Already Exist...For Lunch(Side) - Item : " + strCode1.ToString(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next
            Next

            oMatrix = oForm.Items.Item("31").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim strCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    If index <> intRow Then
                        Dim strCode1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
                        If strCode = strCode1 Then
                            oApplication.Utilities.Message("Item Code Already Exist...For Snack - Item : " + strCode1.ToString(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next
            Next

            oMatrix = oForm.Items.Item("32").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim strCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    If index <> intRow Then
                        Dim strCode1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
                        If strCode = strCode1 Then
                            oApplication.Utilities.Message("Item Code Already Exist...For Dinner - Item : " + strCode1.ToString(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next
            Next

            oMatrix = oForm.Items.Item("33").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim strCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    If index <> intRow Then
                        Dim strCode1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
                        If strCode = strCode1 Then
                            oApplication.Utilities.Message("Item Code Already Exist...For Dinner(Side) - Item : " + strCode1.ToString(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next
            Next

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
            oForm.Items.Item("7_").Enabled = blnEnable
            oForm.Items.Item("_18").Enabled = blnEnable
            oForm.Items.Item("18").Enabled = blnEnable
            oForm.Items.Item("_21").Enabled = blnEnable
            oForm.Items.Item("21").Enabled = blnEnable
            oForm.Items.Item("22").Enabled = blnEnable
            oForm.Items.Item("23").Enabled = blnEnable
            oForm.Items.Item("47").Enabled = blnEnable
            oForm.Items.Item("47_").Enabled = blnEnable

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

#End Region

    Private Function getMatrixItem(ByVal oForm As SAPbouiCOM.Form) As String
        Try
            Dim _retVal As String = String.Empty
            Select Case oForm.PaneLevel
                Case "0", "1"
                    _retVal = "3"
                Case "2"
                    _retVal = "29"
                Case "3"
                    _retVal = "30"
                Case "4"
                    _retVal = "31"
                Case "5"
                    _retVal = "32"
                Case "6"
                    _retVal = "33"
            End Select
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Private Function changePane(ByVal oForm As SAPbouiCOM.Form, ByVal strItem As String) As String
        Try
            Dim _retVal As String = String.Empty
            Select Case strItem
                Case "17"
                    oForm.PaneLevel = 1
                Case "24"
                    oForm.PaneLevel = 2
                Case "25"
                    oForm.PaneLevel = 3
                Case "26"
                    oForm.PaneLevel = 4
                Case "27"
                    oForm.PaneLevel = 5
                Case "28"
                    oForm.PaneLevel = 6
                Case "38"
                    oForm.PaneLevel = 7
            End Select
            _retVal = oForm.PaneLevel
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Private Sub addChooseFromListConditions(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oCFLs = oForm.ChooseFromLists

            oCFL = oCFLs.Item("CFL_1_1")
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

            'MessageBox.Show(oCFL.GetConditions().GetAsXML())

            oCFL = oCFLs.Item("CFL_2_2")
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

            oCFL = oCFLs.Item("CFL_3_3")
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

            oCFL = oCFLs.Item("CFL_4_4")
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

            oCFL = oCFLs.Item("CFL_5_5")
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

            oCFL = oCFLs.Item("CFL_6_6")
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

            strQuery = "Select ItmsGrpCod From OITB Where U_Program = 'Y' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oCFL = oCFLs.Item("CFL_8")
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

            'Item Group Condition
            oCFL = oCFLs.Item("CFL_9")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.[Alias] = "U_Program"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            'Item Group Condition
            oCFL = oCFLs.Item("CFL_10")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.[Alias] = "U_Program"
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

            oForm.Items.Item("34").Width = oForm.Width - 30
            oForm.Items.Item("34").Height = oForm.Items.Item("3").Height + 10

            oForm.Freeze(False)
        Catch ex As Exception
            'oApplication.Log.Trace_DIET_AddOn_Error(ex)
            'oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Private Function copyFile(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim _retVal As Boolean = False
        Try
            'Logic to Move Files from Source to Destination.
            Dim strSource As String = CType(aForm.Items.Item("36").Specific, SAPbouiCOM.EditText).Value
            Dim strDest As String = oApplication.Utilities.getPicturePath() + CType(aForm.Items.Item("37").Specific, SAPbouiCOM.EditText).Value
            My.Computer.FileSystem.CopyFile(strSource, strDest, True)
            _retVal = True
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
        Return _retVal
    End Function

    Public Sub loadImage(ByVal aForm As SAPbouiCOM.Form)
        Try
            Dim strDest As String = oApplication.Utilities.getPicturePath() + CType(oForm.Items.Item("37").Specific, SAPbouiCOM.EditText).Value
            Dim oPicture As SAPbouiCOM.PictureBox
            oPicture = CType(aForm.Items.Item("39").Specific, SAPbouiCOM.PictureBox)
            oPicture.DataBind.SetBound(True, "", "PicSource")
            aForm.DataSources.UserDataSources.Item("PicSource").ValueEx = strDest
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Public Sub clearImage(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.DataSources.UserDataSources.Item("PicSource").ValueEx = ""
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Public Sub enableSession(ByVal aForm As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OMED")
            oForm.Freeze(True)
            oForm.Items.Item("18").Enabled = True
            oForm.Items.Item("24").Enabled = True
            oForm.Items.Item("25").Enabled = True
            oForm.Items.Item("26").Enabled = True
            oForm.Items.Item("27").Enabled = True
            oForm.Items.Item("28").Enabled = True
            If oDBDataSource.GetValue("U_CatType", 0).ToString() = "I" Then
                Dim strProgram As String = oDBDataSource.GetValue("U_PrgCode", 0).ToString()
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strQuery = "Select "
                strQuery += " ISNULL(U_BF,'N') As 'U_BF',ISNULL(U_LN,'N') As 'U_LN',ISNULL(U_LS,'N') As 'U_LS', "
                strQuery += " ISNULL(U_SK,'N') As 'U_SK',ISNULL(U_DN,'N') As 'U_DN',ISNULL(U_DS,'N') As 'U_DS' "
                strQuery += " From OITM Where ItemCode = '" & strProgram & "'"
                oRecordSet.DoQuery(strQuery)
                If Not oRecordSet.EoF Then

                    oForm.Items.Item("17").Enabled = IIf(oRecordSet.Fields.Item("U_BF").Value = "Y", True, False)
                    oForm.Items.Item("24").Enabled = IIf(oRecordSet.Fields.Item("U_LN").Value = "Y", True, False)
                    oForm.Items.Item("25").Enabled = IIf(oRecordSet.Fields.Item("U_LS").Value = "Y", True, False)
                    oForm.Items.Item("26").Enabled = IIf(oRecordSet.Fields.Item("U_SK").Value = "Y", True, False)
                    oForm.Items.Item("27").Enabled = IIf(oRecordSet.Fields.Item("U_DN").Value = "Y", True, False)
                    oForm.Items.Item("28").Enabled = IIf(oRecordSet.Fields.Item("U_DS").Value = "Y", True, False)

                    Dim strNotSelFol As String = IIf(Not oForm.Items.Item("17").Enabled, _
                                                     IIf(Not oForm.Items.Item("24").Enabled, _
                                                         IIf(Not oForm.Items.Item("25").Enabled, _
                                                             IIf(Not oForm.Items.Item("26").Enabled, _
                                                                 IIf(Not oForm.Items.Item("27").Enabled, _
                                                                     IIf(Not oForm.Items.Item("28").Enabled, _
                                                                         "38",
                                                                         oForm.Items.Item("28").UniqueID),
                                                                     oForm.Items.Item("27").UniqueID),
                                                                 oForm.Items.Item("26").UniqueID),
                                                             oForm.Items.Item("25").UniqueID), _
                                                         oForm.Items.Item("24").UniqueID) _
                                                     , oForm.Items.Item("17").UniqueID)

                    Dim intPane As Integer = changePane(oForm, strNotSelFol)
                    oForm.DataSources.UserDataSources.Item("FolderDS").ValueEx = intPane

                    'oForm.Items.Item(strNotSelFol).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    'changePane(oForm, IIf(strNotSelFol = "", "38", strNotSelFol))
                End If
            ElseIf oDBDataSource.GetValue("U_CatType", 0).ToString() = "G" Then
                Dim strGroup As String = oDBDataSource.GetValue("U_GrpCode", 0).ToString()
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strQuery = "Select "
                strQuery += " ISNULL(U_BF,'N') As 'U_BF',ISNULL(U_LN,'N') As 'U_LN',ISNULL(U_LS,'N') As 'U_LS', "
                strQuery += " ISNULL(U_SK,'N') As 'U_SK',ISNULL(U_DN,'N') As 'U_DN',ISNULL(U_DS,'N') As 'U_DS' "
                strQuery += " From OITB Where ItmsGrpCod = '" & strGroup & "'"
                oRecordSet.DoQuery(strQuery)
                If Not oRecordSet.EoF Then

                    oForm.Items.Item("17").Enabled = IIf(oRecordSet.Fields.Item("U_BF").Value = "Y", True, False)
                    oForm.Items.Item("24").Enabled = IIf(oRecordSet.Fields.Item("U_LN").Value = "Y", True, False)
                    oForm.Items.Item("25").Enabled = IIf(oRecordSet.Fields.Item("U_LS").Value = "Y", True, False)
                    oForm.Items.Item("26").Enabled = IIf(oRecordSet.Fields.Item("U_SK").Value = "Y", True, False)
                    oForm.Items.Item("27").Enabled = IIf(oRecordSet.Fields.Item("U_DN").Value = "Y", True, False)
                    oForm.Items.Item("28").Enabled = IIf(oRecordSet.Fields.Item("U_DS").Value = "Y", True, False)

                    Dim strNotSelFol As String = IIf(Not oForm.Items.Item("17").Enabled, _
                                                     IIf(Not oForm.Items.Item("24").Enabled, _
                                                         IIf(Not oForm.Items.Item("25").Enabled, _
                                                             IIf(Not oForm.Items.Item("26").Enabled, _
                                                                 IIf(Not oForm.Items.Item("27").Enabled, _
                                                                     IIf(Not oForm.Items.Item("28").Enabled, _
                                                                         "38",
                                                                         oForm.Items.Item("28").UniqueID),
                                                                     oForm.Items.Item("27").UniqueID),
                                                                 oForm.Items.Item("26").UniqueID),
                                                             oForm.Items.Item("25").UniqueID), _
                                                         oForm.Items.Item("24").UniqueID) _
                                                     , oForm.Items.Item("17").UniqueID)

                    'oForm.Items.Item(strNotSelFol).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    'changePane(oForm, IIf(strNotSelFol = "", "38", strNotSelFol))

                    Dim intPane As Integer = changePane(oForm, strNotSelFol)
                    oForm.DataSources.UserDataSources.Item("FolderDS").ValueEx = intPane

                End If
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm.Freeze(False)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

#End Region

End Class
