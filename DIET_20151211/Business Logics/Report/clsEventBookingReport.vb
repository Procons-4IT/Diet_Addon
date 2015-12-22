Public Class clsEventBookingReport
    Inherits clsBase

    Private oMatrix As SAPbouiCOM.Matrix
    Dim oStatic As SAPbouiCOM.StaticText
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private oEventGrid As SAPbouiCOM.Grid
    Private ocombo As SAPbouiCOM.ComboBoxColumn
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim blnFormLoaded As Boolean = False
    Dim oDBDataSource As SAPbouiCOM.DBDataSource

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_Z_OEVR, frm_Z_OEVR)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            oForm.PaneLevel = 1
            initialize(oForm)
            oForm.Settings.EnableRowFormat = False
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Z_OEVR Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And (oForm.PaneLevel = 2) Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                blnFormLoaded = True
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And (oForm.PaneLevel = 1) Then
                                    oForm.PaneLevel = oForm.PaneLevel + 1
                                    changeLabel(oForm)
                                ElseIf pVal.ItemUID = "3" And (oForm.PaneLevel = 2) Then
                                    loadEvents(oForm)
                                    oEventGrid = oForm.Items.Item("10").Specific
                                    If oEventGrid.DataTable.Rows.Count >= 1 Then
                                        oForm.PaneLevel = oForm.PaneLevel + 1
                                        changeLabel(oForm)
                                    Else
                                        If oEventGrid.DataTable.Rows.Count = 0 Then
                                            oApplication.Utilities.Message("No Events Found for the Selection...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                    End If
                                ElseIf pVal.ItemUID = "4" Then
                                    oForm.Freeze(True)
                                    If oForm.PaneLevel <> 2 Then
                                        oForm.PaneLevel = 2
                                    Else
                                        oForm.PaneLevel = oForm.PaneLevel - 1
                                        changeLabel(oForm)
                                    End If
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "19" Then
                                    oForm.Freeze(True)
                                    If oForm.PaneLevel <> 3 Then
                                        oForm.PaneLevel = 3
                                    End If
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "21" Then
                                    oForm.Freeze(True)
                                    If oForm.PaneLevel <> 4 Then
                                        oForm.PaneLevel = 4
                                    End If
                                    oForm.Freeze(False)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OEBQ")
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim strCode, strName, strHallSpace, strMaxCap As String
                                Dim sCHFL_ID As String
                                Dim intChoice As Integer
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        intChoice = 0
                                        oForm.Freeze(True)
                                        If pVal.ItemUID = "11" Then
                                            strCode = oDataTable.GetValue("U_Code", 0)
                                            strName = oDataTable.GetValue("U_Name", 0)
                                            oDBDataSource.SetValue("U_EType", oDBDataSource.Offset, strCode)
                                            'oDBDataSource.SetValue("U_ETypeN", oDBDataSource.Offset, strName)
                                        ElseIf pVal.ItemUID = "12" Then
                                            strCode = oDataTable.GetValue("U_Code", 0)
                                            strName = oDataTable.GetValue("U_Name", 0)
                                            oDBDataSource.SetValue("U_ELabel", oDBDataSource.Offset, strCode)
                                            'oDBDataSource.SetValue("U_ELabelN", oDBDataSource.Offset, strName)                                       
                                        ElseIf pVal.ItemUID = "13" Then
                                            strCode = oDataTable.GetValue("U_Code", 0)
                                            strName = oDataTable.GetValue("U_Name", 0)
                                            'strHallSpace = oDataTable.GetValue("U_HallSpc", 0)
                                            'strMaxCap = oDataTable.GetValue("U_MaxCap", 0)
                                            oDBDataSource.SetValue("U_FSpace", oDBDataSource.Offset, strCode)
                                            'oDBDataSource.SetValue("U_FSpaceN", oDBDataSource.Offset, strName)
                                            'oDBDataSource.SetValue("U_HallSpc", oDBDataSource.Offset, strHallSpace)
                                            'oDBDataSource.SetValue("U_MaxCap", oDBDataSource.Offset, strMaxCap)
                                        End If
                                        oForm.Freeze(False)
                                    End If
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

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Z_OEVR
                    LoadForm()
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Data Event"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Validations"

    Private Function Validation(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim strEventType, strEventLabel, strFunSpace As String
            strEventType = CType(oForm.Items.Item("11").Specific, SAPbouiCOM.EditText).Value
            strEventLabel = CType(oForm.Items.Item("12").Specific, SAPbouiCOM.EditText).Value
            strFunSpace = CType(oForm.Items.Item("13").Specific, SAPbouiCOM.EditText).Value

            'If strEventType = "" Then
            '    oApplication.Utilities.Message("Select Event Type ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'ElseIf strEventLabel = "" Then
            '    oApplication.Utilities.Message("Select Event Label ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'ElseIf strFunSpace = "" Then
            '    oApplication.Utilities.Message("Select Function Space ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

#End Region

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Items.Item("1").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
            oForm.Items.Item("_13").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
            oForm.DataSources.DataTables.Add("dtEvents")
            changeLabel(oForm)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub loadEvents(ByVal aform As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            Dim strqry As String

            Dim strEventType, strEventLabel, strFunSpace As String
            strEventType = CType(oForm.Items.Item("11").Specific, SAPbouiCOM.EditText).Value
            strEventLabel = CType(oForm.Items.Item("12").Specific, SAPbouiCOM.EditText).Value
            strFunSpace = CType(oForm.Items.Item("13").Specific, SAPbouiCOM.EditText).Value

            oEventGrid = oForm.Items.Item("10").Specific
            oEventGrid.DataTable = oForm.DataSources.DataTables.Item("dtEvents")

            strqry = " Select T1.U_Name As 'Event Type',T2.U_Name As 'Event Label',U_FSpace,U_StartDt,U_EndDt,U_FunDate From [@Z_OEBQ] T0  "
            strqry += " JOIN [@Z_OEVT] T1 On T0.U_EType = T1.U_Code JOIN [@Z_OEVL] T2 On T0.U_ELabel = T2.U_Code "
            strqry += " JOIN OSLP T3 On T3.SlpCode = T0.U_SalEmp Where 1 = 1 "

            If strEventType.Length > 0 Then
                strqry += " And T0.U_EType = '" + strEventType + "'"
            End If

            If strEventLabel.Length > 0 Then
                strqry += " And T0.U_ELabel = '" + strEventLabel + "'"
            End If

            If strFunSpace.Length > 0 Then
                strqry += " And T0.U_FSpace = '" + strFunSpace + "'"
            End If

            oEventGrid.DataTable.ExecuteQuery(strqry)

            oEventGrid.Columns.Item("Event Type").TitleObject.Caption = "Event Type"
            oEventGrid.Columns.Item("Event Label").TitleObject.Caption = "Event Label"
            oEventGrid.Columns.Item("U_FSpace").TitleObject.Caption = "Function Space"
            oEventGrid.Columns.Item("U_StartDt").TitleObject.Caption = "Start Date"
            oEventGrid.Columns.Item("U_EndDt").TitleObject.Caption = "End Date"
            oEventGrid.Columns.Item("U_FunDate").TitleObject.Caption = "Function Date"
            oEventGrid.CollapseLevel = 1

            If oEventGrid.DataTable.Rows.Count > 0 Then
                fillHeader(oForm)
            End If

            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub changeLabel(ByVal oForm As SAPbouiCOM.Form)
        Try
            oStatic = oForm.Items.Item("17").Specific
            oStatic.Caption = "Step " & oForm.PaneLevel & " of 3"
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub fillHeader(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oEventGrid = aForm.Items.Item("10").Specific
            oEventGrid.RowHeaders.TitleObject.Caption = "#"
            For index As Integer = 0 To oEventGrid.DataTable.Rows.Count - 1
                oEventGrid.RowHeaders.SetText(index, (index + 1).ToString())
            Next
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

#End Region

End Class
