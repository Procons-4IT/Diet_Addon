Imports SAPbobsCOM

Public Class clsMissedCall
    Inherits clsBase

    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Dim oStatic As SAPbouiCOM.StaticText
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private oCustomerGrid As SAPbouiCOM.Grid
    Private ocombo As SAPbouiCOM.ComboBoxColumn
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim strQuery As String
    Dim oGrid As SAPbouiCOM.Grid

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Public Sub LoadForm()
        Try
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_OMCT, frm_Z_OMCT)
            oForm = oApplication.SBO_Application.Forms.Item(strUID)
            oForm.Freeze(True)
            oForm.PaneLevel = 1
            initialize(oForm)
            addChooseFromListConditions(oForm)
            FillCombo(oForm)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Z_OMCT Then
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
                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                If pVal.ItemUID = "11" And pVal.ColUID = "CardCode" And pVal.Row > -1 Then
                                    oGrid = oForm.Items.Item("11").Specific
                                    Dim strCardCode As String = oGrid.DataTable.GetValue("CardCode", pVal.Row).ToString()
                                    Dim strCardName As String = oGrid.DataTable.GetValue("CardName", pVal.Row).ToString()
                                    Dim strProgramID As String = oGrid.DataTable.GetValue("DocEntry", pVal.Row).ToString()
                                    If Validation_Customer(oForm, strCardCode) Then
                                        Dim strFromDate, strToDate As String
                                        strFromDate = oForm.Items.Item("24").Specific.value
                                        strToDate = oForm.Items.Item("33").Specific.value
                                        Dim objPreSales As clsPreSalesOrder
                                        objPreSales = New clsPreSalesOrder
                                        objPreSales.LoadForm(strCardCode, strCardName, strFromDate, strToDate, strProgramID)
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        oApplication.Utilities.Message("Calories Or Address Not defined For Selected Customer...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And (oForm.PaneLevel = 1) Then
                                    oForm.PaneLevel = oForm.PaneLevel + 1
                                    changeLabel(oForm)
                                ElseIf pVal.ItemUID = "3" And (oForm.PaneLevel = 2) Then
                                    LoadProgramCustomers(oForm)
                                    oCustomerGrid = oForm.Items.Item("11").Specific
                                    If oCustomerGrid.DataTable.Rows.Count >= 1 Then
                                        oForm.PaneLevel = oForm.PaneLevel + 1
                                        changeLabel(oForm)
                                    Else
                                        If oCustomerGrid.DataTable.Rows.Count = 0 Then
                                            oApplication.Utilities.Message("No Customer Found for the Selection...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim strValue, strName As String
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If pVal.ItemUID = "8" Or pVal.ItemUID = "19" Then
                                        strValue = oDataTable.GetValue(CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).ChooseFromListAlias, 0)
                                        Try
                                            oForm.Items.Item(pVal.ItemUID).Specific.value = strValue
                                        Catch ex As Exception
                                            oApplication.Log.Trace_DIET_AddOn_Error(ex)
                                            oForm.Items.Item(pVal.ItemUID).Specific.value = strValue
                                        End Try
                                    ElseIf pVal.ItemUID = "35" Or pVal.ItemUID = "36" Then
                                        strValue = oDataTable.GetValue("CardCode", 0)
                                        strName = oDataTable.GetValue(CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).ChooseFromListAlias, 0)
                                        Try
                                            oForm.Items.Item(pVal.ItemUID).Specific.value = strName
                                        Catch ex As Exception
                                            oApplication.Log.Trace_DIET_AddOn_Error(ex)
                                            oForm.Items.Item(pVal.ItemUID).Specific.value = strName
                                        End Try
                                        If pVal.ItemUID = "35" Then
                                            oForm.Items.Item("8").Specific.value = strValue
                                        End If
                                        If pVal.ItemUID = "36" Then
                                            oForm.Items.Item("19").Specific.value = strValue
                                        End If
                                    ElseIf pVal.ItemUID = "12" Then
                                        strValue = oDataTable.GetValue(CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).ChooseFromListAlias, 0)
                                        strName = oDataTable.GetValue("ItemName", 0)
                                        Try
                                            oForm.Items.Item(pVal.ItemUID).Specific.value = strValue
                                        Catch ex As Exception
                                            oApplication.Log.Trace_DIET_AddOn_Error(ex)
                                            oForm.Items.Item(pVal.ItemUID).Specific.value = strValue
                                        End Try
                                    ElseIf pVal.ItemUID = "32" Then
                                        strValue = oDataTable.GetValue(CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).ChooseFromListAlias, 0)
                                        strName = oDataTable.GetValue("ItemName", 0)
                                        Try
                                            oForm.Items.Item(pVal.ItemUID).Specific.value = strValue
                                        Catch ex As Exception
                                            oApplication.Log.Trace_DIET_AddOn_Error(ex)
                                            oForm.Items.Item(pVal.ItemUID).Specific.value = strValue
                                        End Try
                                    End If
                                Catch ex As Exception
                                    oApplication.Log.Trace_DIET_AddOn_Error(ex)

                                End Try
                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "35" Then
                                    If CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value = "" Then
                                        CType(oForm.Items.Item("8").Specific, SAPbouiCOM.EditText).Value = ""
                                    End If
                                ElseIf pVal.ItemUID = "36" Then
                                    If CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value = "" Then
                                        CType(oForm.Items.Item("19").Specific, SAPbouiCOM.EditText).Value = ""
                                    End If
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
                Case mnu_Z_OMCT
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
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

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Validations"
    Private Function Validation(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim strFromDate, strToDate As String

            strFromDate = oForm.Items.Item("24").Specific.value
            strToDate = oForm.Items.Item("33").Specific.value
            If strFromDate = "" Then
                oApplication.Utilities.Message("Select From Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strToDate = "" Then
                oApplication.Utilities.Message("Select Till Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Dim strFrDt As String = oForm.Items.Item("24").Specific.value
            Dim strToDt As String = oForm.Items.Item("33").Specific.value
            If strFrDt.Length > 0 And strToDt.Length > 0 Then
                If CInt(strFrDt) > CInt(strToDt) Then
                    oApplication.Utilities.Message("From Date Should be Lesser than or Equal To Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Private Function Validation_Customer(ByVal oForm As SAPbouiCOM.Form, strCardCode As String) As Boolean
        Try
            strQuery = "Select ISNULL(U_CPAdj,0) As U_CPAdj From [@Z_OCPR] Where U_CardCode = '" & strCardCode & "'"
            Dim dblCalories As Double = oApplication.Utilities.getRecordSetValue(strQuery, "U_CPAdj")
            If dblCalories <= 0 Then
                Return False
            End If
            strQuery = "Select Address From CRD1 Where CardCode = '" & strCardCode & "' And AdresType = 'S'"
            Dim strAddress As String = oApplication.Utilities.getRecordSetValueString(strQuery, "Address")
            If strAddress.Length <= 0 Then
                Return False
            End If
            Return True
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#End Region

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Items.Item("1").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
            oForm.Items.Item("17").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
            oForm.DataSources.DataTables.Add("dtCustomers")
            oForm.Items.Item("13").TextStyle = 5
            'oForm.Items.Item("24").TextStyle = 5
            changeLabel(oForm)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub addChooseFromListConditions(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList

            oCFLs = oForm.ChooseFromLists

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

            'strQuery = "Select ItmsGrpCod From OITB Where U_Program = 'Y' "
            'oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRecordSet.DoQuery(strQuery)
            'If Not oRecordSet.EoF Then
            '    Dim strIG As String = oRecordSet.Fields.Item(0).Value
            '    oCFL = oCFLs.Item("CFL_5")
            '    oCons = oCFL.GetConditions()
            '    oCon = oCons.Add()
            '    oCon.Alias = "ItmsGrpCod"
            '    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '    oCon.CondVal = strIG
            '    oCFL.SetConditions(oCons)
            'End If

            strQuery = "Select ItmsGrpCod From OITB Where U_Program = 'Y' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oCFL = oCFLs.Item("CFL_5")
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

            'strQuery = "Select ItmsGrpCod From OITB Where U_Program = 'Y' "
            'oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRecordSet.DoQuery(strQuery)
            'If Not oRecordSet.EoF Then
            '    Dim strIG As String = oRecordSet.Fields.Item(0).Value
            '    oCFL = oCFLs.Item("CFL_6")
            '    oCons = oCFL.GetConditions()
            '    oCon = oCons.Add()
            '    oCon.Alias = "ItmsGrpCod"
            '    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '    oCon.CondVal = strIG
            '    oCFL.SetConditions(oCons)
            'End If

            strQuery = "Select ItmsGrpCod From OITB Where U_Program = 'Y' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oCFL = oCFLs.Item("CFL_6")
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

            oCFL = oCFLs.Item("CFL_7")
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

            oCFL = oCFLs.Item("CFL_8")
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

    Private Sub LoadProgramCustomers(ByVal aform As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            Dim strqry As String
            Dim strFromCust = "", strToCust = "", strProgram1 = "", strProgram2 = "", strCustGroup1 = "", strCustGroup2 = "", _
                strFromDate = "", strToDate As String = "" _
                , strItemGroup1 = "", strItemGroup2 = ""

            strFromCust = oForm.Items.Item("8").Specific.value
            strToCust = oForm.Items.Item("19").Specific.value

            strCustGroup1 = CType(oForm.Items.Item("10").Specific, SAPbouiCOM.ComboBox).Value.Trim()
            strCustGroup2 = CType(oForm.Items.Item("31").Specific, SAPbouiCOM.ComboBox).Value.Trim()

            strItemGroup1 = CType(oForm.Items.Item("34").Specific, SAPbouiCOM.ComboBox).Value.Trim()
            strItemGroup2 = CType(oForm.Items.Item("34_").Specific, SAPbouiCOM.ComboBox).Value.Trim()

            strProgram1 = oForm.Items.Item("12").Specific.value
            strProgram2 = oForm.Items.Item("32").Specific.value

            strFromDate = oForm.Items.Item("24").Specific.value
            strToDate = oForm.Items.Item("33").Specific.value

            oForm.Items.Item("23").Specific.value = strProgram1
            oForm.Items.Item("25").Specific.value = strProgram2

            oCustomerGrid = oForm.Items.Item("11").Specific
            oCustomerGrid.DataTable = oForm.DataSources.DataTables.Item("dtCustomers")

            strqry = " Select DISTINCT T0.CardCode,CardName,T1.DocEntry,T1.U_PrgName,T1.U_PFromDate,T1.U_PToDate,Convert(VarChar(100),T0.aliasname) As 'Dietitian' From OCRD T0  "
            strqry += " JOIN [@Z_OCPM] T1 On T1.U_CardCode = T0.CardCode "
            strqry += " JOIN [@Z_CPM1] T2 On T1.DocEntry = T2.DocEntry "
            strqry += " And T1.U_RemDays > 0 And ISNULL(T1.U_Transfer,'N') = 'N' "
            strqry += " LEFT OUTER JOIN [@Z_OCPR] T3 On T0.CardCode = T3.U_CardCode "
            strqry += " JOIN OITM T4 ON T4.ItemCode = T1.U_PrgCode "
            strqry += " LEFT OUTER JOIN OITB T5 On T4.ItmsGrpCod = T5.ItmsGrpCod "
            strqry += " Where CardType = 'C' And ISNULL(T2.U_ONOFFSTA,'O') = 'O' AND ISNULL(T3.U_ONOFFSTA,'O') = 'O' AND ISNULL(T2.U_AppStatus,'I') = 'I' "

            If strFromCust.Length > 0 And strToCust.Length > 0 Then
                strqry += " And T0.CardCode Between '" + strFromCust + "' AND '" + strToCust + "'"
            End If

            If strProgram1.Length > 0 And strProgram2.Length > 0 Then
                strqry += " And T1.U_PrgCode Between '" + strProgram1 + "' And '" + strProgram2 + "'"
            End If

            If strCustGroup1.Length > 0 And strCustGroup2.Length > 0 Then
                strqry += " And T0.GroupCode Between '" + strCustGroup1 + "' And '" + strCustGroup2 + "'"
            End If

            If strItemGroup1.Length > 0 And strItemGroup2.Length > 0 Then
                strqry += " And T4.ItmsGrpCod Between '" + strItemGroup1 + "' And '" + strItemGroup2 + "'"
            End If

            If strFromDate.Length > 0 And strToDate.Length > 0 Then
                strqry += " And "
                strqry += " ( "
                strqry += " Convert(VarChar(8),T2.U_PrgDate,112) >= '" + strFromDate + "' "
                strqry += " AND Convert(VarChar(8),T2.U_PrgDate,112) <= '" + strToDate + "' "
                strqry += " ) "
            End If
            strqry += "  And U_PrgDate Between T1.U_PFromDate And T1.U_PToDate "
            strqry += "  And  "
            strqry += "  (  "
            strqry += " (T1.U_PFromDate < T3.U_SuFrDt And ISNULL(T3.U_SuToDt,'') = '') "
            strqry += "  OR  "
            strqry += "  (1 = 1)  "
            strqry += "  ) "

            oCustomerGrid.DataTable.ExecuteQuery(strqry)

            oCustomerGrid.Columns.Item("CardCode").TitleObject.Caption = "Customer Code"
            oCustomerGrid.Columns.Item("CardCode").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditTextColumn = oCustomerGrid.Columns.Item("CardCode")
            oEditTextColumn.LinkedObjectType = "2"
            oCustomerGrid.Columns.Item("CardCode").Editable = False

            oCustomerGrid.Columns.Item("CardName").TitleObject.Caption = "Customer Name"
            oCustomerGrid.Columns.Item("CardName").Editable = False

            oCustomerGrid.Columns.Item("U_PrgName").TitleObject.Caption = "Program Name"
            oCustomerGrid.Columns.Item("U_PrgName").Editable = False

            oCustomerGrid.Columns.Item("U_PFromDate").TitleObject.Caption = "Program From"
            oCustomerGrid.Columns.Item("U_PFromDate").Editable = False

            oCustomerGrid.Columns.Item("U_PToDate").TitleObject.Caption = "Program To"
            oCustomerGrid.Columns.Item("U_PToDate").Editable = False


            oCustomerGrid.Columns.Item("DocEntry").Visible = False
            oCustomerGrid.Columns.Item("Dietitian").Editable = False

            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub FillCombo(ByVal aForm As SAPbouiCOM.Form)
        Try
            Dim oTempRec As SAPbobsCOM.Recordset
            oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            oCombobox = aForm.Items.Item("10").Specific
            For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
                oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oCombobox.ValidValues.Add("", "")
            oTempRec.DoQuery("Select GroupCode,GroupName From OCRG Where GroupType = 'C'")
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                oCombobox.ValidValues.Add(oTempRec.Fields.Item("GroupCode").Value, oTempRec.Fields.Item("GroupName").Value)
                oTempRec.MoveNext()
            Next

            oCombobox = aForm.Items.Item("31").Specific
            For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
                oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oCombobox.ValidValues.Add("", "")
            oTempRec.DoQuery("Select GroupCode,GroupName From OCRG Where GroupType = 'C'")
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                oCombobox.ValidValues.Add(oTempRec.Fields.Item("GroupCode").Value, oTempRec.Fields.Item("GroupName").Value)
                oTempRec.MoveNext()
            Next

            oCombobox = aForm.Items.Item("34").Specific
            For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
                oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oCombobox.ValidValues.Add("", "")
            oTempRec.DoQuery("Select ItmsGrpCod,ItmsGrpNam From OITB Where U_Program = 'Y'")
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                oCombobox.ValidValues.Add(oTempRec.Fields.Item("ItmsGrpCod").Value, oTempRec.Fields.Item("ItmsGrpNam").Value)
                oTempRec.MoveNext()
            Next

            oCombobox = aForm.Items.Item("34_").Specific
            For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
                oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oCombobox.ValidValues.Add("", "")
            oTempRec.DoQuery("Select ItmsGrpCod,ItmsGrpNam From OITB Where U_Program = 'Y'")
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                oCombobox.ValidValues.Add(oTempRec.Fields.Item("ItmsGrpCod").Value, oTempRec.Fields.Item("ItmsGrpNam").Value)
                oTempRec.MoveNext()
            Next

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
        End Try
    End Sub

    Private Sub reDrawForm(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)

            oForm.Items.Item("11").Top = oForm.Items.Item("13").Top + oForm.Items.Item("13").Height + 1
            oForm.Items.Item("11").Height = (oForm.Height - 140)
            oForm.Items.Item("11").Width = oForm.Width - 30
            ' oForm.Items.Item("24").Top = oForm.Items.Item("11").Top + oForm.Items.Item("11").Height + 2

            oForm.Freeze(False)
        Catch ex As Exception
            'oApplication.Log.Trace_DIET_AddOn_Error(ex)
            'oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Private Sub changeLabel(ByVal oForm As SAPbouiCOM.Form)
        Try
            oStatic = oForm.Items.Item("17").Specific
            oStatic.Caption = "Step " & oForm.PaneLevel & " of 3"
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

End Class
