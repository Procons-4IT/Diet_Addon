Public Class clsCustomerProfile
    Inherits clsBase

    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oMatrix As SAPbouiCOM.Matrix
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines1 As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines2 As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines3 As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines4 As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines5 As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines6 As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines7 As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines8 As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines9 As SAPbouiCOM.DBDataSource
    Private oDTPrograms As SAPbouiCOM.DataTable
    Private oDTPreSales As SAPbouiCOM.DataTable
    Private oRecordSet As SAPbobsCOM.Recordset
    Private objForm As SAPbouiCOM.Form
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oEditText As SAPbouiCOM.EditText
    Private oMode As SAPbouiCOM.BoFormMode
    Private InvForConsumedItems, count As Integer
    Private blnFlag As Boolean = False
    Dim oGrid As SAPbouiCOM.Grid
    Public intSelectedMatrixrow As Integer = 0
    Private RowtoDelete As Integer
    Private oCombo As SAPbouiCOM.ComboBox
    Private strQuery As String = String.Empty
    Private oLoadForm As SAPbouiCOM.Form

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Sub LoadForm()
        Try
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_OCPR) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_OCPR, frm_Z_OCPR)
            oForm = oApplication.SBO_Application.Forms.Item(strUID)
            oForm.Freeze(True)
            initialize(oForm)

            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            oForm.EnableMenu(mnu_ADD, False)
            oForm.EnableMenu(mnu_FIND, False)
            Try
                loadComboColumn(oForm)
            Catch ex As Exception

            End Try

            oForm.DataSources.DataTables.Add("Programs")
            oForm.DataSources.DataTables.Add("PerSales")
            loadProgramsAndVisits(oForm)
            oForm.Freeze(False)
            oForm.DataSources.UserDataSources.Add("PicSource", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 1000)
        Catch ex As Exception
            oForm.Freeze(False)
        End Try
    End Sub

    Public Sub LoadForm(ByVal strDocEntry As String)
        Try
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_OCPR) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_OCPR, frm_Z_OCPR)
            oForm = oApplication.SBO_Application.Forms.Item(strUID)
            oForm.Freeze(True)
            initialize(oForm)
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            oForm.EnableMenu(mnu_ADD, False)
            oForm.DataSources.DataTables.Add("Programs")
            oForm.DataSources.DataTables.Add("PreSales")
            oForm.Freeze(False)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            oForm.Items.Item("13").Specific.value = strDocEntry
            Try
                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Catch ex As Exception

            End Try
            loadProgramsAndVisits(oForm)
            oForm.EnableMenu(mnu_FIND, False)
            Try
                loadComboColumn(oForm)
            Catch ex As Exception

            End Try
            oForm.Items.Item("13").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("9").Enabled = False
            oForm.Items.Item("10").Enabled = False
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

    Public Sub LoadForm(ByVal strCardCode As String, ByVal strCardName As String)
        Try
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_OCPR) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_OCPR, frm_Z_OCPR)
            oForm = oApplication.SBO_Application.Forms.Item(strUID)
            oForm.Freeze(True)
            initialize(oForm)
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            oForm.EnableMenu(mnu_ADD, False)
            oForm.EnableMenu(mnu_FIND, False)
            oForm.DataSources.DataTables.Add("Programs")
            oForm.DataSources.DataTables.Add("PreSales")
            oForm.Items.Item("9").Specific.value = strCardCode
            oForm.Items.Item("10").Specific.value = strCardName
            oForm.Items.Item("13").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Try
                loadComboColumn(oForm)
            Catch ex As Exception

            End Try
            loadProgramsAndVisits(oForm)
            oForm.Items.Item("9").Enabled = False
            oForm.Items.Item("10").Enabled = False
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Z_OCPR Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    oForm.Freeze(True)
                                    If validation(oForm) = False Then
                                        oForm.Freeze(False)
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
                                            oForm.Freeze(False)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "53" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPR")
                                    Dim objCProgram As clsCustomerProgram
                                    objCProgram = New clsCustomerProgram
                                    Dim strDiscount As String = oApplication.Utilities.getRecordSetValue("Select Discount From OCRD Where CardCode = '" & oDBDataSource.GetValue("U_CardCode", 0).ToString().Trim() & "'", "Discount")
                                    Dim strCurrency As String = oApplication.Utilities.getRecordSetValueString("Select Currency From OCRD Where CardCode = '" & oDBDataSource.GetValue("U_CardCode", 0).ToString().Trim() & "'", "Currency")
                                    objCProgram.LoadForm(oDBDataSource.GetValue("U_CardCode", 0).ToString().Trim() _
                                                         , oDBDataSource.GetValue("U_CardName", 0).ToString().Trim(), strDiscount, oDBDataSource.GetValue("U_DisRemarks", 0).ToString().Trim(), strCurrency)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                If (Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE) Then
                                    alldataSource(oForm)
                                    If (pVal.ItemUID = "24" Or pVal.ItemUID = "25" Or pVal.ItemUID = "26" _
                                             Or pVal.ItemUID = "27" Or pVal.ItemUID = "28" Or pVal.ItemUID = "29" _
                                             Or pVal.ItemUID = "36" Or pVal.ItemUID = "41" Or pVal.ItemUID = "42" Or pVal.ItemUID = "43") Then
                                        oForm.Freeze(True)
                                        'changePane(oForm, pVal.ItemUID)
                                        If pVal.ItemUID = "36" Then
                                            oForm.Items.Item("37").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                End If

                                If (pVal.ItemUID = "3" Or pVal.ItemUID = "4" Or pVal.ItemUID = "5" Or pVal.ItemUID = "6" _
                                        Or pVal.ItemUID = "39" Or pVal.ItemUID = "40" Or pVal.ItemUID = "44" _
                                        Or pVal.ItemUID = "45" Or pVal.ItemUID = "46") Then
                                    If (Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE) Then
                                        intSelectedMatrixrow = pVal.Row
                                        If pVal.ItemUID = "44" Then
                                            oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                            If pVal.Row < oMatrix.VisualRowCount Then
                                                If pVal.ColUID <> "V_7" Then
                                                    BubbleEvent = False
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                If pVal.ItemUID = "8" And pVal.ColUID = "DocEntry" And pVal.Row > -1 Then
                                    oGrid = oForm.Items.Item("8").Specific
                                    Dim strPID As String = oGrid.DataTable.GetValue(pVal.ColUID, pVal.Row).ToString()
                                    Dim objPreSales As clsPreSalesOrder
                                    objPreSales = New clsPreSalesOrder
                                    objPreSales.LoadForm(strPID)
                                    BubbleEvent = False
                                    Exit Sub
                                ElseIf pVal.ItemUID = "7" And pVal.ColUID = "DocEntry" And pVal.Row > -1 Then
                                    oGrid = oForm.Items.Item("7").Specific
                                    Dim strPID As String = oGrid.DataTable.GetValue(pVal.ColUID, pVal.Row).ToString()
                                    Dim objCProgram As clsCustomerProgram
                                    objCProgram = New clsCustomerProgram
                                    objCProgram.LoadForm(strPID)
                                    BubbleEvent = False
                                    Exit Sub
                                ElseIf pVal.ItemUID = "7" And pVal.ColUID = "U_TrnRef" And pVal.Row > -1 Then
                                    oGrid = oForm.Items.Item("7").Specific
                                    Dim strPID As String = oGrid.DataTable.GetValue(pVal.ColUID, pVal.Row).ToString()
                                    Dim objPreSales As clsProgramTransfer
                                    objPreSales = New clsProgramTransfer
                                    objPreSales.LoadForm(strPID)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or _
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or _
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    alldataSource(oForm)

                                    If pVal.ItemUID = "6" And pVal.ColUID = "V_0" And pVal.Row > 0 Then
                                        oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                        Dim strFDate As String = oApplication.Utilities.getMatrixValues(oMatrix, pVal.ColUID, pVal.Row)

                                        If strFDate <> "" Then
                                            Dim oRecordExist As SAPbobsCOM.Recordset
                                            oRecordExist = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                strQuery = " Select LineId From [@Z_CPR4] Where DocEntry = '" & oDBDataSource.GetValue("DocEntry", 0).ToString & "'"
                                                strQuery += " AND '" & strFDate & "' =  Convert(VarChar(8),U_ExDate,112) "
                                                oRecordExist.DoQuery(strQuery)
                                                If oRecordExist.EoF Then
                                                    If Not oApplication.Utilities.validateDate(oForm, strFDate, -1) Then
                                                        Dim strMessage As String = "Exclude From Date Should be Greater than Or Equal Yesterday Date..."
                                                        If Not oApplication.Utilities.valCustomerProgramDate(oForm, oDBDataSource.GetValue("U_CardCode", 0).ToString.Trim(), strFDate) Then
                                                            strMessage = " Open Delivery Document Exists & Linked to Invoice "
                                                            oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                            BubbleEvent = False
                                                            Exit Sub
                                                        End If
                                                    Else
                                                        If Not oApplication.Utilities.valCustomerProgramDate(oForm, oDBDataSource.GetValue("U_CardCode", 0).ToString.Trim(), strFDate) Then
                                                            Dim strMessage As String = " Open Delivery Document Exists & Linked to Invoice "
                                                            oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                            BubbleEvent = False
                                                            Exit Sub
                                                        End If
                                                    End If
                                                End If
                                            ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                                If Not oApplication.Utilities.validateDate(oForm, strFDate, -1) Then
                                                    'oApplication.Utilities.Message("Exclude From Date Should be Greater than Or Equal Yesterday Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    'BubbleEvent = False
                                                    'Exit Sub
                                                    Dim strMessage As String = "Exclude From Date Should be Greater than Or Equal Yesterday Date..."
                                                    If Not oApplication.Utilities.valCustomerProgramDate(oForm, oDBDataSource.GetValue("U_CardCode", 0).ToString.Trim(), strFDate) Then
                                                        strMessage = " Open Delivery Document Exists & Linked to Invoice  "
                                                        oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        BubbleEvent = False
                                                        Exit Sub
                                                    End If

                                                End If
                                            End If

                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                                If Not validateExcludeVsRemoveDates(oForm, strFDate, "45") Then
                                                    oApplication.Utilities.Message("The Selected Date already entered in Remove Date Please Check...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    BubbleEvent = False
                                                    Exit Sub
                                                End If
                                            End If

                                        End If

                                        'ElseIf pVal.ItemUID = "6" And pVal.ColUID = "V_1" And pVal.Row > 0 Then
                                        '    oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                        '    Dim strFDate As String = oApplication.Utilities.getMatrixValues(oMatrix, pVal.ColUID, pVal.Row)
                                        '    If strFDate <> "" Then
                                        '        If Not oApplication.Utilities.validateDate(oForm, strFDate, -1) Then
                                        '            oApplication.Utilities.Message("Exclude From Date Should be Greater than Or Equal Yesterday Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        '            BubbleEvent = False
                                        '            Exit Sub
                                        '        End If
                                        'End If

                                    ElseIf pVal.ItemUID = "45" And (pVal.ColUID = "V_0" Or pVal.ColUID = "V_1") Then
                                        oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                        Dim strFDate As String = oApplication.Utilities.getMatrixValues(oMatrix, pVal.ColUID, pVal.Row)

                                        If strFDate <> "" Then

                                            Dim oRecordExist As SAPbobsCOM.Recordset
                                            oRecordExist = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                strQuery = " Select LineId From [@Z_CPR8] Where DocEntry = '" & oDBDataSource.GetValue("DocEntry", 0).ToString.Trim() & "'"
                                                strQuery += " AND '" & strFDate & "' Between  Convert(VarChar(8),U_FDate,112) And Convert(VarChar(8),U_TDate,112) "
                                                oRecordExist.DoQuery(strQuery)
                                                If oRecordExist.EoF Then
                                                    If Not oApplication.Utilities.validateDate(oForm, strFDate, -1) Then
                                                        'oApplication.Utilities.Message("Exclude From Date Should be Greater than Or Equal Yesterday Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        'BubbleEvent = False
                                                        'Exit Sub
                                                        Dim strMessage As String = "Remove From Date Should be Greater than Or Equal Yesterday Date..."
                                                        If Not oApplication.Utilities.valCustomerProgramDate(oForm, oDBDataSource.GetValue("U_CardCode", 0).ToString.Trim(), strFDate) Then
                                                            strMessage = " Open Delivery Document Exists & Linked to Invoice  "
                                                            oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                            BubbleEvent = False
                                                            Exit Sub
                                                        End If
                                                    Else
                                                        If Not oApplication.Utilities.valCustomerProgramDate(oForm, oDBDataSource.GetValue("U_CardCode", 0).ToString.Trim(), strFDate) Then
                                                            Dim strMessage As String = " Open Delivery Document Exists & Linked to Invoice "
                                                            oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                            BubbleEvent = False
                                                            Exit Sub
                                                        End If
                                                    End If
                                                End If

                                                'strQuery = " Select LineId From [@Z_CPR8] Where DocEntry = '" & oDBDataSource.GetValue("DocEntry", 0).ToString & "'"
                                                'strQuery += " AND '" & strFDate & "' Between  Convert(VarChar(8),U_FDate,112) And Convert(VarChar(8),U_Tdate,112) "
                                                'oRecordExist.DoQuery(strQuery)
                                                'If oRecordExist.EoF Then
                                                '    'Block Remove Date when Invoiced.
                                                '    If Not oApplication.Utilities.valCustomerProgramDate(oForm, oDBDataSource.GetValue("U_CardCode", 0).ToString.Trim(), strFDate) Then
                                                '        Dim strMessage As String = " Open Delivery Document Exists & Linked to Invoice..."
                                                '        oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                '        BubbleEvent = False
                                                '        Exit Sub
                                                '    End If
                                                'End If
                                                

                                            ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                                If Not oApplication.Utilities.validateDate(oForm, strFDate, -1) Then

                                                    'oApplication.Utilities.Message("Remove Date Should be Greater than Or Equal Yesterday Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    'BubbleEvent = False
                                                    'Exit Sub

                                                    Dim strMessage As String = "Remove From Date Should be Greater than Or Equal Yesterday Date..."
                                                    If Not oApplication.Utilities.valCustomerProgramDate(oForm, oDBDataSource.GetValue("U_CardCode", 0).ToString.Trim, strFDate) Then
                                                        strMessage = " Open Delivery Document Exists & Linked to Invoice..."
                                                        oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        BubbleEvent = False
                                                        Exit Sub
                                                    End If

                                                End If
                                            End If

                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                                If Not validateExcludeVsRemoveDates(oForm, strFDate, "6") Then
                                                    oApplication.Utilities.Message("The Selected Date already entered in Exclude Date Please Check...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    BubbleEvent = False
                                                    Exit Sub
                                                End If
                                            End If

                                            Dim strFDate1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", pVal.Row)
                                            Dim strTDate1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", pVal.Row)
                                            If strFDate1 <> "" And strTDate1 <> "" Then
                                                If CInt(strFDate1) > CInt(strTDate1) Then
                                                    oApplication.Utilities.Message("Remove From Date Should be Less than or Equal to To Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    BubbleEvent = False
                                                    Exit Sub
                                                End If
                                            End If

                                        End If
                                    ElseIf pVal.ItemUID = "46" And (pVal.ColUID = "V_0" Or pVal.ColUID = "V_1") Then
                                        oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                        Dim strFDate As String = oApplication.Utilities.getMatrixValues(oMatrix, pVal.ColUID, pVal.Row)
                                        If strFDate <> "" Then

                                            Dim oRecordExist As SAPbobsCOM.Recordset
                                            oRecordExist = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                strQuery = " Select LineId From [@Z_CPR9] Where DocEntry = '" & oDBDataSource.GetValue("DocEntry", 0).ToString & "'"
                                                strQuery += " And '" & strFDate & "' Between  Convert(VarChar(8),U_FDate,112) And Convert(VarChar(8),U_TDate,112) "
                                                oRecordExist.DoQuery(strQuery)
                                                If oRecordExist.EoF Then
                                                    If Not oApplication.Utilities.validateDate(oForm, strFDate, -1) Then
                                                        oApplication.Utilities.Message("Exclude From Date Should be Greater than Or Equal Yesterday Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        BubbleEvent = False
                                                        Exit Sub
                                                    Else
                                                        If Not oApplication.Utilities.valCustomerProgramDate(oForm, oDBDataSource.GetValue("U_CardCode", 0).ToString.Trim(), strFDate) Then
                                                            Dim strMessage As String = " Open Delivery Document Exists & Linked to Invoice "
                                                            oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                            BubbleEvent = False
                                                            Exit Sub
                                                        End If
                                                    End If
                                                End If
                                            ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                                If Not oApplication.Utilities.validateDate(oForm, strFDate, -1) Then
                                                    oApplication.Utilities.Message("Suspend Date Should be Greater than Or Equal Yesterday Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    BubbleEvent = False
                                                    Exit Sub
                                                End If
                                            End If

                                            Dim strFDate1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", pVal.Row)
                                            Dim strTDate1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", pVal.Row)
                                            If strFDate1 <> "" And strTDate1 <> "" Then
                                                If CInt(strFDate1) > CInt(strTDate1) Then
                                                    oApplication.Utilities.Message("Suspend From Date Should be Less than or Equal to To Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    BubbleEvent = False
                                                    Exit Sub
                                                End If
                                            End If

                                        End If
                                    ElseIf pVal.ItemUID = "39" And (pVal.ColUID = "V_0" Or pVal.ColUID = "V_1") Then
                                        oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                        Dim strFDate1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", pVal.Row)
                                        Dim strTDate1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0_1", pVal.Row)
                                        If strFDate1 <> "" And strTDate1 <> "" Then
                                            If CInt(strFDate1) > CInt(strTDate1) Then
                                                oApplication.Utilities.Message("Address From Date Should be Less than or Equal to To Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                                    Case "22"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_ADD_ROW)
                                    Case "23"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                                    Case "1"
                                        If pVal.Action_Success Then
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                                oForm.Close()
                                            ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                oForm.Items.Item("24").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                oApplication.SBO_Application.Menus.Item("1304").Activate()
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
                                        If ((pVal.ItemUID = "3" Or pVal.ItemUID = "4") _
                                                  And (pVal.ColUID = "V_0")) Then
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
                                                        oDBDataSourceLines1.SetValue("U_DLikeItem", pVal.Row + index - 1, oDataTable.GetValue("U_Code", index))
                                                        oDBDataSourceLines1.SetValue("U_Name", pVal.Row + index - 1, oDataTable.GetValue("U_Name", index))
                                                    Next
                                                Case "4"
                                                    For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                        oDBDataSourceLines2.SetValue("LineId", pVal.Row + index - 1, (pVal.Row + index).ToString())
                                                        oDBDataSourceLines2.SetValue("U_MSCode", pVal.Row + index - 1, oDataTable.GetValue("U_Code", index))
                                                        oDBDataSourceLines2.SetValue("U_Name", pVal.Row + index - 1, oDataTable.GetValue("U_Name", index))
                                                    Next
                                            End Select
                                            oMatrix.LoadFromDataSource()
                                            oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf ((pVal.ItemUID = "44") _
                                              And (pVal.ColUID = "V_0_0")) Then
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
                                                Case "44"
                                                    For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                        oDBDataSourceLines7.SetValue("U_CPAdj", pVal.Row + index - 1, oDataTable.GetValue("U_Calories", index))
                                                    Next
                                            End Select
                                            oMatrix.LoadFromDataSource()
                                            oMatrix.Columns.Item("V_0_0").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf pVal.ItemUID = "21" Then
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 0 Then
                                                oDBDataSource.SetValue("U_CPAdj", 0, oDataTable.GetValue("U_Calories", 0))
                                                Dim oRecordSet As SAPbobsCOM.Recordset
                                                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                oRecordSet.DoQuery("Select U_CPAdj From [@Z_OCPR] Where DocEntry = '" & oDBDataSource.GetValue("DocEntry", 0).ToString().Trim() & "'")
                                                If Not oRecordSet.EoF Then
                                                    If oDBDataSource.GetValue("U_CPAdj", 0).ToString().Trim() <> oRecordSet.Fields.Item(0).Value.ToString() Then
                                                        If validate_Calories(oForm) Then
                                                            AddRow_1(oForm, "44")
                                                            fillStandard(oForm, oDBDataSource.GetValue("U_CPAdj", 0).ToString().Trim())
                                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                        End If
                                                    End If
                                                Else
                                                    If validate_Calories(oForm) Then
                                                        AddRow_1(oForm, "44")
                                                        fillStandard(oForm, oDBDataSource.GetValue("U_CPAdj", 0).ToString().Trim())
                                                    End If

                                                End If
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception

                                End Try
                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    If ((pVal.ItemUID = "3" Or pVal.ItemUID = "4") _
                                                And (pVal.ColUID = "V_0") And pVal.Row > 0) Then

                                        alldataSource(oForm)
                                        oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                        oMatrix.FlushToDataSource()
                                        oMatrix.LoadFromDataSource()
                                        oMatrix.FlushToDataSource()
                                        Select Case pVal.ItemUID
                                            Case "3"
                                                Dim strItemCode As String = oDBDataSourceLines1.GetValue("U_DLikeItem", pVal.Row - 1)
                                                Dim strItemName As String = oDBDataSourceLines1.GetValue("U_Name", pVal.Row - 1)
                                                If strItemCode.Trim().Length = 0 And strItemName.Trim().Length > 0 Then
                                                    oDBDataSourceLines1.SetValue("U_Name", pVal.Row - 1, "")
                                                    oMatrix.LoadFromDataSource()
                                                    oMatrix.FlushToDataSource()
                                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                End If
                                            Case "4"
                                                Dim strItemCode As String = oDBDataSourceLines2.GetValue("U_MSCode", pVal.Row - 1)
                                                Dim strItemName As String = oDBDataSourceLines2.GetValue("U_Name", pVal.Row - 1)
                                                If strItemCode.Trim().Length = 0 And strItemName.Trim().Length > 0 Then
                                                    oDBDataSourceLines2.SetValue("U_Name", pVal.Row - 1, "")
                                                    oMatrix.LoadFromDataSource()
                                                    oMatrix.FlushToDataSource()
                                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                End If
                                        End Select
                                        'ElseIf pVal.ItemUID = "21" Then
                                        '    Dim oRecordSet As SAPbobsCOM.Recordset
                                        '    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        '    oRecordSet.DoQuery("Select U_CPAdj From [@Z_OCPR] Where DocEntry = '" & oDBDataSource.GetValue("DocEntry", 0).ToString().Trim() & "'")
                                        '    If Not oRecordSet.EoF Then
                                        '        If oDBDataSource.GetValue("U_CPAdj", 0).ToString().Trim() <> oRecordSet.Fields.Item(0).Value.ToString() Then
                                        '            AddRow(oForm, "44")
                                        '            fillStandard(oForm, oDBDataSource.GetValue("U_CPAdj", 0).ToString().Trim())
                                        '            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        '        End If
                                        '    End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                If oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Or oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    Try
                                        If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                            Dim strActiveItem As String = oForm.ActiveItem
                                            If strActiveItem = "49" Or strActiveItem = "50" Or strActiveItem = "51" Or strActiveItem = "52" Then
                                                Dim intItem As Integer = CInt(strActiveItem)
                                                Dim intItem1 As Integer = CInt(strActiveItem) + 1
                                                If intItem1 >= 49 And intItem1 <= 52 Then

                                                Else
                                                    intItem1 -= 1
                                                    If intItem1.ToString() = strActiveItem Then
                                                        intItem1 -= 1
                                                    End If
                                                End If
                                                oForm.Items.Item(intItem1.ToString()).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                reDrawForm(oForm)
                                                oForm.Items.Item(strActiveItem).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            Else
                                                reDrawForm(oForm)
                                            End If
                                        End If
                                    Catch ex As Exception

                                    End Try
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                If (pVal.ItemUID = "39" Or pVal.ItemUID = "40") And pVal.ColUID = "V_2" And pVal.Row > 0 Then
                                    fillBuilding(oForm, pVal.ItemUID, pVal.Row)
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
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
                        Case mnu_DELETE_ROW
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            If pVal.BeforeAction = True Then
                                Dim strItem As String = getMatrixItem(oForm)
                                If strItem = "6" Then
                                    If intSelectedMatrixrow > 0 Then
                                        oMatrix = oForm.Items.Item(strItem).Specific
                                        'If CType(oMatrix.Columns.Item("V_2").Cells.Item(intSelectedMatrixrow).Specific, SAPbouiCOM.ComboBox).Value = "Y" Then
                                        'oApplication.Utilities.Message("Cannot Delete Row As its Already Applied...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        'BubbleEvent = False
                                        'Exit Sub
                                        'End If
                                    End If
                                ElseIf strItem = "45" Then
                                    If intSelectedMatrixrow > 0 Then
                                        oMatrix = oForm.Items.Item(strItem).Specific
                                        'Dim strFDate As String = CType(oMatrix.Columns.Item("V_0").Cells.Item(intSelectedMatrixrow).Specific, SAPbouiCOM.EditText).Value
                                        'Dim strTDate As String = CType(oMatrix.Columns.Item("V_1").Cells.Item(intSelectedMatrixrow).Specific, SAPbouiCOM.EditText).Value
                                        'If CType(oMatrix.Columns.Item("V_3").Cells.Item(intSelectedMatrixrow).Specific, SAPbouiCOM.ComboBox).Value = "Y" Then
                                        '    oApplication.Utilities.Message("Cannot Delete Row As  Applied...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        '    BubbleEvent = False
                                        '    Exit Sub
                                        'End If
                                    End If
                                ElseIf strItem = "46" Then
                                    If intSelectedMatrixrow > 0 Then
                                        oMatrix = oForm.Items.Item(strItem).Specific
                                        'If CType(oMatrix.Columns.Item("V_3").Cells.Item(intSelectedMatrixrow).Specific, SAPbouiCOM.ComboBox).Value = "Y" Then
                                        'oApplication.Utilities.Message("Cannot Delete Row As its Already Applied...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        'BubbleEvent = False
                                        'Exit Sub
                                        'End If
                                    End If
                                End If
                            End If
                    End Select
                Case False
                    Select Case pVal.MenuUID
                        Case mnu_Z_OCPR
                            LoadForm()
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
                    End Select
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
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            If oForm.TypeEx = frm_Z_OCPR Then
                Select Case BusinessObjectInfo.BeforeAction
                    Case True
                    Case False
                        Select Case BusinessObjectInfo.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If BusinessObjectInfo.ActionSuccess Then
                                    loadProgramsAndVisits(oForm)

                                    'Disable Exclude when Applied
                                    oMatrix = oForm.Items.Item("6").Specific
                                    For index As Integer = 1 To oMatrix.VisualRowCount
                                        If CType(oMatrix.Columns.Item("V_2").Cells.Item(index).Specific, SAPbouiCOM.ComboBox).Value = "Y" Then
                                            'oMatrix.CommonSetting.SetCellEditable(index, 1, False)
                                        Else
                                            oMatrix.CommonSetting.SetCellEditable(index, 1, True)
                                        End If
                                    Next

                                    'Disable Remove when Applied
                                    oMatrix = oForm.Items.Item("45").Specific
                                    For index As Integer = 1 To oMatrix.VisualRowCount
                                        If CType(oMatrix.Columns.Item("V_3").Cells.Item(index).Specific, SAPbouiCOM.ComboBox).Value = "Y" Then
                                            'oMatrix.CommonSetting.SetCellEditable(index, 1, False)
                                            'oMatrix.CommonSetting.SetCellEditable(index, 2, False)
                                        Else
                                            oMatrix.CommonSetting.SetCellEditable(index, 1, True)
                                            oMatrix.CommonSetting.SetCellEditable(index, 2, True)
                                        End If
                                    Next

                                    'Disable Suspend when Applied
                                    oMatrix = oForm.Items.Item("46").Specific
                                    For index As Integer = 1 To oMatrix.VisualRowCount
                                        If CType(oMatrix.Columns.Item("V_3").Cells.Item(index).Specific, SAPbouiCOM.ComboBox).Value = "Y" Then
                                            'oMatrix.CommonSetting.SetCellEditable(index, 1, False)
                                            If CType(oMatrix.Columns.Item("V_1").Cells.Item(index).Specific, SAPbouiCOM.EditText).Value <> "" Then
                                                'oMatrix.CommonSetting.SetCellEditable(index, 2, False)
                                            Else
                                                oMatrix.CommonSetting.SetCellEditable(index, 2, True)
                                            End If
                                        Else
                                            oMatrix.CommonSetting.SetCellEditable(index, 1, True)
                                            oMatrix.CommonSetting.SetCellEditable(index, 2, True)
                                        End If
                                    Next

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, _
                                SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                If BusinessObjectInfo.ActionSuccess Then
                                    oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                                    Dim oXmlDoc As System.Xml.XmlDocument = New Xml.XmlDocument()
                                    oXmlDoc.LoadXml(BusinessObjectInfo.ObjectKey)
                                    Dim strDocEntry As String = oXmlDoc.SelectSingleNode("/Customer_ProfileParams/DocEntry").InnerText

                                    oLoadForm = Nothing
                                    oLoadForm = oApplication.Utilities.LoadMessageForm(xml_Load, frm_Load)
                                    oLoadForm = oApplication.SBO_Application.Forms.ActiveForm()
                                    oLoadForm.Items.Item("3").TextStyle = 4
                                    oLoadForm.Items.Item("4").TextStyle = 5
                                    CType(oLoadForm.Items.Item("3").Specific, SAPbouiCOM.StaticText).Caption = "PLEASE WAIT..."
                                    CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Validating Data..."
                                    Try

                                        Dim strCardCode As String = oApplication.Utilities.getRecordSetValueString("Select U_CardCode From [@Z_OCPR] Where DocEntry = '" & strDocEntry & "'", "U_CardCode")

                                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Updating Open Sales Order...based on Customer Profile exclude/remove/Suspend"
                                        oApplication.Utilities.CloseOrderQuantityRemoveSuspendDates_P(strDocEntry) 'Close Open Order Based On based on Customer Profile exclude/remove/Suspend Dates.

                                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Updating Open Delivery ...based on Customer Profile exclude/remove/Suspend"
                                        oApplication.Utilities.CancelDeliveryQuantityRemoveSuspendDatesAndExclude(strDocEntry) 'Close Open Delivery Based On Remove/Suspend/Exclude Dates.

                                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Updating Open Sales Order...based on Customer Profile exclude/remove/Suspend"
                                        oApplication.Utilities.CloseOrderQuantityRemoveSuspendDates_P(strDocEntry) 'Close Open Order Based On Remove/Suspend/Exclude Dates.


                                        'tricky bit contradictory....20151211
                                        '====================

                                        'Newly Added for Overlapping...20151209
                                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Updating Overlapping Program..."
                                        oApplication.Utilities.updateProgramDates_IfOverLapping(oForm, strDocEntry, strCardCode)

                                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Updating Remove Date For No of Days..."
                                        oApplication.Utilities.UpdateProgramNoofDaysBasedOnRemoveDate(oForm, strDocEntry)

                                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Updating Program To Date..."
                                        oApplication.Utilities.UpdateProgramToDate(oForm, strDocEntry)

                                        '====================

                                        'this should be executed after Overlapping process get executed.
                                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Updating Open Delivery ...based on Customer Profile Overlapping/Exclude"
                                        oApplication.Utilities.CancelDeliveryQuantityExcludeOnOverlapping(strDocEntry) 'Close Open Delivery Based On Overlapping/Exclude Dates.

                                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Updating Open Sales Order...based on Customer Profile Overlapping/Exclude "
                                        oApplication.Utilities.CloseOrderQuantityExcludeOnOverlapping(strDocEntry) 'Close Open Order Based On Overlapping/Exclude Dates.

                                        'Independent Process...based on Program From/ProgramTo the Status will be updated
                                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Updating Program On/Off..."
                                        oApplication.Utilities.UpdateProgramDateOnOffStatus(oForm, strDocEntry) 'Update Customer On/Off In Program

                                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Updating Open Order for Calories Changes ..."
                                        oApplication.Utilities.UpdateOrderQuantityBasedOnCalories(strDocEntry) 'Update Open Order Based On Calories if Differ.

                                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Updating Open Order for Address Changes ..."
                                        oApplication.Utilities.UpdateOpenOrderAddresses(strCardCode) 'Update Open Order Based On Addresses if Differ.

                                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Updating Program Services..."
                                        oApplication.Utilities.updateServiceRegistrationRows(oForm, strCardCode) 'Service Quantity & Applied Date of the Service

                                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Updating Program Registrations..."
                                        oApplication.Utilities.updateRegistrationRows(oForm, strCardCode) 'Row level calculation & Document Level Calculation.

                                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Updating Customer On/Off..."
                                        oApplication.Utilities.UpdateONOFFStatus(oForm, strDocEntry) 'Update Customer On/Off If Program Document Not Exists

                                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "COMPLETED..."

                                        oLoadForm.Close()
                                    Catch ex As Exception
                                        oLoadForm.Close()
                                    End Try

                                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Then
                                        loadProgramsAndVisits(oForm)
                                    End If
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

#End Region

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            If oForm.TypeEx = frm_Z_OCPR Then
                intSelectedMatrixrow = eventInfo.Row
                Dim oMenuItem As SAPbouiCOM.MenuItem
                oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data
                If (eventInfo.BeforeAction = True) Then
                    oMenuItem.SubMenus.Item(mnu_CANCEL).Enabled = False
                    oMenuItem.SubMenus.Item(mnu_CLOSE).Enabled = False
                    If eventInfo.ItemUID = "3" Then
                        'oMenuItem.SubMenus.Item(mnu_ADD_ROW).Enabled = True
                        'oMenuItem.SubMenus.Item(mnu_DELETE_ROW).Enabled = True
                    Else
                        'oMenuItem.SubMenus.Item(mnu_ADD_ROW).Enabled = False
                        'oMenuItem.SubMenus.Item(mnu_DELETE_ROW).Enabled = False
                    End If
                Else
                    oMenuItem.SubMenus.Item(mnu_CANCEL).Enabled = True
                    oMenuItem.SubMenus.Item(mnu_CLOSE).Enabled = True
                End If
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.PaneLevel = 1
            alldataSource(oForm)

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select IsNull(MAX(DocEntry),0) +1 From [@Z_OCPR]")
            If Not oRecordSet.EoF Then
                oApplication.Utilities.setEditText(oForm, "13", oRecordSet.Fields.Item(0).Value.ToString())
                oForm.Items.Item("11").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                oApplication.Utilities.setEditText(oForm, "11", "t")
                oApplication.SBO_Application.SendKeys("{TAB}")
            End If
            oForm.Items.Item("13").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Try
                CType(oForm.Items.Item("39").Specific, SAPbouiCOM.Matrix).Columns.Item("V_1").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                CType(oForm.Items.Item("40").Specific, SAPbouiCOM.Matrix).Columns.Item("V_1").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                CType(oForm.Items.Item("44").Specific, SAPbouiCOM.Matrix).Columns.Item("V_1").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                CType(oForm.Items.Item("44").Specific, SAPbouiCOM.Matrix).Columns.Item("V_2").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                CType(oForm.Items.Item("44").Specific, SAPbouiCOM.Matrix).Columns.Item("V_3").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                CType(oForm.Items.Item("44").Specific, SAPbouiCOM.Matrix).Columns.Item("V_4").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                CType(oForm.Items.Item("44").Specific, SAPbouiCOM.Matrix).Columns.Item("V_5").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                CType(oForm.Items.Item("44").Specific, SAPbouiCOM.Matrix).Columns.Item("V_6").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            Catch ex As Exception

            End Try
            oForm.Update()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim _retVal As Boolean = True
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            alldataSource(aForm)

            'Dim blnDislike As Boolean = False
            'For index As Integer = 0 To oDBDataSourceLines1.Size - 1
            '    If oDBDataSourceLines1.GetValue("U_DLikeItem", index) <> "" Then
            '        blnDislike = True
            '        Exit For
            '    End If
            'Next

            'Dim blnMedical As Boolean = False
            'For index As Integer = 0 To oDBDataSourceLines2.Size - 1
            '    If oDBDataSourceLines2.GetValue("U_MSCode", index) <> "" Then
            '        blnMedical = True
            '        Exit For
            '    End If
            'Next

            'If Not blnDislike Then
            '    oApplication.Utilities.Message("Add Dislike to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If


            'If Not blnMedical Then
            '    oApplication.Utilities.Message("Add Medical to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

            If oDBDataSource.GetValue("U_Sunday", oDBDataSource.Offset).Trim().ToString() = "Y" _
                And oDBDataSource.GetValue("U_Monday", oDBDataSource.Offset).Trim().ToString() = "Y" _
                And oDBDataSource.GetValue("U_Tuesday", oDBDataSource.Offset).Trim().ToString() = "Y" _
                And oDBDataSource.GetValue("U_Wednesday", oDBDataSource.Offset).Trim().ToString() = "Y" _
                And oDBDataSource.GetValue("U_Thursday", oDBDataSource.Offset).Trim().ToString() = "Y" _
                And oDBDataSource.GetValue("U_Friday", oDBDataSource.Offset).Trim().ToString() = "Y" _
                And oDBDataSource.GetValue("U_Saturday", oDBDataSource.Offset).Trim().ToString() = "Y" _
                Then
                oApplication.Utilities.Message("Cannot have all days Exclude...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If


            oMatrix = oForm.Items.Item("3").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim strCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
                If strCode.Trim() = "" And oMatrix.VisualRowCount > 1 Then
                    oApplication.Utilities.Message("Please Remove blank row in Dislike", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    If index <> intRow Then
                        Dim strCode1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
                        If strCode = strCode1 Then
                            oApplication.Utilities.Message("Dislike Code Already Exist...Code : " + strCode1.ToString(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next
            Next

            oMatrix = oForm.Items.Item("4").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim strCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
                If strCode.Trim() = "" And oMatrix.VisualRowCount > 1 Then
                    oApplication.Utilities.Message("Please Remove blank row in Medical", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    If index <> intRow Then
                        Dim strCode1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
                        If strCode = strCode1 Then
                            oApplication.Utilities.Message("Medical Code Already Exist...Code : " + strCode1.ToString(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next
            Next

            oMatrix = oForm.Items.Item("6").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim strCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
                'If Not oApplication.Utilities.validateDate(oForm, strCode, -1) Then
                '    oApplication.Utilities.Message("Exclude From Date Should be Greater than Or Equal Current Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Return False
                'End If
                If strCode.Trim() = "" And oMatrix.VisualRowCount > 1 Then
                    oApplication.Utilities.Message("Please Remove blank row in Exclude", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    If index <> intRow Then
                        Dim strCode1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
                        If strCode = strCode1 Then
                            oApplication.Utilities.Message("Exclude Date : " + strCode1.ToString(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next
            Next

            oMatrix = oForm.Items.Item("39").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim strFromDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
                Dim strToDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0_1", index)
                If strFromDate.Length < 0 And oMatrix.VisualRowCount > 1 Then
                    oApplication.Utilities.Message("Please Remove blank row in Address Tab", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    If index <> intRow Then
                        Dim strFromDate1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
                        Dim strToDate1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0_1", intRow)

                        If CInt(strFromDate1) >= CInt(strFromDate) And CInt(strFromDate1) <= CInt(strToDate) Then
                            For intCol As Integer = 0 To oMatrix.Columns.Count - 1

                                If oMatrix.Columns.Item(intCol).UniqueID = "V_4" Or oMatrix.Columns.Item(intCol).UniqueID = "V_5" Or oMatrix.Columns.Item(intCol).UniqueID = "V_6" _
                                    Or oMatrix.Columns.Item(intCol).UniqueID = "V_7" Or oMatrix.Columns.Item(intCol).UniqueID = "V_8" Or oMatrix.Columns.Item(intCol).UniqueID = "V_9" Then

                                    If CType(oMatrix.Columns.Item(intCol).Cells.Item(index).Specific, SAPbouiCOM.CheckBox).Checked = True _
                                        And CType(oMatrix.Columns.Item(intCol).Cells.Item(intRow).Specific, SAPbouiCOM.CheckBox).Checked = True Then
                                        oApplication.Utilities.Message("Ambiguity in Selection", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Return False
                                    End If
                                End If
                            Next
                        End If

                    End If
                Next
            Next

            oMatrix = oForm.Items.Item("40").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim strCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
                If strCode.Trim() = "" And oMatrix.VisualRowCount > 1 Then
                    oApplication.Utilities.Message("Please Remove blank row in Address(Day)", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    If index <> intRow Then
                        Dim strCode1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
                        If strCode = strCode1 Then
                            For intCol As Integer = 0 To oMatrix.Columns.Count - 1

                                If oMatrix.Columns.Item(intCol).UniqueID = "V_4" Or oMatrix.Columns.Item(intCol).UniqueID = "V_5" Or oMatrix.Columns.Item(intCol).UniqueID = "V_6" _
                                    Or oMatrix.Columns.Item(intCol).UniqueID = "V_7" Or oMatrix.Columns.Item(intCol).UniqueID = "V_8" Or oMatrix.Columns.Item(intCol).UniqueID = "V_9" Then

                                    If CType(oMatrix.Columns.Item(intCol).Cells.Item(index).Specific, SAPbouiCOM.CheckBox).Checked = True _
                                        And CType(oMatrix.Columns.Item(intCol).Cells.Item(intRow).Specific, SAPbouiCOM.CheckBox).Checked = True Then
                                        oApplication.Utilities.Message("Ambiguity in Selection in Address Tab", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Return False
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            Next

            oMatrix = oForm.Items.Item("44").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim strDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
                If strDate.Trim() = "" And oMatrix.VisualRowCount > 1 Then
                    oApplication.Utilities.Message("Please Remove blank row in Calories Ratio", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                Dim strCalories As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0_0", index)
                If strDate <> "" Then
                    If strCalories = "" Then
                        oApplication.Utilities.Message("Enter Calories in Calories Ratio for Date : " + strDate, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If

                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    If index <> intRow Then
                        Dim strDate1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
                        If strDate = strDate1 Then
                            oApplication.Utilities.Message("Program Date Already Exist In Calories Ratio Tab...Date " + strDate1, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next
            Next

            oMatrix = oForm.Items.Item("45").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim strFromDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
                Dim strToDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", index)

                'If Not oApplication.Utilities.validateDate(oForm, strFromDate, -1) Then
                '    oApplication.Utilities.Message("Remove From Date Should be Greater than Or Equal Current Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Return False
                'End If

                'If Not oApplication.Utilities.validateDate(oForm, strToDate, -1) Then
                '    oApplication.Utilities.Message("Remove To Date Should be Greater than Or Equal Current Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Return False
                'End If

                If strFromDate.Length < 0 And oMatrix.VisualRowCount > 1 Then
                    oApplication.Utilities.Message("Please Remove blank row in Remove Tab", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                ElseIf strFromDate.Length > 0 And strToDate.Length = 0 Then
                    oApplication.Utilities.Message("Enter Remove To Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    If index <> intRow Then
                        Dim strFromDate1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
                        Dim strToDate1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", intRow)

                        If strFromDate1.Length > 0 And strFromDate.Length > 0 And strToDate.Length > 0 Then
                            If CInt(strFromDate1) >= CInt(strFromDate) And CInt(strFromDate1) <= CInt(strToDate) Then
                                oApplication.Utilities.Message("Overlapping Dates In Remove Date Tab", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        End If
                    End If
                Next
                Dim strRemarks As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_2", index)

                If strFromDate.Length > 0 And strRemarks.Trim() = "" Then
                    oApplication.Utilities.Message("Enter Remarks In Remove Date Tab", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Next

            oMatrix = oForm.Items.Item("46").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim strFromDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
                Dim strToDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", index)

                'If Not oApplication.Utilities.validateDate(oForm, strFromDate, -1) Then
                '    oApplication.Utilities.Message("Suspend From Date Should be Greater than Or Equal Current Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Return False
                'End If

                'If Not oApplication.Utilities.validateDate(oForm, strToDate, -1) Then
                '    oApplication.Utilities.Message("Suspend To Date Should be Greater than Or Equal Current Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Return False
                'End If

                If strFromDate.Length < 0 And oMatrix.VisualRowCount > 1 Then
                    oApplication.Utilities.Message("Please Remove blank row in Suspend Tab", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    If index <> intRow Then
                        Dim strFromDate1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
                        Dim strToDate1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", intRow)

                        If strFromDate1.Length > 0 And strFromDate.Length > 0 And strToDate.Length > 0 Then
                            If CInt(strFromDate1) >= CInt(strFromDate) And CInt(strFromDate1) <= CInt(strToDate) Then
                                oApplication.Utilities.Message("Overlapping Dates In Suspend Date Tab", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        End If

                    End If
                Next
                Dim strRemarks As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_2", index)

                If strFromDate.Length > 0 And strRemarks.Trim() = "" Then
                    oApplication.Utilities.Message("Enter Remarks In Suspend Date Tab", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Next

            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
        Return _retVal
    End Function

    Private Function validateExcludeVsRemoveDates(ByVal oForm As SAPbouiCOM.Form, ByVal strValues As String, ByVal strCMatrixID As String) As Boolean
        Dim _retVal As Boolean = True
        Dim oCMatrix As SAPbouiCOM.Matrix
        oCMatrix = oForm.Items.Item(strCMatrixID).Specific

        Try
            If strCMatrixID = "45" Then
                For index As Integer = 1 To oCMatrix.RowCount
                    Dim strFDate As String = oApplication.Utilities.getMatrixValues(oCMatrix, "V_0", index)
                    Dim strTDate As String = oApplication.Utilities.getMatrixValues(oCMatrix, "V_0", index)
                    If strValues = strFDate Then
                        _retVal = False
                        Exit For
                    ElseIf strFDate = strValues Then
                        _retVal = False
                        Exit For
                    End If
                Next
            ElseIf strCMatrixID = "6" Then
                For index As Integer = 1 To oCMatrix.RowCount
                    Dim strFDate As String = oApplication.Utilities.getMatrixValues(oCMatrix, "V_0", index)
                    If strValues = strFDate Then
                        _retVal = False
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try

        Return _retVal
    End Function

    Private Function validate_Calories(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim _retVal As Boolean = True
            oMatrix = oForm.Items.Item("44").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim strDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
                Dim strCDate As String = System.DateTime.Now.ToString("yyyyMMdd")
                If strDate = strCDate Then
                    _retVal = False
                    Exit For
                End If
            Next
            Return _retVal
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Private Sub alldataSource(ByVal aForm As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPR")
            oDBDataSourceLines1 = oForm.DataSources.DBDataSources.Item("@Z_CPR1")
            oDBDataSourceLines2 = oForm.DataSources.DBDataSources.Item("@Z_CPR2")
            oDBDataSourceLines3 = oForm.DataSources.DBDataSources.Item("@Z_CPR3")
            oDBDataSourceLines4 = oForm.DataSources.DBDataSources.Item("@Z_CPR4")
            oDBDataSourceLines5 = oForm.DataSources.DBDataSources.Item("@Z_CPR5")
            oDBDataSourceLines6 = oForm.DataSources.DBDataSources.Item("@Z_CPR6")
            oDBDataSourceLines7 = oForm.DataSources.DBDataSources.Item("@Z_CPR7")
            oDBDataSourceLines8 = oForm.DataSources.DBDataSources.Item("@Z_CPR8")
            oDBDataSourceLines9 = oForm.DataSources.DBDataSources.Item("@Z_CPR9")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub addblankRow(ByVal oForm As SAPbouiCOM.Form, ByVal strItem As String)
        Try
            oMatrix = oForm.Items.Item(strItem).Specific
            oMatrix.LoadFromDataSource()
            oMatrix.AddRow(1, -1)
            oMatrix.FlushToDataSource()
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form, ByVal strItem As String)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item(strItem).Specific
            oMatrix.FlushToDataSource()
            Select Case aForm.PaneLevel
                Case "0", "1"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR1")
                Case "2"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR2")
                Case "3"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR4")
                Case "7"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR5")
                Case "8"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR6")
                Case "9"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR7")
                Case "10"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR8")
                Case "11"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR9")
            End Select

            If oMatrix.RowCount <= 0 Then
                oMatrix.AddRow()
            Else
                oMatrix.LoadFromDataSource()
                If oApplication.Utilities.getMatrixValues(oMatrix, "V_0", oMatrix.RowCount) <> "" Then
                    oMatrix.AddRow(1, oMatrix.RowCount + 1)
                    If strItem <> "44" Then
                        oMatrix.ClearRowData(oMatrix.RowCount)
                    End If
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
            AssignLineNo(aForm, strItem)

            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub AddRow_1(ByVal aForm As SAPbouiCOM.Form, ByVal strItem As String)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item(strItem).Specific
            oMatrix.FlushToDataSource()
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR7")

            If oMatrix.RowCount <= 0 Then
                oMatrix.AddRow()
            Else
                oMatrix.LoadFromDataSource()
                If oApplication.Utilities.getMatrixValues(oMatrix, "V_0", oMatrix.RowCount) <> "" Then
                    oMatrix.AddRow(1, oMatrix.RowCount + 1)
                    If strItem <> "44" Then
                        oMatrix.ClearRowData(oMatrix.RowCount)
                    End If
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
            Try
                oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Catch ex As Exception

            End Try
            AssignLineNo(aForm, strItem)

            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

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
                Case "4"
                    For count = 1 To oDBDataSourceLines2.Size
                        oDBDataSourceLines2.SetValue("LineId", count - 1, count)
                    Next
                Case "6"
                    For count = 1 To oDBDataSourceLines4.Size
                        oDBDataSourceLines4.SetValue("LineId", count - 1, count)
                    Next
                Case "39"
                    For count = 1 To oDBDataSourceLines5.Size
                        oDBDataSourceLines5.SetValue("LineId", count - 1, count)
                    Next
                Case "40"
                    For count = 1 To oDBDataSourceLines6.Size
                        oDBDataSourceLines6.SetValue("LineId", count - 1, count)
                    Next
                Case "44"
                    For count = 1 To oDBDataSourceLines7.Size
                        oDBDataSourceLines7.SetValue("LineId", count - 1, count)
                    Next
                Case "45"
                    For count = 1 To oDBDataSourceLines8.Size
                        oDBDataSourceLines8.SetValue("LineId", count - 1, count)
                    Next
                Case "46"
                    For count = 1 To oDBDataSourceLines9.Size
                        oDBDataSourceLines9.SetValue("LineId", count - 1, count)
                    Next
            End Select

            oMatrix.LoadFromDataSource()
            oMatrix.FlushToDataSource()
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form, ByVal strItem As String)
        Try
            oMatrix = aForm.Items.Item(strItem).Specific
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCPR")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR1")

            Select Case aForm.PaneLevel
                Case "0", "1"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR1")
                Case "2"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR2")
                    'Case "3"
                    '    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR3")
                Case "3"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR4")
                Case "7"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR5")
                Case "8"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR6")
                Case "9"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR7")
                Case "10"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR8")
                Case "11"
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_CPR9")
            End Select

            Me.RowtoDelete = intSelectedMatrixrow
            If Me.RowtoDelete - 1 >= 0 Then
                Me.RowtoDelete = intSelectedMatrixrow
                oMatrix.LoadFromDataSource()
                oMatrix.FlushToDataSource()
                oDBDataSourceLines.RemoveRecord(Me.RowtoDelete - 1)
                oMatrix.LoadFromDataSource()
                oMatrix.FlushToDataSource()
                For count = 0 To oDBDataSourceLines.Size - 1
                    oDBDataSourceLines.SetValue("LineId", count, count + 1)
                Next
                oMatrix.LoadFromDataSource()
                oMatrix.FlushToDataSource()
            End If

        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Function getMatrixItem(ByVal oForm As SAPbouiCOM.Form) As String
        Try
            Dim _retVal As String = String.Empty
            Select Case oForm.PaneLevel
                Case "0", "1"
                    _retVal = "3"
                Case "2"
                    _retVal = "4"
                Case "3"
                    _retVal = "6"
                Case "7"
                    _retVal = "39"
                Case "8"
                    _retVal = "40"
                Case "9"
                    _retVal = "44"
                Case "10"
                    _retVal = "45"
                Case "11"
                    _retVal = "46"
            End Select
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function changePane(ByVal oForm As SAPbouiCOM.Form, ByVal strItem As String) As String
        Try
            Dim _retVal As String = String.Empty
            Select Case strItem
                Case "24"
                    oForm.PaneLevel = 1
                Case "25"
                    oForm.PaneLevel = 2
                Case "26"
                    oForm.PaneLevel = 3
                Case "27"
                    oForm.PaneLevel = 4
                Case "28"
                    oForm.PaneLevel = 5
                Case "29"
                    oForm.PaneLevel = 6
                Case "36", "37"
                    oForm.PaneLevel = 7
                Case "38"
                    oForm.PaneLevel = 8
                Case "41"
                    oForm.PaneLevel = 9
                Case "42"
                    oForm.PaneLevel = 10
                Case "43"
                    oForm.PaneLevel = 11
            End Select
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub loadProgramsAndVisits(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim strQuery As String = String.Empty

            Dim strDocEntry As String = CType(oForm.Items.Item("13").Specific, SAPbouiCOM.EditText).Value
            oDTPrograms = oForm.DataSources.DataTables.Item("Programs")

            'Modifed for Phase II On 21-07-2015
            'strQuery = " Select T0.DocEntry, ISNULL(T0.U_InvRef,T2.U_InvRef) As U_InvRef,T0.U_TrnRef,T0.U_PrgCode,T1.ItemName As 'U_PrgName',"
            'strQuery += " T0.U_PFromDate,T0.U_PToDate,T0.U_NoOfDays,T0.U_RemDays,"
            'strQuery += " (Case WHEN ISNULL(T0.U_Transfer,'N') = 'N' THEN 'NO' WHEN ISNULL(T0.U_Transfer,'N') = 'Y' THEN 'YES' END) As U_Transfer, "
            'strQuery += " (Case WHEN ISNULL(T0.U_Cancel,'N') = 'N' THEN 'NO' WHEN ISNULL(T0.U_Cancel,'N') = 'Y' THEN 'YES' END) As U_Cancel   "
            'strQuery += " From [@Z_OCPM] T0 "
            'strQuery += " JOIN OITM T1 On T0.U_PrgCode = T1.ItemCode"
            'strQuery += " LEFT OUTER JOIN [@Z_CPM6] T2 On T0.DocEntry = T2.DocEntry "
            'strQuery += " Where T0.U_CardCode = '" + oForm.Items.Item("9").Specific.value + "'"

            strQuery = " Select Distinct T0.DocEntry, "
            strQuery += " (Case WHEN ISNULL(T0.U_Cancel,'N') = 'N' THEN 'NO' WHEN ISNULL(T0.U_Cancel,'N') = 'Y' THEN 'YES' END) As U_Cancel  "
            strQuery += " ,T0.U_PrgCode,T1.ItemName As 'U_PrgName', "
            'strQuery += " ISNULL(T0.U_InvRef,T2.U_InvRef) As U_InvRef,T0.U_TrnRef,  "
            strQuery += " T0.U_PFromDate As U_PFromDate,T0.U_PToDate As U_PToDate,  "
            strQuery += " ISNULL(T0.U_NoOfDays,0) As 'U_NoOfDays',ISNULL(T0.U_FreeDays,0) As 'U_FreeDays'"
            strQuery += " ,ISNULL(T0.U_RemDays,0) As 'U_RemDays',ISNULL(T0.U_OrdDays,0) As 'U_OrdDays', "
            strQuery += " ISNULL(T0.U_DelDays,0) As 'U_DelDays',ISNULL(T0.U_InvDays,0) As 'U_InvDays',  "
            strQuery += " (Case WHEN ISNULL(T0.U_DocStatus,'O') = 'O' THEN 'OPEN' WHEN ISNULL(T0.U_DocStatus,'O') = 'C' THEN 'CLOSED' WHEN ISNULL(T0.U_DocStatus,'O') = 'L' THEN 'CANCELED' END) As U_DocStatus, "
            strQuery += " (Case WHEN ISNULL(T0.U_Transfer,'N') = 'N' THEN 'NO' WHEN ISNULL(T0.U_Transfer,'N') = 'Y' THEN 'YES' END) As U_Transfer "
            'strQuery += " ,T3.U_PrgDate, "
            'strQuery += " (Case WHEN ISNULL(T3.U_AppStatus,'I') = 'I' THEN 'INCLUDE' WHEN ISNULL(T3.U_AppStatus,'I') = 'E' THEN 'EXCLUDE' END) As U_AppStatus, "
            'strQuery += " (Case WHEN ISNULL(T3.U_ONOFFSTA,'O') = 'O' THEN 'ON' WHEN ISNULL(T3.U_ONOFFSTA,'O') = 'F' THEN 'OFF' END) As U_ONOFFSTA  "
            strQuery += " From [@Z_OCPM] T0  JOIN OITM T1 On T0.U_PrgCode = T1.ItemCode LEFT OUTER JOIN [@Z_CPM6] T2 On T0.DocEntry = T2.DocEntry  "
            'strQuery += " JOIN [@Z_CPM1] T3 On T3.DocEntry = T0.DocEntry And T3.U_PrgDate Is Not Null "
            'strQuery += " And T3.U_PrgDate Between T0.U_PFromDate And T0.U_PToDate "
            strQuery += " Where T0.U_CardCode = '" + oForm.Items.Item("9").Specific.value + "'"
            strQuery += " Order By T0.DocEntry Desc "
            oDTPrograms.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("7").Specific
            oGrid.DataTable = oDTPrograms

            oGrid.Columns.Item("DocEntry").TitleObject.Caption = "Program Ref."
            oGrid.Columns.Item("DocEntry").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditTextColumn = oGrid.Columns.Item("DocEntry")
            oEditTextColumn.LinkedObjectType = "Z_OCPM"
            oGrid.Columns.Item("DocEntry").Editable = False

            'oGrid.Columns.Item("U_InvRef").TitleObject.Caption = "Invoice Ref."
            'oGrid.Columns.Item("U_InvRef").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            'oEditTextColumn = oGrid.Columns.Item("U_InvRef")
            'oEditTextColumn.LinkedObjectType = "13"
            'oGrid.Columns.Item("U_InvRef").Editable = False

            'oGrid.Columns.Item("U_TrnRef").TitleObject.Caption = "Transfer Ref."
            'oGrid.Columns.Item("U_TrnRef").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            'oEditTextColumn = oGrid.Columns.Item("U_TrnRef")
            'oEditTextColumn.LinkedObjectType = "13"
            'oGrid.Columns.Item("U_TrnRef").Editable = False

            oGrid.Columns.Item("U_PrgCode").TitleObject.Caption = "Program Code"
            oGrid.Columns.Item("U_PrgName").TitleObject.Caption = "Program Name"
            oGrid.Columns.Item("U_PFromDate").TitleObject.Caption = "Program From"
            oGrid.Columns.Item("U_PToDate").TitleObject.Caption = "Program To"
            oGrid.Columns.Item("U_NoOfDays").TitleObject.Caption = "No. Of Days(Paid)"
            oGrid.Columns.Item("U_FreeDays").TitleObject.Caption = "No. Of Days(Free)"
            oGrid.Columns.Item("U_RemDays").TitleObject.Caption = "Remaining No. Of Days"
            oGrid.Columns.Item("U_OrdDays").TitleObject.Caption = "No of Order(Days)"
            oGrid.Columns.Item("U_DelDays").TitleObject.Caption = "No of Delivery(Days)"
            oGrid.Columns.Item("U_InvDays").TitleObject.Caption = "No of Invoice(Days)"
            oGrid.Columns.Item("U_DocStatus").TitleObject.Caption = "Program Status"
            oGrid.Columns.Item("U_Transfer").TitleObject.Caption = "Transfer Status"
            oGrid.Columns.Item("U_Cancel").TitleObject.Caption = "Cancel Status"

            oApplication.Utilities.assignLineNo(oGrid, oForm)
            'oGrid.CollapseLevel = 15
            oForm.Items.Item("7").Enabled = False

            strQuery = " Select DocEntry,U_Program,U_FromDate,U_TillDate,U_SalesO,U_NoOfDays From [@Z_OPSL] "
            strQuery += " Where U_CardCode = '" + oForm.Items.Item("9").Specific.value + "'"
            strQuery += " Order By DocEntry Desc "
            oDTPreSales = oForm.DataSources.DataTables.Item("PreSales")
            oDTPreSales.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("8").Specific
            oGrid.DataTable = oDTPreSales

            oGrid.Columns.Item("DocEntry").TitleObject.Caption = "Pre Sales(Ref)"
            oGrid.Columns.Item("DocEntry").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditTextColumn = oGrid.Columns.Item("DocEntry")
            oEditTextColumn.LinkedObjectType = "13"
            oGrid.Columns.Item("DocEntry").Editable = False

            oGrid.Columns.Item("U_Program").TitleObject.Caption = "Program Name"
            oGrid.Columns.Item("U_FromDate").TitleObject.Caption = "PreSales From"
            oGrid.Columns.Item("U_TillDate").TitleObject.Caption = "PreSales To"
            oGrid.Columns.Item("U_NoOfDays").TitleObject.Caption = "No. Of Days"
            oGrid.Columns.Item("U_SalesO").TitleObject.Caption = "Sales Order"

            oApplication.Utilities.assignLineNo(oGrid, oForm)
            oForm.Items.Item("8").Enabled = False

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub EnableControls(ByVal oForm As SAPbouiCOM.Form, ByVal blnEnable As Boolean)
        Try
            oForm.Items.Item("13").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("9").Enabled = blnEnable
            oForm.Items.Item("10").Enabled = blnEnable
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub reDrawForm(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            oForm.Items.Item("34").Width = oForm.Width - 30
            oForm.Items.Item("34").Height = oForm.Items.Item("3").Height + 15
            oForm.Freeze(False)
        Catch ex As Exception
            'oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Private Sub loadComboColumn(ByVal oForm As SAPbouiCOM.Form)
        Try
            oMatrix = oForm.Items.Item("39").Specific
            If oMatrix.RowCount = 0 Then
                oMatrix.AddRow(1, -1)
            End If

            oCombo = oMatrix.Columns.Item("V_1").Cells().Item(1).Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select SlpCode,SlpName From OSLP "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("SlpCode").Value, oRecordSet.Fields.Item("SlpName").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            oCombo = oMatrix.Columns.Item("V_2").Cells().Item(1).Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select [Address] From CRD1 Where CardCode = '" + oForm.Items.Item("9").Specific.value + "' And AdresType = 'S'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("Address").Value, oRecordSet.Fields.Item("Address").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            oMatrix = oForm.Items.Item("40").Specific
            If oMatrix.RowCount = 0 Then
                oMatrix.AddRow(1, -1)
            End If
            oCombo = oMatrix.Columns.Item("V_1").Cells().Item(1).Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select SlpCode,SlpName From OSLP "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("SlpCode").Value, oRecordSet.Fields.Item("SlpName").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            oCombo = oMatrix.Columns.Item("V_2").Cells().Item(1).Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select [Address] From CRD1 Where CardCode = '" + oForm.Items.Item("9").Specific.value + "' And AdresType = 'S'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("Address").Value, oRecordSet.Fields.Item("Address").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            oMatrix = oForm.Items.Item("44").Specific
            If oMatrix.RowCount = 0 Then
                oMatrix.AddRow(1, -1)
            End If
            oCombo = oMatrix.Columns.Item("V_1").Cells().Item(1).Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select U_Code,U_Ratio From [@Z_OCRT] Where U_FType = 'BF' And U_Active = 'Y' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("U_Code").Value, oRecordSet.Fields.Item("U_Ratio").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            oCombo = oMatrix.Columns.Item("V_2").Cells().Item(1).Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select U_Code,U_Ratio From [@Z_OCRT] Where U_FType = 'LN' And U_Active = 'Y' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("U_Code").Value, oRecordSet.Fields.Item("U_Ratio").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            oCombo = oMatrix.Columns.Item("V_3").Cells().Item(1).Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select U_Code,U_Ratio From [@Z_OCRT] Where U_FType = 'LS' And U_Active = 'Y' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("U_Code").Value, oRecordSet.Fields.Item("U_Ratio").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            oCombo = oMatrix.Columns.Item("V_4").Cells().Item(1).Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select U_Code,U_Ratio From [@Z_OCRT] Where U_FType = 'SK' And U_Active = 'Y' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("U_Code").Value, oRecordSet.Fields.Item("U_Ratio").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            oCombo = oMatrix.Columns.Item("V_5").Cells().Item(1).Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select U_Code,U_Ratio From [@Z_OCRT] Where U_FType = 'DI' And U_Active = 'Y' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("U_Code").Value, oRecordSet.Fields.Item("U_Ratio").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            oCombo = oMatrix.Columns.Item("V_6").Cells().Item(1).Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select U_Code,U_Ratio From [@Z_OCRT] Where U_FType = 'DS' And U_Active = 'Y' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("U_Code").Value, oRecordSet.Fields.Item("U_Ratio").Value)
                    oRecordSet.MoveNext()
                End While
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub fillBuilding(ByVal oForm As SAPbouiCOM.Form, ByVal strMatID As String, ByVal intRow As Integer)
        Try
            oMatrix = oForm.Items.Item(strMatID).Specific
            oMatrix.FlushToDataSource()
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select Building From CRD1 Where CardCode = '" + oForm.Items.Item("9").Specific.value + "' And AdresType = 'S' "
            If strMatID = "39" Then
                strQuery += " And Address = '" + oDBDataSourceLines5.GetValue("U_Address", intRow - 1).Trim() + "'"
            ElseIf strMatID = "40" Then
                strQuery += " And Address = '" + oDBDataSourceLines6.GetValue("U_Address", intRow - 1).Trim() + "'"
            End If
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                If strMatID = "39" Then
                    oDBDataSourceLines5.SetValue("U_Building", intRow - 1, oRecordSet.Fields.Item(0).Value.ToString())
                ElseIf strMatID = "40" Then
                    oDBDataSourceLines6.SetValue("U_Building", intRow - 1, oRecordSet.Fields.Item(0).Value.ToString())
                End If
            End If
            oMatrix.LoadFromDataSource()
            oMatrix.FlushToDataSource()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub fillStandard(ByVal oForm As SAPbouiCOM.Form, ByVal strCaloriesID As String)
        Try
            oMatrix = oForm.Items.Item("44").Specific
            oMatrix.FlushToDataSource()
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select  "
            strQuery += " (Select U_Code From [@Z_OCRT] Where U_Ratio = T0.U_BFactor And U_FType = 'BF') As U_BF,  "
            strQuery += " (Select U_Code From [@Z_OCRT] Where U_Ratio = T0.U_LFactor And U_FType = 'LN') As U_LN,  "
            strQuery += " (Select U_Code From [@Z_OCRT] Where U_Ratio = T0.U_LSFactor And U_FType = 'LS') As U_LS,    "
            strQuery += " (Select U_Code From [@Z_OCRT] Where U_Ratio = T0.U_SFactor And U_FType = 'SK') As U_SK,  "
            strQuery += " (Select U_Code From [@Z_OCRT] Where U_Ratio = T0.U_DFactor And U_FType = 'DI') As U_DI,  "
            strQuery += " (Select U_Code From [@Z_OCRT] Where U_Ratio = T0.U_DSFactor And U_FType = 'DS') As U_DS  "
            strQuery += " From [@Z_OCAJ] T0 "
            strQuery += " Where  U_Calories = '" + strCaloriesID + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                Dim intColumns As Int16 = oRecordSet.Fields.Count - 1
                oDBDataSourceLines7.SetValue("U_CPAdj", oMatrix.RowCount - 1, strCaloriesID)
                oDBDataSourceLines7.SetValue("U_PrgDate", oMatrix.RowCount - 1, System.DateTime.Now.ToString("yyyyMMdd"))
                While intColumns >= 0
                    oDBDataSourceLines7.SetValue(oRecordSet.Fields.Item(intColumns).Name, oMatrix.RowCount - 1, oRecordSet.Fields.Item(oRecordSet.Fields.Item(intColumns).Name).Value.ToString())
                    intColumns -= 1
                End While
            End If
            oMatrix.LoadFromDataSource()
            oMatrix.FlushToDataSource()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
