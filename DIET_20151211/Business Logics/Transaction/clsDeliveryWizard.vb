Imports SAPbobsCOM

Public Class clsDeliveryWizard

    Inherits clsBase

    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Dim oDBDataSource1 As SAPbouiCOM.DBDataSource
    Private oLoadForm As SAPbouiCOM.Form
    Private oOrderGrid As SAPbouiCOM.Grid
    Private oSuccessGrid As SAPbouiCOM.Grid
    Private oFailureGrid As SAPbouiCOM.Grid
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oDTSuccess As SAPbouiCOM.DataTable
    Private oDTFailure As SAPbouiCOM.DataTable
    Private oCombo As SAPbouiCOM.ComboBox
    Dim strqry As String
    Dim sQuery As String
    Private oRecordSet As SAPbobsCOM.Recordset


    Public Sub LoadForm()
        Try
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_ODWT, frm_Z_ODWT)
            oForm = oApplication.SBO_Application.Forms.Item(strUID)
            oForm.Freeze(True)
            oForm.PaneLevel = 1
            initialize(oForm)
            addChooseFromListConditions(oForm)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub initialize(ByRef oForm As SAPbouiCOM.Form)
        Try

            oForm.Items.Item("15").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
            oForm.Items.Item("_41").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
            oForm.Items.Item("_19").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
            oForm.Items.Item("_21").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD

            oForm.DataSources.DataTables.Add("dtOrder")
            oForm.DataSources.DataTables.Add("dtSuccess")
            oForm.DataSources.DataTables.Add("dtFailure")

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Z_ODWT
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

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Dim oDataTable As SAPbouiCOM.DataTable
            If pVal.FormTypeEx = frm_Z_ODWT Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                'If (pVal.ItemUID = "18" Or pVal.ItemUID = "3") And oForm.PaneLevel > 1 Then
                                '    If validation(oForm) = False Then
                                '        BubbleEvent = False
                                '        Exit Sub
                                '    Else
                                If pVal.ItemUID = "3" And oForm.PaneLevel = 2 Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If

                                If pVal.ItemUID = "18" Then
                                    If oApplication.SBO_Application.MessageBox("Do you want to Proceed?", , "Yes", "No") = 2 Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        If oForm.PaneLevel = 3 Then
                                            oLoadForm = Nothing
                                            oLoadForm = oApplication.Utilities.LoadMessageForm(xml_Load, frm_Load)
                                            oLoadForm = oApplication.SBO_Application.Forms.ActiveForm()
                                            oLoadForm.Items.Item("3").TextStyle = 4
                                            oLoadForm.Items.Item("4").TextStyle = 5
                                            CType(oLoadForm.Items.Item("3").Specific, SAPbouiCOM.StaticText).Caption = "PLEASE WAIT..."
                                            CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Executing..."
                                            Try
                                                If importDelivery(oForm) Then
                                                    oSuccessGrid = oForm.Items.Item("19").Specific
                                                    oSuccessGrid.DataTable = oForm.DataSources.DataTables.Item("dtSuccess")
                                                    oSuccessGrid = oForm.Items.Item("21").Specific
                                                    oSuccessGrid.DataTable = oForm.DataSources.DataTables.Item("dtFailure")
                                                    gridRFormat(oForm)
                                                    oForm.PaneLevel = 4
                                                Else
                                                    oSuccessGrid = oForm.Items.Item("19").Specific
                                                    oSuccessGrid.DataTable = oForm.DataSources.DataTables.Item("dtSuccess")
                                                    oSuccessGrid = oForm.Items.Item("21").Specific
                                                    oSuccessGrid.DataTable = oForm.DataSources.DataTables.Item("dtFailure")
                                                    gridRFormat(oForm)
                                                    oForm.PaneLevel = 4
                                                End If
                                                oLoadForm.Close()
                                            Catch ex As Exception
                                                oApplication.Log.Trace_DIET_AddOn_Error(ex)
                                                oLoadForm.Close()
                                                oSuccessGrid = oForm.Items.Item("19").Specific
                                                oSuccessGrid.DataTable = oForm.DataSources.DataTables.Item("dtSuccess")
                                                oSuccessGrid = oForm.Items.Item("21").Specific
                                                oSuccessGrid.DataTable = oForm.DataSources.DataTables.Item("dtFailure")
                                                gridRFormat(oForm)
                                                oForm.PaneLevel = 4
                                            End Try
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
                                If pVal.ItemUID = "3" And (oForm.PaneLevel = 1) Then
                                    oForm.PaneLevel = oForm.PaneLevel + 1
                                ElseIf pVal.ItemUID = "17" And (oForm.PaneLevel > 1) Then
                                    oForm.PaneLevel = oForm.PaneLevel - 1
                                ElseIf pVal.ItemUID = "3" And (oForm.PaneLevel = 2) Then
                                    LSOH(oForm)
                                    oForm.PaneLevel = oForm.PaneLevel + 1
                                ElseIf pVal.ItemUID = "20" And (oForm.PaneLevel = 4) Then
                                    oForm.Close()
                                    Exit Sub
                                ElseIf pVal.ItemUID = "38" Then
                                    selectAll(oForm)
                                ElseIf pVal.ItemUID = "39" Then
                                    clearAll(oForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                oDBDataSource = oForm.DataSources.DBDataSources.Item("OINV")

                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent

                                oCFLEvento = pVal
                                oDataTable = oCFLEvento.SelectedObjects

                                If IsNothing(oDataTable) Then
                                    Exit Sub
                                End If

                                If pVal.ItemUID = "8" Then
                                    Dim intAddRows As Integer = oDataTable.Rows.Count
                                    If intAddRows > 0 Then
                                        oDBDataSource.SetValue("CardName", 0, oDataTable.GetValue("CardCode", 0))
                                        oDBDataSource.SetValue("NumAtCard", 0, oDataTable.GetValue("CardName", 0))
                                    End If
                                ElseIf pVal.ItemUID = "9" Then
                                    Dim intAddRows As Integer = oDataTable.Rows.Count
                                    If intAddRows > 0 Then
                                        oDBDataSource.SetValue("CardCode", 0, oDataTable.GetValue("CardCode", 0))
                                        oDBDataSource.SetValue("JrnlMemo", 0, oDataTable.GetValue("CardName", 0))
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "8" Then
                                    If CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value = "" Then
                                        CType(oForm.Items.Item("7").Specific, SAPbouiCOM.EditText).Value = ""
                                    End If
                                ElseIf pVal.ItemUID = "9" Then
                                    If CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value = "" Then
                                        CType(oForm.Items.Item("6").Specific, SAPbouiCOM.EditText).Value = ""
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
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
        End Try
    End Sub
#End Region

    Private Sub LSOH(ByVal aform As SAPbouiCOM.Form) 'Load Sales Order Header
        Try
            oForm.Freeze(True)

            Dim strFromCardCode, strToCardCode, strFromCardName, strToCardName, strFromDeliveryDate, strToDeliveryDate As String
            Dim strType As String = String.Empty
            Dim strFType As String = String.Empty

            strFromCardCode = oForm.Items.Item("7").Specific.value
            strToCardCode = oForm.Items.Item("6").Specific.value

            strFromCardName = oForm.Items.Item("8").Specific.value
            strToCardName = oForm.Items.Item("9").Specific.value

            strFromDeliveryDate = oForm.Items.Item("26").Specific.value
            strToDeliveryDate = oForm.Items.Item("34").Specific.value

            oOrderGrid = oForm.Items.Item("16").Specific
            oOrderGrid.DataTable = oForm.DataSources.DataTables.Item("dtOrder")

            oSuccessGrid = oForm.Items.Item("19").Specific
            oSuccessGrid.DataTable = oForm.DataSources.DataTables.Item("dtSuccess")

            oFailureGrid = oForm.Items.Item("21").Specific
            oFailureGrid.DataTable = oForm.DataSources.DataTables.Item("dtFailure")

            strqry = " Select Distinct Convert(VarChar(1),'Y') As 'Select',T0.CardCode,T0.CardName, "
            strqry += " T0.DocEntry,T0.DocNum,T1.ShipDate,T1.U_DelDate,Convert(VarChar(8),T1.U_DelDate,112) As 'DelDate' "
            strqry += " From ORDR T0 JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry "
            strqry += " Where T1.LineStatus = 'O' "
            strqry += " And ISNULL(T0.U_PSNo,'') <> '' "

            If ((strFromCardCode.Length > 0) And (strToCardCode.Length <= 0)) Or ((strFromCardCode.Length <= 0) And (strToCardCode.Length > 0)) Then
                strqry += " And T0.CardCode = '" + IIf(strFromCardCode.Length > 0, strFromCardCode, strToCardCode) + "'"
            End If
            If (strFromCardCode.Length > 0 And strToCardCode.Length > 0) Then
                strqry += " And T0.CardCode Between '" + strFromCardCode + "' AND '" + strToCardCode + "'"
            End If
            If (strFromDeliveryDate.Length > 0 And strToDeliveryDate.Length > 0) Then
                strqry += " And Convert(VarChar(8),T1.ShipDate,112) Between '" + strFromDeliveryDate + "' AND '" + strToDeliveryDate + "'"
            End If
            If ((strFromDeliveryDate.Length > 0) And (strToDeliveryDate.Length = 0)) Then
                strqry += " And Convert(VarChar(8),T1.ShipDate,112) = '" + strFromDeliveryDate + "'"
            End If
            If ((strFromDeliveryDate.Length = 0) And (strToDeliveryDate.Length > 0)) Then
                strqry += " And Convert(VarChar(8),T1.ShipDate,112) = '" + strToDeliveryDate + "'"
            End If
            strqry += " Order By T0.CardCode,T1.U_DelDate "
            oOrderGrid.DataTable.ExecuteQuery(strqry)

            gridFormat(oForm)
            fillHeader(oForm, "16")
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm.Freeze(False)
        End Try

    End Sub

    Private Sub gridFormat(ByVal oform As SAPbouiCOM.Form)
        Try

            oOrderGrid = oform.Items.Item("16").Specific
            oOrderGrid.DataTable = oform.DataSources.DataTables.Item("dtOrder")

            oOrderGrid.Columns.Item("Select").TitleObject.Caption = "Select"
            oOrderGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oOrderGrid.Columns.Item("Select").Editable = True

            oOrderGrid.Columns.Item("DocEntry").TitleObject.Caption = "Sale Order Ref"
            oOrderGrid.Columns.Item("DocEntry").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditTextColumn = oOrderGrid.Columns.Item("DocEntry")
            oEditTextColumn.LinkedObjectType = "17"
            oOrderGrid.Columns.Item("DocEntry").Editable = False

            oOrderGrid.Columns.Item("DocNum").TitleObject.Caption = "Order No."
            oOrderGrid.Columns.Item("DocNum").Editable = False

            oOrderGrid.Columns.Item("CardCode").TitleObject.Caption = "Card Code"
            oOrderGrid.Columns.Item("CardCode").Editable = False

            oOrderGrid.Columns.Item("CardName").TitleObject.Caption = "Card Name"
            oOrderGrid.Columns.Item("CardName").Editable = False

            oOrderGrid.Columns.Item("CardCode").TitleObject.Caption = "Card Code"
            oOrderGrid.Columns.Item("CardCode").Editable = False

            'oOrderGrid.Columns.Item("LineNum").TitleObject.Caption = "Order Line"
            'oOrderGrid.Columns.Item("LineNum").Visible = False

            oOrderGrid.Columns.Item("ShipDate").TitleObject.Caption = "Delivery Date"
            oOrderGrid.Columns.Item("ShipDate").Editable = False

            oOrderGrid.Columns.Item("U_DelDate").TitleObject.Caption = "Program Date"
            oOrderGrid.Columns.Item("U_DelDate").Editable = False

            oOrderGrid.Columns.Item("DelDate").TitleObject.Caption = "Program Date"
            oOrderGrid.Columns.Item("DelDate").Visible = False

            'oOrderGrid.CollapseLevel = 1

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub fillHeader(ByVal aForm As SAPbouiCOM.Form, ByVal strGridID As String)
        Try
            Dim oGrid As SAPbouiCOM.Grid
            aForm.Freeze(True)
            oGrid = aForm.Items.Item(strGridID).Specific
            oGrid.RowHeaders.TitleObject.Caption = "#"
            For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(index, (index + 1).ToString())
            Next
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            aForm.Freeze(False)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    'Private Sub gridSFFormat(ByVal oform As SAPbouiCOM.Form)
    '    Try
    '        oSuccessGrid = oform.Items.Item("19").Specific
    '        oFailureGrid = oform.Items.Item("21").Specific

    '    Catch ex As Exception 
    'oApplication.Log.Trace_DIET_AddOn_Error(ex)
    '        Throw ex 
    ''oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
    '    End Try
    'End Sub

    Private Sub gridRFormat(ByVal oform As SAPbouiCOM.Form)
        Try
            oSuccessGrid = oform.Items.Item("19").Specific
            oSuccessGrid.DataTable = oform.DataSources.DataTables.Item("dtSuccess")

            oSuccessGrid.Columns.Item("SaleOrderNo").TitleObject.Caption = "SaleOrderNo"
            oSuccessGrid.Columns.Item("SaleOrderNo").Editable = False

            oSuccessGrid.Columns.Item("SaleOrderRef").TitleObject.Caption = "Sale Order Ref"
            oSuccessGrid.Columns.Item("SaleOrderRef").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditTextColumn = oSuccessGrid.Columns.Item("SaleOrderRef")
            oEditTextColumn.LinkedObjectType = "17"
            oSuccessGrid.Columns.Item("SaleOrderRef").Editable = False

            oSuccessGrid.Columns.Item("Customer Code").TitleObject.Caption = "Customer Code"
            oSuccessGrid.Columns.Item("Customer Code").Editable = False

            oSuccessGrid.Columns.Item("Customer Name").TitleObject.Caption = "Customer Name"
            oSuccessGrid.Columns.Item("Customer Name").Editable = False

            oSuccessGrid.Columns.Item("Delivery No.").TitleObject.Caption = "Delivery No."
            oSuccessGrid.Columns.Item("Delivery No.").Editable = False

            oSuccessGrid.Columns.Item("Delivery Ref.").TitleObject.Caption = "Delivery Ref."
            oSuccessGrid.Columns.Item("Delivery Ref.").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditTextColumn = oSuccessGrid.Columns.Item("Delivery Ref.")
            oEditTextColumn.LinkedObjectType = "15"
            oSuccessGrid.Columns.Item("Delivery Ref.").Editable = False

            oSuccessGrid.Columns.Item("DelDate").TitleObject.Caption = "Program Date."
            oSuccessGrid.Columns.Item("DelDate").Editable = False


            oFailureGrid = oform.Items.Item("21").Specific
            oFailureGrid.DataTable = oform.DataSources.DataTables.Item("dtFailure")

            oFailureGrid.Columns.Item("SaleOrderNo").TitleObject.Caption = "SaleOrderNo"
            oFailureGrid.Columns.Item("SaleOrderNo").Editable = False

            oFailureGrid.Columns.Item("SaleOrderRef").TitleObject.Caption = "Sale Order Ref"
            oFailureGrid.Columns.Item("SaleOrderRef").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditTextColumn = oFailureGrid.Columns.Item("SaleOrderRef")
            oEditTextColumn.LinkedObjectType = "17"
            oFailureGrid.Columns.Item("SaleOrderRef").Editable = False

            oFailureGrid.Columns.Item("DelDate").TitleObject.Caption = "Program Date."
            oFailureGrid.Columns.Item("DelDate").Editable = False

            oFailureGrid.Columns.Item("Customer Code").TitleObject.Caption = "Customer Code"
            oFailureGrid.Columns.Item("Customer Code").Editable = False

            oFailureGrid.Columns.Item("Customer Name").TitleObject.Caption = "Customer Name"
            oFailureGrid.Columns.Item("Customer Name").Editable = False

            oFailureGrid.Columns.Item("FailedReason").TitleObject.Caption = "Failed Reason"
            oFailureGrid.Columns.Item("FailedReason").Editable = False


        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub selectAll(ByVal oForm As SAPbouiCOM.Form)
        Try
            oOrderGrid = oForm.Items.Item("16").Specific
            oForm.Freeze(True)
            For index As Integer = 0 To oOrderGrid.DataTable.Rows.Count - 1
                oOrderGrid.DataTable.SetValue("Select", index, "Y")
            Next
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub clearAll(ByVal oForm As SAPbouiCOM.Form)
        Try
            oOrderGrid = oForm.Items.Item("16").Specific
            oForm.Freeze(True)
            For index As Integer = 0 To oOrderGrid.DataTable.Rows.Count - 1
                oOrderGrid.DataTable.SetValue("Select", index, "N")
            Next
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub reDrawForm(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)

            oForm.Items.Item("19").Top = oForm.Items.Item("_19").Top + oForm.Items.Item("_19").Height + 2
            oForm.Items.Item("19").Height = (oForm.Height - 120) / 2
            oForm.Items.Item("19").Width = oForm.Width - 30

            oForm.Items.Item("_21").Top = oForm.Items.Item("19").Top + oForm.Items.Item("19").Height + 5

            oForm.Items.Item("21").Top = oForm.Items.Item("_21").Top + oForm.Items.Item("_21").Height + 5
            oForm.Items.Item("21").Height = oForm.Items.Item("19").Height
            oForm.Items.Item("21").Width = oForm.Items.Item("19").Width

            oForm.Freeze(False)
        Catch ex As Exception
            'oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm.Freeze(False)
        End Try
    End Sub

    Private Function importDelivery(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Try

            Dim _retVal As Boolean = False
            Dim oRecordSet, oISBatchSerial, oBatch As SAPbobsCOM.Recordset
            Dim intDocEntry As Integer
            Dim iRow As Integer = 0
            Dim strDocNum As String = String.Empty
            Dim strCustomerCode As String = String.Empty
            Dim strCustomer As String = String.Empty

            Dim intBatchNo As Integer = 0
            Dim oHashTable As Hashtable

            Dim oDelivery As SAPbobsCOM.Documents
            Dim blnAdd As Boolean = False
            Dim intID_S As Integer = 0
            Dim intID_F As Integer = 0
            Dim intStatus As Integer
            Dim strFromDeliveryDate, strToDeliveryDate As String

            strFromDeliveryDate = oForm.Items.Item("26").Specific.value
            strToDeliveryDate = oForm.Items.Item("34").Specific.value

            oDTSuccess = oForm.DataSources.DataTables.Item("dtSuccess")
            oDTFailure = oForm.DataSources.DataTables.Item("dtFailure")

            If oDTSuccess.Columns.Count = 0 Then
                oDTSuccess.Columns.Add("SaleOrderNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                oDTSuccess.Columns.Add("SaleOrderRef", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                oDTSuccess.Columns.Add("DelDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 30)
                oDTSuccess.Columns.Add("Customer Code", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
                oDTSuccess.Columns.Add("Customer Name", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                oDTSuccess.Columns.Add("Delivery No.", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                oDTSuccess.Columns.Add("Delivery Ref.", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            End If

            If oDTFailure.Columns.Count = 0 Then
                oDTFailure.Columns.Add("SaleOrderNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                oDTFailure.Columns.Add("SaleOrderRef", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                oDTFailure.Columns.Add("DelDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 30)
                oDTFailure.Columns.Add("Customer Code", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
                oDTFailure.Columns.Add("Customer Name", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                oDTFailure.Columns.Add("FailedReason", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)
            End If

            For intRow As Integer = 0 To oOrderGrid.Rows.Count - 1

                If oOrderGrid.DataTable.GetValue("Select", intRow).ToString() = "Y" Then

                    intDocEntry = CInt(oOrderGrid.DataTable.GetValue("DocEntry", intRow).ToString())
                    strDocNum = oOrderGrid.DataTable.GetValue("DocNum", intRow).ToString()
                    strCustomerCode = oOrderGrid.DataTable.GetValue("CardCode", intRow).ToString()
                    strCustomer = oOrderGrid.DataTable.GetValue("CardName", intRow).ToString()

                    strFromDeliveryDate = oOrderGrid.DataTable.GetValue("DelDate", intRow).ToString()
                    strToDeliveryDate = oOrderGrid.DataTable.GetValue("DelDate", intRow).ToString()

                    oDelivery = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    oDelivery.CardCode = strCustomerCode
                    oDelivery.CardName = strCustomer

                    Dim delDate As Date = CDate(strFromDeliveryDate.Substring(0, 4) + "-" + strFromDeliveryDate.Substring(4, 2) + "-" + strFromDeliveryDate.Substring(6, 2))
                    oDelivery.DocDate = delDate
                    oDelivery.DocDueDate = delDate
                    oDelivery.TaxDate = delDate

                    iRow = 0
                    oHashTable = New Hashtable()

                    strqry = " Select T1.ItemCode, T1.Dscription, T1.Quantity, T1.ShipDate, T1.LineNum, T1.WhsCode,T1.U_PaidType "
                    strqry += " From ORDR T0 JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry "
                    strqry += " Where T1.DocEntry = " & intDocEntry
                    strqry += " And T1.LineStatus = 'O' "

                    If (strFromDeliveryDate.Length > 0 And strToDeliveryDate.Length > 0) Then
                        strqry += " And Convert(VarChar(8),T1.U_DelDate,112) Between '" + strFromDeliveryDate + "' AND '" + strToDeliveryDate + "'"
                    End If

                    'If ((strFromDeliveryDate.Length > 0) And (strToDeliveryDate.Length = 0)) Then
                    '    strqry += " And Convert(VarChar(8),T1.ShipDate,112) = '" + strFromDeliveryDate + "'"
                    'End If
                    'If ((strFromDeliveryDate.Length = 0) And (strToDeliveryDate.Length > 0)) Then
                    '    strqry += " And Convert(VarChar(8),T1.ShipDate,112) = '" + strToDeliveryDate + "'"
                    'End If

                    oRecordSet.DoQuery(strqry)

                    If Not oRecordSet.EoF Then
                        While Not oRecordSet.EoF

                            CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Executing Sale Order: " + strDocNum + " For Delivery Date : " + strFromDeliveryDate

                            oISBatchSerial = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            If iRow > 0 Then
                                oDelivery.Lines.Add()
                            End If

                            oDelivery.Lines.ItemCode = oRecordSet.Fields.Item("ItemCode").Value
                            oDelivery.Lines.ItemDescription = oRecordSet.Fields.Item("Dscription").Value
                            oDelivery.Lines.Quantity = oRecordSet.Fields.Item("Quantity").Value
                            oDelivery.Lines.ShipDate = oRecordSet.Fields.Item("ShipDate").Value
                            oDelivery.Lines.WarehouseCode = oRecordSet.Fields.Item("WhsCode").Value

                            oDelivery.Lines.BaseType = 17
                            oDelivery.Lines.BaseEntry = intDocEntry
                            oDelivery.Lines.BaseLine = oRecordSet.Fields.Item("LineNum").Value
                            oDelivery.Lines.UserFields.Fields.Item("U_PaidType").Value = oRecordSet.Fields.Item("U_PaidType").Value

                            sQuery = "Select ManSerNum,ManBtchNum From OITM Where ItemCode = '" + oRecordSet.Fields.Item("ItemCode").Value + "'"
                            oISBatchSerial.DoQuery(sQuery)
                            If Not oISBatchSerial.EoF Then
                                If oISBatchSerial.Fields.Item("ManSerNum").Value = "Y" Then


                                ElseIf oISBatchSerial.Fields.Item("ManBtchNum").Value = "Y" Then
                                    oBatch = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    Dim oBatchNumber As SAPbobsCOM.BatchNumbers = oDelivery.Lines.BatchNumbers
                                    Dim dblBatchAllocated_A As Double = 0


                                    Dim dblRequiredQty As Double = oRecordSet.Fields.Item("Quantity").Value
                                    Dim dblBatchAllocated_B As Double = oRecordSet.Fields.Item("Quantity").Value 'Batch To Be Allocated

                                    sQuery = " Select T0.DistNumber,T1.Quantity,T0.AbsEntry From OBTN T0 "
                                    sQuery += " JOIN OIBT T1 On T0.SysNumber = T1.SysNumber "
                                    sQuery += " And T0.ItemCode = T1.ItemCode  "
                                    sQuery += " Where T1.ItemCode = '" + oRecordSet.Fields.Item("ItemCode").Value + "'"
                                    sQuery += " And T1.WhsCode  = '" + oRecordSet.Fields.Item("WhsCode").Value + "'"
                                    sQuery += " And T1.Quantity > 0 "
                                    sQuery += " Order By T0.SysNumber "

                                    oBatch.DoQuery(sQuery)
                                    If Not oBatch.EoF Then

                                        'Reset Batch Rows
                                        intBatchNo = 0

                                        While Not oBatch.EoF

                                            If oBatch.Fields.Item("Quantity").Value >= dblBatchAllocated_B Then

                                                Dim allocatedQty As Double = 0

                                                Dim Key As ICollection
                                                Key = oHashTable.Keys
                                                For Each k As String In Key
                                                    If oBatch.Fields.Item("AbsEntry").Value = k Then
                                                        allocatedQty = CDbl(oHashTable(k))
                                                        Exit For
                                                    End If
                                                Next

                                                If allocatedQty > 0 Then

                                                    If oHashTable.ContainsKey(oBatch.Fields.Item("AbsEntry").Value.ToString()) Then
                                                        Dim dblBalAvailable As Double = CDbl(oBatch.Fields.Item("Quantity").Value) - allocatedQty


                                                        If dblBalAvailable > 0 Then
                                                            If oRecordSet.Fields.Item("Quantity").Value >= dblBalAvailable Then
                                                                oHashTable.Remove(oBatch.Fields.Item("AbsEntry").Value.ToString())
                                                                oHashTable.Add(oBatch.Fields.Item("AbsEntry").Value.ToString(), allocatedQty + dblBalAvailable)

                                                                oBatchNumber.SetCurrentLine(intBatchNo)
                                                                oBatchNumber.BatchNumber = oBatch.Fields.Item("DistNumber").Value.ToString()
                                                                oBatchNumber.Quantity = dblBalAvailable
                                                                oBatchNumber.Add()

                                                                dblBatchAllocated_B = IIf((dblBatchAllocated_B - dblBalAvailable) <= 0, 0, dblBatchAllocated_B - dblBalAvailable)
                                                                dblBatchAllocated_A = IIf((dblBatchAllocated_A + dblBalAvailable) >= dblRequiredQty, dblRequiredQty, dblBatchAllocated_A + dblBalAvailable)

                                                                intBatchNo += 1
                                                            ElseIf oRecordSet.Fields.Item("Quantity").Value < dblBalAvailable Then
                                                                oHashTable.Remove(oBatch.Fields.Item("AbsEntry").Value.ToString())
                                                                oHashTable.Add(oBatch.Fields.Item("AbsEntry").Value.ToString(), allocatedQty + oRecordSet.Fields.Item("Quantity").Value)

                                                                oBatchNumber.SetCurrentLine(intBatchNo)
                                                                oBatchNumber.BatchNumber = oBatch.Fields.Item("DistNumber").Value.ToString()
                                                                oBatchNumber.Quantity = oRecordSet.Fields.Item("Quantity").Value
                                                                oBatchNumber.Add()

                                                                dblBatchAllocated_B = IIf((dblBatchAllocated_B - oRecordSet.Fields.Item("Quantity").Value) <= 0, 0, dblBatchAllocated_B - oRecordSet.Fields.Item("Quantity").Value)
                                                                dblBatchAllocated_A = IIf((dblBatchAllocated_A + oRecordSet.Fields.Item("Quantity").Value) >= dblRequiredQty, dblRequiredQty, dblBatchAllocated_A + oRecordSet.Fields.Item("Quantity").Value)

                                                                intBatchNo += 1
                                                            End If
                                                            'Else
                                                            '    intBatchNo += 1
                                                            '    oBatch.MoveNext()
                                                            '    Continue While
                                                        End If

                                                    End If
                                                Else

                                                    If Not oHashTable.ContainsKey(oBatch.Fields.Item("AbsEntry").Value.ToString()) Then

                                                        oHashTable.Add(oBatch.Fields.Item("AbsEntry").Value.ToString(), dblBatchAllocated_B)

                                                        oBatchNumber.SetCurrentLine(intBatchNo)
                                                        oBatchNumber.BatchNumber = oBatch.Fields.Item("DistNumber").Value.ToString()
                                                        oBatchNumber.Quantity = dblBatchAllocated_B
                                                        oBatchNumber.Add()

                                                        dblBatchAllocated_B = IIf((dblBatchAllocated_B - oBatch.Fields.Item("Quantity").Value) <= 0, 0, dblBatchAllocated_B - oBatch.Fields.Item("Quantity").Value)
                                                        dblBatchAllocated_A = IIf((dblBatchAllocated_A + oBatch.Fields.Item("Quantity").Value) >= dblRequiredQty, dblRequiredQty, dblBatchAllocated_A + oBatch.Fields.Item("Quantity").Value)
                                                        intBatchNo += 1

                                                    End If

                                                End If

                                            ElseIf oBatch.Fields.Item("Quantity").Value < dblBatchAllocated_B Then

                                                Dim allocatedQty As Double = 0

                                                Dim Key As ICollection
                                                Key = oHashTable.Keys
                                                For Each k As String In Key
                                                    If oBatch.Fields.Item("AbsEntry").Value = k Then
                                                        allocatedQty = CDbl(oHashTable(k))
                                                        Exit For
                                                    End If
                                                Next

                                                If Not oHashTable.ContainsKey(oBatch.Fields.Item("AbsEntry").Value.ToString()) Then
                                                    oHashTable.Add(oBatch.Fields.Item("AbsEntry").Value.ToString(), oBatch.Fields.Item("Quantity").Value)

                                                    oBatchNumber.SetCurrentLine(intBatchNo)
                                                    oBatchNumber.BatchNumber = oBatch.Fields.Item("DistNumber").Value.ToString()
                                                    oBatchNumber.Quantity = oBatch.Fields.Item("Quantity").Value
                                                    oBatchNumber.Add()

                                                    dblBatchAllocated_B -= oBatch.Fields.Item("Quantity").Value
                                                    dblBatchAllocated_A += oBatch.Fields.Item("Quantity").Value

                                                    intBatchNo += 1

                                                End If

                                            End If

                                            If dblBatchAllocated_B = 0 Then
                                                Exit While
                                            End If

                                            oBatch.MoveNext()

                                        End While

                                    End If
                                End If
                            End If

                            blnAdd = True
                            iRow += 1
                            oRecordSet.MoveNext()
                        End While
                    End If

                    If blnAdd = True Then
                        oDelivery.UserFields.Fields.Item("U_IsDWizard").Value = "Y"
                        intStatus = oDelivery.Add()
                    End If

                    If intStatus = 0 Then

                        Dim intDelDE As String = oApplication.Company.GetNewObjectKey()
                        Dim intDelDN As String

                        If oDelivery.GetByKey(intDelDE) Then
                            intDelDN = oDelivery.DocNum

                            oDTSuccess.Rows.Add(1)
                            oDTSuccess.SetValue("SaleOrderNo", intID_S, strDocNum)
                            oDTSuccess.SetValue("SaleOrderRef", intID_S, intDocEntry)
                            oDTSuccess.SetValue("Customer Code", intID_S, strCustomerCode)
                            oDTSuccess.SetValue("Customer Name", intID_S, strCustomer)
                            oDTSuccess.SetValue("Delivery No.", intID_S, intDelDN)
                            oDTSuccess.SetValue("Delivery Ref.", intID_S, intDelDE)
                            oDTSuccess.SetValue("DelDate", intID_S, oOrderGrid.DataTable.GetValue("U_DelDate", intRow).ToString())

                            intID_S += 1

                        End If
                    Else

                        oDTFailure.Rows.Add(1)
                        oDTFailure.SetValue("SaleOrderNo", intID_F, strDocNum)
                        oDTFailure.SetValue("SaleOrderRef", intID_F, intDocEntry)
                        oDTFailure.SetValue("Customer Code", intID_F, strCustomerCode)
                        oDTFailure.SetValue("Customer Name", intID_F, strCustomer)
                        oDTFailure.SetValue("DelDate", intID_F, oOrderGrid.DataTable.GetValue("U_DelDate", intRow).ToString())
                        Dim strError As String = oApplication.Company.GetLastErrorDescription().ToString()
                        Dim strLineNo As String = String.Empty
                        If strError.Contains("Quantity falls") Then
                            Dim strsLine() As String = strError.Split(":")
                            If strsLine.Length > 1 Then
                                Dim strLine() As String = strsLine(1).Split("]")
                                If strLine.Length > 0 Then
                                    strLineNo = strLine(0).Trim()
                                End If
                            End If
                        End If
                        If strLineNo <> "" Then
                            strqry = "Select (ItemCode +'-'+ Dscription) As 'Product' From RDR1 Where DocEntry = '" & intDocEntry.ToString & "'"
                            strqry += " And LineNum = '" & CInt(strLineNo) - 1 & "'"
                            Dim strProduct As String = oApplication.Utilities.getRecordSetValueString(strqry, "Product")
                            strError &= " " & strProduct
                        End If
                        oDTFailure.SetValue("FailedReason", intID_F, strError)
                        intID_F += 1

                    End If
                End If
            Next

            _retVal = True
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Private Function Validation(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim strFromDate, strToDate As String

            strFromDate = oForm.Items.Item("26").Specific.value
            strToDate = oForm.Items.Item("34").Specific.value
            If strFromDate = "" Then
                oApplication.Utilities.Message("Select Delivery From Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strToDate = "" Then
                oApplication.Utilities.Message("Select Delivery To Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Dim strFrDt As String = oForm.Items.Item("26").Specific.value
            Dim strToDt As String = oForm.Items.Item("34").Specific.value
            Dim strCurrentDt As String = System.DateTime.Now.AddDays(0).ToString("yyyyMMdd")

            If strFrDt.Length > 0 And strToDt.Length > 0 Then

                If CInt(strFrDt) > CInt(strToDt) Then
                    oApplication.Utilities.Message("From Date Should be Lesser than or Equal To Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                'From Date should be Lesser then or equal to Current Date
                If CInt(strFrDt) > CInt(strCurrentDt) Then
                    oApplication.Utilities.Message("Delivery From Date Should be Lesser than or Equal Todate Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                'To Date should be Lesser then or equal to Current Date
                If CInt(strToDt) > CInt(strCurrentDt) Then
                    oApplication.Utilities.Message("Delivery To Date Should be Lesser than or Equal Todate Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)

        End Try
    End Sub

End Class
