Imports SAPbobsCOM


Public Class clsInvoiceGeneration
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
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_OIVG, frm_Z_OIVG)
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

            oApplication.Utilities.setEditText(oForm, "42", System.DateTime.Now.ToString("yyyyMMdd"))

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
                Case mnu_Z_OIVG
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
            If pVal.FormTypeEx = frm_Z_OIVG Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
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
                                                If importInvoice_Consolidate(oForm) Then 'If importInvoice(oForm) Then
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
                                    LIVH(oForm)
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

    'Private Sub LIVH(ByVal aform As SAPbouiCOM.Form) 'Load Invoice Header
    '    Try
    '        oForm.Freeze(True)

    '        Dim strFromCardCode, strToCardCode, strFromCardName, strToCardName, strFromInvoiceDate, strToInvoiceDate As String
    '        Dim strType As String = String.Empty
    '        Dim strFType As String = String.Empty

    '        strFromCardCode = oForm.Items.Item("7").Specific.value
    '        strToCardCode = oForm.Items.Item("6").Specific.value

    '        strFromCardName = oForm.Items.Item("8").Specific.value
    '        strToCardName = oForm.Items.Item("9").Specific.value

    '        strFromInvoiceDate = oForm.Items.Item("26").Specific.value
    '        strToInvoiceDate = oForm.Items.Item("34").Specific.value

    '        oOrderGrid = oForm.Items.Item("16").Specific
    '        oOrderGrid.DataTable = oForm.DataSources.DataTables.Item("dtOrder")

    '        oSuccessGrid = oForm.Items.Item("19").Specific
    '        oSuccessGrid.DataTable = oForm.DataSources.DataTables.Item("dtSuccess")

    '        oFailureGrid = oForm.Items.Item("21").Specific
    '        oFailureGrid.DataTable = oForm.DataSources.DataTables.Item("dtFailure")

    '        'strqry = "Select Distinct Convert(VarChar(1),'Y') As 'Select',T0.DocEntry,T0.DocNum, T0.U_CardCode,T0.U_CardName, T0.U_PrgCode, "
    '        'strqry += " (T1.U_DelDays-IsNull(T1.U_InvDays,0)) As 'NoofDays'"
    '        'strqry += ",T1.U_Fdate,T1.U_Edate,T1.U_Price, T1.U_Discount,T1.U_PaidType"
    '        'strqry += " From [@Z_OCPM] T0 JOIN [@Z_CPM6] T1 ON T0.DocEntry = T1.DocEntry "
    '        'strqry += " Where (T1.U_DelDays-IsNull(T1.U_InvDays,0)) > 0 "

    '        'If (strFromInvoiceDate.Length > 0 And strToInvoiceDate.Length > 0) Then
    '        '    strqry += " And Convert(VarChar(8),T1.ShipDate,112) Between '" + strFromInvoiceDate + "' AND '" + strToInvoiceDate + "'"
    '        'End If
    '        'If ((strFromInvoiceDate.Length > 0) And (strToInvoiceDate.Length = 0)) Then
    '        '    strqry += " And Convert(VarChar(8),T1.ShipDate,112) = '" + strFromInvoiceDate + "'"
    '        'End If
    '        'If ((strFromInvoiceDate.Length = 0) And (strToInvoiceDate.Length > 0)) Then
    '        '    strqry += " And Convert(VarChar(8),T1.ShipDate,112) = '" + strToInvoiceDate + "'"
    '        'End If

    '        strqry = " Select Convert(VarChar(1),'Y') As 'Select',T1.U_DocCur,T1.U_DocRate,T1.DocEntry,T4.Item,T2.ItemName, "
    '        strqry += " (Case WHEN ISNULL(U_TrnRef,'') <> '' THEN T1.U_CardCode ELSE U_CardCode END) As U_CardCode "
    '        strqry += ",T1.U_CardName,ISNULL(T1.U_Discount,0) As Discount"
    '        'strqry += " ,Count(T4.U_DelDate) As 'DD', "
    '        strqry += "  ,(Case  When T4.[Type] = 'Program' Then Count(T4.U_DelDate)  Else T4.Qty End) As 'DD' "
    '        strqry += " ,T4.U_Price,T4.RowDiscount,T4.U_TaxCode, "
    '        strqry += " Min(U_DelDate) As 'MinDate',Max(U_DelDate) As 'MaxDate' "
    '        strqry += " ,T4.U_PaidType "
    '        strqry += " ,T4.Qty "
    '        strqry += " ,T4.[Type] "
    '        strqry += " ,Convert(VarChar(8),Min(U_DelDate),112) 'MD' "
    '        strqry += " ,Convert(VarChar(8),Max(U_DelDate),112) 'XD' "
    '        strqry += " ,T3.U_SequenceType "
    '        strqry += " From "
    '        strqry += " ( "
    '        strqry += " Select Distinct(U_DelDate),T0.U_ProgramID,T3.U_PrgCode As 'Item' "
    '        strqry += " ,T2.U_Price,T2.U_Discount As 'RowDiscount',T2.U_TaxCode, "
    '        strqry += " T0.U_PaidType "
    '        strqry += " ,T2.U_NoofDays As 'Qty' "
    '        strqry += " ,'Program' As 'Type' "
    '        strqry += " From [DLN1] T0 "
    '        strqry += " JOIN [ODLN] T1 On T0.DocEntry = T1.DocEntry "
    '        strqry += " And Convert(VarChar(8),T0.U_DelDate,112) Between '" + strFromInvoiceDate + "' AND '" + strToInvoiceDate + "'"
    '        strqry += " JOIN [@Z_CPM6] T2  "
    '        strqry += " On T0.U_ProgramID = T2.DocEntry "
    '        strqry += " JOIN [@Z_OCPM] T3 On T2.DocEntry = T3.DocEntry "
    '        strqry += " And T2.U_PaidType = T0.U_PaidType "
    '        strqry += " And Convert(VarChar(8),T0.U_DelDate,112) 	 "
    '        strqry += " Between Convert(VarChar(8),T2.U_Fdate,112) And Convert(VarChar(8),T2.U_Edate,112) "
    '        strqry += " And ((T0.LineStatus = 'O') And (ISNULL(T1.U_InvRef,'') = '')) "
    '        strqry += " And T2.U_IsIReq = 'Y' "
    '        If ((strFromCardCode.Length > 0) And (strToCardCode.Length <= 0)) Or ((strFromCardCode.Length <= 0) And (strToCardCode.Length > 0)) Then
    '            strqry += " And T1.CardCode = '" + IIf(strFromCardCode.Length > 0, strFromCardCode, strToCardCode) + "'"
    '        End If
    '        If (strFromCardCode.Length > 0 And strToCardCode.Length > 0) Then
    '            strqry += " And T1.CardCode Between '" + strFromCardCode + "' AND '" + strToCardCode + "'"
    '        End If
    '        strqry += " Group By T0.U_ProgramID,T3.U_PrgCode,T0.U_DelDate,T2.U_Price,T2.U_Discount,T2.U_TaxCode,T0.U_PaidType,T2.U_NoofDays,T1.CardCode "
    '        strqry += " Union All  "
    '        strqry += " Select Distinct T2.U_Date,T0.U_ProgramID,T2.U_ItemCode As 'Item' "
    '        strqry += " ,T2.U_Price,T2.U_Discount As 'RowDiscount',T2.U_TaxCode, "
    '        strqry += " 'P' As U_PaidType "
    '        strqry += " ,T2.U_Quantity As 'Qty' "
    '        strqry += " ,'Service' As 'Type' "
    '        strqry += " From [DLN1] T0 "
    '        strqry += " JOIN [ODLN] T1 On T0.DocEntry = T1.DocEntry "
    '        'strqry += " And Convert(VarChar(8),T0.U_DelDate,112) Between '" + strFromInvoiceDate + "' AND '" + strToInvoiceDate + "'"
    '        strqry += " JOIN [@Z_CPM7] T2  "
    '        strqry += " On T0.U_ProgramID = T2.DocEntry "
    '        strqry += " And ISNULL(T2.U_InvCreated,'N') = 'N' "
    '        strqry += " And Convert(VarChar(8),T2.U_Date,112) 	 "
    '        strqry += " Between '" + strFromInvoiceDate + "' AND '" + strToInvoiceDate + "'"
    '        If ((strFromCardCode.Length > 0) And (strToCardCode.Length <= 0)) Or ((strFromCardCode.Length <= 0) And (strToCardCode.Length > 0)) Then
    '            strqry += " And T1.CardCode = '" + IIf(strFromCardCode.Length > 0, strFromCardCode, strToCardCode) + "'"
    '        End If
    '        If (strFromCardCode.Length > 0 And strToCardCode.Length > 0) Then
    '            strqry += " And T1.CardCode Between '" + strFromCardCode + "' AND '" + strToCardCode + "'"
    '        End If
    '        strqry += " Group By T0.U_ProgramID,T2.U_ItemCode,T2.U_Date,T2.U_Price,T2.U_Discount,T2.U_TaxCode,T0.U_PaidType,T2.U_Quantity,T1.CardCode "
    '        strqry += " ) T4  "
    '        strqry += " JOIN [@Z_OCPM] T1 On T4.U_ProgramID = T1.DocEntry "
    '        strqry += " JOIN OITM T2 On T4.Item = T2.ItemCode "
    '        strqry += " JOIN OCRD T3 On T1.U_CardCode = T3.CardCode "
    '        strqry += " Group By T1.DocEntry,T1.U_DocCur,T1.U_DocRate,T4.Item,T2.ItemName,T1.U_CardCode,T1.U_TrnRef,T1.U_CardName,T1.U_Discount,"
    '        strqry += " T4.U_Price,T4.RowDiscount,T4.U_TaxCode,T4.U_PaidType,T4.[Type],T4.Qty,T3.U_SequenceType " ',T4.U_DelDate
    '        strqry += " Order By T1.U_CardName "
    '        oOrderGrid.DataTable.ExecuteQuery(strqry)
    '        gridFormat(oForm)
    '        fillHeader(oForm, "16")
    '        oForm.Freeze(False)
    '    Catch ex As Exception 
    '.Log.Trace_DIET_AddOn_Error(ex)
    '        oForm.Freeze(False)
    '    End Try

    'End Sub

    Private Sub LIVH(ByVal aform As SAPbouiCOM.Form) 'Load Invoice Header
        Try
            oForm.Freeze(True)

            Dim strFromCardCode, strToCardCode, strFromCardName, strToCardName, strFromInvoiceDate, strToInvoiceDate As String
            Dim strType As String = String.Empty
            Dim strFType As String = String.Empty

            strFromCardCode = oForm.Items.Item("7").Specific.value
            strToCardCode = oForm.Items.Item("6").Specific.value

            strFromCardName = oForm.Items.Item("8").Specific.value
            strToCardName = oForm.Items.Item("9").Specific.value

            strFromInvoiceDate = oForm.Items.Item("26").Specific.value
            strToInvoiceDate = oForm.Items.Item("34").Specific.value

            oOrderGrid = oForm.Items.Item("16").Specific
            oOrderGrid.DataTable = oForm.DataSources.DataTables.Item("dtOrder")

            oSuccessGrid = oForm.Items.Item("19").Specific
            oSuccessGrid.DataTable = oForm.DataSources.DataTables.Item("dtSuccess")

            oFailureGrid = oForm.Items.Item("21").Specific
            oFailureGrid.DataTable = oForm.DataSources.DataTables.Item("dtFailure")

            'strqry = "Select Distinct Convert(VarChar(1),'Y') As 'Select',T0.DocEntry,T0.DocNum, T0.U_CardCode,T0.U_CardName, T0.U_PrgCode, "
            'strqry += " (T1.U_DelDays-IsNull(T1.U_InvDays,0)) As 'NoofDays'"
            'strqry += ",T1.U_Fdate,T1.U_Edate,T1.U_Price, T1.U_Discount,T1.U_PaidType"
            'strqry += " From [@Z_OCPM] T0 JOIN [@Z_CPM6] T1 ON T0.DocEntry = T1.DocEntry "
            'strqry += " Where (T1.U_DelDays-IsNull(T1.U_InvDays,0)) > 0 "

            'If (strFromInvoiceDate.Length > 0 And strToInvoiceDate.Length > 0) Then
            '    strqry += " And Convert(VarChar(8),T1.ShipDate,112) Between '" + strFromInvoiceDate + "' AND '" + strToInvoiceDate + "'"
            'End If
            'If ((strFromInvoiceDate.Length > 0) And (strToInvoiceDate.Length = 0)) Then
            '    strqry += " And Convert(VarChar(8),T1.ShipDate,112) = '" + strFromInvoiceDate + "'"
            'End If
            'If ((strFromInvoiceDate.Length = 0) And (strToInvoiceDate.Length > 0)) Then
            '    strqry += " And Convert(VarChar(8),T1.ShipDate,112) = '" + strToInvoiceDate + "'"
            'End If

            strqry = " Select (T1.U_CardName + '-' + Convert(VarChar,T1.DocEntry)) As 'Key',Convert(VarChar(1),'Y') As 'Select',T1.U_DocCur,T1.U_DocRate,T1.DocEntry,T4.Item,T2.ItemName "
            strqry += " ,(Case WHEN ISNULL(U_TrnRef,'') <> '' THEN T1.U_CardCode ELSE U_CardCode END) As U_CardCode, "
            strqry += " T1.U_CardName,ISNULL(T1.U_Discount,0) As Discount"
            'strqry += " ,Count(T4.U_DelDate) As 'DD', "
            strqry += "  ,(Case  When T4.[Type] = 'Program' Then Count(T4.U_DelDate)  Else T4.Qty End) As 'DD' "
            strqry += " ,T4.U_Price,T4.RowDiscount,T4.U_TaxCode, "
            strqry += " Min(U_DelDate) As 'MinDate',Max(U_DelDate) As 'MaxDate' "
            strqry += " ,T4.U_PaidType "
            strqry += " ,T4.Qty "
            strqry += " ,T4.[Type] "
            strqry += " ,Convert(VarChar(8),Min(U_DelDate),112) 'MD' "
            strqry += " ,Convert(VarChar(8),Max(U_DelDate),112) 'XD' "
            strqry += " ,T3.U_SequenceType "
            strqry += " From "
            strqry += " ( "
            strqry += " Select Distinct(U_DelDate),T0.U_ProgramID,T3.U_PrgCode As 'Item' "
            strqry += " ,T2.U_Price,T2.U_Discount As 'RowDiscount',T2.U_TaxCode, "
            strqry += " T0.U_PaidType "
            strqry += " ,T2.U_NoofDays As 'Qty' "
            strqry += " ,'Program' As 'Type' "
            strqry += " From [DLN1] T0 "
            strqry += " JOIN [ODLN] T1 On T0.DocEntry = T1.DocEntry "
            strqry += " And Convert(VarChar(8),T0.U_DelDate,112) Between '" + strFromInvoiceDate + "' AND '" + strToInvoiceDate + "'"
            strqry += " JOIN [@Z_CPM6] T2  "
            strqry += " On T0.U_ProgramID = T2.DocEntry "
            strqry += " JOIN [@Z_OCPM] T3 On T2.DocEntry = T3.DocEntry "
            strqry += " And T2.U_PaidType = T0.U_PaidType "
            strqry += " And Convert(VarChar(8),T0.U_DelDate,112) 	 "
            strqry += " Between Convert(VarChar(8),T2.U_Fdate,112) And Convert(VarChar(8),T2.U_Edate,112) "
            strqry += " And ((T0.LineStatus = 'O') And (ISNULL(T1.U_InvRef,'') = '')) "
            strqry += " And T2.U_IsIReq = 'Y' "
            If ((strFromCardCode.Length > 0) And (strToCardCode.Length <= 0)) Or ((strFromCardCode.Length <= 0) And (strToCardCode.Length > 0)) Then
                strqry += " And T1.CardCode = '" + IIf(strFromCardCode.Length > 0, strFromCardCode, strToCardCode) + "'"
            End If
            If (strFromCardCode.Length > 0 And strToCardCode.Length > 0) Then
                strqry += " And T1.CardCode Between '" + strFromCardCode + "' AND '" + strToCardCode + "'"
            End If
            strqry += " Group By T0.U_ProgramID,T3.U_PrgCode,T0.U_DelDate,T2.U_Price,T2.U_Discount,T2.U_TaxCode,T0.U_PaidType,T2.U_NoofDays,T1.CardCode "
            strqry += " Union All  "
            strqry += " Select Distinct T2.U_Date,T0.U_ProgramID,T2.U_ItemCode As 'Item' "
            strqry += " ,T2.U_Price,T2.U_Discount As 'RowDiscount',T2.U_TaxCode, "
            strqry += " 'P' As U_PaidType "
            strqry += " ,T2.U_Quantity As 'Qty' "
            strqry += " ,'Service' As 'Type' "
            strqry += " From [DLN1] T0 "
            strqry += " JOIN [ODLN] T1 On T0.DocEntry = T1.DocEntry "
            'strqry += " And Convert(VarChar(8),T0.U_DelDate,112) Between '" + strFromInvoiceDate + "' AND '" + strToInvoiceDate + "'"
            strqry += " JOIN [@Z_CPM7] T2  "
            strqry += " On T0.U_ProgramID = T2.DocEntry "
            strqry += " And ISNULL(T2.U_InvCreated,'N') = 'N' "
            strqry += " And Convert(VarChar(8),T2.U_Date,112) 	 "
            strqry += " Between '" + strFromInvoiceDate + "' AND '" + strToInvoiceDate + "'"
            If ((strFromCardCode.Length > 0) And (strToCardCode.Length <= 0)) Or ((strFromCardCode.Length <= 0) And (strToCardCode.Length > 0)) Then
                strqry += " And T1.CardCode = '" + IIf(strFromCardCode.Length > 0, strFromCardCode, strToCardCode) + "'"
            End If
            If (strFromCardCode.Length > 0 And strToCardCode.Length > 0) Then
                strqry += " And T1.CardCode Between '" + strFromCardCode + "' AND '" + strToCardCode + "'"
            End If
            strqry += " Group By T0.U_ProgramID,T2.U_ItemCode,T2.U_Date,T2.U_Price,T2.U_Discount,T2.U_TaxCode,T0.U_PaidType,T2.U_Quantity,T1.CardCode "
            strqry += " ) T4  "
            strqry += " JOIN [@Z_OCPM] T1 On T4.U_ProgramID = T1.DocEntry "
            strqry += " JOIN OITM T2 On T4.Item = T2.ItemCode "
            strqry += " JOIN OCRD T3 On T1.U_CardCode = T3.CardCode "
            strqry += " Group By T1.DocEntry,T1.U_DocCur,T1.U_DocRate,T4.Item,T2.ItemName,T1.U_CardCode,T1.U_TrnRef,T1.U_CardName,T1.U_Discount,"
            strqry += " T4.U_Price,T4.RowDiscount,T4.U_TaxCode,T4.U_PaidType,T4.[Type],T4.Qty,T3.U_SequenceType,T4.U_DelDate " '
            strqry += " Order By T1.U_CardName,T1.DocEntry,T4.Item "
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

            oOrderGrid.Columns.Item("Key").Editable = False

            oOrderGrid.Columns.Item("Select").TitleObject.Caption = "Select"
            oOrderGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oOrderGrid.Columns.Item("Select").Editable = True

            oOrderGrid.Columns.Item("DocEntry").TitleObject.Caption = "Program Code"
            oOrderGrid.Columns.Item("DocEntry").Visible = False

            oOrderGrid.Columns.Item("U_DocCur").TitleObject.Caption = "Doc Currency"
            oOrderGrid.Columns.Item("U_DocCur").Editable = False

            oOrderGrid.Columns.Item("U_DocRate").TitleObject.Caption = "Doc Rate"
            oOrderGrid.Columns.Item("U_DocRate").Editable = False

            oOrderGrid.Columns.Item("U_CardCode").TitleObject.Caption = "Card Code"
            oOrderGrid.Columns.Item("U_CardCode").Visible = False
            oOrderGrid.Columns.Item("U_CardCode").Editable = False

            oOrderGrid.Columns.Item("U_CardName").TitleObject.Caption = "Card Name"
            oOrderGrid.Columns.Item("U_CardName").Editable = False

            oOrderGrid.Columns.Item("Item").TitleObject.Caption = "Program Code"
            oOrderGrid.Columns.Item("Item").Editable = False

            oOrderGrid.Columns.Item("ItemName").TitleObject.Caption = "Program/Service Name"
            oOrderGrid.Columns.Item("ItemName").Visible = False

            oOrderGrid.Columns.Item("DD").TitleObject.Caption = "Invoice(Days)"
            oOrderGrid.Columns.Item("DD").Editable = False

            oOrderGrid.Columns.Item("Discount").TitleObject.Caption = "Document Discount"
            oOrderGrid.Columns.Item("Discount").Editable = False

            oOrderGrid.Columns.Item("U_Price").TitleObject.Caption = "Price"
            oOrderGrid.Columns.Item("U_Price").Editable = False

            oOrderGrid.Columns.Item("RowDiscount").TitleObject.Caption = "Discount"
            oOrderGrid.Columns.Item("RowDiscount").Editable = False

            oOrderGrid.Columns.Item("U_TaxCode").TitleObject.Caption = "Tax"
            oOrderGrid.Columns.Item("U_TaxCode").Editable = False

            oOrderGrid.Columns.Item("MinDate").TitleObject.Caption = "Program From"
            oOrderGrid.Columns.Item("MinDate").Editable = False

            oOrderGrid.Columns.Item("MaxDate").TitleObject.Caption = "Program To"
            oOrderGrid.Columns.Item("MaxDate").Editable = False

            oOrderGrid.Columns.Item("U_PaidType").TitleObject.Caption = "Paid Type"
            oOrderGrid.Columns.Item("U_PaidType").Editable = False

            'oOrderGrid.Columns.Item("Qty").TitleObject.Caption = "Paid Type"
            oOrderGrid.Columns.Item("Qty").Visible = False

            'oOrderGrid.Columns.Item("Type").TitleObject.Caption = "Paid Type"
            oOrderGrid.Columns.Item("Type").Visible = False
            oOrderGrid.Columns.Item("MD").Visible = False
            oOrderGrid.Columns.Item("XD").Visible = False
            oOrderGrid.Columns.Item("U_SequenceType").Visible = False
            oOrderGrid.CollapseLevel = 1

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
            Dim intRowNo As Integer = 0
            For index As Integer = 0 To oGrid.Rows.Count - 1
                If oGrid.GetDataTableRowIndex(index) >= 0 Then
                    intRowNo += 1
                    oGrid.RowHeaders.SetText(index, (intRowNo).ToString())
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

    Private Sub gridRFormat(ByVal oform As SAPbouiCOM.Form)
        Try
            oSuccessGrid = oform.Items.Item("19").Specific
            oSuccessGrid.DataTable = oform.DataSources.DataTables.Item("dtSuccess")

            oSuccessGrid.Columns.Item("Customer Code").TitleObject.Caption = "Customer Code"
            oSuccessGrid.Columns.Item("Customer Code").Editable = False

            oSuccessGrid.Columns.Item("Customer Name").TitleObject.Caption = "Customer Name"
            oSuccessGrid.Columns.Item("Customer Name").Editable = False

            oSuccessGrid.Columns.Item("Invoice No.").TitleObject.Caption = "Invoice No."
            oSuccessGrid.Columns.Item("Invoice No.").Editable = False

            oSuccessGrid.Columns.Item("Invoice Ref.").TitleObject.Caption = "Invoice Ref."
            oSuccessGrid.Columns.Item("Invoice Ref.").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditTextColumn = oSuccessGrid.Columns.Item("Invoice Ref.")
            oEditTextColumn.LinkedObjectType = "13"
            oSuccessGrid.Columns.Item("Invoice Ref.").Editable = False

            oFailureGrid = oform.Items.Item("21").Specific
            oFailureGrid.DataTable = oform.DataSources.DataTables.Item("dtFailure")

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

    Private Function Validation(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim strFromDate, strToDate, strPDate As String

            strFromDate = oForm.Items.Item("26").Specific.value
            strToDate = oForm.Items.Item("34").Specific.value
            strPDate = oForm.Items.Item("42").Specific.value

            If strFromDate = "" Then
                oApplication.Utilities.Message("Select Program From Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strToDate = "" Then
                oApplication.Utilities.Message("Select Program To Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strPDate = "" Then
                oApplication.Utilities.Message("Select Posting Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Dim strFrDt As String = oForm.Items.Item("26").Specific.value
            Dim strToDt As String = oForm.Items.Item("34").Specific.value
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

    Private Function importInvoice(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim _retVal As Boolean = False
            Dim strQuery As String
            Dim intDocEntry As Integer
            Dim iRow As Integer = 0
            Dim strDocNum As String = String.Empty
            Dim strCustomerCode As String = String.Empty
            Dim strCustomer As String = String.Empty
            Dim dblDiscountPercent As Double
            Dim strSeqType As String = String.Empty
            Dim strSeries As String = String.Empty

            Dim oInvoice As SAPbobsCOM.Documents
            Dim blnAdd As Boolean = False
            Dim intID_S As Integer = 0
            Dim intID_F As Integer = 0
            Dim intStatus As Integer
            Dim strFromInvoiceDate, strToInvoiceDate As String
            Dim strProgram, strFProgramdate, strTProgramdate, strPaidType As String
            Dim dblQty, dblPrice, dblRDiscount As String
            Dim strDCurrency As String

            strFromInvoiceDate = oForm.Items.Item("26").Specific.value
            strToInvoiceDate = oForm.Items.Item("34").Specific.value

            oDTSuccess = oForm.DataSources.DataTables.Item("dtSuccess")
            oDTFailure = oForm.DataSources.DataTables.Item("dtFailure")

            If oDTSuccess.Columns.Count = 0 Then
                oDTSuccess.Columns.Add("Customer Code", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
                oDTSuccess.Columns.Add("Customer Name", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                oDTSuccess.Columns.Add("Invoice No.", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                oDTSuccess.Columns.Add("Invoice Ref.", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            End If

            If oDTFailure.Columns.Count = 0 Then
                oDTFailure.Columns.Add("Customer Code", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
                oDTFailure.Columns.Add("Customer Name", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                oDTFailure.Columns.Add("FailedReason", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)
            End If

            For intRow As Integer = 0 To oOrderGrid.Rows.Count - 1

                'If Not oOrderGrid.Rows.IsLeaf(intRow) Then
                '    Continue For
                'End If

                If oOrderGrid.DataTable.GetValue("Select", intRow).ToString() = "Y" Then
                    Try
                        iRow = 0

                        intDocEntry = CInt(oOrderGrid.DataTable.GetValue("DocEntry", intRow).ToString())

                        strCustomerCode = oOrderGrid.DataTable.GetValue("U_CardCode", intRow).ToString()
                        strCustomer = oOrderGrid.DataTable.GetValue("U_CardName", intRow).ToString()
                        dblDiscountPercent = oOrderGrid.DataTable.GetValue("Discount", intRow)
                        strDCurrency = oOrderGrid.DataTable.GetValue("U_DocCur", intRow)
                        strSeqType = oOrderGrid.DataTable.GetValue("U_SequenceType", intRow).ToString()

                        If strSeqType.Trim().Length = 0 Then
                            Throw New Exception("Customer Sequence Not Defined.")
                        End If

                        strProgram = oOrderGrid.DataTable.GetValue("Item", intRow).ToString()
                        Dim strItemType As String
                        If oOrderGrid.DataTable.GetValue("Type", intRow) = "Program" Then
                            dblQty = oOrderGrid.DataTable.GetValue("DD", intRow)
                            strItemType = "P"
                        Else
                            dblQty = oOrderGrid.DataTable.GetValue("Qty", intRow)
                            strItemType = "S"
                        End If
                        dblPrice = oOrderGrid.DataTable.GetValue("U_Price", intRow)
                        dblRDiscount = oOrderGrid.DataTable.GetValue("RowDiscount", intRow)

                        oInvoice = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                        strQuery = "Select Series From NNM1 Where ObjectCode = '13' And BeginStr = '" & strSeqType & "' And Locked = 'N'"
                        strSeries = oApplication.Utilities.getRecordSetValueString_Series(strQuery, "Series")
                        If strSeries.Length > 0 Then
                            oInvoice.Series = CInt(strSeries)
                        Else
                            Throw New Exception("Check the Series for the Customer.")
                        End If

                        oInvoice.CardCode = strCustomerCode
                        oInvoice.CardName = strCustomer
                        oInvoice.DocDate = System.DateTime.Now
                        oInvoice.TaxDate = System.DateTime.Now
                        oInvoice.DocDueDate = System.DateTime.Now
                        oInvoice.DocCurrency = strDCurrency

                        oInvoice.DiscountPercent = dblDiscountPercent
                        oInvoice.Comments = "Program Invoice..."

                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Executing Invoice for Customer : " + strCustomer

                        If iRow > 0 Then
                            oInvoice.Lines.Add()
                        End If

                        oInvoice.Lines.SetCurrentLine(iRow)
                        oInvoice.Lines.ItemCode = strProgram
                        oInvoice.Lines.Quantity = dblQty
                        oInvoice.Lines.UnitPrice = dblPrice
                        oInvoice.Lines.DiscountPercent = dblRDiscount
                        oInvoice.Lines.Currency = strDCurrency
                        oInvoice.Lines.UserFields.Fields.Item("U_ProgramID").Value = oOrderGrid.DataTable.GetValue("DocEntry", intRow).ToString()
                        oInvoice.Lines.UserFields.Fields.Item("U_Program").Value = strProgram
                        oInvoice.Lines.UserFields.Fields.Item("U_Fdate").Value = oOrderGrid.DataTable.GetValue("MinDate", intRow)
                        oInvoice.Lines.UserFields.Fields.Item("U_Edate").Value = oOrderGrid.DataTable.GetValue("MaxDate", intRow)
                        oInvoice.Lines.UserFields.Fields.Item("U_PaidType").Value = oOrderGrid.DataTable.GetValue("U_PaidType", intRow)
                        oInvoice.Lines.UserFields.Fields.Item("U_ItemType").Value = strItemType

                        blnAdd = True
                        iRow += 1

                        If blnAdd = True Then
                            oInvoice.UserFields.Fields.Item("U_IsIWizard").Value = "Y"

                            ''Newly Added for Scope On 31102015
                            'If oApplication.Company.InTransaction Then
                            '    oApplication.Company.EndTransaction(BoWfTransOpt.wf_RollBack)
                            'End If
                            'oApplication.Company.StartTransaction()
                            ''Newly Added for Scope On 31102015

                            intStatus = oInvoice.Add()
                        End If

                        If intStatus = 0 Then

                            Dim intInvoiceDE As String = oApplication.Company.GetNewObjectKey()
                            Dim intInvoiceDN As String = String.Empty

                            If oInvoice.GetByKey(intInvoiceDE) Then
                                intInvoiceDN = oInvoice.DocNum

                                oDTSuccess.Rows.Add(1)
                                oDTSuccess.SetValue("Customer Code", intID_S, strCustomerCode)
                                oDTSuccess.SetValue("Customer Name", intID_S, strCustomer)
                                oDTSuccess.SetValue("Invoice No.", intID_S, intInvoiceDN)
                                oDTSuccess.SetValue("Invoice Ref.", intID_S, intInvoiceDE)

                                intID_S += 1
                            End If

                            If strItemType = "P" Then
                                strqry = " Select Distinct Top " & CInt(dblQty).ToString() & "  T0.DocEntry From DLN1 T0 Inner Join ODLN T1 on T1.DocEntry=T0.DocEntry Where "
                                strqry += " Convert(VarChar(8),T0.U_DelDate,112) "
                                strqry += " Between '" & oOrderGrid.DataTable.GetValue("MD", intRow) & "' AND '" & oOrderGrid.DataTable.GetValue("XD", intRow) & "'"
                                strqry += " And T0.U_ProgramID = '" & oOrderGrid.DataTable.GetValue("DocEntry", intRow).ToString() & "'"
                                strqry += " And T0.TargetType = '-1' and isnull(T1.U_InvNo,'')='' "
                                oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery(strqry)
                                If Not oRecordSet.EoF Then
                                    While Not oRecordSet.EoF
                                        'oApplication.Utilities.CloseDeliveryDocument(oRecordSet.Fields.Item(0).Value)
                                        oApplication.Utilities.UpdateDeliveryDocument(oRecordSet.Fields.Item(0).Value, intInvoiceDE, intInvoiceDN)
                                        oRecordSet.MoveNext()
                                    End While
                                End If
                            Else
                                oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                strqry = " Update [@Z_CPM7] Set "
                                strqry += " U_InvCreated = 'Y' "
                                strqry += " ,U_InvNo = '" + oInvoice.DocNum.ToString() + "'"
                                strqry += " ,U_InvRef = '" + oInvoice.DocEntry.ToString + "'"
                                strqry += " Where "
                                strqry += " Convert(VarChar(8),U_Date,112) "
                                strqry += " Between '" & oOrderGrid.DataTable.GetValue("MD", intRow) & "' AND '" & oOrderGrid.DataTable.GetValue("MD", intRow) & "'"
                                strqry += " And DocEntry = '" & oOrderGrid.DataTable.GetValue("DocEntry", intRow).ToString() & "'"
                                strqry += " And U_ItemCode = '" & oOrderGrid.DataTable.GetValue("Item", intRow).ToString() & "'"
                                oRecordSet.DoQuery(strqry)
                            End If

                            ''Newly Added for Scope On 31102015
                            'oApplication.Company.EndTransaction(BoWfTransOpt.wf_Commit)
                            ''Newly Added for Scope On 31102015

                        Else

                            ''Newly Added for Scope On 31102015
                            'oApplication.Company.EndTransaction(BoWfTransOpt.wf_RollBack)
                            ''Newly Added for Scope On 31102015

                            oDTFailure.Rows.Add(1)
                            oDTFailure.SetValue("Customer Code", intID_F, strCustomerCode)
                            oDTFailure.SetValue("Customer Name", intID_F, strCustomer)
                            oDTFailure.SetValue("FailedReason", intID_F, oApplication.Company.GetLastErrorDescription().ToString())
                            intID_F += 1

                        End If
                    Catch ex As Exception
                        oApplication.Log.Trace_DIET_AddOn_Error(ex)
                        oDTFailure.Rows.Add(1)
                        oDTFailure.SetValue("Customer Code", intID_F, strCustomerCode)
                        oDTFailure.SetValue("Customer Name", intID_F, strCustomer)
                        oDTFailure.SetValue("FailedReason", intID_F, ex.Message.ToString())
                        intID_F += 1
                    End Try
                End If
            Next
            _retVal = True
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)

            ''Newly Added for Scope On 31102015
            'If oApplication.Company.InTransaction Then
            '    oApplication.Company.EndTransaction(BoWfTransOpt.wf_RollBack)
            'End If
            ''Newly Added for Scope On 31102015

            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Private Function importInvoice_Consolidate(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim _retVal As Boolean = False
            Dim strQuery As String
            Dim intDocEntry As Integer
            Dim iRow As Integer = -1
            Dim strDocNum As String = String.Empty
            Dim strCustomerCode As String = String.Empty
            Dim strCustomer As String = String.Empty
            Dim dblDiscountPercent As Double
            Dim strSeqType As String = String.Empty
            Dim strSeries As String = String.Empty

            Dim oInvoice As SAPbobsCOM.Documents
            Dim blnAdd As Boolean = False
            Dim blnNextCust As Boolean = False
            Dim intID_S As Integer = 0
            Dim intID_F As Integer = 0
            Dim intStatus As Integer
            Dim strFromInvoiceDate, strToInvoiceDate, strPostingDate As String
            Dim strProgram, strFProgramdate, strTProgramdate, strPaidType As String
            Dim dblQty, dblPrice, dblRDiscount As String
            Dim strDCurrency As String

            strFromInvoiceDate = oForm.Items.Item("26").Specific.value
            strToInvoiceDate = oForm.Items.Item("34").Specific.value
            strPostingDate = oForm.Items.Item("42").Specific.value

            oDTSuccess = oForm.DataSources.DataTables.Item("dtSuccess")
            oDTFailure = oForm.DataSources.DataTables.Item("dtFailure")

            If oDTSuccess.Columns.Count = 0 Then
                oDTSuccess.Columns.Add("Customer Code", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
                oDTSuccess.Columns.Add("Customer Name", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                oDTSuccess.Columns.Add("Invoice No.", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                oDTSuccess.Columns.Add("Invoice Ref.", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            End If

            If oDTFailure.Columns.Count = 0 Then
                oDTFailure.Columns.Add("Customer Code", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
                oDTFailure.Columns.Add("Customer Name", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                oDTFailure.Columns.Add("FailedReason", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)
            End If

            For intRow As Integer = 0 To oOrderGrid.Rows.Count - 1

                If Not oOrderGrid.Rows.IsLeaf(intRow) Then
                    iRow = -1
                    blnAdd = False
                    blnNextCust = False
                    Continue For
                End If

                Dim intVRow As Integer = oOrderGrid.GetDataTableRowIndex(intRow)

                If oOrderGrid.DataTable.GetValue("Select", intVRow).ToString() = "Y" Then
                    Try
                        iRow += 1

                        intDocEntry = CInt(oOrderGrid.DataTable.GetValue("DocEntry", intVRow).ToString())

                        strCustomerCode = oOrderGrid.DataTable.GetValue("U_CardCode", intVRow).ToString()
                        strCustomer = oOrderGrid.DataTable.GetValue("U_CardName", intVRow).ToString()
                        dblDiscountPercent = oOrderGrid.DataTable.GetValue("Discount", intVRow)
                        strDCurrency = oOrderGrid.DataTable.GetValue("U_DocCur", intVRow)
                        strSeqType = oOrderGrid.DataTable.GetValue("U_SequenceType", intVRow).ToString()

                        If strSeqType.Trim().Length = 0 Then
                            Throw New Exception("Customer Sequence Not Defined.")
                        End If

                        strProgram = oOrderGrid.DataTable.GetValue("Item", intVRow).ToString()
                        Dim strItemType As String
                        If oOrderGrid.DataTable.GetValue("Type", intVRow) = "Program" Then
                            dblQty = oOrderGrid.DataTable.GetValue("DD", intVRow)
                            strItemType = "P"
                        Else
                            dblQty = oOrderGrid.DataTable.GetValue("Qty", intVRow)
                            strItemType = "S"
                        End If
                        dblPrice = oOrderGrid.DataTable.GetValue("U_Price", intVRow)
                        dblRDiscount = oOrderGrid.DataTable.GetValue("RowDiscount", intVRow)


                        If Not blnAdd Then
                            CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Executing Invoice for Customer : " + strCustomer
                            oInvoice = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                            strQuery = "Select Series From NNM1 Where ObjectCode = '13' And BeginStr = '" & strSeqType & "' And Locked = 'N'"
                            strSeries = oApplication.Utilities.getRecordSetValueString_Series(strQuery, "Series")
                            If strSeries.Length > 0 Then
                                oInvoice.Series = CInt(strSeries)
                            Else
                                Throw New Exception("Check the Series for the Customer.")
                            End If
                            oInvoice.CardCode = strCustomerCode
                            oInvoice.CardName = strCustomer
                            Dim PDate As Date = CDate(strPostingDate.Substring(0, 4) + "-" + strPostingDate.Substring(4, 2) + "-" + strPostingDate.Substring(6, 2))
                            oInvoice.DocDate = PDate 'System.DateTime.Now
                            oInvoice.TaxDate = PDate 'System.DateTime.Now
                            oInvoice.DocDueDate = PDate ' System.DateTime.Now
                            oInvoice.DocCurrency = strDCurrency
                            oInvoice.DiscountPercent = dblDiscountPercent
                            oInvoice.Comments = "Program Invoice..."
                            oInvoice.UserFields.Fields.Item("U_IsIWizard").Value = "Y"
                        End If


                        If iRow > 0 Then
                            oInvoice.Lines.Add()
                        End If

                        oInvoice.Lines.SetCurrentLine(iRow)
                        oInvoice.Lines.ItemCode = strProgram
                        oInvoice.Lines.Quantity = dblQty
                        oInvoice.Lines.UnitPrice = dblPrice
                        oInvoice.Lines.DiscountPercent = dblRDiscount
                        oInvoice.Lines.Currency = strDCurrency
                        oInvoice.Lines.UserFields.Fields.Item("U_ProgramID").Value = oOrderGrid.DataTable.GetValue("DocEntry", intVRow).ToString()
                        oInvoice.Lines.UserFields.Fields.Item("U_Program").Value = strProgram
                        oInvoice.Lines.UserFields.Fields.Item("U_Fdate").Value = oOrderGrid.DataTable.GetValue("MinDate", intVRow)
                        oInvoice.Lines.UserFields.Fields.Item("U_Edate").Value = oOrderGrid.DataTable.GetValue("MaxDate", intVRow)
                        oInvoice.Lines.UserFields.Fields.Item("U_PaidType").Value = oOrderGrid.DataTable.GetValue("U_PaidType", intVRow)
                        oInvoice.Lines.UserFields.Fields.Item("U_ItemType").Value = strItemType

                        blnAdd = True

                        'iRow += 1
                        If intRow < oOrderGrid.Rows.Count - 1 Then
                            If Not oOrderGrid.Rows.IsLeaf(intRow + 1) Then
                                blnNextCust = True
                            End If
                        ElseIf intRow = oOrderGrid.Rows.Count - 1 Then
                            blnNextCust = True
                        End If

                        If blnAdd = True And blnNextCust Then


                            ''Newly Added for Scope On 31102015
                            'If oApplication.Company.InTransaction Then
                            '    oApplication.Company.EndTransaction(BoWfTransOpt.wf_RollBack)
                            'End If
                            'oApplication.Company.StartTransaction()
                            ''Newly Added for Scope On 31102015

                            intStatus = oInvoice.Add()

                            If intStatus = 0 Then

                                Dim intInvoiceDE As String = oApplication.Company.GetNewObjectKey()
                                Dim intInvoiceDN As String = String.Empty

                                If oInvoice.GetByKey(intInvoiceDE) Then
                                    intInvoiceDN = oInvoice.DocNum

                                    oDTSuccess.Rows.Add(1)
                                    oDTSuccess.SetValue("Customer Code", intID_S, strCustomerCode)
                                    oDTSuccess.SetValue("Customer Name", intID_S, strCustomer)
                                    oDTSuccess.SetValue("Invoice No.", intID_S, intInvoiceDN)
                                    oDTSuccess.SetValue("Invoice Ref.", intID_S, intInvoiceDE)

                                    intID_S += 1
                                End If

                                If strItemType = "P" Then
                                    'strqry = " Select Distinct Top " & CInt(dblQty).ToString() & "  T0.DocEntry From DLN1 T0 Inner Join ODLN T1 on T1.DocEntry=T0.DocEntry Where "
                                    'strqry += " Convert(VarChar(8),T0.U_DelDate,112) "
                                    'strqry += " Between '" & oOrderGrid.DataTable.GetValue("MD", intVRow) & "' AND '" & oOrderGrid.DataTable.GetValue("XD", intVRow) & "'"
                                    'strqry += " And T0.U_ProgramID = '" & oOrderGrid.DataTable.GetValue("DocEntry", intVRow).ToString() & "'"
                                    'strqry += " And T0.TargetType = '-1' and isnull(T1.U_InvNo,'')='' "
                                    'oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    'oRecordSet.DoQuery(strqry)
                                    'If Not oRecordSet.EoF Then
                                    '    While Not oRecordSet.EoF
                                    '        'oApplication.Utilities.CloseDeliveryDocument(oRecordSet.Fields.Item(0).Value)
                                    '        oApplication.Utilities.UpdateDeliveryDocument(oRecordSet.Fields.Item(0).Value, intInvoiceDE, intInvoiceDN)
                                    '        oRecordSet.MoveNext()
                                    '    End While
                                    'End If
                                Else
                                    'oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    'strqry = " Update [@Z_CPM7] Set "
                                    'strqry += " U_InvCreated = 'Y' "
                                    'strqry += " ,U_InvNo = '" + oInvoice.DocNum.ToString() + "'"
                                    'strqry += " ,U_InvRef = '" + oInvoice.DocEntry.ToString + "'"
                                    'strqry += " Where "
                                    'strqry += " Convert(VarChar(8),U_Date,112) "
                                    'strqry += " Between '" & oOrderGrid.DataTable.GetValue("MD", intVRow) & "' AND '" & oOrderGrid.DataTable.GetValue("MD", intVRow) & "'"
                                    'strqry += " And DocEntry = '" & oOrderGrid.DataTable.GetValue("DocEntry", intVRow).ToString() & "'"
                                    'strqry += " And U_ItemCode = '" & oOrderGrid.DataTable.GetValue("Item", intVRow).ToString() & "'"
                                    'oRecordSet.DoQuery(strqry)
                                End If

                                ''Newly Added for Scope On 31102015
                                'oApplication.Company.EndTransaction(BoWfTransOpt.wf_Commit)
                                ''Newly Added for Scope On 31102015

                            Else

                                ''Newly Added for Scope On 31102015
                                'oApplication.Company.EndTransaction(BoWfTransOpt.wf_RollBack)
                                ''Newly Added for Scope On 31102015

                                oDTFailure.Rows.Add(1)
                                oDTFailure.SetValue("Customer Code", intID_F, strCustomerCode)
                                oDTFailure.SetValue("Customer Name", intID_F, strCustomer)
                                oDTFailure.SetValue("FailedReason", intID_F, oApplication.Company.GetLastErrorDescription().ToString())
                                intID_F += 1

                            End If

                        End If

                    Catch ex As Exception
                        oApplication.Log.Trace_DIET_AddOn_Error(ex)

                        oDTFailure.Rows.Add(1)
                        oDTFailure.SetValue("Customer Code", intID_F, strCustomerCode)
                        oDTFailure.SetValue("Customer Name", intID_F, strCustomer)
                        oDTFailure.SetValue("FailedReason", intID_F, ex.Message.ToString())

                        intID_F += 1

                    End Try
                End If


            Next
            _retVal = True
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)

            ''Newly Added for Scope On 31102015
            'If oApplication.Company.InTransaction Then
            '    oApplication.Company.EndTransaction(BoWfTransOpt.wf_RollBack)
            'End If
            ''Newly Added for Scope On 31102015

            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

End Class

'Public Function AddInvoiceDocument(ByVal oForm As SAPbouiCOM.Form) As Boolean
'    Dim _retVal As Boolean = False
'    Dim oInvoice As SAPbobsCOM.Documents
'    Dim oRecordSet As SAPbobsCOM.Recordset
'    Dim oRecordSet_C As SAPbobsCOM.Recordset
'    Dim oRecordSet_U As SAPbobsCOM.Recordset
'    Dim strQuery As String = String.Empty
'    Dim intStatus As Integer
'    Try
'        oInvoice = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
'        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
'        oRecordSet_C = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
'        oRecordSet_U = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
'        Dim strDocEntry As String = CType(oForm.Items.Item("10").Specific, SAPbouiCOM.EditText).Value

'        strQuery = " Select Distinct T1.* From ( "
'        strQuery += " Select DISTINCT ISNULL(T0.U_SerRef,'') As 'U_Reference',T0.LineId "
'        strQuery += " From [@Z_CPM6] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry "
'        strQuery += " Where T1.DocEntry = '" + strDocEntry + "'"
'        strQuery += " And ISNULL(T0.U_InvRef,'') = ''  "
'        strQuery += " And ISNULL(T0.U_IsIReq,'N') = 'Y'  "
'        strQuery += " UNION ALL "
'        strQuery += " Select DISTINCT T0.U_Reference,T0.U_Reference As 'LineId'  "
'        strQuery += " From [@Z_OISI] T0 JOIN [@Z_ISI1] T1 On T0.DocEntry = T1.DocEntry "
'        strQuery += " JOIN [@Z_CPM6] T2 ON T0.U_Reference = T2.U_SerRef "
'        strQuery += " Where T2.DocEntry = '" + strDocEntry + "'"
'        strQuery += " And ISNULL(T2.U_InvRef,'') = '' "
'        strQuery += " And T1.U_ItemCode <> '' "
'        strQuery += "  ) T1 "
'        oRecordSet.DoQuery(strQuery)
'        If Not oRecordSet.EoF Then
'            While Not oRecordSet.EoF

'                Dim intCurrentLine As Integer = 0
'                oInvoice = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

'                oInvoice.CardCode = oForm.Items.Item("6").Specific.value
'                oInvoice.CardName = oForm.Items.Item("7").Specific.value
'                oInvoice.NumAtCard = oForm.Items.Item("9").Specific.value
'                oInvoice.DocDate = System.DateTime.Now
'                oInvoice.TaxDate = System.DateTime.Now
'                oInvoice.DocDueDate = System.DateTime.Now
'                oInvoice.DiscountPercent = CDbl(IIf(oForm.Items.Item("17").Specific.value = "", 0, oForm.Items.Item("17").Specific.value))
'                oInvoice.Comments = "Program Booking"

'                strQuery = " Select T1.U_PrgCode As 'U_ItemCode',T1.U_PrgName As 'U_ItemName',T0.U_NoofDays As 'U_Quantity' "
'                strQuery += ",T0.U_Price,T0.U_Discount,T0.U_LineTotal,T0.U_PaidType "
'                strQuery += ",T0.U_Fdate,T0.U_Edate,'P' As 'Type' "
'                strQuery += " From [@Z_CPM6] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry "
'                If oRecordSet.Fields.Item("U_Reference").Value.ToString() = "" Then
'                    strQuery += " Where T0.LineId = '" + oRecordSet.Fields.Item("LineId").Value.ToString() + "'"
'                Else
'                    strQuery += " Where T0.U_SerRef = '" + oRecordSet.Fields.Item("U_Reference").Value + "'"
'                End If
'                strQuery += " And T1.DocEntry = '" + strDocEntry + "'"
'                strQuery += " And ISNULL(T0.U_IsIReq,'N') = 'Y'  "
'                strQuery += " UNION ALL "
'                strQuery += " Select T1.U_ItemCode,T1.U_ItemName,T1.U_Quantity,T1.U_Price,T1.U_Discount,T1.U_LineTotal,T2.U_PaidType  "
'                strQuery += ",T2.U_Fdate,T2.U_Edate,'S' As 'Type' "
'                strQuery += " From [@Z_OISI] T0 JOIN [@Z_ISI1] T1 On T0.DocEntry = T1.DocEntry "
'                strQuery += " JOIN [@Z_CPM6] T2 ON T0.U_Reference = T2.U_SerRef "
'                strQuery += " Where T2.U_SerRef = '" + oRecordSet.Fields.Item("U_Reference").Value + "'"
'                strQuery += " And T1.U_ItemCode <> '' "
'                oRecordSet_C.DoQuery(strQuery)
'                If Not oRecordSet_C.EoF Then
'                    While Not oRecordSet_C.EoF
'                        oInvoice.Lines.SetCurrentLine(intCurrentLine)
'                        oInvoice.Lines.ItemCode = oRecordSet_C.Fields.Item("U_ItemCode").Value
'                        oInvoice.Lines.Quantity = oRecordSet_C.Fields.Item("U_Quantity").Value
'                        oInvoice.Lines.UnitPrice = oRecordSet_C.Fields.Item("U_Price").Value
'                        oInvoice.Lines.DiscountPercent = oRecordSet_C.Fields.Item("U_Discount").Value
'                        If oRecordSet_C.Fields.Item("Type").Value = "P" Then
'                            oInvoice.Lines.UserFields.Fields.Item("U_Fdate").Value = oRecordSet_C.Fields.Item("U_Fdate").Value
'                            oInvoice.Lines.UserFields.Fields.Item("U_Edate").Value = oRecordSet_C.Fields.Item("U_Edate").Value
'                            oInvoice.Lines.UserFields.Fields.Item("U_PaidType").Value = oRecordSet_C.Fields.Item("U_PaidType").Value
'                        End If
'                        oInvoice.Lines.Add()
'                        intCurrentLine += 1
'                        oRecordSet_C.MoveNext()
'                    End While
'                End If

'                intStatus = oInvoice.Add
'                If intStatus = 0 Then

'                    Dim strInvoice As String = oApplication.Company.GetNewObjectKey()
'                    oInvoice.GetByKey(strInvoice)
'                    _retVal = True
'                    strQuery = "Update [@Z_CPM6] Set U_InvNo = '" + oInvoice.DocNum.ToString() + "'"
'                    strQuery += " ,U_InvRef = '" + strInvoice + "'"
'                    strQuery += " ,U_InvCreated = 'Y' "

'                    If oRecordSet.Fields.Item("U_Reference").Value.ToString() = "" Then
'                        strQuery += " Where LineId = '" + oRecordSet.Fields.Item("LineId").Value.ToString() + "'"
'                    Else
'                        strQuery += " Where U_SerRef = '" + oRecordSet.Fields.Item("U_Reference").Value + "'"
'                    End If

'                    strQuery += " And ISNULL(U_InvRef,'') = '' "
'                    strQuery += " AND DocEntry = '" + strDocEntry + "'"
'                    oRecordSet_U.DoQuery(strQuery)

'                    strQuery = " Update T1 Set "
'                    strQuery += " T1.U_InvNo = '" + oInvoice.DocNum.ToString() + "'"
'                    strQuery += " ,T1.U_InvRef = '" + strInvoice + "'"
'                    strQuery += " ,U_InvCreated = 'Y' "
'                    strQuery += " From [@Z_OISI] T0 JOIN [@Z_ISI1] T1 On T0.DocEntry = T1.DocEntry "
'                    strQuery += " Where T0.U_Reference =  '" + oRecordSet.Fields.Item("U_Reference").Value + "'"
'                    strQuery += " And ISNULL(T1.U_ItemCode,'') <> '' "
'                    oRecordSet_U.DoQuery(strQuery)

'                Else
'                    oApplication.SBO_Application.MessageBox(oApplication.Company.GetLastErrorDescription(), 1, "Ok", "", "")

'                End If

'                oRecordSet.MoveNext()

'            End While

'        End If

'        Return _retVal
'    Catch ex As Exception 
'oApplication.Log.Trace_DIET_AddOn_Error(ex)
'        Throw ex 
''oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
'    End Try
'End Function