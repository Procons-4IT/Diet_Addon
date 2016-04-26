Imports SAPbobsCOM

Public Class clsModifyOrder
    Inherits clsBase

    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Dim oDBDataSource_1 As SAPbouiCOM.DBDataSource
    Dim oOptionbutton As SAPbouiCOM.OptionBtn
    Private oOrderGrid As SAPbouiCOM.Grid
    Private oSuccessGrid As SAPbouiCOM.Grid
    Private oFailureGrid As SAPbouiCOM.Grid
    Private oSuccessGrid_R As SAPbouiCOM.Grid
    Private oFailureGrid_R As SAPbouiCOM.Grid
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oComoColumn As SAPbouiCOM.ComboBoxColumn
    Private oDTSuccess As SAPbouiCOM.DataTable
    Private oDTFailure As SAPbouiCOM.DataTable
    Private oCombo As SAPbouiCOM.ComboBox
    Dim strqry As String
    Private oRecordSet As SAPbobsCOM.Recordset
    Private oLoadForm As SAPbouiCOM.Form
    Dim strQuery As String

    Public Sub LoadForm()
        Try
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_OMOT, frm_Z_OMOT)
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
            oForm.Items.Item("_19").TextStyle = 5
            oForm.Items.Item("_21").TextStyle = 5
            oForm.Items.Item("_40").TextStyle = 5
            oForm.Items.Item("_41").TextStyle = 5
            oForm.DataSources.DataTables.Add("dtOrder")

            oForm.DataSources.DataTables.Add("dtSuccess_A")
            oForm.DataSources.DataTables.Add("dtFailure_A")
            oForm.DataSources.DataTables.Add("dtSuccess_C")
            oForm.DataSources.DataTables.Add("dtFailure_C")
            oForm.DataSources.DataTables.Add("dtSuccess_D")
            oForm.DataSources.DataTables.Add("dtFailure_D")
            oForm.DataSources.DataTables.Add("dtSuccess_I")
            oForm.DataSources.DataTables.Add("dtFailure_I")
            oForm.DataSources.DataTables.Add("dtSuccess_L")
            oForm.DataSources.DataTables.Add("dtFailure_L")

            oOptionbutton = oForm.Items.Item("12").Specific
            oOptionbutton.Selected = True
            loadCombo(oForm)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Z_OMOT
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
            If pVal.FormTypeEx = frm_Z_OMOT Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "18" Or pVal.ItemUID = "3") And oForm.PaneLevel > 1 Then
                                    If validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
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
                                                    Try
                                                        If UpdateOrder(oForm) Then

                                                            Dim strType As String = String.Empty
                                                            If CType(oForm.Items.Item("12").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                                                                strType = "I"
                                                            ElseIf CType(oForm.Items.Item("13").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                                                                strType = "D"
                                                            ElseIf CType(oForm.Items.Item("14").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                                                                strType = "L"
                                                            ElseIf CType(oForm.Items.Item("24").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                                                                strType = "C"
                                                            ElseIf CType(oForm.Items.Item("43").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                                                                strType = "A"
                                                            End If

                                                            oSuccessGrid = oForm.Items.Item("19").Specific
                                                            oSuccessGrid.DataTable = oForm.DataSources.DataTables.Item("dtSuccess_" & strType)
                                                            oSuccessGrid = oForm.Items.Item("21").Specific
                                                            oSuccessGrid.DataTable = oForm.DataSources.DataTables.Item("dtFailure_" & strType)

                                                            gridSFFormat(oForm)
                                                            oForm.PaneLevel = 4
                                                        End If
                                                        oLoadForm.Close()
                                                    Catch ex As Exception
                                                        oApplication.Log.Trace_DIET_AddOn_Error(ex)
                                                        oLoadForm.Close()
                                                    End Try
                                                End If
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
                                If pVal.ItemUID = "3" And (oForm.PaneLevel = 1) Then
                                    oForm.PaneLevel = oForm.PaneLevel + 1
                                    CType(oForm.Items.Item("24").Specific, SAPbouiCOM.OptionBtn).Selected = True
                                    'checkOB(oForm)
                                ElseIf pVal.ItemUID = "17" And (oForm.PaneLevel > 1) Then
                                    oForm.PaneLevel = oForm.PaneLevel - 1
                                ElseIf pVal.ItemUID = "3" And (oForm.PaneLevel = 2) Then
                                    LoadOrder(oForm)
                                    oForm.PaneLevel = oForm.PaneLevel + 1
                                ElseIf pVal.ItemUID = "_20" And (oForm.PaneLevel = 4) Then
                                    oForm.PaneLevel = 2
                                    CType(oForm.Items.Item("12").Specific, SAPbouiCOM.OptionBtn).Selected = True
                                    checkOB(oForm)
                                    oForm.Items.Item("8").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                ElseIf pVal.ItemUID = "20" And (oForm.PaneLevel = 4) Then
                                    oForm.Close()
                                    Exit Sub
                                ElseIf pVal.ItemUID = "12" Or pVal.ItemUID = "24" Or pVal.ItemUID = "13" Or pVal.ItemUID = "14" Or pVal.ItemUID = "43" Then
                                    oForm.Freeze(True)
                                    checkOB(oForm)
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "38" Then
                                    selectAll(oForm)
                                ElseIf pVal.ItemUID = "39" Then
                                    clearAll(oForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                oDBDataSource = oForm.DataSources.DBDataSources.Item("OINV")
                                oDBDataSource_1 = oForm.DataSources.DBDataSources.Item("RDR1")
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent


                                oCFLEvento = pVal
                                oDataTable = oCFLEvento.SelectedObjects

                                If IsNothing(oDataTable) Then
                                    Exit Sub
                                End If

                                'Dim oCFL As SAPbouiCOM.ChooseFromList
                                'oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID)
                                'If oCFL.ObjectType = "2" Then

                                'End If                         


                                If pVal.ItemUID = "7" Then
                                    Dim intAddRows As Integer = oDataTable.Rows.Count
                                    If intAddRows > 0 Then
                                        oDBDataSource.SetValue("CardName", 0, oDataTable.GetValue("CardCode", 0))
                                        oDBDataSource.SetValue("NumAtCard", 0, oDataTable.GetValue("CardName", 0))
                                    End If
                                ElseIf pVal.ItemUID = "6" Then
                                    Dim intAddRows As Integer = oDataTable.Rows.Count
                                    If intAddRows > 0 Then
                                        oDBDataSource.SetValue("CardCode", 0, oDataTable.GetValue("CardCode", 0))
                                        oDBDataSource.SetValue("JrnlMemo", 0, oDataTable.GetValue("CardName", 0))
                                    End If
                                ElseIf pVal.ItemUID = "8" Then
                                    Dim intAddRows As Integer = oDataTable.Rows.Count
                                    If intAddRows > 0 Then
                                        oDBDataSource.SetValue("NumAtCard", 0, oDataTable.GetValue("CardName", 0))
                                        oDBDataSource.SetValue("CardName", 0, oDataTable.GetValue("CardCode", 0))
                                    End If
                                ElseIf pVal.ItemUID = "9" Then
                                    Dim intAddRows As Integer = oDataTable.Rows.Count
                                    If intAddRows > 0 Then
                                        oDBDataSource.SetValue("JrnlMemo", 0, oDataTable.GetValue("CardName", 0))
                                        oDBDataSource.SetValue("CardCode", 0, oDataTable.GetValue("CardCode", 0))
                                    End If
                                ElseIf pVal.ItemUID = "31" Then
                                    Dim intAddRows As Integer = oDataTable.Rows.Count
                                    If intAddRows > 0 Then
                                        oDBDataSource_1.SetValue("ItemCode", 0, oDataTable.GetValue("ItemCode", 0))
                                        oDBDataSource.SetValue("CertifNum", 0, oDataTable.GetValue("ItemName", 0))
                                    End If
                                ElseIf pVal.ItemUID = "35" Then
                                    Dim intAddRows As Integer = oDataTable.Rows.Count
                                    If intAddRows > 0 Then
                                        oDBDataSource.SetValue("CertifNum", 0, oDataTable.GetValue("ItemName", 0))
                                        oDBDataSource_1.SetValue("ItemCode", 0, oDataTable.GetValue("ItemCode", 0))
                                    End If
                                ElseIf pVal.ItemUID = "33" Then
                                    Dim intAddRows As Integer = oDataTable.Rows.Count
                                    If intAddRows > 0 Then
                                        oDBDataSource_1.SetValue("Dscription", 0, oDataTable.GetValue("ItemCode", 0))
                                        oDBDataSource.SetValue("NTSeTaxNo", 0, oDataTable.GetValue("ItemName", 0))
                                    End If
                                ElseIf pVal.ItemUID = "36" Then
                                    Dim intAddRows As Integer = oDataTable.Rows.Count
                                    If intAddRows > 0 Then
                                        oDBDataSource.SetValue("NTSeTaxNo", 0, oDataTable.GetValue("ItemName", 0))
                                        oDBDataSource_1.SetValue("Dscription", 0, oDataTable.GetValue("ItemCode", 0))
                                    End If
                                End If

                                If pVal.ItemUID = "16" _
                                     And pVal.ColUID = "TItemName" And pVal.Row > -1 Then
                                    oOrderGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    If Not IsNothing(oDataTable) Then
                                        Dim intAddRows As Integer = oDataTable.Rows.Count
                                        If intAddRows > 0 Then
                                            oOrderGrid.DataTable.SetValue("TItemCode", pVal.Row, oDataTable.GetValue("ItemCode", 0).ToString())
                                            oOrderGrid.DataTable.SetValue("TItemName", pVal.Row, oDataTable.GetValue("ItemName", 0).ToString())
                                            oOrderGrid.DataTable.SetValue("UgpEntry", pVal.Row, oDataTable.GetValue("UgpEntry", 0).ToString())
                                            oOrderGrid.DataTable.SetValue("SUoMEntry", pVal.Row, oDataTable.GetValue("SUoMEntry", 0).ToString())

                                            oOrderGrid.DataTable.SetValue("Select", pVal.Row, "Y")
                                        End If
                                    End If
                                End If

                                If pVal.ItemUID = "16" _
                                     And pVal.ColUID = "AItemName" And pVal.Row > -1 Then
                                    oOrderGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    If Not IsNothing(oDataTable) Then
                                        Dim intAddRows As Integer = oDataTable.Rows.Count
                                        If intAddRows > 0 Then
                                            oOrderGrid.DataTable.SetValue("AItemCode", pVal.Row, oDataTable.GetValue("ItemCode", 0).ToString())
                                            oOrderGrid.DataTable.SetValue("AItemName", pVal.Row, oDataTable.GetValue("ItemName", 0).ToString())
                                            oOrderGrid.DataTable.SetValue("UgpEntry", pVal.Row, oDataTable.GetValue("UgpEntry", 0).ToString())
                                            oOrderGrid.DataTable.SetValue("SUoMEntry", pVal.Row, oDataTable.GetValue("SUoMEntry", 0).ToString())

                                            oOrderGrid.DataTable.SetValue("Select", pVal.Row, "Y")
                                        End If
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "16" _
                                 And pVal.ColUID = "CShipDate" And pVal.Row > -1 Then
                                    oOrderGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    Dim shipDate As String
                                    Dim cShipDate As String

                                    shipDate = oOrderGrid.DataTable.GetValue("ShipDate", pVal.Row)
                                    cShipDate = oOrderGrid.DataTable.GetValue("CShipDate", pVal.Row)

                                    If shipDate = cShipDate Then
                                        oOrderGrid.DataTable.SetValue("Select", pVal.Row, "N")
                                    ElseIf cShipDate <> shipDate And cShipDate.Length > 0 Then
                                        oOrderGrid.DataTable.SetValue("Select", pVal.Row, "Y")
                                    End If
                                ElseIf pVal.ItemUID = "8" Then
                                    If CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value = "" Then
                                        CType(oForm.Items.Item("7").Specific, SAPbouiCOM.EditText).Value = ""
                                    End If
                                ElseIf pVal.ItemUID = "9" Then
                                    If CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value = "" Then
                                        CType(oForm.Items.Item("6").Specific, SAPbouiCOM.EditText).Value = ""
                                    End If
                                ElseIf pVal.ItemUID = "35" Then
                                    If CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value = "" Then
                                        CType(oForm.Items.Item("31").Specific, SAPbouiCOM.EditText).Value = ""
                                    End If
                                ElseIf pVal.ItemUID = "36" Then
                                    If CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value = "" Then
                                        CType(oForm.Items.Item("33").Specific, SAPbouiCOM.EditText).Value = ""
                                    End If
                                ElseIf pVal.ItemUID = "10" Then
                                    If CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value <> "" Then
                                        CType(oForm.Items.Item("26").Specific, SAPbouiCOM.EditText).Value = _
                                            CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value
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

    Private Sub LoadOrder(ByVal aform As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)

            Dim strFromCardCode, strToCardCode, strFromCardName, strToCardName, strFromDeliveryDate, strToDeliveryDate As String
            Dim strType As String = String.Empty
            Dim strFType As String = String.Empty
            Dim strConsolidateDate As String = String.Empty
            Dim strItemCode, strItemName, strTItemCode, strTItemName As String

            strFromCardCode = oForm.Items.Item("7").Specific.value
            strToCardCode = oForm.Items.Item("6").Specific.value

            strFromCardName = oForm.Items.Item("8").Specific.value
            strToCardName = oForm.Items.Item("9").Specific.value

            strFromDeliveryDate = oForm.Items.Item("10").Specific.value
            strToDeliveryDate = oForm.Items.Item("34").Specific.value

            strConsolidateDate = oForm.Items.Item("26").Specific.value

            strItemCode = oForm.Items.Item("31").Specific.value
            strItemName = oForm.Items.Item("35").Specific.value

            strTItemCode = oForm.Items.Item("33").Specific.value
            strTItemName = oForm.Items.Item("36").Specific.value

            Dim strSQLFormat As String = String.Empty
            If strConsolidateDate <> "" Then
                strSQLFormat = strConsolidateDate.Substring(0, 4) + "-" + strConsolidateDate.Substring(4, 2) + "-" + strConsolidateDate.Substring(6, 2)
            End If

            Try
                strFType = CType(oForm.Items.Item("28").Specific, SAPbouiCOM.ComboBox).Selected.Value
            Catch ex As Exception
                oApplication.Log.Trace_DIET_AddOn_Error(ex)

            End Try


            oOrderGrid = oForm.Items.Item("16").Specific
            oOrderGrid.DataTable = oForm.DataSources.DataTables.Item("dtOrder")

            If CType(oForm.Items.Item("12").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                strType = "I"
            ElseIf CType(oForm.Items.Item("13").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                strType = "D"
            ElseIf CType(oForm.Items.Item("14").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                strType = "L"
            ElseIf CType(oForm.Items.Item("24").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                strType = "C"
            ElseIf CType(oForm.Items.Item("43").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                strType = "A"
            End If

            oSuccessGrid = oForm.Items.Item("19").Specific
            oSuccessGrid.DataTable = oForm.DataSources.DataTables.Item("dtSuccess_" & strType)

            oFailureGrid = oForm.Items.Item("21").Specific
            oFailureGrid.DataTable = oForm.DataSources.DataTables.Item("dtFailure_" & strType)

            If CType(oForm.Items.Item("24").Specific, SAPbouiCOM.OptionBtn).Selected = True Then

                strType = "C"
                strqry = " Select T1.VisOrder As LineNum, Convert(VarChar(1),'Y') As 'Select', "
                strqry += " T0.DocEntry,T0.DocNum,T0.CardCode,T0.CardName "
                strqry += " ,T1.ItemCode,T1.Dscription, "
                strqry += " T1.Quantity,T1.ShipDate,Convert(DateTime,'" + strSQLFormat + "') As 'CShipDate',T1.U_DelDate,T1.U_FTYpe "
                strqry += " From ORDR T0 JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry "
                'strqry += " JOIN [@Z_OPSL] T2 ON T2.DocEntry = T1.U_PSORef  "
                strqry += " Where T1.LineStatus = 'O' "
                'strqry += " AND T2.U_IsCon = 'Y' "
                'strqry += " And Convert(VarChar(8),T1.ShipDate,112) = '" + strConsolidateDate + "'"

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
                    strqry += " And Convert(VarChar(8),T1.ShipDate,112) >= '" + strFromDeliveryDate + "'"
                End If
                If ((strFromDeliveryDate.Length = 0) And (strToDeliveryDate.Length > 0)) Then
                    strqry += " And Convert(VarChar(8),T1.ShipDate,112) <= '" + strToDeliveryDate + "'"
                End If

            ElseIf CType(oForm.Items.Item("12").Specific, SAPbouiCOM.OptionBtn).Selected = True Then

                strType = "I"
                strqry = "Select T1.VisOrder As LineNum, Convert(VarChar(1),'Y') As 'Select', "
                strqry += "T0.DocEntry,T0.DocNum,T0.CardCode,T0.CardName "
                strqry += ",T1.ItemCode,T1.Dscription, "
                strqry += "Convert(VarChar(20),'" + strTItemCode + "') As 'TItemCode',"
                strqry += "Convert(VarChar(100),'" + strTItemName.Replace("'", "''") + "') As 'TItemName', "
                strqry += "Convert(VarChar(100),'') As 'Remarks', "
                strqry += "T1.Quantity,T1.ShipDate,T1.U_FTYpe,U_DelDate,T2.UgpEntry,T2.SUoMEntry "
                strqry += "From ORDR T0 JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry "
                strqry += " JOIN OITM T2 On T1.ItemCode = T2.ItemCode "
                strqry += "Where T1.LineStatus = 'O' "

                If ((strFromCardCode.Length > 0) And (strToCardCode.Length <= 0)) Or ((strFromCardCode.Length <= 0) And (strToCardCode.Length > 0)) Then
                    strqry += " And T0.CardCode = '" + IIf(strFromCardCode.Length > 0, strFromCardCode, strToCardCode) + "'"
                End If
                If (strFromCardCode.Length > 0 And strToCardCode.Length > 0) Then
                    strqry += " And T0.CardCode Between '" + strFromCardCode + "' AND '" + strToCardCode + "'"
                End If

                If (strItemCode.Length > 0) Then
                    strqry += " And T1.ItemCode = '" + strItemCode + "'"
                End If
                If (strFType.Length > 0 And strFType.Trim() <> "S") Then
                    strqry += " And T1.U_FType = '" + strFType + "'"
                End If
                If (strFromDeliveryDate.Length > 0 And strToDeliveryDate.Length > 0) Then
                    strqry += " And Convert(VarChar(8),T1.ShipDate,112) Between '" + strFromDeliveryDate + "' AND '" + strToDeliveryDate + "'"
                End If
                If ((strFromDeliveryDate.Length > 0) And (strToDeliveryDate.Length = 0)) Then
                    strqry += " And Convert(VarChar(8),T1.ShipDate,112) >= '" + strFromDeliveryDate + "'"
                End If
                If ((strFromDeliveryDate.Length = 0) And (strToDeliveryDate.Length > 0)) Then
                    strqry += " And Convert(VarChar(8),T1.ShipDate,112) <= '" + strToDeliveryDate + "'"
                End If


            ElseIf CType(oForm.Items.Item("13").Specific, SAPbouiCOM.OptionBtn).Selected = True Then

                strType = "D"
                strqry = "Select T1.VisOrder As LineNum, Convert(VarChar(1),'Y') As 'Select', "
                strqry += "T0.DocEntry,T0.DocNum,T0.CardCode,T0.CardName "
                strqry += ",T1.ItemCode,T1.Dscription, "
                strqry += "T1.Quantity,T1.ShipDate,T1.U_DelDate AS 'CShipDate',T1.U_FTYpe,U_DelDate "
                strqry += "From ORDR T0 JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry "
                strqry += "Where T1.LineStatus = 'O' "

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
                    strqry += " And Convert(VarChar(8),T1.ShipDate,112) >= '" + strFromDeliveryDate + "'"
                End If
                If ((strFromDeliveryDate.Length = 0) And (strToDeliveryDate.Length > 0)) Then
                    strqry += " And Convert(VarChar(8),T1.ShipDate,112) <= '" + strToDeliveryDate + "'"
                End If

            ElseIf CType(oForm.Items.Item("14").Specific, SAPbouiCOM.OptionBtn).Selected = True Then

                strType = "L"
                strqry = "Select T1.VisOrder As LineNum, Convert(VarChar(1),'Y') As 'Select', "
                strqry += "T0.DocEntry,T0.DocNum,T0.CardCode,T0.CardName "
                strqry += ",T1.ItemCode,T1.Dscription, "
                strqry += "T1.Quantity,T1.ShipDate,T1.U_FTYpe,U_DelDate "
                strqry += "From ORDR T0 JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry "
                strqry += "Where T1.LineStatus = 'O' "

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
                    strqry += " And Convert(VarChar(8),T1.ShipDate,112) >= '" + strFromDeliveryDate + "'"
                End If
                If ((strFromDeliveryDate.Length = 0) And (strToDeliveryDate.Length > 0)) Then
                    strqry += " And Convert(VarChar(8),T1.ShipDate,112) <= '" + strToDeliveryDate + "'"
                End If

            ElseIf CType(oForm.Items.Item("43").Specific, SAPbouiCOM.OptionBtn).Selected = True Then

                strType = "A"

                strqry = " Select Top 1 * From "
                strqry += " ( "
                strqry += " Select Top 1 T1.VisOrder As LineNum, Convert(VarChar(1),'Y') As 'Select', "
                strqry += "T0.DocEntry,T0.DocNum,T0.CardCode,T0.CardName, "
                strqry += "Convert(VarChar(20),'" + strItemCode + "') As 'AItemCode',"
                strqry += "Convert(VarChar(100),'" + strItemName.Replace("'", "''") + "') As 'AItemName', "
                strqry += " Convert(Decimal(18,2),1) 'Quantity'  "
                strqry += ", Convert(VarChar(100),'') As 'Remarks', "
                strqry += " T1.ShipDate,T1.U_FTYpe,U_DelDate,T2.UgpEntry,T2.SUoMEntry "
                strqry += "From ORDR T0 JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry "
                strqry += " JOIN OITM T2 On T1.ItemCode = T2.ItemCode "
                strqry += "Where T1.LineStatus = 'O' "

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
                    strqry += " And Convert(VarChar(8),T1.ShipDate,112) >= '" + strFromDeliveryDate + "'"
                End If
                If ((strFromDeliveryDate.Length = 0) And (strToDeliveryDate.Length > 0)) Then
                    strqry += " And Convert(VarChar(8),T1.ShipDate,112) <= '" + strToDeliveryDate + "'"
                End If

                If (strFType.Length > 0 And strFType.Trim() <> "S") Then
                    strqry += " And T1.U_FType = '" + strFType + "'"
                End If
                strqry += " UNION ALL "

                strqry += " Select Top 1 T1.VisOrder As LineNum, Convert(VarChar(1),'Y') As 'Select', "
                strqry += "T0.DocEntry,T0.DocNum,T0.CardCode,T0.CardName, "
                strqry += "Convert(VarChar(20),'" + strItemCode + "') As 'AItemCode',"
                strqry += "Convert(VarChar(100),'" + strItemName.Replace("'", "''") + "') As 'AItemName', "
                strqry += " Convert(Decimal(18,2),1) 'Quantity'  "
                strqry += ", Convert(VarChar(100),'') As 'Remarks', "
                strqry += " T1.ShipDate,T1.U_FTYpe,U_DelDate,T2.UgpEntry,T2.SUoMEntry "
                strqry += "From ORDR T0 JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry "
                strqry += " JOIN OITM T2 On T1.ItemCode = T2.ItemCode "
                strqry += "Where T1.LineStatus = 'O' "

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
                    strqry += " And Convert(VarChar(8),T1.ShipDate,112) >= '" + strFromDeliveryDate + "'"
                End If
                If ((strFromDeliveryDate.Length = 0) And (strToDeliveryDate.Length > 0)) Then
                    strqry += " And Convert(VarChar(8),T1.ShipDate,112) <= '" + strToDeliveryDate + "'"
                End If
                strqry += " ) T0 "

            End If

            oOrderGrid.DataTable.ExecuteQuery(strqry)

            gridFormat(oForm, strType)
            fillHeader(oForm, "16")
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm.Freeze(False)
        End Try

    End Sub

    Private Sub gridFormat(ByVal oform As SAPbouiCOM.Form, ByVal strType As String)
        Try
            If strType = "C" Then

                oOrderGrid = oform.Items.Item("16").Specific
                oOrderGrid.DataTable = oform.DataSources.DataTables.Item("dtOrder")

                oOrderGrid.Columns.Item("LineNum").Visible = False

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
                oOrderGrid.Columns.Item("CardCode").Visible = False

                oOrderGrid.Columns.Item("CardName").TitleObject.Caption = "Card Name"
                oOrderGrid.Columns.Item("CardName").Editable = False

                oOrderGrid.Columns.Item("ItemCode").Visible = False

                oOrderGrid.Columns.Item("Dscription").TitleObject.Caption = "Selected Food"
                oOrderGrid.Columns.Item("Dscription").Editable = False

                oOrderGrid.Columns.Item("Quantity").TitleObject.Caption = "Quantity"
                oOrderGrid.Columns.Item("Quantity").Editable = False

                oOrderGrid.Columns.Item("ShipDate").TitleObject.Caption = "Delivery Date"
                oOrderGrid.Columns.Item("ShipDate").Editable = False

                oOrderGrid.Columns.Item("CShipDate").TitleObject.Caption = "Change Delivery Date"
                oOrderGrid.Columns.Item("CShipDate").Editable = False

                oOrderGrid.Columns.Item("U_DelDate").TitleObject.Caption = "Program Date"
                oOrderGrid.Columns.Item("U_DelDate").Editable = False

                oOrderGrid.Columns.Item("U_FTYpe").TitleObject.Caption = "Food Type"
                oOrderGrid.Columns.Item("U_FTYpe").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                oComoColumn = oOrderGrid.Columns.Item("U_FTYpe")
                oComoColumn.ValidValues.Add("BF", "Break Fast")
                oComoColumn.ValidValues.Add("LN", "Lunch")
                oComoColumn.ValidValues.Add("LS", "Lunch Side")
                oComoColumn.ValidValues.Add("SK", "Snack")
                oComoColumn.ValidValues.Add("DI", "Dinner")
                oComoColumn.ValidValues.Add("DS", "Dinner Side")
                oOrderGrid.Columns.Item("U_FTYpe").Editable = False
                oComoColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description

            ElseIf strType = "I" Then

                oOrderGrid = oform.Items.Item("16").Specific
                oOrderGrid.DataTable = oform.DataSources.DataTables.Item("dtOrder")

                oOrderGrid.Columns.Item("LineNum").Visible = False

                oOrderGrid.Columns.Item("Select").TitleObject.Caption = "Select"
                oOrderGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                oOrderGrid.Columns.Item("Select").Editable = True

                oOrderGrid.Columns.Item("TItemName").TitleObject.Caption = "Replaced Food"
                oOrderGrid.Columns.Item("TItemName").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oEditTextColumn = oOrderGrid.Columns.Item("TItemName")
                'oEditTextColumn.LinkedObjectType = "4"
                oEditTextColumn.ChooseFromListUID = "CFL_6"
                oEditTextColumn.ChooseFromListAlias = "ItemName"
                oOrderGrid.Columns.Item("TItemName").Editable = True

                oOrderGrid.Columns.Item("DocEntry").TitleObject.Caption = "Sale Order Ref"
                oOrderGrid.Columns.Item("DocEntry").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oEditTextColumn = oOrderGrid.Columns.Item("DocEntry")
                oEditTextColumn.LinkedObjectType = "17"
                oOrderGrid.Columns.Item("DocEntry").Editable = False

                oOrderGrid.Columns.Item("DocNum").TitleObject.Caption = "Order No."
                oOrderGrid.Columns.Item("DocNum").Editable = False

                oOrderGrid.Columns.Item("CardCode").TitleObject.Caption = "Card Code"
                oOrderGrid.Columns.Item("CardCode").Visible = False

                oOrderGrid.Columns.Item("CardName").TitleObject.Caption = "Card Name"
                oOrderGrid.Columns.Item("CardName").Editable = False

                oOrderGrid.Columns.Item("ItemCode").Visible = False

                oOrderGrid.Columns.Item("Dscription").TitleObject.Caption = "Selected Food"
                oOrderGrid.Columns.Item("Dscription").Editable = False

                oOrderGrid.Columns.Item("Quantity").TitleObject.Caption = "Quantity"
                oOrderGrid.Columns.Item("Quantity").Editable = False

                oOrderGrid.Columns.Item("ShipDate").TitleObject.Caption = "Delivery Date"
                oOrderGrid.Columns.Item("ShipDate").Editable = False

                oOrderGrid.Columns.Item("TItemCode").Visible = False

                oOrderGrid.Columns.Item("U_FTYpe").TitleObject.Caption = "Food Type"

                oOrderGrid.Columns.Item("U_FTYpe").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                oComoColumn = oOrderGrid.Columns.Item("U_FTYpe")
                oComoColumn.ValidValues.Add("BF", "Break Fast")
                oComoColumn.ValidValues.Add("LN", "Lunch")
                oComoColumn.ValidValues.Add("LS", "Lunch Side")
                oComoColumn.ValidValues.Add("SK", "Snack")
                oComoColumn.ValidValues.Add("DI", "Dinner")
                oComoColumn.ValidValues.Add("DS", "Dinner Side")
                oOrderGrid.Columns.Item("U_FTYpe").Editable = False
                oComoColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
                oOrderGrid.Columns.Item("U_DelDate").TitleObject.Caption = "Program Date"
                oOrderGrid.Columns.Item("U_DelDate").Editable = False
                oOrderGrid.Columns.Item("UgpEntry").Editable = False
                oOrderGrid.Columns.Item("SUoMEntry").Editable = False


                oOrderGrid.Columns.Item("Remarks").TitleObject.Caption = "Remarks"
                oOrderGrid.Columns.Item("Remarks").Editable = True
            ElseIf strType = "D" Then


                oOrderGrid = oform.Items.Item("16").Specific
                oOrderGrid.DataTable = oform.DataSources.DataTables.Item("dtOrder")

                oOrderGrid.Columns.Item("LineNum").Visible = False

                oOrderGrid.Columns.Item("Select").TitleObject.Caption = "Select"
                oOrderGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                oOrderGrid.Columns.Item("Select").Editable = True

                oOrderGrid.Columns.Item("CShipDate").TitleObject.Caption = "Replacing Del Date"
                oOrderGrid.Columns.Item("CShipDate").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oEditTextColumn = oOrderGrid.Columns.Item("CShipDate")
                oOrderGrid.Columns.Item("CShipDate").Editable = False

                oOrderGrid.Columns.Item("DocEntry").TitleObject.Caption = "Sale Order Ref"
                oOrderGrid.Columns.Item("DocEntry").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oEditTextColumn = oOrderGrid.Columns.Item("DocEntry")
                oEditTextColumn.LinkedObjectType = "17"
                oOrderGrid.Columns.Item("DocEntry").Editable = False

                oOrderGrid.Columns.Item("DocNum").TitleObject.Caption = "Order No."
                oOrderGrid.Columns.Item("DocNum").Editable = False

                oOrderGrid.Columns.Item("CardCode").TitleObject.Caption = "Card Code"
                oOrderGrid.Columns.Item("CardCode").Visible = False

                oOrderGrid.Columns.Item("CardName").TitleObject.Caption = "Card Name"
                oOrderGrid.Columns.Item("CardName").Editable = False

                oOrderGrid.Columns.Item("ItemCode").Visible = False

                oOrderGrid.Columns.Item("Dscription").TitleObject.Caption = "Selected Food"
                oOrderGrid.Columns.Item("Dscription").Editable = False

                oOrderGrid.Columns.Item("Quantity").TitleObject.Caption = "Quantity"
                oOrderGrid.Columns.Item("Quantity").Editable = False

                oOrderGrid.Columns.Item("ShipDate").TitleObject.Caption = "Ship Date"
                oOrderGrid.Columns.Item("ShipDate").Editable = False

                oOrderGrid.Columns.Item("U_FTYpe").TitleObject.Caption = "Food Type"
                oOrderGrid.Columns.Item("U_FTYpe").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                oComoColumn = oOrderGrid.Columns.Item("U_FTYpe")
                oComoColumn.ValidValues.Add("BF", "Break Fast")
                oComoColumn.ValidValues.Add("LN", "Lunch")
                oComoColumn.ValidValues.Add("LS", "Lunch Side")
                oComoColumn.ValidValues.Add("SK", "Snack")
                oComoColumn.ValidValues.Add("DI", "Dinner")
                oComoColumn.ValidValues.Add("DS", "Dinner Side")
                oOrderGrid.Columns.Item("U_FTYpe").Editable = False
                oComoColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
                oOrderGrid.Columns.Item("U_DelDate").TitleObject.Caption = "Program Date"
                oOrderGrid.Columns.Item("U_DelDate").Editable = False
            ElseIf strType = "L" Then

                oOrderGrid = oform.Items.Item("16").Specific
                oOrderGrid.DataTable = oform.DataSources.DataTables.Item("dtOrder")

                oOrderGrid.Columns.Item("LineNum").Visible = False

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
                oOrderGrid.Columns.Item("CardCode").Visible = False

                oOrderGrid.Columns.Item("CardName").TitleObject.Caption = "Card Name"
                oOrderGrid.Columns.Item("CardName").Editable = False

                oOrderGrid.Columns.Item("ItemCode").Visible = False

                oOrderGrid.Columns.Item("Dscription").TitleObject.Caption = "Selected Food"
                oOrderGrid.Columns.Item("Dscription").Editable = False

                oOrderGrid.Columns.Item("Quantity").TitleObject.Caption = "Quantity"
                oOrderGrid.Columns.Item("Quantity").Editable = False

                oOrderGrid.Columns.Item("ShipDate").TitleObject.Caption = "Ship Date"
                oOrderGrid.Columns.Item("ShipDate").Editable = False

                oOrderGrid.Columns.Item("U_FTYpe").TitleObject.Caption = "Food Type"
                oOrderGrid.Columns.Item("U_FTYpe").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                oComoColumn = oOrderGrid.Columns.Item("U_FTYpe")
                oComoColumn.ValidValues.Add("BF", "Break Fast")
                oComoColumn.ValidValues.Add("LN", "Lunch")
                oComoColumn.ValidValues.Add("LS", "Lunch Side")
                oComoColumn.ValidValues.Add("SK", "Snack")
                oComoColumn.ValidValues.Add("DI", "Dinner")
                oComoColumn.ValidValues.Add("DS", "Dinner Side")
                oOrderGrid.Columns.Item("U_FTYpe").Editable = False
                oComoColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
                oOrderGrid.Columns.Item("U_DelDate").TitleObject.Caption = "Program Date"
                oOrderGrid.Columns.Item("U_DelDate").Editable = False
            ElseIf strType = "A" Then

                oOrderGrid = oform.Items.Item("16").Specific
                oOrderGrid.DataTable = oform.DataSources.DataTables.Item("dtOrder")

                oOrderGrid.Columns.Item("LineNum").Visible = False

                oOrderGrid.Columns.Item("Select").TitleObject.Caption = "Select"
                oOrderGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                oOrderGrid.Columns.Item("Select").Editable = True

                oOrderGrid.Columns.Item("AItemName").TitleObject.Caption = "Include Food"
                oOrderGrid.Columns.Item("AItemName").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oEditTextColumn = oOrderGrid.Columns.Item("AItemName")
                oEditTextColumn.ChooseFromListUID = "CFL_6"
                oEditTextColumn.ChooseFromListAlias = "ItemName"

                oOrderGrid.Columns.Item("AItemName").Editable = True

                oOrderGrid.Columns.Item("DocEntry").TitleObject.Caption = "Sale Order Ref"
                oOrderGrid.Columns.Item("DocEntry").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oEditTextColumn = oOrderGrid.Columns.Item("DocEntry")
                oEditTextColumn.LinkedObjectType = "17"
                oOrderGrid.Columns.Item("DocEntry").Editable = False

                oOrderGrid.Columns.Item("DocNum").TitleObject.Caption = "Order No."
                oOrderGrid.Columns.Item("DocNum").Editable = False

                oOrderGrid.Columns.Item("CardCode").TitleObject.Caption = "Card Code"
                oOrderGrid.Columns.Item("CardCode").Visible = False

                oOrderGrid.Columns.Item("CardName").TitleObject.Caption = "Card Name"
                oOrderGrid.Columns.Item("CardName").Editable = False

                oOrderGrid.Columns.Item("Quantity").TitleObject.Caption = "Quantity"
                oOrderGrid.Columns.Item("Quantity").Editable = True

                oOrderGrid.Columns.Item("ShipDate").TitleObject.Caption = "Delivery Date"
                oOrderGrid.Columns.Item("ShipDate").Editable = False
                oOrderGrid.Columns.Item("AItemCode").Visible = False

                oOrderGrid.Columns.Item("U_FTYpe").TitleObject.Caption = "Food Type"
                oOrderGrid.Columns.Item("U_FTYpe").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                oComoColumn = oOrderGrid.Columns.Item("U_FTYpe")
                oComoColumn.ValidValues.Add("BF", "Break Fast")
                oComoColumn.ValidValues.Add("LN", "Lunch")
                oComoColumn.ValidValues.Add("LS", "Lunch Side")
                oComoColumn.ValidValues.Add("SK", "Snack")
                oComoColumn.ValidValues.Add("DI", "Dinner")
                oComoColumn.ValidValues.Add("DS", "Dinner Side")
                oOrderGrid.Columns.Item("U_FTYpe").Editable = False

                oComoColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
                oOrderGrid.Columns.Item("U_DelDate").TitleObject.Caption = "Program Date"
                oOrderGrid.Columns.Item("U_DelDate").Editable = False
                oOrderGrid.Columns.Item("UgpEntry").Editable = False
                oOrderGrid.Columns.Item("SUoMEntry").Editable = False


                oOrderGrid.Columns.Item("Remarks").TitleObject.Caption = "Remarks"
                oOrderGrid.Columns.Item("Remarks").Editable = True

            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub gridSFFormat(ByVal oform As SAPbouiCOM.Form)
        Try
            oSuccessGrid = oform.Items.Item("19").Specific
            oFailureGrid = oform.Items.Item("21").Specific
            If CType(oform.Items.Item("24").Specific, SAPbouiCOM.OptionBtn).Selected Then
                oSuccessGrid.Columns.Item("OrderNo").TitleObject.Caption = "Order No"
                oSuccessGrid.Columns.Item("OrderLine").Visible = False
                oSuccessGrid.Columns.Item("CustomerName").TitleObject.Caption = "Customer Name"
                oSuccessGrid.Columns.Item("ProgramDate").TitleObject.Caption = "Program Date"
                oSuccessGrid.Columns.Item("ItemCode").Visible = False
                oSuccessGrid.Columns.Item("ItemName").TitleObject.Caption = "Selected Food"
                oSuccessGrid.Columns.Item("ShipDate").TitleObject.Caption = "Delivery Date"
                oSuccessGrid.Columns.Item("CShipDate").TitleObject.Caption = "Changed Delivery Date"

                oFailureGrid.Columns.Item("OrderNo").TitleObject.Caption = "Order No"
                oFailureGrid.Columns.Item("OrderLine").Visible = False
                oFailureGrid.Columns.Item("CustomerName").TitleObject.Caption = "Customer Name"
                oFailureGrid.Columns.Item("ProgramDate").TitleObject.Caption = "Program Date"
                oFailureGrid.Columns.Item("ItemCode").Visible = False
                oFailureGrid.Columns.Item("ItemName").TitleObject.Caption = "Selected Food"
                oFailureGrid.Columns.Item("ShipDate").TitleObject.Caption = "Delivery Date"
                oFailureGrid.Columns.Item("CShipDate").TitleObject.Caption = "Changed Delivery Date"
                oFailureGrid.Columns.Item("FailReason").TitleObject.Caption = "Failure Reason"
            ElseIf CType(oform.Items.Item("12").Specific, SAPbouiCOM.OptionBtn).Selected Then
                oSuccessGrid.Columns.Item("OrderNo").TitleObject.Caption = "Order No"
                oSuccessGrid.Columns.Item("OrderLine").Visible = False
                oSuccessGrid.Columns.Item("CustomerName").TitleObject.Caption = "Customer Name"
                oSuccessGrid.Columns.Item("ProgramDate").TitleObject.Caption = "Program Date"
                oSuccessGrid.Columns.Item("ItemName").TitleObject.Caption = "Selected Food"
                oSuccessGrid.Columns.Item("CItemName").TitleObject.Caption = "Replaced Food"

                oFailureGrid.Columns.Item("OrderNo").TitleObject.Caption = "Order No"
                oFailureGrid.Columns.Item("OrderLine").Visible = False
                oFailureGrid.Columns.Item("CustomerName").TitleObject.Caption = "Customer Name"
                oFailureGrid.Columns.Item("ProgramDate").TitleObject.Caption = "Program Date"
                oFailureGrid.Columns.Item("ItemName").TitleObject.Caption = "Selected Food"
                oFailureGrid.Columns.Item("CItemName").TitleObject.Caption = "Selected Food"
                oFailureGrid.Columns.Item("FailReason").TitleObject.Caption = "Failure Reason"
            ElseIf CType(oform.Items.Item("13").Specific, SAPbouiCOM.OptionBtn).Selected Then
                oSuccessGrid.Columns.Item("OrderNo").TitleObject.Caption = "Order No"
                oSuccessGrid.Columns.Item("OrderLine").Visible = False
                oSuccessGrid.Columns.Item("CustomerName").TitleObject.Caption = "Customer Name"
                oSuccessGrid.Columns.Item("ProgramDate").TitleObject.Caption = "Program Date"
                oSuccessGrid.Columns.Item("ItemName").TitleObject.Caption = "Selected Food"
                oSuccessGrid.Columns.Item("CShipDate").TitleObject.Caption = "Delivery Date"

                oFailureGrid.Columns.Item("OrderNo").TitleObject.Caption = "Order No"
                oFailureGrid.Columns.Item("OrderLine").Visible = False
                oFailureGrid.Columns.Item("CustomerName").TitleObject.Caption = "Customer Name"
                oFailureGrid.Columns.Item("ProgramDate").TitleObject.Caption = "Program Date"
                oFailureGrid.Columns.Item("ItemName").TitleObject.Caption = "Selected Food"
                oFailureGrid.Columns.Item("CShipDate").TitleObject.Caption = "Delivery Date"
                oFailureGrid.Columns.Item("FailReason").TitleObject.Caption = "Failure Reason"
            ElseIf CType(oform.Items.Item("14").Specific, SAPbouiCOM.OptionBtn).Selected Then
                oSuccessGrid.Columns.Item("OrderNo").TitleObject.Caption = "Order No"
                oSuccessGrid.Columns.Item("OrderLine").Visible = False
                oSuccessGrid.Columns.Item("CustomerName").TitleObject.Caption = "Customer Name"
                oSuccessGrid.Columns.Item("ProgramDate").TitleObject.Caption = "Program Date"
                oSuccessGrid.Columns.Item("LineStatus").TitleObject.Caption = "Line Status"

                oFailureGrid.Columns.Item("OrderNo").TitleObject.Caption = "Order No"
                oFailureGrid.Columns.Item("OrderLine").Visible = False
                oFailureGrid.Columns.Item("CustomerName").TitleObject.Caption = "Customer Name"
                oFailureGrid.Columns.Item("ProgramDate").TitleObject.Caption = "Program Date"
                oFailureGrid.Columns.Item("LineStatus").TitleObject.Caption = "Line Status"
                oFailureGrid.Columns.Item("FailReason").TitleObject.Caption = "Failure Reason"
            ElseIf CType(oform.Items.Item("43").Specific, SAPbouiCOM.OptionBtn).Selected Then
                oSuccessGrid.Columns.Item("OrderNo").TitleObject.Caption = "Order No"
                oSuccessGrid.Columns.Item("OrderLine").Visible = False
                oSuccessGrid.Columns.Item("CustomerName").TitleObject.Caption = "Customer Name"
                oSuccessGrid.Columns.Item("ProgramDate").TitleObject.Caption = "Program Date"
                oSuccessGrid.Columns.Item("AItemName").TitleObject.Caption = "Included Food"

                oFailureGrid.Columns.Item("OrderNo").TitleObject.Caption = "Order No"
                oFailureGrid.Columns.Item("OrderLine").Visible = False
                oFailureGrid.Columns.Item("CustomerName").TitleObject.Caption = "Customer Name"
                oFailureGrid.Columns.Item("ProgramDate").TitleObject.Caption = "Program Date"
                oFailureGrid.Columns.Item("AItemName").TitleObject.Caption = "Included Food"
                oFailureGrid.Columns.Item("FailReason").TitleObject.Caption = "Failure Reason"
            End If

            'Fill Header
            fillHeader(oform, "19")
            fillHeader(oform, "21")

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

    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim _retVal As Boolean = True
        Try
            If True Then

            End If

            Dim strFrDt As String = oForm.Items.Item("10").Specific.value
            Dim strToDt As String = oForm.Items.Item("34").Specific.value
            If strFrDt.Length > 0 And strToDt.Length > 0 Then
                If CInt(strFrDt) > CInt(strToDt) Then
                    oApplication.Utilities.Message("From Date Should be Lesser than or Equal To Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If

            If oForm.PaneLevel = 2 Then
                If CType(oForm.Items.Item("24").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                    Dim strConsolidateDate As String = oForm.Items.Item("26").Specific.value
                    If strConsolidateDate.Length = 0 Then
                        oApplication.Utilities.Message("Please Selected the Consolidate Date...to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                ElseIf CType(oForm.Items.Item("12").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                    Dim strItemCode As String = oForm.Items.Item("31").Specific.value
                    Dim strTItemCode As String = oForm.Items.Item("33").Specific.value
                    Dim strFromDate As String = oForm.Items.Item("10").Specific.value
                    Dim strToDate As String = oForm.Items.Item("34").Specific.value

                    If strFromDate.Length = 0 Then
                        oApplication.Utilities.Message("Please Selected the Delivery Date...to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf strItemCode.Length = 0 Then
                        'oApplication.Utilities.Message("Please Selected the From Food...to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'Return False
                    ElseIf strTItemCode.Length = 0 Then
                        'oApplication.Utilities.Message("Please Selected the To Food...to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'Return False
                    ElseIf strItemCode.Trim() = strTItemCode.Trim() Then
                        'oApplication.Utilities.Message("To Food Should be different From From Food...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'Return False
                    End If
                ElseIf CType(oForm.Items.Item("13").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                ElseIf CType(oForm.Items.Item("14").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                    Dim strCustomer As String = oForm.Items.Item("8").Specific.value
                    If strCustomer.Length = 0 Then
                        oApplication.Utilities.Message("Please Select Customer to Procced...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If

                ElseIf CType(oForm.Items.Item("43").Specific, SAPbouiCOM.OptionBtn).Selected = True Then

                    Dim strCustomer As String = oForm.Items.Item("8").Specific.value
                    Dim strFromDate As String = oForm.Items.Item("10").Specific.value
                    Dim strIncFood As String = oForm.Items.Item("31").Specific.value
                    Dim strFType As String = String.Empty
                    Try
                        strFType = CType(oForm.Items.Item("28").Specific, SAPbouiCOM.ComboBox).Selected.Value
                    Catch ex As Exception
                        oApplication.Log.Trace_DIET_AddOn_Error(ex)

                    End Try

                    If strCustomer.Length = 0 Then
                        oApplication.Utilities.Message("Please Select Customer to Procced...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf strFromDate.Length = 0 Then
                        oApplication.Utilities.Message("Please Select Program Date to Procced...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf strIncFood.Length = 0 Then
                        oApplication.Utilities.Message("Please Include Food to Procced...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf strFType.Length = 0 Then
                        oApplication.Utilities.Message("Select Food Type to Procced...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf strFType = "S" Then
                        oApplication.Utilities.Message("Select Specific Food Type to Procced...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If

                End If
            End If

            If oForm.PaneLevel = 3 Then
                oOrderGrid = oForm.Items.Item("16").Specific
                For intRow As Integer = 0 To oOrderGrid.Rows.Count - 1
                    If oOrderGrid.DataTable.GetValue("Select", intRow).ToString() = "Y" Then
                        _retVal = True
                        Exit For
                    End If
                Next

                If _retVal = False Then
                    oApplication.Utilities.Message("No Records Selected", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If

            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Private Function UpdateOrder(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Try

            Dim _retVal As Boolean = False
            Dim intDocEntry As Integer
            Dim intLine As Integer
            Dim strDocNum As String = String.Empty
            Dim strCustomer As String = String.Empty
            Dim strFItemCode As String = String.Empty
            Dim strTItemCode As String = String.Empty
            Dim strFItemName As String = String.Empty
            Dim strTItemName As String = String.Empty
            Dim strShipDate As String = String.Empty
            Dim strCShipDate As String = String.Empty
            Dim strPrgDate As String = String.Empty
            Dim oRecordSet_U As SAPbobsCOM.Recordset
            oRecordSet_U = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            Dim oOrder As SAPbobsCOM.Documents
            Dim strType As String = String.Empty
            Dim intID_S As Integer = 0
            Dim intID_F As Integer = 0
            Dim intStatus As Integer

            If CType(oForm.Items.Item("12").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                strType = "I"
            ElseIf CType(oForm.Items.Item("13").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                strType = "D"
            ElseIf CType(oForm.Items.Item("14").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                strType = "L"
            ElseIf CType(oForm.Items.Item("24").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                strType = "C"
            ElseIf CType(oForm.Items.Item("43").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                strType = "A"
            End If

            oDTSuccess = oForm.DataSources.DataTables.Item("dtSuccess_" & strType)
            oDTFailure = oForm.DataSources.DataTables.Item("dtFailure_" & strType)

            'Consolidate Date Update....
            If strType = "C" Then   'Consolidate


                If oDTSuccess.Columns.Count = 0 Then
                    oDTSuccess.Columns.Add("OrderNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                    oDTSuccess.Columns.Add("OrderLine", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
                    oDTSuccess.Columns.Add("CustomerName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTSuccess.Columns.Add("ProgramDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 25)
                    oDTSuccess.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                    oDTSuccess.Columns.Add("ItemName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTSuccess.Columns.Add("ShipDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 25)
                    oDTSuccess.Columns.Add("CShipDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 25)
                End If

                If oDTFailure.Columns.Count = 0 Then
                    oDTFailure.Columns.Add("OrderNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                    oDTFailure.Columns.Add("OrderLine", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
                    oDTFailure.Columns.Add("CustomerName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTFailure.Columns.Add("ProgramDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 25)
                    oDTFailure.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                    oDTFailure.Columns.Add("ItemName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTFailure.Columns.Add("ShipDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 25)
                    oDTFailure.Columns.Add("CShipDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 25)
                    oDTFailure.Columns.Add("FailReason", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)
                End If



                For intRow As Integer = 0 To oOrderGrid.Rows.Count - 1
                    If oOrderGrid.DataTable.GetValue("Select", intRow).ToString() = "Y" Then

                        intDocEntry = CInt(oOrderGrid.DataTable.GetValue("DocEntry", intRow).ToString())
                        intLine = CInt(oOrderGrid.DataTable.GetValue("LineNum", intRow).ToString())


                        strDocNum = oOrderGrid.DataTable.GetValue("DocNum", intRow).ToString()
                        strCustomer = oOrderGrid.DataTable.GetValue("CardName", intRow).ToString()
                        strFItemCode = oOrderGrid.DataTable.GetValue("ItemCode", intRow).ToString()
                        strFItemName = oOrderGrid.DataTable.GetValue("Dscription", intRow).ToString()
                        strShipDate = oOrderGrid.DataTable.GetValue("ShipDate", intRow)
                        strCShipDate = oOrderGrid.DataTable.GetValue("CShipDate", intRow)
                        strPrgDate = oOrderGrid.DataTable.GetValue("U_DelDate", intRow)

                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Executing Sale Order: " + strDocNum + " For Delivery Date : " + strPrgDate

                        oOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                        If oOrder.GetByKey(intDocEntry) Then
                            Dim blnUpdate As Boolean = False
                            For index As Integer = 0 To oOrder.Lines.Count - 1
                                If intLine = index Then
                                    If strShipDate <> strCShipDate Then
                                        oOrder.Lines.SetCurrentLine(intLine)
                                        'oOrder.Lines.ShipDate = oOrder.Lines.UserFields.Fields.Item("U_DelDate").Value
                                        oOrder.Lines.ShipDate = oOrderGrid.DataTable.GetValue("CShipDate", intRow)
                                        oOrder.Lines.UserFields.Fields.Item("U_ConDate").Value = oOrderGrid.DataTable.GetValue("CShipDate", intRow)
                                        oOrder.Lines.UserFields.Fields.Item("U_IsCon").Value = "Y"
                                        blnUpdate = True
                                    Else
                                        oOrder.Lines.SetCurrentLine(intLine)
                                        oOrder.Lines.UserFields.Fields.Item("U_ConDate").Value = oOrderGrid.DataTable.GetValue("CShipDate", intRow)
                                        oOrder.Lines.UserFields.Fields.Item("U_IsCon").Value = "Y"
                                        blnUpdate = True
                                    End If
                                End If
                            Next
                            If blnUpdate Then
                                intStatus = oOrder.Update()
                            End If
                            If intStatus = 0 Then
                                oDTSuccess.Rows.Add(1)
                                oDTSuccess.SetValue("OrderNo", intID_S, strDocNum)
                                oDTSuccess.SetValue("OrderLine", intID_S, intLine)
                                oDTSuccess.SetValue("CustomerName", intID_S, strCustomer)
                                oDTSuccess.SetValue("ProgramDate", intID_S, strPrgDate)
                                oDTSuccess.SetValue("ItemCode", intID_S, strFItemCode)
                                oDTSuccess.SetValue("ItemName", intID_S, strFItemName)
                                oDTSuccess.SetValue("ShipDate", intID_S, strShipDate)
                                oDTSuccess.SetValue("CShipDate", intID_S, strCShipDate)
                                intID_S += 1
                            Else
                                oDTFailure.Rows.Add(1)
                                oDTFailure.SetValue("OrderNo", intID_F, strDocNum)
                                oDTFailure.SetValue("OrderLine", intID_F, intLine)
                                oDTFailure.SetValue("CustomerName", intID_F, strCustomer)
                                oDTFailure.SetValue("ProgramDate", intID_F, strPrgDate)
                                oDTFailure.SetValue("ItemCode", intID_F, strFItemCode)
                                oDTFailure.SetValue("ItemName", intID_F, strFItemName)
                                oDTFailure.SetValue("ShipDate", intID_F, strShipDate)
                                oDTSuccess.SetValue("CShipDate", intID_S, strShipDate)
                                oDTFailure.SetValue("FailReason", intID_F, oApplication.Company.GetLastErrorDescription().ToString())
                                intID_F += 1
                            End If
                        End If
                    End If
                Next

            End If

            'Food Change 
            intID_S = 0
            intID_F = 0
            If strType = "I" Then

                If oDTSuccess.Columns.Count = 0 Then
                    oDTSuccess.Columns.Add("OrderNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                    oDTSuccess.Columns.Add("OrderLine", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
                    oDTSuccess.Columns.Add("CustomerName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTSuccess.Columns.Add("ProgramDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 25)
                    oDTSuccess.Columns.Add("ItemName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTSuccess.Columns.Add("CItemName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                End If


                If oDTFailure.Columns.Count = 0 Then
                    oDTFailure.Columns.Add("OrderNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                    oDTFailure.Columns.Add("OrderLine", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
                    oDTFailure.Columns.Add("CustomerName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTFailure.Columns.Add("ProgramDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 25)
                    oDTFailure.Columns.Add("ItemName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTFailure.Columns.Add("CItemName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTFailure.Columns.Add("FailReason", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)
                End If


                For intRow As Integer = 0 To oOrderGrid.Rows.Count - 1
                    If oOrderGrid.DataTable.GetValue("Select", intRow).ToString() = "Y" Then

                        intDocEntry = CInt(oOrderGrid.DataTable.GetValue("DocEntry", intRow).ToString())
                        intLine = CInt(oOrderGrid.DataTable.GetValue("LineNum", intRow).ToString())

                        strDocNum = oOrderGrid.DataTable.GetValue("DocNum", intRow).ToString()
                        strCustomer = oOrderGrid.DataTable.GetValue("CardName", intRow).ToString()
                        strFItemCode = oOrderGrid.DataTable.GetValue("ItemCode", intRow).ToString()
                        strFItemName = oOrderGrid.DataTable.GetValue("Dscription", intRow).ToString()
                        strTItemCode = oOrderGrid.DataTable.GetValue("TItemCode", intRow).ToString()
                        strTItemName = oOrderGrid.DataTable.GetValue("TItemName", intRow).ToString()
                        strPrgDate = oOrderGrid.DataTable.GetValue("U_DelDate", intRow)

                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Executing Sale Order: " + strDocNum + " For Delivery Date : " + strPrgDate

                        oOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                        If oOrder.GetByKey(intDocEntry) Then

                            Dim blnUpdate As Boolean = False

                            For index As Integer = 0 To oOrder.Lines.Count - 1
                                If intLine = index Then
                                    If strTItemCode <> "" Then
                                        If strFItemCode <> strTItemCode Then
                                            oOrder.Lines.SetCurrentLine(intLine)

                                            'oOrder.Lines.ItemCode = oOrderGrid.DataTable.GetValue("TItemCode", intRow).ToString()
                                            'If oOrderGrid.DataTable.GetValue("UgpEntry", intRow).ToString().Length > 0 Then
                                            '    If oOrderGrid.DataTable.GetValue("UgpEntry", intRow).ToString() <> "-1" Then
                                            '        oOrder.Lines.UoMEntry = CInt(oOrderGrid.DataTable.GetValue("SUoMEntry", intRow).ToString())
                                            '    End If
                                            'End If
                                            'oOrder.Lines.UnitPrice = 0
                                            'oOrder.Lines.LineTotal = 0

                                            If oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Open Then
                                                'oOrder.Lines.UserFields.Fields.Item("U_CanFrom").Value = "M"
                                                oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close
                                                blnUpdate = True
                                            End If

                                            'blnUpdate = True
                                        End If
                                    End If
                                End If
                            Next

                            'If strTItemCode <> "" Then
                            '    oOrder.Lines.SetCurrentLine(intLine)
                            '    If oOrder.Lines.ItemCode <> strTItemCode Then
                            '        If oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Open Then
                            '            oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close
                            '            blnUpdate = True
                            '        End If
                            '    End If
                            'End If

                            If blnUpdate Then

                                oApplication.Company.StartTransaction()

                                intStatus = oOrder.Update()

                                If intStatus = 0 Then

                                    'Updating the Older Food Cancellation Status to Modified As Per nicole.
                                    strQuery = "Update RDR1 SET U_CanFrom = 'M' "
                                    strQuery += " Where DocEntry = '" & intDocEntry & "'"
                                    strQuery += " And VisOrder = '" & intLine & "'"
                                    oRecordSet_U.DoQuery(strQuery)

                                    Dim oTOrder As SAPbobsCOM.Documents = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                                    If oTOrder.GetByKey(intDocEntry) Then
                                        oTOrder.Lines.Add()
                                        oTOrder.Lines.ItemCode = oOrderGrid.DataTable.GetValue("TItemCode", intRow).ToString()
                                        oTOrder.Lines.Quantity = oOrder.Lines.Quantity
                                        oTOrder.Lines.ShipDate = oOrder.Lines.ShipDate
                                        oTOrder.Lines.SalesPersonCode = oOrder.Lines.SalesPersonCode
                                        oTOrder.Lines.FreeText = oOrderGrid.DataTable.GetValue("Remarks", intRow).ToString()
                                        oTOrder.Lines.UnitPrice = 0
                                        oTOrder.Lines.LineTotal = 0
                                        For index As Integer = 0 To oOrder.Lines.UserFields.Fields.Count - 1
                                            oTOrder.Lines.UserFields.Fields.Item(index).Value = oOrder.Lines.UserFields.Fields.Item(index).Value
                                        Next


                                        Dim strDisLike As String = String.Empty
                                        Dim strMedical As String = String.Empty
                                        Dim strCardCode As String = oTOrder.CardCode
                                        Dim strItemCode As String = oOrderGrid.DataTable.GetValue("TItemCode", intRow).ToString()
                                        If (oApplication.Utilities.hasBOM(strItemCode)) Then
                                            strDisLike = oApplication.Utilities.GetDisLikeItem(strCardCode, strItemCode)
                                            strMedical = oApplication.Utilities.GetMedicalItem(strCardCode, strItemCode)
                                            oApplication.Utilities.get_ChildItems(strCardCode, strItemCode, strDisLike, strMedical)
                                        Else
                                            strDisLike = oApplication.Utilities.GetDisLikeItem(strCardCode, strItemCode)
                                            strMedical = oApplication.Utilities.GetMedicalItem(strCardCode, strItemCode)
                                        End If
                                        oTOrder.Lines.UserFields.Fields.Item("U_Dislike").Value = strDisLike
                                        oTOrder.Lines.UserFields.Fields.Item("U_Medical").Value = strMedical

                                    End If

                                    intStatus = oTOrder.Update()
                                    If intStatus = 0 Then
                                        oApplication.Company.EndTransaction(BoWfTransOpt.wf_Commit)
                                    Else
                                        oApplication.Company.EndTransaction(BoWfTransOpt.wf_RollBack)
                                    End If

                                End If

                            End If


                            If intStatus = 0 And strTItemName <> "" Then

                                oDTSuccess.Rows.Add(1)
                                oDTSuccess.SetValue("OrderNo", intID_S, strDocNum)
                                oDTSuccess.SetValue("OrderLine", intID_S, intLine)
                                oDTSuccess.SetValue("CustomerName", intID_S, strCustomer)
                                oDTSuccess.SetValue("ProgramDate", intID_S, strPrgDate)
                                oDTSuccess.SetValue("ItemName", intID_S, strFItemName)
                                oDTSuccess.SetValue("CItemName", intID_S, strTItemName)
                                intID_S += 1

                            Else

                                oDTFailure.Rows.Add(1)
                                oDTFailure.SetValue("OrderNo", intID_F, strDocNum)
                                oDTFailure.SetValue("OrderLine", intID_F, intLine)
                                oDTFailure.SetValue("CustomerName", intID_F, strCustomer)
                                oDTFailure.SetValue("ProgramDate", intID_F, strPrgDate)
                                oDTFailure.SetValue("ItemName", intID_F, strFItemName)
                                oDTFailure.SetValue("CItemName", intID_F, strTItemName)
                                If strTItemName = "" Then
                                    oDTFailure.SetValue("FailReason", intID_F, "No Item Selected")
                                Else
                                    oDTFailure.SetValue("FailReason", intID_F, oApplication.Company.GetLastErrorDescription().ToString())
                                End If

                                intID_F += 1
                            End If
                        End If

                    End If
                Next
            End If

            'Change Delivery Date
            intID_S = 0
            intID_F = 0
            If strType = "D" Then

                If oDTSuccess.Columns.Count = 0 Then
                    oDTSuccess.Columns.Add("OrderNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                    oDTSuccess.Columns.Add("OrderLine", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
                    oDTSuccess.Columns.Add("CustomerName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTSuccess.Columns.Add("ProgramDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 25)
                    oDTSuccess.Columns.Add("ItemName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTSuccess.Columns.Add("CShipDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 25)
                End If

                If oDTFailure.Columns.Count = 0 Then
                    oDTFailure.Columns.Add("OrderNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                    oDTFailure.Columns.Add("OrderLine", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
                    oDTFailure.Columns.Add("CustomerName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTFailure.Columns.Add("ProgramDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 25)
                    oDTFailure.Columns.Add("ItemName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTFailure.Columns.Add("CShipDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 25)
                    oDTFailure.Columns.Add("FailReason", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)
                End If


                For intRow As Integer = 0 To oOrderGrid.Rows.Count - 1

                    If oOrderGrid.DataTable.GetValue("Select", intRow).ToString() = "Y" Then

                        intDocEntry = CInt(oOrderGrid.DataTable.GetValue("DocEntry", intRow).ToString())
                        intLine = CInt(oOrderGrid.DataTable.GetValue("LineNum", intRow).ToString())

                        strDocNum = oOrderGrid.DataTable.GetValue("DocNum", intRow).ToString()
                        strCustomer = oOrderGrid.DataTable.GetValue("CardName", intRow).ToString()
                        strFItemCode = oOrderGrid.DataTable.GetValue("ItemCode", intRow).ToString()
                        strFItemName = oOrderGrid.DataTable.GetValue("Dscription", intRow).ToString()
                        strShipDate = oOrderGrid.DataTable.GetValue("ShipDate", intRow)
                        strCShipDate = oOrderGrid.DataTable.GetValue("CShipDate", intRow)
                        strPrgDate = oOrderGrid.DataTable.GetValue("U_DelDate", intRow)

                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Executing Sale Order: " + strDocNum + " For Delivery Date : " + strPrgDate

                        oOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)

                        If oOrder.GetByKey(intDocEntry) Then
                            Dim blnUpdate As Boolean = False
                            For index As Integer = 0 To oOrder.Lines.Count - 1
                                If intLine = index Then
                                    If strShipDate <> strPrgDate Then
                                        oOrder.Lines.SetCurrentLine(intLine)
                                        oOrder.Lines.ShipDate = oOrderGrid.DataTable.GetValue("CShipDate", intRow)
                                        blnUpdate = True
                                    End If
                                End If
                            Next
                            If blnUpdate Then
                                intStatus = oOrder.Update()
                            End If
                            If intStatus = 0 Then
                                oDTSuccess.Rows.Add(1)
                                oDTSuccess.SetValue("OrderNo", intID_S, strDocNum)
                                oDTSuccess.SetValue("OrderLine", intID_S, intLine)
                                oDTSuccess.SetValue("CustomerName", intID_S, strCustomer)
                                oDTSuccess.SetValue("ProgramDate", intID_S, strPrgDate)
                                oDTSuccess.SetValue("ItemName", intID_S, strFItemName)
                                oDTSuccess.SetValue("CShipDate", intID_S, strCShipDate)

                                intID_S += 1
                            Else
                                oDTFailure.Rows.Add(1)
                                oDTFailure.SetValue("OrderNo", intID_F, strDocNum)
                                oDTFailure.SetValue("OrderLine", intID_F, intLine)
                                oDTFailure.SetValue("CustomerName", intID_F, strCustomer)
                                oDTFailure.SetValue("ProgramDate", intID_F, strPrgDate)
                                oDTFailure.SetValue("ItemName", intID_F, strFItemName)
                                oDTFailure.SetValue("CShipDate", intID_F, strCShipDate)
                                oDTFailure.SetValue("FailReason", intID_F, oApplication.Company.GetLastErrorDescription().ToString())
                                intID_F += 1
                            End If
                        End If

                    End If

                Next
            End If

            'Line Close
            intID_S = 0
            intID_F = 0
            If strType = "L" Then

                If oDTSuccess.Columns.Count = 0 Then
                    oDTSuccess.Columns.Add("OrderNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                    oDTSuccess.Columns.Add("OrderLine", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
                    oDTSuccess.Columns.Add("CustomerName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTSuccess.Columns.Add("ProgramDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 25)
                    oDTSuccess.Columns.Add("LineStatus", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 25)
                End If

                If oDTFailure.Columns.Count = 0 Then
                    oDTFailure.Columns.Add("OrderNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                    oDTFailure.Columns.Add("OrderLine", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
                    oDTFailure.Columns.Add("CustomerName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTFailure.Columns.Add("ProgramDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 25)
                    oDTFailure.Columns.Add("LineStatus", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 25)
                    oDTFailure.Columns.Add("FailReason", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)
                End If


                For intRow As Integer = 0 To oOrderGrid.Rows.Count - 1
                    If oOrderGrid.DataTable.GetValue("Select", intRow).ToString() = "Y" Then

                        intDocEntry = CInt(oOrderGrid.DataTable.GetValue("DocEntry", intRow).ToString())
                        intLine = CInt(oOrderGrid.DataTable.GetValue("LineNum", intRow).ToString())

                        strDocNum = oOrderGrid.DataTable.GetValue("DocNum", intRow).ToString()
                        strCustomer = oOrderGrid.DataTable.GetValue("CardName", intRow).ToString()
                        strFItemCode = oOrderGrid.DataTable.GetValue("ItemCode", intRow).ToString()
                        strFItemName = oOrderGrid.DataTable.GetValue("Dscription", intRow).ToString()
                        strShipDate = oOrderGrid.DataTable.GetValue("ShipDate", intRow)
                        strPrgDate = oOrderGrid.DataTable.GetValue("U_DelDate", intRow)

                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Executing Sale Order: " + strDocNum + " For Delivery Date : " + strPrgDate

                        oOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                        If oOrder.GetByKey(intDocEntry) Then

                            Dim blnUpdate As Boolean = False
                            Dim blnClose As Boolean = False

                            Dim oRecordSet_P As SAPbobsCOM.Recordset
                            oRecordSet_P = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

                            strQuery = " Select Distinct T3.U_DelDate "
                            strQuery += " From [@Z_OCPR] T0  "
                            strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
                            strQuery += " And T3.LineStatus = 'O' "
                            strQuery += " Where T3.DocEntry = '" & intDocEntry.ToString & "'"
                            oRecordSet_P.DoQuery(strQuery)
                            If oRecordSet_P.RecordCount = 1 Then
                                If oOrder.DocumentStatus = BoStatus.bost_Open Then
                                    blnClose = True
                                End If
                            Else
                                For index As Integer = 0 To oOrder.Lines.Count - 1
                                    If intLine = index Then
                                        oOrder.Lines.SetCurrentLine(intLine)
                                        If oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Open Then
                                            oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close
                                            blnUpdate = True
                                        End If
                                    End If
                                Next
                            End If

                            If blnUpdate Then
                                intStatus = oOrder.Update()
                            ElseIf blnClose Then
                                intStatus = oOrder.Close()
                            End If

                            If intStatus = 0 Then
                                oDTSuccess.Rows.Add(1)

                                oDTSuccess.SetValue("OrderNo", intID_S, strDocNum)
                                oDTSuccess.SetValue("OrderLine", intID_S, intLine)
                                oDTSuccess.SetValue("CustomerName", intID_S, strCustomer)
                                oDTSuccess.SetValue("ProgramDate", intID_S, strPrgDate)
                                oDTSuccess.SetValue("LineStatus", intID_S, "Closed")

                                strQuery = "Update RDR1 SET U_CanFrom = 'M' "
                                strQuery += " Where DocEntry = '" & intDocEntry & "'"
                                strQuery += " And VisOrder = '" & intLine & "'"
                                oRecordSet_U.DoQuery(strQuery)

                                intID_S += 1
                            Else
                                oDTFailure.Rows.Add(1)

                                oDTFailure.SetValue("OrderNo", intID_F, strDocNum)
                                oDTFailure.SetValue("OrderLine", intID_F, intLine)
                                oDTFailure.SetValue("CustomerName", intID_F, strCustomer)
                                oDTFailure.SetValue("ProgramDate", intID_F, strPrgDate)
                                oDTFailure.SetValue("LineStatus", intID_F, "Closed")
                                oDTFailure.SetValue("FailReason", intID_F, oApplication.Company.GetLastErrorDescription().ToString())

                                intID_F += 1
                            End If
                        End If
                    End If
                Next
            End If


            'Food Inclusion
            intID_S = 0
            intID_F = 0
            If strType = "A" Then

                If oDTSuccess.Columns.Count = 0 Then
                    oDTSuccess.Columns.Add("OrderNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                    oDTSuccess.Columns.Add("OrderLine", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
                    oDTSuccess.Columns.Add("CustomerName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTSuccess.Columns.Add("ProgramDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 25)
                    oDTSuccess.Columns.Add("AItemName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                End If


                If oDTFailure.Columns.Count = 0 Then
                    oDTFailure.Columns.Add("OrderNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                    oDTFailure.Columns.Add("OrderLine", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
                    oDTFailure.Columns.Add("CustomerName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTFailure.Columns.Add("ProgramDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 25)
                    oDTFailure.Columns.Add("AItemName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTFailure.Columns.Add("FailReason", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)
                End If


                For intRow As Integer = 0 To oOrderGrid.Rows.Count - 1
                    If oOrderGrid.DataTable.GetValue("Select", intRow).ToString() = "Y" Then

                        intDocEntry = CInt(oOrderGrid.DataTable.GetValue("DocEntry", intRow).ToString())
                        intLine = CInt(oOrderGrid.DataTable.GetValue("LineNum", intRow).ToString())

                        strDocNum = oOrderGrid.DataTable.GetValue("DocNum", intRow).ToString()
                        strCustomer = oOrderGrid.DataTable.GetValue("CardName", intRow).ToString()
                        strTItemCode = oOrderGrid.DataTable.GetValue("AItemCode", intRow).ToString()
                        strTItemName = oOrderGrid.DataTable.GetValue("AItemName", intRow).ToString()
                        strPrgDate = oOrderGrid.DataTable.GetValue("U_DelDate", intRow)

                        CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Executing Sale Order: " & strDocNum & " For Delivery Date : " & strPrgDate & " to Include..."

                        oOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                        If oOrder.GetByKey(intDocEntry) Then

                            Dim blnUpdate As Boolean = False

                            For index As Integer = 0 To oOrder.Lines.Count - 1
                                If intLine = index Then
                                    oOrder.Lines.SetCurrentLine(intLine)
                                    blnUpdate = True
                                End If
                            Next

                            If blnUpdate Then

                                Dim oTOrder As SAPbobsCOM.Documents = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                                If oTOrder.GetByKey(intDocEntry) Then
                                    oTOrder.Lines.Add()
                                    oTOrder.Lines.ItemCode = oOrderGrid.DataTable.GetValue("AItemCode", intRow).ToString()
                                    oTOrder.Lines.Quantity = CDbl(oOrderGrid.DataTable.GetValue("Quantity", intRow).ToString()) 'oOrder.Lines.Quantity
                                    oTOrder.Lines.ShipDate = oOrder.Lines.ShipDate
                                    oTOrder.Lines.SalesPersonCode = oOrder.Lines.SalesPersonCode
                                    oTOrder.Lines.FreeText = oOrderGrid.DataTable.GetValue("Remarks", intRow).ToString()
                                    oTOrder.Lines.UnitPrice = 0
                                    oTOrder.Lines.LineTotal = 0

                                    For index As Integer = 0 To oOrder.Lines.UserFields.Fields.Count - 1
                                        oTOrder.Lines.UserFields.Fields.Item(index).Value = oOrder.Lines.UserFields.Fields.Item(index).Value
                                    Next

                                    Dim strDisLike As String = String.Empty
                                    Dim strMedical As String = String.Empty
                                    Dim strCardCode As String = oTOrder.CardCode
                                    Dim strItemCode As String = oOrderGrid.DataTable.GetValue("AItemCode", intRow).ToString()

                                    If (oApplication.Utilities.hasBOM(strItemCode)) Then
                                        strDisLike = oApplication.Utilities.GetDisLikeItem(strCardCode, strItemCode)
                                        strMedical = oApplication.Utilities.GetMedicalItem(strCardCode, strItemCode)
                                        oApplication.Utilities.get_ChildItems(strCardCode, strItemCode, strDisLike, strMedical)
                                    Else
                                        strDisLike = oApplication.Utilities.GetDisLikeItem(strCardCode, strItemCode)
                                        strMedical = oApplication.Utilities.GetMedicalItem(strCardCode, strItemCode)
                                    End If

                                    oTOrder.Lines.UserFields.Fields.Item("U_Dislike").Value = strDisLike
                                    oTOrder.Lines.UserFields.Fields.Item("U_Medical").Value = strMedical

                                    Dim strFType As String = String.Empty
                                    Try
                                        strFType = CType(oForm.Items.Item("28").Specific, SAPbouiCOM.ComboBox).Selected.Value
                                    Catch ex As Exception
                                        oApplication.Log.Trace_DIET_AddOn_Error(ex)

                                    End Try
                                    If strFType <> "" Then
                                        oTOrder.Lines.UserFields.Fields.Item("U_FType").Value = strFType
                                    End If


                                End If

                                intStatus = oTOrder.Update()
                                If intStatus = 0 Then

                                Else

                                End If

                            End If


                            If intStatus = 0 And strTItemName <> "" Then

                                oDTSuccess.Rows.Add(1)
                                oDTSuccess.SetValue("OrderNo", intID_S, strDocNum)
                                oDTSuccess.SetValue("OrderLine", intID_S, intLine)
                                oDTSuccess.SetValue("CustomerName", intID_S, strCustomer)
                                oDTSuccess.SetValue("ProgramDate", intID_S, strPrgDate)
                                oDTSuccess.SetValue("AItemName", intID_S, strTItemName)
                                intID_S += 1

                            Else

                                oDTFailure.Rows.Add(1)
                                oDTFailure.SetValue("OrderNo", intID_F, strDocNum)
                                oDTFailure.SetValue("OrderLine", intID_F, intLine)
                                oDTFailure.SetValue("CustomerName", intID_F, strCustomer)
                                oDTFailure.SetValue("ProgramDate", intID_F, strPrgDate)
                                oDTFailure.SetValue("AItemName", intID_F, strFItemName)
                                If strTItemName = "" Then
                                    oDTFailure.SetValue("FailReason", intID_F, "No Item Selected")
                                Else
                                    oDTFailure.SetValue("FailReason", intID_F, oApplication.Company.GetLastErrorDescription().ToString())
                                End If

                                intID_F += 1
                            End If
                        End If

                    End If
                Next
            End If

            _retVal = True
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Sub checkOB(ByVal oForm As SAPbouiCOM.Form)
        Try
            rbRefresh(oForm)
            oForm.Items.Item("37").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            If CType(oForm.Items.Item("24").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                'oForm.Items.Item("5").Visible = False
                'oForm.Items.Item("10").Visible = False
                'oForm.Items.Item("34").Visible = False
                oForm.Items.Item("27").Visible = False
                oForm.Items.Item("28").Visible = False
                oForm.Items.Item("29").Visible = False
                oForm.Items.Item("31").Visible = False
                oForm.Items.Item("30").Visible = False
                oForm.Items.Item("33").Visible = False
                oForm.Items.Item("35").Visible = False
                oForm.Items.Item("36").Visible = False
                oForm.Items.Item("_42").Visible = False
                oForm.Items.Item("42").Visible = False
            ElseIf CType(oForm.Items.Item("12").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                oForm.Items.Item("25").Visible = False
                oForm.Items.Item("26").Visible = False
                oForm.Items.Item("_42").Visible = True
                oForm.Items.Item("42").Visible = True
            ElseIf CType(oForm.Items.Item("13").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                oForm.Items.Item("25").Visible = False
                oForm.Items.Item("26").Visible = False
                oForm.Items.Item("27").Visible = False
                oForm.Items.Item("28").Visible = False
                oForm.Items.Item("29").Visible = False
                oForm.Items.Item("31").Visible = False
                oForm.Items.Item("30").Visible = False
                oForm.Items.Item("33").Visible = False
                oForm.Items.Item("35").Visible = False
                oForm.Items.Item("36").Visible = False
                oForm.Items.Item("_42").Visible = False
                oForm.Items.Item("42").Visible = False
            ElseIf CType(oForm.Items.Item("14").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                oForm.Items.Item("25").Visible = False
                oForm.Items.Item("26").Visible = False
                oForm.Items.Item("27").Visible = False
                oForm.Items.Item("28").Visible = False
                oForm.Items.Item("29").Visible = False
                oForm.Items.Item("31").Visible = False
                oForm.Items.Item("30").Visible = False
                oForm.Items.Item("33").Visible = False
                oForm.Items.Item("35").Visible = False
                oForm.Items.Item("36").Visible = False
                oForm.Items.Item("_42").Visible = False
                oForm.Items.Item("42").Visible = False
            ElseIf CType(oForm.Items.Item("43").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                oForm.Items.Item("6").Visible = False
                oForm.Items.Item("9").Visible = False
                oForm.Items.Item("34").Visible = False
                oForm.Items.Item("25").Visible = False
                oForm.Items.Item("26").Visible = False
                'oForm.Items.Item("27").Visible = False
                'oForm.Items.Item("28").Visible = False
                'oForm.Items.Item("29").Visible = False
                oForm.Items.Item("31").Visible = False
                oForm.Items.Item("30").Visible = False
                oForm.Items.Item("33").Visible = False
                'oForm.Items.Item("35").Visible = False
                oForm.Items.Item("36").Visible = False
                oForm.Items.Item("_42").Visible = False
                oForm.Items.Item("42").Visible = False
                CType(oForm.Items.Item("29").Specific, SAPbouiCOM.StaticText).Caption = "Include Food"
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub rbRefresh(ByVal oForm As SAPbouiCOM.Form)

        Try
            oForm.Freeze(True)

            CType(oForm.Items.Item("29").Specific, SAPbouiCOM.StaticText).Caption = "From Food"

            oForm.Items.Item("6").Visible = True
            oForm.Items.Item("5").Visible = True
            oForm.Items.Item("9").Visible = True
            oForm.Items.Item("10").Visible = True
            oForm.Items.Item("34").Visible = True
            oForm.Items.Item("27").Visible = True
            oForm.Items.Item("28").Visible = True
            oForm.Items.Item("29").Visible = True
            oForm.Items.Item("31").Visible = True
            oForm.Items.Item("30").Visible = True
            oForm.Items.Item("33").Visible = True
            oForm.Items.Item("25").Visible = True
            oForm.Items.Item("26").Visible = True
            oForm.Items.Item("35").Visible = True
            oForm.Items.Item("36").Visible = True
            oForm.Items.Item("42").Visible = True


            CType(oForm.Items.Item("6").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("7").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("8").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("9").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("10").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("34").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("28").Specific, SAPbouiCOM.ComboBox).Select("S", SAPbouiCOM.BoSearchKey.psk_ByValue)
            CType(oForm.Items.Item("31").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("33").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("35").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("36").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("26").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("42").Specific, SAPbouiCOM.EditText).Value = ""

            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Private Sub loadCombo(ByVal oForm As SAPbouiCOM.Form)
        Try
            oCombo = oForm.Items.Item("28").Specific
            oCombo.ValidValues.Add("S", "All Food")
            oCombo.ValidValues.Add("BF", "Break Fast")
            oCombo.ValidValues.Add("LN", "Lunch")
            oCombo.ValidValues.Add("LS", "Lunch Side")
            oCombo.ValidValues.Add("SK", "Snack")
            oCombo.ValidValues.Add("DI", "Dinner")
            oCombo.ValidValues.Add("DS", "Dinner Side")
            oCombo.Select("S", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
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

            oCFL = oCFLs.Item("CFL_5")
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

            oCFL = oCFLs.Item("CFL_7")
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

            oCFL = oCFLs.Item("CFL_8")
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

            oCFL = oCFLs.Item("CFL_9")
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

End Class
