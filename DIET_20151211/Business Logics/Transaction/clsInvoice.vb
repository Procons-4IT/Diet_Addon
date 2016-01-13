Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsInvoice
    Inherits clsBase
    Private oEditText As SAPbouiCOM.EditText
    Private oRecordSet As SAPbobsCOM.Recordset
    Private strQuery As String
    Private oMatrix As SAPbouiCOM.Matrix
    Private oDBDataSource As SAPbouiCOM.DBDataSource

    Public Sub New()
        MyBase.New()
    End Sub

    Private Function Validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Try
            Dim strFromDate, strToDate, strCardCode As String
            Dim dtFromdate, dtToDate As Date
            Dim oRecSet As SAPbobsCOM.Recordset
            oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strCardCode = oApplication.Utilities.getEditTextvalue(aform, "4")

            oDBDataSource = aform.DataSources.DBDataSources.Item("OINV")
            Dim strCanStatuas As String = oDBDataSource.GetValue("CANCELED", 0).Trim()

            If strCanStatuas = "N" Then
                oMatrix = aform.Items.Item("38").Specific
                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    If oApplication.Utilities.getMatrixValues(oMatrix, "1", intRow) <> "" Then
                        Dim stritemcode As String = oApplication.Utilities.getMatrixValues(oMatrix, "1", intRow)
                        strFromDate = oApplication.Utilities.getMatrixValues(oMatrix, "U_Fdate", intRow)
                        strToDate = oApplication.Utilities.getMatrixValues(oMatrix, "U_Edate", intRow)
                        Dim oTest As SAPbobsCOM.Recordset
                        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oTest.DoQuery("SELECT ItemCode,ItemName  FROM OITM T0  INNER JOIN OITB T1 ON T0.[ItmsGrpCod] = T1.[ItmsGrpCod] WHERE T1.[U_Program] ='Y' and T0.ItemCode='" & stritemcode & "'")
                        If oTest.RecordCount > 0 Then
                            If strFromDate = "" Then
                                oApplication.Utilities.Message("Program From Date is missing... Line No : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                            If strToDate = "" Then
                                oApplication.Utilities.Message("Program End Date is missing... Line No : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                            dtFromdate = oApplication.Utilities.GetDateTimeValue(strFromDate)
                            dtToDate = oApplication.Utilities.GetDateTimeValue(strToDate)
                            oTest.DoQuery("Select * from [@Z_OCPM] where U_CardCode='" & strCardCode & "' and '" & dtFromdate.ToString("yyyy-MM-dd") & "' between U_PFromDate and U_PToDate And IsNull(U_Cancel,'N') = 'N' And ISNULL(U_Transfer,'N') = 'N' ")
                            If oTest.RecordCount > 0 Then
                                oApplication.Utilities.Message("Program From date is overlapped with another program for selected customer : Line No : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                            oTest.DoQuery("Select * from [@Z_OCPM] where U_CardCode='" & strCardCode & "' and '" & dtToDate.ToString("yyyy-MM-dd") & "' between U_PFromDate and U_PToDate And IsNull(U_Cancel,'N') = 'N' And ISNULL(U_Transfer,'N') = 'N' ")
                            If oTest.RecordCount > 0 Then
                                oApplication.Utilities.Message("Program End date is overlapped with another program for selected customer : Line No : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        End If
                    End If
                Next
            End If
            Return True
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            Select Case pVal.MenuUID
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                Case mnu_ADD
            End Select
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_INVOICES Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("38").Specific
                                If pVal.ItemUID = "38" And (pVal.ColUID = "U_Fdate") Then
                                    oMatrix = oForm.Items.Item("38").Specific
                                    Dim stritemcode As String = oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)
                                    Dim oTest As SAPbobsCOM.Recordset
                                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oTest.DoQuery("SELECT ItemCode,ItemName  FROM OITM T0  INNER JOIN OITB T1 ON T0.[ItmsGrpCod] = T1.[ItmsGrpCod] WHERE T1.[U_Program] ='Y' and T0.ItemCode='" & stritemcode & "'")
                                    If oTest.RecordCount <= 0 Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "38" And pVal.ColUID = "U_Edate" And pVal.CharPressed <> 9 Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("38").Specific
                                If pVal.ItemUID = "38" And (pVal.ColUID = "U_Fdate") And pVal.CharPressed <> 9 Then
                                    oMatrix = oForm.Items.Item("38").Specific
                                    Dim stritemcode As String = oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)
                                    Dim oTest As SAPbobsCOM.Recordset
                                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oTest.DoQuery("SELECT ItemCode,ItemName  FROM OITM T0  INNER JOIN OITB T1 ON T0.[ItmsGrpCod] = T1.[ItmsGrpCod] WHERE T1.[U_Program] ='Y' and T0.ItemCode='" & stritemcode & "'")
                                    If oTest.RecordCount <= 0 Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "38" And pVal.ColUID = "U_Edate" And pVal.CharPressed <> 9 Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("38").Specific
                                If pVal.ItemUID = "38" And (pVal.ColUID = "U_Fdate") And pVal.CharPressed <> 9 Then
                                    oMatrix = oForm.Items.Item("38").Specific
                                    Dim stritemcode As String = oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)
                                    Dim oTest As SAPbobsCOM.Recordset
                                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oTest.DoQuery("SELECT ItemCode,ItemName  FROM OITM T0  INNER JOIN OITB T1 ON T0.[ItmsGrpCod] = T1.[ItmsGrpCod] WHERE T1.[U_Program] ='Y' and T0.ItemCode='" & stritemcode & "'")
                                    If oTest.RecordCount <= 0 Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "38" And pVal.ColUID = "U_Edate" And pVal.CharPressed <> 9 Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
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
                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And (pVal.ColUID = "11" Or pVal.ColUID = "U_Fdate") _
                                    And pVal.Row > 0 Then
                                    oForm.Freeze(True)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    Dim stritemcode As String = oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)
                                    Dim oTest As SAPbobsCOM.Recordset
                                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oTest.DoQuery("SELECT ItemCode,ItemName  FROM OITM T0  INNER JOIN OITB T1 ON T0.[ItmsGrpCod] = T1.[ItmsGrpCod] WHERE T1.[U_Program] ='Y' and T0.ItemCode='" & stritemcode & "'")
                                    If oTest.RecordCount > 0 Then
                                        Dim strCardCode As String = oApplication.Utilities.getEditTextvalue(oForm, "4")
                                        Dim strQuantity As String = oApplication.Utilities.getMatrixValues(oMatrix, "11", pVal.Row)
                                        Dim strPFromDt As String = oApplication.Utilities.getMatrixValues(oMatrix, "U_Fdate", pVal.Row)
                                        If strPFromDt.Trim().Length > 0 Then
                                            Dim strPToDt As String = oApplication.Utilities.getProgramToDate(oForm, strCardCode, strPFromDt, strQuantity)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Edate", pVal.Row, strPToDt)
                                        End If
                                    End If
                                    oForm.Freeze(False)
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Data Events"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            Select Case BusinessObjectInfo.BeforeAction
                Case True

                Case False
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    End Select
            End Select
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Right Click"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            If oForm.TypeEx = frm_ORDR Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                If (eventInfo.BeforeAction = True) Then
                    Try

                    Catch ex As Exception
                        oApplication.Log.Trace_DIET_AddOn_Error(ex)
                        MessageBox.Show(ex.Message)
                    End Try
                Else

                End If
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Function"
    Private Sub initializeControls(ByVal oForm As SAPbouiCOM.Form)
        Try

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

#End Region

End Class
