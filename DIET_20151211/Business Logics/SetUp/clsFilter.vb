Public Class clsFilter
    Inherits clsBase

    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private objMatrix As SAPbouiCOM.Matrix
    Private objForm As SAPbouiCOM.Form
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oMode As SAPbouiCOM.BoFormMode
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Private oRecordSet As SAPbobsCOM.Recordset
    Private strQuery As String = String.Empty
    Private intMatrixSelectedRow As Integer

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Z_OFCI Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "1" Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        If validate() Then
                                            If Not validateData() Then
                                                oApplication.Utilities.Message("Filter Code Already defined...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                BubbleEvent = False
                                            End If
                                        Else
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK, SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("3").Specific
                                intMatrixSelectedRow = pVal.Row
                                If pVal.ItemUID = "3" And pVal.ColUID <> "U_Code" Then
                                    If pVal.Row = oMatrix.VisualRowCount Then
                                        Dim strCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "U_Code", pVal.Row)
                                        If strCode.Trim().Length = 0 Then
                                            oApplication.Utilities.Message("Enter Filter Code to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("3").Specific
                                oMatrix.Columns.Item("DocEntry").Visible = False
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Remove
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = True Then
                        If oApplication.SBO_Application.MessageBox("This action will delete the current document. Are you sure you want to proceed?", , "Yes", "No") = 2 Then
                            BubbleEvent = False
                            Exit Sub
                        Else

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
        If oForm.TypeEx = frm_Z_OFCI Then
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

#Region "Data Event"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            MsgBox(ex.Message)
        End Try
    End Sub
#End Region

#Region "Function"

    Public Function validate() As Boolean
        Try
            Dim _retVal As Boolean = True
            oMatrix = oForm.Items.Item("3").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim strCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "U_Code", index)
                If index = oMatrix.VisualRowCount Then
                    Dim strName As String = oApplication.Utilities.getMatrixValues(oMatrix, "U_Name", index)
                    If strCode.Trim().Length = 0 And (strName.Trim.Length > 0) Then
                        _retVal = False
                        oApplication.Utilities.Message("Fitler Code Cannot be Empty...in Row No :" + index.ToString(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit For
                    End If
                ElseIf strCode.Trim().Length = 0 Then
                    _retVal = False
                    oApplication.Utilities.Message("Fitler Code Cannot be Empty...in Row No :" + index.ToString(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit For
                End If
            Next
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function validateData() As Boolean
        Try
            Dim _retVal As Boolean = True
            oMatrix = oForm.Items.Item("3").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount
                Dim strCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "U_Code", index)
                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    If index <> intRow Then
                        Dim strCode1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "U_Code", intRow)
                        If strCode = strCode1 Then
                            _retVal = False
                            Return _retVal
                        End If
                    End If
                Next
            Next
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

#End Region
End Class
