Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsCustomer
    Inherits clsBase
    Private oRecordSet As SAPbobsCOM.Recordset
    Private strQuery As String
    Private oCombo As SAPbouiCOM.ComboBox

    Public Sub New()
        MyBase.New()
    End Sub

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            Select Case pVal.MenuUID
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                Case mnu_ADD
                Case mnu_Z_OCPR_C
                    oCombo = oForm.Items.Item("40").Specific
                    If oCombo.Selected.Value = "C" Then
                        Dim strDocEntry As String = oApplication.Utilities.GetCustomerProfile(oForm)
                        Dim objCPR As clsCustomerProfile
                        objCPR = New clsCustomerProfile
                        If Not String.IsNullOrEmpty(strDocEntry) Then
                            objCPR.LoadForm(strDocEntry)
                        Else
                            objCPR.LoadForm(oForm.Items.Item("5").Specific.value, oForm.Items.Item("7").Specific.value)
                        End If
                    Else
                        oApplication.Utilities.Message("Select Customer Type for Customer Profile...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                  
                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                    If oMenuItem.SubMenus.Exists(pVal.MenuUID) Then
                        oApplication.SBO_Application.Menus.RemoveEx(pVal.MenuUID)
                    End If
            End Select
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Customer Then
                Select Case pVal.BeforeAction
                    Case True

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
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

                    End Select
            End Select
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Right Click"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            If oForm.TypeEx = frm_Customer Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'

                If (eventInfo.BeforeAction = True) Then
                    Try
                        'Customer Profile 
                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            If Not oMenuItem.SubMenus.Exists(mnu_Z_OCPR_C) Then
                                Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                                oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                oCreationPackage.UniqueID = mnu_Z_OCPR_C
                                oCreationPackage.String = "Customer Profile"
                                oCreationPackage.Enabled = True
                                oMenus = oMenuItem.SubMenus
                                oMenus.AddEx(oCreationPackage)
                            End If
                        End If
                    Catch ex As Exception
                        oApplication.Log.Trace_DIET_AddOn_Error(ex)
                        MessageBox.Show(ex.Message)
                    End Try
                Else
                    If oMenuItem.SubMenus.Exists(mnu_Z_OCPR_C) Then
                        oApplication.SBO_Application.Menus.RemoveEx(mnu_Z_OCPR_C)
                    End If
                End If
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm.Freeze(False)
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
