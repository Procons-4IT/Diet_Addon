Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System


Public Class clsItemGroup
    Inherits clsBase

    Private oRecordSet As SAPbobsCOM.Recordset
    Private strQuery As String
    Private oCheckBox As SAPbouiCOM.CheckBox
    Dim oDBDataSource As SAPbouiCOM.DBDataSource

    Private Sub LoadForm()
        Try
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm.Freeze(False)
        End Try
    End Sub

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.BeforeAction
                Case True

                Case False
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
            If pVal.FormTypeEx = frm_ItemGroup Then
                Select Case pVal.BeforeAction
                    Case True

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oForm.PaneLevel = 1
                                initializeControls(oForm)
                                dataBind(oForm)
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

#Region "Form Data Event"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Dim oForm As SAPbouiCOM.Form
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            Select Case BusinessObjectInfo.FormTypeEx

            End Select

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

    Private Sub initializeControls(ByVal oForm As SAPbouiCOM.Form)
        Try

            oDBDataSource = oForm.DataSources.DBDataSources.Add("OITB")

            oApplication.Utilities.AddControls(oForm, "_125", "123", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 1, 1, "", "Select Food Category", 50, 0, 0, False)
            oForm.Items.Item("_125").Visible = True
            oForm.Items.Item("_125").Left = oForm.Items.Item("123").Left
            oForm.Items.Item("_125").Top = oForm.Items.Item("123").Top + oForm.Items.Item("123").Height + 5
            oForm.Items.Item("_125").Width = oForm.Items.Item("123").Width
            oForm.Items.Item("_125").Height = oForm.Items.Item("123").Height
            oForm.Items.Item("_125").TextStyle = 5

            oApplication.Utilities.AddControls(oForm, "_124", "123", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE, "DOWN", 1, 1, "", "Sample", 50, 0, 0, False)
            oForm.Items.Item("_124").Visible = True
            oForm.Items.Item("_124").Left = oForm.Items.Item("123").Left
            oForm.Items.Item("_124").Top = oForm.Items.Item("123").Top + oForm.Items.Item("123").Height + 25
            oForm.Items.Item("_124").Width = oForm.Items.Item("6").Width + 75
            oForm.Items.Item("_124").Height = oForm.Items.Item("123").Height + 75

            oApplication.Utilities.AddControls(oForm, "_126", "123", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX, "DOWN", 1, 1, "", "Break Fast", 50, 0, 0, False)
            oForm.Items.Item("_126").Visible = True
            oForm.Items.Item("_126").Left = oForm.Items.Item("123").Left + 5
            oForm.Items.Item("_126").Top = oForm.Items.Item("_124").Top + 20
            oForm.Items.Item("_126").Width = oForm.Items.Item("123").Width
            oForm.Items.Item("_126").Height = oForm.Items.Item("123").Height

            oApplication.Utilities.AddControls(oForm, "_127", "123", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX, "DOWN", 1, 1, "", "Lunch", 50, 0, 0, False)
            oForm.Items.Item("_127").Visible = True
            oForm.Items.Item("_127").Left = oForm.Items.Item("123").Left + 5
            oForm.Items.Item("_127").Top = oForm.Items.Item("_126").Top + Form.Items.Item("_126").Height + 1
            oForm.Items.Item("_127").Width = oForm.Items.Item("123").Width
            oForm.Items.Item("_127").Height = oForm.Items.Item("123").Height


            oApplication.Utilities.AddControls(oForm, "_128", "123", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX, "DOWN", 1, 1, "", "Lunch Side", 50, 0, 0, False)
            oForm.Items.Item("_128").Visible = True
            oForm.Items.Item("_128").Left = oForm.Items.Item("123").Left + 5
            oForm.Items.Item("_128").Top = oForm.Items.Item("_127").Top + +oForm.Items.Item("_127").Height + 1
            oForm.Items.Item("_128").Width = oForm.Items.Item("123").Width
            oForm.Items.Item("_128").Height = oForm.Items.Item("123").Height

            oApplication.Utilities.AddControls(oForm, "_129", "123", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX, "DOWN", 1, 1, "", "Snacks", 50, 0, 0, False)
            oForm.Items.Item("_129").Visible = True
            oForm.Items.Item("_129").Left = oForm.Items.Item("_126").Left + oForm.Items.Item("_126").Width + 5
            oForm.Items.Item("_129").Top = oForm.Items.Item("_126").Top
            oForm.Items.Item("_129").Width = oForm.Items.Item("123").Width
            oForm.Items.Item("_129").Height = oForm.Items.Item("123").Height

            oApplication.Utilities.AddControls(oForm, "_130", "123", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX, "DOWN", 1, 1, "", "Dinner", 50, 0, 0, False)
            oForm.Items.Item("_130").Visible = True
            oForm.Items.Item("_130").Left = oForm.Items.Item("_127").Left + oForm.Items.Item("_127").Width + 5
            oForm.Items.Item("_130").Top = oForm.Items.Item("_127").Top
            oForm.Items.Item("_130").Width = oForm.Items.Item("123").Width
            oForm.Items.Item("_130").Height = oForm.Items.Item("123").Height

            oApplication.Utilities.AddControls(oForm, "_131", "123", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX, "DOWN", 1, 1, "", "Dinner Side", 50, 0, 0, False)
            oForm.Items.Item("_131").Visible = True
            oForm.Items.Item("_131").Left = oForm.Items.Item("_128").Left + oForm.Items.Item("_128").Width + 5
            oForm.Items.Item("_131").Top = oForm.Items.Item("_128").Top
            oForm.Items.Item("_131").Width = oForm.Items.Item("123").Width
            oForm.Items.Item("_131").Height = oForm.Items.Item("123").Height

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub dataBind(ByVal oForm As SAPbouiCOM.Form)
        Try
            oCheckBox = oForm.Items.Item("_126").Specific
            oCheckBox.DataBind.SetBound(True, "OITB", "U_BF")

            oCheckBox = oForm.Items.Item("_127").Specific
            oCheckBox.DataBind.SetBound(True, "OITB", "U_LN")

            oCheckBox = oForm.Items.Item("_128").Specific
            oCheckBox.DataBind.SetBound(True, "OITB", "U_LS")

            oCheckBox = oForm.Items.Item("_129").Specific
            oCheckBox.DataBind.SetBound(True, "OITB", "U_SK")

            oCheckBox = oForm.Items.Item("_130").Specific
            oCheckBox.DataBind.SetBound(True, "OITB", "U_DN")

            oCheckBox = oForm.Items.Item("_131").Specific
            oCheckBox.DataBind.SetBound(True, "OITB", "U_DS")

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub
End Class
