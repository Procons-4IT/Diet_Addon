Imports SAPbobsCOM

Public Class clsListener
    Inherits Object

    Private ThreadClose As New Threading.Thread(AddressOf CloseApp)
    Private WithEvents _SBO_Application As SAPbouiCOM.Application
    Private _Company As SAPbobsCOM.Company
    Private _Utilities As clsUtilities
    Private _Collection As Hashtable
    Private _LookUpCollection As Hashtable
    Private _FormUID As String
    Private _Log As clsLog_Error
    Private oMenuObject As Object
    Private oItemObject As Object
    Private oSystemForms As Object
    Dim objFilters As SAPbouiCOM.EventFilters
    Dim objFilter As SAPbouiCOM.EventFilter

#Region "New"

    Public Sub New()
        MyBase.New()
        Try
            _Company = New SAPbobsCOM.Company
            _Utilities = New clsUtilities
            _Collection = New Hashtable(10, 0.5)
            _LookUpCollection = New Hashtable(10, 0.5)
            oSystemForms = New clsSystemForms
            _Log = New clsLog_Error

            SetApplication()

        Catch ex As Exception
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
        End Try
    End Sub

#End Region

#Region "Public Properties"

    Public ReadOnly Property SBO_Application() As SAPbouiCOM.Application
        Get
            Return _SBO_Application
        End Get
    End Property

    Public ReadOnly Property Company() As SAPbobsCOM.Company
        Get
            Return _Company
        End Get
    End Property

    Public ReadOnly Property Utilities() As clsUtilities
        Get
            Return _Utilities
        End Get
    End Property

    Public ReadOnly Property Collection() As Hashtable
        Get
            Return _Collection
        End Get
    End Property

    Public ReadOnly Property LookUpCollection() As Hashtable
        Get
            Return _LookUpCollection
        End Get
    End Property

    Public ReadOnly Property Log() As clsLog_Error
        Get
            Return _Log
        End Get
    End Property

#Region "Filter"

    Public Sub SetFilter(ByVal Filters As SAPbouiCOM.EventFilters)
        oApplication.SBO_Application.SetFilter(Filters)
    End Sub

    Public Sub SetFilter()
        Try
            objFilters = New SAPbouiCOM.EventFilters()

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            objFilter.AddEx(frm_Customer) 'Customer
            objFilter.AddEx(frm_INVOICES) 'Invoice
            objFilter.AddEx(frm_Delivery) 'Delivery
            objFilter.AddEx(frm_SalesOpp) 'Sales Opp


            objFilter.AddEx(frm_Z_ODLK) 'Dislike
            objFilter.AddEx(frm_Z_OMST) 'Medical
            objFilter.AddEx(frm_Z_OCAJ) 'Calories Adjustment
            objFilter.AddEx(frm_Z_OCLP) 'Calories Plan
            objFilter.AddEx(frm_Z_OTTI) 'Check Up Timing
            objFilter.AddEx(frm_Z_OMED) 'Menu Definition
            objFilter.AddEx(frm_Z_OCRG) 'New Registration
            objFilter.AddEx(frm_Z_OCPR) 'Customer Profile
            objFilter.AddEx(frm_Z_OPSL) 'Pre Sales
            'objFilter.AddEx(frm_Z_OPSL_1) 'Select Food
            objFilter.AddEx(frm_Z_OPSL_2) 'Menu
            objFilter.AddEx(frm_Z_OPGT) 'Transfer
            objFilter.AddEx(frm_Z_OCPM) 'Customer Program
            objFilter.AddEx(frm_Z_OCSR) 'Search Wizard
            objFilter.AddEx(frm_Z_OISI) 'Program Service Item
            objFilter.AddEx(frm_Z_OCRT) 'Calories Ratio
            objFilter.AddEx(frm_Z_ODWT) 'Delivery Wizard
            objFilter.AddEx(frm_Z_OMCT) 'Missed Client
            objFilter.AddEx(frm_ITEM_MASTER) 'Item Master
            objFilter.AddEx(frm_ItemGroup) 'Item Group
            objFilter.AddEx(frm_Z_OMOT) 'Modify Order
            objFilter.AddEx(frm_Z_OFCI) 'Filter Prefix
            objFilter.AddEx(frm_Z_OIVG) 'Invoice Generation Wizard

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)
            objFilter.AddEx(frm_Customer) 'Customer
            objFilter.AddEx(frm_Z_ODLK) 'Dislike
            objFilter.AddEx(frm_Z_OMST) 'Medical
            objFilter.AddEx(frm_Z_OCAJ) 'Calories Adjustment
            objFilter.AddEx(frm_Z_OCLP) 'Calories Plan
            objFilter.AddEx(frm_Z_OTTI) 'Check Up Timing
            objFilter.AddEx(frm_Z_OMED) 'Menu Definition
            objFilter.AddEx(frm_Z_OCRG) 'New Registration
            objFilter.AddEx(frm_Z_OCPR) 'Customer Profile
            objFilter.AddEx(frm_Z_OPSL) 'Pre Sales
            objFilter.AddEx(frm_Z_OPGT) 'Transfer
            objFilter.AddEx(frm_Z_OCPM) 'Customer Program
            objFilter.AddEx(frm_Z_OCSR) 'Search Wizard
            objFilter.AddEx(frm_Z_OISI) 'Program Service Item
            objFilter.AddEx(frm_Z_OCRT) 'Calories Ratio
            objFilter.AddEx(frm_Z_ODWT) 'Delivery Wizard
            objFilter.AddEx(frm_Z_OMCT) 'Missed Client
            objFilter.AddEx(frm_Z_OMOT) 'Modify Order
            objFilter.AddEx(frm_Z_OFCI) 'Filter Prefix
            objFilter.AddEx(frm_Z_OIVG) 'Invoice Generation Wizard

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK)
            objFilter.AddEx(frm_Customer) 'Customer
            objFilter.AddEx(frm_Z_ODLK) 'Dislike
            objFilter.AddEx(frm_Z_OMST) 'Medical
            objFilter.AddEx(frm_Z_OCAJ) 'Calories Adjustment
            objFilter.AddEx(frm_Z_OCLP) 'Calories Plan
            objFilter.AddEx(frm_Z_OTTI) 'Check Up Timing
            objFilter.AddEx(frm_Z_OPSL) 'Pre Sales
            objFilter.AddEx(frm_Z_OCRT) 'Calories Ratio
            objFilter.AddEx(frm_Z_OFCI) 'Filter Prefix
            objFilter.AddEx(frm_Z_OCPM) 'Customer Program
            objFilter.AddEx(frm_Z_OCPR) 'Customer Profile
            objFilter.AddEx(frm_Z_OPGT) 'Program Transfer
            objFilter.AddEx(frm_Z_OCRG) 'Customer Registration

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            objFilter.AddEx(frm_INVOICES) 'Invoice
            objFilter.AddEx(frm_Delivery) 'Delivery
            objFilter.AddEx(frm_SalesOpp) 'Sales Opp
            objFilter.AddEx(frm_Z_OCRG) 'New Registration
            objFilter.AddEx(frm_Z_OPGT) 'Transfer
            objFilter.AddEx(frm_Customer) 'Customer
            objFilter.AddEx(frm_Z_OCPR) 'Customer Profile
            objFilter.AddEx(frm_Z_OPSL) 'Pre Sales
            objFilter.AddEx(frm_PickList) 'Pick List
            objFilter.AddEx(frm_Z_OCPM) 'Customer Program
            objFilter.AddEx(frm_Z_ODWT) 'Delivery Wizard
            objFilter.AddEx(frm_ITEM_MASTER) 'Item Master
            objFilter.AddEx(frm_ItemGroup) 'Item Group

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
            objFilter.AddEx(frm_SalesOpp) 'Sales Opp
            objFilter.AddEx(frm_Z_OCRG) 'New Registration
            objFilter.AddEx(frm_Z_OCPR) 'Customer Profile
            objFilter.AddEx(frm_PickList) 'Pick List
            objFilter.AddEx(frm_Z_OCPM) 'Customer Program
            objFilter.AddEx(frm_Z_ODWT) 'Delivery Wizard
            objFilter.AddEx(frm_Customer) 'Customer

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD)
            objFilter.AddEx(frm_Z_ODLK) 'Dislike
            objFilter.AddEx(frm_Z_OMST) 'Medical
            objFilter.AddEx(frm_Z_OCPR) 'Customer Profile
            objFilter.AddEx(frm_Z_OMED) 'Menu Definition
            objFilter.AddEx(frm_Z_OPSL) 'Pre Sales
            objFilter.AddEx(frm_Z_OPGT) 'Transfer
            objFilter.AddEx(frm_Z_OCRG) 'New Registration
            objFilter.AddEx(frm_Z_OCPM) 'Customer Program
            objFilter.AddEx(frm_Z_OISI) 'Program Service Item
            objFilter.AddEx(frm_Z_ODWT) 'Delivery Wizard

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            objFilter.AddEx(frm_Z_OCSR) 'Search Wizard
            objFilter.AddEx(frm_Z_OCAJ) 'Calories Adjustment
            objFilter.AddEx(frm_Z_OCLP) 'Calories Plan
            objFilter.AddEx(frm_Z_OTTI) 'Check Up Timing
            objFilter.AddEx(frm_Z_ODLK) 'Dislike
            objFilter.AddEx(frm_Z_OMST) 'Medical
            objFilter.AddEx(frm_Z_OCPR) 'Customer Profile
            objFilter.AddEx(frm_Z_OMED) 'Menu Definition
            objFilter.AddEx(frm_Z_OPSL) 'Pre Sales
            'objFilter.AddEx(frm_Z_OPSL_1) 'Select Food
            objFilter.AddEx(frm_Z_OPSL_2) 'Menu
            objFilter.AddEx(frm_Z_OPGT) 'Transfer
            objFilter.AddEx(frm_Z_OCRG) 'New Registration
            objFilter.AddEx(frm_INVOICES) 'Invoice
            objFilter.AddEx(frm_Z_OCPM) 'Customer Program
            objFilter.AddEx(frm_Z_OISI) 'Program Service Item
            objFilter.AddEx(frm_Z_OCRT) 'Calories Ratio
            objFilter.AddEx(frm_Z_ODWT) 'Delivery Wizard
            objFilter.AddEx(frm_Z_OMCT) 'Missed Client
            objFilter.AddEx(frm_Z_OMOT) 'Modify Order
            objFilter.AddEx(frm_Z_OFCI) 'Filter Prefix
            objFilter.AddEx(frm_Z_OIVG) 'Invoice Generation Wizard

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_RESIZE)
            objFilter.AddEx(frm_Z_OCSR) 'Search Wizard
            objFilter.AddEx(frm_Z_ODLK) 'Dislike
            objFilter.AddEx(frm_Z_OMST) 'Medical
            objFilter.AddEx(frm_Z_OCPR) 'Customer Profile
            objFilter.AddEx(frm_Z_OMED) 'Menu Definition
            objFilter.AddEx(frm_Z_OPSL) 'Pre Sales
            objFilter.AddEx(frm_Z_OPSL_2) 'Menu
            objFilter.AddEx(frm_Z_OCPM) 'Customer Program
            objFilter.AddEx(frm_Z_ODWT) 'Delivery Wizard
            objFilter.AddEx(frm_Z_OMCT) 'Missed Client
            objFilter.AddEx(frm_Z_OMOT) 'Modify Order
            objFilter.AddEx(frm_Z_OIVG) 'Invoice Generation Wizard

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
            objFilter.AddEx(frm_Z_OCSR) 'Search Wizard
            objFilter.AddEx(frm_Z_ODLK) 'Dislike
            objFilter.AddEx(frm_Z_OMST) 'Medical
            objFilter.AddEx(frm_Z_OCPR) 'Customer Profile
            objFilter.AddEx(frm_Z_OMED) 'Menu Definition
            objFilter.AddEx(frm_Z_OPSL) 'Pre Sales
            objFilter.AddEx(frm_Z_OPSL_2) 'Menu
            objFilter.AddEx(frm_Z_OPGT) 'Transfer
            objFilter.AddEx(frm_Z_OCPM) 'Customer Program
            objFilter.AddEx(frm_Z_OISI) 'Program Service Item
            objFilter.AddEx(frm_Z_ODWT) 'Delivery Wizard
            objFilter.AddEx(frm_Z_OMCT) 'Missed Client
            objFilter.AddEx(frm_Z_OMOT) 'Modify Order
            objFilter.AddEx(frm_Z_OIVG) 'Invoice Generation Wizard

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK)
            objFilter.AddEx(frm_Z_OCAJ) 'Calories Adjustment
            objFilter.AddEx(frm_Z_OCLP) 'Calories Plan
            objFilter.AddEx(frm_Z_OTTI) 'Check Up Timing
            objFilter.AddEx(frm_Z_ODLK) 'Dislike
            objFilter.AddEx(frm_Z_OMST) 'Medical
            objFilter.AddEx(frm_Z_OCPR) 'Customer Profile
            objFilter.AddEx(frm_Z_OMED) 'Menu Definition
            objFilter.AddEx(frm_Z_OPSL_2) 'Menu
            objFilter.AddEx(frm_Z_OPGT) 'Transfer
            objFilter.AddEx(frm_INVOICES) 'Invoice.
            objFilter.AddEx(frm_Z_OCPM) 'Customer Program
            objFilter.AddEx(frm_Z_OFCI) 'Filter Prefix

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
            objFilter.AddEx(frm_Z_ODLK) 'Dislike
            objFilter.AddEx(frm_Z_OMST) 'Medical
            objFilter.AddEx(frm_Z_OCPR) 'Customer Profile
            objFilter.AddEx(frm_Z_OMED) 'Menu Definition
            objFilter.AddEx(frm_Z_OPSL) 'Pre Sales
            objFilter.AddEx(frm_Z_OCRG) 'New Registration
            objFilter.AddEx(frm_INVOICES) 'Invoice
            objFilter.AddEx(frm_Z_OCPM) 'Customer Program
            objFilter.AddEx(frm_Z_OISI) 'Program Service Item
            objFilter.AddEx(frm_Z_OMCT) 'Missed Client
            objFilter.AddEx(frm_Z_OMOT) 'Modify Order
            objFilter.AddEx(frm_Z_ODWT) 'Delivery Wizard
            objFilter.AddEx(frm_Z_OIVG) 'Invoice Generation Wizard

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
            objFilter.AddEx(frm_Z_OCPR) 'Customer Profile
            objFilter.AddEx(frm_Z_OMCT) 'Missed Client

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)
            objFilter.AddEx(frm_Z_OCAJ) 'Calories Adjustment
            objFilter.AddEx(frm_Z_OCLP) 'Calories Plan
            objFilter.AddEx(frm_Z_OTTI) 'Check Up Timing
            objFilter.AddEx(frm_INVOICES)
            objFilter.AddEx(frm_Z_OFCI) 'Filter Prefix
            objFilter.AddEx(frm_Z_OMED) 'Menu Definition

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
            objFilter.AddEx(frm_Z_OCRG) 'New Registration
            objFilter.AddEx(frm_Z_OCPR) 'Customer Profile
            objFilter.AddEx(frm_Z_OCPM) 'Customer Program

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED)
            objFilter.AddEx(frm_Z_OMED) 'Menu Definition
            objFilter.AddEx(frm_INVOICES)
            objFilter.AddEx(frm_Z_OCPM) 'Cutomer Program

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE)
            objFilter.AddEx(frm_Z_OPSL) 'Pre Sales
            objFilter.AddEx(frm_Z_OCPM) 'Customer Program
            objFilter.AddEx(frm_Z_OCPR) 'Customer Profile

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
            objFilter.AddEx(frm_Z_OCPM) 'Customer Program
            objFilter.AddEx(frm_Z_ODWT) 'Delivery Wizard

            SetFilter(objFilters)
        Catch ex As Exception
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
        End Try

    End Sub
#End Region

#End Region

#Region "Data Event"
    Private Sub _SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.FormDataEvent
        Try
            Dim oForm As SAPbouiCOM.Form
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            Select Case BusinessObjectInfo.FormTypeEx
                'Case frm_INVOICES
                '    Select Case BusinessObjectInfo.EventType
                '        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                '            If Not BusinessObjectInfo.BeforeAction Then
                '                If BusinessObjectInfo.ActionSuccess Then
                '                    oApplication.Utilities.AddCustomerProgram(oForm, BusinessObjectInfo.FormTypeEx, BusinessObjectInfo.ObjectKey)
                '                End If
                '            End If
                '    End Select
                Case frm_Delivery
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                            If Not BusinessObjectInfo.BeforeAction Then
                                If BusinessObjectInfo.ActionSuccess Then
                                    oApplication.Utilities.updateCustomerProgram(BusinessObjectInfo.ObjectKey)
                                End If
                            End If
                    End Select
                Case frm_SalesOpp
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                            If Not BusinessObjectInfo.BeforeAction Then
                                If BusinessObjectInfo.ActionSuccess Then
                                    oApplication.Utilities.updateCustomerProfileMesurements(BusinessObjectInfo.ObjectKey)
                                End If
                            End If
                    End Select
                Case frm_Customer
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                            If Not BusinessObjectInfo.BeforeAction Then
                                If BusinessObjectInfo.ActionSuccess Then
                                    oApplication.Utilities.AddCustomerProfileFromCustomer(oForm, BusinessObjectInfo.ObjectKey.ToString())
                                End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                            If Not BusinessObjectInfo.BeforeAction Then
                                If BusinessObjectInfo.ActionSuccess Then
                                    Dim oCustomer As SAPbobsCOM.BusinessPartners = Nothing
                                    oCustomer = oApplication.Company.GetBusinessObject(BoObjectTypes.oBusinessPartners)
                                    If oCustomer.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                                        If oCustomer.CardType = BoCardTypes.cCustomer Then
                                            oApplication.Utilities.UpdateOpenOrderAddresses(oCustomer.CardCode) 'Update Open Order Based On Addresses if Differ.
                                        End If
                                    End If
                                End If
                            End If
                    End Select
                Case frm_PickList
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, _
                            SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                            If Not BusinessObjectInfo.BeforeAction Then
                                If BusinessObjectInfo.ActionSuccess Then

                                End If
                            End If
                    End Select
            End Select
            _FormUID = oForm.UniqueID
            If _Collection.ContainsKey(_FormUID) Then
                Dim objform As SAPbouiCOM.Form
                objform = oApplication.SBO_Application.Forms.ActiveForm()
                If 1 = 1 Then
                    oMenuObject = _Collection.Item(_FormUID)
                    oMenuObject.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                End If
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Dim aForm As SAPbouiCOM.Form
                aForm = oApplication.SBO_Application.Forms.ActiveForm()
                Select Case pVal.MenuUID
                    'Case mnu_Z_OPRM
                    '    oApplication.Utilities.ActivateMenuEvent(oApplication.SBO_Application, pVal, frm_Z_OPRM)
                    Case mnu_Z_OCLP
                        oApplication.Utilities.ActivateMenuEvent(oApplication.SBO_Application, pVal, frm_Z_OCLP)
                    Case mnu_Z_OCAJ
                        oApplication.Utilities.ActivateMenuEvent(oApplication.SBO_Application, pVal, frm_Z_OCAJ)
                    Case mnu_Z_OTTI
                        oApplication.Utilities.ActivateMenuEvent(oApplication.SBO_Application, pVal, frm_Z_OTTI)
                    Case mnu_Z_OCRT
                        oApplication.Utilities.ActivateMenuEvent(oApplication.SBO_Application, pVal, frm_Z_OCRT)
                    Case mnu_Z_OFCI
                        oApplication.Utilities.ActivateMenuEvent(oApplication.SBO_Application, pVal, frm_Z_OFCI)
                    Case mnu_Z_ODLK
                        oMenuObject = New clsDislikeItem
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Z_OMST
                        oMenuObject = New clsMedical
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                        'Case mnu_Z_OEXD
                        '    oMenuObject = New clsDietExclude
                        '    oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Z_OMED
                        oMenuObject = New clsMenuDefinition
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Z_OCRG
                        oMenuObject = New clsRegistration
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Z_OCPR
                        oMenuObject = New clsCustomerProfile
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Z_OCPM
                        oMenuObject = New clsCustomerProgram
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Z_OPSL
                        oMenuObject = New clsPreSalesOrder
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Z_OPGT
                        oMenuObject = New clsProgramTransfer
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Z_OCSR
                        oMenuObject = New clsCustomerSearch
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Z_ODWT
                        oMenuObject = New clsDeliveryWizard
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Z_OMCT
                        oMenuObject = New clsMissedCall
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Z_OMOT
                        oMenuObject = New clsModifyOrder
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Z_OIVG
                        oMenuObject = New clsInvoiceGeneration
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_ADD_ROW, mnu_DELETE_ROW, mnu_ADD, mnu_FIND, mnu_Remove, mnu_FIRST, mnu_LAST, mnu_PREVIOUS, mnu_NEXT
                        If _Collection.ContainsKey(aForm.UniqueID) Then
                            oMenuObject = _Collection.Item(aForm.UniqueID)
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                            Exit Sub
                        End If
                End Select
            Else
                Dim aForm As SAPbouiCOM.Form
                aForm = oApplication.SBO_Application.Forms.ActiveForm()
                Select Case aForm.TypeEx
                    Case frm_Z_OCPR
                        If _Collection.ContainsKey(aForm.UniqueID) And pVal.MenuUID = mnu_DELETE_ROW Then
                            oMenuObject = _Collection.Item(aForm.UniqueID)
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                            Exit Sub
                        End If
                    Case frm_Z_OCPM
                        If _Collection.ContainsKey(aForm.UniqueID) Then
                            If pVal.MenuUID = mnu_ADD Or pVal.MenuUID = mnu_FIND Or pVal.MenuUID = mnu_FIRST Or _
                                pVal.MenuUID = mnu_PREVIOUS Or pVal.MenuUID = mnu_NEXT Or pVal.MenuUID = mnu_LAST _
                                Or pVal.MenuUID = mnu_CANCELCP Or pVal.MenuUID = mnu_DELETE_ROW Or pVal.MenuUID = mnu_ADD_ROW Or pVal.MenuUID = mnu_CLOSECP Then
                                oMenuObject = _Collection.Item(aForm.UniqueID)
                                oMenuObject.MenuEvent(pVal, BubbleEvent)
                                Exit Sub
                            End If
                        End If
                    Case frm_Customer
                        Select Case pVal.MenuUID
                            Case mnu_Z_OCPR_C
                                If _Collection.ContainsKey(aForm.UniqueID) Then
                                    oItemObject = _Collection.Item(aForm.UniqueID)
                                    _Collection.Item(aForm.UniqueID).menuevent(pVal, BubbleEvent)
                                End If
                        End Select
                    Case frm_Z_OPSL
                        Select Case pVal.MenuUID
                            Case mnu_GenerateSO, mnu_ViewSO
                                If _Collection.ContainsKey(aForm.UniqueID) Then
                                    oItemObject = _Collection.Item(aForm.UniqueID)
                                    _Collection.Item(aForm.UniqueID).menuevent(pVal, BubbleEvent)
                                End If
                        End Select
                    Case frm_Z_ODLK, frm_Z_OMST, frm_Z_OCLP, frm_Z_OCAJ, frm_Z_OTTI, frm_Z_OCRT, frm_Z_OFCI
                        Select Case pVal.MenuUID
                            Case mnu_Remove
                                If _Collection.ContainsKey(aForm.UniqueID) Then
                                    oItemObject = _Collection.Item(aForm.UniqueID)
                                    _Collection.Item(aForm.UniqueID).menuevent(pVal, BubbleEvent)
                                End If
                        End Select

                End Select

            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            oMenuObject = Nothing
        End Try
    End Sub

#End Region

#Region "Item Event"

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.ItemEvent
        Try
            _FormUID = FormUID
            If pVal.BeforeAction = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Then
                Select Case pVal.FormTypeEx
                    'Case frm_Z_OPRM
                    '    If Not _Collection.ContainsKey(FormUID) Then
                    '        oItemObject = New clsProgram
                    '        oItemObject.FrmUID = FormUID
                    '        _Collection.Add(FormUID, oItemObject)
                    '    End If
                    Case frm_Z_OFCI
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsFilter
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Z_OCRT
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsCaloriesRatio
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Z_OCLP
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsCaloriesPlan
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Z_OCAJ
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsCaloriesAdjustment
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Z_OTTI
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDietTiming
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Z_ODLK
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDislikeItem
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Z_OMST
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsMedical
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                        'Case frm_Z_OEXD
                        '    If Not _Collection.ContainsKey(FormUID) Then
                        '        oItemObject = New clsDietExclude
                        '        oItemObject.FrmUID = FormUID
                        '        _Collection.Add(FormUID, oItemObject)
                        '    End If
                    Case frm_Z_OMED
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsMenuDefinition
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Z_OCRG
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsRegistration
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Z_OCPR
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsCustomerProfile
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Z_OCPM
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsCustomerProgram
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Z_OPSL
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsPreSalesOrder
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                        'Case frm_Z_OPSL_1
                        '    If Not _Collection.ContainsKey(FormUID) Then
                        '        oItemObject = New clsSelectFood
                        '        oItemObject.FrmUID = FormUID
                        '        _Collection.Add(FormUID, oItemObject)
                        '    End If
                    Case frm_Z_OPSL_2
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsFoodMenu
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Z_OPGT
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsProgramTransfer
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Customer
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsCustomer
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_INVOICES
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsInvoice
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Z_OCSR
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsCustomerSearch
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Delivery
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDelivery
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_SalesOpp
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsSalesOpp
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                        'Case frm_Z_OISI
                        '    If Not _Collection.ContainsKey(FormUID) Then
                        '        oItemObject = New clsInvoiceServiceItem
                        '        oItemObject.FrmUID = FormUID
                        '        _Collection.Add(FormUID, oItemObject)
                        '    End If
                    Case frm_Z_ODWT
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDeliveryWizard
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_Z_OMCT
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsMissedCall
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ItemGroup
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsItemGroup
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_ITEM_MASTER
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsItemMaster
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Z_OMOT
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsModifyOrder
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Z_OIVG
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsInvoiceGeneration
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                End Select
            ElseIf pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD And pVal.BeforeAction = False Then
                If _LookUpCollection.ContainsKey(FormUID) Then
                    oItemObject = _Collection.Item(_LookUpCollection.Item(FormUID))
                    If Not oItemObject Is Nothing Then
                        oItemObject.IsLookUpOpen = False
                    End If
                    _LookUpCollection.Remove(FormUID)
                End If
                If _Collection.ContainsKey(FormUID) Then
                    _Collection.Item(FormUID) = Nothing
                    _Collection.Remove(FormUID)
                End If
            End If
            If _Collection.ContainsKey(FormUID) Then
                oItemObject = _Collection.Item(FormUID)
                If oItemObject.IsLookUpOpen And pVal.BeforeAction = True Then
                    _SBO_Application.Forms.Item(oItemObject.LookUpFormUID).Select()
                    BubbleEvent = False
                    Exit Sub
                End If
                Dim oForm As SAPbouiCOM.Form
                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                If (pVal.FormTypeEx = frm_INVOICES.ToString()) Then
                    'If pVal.ItemUID = "1" And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    '    Dim strMessage As String = String.Empty
                    '    If Not Utilities.validate_Program(oForm, strMessage) Then
                    '        Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        BubbleEvent = False
                    '    End If
                    'End If
                End If
                _Collection.Item(FormUID).ItemEvent(FormUID, pVal, BubbleEvent)
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

#End Region

#Region "Right Click Event"

    Private Sub _SBO_Application_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.RightClickEvent
        Try
            Dim oForm As SAPbouiCOM.Form
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            If (oForm.TypeEx = frm_Customer.ToString()) Then
                oMenuObject = New clsCustomer
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            ElseIf oForm.TypeEx = frm_Z_OPSL Then
                oMenuObject = New clsPreSalesOrder
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            ElseIf oForm.TypeEx = frm_Z_ODLK Then
                oMenuObject = New clsDislikeItem
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            ElseIf oForm.TypeEx = frm_Z_OMST Then
                oMenuObject = New clsMedical
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            ElseIf oForm.TypeEx = frm_Z_OCAJ Then
                oMenuObject = New clsCaloriesAdjustment
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            ElseIf oForm.TypeEx = frm_Z_OCLP Then
                oMenuObject = New clsCaloriesPlan
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            ElseIf oForm.TypeEx = frm_Z_OTTI Then
                oMenuObject = New clsDietTiming
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            ElseIf oForm.TypeEx = frm_Z_OCRT Then
                oMenuObject = New clsCaloriesRatio
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            ElseIf oForm.TypeEx = frm_Z_OFCI Then
                oMenuObject = New clsFilter
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            ElseIf oForm.TypeEx = frm_Z_OCPM Then
                oMenuObject = New clsCustomerProgram
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            ElseIf oForm.TypeEx = frm_Z_OCPR Then
                oMenuObject = New clsCustomerProfile
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            ElseIf oForm.TypeEx = frm_Z_OPGT Then
                oMenuObject = New clsProgramTransfer
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            ElseIf oForm.TypeEx = frm_Z_OCRG Then
                oMenuObject = New clsRegistration
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "Application Event"

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles _SBO_Application.AppEvent
        Try
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                    _Utilities.AddRemoveMenus("RemoveMenus.xml")
                    CloseApp()
            End Select
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            MessageBox.Show(ex.Message, "Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        End Try
    End Sub

#End Region

#Region "Close Application"

    Private Sub CloseApp()
        Try
            If Not _SBO_Application Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_SBO_Application)
            End If

            If Not _Company Is Nothing Then
                If _Company.Connected Then
                    _Company.Disconnect()
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_Company)
            End If

            _Utilities = Nothing
            _Collection = Nothing
            _LookUpCollection = Nothing

            ThreadClose.Sleep(10)
            System.Windows.Forms.Application.Exit()
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        Finally
            oApplication = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

#End Region

#Region "Set Application"

    Private Sub SetApplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        Try
            If Environment.GetCommandLineArgs.Length > 1 Then
                sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
                SboGuiApi = New SAPbouiCOM.SboGuiApi
                SboGuiApi.Connect(sConnectionString)
                _SBO_Application = SboGuiApi.GetApplication()
            Else
                Throw New Exception("Connection string missing.")
            End If

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        Finally
            SboGuiApi = Nothing
        End Try
    End Sub

#End Region

#Region "Finalize"

    Protected Overrides Sub Finalize()
        Try
            MyBase.Finalize()
            '            CloseApp()

            oMenuObject = Nothing
            oItemObject = Nothing
            oSystemForms = Nothing

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            MessageBox.Show(ex.Message, "Addon Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

#End Region

    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window
        Private _hwnd As IntPtr

        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property
    End Class

End Class
