Imports SAPbobsCOM

Public Class clsRegistration
    Inherits clsBase

    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private objForm As SAPbouiCOM.Form
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Private oEditText As SAPbouiCOM.EditText
    Private oMode As SAPbouiCOM.BoFormMode
    Private oCombo As SAPbouiCOM.ComboBox
    Private oRecordSet As SAPbobsCOM.Recordset
    Private strQuery As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Sub LoadForm()
        Try
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_OCRG) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim strUID As String = oApplication.Utilities.LoadForm1(xml_Z_OCRG, frm_Z_OCRG)
            oForm = oApplication.SBO_Application.Forms.Item(strUID)
            initialize(oForm)
            loadCombo(oForm)
            oForm.DataBrowser.BrowseBy = "3"
            oForm.DataSources.UserDataSources.Add("series", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 100)
            loadSeries(oForm)
            oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oForm.Freeze(False)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Z_OCRG Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
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
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "1"
                                        If pVal.Action_Success Then
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                                initialize(oForm)
                                                defloadSeries(oForm)
                                                If oForm.DataSources.UserDataSources.Item("series").ValueEx <> "" Then
                                                    CType(oForm.Items.Item("41").Specific, SAPbouiCOM.ComboBox).Select(oForm.DataSources.UserDataSources.Item("series").ValueEx, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                                End If
                                                oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            End If
                                        End If
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCRG")
                                If pVal.ItemUID = "7" Then
                                    Dim strDOB As String = oDBDataSource.GetValue("U_DOB", 0)
                                    If strDOB.Length > 0 Then
                                        Dim strAge As String = oApplication.Utilities.getAgebyDOB(oForm, strDOB)
                                        oDBDataSource.SetValue("U_Age", 0, strAge)
                                        oForm.Update()
                                    End If
                                ElseIf pVal.ItemUID = "29" Then
                                    'Dim strCardCode As String = "C" + oDBDataSource.GetValue("U_Mobile", 0)
                                    'oDBDataSource.SetValue("U_CardCode", 0, strCardCode)
                                    'oForm.Update()
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                'oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCRG")
                                'Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                'Dim oDataTable As SAPbouiCOM.DataTable
                                'Try
                                '    oCFLEvento = pVal
                                '    oDataTable = oCFLEvento.SelectedObjects
                                '    If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                '        If pVal.ItemUID = "16" Then
                                '            Dim intAddRows As Integer = oDataTable.Rows.Count
                                '            If intAddRows > 0 Then
                                '                oDBDataSource.SetValue("U_PrgCode", 0, oDataTable.GetValue("U_Code", 0))
                                '                oForm.DataSources.UserDataSources.Item("prgName").ValueEx = oDataTable.GetValue("U_Name", 0)
                                '                Dim strItemCode As String = oApplication.Utilities.GetProgramByItem(oForm)
                                '                oDBDataSource.SetValue("U_ItemCode", 0, strItemCode)
                                '            End If
                                '            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                '        End If
                                '    End If
                                'Catch ex As Exception 
                                ' oApplication.Log.Trace_DIET_AddOn_Error(ex)

                                'End Try
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "41" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    isManual(oForm)
                                End If
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
            Select Case pVal.BeforeAction
                Case True
                    Select Case pVal.MenuUID

                    End Select
                Case False
                    Select Case pVal.MenuUID
                        Case mnu_Z_OCRG
                            LoadForm()
                        Case mnu_ADD
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            oForm.Items.Item("4").Enabled = False
                            oForm.Items.Item("5").Enabled = False
                            oForm.Items.Item("41").Enabled = True
                            initialize(oForm)
                            defloadSeries(oForm)
                            If oForm.DataSources.UserDataSources.Item("series").ValueEx <> "" Then
                                CType(oForm.Items.Item("41").Specific, SAPbouiCOM.ComboBox).Select(oForm.DataSources.UserDataSources.Item("series").ValueEx, SAPbouiCOM.BoSearchKey.psk_ByValue)
                            End If
                            oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Case mnu_FIND
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            oForm.Items.Item("4").Enabled = True
                            oForm.Items.Item("5").Enabled = True
                            oForm.Items.Item("41").Enabled = True
                    End Select
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
            Select Case BusinessObjectInfo.BeforeAction
                Case True
                Case False
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                            If BusinessObjectInfo.ActionSuccess Then
                                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                                Dim oXmlDoc As System.Xml.XmlDocument = New Xml.XmlDocument()
                                oXmlDoc.LoadXml(BusinessObjectInfo.ObjectKey)
                                Dim DocEntry As String = oXmlDoc.SelectSingleNode("/New_RegistrationParams/DocEntry").InnerText
                                oApplication.Company.StartTransaction()
                                Try
                                    If oApplication.Utilities.CreateCustomer(oForm, DocEntry) Then
                                        If oApplication.Utilities.CreateSalesOpp(oForm, DocEntry) Then
                                            If oApplication.Utilities.AddCustomerProfile(oForm, DocEntry) Then
                                                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception
                                    oApplication.Log.Trace_DIET_AddOn_Error(ex)
                                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try
                                clearUserDataSource(oForm)
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                            If BusinessObjectInfo.ActionSuccess Then
                                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                                Dim oXmlDoc As System.Xml.XmlDocument = New Xml.XmlDocument()
                                oXmlDoc.LoadXml(BusinessObjectInfo.ObjectKey)
                                Dim DocEntry As String = oXmlDoc.SelectSingleNode("/New_RegistrationParams/DocEntry").InnerText
                                Try
                                    If oApplication.Utilities.UpdateCustomer(oForm, DocEntry) Then

                                    End If
                                Catch ex As Exception
                                    oApplication.Log.Trace_DIET_AddOn_Error(ex)

                                End Try
                                clearUserDataSource(oForm)
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            If BusinessObjectInfo.ActionSuccess Then
                                oForm.Items.Item("4").Enabled = False
                                oForm.Items.Item("5").Enabled = False
                                oForm.Items.Item("41").Enabled = False
                            End If
                    End Select
            End Select
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

#End Region

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            If oForm.TypeEx = frm_Z_OCRG Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data
                If (eventInfo.BeforeAction = True) Then
                    oMenuItem.SubMenus.Item(mnu_Remove).Enabled = False
                    oMenuItem.SubMenus.Item(mnu_CANCEL).Enabled = False
                    oMenuItem.SubMenus.Item(mnu_CLOSE).Enabled = False
                Else
                    oMenuItem.SubMenus.Item(mnu_Remove).Enabled = True
                    oMenuItem.SubMenus.Item(mnu_CANCEL).Enabled = True
                    oMenuItem.SubMenus.Item(mnu_CLOSE).Enabled = True
                End If
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select IsNull(MAX(DocEntry),0) +1 From [@Z_OCRG]")
            If Not oRecordSet.EoF Then
                oApplication.Utilities.setEditText(oForm, "4", oRecordSet.Fields.Item(0).Value.ToString())
            End If
            Dim strCustCode As String = String.Empty
            'oApplication.Utilities.GetCustomerCode(strCustCode)
            'oApplication.Utilities.setEditText(oForm, "5", strCustCode)
            CType(oForm.Items.Item("31").Specific, SAPbouiCOM.ComboBox).ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            'CType(oForm.Items.Item("42").Specific, SAPbouiCOM.CheckBox).Checked = True
            oForm.Update()
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim _retVal As Boolean = True
        Try
            Dim strCardCode, strCardName, strDOB, strCDate, strMobile, strSeries, strOccupation As String
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OCRG")

            strSeries = oDBDataSource.GetValue("U_Series", 0).Trim()
            strCardCode = oDBDataSource.GetValue("U_CardCode", 0).Trim()
            strCardName = oDBDataSource.GetValue("U_CardName", 0).Trim()
            'strGender = oDBDataSource.GetValue("U_Gender", 0)
            strDOB = oDBDataSource.GetValue("U_DOB", 0).Trim()
            strCDate = System.DateTime.Now.ToString("yyyyMMdd").Trim()
            strMobile = oDBDataSource.GetValue("U_Mobile", 0).Trim()
            strOccupation = oDBDataSource.GetValue("U_Occup", 0).Trim()

            If strSeries = "" Then
                oApplication.Utilities.Message("Enter Customer Series ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strCardName = "" Then
                oApplication.Utilities.Message("Enter Customer Name ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                'ElseIf strGender = "" Then
                '    oApplication.Utilities.Message("Select Gender ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Return False
                'ElseIf strDOB = "" Then
                '    oApplication.Utilities.Message("Enter Date of Birth ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Return False
            ElseIf strOccupation = "" Then
                oApplication.Utilities.Message("Select Occupation ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strDOB.Length > 0 And strCDate.Length > 0 And CInt(IIf(strDOB = "", 0, strDOB)) > CInt(strCDate) Then
                oApplication.Utilities.Message("Date of Birth Should be Lesser than or Equal to Current Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strMobile = "" Then
                oApplication.Utilities.Message("Enter Mobile ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Not CType(oForm.Items.Item("42").Specific, SAPbouiCOM.CheckBox).Checked Then
                If strCardCode.Length = 0 Then
                    oApplication.Utilities.Message("For Selected Series need to specify CardCode...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                ElseIf validCardCode(oForm) Then
                    oApplication.Utilities.Message("Customer Code Already Exist in Business Partner...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
        Return _retVal
    End Function

    Public Sub loadUserDataSource(ByVal aForm As SAPbouiCOM.Form)
        Try
            'Dim strDest As String = ""
            'aForm.DataSources.UserDataSources.Item("series").ValueEx = strDest
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Public Sub clearUserDataSource(ByVal aForm As SAPbouiCOM.Form)
        Try
            'aForm.DataSources.UserDataSources.Item("series").ValueEx = ""
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    'Private Sub loadCombo(ByVal oForm As SAPbouiCOM.Form)
    '    Try
    '        oCombo = oForm.Items.Item("13").Specific
    '        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        strQuery = " Select (Code+'-'+Country) As 'Code',Name From OCST "
    '        oRecordSet.DoQuery(strQuery)
    '        If Not oRecordSet.EoF Then
    '            While Not oRecordSet.EoF
    '                oCombo.ValidValues.Add(oRecordSet.Fields.Item("Code").Value, oRecordSet.Fields.Item("Name").Value)
    '                oRecordSet.MoveNext()
    '            End While
    '        End If

    '        oCombo = oForm.Items.Item("14").Specific
    '        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        strQuery = "Select Code,Name From OCRY "
    '        oRecordSet.DoQuery(strQuery)
    '        If Not oRecordSet.EoF Then
    '            While Not oRecordSet.EoF
    '                oCombo.ValidValues.Add(oRecordSet.Fields.Item("Code").Value, oRecordSet.Fields.Item("Name").Value)
    '                oRecordSet.MoveNext()
    '            End While
    '        End If

    '        oCombo = oForm.Items.Item("19").Specific
    '        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        strQuery = "Select ListNum,ListName From OPLN "
    '        oRecordSet.DoQuery(strQuery)
    '        If Not oRecordSet.EoF Then
    '            While Not oRecordSet.EoF
    '                oCombo.ValidValues.Add(oRecordSet.Fields.Item("ListNum").Value, oRecordSet.Fields.Item("ListName").Value)
    '                oRecordSet.MoveNext()
    '            End While
    '        End If

    '        oCombo = oForm.Items.Item("30").Specific
    '        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        strQuery = "Select Num,Descript From OOST "
    '        oRecordSet.DoQuery(strQuery)
    '        If Not oRecordSet.EoF Then
    '            While Not oRecordSet.EoF
    '                oCombo.ValidValues.Add(oRecordSet.Fields.Item("Num").Value, oRecordSet.Fields.Item("Descript").Value)
    '                oRecordSet.MoveNext()
    '            End While
    '        End If

    '        oCombo = oForm.Items.Item("31").Specific
    '        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        strQuery = "Select SlpCode,SlpName From OSLP "
    '        oRecordSet.DoQuery(strQuery)
    '        If Not oRecordSet.EoF Then
    '            While Not oRecordSet.EoF
    '                oCombo.ValidValues.Add(oRecordSet.Fields.Item("SlpCode").Value, oRecordSet.Fields.Item("SlpName").Value)
    '                oRecordSet.MoveNext()
    '            End While
    '        End If

    '        'Series Fill
    '        oCombo = oForm.Items.Item("20").Specific
    '        strQuery = "Select Series,SeriesName From NNM1 Where ObjectCode = '2'"
    '        oRecordSet.DoQuery(strQuery)
    '        If Not oRecordSet.EoF Then
    '            oCombo.ValidValues.Add("", "")
    '            While Not oRecordSet.EoF
    '                oCombo.ValidValues.Add(oRecordSet.Fields.Item("Series").Value, oRecordSet.Fields.Item("SeriesName").Value)
    '                oRecordSet.MoveNext()
    '            End While
    '        End If
    '        oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

    '    Catch ex As Exception 
    'oApplication.Log.Trace_DIET_AddOn_Error(ex)
    '        Throw ex 
    ''oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
    '    End Try
    'End Sub

    Private Sub loadSeries(ByVal oForm As SAPbouiCOM.Form)
        Try
            'Series Fill
            Dim strDSeries As String = String.Empty
            oCombo = oForm.Items.Item("41").Specific
            strQuery = "Select Series,SeriesName,IsManual,Locked From NNM1 Where ObjectCode = '2' And DocSubType = 'C' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oCombo.ValidValues.Add("", "")
                While Not oRecordSet.EoF
                    If oRecordSet.Fields.Item("IsManual").Value.ToString() = "N" And oRecordSet.Fields.Item("Locked").Value.ToString() = "N" Then
                        strDSeries = oRecordSet.Fields.Item("Series").Value.ToString()
                    End If
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("Series").Value, oRecordSet.Fields.Item("SeriesName").Value)
                    oRecordSet.MoveNext()
                End While
            End If
            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            strDSeries = getUserSeries(oForm)
            oCombo.Select(strDSeries, SAPbouiCOM.BoSearchKey.psk_ByValue)
            oForm.DataSources.UserDataSources.Item("series").ValueEx = strDSeries
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub defloadSeries(ByVal oForm As SAPbouiCOM.Form)
        Try
            'Series Fill
            Dim strDSeries As String = String.Empty
            oCombo = oForm.Items.Item("41").Specific
            strQuery = "Select Series,SeriesName,IsManual,Locked From NNM1 Where ObjectCode = '2' And DocSubType = 'C' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    If oRecordSet.Fields.Item("IsManual").Value.ToString() = "N" And oRecordSet.Fields.Item("Locked").Value.ToString() = "N" Then
                        strDSeries = oRecordSet.Fields.Item("Series").Value.ToString()
                    End If
                    oRecordSet.MoveNext()
                End While
            End If
            oCombo.Select(strDSeries, SAPbouiCOM.BoSearchKey.psk_ByValue)
            'oForm.DataSources.UserDataSources.Item("series").ValueEx = strDSeries
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub isManual(ByVal oForm As SAPbouiCOM.Form)
        Try
            oCombo = oForm.Items.Item("41").Specific
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            strQuery = "Select IsManual From NNM1 Where ObjectCode = '2' And Series = '" + oCombo.Selected.Value.ToString() + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                If oRecordSet.Fields.Item("IsManual").Value.ToString() = "Y" Then
                    CType(oForm.Items.Item("42").Specific, SAPbouiCOM.CheckBox).Checked = False
                    oForm.Items.Item("5").Enabled = True
                Else
                    CType(oForm.Items.Item("42").Specific, SAPbouiCOM.CheckBox).Checked = True
                    oForm.Items.Item("28").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    oForm.Items.Item("5").Enabled = False
                End If
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Function validCardCode(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Dim _retval As Boolean = False
        Try
            strQuery = "Select CardCode From OCRD Where CardCode ='" + oForm.Items.Item("5").Specific.value + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                _retval = True
            End If
            Return _retval
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Private Function getUserSeries(ByVal oForm As SAPbouiCOM.Form) As Integer
        Dim _retVal As Integer = 0
        Try
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim oSeriesService As SAPbobsCOM.SeriesService
            Dim oSeries As SAPbobsCOM.Series
            Dim oDocumentTypeParams As SAPbobsCOM.DocumentTypeParams
            oCmpSrv = oApplication.Company.GetCompanyService()
            oSeriesService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.SeriesService)
            oSeries = oSeriesService.GetDataInterface(SAPbobsCOM.SeriesServiceDataInterfaces.ssdiSeries)
            oDocumentTypeParams = oSeriesService.GetDataInterface(SeriesServiceDataInterfaces.ssdiDocumentTypeParams)
            oDocumentTypeParams.Document = 2
            oDocumentTypeParams.DocumentSubType = "C"
            oSeries = oSeriesService.GetDefaultSeries(oDocumentTypeParams)
            _retVal = oSeries.Series
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Private Sub loadCombo(ByVal oForm As SAPbouiCOM.Form)
        Try
            oCombo = oForm.Items.Item("39").Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select GroupCode,GroupName From [OCRG] Where GroupType = 'C' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("GroupCode").Value, oRecordSet.Fields.Item("GroupName").Value)
                    oRecordSet.MoveNext()
                End While
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

#End Region

End Class
