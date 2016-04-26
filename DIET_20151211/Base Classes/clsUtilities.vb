Imports System.Xml
Imports System.Collections.Specialized
Imports System.IO
Imports SAPbobsCOM

Public Class clsUtilities

    Private strThousSep As String = ","
    Private strDecSep As String = "."
    Private intQtyDec As Integer = 3
    Private FormNum As Integer
    Public strSFilePath As String = String.Empty
    Public strDFilePath As String = String.Empty
    Private strFilepath As String = String.Empty
    Private strFileName As String = String.Empty

    Public Sub New()
        MyBase.New()
        FormNum = 1
    End Sub

#Region "Connect to Company"
    Public Sub Connect()
        Dim strCookie As String
        Dim strConnectionContext As String

        Try
            strCookie = oApplication.Company.GetContextCookie
            strConnectionContext = oApplication.SBO_Application.Company.GetConnectionContext(strCookie)

            If oApplication.Company.SetSboLoginContext(strConnectionContext) <> 0 Then
                Throw New Exception("Wrong login credentials.")
            End If

            'Open a connection to company
            If oApplication.Company.Connect() <> 0 Then
                Throw New Exception("Cannot connect to company database. ")
            End If

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub
#End Region

#Region "Genral Functions"

    Public Sub assignLineNo(ByVal aGrid As SAPbouiCOM.Grid, ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
                aGrid.RowHeaders.SetText(intRow, intRow + 1)
            Next
            aGrid.Columns.Item("RowsHeader").TitleObject.Caption = "#"
            aform.Freeze(False)
        Catch ex As Exception
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)

        End Try
    End Sub

#Region "Get MaxCode"
    Public Function getMaxCode(ByVal sTable As String, ByVal sColumn As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(CAST(" & sColumn & " AS Numeric)) FROM [" & sTable & "]"
            ExecuteSQL(oRS, strSQL)

            If Convert.ToString(oRS.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRS.Fields.Item(0).Value + 1
            Else
                MaxCode = 1
            End If

            sCode = Format(MaxCode, "00000000")
            Return sCode
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        Finally
            oRS = Nothing
        End Try
    End Function
#End Region

#Region "Status Message"
    Public Sub Message(ByVal sMessage As String, ByVal StatusType As SAPbouiCOM.BoStatusBarMessageType)
        oApplication.SBO_Application.StatusBar.SetText(sMessage, SAPbouiCOM.BoMessageTime.bmt_Short, StatusType)
    End Sub
#End Region

#Region "Add Choose from List"
    Public Sub AddChooseFromList(ByVal FormUID As String, ByVal CFL_Text As String, ByVal CFL_Button As String, _
                                        ByVal ObjectType As SAPbouiCOM.BoLinkedObject, _
                                            Optional ByVal AliasName As String = "", Optional ByVal CondVal As String = "", _
                                                    Optional ByVal Operation As SAPbouiCOM.BoConditionOperation = SAPbouiCOM.BoConditionOperation.co_EQUAL)

        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Try
            oCFLs = oApplication.SBO_Application.Forms.Item(FormUID).ChooseFromLists
            oCFLCreationParams = oApplication.SBO_Application.CreateObject( _
                                    SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            If ObjectType = SAPbouiCOM.BoLinkedObject.lf_Items Then
                oCFLCreationParams.MultiSelection = True
            Else
                oCFLCreationParams.MultiSelection = False
            End If

            oCFLCreationParams.ObjectType = ObjectType
            oCFLCreationParams.UniqueID = CFL_Text

            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1

            oCons = oCFL.GetConditions()

            If Not AliasName = "" Then
                oCon = oCons.Add()
                oCon.Alias = AliasName
                oCon.Operation = Operation
                oCon.CondVal = CondVal
                oCFL.SetConditions(oCons)
            End If

            oCFLCreationParams.UniqueID = CFL_Button
            oCFL = oCFLs.Add(oCFLCreationParams)

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub
#End Region

#Region "Get Linked Object Type"
    Public Function getLinkedObjectType(ByVal Type As SAPbouiCOM.BoLinkedObject) As String
        Return CType(Type, String)
    End Function

#End Region

#Region "Execute Query"
    Public Sub ExecuteSQL(ByRef oRecordSet As SAPbobsCOM.Recordset, ByVal SQL As String)
        Try
            If oRecordSet Is Nothing Then
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            End If

            oRecordSet.DoQuery(SQL)

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub
#End Region

#Region "Get Application path"
    Public Function getApplicationPath() As String

        Return Application.StartupPath.Trim

        'Return IO.Directory.GetParent(Application.StartupPath).ToString
    End Function

    Public Function getUserTempPath() As String

        Return System.IO.Path.GetTempPath()

        'Return IO.Directory.GetParent(Application.StartupPath).ToString
    End Function
#End Region

#Region "Date Manipulation"

#Region "Convert SBO Date to System Date"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	ConvertStrToDate
    'Parameter          	:   ByVal oDate As String, ByVal strFormat As String
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	07/12/05
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To convert Date according to current culture info
    '********************************************************************
    Public Function ConvertStrToDate(ByVal strDate As String, ByVal strFormat As String) As DateTime
        Try
            Dim oDate As DateTime
            Dim ci As New System.Globalization.CultureInfo("en-GB", False)
            Dim newCi As System.Globalization.CultureInfo = CType(ci.Clone(), System.Globalization.CultureInfo)

            System.Threading.Thread.CurrentThread.CurrentCulture = newCi
            oDate = oDate.ParseExact(strDate, strFormat, ci.DateTimeFormat)

            Return oDate
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try

    End Function
#End Region

#Region " Get SBO Date Format in String (ddmmyyyy)"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	StrSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(ddmmyy value) as applicable to SBO
    '********************************************************************
    Public Function StrSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String, GetDateFormat As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yy"
                Case 1
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yyyy"
                Case 2
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yy"
                Case 3
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yyyy"
                Case 4
                    GetDateFormat = "yyyy" & DateSep & "dd" & DateSep & "MM"
                Case 5
                    GetDateFormat = "dd" & DateSep & "MMM" & DateSep & "yyyy"
            End Select
            Return GetDateFormat

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function
#End Region

#Region "Get SBO date Format in Number"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	IntSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(integer value) as applicable to SBO
    '********************************************************************
    Public Function NumSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    NumSBODateFormat = 3
                Case 1
                    NumSBODateFormat = 103
                Case 2
                    NumSBODateFormat = 1
                Case 3
                    NumSBODateFormat = 120
                Case 4
                    NumSBODateFormat = 126
                Case 5
                    NumSBODateFormat = 130
            End Select
            Return NumSBODateFormat

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function
#End Region

#End Region

#Region "Get Rental Period"
    Public Function getRentalDays(ByVal Date1 As String, ByVal Date2 As String, ByVal IsWeekDaysBilling As Boolean) As Integer
        Dim TotalDays, TotalDaysincSat, TotalBillableDays As Integer
        Dim TotalWeekEnds As Integer
        Dim StartDate As Date
        Dim EndDate As Date
        Dim oRecordset As SAPbobsCOM.Recordset

        StartDate = CType(Date1.Insert(4, "/").Insert(7, "/"), Date)
        EndDate = CType(Date2.Insert(4, "/").Insert(7, "/"), Date)

        TotalDays = DateDiff(DateInterval.Day, StartDate, EndDate)

        If IsWeekDaysBilling Then
            strSQL = " select dbo.WeekDays('" & Date1 & "','" & Date2 & "')"
            oApplication.Utilities.ExecuteSQL(oRecordset, strSQL)
            If oRecordset.RecordCount > 0 Then
                TotalBillableDays = oRecordset.Fields.Item(0).Value
            End If
            Return TotalBillableDays
        Else
            Return TotalDays + 1
        End If

    End Function

    Public Function WorkDays(ByVal dtBegin As Date, ByVal dtEnd As Date) As Long
        Try
            Dim dtFirstSunday As Date
            Dim dtLastSaturday As Date
            Dim lngWorkDays As Long

            ' get first sunday in range
            dtFirstSunday = dtBegin.AddDays((8 - Weekday(dtBegin)) Mod 7)

            ' get last saturday in range
            dtLastSaturday = dtEnd.AddDays(-(Weekday(dtEnd) Mod 7))

            ' get work days between first sunday and last saturday
            lngWorkDays = (((DateDiff(DateInterval.Day, dtFirstSunday, dtLastSaturday)) + 1) / 7) * 5

            ' if first sunday is not begin date
            If dtFirstSunday <> dtBegin Then

                ' assume first sunday is after begin date
                ' add workdays from begin date to first sunday
                lngWorkDays = lngWorkDays + (7 - Weekday(dtBegin))

            End If

            ' if last saturday is not end date
            If dtLastSaturday <> dtEnd Then

                ' assume last saturday is before end date
                ' add workdays from last saturday to end date
                lngWorkDays = lngWorkDays + (Weekday(dtEnd) - 1)

            End If

            WorkDays = lngWorkDays
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            MsgBox(ex.Message)
        End Try


    End Function

#End Region

#Region "Get Item Price with Factor"
    Public Function getPrcWithFactor(ByVal CardCode As String, ByVal ItemCode As String, ByVal RntlDays As Integer, ByVal Qty As Double) As Double
        Dim oItem As SAPbobsCOM.Items
        Dim Price, Expressn As Double
        Dim oDataSet, oRecSet As SAPbobsCOM.Recordset

        oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        oApplication.Utilities.ExecuteSQL(oDataSet, "Select U_RentFac, U_NumDys From [@REN_FACT] order by U_NumDys ")
        If oItem.GetByKey(ItemCode) And oDataSet.RecordCount > 0 Then

            oApplication.Utilities.ExecuteSQL(oRecSet, "Select ListNum from OCRD where CardCode = '" & CardCode & "'")
            oItem.PriceList.SetCurrentLine(oRecSet.Fields.Item(0).Value - 1)
            Price = oItem.PriceList.Price
            Expressn = 0
            oDataSet.MoveFirst()

            While RntlDays > 0

                If oDataSet.EoF Then
                    oDataSet.MoveLast()
                End If

                If RntlDays < oDataSet.Fields.Item(1).Value Then
                    Expressn += (oDataSet.Fields.Item(0).Value * RntlDays * Price * Qty)
                    RntlDays = 0
                    Exit While
                End If
                Expressn += (oDataSet.Fields.Item(0).Value * oDataSet.Fields.Item(1).Value * Price * Qty)
                RntlDays -= oDataSet.Fields.Item(1).Value
                oDataSet.MoveNext()

            End While

        End If
        If oItem.UserFields.Fields.Item("U_Rental").Value = "Y" Then
            Return CDbl(Expressn / Qty)
        Else
            Return Price
        End If


    End Function
#End Region

#Region "Get WareHouse List"
    Public Function getUsedWareHousesList(ByVal ItemCode As String, ByVal Quantity As Double) As DataTable
        Dim oDataTable As DataTable
        Dim oRow As DataRow
        Dim rswhs As SAPbobsCOM.Recordset
        Dim LeftQty As Double
        Try
            oDataTable = New DataTable
            oDataTable.Columns.Add(New System.Data.DataColumn("ItemCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("WhsCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("Quantity"))

            strSQL = "Select WhsCode, ItemCode, (OnHand + OnOrder - IsCommited) As Available From OITW Where ItemCode = '" & ItemCode & "' And " & _
                        "WhsCode Not In (Select Whscode From OWHS Where U_Reserved = 'Y' Or U_Rental = 'Y') Order By (OnHand + OnOrder - IsCommited) Desc "

            ExecuteSQL(rswhs, strSQL)
            LeftQty = Quantity

            While Not rswhs.EoF
                oRow = oDataTable.NewRow()

                oRow.Item("WhsCode") = rswhs.Fields.Item("WhsCode").Value
                oRow.Item("ItemCode") = rswhs.Fields.Item("ItemCode").Value

                LeftQty = LeftQty - CType(rswhs.Fields.Item("Available").Value, Double)

                If LeftQty <= 0 Then
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double) + LeftQty
                    oDataTable.Rows.Add(oRow)
                    Exit While
                Else
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double)
                End If

                oDataTable.Rows.Add(oRow)
                rswhs.MoveNext()
                oRow = Nothing
            End While

            'strSQL = ""
            'For count As Integer = 0 To oDataTable.Rows.Count - 1
            '    strSQL += oDataTable.Rows(count).Item("WhsCode") & " : " & oDataTable.Rows(count).Item("Quantity") & vbNewLine
            'Next
            'MessageBox.Show(strSQL)

            Return oDataTable

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        Finally
            oRow = Nothing
        End Try
    End Function
#End Region

#End Region

#Region "Functions related to Load XML"

#Region "Add/Remove Menus "
    Public Sub AddRemoveMenus(ByVal sFileName As String)
        Dim oXMLDoc As New Xml.XmlDocument
        Dim sFilePath As String
        Try
            sFilePath = getApplicationPath() & "\XML Files\" & sFileName
            oXMLDoc.Load(sFilePath)
            oApplication.SBO_Application.LoadBatchActions(oXMLDoc.InnerXml)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        Finally
            oXMLDoc = Nothing
        End Try
    End Sub
#End Region

#Region "Load XML File "
    Private Function LoadXMLFiles(ByVal sFileName As String) As String
        Dim oXmlDoc As Xml.XmlDocument
        Dim oXNode As Xml.XmlNode
        Dim oAttr As Xml.XmlAttribute
        Dim sPath As String
        Dim FrmUID As String
        Try
            oXmlDoc = New Xml.XmlDocument

            sPath = getApplicationPath() & "\XML Files\" & sFileName

            oXmlDoc.Load(sPath)
            oXNode = oXmlDoc.GetElementsByTagName("form").Item(0)
            oAttr = oXNode.Attributes.GetNamedItem("uid")
            oAttr.Value = oAttr.Value & FormNum
            FormNum = FormNum + 1
            oApplication.SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
            FrmUID = oAttr.Value

            Return FrmUID

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        Finally
            oXmlDoc = Nothing
        End Try
    End Function
#End Region

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String) As SAPbouiCOM.Form
        'Return LoadForm(XMLFile, FormType.ToString(), FormType & "_" & oApplication.SBO_Application.Forms.Count.ToString)
        LoadXMLFiles(XMLFile)
        Return Nothing
    End Function

    Public Function LoadForm1(ByVal XMLFile As String, ByVal FormType As String) As String
        'Return LoadForm(XMLFile, FormType.ToString(), FormType & "_" & oApplication.SBO_Application.Forms.Count.ToString)
        Dim strFormID As String = LoadXMLFiles(XMLFile)
        Return strFormID
    End Function

    Public Function LoadMessageForm(ByVal XMLFile As String, ByVal FormType As String) As SAPbouiCOM.Form
        LoadXMLFiles(XMLFile)
        Return Nothing
    End Function

    '*****************************************************************
    'Type               : Function   
    'Name               : LoadForm
    'Parameter          : XmlFile,FormType,FormUID
    'Return Value       : SBO Form
    'Author             : Senthil Kumar B Senthil Kumar B
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Load XML file 
    '*****************************************************************

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String, ByVal FormUID As String) As SAPbouiCOM.Form

        Dim oXML As System.Xml.XmlDocument
        Dim objFormCreationParams As SAPbouiCOM.FormCreationParams
        Try
            oXML = New System.Xml.XmlDocument
            oXML.Load(XMLFile)
            objFormCreationParams = (oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams))
            objFormCreationParams.XmlData = oXML.InnerXml
            objFormCreationParams.FormType = FormType
            objFormCreationParams.UniqueID = FormUID
            Return oApplication.SBO_Application.Forms.AddEx(objFormCreationParams)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)

        End Try

    End Function



#Region "Load Forms"
    Public Sub LoadForm(ByRef oObject As Object, ByVal XmlFile As String)
        Try
            oObject.FrmUID = LoadXMLFiles(XmlFile)
            oObject.Form = oApplication.SBO_Application.Forms.Item(oObject.FrmUID)
            If Not oApplication.Collection.ContainsKey(oObject.FrmUID) Then
                oApplication.Collection.Add(oObject.FrmUID, oObject)
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub
#End Region

#End Region

#Region "Functions related to System Initilization"

#Region "Create Tables"
    Public Sub CreateTables()
        Dim oCreateTable As clsTable
        Try
            oCreateTable = New clsTable
            oCreateTable.CreateTables()
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        Finally
            oCreateTable = Nothing
        End Try
    End Sub
#End Region

#Region "Notify Alert"
    Public Sub NotifyAlert()
        'Dim oAlert As clsPromptAlert

        'Try
        '    oAlert = New clsPromptAlert
        '    oAlert.AlertforEndingOrdr()
        'Catch ex As Exception 
        'oApplication.Log.Trace_DIET_AddOn_Error(ex)
        '    Throw ex 
        ''oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        'Finally
        '    oAlert = Nothing
        'End Try

    End Sub
#End Region

    Public Sub setEditText(ByVal aForm As SAPbouiCOM.Form, ByVal aItem As String, ByVal aValue As String)
        Dim oEdit As SAPbouiCOM.EditText
        oEdit = aForm.Items.Item(aItem).Specific
        oEdit.String = aValue
    End Sub

#End Region

#Region "Function related to Quantities"

#Region "Get Available Quantity"
    Public Function getAvailableQty(ByVal ItemCode As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset

        strSQL = "Select SUM(T1.OnHand + T1.OnOrder - T1.IsCommited) From OITW T1 Left Outer Join OWHS T3 On T3.Whscode = T1.WhsCode " & _
                    "Where T1.ItemCode = '" & ItemCode & "'"
        Me.ExecuteSQL(rsQuantity, strSQL)

        If rsQuantity.Fields.Item(0) Is System.DBNull.Value Then
            Return 0
        Else
            Return CLng(rsQuantity.Fields.Item(0).Value)
        End If

    End Function
#End Region

#Region "Get Rented Quantity"
    Public Function getRentedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim RentedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_RDR1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_ORDR] Where U_Status = 'R') " & _
                    " and '" & StartDate & "' between [@REN_RDR1].U_ShipDt1 and [@REN_RDR1].U_ShipDt2 "
        '" and [@REN_RDR1].U_ShipDt1 between '" & StartDate & "' and '" & EndDate & "'"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            RentedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return RentedQty

    End Function
#End Region

#Region "Get Reserved Quantity"
    Public Function getReservedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim ReservedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_QUT1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_OQUT] Where U_Status = 'R' And Status = 'O') " & _
                    " and '" & StartDate & "' between [@REN_QUT1].U_ShipDt1 and [@REN_QUT1].U_ShipDt2"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            ReservedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return ReservedQty

    End Function
#End Region

#End Region

#Region "Functions related to Tax"

#Region "Get Tax Codes"
    Public Sub getTaxCodes(ByRef oCombo As SAPbouiCOM.ComboBox)
        Dim rsTaxCodes As SAPbobsCOM.Recordset

        strSQL = "Select Code, Name From OVTG Where Category = 'O' Order By Name"
        Me.ExecuteSQL(rsTaxCodes, strSQL)

        oCombo.ValidValues.Add("", "")
        If rsTaxCodes.RecordCount > 0 Then
            While Not rsTaxCodes.EoF
                oCombo.ValidValues.Add(rsTaxCodes.Fields.Item(0).Value, rsTaxCodes.Fields.Item(1).Value)
                rsTaxCodes.MoveNext()
            End While
        End If
        oCombo.ValidValues.Add("Define New", "Define New")
        'oCombo.Select("")
    End Sub
#End Region

#Region "Get Applicable Code"

    Public Function getApplicableTaxCode1(ByVal CardCode As String, ByVal ItemCode As String, ByVal Shipto As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    strSQL = "select LicTradNum from CRD1 where Address ='" & Shipto & "' and CardCode ='" & CardCode & "'"
                    Me.ExecuteSQL(rsExempt, strSQL)
                    If rsExempt.RecordCount > 0 Then
                        rsExempt.MoveFirst()
                        TaxGroup = rsExempt.Fields.Item(0).Value
                    Else
                        TaxGroup = ""
                    End If
                    'TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If




        Return TaxGroup

    End Function


    Public Function getApplicableTaxCode(ByVal CardCode As String, ByVal ItemCode As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If

        'If oBP.GetByKey(CardCode.Trim) Then
        '    If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
        '        If oBP.VatGroup.Trim <> "" Then
        '            TaxGroup = oBP.VatGroup.Trim
        '        Else
        '            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        '            If oItem.GetByKey(ItemCode.Trim) Then
        '                TaxGroup = oItem.SalesVATGroup.Trim
        '            End If
        '        End If
        '    ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
        '        strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
        '        Me.ExecuteSQL(rsExempt, strSQL)
        '        If rsExempt.RecordCount > 0 Then
        '            rsExempt.MoveFirst()
        '            TaxGroup = rsExempt.Fields.Item(0).Value
        '        Else
        '            TaxGroup = ""
        '        End If
        '    End If
        'End If
        Return TaxGroup

    End Function
#End Region

#End Region

#Region "Log Transaction"
    Public Sub LogTransaction(ByVal DocNum As Integer, ByVal ItemCode As String, _
                                    ByVal FromWhs As String, ByVal TransferedQty As Double, ByVal ProcessDate As Date)
        Dim sCode As String
        Dim sColumns As String
        Dim sValues As String
        Dim rsInsert As SAPbobsCOM.Recordset

        sCode = Me.getMaxCode("@REN_PORDR", "Code")

        sColumns = "Code, Name, U_DocNum, U_WhsCode, U_ItemCode, U_Quantity, U_RetQty, U_Date"
        sValues = "'" & sCode & "','" & sCode & "'," & DocNum & ",'" & FromWhs & "','" & ItemCode & "'," & TransferedQty & ", 0, Convert(DateTime,'" & ProcessDate.ToString("yyyyMMdd") & "')"

        strSQL = "Insert into [@REN_PORDR] (" & sColumns & ") Values (" & sValues & ")"
        oApplication.Utilities.ExecuteSQL(rsInsert, strSQL)

    End Sub

    Public Sub LogCreatedDocument(ByVal DocNum As Integer, ByVal CreatedDocType As SAPbouiCOM.BoLinkedObject, ByVal CreatedDocNum As String, ByVal sCreatedDate As String)
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim sCode As String
        Dim CreatedDate As DateTime
        Try
            oUserTable = oApplication.Company.UserTables.Item("REN_DORDR")

            sCode = Me.getMaxCode("@REN_DORDR", "Code")

            If Not oUserTable.GetByKey(sCode) Then
                oUserTable.Code = sCode
                oUserTable.Name = sCode

                With oUserTable.UserFields.Fields
                    .Item("U_DocNum").Value = DocNum
                    .Item("U_DocType").Value = CInt(CreatedDocType)
                    .Item("U_DocEntry").Value = CInt(CreatedDocNum)

                    If sCreatedDate <> "" Then
                        CreatedDate = CDate(sCreatedDate.Insert(4, "/").Insert(7, "/"))
                        .Item("U_Date").Value = CreatedDate
                    Else
                        .Item("U_Date").Value = CDate(Format(Now, "Long Date"))
                    End If

                End With

                If oUserTable.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        Finally
            oUserTable = Nothing
        End Try
    End Sub
#End Region

    Public Function getLocalCurrency(ByVal strCurrency As String) As Double
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select Maincurrncy from OADM")
        Return oTemp.Fields.Item(0).Value
    End Function

    Public Function getRecordSetValue(ByVal strQuery As String, strColumn As String) As Double
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery(strQuery)
        Return oTemp.Fields.Item(strColumn).Value
    End Function

    Public Function getRecordSetValueString(ByVal strQuery As String, strColumn As String) As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery(strQuery)
        Return oTemp.Fields.Item(strColumn).Value
    End Function

    Public Function getRecordSetValueString_Series(ByVal strQuery As String, strColumn As String) As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery(strQuery)
        If Not oTemp.EoF Then
            If oTemp.RecordCount > 1 Then
                Throw New Exception("Customer Series returns more than one records...")
            Else
                Return oTemp.Fields.Item(strColumn).Value
            End If
        End If
    End Function

#Region "Get ExchangeRate"
    Public Function getExchangeRate(ByVal strCurrency As String) As Double
        Dim oTemp As SAPbobsCOM.Recordset
        Dim dblExchange As Double
        If GetCurrency("Local") = strCurrency Then
            dblExchange = 1
        Else
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery("Select isNull(Rate,0) from ORTT where convert(nvarchar(10),RateDate,101)=Convert(nvarchar(10),getdate(),101) and currency='" & strCurrency & "'")
            dblExchange = oTemp.Fields.Item(0).Value
        End If
        Return dblExchange
    End Function

    Public Function getExchangeRate(ByVal strCurrency As String, ByVal dtdate As Date) As Double
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strSql As String
        Dim dblExchange As Double
        If GetCurrency("Local") = strCurrency Then
            dblExchange = 1
        Else
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSql = "Select isNull(Rate,0) from ORTT where ratedate='" & dtdate.ToString("yyyy-MM-dd") & "' and currency='" & strCurrency & "'"
            oTemp.DoQuery(strSql)
            dblExchange = oTemp.Fields.Item(0).Value
        End If
        Return dblExchange
    End Function
#End Region

    Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
    End Function

    'Public Function GetCustItemPrice(ByVal strCardCode As String, strItemCode As String, ByVal strDocDate As Date) As Double
    '    Dim objBridge As SAPbobsCOM.SBObob
    '    Dim strItemPrice As String = String.Empty
    '    objBridge = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
    '    Dim oRecordSet As Recordset = objBridge.GetItemPrice(strCardCode, strItemCode, 1, System.DateTime.Now)
    '    If oRecordSet.RecordCount > 0 Then
    '        strItemPrice = oRecordSet.Fields.Item(0).Value
    '    End If
    '    Return CDbl(IIf(strItemPrice = "", "0", strItemPrice))
    'End Function

    Public Function GetCustItemPrice(ByVal strCardCode As String, strItemCode As String, ByVal strDocDate As Date,
                                     ByRef dblItemPrice As Double, ByRef strCurrency As String) As Double
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Dim oRecordSet As Recordset = objBridge.GetItemPrice(strCardCode, strItemCode, 1, System.DateTime.Now)
        If oRecordSet.RecordCount > 0 Then
            dblItemPrice = CDbl(oRecordSet.Fields.Item(0).Value)
            strCurrency = oRecordSet.Fields.Item(1).Value
        Else
            dblItemPrice = 0
            strCurrency = ""
        End If

    End Function

#Region "Get DocCurrency"
    Public Function GetDocCurrency(ByVal aDocEntry As Integer) As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select DocCur from OINV where docentry=" & aDocEntry)
        Return oTemp.Fields.Item(0).Value
    End Function
#End Region

#Region "GetEditTextValues"
    Public Function getEditTextvalue(ByVal aForm As SAPbouiCOM.Form, ByVal strUID As String) As String
        Dim oEditText As SAPbouiCOM.EditText
        oEditText = aForm.Items.Item(strUID).Specific
        Return oEditText.Value
    End Function
#End Region

#Region "Get Currency"
    Public Function GetCurrency(ByVal strChoice As String, Optional ByVal aCardCode As String = "") As String
        Dim strCurrQuery, Currency As String
        Dim oTempCurrency As SAPbobsCOM.Recordset
        If strChoice = "Local" Then
            strCurrQuery = "Select MainCurncy from OADM"
        Else
            strCurrQuery = "Select Currency from OCRD where CardCode='" & aCardCode & "'"
        End If
        oTempCurrency = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempCurrency.DoQuery(strCurrQuery)
        Currency = oTempCurrency.Fields.Item(0).Value
        Return Currency
    End Function

#End Region

    Public Function FormatDataSourceValue(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If

            If Value.IndexOf(CompanyThousandSeprator) > -1 Then
                Value = Value.Replace(CompanyThousandSeprator, "")
            End If
        Else
            Value = "0"

        End If

        ' NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue


        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue
    End Function

    Public Function FormatScreenValues(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If
        Else
            Value = "0"
        End If

        'NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue

        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue

    End Function

    Public Function SetScreenValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function

    Public Function SetDBValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function

#Region "AddControls"
    Public Sub AddControls(ByVal objForm As SAPbouiCOM.Form, ByVal ItemUID As String, ByVal SourceUID As String, ByVal ItemType As SAPbouiCOM.BoFormItemTypes, ByVal position As String, Optional ByVal fromPane As Integer = 1, Optional ByVal toPane As Integer = 1, Optional ByVal linkedUID As String = "", Optional ByVal strCaption As String = "", Optional ByVal dblWidth As Double = 0, Optional ByVal dblTop As Double = 0, Optional ByVal Hight As Double = 0, Optional ByVal Enable As Boolean = True)
        Dim objNewItem, objOldItem As SAPbouiCOM.Item
        Dim ostatic As SAPbouiCOM.StaticText
        Dim oButton As SAPbouiCOM.Button
        Dim oCheckbox As SAPbouiCOM.CheckBox
        Dim oEditText As SAPbouiCOM.EditText
        Dim ofolder As SAPbouiCOM.Folder
        objOldItem = objForm.Items.Item(SourceUID)
        objNewItem = objForm.Items.Add(ItemUID, ItemType)
        With objNewItem
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON Then
                .Left = objOldItem.Left - 15
                .Top = objOldItem.Top + 1
                .LinkTo = linkedUID
            Else
                If position.ToUpper = "RIGHT" Then
                    .Left = objOldItem.Left + objOldItem.Width + 5
                    .Top = objOldItem.Top
                ElseIf position.ToUpper = "DOWN" Then
                    If ItemUID = "edWork" Then
                        .Left = objOldItem.Left + 40
                    Else
                        .Left = objOldItem.Left
                    End If
                    .Top = objOldItem.Top + objOldItem.Height + 3

                    .Width = objOldItem.Width
                    .Height = objOldItem.Height
                ElseIf position.ToUpper = "COPY" Then
                    .Top = objOldItem.Top
                    .Left = objOldItem.Left
                    .Height = objOldItem.Height
                    .Width = objOldItem.Width
                End If
            End If
            .FromPane = fromPane
            .ToPane = toPane
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
                .LinkTo = linkedUID
            End If
            .LinkTo = linkedUID
        End With
        If (ItemType = SAPbouiCOM.BoFormItemTypes.it_EDIT Or ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC) Then
            objNewItem.Width = objOldItem.Width
        End If
        If ItemType = SAPbouiCOM.BoFormItemTypes.it_BUTTON Then
            objNewItem.Width = objOldItem.Width '+ 50
            oButton = objNewItem.Specific
            oButton.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_FOLDER Then
            ofolder = objNewItem.Specific
            ofolder.Caption = strCaption
            ofolder.GroupWith(linkedUID)
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
            ostatic = objNewItem.Specific
            ostatic.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX Then
            oCheckbox = objNewItem.Specific
            oCheckbox.Caption = strCaption

        End If
        If dblWidth <> 0 Then
            objNewItem.Width = dblWidth
        End If

        If dblTop <> 0 Then
            objNewItem.Top = objNewItem.Top + dblTop
        End If
        If Hight <> 0 Then
            objNewItem.Height = objNewItem.Height + Hight
        End If
    End Sub
#End Region

#Region "Set / Get Values from Matrix"
    Public Function getMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer) As String
        Return aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value
    End Function
    Public Sub SetMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer, ByVal strvalue As String)
        aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value = strvalue
    End Sub
#End Region

#Region "Add Condition CFL"
    Public Sub AddConditionCFL(ByVal FormUID As String, ByVal strQuery As String, ByVal strQueryField As String, ByVal sCFL As String)
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim Conditions As SAPbouiCOM.Conditions
        Dim oCond As SAPbouiCOM.Condition
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Dim sDocEntry As New ArrayList()
        Dim sDocNum As ArrayList
        Dim MatrixItem As ArrayList
        sDocEntry = New ArrayList()
        sDocNum = New ArrayList()
        MatrixItem = New ArrayList()

        Try
            oCFLs = oApplication.SBO_Application.Forms.Item(FormUID).ChooseFromLists
            oCFLCreationParams = oApplication.SBO_Application.CreateObject( _
                                    SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            oCFL = oCFLs.Item(sCFL)

            Dim oRec As SAPbobsCOM.Recordset
            oRec = DirectCast(oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRec.DoQuery(strQuery)
            oRec.MoveFirst()

            Try
                If oRec.EoF Then
                    sDocEntry.Add("")
                Else
                    While Not oRec.EoF
                        Dim DocNum As String = oRec.Fields.Item(strQueryField).Value.ToString()
                        If DocNum <> "" Then
                            sDocEntry.Add(DocNum)
                        End If
                        oRec.MoveNext()
                    End While
                End If
            Catch generatedExceptionName As Exception
                Throw
            End Try

            'If IsMatrixCondition = True Then
            '    Dim oMatrix As SAPbouiCOM.Matrix
            '    oMatrix = DirectCast(oForm.Items.Item(Matrixname).Specific, SAPbouiCOM.Matrix)

            '    For a As Integer = 1 To oMatrix.RowCount
            '        If a <> pVal.Row Then
            '            MatrixItem.Add(DirectCast(oMatrix.Columns.Item(columnname).Cells.Item(a).Specific, SAPbouiCOM.EditText).Value)
            '        End If
            '    Next
            '    If removelist = True Then
            '        For xx As Integer = 0 To MatrixItem.Count - 1
            '            Dim zz As String = MatrixItem(xx).ToString()
            '            If sDocEntry.Contains(zz) Then
            '                sDocEntry.Remove(zz)
            '            End If
            '        Next
            '    End If
            'End If

            'oCFLs = oForm.ChooseFromLists
            'oCFLCreationParams = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            'If systemMatrix = True Then
            '    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = Nothing
            '    oCFLEvento = DirectCast(pVal, SAPbouiCOM.IChooseFromListEvent)
            '    Dim sCFL_ID As String = Nothing
            '    sCFL_ID = oCFLEvento.ChooseFromListUID
            '    oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
            'Else
            '    oCFL = oForm.ChooseFromLists.Item(sCHUD)
            'End If

            Conditions = New SAPbouiCOM.Conditions()
            oCFL.SetConditions(Conditions)
            Conditions = oCFL.GetConditions()
            oCond = Conditions.Add()
            oCond.BracketOpenNum = 2
            For i As Integer = 0 To sDocEntry.Count - 1
                If i > 0 Then
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCond = Conditions.Add()
                    oCond.BracketOpenNum = 1
                End If

                oCond.[Alias] = strQueryField
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = sDocEntry(i).ToString()
                If i + 1 = sDocEntry.Count Then
                    oCond.BracketCloseNum = 2
                Else
                    oCond.BracketCloseNum = 1
                End If
            Next

            oCFL.SetConditions(Conditions)


        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub
#End Region

#Region "Open Files"
    Public Sub OpenFile(ByVal strPath As String)
        Try
            If File.Exists(strPath) Then
                Dim process As New System.Diagnostics.Process
                Dim filestart As New System.Diagnostics.ProcessStartInfo(strPath)
                filestart.UseShellExecute = True
                filestart.WindowStyle = ProcessWindowStyle.Normal
                process.StartInfo = filestart
                process.Start()
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)

        End Try
    End Sub
#End Region

    Public Function createDTMainAuthorization() As Boolean
        Try
            Dim RetVal As Long
            Dim mUserPermission As SAPbobsCOM.UserPermissionTree
            mUserPermission = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree)
            '//Mandatory field, which is the key of the object.
            '//The partner namespace must be included as a prefix followed by _
            mUserPermission.PermissionID = "DIET"
            '//The Name value that will be displayed in the General Authorization Tree
            mUserPermission.Name = "DIET Addon"
            '//The permission that this object can get
            mUserPermission.Options = SAPbobsCOM.BoUPTOptions.bou_FullReadNone
            '//In case the level is one, there Is no need to set the FatherID parameter.
            '   mUserPermission.Levels = 1
            RetVal = mUserPermission.Add
            If RetVal = 0 Or -2035 Then
                Return True
            Else
                MsgBox(oApplication.Company.GetLastErrorDescription)
                Return False
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)

        End Try
    End Function

    Public Function addChildAuthorization(ByVal aChildID As String, ByVal aChildiDName As String, ByVal aorder As Integer, ByVal aFormType As String, ByVal aParentID As String, ByVal Permission As SAPbobsCOM.BoUPTOptions) As Boolean
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim mUserPermission As SAPbobsCOM.UserPermissionTree
        mUserPermission = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree)

        mUserPermission.PermissionID = aChildID
        mUserPermission.Name = aChildiDName
        mUserPermission.Options = Permission ' SAPbobsCOM.BoUPTOptions.bou_FullReadNone

        '//For level 2 and up you must set the object's father unique ID
        'mUserPermission.Level
        mUserPermission.ParentID = aParentID
        mUserPermission.UserPermissionForms.DisplayOrder = aorder
        '//this object manages forms
        ' If aFormType <> "" Then
        mUserPermission.UserPermissionForms.FormType = aFormType
        ' End If

        RetVal = mUserPermission.Add
        If RetVal = 0 Or RetVal = -2035 Then
            Return True
        Else
            MsgBox(oApplication.Company.GetLastErrorDescription)
            Return False
        End If


    End Function

    Public Sub AuthorizationCreation()
        Try
            addChildAuthorization("DT_Setup", " Setup", 2, "", "DIET", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            addChildAuthorization("DT_Trans", "Transactions", 2, "", "DIET", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            addChildAuthorization("DT_Report", "Report", 2, "", "DIET", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

            'Setup
            addChildAuthorization("Z_OPRM", "Program Type - setup", 3, "frm_Z_OPRM", "DT_Setup", SAPbobsCOM.BoUPTOptions.bou_FullNone)
            addChildAuthorization("Z_ODLK", "Dislike/Allergic - setup", 3, "frm_Z_ODLK", "DT_Setup", SAPbobsCOM.BoUPTOptions.bou_FullNone)
            addChildAuthorization("Z_OCLP", "Calories Plan - setup", 3, "frm_Z_OCLP", "DT_Setup", SAPbobsCOM.BoUPTOptions.bou_FullNone)
            addChildAuthorization("Z_OCAJ", "Calories Adjustments - setup", 3, "frm_Z_OCAJ", "DT_Setup", SAPbobsCOM.BoUPTOptions.bou_FullNone)
            addChildAuthorization("Z_OMST", "Medical Status - setup", 3, "frm_Z_OMST", "DT_Setup", SAPbobsCOM.BoUPTOptions.bou_FullNone)
            addChildAuthorization("Z_OEXD", "Diet Exclude - setup", 3, "frm_Z_OEXD", "DT_Setup", SAPbobsCOM.BoUPTOptions.bou_FullNone)
            addChildAuthorization("Z_OTTI", "Check Up Time - setup", 3, "frm_Z_OTTI", "DT_Setup", SAPbobsCOM.BoUPTOptions.bou_FullNone)
            'Transaction
            addChildAuthorization("Z_OMED", "Menu Definition", 3, "frm_Z_OMED", "DT_Trans", SAPbobsCOM.BoUPTOptions.bou_FullNone)
            addChildAuthorization("Z_OCRG", "New Registration", 3, "frm_Z_OCRG", "DT_Trans", SAPbobsCOM.BoUPTOptions.bou_FullNone)
            addChildAuthorization("Z_OCPR", "Customer Profile", 3, "frm_Z_OCPR", "DT_Trans", SAPbobsCOM.BoUPTOptions.bou_FullNone)
            addChildAuthorization("Z_OCPM", "Customer Program", 3, "frm_Z_OCPM", "DT_Trans", SAPbobsCOM.BoUPTOptions.bou_FullNone)
            addChildAuthorization("Z_OPSL", "Per Sales Order", 3, "frm_Z_OPSL", "DT_Trans", SAPbobsCOM.BoUPTOptions.bou_FullNone)
            addChildAuthorization("Z_OPMT", "Program Transfer", 3, "frm_Z_OPMT", "DT_Trans", SAPbobsCOM.BoUPTOptions.bou_FullNone)
            'Report
            addChildAuthorization("mnu_Z_OCRR", "Event Type - setup", 3, "frm_Z_OCRR", "DT_Report", SAPbobsCOM.BoUPTOptions.bou_FullNone)

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            MsgBox(ex.Message.ToString())
        End Try
    End Sub

    Public Function validateAuthorization(ByVal aUserId As String, ByVal aFormUID As String) As Boolean
        Dim oAuth As SAPbobsCOM.Recordset
        oAuth = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim struserid As String
        '    Return False
        struserid = oApplication.Company.UserName
        oAuth.DoQuery("select * from UPT1 where FormId='" & aFormUID & "'")
        If (oAuth.RecordCount <= 0) Then
            Return True
        Else
            Dim st As String
            st = oAuth.Fields.Item("PermId").Value
            st = "Select * from USR3 where PermId='" & st & "' and UserLink=" & aUserId
            oAuth.DoQuery(st)
            If oAuth.RecordCount > 0 Then
                If oAuth.Fields.Item("Permission").Value = "N" Then
                    Return False
                End If
                Return True
            Else
                Return True
            End If

        End If

        Return True

    End Function

    Public Sub AssignSerialNo(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        For intRow As Integer = 1 To aMatrix.RowCount
            aMatrix.Columns.Item("SlNo").Cells.Item(intRow).Specific.value = intRow
        Next
        aform.Freeze(False)
    End Sub

    Public Sub AssignRowNo(ByVal aMatrix As SAPbouiCOM.Grid, ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        For intRow As Integer = 0 To aMatrix.DataTable.Rows.Count - 1
            aMatrix.RowHeaders.SetText(intRow, intRow + 1)
        Next
        aform.Freeze(False)
    End Sub

#Region "ValidateCode"
    Public Function ValidateCode(ByVal aCode As String, ByVal aModule As String) As Boolean
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strqry As String = ""
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aModule = "Z_OEVT" Then
            strqry = "Select * from ""@Z_OEVT"" where ""U_Code""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Event Type Already Exist...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "Z_OEVL" Then
            strqry = "Select * from ""@Z_OEVL"" where ""U_Code""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Event Level Already Exist...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "Z_OFUS" Then
            strqry = "Select * from ""@Z_OFUS"" where ""U_Code""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Function Space Already Exist...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "Z_OMUT" Then
            strqry = "Select * from ""@Z_OMUT"" where ""U_Code""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Menu Type Already Exist...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        End If
        Return False
    End Function
#End Region

    Public Function AddServiceItemDocument(ByVal oForm As SAPbouiCOM.Form) As String
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGenDataChild As SAPbobsCOM.GeneralData
        Dim oGenDataCollection As SAPbobsCOM.GeneralDataCollection
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim strQuery As String = String.Empty
        Dim strCode As String = String.Empty
        oCompanyService = oApplication.Company.GetCompanyService()
        Try
            oGeneralService = oCompanyService.GetGeneralService("Z_OISI")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGenDataChild = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim intCode As Integer = getMaxCode("@Z_OISI", "DocEntry")
            strCode = String.Format("{0:000000000}", intCode)
            oGeneralData.SetProperty("U_Reference", strCode)
            oGenDataCollection = oGeneralData.Child("Z_ISI1")
            oGeneralService.Add(oGeneralData)
            Return strCode
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
        Return strCode
    End Function

    Public Sub RemoveServiceItemDocument(ByVal oForm As SAPbouiCOM.Form, ByVal strReference As String)
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim oGeneralDataParams As SAPbobsCOM.GeneralDataParams
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim strQuery As String = String.Empty
        Dim strDocEntry As String = String.Empty
        oCompanyService = oApplication.Company.GetCompanyService()
        Try
            oGeneralService = oCompanyService.GetGeneralService("Z_OISI")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGeneralDataParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select DocEntry From [@Z_OISI] Where U_Reference = '" + strReference + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                strDocEntry = oRecordSet.Fields.Item(0).Value
                oGeneralDataParams.SetProperty("DocEntry", strDocEntry)
                oGeneralService.Delete(oGeneralDataParams)
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Public Function AddOrder(ByVal oForm As SAPbouiCOM.Form, ByVal strDocEntry As String) As Boolean
        Dim _retVal As Boolean = False
        Try
            Dim oOrder As SAPbobsCOM.Documents
            Dim oRecordSet_H As SAPbobsCOM.Recordset
            Dim oRecordSet_P As SAPbobsCOM.Recordset
            Dim oRecordSet As SAPbobsCOM.Recordset

            Dim strQuery As String = String.Empty
            Dim intStatus As Integer
            Try
                oOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                oRecordSet_H = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet_P = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                strQuery = "Select U_CardCode,U_CardName,ISNULL(U_IsCon,'N') As  U_IsCon,U_FromDate From [@Z_OPSL] Where DocEntry = '" + strDocEntry + "'"
                oRecordSet_H.DoQuery(strQuery)
                If Not oRecordSet_H.EoF Then
                    Dim strIsCon As String = oRecordSet_H.Fields.Item("U_IsCon").Value

                    strQuery = " Select Distinct Convert(VarChar(8),U_DelDate,112) As 'PrgDate',U_DelDate From [@Z_PSL1] "
                    strQuery += " Where DocEntry = '" + strDocEntry + "'"
                    oRecordSet_P.DoQuery(strQuery)

                    If Not oRecordSet_P.EoF Then
                        While Not oRecordSet_P.EoF

                            oOrder.CardCode = oRecordSet_H.Fields.Item("U_CardCode").Value
                            oOrder.CardName = oRecordSet_H.Fields.Item("U_CardName").Value
                            oOrder.NumAtCard = strDocEntry

                            Dim prgDate As Date = oRecordSet_P.Fields.Item("U_DelDate").Value

                            'oOrder.DocDate = System.DateTime.Now 'CDate(oRecordSet_H.Fields.Item("U_FromDate").Value)
                            'oOrder.TaxDate = System.DateTime.Now
                            'oOrder.DocDueDate = System.DateTime.Now

                            oOrder.DocDate = prgDate 'System.DateTime.Now 'CDate(oRecordSet_H.Fields.Item("U_FromDate").Value)
                            oOrder.TaxDate = prgDate 'System.DateTime.Now
                            oOrder.DocDueDate = prgDate 'System.DateTime.Now

                            oOrder.Comments = "Pre Sales Booking"
                            oOrder.UserFields.Fields.Item("U_PSNo").Value = strDocEntry
                            oOrder.UserFields.Fields.Item("U_IsCon").Value = strIsCon
                            oOrder.UserFields.Fields.Item("U_IsWizard").Value = "Y"
                            If strIsCon = "Y" Then
                                oOrder.UserFields.Fields.Item("U_IsCon").Value = "Y"
                                oOrder.UserFields.Fields.Item("U_ConDate").Value = CDate(oRecordSet_H.Fields.Item("U_FromDate").Value)
                            End If
                            Dim intRow As Integer = 0

                            'strQuery = "  Select T0.DocEntry,T0.LineId,T0.U_ItemCode,T0.U_Quantity,T0.U_DelDate,T0.U_FType,T0.U_Dislike,T0.U_Medical, "
                            'strQuery += " T1.U_Program,T0.U_Remarks,ISNULL(T1.U_IsCon,'N') As  U_IsCon,T1.U_FromDate "
                            'strQuery += " From [@Z_PSL1] T0 JOIN [@Z_OPSL] T1 On T0.DocEntry = T1.DocEntry Where T0.DocEntry = '" + strDocEntry + "' "

                            If strIsCon = "N" Then
                                strQuery = "  Select T0.DocEntry,T0.LineId,T0.U_ItemCode,T0.U_Quantity,T0.U_DelDate,T0.U_FType,T0.U_Dislike,T0.U_Medical, "
                                strQuery += " T1.U_Program,T0.U_Remarks,ISNULL(T1.U_IsCon,'N') As  U_IsCon,T1.U_FromDate, "
                                strQuery += " ISNULL(ISNULL(ISNULL(T3.U_SaleEmp,T4.U_SaleEmp),T5.SlpCode),-1) As 'U_SaleEmp', "
                                strQuery += " ISNULL(ISNULL(ISNULL(T3.U_Address,T4.U_Address),T5.ShipToDef),'') As 'U_Address', "
                                strQuery += " ISNULL(ISNULL(ISNULL(T3.U_Building,T4.U_Building),T5.MailBuildi),'') As 'U_Building' "
                                strQuery += " ,T1.U_ProgramID"
                                strQuery += " From [@Z_PSL1] T0 JOIN [@Z_OPSL] T1 On T0.DocEntry = T1.DocEntry "
                                strQuery += " LEFT OUTER JOIN [@Z_OCPR] T2 On T1.U_CardCode = T2.U_CardCode "
                                strQuery += " LEFT OUTER JOIN [@Z_CPR5] T3 On T2.DocEntry = T3.DocEntry "
                                strQuery += " AND Convert(VarChar(8),T0.U_DelDate,112) Between Convert(VarChar(8),T3.U_DelDate,112) And Convert(VarChar(8),T3.U_TDelDate,112) "
                                strQuery += " And ((T3.U_BF = 'Y' AND T0.U_FType = 'BF') "
                                strQuery += " OR (T3.U_LN = 'Y' AND T0.U_FType = 'LN') "
                                strQuery += " OR (T3.U_LS = 'Y' AND T0.U_FType = 'LS') "
                                strQuery += " OR (T3.U_SK = 'Y' AND T0.U_FType = 'SK') "
                                strQuery += " OR (T3.U_DI = 'Y' AND T0.U_FType = 'DI') "
                                strQuery += " OR (T3.U_DS = 'Y' AND T0.U_FType = 'DS')) "
                                strQuery += " LEFT OUTER JOIN [@Z_CPR6] T4 On T2.DocEntry = T4.DocEntry "
                                strQuery += " And ((T4.U_BF = 'Y' AND T0.U_FType = 'BF') "
                                strQuery += " OR (T4.U_LN = 'Y' AND T0.U_FType = 'LN') "
                                strQuery += " OR (T4.U_LS = 'Y' AND T0.U_FType = 'LS') "
                                strQuery += " OR (T4.U_SK = 'Y' AND T0.U_FType = 'SK') "
                                strQuery += " OR (T4.U_DI = 'Y' AND T0.U_FType = 'DI') "
                                strQuery += " OR (T4.U_DS = 'Y' AND T0.U_FType = 'DS')) "
                                strQuery += " AND T4.U_Day = DatePart(DW,T0.U_DelDate) "
                                strQuery += " JOIN OCRD T5 On T5.CardCode = T1.U_CardCode "
                                strQuery += " Where T0.DocEntry = '" + strDocEntry + "' "
                                strQuery += " And Convert(VarChar(8),T0.U_DelDate,112) = '" & oRecordSet_P.Fields.Item("PrgDate").Value & "'"
                            Else
                                strQuery = "  Select T0.DocEntry,T0.LineId,T0.U_ItemCode,T0.U_Quantity,T0.U_DelDate,T0.U_FType,T0.U_Dislike,T0.U_Medical, "
                                strQuery += " T1.U_Program,T0.U_Remarks,ISNULL(T1.U_IsCon,'N') As  U_IsCon,T1.U_FromDate, "
                                strQuery += " ISNULL(ISNULL(ISNULL(T3.U_SaleEmp,T4.U_SaleEmp),T5.SlpCode),-1) As 'U_SaleEmp', "
                                strQuery += " ISNULL(ISNULL(ISNULL(T3.U_Address,T4.U_Address),T5.ShipToDef),'') As 'U_Address', "
                                strQuery += " ISNULL(ISNULL(ISNULL(T3.U_Building,T4.U_Building),T5.MailBuildi),'') As 'U_Building' "
                                strQuery += " ,T1.U_ProgramID"
                                strQuery += " From [@Z_PSL1] T0 JOIN [@Z_OPSL] T1 On T0.DocEntry = T1.DocEntry "
                                strQuery += " LEFT OUTER JOIN [@Z_OCPR] T2 On T1.U_CardCode = T2.U_CardCode "
                                strQuery += " LEFT OUTER JOIN [@Z_CPR5] T3 On T2.DocEntry = T3.DocEntry "
                                strQuery += " AND Convert(VarChar(8),T1.U_FromDate,112) Between Convert(VarChar(8),T3.U_DelDate,112) And Convert(VarChar(8),T3.U_TDelDate,112) "
                                strQuery += " And ((T3.U_BF = 'Y' AND T0.U_FType = 'BF') "
                                strQuery += " OR (T3.U_LN = 'Y' AND T0.U_FType = 'LN') "
                                strQuery += " OR (T3.U_LS = 'Y' AND T0.U_FType = 'LS') "
                                strQuery += " OR (T3.U_SK = 'Y' AND T0.U_FType = 'SK') "
                                strQuery += " OR (T3.U_DI = 'Y' AND T0.U_FType = 'DI') "
                                strQuery += " OR (T3.U_DS = 'Y' AND T0.U_FType = 'DS')) "
                                strQuery += " LEFT OUTER JOIN [@Z_CPR6] T4 On T2.DocEntry = T4.DocEntry "
                                strQuery += " And ((T4.U_BF = 'Y' AND T0.U_FType = 'BF') "
                                strQuery += " OR (T4.U_LN = 'Y' AND T0.U_FType = 'LN') "
                                strQuery += " OR (T4.U_LS = 'Y' AND T0.U_FType = 'LS') "
                                strQuery += " OR (T4.U_SK = 'Y' AND T0.U_FType = 'SK') "
                                strQuery += " OR (T4.U_DI = 'Y' AND T0.U_FType = 'DI') "
                                strQuery += " OR (T4.U_DS = 'Y' AND T0.U_FType = 'DS')) "
                                strQuery += " AND T4.U_Day = DatePart(DW,T1.U_FromDate) "
                                strQuery += " JOIN OCRD T5 On T5.CardCode = T1.U_CardCode "
                                strQuery += " Where T0.DocEntry = '" + strDocEntry + "' "
                                strQuery += " And Convert(VarChar(8),T0.U_DelDate,112) = '" & oRecordSet_P.Fields.Item("PrgDate").Value & "'"
                            End If

                            oRecordSet.DoQuery(strQuery)
                            If Not oRecordSet.EoF Then
                                While Not oRecordSet.EoF

                                    oOrder.Lines.SetCurrentLine(intRow)

                                    oOrder.Lines.ItemCode = oRecordSet.Fields.Item("U_ItemCode").Value
                                    oOrder.Lines.Quantity = oRecordSet.Fields.Item("U_Quantity").Value
                                    oOrder.Lines.UnitPrice = 0

                                    If oRecordSet.Fields.Item("U_IsCon").Value.ToString() = "Y" Then
                                        oOrder.Lines.ShipDate = CDate(oRecordSet.Fields.Item("U_FromDate").Value)
                                        oOrder.Lines.UserFields.Fields.Item("U_ConDate").Value = CDate(oRecordSet.Fields.Item("U_FromDate").Value)
                                        oOrder.Lines.UserFields.Fields.Item("U_IsCon").Value = "Y"
                                    Else
                                        oOrder.Lines.ShipDate = CDate(oRecordSet.Fields.Item("U_DelDate").Value)
                                    End If

                                    oOrder.Lines.UserFields.Fields.Item("U_DelDate").Value = CDate(oRecordSet.Fields.Item("U_DelDate").Value)
                                    oOrder.Lines.UserFields.Fields.Item("U_PSORef").Value = oRecordSet.Fields.Item("DocEntry").Value.ToString()
                                    oOrder.Lines.UserFields.Fields.Item("U_PSOLine").Value = oRecordSet.Fields.Item("LineId").Value.ToString()
                                    oOrder.Lines.UserFields.Fields.Item("U_FType").Value = oRecordSet.Fields.Item("U_FType").Value.ToString()
                                    oOrder.Lines.UserFields.Fields.Item("U_Dislike").Value = oRecordSet.Fields.Item("U_Dislike").Value.ToString()
                                    oOrder.Lines.UserFields.Fields.Item("U_Medical").Value = oRecordSet.Fields.Item("U_Medical").Value.ToString()
                                    oOrder.Lines.UserFields.Fields.Item("U_Program").Value = oRecordSet.Fields.Item("U_Program").Value.ToString()
                                    oOrder.Lines.FreeText = oRecordSet.Fields.Item("U_Remarks").Value.ToString()

                                    oOrder.Lines.UserFields.Fields.Item("U_Address").Value = oRecordSet.Fields.Item("U_Address").Value.ToString()
                                    oOrder.Lines.UserFields.Fields.Item("U_Building").Value = oRecordSet.Fields.Item("U_Building").Value.ToString()

                                    strQuery = "Select State From CRD1 Where CardCode = '" & oRecordSet_H.Fields.Item("U_CardCode").Value & "' And AdresType = 'S' "
                                    strQuery += " And Address = '" & oRecordSet.Fields.Item("U_Address").Value.ToString() & "'"
                                    Dim strState As String = oApplication.Utilities.getRecordSetValueString(strQuery, "State")
                                    If strState <> "" Then
                                        oOrder.Lines.UserFields.Fields.Item("U_State").Value = strState
                                    End If

                                    oOrder.Lines.SalesPersonCode = CInt(oRecordSet.Fields.Item("U_SaleEmp").Value.ToString())
                                    oOrder.Lines.UserFields.Fields.Item("U_ProgramID").Value = oRecordSet.Fields.Item("U_ProgramID").Value.ToString()

                                    strQuery = "Select U_PaidType From [@Z_CPM6] "
                                    strQuery += " Where DocEntry = '" & oRecordSet.Fields.Item("U_ProgramID").Value.ToString() & "' "
                                    strQuery += " AND '" & CDate(oRecordSet.Fields.Item("U_DelDate").Value).ToString("yyyyMMdd") & "' Between Convert(VarChar(8),U_Fdate,112) And Convert(VarChar(8),U_Edate,112) "
                                    Dim strPayType As String = oApplication.Utilities.getRecordSetValueString(strQuery, "U_PaidType")
                                    oOrder.Lines.UserFields.Fields.Item("U_PaidType").Value = strPayType

                                    oOrder.Lines.Add()
                                    intRow += 1
                                    oRecordSet.MoveNext()

                                End While

                                intStatus = oOrder.Add
                                If intStatus = 0 Then
                                    _retVal = True
                                    Dim strOrder As String = oApplication.Company.GetNewObjectKey()

                                    'Header
                                    strQuery = "Update [@Z_OPSL] Set U_SalesO = ISNULL(U_SalesO,'') + '" & strOrder & ",'"
                                    strQuery += " Where DocEntry = '" + strDocEntry + "'"
                                    oRecordSet.DoQuery(strQuery)

                                    'Rows
                                    'strQuery = "Update [@Z_PSL1] Set U_Status = 'C' Where DocEntry = '" + strDocEntry + "'"
                                    strQuery = "Update [@Z_PSL1] Set U_Status = 'C' "
                                    strQuery += " ,U_SalesO = '" & strOrder & "'"
                                    strQuery += " Where DocEntry = '" & strDocEntry & "'"
                                    strQuery += " And Convert(VarChar(8),U_DelDate,112) = '" & oRecordSet_P.Fields.Item("PrgDate").Value & "'"
                                    oRecordSet.DoQuery(strQuery)
                                Else
                                    _retVal = False
                                    'Throw New Exception(oApplication.Company.GetLastErrorDescription())
                                    oApplication.SBO_Application.MessageBox(oApplication.Company.GetLastErrorDescription(), 1, "OK", "", "")
                                End If
                            End If

                            oRecordSet_P.MoveNext()
                        End While
                    End If

                End If
                Return _retVal
            Catch ex As Exception
                oApplication.Log.Trace_DIET_AddOn_Error(ex)
                _retVal = False
                oApplication.SBO_Application.MessageBox(ex.Message, 1, "OK", "", "")
            End Try
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function AddInvoiceDocument(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Dim _retVal As Boolean = False
        Dim oInvoice As SAPbobsCOM.Documents
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim oRecordSet_C As SAPbobsCOM.Recordset
        Dim oRecordSet_U As SAPbobsCOM.Recordset
        Dim strQuery As String = String.Empty
        Dim intStatus As Integer
        Try
            oInvoice = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet_C = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet_U = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strDocEntry As String = CType(oForm.Items.Item("10").Specific, SAPbouiCOM.EditText).Value

            strQuery = " Select Distinct T1.* From ( "
            strQuery += " Select DISTINCT ISNULL(T0.U_SerRef,'') As 'U_Reference',T0.LineId "
            strQuery += " From [@Z_CPM6] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " Where T1.DocEntry = '" + strDocEntry + "'"
            strQuery += " And ISNULL(T0.U_InvRef,'') = ''  "
            strQuery += " And ISNULL(T0.U_IsIReq,'N') = 'Y'  "
            strQuery += " UNION ALL "
            strQuery += " Select DISTINCT T0.U_Reference,T0.U_Reference As 'LineId'  "
            strQuery += " From [@Z_OISI] T0 JOIN [@Z_ISI1] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN [@Z_CPM6] T2 ON T0.U_Reference = T2.U_SerRef "
            strQuery += " Where T2.DocEntry = '" + strDocEntry + "'"
            strQuery += " And ISNULL(T2.U_InvRef,'') = '' "
            strQuery += " And T1.U_ItemCode <> '' "
            strQuery += "  ) T1 "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF

                    Dim intCurrentLine As Integer = 0
                    oInvoice = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                    oInvoice.CardCode = oForm.Items.Item("6").Specific.value
                    oInvoice.CardName = oForm.Items.Item("7").Specific.value
                    oInvoice.NumAtCard = oForm.Items.Item("9").Specific.value
                    oInvoice.DocDate = System.DateTime.Now
                    oInvoice.TaxDate = System.DateTime.Now
                    oInvoice.DocDueDate = System.DateTime.Now
                    oInvoice.DiscountPercent = CDbl(IIf(oForm.Items.Item("17").Specific.value = "", 0, oForm.Items.Item("17").Specific.value))
                    oInvoice.Comments = "Program Booking"

                    strQuery = " Select T1.U_PrgCode As 'U_ItemCode',T1.U_PrgName As 'U_ItemName',T0.U_NoofDays As 'U_Quantity' "
                    strQuery += ",T0.U_Price,T0.U_Discount,T0.U_LineTotal,T0.U_PaidType "
                    strQuery += ",T0.U_Fdate,T0.U_Edate,'P' As 'Type' "
                    strQuery += " From [@Z_CPM6] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry "
                    If oRecordSet.Fields.Item("U_Reference").Value.ToString() = "" Then
                        strQuery += " Where T0.LineId = '" + oRecordSet.Fields.Item("LineId").Value.ToString() + "'"
                    Else
                        strQuery += " Where T0.U_SerRef = '" + oRecordSet.Fields.Item("U_Reference").Value + "'"
                    End If
                    strQuery += " And T1.DocEntry = '" + strDocEntry + "'"
                    strQuery += " And ISNULL(T0.U_IsIReq,'N') = 'Y'  "
                    strQuery += " UNION ALL "
                    strQuery += " Select T1.U_ItemCode,T1.U_ItemName,T1.U_Quantity,T1.U_Price,T1.U_Discount,T1.U_LineTotal,T2.U_PaidType  "
                    strQuery += ",T2.U_Fdate,T2.U_Edate,'S' As 'Type' "
                    strQuery += " From [@Z_OISI] T0 JOIN [@Z_ISI1] T1 On T0.DocEntry = T1.DocEntry "
                    strQuery += " JOIN [@Z_CPM6] T2 ON T0.U_Reference = T2.U_SerRef "
                    strQuery += " Where T2.U_SerRef = '" + oRecordSet.Fields.Item("U_Reference").Value + "'"
                    strQuery += " And T1.U_ItemCode <> '' "
                    oRecordSet_C.DoQuery(strQuery)
                    If Not oRecordSet_C.EoF Then
                        While Not oRecordSet_C.EoF
                            oInvoice.Lines.SetCurrentLine(intCurrentLine)
                            oInvoice.Lines.ItemCode = oRecordSet_C.Fields.Item("U_ItemCode").Value
                            oInvoice.Lines.Quantity = oRecordSet_C.Fields.Item("U_Quantity").Value
                            oInvoice.Lines.UnitPrice = oRecordSet_C.Fields.Item("U_Price").Value
                            oInvoice.Lines.DiscountPercent = oRecordSet_C.Fields.Item("U_Discount").Value
                            If oRecordSet_C.Fields.Item("Type").Value = "P" Then
                                oInvoice.Lines.UserFields.Fields.Item("U_Fdate").Value = oRecordSet_C.Fields.Item("U_Fdate").Value
                                oInvoice.Lines.UserFields.Fields.Item("U_Edate").Value = oRecordSet_C.Fields.Item("U_Edate").Value
                                oInvoice.Lines.UserFields.Fields.Item("U_PaidType").Value = oRecordSet_C.Fields.Item("U_PaidType").Value
                            End If
                            oInvoice.Lines.Add()
                            intCurrentLine += 1
                            oRecordSet_C.MoveNext()
                        End While
                    End If

                    intStatus = oInvoice.Add
                    If intStatus = 0 Then

                        Dim strInvoice As String = oApplication.Company.GetNewObjectKey()
                        oInvoice.GetByKey(strInvoice)
                        _retVal = True
                        strQuery = "Update [@Z_CPM6] Set U_InvNo = '" + oInvoice.DocNum.ToString() + "'"
                        strQuery += " ,U_InvRef = '" + strInvoice + "'"
                        strQuery += " ,U_InvCreated = 'Y' "

                        If oRecordSet.Fields.Item("U_Reference").Value.ToString() = "" Then
                            strQuery += " Where LineId = '" + oRecordSet.Fields.Item("LineId").Value.ToString() + "'"
                        Else
                            strQuery += " Where U_SerRef = '" + oRecordSet.Fields.Item("U_Reference").Value + "'"
                        End If

                        strQuery += " And ISNULL(U_InvRef,'') = '' "
                        strQuery += " AND DocEntry = '" + strDocEntry + "'"
                        oRecordSet_U.DoQuery(strQuery)

                        strQuery = " Update T1 Set "
                        strQuery += " T1.U_InvNo = '" + oInvoice.DocNum.ToString() + "'"
                        strQuery += " ,T1.U_InvRef = '" + strInvoice + "'"
                        strQuery += " ,U_InvCreated = 'Y' "
                        strQuery += " From [@Z_OISI] T0 JOIN [@Z_ISI1] T1 On T0.DocEntry = T1.DocEntry "
                        strQuery += " Where T0.U_Reference =  '" + oRecordSet.Fields.Item("U_Reference").Value + "'"
                        strQuery += " And ISNULL(T1.U_ItemCode,'') <> '' "
                        oRecordSet_U.DoQuery(strQuery)

                    Else
                        oApplication.SBO_Application.MessageBox(oApplication.Company.GetLastErrorDescription(), 1, "Ok", "", "")

                    End If

                    oRecordSet.MoveNext()

                End While

            End If

            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Sub OpenFileDialogBox(ByVal oForm As SAPbouiCOM.Form, ByVal strPath As String, ByVal strFile As String)
        Dim _retVal As String = String.Empty
        Try
            FileOpen()
            CType(oForm.Items.Item(strPath).Specific, SAPbouiCOM.EditText).Value = strFilepath
            strFileName = Path.GetFileName(strFilepath)
            CType(oForm.Items.Item(strFile).Specific, SAPbouiCOM.EditText).Value = strFileName
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

#Region "FileOpen"
    Private Sub FileOpen()
        Try
            Dim mythr As New System.Threading.Thread(AddressOf ShowFileDialog)
            mythr.SetApartmentState(Threading.ApartmentState.STA)
            mythr.Start()
            mythr.Join()
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Private Sub ShowFileDialog()
        Try
            Dim oDialogBox As New OpenFileDialog
            Dim strMdbFilePath As String
            Dim oProcesses() As Process
            Try
                Dim aform As New System.Windows.Forms.Form
                aform.TopMost = True
                oProcesses = Process.GetProcessesByName("SAP Business One")
                If oProcesses.Length <> 0 Then
                    For i As Integer = 0 To oProcesses.Length - 1
                        Dim MyWindow As New clsListener.WindowWrapper(oProcesses(i).MainWindowHandle)
                        If oDialogBox.ShowDialog(MyWindow) = DialogResult.OK Then
                            strMdbFilePath = oDialogBox.FileName
                            strFilepath = oDialogBox.FileName
                            Exit For
                        Else
                            Exit For
                        End If
                    Next
                End If
            Catch ex As Exception
                oApplication.Log.Trace_DIET_AddOn_Error(ex)
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
            End Try
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub
#End Region

    Public Function getPicturePath()
        Dim _retVal As String = String.Empty
        Try
            Dim oCompanyService As SAPbobsCOM.CompanyService
            oCompanyService = oApplication.Company.GetCompanyService
            _retVal = oCompanyService.GetPathAdmin().PicturesFolderPath
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
        Return _retVal
    End Function

    Public Sub ActivateMenuEvent(oApplication As SAPbouiCOM.Application, ByRef pVal As SAPbouiCOM.MenuEvent, UDOName As String)
        Try
            Dim oMenuItem As SAPbouiCOM.MenuItem
            Dim i As Integer, j As Integer
            oMenuItem = oApplication.Menus.Item("47616")
            j = 47616 + oMenuItem.SubMenus.Count
            For Each var As SAPbouiCOM.MenuItem In oMenuItem.SubMenus
                If var.[String].Contains(UDOName) Then
                    var.Activate()
                End If
            Next
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Public Function CreateCustomer(ByVal oForm As SAPbouiCOM.Form, ByVal oDocEntry As String) As Boolean
        Dim oCustomer As SAPbobsCOM.BusinessPartners = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim _retVal As Boolean = False
        Dim strQuery As String = "Select * From [@Z_OCRG] Where DocEntry = '" + oDocEntry + "'"
        Try

            oCustomer = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then

                oCustomer.CardType = SAPbobsCOM.BoCardTypes.cCustomer

                If oRecordSet.Fields.Item("U_IsAuto").Value.ToString() = "N" Then
                    oCustomer.CardCode = oRecordSet.Fields.Item("U_CardCode").Value
                End If

                oCustomer.Series = oRecordSet.Fields.Item("U_Series").Value
                oCustomer.CardName = oRecordSet.Fields.Item("U_CardName").Value
                oCustomer.Currency = "##"
                oCustomer.GroupCode = oRecordSet.Fields.Item("U_Occup").Value.ToString().Trim()
                oCustomer.Valid = BoYesNoEnum.tYES

                'oCustomer.PriceListNum = oRecordSet.Fields.Item("U_PriceList").Value
                oCustomer.CardForeignName = oRecordSet.Fields.Item("U_CardName").Value
                oCustomer.UserFields.Fields.Item("U_DOB").Value = oRecordSet.Fields.Item("U_DOB").Value
                oCustomer.UserFields.Fields.Item("U_Title").Value = oRecordSet.Fields.Item("U_Title").Value
                oCustomer.UserFields.Fields.Item("U_Occup").Value = oRecordSet.Fields.Item("U_Occup").Value

                oCustomer.Cellular = IIf(oRecordSet.Fields.Item("U_Mobile").Value.ToString() = "", "", oRecordSet.Fields.Item("U_Mobile").Value.ToString())
                oCustomer.Phone1 = IIf(oRecordSet.Fields.Item("U_TeleNo").Value.ToString() = "", "", oRecordSet.Fields.Item("U_TeleNo").Value.ToString())

                'oCustomer.Addresses.AddressName = IIf(oRecordSet.Fields.Item("U_CardName").Value.ToString() = "", "", oRecordSet.Fields.Item("U_CardName").Value.ToString())
                'oCustomer.Addresses.Street = IIf(oRecordSet.Fields.Item("U_Street").Value.ToString() = "", "", oRecordSet.Fields.Item("U_Street").Value.ToString())
                'oCustomer.Addresses.Block = IIf(oRecordSet.Fields.Item("U_Block").Value.ToString() = "", "", oRecordSet.Fields.Item("U_Block").Value.ToString())
                'oCustomer.Addresses.City = IIf(oRecordSet.Fields.Item("U_City").Value.ToString() = "", "", oRecordSet.Fields.Item("U_City").Value.ToString())
                'oCustomer.Addresses.BuildingFloorRoom = IIf(oRecordSet.Fields.Item("U_Address").Value.ToString() = "", "", oRecordSet.Fields.Item("U_Address").Value.ToString())
                'oCustomer.Addresses.Country = "LB"
                'oCustomer.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo
                'oCustomer.Addresses.Add()

                'oCustomer.Addresses.AddressName = IIf(oRecordSet.Fields.Item("U_CardName").Value.ToString() = "", "", oRecordSet.Fields.Item("U_CardName").Value.ToString())
                'oCustomer.Addresses.Street = IIf(oRecordSet.Fields.Item("U_Street").Value.ToString() = "", "", oRecordSet.Fields.Item("U_Street").Value.ToString())
                'oCustomer.Addresses.Block = IIf(oRecordSet.Fields.Item("U_Block").Value.ToString() = "", "", oRecordSet.Fields.Item("U_Block").Value.ToString())
                'oCustomer.Addresses.City = IIf(oRecordSet.Fields.Item("U_City").Value.ToString() = "", "", oRecordSet.Fields.Item("U_City").Value.ToString())
                'oCustomer.Addresses.BuildingFloorRoom = IIf(oRecordSet.Fields.Item("U_Address").Value.ToString() = "", "", oRecordSet.Fields.Item("U_Address").Value.ToString())
                'oCustomer.Addresses.Country = "LB"
                'oCustomer.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo
                'oCustomer.Addresses.Add()

                oCustomer.UserFields.Fields.Item("U_RegNo").Value = oRecordSet.Fields.Item("DocNum").Value.ToString()
                oCustomer.FreeText = oRecordSet.Fields.Item("U_Remarks").Value

                Dim iRes As Integer = oCustomer.Add()
                If iRes = 0 Then
                    _retVal = True
                    Dim strCardCode As String = oApplication.Company.GetNewObjectKey()
                    oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    'Header
                    strQuery = " Update T0 Set T0.U_CardCode = '" + strCardCode + "'"
                    strQuery += " From [@Z_OCRG] T0 "
                    strQuery += " Where T0.DocEntry = '" + oDocEntry + "'"
                    oRecordSet.DoQuery(strQuery)
                Else
                    _retVal = False
                    Throw New Exception(oApplication.Company.GetLastErrorDescription())
                End If
            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oCustomer)
        End Try
    End Function

    Public Function UpdateCustomer(ByVal oForm As SAPbouiCOM.Form, ByVal oDocEntry As String) As Boolean
        Dim oCustomer As SAPbobsCOM.BusinessPartners = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim _retVal As Boolean = False
        Dim strQuery As String = "Select * From [@Z_OCRG] Where DocEntry = '" + oDocEntry + "'"
        Try

            oCustomer = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                If oCustomer.GetByKey(oRecordSet.Fields.Item("U_CardCode").Value) Then
                    oCustomer.CardName = oRecordSet.Fields.Item("U_CardName").Value
                    oCustomer.CardForeignName = oRecordSet.Fields.Item("U_CardName").Value
                    oCustomer.UserFields.Fields.Item("U_DOB").Value = oRecordSet.Fields.Item("U_DOB").Value
                    oCustomer.UserFields.Fields.Item("U_Title").Value = oRecordSet.Fields.Item("U_Title").Value
                    oCustomer.UserFields.Fields.Item("U_Occup").Value = oRecordSet.Fields.Item("U_Occup").Value
                    oCustomer.Cellular = IIf(oRecordSet.Fields.Item("U_Mobile").Value.ToString() = "", "", oRecordSet.Fields.Item("U_Mobile").Value.ToString())
                    oCustomer.Phone1 = IIf(oRecordSet.Fields.Item("U_TeleNo").Value.ToString() = "", "", oRecordSet.Fields.Item("U_TeleNo").Value.ToString())
                    oCustomer.FreeText = oRecordSet.Fields.Item("U_Remarks").Value

                    Dim iRes As Integer = oCustomer.Update()
                    If iRes = 0 Then
                        _retVal = True
                    Else
                        _retVal = False
                        Throw New Exception(oApplication.Company.GetLastErrorDescription())
                    End If
                End If
            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oCustomer)
        End Try
    End Function

    Public Function AddCustomerProfile(ByVal oForm As SAPbouiCOM.Form, ByVal strDocEntry As String) As String
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim oGeneralDataCollection As SAPbobsCOM.GeneralDataCollection
        Dim oChildData As SAPbobsCOM.GeneralData
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim strQuery As String = String.Empty
        Dim strCode As String = String.Empty
        oCompanyService = oApplication.Company.GetCompanyService()
        Try
            oGeneralService = oCompanyService.GetGeneralService("Z_OCPR")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            strQuery = "Select * From [@Z_OCRG] Where DocEntry = '" + strDocEntry + "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then

                Dim strCardCode As String = oRecordSet.Fields.Item("U_CardCode").Value
                Dim strCardName As String = oRecordSet.Fields.Item("U_CardName").Value

                Dim intCode As Integer = getMaxCode("@Z_OCPR", "DocEntry")
                strCode = String.Format("{0:000000000}", intCode)
                oGeneralData.SetProperty("U_CardCode", strCardCode)
                oGeneralData.SetProperty("U_CardName", strCardName)
                oGeneralData.SetProperty("U_RegNo", oRecordSet.Fields.Item("DocNum").Value.ToString())

                strQuery = " Select "
                strQuery += " T0.OpenDate,T0.OpprId,T0.Line,T0.U_Duration,T0.U_Dietitian1,T0.U_Dietitian2 "
                strQuery += " From OPR1 T0 JOIN OOPR T1 On T0.OpprID = T1.OpprID Where T1.U_RegNo = '" + strDocEntry + "'"
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery(strQuery)
                If Not oRecordSet.EoF Then
                    oGeneralDataCollection = oGeneralData.Child("Z_CPR3")
                    Dim intRow As String = 0
                    While Not oRecordSet.EoF
                        oChildData = oGeneralDataCollection.Add()
                        oChildData.SetProperty("U_OpprId", oRecordSet.Fields.Item("OpprId").Value.ToString())
                        oChildData.SetProperty("U_Line", oRecordSet.Fields.Item("Line").Value.ToString())
                        oChildData.SetProperty("U_VisitDate", oRecordSet.Fields.Item("OpenDate").Value)
                        oChildData.SetProperty("U_Duration", oRecordSet.Fields.Item("U_Duration").Value.ToString())
                        oChildData.SetProperty("U_Dietitian1", oRecordSet.Fields.Item("U_Dietitian1").Value.ToString())
                        oChildData.SetProperty("U_Dietitian2", oRecordSet.Fields.Item("U_Dietitian2").Value.ToString())
                        intRow += 1
                        oRecordSet.MoveNext()
                    End While
                    oGeneralService.Add(oGeneralData)
                End If

            End If
            Return strCode
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
        Return strCode
    End Function

    Public Function AddCustomerProfileFromCustomer(ByVal oForm As SAPbouiCOM.Form, ByVal strDocEntry As String) As String
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim strQuery As String = String.Empty
        Dim strCode As String = String.Empty
        Dim oCustomer As SAPbobsCOM.BusinessPartners = Nothing
        oCompanyService = oApplication.Company.GetCompanyService()
        Try
            oGeneralService = oCompanyService.GetGeneralService("Z_OCPR")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oCustomer = oApplication.Company.GetBusinessObject(BoObjectTypes.oBusinessPartners)
            If oCustomer.Browser.GetByKeys(strDocEntry) Then
                If oCustomer.CardType = BoCardTypes.cCustomer Then
                    strQuery = "Select * From [@Z_OCPR] Where U_CardCode = '" + oCustomer.CardCode + "'"
                    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery(strQuery)
                    If oRecordSet.EoF Then
                        Dim strCardCode As String = oCustomer.CardCode
                        Dim strCardName As String = oCustomer.CardName
                        Dim intCode As Integer = getMaxCode("@Z_OCPR", "DocEntry")
                        strCode = String.Format("{0:000000000}", intCode)
                        oGeneralData.SetProperty("U_CardCode", strCardCode)
                        oGeneralData.SetProperty("U_CardName", strCardName)
                        oGeneralService.Add(oGeneralData)
                    End If
                End If
            End If
            Return strCode
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
        Return strCode
    End Function

    Public Function CreateSalesOpp(ByVal oForm As SAPbouiCOM.Form, ByVal oDocEntry As String) As Boolean
        Dim oSalesOp As SAPbobsCOM.SalesOpportunities = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim _retVal As Boolean = False
        Dim strQuery As String = "Select * From [@Z_OCRG] Where DocEntry = '" + oDocEntry + "'"
        Try
            oSalesOp = oApplication.Company.GetBusinessObject(BoObjectTypes.oSalesOpportunities)
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then

                oSalesOp.CardCode = oRecordSet.Fields.Item("U_CardCode").Value
                oSalesOp.CustomerName = oRecordSet.Fields.Item("U_CardName").Value
                oSalesOp.StartDate = oRecordSet.Fields.Item("U_VisitDate").Value
                oSalesOp.UserFields.Fields.Item("U_RegNo").Value = oDocEntry.ToString()
                oSalesOp.Lines.StartDate = oRecordSet.Fields.Item("U_VisitDate").Value
                oSalesOp.Lines.ClosingDate = oRecordSet.Fields.Item("U_VisitDate").Value

                oSalesOp.Lines.StageKey = 1
                oSalesOp.Lines.MaxLocalTotal = 1
                oSalesOp.Lines.UserFields.Fields.Item("U_Duration").Value = oRecordSet.Fields.Item("U_Duration").Value.ToString()
                oSalesOp.Lines.UserFields.Fields.Item("U_Dietitian1").Value = oRecordSet.Fields.Item("U_Dietitian1").Value
                oSalesOp.Lines.UserFields.Fields.Item("U_Dietitian2").Value = oRecordSet.Fields.Item("U_Dietitian2").Value
                oSalesOp.Lines.Add()

                Dim iRes As Integer = oSalesOp.Add()
                If iRes = 0 Then
                    _retVal = True
                Else
                    _retVal = False
                    Throw New Exception(oApplication.Company.GetLastErrorDescription())
                End If
            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function AddCustomerProgram(ByVal oForm As SAPbouiCOM.Form, ByVal strDocType As String, ByVal strObjectKey As String) As String
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim oGenDataCollection As SAPbobsCOM.GeneralDataCollection
        Dim oGenDataCollection_1 As SAPbobsCOM.GeneralDataCollection
        Dim oGenDataCollection_2 As SAPbobsCOM.GeneralDataCollection
        Dim oGenDataChild As SAPbobsCOM.GeneralData
        Dim oGenDataChild_1 As SAPbobsCOM.GeneralData
        Dim oGenDataChild_2 As SAPbobsCOM.GeneralData

        Dim oRecordSet, oRecordSetMain As SAPbobsCOM.Recordset
        Dim strQuery As String = String.Empty
        Dim strCode As String = String.Empty
        Dim oDoc As SAPbobsCOM.Documents = Nothing
        Dim oDoc_Lines As SAPbobsCOM.Document_Lines
        Dim blnLineExists As Boolean
        oCompanyService = oApplication.Company.GetCompanyService()
        Try

            Select Case strDocType
                Case frm_INVOICES
                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            End Select
            oRecordSetMain = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            If oDoc.Browser.GetByKeys(strObjectKey) Then

                '  strQuery = "Select * From [OINV] Where DocEntry = '" + oDoc.DocEntry.ToString() + "'"
                '  oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oRecordSet.DoQuery(strQuery)
                'MessageBox.Show(oDoc.Cancelled = SAPbobsCOM.BoYesNoEnum.tNO
                'oDoc.CancelStatus = SAPbobsCOM.CancelStatusEnum.csCancellation
                '   MsgBox(oDoc.DocumentStatus)
                '  MsgBox(oDoc.Cancelled)


                If (oDoc.DocumentStatus = BoStatus.bost_Open Or oDoc.DocumentStatus = BoStatus.bost_Close) And oDoc.Cancelled = SAPbobsCOM.BoYesNoEnum.tNO Then
                    strQuery = "Select T0.ItemCode,T1.ItemName,Quantity,U_Fdate,U_Edate,DATEDIFF(day,U_Fdate,U_Edate) As 'NoofDays',T4.CardCode,T4.CardName,T0.LineNum From INV1 T0 inner Join OINV T4 on T4.DocEntry=T0.DocEntry "
                    strQuery += " JOIN OITM T1 On T0.ItemCode = T1.ItemCode "
                    strQuery += " JOIN OITB T2 On T1.ItmsGrpCod = T2.ItmsGrpCod And T2.U_Program = 'Y'"
                    strQuery += " Where T0.DocEntry = '" + oDoc.DocEntry.ToString() + "' "
                    oRecordSetMain.DoQuery(strQuery)
                    For index As Integer = 0 To oRecordSetMain.RecordCount - 1
                        ' If Not oRecordSet.EoF Then
                        blnLineExists = False
                        Dim strCardCode As String = oDoc.CardCode 'ordSet.Fields.Item("CardCode").Value
                        Dim strCardName As String = oDoc.CardName ' oRecordSet.Fields.Item("CardName").Value
                        oGeneralService = oCompanyService.GetGeneralService("Z_OCPM")
                        oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                        oGenDataChild = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                        Dim intCode As Integer = getMaxCode("@Z_OCPM", "DocEntry")
                        strCode = String.Format("{0:000000000}", intCode)
                        oGeneralData.SetProperty("U_CardCode", strCardCode)
                        oGeneralData.SetProperty("U_CardName", strCardName)
                        oGeneralData.SetProperty("U_InvRef", oDoc.DocEntry.ToString())
                        oGeneralData.SetProperty("U_OrderNo", oDoc.DocNum.ToString())

                        oGenDataCollection = oGeneralData.Child("Z_CPM1")
                        oGenDataCollection_1 = oGeneralData.Child("Z_CPM4")
                        oGenDataCollection_2 = oGeneralData.Child("Z_CPM5")

                        strQuery = "Select T0.ItemCode,T1.ItemName,Quantity,U_Fdate,U_Edate,DATEDIFF(day,U_Fdate,U_Edate) As 'NoofDays' From INV1 T0 "
                        strQuery += " JOIN OITM T1 On T0.ItemCode = T1.ItemCode "
                        strQuery += " JOIN OITB T2 On T1.ItmsGrpCod = T2.ItmsGrpCod And T2.U_Program = 'Y'"
                        strQuery += " Where T0.DocEntry = '" + oDoc.DocEntry.ToString() + "' and T0.LineNum=" & oRecordSetMain.Fields.Item("LineNum").Value
                        oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery(strQuery)
                        If Not oRecordSet.EoF Then
                            blnLineExists = True
                            Dim intNoofDays As Integer = CInt(oRecordSet.Fields.Item("NoofDays").Value)
                            oGeneralData.SetProperty("U_PrgCode", oRecordSet.Fields.Item("ItemCode").Value)
                            oGeneralData.SetProperty("U_PrgName", oRecordSet.Fields.Item("ItemName").Value)
                            oGeneralData.SetProperty("U_PFromDate", oRecordSet.Fields.Item("U_Fdate").Value)
                            oGeneralData.SetProperty("U_PToDate", oRecordSet.Fields.Item("U_Edate").Value)
                            oGeneralData.SetProperty("U_NoOfDays", CInt(oRecordSet.Fields.Item("Quantity").Value))
                            oGeneralData.SetProperty("U_RemDays", CInt(oRecordSet.Fields.Item("Quantity").Value))
                            Dim dtFromDate As Date = oRecordSet.Fields.Item("U_Fdate").Value
                            For intRow As Integer = 0 To intNoofDays
                                oGenDataChild = oGenDataCollection.Add()
                                oGenDataChild_1 = oGenDataCollection_1.Add()
                                oGenDataChild_2 = oGenDataCollection_2.Add()
                                oGenDataChild.SetProperty("U_PrgDate", dtFromDate.AddDays(intRow))
                                oGenDataChild_1.SetProperty("U_PrgDate", dtFromDate.AddDays(intRow))
                                oGenDataChild_2.SetProperty("U_PrgDate", dtFromDate.AddDays(intRow))
                                Dim strIncStatus As String = checkExclude(strCardCode, dtFromDate.AddDays(intRow))
                                oGenDataChild.SetProperty("U_AppStatus", strIncStatus.Trim())
                                oGenDataChild.SetProperty("U_Remarks", oRecordSet.Fields.Item("ItemCode").Value.ToString)
                            Next
                        End If
                        If blnLineExists = True Then
                            oGeneralService.Add(oGeneralData)
                        End If

                        oRecordSetMain.MoveNext()
                    Next
                    ' End If
                ElseIf oDoc.DocumentStatus = BoStatus.bost_Close Then 'And oDoc.CancelStatus = SAPbobsCOM.CancelStatusEnum.csCancellation Then
                    If oApplication.SBO_Application.MessageBox("Do You want to Cancel All Program related to Cancelled Invoice Document...?", 1, "Yes", "No") = 1 Then
                        strQuery = " Select Distinct BaseEntry From INV1 "
                        strQuery += " Where DocEntry = '" + oDoc.DocEntry.ToString() + "' "
                        oRecordSetMain.DoQuery(strQuery)
                        For index As Integer = 0 To oRecordSetMain.RecordCount - 1
                            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                            ' strQuery = " Update [@Z_OCPM] SET U_Cancel = 'Y', "
                            '  strQuery += " U_RemDays = 0 "
                            strQuery = " Update [@Z_OCPM] SET U_Cancel = 'Y' "
                            strQuery += " Where U_InvRef = '" + oRecordSetMain.Fields.Item(0).Value.ToString() + "'"
                            oRecordSet.DoQuery(strQuery)
                        Next
                    End If
                End If
            End If
            Return strCode
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
        Return strCode
    End Function

    Private Function checkExclude(ByVal strCardCode As String, ByVal dtPrgDate As Date) As String
        Dim _retVal As String = "I"
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            'Based on Day...
            strQuery = "Select U_Monday,U_Tuesday,U_Wednesday,U_Thursday,U_Friday,U_Saturday,U_Sunday From [@Z_OCPR] "
            strQuery += " Where U_CardCode = '" + strCardCode + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                Select Case dtPrgDate.DayOfWeek
                    Case DayOfWeek.Monday
                        If oRecordSet.Fields.Item("U_Monday").Value = "Y" Then
                            _retVal = "E"
                            'Exit Select
                        End If
                    Case DayOfWeek.Tuesday
                        If oRecordSet.Fields.Item("U_Tuesday").Value = "Y" Then
                            _retVal = "E"
                            'Exit Select
                        End If
                    Case DayOfWeek.Wednesday
                        If oRecordSet.Fields.Item("U_Wednesday").Value = "Y" Then
                            _retVal = "E"
                            'Exit Select
                        End If
                    Case DayOfWeek.Thursday
                        If oRecordSet.Fields.Item("U_Thursday").Value = "Y" Then
                            _retVal = "E"
                            'Exit Select
                        End If
                    Case DayOfWeek.Friday
                        If oRecordSet.Fields.Item("U_Friday").Value = "Y" Then
                            _retVal = "E"
                            'Exit Select
                        End If
                    Case DayOfWeek.Saturday
                        If oRecordSet.Fields.Item("U_Saturday").Value = "Y" Then
                            _retVal = "E"
                            'Exit Select
                        End If
                    Case DayOfWeek.Sunday
                        If oRecordSet.Fields.Item("U_Sunday").Value = "Y" Then
                            _retVal = "E"
                            'Exit Select
                        End If
                End Select
            End If

            'Based on Date...
            If _retVal = "I" Then
                strQuery = "Select T0.DocEntry From [@Z_CPR4] T0 JOIN [@Z_OCPR] T1 On T0.DocEntry = T1.DocEntry "
                strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
                strQuery += " And Convert(VarChar(8),T0.U_ExDate,112) = '" + dtPrgDate.ToString("yyyyMMdd") + "'"
                strQuery += " And ISNULL(T0.U_Include,'N') = 'N' "
                oRecordSet.DoQuery(strQuery)
                If Not oRecordSet.EoF Then
                    _retVal = "E"
                Else
                    _retVal = "I"
                End If
            Else
                strQuery = "Select T0.DocEntry From [@Z_CPR4] T0 JOIN [@Z_OCPR] T1 On T0.DocEntry = T1.DocEntry "
                strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
                strQuery += " And Convert(VarChar(8),T0.U_ExDate,112) = '" + dtPrgDate.ToString("yyyyMMdd") + "'"
                strQuery += " And ISNULL(T0.U_Include,'N') = 'Y' "
                oRecordSet.DoQuery(strQuery)
                If Not oRecordSet.EoF Then
                    _retVal = "I"
                End If
            End If

            'Remove Date...
            If _retVal = "I" Then
                strQuery = "Select T0.DocEntry From [@Z_CPR8] T0 JOIN [@Z_OCPR] T1 On T0.DocEntry = T1.DocEntry "
                strQuery += " Where T1.U_CardCode = '" & strCardCode & "'"
                strQuery += " And '" & dtPrgDate.ToString("yyyyMMdd") & "' Between Convert(VarChar(8), T0.U_FDate, 112)  AND Convert(VarChar(8), T0.U_TDate, 112) "
                oRecordSet.DoQuery(strQuery)
                If Not oRecordSet.EoF Then
                    _retVal = "E"
                End If
            End If

            'Suspended Date...
            If _retVal = "I" Then
                strQuery = "Select T0.DocEntry From [@Z_CPR9] T0 JOIN [@Z_OCPR] T1 On T0.DocEntry = T1.DocEntry "
                strQuery += " Where T1.U_CardCode = '" & strCardCode & "'"
                strQuery += " And '" & dtPrgDate.ToString("yyyyMMdd") & "' Between Convert(VarChar(8), T0.U_FDate, 112)  AND Convert(VarChar(8), T0.U_TDate, 112) "
                oRecordSet.DoQuery(strQuery)
                If Not oRecordSet.EoF Then
                    _retVal = "E"
                End If
            End If


            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Private Function checkExcludeAlone(ByVal strCardCode As String, ByVal dtPrgDate As Date) As String
        Dim _retVal As String = "I"
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            'Based on Day...
            strQuery = "Select U_Monday,U_Tuesday,U_Wednesday,U_Thursday,U_Friday,U_Saturday,U_Sunday From [@Z_OCPR] "
            strQuery += " Where U_CardCode = '" + strCardCode + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                Select Case dtPrgDate.DayOfWeek
                    Case DayOfWeek.Monday
                        If oRecordSet.Fields.Item("U_Monday").Value = "Y" Then
                            _retVal = "E"
                            'Exit Select
                        End If
                    Case DayOfWeek.Tuesday
                        If oRecordSet.Fields.Item("U_Tuesday").Value = "Y" Then
                            _retVal = "E"
                            'Exit Select
                        End If
                    Case DayOfWeek.Wednesday
                        If oRecordSet.Fields.Item("U_Wednesday").Value = "Y" Then
                            _retVal = "E"
                            'Exit Select
                        End If
                    Case DayOfWeek.Thursday
                        If oRecordSet.Fields.Item("U_Thursday").Value = "Y" Then
                            _retVal = "E"
                            'Exit Select
                        End If
                    Case DayOfWeek.Friday
                        If oRecordSet.Fields.Item("U_Friday").Value = "Y" Then
                            _retVal = "E"
                            'Exit Select
                        End If
                    Case DayOfWeek.Saturday
                        If oRecordSet.Fields.Item("U_Saturday").Value = "Y" Then
                            _retVal = "E"
                            'Exit Select
                        End If
                    Case DayOfWeek.Sunday
                        If oRecordSet.Fields.Item("U_Sunday").Value = "Y" Then
                            _retVal = "E"
                            'Exit Select
                        End If
                End Select
            End If

            If _retVal = "I" Then
                strQuery = "Select T0.DocEntry From [@Z_CPR4] T0 JOIN [@Z_OCPR] T1 On T0.DocEntry = T1.DocEntry "
                strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
                strQuery += " And Convert(VarChar(8),T0.U_ExDate,112) = '" + dtPrgDate.ToString("yyyyMMdd") + "'"
                strQuery += " And ISNULL(T0.U_Include,'N') = 'N' "
                oRecordSet.DoQuery(strQuery)
                If Not oRecordSet.EoF Then
                    _retVal = "E"
                Else
                    _retVal = "I"
                End If
            Else
                strQuery = "Select T0.DocEntry From [@Z_CPR4] T0 JOIN [@Z_OCPR] T1 On T0.DocEntry = T1.DocEntry "
                strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
                strQuery += " And Convert(VarChar(8),T0.U_ExDate,112) = '" + dtPrgDate.ToString("yyyyMMdd") + "'"
                strQuery += " And ISNULL(T0.U_Include,'N') = 'Y' "
                oRecordSet.DoQuery(strQuery)
                If Not oRecordSet.EoF Then
                    _retVal = "I"
                End If
            End If

            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Private Function checkSuspendONOFF(ByVal strCardCode As String, ByVal dtPrgDate As Date) As String
        Dim _retVal As String = "O"
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            strQuery = "Select T0.DocEntry From [@Z_CPR9] T0 JOIN [@Z_OCPR] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " Where T1.U_CardCode = '" & strCardCode & "'"
            strQuery += " And '" & dtPrgDate.ToString("yyyyMMdd") & "' Between Convert(VarChar(8), T0.U_FDate, 112)  AND Convert(VarChar(8), T0.U_TDate, 112) "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                _retVal = "F"
            Else
                _retVal = "O"
            End If

            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function getAgebyDOB(ByVal oForm As SAPbouiCOM.Form, ByVal strDOB As String) As String
        Dim _retVal As String
        Try
            Dim dtDOB As Date = strDOB.Substring(0, 4) + "-" + strDOB.Substring(4, 2) + "-" + strDOB.Substring(6, 2)
            Dim dtDate As Date = System.DateTime.Now
            Dim strYear As String = strDOB.Substring(0, 4)
            Dim strMonth As String = strDOB.Substring(4, 2)
            Dim strDate As String = strDOB.Substring(6, 2)
            Dim DOB As New DateTime(strYear, strMonth, strDate)
            Dim Years As Integer = Now.Year - DOB.Year 'DateDiff(DateInterval.Year, DOB, Now) - 1
            If (DOB > Now.AddYears(-Years)) Then Years -= 1
            'Dim Months As Integer = DateDiff(DateInterval.Month, DOB, Now) Mod 12
            'Dim days As Integer = DateDiff(DateInterval.Day, DOB, Now) Mod 30 - 10
            _retVal = Years.ToString()
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    'Public Function getDateDiff(ByVal oForm As SAPbouiCOM.Form, ByVal strFromDate As String, ByVal strToDate As String) As String
    '    Dim _retVal As String
    '    Try
    '        Dim dtFromDate As Date = strFromDate.Substring(0, 4) + "-" + strFromDate.Substring(4, 2) + "-" + strFromDate.Substring(6, 2)
    '        Dim dtToDate As Date = strToDate.Substring(0, 4) + "-" + strToDate.Substring(4, 2) + "-" + strToDate.Substring(6, 2)
    '        Dim days As Integer = DateDiff(DateInterval.Day, dtFromDate, dtToDate)
    '        _retVal = days.ToString()
    '        Return _retVal
    '    Catch ex As Exception 
    '.Log.Trace_DIET_AddOn_Error(ex)
    '        Throw ex 
    ''oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
    '    End Try
    'End Function

    Public Function getDateDiff(ByVal oForm As SAPbouiCOM.Form, strCardCode As String, ByVal strFromDate As String, ByVal strToDate As String) As Integer
        Dim _retVal As Integer
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim strQuery As String = String.Empty
        Try
            Dim dtFromDate As Date = strFromDate.Substring(0, 4) + "-" + strFromDate.Substring(4, 2) + "-" + strFromDate.Substring(6, 2)
            Dim dtToDate As Date = strToDate.Substring(0, 4) + "-" + strToDate.Substring(4, 2) + "-" + strToDate.Substring(6, 2)
            Dim days As Integer = 0
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            'strQuery = "Select Count(LineId) From [@Z_CPM1] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry "
            'strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
            'strQuery += " AND Convert(VarChar(8),T0.U_PrgDate,112) >= '" + strFromDate + "'"
            'strQuery += " And Convert(VarChar(8),T0.U_PrgDate,112) <= '" + strToDate + "'"
            'strQuery += " And U_AppStatus = 'I' "
            strQuery += " Select U_RemDays From [@Z_OCPM] Where U_RemDays > 0 "
            strQuery += " And U_CardCode = '" + strCardCode + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                _retVal = CInt(oRecordSet.Fields.Item(0).Value)
            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function GetCustomerProfile(ByVal oForm As SAPbouiCOM.Form) As String
        Dim _retVal As String
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            strQuery = "Select DocEntry From [@Z_OCPR] Where U_CardCode = '" + oForm.Items.Item("5").Specific.value + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                _retVal = oRecordSet.Fields.Item(0).Value
            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    'Public Function GetProgramByItem(ByVal oForm As SAPbouiCOM.Form) As String
    '    Dim _retVal As String
    '    Dim strQuery As String = String.Empty
    '    Try
    '        Dim oRecordSet As SAPbobsCOM.Recordset
    '        oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
    '        strQuery = "Select ItemCode From OITM Where U_PrgCode = '" + oForm.Items.Item("16").Specific.value + "'"
    '        oRecordSet.DoQuery(strQuery)
    '        If Not oRecordSet.EoF Then
    '            _retVal = oRecordSet.Fields.Item(0).Value
    '        End If
    '        Return _retVal
    '    Catch ex As Exception 
    '.Log.Trace_DIET_AddOn_Error(ex)
    '        Throw ex 
    ''oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
    '    End Try
    'End Function

    Public Function GetItemName(ByVal strItemCode As String) As String
        Dim _retVal As String
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            strQuery = "Select ItemName From OITM Where itemCode = '" + strItemCode + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                _retVal = oRecordSet.Fields.Item(0).Value
            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function GetItemPrice(ByVal strItem As String, ByVal strPrice As String) As Double
        Dim _retVal As Double = 0
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            strQuery = "Select Price From ITM1 Where ItemCode = '" + strItem + "' And PriceList = '" + strPrice + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function GetCustomerCode(ByRef strCode As String) As String
        Dim _retVal As String
        Try
            Dim intCode As Integer = getMaxCode("OCRD", "DocEntry")
            strCode = "D" + String.Format("{0:00000000}", intCode)
            _retVal = strCode
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.SBO_Application.SetStatusBarMessage(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
    End Function

    Public Function hasBOM(ByVal strItemCode As String) As Boolean
        Try
            'Dim oBOM As SAPbobsCOM.ProductTrees
            'oBOM = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees)
            'If oBOM.GetByKey(strItemCode) Then
            '    Return True
            'Else
            '    Return False
            'End If
            Dim oISBOM As SAPbobsCOM.Recordset
            oISBOM = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            strSQL = "Select TreeType From OITM Where ItemCode = '" & strItemCode & "'"
            oISBOM.DoQuery(strSQL)
            If Not oISBOM.EoF Then
                If oISBOM.Fields.Item("TreeType").Value = "N" Then
                    Return False
                Else
                    Return True
                End If
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    'Public Sub get_ChildItems(ByVal strCustomer As String, ByVal strItemCode As String, ByRef strDislike As String, ByRef strMedical As String)
    '    Try
    '        Dim oBOM As SAPbobsCOM.ProductTrees
    '        Dim oBOM_Lines As SAPbobsCOM.ProductTrees_Lines
    '        oBOM = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees)
    '        If oBOM.GetByKey(strItemCode) Then
    '            oBOM_Lines = oBOM.Items
    '            For bomlineindex As Integer = 0 To oBOM_Lines.Count - 1
    '                oBOM_Lines.SetCurrentLine(bomlineindex)
    '                Dim strChildItem As String = oBOM_Lines.ItemCode
    '                If hasBOM(strChildItem) Then

    '                    Dim strDislikeItem As String = GetDisLikeItem(strCustomer, strChildItem)
    '                    If strDislikeItem.Trim.Length > 0 Then
    '                        If strDislike.Length = 0 Then
    '                            strDislike = strDislikeItem
    '                        Else
    '                            If Not strDislike.Contains(strDislikeItem) Then
    '                                strDislike += "," + strDislikeItem
    '                            End If
    '                        End If
    '                    End If

    '                    Dim strMedicalItem As String = GetMedicalItem(strCustomer, strChildItem)
    '                    If strMedicalItem.Trim.Length > 0 Then
    '                        If strMedical.Length = 0 Then
    '                            strMedical = strMedicalItem
    '                        Else
    '                            If Not strMedical.Contains(strMedicalItem) Then
    '                                strMedical += "," + strMedicalItem
    '                            End If
    '                        End If
    '                    End If

    '                    get_ChildItems(strCustomer, strChildItem, strDislike, strMedical)
    '                Else

    '                    Dim strDislikeItem As String = GetDisLikeItem(strCustomer, strChildItem)
    '                    If strDislikeItem.Trim.Length > 0 Then
    '                        If strDislike.Length = 0 Then
    '                            strDislike = strDislikeItem
    '                        Else
    '                            If Not strDislike.Contains(strDislikeItem) Then
    '                                strDislike += "," + strDislikeItem
    '                            End If
    '                        End If
    '                    End If

    '                    Dim strMedicalItem As String = GetMedicalItem(strCustomer, strChildItem)
    '                    If strMedicalItem.Trim.Length > 0 Then
    '                        If strMedical.Length = 0 Then
    '                            strMedical = strMedicalItem
    '                        Else
    '                            If Not strMedical.Contains(strMedicalItem) Then
    '                                strMedical += "," + strMedicalItem
    '                            End If
    '                        End If
    '                    End If
    '                End If

    '            Next
    '        Else
    '            Dim strDislikeItem As String = GetDisLikeItem(strCustomer, strItemCode)
    '            If strDislikeItem.Trim.Length > 0 Then
    '                If strDislike.Length = 0 Then
    '                    strDislike = strDislikeItem
    '                Else
    '                    If Not strDislike.Contains(strDislikeItem) Then
    '                        strDislike += "," + strDislikeItem
    '                    End If
    '                End If
    '            End If

    '            Dim strMedicalItem As String = GetMedicalItem(strCustomer, strItemCode)
    '            If strMedicalItem.Trim.Length > 0 Then
    '                If strMedical.Length = 0 Then
    '                    strMedical = strMedicalItem
    '                Else
    '                    If Not strMedical.Contains(strMedicalItem) Then
    '                        strMedical += "," + strMedicalItem
    '                    End If
    '                End If
    '            End If
    '        End If
    '    Catch ex As Exception 
    'oApplication.Log.Trace_DIET_AddOn_Error(ex)
    '        Throw ex 
    ''oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
    '    End Try
    'End Sub

    Public Sub get_ChildItems(ByVal strCustomer As String, ByVal strItemCode As String, ByRef strDislike As String, ByRef strMedical As String)
        Try
            Dim oBOM_Lines As SAPbobsCOM.Recordset
            oBOM_Lines = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSQL = "Select T1.Code From OITT T0 JOIN ITT1 T1 On T0.Code = T1.Father Where T0.Code = '" & strItemCode & "'"
            oBOM_Lines.DoQuery(strSQL)
            If Not oBOM_Lines.EoF Then
                While Not oBOM_Lines.EoF
                    Dim strChildItem As String = oBOM_Lines.Fields.Item("Code").Value
                    If hasBOM(strChildItem) Then
                        Dim strDislikeItem As String = GetDisLikeItem(strCustomer, strChildItem)
                        If strDislikeItem.Trim.Length > 0 Then
                            If strDislike.Length = 0 Then
                                strDislike = strDislikeItem
                            Else
                                If Not strDislike.Contains(strDislikeItem) Then
                                    strDislike += "," + strDislikeItem
                                End If
                            End If
                        End If

                        Dim strMedicalItem As String = GetMedicalItem(strCustomer, strChildItem)
                        If strMedicalItem.Trim.Length > 0 Then
                            If strMedical.Length = 0 Then
                                strMedical = strMedicalItem
                            Else
                                If Not strMedical.Contains(strMedicalItem) Then
                                    strMedical += "," + strMedicalItem
                                End If
                            End If
                        End If

                        get_ChildItems(strCustomer, strChildItem, strDislike, strMedical)
                    Else
                        Dim strDislikeItem As String = GetDisLikeItem(strCustomer, strChildItem)
                        If strDislikeItem.Trim.Length > 0 Then
                            If strDislike.Length = 0 Then
                                strDislike = strDislikeItem
                            Else
                                If Not strDislike.Contains(strDislikeItem) Then
                                    strDislike += "," + strDislikeItem
                                End If
                            End If
                        End If

                        Dim strMedicalItem As String = GetMedicalItem(strCustomer, strChildItem)
                        If strMedicalItem.Trim.Length > 0 Then
                            If strMedical.Length = 0 Then
                                strMedical = strMedicalItem
                            Else
                                If Not strMedical.Contains(strMedicalItem) Then
                                    strMedical += "," + strMedicalItem
                                End If
                            End If
                        End If
                    End If
                    oBOM_Lines.MoveNext()
                End While
            Else
                Dim strDislikeItem As String = GetDisLikeItem(strCustomer, strItemCode)
                If strDislikeItem.Trim.Length > 0 Then
                    If strDislike.Length = 0 Then
                        strDislike = strDislikeItem
                    Else
                        If Not strDislike.Contains(strDislikeItem) Then
                            strDislike += "," + strDislikeItem
                        End If
                    End If
                End If

                Dim strMedicalItem As String = GetMedicalItem(strCustomer, strItemCode)
                If strMedicalItem.Trim.Length > 0 Then
                    If strMedical.Length = 0 Then
                        strMedical = strMedicalItem
                    Else
                        If Not strMedical.Contains(strMedicalItem) Then
                            strMedical += "," + strMedicalItem
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Public Function GetDisLikeItem(ByVal strCardCode As String, ByVal strItem As String) As String
        Dim _retVal As String = String.Empty
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            strQuery = " Select ISNULL(T2.U_Name,'') From [@Z_CPR1] T0  "
            strQuery += " JOIN [@Z_OCPR] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN [@Z_ODLK] T2 On T2.U_Code = T0.U_DLikeItem "
            strQuery += " JOIN [@Z_DLK1] T3 On T3.DocEntry = T2.DocEntry "
            strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
            strQuery += " And T3.U_ItemCode = '" + strItem + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                _retVal = oRecordSet.Fields.Item(0).Value
            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function GetMedicalItem(ByVal strCardCode As String, ByVal strItem As String) As String
        Dim _retVal As String = String.Empty
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            strQuery = " Select ISNULL(T4.FrgnName,T3.U_ItemName) From [@Z_CPR2] T0  "
            strQuery += " JOIN [@Z_OCPR] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN [@Z_OMST] T2 On T2.U_Code = T0.U_MSCode "
            strQuery += " JOIN [@Z_MST1] T3 On T3.DocEntry = T2.DocEntry "
            strQuery += " JOIN OITM T4 On T4.ItemCode = T3.U_ItemCode "
            strQuery += " Where T1.U_CardCode = '" + strCardCode + "'"
            strQuery += " And T3.U_ItemCode = '" + strItem + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                _retVal = oRecordSet.Fields.Item(0).Value
            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function updateProgramDocument(ByVal strDocEntry As String) As Boolean
        Dim _retVal As Boolean = False
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            'Header
            strQuery = " Update T0 Set T0.U_RemDays = 0,T0.U_Transfer = 'Y',T0.U_DocStatus = 'C' "
            strQuery += " From [@Z_OCPM] T0 JOIN [@Z_OPGT] T1 On T0.DocEntry = T1.U_ProgramID "
            strQuery += " Where T1.DocEntry = '" + strDocEntry + "'"
            oRecordSet.DoQuery(strQuery)

            'Lines Staus
            strQuery = " Update T0 Set T0.U_Status = 'T' "
            strQuery += " From [@Z_CPM1] T0 JOIN [@Z_PGT1] T1 "
            strQuery += " On T0.DocEntry = T1.U_PrgNo And T0.LineId = T1.U_PrgLine "
            strQuery += " Where T1.DocEntry = '" + strDocEntry + "'"
            oRecordSet.DoQuery(strQuery)

            _retVal = True
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function AddTransferProgram(ByVal oForm As SAPbouiCOM.Form, ByVal strObjectKey As String) As Boolean
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Dim oGenDataChild As SAPbobsCOM.GeneralData
        Dim oGenDataChild_R As SAPbobsCOM.GeneralData

        Dim oGenDataCollection As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralDataCollection_R As SAPbobsCOM.GeneralDataCollection

        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim strQuery As String = String.Empty
        Dim _retVal As Boolean = False
        Dim strCode As String = String.Empty

        oCompanyService = oApplication.Company.GetCompanyService()
        Try
            oGeneralService = oCompanyService.GetGeneralService("Z_OCPM")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGenDataChild = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

            strQuery = "Select T0.U_TCardCode,T0.U_TCardName,T0.DocEntry,T0.U_NoOfDays,T0.U_PrgCode,T0.U_PFromDate,T0.U_PToDate,"
            strQuery += " T0.U_TrnType,T0.U_TProgramID,T0.U_TNoOfDays,T0.U_TPrgCode,T1.ItemName,T2.Currency,T0.U_CardCode From [@Z_OPGT] T0 JOIN OITM T1 ON T1.ItemCode = T0.U_PrgCode "
            strQuery += " JOIN OCRD T2 On T0.U_TCardCode = T2.CardCode "
            strQuery += " Where T0.DocEntry = '" + strObjectKey.ToString() + "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                Dim strBCardCode As String = oRecordSet.Fields.Item("U_CardCode").Value
                Dim strCardCode As String = oRecordSet.Fields.Item("U_TCardCode").Value
                Dim strCardName As String = oRecordSet.Fields.Item("U_TCardName").Value
                Dim strDocEntry As String = oRecordSet.Fields.Item("DocEntry").Value
                Dim strNooDays As String = oRecordSet.Fields.Item("U_NoOfDays").Value
                Dim strTType As String = oRecordSet.Fields.Item("U_TrnType").Value
                Dim strTProgram As String = oRecordSet.Fields.Item("U_TProgramID").Value
                Dim strTNooDays As String = oRecordSet.Fields.Item("U_TNoOfDays").Value
                Dim strTProg As String = oRecordSet.Fields.Item("U_TPrgCode").Value
                Dim intNoofDays As Integer = CInt(strTNooDays)
                Dim strCurrency As String = oRecordSet.Fields.Item("Currency").Value
                Dim strLCurrency, strSCurrency As String
                strLCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
                strSCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().SystemCurrency

                strQuery = "Select VatGourpSa From OITM Where ItemCode = '" & strTProg & "'"
                Dim strTaxCode As String = oApplication.Utilities.getRecordSetValueString(strQuery, "VatGourpSa")
                Dim dblDiscount As Double = oApplication.Utilities.getRecordSetValue("Select Discount From OCRD Where CardCode = '" & strCardCode & "'", "Discount")

                Dim dtProgramFrom As Date

                Try
                    dtProgramFrom = CDate(oRecordSet.Fields.Item("U_PFromDate").Value)
                Catch ex As Exception
                    oApplication.Log.Trace_DIET_AddOn_Error(ex)

                End Try


                If strTType = "C" Then
                    Dim intCode As Integer = getMaxCode("@Z_OCPM", "DocEntry")
                    strCode = String.Format("{0:000000000}", intCode)

                    oGeneralData.SetProperty("U_CardCode", strCardCode)
                    oGeneralData.SetProperty("U_BCardCode", strBCardCode)
                    oGeneralData.SetProperty("U_CardName", strCardName)
                    oGeneralData.SetProperty("U_NoOfDays", strTNooDays)
                    oGeneralData.SetProperty("U_RemDays", strTNooDays)

                    Dim strItemName As String = GetItemName(strTProg)
                    oGeneralData.SetProperty("U_PrgCode", strTProg)
                    oGeneralData.SetProperty("U_PrgName", strItemName)
                    oGeneralData.SetProperty("U_TrnRef", strObjectKey.ToString())
                    Dim strDocCurrency As String = String.Empty

                    If strCurrency <> "##" Then
                        If strCurrency = strLCurrency Then
                            oGeneralData.SetProperty("U_CurSour", "L")
                        ElseIf strCurrency = strSCurrency Then
                            oGeneralData.SetProperty("U_CurSour", "S")
                        End If
                        oGeneralData.SetProperty("U_VenCur", strCurrency)
                        oGeneralData.SetProperty("U_DocCur", strCurrency)
                        strDocCurrency = strCurrency
                        oGeneralData.SetProperty("U_DocRate", GetCurrencyRate(oForm, strCurrency))
                    ElseIf strCurrency = "##" Then
                        oGeneralData.SetProperty("U_CurSour", "C")
                        oGeneralData.SetProperty("U_VenCur", strCurrency)
                        oGeneralData.SetProperty("U_DocCur", strLCurrency)
                        strDocCurrency = strLCurrency
                        oGeneralData.SetProperty("U_DocRate", 1)
                    End If

                    oGenDataCollection = oGeneralData.Child("Z_CPM1")
                    'oGeneralDataCollection_R = oGeneralData.Child("Z_CPM6")

                    strQuery = "Select U_PFromDate,U_PToDate,U_PrgCode,U_TPrgCode From [@Z_OPGT] T0 "
                    strQuery += " Where T0.DocEntry = '" + strDocEntry.ToString() + "' "
                    oRecordSet.DoQuery(strQuery)

                    'Dim dtFromDate As Date = System.DateTime.Now
                    Dim dtFromDate As Date = dtProgramFrom
                    oGeneralData.SetProperty("U_PFromDate", dtFromDate)
                    Dim intToDays As Integer = 0
                    If Not oRecordSet.EoF Then

                        'For intRow As Integer = 0 To intNoofDays
                        '    oGenDataChild = oGenDataCollection.Add()
                        '    oGenDataChild.SetProperty("U_PrgDate", dtFromDate.AddDays(intRow))
                        '    Dim strIncStatus As String = checkExclude(strCardCode, dtFromDate.AddDays(intRow))

                        '    If strIncStatus = "E" Then
                        '        intRow -= 1
                        '    End If

                        '    oGenDataChild.SetProperty("U_AppStatus", strIncStatus.Trim())
                        '    oGenDataChild.SetProperty("U_Remarks", "Transfer")
                        '    intToDays += 1
                        'Next

                        While intNoofDays > 0
                            dtFromDate.AddDays(intToDays)
                            Dim strIncStatus As String = checkExclude(strCardCode, dtFromDate.AddDays(intToDays))
                            oGenDataChild = oGenDataCollection.Add()
                            oGenDataChild.SetProperty("U_PrgDate", dtFromDate.AddDays(intToDays))
                            If (strIncStatus = "I") Then
                                intNoofDays -= 1
                            End If
                            oGenDataChild.SetProperty("U_AppStatus", strIncStatus.Trim())
                            oGenDataChild.SetProperty("U_Remarks", "Transfer")
                            intToDays += 1
                        End While


                    End If

                    oGeneralDataCollection_R = oGeneralData.Child("Z_CPM6")
                    'Dim intNoofDays As Integer = CInt(strNooDays)
                    Dim dblTBD As Double
                    Dim dblTaxAmount As Double
                    'Dim dblDiscount As Double = oApplication.Utilities.getRecordSetValue("Select Discount From OCRD Where CardCode = '" & strCardCode & "'", "Discount")

                    For intRow As Integer = 0 To 0

                        'Dim dblItemPrice As Double = oApplication.Utilities.GetCustItemPrice(strCardCode, strTProg, System.DateTime.Now.Date)
                        oGenDataChild = oGeneralDataCollection_R.Add()
                        oGenDataChild.SetProperty("U_Fdate", dtFromDate.AddDays(intRow))
                        oGenDataChild.SetProperty("U_Edate", dtFromDate.AddDays(intToDays - 1))
                        oGenDataChild.SetProperty("U_NoofDays", strTNooDays.Trim())

                        Dim dblItemPrice, dblBasePrice As Double
                        Dim strICurrency As String = String.Empty
                        oApplication.Utilities.GetCustItemPrice(strCardCode, _
                                                                strTProg, _
                                                                System.DateTime.Now.Date, dblBasePrice, strICurrency)
                        oGenDataChild.SetProperty("U_Currency", strICurrency)
                        oGenDataChild.SetProperty("U_IPrice", dblBasePrice)
                        oGenDataChild.SetProperty("U_TaxCode", strTaxCode)
                        oGenDataChild.SetProperty("U_IsIReq", "Y")

                        If strICurrency = strLCurrency Then
                            If strICurrency = strDocCurrency Then
                                oGenDataChild.SetProperty("U_Price", dblBasePrice)
                                dblItemPrice = dblBasePrice
                            Else
                                getPrice(strDocCurrency, strICurrency, dblBasePrice, dblItemPrice)
                                oGenDataChild.SetProperty("U_Price", dblItemPrice)
                            End If
                            oGenDataChild.SetProperty("U_LineTotal", (CInt(strNooDays) * (dblItemPrice)))
                            'dblItemPrice = dblBasePrice
                        Else
                            getPrice(strDocCurrency, strICurrency, dblBasePrice, dblItemPrice)
                            oGenDataChild.SetProperty("U_Price", dblItemPrice)
                            oGenDataChild.SetProperty("U_LineTotal", (CInt(strNooDays) * (dblItemPrice)))
                        End If

                        Dim dblTaxRate As Double = 0
                        If strTaxCode <> "" Then
                            strQuery = " Select Rate From OVTG Where Code = '" & strTaxCode & "'"
                            dblTaxRate = oApplication.Utilities.getRecordSetValue(strQuery, "Rate")
                            dblTaxAmount += (dblTaxRate / 100) * (((CInt(strNooDays) * (dblItemPrice))) - ((CInt(strNooDays) * (dblItemPrice)) * (dblDiscount / 100)))
                        End If
                        dblTBD += (CInt(strNooDays) * (dblItemPrice))
                    Next

                    oGeneralData.SetProperty("U_TBDisc", dblTBD)
                    oGeneralData.SetProperty("U_Discount", dblDiscount)
                    Dim dblDisAmount As Double = (CDbl(IIf(dblDiscount.ToString = "", 0, dblDiscount) / 100)) * dblTBD
                    oGeneralData.SetProperty("U_DisAmount", dblDisAmount)
                    oGeneralData.SetProperty("U_TaxAmount", dblTaxAmount)
                    Dim dblTotal As Double = (dblTBD - (dblTBD * (dblDiscount / 100)) + dblTaxAmount)
                    oGeneralData.SetProperty("U_DocTotal", dblTotal)
                    oGeneralData.SetProperty("U_PToDate", dtFromDate.AddDays(intToDays - 1))

                    oGeneralService.Add(oGeneralData)

                    _retVal = True
                ElseIf (strTType = "P") Then

                    'updateProgramNoofDayAndToDate(oForm, strTProgram, CInt(strNooDays))

                    Dim intCode As Integer = getMaxCode("@Z_OCPM", "DocEntry")
                    strCode = String.Format("{0:000000000}", intCode)
                    oGeneralData.SetProperty("U_CardCode", strCardCode)
                    oGeneralData.SetProperty("U_BCardCode", strBCardCode)
                    oGeneralData.SetProperty("U_CardName", strCardName)
                    oGeneralData.SetProperty("U_NoOfDays", strTNooDays)
                    oGeneralData.SetProperty("U_RemDays", strTNooDays)

                    Dim strItemName As String = GetItemName(strTProg)
                    oGeneralData.SetProperty("U_PrgCode", strTProg)
                    oGeneralData.SetProperty("U_PrgName", strItemName)
                    oGeneralData.SetProperty("U_TrnRef", strObjectKey.ToString())

                    Dim strDocCurrency As String = String.Empty

                    If strCurrency <> "##" Then
                        If strCurrency = strLCurrency Then
                            oGeneralData.SetProperty("U_CurSour", "L")
                        ElseIf strCurrency = strSCurrency Then
                            oGeneralData.SetProperty("U_CurSour", "S")
                        End If
                        oGeneralData.SetProperty("U_VenCur", strCurrency)
                        oGeneralData.SetProperty("U_DocCur", strCurrency)
                        strDocCurrency = strCurrency
                        oGeneralData.SetProperty("U_DocRate", GetCurrencyRate(oForm, strCurrency))
                    ElseIf strCurrency = "##" Then
                        oGeneralData.SetProperty("U_CurSour", "C")
                        oGeneralData.SetProperty("U_VenCur", strCurrency)
                        oGeneralData.SetProperty("U_DocCur", strLCurrency)
                        strDocCurrency = strLCurrency
                        oGeneralData.SetProperty("U_DocRate", 1)
                    End If

                    oGenDataCollection = oGeneralData.Child("Z_CPM1")
                    strQuery = "Select U_PFromDate,U_PToDate,U_PrgCode From [@Z_OPGT] T0 "
                    strQuery += " Where T0.DocEntry = '" + strDocEntry.ToString() + "' "
                    oRecordSet.DoQuery(strQuery)

                    'Dim dtFromDate As Date = System.DateTime.Now
                    Dim dtFromDate As Date = dtProgramFrom
                    oGeneralData.SetProperty("U_PFromDate", dtFromDate)
                    Dim intToDays As Integer = 0
                    If Not oRecordSet.EoF Then

                        'Dim intNoofDays As Integer = CInt(strTNooDays)
                        'For intRow As Integer = 0 To intNoofDays
                        '    oGenDataChild = oGenDataCollection.Add()
                        '    oGenDataChild.SetProperty("U_PrgDate", dtFromDate.AddDays(intRow))
                        '    Dim strIncStatus As String = checkExclude(strCardCode, dtFromDate.AddDays(intRow))
                        '    If strIncStatus = "E" Then
                        '        intRow -= 1
                        '    End If
                        '    oGenDataChild.SetProperty("U_AppStatus", strIncStatus.Trim())
                        '    oGenDataChild.SetProperty("U_Remarks", "Transfer")
                        '    intToDays += 1
                        'Next

                        While intNoofDays > 0
                            dtFromDate.AddDays(intToDays)
                            Dim strIncStatus As String = checkExclude(strCardCode, dtFromDate.AddDays(intToDays))
                            oGenDataChild = oGenDataCollection.Add()
                            oGenDataChild.SetProperty("U_PrgDate", dtFromDate.AddDays(intToDays))
                            If (strIncStatus = "I") Then
                                intNoofDays -= 1
                            End If
                            oGenDataChild.SetProperty("U_AppStatus", strIncStatus.Trim())
                            oGenDataChild.SetProperty("U_Remarks", "Transfer")
                            intToDays += 1
                        End While

                    End If

                    oGeneralDataCollection_R = oGeneralData.Child("Z_CPM6")
                    'Dim intNoofDays As Integer = CInt(strNooDays)
                    Dim dblTBD As Double
                    Dim dblTaxAmount As Double

                    For intRow As Integer = 0 To 0
                        'Dim dblItemPrice As Double = oApplication.Utilities.GetCustItemPrice(strCardCode, strTProg, System.DateTime.Now.Date)
                        oGenDataChild = oGeneralDataCollection_R.Add()
                        oGenDataChild.SetProperty("U_Fdate", dtFromDate.AddDays(intRow))
                        ' Dim strIncStatus As String = checkExclude(strCardCode, dtFromDate.AddDays(intRow))
                        'Dim strProgramToDate As String = getProgramToDate(oForm, strCardCode, dtFromDate.ToString("yyyyMMdd"), strNooDays)
                        oGenDataChild.SetProperty("U_Edate", dtFromDate.AddDays(intToDays - 1))
                        oGenDataChild.SetProperty("U_NoofDays", strNooDays.Trim())

                        Dim dblItemPrice, dblBasePrice As Double
                        Dim strICurrency As String = String.Empty
                        oApplication.Utilities.GetCustItemPrice(strCardCode, _
                                                                strTProg, _
                                                                System.DateTime.Now.Date, dblBasePrice, strICurrency)
                        oGenDataChild.SetProperty("U_Currency", strICurrency)
                        oGenDataChild.SetProperty("U_IPrice", dblBasePrice)
                        oGenDataChild.SetProperty("U_TaxCode", strTaxCode)
                        oGenDataChild.SetProperty("U_IsIReq", "Y")

                        If strICurrency = strLCurrency Then
                            If strICurrency = strDocCurrency Then
                                oGenDataChild.SetProperty("U_Price", dblBasePrice)
                                dblItemPrice = dblBasePrice
                            Else
                                getPrice(strDocCurrency, strICurrency, dblBasePrice, dblItemPrice)
                                oGenDataChild.SetProperty("U_Price", dblItemPrice)
                            End If
                            oGenDataChild.SetProperty("U_LineTotal", (CInt(strNooDays) * (dblItemPrice)))
                        Else
                            getPrice(strDocCurrency, strICurrency, dblBasePrice, dblItemPrice)
                            oGenDataChild.SetProperty("U_Price", dblItemPrice)
                            oGenDataChild.SetProperty("U_LineTotal", (CInt(strNooDays) * (dblItemPrice)))
                        End If

                        'oGenDataChild.SetProperty("U_Price", dblItemPrice)
                        'oGenDataChild.SetProperty("U_LineTotal", (CInt(strNooDays) * (dblItemPrice)))

                        Dim dblTaxRate As Double = 0
                        If strTaxCode <> "" Then
                            strQuery = " Select Rate From OVTG Where Code = '" & strTaxCode & "'"
                            dblTaxRate = oApplication.Utilities.getRecordSetValue(strQuery, "Rate")
                            dblTaxAmount += (dblTaxRate / 100) * (((CInt(strNooDays) * (dblItemPrice))) - ((CInt(strNooDays) * (dblItemPrice)) * (dblDiscount / 100)))
                        End If

                        dblTBD += (CInt(strNooDays) * (dblItemPrice))
                    Next


                    oGeneralData.SetProperty("U_TBDisc", dblTBD)
                    oGeneralData.SetProperty("U_Discount", dblDiscount)
                    Dim dblDisAmount As Double = (CDbl(IIf(dblDiscount.ToString = "", 0, dblDiscount) / 100)) * dblTBD
                    oGeneralData.SetProperty("U_DisAmount", dblDisAmount)
                    oGeneralData.SetProperty("U_TaxAmount", dblTaxAmount)
                    Dim dblTotal As Double = (dblTBD - (dblTBD * (dblDiscount / 100)) + dblTaxAmount)
                    oGeneralData.SetProperty("U_DocTotal", dblTotal)
                    oGeneralData.SetProperty("U_PToDate", dtFromDate.AddDays(intToDays - 1))
                    oGeneralService.Add(oGeneralData)

                    _retVal = True
                End If
            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
        Return _retVal
    End Function

    Private Sub getPrice(ByVal strDCurrency As String, strICurrency As String, ByVal dblBasePrice As String, ByRef dblPrice As Double)
        Try
            Dim oExRecordSet As SAPbobsCOM.Recordset
            Dim dblRExRate, dblAExRate As Double
            Dim strQuery As String = String.Empty
            'Dim oRecordSet As SAPbobsCOM.Recordset

            oExRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim dblLocalCurrency As String = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency

            If strDCurrency <> dblLocalCurrency Then
                strQuery = "Select Rate From ORTT Where Currency = '" + strDCurrency + "' And Convert(VarChar(8),RateDate,112) = Convert(VarChar(8),GetDate(),112)"
                oExRecordSet.DoQuery(strQuery)
                If Not oExRecordSet.EoF Then
                    dblRExRate = oExRecordSet.Fields.Item("Rate").Value
                    If strICurrency = dblLocalCurrency Then
                        dblPrice = (dblBasePrice / dblRExRate)
                    Else
                        strQuery = "Select isnull(Rate,1) 'Rate' From ORTT Where Currency = '" + strICurrency + "' And Convert(VarChar(8),RateDate,112) = Convert(VarChar(8),GetDate(),112)"
                        oExRecordSet.DoQuery(strQuery)
                        If Not oExRecordSet.EoF Then
                            dblAExRate = oExRecordSet.Fields.Item("Rate").Value
                            dblPrice = ((dblBasePrice * dblAExRate) / dblRExRate)
                        End If
                    End If
                End If
            ElseIf strDCurrency = dblLocalCurrency Then
                strQuery = "Select Rate From ORTT Where Currency = '" + strICurrency + "' And Convert(VarChar(8),RateDate,112) = Convert(VarChar(8),GetDate(),112)"
                oExRecordSet.DoQuery(strQuery)
                If Not oExRecordSet.EoF Then
                    dblAExRate = oExRecordSet.Fields.Item("Rate").Value
                    dblPrice = (dblBasePrice * dblAExRate)
                End If
            End If

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    'Private Function getItemName(ByVal strItemCode As String) As String
    '    Dim _retVal As String = String.Empty
    '    Dim oRecordSet As SAPbobsCOM.Recordset
    '    Try
    '        oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
    '        oRecordSet.DoQuery("Select ItemName From OITM Where ItemCode = '" + strItemCode + "'")
    '        If Not oRecordSet.EoF Then
    '            _retVal = oRecordSet.Fields.Item(0).Value
    '        End If
    '        Return _retVal
    '    Catch ex As Exception 
    ' oApplication.Log.Trace_DIET_AddOn_Error(ex)
    '        Throw ex 
    ''oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
    '    End Try
    'End Function

    Public Function ValidateRemoveSetup(ByVal aCode As String, ByVal aChoice As String) As Boolean
        Dim oREC As SAPbobsCOM.Recordset
        Dim strString As String = ""
        oREC = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Select Case aChoice
            Case "DisLike"
                strString = "Select * from [@Z_CPR1] where ""U_DLikeItem""='" & aCode & "'"
            Case "Medical"
                strString = "Select * from [@Z_CPR2] where ""U_MSCode""='" & aCode & "'"
            Case "Calories"
                strString = "Select * from [@Z_OCPR] where ""U_CPAdj""='" & aCode & "'"
            Case "CaloriesPlan"
                strString = "Select * from [@Z_OCPR] where ""U_CPCode""='" & aCode & "'"
        End Select
        oREC.DoQuery(strString)
        If oREC.RecordCount > 0 Then
            Return False
        End If
        Return True
    End Function

    Public Function updateCustomerProgram(ByVal strObjectKey As String) As Boolean
        Dim _retVal As Boolean = False
        Dim strQuery As String = String.Empty
        Dim oDoc As SAPbobsCOM.Documents = Nothing

        Dim oRecordSet_S As SAPbobsCOM.Recordset
        Dim oRecordSet As SAPbobsCOM.Recordset
        Try
            oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)

            If oDoc.Browser.GetByKeys(strObjectKey) Then
                strQuery = "Select Distinct U_ProgramID From [DLN1] T0 JOIN [ODLN] T1 ON T0.DocEntry = T1.DocEntry Where T0.DocEntry = '" + oDoc.DocEntry.ToString() + "'"
                strQuery += " And ISNULL(T0.U_ProgramID,'') <> '' And ISNULL(T1.CANCELED,'N') = 'N' "
                'strQuery = "Exec [PROCON_UPDATEPROGRAMDAYS_u] '" + oDoc.DocEntry.ToString() + "'"
                oRecordSet_S = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet_S.DoQuery(strQuery)
                If Not oRecordSet_S.EoF Then
                    'Pre Sales Order
                    While Not oRecordSet_S.EoF
                        Dim strProgramID As String = oRecordSet_S.Fields.Item("U_ProgramID").Value
                        oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                        strQuery = " Select Distinct T0.U_DelDate From [DLN1] T0 JOIN [ODLN] T1 ON T0.DocEntry = T1.DocEntry Where T0.U_ProgramID = '" + strProgramID + "'"
                        strQuery += " And ISNULL(T0.U_ProgramID,'') <> '' And ISNULL(T1.CANCELED,'N') = 'N' "
                        oRecordSet.DoQuery(strQuery)
                        If Not oRecordSet.EoF Then
                            Dim intNoOfDays As Integer = CInt(oRecordSet.RecordCount)
                            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                            strQuery = " Select (ISNULL(T0.U_NoOfDays,0) + ISNULL(T0.U_FreeDays,0)) "
                            strQuery += " From [@Z_OCPM] T0 "
                            strQuery += " Where DocEntry = '" + strProgramID + "'"
                            oRecordSet.DoQuery(strQuery)
                            If Not oRecordSet.EoF Then
                                Dim intRemaining As Integer = CInt(oRecordSet.Fields.Item(0).Value) - intNoOfDays
                                If intRemaining > 0 Then
                                    strQuery = " Update T0 Set T0.U_RemDays = T0.U_NoOfDays - '" + intNoOfDays.ToString() + "' "
                                    strQuery += " From [@Z_OCPM] T0 "
                                    strQuery += " Where DocEntry = '" + strProgramID + "'"
                                    oRecordSet.DoQuery(strQuery)
                                Else
                                    strQuery = " Update T0 Set T0.U_RemDays = T0.U_NoOfDays - '" + intNoOfDays.ToString() + "' "
                                    strQuery += " U_DocStatus = 'C' "
                                    strQuery += " From [@Z_OCPM] T0 "
                                    strQuery += " Where DocEntry = '" + strProgramID + "'"
                                    oRecordSet.DoQuery(strQuery)
                                End If
                                _retVal = True
                            End If
                        End If
                        oRecordSet_S.MoveNext()
                    End While
                End If

            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function updateCustomerProfileMesurements(ByVal strObjectKey As String) As Boolean
        Dim _retVal As Boolean = False
        Dim strQuery As String = String.Empty

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim oGeneralDataCollection As SAPbobsCOM.GeneralDataCollection
        Dim oChildData As SAPbobsCOM.GeneralData
        oCompanyService = oApplication.Company.GetCompanyService()

        Dim oDoc As SAPbobsCOM.SalesOpportunities = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset
        Try
            oGeneralService = oCompanyService.GetGeneralService("Z_OCPR")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)

            oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesOpportunities)

            'Dim oXmlDoc As System.Xml.XmlDocument = New Xml.XmlDocument()
            'oXmlDoc.LoadXml(strObjectKey)

            'Try
            '    strOppID = oXmlDoc.SelectSingleNode("/SalesOpportunityParams/OpprId").InnerText
            'Catch ex As Exception 
            'oApplication.Log.Trace_DIET_AddOn_Error(ex)
            '    strOppID = oXmlDoc.SelectSingleNode("/SalesOpportunityParams/SequentialNo").InnerText
            'End Try

            Dim strOppID As String
            If oDoc.Browser.GetByKeys(strObjectKey) Then

                strOppID = oDoc.SequentialNo

                If oDoc.GetByKey(CInt(strOppID)) Then
                    strQuery = " Select DocEntry From [@Z_OCPR] Where U_CardCode = '" + oDoc.CardCode + "'"
                    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery(strQuery)
                    If Not oRecordSet.EoF Then
                        oGeneralParams.SetProperty("DocEntry", oRecordSet.Fields.Item(0).Value)
                        oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                        strQuery = " Select "
                        strQuery += " OpprId,Line,OpenDate,Convert(VarChar(12),OpenDate,101) As 'OpenDate1',U_Duration,U_Dietitian1,U_Dietitian2, "
                        strQuery += " U_Weight,U_Breast,U_Height,U_UnderBreast,U_Hip,U_Fat,U_BMI,U_Arm,U_Bust,U_Waist,U_Thigh,U_Neck, "
                        strQuery += " U_WH,U_24RCall,U_BC,U_PAD,U_PA,ISNULL(U_Smoking,'N') As U_Smoking "
                        strQuery += " From OPR1 Where OpprID = '" + strOppID + "'"
                        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery(strQuery)
                        If Not oRecordSet.EoF Then
                            oGeneralDataCollection = oGeneralData.Child("Z_CPR3")
                            Dim intRow As Integer = 0
                            Dim blnRecordAdded As Boolean = False

                            While Not oRecordSet.EoF
                                'MessageBox.Show(oGeneralDataCollection.Count)
                                'oChildData = IIf(intRow < oGeneralDataCollection.Count, oGeneralDataCollection.Item(intRow), oGeneralDataCollection.Add())

                                Dim oRecordExist As SAPbobsCOM.Recordset
                                oRecordExist = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                strQuery = " Select DocEntry From [@Z_CPR3] Where U_OpprId = '" + oRecordSet.Fields.Item("OpprId").Value.ToString() + "'"
                                strQuery += " And U_Line = '" + oRecordSet.Fields.Item("Line").Value.ToString() + "'"
                                oRecordExist.DoQuery(strQuery)
                                If oRecordExist.EoF Then

                                    blnRecordAdded = True
                                    oChildData = oGeneralDataCollection.Add()
                                    oChildData.SetProperty("U_OpprId", oRecordSet.Fields.Item("OpprId").Value.ToString())
                                    oChildData.SetProperty("U_Line", oRecordSet.Fields.Item("Line").Value.ToString())

                                    oChildData.SetProperty("U_VisitDate", oRecordSet.Fields.Item("OpenDate").Value)
                                    oChildData.SetProperty("U_Duration", oRecordSet.Fields.Item("U_Duration").Value.ToString())
                                    oChildData.SetProperty("U_Dietitian1", oRecordSet.Fields.Item("U_Dietitian1").Value.ToString())
                                    oChildData.SetProperty("U_Dietitian2", oRecordSet.Fields.Item("U_Dietitian2").Value.ToString())

                                    oChildData.SetProperty("U_Weight", oRecordSet.Fields.Item("U_Weight").Value.ToString())
                                    oChildData.SetProperty("U_Breast", oRecordSet.Fields.Item("U_Breast").Value.ToString())
                                    oChildData.SetProperty("U_Height", oRecordSet.Fields.Item("U_Height").Value.ToString())
                                    oChildData.SetProperty("U_UnderBreast", oRecordSet.Fields.Item("U_UnderBreast").Value.ToString())
                                    oChildData.SetProperty("U_Hip", oRecordSet.Fields.Item("U_Hip").Value.ToString())
                                    oChildData.SetProperty("U_Fat", oRecordSet.Fields.Item("U_Fat").Value.ToString())
                                    oChildData.SetProperty("U_BMI", oRecordSet.Fields.Item("U_BMI").Value.ToString())

                                    oChildData.SetProperty("U_Arm", oRecordSet.Fields.Item("U_Arm").Value.ToString())
                                    oChildData.SetProperty("U_Bust", oRecordSet.Fields.Item("U_Bust").Value.ToString())
                                    oChildData.SetProperty("U_Waist", oRecordSet.Fields.Item("U_Waist").Value.ToString())
                                    oChildData.SetProperty("U_Thigh", oRecordSet.Fields.Item("U_BMI").Value.ToString())
                                    oChildData.SetProperty("U_Neck", oRecordSet.Fields.Item("U_BMI").Value.ToString())

                                    oChildData.SetProperty("U_WH", oRecordSet.Fields.Item("U_WH").Value.ToString())
                                    oChildData.SetProperty("U_24RCall", oRecordSet.Fields.Item("U_24RCall").Value.ToString())
                                    oChildData.SetProperty("U_BC", oRecordSet.Fields.Item("U_BC").Value.ToString())
                                    oChildData.SetProperty("U_PAD", oRecordSet.Fields.Item("U_PAD").Value.ToString())
                                    oChildData.SetProperty("U_PA", oRecordSet.Fields.Item("U_PA").Value.ToString())
                                    oChildData.SetProperty("U_Smoking", oRecordSet.Fields.Item("U_Smoking").Value.ToString())

                                Else

                                    strQuery = " Update T0 SET "
                                    strQuery += " U_VisitDate = '" + oRecordSet.Fields.Item("OpenDate1").Value + "'"
                                    strQuery += " ,U_Duration = '" + oRecordSet.Fields.Item("U_Duration").Value + "'"
                                    strQuery += " ,U_Dietitian1 = '" + oRecordSet.Fields.Item("U_Dietitian1").Value + "'"
                                    strQuery += " ,U_Dietitian2 = '" + oRecordSet.Fields.Item("U_Dietitian2").Value + "'"

                                    strQuery += " ,U_Weight = '" + oRecordSet.Fields.Item("U_Weight").Value + "'"
                                    strQuery += " ,U_Breast = '" + oRecordSet.Fields.Item("U_Breast").Value + "'"
                                    strQuery += " ,U_Height = '" + oRecordSet.Fields.Item("U_Height").Value + "'"
                                    strQuery += " ,U_UnderBreast = '" + oRecordSet.Fields.Item("U_UnderBreast").Value + "'"
                                    strQuery += " ,U_Hip = '" + oRecordSet.Fields.Item("U_Hip").Value + "'"
                                    strQuery += " ,U_Fat = '" + oRecordSet.Fields.Item("U_Fat").Value + "'"
                                    strQuery += " ,U_BMI = '" + oRecordSet.Fields.Item("U_BMI").Value + "'"

                                    strQuery += " ,U_Arm = '" + oRecordSet.Fields.Item("U_Arm").Value + "'"
                                    strQuery += " ,U_Bust = '" + oRecordSet.Fields.Item("U_Bust").Value + "'"
                                    strQuery += " ,U_Waist = '" + oRecordSet.Fields.Item("U_Waist").Value + "'"
                                    strQuery += " ,U_Thigh = '" + oRecordSet.Fields.Item("U_Thigh").Value + "'"
                                    strQuery += " ,U_Neck = '" + oRecordSet.Fields.Item("U_Neck").Value + "'"

                                    strQuery += " ,U_WH = '" + oRecordSet.Fields.Item("U_WH").Value + "'"
                                    strQuery += " ,U_24RCall = '" + oRecordSet.Fields.Item("U_24RCall").Value + "'"
                                    strQuery += " ,U_BC = '" + oRecordSet.Fields.Item("U_BC").Value + "'"
                                    strQuery += " ,U_PAD = '" + oRecordSet.Fields.Item("U_PAD").Value + "'"
                                    strQuery += " ,U_PA = '" + oRecordSet.Fields.Item("U_PA").Value + "'"
                                    strQuery += " ,U_Smoking = '" + oRecordSet.Fields.Item("U_Smoking").Value + "'"

                                    strQuery += "  From [@Z_CPR3] T0 Where T0.U_OpprId = '" + oRecordSet.Fields.Item("OpprId").Value.ToString() + "'"
                                    strQuery += " And T0.U_Line = '" + oRecordSet.Fields.Item("Line").Value.ToString() + "'"
                                    oRecordExist.DoQuery(strQuery)
                                End If

                                intRow += 1
                                oRecordSet.MoveNext()
                            End While
                            If blnRecordAdded Then
                                oGeneralService.Update(oGeneralData)
                            End If
                        End If
                    End If
                End If
            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function getDateDiff_PreSales(ByVal oForm As SAPbouiCOM.Form, strCardCode As String, ByVal strFromDate As String, ByVal strToDate As String _
                                         , strType As String, ByVal strstrRefKey As String) As Integer
        Dim _retVal As Integer
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim strQuery As String = String.Empty
        Try
            Dim dtFromDate As Date = strFromDate.Substring(0, 4) + "-" + strFromDate.Substring(4, 2) + "-" + strFromDate.Substring(6, 2)
            Dim dtToDate As Date = strToDate.Substring(0, 4) + "-" + strToDate.Substring(4, 2) + "-" + strToDate.Substring(6, 2)
            Dim days As Integer = 0
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            If strType = "I" Then
                'Madhu Modified this for Phase II On 20150708.
                'strQuery = " Select DocEntry From [@Z_OCPM] Where U_InvRef = '" + strstrRefKey + "'"
                strQuery = " Select T0.DocEntry From [@Z_OCPM] T0 LEFT OUTER JOIN [@Z_CPM6] T1 On T0.DocEntry = T1.DocEntry Where ISNULL(T0.U_InvRef,T1.U_InvRef) = '" + strstrRefKey.Trim() + "'"
            ElseIf strType = "T" Then
                strQuery = " Select DocEntry From [@Z_OCPM] Where U_TrnRef = '" + strstrRefKey.Trim() + "'"
            ElseIf strType = "P" Then
                strQuery = " Select DocEntry From [@Z_OCPM] Where DocEntry = '" + strstrRefKey.Trim() + "'"
            End If
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                Dim strProRef As String = oRecordSet.Fields.Item(0).Value
                Dim intNoofDays As Integer = DateDiff(DateInterval.Day, dtFromDate, dtToDate) ' CInt(strToDate) - CInt(strFromDate)
                For intRow As Integer = 0 To intNoofDays
                    Dim strIncStatus As String = checkExclude(strCardCode, dtFromDate.AddDays(intRow))
                    If strIncStatus = "I" Then
                        _retVal += 1
                    End If
                Next
            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function updateCustomerProgram(ByVal oForm As SAPbouiCOM.Form, strCardCode As String, strFromDate As String, _
                                          strToDate As String, strType As String, ByVal strstrRefKey As String) As Boolean
        Dim _retVal As Boolean = False
        Dim strQuery As String = String.Empty

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim oGeneralDataCollection As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralDataCollection_Q As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralDataCollection_T As SAPbobsCOM.GeneralDataCollection
        Dim oChildData As SAPbobsCOM.GeneralData
        Dim oChildData_Q As SAPbobsCOM.GeneralData
        Dim oChildData_T As SAPbobsCOM.GeneralData
        oCompanyService = oApplication.Company.GetCompanyService()

        Dim oDoc As SAPbobsCOM.SalesOpportunities = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset
        Try
            oGeneralService = oCompanyService.GetGeneralService("Z_OCPM")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)

            If strType = "I" Then
                'Madhu Modified this for Phase II On 20150708.
                'strQuery = " Select DocEntry From [@Z_OCPM] Where U_InvRef = '" + strstrRefKey + "'"
                strQuery = " Select T0.DocEntry From [@Z_OCPM] T0 LEFT OUTER JOIN [@Z_CPM6] T1 On T0.DocEntry = T1.DocEntry Where ISNULL(T0.U_InvRef,T1.U_InvRef) = '" + strstrRefKey.Trim() + "'"
            ElseIf strType = "T" Then
                strQuery = " Select DocEntry From [@Z_OCPM] Where U_TrnRef = '" + strstrRefKey + "'"
            ElseIf strType = "P" Then
                strQuery = " Select DocEntry From [@Z_OCPM] Where DocEntry = '" + strstrRefKey + "'"
            End If
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)

            If Not oRecordSet.EoF Then
                Dim dtFromDate As Date = strFromDate.Substring(0, 4) + "-" + strFromDate.Substring(4, 2) + "-" + strFromDate.Substring(6, 2)
                Dim dtToDate As Date = strToDate.Substring(0, 4) + "-" + strToDate.Substring(4, 2) + "-" + strToDate.Substring(6, 2)
                Dim intNoofDays As Integer = DateDiff(DateInterval.Day, dtFromDate, dtToDate) + 1 ' CInt(strToDate) - CInt(strFromDate) + 1
                Dim strProRef As String = oRecordSet.Fields.Item(0).Value
                oGeneralParams.SetProperty("DocEntry", oRecordSet.Fields.Item(0).Value)
                oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                oGeneralDataCollection = oGeneralData.Child("Z_CPM1")
                oGeneralDataCollection_Q = oGeneralData.Child("Z_CPM4")
                oGeneralDataCollection_T = oGeneralData.Child("Z_CPM5")
                For intRow As Integer = 0 To intNoofDays
                    strQuery = "Select DocEntry,LineId,U_AppStatus From [@Z_CPM1] Where DocEntry = '" + strProRef + "' And Convert(VarChar(8),U_PrgDate,112) = '" + dtFromDate.AddDays(intRow).ToString("yyyyMMdd") + "'"
                    oRecordSet.DoQuery(strQuery)
                    If oRecordSet.EoF Then
                        oChildData = oGeneralDataCollection.Add()
                        oChildData.SetProperty("U_PrgDate", dtFromDate.AddDays(intRow))
                        Dim strIncStatus As String = checkExclude(strCardCode, dtFromDate.AddDays(intRow))
                        oChildData.SetProperty("U_AppStatus", strIncStatus.Trim())

                        oChildData_Q = oGeneralDataCollection_Q.Add()
                        oChildData_Q.SetProperty("U_PrgDate", dtFromDate.AddDays(intRow))

                        oChildData_T = oGeneralDataCollection_T.Add()
                        oChildData_T.SetProperty("U_PrgDate", dtFromDate.AddDays(intRow))

                    Else
                        Dim strIncStatus As String = checkExclude(strCardCode, dtFromDate.AddDays(intRow))
                        If oRecordSet.Fields.Item("U_AppStatus").Value.ToString() = "I" Then
                            If strIncStatus = "E" Then
                                Dim intLine As Integer = CInt(oRecordSet.Fields.Item("LineId").Value.ToString())
                                oChildData = oGeneralDataCollection.Item(intLine - 1)
                                oChildData.SetProperty("U_AppStatus", strIncStatus.Trim())
                                'strQuery = " Update [@Z_CPM1] Set U_AppStatus = 'E' "
                                'strQuery = " Where DocEntry = '" + oRecordSet.Fields.Item("U_AppStatus").Value.ToString() + "'"
                                'strQuery += " And LineId = '" + oRecordSet.Fields.Item("LineId").Value.ToString() + "'"
                                'oRecordSet.DoQuery(strQuery)
                            End If
                        ElseIf oRecordSet.Fields.Item("U_AppStatus").Value.ToString() = "E" Then
                            If strIncStatus = "I" Then
                                Dim intLine As Integer = CInt(oRecordSet.Fields.Item("LineId").Value.ToString())
                                oChildData = oGeneralDataCollection.Item(intLine - 1)
                                oChildData.SetProperty("U_AppStatus", strIncStatus.Trim())
                            End If
                        End If
                    End If
                Next
            End If
            oGeneralService.Update(oGeneralData)
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function updateCustomerProgramInProgram(ByVal strDocEntry As String) As Boolean
        Dim _retVal As Boolean = False
        Dim strQuery As String = String.Empty

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim oGeneralDataCollection As SAPbobsCOM.GeneralDataCollection
        Dim oChildData As SAPbobsCOM.GeneralData
        oCompanyService = oApplication.Company.GetCompanyService()

        Dim oDoc As SAPbobsCOM.SalesOpportunities = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset
        Try
            oGeneralService = oCompanyService.GetGeneralService("Z_OCPM")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)

            strQuery = " Select DocEntry,U_CardCode,Convert(VarChar(8),U_PFromDate,112)As 'PF',Convert(VarChar(8),U_PToDate,112)As 'PT' From [@Z_OCPM] Where DocEntry = '" + strDocEntry + "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)

            If Not oRecordSet.EoF Then
                Dim strFromDate As String = oRecordSet.Fields.Item("PF").Value
                Dim strToDate As String = oRecordSet.Fields.Item("PT").Value
                Dim strCardCode As String = oRecordSet.Fields.Item("U_CardCode").Value
                Dim dtFromDate As Date = strFromDate.Substring(0, 4) + "-" + strFromDate.Substring(4, 2) + "-" + strFromDate.Substring(6, 2)
                Dim dtToDate As Date = strToDate.Substring(0, 4) + "-" + strToDate.Substring(4, 2) + "-" + strToDate.Substring(6, 2)
                Dim intNoofDays As Integer = DateDiff(DateInterval.Day, dtFromDate, dtToDate) + 1 ' CInt(strToDate) - CInt(strFromDate) + 1
                Dim strProRef As String = oRecordSet.Fields.Item(0).Value
                oGeneralParams.SetProperty("DocEntry", oRecordSet.Fields.Item(0).Value)
                oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                oGeneralDataCollection = oGeneralData.Child("Z_CPM1")
                For intRow As Integer = 0 To intNoofDays - 1
                    strQuery = "Select DocEntry,LineId,VisOrder,U_AppStatus From [@Z_CPM1] Where DocEntry = '" + strProRef + "' And Convert(VarChar(8),U_PrgDate,112) = '" + dtFromDate.AddDays(intRow).ToString("yyyyMMdd") + "'"
                    oRecordSet.DoQuery(strQuery)
                    If oRecordSet.EoF Then
                        oChildData = oGeneralDataCollection.Add()
                        oChildData.SetProperty("U_PrgDate", dtFromDate.AddDays(intRow))
                        Dim strIncStatus As String = checkExclude(strCardCode, dtFromDate.AddDays(intRow))
                        oChildData.SetProperty("U_AppStatus", strIncStatus.Trim())
                        If strIncStatus = "E" Then
                            Dim strCPRef As String = oApplication.Utilities.getRecordSetValueString("Select DocEntry From [@Z_OCPR] Where U_CardCode = '" & strCardCode & "'", "DocEntry")
                            addExcludeDocumentRow(strCardCode, strCPRef, dtFromDate.AddDays(intRow).ToString("yyyyMMdd"))
                        End If
                    Else
                        Dim strIncStatus As String = checkExclude(strCardCode, dtFromDate.AddDays(intRow))
                        If strIncStatus = "E" Then
                            Dim strCPRef As String = oApplication.Utilities.getRecordSetValueString("Select DocEntry From [@Z_OCPR] Where U_CardCode = '" & strCardCode & "'", "DocEntry")
                            addExcludeDocumentRow(strCardCode, strCPRef, dtFromDate.AddDays(intRow).ToString("yyyyMMdd"))
                        End If
                        'oChildData = oGeneralDataCollection.Item(oRecordSet.Fields.Item("VisOrder").Value)
                        'oGeneralDataCollection = oGeneralData.Child("Z_CPM1")
                        'oChildData = oGeneralDataCollection.Item(1)
                        'oChildData.SetProperty("U_AppStatus", strIncStatus.Trim())
                    End If
                Next
            End If

            oGeneralService.Update(oGeneralData)
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function updateCustomerProgramInProgramBasedOnLine(ByVal strDocEntry As String, ByVal strDocLine As String) As Boolean
        Dim _retVal As Boolean = False
        Dim strQuery As String = String.Empty

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim oGeneralDataCollection As SAPbobsCOM.GeneralDataCollection
        Dim oChildData As SAPbobsCOM.GeneralData
        oCompanyService = oApplication.Company.GetCompanyService()

        Dim oDoc As SAPbobsCOM.SalesOpportunities = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset
        Try
            oGeneralService = oCompanyService.GetGeneralService("Z_OCPM")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)

            strQuery = " Select DocEntry,U_CardCode,Convert(VarChar(8),U_PFromDate,112)As 'PF',Convert(VarChar(8),U_PToDate,112)As 'PT' From [@Z_OCPM] Where DocEntry = '" + strDocEntry + "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)

            If Not oRecordSet.EoF Then
                Dim strFromDate As String = oRecordSet.Fields.Item("PF").Value
                Dim strToDate As String = oRecordSet.Fields.Item("PT").Value
                Dim strCardCode As String = oRecordSet.Fields.Item("U_CardCode").Value
                Dim dtFromDate As Date = strFromDate.Substring(0, 4) + "-" + strFromDate.Substring(4, 2) + "-" + strFromDate.Substring(6, 2)
                Dim dtToDate As Date = strToDate.Substring(0, 4) + "-" + strToDate.Substring(4, 2) + "-" + strToDate.Substring(6, 2)
                Dim intNoofDays As Integer = DateDiff(DateInterval.Day, dtFromDate, dtToDate) + 1 ' CInt(strToDate) - CInt(strFromDate) + 1
                Dim strProRef As String = oRecordSet.Fields.Item(0).Value
                oGeneralParams.SetProperty("DocEntry", oRecordSet.Fields.Item(0).Value)
                oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                oGeneralDataCollection = oGeneralData.Child("Z_CPM1")
                For intRow As Integer = 0 To intNoofDays
                    strQuery = "Select DocEntry,LineId,U_AppStatus From [@Z_CPM1] Where DocEntry = '" + strProRef + "' And Convert(VarChar(8),U_PrgDate,112) = '" + dtFromDate.AddDays(intRow).ToString("yyyyMMdd") + "'"
                    oRecordSet.DoQuery(strQuery)
                    If oRecordSet.EoF Then
                        oChildData = oGeneralDataCollection.Add()
                        oChildData.SetProperty("U_PrgDate", dtFromDate.AddDays(intRow))
                        Dim strIncStatus As String = checkExclude(strCardCode, dtFromDate.AddDays(intRow))
                        oChildData.SetProperty("U_AppStatus", strIncStatus.Trim())
                    End If
                Next
            End If
            oGeneralService.Update(oGeneralData)
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function validate_Program(ByVal oForm As SAPbouiCOM.Form, ByRef strMessage As String) As Boolean
        Try
            Dim _retVal As Boolean = True
            Dim strItem As String
            Dim strPFDt As String
            Dim strPTDt As String
            Dim oMatrix As SAPbouiCOM.Matrix
            oMatrix = oForm.Items.Item("38").Specific

            For index As Integer = 1 To oMatrix.RowCount
                strItem = oMatrix.Columns.Item("1").Cells.Item(index).Specific.value
                strPFDt = oMatrix.Columns.Item("U_Fdate").Cells.Item(index).Specific.value
                strPTDt = oMatrix.Columns.Item("U_Edate").Cells.Item(index).Specific.value

                If strItem <> "" Then
                    Dim strQuery As String
                    Dim oRecordSet As SAPbobsCOM.Recordset
                    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    strQuery = " Select T0.ItemCode From OITM T0 JOIN OITB T1 On T0.ItmsGrpCod = T1.ItmsGrpCod "
                    strQuery += " Where T1.U_Program = 'Y' "
                    strQuery += " And T0.ItemCode = '" + strItem + "' "
                    oRecordSet.DoQuery(strQuery)
                    If Not oRecordSet.EoF Then
                        If strPFDt = "" Or strPTDt = "" Then
                            strMessage = "Program From Or Program To For Item : " + strItem + " Not Specified..."
                            _retVal = False
                            Exit For
                        ElseIf CInt(strPFDt) > CInt(strPTDt) Then
                            strMessage = "Program From Should be lesser than Program To For Item : " + strItem + ""
                            _retVal = False
                            Exit For
                        End If
                    End If
                End If
                If Not _retVal Then
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

    Public Function getProgramToDate(ByVal oForm As SAPbouiCOM.Form, ByVal strCardCode As String _
                                 , ByVal oDtProgramF As String, ByVal strNooDays As String) As String
        Try
            Dim dtFromDate As Date = oDtProgramF.Substring(0, 4) + "-" + oDtProgramF.Substring(4, 2) + "-" + oDtProgramF.Substring(6, 2)
            Dim intToDays As Integer = 0
            Dim intNoofDays As Integer = CInt(strNooDays)
            While intNoofDays > 0
                dtFromDate.AddDays(intToDays)
                Dim strIncStatus As String = checkExclude(strCardCode, dtFromDate.AddDays(intToDays))
                If (strIncStatus = "I") Then
                    intNoofDays -= 1
                End If
                intToDays += 1
            End While
            Return dtFromDate.AddDays(intToDays - 1).ToString("yyyyMMdd")
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function UpdateONOFFStatus(ByVal oForm As SAPbouiCOM.Form, ByVal strDocEntry As String) As Boolean
        Dim _retVal As Boolean = False
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim oRecordSet_U As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet_U = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            strQuery = " Select DocEntry,U_CardCode "
            strQuery += " From [@Z_OCPM] Where U_CardCode = (Select U_CardCode From [@Z_OCPR] Where DocEntry = '" + strDocEntry + "')"
            oRecordSet.DoQuery(strQuery)
            If oRecordSet.EoF Then
                strQuery = "Update [@Z_OCPR] SET "
                strQuery += " U_ONOFFSTA = 'F' "
                strQuery += " Where DocEntry = '" + strDocEntry + "'"
                oRecordSet_U.DoQuery(strQuery)
            End If



            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function UpdateReRunStatus(ByVal oForm As SAPbouiCOM.Form, ByVal strDocEntry As String) As Boolean
        Dim _retVal As Boolean = False
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim oRecordSet_U As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet_U = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            strQuery = " Select DocEntry,U_CardCode "
            strQuery += " From [@Z_OCPM] Where U_CardCode = (Select U_CardCode From [@Z_OCPR] Where DocEntry = '" + strDocEntry + "')"
            strQuery += " And U_ReRun = 'Y' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    Dim strDocEntry1 As String = oRecordSet.Fields.Item(0).Value
                    strQuery = " Update [@Z_OCPM] SET "
                    strQuery += " U_ReRun = 'N' "
                    strQuery += " Where DocEntry = '" + strDocEntry1 + "'"
                    oRecordSet_U.DoQuery(strQuery)
                    oRecordSet.MoveNext()
                End While
            End If

            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function UpdateProgramToDate(ByVal oForm As SAPbouiCOM.Form, ByVal strDocEntry As String, Optional ByVal strPRef As String = "") As Boolean
        Dim _retVal As Boolean = False
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim oRecordSet_L As SAPbobsCOM.Recordset
            Dim oRecordSet_U As SAPbobsCOM.Recordset
            Dim oRecordSet_A As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet_U = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet_L = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet_A = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            'Convert(VarChar(8),T0.U_PFromDate,112) As 'FD',Convert(VarChar(8),U_PToDate,112) As 'TD',U_NoOfDays,ISNULL(U_FreeDays,0) As 'U_FreeDays'

            strQuery = " Select T0.DocEntry "
            strQuery += " From [@Z_OCPM] T0  "
            strQuery += " Where T0.U_CardCode = (Select U_CardCode From [@Z_OCPR] Where DocEntry = '" + strDocEntry + "')"
            'strQuery += " And T0.U_RemDays > 0 "
            strQuery += " And ( "
            strQuery += " T0.U_RemDays > 0 OR (Convert(VarChar(8),U_PToDate,112) >='" & System.DateTime.Now.AddDays(-1).ToString("yyyyMMdd") & "') "
            strQuery += " OR ( T0.U_ReRun = 'Y' ) "
            strQuery += " ) "
            strQuery += " And T0.U_DocStatus = 'O' "
            If strPRef <> "" Then
                strQuery += " And T0.DocEntry = '" & strPRef & "'"
            End If
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    Dim strPGRef As String = oRecordSet.Fields.Item("DocEntry").Value

                    strQuery = " Select T0.DocEntry,T0.U_PrgCode,T1.LineId,T0.U_CardCode,Convert(VarChar(8),T1.U_Fdate,112) As 'FD',Convert(VarChar(8),T1.U_Edate,112) As 'TD', "
                    strQuery += " ISNULL(T1.U_NoOfDays,0) As U_NoOfDays,ISNULL(T1.U_OrdDays,0) As U_OrdDays,ISNULL(T1.U_DelDays,0) As U_DelDays,ISNULL(T1.U_InvDays,0) As U_InvDays "
                    strQuery += " From [@Z_OCPM] T0 JOIN [@Z_CPM6] T1 On T0.DocEntry = T1.DocEntry "
                    strQuery += " Where T0.U_CardCode = (Select U_CardCode From [@Z_OCPR] Where DocEntry = '" + strDocEntry + "')"
                    strQuery += " And T0.DocEntry = '" & strPGRef & "'"
                    strQuery += " And ( "
                    strQuery += " T0.U_RemDays > 0 "
                    strQuery += " OR (Convert(VarChar(8),U_PToDate,112) >='" & System.DateTime.Now.AddDays(-1).ToString("yyyyMMdd") & "') "
                    strQuery += " OR (T0.U_ReRun = 'Y') "
                    strQuery += " ) "
                    strQuery += " Order By T1.LineId "
                    oRecordSet_L.DoQuery(strQuery)

                    If Not oRecordSet_L.EoF Then
                        While Not oRecordSet_L.EoF

                            Dim strPFromDt As String = oRecordSet_L.Fields.Item("FD").Value
                            Dim strPToDt As String = oRecordSet_L.Fields.Item("TD").Value
                            Dim strPGLine As String = oRecordSet_L.Fields.Item("LineId").Value
                            Dim strCardCode As String = oRecordSet_L.Fields.Item("U_CardCode").Value
                            Dim strQuantity As String = (CInt(oRecordSet_L.Fields.Item("U_NoOfDays").Value)).ToString ' - CInt(oRecordSet_L.Fields.Item("U_DelDays").Value))

                            Dim strPFromDt1 As String = String.Empty
                            Dim strPToDt1 As String = String.Empty

                            '===

                            If CInt(strPGLine) > 1 Then
                                strQuery = " Select Convert(Varchar(8),(U_Edate + 1),112) As 'ED' From [@Z_CPM6] "
                                strQuery += " Where DocEntry = '" + strPGRef + "'"
                                strQuery += " And LineId = '" + (CInt(strPGLine) - 1).ToString() + "'"
                                Dim oRec_TD As SAPbobsCOM.Recordset
                                oRec_TD = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                oRec_TD.DoQuery(strQuery)
                                If Not oRec_TD.EoF Then
                                    strPFromDt1 = oRec_TD.Fields.Item("ED").Value
                                End If
                            End If

                            strQuery = " Select ISNULL(U_OrdDays,0) As U_OrdDays From [@Z_CPM6] "
                            strQuery += " Where DocEntry = '" + strPGRef + "'"
                            strQuery += " And LineId = '" + (CInt(strPGLine) + 1).ToString() + "'"
                            Dim strNOrderQty As String = oApplication.Utilities.getRecordSetValueString(strQuery, "U_OrdDays")

                            If CInt(IIf(strNOrderQty, "0", strNOrderQty)) > 0 Then

                                strQuery = " Select Distinct Count(U_DelDate) As 'NO',T1.LineID From DLN1 T0  "
                                strQuery += " JOIN [@Z_CPM6] T1 On T0.U_ProgramID = T1.DocEntry "
                                strQuery += " And T0.U_DelDate BetWeen T1.U_Fdate And T1.U_Edate "
                                strQuery += " And ((T0.LineStatus = 'C' And T0.TargetType = '-1')) "
                                strQuery += " JOIN ODLN T2 On T0.DocEntry = T2.DocEntry And T2.DocStatus = 'C' "
                                strQuery += " Where T0.U_ProgramID = '" & strPGRef & "' "
                                strQuery += " T1.LineId >= '" & strPGLine & "'"
                                strQuery += " T1.U_IsCal = 'N' "
                                strQuery += " And T1.U_CanFrom In ('E','S') "
                                Dim strAQty As String '= oApplication.Utilities.getRecordSetValueString(strQuery, "NO")
                                oRecordSet_A.DoQuery(strQuery)
                                If Not oRecordSet_A.EoF Then

                                    While Not oRecordSet_A.EoF
                                        strAQty = oRecordSet_L.Fields.Item("NO").Value

                                        strQuery = " Select Convert(Varchar(8),(U_Edate + 1),112) As 'ED' From [@Z_CPM6] "
                                        strQuery += " Where DocEntry = '" + strPGRef + "'"
                                        strQuery += " And LineId = (Select Max(LineId) From  [@Z_CPM6]  Where DocEntry = '" + strPGRef + "')"
                                        Dim oRec_TD As SAPbobsCOM.Recordset
                                        oRec_TD = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                        oRec_TD.DoQuery(strQuery)
                                        If Not oRec_TD.EoF Then
                                            strPFromDt1 = oRec_TD.Fields.Item("ED").Value
                                        End If

                                        If CInt(IIf(strAQty, "0", strAQty)) > 0 Then
                                            updateProgramDocumentRow(oForm, strCardCode, strPGRef, oRecordSet_L.Fields.Item("LineID").Value, strPFromDt1, strAQty)

                                            'Reset the No of Qty of the Row = No of Qty - strAQty
                                            strQuery = " Update T1 Set U_NoofDays = U_NoofDays - '" & strAQty & "' "
                                            strQuery += " From [@Z_CPM6] T1 "
                                            strQuery += " Where T0.U_ProgramID = '" & strPGRef & "' "
                                            strQuery += " T1.LineId = '" & oRecordSet_L.Fields.Item("LineID").Value & "'"
                                            '
                                        End If

                                        oRecordSet_L.MoveNext()

                                    End While

                                    'Make Status of the above Documents IsCal = 'Y'
                                    strQuery = " Update T0 Set U_IsCal = 'Y' "
                                    strQuery += " JOIN [@Z_CPM6] T1 On T0.U_ProgramID = T1.DocEntry "
                                    strQuery += " And T0.U_DelDate BetWeen T1.U_Fdate And T1.U_Edate "
                                    strQuery += " And ((T0.LineStatus = 'C' And T0.TargetType = '-1')) "
                                    strQuery += " JOIN ODLN T2 On T0.DocEntry = T2.DocEntry And T2.DocStatus = 'C' "
                                    strQuery += " Where T0.U_ProgramID = '" & strPGRef & "' "
                                    strQuery += " T1.LineId >= '" & strPGLine & "'"
                                    strQuery += " T1.U_IsCal = 'N' "
                                    strQuery += " And T1.U_CanFrom In ('E','S') "
                                    oRecordSet_U.DoQuery(strQuery)
                                    'Make Status of the above Documents IsCal = 'Y'
                                    Exit Function

                                End If
                            Else

                                If CInt(oRecordSet_L.Fields.Item("U_DelDays").Value) > 0 Then
                                    strQuantity = (CInt(oRecordSet_L.Fields.Item("U_NoOfDays").Value) - CInt(oRecordSet_L.Fields.Item("U_DelDays").Value)).ToString
                                    strQuery = " Select Convert(Varchar(8),Max(T0.U_DelDate+1),112) As 'FD' From DLN1 T0 "
                                    strQuery += " JOIN [@Z_CPM6] T1 On T0.U_ProgramID = T1.DocEntry "
                                    strQuery += " And T0.U_DelDate BetWeen T1.U_Fdate And T1.U_Edate "
                                    strQuery += " And (T0.LineStatus = 'O' "
                                    strQuery += " Or (T0.LineStatus = 'C' And T0.TargetType = '-1')"
                                    strQuery += " ) "
                                    strQuery += " JOIN ODLN T2 On T0.DocEntry = T2.DocEntry And T2.CANCELED = 'N' "
                                    strQuery += " Where T0.U_ProgramID = '" & strPGRef & "' "
                                    strQuery += " And T1.LineId = '" & strPGLine & "'"
                                    strPFromDt = oApplication.Utilities.getRecordSetValueString(strQuery, "FD")
                                    If strPFromDt <> "" Then
                                        strPToDt1 = oApplication.Utilities.getProgramToDate(oForm, strCardCode, strPFromDt, strQuantity)
                                    Else
                                        strPToDt1 = strPToDt
                                    End If
                                Else
                                    If CInt(strPGLine) > 1 Then
                                        If strPFromDt1 <> "" Then
                                            strPToDt1 = oApplication.Utilities.getProgramToDate(oForm, strCardCode, strPFromDt1, strQuantity)
                                        Else
                                            strPToDt1 = oApplication.Utilities.getProgramToDate(oForm, strCardCode, strPFromDt, strQuantity)
                                        End If
                                    Else
                                        strPToDt1 = oApplication.Utilities.getProgramToDate(oForm, strCardCode, strPFromDt, strQuantity)
                                    End If
                                End If

                                If strPFromDt1 <> "" Then
                                    If strPFromDt1 <> strPFromDt Then
                                        strQuery = "Update [@Z_CPM6] SET "
                                        strQuery += " U_Fdate = '" + strPFromDt1 + "'"
                                        strQuery += " Where DocEntry = '" + strPGRef + "'"
                                        strQuery += " And LineId = '" + strPGLine + "'"
                                        oRecordSet_U.DoQuery(strQuery)
                                    End If
                                End If


                                If strPToDt1 <> strPToDt Then
                                    If strPToDt1 <> "" Then
                                        strQuery = "Update [@Z_CPM6] SET "
                                        strQuery += " U_Edate = '" + strPToDt1 + "'"
                                        strQuery += " Where DocEntry = '" + strPGRef + "'"
                                        strQuery += " And LineId = '" + strPGLine + "'"
                                        oRecordSet_U.DoQuery(strQuery)
                                    End If
                                End If

                            End If

                            oRecordSet_L.MoveNext()
                            '===

                            'If CInt(strPGLine) > 1 Then
                            '    strQuery = " Select Convert(Varchar(8),(U_Edate + 1),112) As 'ED' From [@Z_CPM6] "
                            '    strQuery += " Where DocEntry = '" + strPGRef + "'"
                            '    strQuery += " And LineId = '" + (CInt(strPGLine) - 1).ToString() + "'"
                            '    Dim oRec_TD As SAPbobsCOM.Recordset
                            '    oRec_TD = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                            '    oRec_TD.DoQuery(strQuery)
                            '    If Not oRec_TD.EoF Then
                            '        strPFromDt = oRec_TD.Fields.Item("ED").Value
                            '    End If
                            'End If


                            'If CInt(oRecordSet_L.Fields.Item("U_InvDays").Value) < CInt(oRecordSet_L.Fields.Item("U_NoOfDays").Value) Then
                            '    If CInt(oRecordSet_L.Fields.Item("U_DelDays").Value) > 0 Then
                            '        strQuantity = (CInt(oRecordSet_L.Fields.Item("U_NoOfDays").Value) - CInt(oRecordSet_L.Fields.Item("U_DelDays").Value)).ToString
                            '        strQuery = " Select Convert(Varchar(8),Max(T0.U_DelDate+1),112) As 'FD' From DLN1 T0 "
                            '        strQuery += " JOIN [@Z_CPM6] T1 On T0.U_ProgramID = T1.DocEntry "
                            '        strQuery += " And T0.U_DelDate BetWeen T1.U_Fdate And T1.U_Edate "
                            '        strQuery += " And (T0.LineStatus = 'O' Or (T0.LineStatus = 'C' And T0.TargetType = '-1')) "
                            '        strQuery += " JOIN ODLN T2 On T0.DocEntry = T2.DocEntry And T2.CANCELED = 'N' "
                            '        strQuery += " Where T0.U_ProgramID = '" & strPGRef & "' "
                            '        strPFromDt = oApplication.Utilities.getRecordSetValueString(strQuery, "FD")
                            '        If strPFromDt <> "" Then
                            '            strPToDt1 = oApplication.Utilities.getProgramToDate(oForm, strCardCode, strPFromDt, strQuantity)
                            '        Else
                            '            strPToDt1 = strPToDt
                            '        End If
                            '    Else
                            '        strPToDt1 = oApplication.Utilities.getProgramToDate(oForm, strCardCode, strPFromDt, strQuantity)
                            '    End If
                            'Else
                            '    strPToDt1 = strPToDt
                            'End If


                            'strQuery = "Update [@Z_CPM6] SET "
                            'If CInt(strPGLine) > 1 Then
                            '    strQuery += " U_Fdate = '" + strPFromDt + "',"
                            'End If
                            'strQuery += " U_Edate = '" + strPToDt1 + "'"
                            'strQuery += " Where DocEntry = '" + strPGRef + "'"
                            'strQuery += " And LineId = '" + strPGLine + "'"
                            'oRecordSet_U.DoQuery(strQuery)


                        End While
                    End If

                    strQuery = " Select Convert(VarChar(8),Min(U_Fdate),112) As 'FD',Convert(VarChar(8),Max(U_Edate),112) As 'LD' From [@Z_CPM6] "
                    strQuery += " Where DocEntry = '" + strPGRef + "' And ISNULL(U_NoofDays,0) > 0 "
                    oRecordSet_L.DoQuery(strQuery)
                    If Not oRecordSet_L.EoF Then

                        Dim strPFromDt As String = oRecordSet_L.Fields.Item("FD").Value
                        Dim strPToDt As String = oRecordSet_L.Fields.Item("LD").Value

                        strQuery = "Update [@Z_OCPM] SET "
                        strQuery += " U_PFromDate = '" + strPFromDt + "',"
                        strQuery += " U_PToDate = '" + strPToDt + "'"
                        strQuery += " Where DocEntry = '" + oRecordSet.Fields.Item("DocEntry").Value.ToString() + "'"
                        oRecordSet_U.DoQuery(strQuery)

                    Else

                        strQuery = " Select Convert(VarChar(8),Min(U_Fdate),112) As 'FD',Convert(VarChar(8),Max(U_Edate),112) As 'LD' From [@Z_CPM6] "
                        strQuery += " Where DocEntry = '" + strPGRef + "' And ISNULL(U_NoofDays,0) = 0 "
                        oRecordSet_L.DoQuery(strQuery)
                        If Not oRecordSet_L.EoF Then
                            Dim strPFromDt As String = oRecordSet_L.Fields.Item("FD").Value
                            Dim strPToDt As String = oRecordSet_L.Fields.Item("LD").Value
                            strQuery = "Update [@Z_OCPM] SET "
                            strQuery += " U_PFromDate = '" + strPFromDt + "',"
                            strQuery += " U_PToDate = '" + strPToDt + "'"
                            strQuery += " Where DocEntry = '" + oRecordSet.Fields.Item("DocEntry").Value.ToString() + "'"
                            oRecordSet_U.DoQuery(strQuery)
                        End If

                    End If

                    oApplication.Utilities.updateCustomerProgramInProgram(strPGRef)

                    oRecordSet.MoveNext()
                End While




                'oRecordSet.MoveFirst()
                'While Not oRecordSet.EoF
                '    Dim strPGRef As String = oRecordSet.Fields.Item("DocEntry").Value
                '    Dim strCardCode As String = oRecordSet.Fields.Item("U_CardCode").Value
                '    Dim strQuantity As String = oRecordSet.Fields.Item("U_NoOfDays").Value + oRecordSet.Fields.Item("U_FreeDays").Value
                '    Dim strPFromDt As String = oRecordSet.Fields.Item("FD").Value
                '    Dim strPToDt As String = oApplication.Utilities.getProgramToDate(oForm, strCardCode, strPFromDt, strQuantity)

                '    strQuery = "Update [@Z_OCPM] SET "
                '    strQuery += " U_PToDate = '" + strPToDt + "'"
                '    strQuery += " Where DocEntry = '" + oRecordSet.Fields.Item("DocEntry").Value.ToString() + "'"
                '    oRecordSet_U.DoQuery(strQuery)

                '    oRecordSet.MoveNext()
                'End While

                ''Update Program Row Table for Program Date
                'strQuery = " Select DocEntry,U_CardCode,Convert(VarChar(8),U_PFromDate,112) As 'FD',Convert(VarChar(8),U_PToDate,112) As 'TD',U_NoOfDays,ISNULL(U_FreeDays,0) As 'U_FreeDays' "
                'strQuery += " From [@Z_OCPM] Where U_CardCode = (Select U_CardCode From [@Z_OCPR] Where DocEntry = '" + strDocEntry + "')"
                'strQuery += " And U_RemDays > 0 "
                'oRecordSet.DoQuery(strQuery)
                'While Not oRecordSet.EoF
                '    Dim strPGRef As String = oRecordSet.Fields.Item("DocEntry").Value
                '    oApplication.Utilities.updateCustomerProgramInProgram(strPGRef)
                '    oRecordSet.MoveNext()
                'End While

            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Private Function updateProgramDocumentRow(ByVal oForm As SAPbouiCOM.Form, ByVal strCardCode As String, ByVal refPG As String, _
                                              ByVal refLine As String, ByVal strPFromDt1 As String, ByVal strAQty As String) As Boolean
        Try

            Dim _retVal As Boolean = False
            Dim strQuery As String = String.Empty

            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
            Dim oCompanyService As SAPbobsCOM.CompanyService
            Dim oGeneralDataCollection As SAPbobsCOM.GeneralDataCollection
            Dim oChildData As SAPbobsCOM.GeneralData
            oCompanyService = oApplication.Company.GetCompanyService()

            Dim oDoc As SAPbobsCOM.SalesOpportunities = Nothing
            Dim oRecordSet As SAPbobsCOM.Recordset
            Try
                oGeneralService = oCompanyService.GetGeneralService("Z_OCPM")
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                Dim dtFromDate As Date = CDate(strPFromDt1.Substring(0, 4) + "-" + strPFromDt1.Substring(4, 2) + "-" + strPFromDt1.Substring(6, 2)).AddDays(1)
                Dim strPToDt1 As String
                strQuery = " Select * From [@Z_CPM6] Where DocEntry = '" & refPG & "'"
                strQuery += " And LineId = '" & refLine & "'"
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery(strQuery)

                If Not oRecordSet.EoF Then

                    oGeneralParams.SetProperty("DocEntry", oRecordSet.Fields.Item(0).Value)
                    oGeneralData = oGeneralService.GetByParams(oGeneralParams)

                    'Adding Break Up Row
                    oGeneralDataCollection = oGeneralData.Child("Z_CPM6")
                    oChildData = oGeneralDataCollection.Add()
                    oChildData.SetProperty("U_Fdate", dtFromDate.AddDays(0))
                    strPToDt1 = oApplication.Utilities.getProgramToDate(oForm, strCardCode, dtFromDate.ToString("yyyyMMdd"), strAQty)
                    oChildData.SetProperty("U_Edate", strPToDt1)
                    oChildData.SetProperty("U_NoofDays", strAQty.Trim())
                    oChildData.SetProperty("U_Price", oRecordSet.Fields.Item("U_Price").Value)
                    oChildData.SetProperty("U_Discount", oRecordSet.Fields.Item("U_Discount").Value)
                    Dim dblLineTotal As Double = (CDbl(strAQty.Trim()) * CDbl(oRecordSet.Fields.Item("U_Price").Value)) - (CDbl(oRecordSet.Fields.Item("U_Price").Value) * (CDbl(oRecordSet.Fields.Item("U_Discount").Value) / 100))
                    oChildData.SetProperty("U_LineTotal", dblLineTotal)
                    oGeneralService.Update(oGeneralData)

                    '    'Adding Program Row
                    '    oGeneralDataCollection = oGeneralData.Child("Z_CPM1")
                    '    For intRow As Integer = 0 To CInt(strAQty)
                    '        strQuery = "Select DocEntry,LineId,U_AppStatus From [@Z_CPM1] Where DocEntry = '" + refPG + "' And Convert(VarChar(8),U_PrgDate,112) = '" + dtFromDate.AddDays(intRow).ToString("yyyyMMdd") + "'"
                    '        oRecordSet.DoQuery(strQuery)
                    '        If oRecordSet.EoF Then
                    '            oChildData = oGeneralDataCollection.Add()
                    '            oChildData.SetProperty("U_PrgDate", dtFromDate.AddDays(intRow))
                    '            Dim strIncStatus As String = checkExclude(strCardCode, dtFromDate.AddDays(intRow))
                    '            oChildData.SetProperty("U_AppStatus", strIncStatus.Trim())
                    '        End If
                    '    Next
                    'oGeneralService.Update(oGeneralData)
                End If

                Return _retVal
            Catch ex As Exception
                oApplication.Log.Trace_DIET_AddOn_Error(ex)
                Throw ex
                'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
            End Try
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Private Function addExcludeDocumentRow(ByVal strCardCode As String, ByVal strCPRef As String, ByVal strExDate As String) As Boolean
        Try

            Dim _retVal As Boolean = False
            Dim strQuery As String = String.Empty

            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
            Dim oCompanyService As SAPbobsCOM.CompanyService
            Dim oGeneralDataCollection As SAPbobsCOM.GeneralDataCollection
            Dim oChildData As SAPbobsCOM.GeneralData
            oCompanyService = oApplication.Company.GetCompanyService()

            Dim oDoc As SAPbobsCOM.SalesOpportunities = Nothing
            Dim oRecordSet As SAPbobsCOM.Recordset
            Try
                oGeneralService = oCompanyService.GetGeneralService("Z_OCPR")
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                Dim dtFromDate As Date = CDate(strExDate.Substring(0, 4) + "-" + strExDate.Substring(4, 2) + "-" + strExDate.Substring(6, 2))
                Dim dtCurrDate As Date = CDate(System.DateTime.Now.Year.ToString() + "-" + System.DateTime.Now.Month.ToString() + "-" + System.DateTime.Now.Day.ToString())

                If dtFromDate > dtCurrDate Then
                    strQuery = " Select U_ExDate From [@Z_CPR4] Where DocEntry = '" & strCPRef & "'"
                    strQuery += " And Convert(VarChar(8),U_ExDate,112) = '" & strExDate & "'"
                    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery(strQuery)
                    If oRecordSet.EoF Then
                        oGeneralParams.SetProperty("DocEntry", strCPRef)
                        oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                        oGeneralDataCollection = oGeneralData.Child("Z_CPR4")
                        oChildData = oGeneralDataCollection.Add()
                        oChildData.SetProperty("U_ExDate", dtFromDate.AddDays(0))
                        oGeneralService.Update(oGeneralData)
                    End If
                End If
                
                Return _retVal
            Catch ex As Exception
                oApplication.Log.Trace_DIET_AddOn_Error(ex)
                Throw ex
                'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
            End Try
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function updateProgramNoofDayAndToDate(ByVal oForm As SAPbouiCOM.Form, ByVal strDocEntry As String, ByVal intTNoofDays As Integer) As Boolean
        Dim _retVal As Boolean = False
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim oRecordSet_U As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet_U = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            strQuery = " Select DocEntry,U_CardCode,Convert(VarChar(8),U_PFromDate,112) As 'FD',Convert(VarChar(8),U_PToDate,112) As 'TD',U_RemDays "
            strQuery += " From [@Z_OCPM] Where DocEntry = '" + strDocEntry + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                Dim strPGRef As String = oRecordSet.Fields.Item("DocEntry").Value
                Dim strCardCode As String = oRecordSet.Fields.Item("U_CardCode").Value
                Dim strQuantity As String = CInt(oRecordSet.Fields.Item("U_RemDays").Value) + intTNoofDays
                Dim strPFromDt As String = oRecordSet.Fields.Item("FD").Value
                Dim strPToDt As String = oApplication.Utilities.getProgramToDate(oForm, strCardCode, strPFromDt, strQuantity)
                strQuery = "Update [@Z_OCPM] SET "
                strQuery += " U_PToDate = '" + strPToDt + "'"
                strQuery += ", U_NoOfDays = U_NoOfDays + " + intTNoofDays.ToString() + ""
                strQuery += ", U_RemDays = U_RemDays + " + intTNoofDays.ToString() + ""
                strQuery += " Where DocEntry = '" + oRecordSet.Fields.Item("DocEntry").Value.ToString() + "'"
                oRecordSet_U.DoQuery(strQuery)
            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function UpdateProgramNoofDaysBasedOnRemoveDate(ByVal oForm As SAPbouiCOM.Form, ByVal strDocEntry As String) As Boolean
        Dim _retVal As Boolean = False
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim oRecordSet_L As SAPbobsCOM.Recordset
            Dim oRecordSet_U As SAPbobsCOM.Recordset
            Dim oRecordSet_A As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet_U = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet_L = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet_A = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            strQuery = " Select T0.DocEntry "
            strQuery += " From [@Z_OCPM] T0  "
            strQuery += " Where T0.U_CardCode = (Select U_CardCode From [@Z_OCPR] Where DocEntry = '" + strDocEntry + "')"
            'strQuery += " And (ISNULL(T0.U_RemDays,0) > 0) "
            strQuery += " And (T0.U_RemDays > 0 OR (Convert(VarChar(8),U_PToDate,112) >='" & System.DateTime.Now.AddDays(-1).ToString("yyyyMMdd") & "')) "
            strQuery += " And T0.U_DocStatus = 'O' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF

                    Dim strPGRef As String = oRecordSet.Fields.Item("DocEntry").Value

                    strQuery = " Select T1.LineId "
                    strQuery += " From [@Z_OCPM] T0 JOIN [@Z_CPM6] T1 On T0.DocEntry = T1.DocEntry "
                    strQuery += " Where T0.DocEntry = '" & strPGRef & "'"
                    strQuery += " And  T0.U_RemDays > 0 "
                    strQuery += " Order By T1.LineId "
                    oRecordSet_L.DoQuery(strQuery)

                    If Not oRecordSet_L.EoF Then
                        While Not oRecordSet_L.EoF

                            Dim oRec_TD As SAPbobsCOM.Recordset
                            oRec_TD = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

                            strQuery = " Select T0.U_PrgDate From [@Z_CPM1] T0  "
                            strQuery += " JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry "
                            strQuery += " JOIN [@Z_CPM6] T4 On T4.DocEntry = T1.DocEntry "
                            strQuery += " AND Convert(VarChar(8),T0.U_PrgDate,112) "
                            strQuery += " Between Convert(VarChar(8),T4.U_Fdate,112) "
                            strQuery += " And Convert(VarChar(8),T4.U_Edate,112) "
                            strQuery += " JOIN [@Z_OCPR] T2 On T1.U_CardCode = T2.U_CardCode "
                            strQuery += " JOIN [@Z_CPR8] T3 On T2.DocEntry = T3.DocEntry "
                            strQuery += " AND Convert(VarChar(8),T0.U_PrgDate,112) "
                            strQuery += " Between Convert(VarChar(8),T3.U_FDate,112) "
                            strQuery += " And Convert(VarChar(8),T3.U_TDate,112) "
                            strQuery += " Where (T0.U_AppStatus = 'I' And T0.U_ONOFFSTA = 'O') "
                            strQuery += " AND Convert(VarChar(8),T0.U_PrgDate,112) "
                            strQuery += " Between Convert(VarChar(8),T1.U_PFromDate,112) "
                            strQuery += " And Convert(VarChar(8),T1.U_PToDate,112) "
                            strQuery += " And T1.DocEntry = '" & strPGRef & "'"
                            strQuery += " And T4.LineId = '" & oRecordSet_L.Fields.Item(0).Value.ToString & "' "
                            oRec_TD.DoQuery(strQuery)
                            If Not oRec_TD.EoF Then

                                strQuery = "Update [@Z_CPM6] SET "
                                strQuery += " U_NoofDays = U_NoofDays  - " & oRec_TD.RecordCount.ToString & ","
                                strQuery += " U_RmvDays = - ISNULL(U_RmvDays,0)  - " & oRec_TD.RecordCount.ToString & ""
                                strQuery += " Where DocEntry = '" + strPGRef + "'"
                                strQuery += " And LineId = '" & oRecordSet_L.Fields.Item(0).Value.ToString & "' "
                                oRecordSet_U.DoQuery(strQuery)

                                'Rerun the Program Date 
                                strQuery = " Update [@Z_OCPM] SET "
                                strQuery += " U_ReRun = 'Y' "
                                strQuery += " Where DocEntry = '" + strPGRef + "'"
                                oRecordSet_U.DoQuery(strQuery)

                            End If
                            oRecordSet_L.MoveNext()

                        End While
                    End If

                    'Updating Header Details
                    strQuery = " Select  "
                    strQuery += " (Select SUM(ISNULL(U_NoOfDays,0)) From [@Z_CPM6] T1 Where U_PaidType = 'P' And T1.DocEntry = T0.DocEntry) As 'ND', "
                    strQuery += " (Select SUM(ISNULL(U_NoOfDays,0)) From [@Z_CPM6] T1 Where U_PaidType = 'F' And T1.DocEntry = T0.DocEntry) As 'FD' , "
                    strQuery += " (Select SUM(ISNULL(U_RmvDays,0)) From [@Z_CPM6] T1 Where T1.DocEntry = T0.DocEntry) As 'VD' , "
                    strQuery += " T0.U_DelDays As 'DD' "
                    strQuery += " From [@Z_OCPM] T0 "
                    strQuery += " Where T0.DocEntry = '" + strPGRef + "'"
                    oRecordSet_L.DoQuery(strQuery)
                    If Not oRecordSet_L.EoF Then

                        Dim strND As String = oRecordSet_L.Fields.Item("ND").Value
                        Dim strFD As String = oRecordSet_L.Fields.Item("FD").Value
                        Dim strDD As String = oRecordSet_L.Fields.Item("DD").Value
                        Dim strVD As String = oRecordSet_L.Fields.Item("VD").Value
                        Dim strRD As String = (CInt(strND) + CInt(strFD)) - CInt(strDD)


                        strQuery = "Update [@Z_OCPM] SET "
                        strQuery += " U_NoOfDays = '" + strND + "',"
                        strQuery += " U_FreeDays = '" + strFD + "',"
                        strQuery += " U_RemDays =  '" + strRD + "',"
                        strQuery += " U_RmvDays =  '" + strVD + "'"

                        strQuery += " Where DocEntry = '" + oRecordSet.Fields.Item("DocEntry").Value.ToString() + "'"
                        oRecordSet_U.DoQuery(strQuery)

                    End If

                    oRecordSet.MoveNext()

                End While
            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function updateProgramDates_IfOverLapping(ByVal oForm As SAPbouiCOM.Form, ByVal strDocEntry As String, ByVal strCardCode As String) As Boolean
        Dim _retVal As Boolean = False
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim oRecordSet_C As SAPbobsCOM.Recordset
            Dim oRecordSet_U As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet_U = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet_C = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            strQuery = " Select T0.DocEntry,U_CardCode,U_CardName,Convert(VarChar(8),U_PFromDate,112) As U_PFromDate "
            strQuery += " ,Convert(Varchar(8),U_PToDate,112) As U_PToDate From [@Z_OCPM] T0  "
            strQuery += " Where T0.U_DocStatus = 'O'  "
            strQuery += " And T0.U_CardCode = '" + strCardCode + "'"
            'strQuery += " And T0.U_RemDays > 0 "
            strQuery += " And ( T0.U_RemDays > 0 "
            strQuery += " OR (Convert(VarChar(8),U_PToDate,112) >='" & System.DateTime.Now.AddDays(-1).ToString("yyyyMMdd") & "') "
            strQuery += " ) "
            strQuery += " Order By U_PFromDate "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                If oRecordSet.RecordCount > 1 Then
                    While Not oRecordSet.EoF

                        Dim strPGRef As String = oRecordSet.Fields.Item("DocEntry").Value
                        Dim strPFromDt As String = oRecordSet.Fields.Item("U_PFromDate").Value


                        strQuery = " Select DocEntry,Convert(Varchar(8),U_PToDate,112) As U_PToDate from [@Z_OCPM] where U_CardCode= '" & strCardCode & "'" & _
                        " And '" & strPFromDt & "' between Convert(VarChar(8),U_PFromDate,112) and Convert(VarChar(8),U_PToDate,112) And IsNull(U_Cancel,'N') = 'N' " & _
                        " And ISNULL(U_Transfer,'N') = 'N' " & _
                        " And ISNULL(U_DocStatus,'O') = 'O' " & _
                        " And DocEntry <> '" & strPGRef & "'" & _
                        " And U_RemDays > 0 " & _
                        " Order By U_PFromDate "
                        oRecordSet_C.DoQuery(strQuery)
                        'Overlapping...
                        If oRecordSet_C.RecordCount > 0 Then

                            Dim strPToDt As String = oRecordSet_C.Fields.Item("U_PToDate").Value
                            Dim dtToDate As Date = strPToDt.Substring(0, 4) + "-" + strPToDt.Substring(4, 2) + "-" + strPToDt.Substring(6, 2)

                            strQuery = "Update [@Z_OCPM] SET "
                            strQuery += " U_PFromDate = '" + dtToDate.AddDays(1).ToString("MM-dd-yyyy") + "',"
                            strQuery += " U_IsSequence = 'Y' "
                            strQuery += " Where DocEntry = '" + strPGRef + "'"
                            oRecordSet_U.DoQuery(strQuery)

                            strQuery = "Update [@Z_CPM6] SET "
                            strQuery += " U_Fdate = '" + dtToDate.AddDays(1).ToString("MM-dd-yyyy") + "'"
                            strQuery += " Where DocEntry = '" + strPGRef + "'"
                            strQuery += " And LineId = '1' "
                            oRecordSet_U.DoQuery(strQuery)

                        Else

                            strQuery = " Select Top 1 DocEntry,Convert(Varchar(8),U_PToDate,112) As U_PToDate from [@Z_OCPM] where U_CardCode= '" & strCardCode & "'" & _
                                        " And IsNull(U_Cancel,'N') = 'N' " & _
                                        " And ISNULL(U_Transfer,'N') = 'N' " & _
                                        " And ISNULL(U_DocStatus,'O') = 'O' " & _
                                        " And DocEntry <> '" & strPGRef & "'" & _
                                        " And U_RemDays > 0 " & _
                                        " And Convert(VarChar(8),U_PToDate,112) <  '" & strPFromDt & "'  " & _
                                        " Order By U_PFromDate "
                            oRecordSet_C.DoQuery(strQuery)
                            If oRecordSet_C.RecordCount > 0 Then

                                Dim strPToDt As String = oRecordSet_C.Fields.Item("U_PToDate").Value
                                Dim dtFromDate As Date = strPFromDt.Substring(0, 4) + "-" + strPFromDt.Substring(4, 2) + "-" + strPFromDt.Substring(6, 2)
                                Dim dtToDate As Date = strPToDt.Substring(0, 4) + "-" + strPToDt.Substring(4, 2) + "-" + strPToDt.Substring(6, 2)

                                If dtFromDate.ToString("MM-dd-yyyy") <> dtToDate.AddDays(1).ToString("MM-dd-yyyy") Then

                                    strQuery = "Update [@Z_OCPM] SET "
                                    strQuery += " U_PFromDate = '" + dtToDate.AddDays(1).ToString("MM-dd-yyyy") + "'"
                                    strQuery += " Where DocEntry = '" + strPGRef + "'"
                                    strQuery += " And U_IsSequence = 'Y' "
                                    oRecordSet_U.DoQuery(strQuery)

                                    strQuery = "Update T1 SET "
                                    strQuery += " T1.U_Fdate = '" + dtToDate.AddDays(1).ToString("MM-dd-yyyy") + "'"
                                    strQuery += " From [@Z_OCPM] T0 JOIN [@Z_CPM6] T1 On T0.DocEntry = T1.DocEntry "
                                    strQuery += " Where T1.DocEntry = '" + strPGRef + "'"
                                    strQuery += " And T1.LineId = '1' "
                                    strQuery += " And T0.U_IsSequence = 'Y' "

                                    oRecordSet_U.DoQuery(strQuery)

                                End If
                            End If

                        End If

                        UpdateProgramToDate(oForm, strDocEntry, strPGRef)

                        oRecordSet.MoveNext()
                    End While
                Else
                    Exit Function
                End If


            End If

            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Sub UpdateCustomerFoodMenu(ByVal oForm As SAPbouiCOM.Form, ByVal strGridID As String, ByVal strFType As String, ByVal strSType As String)
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim sCode As String
        Dim oGrid As SAPbouiCOM.Grid
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim strQuery As String = String.Empty
        Try
            Dim strCardCode As String = CType(oForm.Items.Item("31").Specific, SAPbouiCOM.EditText).Value
            Dim strProgramID As String = CType(oForm.Items.Item("36").Specific, SAPbouiCOM.EditText).Value
            Dim strMenuDate As String = CType(oForm.Items.Item("35").Specific, SAPbouiCOM.EditText).Value
            Dim dtPrgDate As Date = strMenuDate.Substring(0, 4) + "-" + strMenuDate.Substring(4, 2) + "-" + strMenuDate.Substring(6, 2)
            Dim strSession As String = CType(oForm.Items.Item("38").Specific, SAPbouiCOM.EditText).Value

            oGrid = oForm.Items.Item(strGridID).Specific

            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1

                oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                oUserTable = oApplication.Company.UserTables.Item("Z_OFSL")
                Dim strSelect As String = oGrid.DataTable.GetValue("Select", intRow).ToString
                Dim strItemCode As String = oGrid.DataTable.GetValue("U_ItemCode", intRow)
                If strItemCode.Trim().Length = 0 Then
                    Continue For
                End If

                Dim dblQty As Double = CDbl(oGrid.DataTable.GetValue("Qty", intRow))
                Dim strRemarks As String = oGrid.DataTable.GetValue("Remarks", intRow)

                Dim strDislike As String = oGrid.DataTable.GetValue("U_Dislike", intRow)
                Dim strMedical As String = oGrid.DataTable.GetValue("U_Medical", intRow)

                sCode = Me.getMaxCode("@Z_OFSL", "Code")
                strQuery = "Select Code From [@Z_OFSL] Where U_ProgramID = '" + strProgramID + "'"
                strQuery += " And Convert(VarChar(8),U_PrgDate,112) = '" + strMenuDate + "'"
                strQuery += " And U_FType = '" + strFType + "'"
                strQuery += " And U_SFood = '" + strSType + "'"
                strQuery += " AND U_ItemCode = '" + strItemCode + "'"
                oRecordSet.DoQuery(strQuery)
                If oRecordSet.EoF Then
                    If Not oUserTable.GetByKey(sCode) Then
                        If strSelect = "Y" Then
                            oUserTable.Code = sCode
                            oUserTable.Name = sCode
                            With oUserTable.UserFields.Fields
                                .Item("U_ProgramID").Value = strProgramID
                                .Item("U_CardCode").Value = strCardCode
                                .Item("U_PrgDate").Value = dtPrgDate
                                .Item("U_ItemCode").Value = strItemCode
                                .Item("U_Quantity").Value = dblQty
                                .Item("U_Dislike").Value = strDislike
                                .Item("U_Medical").Value = strMedical
                                .Item("U_FType").Value = strFType
                                .Item("U_SFood").Value = strSType
                                .Item("U_Select").Value = strSelect
                                .Item("U_Remarks").Value = strRemarks
                                .Item("U_Session").Value = strSession
                            End With
                            If oUserTable.Add <> 0 Then
                                Throw New Exception(oApplication.Company.GetLastErrorDescription)
                            End If
                        Else
                            Continue For
                        End If
                    End If
                ElseIf oUserTable.GetByKey(oRecordSet.Fields.Item(0).Value.ToString()) Then
                    With oUserTable.UserFields.Fields
                        .Item("U_ProgramID").Value = strProgramID
                        .Item("U_CardCode").Value = strCardCode
                        .Item("U_PrgDate").Value = dtPrgDate
                        .Item("U_ItemCode").Value = strItemCode
                        .Item("U_Quantity").Value = dblQty
                        .Item("U_Dislike").Value = strDislike
                        .Item("U_Medical").Value = strMedical
                        .Item("U_FType").Value = strFType
                        .Item("U_SFood").Value = strSType
                        .Item("U_Select").Value = strSelect
                        .Item("U_Remarks").Value = strRemarks
                        .Item("U_Session").Value = strSession
                    End With
                    If oUserTable.Update <> 0 Then
                        Throw New Exception(oApplication.Company.GetLastErrorDescription)
                    End If
                End If

            Next

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        Finally
            oUserTable = Nothing
        End Try
    End Sub

    Public Sub UpdateOrderQuantityBasedOnCalories(ByVal strDocEntry As String)
        Try
            Dim strQuery As String = String.Empty
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            strQuery = " Select T3.DocEntry,T3.VisOrder As LineNum,ISNULL(T2.U_Ratio,0) As 'Ratio' "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_CPR7] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN [@Z_OCRT] T2 On T2.U_Code = T1.U_BF "
            strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " And Convert(VarChar(8),T1.U_PrgDate,112) <= Convert(VarChar(8),T3.U_DelDate,112) "
            strQuery += " And T3.LineStatus = 'O' And ISNULL(T2.U_Ratio,0) - ISNULL(T3.Quantity,0) <> 0 "
            strQuery += " And T3.U_FType = T2.U_FType "
            strQuery += " Where T0.DocEntry = '" & strDocEntry & "'"
            strQuery += " And Convert(VarChar(8),T1.U_PrgDate,112) =  "
            strQuery += " (Select TOP 1 T10.U_PrgDate From [@Z_CPR7] T10 Where T10.DocEntry = '" & strDocEntry & "'"
            strQuery += " AND T10.U_PrgDate <= T3.U_DelDate Order By T10.U_PrgDate DESC) "
            strQuery += " UNION ALL  "
            strQuery += " Select T3.DocEntry,T3.VisOrder As LineNum,ISNULL(T2.U_Ratio,0) As 'Ratio' "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_CPR7] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN [@Z_OCRT] T2 On T2.U_Code = T1.U_LN "
            strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " And Convert(VarChar(8),T1.U_PrgDate,112) <= Convert(VarChar(8),T3.U_DelDate,112) "
            strQuery += " And T3.LineStatus = 'O' And ISNULL(T2.U_Ratio,0) - ISNULL(T3.Quantity,0) <> 0 "
            strQuery += " And T3.U_FType = T2.U_FType "
            strQuery += " Where T0.DocEntry = '" & strDocEntry & "'"
            strQuery += " And Convert(VarChar(8),T1.U_PrgDate,112) =  "
            strQuery += " (Select TOP 1 T10.U_PrgDate From [@Z_CPR7] T10 Where T10.DocEntry = '" & strDocEntry & "'"
            strQuery += " AND T10.U_PrgDate <= T3.U_DelDate Order By T10.U_PrgDate DESC) "
            strQuery += " UNION ALL  "
            strQuery += " Select T3.DocEntry,T3.VisOrder As LineNum,ISNULL(T2.U_Ratio,0) As 'Ratio' "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_CPR7] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN [@Z_OCRT] T2 On T2.U_Code = T1.U_LS "
            strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " And Convert(VarChar(8),T1.U_PrgDate,112) <= Convert(VarChar(8),T3.U_DelDate,112) "
            strQuery += " And T3.LineStatus = 'O' And ISNULL(T2.U_Ratio,0) - ISNULL(T3.Quantity,0) <> 0 "
            strQuery += " And T3.U_FType = T2.U_FType "
            strQuery += " Where T0.DocEntry = '" & strDocEntry & "'"
            strQuery += " And Convert(VarChar(8),T1.U_PrgDate,112) =  "
            strQuery += " (Select TOP 1 T10.U_PrgDate From [@Z_CPR7] T10 Where T10.DocEntry = '" & strDocEntry & "'"
            strQuery += " AND T10.U_PrgDate <= T3.U_DelDate Order By T10.U_PrgDate DESC) "
            strQuery += " UNION ALL  "
            strQuery += " Select T3.DocEntry,T3.VisOrder As LineNum,ISNULL(T2.U_Ratio,0) As 'Ratio' "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_CPR7] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN [@Z_OCRT] T2 On T2.U_Code = T1.U_SK "
            strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " And Convert(VarChar(8),T1.U_PrgDate,112) <= Convert(VarChar(8),T3.U_DelDate,112) "
            strQuery += " And T3.LineStatus = 'O' And ISNULL(T2.U_Ratio,0) - ISNULL(T3.Quantity,0) <> 0 "
            strQuery += " And T3.U_FType = T2.U_FType "
            strQuery += " Where T0.DocEntry = '" & strDocEntry & "'"
            strQuery += " And Convert(VarChar(8),T1.U_PrgDate,112) =  "
            strQuery += " (Select TOP 1 T10.U_PrgDate From [@Z_CPR7] T10 Where T10.DocEntry = '" & strDocEntry & "'"
            strQuery += " AND T10.U_PrgDate <= T3.U_DelDate Order By T10.U_PrgDate DESC) "
            strQuery += " UNION ALL  "
            strQuery += " Select T3.DocEntry,T3.VisOrder As LineNum,ISNULL(T2.U_Ratio,0) As 'Ratio' "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_CPR7] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN [@Z_OCRT] T2 On T2.U_Code = T1.U_DI "
            strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " And Convert(VarChar(8),T1.U_PrgDate,112) <= Convert(VarChar(8),T3.U_DelDate,112) "
            strQuery += " And T3.LineStatus = 'O' And ISNULL(T2.U_Ratio,0) - ISNULL(T3.Quantity,0) <> 0 "
            strQuery += " And T3.U_FType = T2.U_FType "
            strQuery += " Where T0.DocEntry = '" & strDocEntry & "'"
            strQuery += " And Convert(VarChar(8),T1.U_PrgDate,112) =  "
            strQuery += " (Select TOP 1 T10.U_PrgDate From [@Z_CPR7] T10 Where T10.DocEntry = '" & strDocEntry & "'"
            strQuery += " AND T10.U_PrgDate <= T3.U_DelDate Order By T10.U_PrgDate DESC) "
            strQuery += " UNION ALL  "
            strQuery += " Select T3.DocEntry,T3.VisOrder As LineNum,ISNULL(T2.U_Ratio,0) As 'Ratio' "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_CPR7] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN [@Z_OCRT] T2 On T2.U_Code = T1.U_DS "
            strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " And Convert(VarChar(8),T1.U_PrgDate,112) <= Convert(VarChar(8),T3.U_DelDate,112) "
            strQuery += " And T3.LineStatus = 'O' And ISNULL(T2.U_Ratio,0) - ISNULL(T3.Quantity,0) <> 0 "
            strQuery += " And T3.U_FType = T2.U_FType "
            strQuery += " Where T0.DocEntry = '" & strDocEntry & "'"
            strQuery += " And Convert(VarChar(8),T1.U_PrgDate,112) =  "
            strQuery += " (Select TOP 1 T10.U_PrgDate From [@Z_CPR7] T10 Where T10.DocEntry = '" & strDocEntry & "'"
            strQuery += " AND T10.U_PrgDate <= T3.U_DelDate Order By T10.U_PrgDate DESC) "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                Dim oOrder As SAPbobsCOM.Documents
                While Not oRecordSet.EoF
                    oOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                    Dim intDocEntry As Integer = oRecordSet.Fields.Item("DocEntry").Value
                    Dim intLine As Integer = oRecordSet.Fields.Item("LineNum").Value
                    Dim intStatus As Integer = 0
                    If oOrder.GetByKey(intDocEntry) Then
                        Dim blnUpdate As Boolean = False
                        For index As Integer = 0 To oOrder.Lines.Count - 1
                            If intLine = index Then
                                oOrder.Lines.SetCurrentLine(intLine)
                                oOrder.Lines.Quantity = CDbl(oRecordSet.Fields.Item("Ratio").Value)
                                blnUpdate = True
                            End If
                        Next
                        If blnUpdate Then
                            intStatus = oOrder.Update()
                        End If
                        If intStatus = 0 Then

                        End If
                    End If
                    oRecordSet.MoveNext()
                End While
            End If

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Public Sub UpdateOpenOrderAddresses(ByVal strCardCode As String)
        Try
            Dim strQuery As String = String.Empty
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim oRecordSet_U As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet_U = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            strQuery = " Select "
            strQuery += " T0.DocEntry,T0.LineNum,   "
            strQuery += " ISNULL(ISNULL(ISNULL(T3.U_Address,T4.U_Address),T5.ShipToDef),'') As 'U_Address',   "
            strQuery += " ISNULL(ISNULL(ISNULL(T3.U_Building,T4.U_Building),T5.MailBuildi),'') As 'U_Building'  , "
            strQuery += " (Select Top 1 State From CRD1 Where CardCode = T5.CardCode And AdresType = 'S'  "
            strQuery += " And Address = ISNULL(ISNULL(ISNULL(T3.U_Address,T4.U_Address),T5.ShipToDef),'')) As 'U_State' "
            strQuery += " From RDR1 T0 JOIN ORDR T1 On T0.DocEntry = T1.DocEntry   "
            strQuery += " LEFT OUTER JOIN [@Z_OCPR] T2 On T1.CardCode = T2.U_CardCode   "
            strQuery += " LEFT OUTER JOIN [@Z_CPR5] T3 On T2.DocEntry = T3.DocEntry   "
            strQuery += " AND Convert(VarChar(8),T0.U_DelDate,112) Between Convert(VarChar(8),T3.U_DelDate,112)   "
            strQuery += "  And Convert(VarChar(8),T3.U_TDelDate,112)  "
            strQuery += " And ((T3.U_BF = 'Y' AND T0.U_FType = 'BF')  "
            strQuery += " OR (T3.U_LN = 'Y' AND T0.U_FType = 'LN')  "
            strQuery += " OR (T3.U_LS = 'Y' AND T0.U_FType = 'LS')  "
            strQuery += " OR (T3.U_SK = 'Y' AND T0.U_FType = 'SK')  "
            strQuery += " OR (T3.U_DI = 'Y' AND T0.U_FType = 'DI')  "
            strQuery += " OR (T3.U_DS = 'Y' AND T0.U_FType = 'DS'))  "
            strQuery += " LEFT OUTER JOIN [@Z_CPR6] T4 On T2.DocEntry = T4.DocEntry  "
            strQuery += " And ((T4.U_BF = 'Y' AND T0.U_FType = 'BF')  "
            strQuery += " OR (T4.U_LN = 'Y' AND T0.U_FType = 'LN')  "
            strQuery += " OR (T4.U_LS = 'Y' AND T0.U_FType = 'LS')  "
            strQuery += " OR (T4.U_SK = 'Y' AND T0.U_FType = 'SK')  "
            strQuery += " OR (T4.U_DI = 'Y' AND T0.U_FType = 'DI')  "
            strQuery += " OR (T4.U_DS = 'Y' AND T0.U_FType = 'DS'))  "
            strQuery += " AND T4.U_Day = DatePart(DW,T0.U_DelDate)  "
            strQuery += " JOIN OCRD T5 On T5.CardCode = T1.CardCode  "
            strQuery += " And T5.CardCode = '" & strCardCode & "'"
            strQuery += " And T0.LineStatus = 'O'"
            strQuery += " And "
            strQuery += " ("
            strQuery += " (ISNULL(ISNULL(ISNULL(T3.U_Address,T4.U_Address),T5.ShipToDef),'') <> T0.U_Address)"
            strQuery += " Or"
            strQuery += " (ISNULL(ISNULL(ISNULL(T3.U_Building,T4.U_Building),T5.MailBuildi),'') <> Convert(VarChar(500),T0.U_Building))"
            strQuery += " Or"
            strQuery += " (Select Top 1 State From CRD1 Where CardCode = T5.CardCode And AdresType = 'S' "
            strQuery += " And Address = ISNULL(ISNULL(ISNULL(T3.U_Address,T4.U_Address),T5.ShipToDef),'')) <> ISNULL(T0.U_State,'')     "
            strQuery += " )"


            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    Dim strDocEntry, strLineNo, strAddress, strBuilding, strState As String

                    strDocEntry = oRecordSet.Fields.Item("DocEntry").Value
                    strLineNo = oRecordSet.Fields.Item("LineNum").Value
                    strAddress = oRecordSet.Fields.Item("U_Address").Value
                    strBuilding = oRecordSet.Fields.Item("U_Building").Value
                    strState = oRecordSet.Fields.Item("U_State").Value

                    strQuery = "Update RDR1 "
                    strQuery += " Set U_Address = '" & strAddress & "'"
                    strQuery += " , U_Building = N'" & strBuilding & "'"
                    strQuery += " , U_State = '" & strState & "'"
                    strQuery += " Where DocEntry = '" & strDocEntry & "'"
                    strQuery += " And LineNum = '" & strLineNo & "'"
                    oRecordSet_U.DoQuery(strQuery)

                    oRecordSet.MoveNext()

                End While
            End If

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Public Sub CloseOrderQuantityRemoveSuspendDates(ByVal strPRDocEntry As String)
        Try
            Dim ohtConOrder As Hashtable
            ohtConOrder = New Hashtable

            Dim strQuery As String = String.Empty
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim oRecordSet_U As SAPbobsCOM.Recordset
            oRecordSet_U = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            strQuery += " Select T3.DocEntry,T3.VisOrder As LineNum,'E' As 'Type',T1.LineId "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_CPR4] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " And Convert(VarChar(8),T3.U_DelDate,112) BETWEEN Convert(VarChar(8),T1.U_ExDate,112) AND Convert(VarChar(8),T1.U_ExDate,112) "
            strQuery += " And T3.LineStatus = 'O' "
            strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
            strQuery += " UNION ALL  "
            strQuery += " Select T3.DocEntry,T3.VisOrder As LineNum,'R' As 'Type',T1.LineId "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_CPR8] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " And Convert(VarChar(8),T3.U_DelDate,112) BETWEEN Convert(VarChar(8),T1.U_FDate,112) AND Convert(VarChar(8),T1.U_TDate,112) "
            strQuery += " And T3.LineStatus = 'O' "
            strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
            strQuery += " UNION ALL  "
            strQuery += " Select T3.DocEntry,T3.VisOrder As LineNum,'S' As 'Type',T1.LineId "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_CPR9] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " And "
            strQuery += " ( "
            strQuery += " (Convert(VarChar(8),T3.U_DelDate,112) >= Convert(VarChar(8),T1.U_FDate,112) AND T1.U_TDate Is Null) "
            strQuery += " OR "
            strQuery += " (Convert(VarChar(8),T3.U_DelDate,112) BETWEEN Convert(VarChar(8),T1.U_FDate,112) AND Convert(VarChar(8),T1.U_TDate,112)) "
            strQuery += " ) "
            strQuery += " And T3.LineStatus = 'O' "
            strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"

            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                Dim oOrder As SAPbobsCOM.Documents
                While Not oRecordSet.EoF
                    oOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                    Dim intDocEntry As Integer = oRecordSet.Fields.Item("DocEntry").Value
                    Dim intLine As Integer = oRecordSet.Fields.Item("LineNum").Value
                    Dim strPRLine As String = oRecordSet.Fields.Item("LineId").Value

                    Dim intStatus As Integer = 0
                    If oOrder.GetByKey(intDocEntry) Then
                        Dim blnUpdate As Boolean = False
                        For index As Integer = 0 To oOrder.Lines.Count - 1
                            If intLine = index Then
                                oOrder.Lines.SetCurrentLine(intLine)
                                Dim strIsCon As String = oOrder.Lines.UserFields.Fields.Item("U_IsCon").Value
                                Dim strConDate As String = ""
                                If strIsCon = "Y" Then
                                    strConDate = CDate(oOrder.Lines.UserFields.Fields.Item("U_ConDate").Value).ToString("yyyyMMdd")
                                End If
                                If oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Open Then
                                    If strIsCon = "Y" Then
                                        If Not ohtConOrder.ContainsKey(intDocEntry) Then
                                            If strConDate = CDate(oOrder.Lines.UserFields.Fields.Item("U_DelDate").Value).ToString("yyyyMMdd") Then
                                                ohtConOrder.Add(intDocEntry, strConDate)
                                            End If
                                        End If
                                    End If
                                    oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close
                                    'oOrder.Lines.UserFields.Fields.Item("U_CanFrom").Value = oRecordSet.Fields.Item("Type").Value
                                    blnUpdate = True
                                    Exit For
                                End If
                            End If
                        Next
                        If blnUpdate Then
                            intStatus = oOrder.Update()
                        End If
                        If intStatus = 0 Then

                            strQuery = "Update RDR1 SET U_CanFrom = '" & oRecordSet.Fields.Item("Type").Value & "' "
                            strQuery += " Where DocEntry = '" & intDocEntry & "'"
                            strQuery += " And VisOrder = '" & intLine & "'"
                            oRecordSet_U.DoQuery(strQuery)

                            If oRecordSet.Fields.Item("Type").Value = "E" Then
                                strQuery = "Update [@Z_CPR4] SET U_Applied = 'Y' "
                                strQuery += " Where DocEntry = '" & strPRDocEntry & "'"
                                strQuery += " And LineId = '" & strPRLine & "'"
                                oRecordSet_U.DoQuery(strQuery)
                            ElseIf oRecordSet.Fields.Item("Type").Value = "R" Then
                                strQuery = "Update [@Z_CPR8] SET U_Applied = 'Y' "
                                strQuery += " Where DocEntry = '" & strPRDocEntry & "'"
                                strQuery += " And LineId = '" & strPRLine & "'"
                                oRecordSet_U.DoQuery(strQuery)
                            ElseIf oRecordSet.Fields.Item("Type").Value = "S" Then
                                strQuery = "Update [@Z_CPR9] SET U_Applied = 'Y' "
                                strQuery += " Where DocEntry = '" & strPRDocEntry & "'"
                                strQuery += " And LineId = '" & strPRLine & "'"
                                oRecordSet_U.DoQuery(strQuery)
                            End If

                        End If
                    End If

                    oRecordSet.MoveNext()
                End While
            End If

            'Consolidate Date Update....
            If Not IsNothing(ohtConOrder) Then
                If ohtConOrder.Count > 0 Then
                    Dim Item As DictionaryEntry
                    Dim oRecordSet_M As SAPbobsCOM.Recordset
                    oRecordSet_M = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)


                    For Each Item In ohtConOrder

                        strQuery = "Select Min(U_DelDate) From RDR1 Where LineStatus = 'O' "
                        strQuery += " And DocEntry = '" & Item.Key & "'"
                        strQuery += " And Convert(VarChar(8),U_ConDate,112) = '" & Item.Value & "'"
                        strQuery += " And U_ConDate Is Not Null "
                        oRecordSet_M.DoQuery(strQuery)

                        If Not oRecordSet_M.EoF Then

                            strQuery = "Select VisOrder As LineNum From RDR1 Where LineStatus = 'O' "
                            strQuery += " And DocEntry = '" & Item.Key & "'"
                            strQuery += " And Convert(VarChar(8),U_ConDate,112) = '" & Item.Value & "'"
                            strQuery += " And U_ConDate Is Not Null "
                            oRecordSet.DoQuery(strQuery)
                            If Not oRecordSet.EoF Then

                                Dim oOrder As SAPbobsCOM.Documents
                                While Not oRecordSet.EoF
                                    oOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                                    Dim intLine As Integer = oRecordSet.Fields.Item("LineNum").Value
                                    Dim intStatus As Integer = 0
                                    If oOrder.GetByKey(CInt(Item.Key)) Then
                                        Dim blnUpdate As Boolean = False
                                        For index As Integer = 0 To oOrder.Lines.Count - 1
                                            If intLine = index Then
                                                oOrder.Lines.SetCurrentLine(intLine)
                                                If oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Open Then
                                                    oOrder.Lines.ShipDate = CDate(oRecordSet_M.Fields.Item(0).Value)
                                                    oOrder.Lines.UserFields.Fields.Item("U_ConDate").Value = CDate(oRecordSet_M.Fields.Item(0).Value)
                                                    blnUpdate = True
                                                End If
                                            End If
                                        Next
                                        If blnUpdate Then
                                            intStatus = oOrder.Update()
                                        End If
                                    End If
                                    oRecordSet.MoveNext()
                                End While

                            End If
                        End If

                    Next
                End If
            End If


        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Public Sub CloseOrderQuantityRemoveSuspendDates_P(ByVal strPRDocEntry As String, ByVal strCardCode As String)
        Try
            Dim ohtConOrder As Hashtable
            ohtConOrder = New Hashtable

            Dim strQuery As String = String.Empty
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim oRecordSet_L As SAPbobsCOM.Recordset
            oRecordSet_L = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim oRecordSet_U As SAPbobsCOM.Recordset
            oRecordSet_U = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim oRecordSet_P As SAPbobsCOM.Recordset
            oRecordSet_P = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            strQuery = "  Select Distinct T4.DocEntry "
            strQuery += " From "
            strQuery += " ( "
            strQuery += " Select T3.DocEntry "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_CPR4] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " And Convert(VarChar(8),T3.U_DelDate,112) BETWEEN Convert(VarChar(8),T1.U_ExDate,112) AND Convert(VarChar(8),T1.U_ExDate,112) "
            strQuery += " And T3.LineStatus = 'O' "
            strQuery += " And ISNULL(T1.U_Include,'N') = 'N' "
            strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
            strQuery += " UNION ALL  "
            strQuery += " Select T3.DocEntry "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " And "
            strQuery += " ( "
            strQuery += " (DatePart(dw,T3.U_DelDate) = 1 AND T0.U_Sunday = 'Y') OR "
            strQuery += " (DatePart(dw,T3.U_DelDate) = 2 AND T0.U_Monday = 'Y') OR "
            strQuery += " (DatePart(dw,T3.U_DelDate) = 3 AND T0.U_Tuesday = 'Y') OR "
            strQuery += " (DatePart(dw,T3.U_DelDate) = 4 AND T0.U_Wednesday = 'Y') OR "
            strQuery += " (DatePart(dw,T3.U_DelDate) = 5 AND T0.U_Thursday = 'Y') OR "
            strQuery += " (DatePart(dw,T3.U_DelDate) = 6 AND T0.U_Friday = 'Y') OR "
            strQuery += " (DatePart(dw,T3.U_DelDate) = 7 AND T0.U_Saturday = 'Y') "
            strQuery += " ) "
            strQuery += " And T3.LineStatus = 'O' "
            strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
            strQuery += " And T3.U_DelDate Not In  "
            strQuery += " ( "
            strQuery += " Select T0.U_ExDate From [@Z_CPR4] T0 "
            strQuery += " JOIN [@Z_OCPR] T1 On T0.DocEntry = T1.DocEntry   "
            strQuery += " Where ISNULL(T0.U_Include,'N') = 'Y' "
            strQuery += " And T1.U_CardCode = '" & strCardCode & "' "
            strQuery += " ) "
            strQuery += " UNION ALL  "
            strQuery += " Select T3.DocEntry "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_CPR8] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " And Convert(VarChar(8),T3.U_DelDate,112) BETWEEN Convert(VarChar(8),T1.U_FDate,112) AND Convert(VarChar(8),T1.U_TDate,112) "
            strQuery += " And T3.LineStatus = 'O' "
            strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
            strQuery += " UNION ALL  "
            strQuery += " Select T3.DocEntry "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_CPR9] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " And "
            strQuery += " ( "
            strQuery += " (Convert(VarChar(8),T3.U_DelDate,112) >= Convert(VarChar(8),T1.U_FDate,112) AND T1.U_TDate Is Null) "
            strQuery += " OR "
            strQuery += " (Convert(VarChar(8),T3.U_DelDate,112) BETWEEN Convert(VarChar(8),T1.U_FDate,112) AND Convert(VarChar(8),T1.U_TDate,112)) "
            strQuery += " ) "
            strQuery += " And T3.LineStatus = 'O' "
            strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
            'strQuery += " UNION ALL  "
            'strQuery += " Select DISTINCT T3.DocEntry "
            'strQuery += " From [@Z_OCPR] T0  "
            'strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
            'strQuery += " JOIN [@Z_OCPM] T4 On T4.U_CardCode = T3.BaseCard "
            'strQuery += " And T4.DocEntry = T3.U_ProgramID "
            'strQuery += " And Convert(VarChar(8),T3.U_DelDate,112) NOT BETWEEN Convert(VarChar(8),T4.U_PFromDate,112) AND Convert(VarChar(8),T4.U_PToDate,112)   "
            'strQuery += " And T3.LineStatus = 'O' "
            'strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
            strQuery += " ) T4  "

            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                Dim oOrder As SAPbobsCOM.Documents
                While Not oRecordSet.EoF
                    oOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)

                    Dim intDocEntry As Integer = oRecordSet.Fields.Item("DocEntry").Value
                    Dim intStatus As Integer = 0
                    If oOrder.GetByKey(intDocEntry) Then

                        Dim blnUpdate As Boolean = False
                        Dim blnClose As Boolean = False
                        oRecordSet_L = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                        oRecordSet_P = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

                        strQuery = " Select T3.DocEntry,T3.VisOrder As LineNum,'E' As 'Type',T1.LineId "
                        strQuery += " From [@Z_OCPR] T0  "
                        strQuery += " JOIN [@Z_CPR4] T1 On T0.DocEntry = T1.DocEntry "
                        strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
                        strQuery += " And Convert(VarChar(8),T3.U_DelDate,112) BETWEEN Convert(VarChar(8),T1.U_ExDate,112) AND Convert(VarChar(8),T1.U_ExDate,112) "
                        strQuery += " And T3.LineStatus = 'O' "
                        strQuery += " And ISNULL(T1.U_Include,'N') = 'N' "
                        strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
                        strQuery += " And T3.DocEntry = '" & intDocEntry.ToString & "'"
                        strQuery += " UNION ALL  "
                        strQuery += " Select T3.DocEntry,T3.VisOrder As LineNum,'ED' As 'Type',T3.VisOrder As 'LineId' "
                        strQuery += " From [@Z_OCPR] T0  "
                        strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
                        strQuery += " And "
                        strQuery += " ( "
                        strQuery += " (DatePart(dw,T3.U_DelDate) = 1 AND T0.U_Sunday = 'Y') OR "
                        strQuery += " (DatePart(dw,T3.U_DelDate) = 2 AND T0.U_Monday = 'Y') OR "
                        strQuery += " (DatePart(dw,T3.U_DelDate) = 3 AND T0.U_Tuesday = 'Y') OR "
                        strQuery += " (DatePart(dw,T3.U_DelDate) = 4 AND T0.U_Wednesday = 'Y') OR "
                        strQuery += " (DatePart(dw,T3.U_DelDate) = 5 AND T0.U_Thursday = 'Y') OR "
                        strQuery += " (DatePart(dw,T3.U_DelDate) = 6 AND T0.U_Friday = 'Y') OR "
                        strQuery += " (DatePart(dw,T3.U_DelDate) = 7 AND T0.U_Saturday = 'Y') "
                        strQuery += " ) "
                        strQuery += " And T3.LineStatus = 'O' "
                        strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
                        strQuery += " And T3.DocEntry = '" & intDocEntry.ToString & "'"
                        strQuery += " And T3.U_DelDate > GetDate() "
                        strQuery += " And T3.U_DelDate Not In  "
                        strQuery += " ( "
                        strQuery += " Select T0.U_ExDate From [@Z_CPR4] T0 "
                        strQuery += " JOIN [@Z_OCPR] T1 On T0.DocEntry = T1.DocEntry   "
                        strQuery += " Where ISNULL(T0.U_Include,'N') = 'Y' "
                        strQuery += " And T1.U_CardCode = '" & strCardCode & "' "
                        strQuery += " ) "
                        strQuery += " UNION ALL  "
                        strQuery += " Select T3.DocEntry,T3.VisOrder As LineNum,'R' As 'Type',T1.LineId "
                        strQuery += " From [@Z_OCPR] T0  "
                        strQuery += " JOIN [@Z_CPR8] T1 On T0.DocEntry = T1.DocEntry "
                        strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
                        strQuery += " And Convert(VarChar(8),T3.U_DelDate,112) BETWEEN Convert(VarChar(8),T1.U_FDate,112) AND Convert(VarChar(8),T1.U_TDate,112) "
                        strQuery += " And T3.LineStatus = 'O' "
                        strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
                        strQuery += " And T3.DocEntry = '" & intDocEntry.ToString & "'"
                        strQuery += " UNION ALL  "
                        strQuery += " Select T3.DocEntry,T3.VisOrder As LineNum,'S' As 'Type',T1.LineId "
                        strQuery += " From [@Z_OCPR] T0  "
                        strQuery += " JOIN [@Z_CPR9] T1 On T0.DocEntry = T1.DocEntry "
                        strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
                        strQuery += " And "
                        strQuery += " ( "
                        strQuery += " (Convert(VarChar(8),T3.U_DelDate,112) >= Convert(VarChar(8),T1.U_FDate,112) AND T1.U_TDate Is Null) "
                        strQuery += " OR "
                        strQuery += " (Convert(VarChar(8),T3.U_DelDate,112) BETWEEN Convert(VarChar(8),T1.U_FDate,112) AND Convert(VarChar(8),T1.U_TDate,112)) "
                        strQuery += " ) "
                        strQuery += " And T3.LineStatus = 'O' "
                        strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
                        strQuery += " And T3.DocEntry = '" & intDocEntry.ToString & "'"
                        'strQuery += " UNION ALL "
                        'strQuery += " Select T3.DocEntry,T3.VisOrder As LineNum,'' As 'Type',T3.LineNum "
                        'strQuery += " From [@Z_OCPR] T0  "
                        'strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
                        'strQuery += " JOIN [@Z_OCPM] T4 On T4.U_CardCode = T3.BaseCard "
                        'strQuery += " And T4.DocEntry = T3.U_ProgramID "
                        'strQuery += " And Convert(VarChar(8),T3.U_DelDate,112) NOT BETWEEN Convert(VarChar(8),T4.U_PFromDate,112) AND Convert(VarChar(8),T4.U_PToDate,112) "
                        'strQuery += " And T3.LineStatus = 'O' "
                        'strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
                        'strQuery += " And T3.DocEntry = '" & intDocEntry.ToString & "'"

                        oRecordSet_L.DoQuery(strQuery)
                        If Not oRecordSet_L.EoF Then

                            strQuery = " Select Distinct T3.U_DelDate,T3.U_ConDate,T3.U_IsCon "
                            strQuery += " From [@Z_OCPR] T0  "
                            strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
                            strQuery += " And T3.LineStatus = 'O' "
                            strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
                            strQuery += " And T3.DocEntry = '" & intDocEntry.ToString & "'"

                            oRecordSet_P.DoQuery(strQuery)
                            If oRecordSet_P.RecordCount = 1 Then

                                Dim strIsCon As String = oRecordSet_P.Fields.Item("U_IsCon").Value
                                Dim strConDate As String = ""
                                If strIsCon = "Y" Then
                                    strConDate = CDate(oRecordSet_P.Fields.Item("U_ConDate").Value).ToString("yyyyMMdd")
                                End If

                                If oOrder.DocumentStatus = BoStatus.bost_Open Then
                                    If strIsCon = "Y" Then
                                        If Not ohtConOrder.ContainsKey(strConDate) Then
                                            If strConDate = CDate(oRecordSet_P.Fields.Item("U_ConDate").Value).ToString("yyyyMMdd") Then
                                                ohtConOrder.Add(strConDate, oOrder.CardCode)
                                            End If
                                        End If
                                    End If
                                    blnClose = True
                                End If

                            Else

                                While Not oRecordSet_L.EoF

                                    Dim intLine As Integer = oRecordSet_L.Fields.Item("LineNum").Value
                                    Dim strPRLine As String = oRecordSet_L.Fields.Item("LineId").Value

                                    oOrder.Lines.SetCurrentLine(intLine)

                                    Dim strIsCon As String = oOrder.Lines.UserFields.Fields.Item("U_IsCon").Value
                                    Dim strConDate As String = ""
                                    If strIsCon = "Y" Then
                                        strConDate = CDate(oOrder.Lines.UserFields.Fields.Item("U_ConDate").Value).ToString("yyyyMMdd")
                                    End If

                                    If oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Open Then
                                        If strIsCon = "Y" Then
                                            If Not ohtConOrder.ContainsKey(strConDate) Then
                                                If strConDate = CDate(oOrder.Lines.UserFields.Fields.Item("U_DelDate").Value).ToString("yyyyMMdd") Then
                                                    ohtConOrder.Add(strConDate, oOrder.CardCode)
                                                End If
                                            End If
                                        End If
                                        oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close
                                        blnUpdate = True
                                    End If
                                    oRecordSet_L.MoveNext()

                                End While

                            End If

                        End If

                        If blnUpdate Then
                            intStatus = oOrder.Update()
                        ElseIf blnClose Then
                            intStatus = oOrder.Close()
                        End If

                        'Reseting to First record to Update
                        oRecordSet_L.MoveFirst()

                        If intStatus = 0 And (blnUpdate Or blnClose) Then
                            If Not oRecordSet_L.EoF Then
                                While Not oRecordSet_L.EoF
                                    Dim intLine As Integer = oRecordSet_L.Fields.Item("LineNum").Value
                                    Dim strPRLine As String = oRecordSet_L.Fields.Item("LineId").Value

                                    If oRecordSet_L.Fields.Item("Type").Value <> "" Then
                                        Dim strType As String = oRecordSet_L.Fields.Item("Type").Value
                                        If strType <> "ED" Then
                                            strQuery = "Update RDR1 SET U_CanFrom = '" & oRecordSet_L.Fields.Item("Type").Value & "' "
                                            strQuery += " Where DocEntry = '" & intDocEntry & "'"
                                            strQuery += " And VisOrder = '" & intLine & "'"
                                            oRecordSet_U.DoQuery(strQuery)
                                        Else
                                            strQuery = "Update RDR1 SET U_CanFrom = 'E' "
                                            strQuery += " Where DocEntry = '" & intDocEntry & "'"
                                            strQuery += " And VisOrder = '" & intLine & "'"
                                            oRecordSet_U.DoQuery(strQuery)
                                        End If

                                    Else
                                        strQuery = "Update RDR1 SET FreeTxt = 'Cancelled due to Overlapping...' "
                                        strQuery += " Where DocEntry = '" & intDocEntry & "'"
                                        strQuery += " And VisOrder = '" & intLine & "'"
                                        oRecordSet_U.DoQuery(strQuery)
                                    End If

                                    If oRecordSet_L.Fields.Item("Type").Value = "E" Then
                                        strQuery = "Update [@Z_CPR4] SET U_Applied = 'Y' "
                                        strQuery += " Where DocEntry = '" & strPRDocEntry & "'"
                                        strQuery += " And LineId = '" & strPRLine & "'"
                                        oRecordSet_U.DoQuery(strQuery)
                                    ElseIf oRecordSet_L.Fields.Item("Type").Value = "R" Then
                                        strQuery = "Update [@Z_CPR8] SET U_Applied = 'Y' "
                                        strQuery += " Where DocEntry = '" & strPRDocEntry & "'"
                                        strQuery += " And LineId = '" & strPRLine & "'"
                                        oRecordSet_U.DoQuery(strQuery)
                                    ElseIf oRecordSet_L.Fields.Item("Type").Value = "S" Then
                                        strQuery = "Update [@Z_CPR9] SET U_Applied = 'Y' "
                                        strQuery += " Where DocEntry = '" & strPRDocEntry & "'"
                                        strQuery += " And LineId = '" & strPRLine & "'"
                                        oRecordSet_U.DoQuery(strQuery)
                                    End If

                                    oRecordSet_L.MoveNext()
                                End While
                            End If
                        End If
                    End If

                    oRecordSet.MoveNext()
                End While
            End If

            'Consolidate Date Update....
            If Not IsNothing(ohtConOrder) Then
                If ohtConOrder.Count > 0 Then
                    Dim Item As DictionaryEntry
                    Dim oOrder As SAPbobsCOM.Documents
                    Dim oRecordSet_M As SAPbobsCOM.Recordset
                    Dim oRecordSet_O As SAPbobsCOM.Recordset
                    oRecordSet_M = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)


                    For Each Item In ohtConOrder
                        oOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)

                        oRecordSet_M = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                        oRecordSet_L = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                        oRecordSet_O = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

                        strQuery = "Select Min(U_DelDate) From RDR1 Where LineStatus = 'O' "
                        strQuery += " And BaseCard = '" & Item.Value & "'"
                        strQuery += " And Convert(VarChar(8),U_ConDate,112) = '" & Item.Key & "'"
                        strQuery += " And U_ConDate Is Not Null "

                        oRecordSet_M.DoQuery(strQuery)
                        If Not oRecordSet_M.EoF Then

                            strQuery = "Select Distinct DocEntry From RDR1 Where LineStatus = 'O' "
                            strQuery += " And BaseCard = '" & Item.Value & "'"
                            strQuery += " And Convert(VarChar(8),U_ConDate,112) = '" & Item.Key & "'"
                            strQuery += " And U_ConDate Is Not Null "
                            oRecordSet_O.DoQuery(strQuery)
                            If Not oRecordSet_O.EoF Then

                                While Not oRecordSet_O.EoF
                                    Dim blnUpdate As Boolean = False
                                    Dim intStatus As Integer = 0

                                    If oOrder.GetByKey(CInt(oRecordSet_O.Fields.Item("DocEntry").Value)) Then
                                        strQuery = "Select VisOrder As LineNum From RDR1 Where LineStatus = 'O' "
                                        strQuery += " And BaseCard = '" & Item.Value & "'"
                                        strQuery += " And Convert(VarChar(8),U_ConDate,112) = '" & Item.Key & "'"
                                        strQuery += " And U_ConDate Is Not Null "
                                        oRecordSet_L.DoQuery(strQuery)
                                        If Not oRecordSet_L.EoF Then
                                            While Not oRecordSet_L.EoF
                                                Dim intLine As Integer = oRecordSet_L.Fields.Item("LineNum").Value
                                                oOrder.Lines.SetCurrentLine(intLine)
                                                If oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Open Then
                                                    oOrder.Lines.ShipDate = CDate(oRecordSet_M.Fields.Item(0).Value)
                                                    oOrder.Lines.UserFields.Fields.Item("U_ConDate").Value = CDate(oRecordSet_M.Fields.Item(0).Value)
                                                    blnUpdate = True
                                                End If
                                                oRecordSet_L.MoveNext()
                                            End While
                                        End If
                                    End If

                                    If blnUpdate Then
                                        intStatus = oOrder.Update()
                                    End If

                                    oRecordSet_O.MoveNext()
                                End While

                            End If

                        End If
                    Next
                End If
            End If


        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Public Sub CancelDeliveryQuantityRemoveSuspendDatesAndExclude(ByVal strPRDocEntry As String, ByVal strCardCode As String)
        Try
            Dim strQuery As String = String.Empty
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            Dim oRecordSet_U As SAPbobsCOM.Recordset
            oRecordSet_U = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            strQuery = " Select T3.DocEntry,'E' As 'Type',T1.LineId "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_CPR4] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN DLN1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " JOIN ODLN T4 On T3.DocEntry = T4.DocEntry "
            strQuery += " And Convert(VarChar(8),T3.U_DelDate,112) BETWEEN Convert(VarChar(8),T1.U_ExDate,112) AND Convert(VarChar(8),T1.U_ExDate,112) "
            strQuery += " And T3.LineStatus = 'O' And (ISNULL(T4.U_InvRef,'') = '') "
            strQuery += " And ISNULL(T1.U_Include,'N') = 'N' "
            strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
            strQuery += " UNION ALL  "
            strQuery += " Select T3.DocEntry,'ED' As 'Type',T3.LineNum As 'LineId' "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN DLN1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " JOIN ODLN T4 On T3.DocEntry = T4.DocEntry "
            strQuery += " And "
            strQuery += " ( "
            strQuery += " (DatePart(dw,T3.U_DelDate) = 1 AND T0.U_Sunday = 'Y') OR "
            strQuery += " (DatePart(dw,T3.U_DelDate) = 2 AND T0.U_Monday = 'Y') OR "
            strQuery += " (DatePart(dw,T3.U_DelDate) = 3 AND T0.U_Tuesday = 'Y') OR "
            strQuery += " (DatePart(dw,T3.U_DelDate) = 4 AND T0.U_Wednesday = 'Y') OR "
            strQuery += " (DatePart(dw,T3.U_DelDate) = 5 AND T0.U_Thursday = 'Y') OR "
            strQuery += " (DatePart(dw,T3.U_DelDate) = 6 AND T0.U_Friday = 'Y') OR "
            strQuery += " (DatePart(dw,T3.U_DelDate) = 7 AND T0.U_Saturday = 'Y') "
            strQuery += " ) "
            strQuery += " And T3.LineStatus = 'O' And (ISNULL(T4.U_InvRef,'') = '') "
            strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
            strQuery += " And T3.U_DelDate > GetDate() "
            strQuery += " And T3.U_DelDate Not In  "
            strQuery += " ( "
            strQuery += " Select T0.U_ExDate From [@Z_CPR4] T0 "
            strQuery += " JOIN [@Z_OCPR] T1 On T0.DocEntry = T1.DocEntry   "
            strQuery += " Where ISNULL(T0.U_Include,'N') = 'Y' "
            strQuery += " And T1.U_CardCode = '" & strCardCode & "' "
            strQuery += " ) "
            strQuery += " UNION ALL  "
            strQuery += " Select T3.DocEntry,'R' As 'Type',T1.LineId "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_CPR8] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN DLN1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " JOIN ODLN T4 On T3.DocEntry = T4.DocEntry "
            strQuery += " And Convert(VarChar(8),T3.U_DelDate,112) BETWEEN Convert(VarChar(8),T1.U_FDate,112) AND Convert(VarChar(8),T1.U_TDate,112) "
            strQuery += " And ((T3.LineStatus = 'O') And (ISNULL(T4.U_InvRef,'') = '')) "
            strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
            strQuery += " UNION ALL  "
            strQuery += " Select T3.DocEntry,'S' As 'Type',T1.LineId "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_CPR9] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN DLN1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " JOIN ODLN T4 On T3.DocEntry = T4.DocEntry "
            strQuery += " And "
            strQuery += " ( "
            strQuery += " (Convert(VarChar(8),T3.U_DelDate,112) >= Convert(VarChar(8),T1.U_FDate,112) AND T1.U_TDate Is Null) "
            strQuery += " OR "
            strQuery += " (Convert(VarChar(8),T3.U_DelDate,112) BETWEEN Convert(VarChar(8),T1.U_FDate,112) AND Convert(VarChar(8),T1.U_TDate,112)) "
            strQuery += " ) "
            strQuery += " And ((T3.LineStatus = 'O') And (ISNULL(T4.U_InvRef,'') = '')) "
            strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
            'strQuery += " UNION ALL "
            'strQuery += " Select T3.DocEntry,'O' As 'Type',T3.LineNum As 'LineId' "
            'strQuery += " From [@Z_OCPR] T0  "
            'strQuery += " JOIN DLN1 T3 ON T3.BaseCard = T0.U_CardCode "
            'strQuery += " JOIN ODLN T4 On T3.DocEntry = T4.DocEntry "
            'strQuery += " JOIN [@Z_OCPM] T5 On T5.U_CardCode = T3.BaseCard "
            'strQuery += " And T5.DocEntry = T3.U_ProgramID "
            'strQuery += " And Convert(VarChar(8),T3.U_DelDate,112) NOT BETWEEN Convert(VarChar(8),T5.U_PFromDate,112) AND Convert(VarChar(8),T5.U_PToDate,112)   "
            'strQuery += " And T3.LineStatus = 'O' AND  (ISNULL(T4.U_InvRef,'') = '') "
            'strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"

            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                Dim oDelivery As SAPbobsCOM.Documents
                Dim oDelivery_C As SAPbobsCOM.Documents
                While Not oRecordSet.EoF
                    oDelivery = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                    oDelivery_C = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                    Dim intDocEntry As Integer = oRecordSet.Fields.Item("DocEntry").Value
                    Dim strPRLine As String = oRecordSet.Fields.Item("LineId").Value

                    Dim intStatus As Integer = 0
                    If oDelivery.GetByKey(intDocEntry) Then
                        Dim blnUpdate As Boolean = False
                        If oDelivery.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                            If oRecordSet.Fields.Item("Type").Value = "R" Then
                                intStatus = oDelivery.Close()

                                'oDelivery_C = oDelivery.CreateCancellationDocument()
                                'intStatus = oDelivery_C.Add()

                                If intStatus = 0 Then

                                    strQuery = "Update DLN1 SET U_CanFrom = '" & oRecordSet.Fields.Item("Type").Value & "'"
                                    strQuery += " Where DocEntry = '" & intDocEntry & "'"
                                    oRecordSet_U.DoQuery(strQuery)

                                    strQuery = "Update [@Z_CPR8] SET U_Applied = 'Y' "
                                    strQuery += " Where DocEntry = '" & strPRDocEntry & "'"
                                    strQuery += " And LineId = '" & strPRLine & "'"
                                    oRecordSet_U.DoQuery(strQuery)

                                    strQuery += " Exec PROCON_UPDATEORDERDAYS_u '" & oDelivery.Lines.BaseEntry.ToString() & "'"
                                    oRecordSet_U.DoQuery(strQuery)

                                End If
                            ElseIf oRecordSet.Fields.Item("Type").Value = "S" Then
                                oDelivery_C = oDelivery.CreateCancellationDocument()
                                intStatus = oDelivery_C.Add()
                                If intStatus = 0 Then

                                    strQuery = "Update DLN1 SET U_CanFrom = '" & oRecordSet.Fields.Item("Type").Value & "'"
                                    strQuery += " Where DocEntry = '" & intDocEntry & "'"
                                    oRecordSet_U.DoQuery(strQuery)

                                    strQuery = "Update [@Z_CPR9] SET U_Applied = 'Y' "
                                    strQuery += " Where DocEntry = '" & strPRDocEntry & "'"
                                    strQuery += " And LineId = '" & strPRLine & "'"
                                    oRecordSet_U.DoQuery(strQuery)

                                End If
                            ElseIf oRecordSet.Fields.Item("Type").Value = "E" Or oRecordSet.Fields.Item("Type").Value = "ED" Then
                                oDelivery_C = oDelivery.CreateCancellationDocument()
                                intStatus = oDelivery_C.Add()
                                If intStatus = 0 Then

                                    Dim strType As String = oRecordSet.Fields.Item("Type").Value
                                    If strType = "E" Then
                                        strQuery = "Update DLN1 SET U_CanFrom = '" & oRecordSet.Fields.Item("Type").Value & "'"
                                        strQuery += " Where DocEntry = '" & intDocEntry & "'"
                                        oRecordSet_U.DoQuery(strQuery)
                                    Else
                                        strQuery = "Update DLN1 SET U_CanFrom = 'E' "
                                        strQuery += " Where DocEntry = '" & intDocEntry & "'"
                                        oRecordSet_U.DoQuery(strQuery)
                                    End If

                                    If strType = "E" Then
                                        strQuery = "Update [@Z_CPR9] SET U_Applied = 'Y' "
                                        strQuery += " Where DocEntry = '" & strPRDocEntry & "'"
                                        strQuery += " And LineId = '" & strPRLine & "'"
                                        oRecordSet_U.DoQuery(strQuery)
                                    End If

                                End If
                            Else
                                oDelivery_C = oDelivery.CreateCancellationDocument()
                                intStatus = oDelivery_C.Add()
                                If intStatus = 0 Then
                                    strQuery = "Update DLN1 SET FreeTxt = 'Cancelled Due to Overlapping...' "
                                    strQuery += " Where DocEntry = '" & intDocEntry & "'"
                                    oRecordSet_U.DoQuery(strQuery)
                                End If
                            End If
                        End If
                    End If
                    oRecordSet.MoveNext()
                End While
            End If

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Public Sub CloseOrderQuantityExcludeOnOverlapping(ByVal strPRDocEntry As String)
        Try
            Dim ohtConOrder As Hashtable
            ohtConOrder = New Hashtable

            Dim strQuery As String = String.Empty
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim oRecordSet_L As SAPbobsCOM.Recordset
            oRecordSet_L = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim oRecordSet_U As SAPbobsCOM.Recordset
            oRecordSet_U = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim oRecordSet_P As SAPbobsCOM.Recordset
            oRecordSet_P = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            strQuery = "  Select Distinct T4.DocEntry "
            strQuery += " From "
            strQuery += " ( "
            strQuery += " Select DISTINCT T3.DocEntry "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " JOIN [@Z_OCPM] T4 On T4.U_CardCode = T3.BaseCard "
            strQuery += " And T4.DocEntry = T3.U_ProgramID "
            strQuery += " And Convert(VarChar(8),T3.U_DelDate,112) NOT BETWEEN Convert(VarChar(8),T4.U_PFromDate,112) AND Convert(VarChar(8),T4.U_PToDate,112)   "
            strQuery += " And T3.LineStatus = 'O' "
            strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
            strQuery += " And T3.U_DelDate > GetDate() "
            strQuery += " ) T4  "

            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                Dim oOrder As SAPbobsCOM.Documents
                While Not oRecordSet.EoF
                    oOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)

                    Dim intDocEntry As Integer = oRecordSet.Fields.Item("DocEntry").Value
                    Dim intStatus As Integer = 0
                    If oOrder.GetByKey(intDocEntry) Then

                        Dim blnUpdate As Boolean = False
                        Dim blnClose As Boolean = False
                        oRecordSet_L = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                        oRecordSet_P = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

                        strQuery = " Select T3.DocEntry,T3.VisOrder As LineNum,'' As 'Type',T3.LineNum As LineId "
                        strQuery += " From [@Z_OCPR] T0  "
                        strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
                        strQuery += " JOIN [@Z_OCPM] T4 On T4.U_CardCode = T3.BaseCard "
                        strQuery += " And T4.DocEntry = T3.U_ProgramID "
                        strQuery += " And Convert(VarChar(8),T3.U_DelDate,112) NOT BETWEEN Convert(VarChar(8),T4.U_PFromDate,112) AND Convert(VarChar(8),T4.U_PToDate,112) "
                        strQuery += " And T3.LineStatus = 'O' "
                        strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
                        strQuery += " And T3.DocEntry = '" & intDocEntry.ToString & "'"

                        oRecordSet_L.DoQuery(strQuery)
                        If Not oRecordSet_L.EoF Then

                            strQuery = " Select Distinct T3.U_DelDate,T3.U_ConDate,T3.U_IsCon "
                            strQuery += " From [@Z_OCPR] T0  "
                            strQuery += " JOIN RDR1 T3 ON T3.BaseCard = T0.U_CardCode "
                            strQuery += " And T3.LineStatus = 'O' "
                            strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
                            strQuery += " And T3.DocEntry = '" & intDocEntry.ToString & "'"

                            oRecordSet_P.DoQuery(strQuery)
                            If oRecordSet_P.RecordCount = 1 Then

                                Dim strIsCon As String = oRecordSet_P.Fields.Item("U_IsCon").Value
                                Dim strConDate As String = ""
                                If strIsCon = "Y" Then
                                    strConDate = CDate(oRecordSet_P.Fields.Item("U_ConDate").Value).ToString("yyyyMMdd")
                                End If

                                If oOrder.DocumentStatus = BoStatus.bost_Open Then
                                    If strIsCon = "Y" Then
                                        If Not ohtConOrder.ContainsKey(strConDate) Then
                                            If strConDate = CDate(oRecordSet_P.Fields.Item("U_ConDate").Value).ToString("yyyyMMdd") Then
                                                ohtConOrder.Add(strConDate, oOrder.CardCode)
                                            End If
                                        End If
                                    End If
                                    blnClose = True
                                End If

                            Else

                                While Not oRecordSet_L.EoF

                                    Dim intLine As Integer = oRecordSet_L.Fields.Item("LineNum").Value
                                    Dim strPRLine As String = oRecordSet_L.Fields.Item("LineId").Value

                                    oOrder.Lines.SetCurrentLine(intLine)

                                    Dim strIsCon As String = oOrder.Lines.UserFields.Fields.Item("U_IsCon").Value
                                    Dim strConDate As String = ""
                                    If strIsCon = "Y" Then
                                        strConDate = CDate(oOrder.Lines.UserFields.Fields.Item("U_ConDate").Value).ToString("yyyyMMdd")
                                    End If

                                    If oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Open Then
                                        If strIsCon = "Y" Then
                                            If Not ohtConOrder.ContainsKey(strConDate) Then
                                                If strConDate = CDate(oOrder.Lines.UserFields.Fields.Item("U_DelDate").Value).ToString("yyyyMMdd") Then
                                                    ohtConOrder.Add(strConDate, oOrder.CardCode)
                                                End If
                                            End If
                                        End If
                                        oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close
                                        blnUpdate = True
                                    End If
                                    oRecordSet_L.MoveNext()

                                End While

                            End If

                        End If

                        If blnUpdate Then
                            intStatus = oOrder.Update()
                        ElseIf blnClose Then
                            intStatus = oOrder.Close()
                        End If

                        'Reseting to First record to Update
                        oRecordSet_L.MoveFirst()

                        If intStatus = 0 And (blnUpdate Or blnClose) Then
                            If Not oRecordSet_L.EoF Then
                                While Not oRecordSet_L.EoF
                                    Dim intLine As Integer = oRecordSet_L.Fields.Item("LineNum").Value
                                    Dim strPRLine As String = oRecordSet_L.Fields.Item("LineId").Value

                                    If oRecordSet_L.Fields.Item("Type").Value <> "" Then
                                        Dim strType As String = oRecordSet_L.Fields.Item("Type").Value
                                        If strType <> "ED" Then
                                            strQuery = "Update RDR1 SET U_CanFrom = '" & oRecordSet_L.Fields.Item("Type").Value & "' "
                                            strQuery += " Where DocEntry = '" & intDocEntry & "'"
                                            strQuery += " And VisOrder = '" & intLine & "'"
                                            oRecordSet_U.DoQuery(strQuery)
                                        Else
                                            strQuery = "Update RDR1 SET U_CanFrom = 'E' "
                                            strQuery += " Where DocEntry = '" & intDocEntry & "'"
                                            strQuery += " And VisOrder = '" & intLine & "'"
                                            oRecordSet_U.DoQuery(strQuery)
                                        End If

                                    Else
                                        strQuery = "Update RDR1 SET FreeTxt = 'Cancelled due to Overlapping...' "
                                        strQuery += " Where DocEntry = '" & intDocEntry & "'"
                                        strQuery += " And VisOrder = '" & intLine & "'"
                                        oRecordSet_U.DoQuery(strQuery)
                                    End If

                                    If oRecordSet_L.Fields.Item("Type").Value = "E" Then
                                        strQuery = "Update [@Z_CPR4] SET U_Applied = 'Y' "
                                        strQuery += " Where DocEntry = '" & strPRDocEntry & "'"
                                        strQuery += " And LineId = '" & strPRLine & "'"
                                        oRecordSet_U.DoQuery(strQuery)
                                    ElseIf oRecordSet_L.Fields.Item("Type").Value = "R" Then
                                        strQuery = "Update [@Z_CPR8] SET U_Applied = 'Y' "
                                        strQuery += " Where DocEntry = '" & strPRDocEntry & "'"
                                        strQuery += " And LineId = '" & strPRLine & "'"
                                        oRecordSet_U.DoQuery(strQuery)
                                    ElseIf oRecordSet_L.Fields.Item("Type").Value = "S" Then
                                        strQuery = "Update [@Z_CPR9] SET U_Applied = 'Y' "
                                        strQuery += " Where DocEntry = '" & strPRDocEntry & "'"
                                        strQuery += " And LineId = '" & strPRLine & "'"
                                        oRecordSet_U.DoQuery(strQuery)
                                    End If

                                    oRecordSet_L.MoveNext()
                                End While
                            End If
                        End If
                    End If

                    oRecordSet.MoveNext()
                End While
            End If

            'Consolidate Date Update....
            If Not IsNothing(ohtConOrder) Then
                If ohtConOrder.Count > 0 Then
                    Dim Item As DictionaryEntry
                    Dim oOrder As SAPbobsCOM.Documents
                    Dim oRecordSet_M As SAPbobsCOM.Recordset
                    Dim oRecordSet_O As SAPbobsCOM.Recordset
                    oRecordSet_M = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)


                    For Each Item In ohtConOrder
                        oOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)

                        oRecordSet_M = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                        oRecordSet_L = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                        oRecordSet_O = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

                        strQuery = "Select Min(U_DelDate) From RDR1 Where LineStatus = 'O' "
                        strQuery += " And BaseCard = '" & Item.Value & "'"
                        strQuery += " And Convert(VarChar(8),U_ConDate,112) = '" & Item.Key & "'"
                        strQuery += " And U_ConDate Is Not Null "

                        oRecordSet_M.DoQuery(strQuery)
                        If Not oRecordSet_M.EoF Then

                            strQuery = "Select Distinct DocEntry From RDR1 Where LineStatus = 'O' "
                            strQuery += " And BaseCard = '" & Item.Value & "'"
                            strQuery += " And Convert(VarChar(8),U_ConDate,112) = '" & Item.Key & "'"
                            strQuery += " And U_ConDate Is Not Null "
                            oRecordSet_O.DoQuery(strQuery)
                            If Not oRecordSet_O.EoF Then

                                While Not oRecordSet_O.EoF
                                    Dim blnUpdate As Boolean = False
                                    Dim intStatus As Integer = 0

                                    If oOrder.GetByKey(CInt(oRecordSet_O.Fields.Item("DocEntry").Value)) Then
                                        strQuery = "Select VisOrder As LineNum From RDR1 Where LineStatus = 'O' "
                                        strQuery += " And BaseCard = '" & Item.Value & "'"
                                        strQuery += " And Convert(VarChar(8),U_ConDate,112) = '" & Item.Key & "'"
                                        strQuery += " And U_ConDate Is Not Null "
                                        oRecordSet_L.DoQuery(strQuery)
                                        If Not oRecordSet_L.EoF Then
                                            While Not oRecordSet_L.EoF
                                                Dim intLine As Integer = oRecordSet_L.Fields.Item("LineNum").Value
                                                oOrder.Lines.SetCurrentLine(intLine)
                                                If oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Open Then
                                                    oOrder.Lines.ShipDate = CDate(oRecordSet_M.Fields.Item(0).Value)
                                                    oOrder.Lines.UserFields.Fields.Item("U_ConDate").Value = CDate(oRecordSet_M.Fields.Item(0).Value)
                                                    blnUpdate = True
                                                End If
                                                oRecordSet_L.MoveNext()
                                            End While
                                        End If
                                    End If

                                    If blnUpdate Then
                                        intStatus = oOrder.Update()
                                    End If

                                    oRecordSet_O.MoveNext()
                                End While

                            End If

                        End If
                    Next
                End If
            End If


        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Public Sub CancelDeliveryQuantityExcludeOnOverlapping(ByVal strPRDocEntry As String)
        Try
            Dim strQuery As String = String.Empty
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            Dim oRecordSet_U As SAPbobsCOM.Recordset
            oRecordSet_U = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            strQuery = " Select T3.DocEntry,'O' As 'Type',T3.LineNum As 'LineId' "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN DLN1 T3 ON T3.BaseCard = T0.U_CardCode "
            strQuery += " JOIN ODLN T4 On T3.DocEntry = T4.DocEntry "
            strQuery += " JOIN [@Z_OCPM] T5 On T5.U_CardCode = T3.BaseCard "
            strQuery += " And T5.DocEntry = T3.U_ProgramID "
            strQuery += " And Convert(VarChar(8),T3.U_DelDate,112) NOT BETWEEN Convert(VarChar(8),T5.U_PFromDate,112) AND Convert(VarChar(8),T5.U_PToDate,112)   "
            strQuery += " And T3.LineStatus = 'O' AND  (ISNULL(T4.U_InvRef,'') = '') "
            strQuery += " Where T0.DocEntry = '" & strPRDocEntry & "'"
            strQuery += " And T3.U_DelDate > GetDate() "

            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                Dim oDelivery As SAPbobsCOM.Documents
                Dim oDelivery_C As SAPbobsCOM.Documents
                While Not oRecordSet.EoF
                    oDelivery = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                    oDelivery_C = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                    Dim intDocEntry As Integer = oRecordSet.Fields.Item("DocEntry").Value
                    Dim strPRLine As String = oRecordSet.Fields.Item("LineId").Value

                    Dim intStatus As Integer = 0
                    If oDelivery.GetByKey(intDocEntry) Then
                        Dim blnUpdate As Boolean = False
                        If oDelivery.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                            If oRecordSet.Fields.Item("Type").Value = "R" Then
                                intStatus = oDelivery.Close()
                                If intStatus = 0 Then

                                    strQuery = "Update DLN1 SET U_CanFrom = '" & oRecordSet.Fields.Item("Type").Value & "'"
                                    strQuery += " Where DocEntry = '" & intDocEntry & "'"
                                    oRecordSet_U.DoQuery(strQuery)

                                    strQuery = "Update [@Z_CPR8] SET U_Applied = 'Y' "
                                    strQuery += " Where DocEntry = '" & strPRDocEntry & "'"
                                    strQuery += " And LineId = '" & strPRLine & "'"
                                    oRecordSet_U.DoQuery(strQuery)

                                End If
                            ElseIf oRecordSet.Fields.Item("Type").Value = "S" Then
                                oDelivery_C = oDelivery.CreateCancellationDocument()
                                intStatus = oDelivery_C.Add()
                                If intStatus = 0 Then

                                    strQuery = "Update DLN1 SET U_CanFrom = '" & oRecordSet.Fields.Item("Type").Value & "'"
                                    strQuery += " Where DocEntry = '" & intDocEntry & "'"
                                    oRecordSet_U.DoQuery(strQuery)

                                    strQuery = "Update [@Z_CPR9] SET U_Applied = 'Y' "
                                    strQuery += " Where DocEntry = '" & strPRDocEntry & "'"
                                    strQuery += " And LineId = '" & strPRLine & "'"
                                    oRecordSet_U.DoQuery(strQuery)

                                End If
                            ElseIf oRecordSet.Fields.Item("Type").Value = "E" Or oRecordSet.Fields.Item("Type").Value = "ED" Then
                                oDelivery_C = oDelivery.CreateCancellationDocument()
                                intStatus = oDelivery_C.Add()
                                If intStatus = 0 Then

                                    Dim strType As String = oRecordSet.Fields.Item("Type").Value
                                    If strType = "E" Then
                                        strQuery = "Update DLN1 SET U_CanFrom = '" & oRecordSet.Fields.Item("Type").Value & "'"
                                        strQuery += " Where DocEntry = '" & intDocEntry & "'"
                                        oRecordSet_U.DoQuery(strQuery)
                                    Else
                                        strQuery = "Update DLN1 SET U_CanFrom = 'E' "
                                        strQuery += " Where DocEntry = '" & intDocEntry & "'"
                                        oRecordSet_U.DoQuery(strQuery)
                                    End If

                                    If strType = "E" Then
                                        strQuery = "Update [@Z_CPR9] SET U_Applied = 'Y' "
                                        strQuery += " Where DocEntry = '" & strPRDocEntry & "'"
                                        strQuery += " And LineId = '" & strPRLine & "'"
                                        oRecordSet_U.DoQuery(strQuery)
                                    End If

                                End If
                            Else
                                oDelivery_C = oDelivery.CreateCancellationDocument()
                                intStatus = oDelivery_C.Add()
                                If intStatus = 0 Then
                                    strQuery = "Update DLN1 SET FreeTxt = 'Cancelled Due to Overlapping...' "
                                    strQuery += " Where DocEntry = '" & intDocEntry & "'"
                                    oRecordSet_U.DoQuery(strQuery)
                                End If
                            End If
                        End If
                    End If
                    oRecordSet.MoveNext()
                End While
            End If

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Public Sub CloseDeliveryDocument(ByVal strDocEntry As String)
        Try
            Dim oDelivery As SAPbobsCOM.Documents
            oDelivery = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
            If oDelivery.GetByKey(strDocEntry) Then
                Dim blnUpdate As Boolean = False
                If oDelivery.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                    oDelivery.Close()
                End If
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Public Sub UpdateDeliveryDocument(ByVal strDocEntry As String, ByVal strInvRef As String, ByVal strInvNo As String)
        Try
            Dim oDelivery As SAPbobsCOM.Documents
            oDelivery = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
            If oDelivery.GetByKey(strDocEntry) Then
                Dim blnUpdate As Boolean = False
                If oDelivery.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                    oDelivery.UserFields.Fields.Item("U_InvRef").Value = strInvRef
                    oDelivery.UserFields.Fields.Item("U_InvNo").Value = strInvNo
                    oDelivery.Update()
                End If
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Public Function UpdateProgramDateOnOffStatus(ByVal oForm As SAPbouiCOM.Form, ByVal strDocEntry As String) As Boolean
        Dim _retVal As Boolean = False
        Dim strQuery As String = String.Empty
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim oRecordSet_U As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet_U = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            'Exclude Days
            strQuery = " Select T0.U_CardCode,T2.DocEntry,T2.LineId,T2.U_PrgDate,U_AppStatus "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_OCPM] T1 ON T1.U_CardCode = T0.U_CardCode "
            strQuery += " JOIN [@Z_CPM1] T2 ON T2.DocEntry = T1.DocEntry "
            strQuery += " And ISNULL(T1.U_Transfer,'N') = 'N' "
            'strQuery += " And T1.U_RemDays > 0 "
            strQuery += " And (  "
            strQuery += " T1.U_RemDays > 0  "
            strQuery += " OR ( Convert(VarChar(8),U_PToDate,112) >=  '" & System.DateTime.Now.AddDays(-1).ToString("yyyyMMdd") & "' ) "
            strQuery += " OR T1.U_ReRun = 'Y' "
            strQuery += " ) "
            strQuery += " Where T0.DocEntry = '" & strDocEntry & "'"
            strQuery += " And T1.U_DocStatus = 'O' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    Dim strCardCode As String = oRecordSet.Fields.Item("U_CardCode").Value
                    Dim strPDate As Date = oRecordSet.Fields.Item("U_PrgDate").Value
                    Dim strIncStatus As String = checkExcludeAlone(oRecordSet.Fields.Item("U_CardCode").Value, strPDate)
                    Dim strONOFFStatus As String = checkSuspendONOFF(oRecordSet.Fields.Item("U_CardCode").Value, strPDate)
                    Dim strPGRef As String = oRecordSet.Fields.Item("DocEntry").Value
                    Dim strLineID As String = oRecordSet.Fields.Item("LineId").Value
                    strQuery = "Update [@Z_CPM1] SET "

                    If strIncStatus = "E" Then
                        strQuery += " U_AppStatus = '" & strIncStatus & "'"
                        strQuery += " ,U_ONOFFSTA = '" & strONOFFStatus & "'"
                    ElseIf strIncStatus = "I" Then
                        strQuery += " U_AppStatus = '" & strIncStatus & "'"
                        strQuery += " ,U_ONOFFSTA = '" & strONOFFStatus & "'"
                    End If

                    strQuery += " Where DocEntry = '" & strPGRef & "'"
                    strQuery += " AND LineId = '" & strLineID & "'"
                    oRecordSet_U.DoQuery(strQuery)
                    oRecordSet.MoveNext()
                End While
            End If

            'Remove Dates(Update Program Dates in Rows....)
            strQuery = " Select T3.DocEntry,T3.LineId "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_CPR8] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN [@Z_OCPM] T2 ON T2.U_CardCode = T0.U_CardCode "
            strQuery += " JOIN [@Z_CPM1] T3 ON T3.DocEntry = T2.DocEntry "
            strQuery += " And Convert(VarChar(8),T3.U_PrgDate,112) BETWEEN Convert(VarChar(8),T1.U_FDate,112) AND Convert(VarChar(8),T1.U_TDate,112) "
            strQuery += " And ISNULL(T3.U_ONOFFSTA,'O') = 'O' AND ISNULL(T3.U_AppStatus,'I') = 'I' "
            'strQuery += " And T2.U_RemDays > 0 "
            strQuery += " And ( T2.U_RemDays > 0  "
            strQuery += " OR ( Convert(VarChar(8),T2.U_PToDate,112) >=  '" & System.DateTime.Now.AddDays(-1).ToString("yyyyMMdd") & "' ) "
            strQuery += " OR ( T2.U_ReRun = 'Y' ) "
            strQuery += " ) "
            strQuery += " And ISNULL(T2.U_Transfer,'N') = 'N' "
            strQuery += " Where T0.DocEntry = '" & strDocEntry & "'"
            strQuery += " UNION ALL  "
            strQuery += " Select T3.DocEntry,T3.LineId "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_CPR9] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN [@Z_OCPM] T2 ON T2.U_CardCode = T0.U_CardCode "
            strQuery += " JOIN [@Z_CPM1] T3 ON T3.DocEntry = T2.DocEntry "
            strQuery += " And "
            strQuery += " ( "
            strQuery += " (Convert(VarChar(8),T3.U_PrgDate,112) >= Convert(VarChar(8),T1.U_FDate,112) AND T1.U_TDate Is Null) "
            strQuery += " OR "
            strQuery += " (Convert(VarChar(8),T3.U_PrgDate,112) BETWEEN Convert(VarChar(8),T1.U_FDate,112) AND Convert(VarChar(8),T1.U_TDate,112)) "
            strQuery += " ) "
            strQuery += " And ISNULL(T3.U_ONOFFSTA,'O') = 'O' AND ISNULL(T3.U_AppStatus,'I') = 'I' "
            'strQuery += " And T2.U_RemDays > 0 "
            strQuery += " And ( T2.U_RemDays > 0  "
            strQuery += " OR ( Convert(VarChar(8),T2.U_PToDate,112) >=  '" & System.DateTime.Now.AddDays(-1).ToString("yyyyMMdd") & "' ) "
            strQuery += " OR ( T2.U_ReRun = 'Y' ) "
            strQuery += " ) "
            strQuery += " And ISNULL(T2.U_Transfer,'N') = 'N' "
            strQuery += " Where T0.DocEntry = '" & strDocEntry & "'"

            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    Dim strPGRef As String = oRecordSet.Fields.Item("DocEntry").Value
                    Dim strLineID As String = oRecordSet.Fields.Item("LineId").Value
                    strQuery = "Update [@Z_CPM1] SET "
                    strQuery += " U_ONOFFSTA = 'F' "
                    strQuery += " Where DocEntry = '" & strPGRef & "'"
                    strQuery += " AND LineId = '" & strLineID & "'"
                    oRecordSet_U.DoQuery(strQuery)
                    oRecordSet.MoveNext()
                End While
            End If

            'Update Logic ON/OFF Status in Customer Profile.
            strQuery = " Update T0 Set  "
            strQuery += " T0.U_SuFrDt = T1.U_FDate  "
            strQuery += " ,T0.U_SuToDt = T1.U_TDate  "
            strQuery += " ,T0.U_ONOFFSTA = (Case When (T1.U_TDate Is Null And T1.U_FDate Is Not Null And T1.U_FDate <= GetDate()) Then 'F' Else 'O' End)  "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_CPR9] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " Where T0.DocEntry = '" & strDocEntry & "'"
            strQuery += " And T1.LineId = "
            strQuery += " ( "
            strQuery += " Select Max(T10.LineId) From [@Z_CPR9] T10 "
            strQuery += " Where T10.DocEntry = '" & strDocEntry & "'"
            strQuery += " ) "
            oRecordSet_U.DoQuery(strQuery)

            'Update Logic ON/OFF Status in Customer Profile.
            strQuery = "Exec PROCON_UPDATEONOFFSTATUS_u"
            oRecordSet_U.DoQuery(strQuery)

            'Update Logic ON/OFF Status in Customer Program Row Based On ON/OFF Status.
            strQuery = " Select T2.DocEntry,T2.LineId,Convert(VarChar(8),T2.U_PrgDate,112) As 'ProgramDt' "
            strQuery += " From [@Z_OCPR] T0  "
            strQuery += " JOIN [@Z_OCPM] T1 ON T1.U_CardCode = T0.U_CardCode "
            strQuery += " JOIN [@Z_CPM1] T2 ON T2.DocEntry = T1.DocEntry "
            'strQuery += " JOIN [@Z_CPR8] T3 ON T3.DocEntry = T0.DocEntry "
            'strQuery += " And "
            'strQuery += " ( "
            'strQuery += " (Convert(VarChar(8),T2.U_PrgDate,112) NOT Between Convert(VarChar(8),T3.U_FDate,112) AND Convert(VarChar(8),T3.U_TDate,112)) "
            'strQuery += " AND (Convert(VarChar(8),T2.U_PrgDate,112) >=  Convert(VarChar(8),T3.U_FDate,112) AND Convert(VarChar(8),T2.U_PrgDate,112) <= Convert(VarChar(8),T3.U_TDate,112)) "
            'strQuery += " ) "
            strQuery += " And ISNULL(T2.U_ONOFFSTA,'O') = 'F' And ISNULL(T2.U_AppStatus,'I') = 'I' "
            'strQuery += " And T1.U_RemDays > 0 "
            strQuery += " And ( T1.U_RemDays > 0  "
            strQuery += " OR ( Convert(VarChar(8),T1.U_PToDate,112) >=  '" & System.DateTime.Now.AddDays(-1).ToString("yyyyMMdd") & "' ) "
            strQuery += " OR ( T1.U_ReRun = 'Y' ) "
            strQuery += " ) "
            strQuery += " And ISNULL(T1.U_Transfer,'N') = 'N' "
            strQuery += " Where T0.DocEntry = '" & strDocEntry & "'"
            strQuery += " And ISNULL(T0.U_ONOFFSTA,'O') = 'O' "
            'strQuery += " UNION ALL "
            'strQuery += " Select T2.DocEntry,T2.LineId "
            'strQuery += " From [@Z_OCPR] T0  "
            'strQuery += " JOIN [@Z_OCPM] T1 ON T1.U_CardCode = T0.U_CardCode "
            'strQuery += " JOIN [@Z_CPM1] T2 ON T2.DocEntry = T1.DocEntry "
            'strQuery += " JOIN [@Z_CPR9] T3 ON T3.DocEntry = T0.DocEntry "
            'strQuery += " And "
            'strQuery += " ( "
            'strQuery += " (Convert(VarChar(8),T2.U_PrgDate,112) NOT Between Convert(VarChar(8),T3.U_FDate,112) AND Convert(VarChar(8),T3.U_TDate,112)) "
            'strQuery += " ) "
            'strQuery += " And ISNULL(T2.U_ONOFFSTA,'O') = 'F' "
            'strQuery += " And T1.U_RemDays > 0 And ISNULL(T1.U_Transfer,'N') = 'N' "
            'strQuery += " Where T0.DocEntry = '" & strDocEntry & "'"
            'strQuery += " And ISNULL(T0.U_ONOFFSTA,'O') = 'O' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    Dim oRecordsetCheck As SAPbobsCOM.Recordset
                    oRecordsetCheck = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    Dim strPrgDate As String = oRecordSet.Fields.Item("ProgramDt").Value
                    Dim strPGRef As String = oRecordSet.Fields.Item("DocEntry").Value
                    Dim strLineID As String = oRecordSet.Fields.Item("LineId").Value
                    Dim blnSetRON As Boolean = False
                    Dim blnSetSON As Boolean = False

                    strQuery = "Select T0.LineId From [@Z_CPR8] T0  "
                    strQuery += " Where T0.DocEntry = '" & strDocEntry & "'"
                    strQuery += " AND '" & strPrgDate & "' Between Convert(VarChar(8),T0.U_FDate,112) AND Convert(VarChar(8),T0.U_TDate,112) "
                    oRecordsetCheck.DoQuery(strQuery)

                    If oRecordsetCheck.EoF Then
                        blnSetRON = True
                    Else
                        blnSetRON = False
                    End If

                    strQuery = " Select T0.LineId From [@Z_CPR9] T0  "
                    strQuery += " Where T0.DocEntry = '" & strDocEntry & "'"
                    'strQuery += " AND '" & strPrgDate & "' Between Convert(VarChar(8),T0.U_FDate,112) AND Convert(VarChar(8),T0.U_TDate,112) "
                    strQuery += " AND Convert(VarChar(8),T0.U_FDate,112) <= '" & strPrgDate & "'"
                    strQuery += " AND (Convert(VarChar(8),T0.U_TDate,112) >= '" & strPrgDate & "' OR T0.U_TDate Is Null) "
                    oRecordsetCheck.DoQuery(strQuery)

                    If oRecordsetCheck.EoF Then
                        blnSetSON = True
                    Else
                        blnSetSON = False
                    End If

                    If blnSetRON And blnSetSON Then
                        strQuery = "Update [@Z_CPM1] SET "
                        strQuery += " U_ONOFFSTA = 'O' "
                        strQuery += " Where DocEntry = '" & strPGRef & "'"
                        strQuery += " AND LineId = '" & strLineID & "'"
                        oRecordSet_U.DoQuery(strQuery)
                    End If

                    oRecordSet.MoveNext()
                End While
            End If

            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function validateDate(ByVal oForm As SAPbouiCOM.Form, ByVal strFrmDate As String, Optional ByVal intADate As Double = 0)
        Dim _retVal As Boolean = True
        Try
            Dim dtCurrentDate As Date = System.DateTime.Now.Date.AddDays(intADate)
            Dim dtFromDate As Date = CDate(strFrmDate.Substring(0, 4) + "-" + strFrmDate.Substring(4, 2) + "-" + strFrmDate.Substring(6, 2))
            If dtCurrentDate > dtFromDate Then
                _retVal = False
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
        Return _retVal
    End Function

    Public Function valCustomerProgramDate(ByVal oForm As SAPbouiCOM.Form, strCustomer As String, ByVal strFrmDate As String) As Boolean
        Dim _retVal As Boolean = True
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = DirectCast(oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            Dim strQry As String = " Select T0.DocEntry FROM DLN1 T0 JOIN ODLN T1 On T0.DocEntry = T1.DocEntry "
            strQry += " Where Convert(VarChar,T0.U_DelDate,112) = '" & strFrmDate & "'"
            strQry += " AND T0.BaseCard = '" & strCustomer & "' And T0.LineStatus = 'O' And T1.U_InvRef <> '' "
            oRecordSet.DoQuery(strQry)
            If Not oRecordSet.EoF Then
                _retVal = False
            Else
                _retVal = True
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
        Return _retVal
    End Function

    Public Function valCustomerInvoiceDate(ByVal oForm As SAPbouiCOM.Form, strCustomer As String, ByVal strFrmDate As String) As Boolean
        Dim _retVal As Boolean = True
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = DirectCast(oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            Dim strQry As String = " Select DocEntry FROM INV1 "
            strQry += " Where Convert(VarChar,U_FDate,112) = '" & strFrmDate & "'"
            strQry += " AND BaseCard = '" & strCustomer & "' "
            oRecordSet.DoQuery(strQry)
            If Not oRecordSet.EoF Then
                _retVal = True
            Else
                _retVal = False
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
        Return _retVal
    End Function

    Private Function GetCurrencyRate(ByVal oForm As SAPbouiCOM.Form, strCurr As String) As Double
        Dim dblRate As Double = 1
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = DirectCast(oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            Dim strQry As String = " Select RATE FROM ORTT Where Convert(VarChar,RateDate,112) "
            strQry += " = '" & System.DateTime.Now.ToString("yyyyMMdd") & "'"
            strQry += " AND Currency='" & strCurr & "'"
            oRecordSet.DoQuery(strQry)
            If Not oRecordSet.EoF Then
                dblRate = Convert.ToDouble(oRecordSet.Fields.Item("Rate").Value.ToString())
                Return dblRate
            Else
                Return dblRate
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
        Return dblRate
    End Function

    Public Sub PrintUDO(ByVal strFormType As String, strDocNum As String)
        Try
            Dim oReportRecord As SAPbobsCOM.Recordset
            oReportRecord = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            strSQL = "SELECT MenuUID FROM OCMN WHERE Name = '" & strFormType & "' AND Type = 'C'"
            oReportRecord.DoQuery(strSQL)
            If oReportRecord.RecordCount = 0 Then
                oApplication.Utilities.Message("No Report Configured", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                oApplication.SBO_Application.ActivateMenuItem(oReportRecord.Fields.Item(0).Value)
                'System.Threading.Thread.Sleep(2000)
                Dim oRForm As SAPbouiCOM.Form
                oRForm = oApplication.SBO_Application.Forms.ActiveForm
                oRForm.Visible = False
                CType(oRForm.Items.Item("1000003").Specific, SAPbouiCOM.EditText).Value = strDocNum
                oRForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Public Function getDay(ByVal strType As String) As System.DayOfWeek
        Try
            Select Case strType
                Case "15"
                    Return System.DayOfWeek.Monday
                Case "16"
                    Return System.DayOfWeek.Tuesday
                Case "17"
                    Return System.DayOfWeek.Wednesday
                Case "18"
                    Return System.DayOfWeek.Thursday
                Case "19"
                    Return System.DayOfWeek.Friday
                Case "20"
                    Return System.DayOfWeek.Saturday
                Case "21"
                    Return System.DayOfWeek.Sunday
            End Select
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Sub SortColumn(ByVal AForm As SAPbouiCOM.Form, itemID As String, colID As String, Optional ByVal strType As String = "D")
        Try
            Dim oMatrix As SAPbouiCOM.Matrix
            oMatrix = AForm.Items.Item(itemID).Specific
            Dim oColumn As SAPbouiCOM.Column = oMatrix.Columns.Item(colID)
            If strType = "D" Then
                oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Descending)
            Else
                oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub

    Public Function closeOpenOrdersFromProgramifCancelled(ByVal strDocEntry As String) As Boolean
        Try

            Dim _retVal As Boolean = False
            Dim intDocEntry As Integer
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim oOrder As SAPbobsCOM.Documents
            Dim intStatus As Integer
            strSQL = "Select Distinct DocEntry,VisOrder From RDR1 Where U_ProgramID = '" & strDocEntry & "' And LineStatus = 'O' "
            oRecordSet.DoQuery(strSQL)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    Dim blnClose As Boolean = False
                    oOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                    intDocEntry = CInt(oRecordSet.Fields.Item(0).Value)
                    If oOrder.GetByKey(intDocEntry) Then
                        If oOrder.DocumentStatus = BoStatus.bost_Open Then
                            blnClose = True
                        End If
                        If blnClose Then
                            intStatus = oOrder.Close()
                        End If
                    End If
                    oRecordSet.MoveNext()
                End While
            End If
            _retVal = True
            Return _retVal
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function updateRegistrationRows(ByVal oForm As SAPbouiCOM.Form, ByVal strCardCode As String)
        Try
            Dim oRecordSet_H As SAPbobsCOM.Recordset
            Dim oRecordSet_D As SAPbobsCOM.Recordset
            Dim oRecordSet_U As SAPbobsCOM.Recordset
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet_H = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet_D = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet_U = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            Dim strQuery As String
            strQuery = " Select Distinct T1.DocEntry "
            strQuery += "From "
            strQuery += "( "
            strQuery += "Select T0.DocEntry "
            strQuery += "From [@Z_CPM6] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += "Where U_PaidType = 'P'"
            strQuery += "And (ISNULL(T0.U_NoofDays,0)*ISNULL(U_Price,0)) - ((ISNULL(T0.U_NoofDays,0)*ISNULL(T0.U_Price,0))*(ISNULL(T0.U_Discount,0)/100))  <> T0.U_LineTotal "
            strQuery += "And T1.U_DocStatus = 'O' "
            strQuery += "And T1.U_CardCode = '" & strCardCode & "' "
            strQuery += "Union All "
            strQuery += "Select T0.DocEntry "
            strQuery += "From [@Z_CPM7] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry  "
            strQuery += "Where ISNULL(T0.U_InvCreated,'N') = 'N' "
            strQuery += "And (ISNULL(T0.U_Quantity,0)*ISNULL(T0.U_Price,0)) -  "
            strQuery += "((ISNULL(T0.U_Quantity,0)*ISNULL(T0.U_Price,0))*(ISNULL(T0.U_Discount,0)/100))  <> T0.U_LineTotal "
            strQuery += "And T1.U_DocStatus = 'O' "
            strQuery += "And T1.U_CardCode = '" & strCardCode & "' "
            strQuery += " ) T1 "
            oRecordSet_H.DoQuery(strQuery)
            If Not oRecordSet_H.EoF Then
                While Not oRecordSet_H.EoF

                    strQuery = " Select 'P' As 'Type',T0.DocEntry,LineId,T0.U_NoofDays,T0.U_Price,T0.U_TaxCode,ISNULL(T1.U_Discount,0) As 'U_Discount', "
                    strQuery += " (ISNULL(T0.U_NoofDays,0)*ISNULL(U_Price,0))*(ISNULL(T0.U_Discount,0)/100) As 'TotalDiscount',"
                    strQuery += " (ISNULL(T0.U_NoofDays,0)*ISNULL(U_Price,0)) - ((ISNULL(T0.U_NoofDays,0)"
                    strQuery += " *ISNULL(T0.U_Price,0))*(ISNULL(T0.U_Discount,0)/100)) As 'Total'"
                    strQuery += " From [@Z_CPM6] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry "
                    strQuery += " Where U_PaidType = 'P'"
                    strQuery += " And (ISNULL(T0.U_NoofDays,0)*ISNULL(U_Price,0)) - ((ISNULL(T0.U_NoofDays,0)"
                    strQuery += " *ISNULL(T0.U_Price,0))*(ISNULL(T0.U_Discount,0)/100))  <> T0.U_LineTotal"
                    strQuery += " And T1.U_DocStatus = 'O'"
                    strQuery += " And T1.DocEntry = '" & oRecordSet_H.Fields.Item("DocEntry").Value & "' "
                    strQuery += " Union All"
                    strQuery += " Select 'S' As 'Type',T0.DocEntry,T0.LineId,T0.U_Quantity,T0.U_Price, T0.U_TaxCode, ISNULL(T1.U_Discount,0) As 'U_Discount',"
                    strQuery += " (ISNULL(T0.U_Quantity,0)*ISNULL(T0.U_Price,0))*(ISNULL(T0.U_Discount,0)/100) As 'TotalDiscount',"
                    strQuery += " (ISNULL(T0.U_Quantity,0)*ISNULL(T0.U_Price,0)) - ((ISNULL(T0.U_Quantity,0)"
                    strQuery += " * "
                    strQuery += " ISNULL(T0.U_Price,0))*(ISNULL(T0.U_Discount,0)/100)) As 'Total'"
                    strQuery += " From [@Z_CPM7] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry "
                    strQuery += " Where ISNULL(T0.U_InvCreated,'N') = 'N'"
                    strQuery += " And (ISNULL(T0.U_Quantity,0)*ISNULL(T0.U_Price,0)) - ((ISNULL(T0.U_Quantity,0)*"
                    strQuery += " ISNULL(T0.U_Price,0))*(ISNULL(T0.U_Discount,0)/100))  <> T0.U_LineTotal"
                    strQuery += " And T1.U_DocStatus = 'O'"
                    strQuery += " And T1.DocEntry = '" & oRecordSet_H.Fields.Item("DocEntry").Value & "' "
                    oRecordSet_D.DoQuery(strQuery)

                    If Not oRecordSet_D.EoF Then
                        Dim dblTaxRate, dblTaxAmount, dblLineTotal, dblDiscount As Double
                        Dim strTaxCode, strType As String

                        While Not oRecordSet_D.EoF

                            strTaxCode = oRecordSet_D.Fields.Item("U_TaxCode").Value
                            dblLineTotal = CDbl(oRecordSet_D.Fields.Item("Total").Value)
                            dblDiscount = CDbl(oRecordSet_D.Fields.Item("U_Discount").Value)
                            strType = oRecordSet_D.Fields.Item("Type").Value

                            strQuery = " Select Rate From OVTG Where Code = '" & strTaxCode & "'"
                            dblTaxRate = oApplication.Utilities.getRecordSetValue(strQuery, "Rate")
                            dblTaxAmount += (dblTaxRate / 100) * ((dblLineTotal) - (dblLineTotal * (dblDiscount / 100)))

                            If strType = "P" Then
                                strQuery = "Update [@Z_CPM6] "
                                strQuery += " Set U_LineTotal = " & dblLineTotal & ""
                                strQuery += " Where DocEntry = '" & oRecordSet_D.Fields.Item("DocEntry").Value & "'"
                                strQuery += " And LineId = '" & oRecordSet_D.Fields.Item("LineId").Value & "'"
                                oRecordSet_U.DoQuery(strQuery)
                            Else
                                strQuery = "Update [@Z_CPM7] "
                                strQuery += " Set U_LineTotal = " & dblLineTotal & ""
                                strQuery += " Where DocEntry = '" & oRecordSet_D.Fields.Item("DocEntry").Value & "'"
                                strQuery += " And LineId = '" & oRecordSet_D.Fields.Item("LineId").Value & "'"
                                oRecordSet_U.DoQuery(strQuery)
                            End If

                            oRecordSet_D.MoveNext()
                        End While

                        Dim dblProgramTotal, dblServiceTotal, dblTBD, dblDiscount1, dblDisAmt, dblDocTotal As Double

                        strQuery = "Select (ISNULL(U_Discount,0)) From [@Z_OCPM] Where DocEntry = '" & oRecordSet_H.Fields.Item("DocEntry").Value & "'"
                        oRecordSet.DoQuery(strQuery)
                        If Not oRecordSet.EoF Then
                            dblDiscount1 = CDbl(oRecordSet.Fields.Item(0).Value)
                        End If

                        strQuery = "Select Sum(ISNULL(U_LineTotal,0)) From [@Z_CPM6] Where DocEntry = '" & oRecordSet_H.Fields.Item("DocEntry").Value & "'"
                        oRecordSet.DoQuery(strQuery)
                        If Not oRecordSet.EoF Then
                            dblProgramTotal = CDbl(oRecordSet.Fields.Item(0).Value)
                        End If

                        strQuery = "Select Sum(ISNULL(U_LineTotal,0)) From [@Z_CPM7] Where DocEntry = '" & oRecordSet_H.Fields.Item("DocEntry").Value & "'"
                        oRecordSet.DoQuery(strQuery)
                        If Not oRecordSet.EoF Then
                            dblServiceTotal = CDbl(oRecordSet.Fields.Item(0).Value)
                        End If

                        dblTBD = dblProgramTotal + dblServiceTotal
                        dblDisAmt = dblTBD * (dblDiscount1 / 100)
                        dblDocTotal = (dblTBD - (dblTBD * (dblDiscount1 / 100)) + dblTaxAmount)

                        strQuery = "Update [@Z_OCPM] "
                        strQuery += " Set U_TBDisc = " & dblTBD & ""
                        strQuery += " , U_DisAmount = " & dblDisAmt & ""
                        strQuery += " , U_TaxAmount = " & dblTaxAmount & ""
                        strQuery += " , U_DocTotal = " & dblDocTotal & ""
                        strQuery += " Where DocEntry = '" & oRecordSet_H.Fields.Item("DocEntry").Value & "'"
                        oRecordSet_U.DoQuery(strQuery)

                    End If

                    oRecordSet_H.MoveNext()
                End While
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Function updateServiceRegistrationRows(ByVal oForm As SAPbouiCOM.Form, ByVal strCardCode As String)
        Try
            Dim oRecordSet_H As SAPbobsCOM.Recordset
            Dim oRecordSet_D As SAPbobsCOM.Recordset
            Dim oRecordSet_U As SAPbobsCOM.Recordset
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet_H = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet_D = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet_U = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            Dim strQuery As String
            strQuery = " Select Distinct T1.DocEntry "
            strQuery += " From "
            strQuery += " ( "
            strQuery += " Select T0.DocEntry "
            strQuery += " From [@Z_CPM7] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry  "
            strQuery += " Where ISNULL(T0.U_InvCreated,'N') = 'N' "
            strQuery += " And (ISNULL(T0.U_Quantity,0) <> ISNULL(T1.U_NoOfDays,0) "
            strQuery += " OR "
            strQuery += " Convert(VarChar(8),T1.U_PToDate,112) <> Convert(VarChar(8),T0.U_Date,112)  "
            strQuery += " ) "
            strQuery += " And T1.U_DocStatus = 'O' "
            strQuery += " And T1.U_CardCode = '" & strCardCode & "' "
            strQuery += " And T0.U_ItemCode <> '' "
            strQuery += " ) T1 "
            oRecordSet_H.DoQuery(strQuery)
            If Not oRecordSet_H.EoF Then
                While Not oRecordSet_H.EoF

                    strQuery = " Select T0.DocEntry,T0.LineId,T1.U_NoOfDays,Convert(VarChar(8),T1.U_PToDate,112) As 'U_PToDate' "
                    strQuery += " From [@Z_CPM7] T0 JOIN [@Z_OCPM] T1 On T0.DocEntry = T1.DocEntry  "
                    strQuery += " Where ISNULL(T0.U_InvCreated,'N') = 'N' "
                    strQuery += " And (ISNULL(T0.U_Quantity,0) <> ISNULL(T1.U_NoOfDays,0) "
                    strQuery += " OR "
                    strQuery += " Convert(VarChar(8),T1.U_PToDate,112) <> Convert(VarChar(8),T0.U_Date,112)  "
                    strQuery += " ) "
                    strQuery += " And T1.U_DocStatus = 'O' "
                    strQuery += " And T1.U_CardCode = '" & strCardCode & "' "
                    strQuery += " And T0.U_ItemCode <> '' "
                    strQuery += " And T0.DocEntry = '" & oRecordSet_H.Fields.Item("DocEntry").Value & "' "
                    oRecordSet_D.DoQuery(strQuery)

                    If Not oRecordSet_D.EoF Then
                        Dim dblQty As Double

                        While Not oRecordSet_D.EoF

                            dblQty = CDbl(oRecordSet_D.Fields.Item("U_NoOfDays").Value)
                            strQuery = "Update [@Z_CPM7] "
                            strQuery += " Set U_Quantity = " & dblQty & ""
                            strQuery += " , U_Date = '" & (oRecordSet_D.Fields.Item("U_PToDate").Value) & "'"
                            strQuery += " Where DocEntry = '" & oRecordSet_D.Fields.Item("DocEntry").Value & "'"
                            strQuery += " And LineId = '" & oRecordSet_D.Fields.Item("LineId").Value & "'"
                            oRecordSet_U.DoQuery(strQuery)

                            oRecordSet_D.MoveNext()
                        End While

                    End If

                    oRecordSet_H.MoveNext()
                End While
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Function

    Public Sub Trace_Error(ByVal ex As Exception)
        Try
            Dim strFile As String = "\DIET_ADDON_" + System.DateTime.Now.ToString("yyyyMMdd") + ".txt"
            Dim strPath As String = System.Windows.Forms.Application.StartupPath.ToString() + strFile
            If Not File.Exists(strPath) Then
                Dim fileStream As FileStream
                fileStream = New FileStream(strPath, FileMode.Create, FileAccess.Write)
                Dim sw As New StreamWriter(fileStream)
                sw.BaseStream.Seek(0, SeekOrigin.End)
                'sw.WriteLine(strContent)
                Dim strMessage As String = vbCrLf & "Message ---> " & ex.Message & _
                vbCrLf & "HelpLink ---> " & ex.HelpLink & _
                vbCrLf & "Source ---> " & ex.Source & _
                vbCrLf & "StackTrace ---> " & ex.StackTrace & _
                vbCrLf & "TargetSite ---> " & ex.TargetSite.ToString()
                sw.WriteLine("======")
                sw.WriteLine("Log Time : " & System.DateTime.Now.ToLongTimeString() & " Message Stack : " & strMessage)
                sw.Flush()
                sw.Close()
            Else
                Dim fileStream As FileStream
                fileStream = New FileStream(strPath, FileMode.Append, FileAccess.Write)
                Dim sw As New StreamWriter(fileStream)
                sw.BaseStream.Seek(0, SeekOrigin.End)
                'sw.WriteLine(strContent)
                Dim strMessage As String = vbCrLf & "Message ---> " & ex.Message & _
                vbCrLf & "HelpLink ---> " & ex.HelpLink & _
                vbCrLf & "Source ---> " & ex.Source & _
                vbCrLf & "StackTrace ---> " & ex.StackTrace & _
                vbCrLf & "TargetSite ---> " & ex.TargetSite.ToString()
                sw.WriteLine("======")
                sw.WriteLine("Log Time : " & System.DateTime.Now.ToLongTimeString() & " Message Stack : " & strMessage)
                sw.Flush()
                sw.Close()
            End If
        Catch ex1 As Exception
            'Trace_Error(ex)
            'Throw ex
        End Try
    End Sub

    Public Function GetAllowOlderDates(ByVal oForm As SAPbouiCOM.Form, strCardCode As String) As Boolean
        Dim _retVal As Boolean = True
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = DirectCast(oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            Dim strQry As String = " Select ISNULL(U_ALOD,'N') As U_ALOD FROM OCRD Where CardCode   "
            strQry += " = '" & strCardCode & "' And U_ALOD = 'Y' "
            oRecordSet.DoQuery(strQry)
            If Not oRecordSet.EoF Then
                _retVal = False
            Else
                _retVal = True
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
        End Try
        Return _retVal
    End Function

End Class
