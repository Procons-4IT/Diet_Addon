Public NotInheritable Class clsTable

#Region "Private Functions"
    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Tables in DB. This function shall be called by 
    '                     public functions to create a table
    '**************************************************************************************************************
    Private Sub AddTables(ByVal strTab As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoUTBTableType)
        Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
        Try

            oUserTablesMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            'Adding Table
            If Not oUserTablesMD.GetByKey(strTab) Then
                oUserTablesMD.TableName = strTab
                oUserTablesMD.TableDescription = strDesc
                oUserTablesMD.TableType = nType
                If oUserTablesMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        Finally
            oUserTablesMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddFields
    'Parameter          : SstrTab As String,strCol As String,
    '                     strDesc As String,nType As Integer,i,nEditSize,nSubType As Integer
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Fields in DB Tables. This function shall be called by 
    '                     public functions to create a Field
    '**************************************************************************************************************
    Private Sub AddFields(ByVal strTab As String, _
                            ByVal strCol As String, _
                                ByVal strDesc As String, _
                                    ByVal nType As SAPbobsCOM.BoFieldTypes, _
                                        Optional ByVal i As Integer = 0, _
                                            Optional ByVal nEditSize As Integer = 10, _
                                                Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, _
                                                    Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO)
        Dim oUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            If Not (strTab = "ODLN" Or strTab = "RDR1" Or strTab = "OCRD" Or strTab = "OPR1" Or strTab = "INV1" Or strTab = "OCRD" Or strTab = "OOPR" Or strTab = "OWHS" Or strTab = "OITM" Or strTab = "INV1" Or strTab = "OWTR" Or strTab = "OUSR" Or strTab = "OITW" Or strTab = "RDR1" Or strTab = "OINV" Or strTab = "INV1" Or strTab = "OWOR" Or strTab = "ORDR" Or strTab = "OCLG" Or strTab = "ORCT") Then
                strTab = "@" + strTab
            End If

            If Not IsColumnExists(strTab, strCol) Then
                oUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

                oUserFieldMD.Description = strDesc
                oUserFieldMD.Name = strCol
                oUserFieldMD.Type = nType
                oUserFieldMD.SubType = nSubType
                oUserFieldMD.TableName = strTab
                oUserFieldMD.EditSize = nEditSize
                oUserFieldMD.Mandatory = Mandatory
                If oUserFieldMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD)

            End If

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        Finally
            oUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddFields
    'Parameter          : SstrTab As String,strCol As String,
    '                     strDesc As String,nType As Integer,i,nEditSize,nSubType As Integer
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Fields in DB Tables. This function shall be called by 
    '                     public functions to create a Field
    '**************************************************************************************************************
    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If

            If Not (TableName = "ORDR" Or TableName = "OPR1" Or TableName = "INV1" Or TableName = "OITM" Or TableName = "RDR1" Or TableName = "OITB" Or TableName = "OCRD") Then
                TableName = "@" + TableName
            End If

            If (Not IsColumnExists(TableName, ColumnName)) Then
                objUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType
                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                For intLoop = 0 To strValue.GetLength(0) - 1
                    objUserFieldMD.ValidValues.Value = strValue(intLoop)
                    objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                    objUserFieldMD.ValidValues.Add()
                Next
                If objUserFieldMD.Add() <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            Else
            End If
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            MsgBox(ex.Message)

        Finally
            objUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()

        End Try


    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : IsColumnExists
    'Parameter          : ByVal Table As String, ByVal Column As String
    'Return Value       : Boolean
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Function to check if the Column already exists in Table
    '**************************************************************************************************************
    Private Function IsColumnExists(ByVal Table As String, ByVal Column As String) As Boolean
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing

        Try
            strSQL = "SELECT COUNT(*) FROM CUFD WHERE TableID = '" & Table & "' AND AliasID = '" & Column & "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strSQL)

            If oRecordSet.Fields.Item(0).Value = 0 Then
                Return False
            Else
                Return True
            End If

            ' System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            oRecordSet = Nothing
            GC.Collect()
        End Try
    End Function

    Private Sub AddKey(ByVal strTab As String, ByVal strColumn As String, ByVal strKey As String, ByVal i As Integer)
        Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

        Try
            '// The meta-data object must be initialized with a
            '// regular UserKeys object
            oUserKeysMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

            If Not oUserKeysMD.GetByKey("@" & strTab, i) Then

                '// Set the table name and the key name
                oUserKeysMD.TableName = strTab
                oUserKeysMD.KeyName = strKey

                '// Set the column's alias
                oUserKeysMD.Elements.ColumnAlias = strColumn
                oUserKeysMD.Elements.Add()
                oUserKeysMD.Elements.ColumnAlias = "RentFac"

                '// Determine whether the key is unique or not
                oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES

                '// Add the key
                If oUserKeysMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

            End If

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)
            oUserKeysMD = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

    End Sub

    '********************************************************************
    'Type		            :   Function    
    'Name               	:	AddUDO
    'Parameter          	:   
    'Return Value       	:	Boolean
    'Author             	:	
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To Add a UDO for Transaction Tables
    '********************************************************************
    Private Sub AddUDO(ByVal strUDO As String, ByVal strDesc As String, ByVal strTable As String, _
                                Optional ByVal sFind1 As String = "", Optional ByVal sFind2 As String = "", _
                                        Optional ByVal strChildTbl As String = "", _
                                        Optional ByVal nObjectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_Document, _
                                                                                                 Optional ByVal blnDefault As Boolean = False _
                                                                                                     , Optional ByVal strDColumns As String = "")

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Try
            oUserObjectMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(strUDO) = 0 Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES

                If blnDefault Then
                    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
                    oUserObjectMD.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO

                    Dim strColumns As String()
                    strColumns = strDColumns.Split(",")
                    For Each strCol As String In strColumns
                        Dim strColumn As String() = strCol.Split("$")
                        oUserObjectMD.FormColumns.FormColumnAlias = strColumn(0)
                        oUserObjectMD.FormColumns.FormColumnDescription = strColumn(1)
                        oUserObjectMD.FormColumns.Add()
                    Next

                End If

                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES

                If sFind1 <> "" And sFind2 <> "" Then
                    oUserObjectMD.FindColumns.ColumnAlias = sFind1
                    oUserObjectMD.FindColumns.Add()
                    oUserObjectMD.FindColumns.SetCurrentLine(1)
                    oUserObjectMD.FindColumns.ColumnAlias = sFind2
                    oUserObjectMD.FindColumns.Add()
                End If

                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.LogTableName = ""
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ExtensionName = ""

                If strChildTbl <> "" Then
                    oUserObjectMD.ChildTables.TableName = strChildTbl
                End If

                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.Code = strUDO
                oUserObjectMD.Name = strDesc
                oUserObjectMD.ObjectType = nObjectType
                oUserObjectMD.TableName = strTable

                If oUserObjectMD.Add() <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If
            'If oUserObjectMD.GetByKey(strUDO) Then
            '    If blnDefault Then
            '        oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
            '    End If
            '    oUserObjectMD.Update()
            'End If
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    Private Sub AddUDO_1(ByVal strUDO As String, ByVal strDesc As String, ByVal strTable As String, _
                                Optional ByVal sFind1 As String = "", Optional ByVal sFind2 As String = "", _
                                        Optional ByVal blnMultiChild As Boolean = False, _
                                        Optional ByVal strChildTbl As String = "", _
                                        Optional ByVal nObjectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_Document)

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Try
            oUserObjectMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(strUDO) = 0 Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES

                If sFind1 <> "" And sFind2 <> "" Then
                    oUserObjectMD.FindColumns.ColumnAlias = sFind1
                    oUserObjectMD.FindColumns.Add()
                    oUserObjectMD.FindColumns.SetCurrentLine(1)
                    oUserObjectMD.FindColumns.ColumnAlias = sFind2
                    oUserObjectMD.FindColumns.Add()
                End If

                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.LogTableName = ""
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ExtensionName = ""

                If Not blnMultiChild Then
                    If strChildTbl <> "" Then
                        oUserObjectMD.ChildTables.TableName = strChildTbl
                    End If
                Else
                    Dim strChild As String()
                    strChild = strChildTbl.Split(",")
                    For Each strTabl As String In strChild
                        oUserObjectMD.ChildTables.TableName = strTabl
                        oUserObjectMD.ChildTables.Add()
                    Next
                End If

                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.Code = strUDO
                oUserObjectMD.Name = strDesc
                oUserObjectMD.ObjectType = nObjectType
                oUserObjectMD.TableName = strTable

                If oUserObjectMD.Add() <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try

    End Sub

    Private Sub UpdateUDO_1(ByVal strUDO As String, ByVal strDesc As String, ByVal strTable As String, _
                                Optional ByVal sFind1 As String = "", Optional ByVal sFind2 As String = "", _
                                        Optional ByVal blnMultiChild As Boolean = False, _
                                        Optional ByVal strChildTbl As String = "", _
                                        Optional ByVal nObjectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_Document)

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Dim blnUpdate As Boolean = False
        Try
            oUserObjectMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(strUDO) Then

                If oUserObjectMD.Name <> strDesc Then
                    oUserObjectMD.Name = strDesc
                    blnUpdate = True
                End If

                If Not blnMultiChild Then
                    If strChildTbl <> "" Then
                        oUserObjectMD.ChildTables.TableName = strChildTbl
                    End If
                Else
                    Dim strChild As String()
                    strChild = strChildTbl.Split(",")

                    For Each strTabl As String In strChild
                        Dim blnTableExists As Boolean = False
                        For index As Integer = 0 To oUserObjectMD.ChildTables.Count - 1
                            oUserObjectMD.ChildTables.SetCurrentLine(index)
                            If oUserObjectMD.ChildTables.TableName = strTabl Then
                                blnTableExists = True
                            End If
                        Next
                        If Not blnTableExists Then
                            blnUpdate = True
                            oUserObjectMD.ChildTables.Add()
                            oUserObjectMD.ChildTables.SetCurrentLine(oUserObjectMD.ChildTables.Count - 1)
                            oUserObjectMD.ChildTables.TableName = strTabl
                        End If
                    Next
                End If

                If blnUpdate Then
                    If oUserObjectMD.Update() <> 0 Then
                        Throw New Exception(oApplication.Company.GetLastErrorDescription)
                    End If
                End If

            End If

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try

    End Sub

#End Region

#Region "Public Functions"
    '*************************************************************************************************************
    'Type               : Public Function
    'Name               : CreateTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Creating Tables by calling the AddTables & AddFields Functions
    '**************************************************************************************************************
    Public Sub CreateTables()
        Try
            oApplication.SBO_Application.StatusBar.SetText("Initializing Database...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Company.StartTransaction()

            ''Program Type.
            'AddTables("Z_OPRM", "Program Type - SetUp", SAPbobsCOM.BoUTBTableType.bott_Document)
            'AddFields("Z_OPRM", "Code", "Program Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            'AddFields("Z_OPRM", "Name", "Program Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            'addField("Z_OPRM", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "Y")

            AddFields("OITM", "PrgCode", "Program Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            addField("OITB", "Program", "Program", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            addField("OITB", "Service", "Service", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")

            'Allergic / Dislike .
            AddTables("Z_ODLK", "Allergic/Dislike - SetUp", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_ODLK", "Code", "Dislike Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_ODLK", "Name", "Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_ODLK", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)
            addField("Z_ODLK", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "Y")

            'Allergic / Dislike  - Child.
            AddTables("Z_DLK1", "Allergic/Dislike Child - SetUp", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_DLK1", "Code", "Dislike Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_DLK1", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_DLK1", "ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_DLK1", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_DLK1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)

            'Calories Plan
            AddTables("Z_OCLP", "Calories Plan - SetUp", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OCLP", "Code", "Calories Plan Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OCLP", "Name", "Calories Plan Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_OCLP", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "Y")

            'Calories Adjustment.
            AddTables("Z_OCAJ", "Calories Adjustment - SetUp", SAPbobsCOM.BoUTBTableType.bott_Document)
            addField("Z_OCAJ", "Calories", "Calories", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OCAJ", "BFactor", "Breakfast", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage, "", "", "")
            addField("Z_OCAJ", "LFactor", "Lunch", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage, "", "", "")
            addField("Z_OCAJ", "LSFactor", "L-Side", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage, "", "", "")
            addField("Z_OCAJ", "SFactor", "Snack", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage, "", "", "")
            addField("Z_OCAJ", "DFactor", "Dinner", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage, "", "", "")
            addField("Z_OCAJ", "DSFactor", "Dinner-Side", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage, "", "", "")
            AddFields("Z_OCAJ", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)
            addField("Z_OCAJ", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "Y")

            'Medical Status .
            AddTables("Z_OMST", "Medical Status - SetUp", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OMST", "Code", "Medical Status Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OMST", "Name", "Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OMST", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)
            addField("Z_OMST", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "Y")

            'Medical Status  - Child.
            AddTables("Z_MST1", "Medical Status Child - SetUp", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_MST1", "Code", "Medical Status Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_MST1", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_MST1", "ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_MST1", "IActive", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_MST1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)

            ''Diet Exclude .
            'AddTables("Z_OEXD", "Diet Exclude - SetUp", SAPbobsCOM.BoUTBTableType.bott_Document)
            'AddFields("Z_OEXD", "Year", "Medical Status Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 4)
            'AddFields("Z_OEXD", "Code", "Dislike Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            'AddFields("Z_OEXD", "Name", "Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            'AddFields("Z_OEXD", "Sunday", "Sunday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            'AddFields("Z_OEXD", "Monday", "Monday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            'AddFields("Z_OEXD", "Tuesday", "Tuesday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            'AddFields("Z_OEXD", "Wednesday", "Wednesday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            'AddFields("Z_OEXD", "Thursday", "Thursday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            'AddFields("Z_OEXD", "Friday", "Friday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            'AddFields("Z_OEXD", "Saturday", "Saturday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)

            ''Diet Exclude  - Child.
            'AddTables("Z_EXD1", "Diet Exclude Child - SetUp", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            'addField("Z_EXD1", "ExDate", "Exclude Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            'AddFields("Z_EXD1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)

            'Check Up Timing.
            AddTables("Z_OTTI", "Check Up Timing - SetUp", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OTTI", "Code", "Check Up Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OTTI", "Name", "Check Up Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_OTTI", "MinTime", "Minimum Time", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OTTI", "MaxTime", "Maximum Time", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OTTI", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "Y")

            'Menu Definition.
            AddTables("Z_OMED", "Menu Definition - SetUp", SAPbobsCOM.BoUTBTableType.bott_Document)
            addField("Z_OMED", "PrgDate", "Program Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OMED", "CatType", "Catogory Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "I,G", "Program(Item),Item Group", "I")
            AddFields("Z_OMED", "GrpCode", "Group Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OMED", "GrpName", "Group Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OMED", "PrgCode", "Program Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OMED", "PrgName", "Program Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_OMED", "MenuType", "Menu Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "R,A", "Regular,Alternative", "R")
            addField("Z_OMED", "FromDate", "From Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OMED", "ToDate", "To Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OMED", "MenuDate", "Menu Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_OMED", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_OMED", "TOD", "Tip of the day", SAPbobsCOM.BoFieldTypes.db_Memo, , 500)
            AddFields("Z_OMED", "Path", "Image Path", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)
            AddFields("Z_OMED", "File", "Image File", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)

            'Menu Definition - Child1.
            AddTables("Z_MED1", "Menu Definition(BF) - Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_MED1", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_MED1", "ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_MED1", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_MED1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            'Menu Definition - Child2.
            AddTables("Z_MED2", "Menu Definition(BF/S) - Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_MED2", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_MED2", "ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_MED2", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_MED2", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            'Menu Definition - Child3.
            AddTables("Z_MED3", "Menu Definition(L) - Child3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_MED3", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_MED3", "ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_MED3", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_MED3", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            'Menu Definition - Child4.
            AddTables("Z_MED4", "Menu Definition(L/S) - Child4", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_MED4", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_MED4", "ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_MED4", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_MED4", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            'Menu Definition - Child5.
            AddTables("Z_MED5", "Menu Definition(DIN) - Child5", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_MED5", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_MED5", "ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_MED5", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_MED5", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            'Menu Definition - Child6.
            AddTables("Z_MED6", "Menu Definition(SK/F) - Child6", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_MED6", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_MED6", "ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_MED6", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_MED6", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            'New Registration.
            AddTables("Z_OCRG", "New Registration", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OCRG", "Series", "Series", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_OCRG", "CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OCRG", "CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_OCRG", "Title", "Title", SAPbobsCOM.BoFieldTypes.db_Alpha, 4, SAPbobsCOM.BoFldSubTypes.st_None, "MR,MRS,MISS", "Mr.,Mrs.,Miss.", "")
            addField("Z_OCRG", "DOB", "Date of Birth", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_OCRG", "Mobile", "Mobile No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            addField("Z_OCRG", "Age", "Age", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_OCRG", "Occup", "Occupation", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OCRG", "TeleNo", "Telephone", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OCRG", "Street", "Street", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OCRG", "Block", "Block", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OCRG", "Address", "Address", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OCRG", "City", "City", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_OCRG", "Duration", "Duration", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_OCRG", "Dietitian1", "Dietitian1", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OCRG", "Dietitian2", "Dietitian2", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_OCRG", "VisitDate", "Visit Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_OCRG", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_OCRG", "IsAuto", "IsAuto", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "Y")

            'addField("Z_OCRG", "Gender", "Gender", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "M,F", "Male,Female", "")
            'AddFields("Z_OCRG", "ZipCode", "ZipCode", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("Z_OCRG", "State", "State", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            'AddFields("Z_OCRG", "Country", "Country", SAPbobsCOM.BoFieldTypes.db_Alpha, , 3)

            'AddFields("Z_OCRG", "PriceList", "Price List", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            'addField("Z_OCRG", "VisitType", "Visit Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "I,F,L", "In Body,First,Follow Up", "")
            'AddFields("Z_OCRG", "PrgCode", "Program Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            'addField("Z_OCRG", "PrgFrDt", "Program From Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            'addField("Z_OCRG", "PrgToDt", "Program To Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            'AddFields("Z_OCRG", "Sunday", "Sunday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            'AddFields("Z_OCRG", "Monday", "Monday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            'AddFields("Z_OCRG", "Tuesday", "Tuesday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            'AddFields("Z_OCRG", "Wednesday", "Wednesday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            'AddFields("Z_OCRG", "Thursday", "Thursday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            'AddFields("Z_OCRG", "Friday", "Friday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            'AddFields("Z_OCRG", "Saturday", "Saturday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            'AddFields("Z_OCRG", "Stage", "Stage", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            'AddFields("Z_OCRG", "SaleEmp", "Sale Employee", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            'AddFields("Z_OCRG", "ItemCode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)

            'Customer Master
            AddFields("OCRD", "RegNo", "Registration No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            addField("OCRD", "Title", "Title", SAPbobsCOM.BoFieldTypes.db_Alpha, 4, SAPbobsCOM.BoFldSubTypes.st_None, "MR,MRS,MISS", "Mr.,Mrs.,Miss.", "")
            addField("OCRD", "DOB", "Date of Birth", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("OCRD", "Occup", "Occupation", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("OCRD", "sequencetype", "Sequence Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, SAPbobsCOM.BoFldSubTypes.st_None, "P1,P2", "P1,P2", "P1")
            addField("OCRD", "ALOD", "All prof Old Dates", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")

            'Oppurtunity
            AddFields("OOPR", "RegNo", "Registration No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)

            'Customer Profile.
            AddTables("Z_OCPR", "Customer Profile", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OCPR", "CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OCPR", "CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OCPR", "CPCode", "Calories Plan", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OCPR", "Sunday", "Sunday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_OCPR", "Monday", "Monday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_OCPR", "Tuesday", "Tuesday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_OCPR", "Wednesday", "Wednesday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_OCPR", "Thursday", "Thursday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_OCPR", "Friday", "Friday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_OCPR", "Saturday", "Saturday", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_OCPR", "RegNo", "Registration No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_OCPR", "PrgCode", "Program Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            addField("Z_OCPR", "PrgFrDt", "Program From Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OCPR", "PrgToDt", "Program To Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_OCPR", "CPAdj", "Calories Adjustment", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            addField("Z_OCPR", "SuFrDt", "Suspend From Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OCPR", "SuToDt", "Suspend To Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OCPR", "ONOFFSTA", "ON/OFF STATUS", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "O,F", "ON,OFF", "O")

            'Customer Profile - Child10.
            AddFields("Z_OCPR", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, , 254)
            AddFields("Z_OCPR", "DisRemarks", "Discount Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, , 254)
            AddFields("Z_OCPR", "OffRemarks", "Off Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, , 254)
            AddFields("Z_OCPR", "ConRemarks", "Consolidated Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, , 254)

            'Customer Profile - Child1.
            AddTables("Z_CPR1", "Customer Profile - Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_CPR1", "DLikeItem", "DisLike Item", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_CPR1", "Name", "DisLike Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CPR1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            'Customer Profile - Child2.
            AddTables("Z_CPR2", "Customer Profile - Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_CPR2", "MSCode", "Medical Status Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_CPR2", "Name", "Medical Status Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CPR2", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            'Customer Profile - Child3.
            AddTables("Z_CPR3", "Customer Profile - Child3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_CPR3", "OpprId", "OpprId", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CPR3", "Line", "Line", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            addField("Z_CPR3", "VisitDate", "Visit Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_CPR3", "Duration", "Duration", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CPR3", "Dietician", "Dietician", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CPR3", "Dietitian1", "Dietitian1", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CPR3", "Dietitian2", "Dietitian2", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)

            AddFields("Z_CPR3", "Weight", "Weight", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("Z_CPR3", "Height", "Height", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("Z_CPR3", "Breast", "Breast", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("Z_CPR3", "UnderBreast", "UnderBreast", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_CPR3", "Hip", "Hip", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_CPR3", "Arm ", "Arm (cm) ", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("Z_CPR3", "Bust", "Bust (cm) ", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)

            AddFields("Z_CPR3", "Waist", "Waist (cm) ", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("Z_CPR3", "Hip", "Hip (cm) ", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("Z_CPR3", "Thigh", "Thigh (cm) ", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("Z_CPR3", "Neck", "Neck (cm) ", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("Z_CPR3", "Fat", "Fat", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_CPR3", "BMI", "BMI", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)

            AddFields("Z_CPR3", "WH", "Weight History", SAPbobsCOM.BoFieldTypes.db_Memo, , 254)
            AddFields("Z_CPR3", "24RCall", "24 Hour Recall", SAPbobsCOM.BoFieldTypes.db_Memo, , 254)
            AddFields("Z_CPR3", "BC", "BeveragesConsumed", SAPbobsCOM.BoFieldTypes.db_Memo, , 254)
            AddFields("Z_CPR3", "PAD", "Physical Activity Desc", SAPbobsCOM.BoFieldTypes.db_Memo, , 254)

            addField("Z_CPR3", "PA", "Physical Activity", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, SAPbobsCOM.BoFldSubTypes.st_None, "VA,MA,LA,SE", "Very Active,Moderately Active,Low Active,Sedentary", "")
            addField("Z_CPR3", "Smoking", "Smoking", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")

            addField("Z_CPR3", "Transfer", "From Transfer", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            AddFields("Z_CPR3", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            'Customer Profile - Child4.
            AddTables("Z_CPR4", "Customer Profile - Child4", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            addField("Z_CPR4", "ExDate", "Exclude Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_CPR4", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_CPR4", "Applied", "Is Applied", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            AddFields("Z_CPR4", "Include", "Include Food", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)

            'Customer Profile - Child5.
            AddTables("Z_CPR5", "Customer Profile - Child5", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            addField("Z_CPR5", "DelDate", "Delivery Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_CPR5", "TDelDate", "To Delivery Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_CPR5", "SaleEmp", "Sector", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_CPR5", "Address", "Address", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_CPR5", "Building", "Building", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)
            AddFields("Z_CPR5", "BF", "Break Fast", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_CPR5", "LN", "Lunch", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_CPR5", "LS", "Lunch Side", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_CPR5", "SK", "Snack", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_CPR5", "DI", "Dinner", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_CPR5", "DS", "Dinner Side", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_CPR5", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)

            'Customer Profile - Child6.
            AddTables("Z_CPR6", "Customer Profile - Child6", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            addField("Z_CPR6", "Day", "Day", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, SAPbobsCOM.BoFldSubTypes.st_None, "1,2,3,4,5,6,7", "Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday", "")
            AddFields("Z_CPR6", "SaleEmp", "Sector", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_CPR6", "Address", "Address", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_CPR6", "Building", "Building", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)
            AddFields("Z_CPR6", "BF", "Break Fast", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_CPR6", "LN", "Lunch", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_CPR6", "LS", "Lunch Side", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_CPR6", "SK", "Snack", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_CPR6", "DI", "Dinner", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_CPR6", "DS", "Dinner Side", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_CPR6", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)

            'Customer Profile - Child7.
            AddTables("Z_CPR7", "Customer Profile - Child7", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            addField("Z_CPR7", "PrgDate", "Program Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_CPR7", "CPAdj", "Calories Adjustment", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CPR7", "BF", "Break Fast", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CPR7", "LN", "Lunch", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CPR7", "LS", "Lunch Side", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CPR7", "SK", "Snack", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CPR7", "DI", "Dinner", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CPR7", "DS", "Dinner Side", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CPR7", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            'Customer Profile - Child8.
            AddTables("Z_CPR8", "Customer Profile - Child8", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            addField("Z_CPR8", "FDate", "From Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_CPR8", "TDate", "To Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_CPR8", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_CPR8", "Applied", "Is Applied", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")

            'Customer Profile - Child8.
            AddTables("Z_CPR9", "Customer Profile - Child9", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            addField("Z_CPR9", "FDate", "From Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_CPR9", "TDate", "To Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_CPR9", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_CPR9", "Applied", "Is Applied", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")

            'Customer Program.
            AddTables("Z_OCPM", "Customer Program", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OCPM", "CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OCPM", "CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OCPM", "PrgCode", "Program Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OCPM", "OrderNo", "SaleOrder", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_OCPM", "PrgName", "Program Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_OCPM", "PFromDate", "Program Frm Dt", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OCPM", "PToDate", "Program To Dt", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OCPM", "NoOfDays", "No of Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OCPM", "RemDays", "Remaining Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_OCPM", "InvRef", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            addField("Z_OCPM", "Transfer", "From Transfer", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            AddFields("Z_OCPM", "TrnRef", "Transfer Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            addField("Z_OCPM", "Cancel", "Cancel", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            AddFields("Z_OCPM", "BCardCode", "Base Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)

            addField("Z_OCPM", "FreeDays", "Free Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_OCPM", "Remarks", "Document Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)
            addField("Z_OCPM", "PaidSta", "Paid Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "O,P", "Open,Paid", "O")
            addField("Z_OCPM", "DocStatus", "Document Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "O,L,C", "Open,Canceled,Closed", "O")
            addField("Z_OCPM", "OrdDays", "Total Order Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OCPM", "DelDays", "Total Delivery Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OCPM", "InvDays", "Total Invoice Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")

            addField("Z_OCPM", "TBDisc", "Total Before Discount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            addField("Z_OCPM", "Discount", "Discount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage, "", "", "")
            addField("Z_OCPM", "DisAmount", "Document Amount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            addField("Z_OCPM", "TaxAmount", "Tax Amount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            addField("Z_OCPM", "DocTotal", "Document Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            addField("Z_OCPM", "CurSour", "Currency Source", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "L,S,C", "Local Currency,System Currency,BP Currency", "L")
            AddFields("Z_OCPM", "VenCur", "Vendor Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 6)
            AddFields("Z_OCPM", "DocCur", "Document Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 6)
            addField("Z_OCPM", "DocRate", "Document Rate", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate, "", "", "")
            AddFields("Z_OCPM", "CRemarks", "Cancellation Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)
            addField("Z_OCPM", "IsSequence", "IsSequence", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "NO,YES", "N")
            addField("Z_OCPM", "RmvDays", "Remove Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_OCPM", "ReRun", "ReRun", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)

            'Customer Program - Child1.
            AddTables("Z_CPM1", "Customer Program - Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            addField("Z_CPM1", "PrgDate", "Program Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_CPM1", "BF", "Break Fast", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_CPM1", "BFS", "Break Fast Side", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_CPM1", "Lunch", "Lunch", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_CPM1", "LunchS", "Lunch Side", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_CPM1", "Dinner", "Dinner", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_CPM1", "DinnerS", "DinnerS", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_CPM1", "Snack", "Snack", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            addField("Z_CPM1", "AppStatus", "AppStatus", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "I,E", "Include,Exclude", "I")
            addField("Z_CPM1", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "O,C,L,T", "Open,Closed,Cancelled,Transfer", "O")
            AddFields("Z_CPM1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_CPM1", "ONOFFSTA", "ON/OFF STATUS", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "O,F", "ON,OFF", "O")

            'Customer Program - Child2.
            AddTables("Z_CPM2", "Customer Program - Child3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            addField("Z_CPM2", "VisitDate", "Visit Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_CPM2", "Duration", "Duration", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CPM2", "Dietician", "Dietician", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CPM2", "Dietitian1", "Dietitian1", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CPM2", "Dietitian2", "Dietitian2", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            addField("Z_CPM2", "Weight", "Weight", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Measurement, "", "", "")
            addField("Z_CPM2", "Breast", "Breast", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Measurement, "", "", "")
            addField("Z_CPM2", "Belly", "Belly", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Measurement, "", "", "")
            addField("Z_CPM2", "Height", "Height", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Measurement, "", "", "")
            AddFields("Z_CPM2", "UnderBreast", "UnderBreast", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_CPM2", "Hip", "Hip", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_CPM2", "Fat", "Fat", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_CPM2", "BMI", "BMI", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_CPM2", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            'Customer Program - Child3.
            AddTables("Z_CPM3", "Customer Program - Child4", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            addField("Z_CPM3", "ExDate", "Exclude Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")

            'Customer Program - Child4
            AddTables("Z_CPM4", "Customer Program -Qty", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            addField("Z_CPM4", "PrgDate", "Program Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_CPM4", "BF", "Break Fast", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity, "", "", "")
            addField("Z_CPM4", "Lunch", "Lunch", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity, "", "", "")
            addField("Z_CPM4", "LunchS", "Lunch Side", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity, "", "", "")
            addField("Z_CPM4", "Dinner", "Dinner", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity, "", "", "")
            addField("Z_CPM4", "DinnerS", "Dinner Side", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity, "", "", "")
            addField("Z_CPM4", "Snack", "Snack", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity, "", "", "")

            'Customer Program - Child5
            AddTables("Z_CPM5", "Program Food -Type", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            addField("Z_CPM5", "PrgDate", "Program Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_CPM5", "BF", "Break Fast", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "R,A,C", "Regular,Alternative,Custom", "")
            addField("Z_CPM5", "Lunch", "Lunch", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "R,A,C", "Regular,Alternative,Custom", "")
            addField("Z_CPM5", "LunchS", "Lunch-Side", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "R,A,C", "Regular,Alternative,Custom", "")
            addField("Z_CPM5", "Dinner", "Dinner", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "R,A,C", "Regular,Alternative,Custom", "")
            addField("Z_CPM5", "DinnerS", "Dinner Side", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "R,A,C", "Regular,Alternative,Custom", "")
            addField("Z_CPM5", "Snack", "Snack", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "R,A,C", "Regular,Alternative,Custom", "")

            'New Table for Invoice Break Up.
            AddTables("Z_CPM6", "Invoice Break Up", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            addField("Z_CPM6", "Fdate", "Invoice From Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_CPM6", "Edate", "Invoice To Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_CPM6", "NoofDays", "No of Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_CPM6", "Price", "Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            addField("Z_CPM6", "Discount", "Discount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage, "", "", "")
            addField("Z_CPM6", "LineTotal", "Line Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            AddFields("Z_CPM6", "IsIReq", "Invoice Required", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            addField("Z_CPM6", "InvCreated", "InvCreated", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            addField("Z_CPM6", "PaidType", "Paid Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "P,F", "Paid,Free", "P")
            AddFields("Z_CPM6", "InvRef", "Invoice Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CPM6", "InvNo", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CPM6", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)
            AddFields("Z_CPM6", "SerRef", "Service Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            addField("Z_CPM6", "OrdDays", "Total Order Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_CPM6", "DelDays", "Total Delivery Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_CPM6", "InvDays", "Total Invoice Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_CPM6", "TaxCode", "Tax Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            addField("Z_CPM6", "IPrice", "Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            AddFields("Z_CPM6", "Currency", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 6)
            addField("Z_CPM6", "RmvDays", "Remove Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")

            AddTables("Z_CPM7", "Invoice Service - Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            addField("Z_CPM7", "Date", "Applied Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_CPM7", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_CPM7", "ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_CPM7", "Price", "Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            addField("Z_CPM7", "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity, "", "", "")
            addField("Z_CPM7", "Discount", "Discount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage, "", "", "")
            addField("Z_CPM7", "LineTotal", "Line Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            addField("Z_CPM7", "InvCreated", "InvCreated", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            AddFields("Z_CPM7", "InvRef", "Invoice Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CPM7", "InvNo", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CPM7", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)
            AddFields("Z_CPM7", "TaxCode", "Tax Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            addField("Z_CPM7", "IPrice", "Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            AddFields("Z_CPM7", "Currency", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 6)

            addField("INV1", "PaidType", "Paid Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "P,F", "Paid,Free", "P")
            addField("INV1", "ItemType", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "P,S", "Program,Service", "P")

            'Invoice Service Items.

            AddTables("Z_OISI", "Invoice Service Items", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OISI", "Reference", "Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_OISI", "InvRef", "Invoice Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_OISI", "InvLine", "Invoice Line", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_OISI", "CardCode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)

            AddTables("Z_ISI1", "Invoice Service - Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_ISI1", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_ISI1", "ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            addField("Z_ISI1", "Price", "Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            addField("Z_ISI1", "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity, "", "", "")
            addField("Z_ISI1", "Discount", "Discount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage, "", "", "")
            addField("Z_ISI1", "LineTotal", "Line Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            addField("Z_ISI1", "InvCreated", "InvCreated", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            AddFields("Z_ISI1", "InvRef", "Invoice Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_ISI1", "InvNo", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_ISI1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)

            'ARInvoice
            AddFields("INV1", "Program", "Program Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            addField("INV1", "Fdate", "Program From Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("INV1", "Edate", "Program To Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")

            'Oppurtunities
            AddFields("OPR1", "Duration", "Duration", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("OPR1", "Dietitian1", "Dietitian1", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("OPR1", "Dietitian2", "Dietitian2", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("OPR1", "Weight", "Weight", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("OPR1", "Breast", "Breast", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("OPR1", "Height", "Height", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("OPR1", "UnderBreast", "UnderBreast", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("OPR1", "Hip", "Hip", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("OPR1", "Arm ", "Arm (cm) ", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("OPR1", "Bust", "Bust (cm) ", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("OPR1", "Waist", "Waist (cm) ", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("OPR1", "Hip", "Hip (cm) ", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("OPR1", "Thigh", "Thigh (cm) ", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("OPR1", "Neck", "Neck (cm) ", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("OPR1", "Fat", "Fat", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("OPR1", "BMI", "BMI", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("OPR1", "ProgramType", "Program Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            addField("OPR1", "VisitType", "Visit Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "I,F,L", "In Body,First,Follow Up", "")
            AddFields("OPR1", "WH", "Weight History", SAPbobsCOM.BoFieldTypes.db_Memo, , 254)
            AddFields("OPR1", "24RCall", "24 Hour Recall", SAPbobsCOM.BoFieldTypes.db_Memo, , 254)
            AddFields("OPR1", "BC", "BeveragesConsumed", SAPbobsCOM.BoFieldTypes.db_Memo, , 254)
            AddFields("OPR1", "PAD", "Physical Activity Desc", SAPbobsCOM.BoFieldTypes.db_Memo, , 254)
            addField("OPR1", "PA", "Physical Activity", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, SAPbobsCOM.BoFldSubTypes.st_None, "VA,MA,LA,SE", "Very Active,Moderately Active,Low Active,Sedentary", "")
            addField("OPR1", "Smoking", "Smoking", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")

            'Pre Sales Order.
            AddTables("Z_OPSL", "Pre Sales Order", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OPSL", "RegNo", "Registration No.", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OPSL", "CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OPSL", "CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OPSL", "Program", "Program", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OPSL", "ProgramID", "Program", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            addField("Z_OPSL", "FromDate", "Progrm From Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OPSL", "TillDate", "Till Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_OPSL", "InvoiceRef", "Invoice No.", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_OPSL", "InvoiceNo", "Invoice Ref.", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            addField("Z_OPSL", "NoOfDays", "No of Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OPSL", "RNoOfDays", "Rem No of Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_OPSL", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OPSL", "SalesO", "Sales Order", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            addField("Z_OPSL", "Type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "I,T,P", "Invoice,Transfer,Program", "I")
            AddFields("Z_OPSL", "TranNo", "Transfer No.", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_OPSL", "IsCon", "Consolidate Delivery", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)

            'Pre Sales Order - Child2.
            AddTables("Z_PSL1", "Pre Sales Order", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            addField("Z_PSL1", "DelDate", "Delivery Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_PSL1", "PrgCode", "Program Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            addField("Z_PSL1", "FType", "Food Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, SAPbobsCOM.BoFldSubTypes.st_None, "BF,LN,LS,SK,DI,DS", "Break Fast,Lunch,Lunch Side,Snack,Dinner,Dinner Side", "")
            AddFields("Z_PSL1", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_PSL1", "ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_PSL1", "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity, "", "", "")
            addField("Z_PSL1", "UnitPrice", "UnitPrice", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            AddFields("Z_PSL1", "Dislike", "Dislike", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)
            AddFields("Z_PSL1", "Medical", "Medical", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)
            AddFields("Z_PSL1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_PSL1", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "O,C", "Open,Closed", "O")
            addField("Z_PSL1", "SFood", "Select Food Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "R,A,C", "Regular,Alternative,Custom", "")
            AddFields("Z_PSL1", "SalesO", "Sales Order", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)


            AddFields("ORDR", "PSNo", "Pre Sales No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            addField("ORDR", "IsCon", "Is Consolidate", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            addField("ORDR", "ConDate", "Consolidate Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("RDR1", "IsCon", "Is Consolidate", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            addField("RDR1", "CanFrom", "Cancel From", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,M,E,R,S", "No,Modify,Exclude,Remove,Suspend", "N")
            addField("RDR1", "IsCal", "Is Calculated", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            AddFields("ORDR", "IsWizard", "Is Order Wizard", SAPbobsCOM.BoFieldTypes.db_Alpha, 1)
            AddFields("ORDR", "IsDWizard", "Is Delivery Wizard", SAPbobsCOM.BoFieldTypes.db_Alpha, 1)
            AddFields("ORDR", "IsIWizard", "Is Invoice Wizard", SAPbobsCOM.BoFieldTypes.db_Alpha, 1)

            addField("RDR1", "ConDate", "Consolidate Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("RDR1", "FType", "Food Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, SAPbobsCOM.BoFldSubTypes.st_None, "BF,LN,LS,SK,DI,DS", "Break Fast,Lunch,Lunch Side,Snack,Dinner,Dinner Side", "")
            AddFields("RDR1", "PSORef", "Pre Sale Order Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("RDR1", "PSOLine", "Pre Sale Order Line", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("RDR1", "Dislike", "Dislike", SAPbobsCOM.BoFieldTypes.db_Memo, , 254)
            AddFields("RDR1", "Medical", "Medical", SAPbobsCOM.BoFieldTypes.db_Memo, , 254)
            addField("RDR1", "DelDate", "Delivery Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("RDR1", "Address", "Address", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("RDR1", "Building", "Building", SAPbobsCOM.BoFieldTypes.db_Memo, , 254)
            AddFields("RDR1", "State", "State", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("RDR1", "ProgramID", "ProgramID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)

            'Program Transfer.
            AddTables("Z_OPGT", "Program Transfer", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OPGT", "ProgramID", "Program Ref No.", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OPGT", "PrgCode", "Program Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OPGT", "PrgName", "Program Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OPGT", "CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OPGT", "CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OPGT", "TCardCode", "To Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OPGT", "TCardName", "To Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_OPGT", "PFromDate", "Visit Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OPGT", "PToDate", "Visit Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OPGT", "NoOfDays", "No of Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OPGT", "TrnType", "Transfer Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "C,P", "Customer,Program", "C")
            AddFields("Z_OPGT", "TProgramID", "To Program Ref No.", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OPGT", "TPrgCode", "To Program Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OPGT", "TPrgName", "To Program Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_OPGT", "TNoOfDays", "No of Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")

            'Program Transfer - Child1.
            AddTables("Z_PGT1", "Program Transfer - Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            addField("Z_PGT1", "PrgDate", "Program Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_PGT1", "PrgNo", "Program No.", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_PGT1", "PrgLine", "Program Line No.", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_PGT1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddTables("Z_OFSL", "Food Selection", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_OFSL", "ProgramID", "Program No.", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_OFSL", "CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            addField("Z_OFSL", "PrgDate", "Program Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_OFSL", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            addField("Z_OFSL", "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity, "", "", "")
            AddFields("Z_OFSL", "Dislike", "Dislike", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)
            AddFields("Z_OFSL", "Medical", "Medical", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)
            addField("Z_OFSL", "FType", "Food Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, SAPbobsCOM.BoFldSubTypes.st_None, "BF,LN,LS,SK,DI,DS", "Break Fast,Lunch,Lunch Side,Snack,Dinner,Dinner Side", "")
            addField("Z_OFSL", "Select", "Select", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            addField("Z_OFSL", "SFood", "Select Food Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "R,A,C", "Regular,Alternative,Custom", "")
            AddFields("Z_OFSL", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OFSL", "Session", "Session", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddTables("Z_OCRT", "Calories Ratio - SetUp", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OCRT", "Code", "Calories Ratio Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OCRT", "Name", "Calories Ratio Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_OCRT", "Ratio", "Ratio", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage, "", "", "")
            addField("Z_OCRT", "FType", "Food Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, SAPbobsCOM.BoFldSubTypes.st_None, "BF,LN,LS,SK,DI,DS", "Break Fast,Lunch,Lunch Side,Snack,Dinner,Dinner Side", "")
            addField("Z_OCRT", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "Y")

            'For Item Groups
            addField("OITB", "BF", "BreakFast", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            addField("OITB", "LN", "Lunch", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            addField("OITB", "LS", "LunchSide", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            addField("OITB", "SK", "Snacks", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            addField("OITB", "DN", "Dinner", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            addField("OITB", "DS", "DinnerSide", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")


            'For Item Master
            addField("OITM", "BF", "BreakFast", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            addField("OITM", "LN", "Lunch", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            addField("OITM", "LS", "LunchSide", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            addField("OITM", "SK", "Snacks", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            addField("OITM", "DN", "Dinner", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            addField("OITM", "DS", "DinnerSide", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")

            'Calories Plan
            AddTables("Z_OFCI", "Filter Cust/Item - SetUp", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OFCI", "Code", "Filter Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OFCI", "Name", "Filter Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OFCI", "Prefix", "Prefix", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            addField("Z_OFCI", "Type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "C,I", "Customer,Item", "C")
            addField("Z_OFCI", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "Y")

            addField("OITM", "ISFOOD", "IS FOOD", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            addField("OITB", "finishedfood", "finished food", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "Y")

            AddFields("ODLN", "InvRef", "Invoice Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("ODLN", "InvNo", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)

            AddTables("Z_OERR", "ERROR LOG", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            addField("Z_OERR", "DATE", "ERROR DATE", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_OERR", "ERROR", "ERROR", SAPbobsCOM.BoFieldTypes.db_Memo, , 254)
            AddFields("Z_OERR", "USER", "SAP USER", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            addField("Z_OERR", "TYPE", "TYPE", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "A,E", "Add-On,External", "A")

            '---- User Defined Object
            CreateUDO()

            'If oApplication.Company.InTransaction() Then
            '    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            'End If
            Dim oUMRecordSet As SAPbobsCOM.Recordset
            oUMRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oUMRecordSet.DoQuery("Update [@Z_OMED] SET U_CatType = 'I' Where U_CatType Is Null")

            oApplication.SBO_Application.StatusBar.SetText("Database creation completed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            'If oApplication.Company.InTransaction() Then
            '    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            'End If
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Public Sub CreateUDO()
        Try

            'AddUDO("Z_OPRM", "Program Type - setup", "Z_OPRM", "U_Code", "U_Name", "", SAPbobsCOM.BoUDOObjType.boud_Document, True, "DocEntry$DocEntry,U_Code$Program Code,U_Name $ Program Name,U_Active $ Active") ' Program Type
            AddUDO("Z_OCLP", "Calories_Plan", "Z_OCLP", "U_Code", "U_Name", "", SAPbobsCOM.BoUDOObjType.boud_Document, True, "DocEntry$DocEntry,U_Code$ Calories Plan Code,U_Name $ Calories Plan Name,U_Active $ Active") ' Calories Plan
            AddUDO("Z_OCAJ", "Calories_Adjustments", "Z_OCAJ", "U_Calories", "U_BFactor", "", SAPbobsCOM.BoUDOObjType.boud_Document, True, "DocEntry$DocEntry,U_Calories $ Calories,U_BFactor $ Breakfast,U_LFactor $ Lunch,U_LSFactor $ L-Side,U_SFactor $ Snack,U_DFactor $ Dinner,U_DSFactor $ Dinner-Side,U_Remarks $ Remarks,U_Active $  Active") ' Calories Adjustments
            AddUDO("Z_OTTI", "Diet_Checkup", "Z_OTTI", "U_Code", "U_Name", "", SAPbobsCOM.BoUDOObjType.boud_Document, True, "DocEntry$DocEntry,U_Code$ Checkup Session Code,U_Name $ Name,U_MinTime $ Minimum Time(min),U_MaxTime $ Maximum Time(min),U_Active $ Active") ' Check Up Timing
            AddUDO("Z_OCRT", "Calories_Ratio", "Z_OCRT", "U_Code", "U_Name", "", SAPbobsCOM.BoUDOObjType.boud_Document, True, "DocEntry$DocEntry,U_Code$ Calories Ratio Code,U_Name $ Calories Ratio Name,U_Ratio $ Ratio,U_FType $ FoodType,U_Active $  Active") ' Calories Adjustments
            AddUDO("Z_OFCI", "Filter_Prefix", "Z_OFCI", "U_Code", "U_Name", "", SAPbobsCOM.BoUDOObjType.boud_Document, True, "DocEntry$DocEntry,U_Code$ Filter Code,U_Name $ Filter Name,U_Prefix $ Prefix,U_Type $ Type,U_Active $ Active") ' Calories Plan

            AddUDO("Z_ODLK", "Dislike_Setup", "Z_ODLK", "U_Code", "U_Name", "Z_DLK1", SAPbobsCOM.BoUDOObjType.boud_Document) ' Dislike
            AddUDO("Z_OMST", "Medical_Status", "Z_OMST", "U_Code", "U_Name", "Z_MST1", SAPbobsCOM.BoUDOObjType.boud_Document) ' Medical Status
            'AddUDO("Z_OEXD", "Diet Exclude - setup", "Z_OEXD", "U_Year", "", "Z_EXD1", SAPbobsCOM.BoUDOObjType.boud_Document) ' Diet Exclude

            AddUDO_1("Z_OMED", "Menu_Definition", "Z_OMED", "U_PrgCode", "U_PrgName", True, "Z_MED1,Z_MED2,Z_MED3,Z_MED4,Z_MED5,Z_MED6", SAPbobsCOM.BoUDOObjType.boud_Document) ' Menu Definition
            AddUDO("Z_OCRG", "New_Registration", "Z_OCRG", "DocEntry", "U_CardCode", "", SAPbobsCOM.BoUDOObjType.boud_Document) ' New Registration
            AddUDO_1("Z_OCPR", "Customer_Profile", "Z_OCPR", "U_CardCode", "U_CardName", True, "Z_CPR1,Z_CPR2,Z_CPR3,Z_CPR4,Z_CPR5,Z_CPR6", SAPbobsCOM.BoUDOObjType.boud_Document) ' Customer Profile
            AddUDO_1("Z_OCPM", "Customer_Program", "Z_OCPM", "U_PrgCode", "U_PrgName", True, "Z_CPM1,Z_CPM2,Z_CPM3,Z_CPM4,Z_CPM5", SAPbobsCOM.BoUDOObjType.boud_Document) ' Customer Program
            AddUDO_1("Z_OISI", "Service_Invoice", "Z_OISI", "DocEntry", "U_Reference", True, "Z_ISI1", SAPbobsCOM.BoUDOObjType.boud_Document) ' Service Invoice

            AddUDO("Z_OPSL", "Pre_Sales_Order", "Z_OPSL", "U_CardCode", "U_CardName", "Z_PSL1", SAPbobsCOM.BoUDOObjType.boud_Document) ' Pre Sales Order
            AddUDO("Z_OPGT", "Program_Transfer", "Z_OPGT", "U_CardCode", "U_CardName", "Z_PGT1", SAPbobsCOM.BoUDOObjType.boud_Document) ' Program Transfer

            'Update UDO II Phase
            UpdateUDO_1("Z_OCPM", "Customer_Program", "Z_OCPM", "", "", True, "Z_CPM6", SAPbobsCOM.BoUDOObjType.boud_Document)
            UpdateUDO_1("Z_OCPR", "Customer_Profile", "Z_OCPR", "", "", True, "Z_CPR7,Z_CPR8,Z_CPR9", SAPbobsCOM.BoUDOObjType.boud_Document)
            UpdateUDO_1("Z_OCPM", "Customer_Program", "Z_OCPM", "", "", True, "Z_CPM7", SAPbobsCOM.BoUDOObjType.boud_Document)

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw ex
            'oApplication.Log.oApplication.Log.Trace_DIET_AddOn_Error(ex)
        End Try
    End Sub
#End Region

End Class
