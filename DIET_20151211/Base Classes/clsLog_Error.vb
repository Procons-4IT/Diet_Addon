Imports System.IO

Public Class clsLog_Error
    Inherits Object

    Private Const log_PROCESS_ORDERS As String = "Log_ProcessOrders.txt"
    Private Const log_INVOICING As String = "Log_Invoicing.txt"
    Private Const log_DIET As String = "Log_DIET.txt"

    Private oFSO As Scripting.FileSystemObject

    Public Enum Log As Integer
        lg_PROCESS_ORDER = 1
        lg_INVOICING
        lg_DIET = 3
    End Enum

    Public Sub New()
        MyBase.New()
        oFSO = New Scripting.FileSystemObject
    End Sub

    Public Sub WriteToLog(ByVal sText As String, ByVal Type As Log)
        Dim sLogPath As String
        Dim sLogPath_T As String
        Dim sLogFilePath As String = String.Empty
        Dim oStream As Scripting.TextStream
        Try
            sLogPath = oApplication.Utilities.getApplicationPath() & "\Log"

            sLogPath_T = oApplication.Utilities.getUserTempPath() & "\Log"

            If Not oFSO.FolderExists(sLogPath) Then
                oFSO.CreateFolder(sLogPath)
            End If

            Select Case Type
                Case Log.lg_PROCESS_ORDER
                    sLogFilePath = sLogPath & "\" & log_PROCESS_ORDERS
                Case Log.lg_INVOICING
                    sLogFilePath = sLogPath & "\" & log_INVOICING
                Case Log.lg_DIET
                    sLogFilePath = sLogPath_T & "\" & log_DIET
            End Select

            If Not oFSO.FileExists(sLogFilePath) Then
                oStream = oFSO.CreateTextFile(sLogFilePath, True)
            Else
                oStream = oFSO.OpenTextFile(sLogFilePath, Scripting.IOMode.ForAppending, True, Scripting.Tristate.TristateUseDefault)
            End If

            sText = sText & vbCrLf & vbCrLf
            oStream.Write(sText)

        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            Throw (ex)
        Finally
            oStream.Close()
            oStream = Nothing
        End Try
    End Sub

    Public Sub Trace_DIET_AddOn_Error(ByVal ex As Exception)
        Try
            Dim oUserTable As SAPbobsCOM.UserTable
            Dim sCode As String
            Try
                oUserTable = oApplication.Company.UserTables.Item("Z_OERR")
                sCode = oApplication.Utilities.getMaxCode("@Z_OERR", "Code")

                Dim strMessage As String = vbCrLf & "Message ---> " & ex.Message & _
                vbCrLf & "HelpLink ---> " & ex.HelpLink & _
                vbCrLf & "Source ---> " & ex.Source & _
                vbCrLf & "StackTrace ---> " & ex.StackTrace & _
                vbCrLf & "TargetSite ---> " & ex.TargetSite.ToString()

                If Not oUserTable.GetByKey(sCode) Then

                    oUserTable.Code = sCode
                    oUserTable.Name = sCode
                    With oUserTable.UserFields.Fields
                        .Item("U_DATE").Value = System.DateTime.Now
                        .Item("U_ERROR").Value = strMessage
                        .Item("U_USER").Value = oApplication.Company.UserName
                    End With
                    If oUserTable.Add <> 0 Then
                        Throw New Exception(oApplication.Company.GetLastErrorDescription)
                    End If

                End If
            Catch ex1 As Exception
            Finally
                oUserTable = Nothing
            End Try
            Dim strFile As String = "\DIET_" + System.DateTime.Now.ToString("yyyyMMdd") + ".txt"
            Dim strPath As String = oApplication.Utilities.getUserTempPath() + strFile
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
        End Try
    End Sub

    Public Sub DeleteFile(ByVal Type As Log)
        Dim sLogFilePath As String

        Select Case Type
            Case Log.lg_PROCESS_ORDER
                sLogFilePath = oApplication.Utilities.getApplicationPath() & "\Log\" & log_PROCESS_ORDERS

            Case Log.lg_INVOICING
                sLogFilePath = oApplication.Utilities.getApplicationPath() & "\Log\" & log_INVOICING

        End Select
        If oFSO.FileExists(sLogFilePath) Then
            oFSO.DeleteFile(sLogFilePath)
        End If
    End Sub

    Public Sub ShowLogFile(ByVal Type As Log)
        Dim sLogFilePath As String

        Select Case Type
            Case Log.lg_PROCESS_ORDER
                sLogFilePath = oApplication.Utilities.getApplicationPath() & "\Log\" & log_PROCESS_ORDERS

            Case Log.lg_INVOICING
                sLogFilePath = oApplication.Utilities.getApplicationPath() & "\Log\" & log_INVOICING

        End Select

        Shell("Notepad.exe " & sLogFilePath, AppWinStyle.NormalFocus)

    End Sub

    Protected Overrides Sub Finalize()
        oFSO = Nothing
    End Sub

End Class
