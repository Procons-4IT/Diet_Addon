Imports System.IO

Public Class clsStart

    Shared Sub Main()
        Try
            Dim oMenuItem As SAPbouiCOM.MenuItem
            Dim strQuery As String = String.Empty

            Try

                oApplication = New clsListener
                oApplication.Utilities.Connect()
                oApplication.SetFilter()
                With oApplication.Company.GetCompanyService
                    CompanyDecimalSeprator = .GetAdminInfo.DecimalSeparator
                    CompanyThousandSeprator = .GetAdminInfo.ThousandsSeparator
                End With
            Catch ex As Exception
                oApplication.Log.Trace_DIET_AddOn_Error(ex)
                MessageBox.Show(ex.Message, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                Exit Sub
            End Try
            oApplication.Utilities.CreateTables()
            oApplication.Utilities.createDTMainAuthorization()
            oApplication.Utilities.AuthorizationCreation()
            oApplication.Utilities.AddRemoveMenus("Menu.xml")

            Dim strPath As String = System.Windows.Forms.Application.StartupPath & "\Script\Script.txt"
            strQuery = File.ReadAllText(strPath)
            Dim oRec_ExeSP As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec_ExeSP.DoQuery(strQuery)

            oMenuItem = oApplication.SBO_Application.Menus.Item("mnu_Z_DT")
            oMenuItem.Image = Application.StartupPath & "\DT.PNG"
            oApplication.Utilities.Message("DIET Addon Connected successfully...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oApplication.Utilities.NotifyAlert()
            System.Windows.Forms.Application.Run()
        Catch ex As Exception
            oApplication.Log.Trace_DIET_AddOn_Error(ex)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

End Class
