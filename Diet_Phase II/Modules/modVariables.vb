Public Module modVariables
    Public oApplication As clsListener
    Public strSQL As String
    Public cfl_Text As String
    Public cfl_Btn As String
    Public CompanyDecimalSeprator As String
    Public CompanyThousandSeprator As String
    Public frmSourceMatrix As SAPbouiCOM.Matrix

    Public Enum ValidationResult As Integer
        CANCEL = 0
        OK = 1
    End Enum

    Public Enum DocType
        Booking = 1
    End Enum

    Public Const frm_WAREHOUSES As Integer = 62
    Public Const frm_ITEM_MASTER As Integer = 150
    Public Const frm_INVOICES As Integer = 133
    Public Const frm_GRPO As Integer = 143
    Public Const frm_ORDR As Integer = 139
    Public Const frm_GR_INVENTORY As Integer = 721
    Public Const frm_Project As Integer = 711
    'Public Const frm_ProdReceipt As Integer = 65214
    Public Const frm_Delivery As Integer = 140
    Public Const frm_SaleReturn As Integer = 180
    Public Const frm_ARCreditMemo As Integer = 179
    Public Const frm_Customer As Integer = 134
    Public Const frm_Production As Integer = 65211
    Public Const frm_SalesOpp As Integer = 320
    Public Const frm_PickList As Integer = 85
    Public Const frm_ItemGroup As Integer = 63

    Public Const mnu_FIND As String = "1281"
    Public Const mnu_ADD As String = "1282"
    Public Const mnu_Remove As String = "1283"
    Public Const mnu_CANCEL As String = "1284"
    Public Const mnu_CLOSE As String = "1286"
    Public Const mnu_NEXT As String = "1288"
    Public Const mnu_PREVIOUS As String = "1289"
    Public Const mnu_FIRST As String = "1290"
    Public Const mnu_LAST As String = "1291"
    Public Const mnu_ADD_ROW As String = "1292"
    Public Const mnu_DELETE_ROW As String = "1293"
    Public Const mnu_TAX_GROUP_SETUP As String = "8458"
    Public Const mnu_DEFINE_ALTERNATIVE_ITEMS As String = "11531"

    Public Const xml_MENU As String = "Menu.xml"
    Public Const xml_MENU_REMOVE As String = "RemoveMenus.xml"

    Public Const mnu_ViewCP As String = "mnu_ViewCP"

    'Public Const mnu_Z_OPRM As String = "mnu_Z_OPRM"
    'Public Const frm_Z_OPRM As String = "Z_OPRM"
    'Public Const xml_Z_OPRM As String = "frm_Z_OPRM.xml"

    Public Const mnu_Z_ODLK As String = "mnu_Z_ODLK"
    Public Const frm_Z_ODLK As String = "frm_Z_ODLK"
    Public Const xml_Z_ODLK As String = "frm_Z_ODLK.xml"

    Public Const mnu_Z_OCLP As String = "mnu_Z_OCLP"
    Public Const frm_Z_OCLP As String = "Z_OCLP"
    Public Const xml_Z_OCLP As String = "frm_Z_OCLP.xml"

    Public Const mnu_Z_OCAJ As String = "mnu_Z_OCAJ"
    Public Const frm_Z_OCAJ As String = "Z_OCAJ"
    Public Const xml_Z_OCAJ As String = "frm_Z_OCAJ.xml"

    Public Const mnu_Z_OMST As String = "mnu_Z_OMST"
    Public Const frm_Z_OMST As String = "frm_Z_OMST"
    Public Const xml_Z_OMST As String = "frm_Z_OMST.xml"

    'Public Const mnu_Z_OEXD As String = "mnu_Z_OEXD"
    'Public Const frm_Z_OEXD As String = "frm_Z_OEXD"
    'Public Const xml_Z_OEXD As String = "frm_Z_OEXD.xml"

    Public Const mnu_Z_OTTI As String = "mnu_Z_OTTI"
    Public Const frm_Z_OTTI As String = "Z_OTTI"
    Public Const xml_Z_OTTI As String = "frm_Z_OTTI.xml"

    Public Const mnu_Z_OMED As String = "mnu_Z_OMED"
    Public Const frm_Z_OMED As String = "frm_Z_OMED"
    Public Const xml_Z_OMED As String = "frm_Z_OMED.xml"

    Public Const mnu_Z_OCRG As String = "mnu_Z_OCRG"
    Public Const frm_Z_OCRG As String = "frm_Z_OCRG"
    Public Const xml_Z_OCRG As String = "frm_Z_OCRG.xml"

    Public Const mnu_Z_OCPR As String = "mnu_Z_OCPR"
    Public Const frm_Z_OCPR As String = "frm_Z_OCPR"
    Public Const xml_Z_OCPR As String = "frm_Z_OCPR.xml"

    Public Const mnu_Z_OCPM As String = "mnu_Z_OCPM"
    Public Const frm_Z_OCPM As String = "frm_Z_OCPM"
    Public Const xml_Z_OCPM As String = "frm_Z_OCPM.xml"

    Public Const mnu_Z_OPSL As String = "mnu_Z_OPSL"
    Public Const frm_Z_OPSL As String = "frm_Z_OPSL"
    Public Const xml_Z_OPSL As String = "frm_Z_OPSL.xml"

    'Public Const mnu_Z_OPSL_1 As String = "mnu_Z_OPSL_1"
    'Public Const frm_Z_OPSL_1 As String = "frm_Z_OPSL_1"
    'Public Const xml_Z_OPSL_1 As String = "frm_Z_OPSL_1.xml"

    Public Const mnu_Z_OPSL_2 As String = "mnu_Z_OPSL_2"
    Public Const frm_Z_OPSL_2 As String = "frm_Z_OPSL_2"
    Public Const xml_Z_OPSL_2 As String = "frm_Z_OPSL_2.xml"

    Public Const mnu_Z_OPGT As String = "mnu_Z_OPGT"
    Public Const frm_Z_OPGT As String = "frm_Z_OPGT"
    Public Const xml_Z_OPGT As String = "frm_Z_OPGT.xml"

    Public Const mnu_Z_OCSR As String = "mnu_Z_OCSR"
    Public Const frm_Z_OCSR As String = "frm_Z_OCSR"
    Public Const xml_Z_OCSR As String = "frm_Z_OCSR.xml"

    Public EntryChoice As DocType
    Public Const mnu_Z_OCPR_C As String = "Z_OCPR_C"

    Public Const mnu_GenerateSO As String = "mnu_GenerateSO"
    Public Const mnu_ViewSO As String = "mnu_ViewSO"
    Public Const mnu_CANCELCP As String = "mnu_CancelCP"
    Public Const mnu_CLOSECP As String = "mnu_CloseCP"

    Public Const frm_Z_OISI As String = "frm_Z_OISI"
    Public Const xml_Z_OISI As String = "frm_Z_OISI.xml"

    Public Const mnu_Z_OCRT As String = "mnu_Z_OCRT"
    Public Const frm_Z_OCRT As String = "Z_OCRT"
    Public Const xml_Z_OCRT As String = "frm_Z_OCRT.xml"

    Public Const mnu_Z_ODWT As String = "mnu_Z_ODWT"
    Public Const frm_Z_ODWT As String = "frm_Z_ODWT"
    Public Const xml_Z_ODWT As String = "frm_Z_ODWT.xml"

    Public Const mnu_Z_OMCT As String = "mnu_Z_OMCT"
    Public Const frm_Z_OMCT As String = "frm_Z_OMCT"
    Public Const xml_Z_OMCT As String = "frm_Z_OMCT.xml"

    Public Const frm_Load As String = "frm_Load"
    Public Const xml_Load As String = "Load.xml"

    Public Const mnu_Z_OMOT As String = "mnu_Z_OMOT"
    Public Const frm_Z_OMOT As String = "frm_Z_OMOT"
    Public Const xml_Z_OMOT As String = "frm_Z_OMOT.xml"

    Public Const mnu_Z_OFCI As String = "mnu_Z_OFCI"
    Public Const frm_Z_OFCI As String = "Z_OFCI"
    Public Const xml_Z_OFCI As String = "frm_Z_OFCI.xml"

    Public Const mnu_Z_OIVG As String = "mnu_Z_OIVG"
    Public Const frm_Z_OIVG As String = "frm_Z_OIVG"
    Public Const xml_Z_OIVG As String = "frm_Z_OIVG.xml"

End Module
