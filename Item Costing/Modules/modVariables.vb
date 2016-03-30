Public Module modVariables
    Public oApplication As clsListener
    Public strSQL As String
    Public cfl_Text As String
    Public cfl_Btn As String
    Public CompanyDecimalSeprator As String
    Public CompanyThousandSeprator As String
    Public frmSourceMatrix As SAPbouiCOM.Matrix
    Public frmFreightType As String
    Public frmFreightRef As String

    Public Enum ValidationResult As Integer
        CANCEL = 0
        OK = 1
    End Enum

    Public Const frm_WAREHOUSES As Integer = 62
    Public Const frm_ITEM_MASTER As Integer = 150
    Public Const frm_INVOICES As Integer = 133
    Public Const frm_GRPO As Integer = 143
    'Public Const frm_ORDR As Integer = 139
    Public Const frm_GR_INVENTORY As Integer = 721
    'Public Const frm_Project As Integer = 711
    Public Const frm_ProdReceipt As Integer = 65214
    Public Const frm_Delivery As Integer = 140
    Public Const frm_SaleReturn As Integer = 180
    Public Const frm_ARCreditMemo As Integer = 179
    'Public Const frm_Customer As Integer = 134
    'Public Const frm_Banking As Integer = 705
    'Public Const frm_IncomingPayment As Integer = 170
    'Public Const frm_OutPayment As Integer = 426
    'Public Const frm_Deposits As Integer = 606
    'Public Const frm_Freight As Integer = 890
    'Public Const frm_DocumentFreight As Integer = 3007
    'Public Const frm_Quotation As Integer = 149
    Public Const frm_INVOICESPAYMENT As Integer = 60090
    Public Const frm_ARReverseInvoice As Integer = 60091
    Public Const frm_GI_INVENTORY As Integer = 720
    Public Const frm_I_Transfer As Integer = 940
    Public Const frm_PurReturn As Integer = 182
    Public Const frm_APCreditMemo As Integer = 181

    Public Const mnu_FIND As String = "1281"
    Public Const mnu_ADD As String = "1282"
    Public Const mnu_CLOSE As String = "1286"
    Public Const mnu_NEXT As String = "1288"
    Public Const mnu_PREVIOUS As String = "1289"
    Public Const mnu_FIRST As String = "1290"
    Public Const mnu_LAST As String = "1291"
    Public Const mnu_ADD_ROW As String = "1292"
    Public Const mnu_DELETE_ROW As String = "1293"
    Public Const mnu_TAX_GROUP_SETUP As String = "8458"
    Public Const mnu_DEFINE_ALTERNATIVE_ITEMS As String = "11531"
    Public Const mnu_DUPLICATE As String = "1287"

    Public Const mnu_OPSP As String = "MNU_OPSP"
    'Public Const mnu_PRJL As String = "PrjList"
    'Public Const mnu_PRJL_C As String = "PrjList_C"
    'Public Const mnu_COMM_I As String = "ComChargesI"
    'Public Const mnu_COMM_O As String = "ComChargesO"
    'Public Const mnu_COMM_D As String = "ComChargesD"
    'Public Const mnu_OPRM As String = "MNU_OPRM"
    'Public Const mnu_OCPR As String = "MNU_OCPR"
    'Public Const mnu_CPRL As String = "MNU_CPRL"
    'Public Const mnu_CPRL_O As String = "CPRL_O"
    'Public Const mnu_CPRL_C As String = "CPRL_C"
    'Public Const mnu_OPRT As String = "MNU_OPRT"

    Public Const xml_MENU As String = "Menu.xml"
    Public Const xml_MENU_REMOVE As String = "RemoveMenus.xml"

    Public Const frm_OPSP As String = "frm_OPSP"
    Public Const xml_OPSP As String = "frm_OPSP.xml"

    'Public Const frm_PSPL As String = "frm_PSPL"
    'Public Const xml_PSPL As String = "frm_PSPL.xml"

    'Public Const frm_ComType As String = "frm_ComType"
    'Public Const xml_ComType As String = "frm_ComType.xml"
    'Public Const mnu_ComType As String = "MNU_OCMT"

    'Public Const frm_CommCharges As String = "frm_CommCharges"
    'Public Const xml_CommCharges As String = "frm_CommCharges.xml"

    'Public Const frm_OPRM As String = "frm_OPRM"
    'Public Const xml_OPRM As String = "frm_OPRM.xml"

    'Public Const frm_OCPR As String = "frm_OCPR"
    'Public Const xml_OCPR As String = "frm_OCPR.xml"

    'Public Const frm_CPRL As String = "frm_CPRL"
    'Public Const xml_CPRL As String = "frm_CPRL.xml"

    'Public Const frm_OPRT As String = "frm_OPRT"
    'Public Const xml_OPRT As String = "frm_OPRT.xml"

    'Public Const frm_PRT2 As String = "frm_PRT2"
    'Public Const xml_PRT2 As String = "frm_PRT2.xml"

    'Public Const frm_FRT1 As String = "frm_FRT1"
    'Public Const xml_FRT1 As String = "frm_FRT1.xml"

End Module
