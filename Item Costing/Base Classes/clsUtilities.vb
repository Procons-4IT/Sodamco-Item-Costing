Public Class clsUtilities

    Private strThousSep As String = ","
    Private strDecSep As String = "."
    Private intQtyDec As Integer = 3
    Private FormNum As Integer
    Private oRecordSet As SAPbobsCOM.Recordset
    Private oEditText As SAPbouiCOM.EditText

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
            Throw ex
        End Try
    End Sub
#End Region
  
#Region "Genral Functions"

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
            Throw ex
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
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Application path"
    Public Function getApplicationPath() As String

        Return Application.StartupPath.Trim

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
            Throw ex
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
            Throw ex
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
            Throw ex
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
            Throw ex
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
            Throw ex
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
            Throw ex
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
            Throw ex

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
            Throw ex
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
            Throw ex
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
        '    Throw ex
        'Finally
        '    oAlert = Nothing
        'End Try

    End Sub
#End Region

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
            Throw ex
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

#Region "Get MaxCode"

    Public Function getMaxCode(ByVal sTable As String, ByVal sColumn As String) As Integer
        Dim oRS As SAPbobsCOM.Recordset = Nothing
        Dim MaxCode As Integer
        Dim sCode As Integer
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(CAST(" & sColumn & " AS Numeric)) FROM [" & sTable & "]"
            strSQL = "select Top 1 convert(numeric," & sColumn & ") from [" & sTable & "] order by convert(numeric," & sColumn & ") desc"
            ExecuteSQL(oRS, strSQL)
            If Convert.ToString(oRS.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRS.Fields.Item(0).Value + 1
            Else
                MaxCode = 1
            End If
            sCode = MaxCode
            Return sCode
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function

#End Region

#Region "Get FormatCode"
    Public Function getFormatCode(ByVal sTable As String, ByVal sColumn As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(CAST(" & sColumn & " AS Numeric)) FROM [" & sTable & "]"
            strSQL = "select Top 1 convert(numeric," & sColumn & ") from [" & sTable & "] order by convert(numeric," & sColumn & ") desc"
            ExecuteSQL(oRS, strSQL)
            If Convert.ToString(oRS.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRS.Fields.Item(0).Value + 1
            Else
                MaxCode = 1
            End If
            sCode = Format(MaxCode, "00000000")
            Return sCode
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function
#End Region

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
                ElseIf position.ToUpper = "LEFT" Then
                    .Left = objOldItem.Left - objOldItem.Width - 3
                    .Top = objOldItem.Top
                ElseIf position.ToUpper = "DOWN" Then
                    If ItemUID = "edWork" Then
                        .Left = objOldItem.Left + 40
                    Else
                        .Left = objOldItem.Left
                    End If
                    .Top = objOldItem.Top + objOldItem.Height + 1
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
            Throw ex
        End Try
    End Sub
#End Region

#Region "Validate Accounts"
    Public Function validate_Accounts(ByVal oForm As SAPbouiCOM.Form, ByRef strMessage As String) As Boolean
        Try
            Dim _retVal As Boolean = True
            Dim strItem, strWhs As String
            Dim oMatrix As SAPbouiCOM.Matrix
            Dim oItem As SAPbobsCOM.Items
            Dim oItemWhs As SAPbobsCOM.ItemWarehouseInfo
            oMatrix = oForm.Items.Item("38").Specific

            For index As Integer = 1 To oMatrix.RowCount
                strItem = oMatrix.Columns.Item("1").Cells.Item(index).Specific.value
                strWhs = oMatrix.Columns.Item("24").Cells.Item(index).Specific.value
                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

                If strItem <> "" Then
                    If oItem.GetByKey(strItem) Then
                        If oItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tYES Then
                            oItemWhs = oItem.WhsInfo
                            For intWhsIndex As Integer = 0 To oItemWhs.Count - 1
                                oItemWhs.SetCurrentLine(intWhsIndex)
                                If oItemWhs.WarehouseCode = strWhs Then

                                    'WareHouse Level Accounts for JE Postings
                                    Dim strWQuery, strRMAcct, strFLAcct, strVLAcct As String
                                    Dim oWhsRecordSet As SAPbobsCOM.Recordset
                                    oWhsRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    strWQuery = "Select U_RMAcct,U_FLAcct,U_VLAcct From OWHS Where WhsCode = '" & strWhs & "'"
                                    oWhsRecordSet.DoQuery(strWQuery)
                                    If Not oWhsRecordSet.EoF Then
                                        strRMAcct = oWhsRecordSet.Fields.Item("U_RMAcct").Value
                                        strFLAcct = oWhsRecordSet.Fields.Item("U_FLAcct").Value
                                        strVLAcct = oWhsRecordSet.Fields.Item("U_VLAcct").Value
                                    End If

                                    If oItemWhs.UserFields.Fields.Item("U_RMAcct").Value = "" And strRMAcct = "" Then
                                        strMessage = oItemWhs.UserFields.Fields.Item("U_RMAcct").Description + " For Item : " + strItem + " Not Specified "
                                        _retVal = False
                                        Exit For
                                    ElseIf oItemWhs.UserFields.Fields.Item("U_FLAcct").Value = "" And strFLAcct = "" Then
                                        strMessage = oItemWhs.UserFields.Fields.Item("U_FLAcct").Description + " For Item : " + strItem + " Not Specified "
                                        _retVal = False
                                        Exit For
                                    ElseIf oItemWhs.UserFields.Fields.Item("U_VLAcct").Value = "" And strVLAcct = "" Then
                                        strMessage = oItemWhs.UserFields.Fields.Item("U_VLAcct").Description + " For Item : " + strItem + " Not Specified "
                                        _retVal = False
                                        Exit For
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If

               
                If Not _retVal Then
                    Exit For
                End If
            Next
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function validate_Accounts_IGI(ByVal oForm As SAPbouiCOM.Form, ByRef strMessage As String) As Boolean
        Try
            Dim _retVal As Boolean = True
            Dim strItem, strWhs As String
            Dim oMatrix As SAPbouiCOM.Matrix
            Dim oItem As SAPbobsCOM.Items
            Dim oItemWhs As SAPbobsCOM.ItemWarehouseInfo
            oMatrix = oForm.Items.Item("13").Specific

            For index As Integer = 1 To oMatrix.RowCount
                strItem = oMatrix.Columns.Item("1").Cells.Item(index).Specific.value
                strWhs = oMatrix.Columns.Item("15").Cells.Item(index).Specific.value
                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

                If strItem <> "" Then
                    If oItem.GetByKey(strItem) Then
                        If oItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tYES Then
                            oItemWhs = oItem.WhsInfo
                            For intWhsIndex As Integer = 0 To oItemWhs.Count - 1
                                oItemWhs.SetCurrentLine(intWhsIndex)
                                If oItemWhs.WarehouseCode = strWhs Then

                                    'WareHouse Level Accounts for JE Postings
                                    Dim strWQuery, strRMAcct, strFLAcct, strVLAcct As String
                                    Dim oWhsRecordSet As SAPbobsCOM.Recordset
                                    oWhsRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    strWQuery = "Select U_RMAcct,U_FLAcct,U_VLAcct From OWHS Where WhsCode = '" & strWhs & "'"
                                    oWhsRecordSet.DoQuery(strWQuery)
                                    If Not oWhsRecordSet.EoF Then
                                        strRMAcct = oWhsRecordSet.Fields.Item("U_RMAcct").Value
                                        strFLAcct = oWhsRecordSet.Fields.Item("U_FLAcct").Value
                                        strVLAcct = oWhsRecordSet.Fields.Item("U_VLAcct").Value
                                    End If

                                    If oItemWhs.UserFields.Fields.Item("U_RMAcct").Value = "" And strRMAcct = "" Then
                                        strMessage = oItemWhs.UserFields.Fields.Item("U_RMAcct").Description + " For Item : " + strItem + " Not Specified "
                                        _retVal = False
                                        Exit For
                                    ElseIf oItemWhs.UserFields.Fields.Item("U_FLAcct").Value = "" And strFLAcct = "" Then
                                        strMessage = oItemWhs.UserFields.Fields.Item("U_FLAcct").Description + " For Item : " + strItem + " Not Specified "
                                        _retVal = False
                                        Exit For
                                    ElseIf oItemWhs.UserFields.Fields.Item("U_VLAcct").Value = "" And strVLAcct = "" Then
                                        strMessage = oItemWhs.UserFields.Fields.Item("U_VLAcct").Description + " For Item : " + strItem + " Not Specified "
                                        _retVal = False
                                        Exit For
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If


                If Not _retVal Then
                    Exit For
                End If
            Next
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Get Acccount"
    Public Function getAccount(ByVal strFormatCode As String) As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select AcctCode From OACT Where FormatCode = '" + strFormatCode + "'")
        Return oTemp.Fields.Item(0).Value
    End Function
#End Region

#Region " Post Journal Entry "

    Public Function post_JournalEntry(ByVal oForm As SAPbouiCOM.Form, ByVal strDocType As String, ByVal strObjectKey As String)
        Dim _retVal As Boolean = True
        Dim strTable As String = String.Empty
        Dim strQuery As String
        Try
            Dim oDoc As SAPbobsCOM.Documents = Nothing
            Dim oDoc_Lines As SAPbobsCOM.Document_Lines
            Dim oItem As SAPbobsCOM.Items
            Dim oJE As SAPbobsCOM.JournalEntries
            Dim oItemWhs As SAPbobsCOM.ItemWarehouseInfo
            Dim blnStatus As Boolean = True
            Dim intMul As Integer = 1
            Dim JERemarks As String

            Select Case strDocType
                Case frm_Delivery
                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                    strTable = "DLN1"
                    JERemarks = "Production Costing - Delivery"
                Case frm_SaleReturn
                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oReturns)
                    strTable = "RDN1"
                    JERemarks = "Production Costing - Sale Return"
                Case frm_INVOICES, frm_INVOICESPAYMENT
                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                    strTable = "INV1"
                    JERemarks = "Production Costing - Sale Invoice"
                Case frm_ARCreditMemo
                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                    strTable = "RIN1"
                    JERemarks = "Production Costing - Sale Credit Memo"
                Case frm_GI_INVENTORY
                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                    strTable = "IGE1"
                    JERemarks = "Production Costing - Inventory Goods Issue"
            End Select

            If oDoc.Browser.GetByKeys(strObjectKey) Then

                oApplication.Company.StartTransaction()

                Dim intCurrentLine As Integer = 0
                Dim blnJEAdd As Boolean = False
                oDoc_Lines = oDoc.Lines

                oJE = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                oJE.ReferenceDate = oDoc.DocDate
                oJE.TaxDate = oDoc.TaxDate
                oJE.DueDate = oDoc.DocDueDate

                oJE.Memo = JERemarks
                oJE.Reference = JERemarks

                oJE.Reference2 = oDoc.DocEntry
                oJE.Reference3 = oDoc_Lines.LineNum.ToString()

                For docLineindex As Integer = 0 To oDoc_Lines.Count - 1
                    oDoc_Lines.SetCurrentLine(docLineindex)
                    Dim dblBaseQty As Double = 0

                    If intCurrentLine = 0 Then
                        oJE.Lines.SetCurrentLine(intCurrentLine)
                    Else
                        oJE.Lines.Add()
                        oJE.Lines.SetCurrentLine(intCurrentLine)
                    End If

                    If strDocType = frm_GI_INVENTORY Then
                        oJE.Lines.AccountCode = oDoc_Lines.AccountCode
                    Else
                        oJE.Lines.AccountCode = oDoc_Lines.COGSAccountCode
                    End If

                    'Skip in Invoice if its BaseType is Delivery
                    If strDocType = frm_INVOICES Or strDocType = frm_INVOICESPAYMENT Then
                        If oDoc_Lines.BaseType.ToString = "15" Then
                            Continue For
                        End If
                    ElseIf (strDocType = frm_SaleReturn Or strDocType = frm_ARCreditMemo) Then
                        intMul = -1
                        'Newly added for partial return Or A/R Invoice. On '17-09-2014
                        strQuery = "Select BaseQty From " + strTable + " Where DocEntry = '" + oDoc.DocEntry.ToString + "' And LineNum = '" + oDoc_Lines.BaseLine.ToString() + "'"
                        oRecordSet.DoQuery(strQuery)
                        If Not oRecordSet.EoF Then
                            dblBaseQty = CDbl(oRecordSet.Fields.Item(0).Value)
                        End If
                    End If

                    If (oDoc_Lines.UserFields.Fields.Item("U_JEDocEty").Value.ToString.Length = 0) Or (intMul = -1 And dblBaseQty <> oDoc_Lines.Quantity) Then
                        Dim dblCreditAmt As Double = 0
                        oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

                        If oItem.GetByKey(oDoc_Lines.ItemCode) Then
                            oItemWhs = oItem.WhsInfo
                            If oItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tYES Then
                                For intWhsIndex As Integer = 0 To oItemWhs.Count - 1
                                    oItemWhs.SetCurrentLine(intWhsIndex)
                                    If oItemWhs.WarehouseCode = oDoc_Lines.WarehouseCode Then

                                        'WareHouse Level Accounts for JE Postings
                                        Dim strWQuery, strRMAcct, strFLAcct, strVLAcct As String
                                        Dim oWhsRecordSet As SAPbobsCOM.Recordset

                                        oWhsRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        strWQuery = "Select U_RMAcct,U_FLAcct,U_VLAcct From OWHS Where WhsCode = '" & oDoc_Lines.WarehouseCode & "'"
                                        oWhsRecordSet.DoQuery(strWQuery)
                                        If Not oWhsRecordSet.EoF Then
                                            strRMAcct = oWhsRecordSet.Fields.Item("U_RMAcct").Value
                                            strFLAcct = oWhsRecordSet.Fields.Item("U_FLAcct").Value
                                            strVLAcct = oWhsRecordSet.Fields.Item("U_VLAcct").Value
                                        End If

                                        dblCreditAmt = intMul * CDbl(oItemWhs.UserFields.Fields.Item("U_AvgRMCst").Value + oItemWhs.UserFields.Fields.Item("U_AvgFLbCst").Value + oItemWhs.UserFields.Fields.Item("U_AvgVLbCst").Value) * CDbl(oDoc_Lines.Quantity)

                                        If dblCreditAmt > 0 Then
                                            blnJEAdd = True
                                            oJE.Lines.Credit = dblCreditAmt
                                            intCurrentLine += 1

                                            If CDbl(oItemWhs.UserFields.Fields.Item("U_AvgRMCst").Value) > 0 Then
                                                oJE.Lines.Add()
                                                oJE.Lines.SetCurrentLine(intCurrentLine)

                                                If oItemWhs.UserFields.Fields.Item("U_RMAcct").Value <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(oItemWhs.UserFields.Fields.Item("U_RMAcct").Value)
                                                ElseIf strRMAcct <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(strRMAcct)
                                                End If

                                                oJE.Lines.Debit = intMul * CDbl(oItemWhs.UserFields.Fields.Item("U_AvgRMCst").Value) * CDbl(oDoc_Lines.Quantity)
                                                intCurrentLine += 1
                                            End If

                                            If CDbl(oItemWhs.UserFields.Fields.Item("U_AvgFLbCst").Value) > 0 Then
                                                oJE.Lines.Add()
                                                oJE.Lines.SetCurrentLine(intCurrentLine)

                                                If oItemWhs.UserFields.Fields.Item("U_FLAcct").Value <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(oItemWhs.UserFields.Fields.Item("U_FLAcct").Value)
                                                ElseIf strFLAcct <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(strFLAcct)
                                                End If

                                                oJE.Lines.Debit = intMul * CDbl(oItemWhs.UserFields.Fields.Item("U_AvgFLbCst").Value) * CDbl(oDoc_Lines.Quantity)
                                                intCurrentLine += 1
                                            End If

                                            If CDbl(oItemWhs.UserFields.Fields.Item("U_AvgVLbCst").Value) > 0 Then
                                                oJE.Lines.Add()
                                                oJE.Lines.SetCurrentLine(intCurrentLine)

                                                If oItemWhs.UserFields.Fields.Item("U_VLAcct").Value <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(oItemWhs.UserFields.Fields.Item("U_VLAcct").Value)
                                                ElseIf strVLAcct <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(strVLAcct)
                                                End If

                                                oJE.Lines.Debit = intMul * CDbl(oItemWhs.UserFields.Fields.Item("U_AvgVLbCst").Value) * CDbl(oDoc_Lines.Quantity)
                                                intCurrentLine += 1
                                            End If

                                        ElseIf (dblCreditAmt < 0) Then
                                            blnJEAdd = True
                                            oJE.Lines.Debit = -1 * dblCreditAmt
                                            intCurrentLine += 1
                                            'MessageBox.Show(oJE.Lines.Credit)
                                            If CDbl(oItemWhs.UserFields.Fields.Item("U_AvgRMCst").Value) > 0 Then
                                                oJE.Lines.Add()
                                                oJE.Lines.SetCurrentLine(intCurrentLine)

                                                If oItemWhs.UserFields.Fields.Item("U_RMAcct").Value <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(oItemWhs.UserFields.Fields.Item("U_RMAcct").Value)
                                                ElseIf strRMAcct <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(strRMAcct)
                                                End If
                                                oJE.Lines.Credit = CDbl(oItemWhs.UserFields.Fields.Item("U_AvgRMCst").Value) * CDbl(oDoc_Lines.Quantity)
                                                'MessageBox.Show(oJE.Lines.Debit)
                                                intCurrentLine += 1
                                            End If

                                            If CDbl(oItemWhs.UserFields.Fields.Item("U_AvgFLbCst").Value) > 0 Then
                                                oJE.Lines.Add()
                                                oJE.Lines.SetCurrentLine(intCurrentLine)

                                                If oItemWhs.UserFields.Fields.Item("U_FLAcct").Value <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(oItemWhs.UserFields.Fields.Item("U_FLAcct").Value)
                                                ElseIf strFLAcct <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(strFLAcct)
                                                End If

                                                oJE.Lines.Credit = CDbl(oItemWhs.UserFields.Fields.Item("U_AvgFLbCst").Value) * CDbl(oDoc_Lines.Quantity)
                                                intCurrentLine += 1
                                            End If

                                            If CDbl(oItemWhs.UserFields.Fields.Item("U_AvgVLbCst").Value) > 0 Then
                                                oJE.Lines.Add()
                                                oJE.Lines.SetCurrentLine(intCurrentLine)

                                                If oItemWhs.UserFields.Fields.Item("U_VLAcct").Value <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(oItemWhs.UserFields.Fields.Item("U_VLAcct").Value)
                                                ElseIf strVLAcct <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(strVLAcct)
                                                End If
                                                oJE.Lines.Credit = CDbl(oItemWhs.UserFields.Fields.Item("U_AvgVLbCst").Value) * CDbl(oDoc_Lines.Quantity)
                                                intCurrentLine += 1
                                            End If
                                        End If
                                    End If
                                Next
                            ElseIf (oItem.ItemType = SAPbobsCOM.ItemTypeEnum.itLabor And oDoc_Lines.TreeType = SAPbobsCOM.BoItemTreeTypes.iIngredient) Then
                                Dim strWQuery, strRMAcct, strFLAcct, strVLAcct As String
                                Dim oWhsRecordSet As SAPbobsCOM.Recordset


                                For intWhsIndex As Integer = 0 To oItemWhs.Count - 1
                                    oItemWhs.SetCurrentLine(intWhsIndex)
                                    If oItemWhs.WarehouseCode = oDoc_Lines.WarehouseCode Then
                                        oWhsRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        strWQuery = "Select U_RMAcct,U_FLAcct,U_VLAcct From OWHS Where WhsCode = '" & oDoc_Lines.WarehouseCode & "'"
                                        oWhsRecordSet.DoQuery(strWQuery)
                                        If Not oWhsRecordSet.EoF Then
                                            strRMAcct = oWhsRecordSet.Fields.Item("U_RMAcct").Value
                                            strFLAcct = oWhsRecordSet.Fields.Item("U_FLAcct").Value
                                            strVLAcct = oWhsRecordSet.Fields.Item("U_VLAcct").Value
                                        End If

                                        dblCreditAmt = intMul * oItemWhs.StandardAveragePrice * CDbl(oDoc_Lines.Quantity)

                                        If dblCreditAmt > 0 Then
                                            blnJEAdd = True
                                            oJE.Lines.Credit = dblCreditAmt
                                            intCurrentLine += 1
                                            If oItem.UserFields.Fields.Item("U_LabType").Value = "F" Then
                                                oJE.Lines.Add()
                                                oJE.Lines.SetCurrentLine(intCurrentLine)

                                                If oItemWhs.UserFields.Fields.Item("U_FLAcct").Value <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(oItemWhs.UserFields.Fields.Item("U_FLAcct").Value)
                                                ElseIf strFLAcct <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(strFLAcct)
                                                End If
                                                oJE.Lines.Debit = intMul * oItemWhs.StandardAveragePrice * CDbl(oDoc_Lines.Quantity)
                                                intCurrentLine += 1
                                            ElseIf (oItem.UserFields.Fields.Item("U_LabType").Value = "V") Then
                                                oJE.Lines.Add()
                                                oJE.Lines.SetCurrentLine(intCurrentLine)

                                                If oItemWhs.UserFields.Fields.Item("U_FLAcct").Value <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(oItemWhs.UserFields.Fields.Item("U_VLAcct").Value)
                                                ElseIf strFLAcct <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(strVLAcct)
                                                End If
                                                oJE.Lines.Debit = intMul * oItemWhs.StandardAveragePrice * CDbl(oDoc_Lines.Quantity)
                                                intCurrentLine += 1
                                            End If
                                        ElseIf (dblCreditAmt < 0) Then
                                            blnJEAdd = True
                                            oJE.Lines.Debit = -1 * dblCreditAmt
                                            intCurrentLine += 1
                                            If oItem.UserFields.Fields.Item("U_LabType").Value = "F" Then
                                                oJE.Lines.Add()
                                                oJE.Lines.SetCurrentLine(intCurrentLine)
                                                If oItemWhs.UserFields.Fields.Item("U_FLAcct").Value <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(oItemWhs.UserFields.Fields.Item("U_FLAcct").Value)
                                                ElseIf strFLAcct <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(strFLAcct)
                                                End If
                                                oJE.Lines.Credit = oItemWhs.StandardAveragePrice * CDbl(oDoc_Lines.Quantity)
                                                intCurrentLine += 1
                                            ElseIf (oItem.UserFields.Fields.Item("U_LabType").Value = "V") Then
                                                oJE.Lines.Add()
                                                oJE.Lines.SetCurrentLine(intCurrentLine)

                                                If oItemWhs.UserFields.Fields.Item("U_FLAcct").Value <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(oItemWhs.UserFields.Fields.Item("U_VLAcct").Value)
                                                ElseIf strFLAcct <> "" Then
                                                    oJE.Lines.AccountCode = getAccount(strVLAcct)
                                                End If
                                                oJE.Lines.Credit = oItemWhs.StandardAveragePrice * CDbl(oDoc_Lines.Quantity)
                                                intCurrentLine += 1
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    ElseIf (oDoc_Lines.UserFields.Fields.Item("U_JEDocEty").Value.ToString.Length > 0) Then
                        If oDoc_Lines.BaseType.ToString() = "15" Or oDoc_Lines.BaseType.ToString() = "13" Then
                            Dim intCode1 As Integer
                            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            strQuery = "Select ISNULL(StornoToTr,'0') From OJDT Where StornoToTr = '" + oDoc_Lines.UserFields.Fields.Item("U_JEDocEty").Value.ToString + "'"
                            oRecordSet.DoQuery(strQuery)
                            If oRecordSet.EoF Then
                                oJE = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                                If oJE.GetByKey(oDoc_Lines.UserFields.Fields.Item("U_JEDocEty").Value.ToString) Then
                                    intCode1 = oJE.Cancel()
                                    If intCode1 <> 0 Then
                                        oApplication.SBO_Application.SetStatusBarMessage(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        _retVal = False
                                        blnStatus = False
                                    Else
                                        Dim intJE As Integer
                                        intJE = oApplication.Company.GetNewObjectKey()
                                        If oJE.GetByKey(intJE) Then
                                            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            strQuery = "Update " + strTable + " Set U_JEDocEty = '" + intJE.ToString + "' Where DocEntry = '" + oDoc.DocEntry.ToString + "' And LineNum = '" + oDoc_Lines.LineNum.ToString() + "'"
                                            oRecordSet.DoQuery(strQuery)
                                        End If
                                        oApplication.SBO_Application.SetStatusBarMessage("Production Costing Posted Sucessfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next

                If blnJEAdd Then
                    Dim intCode As Integer = oJE.Add()
                    If intCode <> 0 Then
                        oApplication.SBO_Application.SetStatusBarMessage(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        _retVal = False
                        blnStatus = False
                    Else
                        Dim intJE As Integer
                        intJE = oApplication.Company.GetNewObjectKey()
                        If oJE.GetByKey(intJE) Then
                            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            strQuery = "Update " + strTable + " Set U_JEDocEty = '" + intJE.ToString + "' Where DocEntry = '" + oDoc.DocEntry.ToString + "' And ISNULL(U_JEDocEty,'') = ''"
                            oRecordSet.DoQuery(strQuery)
                        End If
                        oApplication.SBO_Application.SetStatusBarMessage("Production Costing Posted Sucessfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    End If
                End If

                If blnStatus Then
                    If oApplication.Company.InTransaction Then
                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    End If
                Else
                    If oApplication.Company.InTransaction Then
                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If
                End If
            End If
        Catch ex As Exception
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            oApplication.SBO_Application.SetStatusBarMessage(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        Return _retVal
    End Function

#End Region

#Region "Update Costing "

    Public Function update_ProductionCosting(ByVal oForm As SAPbouiCOM.Form, ByVal strDocType As String, ByVal strObjectKey As String)
        Dim _retVal As Boolean = True
        Try
            Dim strQuery As String
            Dim oDoc As SAPbobsCOM.Documents = Nothing
            Dim oDoc_Lines As SAPbobsCOM.Document_Lines
            Dim oProduction As SAPbobsCOM.ProductionOrders
            Dim oProduction_Lines As SAPbobsCOM.ProductionOrders_Lines
            Dim dblRMCost, dblFixLbCost, dblVarLbCost As Double
            Dim oProdItem As SAPbobsCOM.Items
            Dim oProdWhs As SAPbobsCOM.ItemWarehouseInfo

            Dim oItem As SAPbobsCOM.Items
            Dim oItemWhs As SAPbobsCOM.ItemWarehouseInfo
            oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
            If oDoc.Browser.GetByKeys(strObjectKey) Then
                oDoc_Lines = oDoc.Lines

                For docLineindex As Integer = 0 To oDoc_Lines.Count - 1
                    oProduction = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                    oDoc_Lines.SetCurrentLine(docLineindex)
                    dblRMCost = 0
                    dblFixLbCost = 0
                    dblVarLbCost = 0

                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And oDoc_Lines.UserFields.Fields.Item("U_ITCost").Value = "Y" Then
                        Continue For
                    End If

                    If oProduction.GetByKey(oDoc_Lines.BaseEntry) Then

                        oProduction_Lines = oProduction.Lines
                        For proLineindex As Integer = 0 To oProduction_Lines.Count - 1
                            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            oProduction_Lines.SetCurrentLine(proLineindex)
                            If oItem.GetByKey(oProduction_Lines.ItemNo) Then
                                oItemWhs = oItem.WhsInfo
                                For intWhsIndex As Integer = 0 To oItemWhs.Count - 1
                                    oItemWhs.SetCurrentLine(intWhsIndex)
                                    If oItemWhs.WarehouseCode = oDoc_Lines.WarehouseCode Then
                                        If (oItem.ItemType = SAPbobsCOM.ItemTypeEnum.itItems) Then
                                            If hasBOM(oProduction_Lines.ItemNo) Then
                                                Dim dblRMCost_SB, dblFixLbCost_SB, dblVarLbCost_SB As Double

                                                'Loops to All Sub Bill of Materials to get RM,FC,VC
                                                'get_Cost(oProduction_Lines.ItemNo, oDoc_Lines.WarehouseCode, dblRMCost_SB, dblFixLbCost_SB, dblVarLbCost_SB)
                                                'else below

                                                dblRMCost_SB = CDbl(oItemWhs.UserFields.Fields.Item("U_AvgRMCst").Value)
                                                dblFixLbCost_SB = CDbl(oItemWhs.UserFields.Fields.Item("U_AvgFLbCst").Value)
                                                dblVarLbCost_SB = CDbl(oItemWhs.UserFields.Fields.Item("U_AvgVLbCst").Value)

                                                dblRMCost += CDbl(oProduction_Lines.BaseQuantity) * dblRMCost_SB
                                                dblFixLbCost += CDbl(oProduction_Lines.BaseQuantity) * dblFixLbCost_SB
                                                dblVarLbCost += CDbl(oProduction_Lines.BaseQuantity) * dblVarLbCost_SB
                                            Else
                                                dblRMCost += CDbl(oProduction_Lines.BaseQuantity) * CDbl(oItemWhs.StandardAveragePrice)
                                                'dblFixLbCost += CDbl(oProduction_Lines.BaseQuantity) * CDbl(oItemWhs.UserFields.Fields.Item("U_AvgFLbCst").Value)
                                                'dblVarLbCost += CDbl(oProduction_Lines.BaseQuantity) * CDbl(oItemWhs.UserFields.Fields.Item("U_AvgVLbCst").Value)
                                            End If
                                        ElseIf (oItem.ItemType = SAPbobsCOM.ItemTypeEnum.itLabor) Then
                                            If oItem.UserFields.Fields.Item("U_LabType").Value = "F" Then
                                                dblFixLbCost += CDbl(oProduction_Lines.BaseQuantity) * CDbl(oItemWhs.StandardAveragePrice)
                                            Else
                                                dblVarLbCost += CDbl(oProduction_Lines.BaseQuantity) * CDbl(oItemWhs.StandardAveragePrice)
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        Next

                        If oProduction.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard Then
                            'Update of Production Item Costs...For Standard Order
                            oProdItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            If oProdItem.GetByKey(oDoc_Lines.ItemCode) Then
                                oProdWhs = oProdItem.WhsInfo
                                For intWhsIndex As Integer = 0 To oProdWhs.Count - 1
                                    oProdWhs.SetCurrentLine(intWhsIndex)
                                    If oProdWhs.WarehouseCode = oDoc_Lines.WarehouseCode Then
                                        Dim dblAvgRMCost, dblAvgFLbCost, dblAgvVLbCost As Double

                                        dblAvgRMCost = CDbl(oProdWhs.UserFields.Fields.Item("U_AvgRMCst").Value)
                                        dblAvgFLbCost = CDbl(oProdWhs.UserFields.Fields.Item("U_AvgFLbCst").Value)
                                        dblAgvVLbCost = CDbl(oProdWhs.UserFields.Fields.Item("U_AvgVLbCst").Value)

                                        oProdWhs.UserFields.Fields.Item("U_AvgRMCst").Value = CDbl(((dblRMCost * oDoc_Lines.Quantity) + (dblAvgRMCost * (oProdWhs.InStock - oDoc_Lines.Quantity))) / oProdWhs.InStock)
                                        oProdWhs.UserFields.Fields.Item("U_AvgFLbCst").Value = CDbl(((dblFixLbCost * oDoc_Lines.Quantity) + (dblAvgFLbCost * (oProdWhs.InStock - oDoc_Lines.Quantity))) / oProdWhs.InStock)
                                        oProdWhs.UserFields.Fields.Item("U_AvgVLbCst").Value = CDbl(((dblVarLbCost * oDoc_Lines.Quantity) + (dblAgvVLbCost * (oProdWhs.InStock - oDoc_Lines.Quantity))) / oProdWhs.InStock)

                                        oApplication.Company.StartTransaction()
                                        Dim intCode As Integer = oProdItem.Update()
                                        If intCode = 0 Then
                                            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            strQuery = "Update IGN1 Set U_ITCost = 'Y' Where DocEntry = '" + oDoc.DocEntry.ToString + "' And LineNum = '" + oDoc_Lines.LineNum.ToString + "'"
                                            oRecordSet.DoQuery(strQuery)
                                            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                        Else
                                            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                        End If
                                    End If
                                Next
                            End If
                        ElseIf (oProduction.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotDisassembly) Then
                            'Update of Production Item Costs...For Disassembly Order
                            oProdItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            If oProdItem.GetByKey(oDoc_Lines.ItemCode) Then
                                oProdWhs = oProdItem.WhsInfo
                                For intWhsIndex As Integer = 0 To oProdWhs.Count - 1
                                    oProdWhs.SetCurrentLine(intWhsIndex)
                                    If oProdWhs.WarehouseCode = oDoc_Lines.WarehouseCode Then
                                        Dim dblAvgRMCost, dblAvgFLbCost, dblAgvVLbCost As Double

                                        dblAvgRMCost = CDbl(oProdWhs.UserFields.Fields.Item("U_AvgRMCst").Value)
                                        dblAvgFLbCost = CDbl(oProdWhs.UserFields.Fields.Item("U_AvgFLbCst").Value)
                                        dblAgvVLbCost = CDbl(oProdWhs.UserFields.Fields.Item("U_AvgVLbCst").Value)

                                        oProdWhs.UserFields.Fields.Item("U_AvgRMCst").Value = CDbl(((dblAvgRMCost * (oProdWhs.InStock)) - (dblRMCost * oDoc_Lines.Quantity)) / oProdWhs.InStock)
                                        oProdWhs.UserFields.Fields.Item("U_AvgFLbCst").Value = CDbl((dblAvgFLbCost * (oProdWhs.InStock)) - ((dblFixLbCost * oDoc_Lines.Quantity)) / oProdWhs.InStock)
                                        oProdWhs.UserFields.Fields.Item("U_AvgVLbCst").Value = CDbl(((dblAgvVLbCost * (oProdWhs.InStock)) - (dblVarLbCost * oDoc_Lines.Quantity)) / oProdWhs.InStock)

                                        oApplication.Company.StartTransaction()
                                        Dim intCode As Integer = oProdItem.Update()
                                        If intCode = 0 Then
                                            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            strQuery = "Update IGN1 Set U_ITCost = 'Y' Where DocEntry = '" + oDoc.DocEntry.ToString + "' And LineNum = '" + oDoc_Lines.LineNum.ToString + "'"
                                            oRecordSet.DoQuery(strQuery)
                                            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                        Else
                                            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        End Try
        Return _retVal
    End Function

    Public Sub get_Cost(ByVal strItemCode As String, ByVal strWhsCode As String, ByRef dblRMCost As Double, ByRef dblFixCost As Double, ByRef dblVarCost As Double)
        Try
            Dim oBOM As SAPbobsCOM.ProductTrees
            Dim oItem As SAPbobsCOM.Items
            Dim oItemWhs As SAPbobsCOM.ItemWarehouseInfo
            Dim oBOM_Lines As SAPbobsCOM.ProductTrees_Lines
            oBOM = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees)
            If oBOM.GetByKey(strItemCode) Then
                oBOM_Lines = oBOM.Items
                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                For bomlineindex As Integer = 0 To oBOM_Lines.Count - 1
                    oBOM_Lines.SetCurrentLine(bomlineindex)
                    If oItem.GetByKey(oBOM_Lines.ItemCode) Then
                        oItemWhs = oItem.WhsInfo
                        For intWhsIndex As Integer = 0 To oItemWhs.Count - 1
                            oItemWhs.SetCurrentLine(intWhsIndex)
                            If oItemWhs.WarehouseCode = oBOM_Lines.Warehouse Then
                                If (oItem.ItemType = SAPbobsCOM.ItemTypeEnum.itItems) Then
                                    If hasBOM(oBOM_Lines.ItemCode) Then
                                        get_Cost(oBOM_Lines.ItemCode, oBOM_Lines.Warehouse, dblRMCost, dblFixCost, dblVarCost)
                                    Else
                                        dblRMCost += (CDbl(oBOM_Lines.Quantity) / CDbl(oBOM.Quantity)) * CDbl(oItemWhs.StandardAveragePrice)
                                    End If
                                ElseIf (oItem.ItemType = SAPbobsCOM.ItemTypeEnum.itLabor) Then
                                    If oItem.UserFields.Fields.Item("U_LabType").Value = "F" Then
                                        dblFixCost += (CDbl(oBOM_Lines.Quantity) / CDbl(oBOM.Quantity)) * CDbl(oItemWhs.StandardAveragePrice)
                                    Else
                                        dblVarCost += (CDbl(oBOM_Lines.Quantity) / CDbl(oBOM.Quantity)) * CDbl(oItemWhs.StandardAveragePrice)
                                    End If
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function hasBOM(ByVal strItemCode As String) As Boolean
        Try
            Dim oBOM As SAPbobsCOM.ProductTrees
            oBOM = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees)
            If oBOM.GetByKey(strItemCode) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function update_RMPCosting_Purchase_In(ByVal oForm As SAPbouiCOM.Form, ByVal strDocType As String, ByVal strObjectKey As String)
        Dim _retVal As Boolean = True
        Try
            Dim strQuery As String
            Dim oDoc As SAPbobsCOM.Documents = Nothing
            Dim oDoc_Lines As SAPbobsCOM.Document_Lines
            Dim dblICost As Double
            Dim oDocItem As SAPbobsCOM.Items
            Dim oDocWhs As SAPbobsCOM.ItemWarehouseInfo
            Dim strTable As String = String.Empty

            Select Case strDocType
                Case frm_GRPO
                    strTable = "PDN1"
                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                Case frm_GR_INVENTORY
                    strTable = "IGN1"
                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
            End Select

            If oDoc.Browser.GetByKeys(strObjectKey) Then
                oDoc_Lines = oDoc.Lines

                For docLineindex As Integer = 0 To oDoc_Lines.Count - 1
                    oDoc_Lines.SetCurrentLine(docLineindex)
                    dblICost = 0

                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And oDoc_Lines.UserFields.Fields.Item("U_ITCost").Value = "Y" Then
                        Continue For
                    End If

                    'RMP Update
                    oDocItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                    If oDocItem.GetByKey(oDoc_Lines.ItemCode) Then
                        oDocWhs = oDocItem.WhsInfo
                        For intWhsIndex As Integer = 0 To oDocWhs.Count - 1
                            oDocWhs.SetCurrentLine(intWhsIndex)
                            If oDocWhs.WarehouseCode = oDoc_Lines.WarehouseCode Then
                                Dim dblAvgRMCost, dblAvgFLbCost, dblAgvVLbCost As Double

                                If oDoc_Lines.Currency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency Then
                                    dblICost = CDbl(oDoc_Lines.UnitPrice)
                                Else
                                    dblICost = CDbl(oDoc_Lines.UnitPrice) * CDbl(oDoc_Lines.Rate)
                                End If

                                dblAvgRMCost = CDbl(oDocWhs.UserFields.Fields.Item("U_AvgRMCst").Value)
                                dblAvgFLbCost = CDbl(oDocWhs.UserFields.Fields.Item("U_AvgFLbCst").Value)
                                dblAgvVLbCost = CDbl(oDocWhs.UserFields.Fields.Item("U_AvgVLbCst").Value)

                                oDocWhs.UserFields.Fields.Item("U_AvgRMCst").Value = CDbl(((dblICost * oDoc_Lines.Quantity) + (dblAvgRMCost * (oDocWhs.InStock - oDoc_Lines.Quantity))) / oDocWhs.InStock)
                                oDocWhs.UserFields.Fields.Item("U_AvgFLbCst").Value = CDbl(((0 * oDoc_Lines.Quantity) + (dblAvgFLbCost * (oDocWhs.InStock - oDoc_Lines.Quantity))) / oDocWhs.InStock)
                                oDocWhs.UserFields.Fields.Item("U_AvgVLbCst").Value = CDbl(((0 * oDoc_Lines.Quantity) + (dblAgvVLbCost * (oDocWhs.InStock - oDoc_Lines.Quantity))) / oDocWhs.InStock)

                                oApplication.Company.StartTransaction()

                                Dim intCode As Integer = oDocItem.Update()
                                If intCode = 0 Then
                                    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    strQuery = "Update " + strTable + " Set U_ITCost = 'Y' Where DocEntry = '" + oDoc.DocEntry.ToString + "' And LineNum = '" + oDoc_Lines.LineNum.ToString + "'"
                                    oRecordSet.DoQuery(strQuery)
                                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                Else
                                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        End Try
        Return _retVal
    End Function

    Public Function update_RMPCosting_Purchase_Out(ByVal oForm As SAPbouiCOM.Form, ByVal strDocType As String, ByVal strObjectKey As String)
        Dim _retVal As Boolean = True
        Try
            Dim strQuery As String
            Dim oDoc As SAPbobsCOM.Documents = Nothing
            Dim oDoc_Lines As SAPbobsCOM.Document_Lines
            Dim dblICost As Double
            Dim oDocItem As SAPbobsCOM.Items
            Dim oDocWhs As SAPbobsCOM.ItemWarehouseInfo
            Dim strTable As String = String.Empty

            Select Case strDocType
                Case frm_PurReturn
                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseReturns)
                Case frm_APCreditMemo
                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
            End Select

            If oDoc.Browser.GetByKeys(strObjectKey) Then
                oDoc_Lines = oDoc.Lines

                For docLineindex As Integer = 0 To oDoc_Lines.Count - 1
                    oDoc_Lines.SetCurrentLine(docLineindex)
                    dblICost = 0

                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        Continue For
                    End If

                    'RMP Update
                    oDocItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                    If oDocItem.GetByKey(oDoc_Lines.ItemCode) Then
                        If (oDoc_Lines.BaseType.ToString() <> "" And oDoc_Lines.BaseType.ToString() <> "-1") Then
                            oDocWhs = oDocItem.WhsInfo
                            For intWhsIndex As Integer = 0 To oDocWhs.Count - 1
                                oDocWhs.SetCurrentLine(intWhsIndex)
                                If oDocWhs.WarehouseCode = oDoc_Lines.WarehouseCode Then
                                    Dim dblAvgRMCost, dblAvgFLbCost, dblAgvVLbCost As Double

                                    If oDoc_Lines.Currency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency Then
                                        dblICost = CDbl(oDoc_Lines.UnitPrice)
                                    Else
                                        dblICost = CDbl(oDoc_Lines.UnitPrice) * CDbl(oDoc_Lines.Rate)
                                    End If

                                    dblAvgRMCost = CDbl(oDocWhs.UserFields.Fields.Item("U_AvgRMCst").Value)
                                    dblAvgFLbCost = CDbl(oDocWhs.UserFields.Fields.Item("U_AvgFLbCst").Value)
                                    dblAgvVLbCost = CDbl(oDocWhs.UserFields.Fields.Item("U_AvgVLbCst").Value)

                                    oDocWhs.UserFields.Fields.Item("U_AvgRMCst").Value = CDbl(((dblAvgRMCost * (oDocWhs.InStock + oDoc_Lines.Quantity)) - ((dblICost * oDoc_Lines.Quantity))) / oDocWhs.InStock)
                                    oDocWhs.UserFields.Fields.Item("U_AvgFLbCst").Value = CDbl(((dblAvgFLbCost * (oDocWhs.InStock + oDoc_Lines.Quantity))) / oDocWhs.InStock)
                                    oDocWhs.UserFields.Fields.Item("U_AvgVLbCst").Value = CDbl(((dblAgvVLbCost * (oDocWhs.InStock + oDoc_Lines.Quantity))) / oDocWhs.InStock)

                                    oApplication.Company.StartTransaction()

                                    Dim intCode As Integer = oDocItem.Update()
                                    If intCode = 0 Then
                                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                    Else
                                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        End Try
        Return _retVal
    End Function

    Public Function update_RMPCosting_Sale_In(ByVal oForm As SAPbouiCOM.Form, ByVal strDocType As String, ByVal strObjectKey As String)
        Dim _retVal As Boolean = True
        Try
            Dim strQuery As String
            Dim oDoc As SAPbobsCOM.Documents = Nothing
            Dim oDoc_Lines As SAPbobsCOM.Document_Lines
            Dim dblICost As Double
            Dim oDocItem As SAPbobsCOM.Items
            Dim oDocWhs As SAPbobsCOM.ItemWarehouseInfo
            Dim strTable As String = String.Empty

            Select Case strDocType
                Case frm_SaleReturn
                    strTable = "RDN1"
                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oReturns)
                Case frm_ARCreditMemo
                    strTable = "RIN1"
                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
            End Select

            If oDoc.Browser.GetByKeys(strObjectKey) Then
                oDoc_Lines = oDoc.Lines

                For docLineindex As Integer = 0 To oDoc_Lines.Count - 1
                    oDoc_Lines.SetCurrentLine(docLineindex)
                    dblICost = 0

                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oDoc_Lines.BaseType.ToString() = "" Then
                        Continue For
                    End If

                    'RMP Update
                    oDocItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                    If oDocItem.GetByKey(oDoc_Lines.ItemCode) Then
                        If (oDoc_Lines.BaseType.ToString() <> "" And oDoc_Lines.BaseType.ToString() <> "-1") Then

                            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            strQuery = "Select StockPrice From " + strTable + " Where DocEntry = '" + oDoc.DocEntry.ToString + "' And LineNum = '" + oDoc_Lines.LineNum.ToString + "'"
                            oRecordSet.DoQuery(strQuery)
                            If Not oRecordSet.EoF Then
                                dblICost = CDbl(oRecordSet.Fields.Item(0).Value)
                            End If

                            oDocWhs = oDocItem.WhsInfo
                            For intWhsIndex As Integer = 0 To oDocWhs.Count - 1
                                oDocWhs.SetCurrentLine(intWhsIndex)
                                If oDocWhs.WarehouseCode = oDoc_Lines.WarehouseCode Then
                                    Dim dblAvgRMCost, dblAvgFLbCost, dblAgvVLbCost As Double

                                    dblAvgRMCost = CDbl(oDocWhs.UserFields.Fields.Item("U_AvgRMCst").Value)
                                    dblAvgFLbCost = CDbl(oDocWhs.UserFields.Fields.Item("U_AvgFLbCst").Value)
                                    dblAgvVLbCost = CDbl(oDocWhs.UserFields.Fields.Item("U_AvgVLbCst").Value)

                                    oDocWhs.UserFields.Fields.Item("U_AvgRMCst").Value = CDbl(((dblICost * oDoc_Lines.Quantity) + (dblAvgRMCost * (oDocWhs.InStock - oDoc_Lines.Quantity))) / oDocWhs.InStock)
                                    oDocWhs.UserFields.Fields.Item("U_AvgFLbCst").Value = CDbl(((0 * oDoc_Lines.Quantity) + (dblAvgFLbCost * (oDocWhs.InStock - oDoc_Lines.Quantity))) / oDocWhs.InStock)
                                    oDocWhs.UserFields.Fields.Item("U_AvgVLbCst").Value = CDbl(((0 * oDoc_Lines.Quantity) + (dblAgvVLbCost * (oDocWhs.InStock - oDoc_Lines.Quantity))) / oDocWhs.InStock)

                                    oApplication.Company.StartTransaction()
                                    Dim intCode As Integer = oDocItem.Update()
                                    If intCode = 0 Then
                                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                    Else
                                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        End Try
        Return _retVal
    End Function

    Public Function update_TransferCosting(ByVal oForm As SAPbouiCOM.Form, ByVal strDocType As String, ByVal strObjectKey As String)
        Dim _retVal As Boolean = True
        Try
            Dim strQuery As String
            Dim oDoc As SAPbobsCOM.StockTransfer = Nothing
            Dim oDoc_Lines As SAPbobsCOM.StockTransfer_Lines 
            Dim dblTCost As Double
            Dim dlbItemCost As Double
            Dim oDocItem As SAPbobsCOM.Items
            Dim oDocWhs As SAPbobsCOM.ItemWarehouseInfo
            Dim strTable As String = String.Empty

            Select Case strDocType
                Case frm_I_Transfer
                    strTable = "WTR1"
            End Select

            oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
            If oDoc.Browser.GetByKeys(strObjectKey) Then
                oDoc_Lines = oDoc.Lines

                For docLineindex As Integer = 0 To oDoc_Lines.Count - 1
                    oDoc_Lines.SetCurrentLine(docLineindex)
                    dblTCost = 0

                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And oDoc_Lines.UserFields.Fields.Item("U_ITCost").Value = "Y" Then
                        Continue For
                    End If

                    'Update Transfer
                    oDocItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                    If oDocItem.GetByKey(oDoc_Lines.ItemCode) Then
                        oDocWhs = oDocItem.WhsInfo
                        For intWhsIndex As Integer = 0 To oDocWhs.Count - 1
                            oDocWhs.SetCurrentLine(intWhsIndex)
                            If oDocWhs.WarehouseCode = oDoc.FromWarehouse Then
                                Dim dblAvgRMCost_F, dblAvgFLbCost_F, dblAgvVLbCost_F As Double
                                Dim dblAvgRMCost, dblAvgFLbCost, dblAgvVLbCost As Double
                                oDocWhs = oDocItem.WhsInfo
                                'dblTCost = CDbl(oDocWhs.StandardAveragePrice)

                                dblAvgRMCost_F = CDbl(oDocWhs.UserFields.Fields.Item("U_AvgRMCst").Value)
                                dblAvgFLbCost_F = CDbl(oDocWhs.UserFields.Fields.Item("U_AvgFLbCst").Value)
                                dblAgvVLbCost_F = CDbl(oDocWhs.UserFields.Fields.Item("U_AvgVLbCst").Value)

                                For intWhsToWhs As Integer = 0 To oDocWhs.Count - 1
                                    oDocWhs.SetCurrentLine(intWhsToWhs)
                                    If oDocWhs.WarehouseCode = oDoc_Lines.WarehouseCode Then

                                        dblAvgRMCost = CDbl(oDocWhs.UserFields.Fields.Item("U_AvgRMCst").Value)
                                        dblAvgFLbCost = CDbl(oDocWhs.UserFields.Fields.Item("U_AvgFLbCst").Value)
                                        dblAgvVLbCost = CDbl(oDocWhs.UserFields.Fields.Item("U_AvgVLbCst").Value)

                                        oDocWhs.UserFields.Fields.Item("U_AvgRMCst").Value = CDbl(((dblAvgRMCost_F * oDoc_Lines.Quantity) + (dblAvgRMCost * (oDocWhs.InStock - oDoc_Lines.Quantity))) / oDocWhs.InStock)
                                        oDocWhs.UserFields.Fields.Item("U_AvgFLbCst").Value = CDbl(((dblAvgFLbCost_F * oDoc_Lines.Quantity) + (dblAvgFLbCost * (oDocWhs.InStock - oDoc_Lines.Quantity))) / oDocWhs.InStock)
                                        oDocWhs.UserFields.Fields.Item("U_AvgVLbCst").Value = CDbl(((dblAgvVLbCost_F * oDoc_Lines.Quantity) + (dblAgvVLbCost * (oDocWhs.InStock - oDoc_Lines.Quantity))) / oDocWhs.InStock)

                                    End If
                                Next

                                oApplication.Company.StartTransaction()

                                Dim intCode As Integer = oDocItem.Update()
                                If intCode = 0 Then
                                    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    strQuery = "Update " + strTable + " Set U_ITCost = 'Y' Where DocEntry = '" + oDoc.DocEntry.ToString + "' And LineNum = '" + oDoc_Lines.LineNum.ToString + "'"
                                    oRecordSet.DoQuery(strQuery)
                                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                Else
                                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                End If
                            End If

                        Next
                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        End Try
        Return _retVal
    End Function

#End Region

    Public Sub assignMatrixLineno(ByVal aGrid As SAPbouiCOM.Grid, ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        For intNo As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            aGrid.RowHeaders.SetText(intNo, intNo + 1)
        Next
        aGrid.Columns.Item("RowsHeader").TitleObject.Caption = "#"
        aform.Freeze(False)
    End Sub

#Region " Post Commission "

    'Public Function post_Commission(ByVal oForm As SAPbouiCOM.Form, ByVal strDocType As String, ByVal strObjectKey As String)
    '    Dim _retVal As Boolean = True
    '    Try
    '        Dim strCodeType As String = String.Empty
    '        Dim oPayment As SAPbobsCOM.Payments
    '        Dim oDeposit As SAPbobsCOM.Deposit
    '        Dim dtDocDate As Date
    '        Dim oService As SAPbobsCOM.CompanyService = oApplication.Company.GetCompanyService()
    '        Dim dpService As SAPbobsCOM.DepositsService = CType(oService.GetBusinessService(SAPbobsCOM.ServiceTypes.DepositsService), SAPbobsCOM.DepositsService)
    '        Dim dpsParams As SAPbobsCOM.DepositParams = CType(dpService.GetDataInterface(SAPbobsCOM.DepositsServiceDataInterfaces.dsDepositParams), SAPbobsCOM.DepositParams)

    '        Select Case strDocType
    '            Case frm_IncomingPayment
    '                oPayment = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
    '                If (oPayment.Browser.GetByKeys(strObjectKey)) Then
    '                    strCodeType = oPayment.UserFields.Fields.Item("U_RefCode").Value
    '                    dtDocDate = oPayment.DocDate
    '                End If
    '            Case frm_OutPayment
    '                oPayment = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
    '                If (oPayment.Browser.GetByKeys(strObjectKey)) Then
    '                    strCodeType = oPayment.UserFields.Fields.Item("U_RefCode").Value
    '                    dtDocDate = oPayment.DocDate
    '                End If
    '            Case frm_Deposits
    '                oDeposit = CType(dpService.GetDataInterface(SAPbobsCOM.DepositsServiceDataInterfaces.dsDeposit), SAPbobsCOM.Deposit)
    '                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
    '                    Dim strDQuery As String = " Select Top 1 Convert(VarChar(8),DeposDate,112) As DeposDate,U_RefCode From ODPS Where UserSign = '" + oApplication.Company.UserSignature.ToString() + "' Order By DeposId Desc "
    '                    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    oRecordSet.DoQuery(strDQuery)
    '                    If Not oRecordSet.EoF Then
    '                        strCodeType = oRecordSet.Fields.Item("U_RefCode").Value
    '                        dtDocDate = GetDateTimeValue(oRecordSet.Fields.Item("DeposDate").Value)
    '                    End If
    '                Else
    '                    strCodeType = oForm.Items.Item("_32").Specific.value
    '                    dtDocDate = GetDateTimeValue(oForm.Items.Item("9").Specific.value)
    '                End If
    '        End Select

    '        Dim strQuery As String = "Select Code,U_JourRem From [@OCMR] Where U_RefCode = '" & strCodeType & "' And ISNULL(U_JERef,'') = '' "
    '        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRecordSet.DoQuery(strQuery)

    '        Dim strLocalCurrency As String = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
    '        If Not oRecordSet.EoF Then
    '            Try
    '                Dim oCode As String = oRecordSet.Fields.Item("Code").Value
    '                Dim strRemarks As String = oRecordSet.Fields.Item("U_JourRem").Value

    '                Dim oJE As SAPbobsCOM.JournalEntries
    '                Dim blnStatus As Boolean = True
    '                Dim oUserTable As SAPbobsCOM.UserTable
    '                oUserTable = oApplication.Company.UserTables.Item("OCMR")
    '                oApplication.Company.StartTransaction()

    '                If oUserTable.GetByKey(oCode) Then
    '                    Dim intCurrentLine As Integer = 0

    '                    oJE = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
    '                    oJE.ReferenceDate = System.DateTime.Now
    '                    oJE.TaxDate = dtDocDate.Date
    '                    oJE.DueDate = dtDocDate.Date

    '                    oJE.Reference = "Commission Charges"
    '                    oJE.Memo = strRemarks
    '                    oJE.Reference2 = strCodeType
    '                    oJE.Reference3 = strCodeType

    '                    If CDbl(oUserTable.UserFields.Fields.Item("U_CommCh").Value) > 0 Then
    '                        oJE.Lines.SetCurrentLine(intCurrentLine)
    '                        oJE.Lines.AccountCode = getAccount(oUserTable.UserFields.Fields.Item("U_BankGL").Value)

    '                        If strLocalCurrency = oUserTable.UserFields.Fields.Item("U_Currency").Value Then
    '                            oJE.Lines.Credit = CDbl(oUserTable.UserFields.Fields.Item("U_CommCh").Value)
    '                        Else
    '                            oJE.Lines.FCCurrency = oUserTable.UserFields.Fields.Item("U_Currency").Value
    '                            oJE.Lines.FCCredit = CDbl(oUserTable.UserFields.Fields.Item("U_CommCh").Value)
    '                        End If

    '                        intCurrentLine += 1
    '                        oJE.Lines.Add()
    '                        oJE.Lines.SetCurrentLine(intCurrentLine)
    '                        oJE.Lines.AccountCode = getAccount(oUserTable.UserFields.Fields.Item("U_CommGL").Value)

    '                        If strLocalCurrency = oUserTable.UserFields.Fields.Item("U_Currency").Value Then
    '                            oJE.Lines.Debit = CDbl(oUserTable.UserFields.Fields.Item("U_CommCh").Value)
    '                        Else
    '                            oJE.Lines.FCCurrency = oUserTable.UserFields.Fields.Item("U_Currency").Value
    '                            oJE.Lines.FCDebit = CDbl(oUserTable.UserFields.Fields.Item("U_CommCh").Value)
    '                        End If
    '                        intCurrentLine += 1
    '                    End If

    '                    Dim intCode As Integer = oJE.Add()
    '                    If intCode <> 0 Then
    '                        oApplication.SBO_Application.SetStatusBarMessage(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
    '                        _retVal = False
    '                        blnStatus = False
    '                    Else
    '                        Dim intJE As Integer
    '                        intJE = oApplication.Company.GetNewObjectKey()
    '                        If oJE.GetByKey(intJE) Then
    '                            oUserTable.UserFields.Fields.Item("U_JERef").Value = intJE.ToString()
    '                            If oUserTable.Update() <> 0 Then
    '                                blnStatus = False
    '                            End If
    '                        End If
    '                        oApplication.SBO_Application.SetStatusBarMessage("Commission Charges Posted Sucessfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
    '                    End If
    '                End If
    '                If blnStatus Then
    '                    If oApplication.Company.InTransaction Then
    '                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
    '                    End If
    '                End If
    '            Catch ex As Exception
    '                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
    '                oApplication.SBO_Application.SetStatusBarMessage(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
    '            End Try
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    '    Return _retVal
    'End Function

#End Region

#Region "Create Commission Reference"

    Public Sub addCommissionReference(ByRef strCode As String)
        Try
            Dim oCommCharges As SAPbobsCOM.UserTable
            oCommCharges = oApplication.Company.UserTables.Item("OCMR")
            Dim intCode As Integer = getMaxCode("@OCMR", "Code")
            oCommCharges.Code = intCode.ToString()
            oCommCharges.Name = intCode.ToString()
            oCommCharges.UserFields.Fields.Item("U_RefCode").Value = String.Format("{0:000000000}", intCode)
            Dim intStatus As Integer = oCommCharges.Add()
            If intStatus = 0 Then
                strCode = String.Format("{0:000000000}", intCode)
            End If
        Catch ex As Exception
            oApplication.SBO_Application.SetStatusBarMessage(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
    End Sub

#End Region

    Public Function validateRefExist(ByVal strRef As String)
        Try
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Code From [@OCMR] Where U_RefCode = '" & strRef & "'")
            If Not oRecordSet.EoF Then
                Return True
            End If
            Return False
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function removeCommission(ByVal strRef As String)
        Dim _retVal As Boolean = True
        Try
            Dim oUserTable As SAPbobsCOM.UserTable
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Code From [@OCMR] Where U_RefCode = '" & strRef & "'")
            If Not oRecordSet.EoF Then
                oUserTable = oApplication.Company.UserTables.Item("OCMR")
                If oUserTable.GetByKey(oRecordSet.Fields.Item("Code").Value) Then
                    If oUserTable.UserFields.Fields.Item("U_JERef").Value.ToString().Trim = "" Then
                        If oUserTable.Remove() <> 0 Then

                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return _retVal
    End Function

#Region "GetAccount"

    Public Sub getBankAccount(ByVal strAbsEntry As String, ByRef strAccount As String)
        Try
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ' oRecordSet.DoQuery(" Select GLAccount From DSC1 Where AbsEntry = '" & strAbsEntry & "'")
            oRecordSet.DoQuery("Select GLAccount,isnull(FormatCode ,'') 'FormatCode' From DSC1 T0 left outer join OACT T1  on T1.AcctCode=T0.GLAccount  where T0.AbsEntry='" & strAbsEntry & "'")

            If Not oRecordSet.EoF Then
                ' strAccount = oRecordSet.Fields.Item("GLAccount").Value
                strAccount = oRecordSet.Fields.Item("FormatCode").Value
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub getCommissionAccount(ByVal strCode As String, ByRef strAccount As String)
        Try
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select U_CMAcct From [@OCMT] Where Code = '" & strCode & "'")
            If Not oRecordSet.EoF Then
                strAccount = oRecordSet.Fields.Item("U_CMAcct").Value
            End If
        Catch ex As Exception

        End Try
    End Sub
#End Region

    Public Sub addReference(ByRef strCode As String)
        Try
            Dim oUDT As SAPbobsCOM.UserTable
            oUDT = oApplication.Company.UserTables.Item("PRT2")
            Dim intCode As Integer = getMaxCode("@PRT2", "Code")
            oUDT.Code = intCode.ToString()
            oUDT.Name = intCode.ToString()
            oUDT.UserFields.Fields.Item("U_Reference").Value = String.Format("{0:000000000}", intCode)
            Dim intStatus As Integer = oUDT.Add()
            If intStatus = 0 Then
                strCode = String.Format("{0:000000000}", intCode)
            End If
        Catch ex As Exception
            oApplication.SBO_Application.SetStatusBarMessage(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
    End Sub

    Public Sub convertCurrency(ByVal strCurrency As String, ByVal dblAmount As String, ByVal dtDate As Date, _
 ByVal strRCurrency As String, ByRef dblForeignAmt As Double)
        Try
            Try
                Dim oExRecordSet As SAPbobsCOM.Recordset
                Dim dblRExRate, dblAExRate As Double
                oExRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim dblLocalCurrency As String = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
                If strCurrency <> dblLocalCurrency Then
                    oExRecordSet.DoQuery("Select Rate From ORTT Where Currency = '" + strCurrency + "' And Convert(VarChar(8),RateDate,112) = Convert(VarChar(8),GetDate(),112)")
                    If Not oExRecordSet.EoF Then
                        oExRecordSet.DoQuery("Select isnull(Rate,1) 'Rate' From ORTT Where Currency = '" + strRCurrency + "' And Convert(VarChar(8),RateDate,112) = Convert(VarChar(8),GetDate(),112)")
                        If Not oExRecordSet.EoF Then
                            dblRExRate = oExRecordSet.Fields.Item("Rate").Value
                            If strRCurrency = dblLocalCurrency Then
                                dblForeignAmt = dblAmount / dblRExRate
                            Else
                                oExRecordSet.DoQuery("Select isnull(Rate,1) 'Rate' From ORTT Where Currency = '" + strCurrency + "' And Convert(VarChar(8),RateDate,112) = Convert(VarChar(8),GetDate(),112)")
                                If Not oExRecordSet.EoF Then
                                    dblAExRate = oExRecordSet.Fields.Item("Rate").Value
                                    dblForeignAmt = ((dblAmount * dblAExRate) / dblRExRate)
                                End If
                            End If
                        End If
                    End If
                Else
                    oExRecordSet.DoQuery("Select Rate From ORTT Where Currency = '" + strCurrency + "' And Convert(VarChar(8),RateDate,112) = Convert(VarChar(8),GetDate(),112)")
                    If Not oExRecordSet.EoF Then
                        dblAExRate = oExRecordSet.Fields.Item("Rate").Value
                        dblForeignAmt = dblAmount * dblAExRate
                    End If
                End If
            Catch ex As Exception
                Throw ex
            End Try
        Catch ex As Exception

        End Try
    End Sub

    'Public Sub AddFreightControls(ByVal oForm As SAPbouiCOM.Form, ByVal strTable As String)
    '    Try
    '        oApplication.Utilities.AddControls(oForm, "_53", "230", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 0, 0, "", "Freight Reference", 0, 0, 0, False)
    '        oApplication.Utilities.AddControls(oForm, "_52", "222", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 0, 0, "", "", 0, 0, 0, True)
    '        oForm.Items.Item("_53").Visible = True
    '        oForm.Items.Item("_52").Visible = True
    '        oForm.Items.Item("_53").LinkTo = "_52"
    '        dataBind(oForm, strTable)
    '        oForm.Items.Item("_52").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    Private Sub dataBind(ByVal oForm As SAPbouiCOM.Form, ByVal strTable As String)
        Try
            oEditText = oForm.Items.Item("_52").Specific
            oEditText.DataBind.SetBound(True, strTable, "U_RefCode")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub addReference(ByRef strCode As String, ByVal strDocType As String)
        Try
            Dim oUDT As SAPbobsCOM.UserTable
            oUDT = oApplication.Company.UserTables.Item("OFRT")
            Dim intCode As Integer = getMaxCode("@OFRT", "Code")
            oUDT.Code = intCode.ToString()
            oUDT.Name = intCode.ToString()
            oUDT.UserFields.Fields.Item("U_RefCode").Value = String.Format("{0:000000000}", intCode)
            oUDT.UserFields.Fields.Item("U_DocType").Value = strDocType
            Dim intStatus As Integer = oUDT.Add()
            If intStatus = 0 Then
                strCode = String.Format("{0:000000000}", intCode)
            End If
        Catch ex As Exception
            oApplication.SBO_Application.SetStatusBarMessage(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
    End Sub

    Public Function getDocType(ByVal strFrmType) As String
        Dim _retVal As String = String.Empty
        Try
            Select Case strFrmType
                'Case frm_Quotation
                '    _retVal = "23"
                'Case frm_ORDR
                '    _retVal = "23"
                Case frm_Delivery
                    _retVal = "15"
                Case frm_SaleReturn
                    _retVal = "16"
                Case frm_INVOICES, frm_INVOICESPAYMENT
                    _retVal = "13"
                Case frm_ARCreditMemo
                    _retVal = "14"
                Case frm_ARReverseInvoice
                    _retVal = "14"
            End Select
        Catch ex As Exception
            oApplication.SBO_Application.SetStatusBarMessage(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        Return _retVal
    End Function

    Public Sub addFreightAmt(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal strRefCode As String)
        Try
            Dim oUDT As SAPbobsCOM.UserTable
            oUDT = oApplication.Company.UserTables.Item("FRT1")

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Code,U_FreID From [@FRT1] Where U_RefCode = '" & strRefCode & "'")
            If Not oRecordSet.EoF Then
                oUDT = oApplication.Company.UserTables.Item("FRT1")
                For index As Integer = 1 To oMatrix.RowCount
                    Dim blnRecExist As Boolean = False
                    oRecordSet.MoveFirst()
                    While Not oRecordSet.EoF
                        If CType(oMatrix.Columns.Item("1").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_FreID").Value Then
                            blnRecExist = True
                            If oUDT.GetByKey(oRecordSet.Fields.Item("Code").Value) Then
                                oUDT.UserFields.Fields.Item("U_Currency").Value = CType(oMatrix.Columns.Item("U_Currency").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                                oUDT.UserFields.Fields.Item("U_PAmount").Value = CType(oMatrix.Columns.Item("U_PAmount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                                oUDT.UserFields.Fields.Item("U_PDiscount").Value = CType(oMatrix.Columns.Item("U_PDiscount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                                oUDT.UserFields.Fields.Item("U_Total").Value = CType(oMatrix.Columns.Item("3").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                                Dim intStatus As Integer = oUDT.Update()
                                If intStatus = 0 Then

                                End If
                            End If
                        End If
                        oRecordSet.MoveNext()
                    End While

                    If Not blnRecExist Then
                        Dim intCode As Integer = getMaxCode("@FRT1", "Code")
                        oUDT.Code = intCode.ToString()
                        oUDT.Name = intCode.ToString()
                        oUDT.UserFields.Fields.Item("U_RefCode").Value = strRefCode
                        oUDT.UserFields.Fields.Item("U_FreID").Value = CType(oMatrix.Columns.Item("1").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                        oUDT.UserFields.Fields.Item("U_Currency").Value = CType(oMatrix.Columns.Item("U_Currency").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                        oUDT.UserFields.Fields.Item("U_PAmount").Value = CType(oMatrix.Columns.Item("U_PAmount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                        oUDT.UserFields.Fields.Item("U_PDiscount").Value = CType(oMatrix.Columns.Item("U_PDiscount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                        oUDT.UserFields.Fields.Item("U_Total").Value = CType(oMatrix.Columns.Item("3").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                        Dim intStatus As Integer = oUDT.Add()
                        If intStatus = 0 Then

                        End If
                    End If
                Next
            Else
                For index As Integer = 1 To oMatrix.RowCount
                    Dim intCode As Integer = getMaxCode("@FRT1", "Code")
                    oUDT.Code = intCode.ToString()
                    oUDT.Name = intCode.ToString()
                    oUDT.UserFields.Fields.Item("U_RefCode").Value = strRefCode
                    oUDT.UserFields.Fields.Item("U_FreID").Value = CType(oMatrix.Columns.Item("1").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                    oUDT.UserFields.Fields.Item("U_Currency").Value = CType(oMatrix.Columns.Item("U_Currency").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                    oUDT.UserFields.Fields.Item("U_PAmount").Value = CType(oMatrix.Columns.Item("U_PAmount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                    oUDT.UserFields.Fields.Item("U_PDiscount").Value = CType(oMatrix.Columns.Item("U_PDiscount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                    oUDT.UserFields.Fields.Item("U_Total").Value = CType(oMatrix.Columns.Item("3").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                    Dim intStatus As Integer = oUDT.Add()
                    If intStatus = 0 Then

                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.SBO_Application.SetStatusBarMessage(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
    End Sub

    Public Function removeFreight(ByVal strRef As String)
        Dim _retVal As Boolean = True
        Try
            Dim oUserTable As SAPbobsCOM.UserTable
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Code From [@OFRT] Where U_RefCode = '" & strRef & "'")
            If Not oRecordSet.EoF Then

                'Parent
                oUserTable = oApplication.Company.UserTables.Item("OFRT")
                If oUserTable.GetByKey(oRecordSet.Fields.Item("Code").Value) Then
                    If oUserTable.Remove() <> 0 Then

                    End If
                End If

                'Child
                oRecordSet.DoQuery("Select Code From [@FRT1] Where U_RefCode = '" & strRef & "'")
                If Not oRecordSet.EoF Then
                    While Not oRecordSet.EoF
                        oUserTable = oApplication.Company.UserTables.Item("FRT1")
                        If oUserTable.GetByKey(oRecordSet.Fields.Item("Code").Value) Then
                            If oUserTable.Remove() <> 0 Then

                            End If
                        End If
                        oRecordSet.MoveNext()
                    End While
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return _retVal
    End Function

    Public Sub addPromotionReference(ByRef strCode As String)
        Try
            Dim oUDT As SAPbobsCOM.UserTable
            oUDT = oApplication.Company.UserTables.Item("OPRF")
            Dim intCode As Integer = getMaxCode("@OPRF", "Code")
            oUDT.Code = String.Format("{0:000000000}", intCode)
            oUDT.Name = String.Format("{0:000000000}", intCode)
            Dim intStatus As Integer = oUDT.Add()
            If intStatus = 0 Then
                strCode = String.Format("{0:000000000}", intCode)
            End If
        Catch ex As Exception
            oApplication.SBO_Application.SetStatusBarMessage(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
    End Sub

    Public Function removePromotion(ByVal oMatrix As SAPbouiCOM.Matrix)
        Dim _retVal As Boolean = True
        Try
            Dim oUserTable As SAPbobsCOM.UserTable
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            For index As Integer = 1 To oMatrix.RowCount
                Dim strRef As String = CType(oMatrix.Columns.Item("U_PrRef").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value

                If strRef.Length > 0 Then
                    oRecordSet.DoQuery("Select Code From [@OPRF] Where Code = '" & strRef & "'")
                    If Not oRecordSet.EoF Then
                        oUserTable = oApplication.Company.UserTables.Item("OPRF")
                        If oUserTable.GetByKey(oRecordSet.Fields.Item("Code").Value) Then
                            If oUserTable.Remove() <> 0 Then

                            End If
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
        Return _retVal
    End Function

End Class
