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
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
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

            If Not (strTab = "OWHS" Or strTab = "OITM" Or strTab = "INV1" Or strTab = "OWTR" Or strTab = "OUSR" Or strTab = "OITW" Or strTab = "RDR1" Or strTab = "DLN1" Or strTab = "IGN1" Or strTab = "ODSC" Or strTab = "ORCT" Or strTab = "ODPS" Or strTab = "OPRJ" Or strTab = "OEXD" Or strTab = "RDR3" Or strTab = "ORDR") Then
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
            Throw ex
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
                If (objUserFieldMD.Add() <> 0) Then
                    MsgBox(oApplication.Company.GetLastErrorDescription)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            Else
            End If

        Catch ex As Exception
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
        Dim oRecordSet As SAPbobsCOM.Recordset
        Try
            If Table = "OCMT" Or Table = "PSP1" Or Table = "OPRT" Then
                Table = "@" + Table
            End If
            strSQL = "SELECT COUNT(*) FROM CUFD WHERE TableID = '" & Table & "' AND AliasID = '" & Column & "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strSQL)

            If oRecordSet.Fields.Item(0).Value = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
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
            Throw ex

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
                                        Optional ByVal strChildTbl As String = "", Optional ByVal nObjectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_Document)

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

        Catch ex As Exception
            Throw ex

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
            oApplication.Company.StartTransaction()

            'addField("OUSR", "HideAmt", "Hide Total Amount", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            'addField("OPRJ", "CardCode", "Customer", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            'addField("OPRJ", "CardName", "Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            'AddFields("OPRJ", "CreLimit", "Credit Limit", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            'AddFields("OPRJ", "ComLimit", "Commitment Limit", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            'AddFields("OPRJ", "Currency", "BP Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 3)

            'AddTables("OPSP", "Project Special Price", SAPbobsCOM.BoUTBTableType.bott_Document)
            'AddFields("OPSP", "PrjCode", "Project Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("OPSP", "PrjName", "Project Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            'AddFields("OPSP", "CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("OPSP", "CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            'AddFields("OPSP", "EffFrom", "Effective From", SAPbobsCOM.BoFieldTypes.db_Date)
            'AddFields("OPSP", "EffTo", "Effective To", SAPbobsCOM.BoFieldTypes.db_Date)

            'AddTables("PSP1", "Project Special Items", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            'AddFields("PSP1", "ItmCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("PSP1", "ItmName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            'AddFields("PSP1", "Currency", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 3)
            'AddFields("PSP1", "PriceList", "Price List", SAPbobsCOM.BoFieldTypes.db_Alpha, , 3)
            'AddFields("PSP1", "UnitPrice", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            'addField("PSP1", "DisType", "Discount Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "D,P", "Discount,Price", "")
            'AddFields("PSP1", "Discount", "Discount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            'AddFields("PSP1", "DisPrice", "Price After Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddTables("OPSP", "Project Setup", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("OPSP", "UserName", "User Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("OPSP", "Password", "Password", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddFields("OITW", "AvgRMCst", "Average RM Cost", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("OITW", "AvgFLbCst", "Average Fixed Labour  Cost", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("OITW", "AvgVLbCst", "Average Variable Labour  Cost", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("OITW", "RMAcct", "Raw Material Acct", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("OITW", "FLAcct", "Fixed Labour Acct", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("OITW", "VLAcct", "Variable Labour Acct", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            AddFields("OWHS", "RMAcct", "Raw Material Acct", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("OWHS", "FLAcct", "Fixed Labour Acct", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("OWHS", "VLAcct", "Variable Labour Acct", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            addField("OITM", "LabType", "Labour Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "F,V", "Fixed,Variable", "")
            AddFields("DLN1", "JEDocEty", "Journal DE", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("IGN1", "ITCost", "Item Cost Update", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)

            'AddFields("RDR1", "SPDocEty", "Special Price DE", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            'Commission Charge Account in Banking...
            'AddFields("ODSC", "CMAcct", "Commission Acct", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            'Commission Type Table
            'AddTables("OCMT", "Commission Type", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            'addField("OCMT", "DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "P,D", "Payment,Deposit", "P")
            'AddFields("OCMT", "CMAcct", "Commission Acct", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            'Commission Charges Reference Table
            'AddTables("OCMR", "Commission Reference", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            'AddFields("OCMR", "RefCode", "Commission Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            'AddFields("OCMR", "BankCode", "Bank Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            'AddFields("OCMR", "BankGL", "Bank GL", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            'AddFields("OCMR", "CMType", "Commission Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            'AddFields("OCMR", "CommGL", "Commission GL", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            'AddFields("OCMR", "Currency", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 3)
            'AddFields("OCMR", "CommCh", "Commission Charges", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            'AddFields("OCMR", "JourRem", "Journal Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            'AddFields("OCMR", "JERef", "Journal Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)

            'Commission Reference
            'AddFields("ORCT", "RefCode", "Commission Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            'AddFields("ODPS", "RefCode", "Commission Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)

            'Promotion Template
            'AddTables("OPRM", "Promotion Template", SAPbobsCOM.BoUTBTableType.bott_Document)
            'AddFields("OPRM", "PrCode", "Promotion Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("OPRM", "PrName", "Promotion Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            'AddFields("OPRM", "EffFrom", "Effective From", SAPbobsCOM.BoFieldTypes.db_Date)
            'AddFields("OPRM", "EffTo", "Effective To", SAPbobsCOM.BoFieldTypes.db_Date)
            'AddFields("OPRM", "Active", "Is Active", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)

            ''Promotion Items
            'AddTables("PRM1", "Promotion Items", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            'AddFields("PRM1", "ItmCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("PRM1", "ItmName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            'AddFields("PRM1", "Qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            'AddFields("PRM1", "OffCode", "Offer Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("PRM1", "OffName", "Offer Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            'AddFields("PRM1", "OQty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            'AddFields("PRM1", "ODis", "Discount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)

            ''Commission Charges Reference Table
            'AddTables("OCPR", "Customer Promotion Mapping", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            'AddFields("OCPR", "CustCode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("OCPR", "PrCode", "Promotion Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("OCPR", "EffFrom", "Effective From", SAPbobsCOM.BoFieldTypes.db_Date)
            'AddFields("OCPR", "EffTo", "Effective To", SAPbobsCOM.BoFieldTypes.db_Date)
            'AddFields("OCPR", "ItmCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("OCPR", "ItmName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            'AddFields("OCPR", "Qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            'AddFields("OCPR", "OffCode", "Offer Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("OCPR", "OffName", "Offer Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            'AddFields("OCPR", "OQty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            'AddFields("OCPR", "ODis", "Discount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)

            'AddFields("RDR1", "PrCode", "Promotion Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'addField("RDR1", "PrmApp", "Promotion Applied", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            'AddFields("RDR1", "PrRef", "Promotion Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            'addField("RDR1", "IType", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "R,F", "Regular,Free", "R")

            'AddFields("OEXD", "Currency", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 3)
            'AddFields("OEXD", "Name", "Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("OEXD", "PAmount", "Predetermined Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            'AddFields("OEXD", "PDiscount", "Predetermined Discount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)

            'AddFields("RDR3", "Currency", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 3)
            'AddFields("RDR3", "PAmount", "Predetermined Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            'AddFields("RDR3", "PDiscount", "Predetermined Discount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            ''AddFields("RDR3", "DAmount", "Discount Amount", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            'addField("RDR3", "FCalcu", "Freight Calculated", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")

            'AddTables("OPRF", "Promotion Reference", SsssAPbobsCOM.BoUTBTableType.bott_NoObject)

            'AddFields("ORDR", "RefCode", "Freight Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            'AddTables("OFRT", "Freight Document", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            'AddFields("OFRT", "RefCode", "Freight Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            'AddFields("OFRT", "DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 3)

            'AddTables("FRT1", "Freight Document", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            'AddFields("FRT1", "RefCode", "Freight Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            'AddFields("FRT1", "FreID", "Freight Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 8)
            'AddFields("FRT1", "Currency", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 3)
            'AddFields("FRT1", "PAmount", "Predetermined Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            'AddFields("FRT1", "PDiscount", "Predetermined Discount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            ''AddFields("FRT1", "DAmount", "Discount Amount", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            'AddFields("FRT1", "Total", "Freight Total", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)

            ''Promotion Template (Header)
            'AddTables("OPRT", "Promotion Template", SAPbobsCOM.BoUTBTableType.bott_Document)
            'AddFields("OPRT", "PrmCode", "Promotion Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("OPRT", "PrmName", "Promotion Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            'AddFields("OPRT", "EffFrom", "Effective From", SAPbobsCOM.BoFieldTypes.db_Date)
            'AddFields("OPRT", "EffTo", "Effective To", SAPbobsCOM.BoFieldTypes.db_Date)
            'addField("OPRT", "PrmType", "Promotion Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "I,Q,V", "Item Based,Quantity Based,Volume Based", "I")
            'AddFields("OPRT", "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            'AddFields("OPRT", "TotalAmt", "Document Total", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            ''Will be Only for Quantity Or Volume Promotion Type
            'AddFields("OPRT", "PrmMet", "Promotion Method", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1) 'F-Free of Goods,D-Discount
            ''This will be applied only for Discount
            'AddFields("OPRT", "Discount", "Discount %", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            'AddFields("OPRT", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)
            'AddFields("OPRT", "Reference", "Additional Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)

            ''Promotion Template - Child - 1
            'AddTables("PRT1", "Promotion Child - 1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            'AddFields("PRT1", "ItmCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("PRT1", "ItmDesc", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            'AddFields("PRT1", "MinQty", "Minimum Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            'AddFields("PRT1", "Reference", "Additional Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)

            ''Promotion Template - Child - 2
            'AddTables("PRT2", "Promotion Child - 2", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            'AddFields("PRT2", "Reference", "Additional Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            'AddFields("PRT2", "ItmCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("PRT2", "ItmDesc", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            'AddFields("PRT2", "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            'AddFields("PRT2", "Discount", "Discount %", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)

            'AddTables("PRT3", "Promotion Child - 3", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            'AddFields("PRT3", "CustCode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("PRT3", "PrmCode", "Promotion Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("PRT3", "EffFrom", "Effective From", SAPbobsCOM.BoFieldTypes.db_Date)
            'AddFields("PRT3", "EffTo", "Effective To", SAPbobsCOM.BoFieldTypes.db_Date)

            '---- User Defined Object
            CreateUDO()

            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
            oApplication.SBO_Application.StatusBar.SetText("Database creation completed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Catch ex As Exception
            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            Throw ex
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Public Sub CreateUDO()
        Try
            AddUDO("UDO_OPSP", "Project Setup", "OPSP", "U_UserName", , , SAPbobsCOM.BoUDOObjType.boud_Document)
            'AddUDO("UDO_OPSP", "Project Special Price", "OPSP", "U_PrjCode", "U_PrjName", "PSP1", SAPbobsCOM.BoUDOObjType.boud_Document)
            'AddUDO("UDO_OPRM", "Promotion Template", "OPRM", "U_PrCode", "U_PrName", "PRM1", SAPbobsCOM.BoUDOObjType.boud_Document)
            'AddUDO("UDO_OPRT", "Promotion Template", "OPRT", "U_PrmCode", "U_PrmName", "PRT1", SAPbobsCOM.BoUDOObjType.boud_Document)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

End Class
