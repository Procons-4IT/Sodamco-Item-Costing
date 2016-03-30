Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsCommissionCharges
    Inherits clsBase
    Private oGrid As SAPbouiCOM.Grid
    Private oDtSpecialPriceList As SAPbouiCOM.DataTable
    Private strQuery As String
    Private oEditText As SAPbouiCOM.EditText
    Private oComboBox As SAPbouiCOM.ComboBox
    Private oRecordSet As SAPbobsCOM.Recordset
    Private strCode As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub LoadForm(ByVal strBankRefCode As String, ByVal strDocType As String)
        Try
            oForm = oApplication.Utilities.LoadForm(xml_CommCharges, frm_CommCharges)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            initializeDataSource(oForm)
            addChooseFromList(oForm)
            loadValues(oForm, strDocType)
            initialize(oForm, strBankRefCode)
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_CommCharges Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "13" Then
                                    If Not validate(oForm) Then
                                        BubbleEvent = False
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                                removeCommission(oForm)
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "13" And Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    If (UpdateValues(oForm)) Then
                                        oForm.Close()
                                    End If
                                ElseIf pVal.ItemUID = "13" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oForm.Close()
                                ElseIf pVal.ItemUID = "14" Then
                                    oForm.Close()
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim strAccount As String
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If pVal.ItemUID = "6" Or pVal.ItemUID = "10" Then
                                        strAccount = oDataTable.GetValue("FormatCode", 0)
                                        Try
                                            oForm.Items.Item(pVal.ItemUID).Specific.value = strAccount
                                        Catch ex As Exception
                                            oForm.Items.Item(pVal.ItemUID).Specific.value = strAccount
                                        End Try
                                    End If
                                Catch ex As Exception
                                    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim strAccount As String = String.Empty
                                If pVal.ItemUID = "4" Then
                                    oApplication.Utilities.getBankAccount(CType(oForm.Items.Item("4").Specific, SAPbouiCOM.ComboBox).Value, strAccount)
                                    oForm.Items.Item("6").Enabled = True
                                    oForm.Items.Item("6").Specific.value = strAccount
                                    oForm.Items.Item("12").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oForm.Items.Item("6").Enabled = False
                                ElseIf pVal.ItemUID = "8" Then
                                    oApplication.Utilities.getCommissionAccount(CType(oForm.Items.Item("8").Specific, SAPbouiCOM.ComboBox).Value, strAccount)
                                    oForm.Items.Item("10").Enabled = True
                                    oForm.Items.Item("10").Specific.value = strAccount
                                    oForm.Items.Item("12").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oForm.Items.Item("10").Enabled = False
                                    oForm.Items.Item("12").Specific.value = CType(oForm.Items.Item("8").Specific, SAPbouiCOM.ComboBox).Selected.Description
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form, ByVal strRefCode As String)
        Try
            oComboBox = oForm.Items.Item("4").Specific
            Dim oUserTable As SAPbobsCOM.UserTable
            oUserTable = oApplication.Company.UserTables.Item("OCMR")
            Dim strQuery As String = "Select Code,U_RefCode,U_BankCode,U_BankGL,U_CMType,U_CommGL,U_CommCh,U_JourRem,U_JERef,U_Currency From [@OCMR] Where U_RefCode = '" & strRefCode & "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oForm.Items.Item("18").Specific.Value = oRecordSet.Fields.Item("Code").Value
                oForm.Items.Item("2").Specific.Value = oRecordSet.Fields.Item("U_RefCode").Value
                oComboBox.Select(oRecordSet.Fields.Item("U_BankCode").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                oForm.Items.Item("6").Specific.Value = oRecordSet.Fields.Item("U_BankGL").Value
                oComboBox = oForm.Items.Item("8").Specific
                oComboBox.Select(oRecordSet.Fields.Item("U_CMType").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                oForm.Items.Item("10").Specific.Value = oRecordSet.Fields.Item("U_CommGL").Value
                oForm.Items.Item("16").Specific.Value = oRecordSet.Fields.Item("U_CommCh").Value
                oForm.Items.Item("12").Specific.Value = oRecordSet.Fields.Item("U_JourRem").Value
                oForm.Items.Item("20").Specific.Value = oRecordSet.Fields.Item("U_JERef").Value
                oComboBox = oForm.Items.Item("23").Specific
                oComboBox.Select(oRecordSet.Fields.Item("U_Currency").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                If oForm.Items.Item("20").Specific.Value.ToString.Length > 0 Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                    CType(oForm.Items.Item("13").Specific, SAPbouiCOM.Button).Caption = "Ok"
                    enable(oForm, False)
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub loadValues(ByVal oForm As SAPbouiCOM.Form, ByVal strDocType As String)
        Try
            Dim strQuery As String = " Select AbsEntry,BankCode+'-'+Account As BankCode From DSC1 Order By AbsEntry "
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)
            oComboBox = oForm.Items.Item("4").Specific
            oComboBox.ValidValues.Add("", "")
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oComboBox.ValidValues.Add(oRecordSet.Fields.Item("AbsEntry").Value, oRecordSet.Fields.Item("BankCode").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            strQuery = "Select Code,Name From [@OCMT] Where U_DocType = '" & strDocType.Trim() & "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)
            oComboBox = oForm.Items.Item("8").Specific
            oComboBox.ValidValues.Add("", "")
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oComboBox.ValidValues.Add(oRecordSet.Fields.Item("Code").Value, oRecordSet.Fields.Item("Name").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            strQuery = "Select CurrCode,CurrName From OCRN"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)
            oComboBox = oForm.Items.Item("23").Specific
            oComboBox.ValidValues.Add("", "")
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oComboBox.ValidValues.Add(oRecordSet.Fields.Item("CurrCode").Value, oRecordSet.Fields.Item("CurrName").Value)
                    oRecordSet.MoveNext()
                End While
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub initializeDataSource(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.DataSources.UserDataSources.Add("udsComCod", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 50)
            oForm.DataSources.UserDataSources.Add("udsComRef", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 50)
            oForm.DataSources.UserDataSources.Add("udsBank", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 50)
            oForm.DataSources.UserDataSources.Add("udsBankGL", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200)
            oForm.DataSources.UserDataSources.Add("udsCommTyp", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 30)
            oForm.DataSources.UserDataSources.Add("udsCommGL", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200)
            oForm.DataSources.UserDataSources.Add("udsCharges", SAPbouiCOM.BoDataType.dt_PRICE, 50)
            oForm.DataSources.UserDataSources.Add("udsCurren", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3)
            oForm.DataSources.UserDataSources.Add("udsRemarks", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 50)
            oForm.DataSources.UserDataSources.Add("udsJourRmk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)

            oEditText = oForm.Items.Item("18").Specific
            oEditText.DataBind.SetBound(True, "", "udsComCod")
            oEditText = oForm.Items.Item("2").Specific
            oEditText.DataBind.SetBound(True, "", "udsComRef")
            oComboBox = oForm.Items.Item("4").Specific
            oComboBox.DataBind.SetBound(True, "", "udsBank")
            oEditText = oForm.Items.Item("6").Specific
            oEditText.DataBind.SetBound(True, "", "udsBankGL")
            oComboBox = oForm.Items.Item("8").Specific
            oComboBox.DataBind.SetBound(True, "", "udsCommTyp")

            oEditText = oForm.Items.Item("10").Specific
            oEditText.DataBind.SetBound(True, "", "udsCommGL")
            oComboBox = oForm.Items.Item("23").Specific
            oComboBox.DataBind.SetBound(True, "", "udsCurren")
            oEditText = oForm.Items.Item("16").Specific
            oEditText.DataBind.SetBound(True, "", "udsCharges")
            oEditText = oForm.Items.Item("12").Specific
            oEditText.DataBind.SetBound(True, "", "udsRemarks")

            oEditText = oForm.Items.Item("20").Specific
            oEditText.DataBind.SetBound(True, "", "udsJourRmk")

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub enable(ByVal oForm As SAPbouiCOM.Form, ByVal blnStatus As Boolean)
        Try
            oForm.Items.Item("12").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("18").Enabled = False
            oForm.Items.Item("2").Enabled = False
            oForm.Items.Item("4").Enabled = False
            oForm.Items.Item("6").Enabled = False
            oForm.Items.Item("8").Enabled = False
            oForm.Items.Item("10").Enabled = False
            oForm.Items.Item("23").Enabled = False
            'oForm.Items.Item("12").Enabled = False
            oForm.Items.Item("16").Enabled = False
            oForm.Items.Item("20").Enabled = False
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function validate(ByVal oForm As SAPbouiCOM.Form)
        Dim _retVal As Boolean = True
        Try
            If CType(oForm.Items.Item("4").Specific, SAPbouiCOM.ComboBox).Value = "" Then
                oApplication.Utilities.Message("Select Bank to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            ElseIf CType(oForm.Items.Item("6").Specific, SAPbouiCOM.EditText).Value = "" Then
                oApplication.Utilities.Message("Select Bank Account to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            ElseIf CType(oForm.Items.Item("8").Specific, SAPbouiCOM.ComboBox).Value = "" Then
                oApplication.Utilities.Message("Select Commission Type to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            ElseIf CType(oForm.Items.Item("10").Specific, SAPbouiCOM.EditText).Value = "" Then
                oApplication.Utilities.Message("Select Commission Account to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            ElseIf CType(oForm.Items.Item("23").Specific, SAPbouiCOM.ComboBox).Value = "" Then
                oApplication.Utilities.Message("Select Currency to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            ElseIf CDbl(CType(oForm.Items.Item("16").Specific, SAPbouiCOM.EditText).Value) = 0 Then
                oApplication.Utilities.Message("Select Commission Charges to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return _retVal
    End Function

    Private Function UpdateValues(ByVal oForm As SAPbouiCOM.Form)
        Dim _retVal As Boolean = True
        Try
            Dim oUserTable As SAPbobsCOM.UserTable
            oUserTable = oApplication.Company.UserTables.Item("OCMR")
            oComboBox = CType(oForm.Items.Item("8").Specific, SAPbouiCOM.ComboBox)
            If oUserTable.GetByKey(oForm.Items.Item("18").Specific.value) Then
                oUserTable.UserFields.Fields.Item("U_BankCode").Value = CType(oForm.Items.Item("4").Specific, SAPbouiCOM.ComboBox).Value
                oUserTable.UserFields.Fields.Item("U_BankGL").Value = oForm.Items.Item("6").Specific.value
                oUserTable.UserFields.Fields.Item("U_CMType").Value = CType(oForm.Items.Item("8").Specific, SAPbouiCOM.ComboBox).Value
                oUserTable.UserFields.Fields.Item("U_CommGL").Value = oForm.Items.Item("10").Specific.value
                oUserTable.UserFields.Fields.Item("U_Currency").Value = CType(oForm.Items.Item("23").Specific, SAPbouiCOM.ComboBox).Value.Trim()
                oUserTable.UserFields.Fields.Item("U_CommCh").Value = oForm.Items.Item("16").Specific.value
                oUserTable.UserFields.Fields.Item("U_JourRem").Value = oForm.Items.Item("12").Specific.value
                If oUserTable.Update() <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return _retVal
    End Function

    Public Function removeCommission(ByVal oForm As SAPbouiCOM.Form)
        Dim _retVal As Boolean = True
        Try
            Dim oUserTable As SAPbobsCOM.UserTable
            oUserTable = oApplication.Company.UserTables.Item("OCMR")
            If oUserTable.GetByKey(oForm.Items.Item("18").Specific.value) Then
                If oUserTable.UserFields.Fields.Item("U_BankCode").Value.ToString().Trim = "" Then
                    If oUserTable.Remove() <> 0 Then

                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return _retVal
    End Function

    Private Sub addChooseFromList(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition

            
            oCFLs = oForm.ChooseFromLists

            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams

            'Adding Customer CFL for RM Account
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "1"
            oCFLCreationParams.UniqueID = "CFL_PR_1"

            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            oEditText = oForm.Items.Item("6").Specific
            oEditText.ChooseFromListUID = "CFL_PR_1"
            oEditText.ChooseFromListAlias = "FormatCode"

            'Adding Customer CFL for Fixed Labor Account
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "1"
            oCFLCreationParams.UniqueID = "CFL_PR_2"

            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()

            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            oEditText = oForm.Items.Item("10").Specific
            oEditText.ChooseFromListUID = "CFL_PR_2"
            oEditText.ChooseFromListAlias = "FormatCode"

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
