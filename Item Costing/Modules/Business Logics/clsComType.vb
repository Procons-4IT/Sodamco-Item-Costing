Public Class clsComType
    Inherits clsBase
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oMatrix As SAPbouiCOM.Matrix
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private oTemp As SAPbobsCOM.Recordset
    Private ocombo As SAPbouiCOM.ComboBoxColumn
    Private oEditText As SAPbouiCOM.EditText
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_ComType, frm_ComType)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            AddChooseFromList(oForm)
            Databind(oForm)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

#Region "Databind"
    Private Sub Databind(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            oGrid = aform.Items.Item("5").Specific
            dtTemp = oGrid.DataTable
            dtTemp.ExecuteQuery("Select *,Code 'Ref' from [@OCMT] order by Code")
            oGrid.DataTable = dtTemp
            Formatgrid(oGrid)
            oApplication.Utilities.assignMatrixLineno(oGrid, aform)
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub
#End Region

#Region "FormatGrid"
    Private Sub Formatgrid(ByVal agrid As SAPbouiCOM.Grid)
        Try
            agrid.Columns.Item(0).TitleObject.Caption = "Commission Code"
            agrid.Columns.Item(1).TitleObject.Caption = "Commission Name"
            agrid.Columns.Item(2).TitleObject.Caption = "Document Type"
            agrid.Columns.Item(3).TitleObject.Caption = "Commission Account"

            agrid.Columns.Item(4).Visible = False

            agrid.Columns.Item(2).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            ocombo = agrid.Columns.Item(2)
            ocombo.ValidValues.Add("P", "Payment")
            ocombo.ValidValues.Add("D", "Deposit")
            ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both

            agrid.Columns.Item(3).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditTextColumn = agrid.Columns.Item(3)
            oEditTextColumn.ChooseFromListUID = "CFL1"
            oEditTextColumn.ChooseFromListAlias = "FormatCode"
            oEditTextColumn.LinkedObjectType = "1"

            agrid.AutoResizeColumns()
            agrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "AddRow"
    Private Sub AddEmptyRow(ByVal aGrid As SAPbouiCOM.Grid)
        Try
            If aGrid.DataTable.GetValue("Code", aGrid.DataTable.Rows.Count - 1) <> "" Then
                aGrid.DataTable.Rows.Add()
                aGrid.Columns.Item(0).Click(aGrid.DataTable.Rows.Count - 1, False)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "CommitTrans"
    Private Sub Committrans(ByVal strChoice As String)
        Try
            Dim oTemprec, oItemRec As SAPbobsCOM.Recordset
            oTemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oItemRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If strChoice = "Cancel" Then
                oTemprec.DoQuery("Update [@OCMT] set Name = Code where Name Like '%DX'")
            Else
                oTemprec.DoQuery("Delete from  [@OCMT]  where Name Like '%DX'")
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "AddtoUDT"

    Private Function AddtoUDT1(ByVal aform As SAPbouiCOM.Form) As Boolean
        Try
            Dim oUserTable As SAPbobsCOM.UserTable
            Dim strCode, strECode, strDocType, strAcctNo As String
            oGrid = aform.Items.Item("5").Specific
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue(0, intRow) <> "" Or oGrid.DataTable.GetValue(1, intRow) <> "" Then
                    strCode = oGrid.DataTable.GetValue(0, intRow)
                    strECode = oGrid.DataTable.GetValue(1, intRow)
                    strDocType = oGrid.DataTable.GetValue(2, intRow)
                    strAcctNo = oGrid.DataTable.GetValue(3, intRow)
                    oUserTable = oApplication.Company.UserTables.Item("OCMT")
                    If oUserTable.GetByKey(strCode) = False Then
                        oUserTable.Code = strCode
                        oUserTable.Name = strECode
                        oUserTable.UserFields.Fields.Item("U_DocType").Value = strDocType
                        oUserTable.UserFields.Fields.Item("U_CMAcct").Value = strAcctNo
                        If oUserTable.Add() <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Committrans("Cancel")
                            Return False
                        End If
                    Else
                        strCode = oGrid.DataTable.GetValue(0, intRow)
                        If oUserTable.GetByKey(strCode) Then
                            oUserTable.Code = strCode
                            oUserTable.Name = strECode
                            oUserTable.UserFields.Fields.Item("U_DocType").Value = strDocType
                            oUserTable.UserFields.Fields.Item("U_CMAcct").Value = strAcctNo
                            If oUserTable.Update() <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Committrans("Cancel")
                                Return False
                            End If
                        End If
                    End If
                End If
            Next
            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Committrans("Add")
            Databind(aform)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region

#Region "Remove Row"
    Private Sub RemoveRow(ByVal intRow As Integer, ByVal agrid As SAPbouiCOM.Grid)
        Try
            Dim strCode As String
            Dim otemprec As SAPbobsCOM.Recordset
            For intRow = 0 To agrid.DataTable.Rows.Count - 1
                If agrid.Rows.IsSelected(intRow) Then
                    strCode = agrid.DataTable.GetValue(0, intRow)
                    otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oApplication.Utilities.ExecuteSQL(oTemp, "update [@OCMT] set  NAME = NAME +'DX'  where CODE='" & strCode & "'")
                    agrid.DataTable.Rows.Remove(intRow)
                    Exit Sub
                End If
            Next
            oApplication.Utilities.Message("No row selected", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Validate Grid details"
    Private Function validation(ByVal aGrid As SAPbouiCOM.Grid) As Boolean
        Try
            Dim strECode, strECode1, strEname, strEname1 As String
            For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
                strECode = aGrid.DataTable.GetValue(0, intRow)
                strEname = aGrid.DataTable.GetValue(1, intRow)
                For intInnerLoop As Integer = intRow To aGrid.DataTable.Rows.Count - 1
                    strECode1 = aGrid.DataTable.GetValue(0, intInnerLoop)
                    strEname1 = aGrid.DataTable.GetValue(1, intInnerLoop)
                    If strECode1 <> "" And strEname1 = "" Then
                        oApplication.Utilities.Message("Name can not be empty ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                    If strECode1 = "" And strEname1 <> "" Then
                        oApplication.Utilities.Message("Code can not be empty ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                    If strECode = strECode1 And intRow <> intInnerLoop Then
                        oApplication.Utilities.Message("This entry already exists. Code no : " & strECode, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        aGrid.Columns.Item(0).Click(intInnerLoop, , 1)
                        Return False
                    End If
                Next
            Next
        Catch ex As Exception
            Throw ex
        End Try
        Return True
    End Function
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_ComType Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "2" Then
                                    Committrans("Cancel")
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oGrid = oForm.Items.Item("5").Specific
                                If pVal.ItemUID = "5" And pVal.ColUID = "Code" And pVal.CharPressed <> 9 Then
                                    If oGrid.DataTable.GetValue("Ref", pVal.Row) <> "" Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "13" Then
                                    oGrid = oForm.Items.Item("5").Specific
                                    If validation(oGrid) = True Then
                                        AddtoUDT1(oForm)
                                    End If
                                End If
                                If pVal.ItemUID = "3" Then
                                    oGrid = oForm.Items.Item("5").Specific
                                    AddEmptyRow(oGrid)
                                End If
                                If pVal.ItemUID = "4" Then
                                    oGrid = oForm.Items.Item("5").Specific
                                    RemoveRow(pVal.Row, oGrid)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim strAccount As String
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If pVal.ItemUID = "5" And (pVal.ColUID = "U_CMAcct") Then
                                        strAccount = oDataTable.GetValue("FormatCode", 0)
                                        Try
                                            oGrid = oForm.Items.Item("5").Specific
                                            oGrid.DataTable.SetValue("U_CMAcct", pVal.Row, strAccount)
                                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        Catch ex As Exception
                                            oGrid.DataTable.SetValue("U_CMAcct", pVal.Row, strAccount)
                                        End Try
                                    End If
                                Catch ex As Exception

                                End Try
                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_ComType
                    LoadForm()
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    oGrid = oForm.Items.Item("5").Specific
                    If pVal.BeforeAction = False Then
                        AddEmptyRow(oGrid)
                        oApplication.Utilities.assignMatrixLineno(oGrid, oForm)
                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    oGrid = oForm.Items.Item("5").Specific
                    If pVal.BeforeAction = True Then
                        RemoveRow(1, oGrid)
                        oApplication.Utilities.assignMatrixLineno(oGrid, oForm)
                        BubbleEvent = False
                        Exit Sub
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                   
                End Select
            End If
        Catch ex As Exception
        End Try
    End Sub

#Region "Add Choose From List"
    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFL = oCFLs.Item("CFL1")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

End Class
