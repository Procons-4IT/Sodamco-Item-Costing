Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsFreeItems
    Inherits clsBase
    Private oGrid As SAPbouiCOM.Grid
    Private oGridColumn As SAPbouiCOM.GridColumn
    Private oDtFreeItemList As SAPbouiCOM.DataTable
    Private strQuery As String
    Private oEditText As SAPbouiCOM.EditText
    Private oRecordSet As SAPbobsCOM.Recordset
    Private strCode As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub LoadForm(ByVal _strRFormID As String, ByVal _strReference As String, Optional ByVal _strRRowID As Integer = -1)
        Try
            oForm = oApplication.Utilities.LoadForm(xml_PRT2, frm_PRT2)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            oForm.Items.Item("4").Specific.value = _strReference
            initialize(oForm, _strReference)
            addChooseFromList(oForm)
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
            If pVal.FormTypeEx = frm_PRT2 Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    If Not validate(oForm) Then
                                        BubbleEvent = False
                                    Else
                                        If (UpdateValues(oForm)) Then

                                        End If
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                               If pVal.ItemUID = "5" Then
                                    oGrid = oForm.Items.Item("3").Specific
                                    If oGrid.DataTable.Rows.Count > 0 Then
                                        oGrid.DataTable.Rows.Add(1)
                                    End If
                                ElseIf pVal.ItemUID = "6" Then
                                    oGrid = oForm.Items.Item("3").Specific
                                    If oGrid.GetCellFocus().rowIndex > -1 Then
                                        oGrid.DataTable.Rows.Remove(oGrid.GetCellFocus().rowIndex)
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oGrid = oForm.Items.Item("3").Specific
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If pVal.ItemUID = "1" Or pVal.ColUID = "U_ItmCode" Then
                                        For index As Integer = 0 To oDataTable.Rows.Count - 1
                                            oGrid.DataTable.SetValue("U_ItmCode", pVal.Row, oDataTable.GetValue("ItemCode", index))
                                            oGrid.DataTable.SetValue("U_ItmDesc", pVal.Row, oDataTable.GetValue("ItemName", index))
                                            oGrid.DataTable.SetValue("U_Quantity", pVal.Row, "1")
                                        Next
                                    End If
                                Catch ex As Exception
                                    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try
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
            oGrid = oForm.Items.Item("3").Specific
            oForm.DataSources.DataTables.Add("FreeItemList")
            oDtFreeItemList = oForm.DataSources.DataTables.Item("FreeItemList")
            Dim strQuery As String = "Select Code,U_ItmCode,U_ItmDesc,U_Quantity,U_Discount,U_Reference From [@PRT2] Where U_Reference = '" & strRefCode & "'"
            oDtFreeItemList.ExecuteQuery(strQuery)
            oGrid.DataTable = oDtFreeItemList
            oGrid.Columns.Item("Code").TitleObject.Caption = "Code"
            oGrid.Columns.Item("Code").Visible = False
            oGrid.Columns.Item("U_ItmCode").TitleObject.Caption = "Free Item"
            oGrid.Columns.Item("U_ItmDesc").TitleObject.Caption = "Free Description"
            oGrid.Columns.Item("U_ItmDesc").Editable = False
            oGrid.Columns.Item("U_Quantity").TitleObject.Caption = "Quantity"
            oGrid.Columns.Item("U_Quantity").RightJustified = True
            oGrid.Columns.Item("U_Discount").TitleObject.Caption = "Discount %"
            oGrid.Columns.Item("U_Discount").RightJustified = True
            oGrid.Columns.Item("U_Reference").TitleObject.Caption = "Reference"
            oGrid.Columns.Item("U_Reference").Visible = False
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function validate(ByVal oForm As SAPbouiCOM.Form)
        Dim _retVal As Boolean = True
        Try
            
        Catch ex As Exception
            Throw ex
        End Try
        Return _retVal
    End Function

    Private Function UpdateValues(ByVal oForm As SAPbouiCOM.Form)
        Dim _retVal As Boolean = True
        Try
            oGrid = oForm.Items.Item("3").Specific
            Dim oUserTable As SAPbobsCOM.UserTable
            oUserTable = oApplication.Company.UserTables.Item("PRT2")
            oApplication.Company.StartTransaction()

            'Remove Existing Free Items
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Code From [@PRT2] Where U_Reference = '" + oApplication.Utilities.getEditTextvalue(oForm, "4") + "'")
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    If oUserTable.GetByKey(oRecordSet.Fields.Item("Code").Value) Then
                        If oUserTable.Update() <> 0 Then
                            _retVal = False
                            Exit While
                        End If
                    End If
                    oRecordSet.MoveNext()
                End While
            End If


            If _retVal Then
                For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    If oGrid.DataTable.GetValue("Code", index) <> "" Then
                        If oUserTable.GetByKey(oGrid.DataTable.GetValue("Code", index)) Then
                            oUserTable.UserFields.Fields.Item("U_Reference").Value = oApplication.Utilities.getEditTextvalue(oForm, "4")
                            oUserTable.UserFields.Fields.Item("U_ItmCode").Value = oGrid.DataTable.GetValue("U_ItmCode", index)
                            oUserTable.UserFields.Fields.Item("U_ItmDesc").Value = oGrid.DataTable.GetValue("U_ItmDesc", index)
                            oUserTable.UserFields.Fields.Item("U_Quantity").Value = oGrid.DataTable.GetValue("U_Quantity", index)
                            oUserTable.UserFields.Fields.Item("U_Discount").Value = oGrid.DataTable.GetValue("U_Discount", index)
                            If oUserTable.Update() <> 0 Then
                                _retVal = False
                                Exit For
                            End If
                        End If
                    Else
                        If oGrid.DataTable.GetValue("U_ItmCode", index) <> "" Then
                            Dim intCode As Integer = oApplication.Utilities.getMaxCode("@PRT2", "Code")
                            oUserTable.Code = intCode.ToString()
                            oUserTable.Name = intCode.ToString()
                            oUserTable.UserFields.Fields.Item("U_Reference").Value = oApplication.Utilities.getEditTextvalue(oForm, "4")
                            oUserTable.UserFields.Fields.Item("U_ItmCode").Value = oGrid.DataTable.GetValue("U_ItmCode", index)
                            oUserTable.UserFields.Fields.Item("U_ItmDesc").Value = oGrid.DataTable.GetValue("U_ItmDesc", index)
                            oUserTable.UserFields.Fields.Item("U_Quantity").Value = oGrid.DataTable.GetValue("U_Quantity", index)
                            oUserTable.UserFields.Fields.Item("U_Discount").Value = oGrid.DataTable.GetValue("U_Discount", index)
                            If oUserTable.Add() <> 0 Then
                                _retVal = False
                                Exit For
                            End If
                        End If
                    End If
                Next
            End If

            If _retVal Then
                If (oApplication.Company.InTransaction) Then
                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                End If
            Else
                If (oApplication.Company.InTransaction) Then
                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
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
            oCFLCreationParams.MultiSelection = True
            oCFLCreationParams.ObjectType = "4"
            oCFLCreationParams.UniqueID = "CFL_PR_1"

            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "InvntItem"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            oGrid = oForm.Items.Item("3").Specific
            oGridColumn = oGrid.Columns.Item("U_ItmCode")
            oGridColumn.ChooseFromListUID = "CFL_PR_1"
            oGridColumn.ChooseFromListAlias = "ItemCode"
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
