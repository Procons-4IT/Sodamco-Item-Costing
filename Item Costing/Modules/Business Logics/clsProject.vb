Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsProject
    Inherits clsBase
    Private oMatrix As SAPbouiCOM.Matrix
    Private oRecordSet As SAPbobsCOM.Recordset
    Private strQuery As String

    Public Sub New()
        MyBase.New()
    End Sub

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Select Case pVal.MenuUID
            Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            Case mnu_ADD
        End Select
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Project Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                'If pVal.ItemUID = "3" And pVal.ColUID = "U_Currency" And pVal.Row > 0 Then
                                '    Dim oBP As SAPbobsCOM.BusinessPartners
                                '    oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                                '    If oBP.GetByKey(oMatrix.Columns.Item("U_CardCode").Cells.Item(pVal.Row).Specific.value) Then
                                '        If oBP.Currency <> "##" And oMatrix.Columns.Item("U_Currency").Cells.Item(pVal.Row).Specific.value <> "" Then
                                '            'oApplication.Utilities.Message("Option for only Multiple Currency Customer...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                '            BubbleEvent = False
                                '            Exit Sub
                                '        End If
                                '    End If
                                'End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "1" Then
                                    If Not validate(oForm) Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                If pVal.ItemUID = "3" And (pVal.ColUID = "U_CardCode" Or pVal.ColUID = "U_CardName") And pVal.Row > 0 Then
                                    CType(oForm.Items.Item("_101").Specific, SAPbouiCOM.EditText).Value = oMatrix.Columns.Item("U_CardCode").Cells.Item(pVal.Row).Specific.value
                                    oForm.Items.Item("_102").Visible = True
                                    oForm.Items.Item("_102").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oForm.Items.Item("_102").Visible = False
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                initialize(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim strCustomer, strName, strCurrency As String
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If pVal.ItemUID = "3" And (pVal.ColUID = "U_CardCode" Or pVal.ColUID = "U_CardName") And Not IsNothing(oDataTable) Then
                                        strCustomer = oDataTable.GetValue("CardCode", 0)
                                        strName = oDataTable.GetValue("CardName", 0)
                                        strCurrency = oDataTable.GetValue("Currency", 0)
                                        Try
                                            oMatrix = oForm.Items.Item("3").Specific
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_CardCode", pVal.Row, strCustomer)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_CardName", pVal.Row, strName)
                                            If strCurrency <> "##" Then
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "U_Currency", pVal.Row, strCurrency)
                                            Else
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "U_Currency", pVal.Row, oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency)
                                            End If
                                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        Catch ex As Exception
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_CardCode", pVal.Row, strCustomer)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_CardName", pVal.Row, strName)
                                            If strCurrency <> "##" Then
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "U_Currency", pVal.Row, strCurrency)
                                            Else
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "U_Currency", pVal.Row, oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency)
                                            End If
                                        End Try
                                    ElseIf (pVal.ItemUID = "3" And (pVal.ColUID = "U_Currency")) And Not IsNothing(oDataTable) Then
                                        strCurrency = oDataTable.GetValue("CurrCode", 0)
                                        oMatrix = oForm.Items.Item("3").Specific
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Currency", pVal.Row, strCurrency)
                                    End If
                                Catch ex As Exception

                                End Try
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                Try
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    If pVal.ItemUID = "_100" Then
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            oMatrix = oForm.Items.Item("3").Specific
                                            Dim intRowSelected As Integer
                                            Dim strProject, strProjName, strCust, strCustNam, strValFrm, strValTo As String
                                            For index As Integer = 1 To oMatrix.RowCount
                                                If oMatrix.IsRowSelected(index) Then
                                                    intRowSelected = index
                                                End If
                                            Next
                                            strProject = oMatrix.Columns.Item("PrjCode").Cells.Item(intRowSelected).Specific.value
                                            strProjName = oMatrix.Columns.Item("PrjName").Cells.Item(intRowSelected).Specific.value
                                            strCust = oMatrix.Columns.Item("U_CardCode").Cells.Item(intRowSelected).Specific.value
                                            strCustNam = oMatrix.Columns.Item("U_CardName").Cells.Item(intRowSelected).Specific.value
                                            strValFrm = oMatrix.Columns.Item("ValidFrom").Cells.Item(intRowSelected).Specific.value
                                            strValTo = oMatrix.Columns.Item("ValidTo").Cells.Item(intRowSelected).Specific.value

                                            strQuery = "Select DocEntry From [@OPSP] Where U_PrjCode = '" + strProject + "'"
                                            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRecordSet.DoQuery(strQuery)
                                            If Not oRecordSet.EoF Then
                                                While Not oRecordSet.EoF
                                                    Dim objSpecialPrice As clsSpecialPrice
                                                    objSpecialPrice = New clsSpecialPrice()
                                                    objSpecialPrice.LoadForm(oRecordSet.Fields.Item("DocEntry").Value.ToString())
                                                    oRecordSet.MoveNext()
                                                End While
                                            Else
                                                Dim objSpecialPrice As clsSpecialPrice
                                                objSpecialPrice = New clsSpecialPrice()
                                                objSpecialPrice.LoadForm(strProject, strProjName, strCust, strCustNam, strValFrm, strValTo)
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception

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

#Region "Data Events"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            Select Case BusinessObjectInfo.BeforeAction
                Case True

                Case False
                    Select Case BusinessObjectInfo.EventType
                    End Select
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Function"
    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            addChooseFromList(oForm)
            addControls(oForm)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub addChooseFromList(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCustCol As SAPbouiCOM.Column
            Dim oNameCol As SAPbouiCOM.Column
            Dim oCurrCol As SAPbouiCOM.Column

            oMatrix = oForm.Items.Item("3").Specific
            oCustCol = oMatrix.Columns.Item("U_CardCode")
            oNameCol = oMatrix.Columns.Item("U_CardName")
            oCurrCol = oMatrix.Columns.Item("U_Currency")

            oCFLs = oForm.ChooseFromLists

            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams

            'Adding Customer CFL for Customer Column
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFL_PR_1"

            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"

            oCustCol.ChooseFromListUID = "CFL_PR_1"
            oCustCol.ChooseFromListAlias = "CardCode"

            'Adding Customer CFL for Customer Column
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFL_PR_2"

            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"

            oNameCol.ChooseFromListUID = "CFL_PR_2"
            oNameCol.ChooseFromListAlias = "CardName"

            'Adding Customer CFL for Customer Column
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "37"
            oCFLCreationParams.UniqueID = "CFL_PR_3"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCurrCol.ChooseFromListUID = "CFL_PR_3"
            oCurrCol.ChooseFromListAlias = "CurrCode"

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub addControls(ByVal oForm As SAPbouiCOM.Form)
        Try
            oApplication.Utilities.AddControls(oForm, "_100", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", 1, 1, "2", "Special Price", 0, 0, 0, True)
            oForm.Items.Item("_100").Visible = True
            oApplication.Utilities.AddControls(oForm, "_101", "_100", SAPbouiCOM.BoFormItemTypes.it_EDIT, "RIGHT", 1, 1, "_100", "", 1, 0, 1, False)
            oForm.DataSources.UserDataSources.Add("udsCCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            CType(oForm.Items.Item("_101").Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "", "udsCCode")
            oApplication.Utilities.AddControls(oForm, "_102", "_101", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON, "RIGHT", 1, 1, "_101", "", 1, 0, 1, False)
            CType(oForm.Items.Item("_102").Specific, SAPbouiCOM.LinkedButton).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner

            oForm.Items.Item("_101").AffectsFormMode = False
            oForm.Items.Item("_101").Left = oForm.Items.Item("_100").Left + oForm.Items.Item("_100").Width + 10
            oForm.Items.Item("_101").Width = 1
            oForm.Items.Item("_101").Height = 1
            oForm.Items.Item("_101").Visible = True
            oForm.Items.Item("_102").Width = 1
            oForm.Items.Item("_102").Height = 1
            oForm.Items.Item("_102").Visible = False
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function validate(ByVal oForm As SAPbouiCOM.Form)
        Dim _retVal As Boolean = True
        Try
            oMatrix = oForm.Items.Item("3").Specific
            For index As Integer = 1 To oMatrix.RowCount
                Dim oBP As SAPbobsCOM.BusinessPartners
                oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                Dim strCustomer As String = oMatrix.Columns.Item("U_CardCode").Cells.Item(index).Specific.value
                If oBP.GetByKey(oMatrix.Columns.Item("U_CardCode").Cells.Item(index).Specific.value) Then
                    If oMatrix.Columns.Item("U_Currency").Cells.Item(index).Specific.value <> "" Then
                        If oBP.Currency <> "##" Then
                            If oMatrix.Columns.Item("U_Currency").Cells.Item(index).Specific.value <> oBP.Currency Then
                                oApplication.Utilities.Message("Default Currency not matching for ...:" + strCustomer, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                _retVal = False
                                Exit For
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

#End Region

End Class
