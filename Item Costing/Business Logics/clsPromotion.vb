Public Class clsPromotion
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private InvForConsumedItems, count As Integer
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines As SAPbouiCOM.DBDataSource
    Public intSelectedMatrixrow As Integer = 0
    Private RowtoDelete As Integer
    Private MatrixId As String
    Private oRecordSet As SAPbobsCOM.Recordset
    Private dtValidFrom, dtValidTo As Date
    Private strQuery As String

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_OPRM, frm_OPRM)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            initialize(oForm)
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

    Public Sub LoadForm(ByVal strDocEntry As String)
        Try
            oForm = oApplication.Utilities.LoadForm(xml_OPRM, frm_OPRM)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            initialize(oForm)
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            oForm.Freeze(False)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            oForm.Items.Item("16").Specific.value = strDocEntry
            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_OPRM Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        Else
                                            If validation(oForm) = False Then
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@OPRM")
                                If pVal.ItemUID = "3" Then
                                    intSelectedMatrixrow = pVal.Row
                                    If (oDBDataSource.GetValue("U_PrCode", 0).ToString() = "") Then
                                        BubbleEvent = False
                                        oApplication.SBO_Application.SetStatusBarMessage("Select Promotion Code to Proceed...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    ElseIf (oDBDataSource.GetValue("U_EffFrom", 0).ToString() = "" Or oDBDataSource.GetValue("U_EffTo", 0).ToString() = "") Then
                                        BubbleEvent = False
                                        oApplication.SBO_Application.SetStatusBarMessage("Enter Effective From & To Date to Proceed...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "14"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_ADD_ROW)
                                    Case "15"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@OPRM")
                                oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@PRM1")
                                oMatrix = oForm.Items.Item("3").Specific
                                oMatrix.FlushToDataSource()
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim strCode, strName, strCustomer, strCustName As String
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects

                                    If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If pVal.ItemUID = "6" Then
                                            strCode = oDataTable.GetValue("PrCode", 0)
                                            strName = oDataTable.GetValue("PrName", 0)
                                            strCustomer = oDataTable.GetValue("U_CardCode", 0)
                                            strCustName = oDataTable.GetValue("U_CardName", 0)
                                            Try
                                                oDBDataSource.SetValue("U_PrCode", oDBDataSource.Offset, strCode)
                                                oDBDataSource.SetValue("U_PrName", oDBDataSource.Offset, strName)
                                            Catch ex As Exception

                                            End Try
                                        ElseIf (pVal.ItemUID = "3" And (pVal.ColUID = "V_0" Or pVal.ColUID = "V_1")) Then
                                            oMatrix = oForm.Items.Item("3").Specific
                                            oMatrix.LoadFromDataSource()
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 1 Then
                                                intAddRows -= 1
                                                oMatrix.AddRow(intAddRows, pVal.Row - 1)
                                            End If
                                            oMatrix.FlushToDataSource()
                                            For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                oDBDataSourceLines.SetValue("LineId", pVal.Row + index - 1, (pVal.Row + index).ToString())
                                                oDBDataSourceLines.SetValue("U_ItmCode", pVal.Row + index - 1, oDataTable.GetValue("ItemCode", index))
                                                oDBDataSourceLines.SetValue("U_ItmName", pVal.Row + index - 1, oDataTable.GetValue("ItemName", index))
                                                oDBDataSourceLines.SetValue("U_Qty", pVal.Row + index - 1, "1")
                                            Next
                                            oMatrix.LoadFromDataSource()
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf (pVal.ItemUID = "3" And pVal.ColUID = "V_3") Then
                                            oMatrix = oForm.Items.Item("3").Specific
                                            oMatrix.LoadFromDataSource()
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 1 Then
                                                intAddRows -= 1
                                                oMatrix.AddRow(intAddRows, pVal.Row - 1)
                                            End If
                                            oMatrix.FlushToDataSource()
                                            For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                oDBDataSourceLines.SetValue("LineId", pVal.Row + index - 1, (pVal.Row + index).ToString())
                                                oDBDataSourceLines.SetValue("U_OffCode", pVal.Row + index - 1, oDataTable.GetValue("ItemCode", index))
                                                oDBDataSourceLines.SetValue("U_OffName", pVal.Row + index - 1, oDataTable.GetValue("ItemName", index))
                                            Next
                                            oMatrix.LoadFromDataSource()
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                    End If
                                Catch ex As Exception

                                End Try
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_OPRM
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then

                    End If
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        AddRow(oForm)
                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        RefereshDeleteRow(oForm)
                    End If
                Case mnu_ADD
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        initialize(oForm)
                        EnableControls(oForm, True)
                    End If
                Case mnu_FIND
                    If pVal.BeforeAction = False Then

                    End If
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Data Events"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            If oForm.TypeEx = frm_OPRM Then
                Select Case BusinessObjectInfo.BeforeAction
                    Case True

                    Case False
                        Select Case BusinessObjectInfo.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                EnableControls(oForm, False)
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@OPRM")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@PRM1")

            oMatrix = oForm.Items.Item("3").Specific
            oMatrix.LoadFromDataSource()
            oMatrix.AddRow(1, -1)
            oMatrix.FlushToDataSource()

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select IsNull(MAX(DocEntry),1) From [@OPRM]")
            If Not oRecordSet.EoF Then
                oDBDataSource.SetValue("DocNum", 0, oRecordSet.Fields.Item(0).Value.ToString())
            End If
            oDBDataSource.SetValue("U_Active", 0, "Y")
            MatrixId = "3"
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#Region "Methods"
    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("3").Specific
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@PRM1")
            oMatrix.FlushToDataSource()
            For count = 1 To oDBDataSourceLines.Size
                oDBDataSourceLines.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub
#End Region

    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Select Case aForm.PaneLevel
                Case "0"
                    oMatrix = aForm.Items.Item("3").Specific
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@PRM1")
                    If oMatrix.RowCount <= 0 Then
                        oMatrix.AddRow()
                    Else
                        oMatrix.AddRow(1, oMatrix.RowCount + 1)
                        oMatrix.ClearRowData(oMatrix.RowCount)
                        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If
                    End If
                    oMatrix.FlushToDataSource()
                    For count = 1 To oDBDataSourceLines.Size
                        oDBDataSourceLines.SetValue("LineId", count - 1, count)
                    Next
                    oMatrix.LoadFromDataSource()
                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    AssignLineNo(aForm)
            End Select
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub deleterow(ByVal aForm As SAPbouiCOM.Form)
        Try
            Select Case aForm.PaneLevel
                Case "0"
                    oMatrix = aForm.Items.Item("11").Specific
                    oDBDataSourceLines = aForm.DataSources.DBDataSources.Item("@PSP1")
            End Select
            oMatrix.FlushToDataSource()
            For introw As Integer = 1 To oMatrix.RowCount
                If oMatrix.IsRowSelected(introw) Then
                    oMatrix.DeleteRow(introw)
                    oDBDataSourceLines.RemoveRecord(introw - 1)
                    oMatrix.FlushToDataSource()
                    For count As Integer = 1 To oDBDataSourceLines.Size
                        oDBDataSourceLines.SetValue("LineId", count - 1, count)
                    Next
                    Select Case aForm.PaneLevel
                        Case "0"
                            oMatrix = aForm.Items.Item("3").Specific
                            oDBDataSourceLines = aForm.DataSources.DBDataSources.Item("@PSP1")
                            AssignLineNo(aForm)
                    End Select
                    oMatrix.LoadFromDataSource()
                    If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

#Region "Delete Row"
    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            oMatrix = aForm.Items.Item("3").Specific
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@OPRM")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@PRM1")

            If Me.MatrixId = "3" Then
                oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@PRM1")
            End If

            Me.RowtoDelete = intSelectedMatrixrow
            oDBDataSourceLines.RemoveRecord(Me.RowtoDelete - 1)
            oMatrix.LoadFromDataSource()
            oMatrix.FlushToDataSource()
            For count = 1 To oDBDataSourceLines.Size - 1
                oDBDataSourceLines.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub
#End Region

#Region "Validations"
    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            oMatrix = oForm.Items.Item("3").Specific
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@OPRM")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@PRM1")

            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oApplication.Utilities.getEditTextvalue(aForm, "6") = "" Then
                oApplication.Utilities.Message("Enter Promotion Code...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "7") = "" Then
                oApplication.Utilities.Message("Enter Promotion Name...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "10") = "" Then
                oApplication.Utilities.Message("Enter Effective From Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "11") = "" Then
                oApplication.Utilities.Message("Enter Effective To Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Dim dtFromDt As Integer
            Dim dtToDt As Integer
            dtFromDt = oApplication.Utilities.getEditTextvalue(aForm, "10")
            dtToDt = oApplication.Utilities.getEditTextvalue(aForm, "11")
            If dtFromDt > dtToDt Then
                oApplication.Utilities.Message("Effective To Date Should be Greater than Effective From Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select 1 As 'Return',DocEntry From [@OPRM]"
            strQuery += " Where "
            strQuery += " U_PrCode = '" + oDBDataSource.GetValue("U_PrCode", 0).Trim() + "' And DocEntry <> '" + oDBDataSource.GetValue("DocEntry", 0).ToString() + "'"
            oRecordSet.DoQuery(strQuery)

            If Not oRecordSet.EoF Then
                oApplication.Utilities.Message("Promotion Code Already Exist...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Function
#End Region

#Region "Disable Controls"

    Private Sub EnableControls(ByVal oForm As SAPbouiCOM.Form, ByVal blnEnable As Boolean)
        Try
            oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("6").Enabled = blnEnable
            oForm.Items.Item("7").Enabled = blnEnable
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

#End Region

End Class
