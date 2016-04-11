Public Class clsSpecialPrice
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
            oForm = oApplication.Utilities.LoadForm(xml_OSUS, frm_OSUS)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            initializeDataSource(oForm)
            initialize(oForm)
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            FillCombo(oForm)
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

    Public Sub LoadForm(ByVal strProjectCode As String, ByVal strProjectName As String, ByVal strCust As String, ByVal strName As String, ByVal strEffFrom As String, ByVal strEffTo As String)
        Try
            oForm = oApplication.Utilities.LoadForm(xml_OSUS, frm_OSUS)
            oForm = oApplication.SBO_Application.Forms.ActiveForm
            oForm.Freeze(True)
            initializeDataSource(oForm)
            initialize(oForm)
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            FillCombo(oForm)
            oForm.Items.Item("6").Specific.value = strProjectCode
            oForm.Items.Item("7").Specific.value = strProjectName
            oForm.Items.Item("19").Specific.value = strCust
            oForm.Items.Item("20").Specific.value = strName
            oForm.Items.Item("10").Specific.value = strEffFrom
            oForm.Items.Item("11").Specific.value = strEffTo
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

    Public Sub LoadForm(ByVal strDocEntry As String)
        Try
            oForm = oApplication.Utilities.LoadForm(xml_OSUS, frm_OSUS)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            initializeDataSource(oForm)
            initialize(oForm)
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            FillCombo(oForm)
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

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@OSUS")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@PSP1")

            oMatrix = oForm.Items.Item("3").Specific
            oMatrix.LoadFromDataSource()
            oMatrix.AddRow(1, -1)
            oMatrix.FlushToDataSource()

            clearDataSource(oForm)

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select IsNull(MAX(DocEntry),1) From [@OSUS]")
            If Not oRecordSet.EoF Then
                oDBDataSource.SetValue("DocNum", 0, oRecordSet.Fields.Item(0).Value.ToString())
            End If

            enableControl(oForm, True)
            'modeforControl(oForm)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub FillCombo(ByVal aForm As SAPbouiCOM.Form)
        Try
            Dim oTempRec As SAPbobsCOM.Recordset
            Dim oMatrix As SAPbouiCOM.Matrix
            oMatrix = aForm.Items.Item("3").Specific
            oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oColum As SAPbouiCOM.Column
            'Currency
            oColum = oMatrix.Columns.Item("V_2")
            For intRow As Integer = oColum.ValidValues.Count - 1 To 0 Step -1
                oColum.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oColum.ValidValues.Add("", "")
            oTempRec.DoQuery("Select CurrCode,CurrName From OCRN")
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                oColum.ValidValues.Add(oTempRec.Fields.Item("CurrCode").Value, oTempRec.Fields.Item("CurrName").Value)
                oTempRec.MoveNext()
            Next
            oColum.DisplayDesc = True

            'Price List
            oColum = oMatrix.Columns.Item("V_3")
            For intRow As Integer = oColum.ValidValues.Count - 1 To 0 Step -1
                oColum.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oColum.ValidValues.Add("", "")
            oTempRec.DoQuery("Select ListNum,ListName From OPLN Order By ListNum")
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                oColum.ValidValues.Add(oTempRec.Fields.Item("ListNum").Value, oTempRec.Fields.Item("ListName").Value)
                oTempRec.MoveNext()
            Next
            oColum.DisplayDesc = True
            oMatrix.AutoResizeColumns()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub initializeDataSource(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.DataSources.UserDataSources.Add("udsCust", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 40)
            oForm.DataSources.UserDataSources.Add("udsName", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 200)
            'oEditText = oForm.Items.Item("19").Specific
            'oEditText.DataBind.SetBound(True, "", "udsCust")
            'oEditText = oForm.Items.Item("20").Specific
            'oEditText.DataBind.SetBound(True, "", "udsName")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub clearDataSource(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Items.Item("19").Specific.value = ""
            oForm.Items.Item("20").Specific.value = ""
        Catch ex As Exception

        End Try
    End Sub

#Region "Methods"
    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("3").Specific
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@PSP1")
            oMatrix.FlushToDataSource()
            For count = 1 To oDBDataSourceLines.Size
                oDBDataSourceLines.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
#End Region

    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Dim oCombobox As SAPbouiCOM.ComboBox
            Select Case aForm.PaneLevel
                Case "0"
                    oMatrix = aForm.Items.Item("3").Specific
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@PSP1")
                    If oMatrix.RowCount <= 0 Then
                        oMatrix.AddRow()
                    Else
                        oMatrix.AddRow(1, oMatrix.RowCount + 1)
                        oMatrix.ClearRowData(oMatrix.RowCount)
                        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If
                    End If
                    oCombobox = oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific
                    Try
                        If oCombobox.Value <> "" Then
                            oMatrix.AddRow()
                            Select Case aForm.PaneLevel
                                Case "0"
                                    oCombobox = oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific
                                    oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, "0")
                            End Select
                        End If
                    Catch ex As Exception
                        aForm.Freeze(False)
                    End Try
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
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
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
            Throw ex
        End Try
    End Sub

#Region "Delete Row"
    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            oMatrix = aForm.Items.Item("3").Specific
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@OSUS")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@PSP1")
            If Me.MatrixId = "3" Then
                oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@PSP1")
            End If
            Me.RowtoDelete = intSelectedMatrixrow
            If Me.RowtoDelete > 0 Then
                oDBDataSourceLines.RemoveRecord(Me.RowtoDelete - 1)
            End If
            oMatrix.LoadFromDataSource()
            oMatrix.FlushToDataSource()
            For count = 1 To oDBDataSourceLines.Size - 1
                oDBDataSourceLines.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Validations"
    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oMatrix = oForm.Items.Item("3").Specific
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@OSUS")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@PSP1")

            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oApplication.Utilities.getEditTextvalue(aForm, "6") = "" Then
                oApplication.Utilities.Message("Enter Project Code...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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


            strQuery = "Select ValidFrom,ValidTo,U_CardCode,U_CardName From OPRJ Where PrjCode = '" + oApplication.Utilities.getEditTextvalue(aForm, "6") + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                dtValidFrom = oRecordSet.Fields.Item("ValidFrom").Value
                dtValidTo = oRecordSet.Fields.Item("ValidTo").Value
            End If

            If dtFromDt > dtToDt Then
                oApplication.Utilities.Message("Effective To Date Should be Greater than Effective From Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf (dtFromDt < dtValidFrom.ToString("yyyyMMdd") Or dtFromDt > dtValidTo.ToString("yyyyMMdd")) Then
                oApplication.Utilities.Message("Effective From Date Should Be Between Valid From & To Date of Project...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf (dtToDt < dtValidFrom.ToString("yyyyMMdd") Or dtToDt > dtValidTo.ToString("yyyyMMdd")) Then
                oApplication.Utilities.Message("Effective To Date Should Be Between Valid From & To Date of Project...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            oMatrix.FlushToDataSource()
            For index As Integer = 0 To oDBDataSourceLines.Size - 1
                If oDBDataSourceLines.GetValue("U_ItmCode", index).Trim() <> "" Then
                    If oDBDataSourceLines.GetValue("U_Currency", index).Trim() = "" Then
                        oApplication.Utilities.Message("Enter Currency Code for Item : " + oDBDataSourceLines.GetValue("U_ItmCode", index).Trim(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf CDbl(oDBDataSourceLines.GetValue("U_UnitPrice", index).Trim()) = 0 Then
                        oApplication.Utilities.Message("Enter UnitPrice for Item : " + oDBDataSourceLines.GetValue("U_ItmCode", index).Trim(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf oDBDataSourceLines.GetValue("U_DisType", index).Trim() = "" Then
                        oApplication.Utilities.Message("Enter Discount Type for Item : " + oDBDataSourceLines.GetValue("U_ItmCode", index).Trim(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf CDbl(oDBDataSourceLines.GetValue("U_Discount", index).Trim()) = 0 Then
                        oApplication.Utilities.Message("Enter Discount / Price for Item : " + oDBDataSourceLines.GetValue("U_ItmCode", index).Trim(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf oDBDataSourceLines.GetValue("U_DisType", index).Trim() = "D" And CInt(oDBDataSourceLines.GetValue("U_Discount", index).Trim()) > 100 Then
                        oApplication.Utilities.Message("Discount for Item : " + oDBDataSourceLines.GetValue("U_ItmCode", index).Trim() + " Should be Less Than Or Equal to 100 % ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            Next
            oMatrix.LoadFromDataSource()
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select 1 As 'Return',DocEntry From [@OSUS]"
            strQuery += " Where "
            strQuery += " (('" + oDBDataSource.GetValue("U_EffFrom", 0).ToString() + "' Between Convert(VarChar(12),U_EffFrom,112) And Convert(VarChar(12),U_EffTo,112)) "
            strQuery += " OR "
            strQuery += " ('" + oDBDataSource.GetValue("U_EffTo", 0).ToString() + "' Between Convert(VarChar(12),U_EffFrom,112) And Convert(VarChar(12),U_EffTo,112))) "
            strQuery += " And U_PrjCode = '" + oDBDataSource.GetValue("U_PrjCode", 0).Trim() + "' And DocEntry <> '" + oDBDataSource.GetValue("DocEntry", 0).ToString() + "'"
            oRecordSet.DoQuery(strQuery)

            If Not oRecordSet.EoF Then
                oApplication.Utilities.Message("Over Lapping of Dates already Exists...for Specified Project..", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_OSUS Then
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
                                If pVal.ItemUID = "3" Then
                                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@OSUS")
                                    Me.intSelectedMatrixrow = pVal.Row
                                    If (oDBDataSource.GetValue("U_PrjCode", 0).ToString() = "") Then
                                        oApplication.SBO_Application.SetStatusBarMessage("Select Project Code to Proceed...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        BubbleEvent = False
                                    ElseIf (oDBDataSource.GetValue("U_EffFrom", 0).ToString() = "" Or oDBDataSource.GetValue("U_EffTo", 0).ToString() = "") Then
                                        oApplication.SBO_Application.SetStatusBarMessage("Enter Effective From & To Date to Proceed...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        BubbleEvent = False
                                    End If
                                End If
                                '    Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                '        If pVal.ItemUID = "3" And (pVal.ColUID = "V_6") And pVal.Row > 0 Then
                                '            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                '            oMatrix = oForm.Items.Item("3").Specific
                                '            oMatrix.FlushToDataSource()
                                '            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@PSP1")
                                '            Dim dblDiscount As Double
                                '            Dim strType As String
                                '            strType = oDBDataSourceLines.GetValue("U_DisType", pVal.Row - 1)
                                '            dblDiscount = oDBDataSourceLines.GetValue("U_Discount", pVal.Row - 1)
                                '            If strType = "D" Then
                                '                If dblDiscount > 100 Then
                                '                    oApplication.SBO_Application.SetStatusBarMessage("Enter Discount Percentage Should be Less than Or Equal to 100...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                '                    BubbleEvent = False
                                '                End If
                                '            End If
                                '        End If
                                'End Select
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
                                    Case "1"
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            If pVal.Action_Success Then
                                                clearDataSource(oForm)
                                            End If
                                        End If
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@OSUS")
                                oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@PSP1")
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim strCode, strName, strCustomer, strCustName As String
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects

                                    If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If pVal.ItemUID = "6" Then
                                            strCode = oDataTable.GetValue("PrjCode", 0)
                                            strName = oDataTable.GetValue("PrjName", 0)
                                            strCustomer = oDataTable.GetValue("U_CardCode", 0)
                                            strCustName = oDataTable.GetValue("U_CardName", 0)
                                            Try
                                                oDBDataSource.SetValue("U_PrjCode", oDBDataSource.Offset, strCode)
                                                oDBDataSource.SetValue("U_PrjName", oDBDataSource.Offset, strName)
                                                oForm.DataSources.UserDataSources.Item("udsCust").ValueEx = strCustomer
                                                oForm.DataSources.UserDataSources.Item("udsName").ValueEx = strCustName
                                                dtValidFrom = oApplication.Utilities.GetDateTimeValue(oDataTable.GetValue("ValidFrom", 0))
                                                dtValidTo = oApplication.Utilities.GetDateTimeValue(oDataTable.GetValue("ValidTo", 0))
                                            Catch ex As Exception

                                            End Try
                                        ElseIf (pVal.ItemUID = "3" And (pVal.ColUID = "V_0" Or pVal.ColUID = "V_1")) Then
                                            oMatrix = oForm.Items.Item("3").Specific
                                            Dim strCurrency As String = String.Empty
                                            Dim strPricList As String = String.Empty
                                            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            strQuery = "Select Currency,ListNum From OCRD Where CardCode = '" + oForm.DataSources.UserDataSources.Item("udsCust").ValueEx + "'"
                                            oRecordSet.DoQuery(strQuery)
                                            If Not oRecordSet.EoF Then
                                                strCurrency = oRecordSet.Fields.Item("Currency").Value
                                                strPricList = oRecordSet.Fields.Item("ListNum").Value
                                            End If
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
                                                If strCurrency <> "##" Then
                                                    oDBDataSourceLines.SetValue("U_Currency", pVal.Row + index - 1, strCurrency)
                                                End If
                                                oDBDataSourceLines.SetValue("U_PriceList", pVal.Row + index - 1, strPricList)
                                                oDBDataSourceLines.SetValue("U_DisType", pVal.Row + index - 1, "D")
                                                Dim dblPrice As Double
                                                getPrice(oDataTable.GetValue("ItemCode", index), strPricList, strCurrency, dblPrice)
                                                oDBDataSourceLines.SetValue("U_UnitPrice", pVal.Row + index - 1, dblPrice)
                                                oDBDataSourceLines.SetValue("U_DisPrice", pVal.Row + index - 1, dblPrice)
                                            Next
                                            oMatrix.LoadFromDataSource()
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                    End If
                                Catch ex As Exception

                                End Try
                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And (pVal.ColUID = "V_5" Or pVal.ColUID = "V_6") And pVal.Row > 0 Then
                                    oForm.Freeze(True)
                                    calculatePriceAfterDis(oForm, pVal.Row)
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    oForm.Freeze(False)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And (pVal.ColUID = "V_2" Or pVal.ColUID = "V_3") And pVal.Row > 0 Then
                                    oMatrix = oForm.Items.Item("3").Specific
                                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@PSP1")
                                    oForm.Freeze(True)
                                    Dim dblPrice As Double
                                    oMatrix.FlushToDataSource()
                                    oMatrix.LoadFromDataSource()
                                    getPrice(oDBDataSourceLines.GetValue("U_ItmCode", pVal.Row - 1).Trim(), oDBDataSourceLines.GetValue("U_PriceList", pVal.Row - 1).Trim(), oDBDataSourceLines.GetValue("U_Currency", pVal.Row - 1).Trim(), dblPrice)
                                    oDBDataSourceLines.SetValue("U_UnitPrice", pVal.Row - 1, dblPrice)
                                    oDBDataSourceLines.SetValue("U_DisPrice", pVal.Row - 1, dblPrice)
                                    oMatrix.LoadFromDataSource()
                                    calculatePriceAfterDis(oForm, pVal.Row)
                                    oMatrix.LoadFromDataSource()
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    oForm.Freeze(False)
                                End If
                        End Select
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
                Case mnu_OSUS
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then

                    End If
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If oForm.TypeEx = frm_OSUS Then

                        oDBDataSource = oForm.DataSources.DBDataSources.Item("@OSUS")
                        If oDBDataSource.GetValue("Status", 0).Trim() = "C" Then
                            oApplication.Utilities.Message("Document Status Closed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            If pVal.BeforeAction = False Then
                                AddRow(oForm)
                            End If
                        End If
                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If oForm.TypeEx = frm_OSUS Then
                        oDBDataSource = oForm.DataSources.DBDataSources.Item("@OSUS")
                        If oDBDataSource.GetValue("Status", 0).Trim() = "C" Then
                            oApplication.Utilities.Message("Document Status Closed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            If pVal.BeforeAction = False Then
                                'RefereshDeleteRow(oForm)
                            End If
                        End If
                    End If
                Case mnu_ADD
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        If oForm.TypeEx = frm_OSUS Then
                            initialize(oForm)
                            oForm.Items.Item("7").Enabled = False
                            oForm.Items.Item("19").Enabled = False
                            oForm.Items.Item("20").Enabled = False
                        End If
                    End If
                Case mnu_FIND
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        If oForm.TypeEx = frm_OSUS Then
                            clearDataSource(oForm)
                            enableControl(oForm, True)
                            oForm.Items.Item("7").Enabled = True
                            oForm.Items.Item("19").Enabled = True
                            oForm.Items.Item("20").Enabled = True
                        End If
                    End If
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            If oForm.TypeEx = frm_OSUS Then
                Select Case BusinessObjectInfo.BeforeAction
                    Case True

                    Case False
                        Select Case BusinessObjectInfo.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                getDefaultValues(oForm.Items.Item("6").Specific.value)
                                If CType(oForm.Items.Item("24").Specific, SAPbouiCOM.ComboBox).Value = "C" Then
                                    enableControl(oForm, False)
                                Else
                                    enableControl(oForm, True)
                                End If
                                oForm.Items.Item("6").Enabled = False
                                oForm.Items.Item("7").Enabled = False
                                oForm.Items.Item("19").Enabled = False
                                oForm.Items.Item("20").Enabled = False
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub getDefaultValues(ByVal strProject As String)
        Try
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select ValidFrom,ValidTo,U_CardCode,U_CardName From OPRJ Where PrjCode = '" + strProject + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                dtValidFrom = oRecordSet.Fields.Item("ValidFrom").Value
                dtValidTo = oRecordSet.Fields.Item("ValidTo").Value
                oForm.Items.Item("19").Specific.value = oRecordSet.Fields.Item("U_CardCode").Value.ToString()
                oForm.Items.Item("20").Specific.value = oRecordSet.Fields.Item("U_CardName").Value.ToString()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getPrice(ByVal strItemCode As String, ByVal strPriceList As String, ByVal strCurrency As String, ByRef dblPrice As Double)
        Try
            Dim oExRecordSet As SAPbobsCOM.Recordset
            Dim dblRExRate, dblAExRate As Double
            Dim strItemCurrency As String
            oExRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select Currency,Price From ITM1 Where ItemCode = '" + strItemCode.Trim() + "' And PriceList = '" + strPriceList.Trim() + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                strItemCurrency = oRecordSet.Fields.Item("Currency").Value

                If strCurrency = oRecordSet.Fields.Item("Currency").Value Then
                    dblPrice = oRecordSet.Fields.Item("Price").Value
                Else
                    Dim dblLocalCurrency As String = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
                    If strCurrency <> dblLocalCurrency Then
                        oExRecordSet.DoQuery("Select Rate From ORTT Where Currency = '" + strCurrency + "' And Convert(VarChar(8),RateDate,112) = Convert(VarChar(8),GetDate(),112)")
                        If Not oExRecordSet.EoF Then
                            dblRExRate = oExRecordSet.Fields.Item("Rate").Value
                            If oRecordSet.Fields.Item("Currency").Value = dblLocalCurrency Then
                                dblPrice = oRecordSet.Fields.Item("Price").Value / dblRExRate
                            Else
                                oExRecordSet.DoQuery("Select isnull(Rate,1) 'Rate' From ORTT Where Currency = '" + strItemCurrency + "' And Convert(VarChar(8),RateDate,112) = Convert(VarChar(8),GetDate(),112)")
                                If Not oExRecordSet.EoF Then
                                    dblAExRate = oExRecordSet.Fields.Item("Rate").Value
                                    dblPrice = ((oRecordSet.Fields.Item("Price").Value * dblAExRate) / dblRExRate)
                                End If
                            End If
                        End If
                    Else
                        oExRecordSet.DoQuery("Select Rate From ORTT Where Currency = '" + oRecordSet.Fields.Item("Currency").Value + "' And Convert(VarChar(8),RateDate,112) = Convert(VarChar(8),GetDate(),112)")
                        If Not oExRecordSet.EoF Then
                            dblAExRate = oExRecordSet.Fields.Item("Rate").Value
                            dblPrice = (oRecordSet.Fields.Item("Price").Value * dblAExRate)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub calculatePriceAfterDis(ByVal oForm As SAPbouiCOM.Form, ByVal intRow As Integer)
        Try
            oMatrix = oForm.Items.Item("3").Specific
            oMatrix.FlushToDataSource()
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@PSP1")
            Dim dblPrice, dblDiscount, dblPriceAfterDis As Double
            Dim strType As String
            strType = oDBDataSourceLines.GetValue("U_DisType", intRow - 1)
            dblPrice = oDBDataSourceLines.GetValue("U_UnitPrice", intRow - 1)
            dblDiscount = oDBDataSourceLines.GetValue("U_Discount", intRow - 1)
            If strType = "D" Then
                dblPriceAfterDis = dblPrice - ((dblPrice * dblDiscount) / 100)
            Else
                dblPriceAfterDis = (dblPrice - dblDiscount)
            End If
            oDBDataSourceLines.SetValue("U_DisPrice", intRow - 1, dblPriceAfterDis)
            oMatrix.LoadFromDataSource()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub modeforControl(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 8, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            'oForm.Items.Item("6").Enabled = blnStatus
            'oForm.Items.Item("10").Enabled = blnStatus
            'oForm.Items.Item("11").Enabled = blnStatus
            'oForm.Items.Item("3").Enabled = blnStatus
            'oForm.Items.Item("14").Enabled = blnStatus
            'oForm.Items.Item("15").Enabled = blnStatus
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub enableControl(ByVal oForm As SAPbouiCOM.Form, ByVal blnStatus As Boolean)
        Try
            oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("6").Enabled = blnStatus
            oForm.Items.Item("10").Enabled = blnStatus
            oForm.Items.Item("11").Enabled = blnStatus
            oForm.Items.Item("3").Enabled = blnStatus
            oForm.Items.Item("14").Enabled = blnStatus
            oForm.Items.Item("15").Enabled = blnStatus
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Class
