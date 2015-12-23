Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsDocumentFreight
    Inherits clsBase
    Private oMatrix As SAPbouiCOM.Matrix
    Private oRecordSet As SAPbobsCOM.Recordset
    Private strQuery As String
    Private blnIsLoad As Boolean = False

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
            If pVal.FormTypeEx = frm_DocumentFreight Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_CLICK, SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("3").Specific
                                If pVal.ItemUID = "3" And (pVal.ColUID = "U_Currency" Or pVal.ColUID = "U_FCalcu" Or pVal.ColUID = "U_DAmount") And Not blnIsLoad Then
                                    BubbleEvent = False
                                    Exit Sub
                                ElseIf (pVal.ItemUID = "3" And pVal.ColUID = "3" And pVal.Row > 0) And Not blnIsLoad Then
                                    If CDbl(CType(oMatrix.Columns.Item("U_DAmount").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value) <> 0 Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                If pVal.ItemUID = "3" And (pVal.ColUID = "U_PDiscount") And pVal.Row > 0 And Not blnIsLoad Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oMatrix = oForm.Items.Item("3").Specific
                                    If CInt(CType(oMatrix.Columns.Item("U_PDiscount").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value) < 0 Or CInt(CType(oMatrix.Columns.Item("U_PDiscount").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value) > 100 Then
                                        oApplication.Utilities.Message("Discount percentage should be greater then 0 and less than 100...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "_2" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE) Then
                                    Dim objFreightList As clsFreightList
                                    objFreightList = New clsFreightList
                                    objFreightList.LoadForm(modVariables.frmFreightRef)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                                modVariables.frmFreightType = ""
                                modVariables.frmFreightRef = ""
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                initializeControls(oForm)
                                blnIsLoad = True
                                initialize(oForm)
                                blnIsLoad = False
                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("3").Specific
                                If pVal.ItemUID = "3" And (pVal.ColUID = "U_PDiscount" Or pVal.ColUID = "U_PAmount") And pVal.Row > 0 And Not blnIsLoad Then
                                    calculateDiscount(oForm, pVal.Row)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "1" Then
                                    oMatrix = oForm.Items.Item("3").Specific
                                    oApplication.Utilities.addFreightAmt(oMatrix, modVariables.frmFreightRef)
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

    Private Sub initializeControls(ByVal oForm As SAPbouiCOM.Form)
        Try
            oApplication.Utilities.AddControls(oForm, "_2", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", 0, 0, "2", "View Freights", 0, 0, 0, False)
            oForm.Items.Item("_2").Width = "140"
            oForm.Items.Item("_2").Enabled = True
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oMatrix = oForm.Items.Item("3").Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If oForm.Items.Item("3").Enabled Then
                Dim strCurrency As String
                Dim dblTotal, dblAmount, dblDiscount As Double
                Dim sQuery As String
                oForm.Freeze(True)

                'For index As Integer = 1 To oMatrix.VisualRowCount
                '    If CType(oMatrix.Columns.Item("55").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value.Length > 0 Then
                '        'CType(oForm.Items.Item("4").Specific, SAPbouiCOM.CheckBox).Checked = True
                '        Exit Sub
                '    End If
                'Next

                If (CType(oForm.Items.Item("4").Specific, SAPbouiCOM.CheckBox).Checked) Then
                    CType(oForm.Items.Item("4").Specific, SAPbouiCOM.CheckBox).Checked = False
                End If


                sQuery = "Select U_FreID,U_Currency,U_PAmount,U_PDiscount From [@FRT1] Where U_RefCode = '" + modVariables.frmFreightRef + "'"
                oRecordSet.DoQuery(sQuery)

                If oRecordSet.EoF Then
                    For index As Integer = 1 To oMatrix.VisualRowCount
                        Dim strFrCode As String = CType(oMatrix.Columns.Item("1").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                        If CType(oMatrix.Columns.Item("U_FCalcu").Cells().Item(index).Specific, SAPbouiCOM.ComboBox).Selected.Value = "N" Then
                            sQuery = "Select ExpnsCode,U_Currency,U_PAmount,U_PDiscount From OEXD Where ExpnsCode = '" + strFrCode + "'"
                            oRecordSet.DoQuery(sQuery)
                            If Not oRecordSet.EoF Then
                                If oRecordSet.Fields.Item("ExpnsCode").Value = strFrCode Then
                                    If oRecordSet.Fields.Item("U_Currency").Value <> "" And CDbl(oRecordSet.Fields.Item("U_PAmount").Value) > 0 Then

                                        'CType(oMatrix.Columns.Item("U_Currency").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_Currency").Value
                                        'CType(oMatrix.Columns.Item("U_PAmount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_PAmount").Value
                                        'CType(oMatrix.Columns.Item("U_PDiscount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_PDiscount").Value

                                        If CType(oMatrix.Columns.Item("U_Currency").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value <> _
                        oRecordSet.Fields.Item("U_Currency").Value Then
                                            CType(oMatrix.Columns.Item("U_Currency").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_Currency").Value
                                        End If

                                        If CType(oMatrix.Columns.Item("U_PAmount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value <> _
                                                                oRecordSet.Fields.Item("U_PAmount").Value Then
                                            CType(oMatrix.Columns.Item("U_PAmount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_PAmount").Value
                                        End If

                                        If CType(oMatrix.Columns.Item("U_PDiscount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value <> _
                        oRecordSet.Fields.Item("U_PDiscount").Value Then
                                            CType(oMatrix.Columns.Item("U_PDiscount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_PDiscount").Value
                                        End If


                                        'dblAmount = oRecordSet.Fields.Item("U_PAmount").Value
                                        'dblDiscount = oRecordSet.Fields.Item("U_PDiscount").Value
                                        'dblTotal = dblAmount - (dblAmount * (dblDiscount / 100))

                                        'If dblTotal > 0 Then
                                        '    CType(oMatrix.Columns.Item("3").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_Currency").Value + dblTotal.ToString()
                                        '    CType(oMatrix.Columns.Item("U_FCalcu").Cells().Item(index).Specific, SAPbouiCOM.ComboBox).Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                        'End If

                                        strCurrency = CType(oMatrix.Columns.Item("U_Currency").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                                        If strCurrency.Length > 0 Then

                                            dblAmount = CType(oMatrix.Columns.Item("U_PAmount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                                            dblDiscount = CType(oMatrix.Columns.Item("U_PDiscount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value

                                            dblTotal = dblAmount - (dblAmount * (IIf(dblDiscount.ToString() = "", 0, dblDiscount) / 100))
                                            CType(oMatrix.Columns.Item("3").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = strCurrency + dblTotal.ToString()
                                            CType(oMatrix.Columns.Item("U_FCalcu").Cells().Item(index).Specific, SAPbouiCOM.ComboBox).Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)

                                        End If

                                    End If
                                End If
                            End If
                        End If
                    Next
                Else
                    For index As Integer = 1 To oMatrix.VisualRowCount
                        Dim strFrCode As String = CType(oMatrix.Columns.Item("1").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                        oRecordSet.MoveFirst()
                        While Not oRecordSet.EoF
                            If CType(oMatrix.Columns.Item("1").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_FreID").Value Then

                                Dim blnAlreadyExists As Boolean = False
                                For intRow As Integer = 1 To oMatrix.RowCount
                                    If intRow <> index Then
                                        If CType(oMatrix.Columns.Item("1").Cells().Item(intRow).Specific, SAPbouiCOM.EditText).Value _
                                                = CType(oMatrix.Columns.Item("1").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value Then
                                            blnAlreadyExists = True
                                        End If
                                    End If
                                Next

                                If Not blnAlreadyExists Then

                                    'CType(oMatrix.Columns.Item("U_Currency").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_Currency").Value
                                    'CType(oMatrix.Columns.Item("U_PAmount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_PAmount").Value
                                    'CType(oMatrix.Columns.Item("U_PDiscount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_PDiscount").Value

                                    If CType(oMatrix.Columns.Item("U_Currency").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value <> _
                       oRecordSet.Fields.Item("U_Currency").Value Then
                                        CType(oMatrix.Columns.Item("U_Currency").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_Currency").Value
                                    End If

                                    If CType(oMatrix.Columns.Item("U_PAmount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value <> _
                                                            oRecordSet.Fields.Item("U_PAmount").Value Then
                                        CType(oMatrix.Columns.Item("U_PAmount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_PAmount").Value
                                    End If

                                    If CType(oMatrix.Columns.Item("U_PDiscount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value <> _
                    oRecordSet.Fields.Item("U_PDiscount").Value Then
                                        CType(oMatrix.Columns.Item("U_PDiscount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_PDiscount").Value
                                    End If

                                    'CType(oMatrix.Columns.Item("U_DAmount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_DAmount").Value
                                    'dblAmount = oRecordSet.Fields.Item("U_PAmount").Value
                                    'dblDiscount = oRecordSet.Fields.Item("U_PDiscount").Value
                                    'dblTotal = dblAmount - (dblAmount * (dblDiscount / 100))

                                    'If dblTotal > 0 Then
                                    '    CType(oMatrix.Columns.Item("3").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_Currency").Value + dblTotal.ToString()
                                    '    CType(oMatrix.Columns.Item("U_FCalcu").Cells().Item(index).Specific, SAPbouiCOM.ComboBox).Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                    'End If

                                    strCurrency = CType(oMatrix.Columns.Item("U_Currency").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                                    If strCurrency.Length > 0 Then

                                        dblAmount = CType(oMatrix.Columns.Item("U_PAmount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                                        dblDiscount = CType(oMatrix.Columns.Item("U_PDiscount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value

                                        dblTotal = dblAmount - (dblAmount * (IIf(dblDiscount.ToString() = "", 0, dblDiscount) / 100))
                                        CType(oMatrix.Columns.Item("3").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = strCurrency + dblTotal.ToString()
                                        CType(oMatrix.Columns.Item("U_FCalcu").Cells().Item(index).Specific, SAPbouiCOM.ComboBox).Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                    End If

                                End If
                            End If
                            oRecordSet.MoveNext()
                        End While
                    Next
                End If
                oForm.Freeze(False)
            End If
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub calculateDiscount(ByVal oForm As SAPbouiCOM.Form, ByVal intRow As Integer)
        Try
            oForm.Freeze(True)
            If 1 = 1 Then
                oMatrix = oForm.Items.Item("3").Specific
                Dim strCurrency, strFreight As String
                Dim dblTotal, dblAmount, dblDiscount As Double
                strFreight = CType(oMatrix.Columns.Item("1").Cells().Item(intRow).Specific, SAPbouiCOM.EditText).Value
                strCurrency = CType(oMatrix.Columns.Item("U_Currency").Cells().Item(intRow).Specific, SAPbouiCOM.EditText).Value

                If strCurrency.Length > 0 Then

                    dblAmount = CType(oMatrix.Columns.Item("U_PAmount").Cells().Item(intRow).Specific, SAPbouiCOM.EditText).Value
                    dblDiscount = CType(oMatrix.Columns.Item("U_PDiscount").Cells().Item(intRow).Specific, SAPbouiCOM.EditText).Value

                    dblTotal = dblAmount - (dblAmount * (IIf(dblDiscount.ToString() = "", 0, dblDiscount) / 100))
                    CType(oMatrix.Columns.Item("3").Cells().Item(intRow).Specific, SAPbouiCOM.EditText).Value = strCurrency + dblTotal.ToString()
                    CType(oMatrix.Columns.Item("U_FCalcu").Cells().Item(intRow).Specific, SAPbouiCOM.ComboBox).Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)

                    'Dim dblDisountAmt As Double
                    'dblDisountAmt = (dblAmount * (IIf(dblDiscount.ToString() = "", 0, dblDiscount) / 100))
                    'Dim dblReqAmt As Double
                    'oApplication.Utilities.convertCurrency(strCurrency, dblDisountAmt, System.DateTime.Now.Date, modVariables.frmFreightCurr, dblReqAmt)
                    'CType(oMatrix.Columns.Item("U_DAmount").Cells().Item(intRow).Specific, SAPbouiCOM.EditText).Value = modVariables.frmFreightCurr + " " + dblReqAmt.ToString("0.00")

                End If
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

#End Region

End Class
