Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsFreight
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
            If pVal.FormTypeEx = frm_Freight Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "1" And Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    If Not validate(oForm) Then
                                        BubbleEvent = False
                                    End If
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
                                Dim strCurrency, strName As String
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If pVal.ItemUID = "3" And (pVal.ColUID = "U_Currency" Or pVal.ColUID = "U_Name") Then
                                        strCurrency = oDataTable.GetValue("CurrCode", 0)
                                        strName = oDataTable.GetValue("CurrName", 0)
                                        Try
                                            oMatrix = oForm.Items.Item("3").Specific
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Currency", pVal.Row, strCurrency)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Name", pVal.Row, strName)
                                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        Catch ex As Exception
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Currency", pVal.Row, strCurrency)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Name", pVal.Row, strName)
                                        End Try
                                    End If
                                Catch ex As Exception
                                    'Throw ex
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
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub addChooseFromList(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCurrCol, oCurrNCol As SAPbouiCOM.Column
            oMatrix = oForm.Items.Item("3").Specific
            oCurrCol = oMatrix.Columns.Item("U_Currency")
            oCurrNCol = oMatrix.Columns.Item("U_Name")

            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams

            'Adding Currency CFL..
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "37"
            oCFLCreationParams.UniqueID = "CFL_PR_1"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCurrCol.ChooseFromListUID = "CFL_PR_1"
            oCurrCol.ChooseFromListAlias = "CurrCode"

            'Adding Currency Name CFL..
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "37"
            oCFLCreationParams.UniqueID = "CFL_PR_2"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCurrNCol.ChooseFromListUID = "CFL_PR_2"
            oCurrNCol.ChooseFromListAlias = "CurrName"

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub addControls(ByVal oForm As SAPbouiCOM.Form)
        Try
            
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function validate(ByVal oForm As SAPbouiCOM.Form)
        Dim _retVal As Boolean = True
        oMatrix = oForm.Items.Item("3").Specific
        Try
            For index As Integer = 1 To oMatrix.RowCount
                Dim strAmount As String = CType(oMatrix.Columns().Item("U_PAmount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                Dim dblAmount As Double
                Dim strCurrency As String
                strCurrency = CType(oMatrix.Columns().Item("U_Currency").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                dblAmount = IIf(strAmount.Length > 0, CDbl(CType(oMatrix.Columns().Item("U_PAmount").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value), 0)

                If strCurrency.Length > 0 Then
                    If dblAmount = 0 Then
                        oApplication.Utilities.Message("Enter Predefined Amount to Proceed... in Row: " + index.ToString, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        _retVal = False
                        Exit For
                    End If
                End If
                If dblAmount > 0 Then
                    If strCurrency.Length = 0 Then
                        oApplication.Utilities.Message("Select Currency to Proceed... in Row: " + index.ToString, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        _retVal = False
                        Exit For
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
