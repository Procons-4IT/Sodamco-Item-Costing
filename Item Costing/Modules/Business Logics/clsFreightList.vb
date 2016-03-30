Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsFreightList
    Inherits clsBase
    Private oGrid As SAPbouiCOM.Grid
    Private oDtFreightList As SAPbouiCOM.DataTable
    Private strQuery As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub LoadForm(ByVal strRefCode As String)
        Try
            oForm = oApplication.Utilities.LoadForm(xml_FRT1, frm_FRT1)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            initialize(oForm, strRefCode)
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
            If pVal.FormTypeEx = frm_FRT1 Then
                Select Case pVal.BeforeAction
                    Case True

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
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

            strQuery = " Select ExpnsName,T0.U_Currency,T0.U_PAmount,T0.U_PDiscount,ISNULL(T0.U_Total,(T0.U_Currency + ' ' + Convert(VarChar,0.00))) As U_Total From [@FRT1] T0 JOIN OEXD T1 On T0.U_FreID = T1.ExpnsCode "
            strQuery += " Where T0.U_RefCode = '" + strRefCode + "' "
            oForm.DataSources.DataTables.Add("dtFreightList")

            oDtFreightList = oForm.DataSources.DataTables.Item(0)
            oDtFreightList.ExecuteQuery(strQuery)
            oGrid.DataTable = oDtFreightList

            'Format
            oGrid.Columns.Item("ExpnsName").TitleObject.Caption = "Freight Type"
            oGrid.Columns.Item("U_Currency").TitleObject.Caption = "Currency"
            oGrid.Columns.Item("U_PDiscount").TitleObject.Caption = "Predetermined Discount"
            oGrid.Columns.Item("U_PAmount").TitleObject.Caption = "Predetermind Amount"
            'oGrid.Columns.Item("U_DAmount").TitleObject.Caption = "Discount Amount"
            oGrid.Columns.Item("U_Total").TitleObject.Caption = "Total"

            oGrid.Columns.Item("U_PDiscount").RightJustified = True
            oGrid.Columns.Item("U_PAmount").RightJustified = True
            'oGrid.Columns.Item("U_DAmount").RightJustified = True
            oGrid.Columns.Item("U_Total").RightJustified = True

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
