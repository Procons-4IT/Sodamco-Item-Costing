Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsSpecialPriceList
    Inherits clsBase
    Private oGrid As SAPbouiCOM.Grid
    Private oDtSpecialPriceList As SAPbouiCOM.DataTable
    Private strQuery As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub LoadForm(ByVal strCust As String, Optional ByVal strProject As String = "")
        Try
            oForm = oApplication.Utilities.LoadForm(xml_PSPL, frm_PSPL)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            initialize(oForm, strCust, strProject)
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
            If pVal.FormTypeEx = frm_PSPL Then
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
    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form, ByVal strCust As String, ByVal strProject As String)
        Try
            oGrid = oForm.Items.Item("3").Specific
            strQuery = " Select T1.U_PrjCode,T1.U_PrjName,T1.U_EffFrom,T1.U_EffTo,Case When IsNull(Status,'O') = 'O' Then 'Open' Else 'Close' End As DocStatus,T0.U_ItmCode,T0.U_ItmName ,T3.CurrName,T0.U_UnitPrice ,Case When T0.U_DisType = 'D' Then 'Discount' Else 'Price' End As DisType,U_Discount,U_DisPrice"
            strQuery += " From [@PSP1] T0 JOIN [@OPSP] T1 On T0.DocEntry = T1.DocEntry JOIN OPRJ T2 On T1.U_PrjCode = T2.PrjCode "
            strQuery += " JOIN OCRN T3 On T0.U_Currency = T3.CurrCode  "
            strQuery += " Where T2.U_CardCode = '" + strCust + "' "
            If strProject.Length > 0 Then strQuery += " And T2.U_PrjCode = '" + strProject + "' "
            strQuery += " Order By U_PrjCode "
            oForm.DataSources.DataTables.Add("dtSpecialPriceList")
            oDtSpecialPriceList = oForm.DataSources.DataTables.Item(0)
            oDtSpecialPriceList.ExecuteQuery(strQuery)
            oGrid.DataTable = oDtSpecialPriceList

            'Format
            oGrid.Columns.Item("U_PrjCode").TitleObject.Caption = "Project"
            oGrid.Columns.Item("U_PrjName").TitleObject.Caption = "Project Name"
            oGrid.Columns.Item("U_EffFrom").TitleObject.Caption = "Effective From"
            oGrid.Columns.Item("U_EffTo").TitleObject.Caption = "Effective To"
            oGrid.Columns.Item("DocStatus").TitleObject.Caption = "Status"
            oGrid.Columns.Item("U_ItmCode").TitleObject.Caption = "Item Code"
            oGrid.Columns.Item("U_ItmName").TitleObject.Caption = "Item Name"
            oGrid.Columns.Item("CurrName").TitleObject.Caption = "Currency"
            oGrid.Columns.Item("U_UnitPrice").TitleObject.Caption = "Unit Price"
            oGrid.Columns.Item("DisType").TitleObject.Caption = "Discount Type"
            oGrid.Columns.Item("U_Discount").TitleObject.Caption = "Discount / Price"
            oGrid.Columns.Item("U_DisPrice").TitleObject.Caption = "Price After Discount"

            oGrid.Columns.Item("U_UnitPrice").RightJustified = True
            oGrid.Columns.Item("U_Discount").RightJustified = True
            oGrid.Columns.Item("U_DisPrice").RightJustified = True

            'Collapse Level By Project
            oGrid.CollapseLevel = 1

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub loadValues(ByVal oForm As SAPbouiCOM.Form, ByVal blnStatus As Boolean)
        Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
