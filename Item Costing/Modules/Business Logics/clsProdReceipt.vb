Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsProdReceipt
    Inherits clsBase

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
            If pVal.FormTypeEx = frm_ProdReceipt Then
                Select Case pVal.BeforeAction
                    Case True

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                'oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                '    If pVal.ItemUID = "1" Then
                                '        If pVal.Action_Success Then
                                '            updateCosting(oForm)
                                '        End If
                                '    End If
                                'End If
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
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                            If BusinessObjectInfo.ActionSuccess Then
                                'oApplication.Utilities.update_ProductionCosting(oForm, BusinessObjectInfo.FormTypeEx, BusinessObjectInfo.ObjectKey)
                            End If
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

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub updateCosting(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oUpdateRecordSet As SAPbobsCOM.Recordset
            oUpdateRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select T0.ItemCode,T0.WhsCode,T0.Quantity As 'CompQty',T3.ItemType,ISNULL(T3.U_LabType,'I') As 'Labour Type',"
            strQuery += " SUM((T1.BaseQty) * T2.AvgPrice) As ActualCost From IGN1 T0 JOIN WOR1 T1 On T0.BaseEntry = T1.DocEntry JOIN "
            strQuery += " OITW T2 On T2.ItemCode = T1.ItemCode And T2.WhsCode = T1.wareHouse JOIN OITM T3 On T2.ItemCode = T3.ItemCode  "
            strQuery += " Where T0.BaseType = 202 And T0.DocEntry = (Select Max(DocEntry) From OIGN Where UserSign = '" + oApplication.Company.UserSignature.ToString() + "') Group By  "
            strQuery += " T0.ItemCode,T0.WhsCode,T0.Quantity,T3.ItemType,T3.U_LabType "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    Dim strItemCode As String = oRecordSet.Fields.Item("ItemCode").Value
                    Dim strWhsCode As String = oRecordSet.Fields.Item("WhsCode").Value
                    Dim dblActualCost As Double = oRecordSet.Fields.Item("ActualCost").Value
                    Dim dblCompQty As Double = oRecordSet.Fields.Item("CompQty").Value
                    strQuery = String.Empty

                    If oRecordSet.Fields.Item("ItemType").Value = "I" And oRecordSet.Fields.Item("Labour Type").Value = "I" Then
                        strQuery = "Update T0 Set U_AvgRMCst = ((" + dblActualCost.ToString() + " * " + dblCompQty.ToString() + ")+(ISNULL(U_AvgRMCst,0) * (T0.OnHand - " + dblCompQty.ToString() + ")))"
                    ElseIf (oRecordSet.Fields.Item("ItemType").Value = "L" And oRecordSet.Fields.Item("Labour Type").Value = "F") Then
                        strQuery = "Update T0 Set U_AvgFLbCst = ((" + dblActualCost.ToString() + " * " + dblCompQty.ToString() + ")+(ISNULL(U_AvgFLbCst,0) * (T0.OnHand - " + dblCompQty.ToString() + ")))"
                    ElseIf (oRecordSet.Fields.Item("ItemType").Value = "L" And oRecordSet.Fields.Item("Labour Type").Value = "V") Then
                        strQuery = "Update T0 Set U_AvgVLbCst = ((" + dblActualCost.ToString() + " * " + dblCompQty.ToString() + ")+(ISNULL(U_AvgVLbCst,0) * (T0.OnHand - " + dblCompQty.ToString() + ")))"
                    End If
                    strQuery += "/T0.OnHand From OITW T0 Where T0.ItemCode = '" + strItemCode + "' And WhsCode = '" + strWhsCode + "'"
                    oUpdateRecordSet.DoQuery(strQuery)
                    oRecordSet.MoveNext()
                End While
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
