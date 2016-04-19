
Public Class clsPickList
    Inherits clsBase
    Dim WithEvents CreateComboBox As SAPbouiCOM.ComboBox
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try

            If pVal.FormTypeEx = frm_PickList Then
                Select Case pVal.BeforeAction
                    Case True

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                initialize(oForm)

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
    Public Sub New()
        MyBase.New()
    End Sub
    Public Sub initialize(form As SAPbouiCOM.Form)
        Try
            CreateComboBox = form.Items.Item("13").Specific

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub CreateComboxEvent(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles CreateComboBox.ComboSelectAfter

        If pVal.PopUpIndicator = 1 Then
            Dim strKey
            oApplication.Company.GetNewObjectCode(strKey)
            Dim recordSet As SAPbobsCOM.Recordset = oApplication.AdminCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            recordSet.DoQuery("SELECT a.DocEntry from ODLN a where MONTH(a.DocDate) = MONTH(GETDATE()) AND DAY(a.DocDate) = DAY(GETDATE())order by DocEntry desc")
            strKey = recordSet.Fields.Item(0).Value.ToString
            If Not String.IsNullOrEmpty(strKey) Or strKey = "0" Then
                strKey = "<?xml version=""1.0"" encoding=""UTF-16"" ?><DocumentParams><DocEntry>" + strKey + "</DocEntry></DocumentParams>"
                oApplication.Utilities.post_JournalEntryWithCostCenter(oForm, frm_Delivery, strKey)
            Else
                oApplication.Utilities.Message("Error Journal Entry Could Not be Created", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

        End If
    End Sub

End Class
