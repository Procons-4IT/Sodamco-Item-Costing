Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsGRPO
    Inherits clsBase
    Private oEditText As SAPbouiCOM.EditText
    
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
            If pVal.FormTypeEx = frm_GRPO Then
                Select Case pVal.BeforeAction
                    Case True

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                'Dim oUser As SAPbobsCOM.Users
                                'oUser = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers)
                                'If oUser.GetByKey(oApplication.Company.UserSignature) Then
                                '    If oUser.UserFields.Fields.Item("U_HideAmt").Value = "Y" Then
                                '        initializeControls(oForm)
                                '        enableControls(oForm, False)
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

                    End Select
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Function"

    'Private Sub initializeControls(ByVal oForm As SAPbouiCOM.Form)
    '    Try
    '        oApplication.Utilities.AddControls(oForm, "_22", "22", SAPbouiCOM.BoFormItemTypes.it_EDIT, "COPY", 1, 10, "22", "****", 0, 0, 0, False)
    '        oApplication.Utilities.AddControls(oForm, "_24", "24", SAPbouiCOM.BoFormItemTypes.it_EDIT, "COPY", 1, 10, "22", "****", 0, 0, 0, False)
    '        oApplication.Utilities.AddControls(oForm, "_42", "42", SAPbouiCOM.BoFormItemTypes.it_EDIT, "COPY", 1, 10, "22", "****", 0, 0, 0, False)
    '        oApplication.Utilities.AddControls(oForm, "_89", "89", SAPbouiCOM.BoFormItemTypes.it_EDIT, "COPY", 1, 10, "22", "****", 0, 0, 0, False)
    '        oApplication.Utilities.AddControls(oForm, "_103", "103", SAPbouiCOM.BoFormItemTypes.it_EDIT, "COPY", 1, 10, "22", "****", 0, 0, 0, False)
    '        oApplication.Utilities.AddControls(oForm, "_27", "27", SAPbouiCOM.BoFormItemTypes.it_EDIT, "COPY", 1, 10, "22", "****", 0, 0, 0, False)
    '        oApplication.Utilities.AddControls(oForm, "_29", "29", SAPbouiCOM.BoFormItemTypes.it_EDIT, "COPY", 1, 10, "22", "****", 0, 0, 0, False)

    '        oForm.Items.Item("_22").RightJustified = True
    '        oEditText = oForm.Items.Item("_22").Specific
    '        oEditText.IsPassword = True
    '        oEditText.Value = "abcde"

    '        oForm.Items.Item("_24").RightJustified = True
    '        oEditText = oForm.Items.Item("_24").Specific
    '        oEditText.IsPassword = True
    '        oEditText.Value = "abcde"

    '        oForm.Items.Item("_42").RightJustified = True
    '        oEditText = oForm.Items.Item("_42").Specific
    '        oEditText.IsPassword = True
    '        oEditText.Value = "abcde"

    '        oForm.Items.Item("_89").RightJustified = True
    '        oEditText = oForm.Items.Item("_89").Specific
    '        oEditText.IsPassword = True
    '        oEditText.Value = "abcde"

    '        oForm.Items.Item("_103").RightJustified = True
    '        oEditText = oForm.Items.Item("_103").Specific
    '        oEditText.IsPassword = True
    '        oEditText.Value = "abcde"

    '        oForm.Items.Item("_27").RightJustified = True
    '        oEditText = oForm.Items.Item("_27").Specific
    '        oEditText.IsPassword = True
    '        oEditText.Value = "abcde"

    '        oForm.Items.Item("_29").RightJustified = True
    '        oEditText = oForm.Items.Item("_29").Specific
    '        oEditText.IsPassword = True
    '        oEditText.Value = "abcde"

    '        oForm.Items.Item("_22").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
    '        oForm.Items.Item("_24").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
    '        oForm.Items.Item("_42").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
    '        oForm.Items.Item("_89").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
    '        oForm.Items.Item("_103").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
    '        oForm.Items.Item("_27").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
    '        oForm.Items.Item("_29").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

    '        oForm.Items.Item("105").Visible = False

    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    'Private Sub enableControls(ByVal oForm As SAPbouiCOM.Form, ByVal blnStatus As Boolean)
    '    Try
    '        oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
    '        oForm.Items.Item("_22").Enabled = blnStatus
    '        oForm.Items.Item("_24").Enabled = blnStatus
    '        oForm.Items.Item("_42").Enabled = blnStatus
    '        oForm.Items.Item("_89").Enabled = blnStatus
    '        oForm.Items.Item("_103").Enabled = blnStatus
    '        oForm.Items.Item("_27").Enabled = blnStatus
    '        oForm.Items.Item("_29").Enabled = blnStatus
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

#End Region

End Class
