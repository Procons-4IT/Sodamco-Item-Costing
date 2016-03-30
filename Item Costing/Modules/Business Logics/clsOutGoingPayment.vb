Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsOutGoingPayment
    Inherits clsBase
    Private oEditText As SAPbouiCOM.EditText

    Public Sub New()
        MyBase.New()
    End Sub

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            Select Case pVal.MenuUID
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                Case mnu_ADD
                Case mnu_COMM_O
                    callCommission(oForm)
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_OutPayment Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                                If oApplication.SBO_Application.Menus.Exists(mnu_COMM_O) Then
                                    oApplication.SBO_Application.Menus.RemoveEx(mnu_COMM_O)
                                End If
                                oApplication.Utilities.removeCommission(oForm.Items.Item("_52").Specific.value)
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                initializeControls(oForm)
                                dataBind(oForm)
                                enableControls(oForm, False)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "_14" Then
                                    Dim ocheck As SAPbouiCOM.OptionBtn
                                    ocheck = oForm.Items.Item("58").Specific
                                    If oApplication.Utilities.getEditTextvalue(oForm, "5") <> "" Or ocheck.Selected = True Then
                                        callCommission(oForm)
                                    Else
                                        oApplication.Utilities.Message("Select Customer / Supplier / Account Option", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
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

#Region "Right Click"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            If oForm.TypeEx = frm_OutPayment Then
                If (eventInfo.BeforeAction = True) Then
                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    Dim oMenus As SAPbouiCOM.Menus
                    Try
                        Dim ocheck As SAPbouiCOM.OptionBtn
                        ocheck = oForm.Items.Item("58").Specific
                        If oApplication.Utilities.getEditTextvalue(oForm, "5") <> "" Or ocheck.Selected = True Then
                            oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                            If Not oMenuItem.SubMenus.Exists(mnu_COMM_O) Then
                                Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                                oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                oCreationPackage.UniqueID = "ComChargesO"
                                oCreationPackage.String = "Commission Charges"
                                oCreationPackage.Enabled = True
                                oMenus = oMenuItem.SubMenus
                                oMenus.AddEx(oCreationPackage)
                            Else
                                If oApplication.SBO_Application.Menus.Exists(mnu_COMM_O) Then
                                    oApplication.SBO_Application.Menus.RemoveEx(mnu_COMM_O)
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                Else
                    If oApplication.SBO_Application.Menus.Exists(mnu_COMM_O) Then
                        oApplication.SBO_Application.Menus.RemoveEx(mnu_COMM_O)
                    End If
                End If
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Function"
    Private Sub initializeControls(ByVal oForm As SAPbouiCOM.Form)
        Try
            oApplication.Utilities.AddControls(oForm, "_53", "53", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 0, 0, "53", "Commission Ref", 0, 0, 0, False)
            oApplication.Utilities.AddControls(oForm, "_52", "52", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 0, 0, "52", "", 0, 0, 0, True)
            oApplication.Utilities.AddControls(oForm, "_14", "10001005", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "LEFT", 0, 0, "10001005", "Commission", 0, 0, 0, False)
            oForm.Items.Item("_53").Visible = True
            oForm.Items.Item("_52").Visible = True
            oForm.Items.Item("_53").LinkTo = "_52"
            oForm.Items.Item("_52").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub enableControls(ByVal oForm As SAPbouiCOM.Form, ByVal blnStatus As Boolean)
        Try
            oForm.Items.Item("26").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("_52").Enabled = blnStatus
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub dataBind(ByVal oForm As SAPbouiCOM.Form)
        Try
            oEditText = oForm.Items.Item("_52").Specific
            oEditText.DataBind.SetBound(True, "OVPM", "U_RefCode")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub callCommission(ByVal oForm As SAPbouiCOM.Form)
        Try
            If oForm.Items.Item("_52").Specific.value.ToString().Length = 0 Then
                Dim strRef As String = String.Empty
                oApplication.Utilities.addCommissionReference(strRef)
                enableControls(oForm, True)
                oForm.Items.Item("_52").Specific.value = strRef
                enableControls(oForm, False)
                Dim objCommCharge As clsCommissionCharges
                objCommCharge = New clsCommissionCharges
                objCommCharge.LoadForm(oForm.Items.Item("_52").Specific.value, "P")
            ElseIf oForm.Items.Item("_52").Specific.value.ToString().Length > 0 Then
                If oApplication.Utilities.validateRefExist(oForm.Items.Item("_52").Specific.value) Then
                    Dim objCommCharge As clsCommissionCharges
                    objCommCharge = New clsCommissionCharges
                    objCommCharge.LoadForm(oForm.Items.Item("_52").Specific.value, "P")
                Else
                    Dim strRef As String = String.Empty
                    oApplication.Utilities.addCommissionReference(strRef)
                    If strRef.Length > 0 Then
                        enableControls(oForm, True)
                        oForm.Items.Item("_52").Specific.value = strRef
                        enableControls(oForm, False)
                        Dim objCommCharge As clsCommissionCharges
                        objCommCharge = New clsCommissionCharges
                        objCommCharge.LoadForm(oForm.Items.Item("_52").Specific.value, "P")
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

End Class
