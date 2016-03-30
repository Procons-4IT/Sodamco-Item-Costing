Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsCustomer
    Inherits clsBase
    Private oRecordSet As SAPbobsCOM.Recordset
    Private strQuery As String

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
                Case mnu_PRJL_C
                    If Not oForm.Items.Item("5").Specific.value = "" And oForm.Items.Item("40").Specific.value = "C" Then
                        Dim objSpecialPriceList As clsSpecialPriceList
                        objSpecialPriceList = New clsSpecialPriceList
                        objSpecialPriceList.LoadForm(oForm.Items.Item("5").Specific.value, "")
                    Else
                        oApplication.Utilities.Message("Select Customer to Get Special Price List...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                    If oMenuItem.SubMenus.Exists(pVal.MenuUID) Then
                        oApplication.SBO_Application.Menus.RemoveEx(pVal.MenuUID)
                    End If
                Case mnu_CPRL_C
                    If Not oForm.Items.Item("5").Specific.value = "" Then
                        Dim objPromList As clsCustPromotionList
                        objPromList = New clsCustPromotionList
                        objPromList.LoadForm(oForm.Items.Item("5").Specific.value)
                    Else
                        oApplication.Utilities.Message("Select Customer to Get Promotion List...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                    If oMenuItem.SubMenus.Exists(pVal.MenuUID) Then
                        oApplication.SBO_Application.Menus.RemoveEx(pVal.MenuUID)
                    End If
            End Select
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Delivery Then
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
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Right Click"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            If oForm.TypeEx = frm_Customer Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'

                If (eventInfo.BeforeAction = True) Then
                    Try
                        'Project List
                        If Not oMenuItem.SubMenus.Exists(mnu_PRJL_C) Then
                            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                            oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                            oCreationPackage.UniqueID = mnu_PRJL_C
                            oCreationPackage.String = "Special Price List"
                            oCreationPackage.Enabled = True
                            oMenus = oMenuItem.SubMenus
                            oMenus.AddEx(oCreationPackage)
                        End If

                        'Promotion List
                        If CType(oForm.Items.Item("40").Specific, SAPbouiCOM.ComboBox).Value = "C" Then
                            If Not oMenuItem.SubMenus.Exists(mnu_CPRL_C) Then
                                Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                                oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                oCreationPackage.UniqueID = mnu_CPRL_C
                                oCreationPackage.String = "Promotion List"
                                oCreationPackage.Enabled = True
                                oMenus = oMenuItem.SubMenus
                                oMenus.AddEx(oCreationPackage)
                            End If
                        End If
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                Else
                    If oMenuItem.SubMenus.Exists(mnu_PRJL_C) Then
                        oApplication.SBO_Application.Menus.RemoveEx(mnu_PRJL_C)
                    End If

                    If oMenuItem.SubMenus.Exists(mnu_CPRL_C) Then
                        oApplication.SBO_Application.Menus.RemoveEx(mnu_CPRL_C)
                    End If
                End If
            End If
        Catch ex As Exception
            oForm.Freeze(False)
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

#End Region

End Class
