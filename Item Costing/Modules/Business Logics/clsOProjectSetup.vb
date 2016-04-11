Public Class clsOProjectSetup
    Inherits clsBase
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Private oRecordSet As SAPbobsCOM.Recordset
    Private oPassword As SAPbouiCOM.EditText
    Private oUsername As SAPbouiCOM.EditText
    Private oDocEntry As SAPbouiCOM.EditText

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_OSUS, frm_OSUS)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            initializeDataSource(oForm)
            initialize(oForm)
            'oForm.EnableMenu(mnu_ADD_ROW, True)
            'oForm.EnableMenu(mnu_DELETE_ROW, True)
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub
    Private Sub initializeDataSource(ByVal oForm As SAPbouiCOM.Form)
        Try
            'oForm.DataSources.UserDataSources.Add("udsCust", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 40)
            'oForm.DataSources.UserDataSources.Add("udsName", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 200)
            'oEditText = oForm.Items.Item("19").Specific
            'oEditText.DataBind.SetBound(True, "", "udsCust")
            'oEditText = oForm.Items.Item("20").Specific
            'oEditText.DataBind.SetBound(True, "", "udsName")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@OSUS")
            oUsername = oForm.Items.Item("Item_0").Specific
            oPassword = oForm.Items.Item("Item_2").Specific
            oDocEntry = oForm.Items.Item("docEntry").Specific
            oDocEntry.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Add), SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oDocEntry.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Ok), SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@PSP1")
            'oMatrix = oForm.Items.Item("3").Specific
            'oMatrix.LoadFromDataSource()
            'oMatrix.AddRow(1, -1) 
            'oMatrix.FlushToDataSource()

            'clearDataSource(oForm)

            'oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRecordSet.DoQuery("Select IsNull(MAX(DocEntry),1) From [@OSUS]")
            'If Not oRecordSet.EoF Then
            '    oDBDataSource.SetValue("DocNum", 0, oRecordSet.Fields.Item(0).Value.ToString())
            'End If

            enableControl(oForm, True)
            'modeforControl(oForm)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub clearDataSource(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Items.Item("Item_0").Specific.value = ""
            oForm.Items.Item("Item_2").Specific.value = ""
        Catch ex As Exception

        End Try
    End Sub
    Private Sub enableControl(ByVal oForm As SAPbouiCOM.Form, ByVal blnStatus As Boolean)
        Try
            'oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("Item_0").Enabled = blnStatus
            oForm.Items.Item("Item_2").Enabled = blnStatus
            'oForm.Items.Item("11").Enabled = blnStatus
            'oForm.Items.Item("3").Enabled = blnStatus
            'oForm.Items.Item("14").Enabled = blnStatus
            'oForm.Items.Item("15").Enabled = blnStatus
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

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
                                'AddRow(oForm)
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
                                        Dim crypt As CryptoUtility = New CryptoUtility
                                        Dim encryptedPassword = crypt.Encrypt(oPassword.Value)
                                        oPassword.Value = encryptedPassword
                                        'If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
                                        '    BubbleEvent = False
                                        '    Exit Sub
                                        'Else

                                        'End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK

                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    initialize(oForm)
                                    'Dim crypt As CryptoUtility = New CryptoUtility
                                    'Dim encryptedPassword = crypt.Encrypt(oPassword.Value)
                                    'oPassword.Value = encryptedPassword
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "1"
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            If pVal.Action_Success Then
                                                clearDataSource(oForm)
                                            End If
                                        End If
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
End Class
