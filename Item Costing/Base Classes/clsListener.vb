Public Class clsListener
    Inherits Object

    Private ThreadClose As New Threading.Thread(AddressOf CloseApp)
    Private WithEvents _SBO_Application As SAPbouiCOM.Application
    Private _Company As SAPbobsCOM.Company
    Private _SUserCompany As SAPbobsCOM.Company
    Public ReadOnly Property AdminCompany() As SAPbobsCOM.Company
        Get

            If _SUserCompany Is Nothing Then
                _SUserCompany = Utilities.connectSUserCompany(_Company)
            End If
            Return _SUserCompany
        End Get
    End Property
    Private _Utilities As clsUtilities
    Private _Collection As Hashtable
    Private _LookUpCollection As Hashtable
    Private _FormUID As String
    Private _Log As clsLog_Error
    Private oMenuObject As Object
    Private oItemObject As Object
    Private oSystemForms As Object
    Dim objFilters As SAPbouiCOM.EventFilters
    Dim objFilter As SAPbouiCOM.EventFilter

#Region "New"
    Public Sub New()
        MyBase.New()
        Try
            _Company = New SAPbobsCOM.Company
            _Utilities = New clsUtilities
            _Collection = New Hashtable(10, 0.5)
            _LookUpCollection = New Hashtable(10, 0.5)
            oSystemForms = New clsSystemForms
            _Log = New clsLog_Error

            SetApplication()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Public Properties"

    Public ReadOnly Property SBO_Application() As SAPbouiCOM.Application
        Get
            Return _SBO_Application
        End Get
    End Property

    Public Property Company() As SAPbobsCOM.Company
        Get
            Return _Company
        End Get
        Set(value As SAPbobsCOM.Company)
            _Company = value
        End Set

    End Property

    Public ReadOnly Property Utilities() As clsUtilities
        Get
            Return _Utilities
        End Get
    End Property

    Public ReadOnly Property Collection() As Hashtable
        Get
            Return _Collection
        End Get
    End Property

    Public ReadOnly Property LookUpCollection() As Hashtable
        Get
            Return _LookUpCollection
        End Get
    End Property

    Public ReadOnly Property Log() As clsLog_Error
        Get
            Return _Log
        End Get
    End Property

#Region "Filter"

    Public Sub SetFilter(ByVal Filters As SAPbouiCOM.EventFilters)
        oApplication.SBO_Application.SetFilter(Filters)
    End Sub

    Public Sub SetFilter()
        Try
            ''Form Load
            objFilters = New SAPbouiCOM.EventFilters()

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            objFilter.Add(frm_GRPO) 'GRPO
            'objFilter.Add(frm_Project) 'Project
            objFilter.Add(frm_ITEM_MASTER) ' Item Master
            objFilter.AddEx(frm_Delivery) 'Delivery
            objFilter.AddEx(frm_INVOICES) 'Invoice
            objFilter.AddEx(frm_SaleReturn) 'Return
            objFilter.AddEx(frm_ARCreditMemo) 'AR Credit Memo
            'objFilter.AddEx(frm_Customer) 'Customer
            'objFilter.AddEx(frm_ORDR) 'Order
            objFilter.AddEx(frm_OSUS) 'Special Price
            'objFilter.AddEx(frm_Banking) ' Banking
            'objFilter.AddEx(frm_ComType) 'Commission Type
            'objFilter.AddEx(frm_CommCharges) 'Commission Charges
            'objFilter.AddEx(frm_IncomingPayment) 'Incoming Payment
            'objFilter.AddEx(frm_OutPayment) 'OutGoing Payment
            'objFilter.AddEx(frm_Deposits) 'Deposits
            'objFilter.AddEx(frm_OPRM) 'Promotion
            'objFilter.AddEx(frm_OCPR) 'Promotion Mapping
            'objFilter.AddEx(frm_OPRT) 'Promotion Template
            'objFilter.AddEx(frm_PRT2) 'Promotion Child - 2
            'objFilter.AddEx(frm_Freight) 'Freight
            'objFilter.AddEx(frm_DocumentFreight) 'Document Freight
            'objFilter.AddEx(frm_Quotation) ' Quotation
            objFilter.AddEx(frm_INVOICESPAYMENT) ' Invoice + Payment
            objFilter.AddEx(frm_ARReverseInvoice) 'Reverse Invoice
            objFilter.AddEx(frm_GI_INVENTORY) ' Inventory Goods Issue
            objFilter.AddEx(frm_GR_INVENTORY) ' Inventory Goods Receipt
            objFilter.AddEx(frm_I_Transfer) ' Inventory Transfer
            objFilter.AddEx(frm_PurReturn) ' Purchase Return
            objFilter.AddEx(frm_APCreditMemo) ' AP Credit Memo

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE)
            'objFilter.AddEx(frm_IncomingPayment) 'Incoming Payment
            'objFilter.AddEx(frm_OutPayment) 'OutGoing Payment
            'objFilter.AddEx(frm_Deposits) 'Deposits
            'objFilter.AddEx(frm_CommCharges) 'Commission Charges
            'objFilter.AddEx(frm_DocumentFreight) 'Document Freight
            'objFilter.AddEx(frm_Quotation) ' Quotation
            'objFilter.AddEx(frm_ORDR) 'Order
            objFilter.AddEx(frm_Delivery) 'Delivery
            objFilter.AddEx(frm_SaleReturn) 'Return
            objFilter.AddEx(frm_INVOICES) 'Invoice
            objFilter.AddEx(frm_INVOICESPAYMENT) ' Invoice + Payment
            objFilter.AddEx(frm_ARCreditMemo) 'AR Credit Memo
            objFilter.AddEx(frm_ARReverseInvoice) 'Reverse Invoice
            objFilter.AddEx(frm_GI_INVENTORY) ' Inventory Goods Issue


            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)
            'objFilter.AddEx(frm_ORDR) 'Order
            'objFilter.AddEx(frm_Customer) 'Customer
            'objFilter.AddEx(frm_IncomingPayment) 'Incoming Payment
            'objFilter.AddEx(frm_OutPayment) 'OutGoing Payment
            'objFilter.AddEx(frm_Deposits) 'Deposits
            'objFilter.AddEx(frm_OCPR) 'Promotion Mapping
            objFilter.AddEx(frm_OSUS) 'Special Price

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            objFilter.AddEx(frm_OSUS) 'Special Price
            objFilter.AddEx(frm_ProdReceipt) 'Production Receipt
            objFilter.AddEx(frm_Delivery) 'Delivery
            objFilter.AddEx(frm_INVOICES) 'Invoice
            objFilter.AddEx(frm_SaleReturn) 'Return
            objFilter.AddEx(frm_ARCreditMemo) 'AR Credit Memo
            'objFilter.AddEx(frm_ComType) 'Commission Type
            'objFilter.AddEx(frm_CommCharges) 'Commission Charges
            'objFilter.AddEx(frm_OPRM) 'Promotion
            'objFilter.AddEx(frm_OCPR) 'Promotion Mapping
            'objFilter.AddEx(frm_ORDR) 'Order
            'objFilter.Add(frm_Project) 'Project
            'objFilter.AddEx(frm_IncomingPayment) 'Incoming Payment
            'objFilter.AddEx(frm_OutPayment) 'OutGoing Payment
            'objFilter.AddEx(frm_Deposits) 'Deposits
            'objFilter.AddEx(frm_OPRT) 'Promotion Template
            'objFilter.AddEx(frm_PRT2) 'Promotion Child - 2
            'objFilter.AddEx(frm_Freight) 'Freight
            'objFilter.AddEx(frm_DocumentFreight) 'Document Freight
            objFilter.AddEx(frm_GI_INVENTORY) ' Inventory Goods Issue
            objFilter.AddEx(frm_INVOICESPAYMENT) ' Invoice + Payment

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
            'objFilter.AddEx(frm_CommCharges) 'Commission Charges
            objFilter.AddEx(frm_OSUS) 'Special Price
            ''objFilter.AddEx(frm_OPRT) 'Promotion Template

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
            'objFilter.Add(frm_Project) ' Project
            objFilter.AddEx(frm_OSUS) 'Special Price
            objFilter.AddEx(frm_ITEM_MASTER) ' Item Master
            'objFilter.AddEx(frm_Banking) ' Banking
            'objFilter.AddEx(frm_CommCharges) 'Commission Charges
            'objFilter.AddEx(frm_OPRM) 'Promotion
            'objFilter.AddEx(frm_OCPR) 'Promotion Mapping
            'objFilter.AddEx(frm_ComType) 'Commission Type
            'objFilter.AddEx(frm_ORDR) 'Order
            'objFilter.AddEx(frm_OPRT) 'Promotion Template
            'objFilter.AddEx(frm_PRT2) 'Promotion Child - 2
            'objFilter.AddEx(frm_Freight) 'Promotion Mapping

            'objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)
            'objFilter.AddEx(frm_ORDR) ' Order
            'objFilter.AddEx(frm_ComType) 'Commission Type
            'objFilter.AddEx(frm_DocumentFreight) 'Document Freight

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
            'objFilter.AddEx(frm_ORDR) ' Order
            objFilter.AddEx(frm_OSUS) 'Special Price
            'objFilter.AddEx(frm_DocumentFreight) 'Document Freight

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD)
            objFilter.AddEx(frm_OSUS) 'Special Price
            'objFilter.AddEx(frm_OPRM) 'Promotion
            'objFilter.AddEx(frm_ORDR) ' Order
            ''objFilter.AddEx(frm_OPRT) 'Promotion Template

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK)
            objFilter.AddEx(frm_OSUS) 'Special Price
            'objFilter.AddEx(frm_Delivery) 'Delivery
            'objFilter.AddEx(frm_OPRM) 'Promotion
            'objFilter.AddEx(frm_ORDR) 'Order
            objFilter.Add(frm_ITEM_MASTER) 'Item Master
            'objFilter.AddEx(frm_OPRT) 'Promotion Template
            'objFilter.AddEx(frm_DocumentFreight) 'Document Freight
            objFilter.AddEx(frm_INVOICES) 'Invoice
            'objFilter.AddEx(frm_Quotation) ' Quotation
            objFilter.AddEx(frm_SaleReturn) ' Sale Return
            objFilter.AddEx(frm_INVOICESPAYMENT) ' Invoice + Payment
            objFilter.AddEx(frm_ARCreditMemo) ' AR Credit Memo
            objFilter.AddEx(frm_ARReverseInvoice) 'Reverse Invoice

            'objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK)
            'objFilter.AddEx(frm_ORDR) 'Order
            'objFilter.AddEx(frm_Customer) 'Customer
            'objFilter.AddEx(frm_IncomingPayment) 'Incoming Payment
            'objFilter.AddEx(frm_OutPayment) 'OutGoing Payment
            'objFilter.AddEx(frm_Deposits) 'Deposits

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            objFilter.AddEx(frm_Delivery) 'Delivery
            objFilter.AddEx(frm_SaleReturn) 'Return
            objFilter.AddEx(frm_INVOICES) 'Invoice
            objFilter.AddEx(frm_ARCreditMemo) 'Credit Memo
            objFilter.AddEx(frm_ProdReceipt) 'Production Receipt
            'objFilter.AddEx(frm_IncomingPayment) 'Incoming Payment
            'objFilter.AddEx(frm_OutPayment) 'OutGoing Payment
            'objFilter.AddEx(frm_Deposits) 'Deposits
            objFilter.AddEx(frm_GI_INVENTORY) ' Inventory Goods Issue
            objFilter.AddEx(frm_INVOICESPAYMENT) ' Invoice + Payment
            objFilter.AddEx(frm_GR_INVENTORY) ' Inventory Goods Receipt
            objFilter.Add(frm_GRPO) 'GRPO
            objFilter.AddEx(frm_I_Transfer) ' Inventory Transfer
            objFilter.AddEx(frm_PurReturn) ' Purchase Return
            objFilter.AddEx(frm_APCreditMemo) ' AP Credit Memo

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
            objFilter.AddEx(frm_Delivery) 'Delivery
            objFilter.AddEx(frm_SaleReturn) 'Return
            objFilter.AddEx(frm_INVOICES) 'Invoice
            objFilter.AddEx(frm_ARCreditMemo) 'Credit Memo
            'objFilter.AddEx(frm_IncomingPayment) 'Incoming Payment
            'objFilter.AddEx(frm_OutPayment) 'OutGoing Payment
            'objFilter.AddEx(frm_Deposits) 'Deposits
            objFilter.AddEx(frm_GI_INVENTORY) ' Inventory Goods Issue
            objFilter.AddEx(frm_INVOICESPAYMENT) ' Invoice + Payment
            objFilter.AddEx(frm_ProdReceipt) 'Production Receipt
            objFilter.AddEx(frm_GR_INVENTORY) ' Inventory Goods Receipt
            objFilter.AddEx(frm_I_Transfer) ' Inventory Transfer

            'objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_RESIZE)
            'objFilter.AddEx(frm_OCPR) 'Promotion

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE)
            objFilter.AddEx(frm_OSUS) 'Special Price
            'objFilter.AddEx(frm_DocumentFreight) 'Document Freight

            'objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
            ''objFilter.AddEx(frm_OPRT) 'Promotion Template           
            ''objFilter.Add(frm_Project) 'Project

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED)
            'objFilter.Add(frm_Project) 'Project
            objFilter.Add(frm_ITEM_MASTER) 'Item Master

            SetFilter(objFilters)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

#End Region

#End Region

#Region "Data Event"

    Private Sub _SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.FormDataEvent
        Try
            Dim oForm As SAPbouiCOM.Form
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            Select Case BusinessObjectInfo.FormTypeEx
                'Case frm_OSUS
                '    Dim objOSPS As clsSpecialPrice
                '    objOSPS = New clsSpecialPrice
                '    objOSPS.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case frm_Delivery, frm_INVOICES, frm_INVOICESPAYMENT, frm_GI_INVENTORY
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                            If Not BusinessObjectInfo.BeforeAction Then
                                If BusinessObjectInfo.ActionSuccess Then
                                    Utilities.Message("Attempting to Create Journal Entry", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    oApplication.Utilities.post_JournalEntryWithCostCenter(oForm, BusinessObjectInfo.FormTypeEx, BusinessObjectInfo.ObjectKey)
                                End If
                            End If
                    End Select
                Case frm_SaleReturn, frm_ARCreditMemo
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                            If Not BusinessObjectInfo.BeforeAction Then
                                If BusinessObjectInfo.ActionSuccess Then
                                    Utilities.Message("Attempting to Create Journal Entry", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    oApplication.Utilities.update_RMPCosting_Sale_In(oForm, BusinessObjectInfo.FormTypeEx, BusinessObjectInfo.ObjectKey)
                                    oApplication.Utilities.post_JournalEntryWithCostCenter(oForm, BusinessObjectInfo.FormTypeEx, BusinessObjectInfo.ObjectKey)
                                End If
                            End If
                    End Select
                Case frm_PurReturn, frm_APCreditMemo
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                            If Not BusinessObjectInfo.BeforeAction Then
                                If BusinessObjectInfo.ActionSuccess Then
                                    oApplication.Utilities.update_RMPCosting_Purchase_Out(oForm, BusinessObjectInfo.FormTypeEx, BusinessObjectInfo.ObjectKey)
                                End If
                            End If
                    End Select
                Case frm_GRPO, frm_GR_INVENTORY
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                            If Not BusinessObjectInfo.BeforeAction Then
                                If BusinessObjectInfo.ActionSuccess Then
                                    oApplication.Utilities.update_RMPCosting_Purchase_In(oForm, BusinessObjectInfo.FormTypeEx, BusinessObjectInfo.ObjectKey)
                                End If
                            End If
                    End Select
                Case frm_I_Transfer
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                            If Not BusinessObjectInfo.BeforeAction Then
                                If BusinessObjectInfo.ActionSuccess Then
                                    oApplication.Utilities.update_TransferCosting(oForm, BusinessObjectInfo.FormTypeEx, BusinessObjectInfo.ObjectKey)
                                End If
                            End If
                    End Select
                Case frm_ProdReceipt
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                            If Not BusinessObjectInfo.BeforeAction Then
                                If BusinessObjectInfo.ActionSuccess Then
                                    oApplication.Utilities.update_ProductionCosting(oForm, BusinessObjectInfo.FormTypeEx, BusinessObjectInfo.ObjectKey)
                                End If
                            End If
                    End Select
                    'Case frm_IncomingPayment, frm_OutPayment, frm_Deposits
                    '    Select Case BusinessObjectInfo.EventType
                    '        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    '            If Not BusinessObjectInfo.BeforeAction Then
                    '                If BusinessObjectInfo.ActionSuccess Then
                    '                    oApplication.Utilities.post_Commission(oForm, BusinessObjectInfo.FormTypeEx, BusinessObjectInfo.ObjectKey)
                    '                End If
                    '            End If
                    '    End Select
                    'Case frm_OSUS
                    '    Dim objOSPS As clsSpecialPrice
                    '    objOSPS = New clsSpecialPrice
                    '    objOSPS.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                    'Case frm_ORDR
                    '    Dim objOrder As clsOrder
                    '    objOrder = New clsOrder
                    '    objOrder.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                    'Case frm_OPRM
                    '    Dim objPromotion As clsPromotion
                    '    objPromotion = New clsPromotion
                    '    objPromotion.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                    'Case frm_OPRT
                    '    Dim objPromTemp As clsPromotionTemplate
                    '    objPromTemp = New clsPromotionTemplate
                    '    objPromTemp.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            End Select
            'If _Collection.ContainsKey(_FormUID) Then
            '    Dim objform As SAPbouiCOM.Form
            '    objform = oApplication.SBO_Application.Forms.ActiveForm()
            '    If 1 = 1 Then
            '        oMenuObject = _Collection.Item(_FormUID)
            '        oMenuObject.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            '    End If
            'End If
        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "Menu Event"

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case mnu_OSUS
                        oMenuObject = New clsOProjectSetup
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                        'Case mnu_PRJL, mnu_CPRL_O
                        '    oMenuObject = New clsOrder
                        '    oMenuObject.MenuEvent(pVal, BubbleEvent)
                        'Case mnu_PRJL_C, mnu_CPRL_C
                        '    oMenuObject = New clsCustomer
                        '    oMenuObject.MenuEvent(pVal, BubbleEvent)
                        'Case mnu_ComType
                        '    oMenuObject = New clsComType
                        '    oMenuObject.MenuEvent(pVal, BubbleEvent)
                        'Case mnu_COMM_I
                        '    oMenuObject = New clsIncomingPayment
                        '    oMenuObject.MenuEvent(pVal, BubbleEvent)
                        'Case mnu_COMM_O
                        '    oMenuObject = New clsOutGoingPayment
                        '    oMenuObject.MenuEvent(pVal, BubbleEvent)
                        'Case mnu_COMM_D
                        '    oMenuObject = New clsDeposit
                        '    oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_ADD, mnu_FIND, mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS, mnu_ADD_ROW, mnu_DELETE_ROW
                        If IsNothing(_FormUID) Then
                        Else
                            If _Collection.ContainsKey(_FormUID) Then
                                oMenuObject = _Collection.Item(_FormUID)
                                oMenuObject.MenuEvent(pVal, BubbleEvent)
                            End If
                        End If
                        'Case mnu_OPRM
                        '    oMenuObject = New clsPromotion
                        '    oMenuObject.MenuEvent(pVal, BubbleEvent)
                        'Case mnu_OCPR
                        '    oMenuObject = New clsPromotionMapping
                        '    oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_CLOSE
                        If IsNothing(_FormUID) Then

                        Else
                            If _Collection.ContainsKey(_FormUID) Then
                                oMenuObject = _Collection.Item(_FormUID)
                                oMenuObject.MenuEvent(pVal, BubbleEvent)
                            End If
                        End If
                End Select
            Else
                Select Case pVal.MenuUID
                    'Case mnu_DUPLICATE
                    '    Dim oForm As SAPbouiCOM.Form
                    '    Dim oMatrix As SAPbouiCOM.Matrix
                    '    If IsNothing(_FormUID) Then
                    '    Else
                    '        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    '        If ((oForm.TypeEx = frm_Quotation.ToString Or oForm.TypeEx = frm_ORDR.ToString Or oForm.TypeEx = frm_Delivery.ToString Or oForm.TypeEx = frm_SaleReturn.ToString Or oForm.TypeEx = frm_INVOICES.ToString Or oForm.TypeEx = frm_INVOICESPAYMENT.ToString Or oForm.TypeEx = frm_ARCreditMemo.ToString Or oForm.TypeEx = frm_ARReverseInvoice.ToString)) Then
                    '            CType(oForm.Items.Item("_52").Specific, SAPbouiCOM.EditText).Value = ""
                    '        End If
                    '    End If
                    'Case mnu_DELETE_ROW
                    '    Dim oForm As SAPbouiCOM.Form
                    '    If IsNothing(_FormUID) Then
                    '    Else
                    '        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    '        If (oForm.TypeEx = frm_ORDR.ToString) Then
                    '            Dim oMatrix As SAPbouiCOM.Matrix
                    '            oMatrix = oForm.Items.Item("38").Specific
                    '            Dim intRow = oMatrix.GetCellFocus().rowIndex
                    '            If oMatrix.Columns.Item("U_PrRef").Cells().Item(intRow).Specific.value <> "" Then
                    '                If oApplication.SBO_Application.MessageBox("Promotion Items Link for row about to delete?", , "Yes", "No") = 1 Then
                    '                    Dim strRef As String = oMatrix.Columns.Item("U_PrRef").Cells().Item(intRow).Specific.value

                    '                    'Delete Row
                    '                    Dim intRowCount As Integer = oMatrix.RowCount
                    '                    While intRowCount > 0
                    '                        If strRef = oMatrix.Columns.Item("U_PrRef").Cells().Item(intRowCount).Specific.value And CType(oMatrix.Columns.Item("U_IType").Cells().Item(intRowCount).Specific, SAPbouiCOM.ComboBox).Selected.Value = "F" Then
                    '                            oMatrix.DeleteRow(intRowCount)
                    '                        End If
                    '                        intRowCount -= 1
                    '                    End While
                    '                Else
                    '                    BubbleEvent = False
                    '                End If
                    '            End If
                    '        End If
                    '    End If
                End Select
            End If
        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            oMenuObject = Nothing
        End Try
    End Sub

#End Region

#Region "Item Event"

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.ItemEvent

        Try
            _FormUID = FormUID
            If pVal.BeforeAction = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Then
                Select Case pVal.FormTypeEx
                    Case frm_GRPO
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsGRPO
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_OSUS
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsOProjectSetup
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                        'Case frm_Quotation
                        '    If Not _Collection.ContainsKey(FormUID) Then
                        '        oItemObject = New clsQuotation
                        '        oItemObject.FrmUID = FormUID
                        '        _Collection.Add(FormUID, oItemObject)
                        '    End If
                        'Case frm_ORDR
                        '    If Not _Collection.ContainsKey(FormUID) Then
                        '        oItemObject = New clsOrder
                        '        oItemObject.FrmUID = FormUID
                        '        _Collection.Add(FormUID, oItemObject)
                        '    End If
                        'Case frm_Project
                        '    If Not _Collection.ContainsKey(FormUID) Then
                        '        oItemObject = New clsProject
                        '        oItemObject.FrmUID = FormUID
                        '        _Collection.Add(FormUID, oItemObject)
                        '    End If
                        'Case frm_OSUS
                        '    If Not _Collection.ContainsKey(FormUID) Then
                        '        oItemObject = New clsSpecialPrice
                        '        oItemObject.FrmUID = FormUID
                        '        _Collection.Add(FormUID, oItemObject)
                        '    End If
                    Case frm_ProdReceipt
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsProdReceipt
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Delivery
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDelivery
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_INVOICES
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsInvoice
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_INVOICESPAYMENT
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsInvoicePayment
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_SaleReturn
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsReturn
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ARCreditMemo
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsARCreditMemo
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ARReverseInvoice
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsReverseInvoice
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ITEM_MASTER
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsItemMaster
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                        'Case frm_Banking
                        '    If Not _Collection.ContainsKey(FormUID) Then
                        '        oItemObject = New clsBank
                        '        oItemObject.FrmUID = FormUID
                        '        _Collection.Add(FormUID, oItemObject)
                        '    End If
                        'Case frm_ComType
                        '    If Not _Collection.ContainsKey(FormUID) Then
                        '        oItemObject = New clsComType
                        '        oItemObject.FrmUID = FormUID
                        '        _Collection.Add(FormUID, oItemObject)
                        '    End If
                        'Case frm_CommCharges
                        '    If Not _Collection.ContainsKey(FormUID) Then
                        '        oItemObject = New clsCommissionCharges
                        '        oItemObject.FrmUID = FormUID
                        '        _Collection.Add(FormUID, oItemObject)
                        '    End If
                        'Case frm_IncomingPayment
                        '    If Not _Collection.ContainsKey(FormUID) Then
                        '        oItemObject = New clsIncomingPayment
                        '        oItemObject.FrmUID = FormUID
                        '        _Collection.Add(FormUID, oItemObject)
                        '    End If
                        'Case frm_OutPayment
                        '    If Not _Collection.ContainsKey(FormUID) Then
                        '        oItemObject = New clsOutGoingPayment
                        '        oItemObject.FrmUID = FormUID
                        '        _Collection.Add(FormUID, oItemObject)
                        '    End If
                        'Case frm_Deposits
                        '    If Not _Collection.ContainsKey(FormUID) Then
                        '        oItemObject = New clsDeposit
                        '        oItemObject.FrmUID = FormUID
                        '        _Collection.Add(FormUID, oItemObject)
                        '    End If
                        'Case frm_OPRM
                        '    If Not _Collection.ContainsKey(FormUID) Then
                        '        oItemObject = New clsPromotion
                        '        oItemObject.FrmUID = FormUID
                        '        _Collection.Add(FormUID, oItemObject)
                        '    End If
                        'Case frm_OCPR
                        '    If Not _Collection.ContainsKey(FormUID) Then
                        '        oItemObject = New clsPromotionMapping
                        '        oItemObject.FrmUID = FormUID
                        '        _Collection.Add(FormUID, oItemObject)
                        '    End If
                        '    'Case frm_OPRT
                        '    '    If Not _Collection.ContainsKey(FormUID) Then
                        '    '        oItemObject = New clsPromotionTemplate
                        '    '        oItemObject.FrmUID = FormUID
                        '    '        _Collection.Add(FormUID, oItemObject)
                        '    '    End If
                        '    'Case frm_PRT2
                        '    '    If Not _Collection.ContainsKey(FormUID) Then
                        '    '        oItemObject = New clsFreeItems
                        '    '        oItemObject.FrmUID = FormUID
                        '    '        _Collection.Add(FormUID, oItemObject)
                        '    '    End If
                        'Case frm_Freight
                        '    If Not _Collection.ContainsKey(FormUID) Then
                        '        oItemObject = New clsFreight
                        '        oItemObject.FrmUID = FormUID
                        '        _Collection.Add(FormUID, oItemObject)
                        '    End If
                        'Case frm_DocumentFreight
                        '    If Not _Collection.ContainsKey(FormUID) Then
                        '        oItemObject = New clsDocumentFreight
                        '        oItemObject.FrmUID = FormUID
                        '        _Collection.Add(FormUID, oItemObject)
                        '    End If
                    Case frm_GI_INVENTORY
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsIGI
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_I_Transfer
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsITransfer
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                End Select
            End If

            If _Collection.ContainsKey(FormUID) Then
                oItemObject = _Collection.Item(FormUID)
                If oItemObject.IsLookUpOpen And pVal.BeforeAction = True Then
                    _SBO_Application.Forms.Item(oItemObject.LookUpFormUID).Select()
                    BubbleEvent = False
                    Exit Sub
                End If

                Dim oForm As SAPbouiCOM.Form
                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                If (pVal.FormTypeEx = frm_Delivery.ToString Or pVal.FormTypeEx = frm_INVOICES.ToString Or pVal.FormTypeEx = frm_INVOICESPAYMENT.ToString Or pVal.FormTypeEx = frm_SaleReturn.ToString Or pVal.FormTypeEx = frm_ARCreditMemo.ToString) And (pVal.BeforeAction = True And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) Then 'Validate Accounts for Production Costing
                    If 1 = 1 Then
                        If pVal.ItemUID = "1" And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            Dim strMessage As String = String.Empty
                            If Not Utilities.validate_Accounts(oForm, strMessage) Then
                                Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                            End If
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If oForm.TypeEx = frm_Delivery.ToString Then
                                    Dim oMatrix As SAPbouiCOM.Matrix
                                    oMatrix = oForm.Items.Item("38").Specific
                                    For index As Integer = 1 To oMatrix.RowCount
                                        oMatrix.Columns.Item("U_JEDocEty").Cells().Item(index).Specific.value = ""
                                    Next
                                ElseIf (oForm.TypeEx = frm_INVOICES.ToString Or oForm.TypeEx = frm_INVOICESPAYMENT.ToString) Then
                                    Dim oMatrix As SAPbouiCOM.Matrix
                                    oMatrix = oForm.Items.Item("38").Specific
                                    For index As Integer = 1 To oMatrix.RowCount
                                        If oMatrix.Columns.Item("43").Cells().Item(index).Specific.value <> "15" Then
                                            oMatrix.Columns.Item("U_JEDocEty").Cells().Item(index).Specific.value = ""
                                        End If
                                    Next
                                ElseIf (oForm.TypeEx = frm_SaleReturn.ToString Or oForm.TypeEx = frm_ARCreditMemo.ToString) Then
                                    Dim oMatrix As SAPbouiCOM.Matrix
                                    oMatrix = oForm.Items.Item("38").Specific
                                    For index As Integer = 1 To oMatrix.RowCount
                                        If oMatrix.Columns.Item("43").Cells().Item(index).Specific.value = "" Then
                                            oMatrix.Columns.Item("U_JEDocEty").Cells().Item(index).Specific.value = ""
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    End If
                ElseIf (pVal.FormTypeEx = frm_GI_INVENTORY.ToString) And (pVal.BeforeAction = True And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) Then
                    If 1 = 1 Then
                        If pVal.ItemUID = "1" And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            Dim strMessage As String = String.Empty
                            If Not Utilities.validate_Accounts_IGI(oForm, strMessage) Then
                                Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                            End If
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If oForm.TypeEx = frm_GI_INVENTORY.ToString() Then
                                    Dim oMatrix As SAPbouiCOM.Matrix
                                    oMatrix = oForm.Items.Item("13").Specific
                                    For index As Integer = 1 To oMatrix.RowCount
                                        oMatrix.Columns.Item("U_JEDocEty").Cells().Item(index).Specific.value = ""
                                    Next
                                End If
                            End If
                        End If
                    End If
                    'ElseIf ((pVal.FormTypeEx = frm_Quotation.ToString Or pVal.FormTypeEx = frm_ORDR.ToString Or pVal.FormTypeEx = frm_Delivery.ToString Or pVal.FormTypeEx = frm_SaleReturn.ToString Or pVal.FormTypeEx = frm_INVOICES.ToString Or pVal.FormTypeEx = frm_INVOICESPAYMENT.ToString Or pVal.FormTypeEx = frm_ARCreditMemo.ToString Or pVal.FormTypeEx = frm_ARReverseInvoice.ToString) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.ItemUID = "91") Then
                    '    Dim strFRef As String = CType(oForm.Items.Item("_52").Specific, SAPbouiCOM.EditText).Value
                    '    Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("38").Specific
                    '    If oMatrix.RowCount > 0 Then
                    '        If CType(oMatrix.Columns.Item("1").Cells.Item(1).Specific, SAPbouiCOM.EditText).Value <> "" Then
                    '            If strFRef.Length = 0 Then
                    '                oApplication.Utilities.addReference(strFRef, oApplication.Utilities.getDocType(pVal.FormTypeEx))
                    '                If strFRef.Length > 0 Then
                    '                    CType(oForm.Items.Item("_52").Specific, SAPbouiCOM.EditText).Value = strFRef
                    '                End If
                    '            End If

                    '            modVariables.frmFreightRef = strFRef
                    '            modVariables.frmFreightType = pVal.FormTypeEx.ToString()
                    '            'modVariables.frmFreightCurr = CType(oForm.Items.Item("63").Specific, SAPbouiCOM.ComboBox).Selected.Value
                    '        Else
                    '            Utilities.Message("Select Items to Proceed....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            BubbleEvent = False
                    '        End If
                    '    End If
                    'ElseIf ((pVal.FormTypeEx = frm_Quotation.ToString Or pVal.FormTypeEx = frm_ORDR.ToString Or pVal.FormTypeEx = frm_Delivery.ToString Or pVal.FormTypeEx = frm_SaleReturn.ToString Or pVal.FormTypeEx = frm_INVOICES.ToString Or pVal.FormTypeEx = frm_INVOICESPAYMENT.ToString Or pVal.FormTypeEx = frm_ARCreditMemo.ToString Or pVal.FormTypeEx = frm_ARReverseInvoice.ToString) And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) Then
                    '    Dim strFRef As String = CType(oForm.Items.Item("_52").Specific, SAPbouiCOM.EditText).Value
                    '    If strFRef.Length > 0 Then
                    '        oApplication.Utilities.removeFreight(strFRef)
                    '    End If
                    '    Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("38").Specific
                    '    If pVal.FormTypeEx = frm_ORDR.ToString Then
                    '        oApplication.Utilities.removePromotion(oMatrix)
                    '    End If
                Else
                    'If pVal.FormTypeEx = frm_DocumentFreight.ToString Then
                    '    If Not IsNothing(modVariables.frmFreightType) Then
                    '        If modVariables.frmFreightType.Length > 0 Then
                    '            _Collection.Item(FormUID).ItemEvent(FormUID, pVal, BubbleEvent)
                    '        End If
                    '    End If
                    'Else
                    _Collection.Item(FormUID).ItemEvent(FormUID, pVal, BubbleEvent)
                    'End If
                End If
            End If
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD And pVal.BeforeAction = False Then
                If _LookUpCollection.ContainsKey(FormUID) Then
                    oItemObject = _Collection.Item(_LookUpCollection.Item(FormUID))
                    If Not oItemObject Is Nothing Then
                        oItemObject.IsLookUpOpen = False
                    End If
                    _LookUpCollection.Remove(FormUID)
                End If
                If _Collection.ContainsKey(FormUID) Then
                    _Collection.Item(FormUID) = Nothing
                    _Collection.Remove(FormUID)
                End If
            End If
        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

#End Region

#Region "Right Click Event"

    Private Sub _SBO_Application_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.RightClickEvent
        Try
            Dim oForm As SAPbouiCOM.Form
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            'If oForm.TypeEx = frm_ORDR Then
            '    oMenuObject = New clsOrder
            '    oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            'ElseIf (oForm.TypeEx = frm_Customer) Then
            '    oMenuObject = New clsCustomer
            '    oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            'ElseIf (oForm.TypeEx = frm_IncomingPayment) Then
            '    oMenuObject = New clsIncomingPayment
            '    oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            'ElseIf (oForm.TypeEx = frm_OutPayment) Then
            '    oMenuObject = New clsOutGoingPayment
            '    oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            'ElseIf (oForm.TypeEx = frm_Deposits) Then
            '    oMenuObject = New clsDeposit
            '    oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            'End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "Application Event"

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles _SBO_Application.AppEvent
        Try
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                    _Utilities.AddRemoveMenus("RemoveMenus.xml")
                    CloseApp()
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        End Try
    End Sub

#End Region

#Region "Close Application"

    Private Sub CloseApp()
        Try
            If Not _SBO_Application Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_SBO_Application)
            End If

            If Not _Company Is Nothing Then
                If _Company.Connected Then
                    _Company.Disconnect()
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_Company)
            End If

            _Utilities = Nothing
            _Collection = Nothing
            _LookUpCollection = Nothing

            ThreadClose.Sleep(10)
            System.Windows.Forms.Application.Exit()
        Catch ex As Exception
            Throw ex
        Finally
            oApplication = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

#End Region

#Region "Set Application"

    Private Sub SetApplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        Try
            If Environment.GetCommandLineArgs.Length > 1 Then
                sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
                SboGuiApi = New SAPbouiCOM.SboGuiApi
                SboGuiApi.Connect(sConnectionString)
                _SBO_Application = SboGuiApi.GetApplication()
            Else
                Throw New Exception("Connection string missing.")
            End If

        Catch ex As Exception
            Throw ex
        Finally
            SboGuiApi = Nothing
        End Try
    End Sub

#End Region

#Region "Finalize"

    Protected Overrides Sub Finalize()
        Try
            MyBase.Finalize()
            '            CloseApp()

            oMenuObject = Nothing
            oItemObject = Nothing
            oSystemForms = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Addon Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

#End Region

End Class
