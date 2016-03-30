Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsOrder
    Inherits clsBase
    Private oEditText As SAPbouiCOM.EditText
    Private oRecordSet As SAPbobsCOM.Recordset
    Private strQuery As String
    Private oMatrix As SAPbouiCOM.Matrix

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
                    oForm.Items.Item("38").Enabled = True
                    oForm.Items.Item("_2").Enabled = True
                Case mnu_PRJL
                    If Not oForm.Items.Item("4").Specific.value = "" Then
                        Dim objSpecialPriceList As clsSpecialPriceList
                        objSpecialPriceList = New clsSpecialPriceList
                        objSpecialPriceList.LoadForm(oForm.Items.Item("4").Specific.value, "")
                    Else
                        oApplication.Utilities.Message("Select Customer to Get Special Price List...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                    If oMenuItem.SubMenus.Exists(pVal.MenuUID) Then
                        oApplication.SBO_Application.Menus.RemoveEx(pVal.MenuUID)
                    End If
                Case mnu_CPRL_O
                    If Not oForm.Items.Item("4").Specific.value = "" Then
                        Dim objPromList As clsCustPromotionList
                        objPromList = New clsCustPromotionList
                        objPromList.LoadForm(oForm.Items.Item("4").Specific.value)
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
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_ORDR Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("38").Specific
                                If pVal.ItemUID = "38" And pVal.ColUID = "31" Then
                                    filterProjectChooseFromList(oForm, oMatrix.Columns.Item("31").ChooseFromListUID)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK, SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("38").Specific
                                If pVal.ItemUID = "38" And (pVal.ColUID = "U_PrmApp" Or pVal.ColUID = "U_PrCode" Or pVal.ColUID = "U_SPDocEty" Or pVal.ColUID = "U_PrRef" Or pVal.ColUID = "U_PrLine") Then
                                    BubbleEvent = False
                                    Exit Sub
                                ElseIf pVal.ItemUID = "38" And pVal.Row > 0 Then
                                    If CType(oMatrix.Columns.Item("U_PrmApp").Cells().Item(pVal.Row).Specific, SAPbouiCOM.ComboBox).Value = "Y" Then
                                        If pVal.ColUID = "1" Or pVal.ColUID = "2" Or pVal.ColUID = "3" Or pVal.ColUID = "11" Then
                                            oApplication.Utilities.Message("Promotion details already applied for this Row. You should delete the  .", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                        End If
                                    ElseIf (CType(oMatrix.Columns.Item("U_PrCode").Cells().Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.Trim().Length > 0) Then
                                        If pVal.ColUID = "1" Or pVal.ColUID = "2" Or pVal.ColUID = "3" Or pVal.ColUID = "11" Or pVal.ColUID = "14" Or pVal.ColUID = "15" Or pVal.ColUID = "21" Then
                                            oApplication.Utilities.Message("Promotion details already applied for this Row. You should delete the  .", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                        End If
                                    ElseIf pVal.ItemUID = "38" And (pVal.ColUID = "14" Or pVal.ColUID = "15" Or pVal.ColUID = "17" Or pVal.ColUID = "21") Then
                                        If oApplication.Utilities.getMatrixValues(oMatrix, "31", pVal.Row) <> "" And oApplication.Utilities.getMatrixValues(oMatrix, "U_SPDocEty", pVal.Row) <> "" Then
                                            oApplication.Utilities.Message("Special Price is Linked to Selected Row...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "_2" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    oMatrix = oForm.Items.Item("38").Specific
                                    If oForm.Items.Item("12").Specific.value.ToString.Length = 0 Then
                                        oApplication.Utilities.Message("Please Enter Delivery Date to Apply Promotion...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    ElseIf oMatrix.RowCount = 1 Then
                                        oApplication.Utilities.Message("Add Items To Apply Promotion...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If

                                    If oApplication.SBO_Application.MessageBox("No Possible to Change Items When Promotion Applied Want to Continue?", , "Yes", "No") = 2 Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        oForm.Freeze(True)
                                        applyPromotion(oForm)
                                        oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        oForm.Freeze(False)
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                initializeControls(oForm)
                                oApplication.Utilities.AddFreightControls(oForm, "ORDR")
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    If pVal.ItemUID = "38" And (pVal.ColUID = "31" Or pVal.ColUID = "1") Then
                                        If pVal.CharPressed = "9" Then
                                            oForm.Freeze(True)
                                            changePrice(oForm, pVal.Row)
                                            oForm.Freeze(False)
                                        End If
                                    ElseIf (pVal.ItemUID = "4" Or pVal.ItemUID = "46") Then
                                        If pVal.CharPressed = "9" Then
                                            oForm.Freeze(True)
                                            changePrice(oForm)
                                            oForm.Freeze(False)
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    If pVal.ItemUID = "38" And (pVal.ColUID = "1" Or pVal.ColUID = "31") Then
                                        oForm.Freeze(True)
                                        changePrice(oForm, pVal.Row)
                                        oForm.Freeze(False)
                                    ElseIf (pVal.ItemUID = "4" Or pVal.ItemUID = "46") Then
                                        oForm.Freeze(True)
                                        changePrice(oForm)
                                        oForm.Freeze(False)
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim sCHFL_ID, val As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    Dim oDataTable As SAPbouiCOM.DataTable
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If pVal.ColUID = "31" And pVal.ItemUID = "38" Then
                                        If IsNothing(oCFLEvento.SelectedObjects) Then
                                            val = ""
                                        Else
                                            val = oDataTable.GetValue("PrjCode", 0)
                                        End If

                                        oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                        If val = "" Then
                                            Try
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "U_SPDocEty", pVal.Row, "")
                                            Catch ex As Exception
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "U_SPDocEty", pVal.Row, "")
                                            End Try
                                        Else
                                            changePrice(oForm, pVal.Row, val)
                                        End If
                                        If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                            If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception

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
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            'oForm.Items.Item("_2").Enabled = False
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
            If oForm.TypeEx = frm_ORDR Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                If (eventInfo.BeforeAction = True) Then
                    Try
                        'Project List
                        If Not oMenuItem.SubMenus.Exists(mnu_PRJL) Then
                            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                            oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                            oCreationPackage.UniqueID = mnu_PRJL
                            oCreationPackage.String = "Special Price List"
                            oCreationPackage.Enabled = True
                            oMenus = oMenuItem.SubMenus
                            oMenus.AddEx(oCreationPackage)
                        End If

                        'Promotion List
                        If Not oMenuItem.SubMenus.Exists(mnu_CPRL_O) Then
                            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                            oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                            oCreationPackage.UniqueID = mnu_CPRL_O
                            oCreationPackage.String = "Promotion List"
                            oCreationPackage.Enabled = True
                            oMenus = oMenuItem.SubMenus
                            oMenus.AddEx(oCreationPackage)
                        End If
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                Else
                    If oMenuItem.SubMenus.Exists(mnu_PRJL) Then
                        oApplication.SBO_Application.Menus.RemoveEx(mnu_PRJL)
                    End If

                    If oMenuItem.SubMenus.Exists(mnu_CPRL_O) Then
                        oApplication.SBO_Application.Menus.RemoveEx(mnu_CPRL_O)
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
            oApplication.Utilities.AddControls(oForm, "_2", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", 0, 0, "2", "Apply Promotion", 0, 0, 0, False)
            oForm.Items.Item("_2").Width = "140"
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub changePrice(ByVal oForm As SAPbouiCOM.Form)
        Try
            oMatrix = oForm.Items.Item("38").Specific
            Dim strItemCode, strProject, strPrice, strCustomer, strDocDate, strDocEntry, strDisType As String
            Dim dblDiscount As Double

            For intRow As Integer = 1 To oMatrix.RowCount
                strItemCode = oMatrix.Columns.Item("1").Cells().Item(intRow).Specific.value
                strProject = oMatrix.Columns.Item("31").Cells().Item(intRow).Specific.value
                strCustomer = oForm.Items.Item("4").Specific.Value
                strDocDate = oForm.Items.Item("10").Specific.Value
                If strProject.Length > 0 Then
                    getSpecialPrice(oForm, strCustomer, strItemCode, strProject, strDocDate, strDocEntry, strPrice, strDisType, dblDiscount)
                    If strPrice <> "" Then
                        If strDisType = "D" Then
                            oMatrix.Columns.Item("15").Cells().Item(intRow).Specific.value = dblDiscount
                        Else
                            oMatrix.Columns.Item("15").Cells().Item(intRow).Specific.value = "0"
                            oMatrix.Columns.Item("14").Cells().Item(intRow).Specific.value = strPrice
                        End If
                        oMatrix.Columns.Item("U_SPDocEty").Cells().Item(intRow).Specific.value = strDocEntry
                    End If
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub changePrice(ByVal oForm As SAPbouiCOM.Form, ByVal intRow As Integer, Optional ByVal aProjectCode As String = "")
        Try
            oMatrix = oForm.Items.Item("38").Specific
            Dim strItemCode, strProject, strPrice, strCustomer, strDocDate, strDocEntry, strDisType As String
            Dim dblDiscount As Double

            strItemCode = oMatrix.Columns.Item("1").Cells().Item(intRow).Specific.value
            If aProjectCode = "" Then
                strProject = oMatrix.Columns.Item("31").Cells().Item(intRow).Specific.value
            Else
                strProject = aProjectCode
            End If

            strCustomer = oForm.Items.Item("4").Specific.Value
            strDocDate = oForm.Items.Item("10").Specific.Value
            getSpecialPrice(oForm, strCustomer, strItemCode, strProject, strDocDate, strDocEntry, strPrice, strDisType, dblDiscount)
            If strPrice <> "" Then
                If strDisType = "D" Then
                    oMatrix.Columns.Item("15").Cells().Item(intRow).Specific.value = dblDiscount
                Else
                    oMatrix.Columns.Item("15").Cells().Item(intRow).Specific.value = "0"
                    oMatrix.Columns.Item("14").Cells().Item(intRow).Specific.value = strPrice
                End If
                oMatrix.Columns.Item("U_SPDocEty").Cells().Item(intRow).Specific.value = strDocEntry
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getSpecialPrice(ByVal oForm As SAPbouiCOM.Form, ByVal strCustomer As String, ByVal strItemCode As String, ByVal strProject As String, _
ByVal strDocDate As String, ByRef strDE As String, ByRef strPrice As String, ByRef strDisType As String, ByRef dblDicount As Double)
        Try
            Dim _retVal As String
            Dim oSPRecordSet As SAPbobsCOM.Recordset
            oSPRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select U_Discount,U_DisType,T0.U_Currency + ' ' + Convert(VarChar,IsNull(U_DisPrice,0)) As 'SP',T1.DocEntry From [@PSP1] T0 Join [@OPSP] T1 On T0.DocEntry = T1.DocEntry Join OPRJ T2 On T1.U_PrjCode = T2.PrjCode Where T2.U_CardCode = '" + strCustomer + "' And T0.U_ItmCode = '" + strItemCode + "' And T2.PrjCode =  '" + strProject + "' And Convert(Varchar(8),T1.U_EffFrom,112) <=  '" + strDocDate + "' And Convert(VarChar(8),T1.U_EffTo,112) >=  '" + strDocDate + "' And ISNULL(Status,'O') = 'O'"
            oSPRecordSet.DoQuery(strQuery)
            If Not oSPRecordSet.EoF Then
                _retVal = oSPRecordSet.Fields.Item("SP").Value
                strPrice = _retVal
                strDE = oSPRecordSet.Fields.Item("DocEntry").Value
                dblDicount = oSPRecordSet.Fields.Item("U_Discount").Value
                strDisType = oSPRecordSet.Fields.Item("U_DisType").Value
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub applyPromotion(ByVal oForm As SAPbouiCOM.Form)
        Try
            oMatrix = oForm.Items.Item("38").Specific
            Dim strItemCode As String
            Dim dblQty As Double
            Dim strCustomer As String
            Dim strDocDate As String
            Dim strDocEntry As String = String.Empty
            Dim strStatus As String = String.Empty

            strCustomer = oForm.Items.Item("4").Specific.Value
            strDocDate = oForm.Items.Item("12").Specific.Value

            'Delete Promotion Items if Line Status is Open
            Dim intRowCount As Integer = oMatrix.RowCount
            While intRowCount >= 1
                strStatus = oMatrix.Columns.Item("40").Cells().Item(intRowCount).Specific.value
                If strStatus = "O" Then
                    If CType(oMatrix.Columns.Item("U_PrCode").Cells().Item(intRowCount).Specific, SAPbouiCOM.EditText).Value.Trim().Length > 0 Then
                        oMatrix.DeleteRow(intRowCount)
                    End If
                End If
                intRowCount -= 1
            End While

            oForm.Refresh()
            For intRow As Integer = 1 To oMatrix.RowCount - 1
                strItemCode = oMatrix.Columns.Item("1").Cells().Item(intRow).Specific.value
                dblQty = oMatrix.Columns.Item("11").Cells().Item(intRow).Specific.value
                getFreeOfGoods(oForm, strCustomer, strDocDate, strItemCode, dblQty, intRow, strStatus)
            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getFreeOfGoods(ByVal oForm As SAPbouiCOM.Form, ByVal strCustomer As String, ByVal strDocDate As String, ByVal strItemCode As String, _
  ByRef dblQuantity As Double, ByVal intRow As Integer, ByVal strStatus As String)
        Try
            oMatrix = oForm.Items.Item("38").Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select Top 1 U_OffCode,U_OQty,U_ODis,U_PrCode,Code From [@OCPR] "
            strQuery += " Where '" & strDocDate & "' Between Convert(VarChar(8),U_EffFrom,112) And Convert(VarChar(8),U_EffTo,112) "
            strQuery += " And U_CustCode = '" & strCustomer & "'"
            strQuery += " And U_ItmCode = '" & strItemCode & "'"
            strQuery += " And U_Qty <= '" & dblQuantity & "'"
            strQuery += " Order By U_Qty Desc,Code Desc"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oMatrix.AddRow(1, oMatrix.RowCount)

                Try
                    Dim strRef As String = String.Empty
                    oApplication.Utilities.addPromotionReference(strRef)

                    'Regular Item
                    CType(oMatrix.Columns.Item("U_PrmApp").Cells().Item(intRow).Specific, SAPbouiCOM.ComboBox).Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    CType(oMatrix.Columns.Item("U_IType").Cells().Item(intRow).Specific, SAPbouiCOM.ComboBox).Select("R", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    oMatrix.Columns.Item("U_PrRef").Cells().Item(intRow).Specific.value = strRef

                    'Free Item
                    oMatrix.Columns.Item("1").Cells().Item(oMatrix.RowCount - 1).Specific.value = oRecordSet.Fields.Item("U_OffCode").Value
                    oMatrix.Columns.Item("11").Cells().Item(oMatrix.RowCount - 1).Specific.value = oRecordSet.Fields.Item("U_OQty").Value
                    oMatrix.Columns.Item("15").Cells().Item(oMatrix.RowCount - 1).Specific.value = oRecordSet.Fields.Item("U_ODis").Value
                    oMatrix.Columns.Item("U_PrCode").Cells().Item(oMatrix.RowCount - 1).Specific.value = oRecordSet.Fields.Item("U_PrCode").Value
                    oMatrix.Columns.Item("U_PrRef").Cells().Item(oMatrix.RowCount - 1).Specific.value = strRef
                    CType(oMatrix.Columns.Item("U_IType").Cells().Item(oMatrix.RowCount - 1).Specific, SAPbouiCOM.ComboBox).Select("F", SAPbouiCOM.BoSearchKey.psk_ByValue)

                Catch ex As Exception

                End Try
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub filterProjectChooseFromList(ByVal oForm As SAPbouiCOM.Form, ByVal strCFLID As String)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList

            oCFLs = oForm.ChooseFromLists
            oCFL = oCFLs.Item(strCFLID)

            oCons = oCFL.GetConditions()

            If oCons.Count = 0 Then
                oCon = oCons.Add()
            Else
                oCon = oCons.Item(0)
            End If

            oCon.Alias = "U_CardCode"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = oForm.Items.Item("4").Specific.value
            oCFL.SetConditions(oCons)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
