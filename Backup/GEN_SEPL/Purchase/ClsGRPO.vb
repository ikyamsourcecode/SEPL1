Public Class ClsGRPO

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm, objSubForm, objSForm As SAPbouiCOM.Form
    Dim objItem, objOldItem, TempItem As SAPbouiCOM.Item
    Dim objMatrix, objSubMatrix, objBatchMatrix As SAPbouiCOM.Matrix
    Dim SZDBHead As SAPbouiCOM.DBDataSource
    Dim SZDBDetail As SAPbouiCOM.DBDataSource
    Dim SMDBHead As SAPbouiCOM.DBDataSource
    Dim SMDBDetail As SAPbouiCOM.DBDataSource
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim RS1, RS2 As SAPbobsCOM.Recordset
    Dim ModalForm As Boolean = False
    Dim ChildModalForm As Boolean = False
    Dim PARENT_FORM As String
    Dim RowNo As Integer
    Dim orderno, hwid As String
    Dim sorderno, shwid, sitemcode As String
    Dim Mode As Integer
    Dim TotQty As Double
    Dim GSONO, GMACID, GITEMCODE As String
    Dim RowID As Integer
    Dim DeleteItemCode As String
    Dim objBtnCmb As SAPbouiCOM.ButtonCombo
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            objOldItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("insstat", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Specific.databind.setbound(True, "OPDN", "u_insstat")
            objItem.LinkTo = "46"
            objItem.Specific.TabOrder = objOldItem.Specific.TabOrder + 1
            objItem.DisplayDesc = True
            objItem.Visible = False
            objOldItem = objForm.Items.Item("2")
            objItem = objForm.Items.Add("btnit", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Left = objOldItem.Left + objOldItem.Width + 5
            objItem.Width = objOldItem.Width + 25
            objItem.Height = objOldItem.Height
            objItem.Top = objOldItem.Top
            objItem.Specific.caption = "Move to Main Wh"
            objItem.LinkTo = "2"
            objItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objOldItem = objForm.Items.Item("10000329")
            objItem = objForm.Items.Add("btncmb", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO)
            objItem.Left = objOldItem.Left
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Top = objOldItem.Top
            objItem.Specific.caption = "Copy To"
            objBtnCmb = objItem.Specific
            objBtnCmb.ValidValues.Add("1", "A/P Invoice")
            objBtnCmb.ValidValues.Add("2", "Goods Return")
            objItem.LinkTo = "10000329"
            objOldItem.Visible = False
            objItem.AffectsFormMode = False
            objForm = oApplication.Forms.Item(FormUID)
            TempItem = objForm.Items.Item("86")
            objItem = objForm.Items.Add("spc", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = TempItem.Top + TempItem.Height + 20
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.Caption = "Unit"
            objItem.LinkTo = "86"
            TempItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("cpc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = TempItem.Top + TempItem.Height + 20
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.databind.setbound(True, "OPDN", "u_unit")
            objItem.Specific.TabOrder = TempItem.Specific.TabOrder + 1
            objItem.LinkTo = "46"
            objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    If pVal.BeforeAction = True Then
                        If pVal.FormTypeCount = 1 Then
                            Me.CreateForm(FormUID)
                        Else
                            BubbleEvent = False
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "btncmb" And pVal.BeforeAction = False Then
                        Dim oCombo As SAPbouiCOM.ButtonCombo
                        Dim TRGTCombo As SAPbouiCOM.ComboBox
                        oCombo = objForm.Items.Item("btncmb").Specific
                        TRGTCombo = objForm.Items.Item("10000329").Specific
                        If Trim(oCombo.Selected.Value) = "1" Then
                            If Trim(objForm.Items.Item("insstat").Specific.selected.value) <> "Closed" Then
                                oApplication.StatusBar.SetText("The inspection has to be done before you can create A/P Invoice", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                            Else
                                If objForm.Items.Item("10000329").Enabled = True Then
                                    TRGTCombo.Select("A/P Invoice", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                End If
                            End If
                        End If
                        If Trim(oCombo.Selected.Value) = "2" Then
                            If objForm.Items.Item("10000329").Enabled = True Then
                                TRGTCombo.Select("G. Return", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            End If
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If pVal.ItemUID = "cpc" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = True Then
                        If Trim(objForm.Items.Item("cpc").Specific.value) <> "" Then
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("Select u_unit From [@GEN_USR_UNIT] Where u_user = '" + Trim(oCompany.UserName.ToString) + "' And u_unit = '" + Trim(objForm.Items.Item("cpc").Specific.value) + "'")
                            If oRecordSet.RecordCount = 0 Then
                                oApplication.StatusBar.SetText("Please select the correct Unit for the user", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    If pVal.BeforeAction = True Then
                        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery("Select u_unit From [@GEN_USR_UNIT] Where u_user = '" + oCompany.UserName.ToString.Trim + "'")
                        If oRecordSet.RecordCount = 0 Then
                            oApplication.StatusBar.SetText("No PC assigned to User", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                    If pVal.BeforeAction = False Then
                        Dim objForm As SAPbouiCOM.Form
                        objForm = oApplication.Forms.Item(FormUID)
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                        Dim CFL_Id As String
                        CFL_Id = CFLEvent.ChooseFromListUID
                        oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                        Dim oDT As SAPbouiCOM.DataTable
                        oDT = CFLEvent.SelectedObjects
                        If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            If oCFL.UniqueID = "2" Then
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet.DoQuery("Select u_unit From OCRD Where CardCode = '" + Trim(oDT.GetValue("CardCode", 0)) + "'")
                                objForm.Items.Item("cpc").Specific.value = oRSet.Fields.Item("u_unit").Value
                                objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            End If
                            If oCFL.UniqueID = "3" Then
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet.DoQuery("Select u_unit From OCRD Where CardCode = '" + Trim(oDT.GetValue("CardCode", 0)) + "'")
                                objForm.Items.Item("cpc").Specific.value = oRSet.Fields.Item("u_unit").Value
                                objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            End If
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = True Then
                        If Trim(objForm.Items.Item("cpc").Specific.value) = "" Then
                            oApplication.StatusBar.SetText("Please select Unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                        objMatrix = objForm.Items.Item("38").Specific
                        Dim ErrFlag As Boolean = False
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        For i As Integer = 1 To objMatrix.VisualRowCount
                            If Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) <> "" Then
                                oRSet.DoQuery("Select OpenQty From POR1 WHere DocEntry = '" + Trim(objMatrix.Columns.Item("45").Cells.Item(i).Specific.value) + "' And LineNum = '" + Trim(objMatrix.Columns.Item("46").Cells.Item(i).Specific.Value) + "'")
                                If oRSet.RecordCount > 0 Then
                                    If CDbl(objMatrix.Columns.Item("11").Cells.Item(i).Specific.value) < CDbl(oRSet.Fields.Item("OpenQty").Value) Then
                                        ErrFlag = True
                                    End If
                                End If
                            End If
                        Next
                        If ErrFlag = True Then
                            If oApplication.MessageBox("GRPO quantity is less than PO quantity. Do you still want to continue? ", 2, "Yes", "No") = 2 Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If
                    If pVal.ItemUID = "btnit" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If pVal.BeforeAction = True Then
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select Distinct A.DocNum,B.ItemCode,B.WhsCode,B.Quantity From OPDN A Inner Join PDN1 B On A.DocEntry = B.DocEntry Where A.DocNum = '" + Trim(objForm.Items.Item("8").Specific.value) + "' And IsNull(B.u_insstat,'Open') = 'Open'")
                            If oRSet.RecordCount = 0 Then
                                oApplication.StatusBar.SetText("Items already moved to main warehouse for this GRN", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                        If pVal.BeforeAction = False Then
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select Distinct A.DocNum,B.ItemCode,B.WhsCode,(B.Quantity * B.NumPerMsr) - IsNull(B.u_openqty,0) AS 'Qty',B.LineNum From OPDN A Inner Join PDN1 B On A.DocEntry = B.DocEntry Where A.DocNum = '" + Trim(objForm.Items.Item("8").Specific.value) + "' And IsNull(B.u_insstat,'Open') = 'Open'")
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            If oRSet.RecordCount > 0 Then
                                oApplication.ActivateMenuItem("3080")
                                Dim ITForm As SAPbouiCOM.Form
                                Dim ITMatrix As SAPbouiCOM.Matrix
                                ITForm = oApplication.Forms.GetForm("940", oApplication.Forms.ActiveForm.TypeCount)
                                ITMatrix = ITForm.Items.Item("23").Specific
                                ITForm.Items.Item("grnno").Specific.value = oRSet.Fields.Item("DocNum").Value
                                ITForm.Items.Item("18").Specific.value = oRSet.Fields.Item("WhsCode").Value
                                ITMatrix.Columns.Item("U_grnno").Editable = True
                                ITMatrix.Columns.Item("U_grnlnid").Editable = True
                                For i As Integer = 1 To oRSet.RecordCount
                                    Try
                                        ITMatrix.Columns.Item("1").Cells.Item(i).Specific.value = oRSet.Fields.Item("ItemCode").Value
                                        RS.DoQuery("Select DfltWh From OITM Where ItemCode = '" + Trim(oRSet.Fields.Item("ItemCode").Value) + "'")
                                        If RS.RecordCount > 0 Then
                                            ITMatrix.Columns.Item("5").Cells.Item(i).Specific.value = RS.Fields.Item("DfltWh").Value
                                        End If
                                        If oRSet.Fields.Item("Qty").Value > 0 Then
                                            ITMatrix.Columns.Item("U_BAL_QTY").Editable = True
                                            ITMatrix.Columns.Item("U_BAL_QTY").Cells.Item(i).Specific.value = oRSet.Fields.Item("Qty").Value
                                            ITMatrix.Columns.Item("10").Cells.Item(i).Specific.value = oRSet.Fields.Item("Qty").Value

                                            ITMatrix.Columns.Item("U_BAL_QTY").Editable = False
                                        Else
                                            ITMatrix.Columns.Item("10").Cells.Item(i).Specific.value = 1
                                        End If
                                        ITMatrix.Columns.Item("U_grnno").Cells.Item(i).Specific.value = oRSet.Fields.Item("DocNum").Value
                                        ITMatrix.Columns.Item("U_grnlnid").Cells.Item(i).Specific.value = oRSet.Fields.Item("LineNum").Value
                                        ITMatrix.Columns.Item("1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        oRSet.MoveNext()
                                    Catch ex As Exception
                                        oApplication.StatusBar.SetText(ex.Message)
                                    End Try
                                Next
                                ITMatrix.Columns.Item("U_grnno").Editable = False
                                ITMatrix.Columns.Item("U_grnlnid").Editable = False
                            End If
                        End If
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = True Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "1288", "1289", "1290", "1291"
                        If objForm.TypeEx = "143" Then
                            BubbleEvent = False
                        End If
                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    End If
            End Select
        Catch ex As Exception

        End Try
    End Sub

End Class
