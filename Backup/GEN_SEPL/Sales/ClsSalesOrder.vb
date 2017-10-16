Public Class ClsSalesOrder

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm, objSubForm, objSForm As SAPbouiCOM.Form
    Dim objItem, objOldItem As SAPbouiCOM.Item
    Dim objMatrix, objSubMatrix, objBatchMatrix As SAPbouiCOM.Matrix
    Dim oCombo As SAPbouiCOM.ComboBox
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
    Dim GSONO, GMACID, GITEMCODE, GASRTCODE As String
    Dim DeleteItemCode As String
    Dim RowID As Integer
    Dim SONO As String
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            objOldItem = objForm.Items.Item("86")
            objItem = objForm.Items.Add("spc", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = objOldItem.Top + objOldItem.Height + 20
            objItem.Left = objOldItem.Left
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.Caption = "Unit"
            objItem.LinkTo = "86"
            objOldItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("cpc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = objOldItem.Top + objOldItem.Height + 20
            objItem.Left = objOldItem.Left
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.databind.setbound(True, "ORDR", "u_unit")
            objItem.Specific.TabOrder = objOldItem.Specific.TabOrder + 1
            objItem.LinkTo = "46"
            objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objOldItem = objForm.Items.Item("2")
            objItem = objForm.Items.Add("btnsize", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width + 20
            objItem.Left = objOldItem.Left + objOldItem.Width + 5
            objItem.Top = objOldItem.Top
            objItem.Specific.caption = "Allocate Size"
            objItem.LinkTo = "2"
            objOldItem = objForm.Items.Item("86")
            objItem = objForm.Items.Add("sseason", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Specific.Caption = "Season"
            objItem.LinkTo = "86"
            objOldItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("tseason", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Specific.databind.setbound(True, "ORDR", "u_season")
            objItem.LinkTo = "46"
            objItem.Specific.TabOrder = objOldItem.Specific.TabOrder + 1
            objItem.DisplayDesc = True
            ObjOldItem = objForm.Items.Item("15")
            objItem = objForm.Items.Add("sdoccur", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = ObjOldItem.Top + ObjOldItem.Height + 15
            objItem.Left = ObjOldItem.Left
            objItem.Width = ObjOldItem.Width
            objItem.Height = ObjOldItem.Height
            objItem.Specific.Caption = "Buyer Currency"
            objItem.LinkTo = "15"
            ObjOldItem = objForm.Items.Item("15")
            objItem = objForm.Items.Add("sdocrate", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = ObjOldItem.Top + ObjOldItem.Height + 30
            objItem.Left = ObjOldItem.Left
            objItem.Width = ObjOldItem.Width
            objItem.Height = ObjOldItem.Height
            objItem.Specific.Caption = "Document Rate"
            objItem.LinkTo = "15"
            ObjOldItem = objForm.Items.Item("14")
            objItem = objForm.Items.Add("doccur", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = ObjOldItem.Top + ObjOldItem.Height + 15
            objItem.Left = ObjOldItem.Left
            objItem.Width = ObjOldItem.Width
            objItem.Height = ObjOldItem.Height
            objItem.Specific.databind.setbound(True, "ORDR", "u_doccur")
            objItem.Specific.TabOrder = ObjOldItem.Specific.TabOrder + 1
            objItem.LinkTo = "14"
            ObjOldItem = objForm.Items.Item("14")
            objItem = objForm.Items.Add("docrate", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = ObjOldItem.Top + ObjOldItem.Height + 30
            objItem.Left = ObjOldItem.Left
            objItem.Width = ObjOldItem.Width
            objItem.Height = ObjOldItem.Height
            objItem.Specific.databind.setbound(True, "ORDR", "u_docrate")
            objItem.Specific.TabOrder = ObjOldItem.Specific.TabOrder + 1
            objItem.LinkTo = "14"
            objForm.Items.Item("doccur").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("docrate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objOldItem = objForm.Items.Item("btnsize")
            objItem = objForm.Items.Add("btn", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Top = ObjOldItem.Top
            objItem.Left = ObjOldItem.Left + ObjOldItem.Width + 10
            objItem.Width = objOldItem.Width + 30
            objItem.Height = objOldItem.Height
            objItem.Specific.caption = "Generate FC value"
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Dim CHILD_FORM As String = "SOORD@" & pVal.FormUID
            Dim ModalForm As Boolean = False
            For i As Integer = 0 To oApplication.Forms.Count - 1
                If oApplication.Forms.Item(i).UniqueID = (CHILD_FORM) Then
                    objSubForm = oApplication.Forms.Item("SOORD@" & pVal.FormUID)
                    ModalForm = True
                    Exit For
                End If
            Next
            If ModalForm = False Then
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
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        If (pVal.ItemUID = "10" Or pVal.ItemUID = "doccur") And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = False Then
                            Try
                                If Trim(objForm.Items.Item("doccur").Specific.value) <> "" Then
                                    Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRSet.DoQuery("Select Rate From ORTT Where RateDate = '" + Trim(objForm.Items.Item("10").Specific.value) + "' And Currency = '" + Trim(objForm.Items.Item("doccur").Specific.value) + "'")
                                    objForm.Items.Item("docrate").Specific.value = oRSet.Fields.Item("Rate").Value
                                Else
                                    objForm.Items.Item("docrate").Specific.value = ""
                                End If
                            Catch ex As Exception
                                oApplication.StatusBar.SetText(ex.Message)
                            End Try
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
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        If pVal.ItemUID = "38" And pVal.BeforeAction = True Then
                            objMatrix = objForm.Items.Item("38").Specific
                            If pVal.Row > 0 And pVal.Row <= objMatrix.VisualRowCount Then
                                RowID = pVal.Row
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                        If pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            RS.DoQuery("Delete From [@GEN_SZ_ORDR] Where u_sono = '" + Trim(objForm.Items.Item("8").Specific.value) + "' And u_macid = '" + MAC_ID + "'")
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                            SONO = objForm.Items.Item("8").Specific.Value
                        End If
                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = True Then
                            If Trim(objForm.Items.Item("cpc").Specific.value) = "" Then
                                oApplication.StatusBar.SetText("Please select Unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                        If pVal.ItemUID = "btn" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            objMatrix = objForm.Items.Item("38").Specific
                            If pVal.BeforeAction = True Then
                                If Trim(objForm.Items.Item("doccur").Specific.value) = "" Or Trim(objForm.Items.Item("docrate").Specific.value) = 0 Then
                                    oApplication.StatusBar.SetText("Please select appropriate buyer currency and rate", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                End If
                            Else
                                For i As Integer = 1 To objMatrix.VisualRowCount
                                    If Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) <> "" And Trim(objMatrix.Columns.Item("14").Cells.Item(i).Specific.value) <> "" Then
                                        objMatrix.Columns.Item("U_pricefc").Cells.Item(i).Specific.value = CDbl(objMatrix.Columns.Item("14").Cells.Item(i).Specific.value.ToString.Substring(3)) / objForm.Items.Item("docrate").Specific.value
                                        objMatrix.Columns.Item("U_totalfc").Cells.Item(i).Specific.value = (CDbl(objMatrix.Columns.Item("14").Cells.Item(i).Specific.value.ToString.Substring(3)) / objForm.Items.Item("docrate").Specific.value) * objMatrix.Columns.Item("11").Cells.Item(i).Specific.value
                                    End If
                                Next
                            End If
                        End If
                        If pVal.ItemUID = "btnsize" Then
                            If pVal.BeforeAction = True Then
                                objMatrix = objForm.Items.Item("38").Specific
                                If objMatrix.VisualRowCount < 1 Then
                                    oApplication.StatusBar.SetText("Please select items", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                Else
                                    Mode = pVal.FormMode
                                    Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                                    Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim RS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    RS1.DoQuery("Delete From TMP_ORDR_ITEMS Where ordrno = '" + Trim(objForm.Items.Item("8").Specific.Value) + "' And macid = '" + MAC_ID + "'")
                                    'RS.DoQuery("Select DocEntry From [@GEN_SIZE_ORDR] Where u_ordrno = '" + Trim(objForm.Items.Item("8").Specific.value) + "' And u_macid = '" + MAC_ID + "'")
                                    ''RS1.DoQuery("Delete From [@GEN_SIZE_ORDR_D0] Where DocEntry = '" + Trim(RS.Fields.Item("DocEntry").Value) + "'")
                                    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                                        If Trim(objMatrix.Columns.Item("U_asrtcode").Cells.Item(i).Specific.value) = "" Then
                                            oApplication.StatusBar.SetText("Please select Assorted code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    Next
                                    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                                        oRecordSet.DoQuery("Insert Into TMP_ORDR_ITEMS(ordrno,itemcode,asrtcode,qty,macid) Values('" + Trim(objForm.Items.Item("8").Specific.value) + "','" + Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) + "','" + Trim(objMatrix.Columns.Item("U_asrtcode").Cells.Item(i).Specific.value) + "','" + Trim(objMatrix.Columns.Item("11").Cells.Item(i).Specific.value) + "','" + MAC_ID + "')")
                                    Next
                                    If objForm.Items.Item("81").Specific.selected.value = "1" Or objForm.Items.Item("81").Specific.selected.value = "2" Then
                                        Me.Open_Order_Allocation_Form(pVal.FormUID, Trim(objForm.Items.Item("8").Specific.value), MAC_ID)
                                    End If
                                End If
                            End If
                        End If
                        If pVal.ItemUID = "2" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            RS.DoQuery("Delete From [@GEN_SZ_ORDR] Where u_sono = '" + Trim(objForm.Items.Item("8").Specific.value) + "' And u_macid = '" + MAC_ID + "'")
                        End If
                End Select
            ElseIf pVal.BeforeAction = True And ModalForm = True And pVal.FormUID = (objSubForm.UniqueID.Substring(objSubForm.UniqueID.IndexOf("@") + 1)) Then
                objSubForm = oApplication.Forms.Item("SOORD@" & pVal.FormUID)
                objSubForm.Select()
                BubbleEvent = False
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.ActionSuccess = True Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRSet.DoQuery("Select UserId From OUSR Where User_Code = '" + oCompany.UserName.Trim + "        '")
                        oRS.DoQuery("Select Max(DocEntry) AS 'DocEntry' From ORDR Where UserSign = '" + Trim(oRSet.Fields.Item("UserID").Value) + "'")
                        oRSet.DoQuery("Select DocNum From  ORDR Where DocEntry = '" + Trim(oRS.Fields.Item("DocEntry").Value) + "'")
                        oRS.DoQuery("Update [@GEN_SZ_ORDR] Set u_sono = '" + Trim(oRSet.Fields.Item("DocNum").Value) + "' Where u_sono = '" + SONO + "' And u_macid = '" + MAC_ID + "'")
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation_SalesOrder_Allocation(ByVal FormUID As String) As Boolean
        Try
            objSubForm = oApplication.Forms.Item(FormUID)
            Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim ogrd As SAPbouiCOM.Grid
            ogrd = objSubForm.Items.Item("grd").Specific
            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
            For i As Integer = 0 To ogrd.Rows.Count - 1
                RS.DoQuery("Select IsNull(Sum(Convert(money,u_qty)),0) AS 'Qty' From [@GEN_SZ_ORDR] Where u_sono = '" + ogrd.DataTable.Columns.Item("OrderNo").Cells.Item(i).Value.ToString.Trim + "' And u_itemcode = '" + ogrd.DataTable.Columns.Item("ItemCode").Cells.Item(i).Value.ToString.Trim + "' And u_macid = '" + MAC_ID + "' And u_asrtcode = '" + ogrd.DataTable.Columns.Item("AssortedCode").Cells.Item(i).Value.ToString.Trim + "'")
                If CDbl(ogrd.DataTable.Columns.Item("Qty").Cells.Item(i).Value) <> CDbl(RS.Fields.Item("Qty").Value) Then
                    oApplication.StatusBar.SetText("Please enter sizes for the items", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                    Exit Function
                End If
            Next
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Sub ItemEvent_SalesOrder_Allocation(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Dim CHILD_FORM As String = "GEN_SZ_ORDR@" & pVal.FormUID
            Dim ChildModalForm As Boolean = False
            For i As Integer = 0 To oApplication.Forms.Count - 1
                If oApplication.Forms.Item(i).UniqueID = (CHILD_FORM) Then
                    objSubForm = oApplication.Forms.Item("GEN_SZ_ORDR@" & pVal.FormUID)
                    ChildModalForm = True
                    Exit For
                End If
            Next
            If ChildModalForm = False Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                        objSubForm = oApplication.Forms.Item(pVal.FormUID)
                        PARENT_FORM = (pVal.FormUID.Substring(objSubForm.UniqueID.IndexOf("@") + 1))
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                            If Me.Validation_SalesOrder_Allocation(pVal.FormUID) = False Then
                                BubbleEvent = False
                            End If
                        End If
                        If pVal.ItemUID = "2" And pVal.BeforeAction = False Then
                            ModalForm = False
                        End If
                        If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            ModalForm = False
                            objSubForm.Close()
                            objForm = oApplication.Forms.ActiveForm
                            'Dim oMatrix As SAPbouiCOM.Matrix
                            'oMatrix = objForm.Items.Item("38").Specific
                            'oMatrix.Columns.Item("U_sizemtx").Editable = True
                            'For i As Integer = 1 To oMatrix.VisualRowCount - 1
                            '    oMatrix.Columns.Item("U_sizemtx").Cells.Item(i).Specific.value = "Yes"
                            'Next
                            'oMatrix.Columns.Item("U_sizemtx").Editable = False
                        End If
                        If pVal.ItemUID = "chs" And pVal.BeforeAction = True Then
                            Dim flg As Boolean = False
                            Dim slflag As Boolean = False
                            Dim ogrd As SAPbouiCOM.Grid
                            ogrd = objSubForm.Items.Item("grd").Specific
                            For i As Integer = 0 To ogrd.Rows.Count - 1
                                If ogrd.Rows.IsSelected(i) = True Then
                                    flg = True
                                End If
                            Next
                            If flg = False Then
                                oApplication.StatusBar.SetText("Please select items", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                        If pVal.ItemUID = "chs" And pVal.BeforeAction = False Then
                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            Dim ogrd As SAPbouiCOM.Grid
                            ogrd = objSubForm.Items.Item("grd").Specific
                            For i As Integer = 0 To ogrd.Rows.Count - 1
                                If ogrd.Rows.IsSelected(i) = True Then
                                    Me.Open_Size_Matrix_Form(pVal.FormUID, CStr(ogrd.DataTable.Columns.Item("OrderNo").Cells.Item(i).Value).Trim(), CStr(ogrd.DataTable.Columns.Item("ItemCode").Cells.Item(i).Value).Trim(), CStr(ogrd.DataTable.Columns.Item("Qty").Cells.Item(i).Value).Trim(), MAC_ID, CStr(ogrd.DataTable.Columns.Item("AssortedCode").Cells.Item(i).Value).Trim())
                                End If
                            Next
                        End If
                End Select
            ElseIf pVal.BeforeAction = True And ChildModalForm = True And pVal.FormUID = (objSubForm.UniqueID.Substring(objSubForm.UniqueID.IndexOf("@") + 1)) Then
                objSubForm = oApplication.Forms.Item("GEN_SZ_ORDR@" & pVal.FormUID)
                objSubForm.Select()
                BubbleEvent = False
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FormDataEvent_SalesOrder_Allocation(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.ActionSuccess = True Then
                        'objSubForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        'objSubForm.EnableMenu("1281", True)
                        'objSubForm.Items.Item("ordrno").Specific.Value = orderno
                        'objSubForm.Items.Item("macid").Specific.value = hwid
                        'objSubForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        'objSubForm.EnableMenu("1281", False)
                    End If
            End Select

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation_SalesOrder_SizeMatrix(ByVal FormUID As String) As Boolean
        Try
            Dim errflag As Boolean = False
            Dim Total As Double
            objSForm = oApplication.Forms.Item(FormUID)
            Dim ogrd As SAPbouiCOM.Grid
            ogrd = objSForm.Items.Item("grd").Specific
            For i As Integer = 0 To ogrd.Rows.Count - 1
                Total = Total + ogrd.DataTable.Columns.Item("Qty").Cells.Item(i).Value
            Next
            If Total <> TotQty Then
                oApplication.StatusBar.SetText("Total quantity should be equal to the amount in sales order", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Sub ItemEvent_SalesOrder_SizeMatrix(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objSForm = oApplication.Forms.Item(pVal.FormUID)
                    PARENT_FORM = (pVal.FormUID.Substring(objSForm.UniqueID.IndexOf("@") + 1))
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation_SalesOrder_SizeMatrix(pVal.FormUID) = False Then
                            BubbleEvent = False
                        End If
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        'Dim grd As SAPbouiCOM.Grid = objSForm.Items.Item("grd").Specific
                        'Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'Dim Code As String = RS.Fields.Item("Code").Value
                        'For i As Integer = 0 To grd.Rows.Count - 1
                        '    If grd.DataTable.Columns.Item("Qty").Cells.Item(i).Value > 0 Then
                        '        RS.DoQuery("Select Count(*) + 1 AS 'Code' From [@GEN_SZ_ORDR]")
                        '        oRSet.DoQuery("Insert Into [@GEN_SZ_ORDR] (Code,u_sono,u_itemcode,u_macid,u_size,u_qty) Values('" + Trim(RS.Fields.Item("Code").Value) + "','" + grd.DataTable.Columns.Item("SalesOrderNO").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("ItemCode").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("Size").Cells.Item(i).Value.ToString.Trim + "'," + grd.DataTable.Columns.Item("Qty").Cells.Item(i).Value + ") ")
                        '    End If
                        'Next
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Dim grd As SAPbouiCOM.Grid = objSForm.Items.Item("grd").Specific
                        Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' Dim Code As String = RS.Fields.Item("Code").Value
                        oRSet.DoQuery("Delete From [@GEN_SZ_ORDR] Where u_sono = '" + GSONO + "' And u_macid = '" + GMACID + "' And u_itemcode = '" + GITEMCODE + "' And u_asrtcode = '" + GASRTCODE + "'")
                        For i As Integer = 0 To grd.Rows.Count - 1
                            If grd.DataTable.Columns.Item("Qty").Cells.Item(i).Value > 0 Then
                                RS.DoQuery("Select Convert(VarChar,Count(*) + 1) AS 'Code' From [@GEN_SZ_ORDR]")
                                oRSet.DoQuery("Insert Into [@GEN_SZ_ORDR] (Code,Name,u_sono,u_itemcode,u_asrtcode,u_macid,u_size,u_qty,u_cutqty) Values('" + Trim(RS.Fields.Item("Code").Value) + "','" + Trim(RS.Fields.Item("Code").Value) + "','" + grd.DataTable.Columns.Item("SalesOrderNO").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("ItemCode").Cells.Item(i).Value.ToString.Trim + "','" + GASRTCODE + "','" + MAC_ID + "','" + grd.DataTable.Columns.Item("Size").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("Qty").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("CutQty").Cells.Item(i).Value.ToString.Trim + "') ")
                            End If
                        Next
                        objSForm.Close()
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FormDataEvent_SalesOrder_SizeMatrix(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.ActionSuccess = True Then
                        objSForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        objSForm.EnableMenu("1281", True)
                        Dim DBSource As SAPbouiCOM.DBDataSource
                        DBSource = objSForm.DataSources.DBDataSources.Item("@GEN_SIZE_MX")
                        DBSource.SetValue("U_ordrno", 0, sorderno)
                        DBSource.SetValue("U_macid", 0, shwid)
                        DBSource.SetValue("U_itemcode", 0, sitemcode)
                        objSForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objSForm.EnableMenu("1281", False)
                    End If
            End Select

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub Open_Order_Allocation_Form(ByVal FormUID As String, ByVal ordrno As String, ByVal macid As String)
        Try
            PARENT_FORM = FormUID
            Dim RS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim CHILD_FORM As String = "SOORD@" & FormUID
            Dim oBool As Boolean = False
            For i As Integer = 0 To oApplication.Forms.Count - 1
                If oApplication.Forms.Item(i).UniqueID = (CHILD_FORM) Then
                    objSubForm = oApplication.Forms.Item(CHILD_FORM)
                    objSubForm.Select()
                    oBool = True
                    Exit For
                End If
            Next
            If oBool = False Then
                oUtilities.SAPXML("SOORD.xml", CHILD_FORM)
                objSubForm = oApplication.Forms.Item(CHILD_FORM)
                objSubForm.Select()
            End If
            ChildModalForm = True
            Dim ogrid As SAPbouiCOM.Grid
            ogrid = objSubForm.Items.Item("grd").Specific
            objSubForm.DataSources.DataTables.Add("MyDataTable")
            objSubForm.DataSources.DataTables.Item(0).ExecuteQuery("Select Distinct A.OrdrNo 'OrderNo',A.itemcode 'ItemCode',B.ItemName 'ItemName',A.asrtcode 'AssortedCode',Sum(Convert(Money,qty)) 'Qty' From tmp_ordr_items A Inner Join OITM B On A.ItemCOde = B.ItemCode  Where A.ordrno = '" + ordrno + "' And A.macid = '" + macid + "' Group By A.OrdrNo,A.ItemCode,B.ItemName,A.asrtcode")
            ogrid.DataTable = objSubForm.DataSources.DataTables.Item("MyDataTable")
            For i As Integer = 0 To ogrid.Columns.Count - 1
                ogrid.Columns.Item(i).Editable = False
            Next
            ogrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub Open_Size_Matrix_Form(ByVal FormUID As String, ByVal ordrno As String, ByVal itemno As String, ByVal quantity As String, ByVal macid As String, ByVal asrtcode As String)
        Try
            PARENT_FORM = FormUID
            Dim CHILD_FORM As String = "GEN_SZ_ORDR@" & FormUID
            Dim oBool As Boolean = False
            For i As Integer = 0 To oApplication.Forms.Count - 1
                If oApplication.Forms.Item(i).UniqueID = (CHILD_FORM) Then
                    objSForm = oApplication.Forms.Item(CHILD_FORM)
                    objSForm.Select()
                    oBool = True
                    Exit For
                End If
            Next
            If oBool = False Then
                oUtilities.SAPXML("SizeMatrix.xml", CHILD_FORM)
                objSForm = oApplication.Forms.Item(CHILD_FORM)
                objSForm.Select()
            End If
            ChildModalForm = True
            TotQty = quantity
            GSONO = ordrno
            GMACID = macid
            GITEMCODE = itemno
            GASRTCODE = asrtcode
            Dim ogrid As SAPbouiCOM.Grid
            ogrid = objSForm.Items.Item("grd").Specific
            objSForm.DataSources.DataTables.Add("MyDataTable")
            objSForm.DataSources.DataTables.Item(0).ExecuteQuery("Exec Allocate_Size '" + ordrno + "','" + itemno + "','" + quantity + "','" + macid + "','" + asrtcode + "'")
            ogrid.DataTable = objSForm.DataSources.DataTables.Item("MyDataTable")
            For i As Integer = 0 To ogrid.Columns.Count - 1
                ogrid.Columns.Item(i).Editable = False
            Next
            ogrid.Columns.Item(ogrid.Columns.Count - 1).Editable = True
            ogrid.Columns.Item(ogrid.Columns.Count - 2).Editable = True
            ogrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefaultSEM(ByVal FormUID As String)
        Try
            objSubForm = oApplication.Forms.Item(FormUID)
            objSubForm.Freeze(True)
            If objSubForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                objSubForm.EnableMenu("1282", False)
            End If
            objSubForm.PaneLevel = 1
            objSubMatrix = objSubForm.Items.Item("OrdrMatrix").Specific
            objSubMatrix.Clear()
            objSubMatrix.FlushToDataSource()
            objSubMatrix.Clear()
            objSubMatrix.AddRow()
            objSubMatrix.FlushToDataSource()
            Me.SetNewLineSEM(objSubForm.UniqueID, objSubMatrix.VisualRowCount, objSubMatrix)
            objSubForm.Freeze(False)
        Catch ex As Exception
            objSubForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLineSEM(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objSubForm = oApplication.Forms.Item(FormUID)
            SZDBDetail = objSubForm.DataSources.DBDataSources.Item("@ORDR_ITEMS")
            objMatrix = oMatrix
            objSubForm.Freeze(True)
            objMatrix.FlushToDataSource()
            SZDBDetail.Offset = Row - 1
            SZDBDetail.SetValue("u_sono", SZDBDetail.Offset, "")
            SZDBDetail.SetValue("u_itemcode", SZDBDetail.Offset, "")
            SZDBDetail.SetValue("u_itemname", SZDBDetail.Offset, "")
            SZDBDetail.SetValue("u_qty", SZDBDetail.Offset, "")
            SZDBDetail.SetValue("u_asrtcode", SZDBDetail.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objSubForm.Freeze(False)
        Catch ex As Exception
            objSubForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLineSZ(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objSForm = oApplication.Forms.Item(FormUID)
            SMDBDetail = objSForm.DataSources.DBDataSources.Item("@GEN_SIZE_MX_D0")
            objMatrix = oMatrix
            objSForm.Freeze(True)
            objMatrix.FlushToDataSource()
            SMDBDetail.Offset = Row - 1
            SMDBDetail.SetValue("LineId", SMDBDetail.Offset, objMatrix.VisualRowCount)
            SMDBDetail.SetValue("u_size", SMDBDetail.Offset, "")
            SMDBDetail.SetValue("u_qty", SMDBDetail.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objSForm.Freeze(False)
        Catch ex As Exception
            objSForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = True Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "1293"
                        If objForm.TypeEx = "139" Then
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            objMatrix = objForm.Items.Item("38").Specific
                            DeleteItemCode = objMatrix.Columns.Item("1").Cells.Item(RowID).Specific.Value
                            oRecordSet.DoQuery("Delete From [@GEN_SZ_ORDR] Where u_itemcode = '" + DeleteItemCode + "' And u_sono = '" + Trim(objForm.Items.Item("8").Specific.value) + "'")
                        End If
                    Case "1288", "1289", "1290", "1291"
                        If objForm.TypeEx = "139" Then
                            BubbleEvent = False
                        End If
                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            RowID = eventInfo.Row
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


End Class
