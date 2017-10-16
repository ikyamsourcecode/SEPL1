Public Class ClsAPCreditMemo

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
    Dim oBool As Boolean = False
    Dim DOCNUM As String = ""
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            objOldItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("invno", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = objOldItem.Top
            objItem.Left = objOldItem.Left
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width
            objItem.Specific.databind.setbound(True, "ORPC", "u_invno")
            objItem.Visible = False
            objItem.LinkTo = "46"
            objOldItem = objForm.Items.Item("2")
            objItem = objForm.Items.Add("crpo", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Top = objOldItem.Top
            objItem.Left = objOldItem.Left + objOldItem.Width + 5
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width
            objItem.Specific.Caption = "Create PO"
            objItem.LinkTo = "2"

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
            objItem.Specific.databind.setbound(True, "ORPC", "u_unit")
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
                        Me.CreateForm(FormUID)
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
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = True Then
                        If Trim(objForm.Items.Item("cpc").Specific.value) = "" Then
                            oApplication.StatusBar.SetText("Please select Unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                    If pVal.ItemUID = "crpo" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If pVal.BeforeAction = True Then
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select G.DocNum From ORPC A Inner Join RPC1 B On A.DocEntry = B.DocEntry Inner Join OPCH C On A.u_invno = C.DocNum Inner Join PCH1 D On C.DocEntry = D.DocEntry Inner Join PDN1 E On D.BaseEntry = E.DocEntry And D.BaseType = '20' Inner Join POR1 F On E.BaseEntry = F.DocEntry And E.BaseType = '22' Inner Join OPOR G On F.DocEntry = G.DocEntry Where A.DocNum = '" + Trim(objForm.Items.Item("8").Specific.value) + "'")
                            Dim PO As String = oRSet.Fields.Item("DocNum").Value
                            'oRecordSet.DoQuery("Select A.CardCode,A.NumAtCard,B.ItemCode,B.Price,B.Quantity From ORPC A INNER JOIN RPC1 B ON A.DocEntry = B.DocEntry Where A.DocNum = '" + Trim(objForm.Items.Item("8").Specific.value) + "'")
                            oRecordSet.DoQuery("Select DocNum From OPOR Where u_opono = '" + PO + "'")
                            If oRecordSet.RecordCount > 0 Then
                                oApplication.StatusBar.SetText("PO already created for AP Credit Memo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                            End If
                        Else
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select G.DocNum From ORPC A Inner Join RPC1 B On A.DocEntry = B.DocEntry Inner Join OPCH C On A.u_invno = C.DocNum Inner Join PCH1 D On C.DocEntry = D.DocEntry Inner Join PDN1 E On D.BaseEntry = E.DocEntry And D.BaseType = '20' Inner Join POR1 F On E.BaseEntry = F.DocEntry And E.BaseType = '22' Inner Join OPOR G On F.DocEntry = G.DocEntry Where A.DocNum = '" + Trim(objForm.Items.Item("8").Specific.value) + "'")
                            Dim PO As String = oRSet.Fields.Item("DocNum").Value
                            oRecordSet.DoQuery("Select A.CardCode,A.NumAtCard,B.ItemCode,B.Price,B.Quantity From ORPC A INNER JOIN RPC1 B ON A.DocEntry = B.DocEntry Where A.DocNum = '" + Trim(objForm.Items.Item("8").Specific.value) + "'")
                            If oRecordSet.RecordCount = 0 Then
                                Exit Sub
                            End If
                            oApplication.ActivateMenuItem("2305")
                            Dim POForm As SAPbouiCOM.Form
                            Dim POMatrix As SAPbouiCOM.Matrix
                            POForm = oApplication.Forms.ActiveForm
                            POMatrix = POForm.Items.Item("38").Specific
                            POForm.Items.Item("4").Specific.value = oRecordSet.Fields.Item("CardCode").Value
                            POForm.Items.Item("opono").Specific.value = PO
                            POForm.Items.Item("14").Specific.value = oRecordSet.Fields.Item("NumatCard").Value
                            POForm.Items.Item("3").Specific.Select("I", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            For i As Integer = 1 To oRecordSet.RecordCount
                                POMatrix.Columns.Item("1").Cells.Item(i).Specific.value = oRecordSet.Fields.Item("ItemCode").Value
                                POMatrix.Columns.Item("11").Cells.Item(i).Specific.value = oRecordSet.Fields.Item("Quantity").Value
                                POMatrix.Columns.Item("14").Cells.Item(i).Specific.value = oRecordSet.Fields.Item("Price").Value
                                oRecordSet.MoveNext()
                            Next
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
                        If objForm.TypeEx = "181" Then
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
