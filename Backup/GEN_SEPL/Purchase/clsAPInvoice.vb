Public Class ClsAPInvoice

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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("OPCH")
            Me.SetChooseFromList(FormUID)
            objOldItem = objForm.Items.Item("10000330")
            objItem = objForm.Items.Add("CopyFrom", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Top = objOldItem.Top
            objItem.Left = objOldItem.Left - objOldItem.Width - 5
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width
            objItem.Specific.Caption = "Copy From SC.GRN"
            objItem.Specific.ChooseFromListUID = "SC_GRN_CFL"
            objItem.LinkTo = "10000330"
            objForm.Items.Item("CopyFrom").Enabled = False
            objOldItem = objForm.Items.Item("2")
            objItem = objForm.Items.Add("dnote", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Top = objOldItem.Top
            objItem.Left = objOldItem.Left + objOldItem.Width + 5
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width
            objItem.Specific.caption = "Debit Note"
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
            objItem.Specific.databind.setbound(True, "OPCH", "u_unit")
            objItem.Specific.TabOrder = TempItem.Specific.TabOrder + 1
            objItem.LinkTo = "46"
            objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetChooseFromList(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = True
            oCFLCreationParams.ObjectType = "GEN_SC_GRPO"
            oCFLCreationParams.UniqueID = "SC_GRN_CFL"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetConditionToGRN(ByVal FormUID As String, ByVal CardCode As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("SC_GRN_CFL")
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Select DISTINCT  DocEntry from [@GEN_SC_GRPO] Where U_CardCode='" & CardCode & "' and ISNULL(U_PayNum,'')=''")
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            For Row As Integer = 1 To oRS.RecordCount
                If Row > 1 Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                oCon = oCons.Add()
                oCon.Alias = "DocEntry"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = Trim(oRS.Fields.Item("DocEntry").Value)
                oRS.MoveNext()
            Next
            If oRS.RecordCount = 0 Then
                oCon = oCons.Add()
                oCon.Alias = "DocEntry"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "-1"
            End If
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub FilterGRPO(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("12")
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Distinct DocNum From OPDN Where DocStatus = 'O' And IsNull(u_insstat,'Open') = 'Closed'")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            For IntICount As Integer = 0 To oRecordSet.RecordCount - 1
                If IntICount = (oRecordSet.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "DocNum"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("DocNum").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "DocNum"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("DocNum").Value
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                oRecordSet.MoveNext()
            Next
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadItems(ByVal FormUID As String, ByVal GRNNo As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Select T1.DocEntry,T1.LineID,T1.U_ItemCode + ' - ' + T1.U_ItemDesc ItemCode,T1.U_RecdQty Quantity,T1.U_SerPrice Price,T1.U_TaxCode TaxCode,T2.U_SubAcct AcctCode,T0.U_PayTrms PaymentTerms from [@GEN_SC_GRPO] T0 INNER JOIN [@GEN_SC_GRPO_D0] T1 ON T0.DocEntry=T1.DocEntry INNER JOIN OCRD T2 ON T2.CardCode=T0.U_CardCode Where T1.DocEntry IN(" & GRNNo & ")")
            'oRS.DoQuery("Select T0.DocEntry,T0.DocNum,T1.LineNum,T2.U_BaseLine,T1.ItemCode,T1.OpenQty Quantity from OPDN T0 INNER JOIN PDN1 T1 ON T0.DocEntry=T1.DocEntry LEFT JOIN WTR1 T2 ON T2.U_BaseType='GRN' and T2.U_BaseRef=T0.DocNum and T2.U_BaseLine=T1.LineNum Where T0.DocNum IN(" & GRNNo & ") and ISNULL(T2.U_BaseLine,'')=''")
            objMatrix = objForm.Items.Item("39").Specific
            objMatrix.Clear()
            objMatrix.AddRow()
            oBool = True
            oRS.MoveFirst()
            For Row As Integer = 1 To oRS.RecordCount
                objMatrix.Columns.Item("1").Cells.Item(Row).Specific.Value = Trim(oRS.Fields.Item("ItemCode").Value)
                objMatrix.Columns.Item("2").Cells.Item(Row).Specific.Value = Trim(oRS.Fields.Item("AcctCode").Value)
                objMatrix.Columns.Item("95").Cells.Item(Row).Specific.Value = Trim(oRS.Fields.Item("TaxCode").Value)
                objMatrix.Columns.Item("12").Cells.Item(Row).Specific.Value = Trim(oRS.Fields.Item("Price").Value)
                objMatrix.Columns.Item("U_SGRNNo").Cells.Item(Row).Specific.Value = Trim(oRS.Fields.Item("DocEntry").Value)
                objMatrix.Columns.Item("U_SGRNLine").Cells.Item(Row).Specific.Value = Trim(oRS.Fields.Item("LineID").Value)
                objMatrix.Columns.Item("U_SGRNQty").Cells.Item(Row).Specific.Value = CDbl(oRS.Fields.Item("Quantity").Value)
                oRS.MoveNext()
            Next
            oBool = False
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                    If pVal.BeforeAction = False Then
                        If Trim(DOCNUM).Equals("") = False Then
                            objForm = oApplication.Forms.Item(FormUID)
                            objForm.Freeze(True)
                            Me.LoadItems(FormUID, DOCNUM)
                            DOCNUM = ""
                            objForm.Freeze(False)
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "3" And pVal.BeforeAction = False Then
                        objForm = oApplication.Forms.Item(FormUID)
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("OPCH")
                        If Trim(oDBs_Head.GetValue("DocType", 0)).Equals("S") = True Then
                            If Trim(objForm.Items.Item("4").Specific.Value).Equals("") = True Then
                                objForm.Items.Item("CopyFrom").Enabled = False
                            Else
                                objForm.Items.Item("CopyFrom").Enabled = True
                            End If
                        Else
                            objForm.Items.Item("CopyFrom").Enabled = False
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    If pVal.BeforeAction = True Then
                        If pVal.FormTypeCount = 1 Then
                            Me.CreateForm(FormUID)
                        Else
                            BubbleEvent = False
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
                    If pVal.ItemUID = "dnote" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If pVal.BeforeAction = True Then
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select DocEntry From ORPC WHere u_invno = '" + Trim(objForm.Items.Item("8").Specific.value) + "'")
                            If oRSet.RecordCount > 0 Then
                                oApplication.StatusBar.SetText("Debit Note already raised for this invoice", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        Else
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select A.DocNum,A.CardCode,A.DocDate,A.TaxDate,A.NumatCard,B.ItemCode,B.u_shqty,B.u_rejqty,B.U_exqty,B.Price,B.TaxCode From OPCH A Inner Join PCH1 B On A.DocEntry = B.DocEntry Where A.DocNum = '" + Trim(objForm.Items.Item("8").Specific.value) + "' And (B.u_shqty > 0  or B.u_rejqty > 0)")
                            If oRSet.RecordCount = 0 Then
                                Exit Sub
                            End If
                            oApplication.ActivateMenuItem("2309")
                            Dim DebitNoteForm As SAPbouiCOM.Form
                            Dim DebitNoteMatrix As SAPbouiCOM.Matrix
                            DebitNoteForm = oApplication.Forms.ActiveForm
                            DebitNoteMatrix = DebitNoteForm.Items.Item("38").Specific
                            DebitNoteForm.Items.Item("4").Specific.value = oRSet.Fields.Item("CardCode").Value
                            DebitNoteForm.Items.Item("invno").Specific.value = oRSet.Fields.Item("DocNum").Value
                            DebitNoteForm.Items.Item("14").Specific.value = oRSet.Fields.Item("NumatCard").Value
                            DebitNoteForm.Items.Item("3").Specific.Select("I", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            For i As Integer = 1 To oRSet.RecordCount
                                DebitNoteMatrix.Columns.Item("1").Cells.Item(i).Specific.value = oRSet.Fields.Item("ItemCode").Value
                                If oRSet.Fields.Item("u_shqty").Value > 0 Then
                                    DebitNoteMatrix.Columns.Item("11").Cells.Item(i).Specific.value = oRSet.Fields.Item("u_shqty").Value
                                End If
                                If oRSet.Fields.Item("u_rejqty").Value > 0 Then
                                    DebitNoteMatrix.Columns.Item("11").Cells.Item(i).Specific.value = Convert.ToDouble(oRSet.Fields.Item("u_rejqty").Value) + Convert.ToDouble(oRSet.Fields.Item("u_shqty").Value)
                                End If
                                'Vijeesh
                                If oRSet.Fields.Item("U_exqty").Value > 0 Then
                                    DebitNoteMatrix.Columns.Item("11").Cells.Item(i).Specific.value = Convert.ToDouble(oRSet.Fields.Item("U_exqty").Value) + Convert.ToDouble(oRSet.Fields.Item("u_rejqty").Value)
                                End If
                                'Vijeesh
                                DebitNoteMatrix.Columns.Item("14").Cells.Item(i).Specific.value = oRSet.Fields.Item("Price").Value
                                oRSet.MoveNext()
                            Next
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    If pVal.ItemUID = "39" And (pVal.ColUID = "U_SGRNNo" Or pVal.ColUID = "U_SGRNLine" Or pVal.ColUID = "U_SGRNQty") And pVal.CharPressed <> 9 And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If oBool = False Then
                            BubbleEvent = False
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If (pVal.ItemUID = "4" Or pVal.ItemUID = "54") And pVal.BeforeAction = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        objForm = oApplication.Forms.Item(FormUID)
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("OPCH")
                        If Trim(oDBs_Head.GetValue("DocType", 0)).Equals("S") = True Then
                            If Trim(objForm.Items.Item("4").Specific.Value).Equals("") = True Then
                                objForm.Items.Item("CopyFrom").Enabled = False
                            Else
                                objForm.Items.Item("CopyFrom").Enabled = True
                            End If
                        Else
                            objForm.Items.Item("CopyFrom").Enabled = False
                        End If
                    End If
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
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    If pVal.BeforeAction = True Then
                        If oCFL.UniqueID = "12" And Trim(objForm.Items.Item("3").Specific.selected.value) = "I" Then
                            Me.FilterGRPO(FormUID)
                            'Me.PostPayment()
                        End If
                        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery("Select u_unit From [@GEN_USR_UNIT] Where u_user = '" + oCompany.UserName.ToString.Trim + "'")
                        If oRecordSet.RecordCount = 0 Then
                            oApplication.StatusBar.SetText("No PC assigned to User", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("OPCH")
                        If (pVal.ItemUID = "4" Or pVal.ItemUID = "54") Then
                            objForm = oApplication.Forms.Item(FormUID)
                            If Trim(oDBs_Head.GetValue("DocType", 0)).Equals("S") = True Then
                                objForm.Items.Item("CopyFrom").Enabled = True
                                Me.SetConditionToGRN(FormUID, oDT.GetValue("CardCode", 0))
                            Else
                                objForm.Items.Item("CopyFrom").Enabled = False
                            End If
                        End If

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

                        If oCFL.UniqueID = "SC_GRN_CFL" Then
                            For i As Integer = 0 To oDT.Rows.Count - 1
                                DOCNUM = DOCNUM & "'" & Trim(oDT.GetValue("DocNum", i)) & "',"
                            Next
                            DOCNUM = DOCNUM.Substring(0, Len(DOCNUM) - 1)
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
                        If objForm.TypeEx = "141" Then
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
