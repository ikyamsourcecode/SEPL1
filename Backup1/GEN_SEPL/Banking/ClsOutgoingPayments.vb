Public Class ClsOutgoingPayments

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm, objSubForm As SAPbouiCOM.Form
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim objItem, TempItem As SAPbouiCOM.Item
    Dim objMatrix, oMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim PARENT_FORM As String
    Dim oRS As SAPbobsCOM.Recordset
    Dim oRSet As SAPbobsCOM.Recordset
    Dim ROW_ID As Integer = 0
    Dim ITEM_ID As String
    Dim RowCount As Integer
    Dim enableflag As Boolean = False
    Dim ModalForm As Boolean = False
    Dim docno As String
    Dim PAYNUM As String
    Dim DocDate, RefDate, JVRem As String
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oMatrix = objForm.Items.Item("71").Specific
            TempItem = objForm.Items.Item("96")
            objItem = objForm.Items.Add("sbc", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = TempItem.Top + 30
            objItem.Left = TempItem.Left
            objItem.Height = TempItem.Height
            objItem.Width = TempItem.Width
            objItem.LinkTo = "96"
            objItem.Specific.Caption = "Other Charges - JV"
            TempItem = objForm.Items.Item("95")
            objItem = objForm.Items.Add("ebc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = TempItem.Top + 30
            objItem.Left = TempItem.Left
            objItem.Height = TempItem.Height
            objItem.Width = TempItem.Width
            objItem.Specific.Databind.setbound(True, "OVPM", "u_bcjv")
            objItem.Specific.taborder = TempItem.Specific.taborder - 1
            objItem.LinkTo = "95"
            objItem = objForm.Items.Add("lbc", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)

            Dim LB As SAPbouiCOM.LinkedButton
            LB = objItem.Specific
            LB.LinkedObjectType = "30"
            objItem.Top = TempItem.Top + 30
            objItem.Left = TempItem.Left - 20
            objItem.Width = 20
            objItem.Height = 14
            objItem.LinkTo = "ebc"
            TempItem = objForm.Items.Item("2")
            objItem = objForm.Items.Add("btn", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Top = TempItem.Top
            objItem.Left = TempItem.Left + TempItem.Width + 5
            objItem.Height = TempItem.Height
            objItem.Width = TempItem.Width + 30
            objItem.Specific.caption = "Other Charges"
            objItem.LinkTo = "2"
            'Rajkumar ----- Forward Cover------- 26.08.14
            'TempItem = objForm.Items.Item("151")
            'objItem = objForm.Items.Add("rts", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'objItem.Top = TempItem.Top + 30
            'objItem.Left = TempItem.Left
            'objItem.Height = TempItem.Height
            'objItem.Width = TempItem.Width
            'objItem.LinkTo = "151"
            'objItem.Specific.Caption = "Rate Type"
            'TempItem = objForm.Items.Item("152")
            'objItem = objForm.Items.Add("rte", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            'objItem.Top = TempItem.Top + 30
            'objItem.Left = TempItem.Left
            'objItem.Height = TempItem.Height
            'objItem.Width = TempItem.Width
            'objItem.Specific.Databind.setbound(True, "OVPM", "U_Frgn")
            'objItem.Specific.taborder = TempItem.Specific.taborder - 1
            'objItem.LinkTo = "152"

            objForm.Items.Item("ebc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("btn").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("btn").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Dim CHILD_FORM As String = "OVPM_JV@" & pVal.FormUID
            Dim ModalForm As Boolean = False
            ' Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item(71).Specific
            For i As Integer = 0 To oApplication.Forms.Count - 1
                If oApplication.Forms.Item(i).UniqueID = (CHILD_FORM) Then
                    objSubForm = oApplication.Forms.Item("OVPM_JV@" & pVal.FormUID)
                    ModalForm = True
                    Exit For
                End If
            Next
            If ModalForm = False Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                        objForm = oApplication.Forms.ActiveForm
                        objMatrix = objForm.Items.Item("71").Specific
                        objSubForm = oApplication.Forms.Item(pVal.FormUID)
                        PARENT_FORM = (pVal.FormUID.Substring(objSubForm.UniqueID.IndexOf("@") + 1))
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        If pVal.BeforeAction = True Then
                            If pVal.FormTypeCount = 1 Then
                                Me.CreateForm(FormUID)
                            Else
                                BubbleEvent = False
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        Dim USER_NAME As String = oCompany.UserName
                        If pVal.ItemUID = "2" And pVal.BeforeAction = False Then
                            ModalForm = False
                            objForm.Close()
                        End If
                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("OVPM")
                            Dim DocType As String = oDBs_Head.GetValue("DocType", 0).Trim().ToString()
                            If DocType = "A" Then
                                If USER_NAME <> "manager" Then
                                    Dim GLAccount As String
                                    For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                                        GLAccount = objMatrix.Columns.Item("1").Cells.Item(Row).Specific.value
                                        Dim GLacc As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        GLacc.DoQuery("Select (substring('" + GLAccount + "', 1, len('" + GLAccount + "')-3)+RIGHT('" + GLAccount + "',2))")
                                        Dim ManualJE As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        ManualJE.DoQuery("Select count(Code) From [@GEN_M_JE] where code='" + GLacc.Fields.Item(0).Value + "'")
                                        If ManualJE.Fields.Item(0).Value > 0 Then
                                            oApplication.StatusBar.SetText("You Are Not Permitted To Perform This Action", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    Next
                                End If
                                For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                                    Dim GAccnt As String
                                    GAccnt = objMatrix.Columns.Item("1").Cells.Item(Row).Specific.value
                                    Dim Gacc As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim Gacc_COA As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Gacc.DoQuery("Select (substring('" + GAccnt + "', 1, len('" + GAccnt + "')-3)+RIGHT('" + GAccnt + "',2))")
                                    Gacc_COA.DoQuery("Select U_ccentre From OACT where FormatCode='" + Gacc.Fields.Item(0).Value + "'")
                                    If Gacc_COA.Fields.Item(0).Value = "N" Or Gacc_COA.Fields.Item(0).Value = "" Then
                                        If objMatrix.Columns.Item("10000044").Cells.Item(Row).Specific.value = "" Then
                                            Dim Rowval As Integer = Convert.ToInt32(Int(Row))
                                            oApplication.StatusBar.SetText("Please select CostCentre In Row - " & Row & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If

                                Next

                            End If
                            Dim oPC As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim oSeries As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oPC.DoQuery("Select u_unit from [@GEN_USR_unit] where u_user='" + oCompany.UserName + "'")
                            oSeries.DoQuery("Select (case when SeriesName like '%U1%' then 'UNIT1' when SeriesName like '%U2%' then 'UNIT2' when SeriesName like '%U3%' then 'UNIT3' when SeriesName like '%LU%' then 'LG-UNIT1' end) from NNM1 where series='" + Trim(objForm.Items.Item("87").Specific.value) + "'")
                            For i As Integer = 1 To oPC.RecordCount
                                If i <> oPC.RecordCount Then
                                    If oSeries.Fields.Item(0).Value <> oPC.Fields.Item(0).Value Then
                                        oPC.MoveNext()
                                    End If
                                Else
                                    If oSeries.Fields.Item(0).Value <> oPC.Fields.Item(0).Value Then
                                        oApplication.StatusBar.SetText("Please Select Correct Series", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Next
                        End If
                        If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            ModalForm = False
                        End If
                        If pVal.ItemUID = "2" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            RS.DoQuery("Delete From [@OVPM_JV] Where u_docno = '" + Trim(objForm.Items.Item("3").Specific.value) + "' And u_macid = '" + MAC_ID + "'")
                        End If
                        If pVal.ItemUID = "btn" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            docno = objForm.Items.Item("3").Specific.value
                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            Me.Open_OtherCharges_Form(pVal.FormUID, Trim(objForm.Items.Item("3").Specific.value), MAC_ID)
                        End If
                        If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            PAYNUM = objForm.Items.Item("3").Specific.value
                            DocDate = objForm.Items.Item("10").Specific.value
                            RefDate = objForm.Items.Item("90").Specific.value
                            JVRem = objForm.Items.Item("59").Specific.value
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                        If pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            RS.DoQuery("Delete From [@OVPM_JV] Where u_docno = '" + Trim(objForm.Items.Item("3").Specific.value) + "' And u_macid = '" + MAC_ID + "'")
                        End If
                End Select
            End If
            
        Catch ex As Exception
            '  oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then
                        objForm.Items.Item("btn").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.BeforeAction = False Then
                        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim DocEntry As String
                        Dim UserSign As String
                        Dim DocNum As String
                        Dim CardCode As String
                        oRSet.DoQuery("Select UserID From OUSR Where User_Code = '" + oCompany.UserName.ToString.Trim + "'")
                        UserSign = oRSet.Fields.Item("UserID").Value
                        oRSet.DoQuery("Select Max(DocEntry) As 'DocEntry' From OVPM Where UserSign = '" + UserSign + "'")
                        DocEntry = oRSet.Fields.Item("DocEntry").Value
                        oRSet.DoQuery("Select DocNum,CardCode From OVPM Where DocEntry = '" + DocEntry + "'")
                        DocNum = oRSet.Fields.Item("DocNum").Value
                        CardCode = oRSet.Fields.Item("CardCode").Value
                        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                        oRSet.DoQuery("Update [@OVPM_JV] Set u_docno = '" + DocNum + "' Where u_docno = '" + PAYNUM + "' And u_macid = '" + MAC_ID + "'")
                        oRSet.DoQuery("Select u_acctcode,u_debit as 'Debit',u_credit As 'Credit' From [@OVPM_JV] Where u_docno = '" + DocNum + "' And u_macid = '" + MAC_ID + "' ANd u_acctcode <> ''")
                        If oRSet.RecordCount > 0 Then
                            Dim oJE As SAPbobsCOM.JournalEntries = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                            oJE.TaxDate = DateTime.ParseExact(RefDate, "yyyyMMdd", Nothing)
                            oJE.ReferenceDate = DateTime.ParseExact(DocDate, "yyyyMMdd", Nothing)
                            oJE.Memo = JVRem
                            oJE.TransactionCode = "102"
                            oJE.Reference = DocNum
                            For k As Integer = 1 To oRSet.RecordCount
                                'Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("Select AcctCode From OACT WHere FormatCode = '" + Trim(oRSet.Fields.Item("u_acctcode").Value) + "'")
                                If oRSet.Fields.Item("Debit").Value > 0 Then
                                    oJE.Lines.AccountCode = oRecordSet.Fields.Item("AcctCode").Value
                                    oJE.Lines.Debit = oRSet.Fields.Item("Debit").Value
                                    oJE.Lines.Credit = 0
                                End If
                                If oRSet.Fields.Item("Credit").Value > 0 Then
                                    oJE.Lines.AccountCode = oRecordSet.Fields.Item("AcctCode").Value
                                    oJE.Lines.Credit = oRSet.Fields.Item("Credit").Value
                                    oJE.Lines.Debit = 0
                                End If
                                oJE.Lines.Add()
                                oJE.Lines.SetCurrentLine(oJE.Lines.Count - 1)
                                oRSet.MoveNext()
                            Next
                            Dim errflg As Integer = oJE.Add()
                            If errflg <> 0 Then
                                oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription.ToString)
                            Else
                                Dim Key As String = oCompany.GetNewObjectKey
                                oRSet.DoQuery("Update OVPM Set u_bcjv = '" + Key + "' Where DocNum = '" + DocNum + "'")
                                oRSet.DoQuery("Delete From [@OVPM_JV] Where u_docno = '" + DocNum + "' And u_macid = '" + MAC_ID + "'")
                            End If
                        End If
                        oRSet.DoQuery("Select Sum(T1.SumApplied) As  'Applied' From OVPM T0  INNER JOIN VPM2 T1 ON T0.DocEntry = T1.DocNum Where T0.DocEntry = '" + DocEntry + "' And T1.InvType = '204' And T0.DocType = 'S'")
                        If oRSet.RecordCount > 0 Then
                            Dim oBP As SAPbobsCOM.BusinessPartners
                            oBP = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                            oBP.GetByKey(CardCode)
                            oBP.UserFields.Fields.Item("U_dpmadv").Value = oBP.UserFields.Fields.Item("U_dpmadv").Value + oRSet.Fields.Item("Applied").Value
                            If oBP.Update <> 0 Then
                                oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription.ToString)
                            End If
                        End If
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub Open_OtherCharges_Form(ByVal FormUID As String, ByVal docno As String, ByVal macid As String)
        Try
            PARENT_FORM = FormUID
            Dim RS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oMatrix As SAPbouiCOM.Matrix
            Dim CHILD_FORM As String = "OVPM_JV@" & FormUID
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
                oUtilities.SAPXML("OTHERCHRGSOP.xml", CHILD_FORM)
                objSubForm = oApplication.Forms.Item(CHILD_FORM)
                objSubForm.Select()
            End If
            oMatrix = objSubForm.Items.Item("mtx").Specific
            RS1.DoQuery("Select Distinct code,u_acctcode,u_acctname,u_debit,u_credit From [@OVPM_JV] Where u_docno = '" + docno + "' and u_macid = '" + macid + "' And Isnull(u_acctcode,'') <> ''")
            If RS1.RecordCount > 0 Then
                oMatrix.AddRow(1)
                Me.SetNewLine(objSubForm.UniqueID, oMatrix.VisualRowCount, oMatrix)
                For k As Integer = 1 To RS1.RecordCount
                    ' objSubForm = oApplication.Forms.Item(FormUID)
                    oDBs_Detail = objSubForm.DataSources.DBDataSources.Item("@OVPM_JV")
                    oDBs_Detail.Offset = k - 1
                    oDBs_Detail.SetValue("u_acctcode", oDBs_Detail.Offset, RS1.Fields.Item("u_acctcode").Value)
                    oDBs_Detail.SetValue("u_acctname", oDBs_Detail.Offset, RS1.Fields.Item("u_acctname").Value)
                    oDBs_Detail.SetValue("u_debit", oDBs_Detail.Offset, RS1.Fields.Item("u_debit").Value)
                    oDBs_Detail.SetValue("u_credit", oDBs_Detail.Offset, RS1.Fields.Item("u_credit").Value)
                    oMatrix.SetLineData(oMatrix.VisualRowCount)
                    RS1.MoveNext()
                    oMatrix.AddRow(1, oMatrix.VisualRowCount)
                    Me.SetNewLine(objSubForm.UniqueID, oMatrix.VisualRowCount, oMatrix)
                Next
                Dim totcr, totdr As Double
                For i As Integer = 1 To oMatrix.VisualRowCount
                    totcr = totcr + oMatrix.Columns.Item("credit").Cells.Item(i).Specific.value
                    totdr = totdr + oMatrix.Columns.Item("debit").Cells.Item(i).Specific.value
                Next
                objSubForm.Items.Item("drtot").Specific.value = totcr
                objSubForm.Items.Item("crtot").Specific.value = totdr
            End If
            If RS1.RecordCount = 0 Then
                oMatrix.AddRow(1)
                Me.SetNewLine(objSubForm.UniqueID, oMatrix.VisualRowCount, oMatrix)
            End If
            Me.SetConditionToSO(objSubForm.UniqueID)
            ModalForm = True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetConditionToSO(ByVal FormUID As String)
        Try
            objSubForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objSubForm.ChooseFromLists.Item("ACTCFL")
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objSubForm = oApplication.Forms.Item(FormUID)
            oDBs_Detail = objSubForm.DataSources.DBDataSources.Item("@OVPM_JV")
            objMatrix = oMatrix
            objMatrix.FlushToDataSource()
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("u_acctcode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_acctname", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_debit", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_credit", oDBs_Detail.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent_OtherCharges(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objSubForm = oApplication.Forms.Item(pVal.FormUID)
                    PARENT_FORM = (pVal.FormUID.Substring(objSubForm.UniqueID.IndexOf("@") + 1))
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If pVal.BeforeAction = False Then
                        oDBs_Detail = objSubForm.DataSources.DBDataSources.Item("@OVPM_JV")
                        Dim oMatrix As SAPbouiCOM.Matrix = objSubForm.Items.Item("mtx").Specific
                        Dim totdebit As Double
                        Dim totcredit As Double
                        Try
                            If pVal.ItemUID = "mtx" And (pVal.ColUID = "debit" Or pVal.ColUID = "credit") Then
                                For i As Integer = 1 To oMatrix.VisualRowCount
                                    If oMatrix.Columns.Item("debit").Cells.Item(i).Specific.value <> "" Then
                                        totdebit = totdebit + oMatrix.Columns.Item("debit").Cells.Item(i).Specific.value
                                    End If
                                    If oMatrix.Columns.Item("credit").Cells.Item(i).Specific.value <> "" Then
                                        totcredit = totcredit + oMatrix.Columns.Item("credit").Cells.Item(i).Specific.value
                                    End If
                                Next
                                objSubForm.Items.Item("drtot").Specific.value = totdebit
                                objSubForm.Items.Item("crtot").Specific.value = totcredit
                            End If
                        Catch ex As Exception
                            oApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "btn" And pVal.BeforeAction = True Then
                        If Me.Validation_OtherCharges(pVal.FormUID) = False Then
                            BubbleEvent = False
                        End If
                    End If
                    If pVal.ItemUID = "btn" And pVal.BeforeAction = False Then
                        Dim oMatrix As SAPbouiCOM.Matrix = objSubForm.Items.Item("mtx").Specific
                        Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' Dim Code As String = RS.Fields.Item("Code").Value
                        oRSet.DoQuery("Delete From [@OVPM_JV] Where u_docno = '" + docno + "' And u_macid = '" + MAC_ID + "'")
                        For i As Integer = 1 To oMatrix.VisualRowCount
                            If oMatrix.Columns.Item("acctcode").Cells.Item(i).Specific.Value <> "" Then
                                RS.DoQuery("SELECT isnull(MAX(CAST( Code AS int)),0) +1 AS Code  FROM [@OVPM_JV]")
                                Dim dbt, cdt As String
                                dbt = oMatrix.Columns.Item("debit").Cells.Item(i).Specific.value
                                cdt = oMatrix.Columns.Item("credit").Cells.Item(i).Specific.value
                                oRSet.DoQuery("Insert Into [@OVPM_JV] (Code,Name,u_docno,u_acctcode,u_acctname,u_debit,u_credit,u_macid) Values('" + Trim(RS.Fields.Item("Code").Value) + "','" + Trim(RS.Fields.Item("Code").Value) + "','" + docno + "','" + objMatrix.Columns.Item("acctcode").Cells.Item(i).Specific.value.ToString.Trim + "','" + objMatrix.Columns.Item("acctname").Cells.Item(i).Specific.value.ToString.Trim + "','" + dbt + "','" + cdt + "','" + MAC_ID + "')")
                            End If
                        Next
                        objSubForm.Close()
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objSubForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objSubForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If oCFL.UniqueID = "ACTCFL" Then
                            objMatrix = objSubForm.Items.Item("mtx").Specific
                            oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim Total As Double = 0
                            oDBs_Detail = objSubForm.DataSources.DBDataSources.Item("@OVPM_JV")
                            Dim Flag As Boolean = False
                            Dim errflag As Boolean = False
                            If objMatrix.VisualRowCount = 1 Or pVal.Row = objMatrix.VisualRowCount Then
                                Flag = True
                            End If
                            For i As Integer = 0 To oDT.Rows.Count - 1
                                Dim cflSelectedcount As Integer = oDT.Rows.Count
                                If i < cflSelectedcount - 1 Then
                                    objMatrix.AddRow(1, pVal.Row)
                                    oDBs_Detail.InsertRecord(pVal.Row + i - 1)
                                End If
                                oDBs_Detail.Offset = pVal.Row - 1 + i
                                oDBs_Detail.SetValue("u_acctcode", oDBs_Detail.Offset, oDT.GetValue("FormatCode", 0))
                                oDBs_Detail.SetValue("u_acctname", oDBs_Detail.Offset, oDT.GetValue("AcctName", 0))
                                objMatrix.SetLineData(pVal.Row + i)
                                objSubForm.EnableMenu("1293", True)
                            Next
                            If Flag = True Then
                                objMatrix.AddRow(1, objMatrix.VisualRowCount)
                                Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
                            End If
                        End If
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation_OtherCharges(ByVal FormUID As String) As Boolean
        Try
            Dim errflag As Boolean = False
            Dim CreditTotal, DebitTotal As Double
            objSubForm = oApplication.Forms.Item(FormUID)
            Dim oMatrix As SAPbouiCOM.Matrix
            oMatrix = objSubForm.Items.Item("mtx").Specific
            For i As Integer = 1 To oMatrix.VisualRowCount
                If Trim(oMatrix.Columns.Item("acctcode").Cells.Item(i).Specific.value) <> "" Then
                    If CDbl(oMatrix.Columns.Item("debit").Cells.Item(i).Specific.value) > 0 And CDbl(oMatrix.Columns.Item("credit").Cells.Item(i).Specific.value) > 0 Then
                        oApplication.StatusBar.SetText("Cannot enter debit and credit in the same row", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                        Exit Function
                    End If
                    CreditTotal = CreditTotal + oMatrix.Columns.Item("credit").Cells.Item(i).Specific.value
                    DebitTotal = DebitTotal + oMatrix.Columns.Item("debit").Cells.Item(i).Specific.value
                End If
            Next
            If DebitTotal <> CreditTotal Then
                oApplication.StatusBar.SetText("Credit and Debit does not match", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

End Class
