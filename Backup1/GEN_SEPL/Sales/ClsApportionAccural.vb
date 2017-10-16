Public Class ClsApportionAccural

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset

#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            oUtilities.SAPXML("GEN_FRM_APP_ACC.xml")
            objForm = oApplication.Forms.GetForm("GEN_FRM_APP_ACC", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_APP_ACC")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_APP_ACC_D0")

            objForm.DataBrowser.BrowseBy = "ETDOCNUM"
            ''objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            objForm.Items.Item("ETDOCNUM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("ETDOCNUM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            ''SetDefault(objForm.UniqueID)

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_APP_ACC")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_APP_ACC_D0")

            If BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD Then
                If oDBs_Head.GetValue("U_JV_NO", 0).Trim <> "" Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                Else
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                    ''BubbleEvent = False
                End If
            End If

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "GEN_MN_APP_ACC"
                        Me.CreateForm(objForm.UniqueID)
                    Case "1282"
                        If objForm.TypeEx = "GEN_FRM_APP_ACC" Then
                            Me.SetDefault(objForm.UniqueID)
                            objForm.Items.Item("ETDOCNUM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        End If
                    Case "1281"
                        If objForm.TypeEx = "GEN_FRM_APP_ACC" Then
                            objForm.EnableMenu("1282", True)
                            objForm.Items.Item("ETDOCNUM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        End If
                    Case "1288", "1289", "1290", "1291"
                        If objForm.TypeEx = "GEN_FRM_APP_ACC" Then
                            objForm.EnableMenu("1282", True)
                        End If
                        ''Case "1293"
                        ''    If objForm.TypeEx = "GEN_FRM_APP_ACC" Then
                        ''        'If ITEM_ID.Equals("mtx") = True Then
                        ''        Dim Total As Double
                        ''        objMatrix = objForm.Items.Item("MTDTL").Specific
                        ''        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_APP_ACC")
                        ''        oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_APP_ACC_D0")
                        ''        For Row As Integer = 1 To objMatrix.VisualRowCount
                        ''            objMatrix.GetLineData(Row)
                        ''            oDBs_Detail.Offset = Row - 1
                        ''            oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, Row)
                        ''            oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, objMatrix.Columns.Item("itemcode").Cells.Item(Row).Specific.value)
                        ''            oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, objMatrix.Columns.Item("itemname").Cells.Item(Row).Specific.value)
                        ''            oDBs_Detail.SetValue("u_quantity", oDBs_Detail.Offset, objMatrix.Columns.Item("quantity").Cells.Item(Row).Specific.value)
                        ''            oDBs_Detail.SetValue("u_price", oDBs_Detail.Offset, objMatrix.Columns.Item("price").Cells.Item(Row).Specific.value)
                        ''            oDBs_Detail.SetValue("u_rowtotal", oDBs_Detail.Offset, objMatrix.Columns.Item("rowtotal").Cells.Item(Row).Specific.value)
                        ''            objMatrix.SetLineData(Row)
                        ''        Next
                        ''        objMatrix.FlushToDataSource()
                        ''        oDBs_Detail.RemoveRecord(oDBs_Detail.Size - 1)
                        ''        objMatrix.LoadFromDataSource()
                        ''        'End If
                        ''    End If
                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        ''If Me.Validation(FormUID) = False Then BubbleEvent = False

                        Dim oMatrix As SAPbouiCOM.Matrix
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_APP_ACC")
                        oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_APP_ACC_D0")
                        oMatrix = objForm.Items.Item("MTDTL").Specific

                        oMatrix.FlushToDataSource()

                        For li_row As Integer = 0 To oDBs_Detail.Size - 1
                            If oDBs_Detail.Size >= li_row + 1 Then

                                oDBs_Detail.Offset = li_row

                                If Trim(oDBs_Detail.GetValue("U_SELECT", oDBs_Detail.Offset)) <> "Y" Then
                                    oDBs_Detail.RemoveRecord(oDBs_Detail.Offset)
                                    li_row = li_row - 1
                                End If

                            End If
                        Next

                        oMatrix.LoadFromDataSource()

                        ''Me.SetDefault(FormUID)
                    ElseIf pVal.ItemUID = "BTGEN" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                        Dim oMatrix As SAPbouiCOM.Matrix
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_APP_ACC")
                        oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_APP_ACC_D0")
                        oMatrix = objForm.Items.Item("MTDTL").Specific

                        If oDBs_Head.GetValue("U_JV_NO", 0).Trim <> "" Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                            BubbleEvent = False
                            Exit Sub
                        End If

                        oMatrix.FlushToDataSource()

                        For li_row As Integer = 0 To oDBs_Detail.Size - 1
                            If oDBs_Detail.Size >= li_row + 1 Then

                                oDBs_Detail.Offset = li_row

                                If Trim(oDBs_Detail.GetValue("U_SELECT", oDBs_Detail.Offset)) <> "Y" Then
                                    oDBs_Detail.RemoveRecord(oDBs_Detail.Offset)
                                    li_row = li_row - 1
                                End If
                            End If
                        Next

                        oMatrix.LoadFromDataSource()

                        ''Create Journal Voucher
                        Dim ls_Inv_DocEntry As String = ""
                        Dim ls_Inv_DocNum As String = ""
                        Dim ls_Cre_DocEntry As String = ""
                        Dim ls_Cre_DocNum As String = ""

                        For li_row As Integer = 0 To oDBs_Detail.Size - 1
                            oDBs_Detail.Offset = li_row
                            If Trim(oDBs_Detail.GetValue("U_DOC_TYPE", oDBs_Detail.Offset)) = "Invoice" Then

                                If li_row > 0 Then
                                    ls_Inv_DocEntry += ",'" + Trim(oDBs_Detail.GetValue("U_INV_ENT", oDBs_Detail.Offset)) + "'"
                                    ls_Inv_DocNum += "," + Trim(oDBs_Detail.GetValue("U_INV_NO", oDBs_Detail.Offset))
                                Else
                                    ls_Inv_DocEntry = "'" + Trim(oDBs_Detail.GetValue("U_INV_ENT", oDBs_Detail.Offset)) + "'"
                                    ls_Inv_DocNum = Trim(oDBs_Detail.GetValue("U_INV_NO", oDBs_Detail.Offset))
                                End If
                            End If

                        Next

                        For li_row As Integer = 0 To oDBs_Detail.Size - 1
                            oDBs_Detail.Offset = li_row
                            If Trim(oDBs_Detail.GetValue("U_DOC_TYPE", oDBs_Detail.Offset)) = "Credit" Then

                                If ls_Cre_DocEntry.Length > 0 Then
                                    ls_Cre_DocEntry += ",'" + Trim(oDBs_Detail.GetValue("U_INV_ENT", oDBs_Detail.Offset)) + "'"
                                    ls_Cre_DocNum += "," + Trim(oDBs_Detail.GetValue("U_INV_NO", oDBs_Detail.Offset))
                                Else
                                    ls_Cre_DocEntry = "'" + Trim(oDBs_Detail.GetValue("U_INV_ENT", oDBs_Detail.Offset)) + "'"
                                    ls_Cre_DocNum = Trim(oDBs_Detail.GetValue("U_INV_NO", oDBs_Detail.Offset))
                                End If
                            End If

                        Next
                        Dim ls_Query As String

                        ''**************************99 PC*******************

                        ls_Query = "Select T.Code,T.U_unit,T.Segment_0,T.Account,Sum(T.Debit) Debit,Sum(T.Credit) Credit" & _
                        " From (select (Select A1.AcctCode From OACT A1 Where A1.Segment_0 = A.Segment_0 and  A1.Segment_1 = P.Code) Account2 " & _
                        " ,Unit.OcrCode " & _
                        " ,P.Code,N.U_unit, A.Segment_0, J1.Account,(J1.Debit) Debit,(J1.Credit) Credit From OINV N,OJDT J,JDT1 J1,OACT A,OASC P" & _
                        " ,(Select INV1.LineNum,INV1.DocEntry,INV1.OcrCode From INV1 Where INV1.DocEntry  in (" + ls_Inv_DocEntry + ") ) Unit " & _
                        " Where J.TransId = N.TransId " & _
                        " and J.TransId =  J1.TransId " & _
                        " and P.ShortName = N.U_UNIT " & _
                        " and A.AcctCode = J1.Account " & _
                        " and A.Segment_1 = '99' " & _
                        " and A.LocManTran = 'N' " & _
                        " and N.DocEntry in (" + ls_Inv_DocEntry + ") and Unit.DocEntry = N.DocEntry  and Cast(Unit.LineNum as Int) = (Select Min(Cast(INV1.LineNum as Int)) From INV1 Where INV1.DocEntry = Unit.DocEntry )) T" & _
                        " Group By T.Code,T.U_unit, T.Segment_0,T.Account "

                        If ls_Cre_DocEntry.Length > 0 Then
                            ls_Query += " Union all Select T.U_unit,T.Segment_0,T.Account,Sum(T.Debit) Debit,Sum(T.Credit) Credit" & _
                            " From (select (Select A1.AcctCode From OACT A1 Where A1.Segment_0 = A.Segment_0 and  A1.Segment_1 = P.Code) Account2 " & _
                            " ,Unit.OcrCode " & _
                            " ,P.Code,N.U_unit, A.Segment_0, J1.Account,Sum(J1.Debit) Debit,Sum(J1.Credit) Credit From ORIN N,OJDT J,JDT1 J1,OACT A,OASC P" & _
                            " ,(Select RIN1.LineNum,RIN1.DocEntry,RIN1.OcrCode From RIN1 Where RIN1.DocEntry  in (" + ls_Cre_DocEntry + ") ) Unit " & _
                            " Where J.TransId = N.TransId " & _
                            " and J.TransId =  J1.TransId " & _
                            " and P.ShortName = N.U_UNIT " & _
                            " and A.AcctCode = J1.Account " & _
                            " and A.Segment_1 = '99' " & _
                            " and A.LocManTran = 'N' " & _
                            " and N.DocEntry in (" + ls_Cre_DocEntry + ") and Unit.DocEntry = N.DocEntry  and Cast(Unit.LineNum as Int)  = (Select Min(Cast(INV1.LineNum as Int)) From INV1 Where INV1.DocEntry = Unit.DocEntry )) T" & _
                            " Group By T.Code,T.U_unit,T.Segment_0,T.Account"

                        End If

                        Dim oRecordset As SAPbobsCOM.Recordset
                        Dim oJournalVouchers As SAPbobsCOM.JournalVouchers
                        Dim ldt_PostingDate As Date
                        Dim oSBObob As SAPbobsCOM.SBObob

                        oSBObob = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)

                        oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        If objForm.Items.Item("ETPOS_DATE").Specific.ToString = "" Then
                            oApplication.SetStatusBarMessage("Posting Date Can't be blank", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        oRecordset = oSBObob.Format_StringToDate(objForm.Items.Item("ETPOS_DATE").Specific.ToString)
                        ldt_PostingDate = oRecordset.Fields.Item(0).Value
                        oRecordset.DoQuery(ls_Query)

                        oJournalVouchers = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers)
                        oJournalVouchers.JournalEntries.DueDate = ldt_PostingDate

                        For li_Row As Integer = 1 To oRecordset.RecordCount

                            ''PC - 99
                            oJournalVouchers.JournalEntries.Lines.AccountCode = oRecordset.Fields.Item("Account").Value

                            If oRecordset.Fields.Item("Credit").Value > 0 Then
                                oJournalVouchers.JournalEntries.Lines.Debit = oRecordset.Fields.Item("Credit").Value

                            Else
                                oJournalVouchers.JournalEntries.Lines.Credit = oRecordset.Fields.Item("Debit").Value
                            End If

                            oJournalVouchers.JournalEntries.Lines.Add()

                            oRecordset.MoveNext()

                        Next

                        ''************Invoice PC***************************

                        ls_Query = "Select T.OcrCode,T.Account2,T.Code,T.U_unit,T.Segment_0,Sum(T.Debit) Debit,Sum(T.Credit) Credit" & _
                        " From (select (Select A1.AcctCode From OACT A1 Where A1.Segment_0 = A.Segment_0 and  A1.Segment_1 = P.Code) Account2 " & _
                        " ,Unit.OcrCode " & _
                        " ,P.Code,N.U_unit, A.Segment_0, J1.Account,(J1.Debit) Debit,(J1.Credit) Credit From OINV N,OJDT J,JDT1 J1,OACT A,OASC P" & _
                        " ,(Select INV1.LineNum,INV1.DocEntry,INV1.OcrCode From INV1 Where INV1.DocEntry  in (" + ls_Inv_DocEntry + ") ) Unit " & _
                        " Where J.TransId = N.TransId " & _
                        " and J.TransId =  J1.TransId " & _
                        " and P.ShortName = N.U_UNIT " & _
                        " and A.AcctCode = J1.Account " & _
                        " and A.Segment_1 = '99' " & _
                        " and A.LocManTran = 'N' " & _
                        " and N.DocEntry in (" + ls_Inv_DocEntry + ") and Unit.DocEntry = N.DocEntry  and Cast(Unit.LineNum as Int) = (Select Min(Cast(INV1.LineNum as Int)) From INV1 Where INV1.DocEntry = Unit.DocEntry )) T" & _
                        " Group By T.Code,T.U_unit, T.Segment_0,T.Account2,T.OcrCode "
                        If ls_Cre_DocEntry.Length = 0 Then
                            ls_Query += " Order By T.OcrCode"
                        End If
                        If ls_Cre_DocEntry.Length > 0 Then
                            ls_Query += " Union all Select T.Account2,T.OcrCode,T.Code,T.U_unit,T.Segment_0,Sum(T.Debit) Debit,Sum(T.Credit) Credit" & _
                            " From (select (Select A1.AcctCode From OACT A1 Where A1.Segment_0 = A.Segment_0 and  A1.Segment_1 = P.Code) Account2 " & _
                            " ,Unit.OcrCode " & _
                            " ,P.Code,N.U_unit, A.Segment_0, J1.Account,Sum(J1.Debit) Debit,Sum(J1.Credit) Credit From ORIN N,OJDT J,JDT1 J1,OACT A,OASC P" & _
                            " ,(Select RIN1.LineNum,RIN1.DocEntry,RIN1.OcrCode From RIN1 Where RIN1.DocEntry  in (" + ls_Cre_DocEntry + ") ) Unit " & _
                            " Where J.TransId = N.TransId " & _
                            " and J.TransId =  J1.TransId " & _
                            " and P.ShortName = N.U_UNIT " & _
                            " and A.AcctCode = J1.Account " & _
                            " and A.Segment_1 = '99' " & _
                            " and A.LocManTran = 'N' " & _
                            " and N.DocEntry in (" + ls_Cre_DocEntry + ") and Unit.DocEntry = N.DocEntry  and Cast(Unit.LineNum as Int) = (Select Min(Cast(INV1.LineNum as Int)) From INV1 Where INV1.DocEntry = Unit.DocEntry )) T" & _
                            " Group By T.Code,T.U_unit, T.Segment_0,T.Account2,T.OcrCode Order By T.OcrCode"

                        End If

                        oRecordset.DoQuery(ls_Query)

                        For li_Row As Integer = 1 To oRecordset.RecordCount
                            ''Invoice PC 
                            If oRecordset.Fields.Item("Account2").Value = "" Then
                                oApplication.MessageBox(oRecordset.Fields.Item("Segment_0").Value + " : Account Not Found.")
                                Exit Sub
                            End If
                            oJournalVouchers.JournalEntries.Lines.AccountCode = oRecordset.Fields.Item("Account2").Value
                            oJournalVouchers.JournalEntries.Lines.CostingCode = oRecordset.Fields.Item("OcrCode").Value
                            If oRecordset.Fields.Item("Credit").Value > 0 Then
                                oJournalVouchers.JournalEntries.Lines.Credit = oRecordset.Fields.Item("Credit").Value
                            Else
                                oJournalVouchers.JournalEntries.Lines.Debit = oRecordset.Fields.Item("Debit").Value
                            End If

                            oJournalVouchers.JournalEntries.Lines.Add()

                            oRecordset.MoveNext()

                        Next


                        oJournalVouchers.JournalEntries.Memo = "From Apportion Accural"


                        Dim li_ret As Integer
                        Dim ls_Ret As String

                        li_ret = oJournalVouchers.Add()

                        oCompany.GetNewObjectCode(ls_Ret)

                        If li_ret < 0 Then
                            oApplication.MessageBox(oCompany.GetLastErrorDescription)
                        Else
                            oDBs_Head.SetValue("U_JV_NO", 0, ls_Ret.Split()(0))
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                            oApplication.MessageBox("Journal Voucher is created successfully")
                        End If


                    ElseIf pVal.ItemUID = "BTFETCH" And pVal.BeforeAction = True Then
                        Dim ls_Query As String
                        Dim oMatrix As SAPbouiCOM.Matrix
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_APP_ACC")
                        oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_APP_ACC_D0")

                        ls_Query = "select 'Invoice' DocType,N.DocNum,N.DocEntry,convert(Char(10),N.DocDate,112) as DocDate,N.U_unit,N.Series " & _
                                    " ,(Select Top 1 INV1.OcrCode From INV1 Where INV1.DocEntry = N.DocEntry and ISNULL(INV1.OcrCode,'') <> '') Unit " & _
                                    " From OINV N " & _
                                    " WHERE N.DocEntry not in " & _
                                    " (Select isnull(A.U_INV_ENT,0) From [@GEN_APP_ACC_D0] A) " & _
                                    " And N.DocDate >= '" + oDBs_Head.GetValue("U_DATEFROM", 0) + "'" & _
                                    " And N.DocDate <= '" + oDBs_Head.GetValue("U_DATE_TO", 0) + "' And N.DocType = 'I'"

                        If oDBs_Head.GetValue("U_PC", 0).Trim <> "" Then
                            ls_Query += " And N.U_UNIT = '" + oDBs_Head.GetValue("U_PC", 0).Trim + "' AND N.U_tinv IN ('Direct','Local')"
                        Else
                            oApplication.MessageBox("PC Can't be blank")
                            Exit Sub
                        End If


                        'ls_Query += " Union All select 'Credit' DocType,N.DocNum,N.DocEntry,convert(Char(10),N.DocDate,112) as DocDate,N.U_unit,N.Series " & _
                        '            " ,(Select Top 1 RIN1.OcrCode From RIN1 Where RIN1.DocEntry = N.DocEntry and ISNULL(RIN1.OcrCode,'') <> '') Unit " & _
                        '            " From ORIN N " & _
                        '            " WHERE N.DocEntry not in " & _
                        '            " (Select isnull(A.U_INV_ENT,0) From [@GEN_APP_ACC_D0] A) " & _
                        '            " And N.DocDate >= '" + oDBs_Head.GetValue("U_DATEFROM", 0) + "'" & _
                        '            " And N.DocDate <= '" + oDBs_Head.GetValue("U_DATE_TO", 0) + "'"

                        'If oDBs_Head.GetValue("U_PC", 0).Trim <> "" Then
                        '    ls_Query += " And N.U_UNIT = '" + oDBs_Head.GetValue("U_PC", 0).Trim + "' AND N.U_tinv = 'Direct'"
                        'Else
                        '    oApplication.MessageBox("PC Can't be blank")
                        '    Exit Sub
                        'End If

                        Dim oRecordset As SAPbobsCOM.Recordset
                        oMatrix = objForm.Items.Item("MTDTL").Specific
                        oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordset.DoQuery(ls_Query)
                        oDBs_Detail.Clear()
                        oMatrix.LoadFromDataSource()
                        If oRecordset.RecordCount > 0 Then

                            For li_ROw As Integer = 0 To oRecordset.RecordCount - 1
                                oDBs_Detail.InsertRecord(li_ROw)

                                oDBs_Detail.SetValue("U_UNIT_NO", li_ROw, oRecordset.Fields.Item("Unit").Value)
                                oDBs_Detail.SetValue("U_DOC_TYPE", li_ROw, oRecordset.Fields.Item("DocType").Value)
                                oDBs_Detail.SetValue("U_INV_NO", li_ROw, oRecordset.Fields.Item("DocNum").Value)
                                oDBs_Detail.SetValue("U_INV_ENT", li_ROw, oRecordset.Fields.Item("DocEntry").Value)
                                oDBs_Detail.SetValue("U_PC", li_ROw, oRecordset.Fields.Item("U_UNIT").Value)
                                oDBs_Detail.SetValue("U_INV_SER", li_ROw, oRecordset.Fields.Item("Series").Value)
                                oDBs_Detail.SetValue("U_INV_DATE", li_ROw, oRecordset.Fields.Item("DocDate").Value)
                                oRecordset.MoveNext()
                            Next
                        Else
                            oApplication.MessageBox("No data")
                        End If
                        oMatrix.LoadFromDataSource()
                    End If

            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            objForm.EnableMenu("1282", False)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_APP_ACC")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_APP_ACC_D0")
            oUtilities.GetSeries(FormUID, "COSERIES", "GEN_UDO_APP_ACC")
            oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("COSERIES").Specific.Selected.Value, "GEN_UDO_APP_ACC"))
            oDBs_Head.SetValue("U_DOC_DATE", 0, DateTime.Today.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_POS_DATE", 0, DateTime.Today.ToString("yyyyMMdd"))

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class

