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
    Dim oDBs_Head, oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim RS1, RS2 As SAPbobsCOM.Recordset
    Dim ModalForm As Boolean = False
    Dim ChildModalForm As Boolean = False
    Dim PARENT_FORM As String
    Dim RowNo As Integer
    Dim oCheck As SAPbouiCOM.CheckBox
    Dim Checked As Integer
    Dim orderno, hwid As String
    Dim sorderno, shwid, sitemcode As String
    Dim Mode As Integer
    Dim TotQty As Double
    Dim GSONO, GMACID, GITEMCODE As String
    Dim RowID As Integer
    Dim DeleteItemCode As String
    Dim oBool As Boolean = False
    Dim DOCNUM As String = ""
    Dim DOCNUM_LC As String = ""
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

            Me.SetChooseFromList_LC(FormUID)
            objOldItem = objForm.Items.Item("10000330")
            objItem = objForm.Items.Add("FromLC", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Top = objOldItem.Top
            objItem.Left = objOldItem.Left - objOldItem.Width - 120
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width
            objItem.Specific.Caption = "Copy From LC GRPO"
            objItem.Specific.ChooseFromListUID = "LC_GRN_CFL"
            objItem.LinkTo = "10000330"
            objForm.Items.Item("FromLC").Enabled = False

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

            TempItem = objForm.Items.Item("86")
            objItem = objForm.Items.Add("tgrpodc", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = TempItem.Top + TempItem.Height + 2
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.Caption = "Sub Contracting GRPO DocNum"
            objItem.LinkTo = "86"
            objItem.Visible = False
            TempItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("grpodoc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = TempItem.Top + TempItem.Height + 2
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.databind.setbound(True, "OPCH", "u_grpodoc")
            objItem.Specific.TabOrder = TempItem.Specific.TabOrder + 1
            objItem.LinkTo = "46"
            objItem.Visible = False

            TempItem = objForm.Items.Item("86")
            objItem = objForm.Items.Add("tgrpodc2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = TempItem.Top + TempItem.Height + 2
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.Caption = "GRPO DocNum"
            objItem.LinkTo = "86"
            objItem.Visible = False
            TempItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("grpono", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = TempItem.Top + TempItem.Height + 2
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.databind.setbound(True, "OPCH", "u_grpono")
            objItem.Specific.TabOrder = TempItem.Specific.TabOrder + 1
            objItem.LinkTo = "46"
            objItem.Visible = False



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

    Sub SetChooseFromList_LC(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = True
            oCFLCreationParams.ObjectType = "20"
            oCFLCreationParams.UniqueID = "LC_GRN_CFL"
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
            oRS.DoQuery("Exec SEPL_SC_GRPO_List '" + CardCode + "'")
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

    Sub SetConditionToGRN_LC(ByVal FormUID As String, ByVal CardCode As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("LC_GRN_CFL")
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Exec SEPL_LC_GRPO_List '" + CardCode + "'")
            If oRS.RecordCount > 0 Then
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
            Else
                oRS.DoQuery("select t0.DocEntry,T0.DocNum,T1.U_grpono  from OPDN T0 inner join [@GEN_GRPO_LCOSTS] T1 on T0.DocNum=T1.U_grpono where  U_status='Open' and U_cardcode='NA'")
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
            End If

            
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
        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Exec SEPL_SC_GRPO_List_Documents '" + MAC_ID + "'")
            'oRS.DoQuery("Select T1.DocEntry,T1.LineID,T1.U_ItemCode + ' - ' + T1.U_ItemDesc ItemCode,T1.U_RecdQty Quantity,T1.U_SerPrice Price,T1.U_TaxCode TaxCode,T2.U_SubAcct AcctCode,T0.U_PayTrms PaymentTerms from [@GEN_SC_GRPO] T0 INNER JOIN [@GEN_SC_GRPO_D0] T1 ON T0.DocEntry=T1.DocEntry INNER JOIN OCRD T2 ON T2.CardCode=T0.U_CardCode Where T1.DocEntry IN(" & GRNNo & ")")
            'oRS.DoQuery("Select T0.DocEntry,T0.DocNum,T1.LineNum,T2.U_BaseLine,T1.ItemCode,T1.OpenQty Quantity from OPDN T0 INNER JOIN PDN1 T1 ON T0.DocEntry=T1.DocEntry LEFT JOIN WTR1 T2 ON T2.U_BaseType='GRN' and T2.U_BaseRef=T0.DocNum and T2.U_BaseLine=T1.LineNum Where T0.DocNum IN(" & GRNNo & ") and ISNULL(T2.U_BaseLine,'')=''")
            objMatrix = objForm.Items.Item("39").Specific
            objMatrix.Clear()
            'objMatrix.AddRow()
            oBool = True
            oRS.MoveFirst()
            Dim Docnum As String = ""
            Dim Reference As String = ""
            ' Dim Docnum1 As String = ""
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("PCH1")
            Dim oRS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oCombo, oCombo_Loc As SAPbouiCOM.ComboBox

            oRS1.DoQuery("Select WtLiable from OCRD Where Cardcode='" + objForm.Items.Item("4").Specific.value + "'")
            For Row As Integer = 1 To oRS.RecordCount
                'objMatrix.GetLineData(Row)
                objMatrix.AddRow()
                ''objMatrix.Columns.Item("0").Cells.Item(Row).Specific.Value = Row
                ' objMatrix.Columns.Item("0").Cells.Item(Row).Specific.Value = Row
                objMatrix.Columns.Item("U_itemcode").Cells.Item(Row).Specific.Value = Trim(oRS.Fields.Item("ItemCode").Value)
                objMatrix.Columns.Item("94").Cells.Item(Row).Specific.Value = Trim(oRS.Fields.Item("AcctCode").Value)
                objMatrix.Columns.Item("95").Cells.Item(Row).Specific.Value = Trim(oRS.Fields.Item("TaxCode").Value)
                oCombo = objMatrix.Columns.Item("101").Cells.Item(Row).Specific
                If oRS1.Fields.Item(0).Value = "Y" Then
                    oCombo.Select(oRS1.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                Else
                    'oCombo.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    'objMatrix.Columns.Item("101").Cells.Item(Row).Specific.value = "No"
                End If
                oCombo_Loc = objMatrix.Columns.Item("2000002028").Cells.Item(Row).Specific
                oCombo_Loc.Select("Bangalore", SAPbouiCOM.BoSearchKey.psk_ByValue)

                '  oCombo = RPMatrix.Columns.Item("69").Cells.Item(1).Specific
                ' oCombo.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue)
                objMatrix.Columns.Item("U_SCGRNNo").Cells.Item(Row).Specific.Value = Trim(oRS.Fields.Item("DocEntry").Value)
                objMatrix.Columns.Item("U_SCGRNLine").Cells.Item(Row).Specific.Value = Trim(oRS.Fields.Item("LineNum").Value)
                objMatrix.Columns.Item("U_SCGRNQty").Cells.Item(Row).Specific.Value = CDbl(oRS.Fields.Item("RemQty").Value)
                objMatrix.Columns.Item("U_price").Cells.Item(Row).Specific.Value = CDbl(oRS.Fields.Item("POPrice").Value)
                objMatrix.Columns.Item("12").Cells.Item(Row).Specific.Value = (CDbl(oRS.Fields.Item("Remqty").Value) * CDbl(oRS.Fields.Item("POPrice").Value))
                'If objMatrix.Columns.Item("94").Cells.Item(Row + 1).Specific.value <> "" Then
                'objMatrix.AddRow()
                ''objMatrix.DeleteRow(Row + 1)''   un comment
                'End If
                'objMatrix.DeleteRow(Row + 1)

                ' objForm.Items.Item("14").Specific.value()
                ' Docnum = Trim(oRS.Fields.Item("DocNum").Value)
                If Row = 1 Then
                    Docnum = Trim(oRS.Fields.Item("DocNum").Value) '+ "," + Docnum
                Else
                    Docnum = Docnum + "," + Trim(oRS.Fields.Item("DocNum").Value)
                End If

                oRS.MoveNext()
                'objMatrix.SetLineData(Row)
            Next
            'objMatrix.FlushToDataSource()
            '  objMatrix.DeleteRow(objMatrix.VisualRowCount - 1)
            ' objMatrix.LoadFromDataSource()
            objForm.Items.Item("grpodoc").Specific.value = Docnum
            'oRS.DoQuery("Update OPCH set U_grpodoc='" + Docnum + "' where DocNum='" + objForm.Items.Item("8").Specific.value + "'")

            '  For i As Integer = 1 To oRS.RecordCount

            '            Next





            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSet.DoQuery("Delete From TMP_GRPO_SC Where MacId = '" + MAC_ID + "'")
            oBool = False
        Catch ex As Exception
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSet.DoQuery("Delete From TMP_GRPO_SC Where MacId = '" + MAC_ID + "'")
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub LoadItems_LC(ByVal FormUID As String, ByVal GRNNo As String)
        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRS2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Exec SEPL_LC_GRPO_List_Documents '" + MAC_ID + "'")
            'oRS.DoQuery("Select T1.DocEntry,T1.LineID,T1.U_ItemCode + ' - ' + T1.U_ItemDesc ItemCode,T1.U_RecdQty Quantity,T1.U_SerPrice Price,T1.U_TaxCode TaxCode,T2.U_SubAcct AcctCode,T0.U_PayTrms PaymentTerms from [@GEN_SC_GRPO] T0 INNER JOIN [@GEN_SC_GRPO_D0] T1 ON T0.DocEntry=T1.DocEntry INNER JOIN OCRD T2 ON T2.CardCode=T0.U_CardCode Where T1.DocEntry IN(" & GRNNo & ")")
            'oRS.DoQuery("Select T0.DocEntry,T0.DocNum,T1.LineNum,T2.U_BaseLine,T1.ItemCode,T1.OpenQty Quantity from OPDN T0 INNER JOIN PDN1 T1 ON T0.DocEntry=T1.DocEntry LEFT JOIN WTR1 T2 ON T2.U_BaseType='GRN' and T2.U_BaseRef=T0.DocNum and T2.U_BaseLine=T1.LineNum Where T0.DocNum IN(" & GRNNo & ") and ISNULL(T2.U_BaseLine,'')=''")
            objMatrix = objForm.Items.Item("39").Specific
            objMatrix.Clear()
            'objMatrix.AddRow()
            oBool = True
            oRS.MoveFirst()
            Dim Docnum As String = ""
            Dim Reference As String = ""
            ' Dim Docnum1 As String = ""
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("PCH1")
            Dim oRS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oCombo, oCombo_Loc As SAPbouiCOM.ComboBox
            objForm.Items.Item("14").Specific.value = Trim(oRS.Fields.Item("Ref").Value)
            oRS1.DoQuery("Select WtLiable from OCRD Where Cardcode='" + objForm.Items.Item("4").Specific.value + "'")
            For Row As Integer = 1 To oRS.RecordCount
                'objMatrix.GetLineData(Row)
                objMatrix.AddRow()
                ' oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, objMatrix.VisualRowCount)
                ' objMatrix.Columns.Item("0").Cells.Item(Row).Specific.Value = Row
                'objMatrix.Columns.Item("U_itemcode").Cells.Item(Row).Specific.Value = Trim(oRS.Fields.Item("ItemCode").Value)
                objMatrix.Columns.Item("94").Cells.Item(Row).Specific.Value = Trim(oRS.Fields.Item("GLAccount").Value)
                objMatrix.Columns.Item("1").Cells.Item(Row).Specific.Value = Trim(oRS.Fields.Item("AcctName").Value)
                objMatrix.Columns.Item("12").Cells.Item(Row).Specific.Value = Trim(oRS.Fields.Item("Total").Value)
                objMatrix.Columns.Item("U_qty").Cells.Item(Row).Specific.Value = Trim(oRS.Fields.Item("Quantity").Value)
                oCombo = objMatrix.Columns.Item("101").Cells.Item(Row).Specific
                If oRS1.Fields.Item(0).Value = "Y" Then
                    oCombo.Select(oRS1.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                Else
                    'oCombo.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    'objMatrix.Columns.Item("101").Cells.Item(Row).Specific.value = "No"
                End If
                oCombo_Loc = objMatrix.Columns.Item("2000002028").Cells.Item(Row).Specific
                oCombo_Loc.Select("Bangalore", SAPbouiCOM.BoSearchKey.psk_ByValue)

                '  oCombo = RPMatrix.Columns.Item("69").Cells.Item(1).Specific
                ' oCombo.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue)
                objMatrix.Columns.Item("U_grpode").Cells.Item(Row).Specific.Value = Trim(oRS.Fields.Item("DocEntry").Value)
                'objMatrix.Columns.Item("U_SCGRNLine").Cells.Item(Row).Specific.Value = Trim(oRS.Fields.Item("LineNum").Value)
                'objMatrix.Columns.Item("U_SCGRNQty").Cells.Item(Row).Specific.Value = CDbl(oRS.Fields.Item("RemQty").Value)
                'objMatrix.Columns.Item("U_price").Cells.Item(Row).Specific.Value = CDbl(oRS.Fields.Item("POPrice").Value)
                'objMatrix.Columns.Item("12").Cells.Item(Row).Specific.Value = (CDbl(oRS.Fields.Item("Remqty").Value) * CDbl(oRS.Fields.Item("POPrice").Value))
                ''If objMatrix.Columns.Item("94").Cells.Item(Row + 1).Specific.value <> "" Then
                'objMatrix.AddRow()
                objMatrix.DeleteRow(Row + 1)
                'End If
                'objMatrix.DeleteRow(Row + 1)

                ' objForm.Items.Item("14").Specific.value()
                ' Docnum = Trim(oRS.Fields.Item("DocNum").Value)
                If Row = 1 Then
                    Docnum = Trim(oRS.Fields.Item("DocNum").Value) '+ "," + Docnum
                Else
                    Docnum = Docnum + "," + Trim(oRS.Fields.Item("DocNum").Value)
                End If

               

                oRS.MoveNext()
                'objMatrix.SetLineData(Row)
            Next
            'objMatrix.FlushToDataSource()
            '  objMatrix.DeleteRow(objMatrix.VisualRowCount - 1)
            ' objMatrix.LoadFromDataSource()
            oRs2.DoQuery("Select Docnum from TMP_GRPO_LC where Macid='" + MAC_ID + "'")
            For i As Integer = 1 To oRs2.RecordCount
                If i = 1 Then
                    Docnum = Trim(oRS2.Fields.Item("DocNum").Value) '+ "," + Docnum
                Else
                    Docnum = Docnum + "," + Trim(oRS2.Fields.Item("DocNum").Value)
                End If
                oRS2.MoveNext()
                'Dim Reference As Integer              
            Next
            oRS.DoQuery("select T0.DocNum,T0.U_lcno'Ref' from OPDN T0 inner join [@GEN_GRPO_LCOSTS] T1 on T0.DocNum=T1.U_grpono inner join [@GEN_GRPO_LCOSTS_D0] T2 on T1.Code=T2.Code where DocNum in ( Select DocNum From TMP_GRPO_LC Where MacId = '" + MAC_ID + "' ) and (T2.U_glacct<> '' and T2.U_lname<>'NA')")
            For i As Integer = 1 To oRS.RecordCount
                If i = 1 Then
                    Reference = Trim(oRS.Fields.Item("Ref").Value) '+ "," + Docnum
                Else
                    Reference = Reference + "," + Trim(oRS.Fields.Item("Ref").Value)
                End If
                oRS.MoveNext()
            Next
           
            objForm.Items.Item("grpono").Specific.value = Docnum
            objForm.Items.Item("14").Specific.value = Reference
            'oRS.DoQuery("Update OPCH set U_grpodoc='" + Docnum + "' where DocNum='" + objForm.Items.Item("8").Specific.value + "'")

            '  For i As Integer = 1 To oRS.RecordCount

            '            Next
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSet.DoQuery("Delete From TMP_GRPO_LC Where MacId = '" + MAC_ID + "'")
            oBool = False
        Catch ex As Exception
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSet.DoQuery("Delete From TMP_GRPO_LC Where MacId = '" + MAC_ID + "'")
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
                        If Trim(DOCNUM_LC).Equals("") = False Then
                            objForm = oApplication.Forms.Item(FormUID)
                            objForm.Freeze(True)
                            Me.LoadItems_LC(FormUID, DOCNUM_LC)
                            DOCNUM_LC = ""
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
                                objForm.Items.Item("FromLC").Enabled = False
                            Else
                                objForm.Items.Item("CopyFrom").Enabled = True
                                objForm.Items.Item("FromLC").Enabled = True
                            End If
                        Else
                            objForm.Items.Item("CopyFrom").Enabled = False
                            objForm.Items.Item("FromLC").Enabled = False
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
                    objMatrix = objForm.Items.Item("39").Specific
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = True Then
                        If Trim(objForm.Items.Item("cpc").Specific.value) = "" Then
                            oApplication.StatusBar.SetText("Please select Unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And objForm.Items.Item("3").Specific.value = "S" And pVal.ActionSuccess = True Then
                        Dim SCGRN As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        SCGRN.DoQuery("select u_scgrnno,sum(U_SCGRNQty),U_itemcode,Docentry from PCH1 where DocEntry=(Select Top 1 Docentry from opch order by Docentry Desc) group by U_SCGRNNo,U_itemcode,DocEntry")
                        For i As Integer = 1 To SCGRN.RecordCount
                            Dim GRN As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim SCGRN_Old As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            SCGRN_Old.DoQuery("select sum(U_SCGRNQty) from PCH1 where U_SCGRNNo='" + SCGRN.Fields.Item(0).Value + "' and U_itemcode='" + SCGRN.Fields.Item(2).Value + "' and DocEntry <>(Select Top 1 Docentry from opch order by Docentry Desc)")
                            GRN.DoQuery("select sum(U_RecdQty),DocEntry from [@GEN_SC_GRPO_D0] where DocEntry='" + SCGRN.Fields.Item(0).Value + "' and U_ItemCode='" + SCGRN.Fields.Item(2).Value + "' Group by Docentry")
                            If CDbl(SCGRN.Fields.Item(1).Value + SCGRN_Old.Fields.Item(0).Value) >= GRN.Fields.Item(0).Value Then
                                Dim UPDATE As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                UPDATE.DoQuery("Update [@GEN_SC_GRPO] set u_status='Closed' Where Docentry='" + SCGRN.Fields.Item(0).Value + "'")
                            End If
                            SCGRN.MoveNext()
                        Next

                        Dim SCGRPO As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        SCGRPO.DoQuery("select u_grpode,sum(U_Qty),Docentry,dscription from PCH1 where DocEntry=(Select Top 1 Docentry from opch order by Docentry Desc) group by U_grpode,DocEntry,dscription")
                        For i As Integer = 1 To SCGRPO.RecordCount
                            Dim GRPO As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim SCGRPO_Old As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            SCGRPO_Old.DoQuery("select sum(U_Qty) from PCH1 where U_grpode='" + SCGRPO.Fields.Item(0).Value + "' and dscription='" + SCGRPO.Fields.Item(3).Value + "'  and DocEntry <>(Select Top 1 Docentry from opch order by Docentry Desc)")
                            GRPO.DoQuery("select (U_qty),t2.DocEntry,U_lname from [@GEN_GRPO_LCOSTS_D0] T0 inner join [@GEN_GRPO_LCOSTS] T1 on T0.Code=T1.Code inner join OPDN T2 on T2.DocNum=T1.U_grpono where t2.DocEntry='" + SCGRPO.Fields.Item(0).Value + "' AND T0.U_lname='" + SCGRPO.Fields.Item(3).Value + "' ")
                            '   For j As Integer = 1 To GRPO.RecordCount
                            If CDbl(SCGRPO.Fields.Item(1).Value + SCGRPO_Old.Fields.Item(0).Value).ToString >= GRPO.Fields.Item(0).Value.ToString Then
                                Dim UPDATE As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                UPDATE.DoQuery("Update opdn set u_status='Closed' Where Docentry='" + SCGRPO.Fields.Item(0).Value + "'")
                            End If
                            'GRPO.MoveNext()
                            'Next

                            SCGRPO.MoveNext()
                        Next
                    End If
                    Dim USER_NAME As String = oCompany.UserName
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                        If Trim(objForm.Items.Item("3").Specific.value) = "S" Then

                            If USER_NAME <> "manager" Then
                                Dim GLAccount As String
                                For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                                    GLAccount = objMatrix.Columns.Item("94").Cells.Item(Row).Specific.value
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
                        End If

                        If Trim(objForm.Items.Item("3").Specific.value) = "S" Then
                            For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                                Dim GAccnt As String
                                GAccnt = objMatrix.Columns.Item("94").Cells.Item(Row).Specific.value
                                Dim Gacc As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim Gacc_COA As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Gacc.DoQuery("Select (substring('" + GAccnt + "', 1, len('" + GAccnt + "')-3)+RIGHT('" + GAccnt + "',2))")
                                Gacc_COA.DoQuery("Select U_ccentre From OACT where FormatCode='" + Gacc.Fields.Item(0).Value + "'")
                                If Gacc_COA.Fields.Item(0).Value = "N" Or Gacc_COA.Fields.Item(0).Value = "" Then
                                    If objMatrix.Columns.Item("10002026").Cells.Item(Row).Specific.Value = "" Then
                                        Dim Rowval As Integer = Convert.ToInt32(Int(Row))
                                        oApplication.StatusBar.SetText("Please select CostCentre In Row - " & Row & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Next
                        End If
                    End If

                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess = True Then
                        Dim DPay As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim DPay1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim DPayJE As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        DPay.DoQuery("Select T0.DocNum,T1.BaseDocNum,T0.NumAtCard,(CONVERT(VARCHAR(8), T0.DocDate, 3))'DocDate' From OPCH T0 Inner Join PCH9 T1 On T0.DocEntry=T1.DocEntry Where T0.DocEntry=(Select Top 1(DocEntry) From OPCH order by DocEntry Desc)")
                        For Row As Integer = 1 To DPay.RecordCount
                            DPay1.DoQuery("Select T0.DocType From ODPO T0 where T0.DocNum='" & DPay.Fields.Item(1).Value & "' ")
                            If DPay1.Fields.Item(0).Value = "S" Then
                                DPayJE.DoQuery("Update OJDT Set Ref1='" & DPay.Fields.Item(0).Value & "' Where OJDT.BaseRef='" & DPay.Fields.Item(1).Value & "' ")
                                DPayJE.DoQuery("Update OJDT Set Ref2='" & DPay.Fields.Item(2).Value & "' Where OJDT.BaseRef='" & DPay.Fields.Item(1).Value & "' ")
                                DPayJE.DoQuery("Update OJDT Set Ref3='" & DPay.Fields.Item("DocDate").Value & "' Where OJDT.BaseRef='" & DPay.Fields.Item(1).Value & "' ")
                            End If
                        Next
                        Dim oRS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRS2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRS3 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRS1.DoQuery("Select cardcode From opch Where docentry=(Select Top 1(docentry) from opch)")
                        oRS2.DoQuery("Select WtLiable From OCRD Where Cardcode='" + oRS1.Fields.Item(0).Value + "'")
                        If oRS2.Fields.Item(0).Value = "N" Then
                            oRs3.DoQuery("Update PCH1 Set wtliable='N' where docentry=(Select Top 1(docentry) from opch)")
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
                    'If pVal.ItemUID = "39" And (pVal.ColUID = "U_SGRNQty" Or pVal.ColUID = "5") Then
                    '    For Row As Integer = 1 To objMatrix.VisualRowCount
                    '        objMatrix.Columns.Item("12").Cells.Item(Row).Specific.Value = CDbl(CDbl(objMatrix.Columns.Item("U_SGRNQty").Cells.Item(Row).Specific.Value) * (objMatrix.Columns.Item("5").Cells.Item(Row).Specific.Value))
                    '    Next

                    'End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If pVal.ItemUID = "39" And (pVal.ColUID = "U_SCGRNQty" Or pVal.ColUID = "U_price") And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = False Then
                        Dim objMatrix As SAPbouiCOM.Matrix
                        objMatrix = objForm.Items.Item("39").Specific
                        'Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'oRSet.DoQuery("Select U_POPrice From [@GEN_SC_GRPO_D0] Where DocEntry = '" + Trim(objMatrix.Columns.Item("U_SGRNNo").Cells.Item(pVal.Row).Specific.value) + "' And LineId = '" + Trim(objMatrix.Columns.Item("U_SGRNLine").Cells.Item(pVal.Row).Specific.value) + "'")
                        ' Dim a As String = (objMatrix.Columns.Item("U_SCGRNQty").Cells.Item(pVal.Row).Specific.Value * objMatrix.Columns.Item("U_price").Cells.Item(pVal.Row).Specific.Value)
                        objMatrix.Columns.Item("12").Cells.Item(pVal.Row).Specific.value = (objMatrix.Columns.Item("U_SCGRNQty").Cells.Item(pVal.Row).Specific.Value * objMatrix.Columns.Item("U_price").Cells.Item(pVal.Row).Specific.Value)
                        ' objMatrix.Columns.Item("12").Cells.Item(pVal.Row).Specific.Value = (objMatrix.Columns.Item("U_price").Cells.Item(pVal.Row).Specific.Value * objMatrix.Columns.Item("U_SCGRNQty").Cells.Item(pVal.Row).Specific.Value)
                    End If

                    If pVal.ItemUID = "4" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = False Then
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRSet.DoQuery("Select u_unit From OCRD Where CardCode = '" + Trim(objForm.Items.Item("4").Specific.value) + "'")
                        Dim unit As String
                        unit = oRSet.Fields.Item("u_unit").Value
                        '   oDT.SetValue("cpc", 0, oRSet.Fields.Item(0).Value)
                        '  oDBs_Head.SetValue("u_unit", 0, oRSet.Fields.Item(0).Value)
                        ' objForm.Items.Item("cpc").Specific.value = unit
                        objForm.Items.Item("cpc").Specific.value = oRSet.Fields.Item("u_unit").Value
                        objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    End If
                    If pVal.ItemUID = "39" And pVal.ColUID = "94" And pVal.BeforeAction = False Then
                        Dim USER_NAME As String = oCompany.UserName
                        If USER_NAME <> "manager" And objForm.Items.Item("3").Specific.Value = "S" Then
                            Dim GLAccount As String
                            objMatrix = objForm.Items.Item("39").Specific
                            For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                                GLAccount = objMatrix.Columns.Item("94").Cells.Item(Row).Specific.value
                                If GLAccount <> "" Then


                                    Dim str As String = "Select (substring('" + GLAccount + "', 1, len('" + GLAccount + "')-3)+RIGHT('" + GLAccount + "',2))"
                                    Dim GLacc As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    GLacc.DoQuery("Select (substring('" + GLAccount + "', 1, len('" + GLAccount + "')-3)+RIGHT('" + GLAccount + "',2))")
                                    Dim ManualJE As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    ManualJE.DoQuery("Select count(Code) From [@GEN_M_JE] where code='" + GLacc.Fields.Item(0).Value + "'")
                                    If ManualJE.Fields.Item(0).Value > 0 Then
                                        oApplication.StatusBar.SetText("You Are Not Permitted To Perform This Action-1", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        objMatrix.Columns.Item("94").Cells.Item(Row).Specific.value = ""
                                        Exit Sub
                                    End If
                                End If
                            Next
                        End If
                    End If

                    ' If pVal.ItemUID = "39" And (pVal.ColUID = "U_price") And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '   For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                    '    Dim Price As Double
                    '    ODBs_Detail.SetValue("Rate", Row, ODBs_Detail.GetValue(5, 0))
                    '    'Price = ODBs_Detail.GetValue(5, 0)

                    'Next
                    'End If

                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If (pVal.ItemUID = "4" Or pVal.ItemUID = "54") And pVal.BeforeAction = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        objForm = oApplication.Forms.Item(FormUID)
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("OPCH")
                        If Trim(oDBs_Head.GetValue("DocType", 0)).Equals("S") = True Then
                            If Trim(objForm.Items.Item("4").Specific.Value).Equals("") = True Then
                                objForm.Items.Item("CopyFrom").Enabled = False
                                objForm.Items.Item("FromLC").Enabled = False
                            Else
                                objForm.Items.Item("CopyFrom").Enabled = True
                                objForm.Items.Item("FromLC").Enabled = True
                            End If
                        Else
                            objForm.Items.Item("CopyFrom").Enabled = False
                            objForm.Items.Item("FromLC").Enabled = False
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
                    If oCFL.UniqueID = "SC_GRN_CFL" Then
                        If Trim(oDBs_Head.GetValue("DocType", 0)).Equals("S") = True And Trim(objForm.Items.Item("4").Specific.value) <> "" Then
                            objForm.Items.Item("CopyFrom").Enabled = True
                            Me.SetConditionToGRN(FormUID, objForm.Items.Item("4").Specific.value)
                        Else
                            objForm.Items.Item("CopyFrom").Enabled = False
                        End If
                    End If
                    If oCFL.UniqueID = "LC_GRN_CFL" Then
                        If Trim(oDBs_Head.GetValue("DocType", 0)).Equals("S") = True And Trim(objForm.Items.Item("4").Specific.value) <> "" Then
                            objForm.Items.Item("FromLC").Enabled = True
                            Me.SetConditionToGRN_LC(FormUID, objForm.Items.Item("4").Specific.value)
                        Else
                            objForm.Items.Item("FromLC").Enabled = False
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

                        oDBs_Head = objForm.DataSources.DBDataSources.Item("OPCH")
                        'If oCFL.UniqueID = "2" Then
                        '    Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '    oRSet.DoQuery("Select u_unit From OCRD Where CardCode = '" + Trim(oDT.GetValue("CardCode", 0)) + "'")
                        '    Dim unit As String
                        '    unit = oRSet.Fields.Item("u_unit").Value
                        '    '   oDT.SetValue("cpc", 0, oRSet.Fields.Item(0).Value)
                        '    '  oDBs_Head.SetValue("u_unit", 0, oRSet.Fields.Item(0).Value)
                        '    ' objForm.Items.Item("cpc").Specific.value = unit
                        '    objForm.Items.Item("cpc").Specific.value = oRSet.Fields.Item("u_unit").Value
                        '    objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        'End If
                        If oCFL.UniqueID = "3" Then
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select u_unit From OCRD Where CardCode = '" + Trim(oDT.GetValue("CardCode", 0)) + "'")
                            objForm.Items.Item("cpc").Specific.value = oRSet.Fields.Item("u_unit").Value
                            objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        End If

                        If oCFL.UniqueID = "SC_GRN_CFL" Then
                            'For i As Integer = 0 To oDT.Rows.Count - 1
                            '    DOCNUM = DOCNUM & "'" & Trim(oDT.GetValue("DocNum", i)) & "',"
                            'Next
                            'DOCNUM = DOCNUM.Substring(0, Len(DOCNUM) - 1)

                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Delete From TMP_GRPO_SC Where MacId = '" + MAC_ID + "'")
                            For i As Integer = 0 To oDT.Rows.Count - 1
                                DOCNUM = Trim(oDT.GetValue("DocNum", i))
                                oRSet.DoQuery("Insert Into TMP_GRPO_SC(DocNum,MacId) Values('" + Trim(oDT.GetValue("DocNum", i)) + "','" + MAC_ID + "')")
                            Next
                        End If
                    End If
                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("OPCH")
                        If (pVal.ItemUID = "4" Or pVal.ItemUID = "54") Then
                            objForm = oApplication.Forms.Item(FormUID)
                            If Trim(oDBs_Head.GetValue("DocType", 0)).Equals("S") = True Then
                                objForm.Items.Item("FromLC").Enabled = True
                                Me.SetConditionToGRN_LC(FormUID, oDT.GetValue("CardCode", 0))
                            Else
                                objForm.Items.Item("FromLC").Enabled = False
                            End If
                        End If

                        oDBs_Head = objForm.DataSources.DBDataSources.Item("OPCH")
                        'If oCFL.UniqueID = "2" Then
                        '    Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '    oRSet.DoQuery("Select u_unit From OCRD Where CardCode = '" + Trim(oDT.GetValue("CardCode", 0)) + "'")
                        '    Dim unit As String
                        '    unit = oRSet.Fields.Item("u_unit").Value
                        '    '   oDT.SetValue("cpc", 0, oRSet.Fields.Item(0).Value)
                        '    '  oDBs_Head.SetValue("u_unit", 0, oRSet.Fields.Item(0).Value)
                        '    ' objForm.Items.Item("cpc").Specific.value = unit
                        '    objForm.Items.Item("cpc").Specific.value = oRSet.Fields.Item("u_unit").Value
                        '    objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        'End If
                        If oCFL.UniqueID = "3" Then
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select u_unit From OCRD Where CardCode = '" + Trim(oDT.GetValue("CardCode", 0)) + "'")
                            objForm.Items.Item("cpc").Specific.value = oRSet.Fields.Item("u_unit").Value
                            objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        End If

                        If oCFL.UniqueID = "LC_GRN_CFL" Then
                            'For i As Integer = 0 To oDT.Rows.Count - 1
                            '    DOCNUM = DOCNUM & "'" & Trim(oDT.GetValue("DocNum", i)) & "',"
                            'Next
                            'DOCNUM = DOCNUM.Substring(0, Len(DOCNUM) - 1)

                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Delete From TMP_GRPO_LC Where MacId = '" + MAC_ID + "'")
                            For i As Integer = 0 To oDT.Rows.Count - 1
                                DOCNUM_LC = Trim(oDT.GetValue("DocNum", i))
                                oRSet.DoQuery("Insert Into TMP_GRPO_LC(DocNum,MacId) Values('" + Trim(oDT.GetValue("DocNum", i)) + "','" + MAC_ID + "')")
                            Next
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
