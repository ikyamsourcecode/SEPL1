Public Class ClsMaterialRequisition

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim oRS As SAPbobsCOM.Recordset
    Dim oRSet As SAPbobsCOM.Recordset
    Dim ROW_ID As Integer = 0
    Dim ITEM_ID As String
    Dim RowCount As Integer
    Dim AlertWhs As String
    Dim AlertDocNum As String
    Dim enableflag As Boolean = False
    Dim UpdMode As Boolean = False
    Dim DocStatus As String
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            oUtilities.SAPXML("MaterialRequisition.xml")
            objForm = oApplication.Forms.GetForm("GEN_MREQ", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MREQ")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_MREQ_D0")
            objForm.EnableMenu("1288", True)
            objForm.EnableMenu("1289", True)
            objForm.EnableMenu("1290", True)
            objForm.EnableMenu("1291", True)
            objForm.Items.Item("btnis").Enabled = False
            objForm.Items.Item("btnret").Enabled = False
            objForm.DataBrowser.BrowseBy = "docnum"
            objForm.Items.Item("sono").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("itemcode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("sfgcode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("type").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("docdate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("empname").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("unit").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("whs").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("wipwhs").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            SetDefault(objForm.UniqueID)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            objForm.EnableMenu("1282", False)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MREQ")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_MREQ_D0")
            oUtilities.GetSeries(FormUID, "series", "GEN_MREQ")
            oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("series").Specific.Selected.Value, "GEN_MREQ"))
            oDBs_Head.SetValue("U_docdate", 0, DateTime.Today.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_status", 0, "Open")
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSet.DoQuery("Select U_Name From OUSR WHere User_Code = '" + oCompany.UserName.Trim + "'")
            oDBs_Head.SetValue("U_empname", 0, oRSet.Fields.Item("U_Name").Value)
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_MREQ_D0")
            objMatrix = oMatrix
            objMatrix.FlushToDataSource()
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, objMatrix.VisualRowCount)
            oDBs_Detail.SetValue("u_chk", oDBs_Detail.Offset, "N")
            oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_rqstqty", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_tol", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_uom", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_reqdqty", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_totavlbl", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_wipavlbl", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_totis", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_issued", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_whs", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_remarks", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_stat", oDBs_Detail.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterSO(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("SOCFL")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            oCon = oCons.Add()
            oCon.Alias = "DocStatus"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "O"
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterSOItems(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("ITCFL")
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Distinct B.ItemCode From ORDR A Inner Join RDR1 B On A.DocEntry = B.DocEntry Where A.DocEntry = '" + Trim(objForm.Items.Item("soentry").Specific.value) + "'")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            For IntICount As Integer = 0 To oRecordSet.RecordCount - 1
                If IntICount = (oRecordSet.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("ItemCode").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("ItemCode").Value
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                oRecordSet.MoveNext()
            Next
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterBOMItems(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("BOMCFL")
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Exec BOM_List '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "'")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            oCon = oCons.Add()
            oCon.Alias = "ItemCode"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = objForm.Items.Item("itemcode").Specific.value
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
            For IntICount As Integer = 0 To oRecordSet.RecordCount - 1
                If IntICount = (oRecordSet.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("Code").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("Code").Value
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                oRecordSet.MoveNext()
            Next
            oCFL.SetConditions(oCons)

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterItems(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("ITRCFL")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            oCon = oCons.Add()
            oCon.Alias = "IssueMthd"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "B"
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterItemsSample(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("ITRCFL")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try
            Dim WhsFlag As Boolean = False
            Dim RowWhs As String
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objForm = oApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("mtx").Specific
            Dim oCheck As SAPbouiCOM.CheckBox
            Dim Checked As Integer
            For I As Integer = 1 To objMatrix.VisualRowCount
                oCheck = objMatrix.Columns.Item("chk").Cells.Item(I).Specific
                If oCheck.Checked = False Then
                    Checked = Checked + 1
                    If Checked = objMatrix.VisualRowCount Then
                        oApplication.StatusBar.SetText("Select Items")
                        Return False
                    End If
                End If
            Next
            Dim DeletedRowNo As Integer = 0
            Dim InitialRow As Integer = objMatrix.VisualRowCount
            Checked = 0
            For I As Integer = 1 To InitialRow - 1
                oCheck = objMatrix.Columns.Item("chk").Cells.Item(I).Specific
                If oCheck.Checked = True Then
                    Checked = Checked + 1
                End If
                If oCheck.Checked = False Then
                    objMatrix.DeleteRow(I)
                    DeletedRowNo = DeletedRowNo + 1
                    I = I - 1
                End If
                If InitialRow = Checked + DeletedRowNo Then
                    Exit For
                End If
            Next
            For I As Integer = 1 To objMatrix.VisualRowCount
                oCheck = objMatrix.Columns.Item("chk").Cells.Item(I).Specific
                If oCheck.Checked = False Then
                    objMatrix.DeleteRow(I)
                End If
            Next
            For I As Integer = 1 To objMatrix.VisualRowCount
                objMatrix.Columns.Item("lineid").Cells.Item(I).Specific.Value = I
            Next
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MREQ")
            If Trim(oDBs_Head.GetValue("u_type", 0)) = "" Then
                oApplication.StatusBar.SetText("Please select document type", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            If Trim(oDBs_Head.GetValue("u_wipwhs", 0)) = "" And (Trim(oDBs_Head.GetValue("u_type", 0)) = "Regular" Or Trim(oDBs_Head.GetValue("u_type", 0)) = "Excess" Or Trim(oDBs_Head.GetValue("u_type", 0)) = "Consumable") Then
                oApplication.StatusBar.SetText("Please select WIP warehouse", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            'If Trim(oDBs_Head.GetValue("u_type", 0)) = "Regular" Then
            '    oRSet.DoQuery("Select A.DocNum From [@GEN_MREQ] A Where A.u_sono = '" + Trim(objForm.Items.Item("sono").Specific.value) + "' And A.u_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' And A.u_sfgcode = '" + Trim(objForm.Items.Item("sfgcode").Specific.value) + "' And A.u_status = 'Closed' And A.DocNum <> '" + Trim(objForm.Items.Item("docnum").Specific.value) + "' ANd A.u_type = 'Regular'")
            '    If oRSet.RecordCount > 0 Then
            '        oApplication.StatusBar.SetText("MRN already generated for this sales order,itemcode and Semi-FG combination", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        Return False
            '    End If
            'End If
            If Trim(oDBs_Head.GetValue("u_type", 0)) = "Regular" Or Trim(oDBs_Head.GetValue("u_type", 0)) = "Excess" Or Trim(oDBs_Head.GetValue("u_type", 0)) = "Production Consumable" Then
                If objMatrix.VisualRowCount < 1 Then
                    oApplication.StatusBar.SetText("Please select sales order no, item code and semi finished goods", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                    Exit Function
                End If
                For i As Integer = 1 To objMatrix.VisualRowCount
                    If Trim(oDBs_Head.GetValue("u_type", 0)) = "Regular" Or Trim(oDBs_Head.GetValue("u_type", 0)) = "Production Consumable" Then
                        If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) <> "" And Trim(objForm.Items.Item("status").Specific.Selected.Value) <> "Closed" Then
                            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRS.DoQuery("Select Distinct u_per From [@GEN_CUST_BOM] A INNER JOIN [@GEN_CUST_BOM_D0] B ON A.DOCENTRY = B.DOCENTRY Where B.u_itemcode = '" + Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) + "' And A.u_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "'")
                            'If (CDbl(objMatrix.Columns.Item("rqstqty").Cells.Item(i).Specific.Value) + CDbl(objMatrix.Columns.Item("totis").Cells.Item(i).Specific.Value)) > CDbl(objMatrix.Columns.Item("reqdqty").Cells.Item(i).Specific.Value) + ((CDbl(objMatrix.Columns.Item("reqdqty").Cells.Item(i).Specific.Value) * CDbl(oRS.Fields.Item("u_per").Value)) / 100) Then
                            If (CDbl(objMatrix.Columns.Item("rqstqty").Cells.Item(i).Specific.Value) + CDbl(objMatrix.Columns.Item("totis").Cells.Item(i).Specific.Value)) > CDbl(objMatrix.Columns.Item("reqdqty").Cells.Item(i).Specific.Value) Then
                                oApplication.StatusBar.SetText(" " + Trim(oDBs_Head.GetValue("u_type", 0)) + " MRN Cannot be Generated for Item Code --> " + Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) + " ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                                Exit Function
                            End If
                        End If
                    End If
                    If Trim(oDBs_Head.GetValue("u_type", 0)) = "Excess" Then
                        'If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) <> "" Then
                        '    'Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '    'oRS.DoQuery("Select u_tol From OITM Where ItemCode = '" + Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) + "'")
                        '    'If (CDbl(objMatrix.Columns.Item("rqstqty").Cells.Item(i).Specific.Value) + CDbl(objMatrix.Columns.Item("totis").Cells.Item(i).Specific.Value)) <= CDbl(objMatrix.Columns.Item("reqdqty").Cells.Item(i).Specific.Value) + ((CDbl(objMatrix.Columns.Item("reqdqty").Cells.Item(i).Specific.Value) * CDbl(oRS.Fields.Item("u_tol").Value)) / 100) Then
                        '    '    oApplication.StatusBar.SetText("Excess MRN cannot be generated", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    '    Return False
                        '    '    Exit Function
                        '    'End If
                        'End If
                    End If
                Next
                For i As Integer = 1 To objMatrix.VisualRowCount
                    If Trim(objMatrix.Columns.Item("whs").Cells.Item(i).Specific.Value) = "" Then
                        oApplication.StatusBar.SetText("Please enter warehouse for all rows", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                        Exit Function
                    End If
                    If CDbl(objMatrix.Columns.Item("rqstqty").Cells.Item(i).Specific.Value) <= 0 Then
                        oApplication.StatusBar.SetText("Please enter requested quantity more than 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                        Exit Function
                    End If
                Next
                For i As Integer = 1 To objMatrix.VisualRowCount
                    If Trim(objForm.Items.Item("wipwhs").Specific.value) = Trim(objMatrix.Columns.Item("whs").Cells.Item(i).Specific.value) Then
                        WhsFlag = True
                    End If
                Next
            End If

            'Vijeesh
            If Trim(oDBs_Head.GetValue("u_type", 0)) = "Sampling" Then
                For i As Integer = 1 To objMatrix.VisualRowCount
                    If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) = "" Then
                        oApplication.StatusBar.SetText("Please select items", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                        Exit Function
                    End If
                    If Trim(objMatrix.Columns.Item("whs").Cells.Item(i).Specific.Value) = "" And Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.Value) <> "" Then
                        oApplication.StatusBar.SetText("Please enter warehouse for all rows", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                        Exit Function
                    End If
                Next
            End If
            'Vijeesh

            If Trim(oDBs_Head.GetValue("u_type", 0)) = "Consumable" Then
                For i As Integer = 1 To objMatrix.VisualRowCount
                    If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) = "" Then
                        oApplication.StatusBar.SetText("Please select items", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                        Exit Function
                    End If
                    If Trim(objMatrix.Columns.Item("whs").Cells.Item(i).Specific.Value) = "" And Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.Value) <> "" Then
                        oApplication.StatusBar.SetText("Please enter warehouse for all rows", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                        Exit Function
                    End If
                Next
                For i As Integer = 1 To objMatrix.VisualRowCount
                    If Trim(objForm.Items.Item("wipwhs").Specific.value) = Trim(objMatrix.Columns.Item("whs").Cells.Item(i).Specific.value) Then
                        WhsFlag = True
                    End If
                Next
            End If
            If WhsFlag = True Then
                oApplication.StatusBar.SetText("WIP warehouse and row level warehouse cannot be same", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            'RowWhs = objMatrix.Columns.Item("whs").Cells.Item(1).Specific.value
            'For i As Integer = 1 To objMatrix.VisualRowCount
            '    If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) <> "" Then
            '        If Trim(objMatrix.Columns.Item("whs").Cells.Item(i).Specific.value) <> RowWhs Then
            '            oApplication.StatusBar.SetText("All rows should have the same warehouse", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '            Return False
            '            Exit Function
            '        End If
            '    End If
            'Next
            If UpdMode = False Then
                If Trim(objForm.Items.Item("status").Specific.Selected.Value) = "Closed" Then
                    oApplication.StatusBar.SetText("Document cannot be closed in Add mode", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                    Exit Function
                End If
                If Trim(oDBs_Head.GetValue("u_empname", 0)) = "" Then
                    oApplication.StatusBar.SetText("Please select Employee Name", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                    Exit Function
                End If
            End If
            If UpdMode = True Then
                If Trim(objForm.Items.Item("status").Specific.Selected.Value) = "Closed" Then
                    If oApplication.MessageBox("You cannot make further issues for this document. Do you still want to continue? ", 2, "Yes", "No") = 2 Then
                        Return False
                        Exit Function
                    End If
                End If
            End If
            'AlertDocNum = objForm.Items.Item("docnum").Specific.value
            'AlertWhs = objMatrix.Columns.Item("whs").Cells.Item(1).Specific.value
            DocStatus = objForm.Items.Item("status").Specific.selected.value
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    'If pVal.ItemUID = "wipwhs" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '    If Trim(oDBs_Head.GetValue("u_wipwhs", 0)) <> "" And Trim(oDBs_Head.GetValue("u_type", 0)) = "Sampling" Then
                    '        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '        oRecordSet.DoQuery("Select WhsCode From OWHS WHere IsNull(U_sampwhs,'NO') = 'YES' And WhsCode = '" + Trim(oDBs_Head.GetValue("u_wipwhs", 0)) + "'")
                    '        If oRecordSet.RecordCount = 0 Then
                    '            oApplication.StatusBar.SetText("Please select Sampling Wareshouse for this transaction", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            BubbleEvent = False
                    '            Exit Sub
                    '        End If
                    '    End If
                    'End If
                    If pVal.ItemUID = "excsqty" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim _dbl_Ord_Qty As Double
                        _dbl_Ord_Qty = CDbl(objForm.Items.Item("ordrqty").Specific.value) + (CDbl(objForm.Items.Item("ordrqty").Specific.value) * 10 / 100)
                        'If CDbl(objForm.Items.Item("excsqty").Specific.value) > CDbl(objForm.Items.Item("ordrqty").Specific.value) Then
                        If CDbl(objForm.Items.Item("excsqty").Specific.value) > _dbl_Ord_Qty Then
                            oApplication.StatusBar.SetText("MRN Quantity cannot be greater than order quantity", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If

                    'Vijeesh
                    If pVal.ItemUID = "mtx" And pVal.ColUID = "whs" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.InnerEvent = False Then
                        If Trim(oDBs_Head.GetValue("u_type", 0)) = "Production Consumable" Or Trim(oDBs_Head.GetValue("u_type", 0)) = "Production Consumable Excess" Then
                            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            objMatrix = objForm.Items.Item("mtx").Specific
                            objMatrix.GetLineData(pVal.Row)
                            
                            oRS.DoQuery("Select SUM(M1.U_issued)IssQty from [@GEN_MREQ] M " _
                                                    & "inner join [@GEN_MREQ_D0] M1 on M.DocEntry = M1.DocEntry " _
                                                    & "where M.U_itemcode ='" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' " _
                                                    & "and M.U_sono ='" + Trim(objForm.Items.Item("sono").Specific.value) + "' " _
                                                    & "and M1.U_itemcode ='" + Trim(objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.Value) + "' " _
                                                    & "and M1.U_whs ='" + Trim(objMatrix.Columns.Item("whs").Cells.Item(pVal.Row).Specific.Value) + "' " _
                                                    & "and (M.U_type ='Regular' or M.U_type ='Production Consumable')")
                            Dim IssQty As Double
                            IssQty = oRS.Fields.Item("IssQty").Value
                            
                            oRS.DoQuery("Select SUM(M1.U_returned)RetQty from [@GEN_MREQ] M " _
                                                    & "inner join [@GEN_MREQ_D0] M1 on M.DocEntry = M1.DocEntry " _
                                                    & "where M.U_itemcode ='" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' " _
                                                    & "and M.U_sono ='" + Trim(objForm.Items.Item("sono").Specific.value) + "' " _
                                                    & "and M1.U_itemcode ='" + Trim(objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.Value) + "' " _
                                                    & "and M1.U_whs ='" + Trim(objMatrix.Columns.Item("whs").Cells.Item(pVal.Row).Specific.Value) + "' " _
                                                    & "and (M.U_type ='Regular' or M.U_type ='Production Consumable')")
                            Dim RetdQty As Double
                            RetdQty = oRS.Fields.Item("RetQty").Value
                            oDBs_Detail.SetValue("u_totis", pVal.Row - 1, IssQty - RetdQty)
                            objMatrix.SetLineData(pVal.Row)
                        End If
                    End If
                    If pVal.ItemUID = "excsqty" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.InnerEvent = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If Trim(oDBs_Head.GetValue("u_unit", 0)) = "" Then
                            oApplication.StatusBar.SetText("Please Select Unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If Trim(oDBs_Head.GetValue("u_type", 0)) = "Production Consumable" Or Trim(oDBs_Head.GetValue("u_type", 0)) = "Production Consumable Excess" Then
                            If Trim(oDBs_Head.GetValue("u_itemcode", 0)) = "" Then
                                oApplication.StatusBar.SetText("Please Select ItemCode", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            Else
                                objMatrix = objForm.Items.Item("mtx").Specific
                                objMatrix.Clear()
                                Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim RSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                Dim _str_Query As String = "Select B1.U_itemcode [ChildItem],OITM.ItemName,B1.U_qty [BOMQty],B1.U_process, 1 Quantity , " _
                                                            & "B1.U_per ,B1.U_uom,B1.U_process  from [@GEN_CUST_BOM]B " _
                                                            & "inner join [@GEN_CUST_BOM_D0] B1 on B.DocEntry = B1.DocEntry " _
                                                            & "inner join OITM on B1.U_itemcode = OITM.ItemCode " _
                                                            & "where B.U_itemcode ='" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' " _
                                                            & "and B1.U_process not in (Select U_process from [@GEN_PROD_PRCS] inner join [@GEN_PROD_PRCS_D0] on [@GEN_PROD_PRCS].CODe = [@GEN_PROD_PRCS_D0].Code and [@GEN_PROD_PRCS].U_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "') " _
                                                            & "and B1.U_issmthd ='M' and B.U_sono ='" + Trim(objForm.Items.Item("sono").Specific.value) + "' " _
                                                            & "and B1.U_status <>'DELETE' and isnull(B.U_closed,'N')='N' "
                                oRecordSet.DoQuery(_str_Query)
                                If oRecordSet.RecordCount <> "0" Then
                                    For i As Integer = 1 To oRecordSet.RecordCount
                                        objMatrix.AddRow(1, objMatrix.VisualRowCount)
                                        Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
                                        oDBs_Detail.SetValue("u_itemcode", i - 1, oRecordSet.Fields.Item("ChildItem").Value)
                                        oDBs_Detail.SetValue("u_itemname", i - 1, oRecordSet.Fields.Item("Itemname").Value)
                                        oDBs_Detail.SetValue("u_tol", i - 1, oRecordSet.Fields.Item("u_per").Value)
                                        oDBs_Detail.SetValue("u_uom", i - 1, oRecordSet.Fields.Item("U_uom").Value)
                                        oRSet.DoQuery("Select B.u_inwhs,B.u_outwhs,B.u_stwhs From [@GEN_UNIT_MST] A Inner Join [@GEN_UNIT_MST_D0] B ON A.CODE = B.CODE Where A.Name = '" + Trim(objForm.Items.Item("unit").Specific.value) + "' And B.u_process = '" + Trim(oRecordSet.Fields.Item("U_process").Value) + "'")
                                        oDBs_Detail.SetValue("u_whs", i - 1, oRSet.Fields.Item("u_stwhs").Value)
                                        Dim sfgordrqty As Double
                                        'sfgordrqty = (CDbl(objForm.Items.Item("ordrqty").Specific.value) / oRecordSet.Fields.Item("Quantity").Value) * oRecordSet.Fields.Item("BOMQty").Value
                                        'oDBs_Detail.SetValue("u_reqdqty", i - 1, (sfgordrqty))
                                        sfgordrqty = (CDbl(objForm.Items.Item("excsqty").Specific.value) / oRecordSet.Fields.Item("Quantity").Value) * oRecordSet.Fields.Item("BOMQty").Value
                                        oDBs_Detail.SetValue("u_reqdqty", i - 1, (sfgordrqty))
                                        oDBs_Detail.SetValue("u_rqstqty", i - 1, sfgordrqty)


                                        oRS.DoQuery("Select SUM(M1.U_issued)IssQty from [@GEN_MREQ] M " _
                                                    & "inner join [@GEN_MREQ_D0] M1 on M.DocEntry = M1.DocEntry " _
                                                    & "where M.U_itemcode ='" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' " _
                                                    & "and M.U_sono ='" + Trim(objForm.Items.Item("sono").Specific.value) + "' " _
                                                    & "and M1.U_itemcode ='" + oRecordSet.Fields.Item("ChildItem").Value + "' " _
                                                    & "and M1.U_whs ='" + oRSet.Fields.Item("u_stwhs").Value + "' " _
                                                    & "and (M.U_type ='Regular' or M.U_type ='Production Consumable')")
                                        Dim IssQty As Double
                                        IssQty = oRS.Fields.Item("IssQty").Value

                                        oRS.DoQuery("Select SUM(M1.U_returned)RetQty from [@GEN_MREQ] M " _
                                                    & "inner join [@GEN_MREQ_D0] M1 on M.DocEntry = M1.DocEntry " _
                                                    & "where M.U_itemcode ='" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' " _
                                                    & "and M.U_sono ='" + Trim(objForm.Items.Item("sono").Specific.value) + "' " _
                                                    & "and M1.U_itemcode ='" + oRecordSet.Fields.Item("ChildItem").Value + "' " _
                                                    & "and M1.U_whs ='" + oRSet.Fields.Item("u_stwhs").Value + "'" _
                                                    & "and (M.U_type ='Regular' or M.U_type ='Production Consumable')")
                                        Dim RetdQty As Double
                                        RetdQty = oRS.Fields.Item("RetQty").Value
                                        oDBs_Detail.SetValue("u_totis", i - 1, IssQty - RetdQty)
                                        oRS.DoQuery("Select OnHand From OITW Where Itemcode = '" + Trim(oRecordSet.Fields.Item("ChildItem").Value) + "' And WhsCode = '" + oRSet.Fields.Item("u_stwhs").Value + "'")
                                        oDBs_Detail.SetValue("u_totavlbl", i - 1, oRS.Fields.Item("OnHand").Value)
                                        objMatrix.SetLineData(i)
                                        oRecordSet.MoveNext()
                                    Next
                                Else
                                    oApplication.StatusBar.SetText("There is No Custom BOM for the Corresponding ItemCode and Sale Order No.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                    'Vijeesh
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            UpdMode = True
                        Else
                            UpdMode = False
                        End If
                        If Me.Validation(FormUID) = False Then BubbleEvent = False
                    ElseIf pVal.ItemUID = "1" And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim UserID As String
                        Dim DEntry As String
                        oRSet.DoQuery("Select UserID From  OUSR Where User_Code = '" + oCompany.UserName.ToString.Trim + "'")
                        UserID = oRSet.Fields.Item("UserID").Value
                        oRSet.DoQuery("Select Max(DocEntry) As 'DocEntry' From [@GEN_MREQ] Where UserSign = '" + UserID + "'")
                        DEntry = oRSet.Fields.Item("DocEntry").Value
                        oRSet.DoQuery("Select Distinct DocNum From [@GEN_MREQ] Where DocEntry = '" + DEntry + "'")
                        AlertDocNum = oRSet.Fields.Item("DocNum").Value
                        'oRSet.DoQuery("Select Top")
                        oApplication.MessageBox("The MRN No of the last added document: - " + AlertDocNum + "")
                        'Me.PostMessage(FormUID)
                        Me.SetDefault(FormUID)
                    End If
                    If pVal.ItemUID = "btnis" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If Trim(oDBs_Head.GetValue("u_type", 0)) = "Excess" Or Trim(oDBs_Head.GetValue("u_type", 0)) = "Production Consumable Excess" Then
                            If Trim(oDBs_Head.GetValue("u_approve", 0)) = "N" Or Trim(oDBs_Head.GetValue("u_approve", 0)) = "" Then
                                oApplication.StatusBar.SetText(" " + Trim(oDBs_Head.GetValue("u_type", 0)) + " MRN should be approved before it can be issued", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                        'Vijeesh
                        Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("mtx").Specific
                        Dim Flag As Boolean = False
                        Dim oCheck As SAPbouiCOM.CheckBox
                        For i As Integer = 1 To oMatrix.VisualRowCount
                            oCheck = oMatrix.Columns.Item("minchk").Cells.Item(i).Specific
                            If oCheck.Checked = False Then
                                Flag = True
                            Else
                                Flag = False
                                Exit Sub
                            End If
                        Next
                        If Flag = True Then
                            oApplication.StatusBar.SetText("Please Select Items", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        'Vijeesh
                    End If
                    If pVal.ItemUID = "btnis" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Dim DocEntry As String
                        Dim _str_Type As String = ""
                        Dim _str_Account As String = ""

                        DocEntry = objForm.Items.Item("docnum").Specific.value
                        _str_Type = objForm.Items.Item("type").Specific.value

                        objForm.Close()
                        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRSWHcount As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRSUnit As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        oRSWHcount.DoQuery("Select Distinct B.u_whs From [@GEN_MREQ] A Inner Join [@GEN_MREQ_D0] B On A.DocEntry = B.DocEntry Where A.DocNum = '" + DocEntry + "' And B.u_whs in (Select Distinct u_whs From [@GEN_WHS_USR] Where u_user = '" + oCompany.UserName.Trim + "') Group By B.u_whs")
                        For cnt As Integer = 1 To oRSWHcount.RecordCount
                            oRecordSet.DoQuery("Select Distinct A.u_type,A.u_sono,A.u_itemcode,A.u_sfgcode,A.u_wipwhs,A.DocNum,B.u_itemcode as 'RItemCode' ,B.u_rqstqty-IsNull(B.u_issued,0) as 'rqstqty',B.u_rqstqty,B.u_issued,B.u_whs,B.lineid From [@GEN_MREQ] A INNER JOIN [@GEN_MREQ_D0] B On A.DocEntry = B.DocEntry And A.Docnum = '" + DocEntry + "' And B.u_whs = '" + Trim(oRSWHcount.Fields.Item("u_whs").Value) + "' And IsNull(B.u_stat,'Open') != 'Closed' And isnull(B.U_minchk,'N')='Y'")
                            'Vijeesh
                            Dim _int_RecordCount As Integer = oRecordSet.RecordCount
                            If _int_RecordCount > 0 Then
                                'Inventory Transfer
                                If _str_Type = "Regular" Or _str_Type = "Excess" Or _str_Type = "Consumable" Then
                                    oApplication.ActivateMenuItem("3080")
                                    Dim ITForm As SAPbouiCOM.Form
                                    Dim ITMatrix As SAPbouiCOM.Matrix
                                    ITForm = oApplication.Forms.GetForm("940", oApplication.Forms.ActiveForm.TypeCount)
                                    ITMatrix = ITForm.Items.Item("23").Specific
                                    Try
                                        ITForm.Items.Item("18").Specific.Value = oRecordSet.Fields.Item("u_whs").Value
                                        ITForm.Items.Item("sono").Specific.value = oRecordSet.Fields.Item("u_sono").Value
                                        ITForm.Items.Item("type").Specific.value = oRecordSet.Fields.Item("u_type").Value
                                        ITForm.Items.Item("mrnno").Specific.value = oRecordSet.Fields.Item("DocNum").Value
                                        ITForm.Items.Item("itemcode").Specific.value = oRecordSet.Fields.Item("u_itemcode").Value
                                        ITForm.Items.Item("sfgcode").Specific.value = oRecordSet.Fields.Item("u_sfgcode").Value
                                        ITForm.Items.Item("isstyp").Specific.value = "I"
                                        ITMatrix.Columns.Item("U_mrnlid").Editable = True
                                        ITMatrix.Columns.Item("U_rqstqty").Editable = True
                                        ITMatrix.Columns.Item("U_issued").Editable = True
                                        For i As Integer = 1 To oRecordSet.RecordCount
                                            Try
                                                ITMatrix.Columns.Item("1").Cells.Item(i).Specific.value = oRecordSet.Fields.Item("RItemCode").Value
                                                ITMatrix.Columns.Item("5").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("u_wipwhs").Value
                                                If oRecordSet.Fields.Item("rqstqty").Value > 0 Then
                                                    ITMatrix.Columns.Item("10").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("rqstqty").Value
                                                Else
                                                    ITMatrix.Columns.Item("10").Cells.Item(i).Specific.Value = 1
                                                End If
                                                ITMatrix.Columns.Item("U_mrnlid").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("lineid").Value
                                                ITMatrix.Columns.Item("U_rqstqty").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("u_rqstqty").Value
                                                ITMatrix.Columns.Item("U_issued").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("u_issued").Value
                                                ITMatrix.Columns.Item("1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            Catch ex As Exception
                                            End Try
                                            oRecordSet.MoveNext()
                                        Next
                                        ITMatrix.Columns.Item("U_mrnlid").Editable = False
                                        ITMatrix.Columns.Item("U_rqstqty").Editable = False
                                        ITMatrix.Columns.Item("U_issued").Editable = False
                                    Catch ex As Exception
                                    End Try
                                    'Next
                                Else
                                    ' Goods Issue
                                    '25-6-2012 - modified for Sampling Process
                                    oApplication.ActivateMenuItem("3079")
                                    Dim GIForm As SAPbouiCOM.Form
                                    Dim GIMatrix As SAPbouiCOM.Matrix
                                    GIForm = oApplication.Forms.GetForm("720", oApplication.Forms.ActiveForm.TypeCount)
                                    GIMatrix = GIForm.Items.Item("13").Specific
                                    Try
                                        GIForm.Items.Item("sono").Specific.value = oRecordSet.Fields.Item("u_sono").Value
                                        GIForm.Items.Item("type").Specific.value = oRecordSet.Fields.Item("u_type").Value
                                        GIForm.Items.Item("mrnno").Specific.value = oRecordSet.Fields.Item("DocNum").Value
                                        GIForm.Items.Item("itemcode").Specific.value = oRecordSet.Fields.Item("u_itemcode").Value
                                        GIForm.Items.Item("sfgcode").Specific.value = oRecordSet.Fields.Item("u_sfgcode").Value
                                        GIForm.Items.Item("isstyp").Specific.value = "I"
                                        GIMatrix.Columns.Item("U_mrnlid").Editable = True
                                        GIMatrix.Columns.Item("U_rqstqty").Editable = True
                                        GIMatrix.Columns.Item("U_issued").Editable = True
                                        For i As Integer = 1 To oRecordSet.RecordCount
                                            Try
                                                GIMatrix.Columns.Item("1").Cells.Item(i).Specific.value = oRecordSet.Fields.Item("RItemCode").Value
                                                GIMatrix.Columns.Item("15").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("u_whs").Value
                                                If oRecordSet.Fields.Item("rqstqty").Value > 0 Then
                                                    GIMatrix.Columns.Item("9").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("rqstqty").Value
                                                Else
                                                    GIMatrix.Columns.Item("9").Cells.Item(i).Specific.Value = 1
                                                End If
                                                oRSUnit.DoQuery("Select isnull(U_unit,'')unit  from OWHS where WhsCode ='" + oRecordSet.Fields.Item("u_whs").Value + "'")
                                                If oRSUnit.Fields.Item("unit").Value = "UNIT1" Then
                                                    If _str_Type = "Sampling" Then
                                                        GIMatrix.Columns.Item("57").Cells.Item(i).Specific.Value = "400024-01"
                                                    Else
                                                        GIMatrix.Columns.Item("57").Cells.Item(i).Specific.Value = "400023-01"
                                                    End If
                                                ElseIf oRSUnit.Fields.Item("unit").Value = "UNIT2" Then
                                                    If _str_Type = "Sampling" Then
                                                        GIMatrix.Columns.Item("57").Cells.Item(i).Specific.Value = "400024-02"
                                                    Else
                                                        GIMatrix.Columns.Item("57").Cells.Item(i).Specific.Value = "400023-02"
                                                    End If
                                                Else
                                                    GIMatrix.Columns.Item("57").Cells.Item(i).Specific.Value = ""
                                                End If
                                                GIMatrix.Columns.Item("U_mrnlid").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("lineid").Value
                                                GIMatrix.Columns.Item("U_rqstqty").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("u_rqstqty").Value
                                                GIMatrix.Columns.Item("U_issued").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("u_issued").Value
                                                GIMatrix.Columns.Item("1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            Catch ex As Exception
                                            End Try
                                            oRecordSet.MoveNext()
                                        Next
                                        GIMatrix.Columns.Item("U_mrnlid").Editable = False
                                        GIMatrix.Columns.Item("U_rqstqty").Editable = False
                                        GIMatrix.Columns.Item("U_issued").Editable = False
                                    Catch ex As Exception
                                    End Try
                                End If
                            End If
                            oRSWHcount.MoveNext()
                        Next

                    End If
                    If pVal.ItemUID = "btnret" And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE) Then
                        Dim DocEntry As String
                        Dim _str_Type As String = ""
                        Dim _str_Account As String = ""
                        DocEntry = objForm.Items.Item("docnum").Specific.value
                        _str_Type = objForm.Items.Item("type").Specific.value
                        objForm.Close()

                        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRSWHcount As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRSUnit As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        If _str_Type = "Production Consumable" Or _str_Type = "Production Consumable Excess" Or _str_Type = "Sampling" Then
                            'Goods Reciept
                            '25-6-2012 modified for Sampling MRN
                            oRSWHcount.DoQuery("Select Distinct B.u_whs From [@GEN_MREQ] A Inner Join [@GEN_MREQ_D0] B On A.DocEntry = B.DocEntry Where A.DocNum = '" + DocEntry + "' And B.u_whs in (Select Distinct u_whs From [@GEN_WHS_USR] Where u_user = '" + oCompany.UserName.ToString.Trim + "') Group By B.u_whs")
                            While Not oRSWHcount.EoF
                                'objForm.Close()
                                oRecordSet.DoQuery("Select Distinct A.u_type,A.u_sono,A.u_itemcode,A.u_sfgcode,A.u_wipwhs,B.U_whs ,A.DocNum,B.u_itemcode as 'RItemCode' , " _
                                                        & "IsNull(B.u_issued,0) - IsNull(B.u_returned,0) as 'retqty', " _
                                                        & "B.u_rqstqty, B.u_issued, B.u_whs, B.lineid " _
                                                        & "From [@GEN_MREQ] A INNER JOIN [@GEN_MREQ_D0] B On A.DocEntry = B.DocEntry " _
                                                        & "And A.Docnum = '" + DocEntry + "' " _
                                                        & "And IsNull(B.u_issued,0) - IsNUll(B.u_returned,0) > 0 " _
                                                        & "And B.U_whs = '" + Trim(oRSWHcount.Fields.Item("u_whs").Value) + "'")
                                While Not oRecordSet.EoF
                                    oApplication.ActivateMenuItem("3078")
                                    Dim GIForm As SAPbouiCOM.Form
                                    Dim GIMatrix As SAPbouiCOM.Matrix
                                    GIForm = oApplication.Forms.GetForm("721", oApplication.Forms.ActiveForm.TypeCount)
                                    GIMatrix = GIForm.Items.Item("13").Specific
                                    Try
                                        GIForm.Items.Item("sono").Specific.value = oRecordSet.Fields.Item("u_sono").Value
                                        GIForm.Items.Item("type").Specific.value = oRecordSet.Fields.Item("u_type").Value
                                        GIForm.Items.Item("mrnno").Specific.value = oRecordSet.Fields.Item("DocNum").Value
                                        GIForm.Items.Item("itemcode").Specific.value = oRecordSet.Fields.Item("u_itemcode").Value
                                        GIForm.Items.Item("sfgcode").Specific.value = oRecordSet.Fields.Item("u_sfgcode").Value
                                        GIForm.Items.Item("isstyp").Specific.value = "R"
                                        GIMatrix.Columns.Item("U_mrnlid").Editable = True
                                        GIMatrix.Columns.Item("U_rqstqty").Editable = True
                                        GIMatrix.Columns.Item("U_issued").Editable = True
                                        For i As Integer = 1 To oRecordSet.RecordCount
                                            Try
                                                GIMatrix.Columns.Item("1").Cells.Item(i).Specific.value = oRecordSet.Fields.Item("RItemCode").Value
                                                GIMatrix.Columns.Item("15").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("u_whs").Value
                                                If oRecordSet.Fields.Item("retqty").Value > 0 Then
                                                    GIMatrix.Columns.Item("9").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("retqty").Value
                                                Else
                                                    GIMatrix.Columns.Item("9").Cells.Item(i).Specific.Value = 1
                                                End If
                                                oRSUnit.DoQuery("Select isnull(U_unit,'')unit  from OWHS where WhsCode ='" + oRecordSet.Fields.Item("u_whs").Value + "'")
                                                If oRSUnit.Fields.Item("unit").Value = "UNIT1" Then
                                                    If _str_Type = "Sampling" Then
                                                        GIMatrix.Columns.Item("57").Cells.Item(i).Specific.Value = "400024-01"
                                                    Else
                                                        GIMatrix.Columns.Item("57").Cells.Item(i).Specific.Value = "400023-01"
                                                    End If
                                                ElseIf oRSUnit.Fields.Item("unit").Value = "UNIT2" Then
                                                    If _str_Type = "Sampling" Then
                                                        GIMatrix.Columns.Item("57").Cells.Item(i).Specific.Value = "400024-02"
                                                    Else
                                                        GIMatrix.Columns.Item("57").Cells.Item(i).Specific.Value = "400023-02"
                                                    End If
                                                Else
                                                    GIMatrix.Columns.Item("57").Cells.Item(i).Specific.Value = ""
                                                End If
                                                'GIMatrix.Columns.Item("57").Cells.Item(i).Specific.Value = _str_Account
                                                GIMatrix.Columns.Item("U_mrnlid").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("lineid").Value
                                                GIMatrix.Columns.Item("1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            Catch ex As Exception
                                            End Try
                                            oRecordSet.MoveNext()
                                        Next
                                        GIMatrix.Columns.Item("U_mrnlid").Editable = False
                                        GIMatrix.Columns.Item("U_rqstqty").Editable = False
                                        GIMatrix.Columns.Item("U_issued").Editable = False
                                    Catch ex As Exception
                                    End Try
                                End While
                                oRSWHcount.MoveNext()
                            End While
                        Else
                            oRSWHcount.DoQuery("Select Distinct A.u_wipwhs From [@GEN_MREQ] A Inner Join [@GEN_MREQ_D0] B On A.DocEntry = B.DocEntry Where A.DocNum = '" + DocEntry + "' And A.u_wipwhs in (Select Distinct u_whs From [@GEN_WHS_USR] Where u_user = '" + oCompany.UserName.ToString.Trim + "') Group By A.u_wipwhs")
                            While Not oRSWHcount.EoF
                                objForm.Close()
                                oRecordSet.DoQuery("Select Distinct A.u_type,A.u_sono,A.u_itemcode,A.u_sfgcode,A.u_wipwhs,A.DocNum,B.u_itemcode as 'RItemCode' ,IsNull(B.u_issued,0) - IsNull(B.u_returned,0) as 'retqty',B.u_rqstqty,B.u_issued,B.u_whs,B.lineid From [@GEN_MREQ] A INNER JOIN [@GEN_MREQ_D0] B On A.DocEntry = B.DocEntry And A.Docnum = '" + DocEntry + "' And A.u_wipwhs = '" + Trim(oRSWHcount.Fields.Item("u_wipwhs").Value) + "' And IsNull(B.u_issued,0) - IsNUll(B.u_returned,0) > 0")
                                While Not oRecordSet.EoF
                                    oApplication.ActivateMenuItem("3080")
                                    Dim ITForm As SAPbouiCOM.Form
                                    Dim ITMatrix As SAPbouiCOM.Matrix
                                    ITForm = oApplication.Forms.GetForm("940", oApplication.Forms.ActiveForm.TypeCount)
                                    ITMatrix = ITForm.Items.Item("23").Specific
                                    Try
                                        ITForm.Items.Item("18").Specific.Value = oRecordSet.Fields.Item("u_wipwhs").Value
                                        ITForm.Items.Item("sono").Specific.value = oRecordSet.Fields.Item("u_sono").Value
                                        ITForm.Items.Item("type").Specific.value = oRecordSet.Fields.Item("u_type").Value
                                        ITForm.Items.Item("mrnno").Specific.value = oRecordSet.Fields.Item("DocNum").Value
                                        ITForm.Items.Item("itemcode").Specific.value = oRecordSet.Fields.Item("u_itemcode").Value
                                        ITForm.Items.Item("sfgcode").Specific.value = oRecordSet.Fields.Item("u_sfgcode").Value
                                        ITForm.Items.Item("isstyp").Specific.value = "R"
                                        ITMatrix.Columns.Item("U_mrnlid").Editable = True
                                        ITMatrix.Columns.Item("U_rqstqty").Editable = True
                                        ITMatrix.Columns.Item("U_issued").Editable = True
                                        For i As Integer = 1 To oRecordSet.RecordCount
                                            Try
                                                ITMatrix.Columns.Item("1").Cells.Item(i).Specific.value = oRecordSet.Fields.Item("RItemCode").Value
                                                ITMatrix.Columns.Item("5").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("u_whs").Value
                                                If oRecordSet.Fields.Item("retqty").Value > 0 Then
                                                    ITMatrix.Columns.Item("10").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("retqty").Value
                                                Else
                                                    ITMatrix.Columns.Item("10").Cells.Item(i).Specific.Value = 1
                                                End If
                                                ITMatrix.Columns.Item("U_mrnlid").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("lineid").Value
                                                ITMatrix.Columns.Item("1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            Catch ex As Exception
                                            End Try
                                            oRecordSet.MoveNext()
                                        Next
                                        ITMatrix.Columns.Item("U_mrnlid").Editable = False
                                        ITMatrix.Columns.Item("U_rqstqty").Editable = False
                                        ITMatrix.Columns.Item("U_issued").Editable = False
                                    Catch ex As Exception
                                    End Try
                                End While
                                oRSWHcount.MoveNext()
                            End While
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "type" And pVal.BeforeAction = False Then
                        If Trim(oDBs_Head.GetValue("u_type", 0)) = "Regular" Or Trim(oDBs_Head.GetValue("u_type", 0)) = "Excess" Then
                            objForm.Items.Item("sono").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("itemcode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("sfgcode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("wipwhs").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("ETSLEMP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            If Trim(oDBs_Head.GetValue("u_type", 0)) = "Regular" Then
                                objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                oDBs_Head.SetValue("u_approve", 0, "Y")
                            End If
                            If Trim(oDBs_Head.GetValue("u_type", 0)) = "Excess" Then
                                oDBs_Head.SetValue("u_approve", 0, "N")
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet.DoQuery("Select USER_CODE From [OUSR] Where USER_CODE = '" + oCompany.UserName.ToString.Trim + "' And IsNull(u_approve,'N') = 'Y'")
                                If oRSet.RecordCount > 0 Then
                                    objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                Else
                                    objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                End If
                            End If
                            oDBs_Head.SetValue("u_itemcode", 0, "")
                            oDBs_Head.SetValue("u_itemname", 0, "")
                            oDBs_Head.SetValue("u_ordrqty", 0, "")
                            oDBs_Head.SetValue("u_sfgcode", 0, "")
                            oDBs_Head.SetValue("u_sfgname", 0, "")
                            oDBs_Head.SetValue("u_whs", 0, "")
                            oDBs_Head.SetValue("u_wipwhs", 0, "")
                            oDBs_Head.SetValue("u_sono", 0, "")
                            oDBs_Head.SetValue("u_soentry", 0, "")
                            oDBs_Head.SetValue("u_soref", 0, "")
                            oDBs_Head.SetValue("u_excsqty", 0, "")
                            oDBs_Head.SetValue("u_process", 0, "")
                            oDBs_Head.SetValue("u_unit", 0, "")
                            objMatrix = objForm.Items.Item("mtx").Specific
                            objMatrix.Columns.Item("reqdqty").Editable = False
                            objMatrix.Columns.Item("totis").Editable = False
                            If Trim(oDBs_Head.GetValue("u_type", 0)) = "Excess" Or Trim(oDBs_Head.GetValue("u_type", 0)) = "Regular" Then
                                objForm.Items.Item("excsqty").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            Else
                                objForm.Items.Item("excsqty").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            End If
                            objMatrix.Clear()
                        End If
                        If Trim(oDBs_Head.GetValue("u_type", 0)) = "Consumable" Or Trim(oDBs_Head.GetValue("u_type", 0)) = "Sampling" Then
                            objForm.Items.Item("sono").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Items.Item("itemcode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("sfgcode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Items.Item("wipwhs").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("excsqty").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Items.Item("ETSLEMP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            oDBs_Head.SetValue("u_approve", 0, "Y")
                            oDBs_Head.SetValue("u_itemcode", 0, "")
                            oDBs_Head.SetValue("u_itemname", 0, "")
                            oDBs_Head.SetValue("u_ordrqty", 0, "")
                            oDBs_Head.SetValue("u_excsqty", 0, "")
                            oDBs_Head.SetValue("u_sfgcode", 0, "")
                            oDBs_Head.SetValue("u_sfgname", 0, "")
                            oDBs_Head.SetValue("u_whs", 0, "")
                            oDBs_Head.SetValue("u_wipwhs", 0, "")
                            oDBs_Head.SetValue("u_sono", 0, "")
                            oDBs_Head.SetValue("u_soentry", 0, "")
                            oDBs_Head.SetValue("u_soref", 0, "")
                            oDBs_Head.SetValue("u_process", 0, "")
                            oDBs_Head.SetValue("u_unit", 0, "")
                            objMatrix = objForm.Items.Item("mtx").Specific
                            objMatrix.Clear()
                            objMatrix.AddRow(1, objMatrix.VisualRowCount)
                            Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
                            objMatrix = objForm.Items.Item("mtx").Specific
                            objMatrix.Columns.Item("reqdqty").Editable = False
                            objMatrix.Columns.Item("totis").Editable = False
                        End If
                        '---> Vijeesh
                        If Trim(oDBs_Head.GetValue("u_type", 0)) = "Production Consumable" Or Trim(oDBs_Head.GetValue("u_type", 0)) = "Production Consumable Excess" Then
                            objForm.Items.Item("sono").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("itemcode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("sfgcode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Items.Item("wipwhs").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("excsqty").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("ETSLEMP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            If Trim(oDBs_Head.GetValue("u_type", 0)) = "Production Consumable" Then
                                objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                oDBs_Head.SetValue("u_approve", 0, "Y")
                            End If
                            If Trim(oDBs_Head.GetValue("u_type", 0)) = "Production Consumable Excess" Then
                                oDBs_Head.SetValue("u_approve", 0, "N")
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet.DoQuery("Select USER_CODE From [OUSR] Where USER_CODE = '" + oCompany.UserName.ToString.Trim + "' And IsNull(u_approve,'N') = 'Y'")
                                If oRSet.RecordCount > 0 Then
                                    objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                Else
                                    objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                End If
                            End If
                            oDBs_Head.SetValue("u_itemcode", 0, "")
                            oDBs_Head.SetValue("u_itemname", 0, "")
                            oDBs_Head.SetValue("u_ordrqty", 0, "")
                            oDBs_Head.SetValue("u_excsqty", 0, "")
                            oDBs_Head.SetValue("u_sfgcode", 0, "")
                            oDBs_Head.SetValue("u_sfgname", 0, "")
                            oDBs_Head.SetValue("u_whs", 0, "")
                            oDBs_Head.SetValue("u_wipwhs", 0, "")
                            oDBs_Head.SetValue("u_sono", 0, "")
                            oDBs_Head.SetValue("u_soentry", 0, "")
                            oDBs_Head.SetValue("u_soref", 0, "")
                            oDBs_Head.SetValue("u_process", 0, "")
                            oDBs_Head.SetValue("u_unit", 0, "")
                            objMatrix = objForm.Items.Item("mtx").Specific
                            objMatrix.Columns.Item("reqdqty").Editable = False
                            objMatrix.Columns.Item("totis").Editable = False
                        End If
                        '---> Vijeesh
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If (pVal.ItemUID = "sono" Or pVal.ItemUID = "ETSLEMP" Or pVal.ItemUID = "itemcode" Or pVal.ItemUID = "excsqty" Or pVal.ItemUID = "isstype" Or pVal.ItemUID = "sfgcode" Or pVal.ItemUID = "series" Or pVal.ItemUID = "docdate" Or pVal.ItemUID = "unit") _
                       And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If objForm.DataSources.DBDataSources.Item("@GEN_MREQ").GetValue("u_type", 0) = Nothing Then
                            oApplication.StatusBar.SetText("Please select Document Type", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                        End If
                    End If
                    If pVal.ItemUID = "mtx" And pVal.BeforeAction = True Then
                        If oDBs_Head.GetValue("u_itemcode", 0) = "" And Trim(oDBs_Head.GetValue("u_type", 0)) <> "Consumable" Then
                            oApplication.StatusBar.SetText("Please select item code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                        End If
                    End If
                    'Vijeesh
                    If pVal.ItemUID = "mtx" And pVal.ColUID = "minchk" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) And pVal.BeforeAction = True Then
                        Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("mtx").Specific
                        BubbleEvent = False
                    End If
                    'Vijeesh
                    'If pVal.ItemUID = "mtx" And pVal.ColUID <> "minchk" And pVal.ColUID = "itemcode" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = True Then
                    '    Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("mtx").Specific
                    '    If oMatrix.Columns.Item("issued").Cells.Item(pVal.Row).Specific.value > 0 Then
                    '        oApplication.StatusBar.SetText("Cannot modify document", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        BubbleEvent = False
                    '        Exit Sub
                    '    End If
                    'End If
                    If pVal.ItemUID = "mtx" And pVal.ColUID <> "minchk" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = True Then
                        'Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("mtx").Specific
                        'Dim Flag As Boolean = False
                        'For i As Integer = 1 To oMatrix.VisualRowCount
                        '    If oMatrix.Columns.Item("issued").Cells.Item(i).Specific.value > 0 Then
                        '        Flag = True
                        '    Else
                        '        Flag = False
                        '    End If
                        'Next
                        'If Flag = True Then
                        '    oApplication.StatusBar.SetText("Cannot modify document", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If
                        oApplication.StatusBar.SetText("Cannot modify document", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False
                        Exit Sub
                    End If
                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    If (pVal.CharPressed = 9 Or pVal.CharPressed = 13) And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = True Then
                        If pVal.ItemUID <> "doctype" Then
                            If objForm.DataSources.DBDataSources.Item("@GEN_MREQ").GetValue("u_type", 0) = Nothing Then
                                oApplication.StatusBar.SetText("Please select Document Type", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                            End If
                        End If
                    End If
                    If pVal.ItemUID = "itemcode" Then
                        oCombo = objForm.Items.Item("type").Specific
                        If oCombo.Selected.Value = "Regular" Or oCombo.Selected.Value = "Excess" Then
                            If objForm.Items.Item("sono").Specific.value = "" Then
                                oApplication.StatusBar.SetText("Please enter Document Number", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If
                    If pVal.ItemUID = "mtx" And pVal.ColUID = "rqstqty" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = True Then
                        Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("mtx").Specific
                        Dim Flag As Boolean = False
                        For i As Integer = 1 To oMatrix.VisualRowCount
                            If oMatrix.Columns.Item("issued").Cells.Item(i).Specific.value > 0 Then
                                Flag = True
                            End If
                        Next
                        If Flag = True Then
                            oApplication.StatusBar.SetText("Cannot modify document", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
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
                        If oCFL.UniqueID = "SOCFL" Then
                            Me.FilterSO(FormUID)
                        End If
                        If oCFL.UniqueID = "ITCFL" Then
                            Me.FilterSOItems(FormUID)
                        End If
                        If oCFL.UniqueID = "ITRCFL" Then
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MREQ")
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_MREQ_D0")
                            If Trim(oDBs_Head.GetValue("u_wipwhs", 0)) = "" And Trim(oDBs_Head.GetValue("U_type", 0)) <> "Sampling" And Trim(oDBs_Head.GetValue("U_type", 0)) <> "Production Consumable" And Trim(oDBs_Head.GetValue("U_type", 0)) <> "Production Consumable Excess" Then
                                oApplication.StatusBar.SetText("Please select WIP Warehouse", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            If Trim(oDBs_Head.GetValue("U_type", 0)) = "Consumable" Then
                                Me.FilterItems(FormUID)
                            End If
                            If Trim(oDBs_Head.GetValue("U_type", 0)) = "Sampling" Then
                                Me.FilterItemsSample(FormUID)
                            End If
                        End If


                        If oCFL.UniqueID = "BOMCFL" Then
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MREQ")
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_MREQ_D0")
                            If Trim(oDBs_Head.GetValue("u_itemcode", 0)) = "" And (Trim(oDBs_Head.GetValue("u_type", 0)) = "Regular" Or Trim(oDBs_Head.GetValue("u_type", 0)) = "Excess") Then
                                oApplication.StatusBar.SetText("Please select item code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            If Trim(oDBs_Head.GetValue("u_ordrqty", 0)) = 0 And (Trim(oDBs_Head.GetValue("u_type", 0)) = "Regular" Or Trim(oDBs_Head.GetValue("u_type", 0)) = "Excess") Then
                                oApplication.StatusBar.SetText("Please select ordered quantity", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            If Trim(oDBs_Head.GetValue("u_unit", 0)) = "" Then
                                oApplication.StatusBar.SetText("Please select Unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            If Trim(oDBs_Head.GetValue("u_type", 0)) = "Excess" Or Trim(oDBs_Head.GetValue("u_type", 0)) = "Regular" Then
                                If Trim(oDBs_Head.GetValue("u_excsqty", 0)) <= 0 Then
                                    oApplication.StatusBar.SetText("Please enter the MRN quantity", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                            objForm.Items.Item("excsqty").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            Me.FilterBOMItems(FormUID)
                        End If
                    Else
                        If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MREQ")
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_MREQ_D0")
                            If oCFL.UniqueID = "ITCFL" Then
                                oDBs_Head.SetValue("u_itemcode", 0, oDT.GetValue("ItemCode", 0))
                                oDBs_Head.SetValue("u_itemname", 0, oDT.GetValue("ItemName", 0))
                                oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("Select Sum(B.Quantity) As 'Qty' From ORDR A Inner Join RDR1 B On A.DocEntry = B.DocEntry And A.DocEntry = '" + Trim(objForm.Items.Item("soentry").Specific.value) + "' And B.ItemCode = '" + oDT.GetValue("ItemCode", 0) + "'")
                                oDBs_Head.SetValue("u_ordrqty", 0, oRecordSet.Fields.Item("Qty").Value)
                            End If
                            If oCFL.UniqueID = "W1CFL" Then
                                oDBs_Head.SetValue("u_whs", 0, oDT.GetValue("WhsCode", 0))
                            End If
                            If oCFL.UniqueID = "W2CFL" Then
                                oDBs_Head.SetValue("u_wipwhs", 0, oDT.GetValue("WhsCode", 0))
                            End If
                            '' Added By vivek for Buyer'S Detail
                            If oCFL.UniqueID = "EMP_ID" And oDBs_Head.GetValue("u_type", 0).Trim = "Sampling" Then
                                oDBs_Head.SetValue("U_EMP_ID", 0, oDT.GetValue("SlpCode", 0))
                                oDBs_Head.SetValue("U_EMP_NAME", 0, oDT.GetValue("SlpName", 0))
                            End If
                            If oCFL.UniqueID = "UNTCFL" Then
                                oDBs_Head.SetValue("u_unit", 0, oDT.GetValue("Name", 0))
                            End If
                            If oCFL.UniqueID = "SOCFL" Then
                                oDBs_Head.SetValue("u_sono", 0, oDT.GetValue("DocNum", 0))
                                oDBs_Head.SetValue("u_soentry", 0, oDT.GetValue("DocEntry", 0))
                                oDBs_Head.SetValue("u_soref", 0, oDT.GetValue("NumAtCard", 0))
                                oDBs_Head.SetValue("u_itemcode", 0, "")
                                oDBs_Head.SetValue("u_itemname", 0, "")
                                oDBs_Head.SetValue("u_ordrqty", 0, "")
                                oDBs_Head.SetValue("u_sfgcode", 0, "")
                                oDBs_Head.SetValue("u_sfgname", 0, "")
                                oDBs_Head.SetValue("u_excsqty", 0, "")
                                oDBs_Head.SetValue("u_whs", 0, "")
                                oDBs_Head.SetValue("u_unit", 0, "")
                                oDBs_Head.SetValue("u_process", 0, "")
                                oDBs_Head.SetValue("u_wipwhs", 0, "")
                                oDBs_Head.SetValue("U_EMP_ID", 0, oDT.GetValue("SlpCode", 0).ToString.Trim)

                                Dim oRecordset As SAPbobsCOM.Recordset

                                oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                oRecordset.DoQuery("Select SlpName From OSLP Where SlpCode = '" + oDT.GetValue("SlpCode", 0).ToString.Trim + "'")
                                oDBs_Head.SetValue("U_EMP_NAME", 0, oRecordset.Fields.Item("SlpName").Value.ToString.Trim)

                                objMatrix = objForm.Items.Item("mtx").Specific
                                objMatrix.Clear()
                            End If
                            If oCFL.UniqueID = "BOMCFL" Then
                                Dim StWhs As String
                                objMatrix = objForm.Items.Item("mtx").Specific
                                objMatrix.Clear()
                                oDBs_Head.SetValue("u_sfgcode", 0, oDT.GetValue("ItemCode", 0))
                                oDBs_Head.SetValue("u_sfgname", 0, oDT.GetValue("ItemName", 0))
                                Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim RSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet.DoQuery("Select B.u_process From [@GEN_PROD_PRCS] A Inner Join [@GEN_PROD_PRCS_D0] B On A.Code = B.Code Where A.u_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' And B.u_itemcode = '" + Trim(oDT.GetValue("ItemCode", 0)) + "'")
                                oDBs_Head.SetValue("u_process", 0, oRSet.Fields.Item("u_process").Value)
                                If Trim(oDBs_Head.GetValue("u_itemcode", 0)) <> Trim(oDT.GetValue("ItemCode", 0)) Then
                                    oRS.DoQuery("Select B.u_inwhs,B.u_outwhs,B.u_stwhs From [@GEN_UNIT_MST] A Inner Join [@GEN_UNIT_MST_D0] B ON A.CODE = B.CODE Where A.Name = '" + Trim(objForm.Items.Item("unit").Specific.value) + "' And B.u_process = '" + Trim(oRSet.Fields.Item("u_process").Value) + "'")
                                    StWhs = oRS.Fields.Item("u_stwhs").Value
                                    oDBs_Head.SetValue("u_whs", 0, oRS.Fields.Item("u_outwhs").Value)
                                    oDBs_Head.SetValue("u_wipwhs", 0, oRS.Fields.Item("u_inwhs").Value)
                                End If
                                If Trim(oDBs_Head.GetValue("u_itemcode", 0)) = Trim(oDT.GetValue("ItemCode", 0)) Then
                                    oRS.DoQuery("Select A.u_inwhs,A.u_outwhs,A.u_stwhs From [@GEN_PROD_PRCS] A Where A.u_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "'")
                                    StWhs = oRS.Fields.Item("u_stwhs").Value
                                    oDBs_Head.SetValue("u_whs", 0, oRS.Fields.Item("u_outwhs").Value)
                                    oDBs_Head.SetValue("u_wipwhs", 0, oRS.Fields.Item("u_inwhs").Value)
                                End If
                                Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("Select Distinct A.Code,B.ItemName,IsNull(A.u_tol,0) As 'Tol',B.InvntryUOM,X.Qauntity,A.Quantity,A.Warehouse From OITT X Inner Join ITT1 A On X.Code = A.Father Inner Join OITM B ON A.Code = B.ItemCode Where A.Father = '" + Trim(oDT.GetValue("ItemCode", 0)) + "' And A.IssueMthd = 'M' Order By A.Warehouse")
                                For i As Integer = 1 To oRecordSet.RecordCount
                                    objMatrix.AddRow(1, objMatrix.VisualRowCount)
                                    Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
                                    oDBs_Detail.SetValue("u_itemcode", i - 1, oRecordSet.Fields.Item("Code").Value)
                                    oDBs_Detail.SetValue("u_itemname", i - 1, oRecordSet.Fields.Item("Itemname").Value)
                                    oRS.DoQuery("Select B.u_per From [@GEN_CUST_BOM] A INNER JOIN [@GEN_CUST_BOM_D0] B On A.DocEntry = B.DocEntry Where A.u_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' And B.u_itemcode = '" + Trim(oRecordSet.Fields.Item("Code").Value) + "'")
                                    oDBs_Detail.SetValue("u_tol", i - 1, oRS.Fields.Item("u_per").Value)
                                    oDBs_Detail.SetValue("u_uom", i - 1, oRecordSet.Fields.Item("InvntryUOM").Value)
                                    'If Trim(objForm.Items.Item("itemcode").Specific.value) = Trim(oDT.GetValue("ItemCode", 0)) Then
                                    '    oDBs_Detail.SetValue("u_reqdqty", i - 1, (CDbl(objForm.Items.Item("ordrqty").Specific.value) / oRecordSet.Fields.Item("Qauntity").Value) * oRecordSet.Fields.Item("Quantity").Value)
                                    'End If
                                    'If Trim(objForm.Items.Item("itemcode").Specific.value) <> Trim(oDT.GetValue("ItemCode", 0)) Then
                                    '    Dim sfgordrqty As Double
                                    '    RSet.DoQuery("Exec Child_BOM '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "','" + Trim(oDT.GetValue("ItemCode", 0)) + "'")
                                    '    sfgordrqty = (CDbl(objForm.Items.Item("ordrqty").Specific.value) / RSet.Fields.Item("Qauntity").Value) * RSet.Fields.Item("Quantity").Value
                                    '    oDBs_Detail.SetValue("u_reqdqty", i - 1, (sfgordrqty / oRecordSet.Fields.Item("Qauntity").Value) * oRecordSet.Fields.Item("Quantity").Value)
                                    'End If
                                    If oDBs_Head.GetValue("u_itemcode", 0) = Trim(oDT.GetValue("ItemCode", 0)) Then
                                        oDBs_Detail.SetValue("u_whs", i - 1, oRecordSet.Fields.Item("Warehouse").Value)
                                        RSet.DoQuery("Select Warehouse From ITT1 Where Father = '" + oDT.GetValue("ItemCode", 0) + "' And Code = '" + Trim(oRecordSet.Fields.Item("Code").Value) + "'")
                                        oRS.DoQuery("Select OnHand From OITW Where Itemcode = '" + Trim(oDT.GetValue("ItemCode", 0)) + "' And WhsCode = '" + Trim(oRecordSet.Fields.Item("Warehouse").Value) + "'")
                                        oDBs_Detail.SetValue("u_totavlbl", i - 1, oRS.Fields.Item("OnHand").Value)
                                        oRS.DoQuery("Select OnHand From OITW Where Itemcode = '" + Trim(oDT.GetValue("ItemCode", 0)) + "' And WhsCode = '" + Trim(objForm.Items.Item("wipwhs").Specific.value) + "'")
                                        oDBs_Detail.SetValue("u_wipavlbl", i - 1, oRS.Fields.Item("OnHand").Value)
                                    End If
                                    If oDBs_Head.GetValue("u_itemcode", 0) <> Trim(oDT.GetValue("ItemCode", 0)) Then
                                        RSet.DoQuery("Select Warehouse From ITT1 Where Father = '" + oDT.GetValue("ItemCode", 0) + "' And Code = '" + Trim(oRecordSet.Fields.Item("Code").Value) + "'")
                                        oDBs_Detail.SetValue("u_whs", i - 1, StWhs)
                                        oRS.DoQuery("Select OnHand From OITW Where Itemcode = '" + Trim(oRecordSet.Fields.Item("Code").Value) + "' And WhsCode = '" + StWhs + "'")
                                        oDBs_Detail.SetValue("u_totavlbl", i - 1, oRS.Fields.Item("OnHand").Value)
                                        oRS.DoQuery("Select OnHand From OITW Where Itemcode = '" + Trim(oRecordSet.Fields.Item("Code").Value) + "' And WhsCode = '" + Trim(objForm.Items.Item("wipwhs").Specific.value) + "'")
                                        oDBs_Detail.SetValue("u_wipavlbl", i - 1, oRS.Fields.Item("OnHand").Value)
                                    End If
                                    'oDBs_Detail.SetValue("u_whs", i - 1, oRecordSet.Fields.Item("Warehouse").Value)
                                    'oRS.DoQuery("Select OnHand From OITW Where Itemcode = '" + Trim(oDT.GetValue("ItemCode", 0)) + "' And WhsCOde = '" + Trim(oRecordSet.Fields.Item("Warehouse").Value) + "'")
                                    'oDBs_Detail.SetValue("u_totavlbl", i - 1, oRS.Fields.Item("OnHand").Value)
                                    'RSet.DoQuery("Select Warehouse From ITT1 Where Father = '" + oDT.GetValue("ItemCode", 0) + "' And Code = '" + Trim(oRecordSet.Fields.Item("Code").Value) + "'")
                                    oRS.DoQuery("Select IsNull(Sum(B.Quantity),0) AS 'IssdQty' From OWTR A Inner Join WTR1 B On A.DocEntry = B.DocEntry Where A.u_sono = '" + Trim(objForm.Items.Item("sono").Specific.value) + "' and A.u_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' And A.u_sfgcode = '" + Trim(objForm.Items.Item("sfgcode").Specific.value) + "' And A.Filler = '" + StWhs + "' And B.WhsCode = '" + Trim(objForm.Items.Item("wipwhs").Specific.value) + "' And B.ItemCode = '" + Trim(oRecordSet.Fields.Item("Code").Value) + "'")
                                    Dim IssQty As Double
                                    IssQty = oRS.Fields.Item("IssdQty").Value
                                    oRS.DoQuery("Select IsNull(Sum(B.Quantity),0) AS 'RetdQty' From OWTR A Inner Join WTR1 B On A.DocEntry = B.DocEntry Where A.u_sono = '" + Trim(objForm.Items.Item("sono").Specific.value) + "' and A.u_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' And A.u_sfgcode = '" + Trim(objForm.Items.Item("sfgcode").Specific.value) + "' And A.Filler = '" + Trim(objForm.Items.Item("wipwhs").Specific.value) + "' And B.WhsCode = '" + StWhs + "' And B.ItemCode = '" + Trim(oRecordSet.Fields.Item("Code").Value) + "'")
                                    Dim RetdQty As Double
                                    RetdQty = oRS.Fields.Item("RetdQty").Value
                                    oDBs_Detail.SetValue("u_totis", i - 1, IssQty - RetdQty)
                                    If Trim(objForm.Items.Item("itemcode").Specific.value) = Trim(oDT.GetValue("ItemCode", 0)) Then
                                        If Trim(oDBs_Head.GetValue("u_type", 0)) = "Regular" Then
                                            oDBs_Detail.SetValue("u_rqstqty", i - 1, ((CDbl(objForm.Items.Item("excsqty").Specific.value) / oRecordSet.Fields.Item("Qauntity").Value) * oRecordSet.Fields.Item("Quantity").Value) - (IssQty - RetdQty))
                                            'Vijeesh
                                            oDBs_Detail.SetValue("u_reqdqty", i - 1, ((CDbl(objForm.Items.Item("excsqty").Specific.value) / oRecordSet.Fields.Item("Qauntity").Value) * oRecordSet.Fields.Item("Quantity").Value) - (IssQty - RetdQty))
                                        End If
                                        If Trim(oDBs_Head.GetValue("u_type", 0)) = "Excess" Then
                                            oDBs_Detail.SetValue("u_rqstqty", i - 1, ((CDbl(objForm.Items.Item("excsqty").Specific.value) / oRecordSet.Fields.Item("Qauntity").Value) * oRecordSet.Fields.Item("Quantity").Value))
                                            'Vijeesh
                                            oDBs_Detail.SetValue("u_reqdqty", i - 1, ((CDbl(objForm.Items.Item("excsqty").Specific.value) / oRecordSet.Fields.Item("Qauntity").Value) * oRecordSet.Fields.Item("Quantity").Value))
                                        End If
                                    End If
                                    If Trim(objForm.Items.Item("itemcode").Specific.value) <> Trim(oDT.GetValue("ItemCode", 0)) Then
                                        Dim sfgordrqty As Double
                                        If Trim(oDBs_Head.GetValue("u_type", 0)) = "Regular" Then
                                            RSet.DoQuery("Exec Child_BOM '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "','" + Trim(oDT.GetValue("ItemCode", 0)) + "'")
                                            sfgordrqty = (CDbl(objForm.Items.Item("excsqty").Specific.value) / RSet.Fields.Item("Qauntity").Value) * RSet.Fields.Item("Quantity").Value
                                            oDBs_Detail.SetValue("u_rqstqty", i - 1, ((sfgordrqty / oRecordSet.Fields.Item("Qauntity").Value) * oRecordSet.Fields.Item("Quantity").Value))
                                            'Vijeesh
                                            oDBs_Detail.SetValue("u_reqdqty", i - 1, ((sfgordrqty / oRecordSet.Fields.Item("Qauntity").Value) * oRecordSet.Fields.Item("Quantity").Value))
                                        End If
                                        If Trim(oDBs_Head.GetValue("u_type", 0)) = "Excess" Then
                                            RSet.DoQuery("Exec Child_BOM '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "','" + Trim(oDT.GetValue("ItemCode", 0)) + "'")
                                            sfgordrqty = (CDbl(objForm.Items.Item("excsqty").Specific.value) / RSet.Fields.Item("Qauntity").Value) * RSet.Fields.Item("Quantity").Value
                                            oDBs_Detail.SetValue("u_rqstqty", i - 1, ((sfgordrqty / oRecordSet.Fields.Item("Qauntity").Value) * oRecordSet.Fields.Item("Quantity").Value))
                                            'Vijeesh
                                            oDBs_Detail.SetValue("u_reqdqty", i - 1, ((sfgordrqty / oRecordSet.Fields.Item("Qauntity").Value) * oRecordSet.Fields.Item("Quantity").Value))
                                        End If
                                    End If
                                    objMatrix.SetLineData(i)
                                    oRecordSet.MoveNext()
                                Next
                            End If
                            If oCFL.UniqueID = "WHCFL" Then
                                objMatrix = objForm.Items.Item("mtx").Specific
                                Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("Select OnHand From OITW Where WhsCode = '" + Trim(oDT.GetValue("WhsCode", 0)) + "' And ItemCode = '" + Trim(objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.Value) + "'")
                                oDBs_Detail.Offset = pVal.Row - 1
                                oDBs_Detail.SetValue("u_whs", oDBs_Detail.Offset, oDT.GetValue("WhsCode", 0))
                                oDBs_Detail.SetValue("u_totavlbl", oDBs_Detail.Offset, oRecordSet.Fields.Item("OnHand").Value)
                                objMatrix.SetLineData(pVal.Row)
                            End If
                            If oCFL.UniqueID = "ITRCFL" Then
                                objMatrix = objForm.Items.Item("mtx").Specific
                                oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
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
                                    oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                                    oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, oDT.GetValue("ItemCode", i))
                                    oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, oDT.GetValue("ItemName", i))
                                    oDBs_Detail.SetValue("u_rqstqty", oDBs_Detail.Offset, 1)
                                    oRS.DoQuery("Select IsNUll(u_tol,0) As 'Tol',InvntryUom,DfltWh From OITM Where ItemCode = '" + Trim(oDT.GetValue("ItemCode", i)) + "'")
                                    oDBs_Detail.SetValue("u_uom", oDBs_Detail.Offset, oRS.Fields.Item("InvntryUOM").Value)
                                    oDBs_Detail.SetValue("u_tol", oDBs_Detail.Offset, oRS.Fields.Item("Tol").Value)
                                    oDBs_Detail.SetValue("u_whs", oDBs_Detail.Offset, oRS.Fields.Item("DfltWh").Value)
                                    Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRSet.DoQuery("Select OnHand From OITW Where ItemCode = '" + Trim(oDT.GetValue("ItemCode", i)) + "' And WhsCode = '" + Trim(oRS.Fields.Item("DfltWh").Value) + "'")
                                    oDBs_Detail.SetValue("u_totavlbl", oDBs_Detail.Offset, oRSet.Fields.Item("OnHand").Value)
                                    oRSet.DoQuery("Select OnHand From OITW Where ItemCode = '" + Trim(oDT.GetValue("ItemCode", i)) + "' And WhsCode = '" + Trim(objForm.Items.Item("wipwhs").Specific.value) + "'")
                                    oDBs_Detail.SetValue("u_wipavlbl", oDBs_Detail.Offset, oRSet.Fields.Item("OnHand").Value)
                                    objMatrix.SetLineData(pVal.Row + i)
                                    objForm.EnableMenu("1293", True)
                                Next
                                If Flag = True Then
                                    objMatrix.AddRow(1, objMatrix.VisualRowCount)
                                    Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
                                End If
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
            If pVal.BeforeAction = False Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "GEN_MREQ"
                        Me.CreateForm(objForm.UniqueID)
                    Case "1282"
                        If objForm.TypeEx = "GEN_MREQ" Then
                            Me.SetDefault(objForm.UniqueID)
                            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        End If
                    Case "1281"
                        If objForm.TypeEx = "GEN_MREQ" Then
                            objForm.EnableMenu("1282", True)
                            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        End If
                    Case "1288", "1289", "1290", "1291"
                        If objForm.TypeEx = "GEN_MREQ" Then
                            objForm.EnableMenu("1282", True)
                        End If
                    Case "1293"
                        If objForm.TypeEx = "GEN_MREQ" Then
                            If ITEM_ID.Equals("mtx") = True Then
                                objMatrix = objForm.Items.Item("mtx").Specific
                                oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MREQ")
                                oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_MREQ_D0")
                                For Row As Integer = 1 To objMatrix.VisualRowCount
                                    objMatrix.GetLineData(Row)
                                    oDBs_Detail.Offset = Row - 1
                                    oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, Row)
                                    Dim ochk As SAPbouiCOM.CheckBox
                                    ochk = objMatrix.Columns.Item("chk").Cells.Item(Row).Specific
                                    If ochk.Checked = True Then
                                        oDBs_Detail.SetValue("u_chk", oDBs_Detail.Offset, "Y")
                                    Else
                                        oDBs_Detail.SetValue("u_chk", oDBs_Detail.Offset, "N")
                                    End If
                                    oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, objMatrix.Columns.Item("itemcode").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, objMatrix.Columns.Item("itemname").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_rqstqty", oDBs_Detail.Offset, objMatrix.Columns.Item("rqstqty").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_tol", oDBs_Detail.Offset, objMatrix.Columns.Item("tol").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_uom", oDBs_Detail.Offset, objMatrix.Columns.Item("uom").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_reqdqty", oDBs_Detail.Offset, objMatrix.Columns.Item("reqdqty").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_totavlbl", oDBs_Detail.Offset, objMatrix.Columns.Item("totavlbl").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_totis", oDBs_Detail.Offset, objMatrix.Columns.Item("totis").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_whs", oDBs_Detail.Offset, objMatrix.Columns.Item("whs").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_remarks", oDBs_Detail.Offset, objMatrix.Columns.Item("remarks").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_stat", oDBs_Detail.Offset, objMatrix.Columns.Item("stat").Cells.Item(Row).Specific.value)
                                    objMatrix.SetLineData(Row)
                                Next
                                objMatrix.FlushToDataSource()
                                oDBs_Detail.RemoveRecord(oDBs_Detail.Size - 1)
                                objMatrix.LoadFromDataSource()
                            End If
                        End If
                End Select
            ElseIf pVal.BeforeAction = True Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "519"
                        Try
                            If objForm.TypeEx = "GEN_MREQ" Then
                                BubbleEvent = False
                                sDocNum = objForm.Items.Item("docnum").Specific.Value
                                sRptName = "GEN_SEPL_MRN.rpt"
                                Me.Report1()
                            End If
                        Catch ex As Exception

                        End Try
                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            RowCount = eventInfo.Row
            Dim oForm As SAPbouiCOM.Form = oApplication.Forms.Item(eventInfo.FormUID)
            If eventInfo.Row > 0 Then
                ITEM_ID = eventInfo.ItemUID
                objMatrix = oForm.Items.Item("mtx").Specific
                If objMatrix.VisualRowCount > 1 Then
                    oForm.EnableMenu("1293", True)
                Else
                    oForm.EnableMenu("1293", False)
                End If
            Else
                ITEM_ID = ""
            End If
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                oForm.EnableMenu("1283", False)
                'BubbleEvent = False
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
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MREQ")
                        If Trim(oDBs_Head.GetValue("U_status", 0)) = "Open" Then
                            objForm.Items.Item("btnis").Enabled = True
                            objForm.Items.Item("btnret").Enabled = True
                            objForm.Items.Item("whs").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Items.Item("wipwhs").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            If Trim(oDBs_Head.GetValue("u_type", 0)) = "Excess" Or Trim(oDBs_Head.GetValue("u_type", 0)) = "Production Consumable Excess" Then
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet.DoQuery("Select USER_CODE From [OUSR] Where USER_CODE = '" + oCompany.UserName.ToString.Trim + "' And IsNull(u_approve,'N') = 'Y'")
                                If oRSet.RecordCount > 0 Then
                                    objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                Else
                                    objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                End If
                            Else
                                objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            End If
                            'If Trim(oDBs_Head.GetValue("u_type", 0)) <> "Excess" Or Trim(oDBs_Head.GetValue("u_type", 0)) <> "Production Consumable Excess" Then
                            '    objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            'End If
                        End If
                        If Trim(oDBs_Head.GetValue("U_status", 0)) = "Closed" Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                            objForm.Items.Item("btnis").Enabled = False
                            objForm.Items.Item("btnret").Enabled = True
                        End If
                        'If Trim(oDBs_Head.GetValue("U_type", 0)) = "Consumable" Then
                        '    objMatrix = objForm.Items.Item("mtx").Specific
                        '    objMatrix.AddRow()
                        '    objMatrix.FlushToDataSource()
                        '    Me.SetNewLine(BusinessObjectInfo.FormUID, objMatrix.VisualRowCount, objMatrix)
                        'End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        If UpdMode = True And DocStatus = "Closed" Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            oApplication.ActivateMenuItem("1281")
                        End If
                    End If

                    'If BusinessObjectInfo.BeforeAction = True Then
                    '    objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    '    oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MREQ")
                    '    objMatrix = objForm.Items.Item("mtx").Specific
                    '    If Trim(oDBs_Head.GetValue("u_type", 0)) = "Consumable" Then
                    '        objMatrix.DeleteRow(objMatrix.VisualRowCount)
                    '        objMatrix.FlushToDataSource()
                    '    End If
                    'ElseIf BusinessObjectInfo.ActionSuccess = True Then
                    '    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Then
                    '        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MREQ")
                    '        objMatrix = objForm.Items.Item("mtx").Specific
                    '        If Trim(oDBs_Head.GetValue("u_type", 0)) = "Consumable" Then
                    '            objMatrix.AddRow()
                    '            objMatrix.FlushToDataSource()
                    '            Me.SetNewLine(BusinessObjectInfo.FormUID, objMatrix.VisualRowCount, objMatrix)
                    '        End If
                    '    End If
                    'End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub PostMessage(ByVal FormUID As String)
        Try
            Dim I As Int16
            Dim Note As String = ""
            Dim oRecordSet1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRecordSet2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oApplication.StatusBar.SetText("Sending alert to users Please wait........", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Dim oMatrix As SAPbouiCOM.Matrix
            objForm = oApplication.Forms.Item(FormUID)
            oMatrix = objForm.Items.Item("mtx").Specific
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim oMessageService As SAPbobsCOM.MessagesService
            Dim oMessage As SAPbobsCOM.Message
            Dim pMessageDataColumns As SAPbobsCOM.MessageDataColumns
            Dim pMessageDataColumn As SAPbobsCOM.MessageDataColumn
            Dim oLines As SAPbobsCOM.MessageDataLines
            Dim oLine As SAPbobsCOM.MessageDataLine
            Dim oRecipientCollection As SAPbobsCOM.RecipientCollection
            oCmpSrv = oCompany.GetCompanyService
            oMessageService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService)
            oMessage = oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage)
            oRecordSet.DoQuery("Select distinct U_User From [@GEN_WHS_USR] Where u_whs = '" + AlertWhs + "'")
            oRecordSet.MoveFirst()
            If oRecordSet.RecordCount > 0 Then
                oMessage.Subject = "Material requisition note generated"
                oRecordSet1.MoveFirst()
                oMessage.Text = "MRN Document Number:" + AlertDocNum
                oMessage.Priority = SAPbobsCOM.BoMsgPriorities.pr_High
                For I = 0 To oRecordSet.RecordCount - 1
                    oRecipientCollection = oMessage.RecipientCollection
                    oRecipientCollection.Add()
                    oRecipientCollection.Item(I).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
                    oRecipientCollection.Item(I).UserCode = oRecordSet.Fields.Item("U_User").Value
                    oRecordSet.MoveNext()
                Next
                oApplication.StatusBar.SetText("Sending alerts to users Please wait........", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                pMessageDataColumns = oMessage.MessageDataColumns
                pMessageDataColumn = pMessageDataColumns.Add()
                pMessageDataColumn.ColumnName = "MRN Number"
                oLines = pMessageDataColumn.MessageDataLines()
                oLine = oLines.Add()
                oLine.Value = AlertDocNum + "( Click here to open MRN)"
                oMessageService.SendMessage(oMessage)
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet1)
            oApplication.StatusBar.SetText("Notification has been sent successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oApplication.MessageBox(ex.Message)
        End Try
    End Sub

    Private Sub Report1()
        Dim oThread As New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Report1Thread))
        oThread.SetApartmentState(System.Threading.ApartmentState.STA)
        oThread.Start()
    End Sub

    Private Sub Report1Thread()
        Try
            Dim oCRForm As New Crystal_Form
            oCRForm.ShowDialog()
        Catch ex As Exception
            oApplication.MessageBox(ex.Message.ToString)
        End Try
    End Sub

End Class
