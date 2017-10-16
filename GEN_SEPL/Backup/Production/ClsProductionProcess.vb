Public Class ClsProductionProcess

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
            oUtilities.SAPXML("GEN_PROD_PRCS.xml")
            objForm = oApplication.Forms.GetForm("GEN_PROD_PRCS", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_PROD_PRCS")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_PROD_PRCS_D0")
            objForm.DataBrowser.BrowseBy = "code"
            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            objForm.Items.Item("itemcode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_PROD_PRCS_D0")
            objMatrix = oMatrix
            objMatrix.FlushToDataSource()
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, objMatrix.VisualRowCount)
            oDBs_Detail.SetValue("u_process", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_sfgcode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_sfgname", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_sfgqty", oDBs_Detail.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try
            objMatrix = objForm.Items.Item("mtx").Specific
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_PROD_PRCS")
            If Trim(oDBs_Head.GetValue("u_itemcode", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter style code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            For I As Integer = 1 To objMatrix.VisualRowCount
                If Trim(objMatrix.Columns.Item("process").Cells.Item(I).Specific.value) <> "" Then
                    If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(I).Specific.value) = "" Then
                        oApplication.StatusBar.SetText("Please enter item code, sub item code and quantity in rows", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            Next
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSet.DoQuery("Select Code From [@GEN_PROD_PRCS] Where u_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' And Code <> '" + Trim(objForm.Items.Item("code").Specific.value) + "'")
            If oRSet.RecordCount > 0 Then
                oApplication.StatusBar.SetText("Production Process already defined for this item", , SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If Trim(objForm.Items.Item("stwhs").Specific.value) = "" Or Trim(objForm.Items.Item("inwhs").Specific.value) = "" Or Trim(objForm.Items.Item("outwhs").Specific.value) = "" Then
                oApplication.StatusBar.SetText("Please enter in,out and stored warehouses", , SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Sub FilterBOMDoc(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("CSTBOM")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            oCon = oCons.Add()
            oCon.Alias = "U_itemcode"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = Trim(objForm.Items.Item("itemcode").Specific.value)
            oCFL.SetConditions(oCons)
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
                        If Me.Validation(FormUID) = False Then BubbleEvent = False
                    ElseIf pVal.ItemUID = "1" And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRSet.DoQuery("Select Max(Convert(Integer,Code)) + 1 AS 'Count' From [@GEN_PROD_PRCS]")
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_PROD_PRCS")
                        oDBs_Head.SetValue("Code", 0, oRSet.Fields.Item("Count").Value)
                    End If
                    If pVal.ItemUID = "upld" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If pVal.BeforeAction = True Then
                            If Trim(objForm.Items.Item("cstbom").Specific.value) = "" Then
                                oApplication.StatusBar.SetText("Please select BOM No", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select Code From [OITT] Where Code = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "'")
                            If oRSet.RecordCount > 0 Then
                                oApplication.StatusBar.SetText("BOM already created for this style, Please click on update button to modify BOMs", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            oRSet.DoQuery("Select DocEntry From [@GEN_CUST_BOM] Where u_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' And u_status = 'NEW'")
                            If oRSet.RecordCount = 0 Then
                                oApplication.StatusBar.SetText("Please create BOM for this style", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        Else
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            Try
                                Dim RS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim ToWhs As String
                                Dim DocEntry As String
                                Dim DocNum As String
                                Dim ItemCode As String
                                Dim BOMFlag As Boolean = False
                                oRecordSet.DoQuery("Delete From TMP_CST_BOM Where MacId = '" + MAC_ID + "'")
                                Dim oPRODTREE As SAPbobsCOM.ProductTrees
                                Dim oPRODTREE_LINES As SAPbobsCOM.ProductTrees_Lines
                                RS1.DoQuery("Select A.DocEntry,A.u_sono,A.DocNum,A.u_itemcode,A.u_qty,A.u_unit,B.u_itemcode AS 'LITEMCODE',B.u_qty AS 'LQTY',B.u_process,B.u_issmthd,B.u_status,B.u_deleted From [@GEN_CUST_BOM] A INNER JOIN [@GEN_CUST_BOM_D0] B ON A.DocEntry = B.DocEntry and A.u_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' And A.DocNum = '" + Trim(objForm.Items.Item("cstbom").Specific.value) + "' Order By B.LineID")
                                DocEntry = RS1.Fields.Item("DocEntry").Value
                                DocNum = RS1.Fields.Item("DocNum").Value
                                For i As Integer = 1 To RS1.RecordCount
                                    oRSet.DoQuery("Insert Into TMP_CST_BOM(SONO,DOCNUM,ItemCode,SOQty,Unit,ChldItem,Qty,Process,IssMthd,Status,MacId,Deleted) Values('" + Trim(RS1.Fields.Item("u_sono").Value) + "','" + Trim(RS1.Fields.Item("DocNum").Value) + "','" + Trim(RS1.Fields.Item("u_itemcode").Value) + "','" + Trim(RS1.Fields.Item("u_qty").Value) + "','" + Trim(RS1.Fields.Item("u_unit").Value) + "','" + Trim(RS1.Fields.Item("LITEMCODE").Value) + "','" + Trim(RS1.Fields.Item("LQTY").Value) + "','" + Trim(RS1.Fields.Item("u_process").Value) + "','" + Trim(RS1.Fields.Item("u_issmthd").Value) + "','" + Trim(RS1.Fields.Item("u_status").Value) + "','" + MAC_ID + "','" + Trim(RS1.Fields.Item("u_deleted").Value) + "')")
                                    RS1.MoveNext()
                                Next
                                Dim Counter As Integer
                                ItemCode = objForm.Items.Item("itemcode").Specific.value
                                oRecordSet.DoQuery("Select Distinct B.u_itemcode,B.u_process,A.Code From [@GEN_PROD_PRCS] A INNER JOIN [@GEN_PROD_PRCS_D0] B ON A.Code = B.Code Where A.u_itemcode = '" + ItemCode + "'")
                                For k As Integer = 1 To oRecordSet.RecordCount
                                    RS1.DoQuery("Select Code FROM OITT Where Code = '" + Trim(oRecordSet.Fields.Item("u_itemcode").Value) + "'")
                                    If RS1.RecordCount = 0 Then
                                        oPRODTREE = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees)
                                        oRSet.DoQuery("Select Distinct SONO,DOCNUm,ItemCode,SOQty,Unit,ChldItem,Qty,Process,IssMthd,Status,MacId From TMP_CST_BOM Where ItemCode = '" + ItemCode + "' And Process = '" + Trim(oRecordSet.Fields.Item("u_process").Value) + "'")
                                        oPRODTREE.TreeCode = oRecordSet.Fields.Item("u_itemcode").Value
                                        oPRODTREE.TreeType = SAPbobsCOM.BoItemTreeTypes.iProductionTree
                                        oPRODTREE.Quantity = 1
                                        oPRODTREE_LINES = oPRODTREE.Items
                                        Dim ss As String = "Select B.u_inwhs From [@GEN_UNIT_MST] A INNER JOIN [@GEN_UNIT_MST_D0] B ON A.CODE = B.CODE Where A.Name = '" + Trim(oRSet.Fields.Item("Unit").Value) + "' And B.u_process = '" + Trim(oRSet.Fields.Item("Process").Value) + "'"

                                        RS1.DoQuery("Select B.u_inwhs From [@GEN_UNIT_MST] A INNER JOIN [@GEN_UNIT_MST_D0] B ON A.CODE = B.CODE Where A.Name = '" + Trim(oRSet.Fields.Item("Unit").Value) + "' And B.u_process = '" + Trim(oRSet.Fields.Item("Process").Value) + "'")
                                        ToWhs = RS1.Fields.Item("u_inwhs").Value

                                        For J As Integer = 1 To oRSet.RecordCount
                                            oPRODTREE_LINES.ParentItem = oRecordSet.Fields.Item("u_itemcode").Value
                                            oPRODTREE_LINES.ItemCode = oRSet.Fields.Item("ChldItem").Value
                                            oPRODTREE_LINES.Quantity = oRSet.Fields.Item("Qty").Value
                                            oPRODTREE_LINES.Warehouse = ToWhs
                                            If Trim(oRSet.Fields.Item("IssMthd").Value) = "B" Then
                                                oPRODTREE_LINES.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                oPRODTREE_LINES.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            oRSet.MoveNext()
                                            oPRODTREE_LINES.SetCurrentLine(J - 1)
                                            oPRODTREE_LINES.Add()
                                            Counter = J
                                        Next
                                        Dim str As String = "Select Distinct B.u_sfgcode,B.u_sfgqty From [@GEN_PROD_PRCS] A INNER JOIN [@GEN_PROD_PRCS_D0] B ON A.Code = B.Code Where A.u_itemcode = '" + ItemCode + "' And B.u_itemcode = '" + Trim(oRecordSet.Fields.Item("u_itemcode").Value) + "' And (B.u_sfgcode <> '' Or B.u_sfgcode is not null) And A.u_itemcode = '" + ItemCode + "'"


                                        RS.DoQuery("Select Distinct B.u_sfgcode,B.u_sfgqty From [@GEN_PROD_PRCS] A INNER JOIN [@GEN_PROD_PRCS_D0] B ON A.Code = B.Code Where A.u_itemcode = '" + ItemCode + "' And B.u_itemcode = '" + Trim(oRecordSet.Fields.Item("u_itemcode").Value) + "' And (B.u_sfgcode <> '' Or B.u_sfgcode is not null) And A.u_itemcode = '" + ItemCode + "'")
                                        For L As Integer = 1 To RS.RecordCount
                                            Dim SFGCODE As String = Trim(RS.Fields.Item("u_sfgcode").Value)
                                            If SFGCODE <> "" Then
                                                oPRODTREE_LINES.ParentItem = oRecordSet.Fields.Item("u_itemcode").Value
                                                oPRODTREE_LINES.ItemCode = RS.Fields.Item("u_sfgcode").Value
                                                If CDbl(RS.Fields.Item("u_sfgqty").Value) < 0 Then
                                                    oPRODTREE_LINES.Quantity = RS.Fields.Item("u_sfgqty").Value
                                                    oPRODTREE_LINES.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Backflush
                                                End If
                                                If RS.Fields.Item("u_sfgqty").Value > 0 Then
                                                    oPRODTREE_LINES.Quantity = RS.Fields.Item("u_sfgqty").Value
                                                    oPRODTREE_LINES.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Manual
                                                End If
                                                oPRODTREE_LINES.Warehouse = ToWhs
                                                oPRODTREE_LINES.SetCurrentLine(Counter)
                                                Counter = Counter + 1
                                                oPRODTREE_LINES.Add()
                                            End If
                                            RS.MoveNext()
                                        Next
                                        oCompany.StartTransaction()
                                        Dim Err As Integer = oPRODTREE.Add()
                                        If Err <> 0 Then
                                            oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                            BOMFlag = False
                                            BubbleEvent = False
                                            Exit Sub
                                        Else
                                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                            BOMFlag = True
                                        End If
                                        oRecordSet.MoveNext()
                                    End If
                                Next
                                'oRecordSet.DoQuery("Select Distinct B.u_itemcode,B.u_process From [@GEN_PROD_PRCS] A INNER JOIN [@GEN_PROD_PRCS_D0] B ON A.Code = B.Code Where A.u_itemcode = '" + ItemCode + "' And B.u_process = 'Finishing New'")
                                oRecordSet.DoQuery("Select Distinct B.u_itemcode,B.u_process From [@GEN_PROD_PRCS] A INNER JOIN [@GEN_PROD_PRCS_D0] B ON A.Code = B.Code Where A.u_itemcode = '" + ItemCode + "' And B.u_process = 'FINISHING STD AND MTRL COST'")
                                oRSet.DoQuery("Select Top 1 Unit AS 'Unit' From TMP_CST_BOM Where DOCNUM = '" + DocNum + "'")
                                RS1.DoQuery("Select B.u_inwhs From [@GEN_UNIT_MST] A INNER JOIN [@GEN_UNIT_MST_D0] B ON A.CODE = B.CODE Where A.Name = '" + Trim(oRSet.Fields.Item("Unit").Value) + "' And B.u_process = '" + Trim(oRecordSet.Fields.Item("u_process").Value) + "'")
                                RS.DoQuery("Select Code From OITT WHere Code = '" + ItemCode + "'")
                                If RS.RecordCount = 0 Then
                                    ToWhs = RS1.Fields.Item("u_inwhs").Value
                                    oPRODTREE = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees)
                                    oPRODTREE.TreeCode = ItemCode
                                    oPRODTREE.TreeType = SAPbobsCOM.BoItemTreeTypes.iProductionTree
                                    oPRODTREE.Quantity = 1
                                    oPRODTREE_LINES = oPRODTREE.Items
                                    For N As Integer = 1 To oRecordSet.RecordCount
                                        oPRODTREE_LINES.ParentItem = ItemCode
                                        oPRODTREE_LINES.ItemCode = oRecordSet.Fields.Item("u_itemcode").Value
                                        oPRODTREE_LINES.Quantity = 1
                                        oPRODTREE_LINES.Warehouse = ToWhs
                                        oPRODTREE_LINES.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Manual
                                        oRecordSet.MoveNext()
                                        oPRODTREE_LINES.SetCurrentLine(N - 1)
                                        oPRODTREE_LINES.Add()
                                    Next
                                    oCompany.StartTransaction()
                                    Dim ErrFlg As Integer = oPRODTREE.Add()
                                    If ErrFlg <> 0 Then
                                        oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BOMFlag = False
                                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        BOMFlag = True
                                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                    End If
                                End If
                                oRecordSet.DoQuery("Delete From TMP_CST_BOM Where MacId = '" + MAC_ID + "'")
                                oRecordSet.DoQuery("Update [@GEN_CUST_BOM_D0] Set u_status = 'ACTIVE' Where DocEntry = '" + DocEntry + "' And u_status = 'NEW'")
                                oRecordSet.DoQuery("Update [@GEN_CUST_BOM] Set u_status = 'ACTIVE' Where DocEntry = '" + DocEntry + "'")
                                If BOMFlag = True Then
                                    oApplication.StatusBar.SetText("BOMs Created Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                End If
                            Catch ex As Exception
                                oRecordSet.DoQuery("Delete From TMP_CST_BOM Where MacId = '" + MAC_ID + "'")
                                oApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                    End If
                    If pVal.ItemUID = "updt" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If pVal.BeforeAction = True Then
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            If Trim(objForm.Items.Item("cstbom").Specific.value) = "" Then
                                oApplication.StatusBar.SetText("Please select BOM No", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            oRSet.DoQuery("Select Code From [OITT] Where Code = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "'")
                            If oRSet.RecordCount = 0 Then
                                oApplication.StatusBar.SetText("BOM does not exist, Please upload the BOM before you can update", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        Else
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            Try
                                Dim RS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim ToWhs As String
                                Dim DocEntry As String
                                Dim DocNum As String
                                Dim ItemCode As String
                                Dim Status As String
                                Dim ErrUpd As Integer
                                Dim ErrAdd As Integer
                                Dim ErrDel As Integer
                                Dim BOMFlag As Boolean = False
                                oRecordSet.DoQuery("Delete From TMP_CST_BOM Where MacId = '" + MAC_ID + "'")
                                Dim oPRODTREE As SAPbobsCOM.ProductTrees
                                Dim oPRODTREE_LINES As SAPbobsCOM.ProductTrees_Lines
                                RS1.DoQuery("Select A.DocEntry,A.u_sono,A.DocNum,A.u_itemcode,A.u_status as 'HSTAT',A.u_qty,A.u_unit,B.u_itemcode AS 'LITEMCODE',B.u_qty AS 'LQTY',B.u_process,B.u_issmthd,B.u_status,B.u_deleted From [@GEN_CUST_BOM] A INNER JOIN [@GEN_CUST_BOM_D0] B ON A.DocEntry = B.DocEntry and A.u_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' ANd A.DocNum = '" + Trim(objForm.Items.Item("cstbom").Specific.value) + "'")
                                DocEntry = RS1.Fields.Item("DocEntry").Value
                                Status = RS1.Fields.Item("HSTAT").Value
                                DocNum = RS1.Fields.Item("DocNum").Value
                                For i As Integer = 1 To RS1.RecordCount
                                    oRSet.DoQuery("Insert Into TMP_CST_BOM(SONO,DOCNUM,ItemCode,SOQty,Unit,ChldItem,Qty,Process,IssMthd,Status,MacId,Deleted) Values('" + Trim(RS1.Fields.Item("u_sono").Value) + "','" + Trim(RS1.Fields.Item("DocNum").Value) + "','" + Trim(RS1.Fields.Item("u_itemcode").Value) + "','" + Trim(RS1.Fields.Item("u_qty").Value) + "','" + Trim(RS1.Fields.Item("u_unit").Value) + "','" + Trim(RS1.Fields.Item("LITEMCODE").Value) + "','" + Trim(RS1.Fields.Item("LQTY").Value) + "','" + Trim(RS1.Fields.Item("u_process").Value) + "','" + Trim(RS1.Fields.Item("u_issmthd").Value) + "','" + Trim(RS1.Fields.Item("u_status").Value) + "','" + MAC_ID + "','" + Trim(RS1.Fields.Item("u_deleted").Value) + "')")
                                    RS1.MoveNext()
                                Next
                                ItemCode = objForm.Items.Item("itemcode").Specific.value
                                oPRODTREE = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees)
                                If Status = "CHANGE" Then
                                    oRSet.DoQuery("Select Distinct SONO,DOCNUM,ItemCode,SOQty,Unit,ChldItem,Qty,Process,IssMthd,Status,MacId,Deleted From TMP_CST_BOM Where ItemCode = '" + ItemCode + "' Order by Status")
                                    ''oRSet.DoQuery(" Select A.DocEntry,A.u_sono,A.DocNum,A.u_itemcode,A.u_status as 'HSTAT' " & _
                                    ''" ,A.u_qty,A.u_unit,'' as 'LITEMCODE',0 AS 'LQTY','' as u_process,'' as u_issmthd,'' as u_status,'' as u_deleted " & _
                                    ''" From [@GEN_CUST_BOM] A Where A.u_itemcode = '" + ItemCode + "'")

                                    For i As Integer = 1 To oRSet.RecordCount
                                        If Trim(oRSet.Fields.Item("Status").Value) = "CHANGE" Then
                                            oRecordSet.DoQuery("Select Distinct B.u_itemcode From [@GEN_PROD_PRCS] A INNER JOIN [@GEN_PROD_PRCS_D0] B ON A.CODE = B.CODE Where A.u_itemcode = '" + ItemCode + "' And B.u_process = '" + Trim(oRSet.Fields.Item("Process").Value) + "'")
                                            oPRODTREE.GetByKey(Trim(oRecordSet.Fields.Item("u_itemcode").Value))
                                            oPRODTREE_LINES = oPRODTREE.Items
                                            RS.DoQuery("Select Distinct ChildNum From ITT1 Where Father = '" + Trim(oRecordSet.Fields.Item("u_itemcode").Value) + "' And Code = '" + Trim(oRSet.Fields.Item("ChldItem").Value) + "'")
                                            oPRODTREE_LINES.SetCurrentLine(Trim(RS.Fields.Item("ChildNum").Value))
                                            oPRODTREE_LINES.Quantity = oRSet.Fields.Item("Qty").Value
                                            If Trim(oRSet.Fields.Item("IssMthd").Value) = "B" Then
                                                oPRODTREE_LINES.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                oPRODTREE_LINES.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            oCompany.StartTransaction()
                                            ErrUpd = oPRODTREE.Update
                                            If ErrUpd <> 0 Then
                                                oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                BubbleEvent = False
                                                Exit Try
                                            Else
                                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                            End If
                                        End If
                                        If Trim(oRSet.Fields.Item("Status").Value) = "NEW" Then
                                            oRecordSet.DoQuery("Select Distinct B.u_itemcode,B.u_process From [@GEN_PROD_PRCS] A INNER JOIN [@GEN_PROD_PRCS_D0] B ON A.CODE = B.CODE Where A.u_itemcode = '" + ItemCode + "' And B.u_process = '" + Trim(oRSet.Fields.Item("Process").Value) + "'")
                                            RS1.DoQuery("Select B.u_inwhs From [@GEN_UNIT_MST] A INNER JOIN [@GEN_UNIT_MST_D0] B ON A.CODE = B.CODE Where B.u_process = '" + Trim(oRecordSet.Fields.Item("u_process").Value) + "'")
                                            ''RS1.DoQuery("Select B.u_inwhs From [@GEN_UNIT_MST] A INNER JOIN [@GEN_UNIT_MST_D0] B ON A.CODE = B.CODE Where A.Name = '" + Trim(objForm.Items.Item("unit").Specific.value) + "' And B.u_process = '" + Trim(oRecordSet.Fields.Item("u_process").Value) + "'")
                                            ToWhs = RS1.Fields.Item("u_inwhs").Value
                                            oPRODTREE.GetByKey(Trim(oRecordSet.Fields.Item("u_itemcode").Value))
                                            oPRODTREE_LINES = oPRODTREE.Items
                                            RS.DoQuery("Select Distinct Max(ChildNum) As 'ChildNum' From ITT1 Where Father = '" + Trim(oRecordSet.Fields.Item("u_itemcode").Value) + "' And Code = '" + Trim(oRSet.Fields.Item("ChldItem").Value) + "'")
                                            oPRODTREE_LINES.SetCurrentLine(oPRODTREE.Items.Count - 1)
                                            oPRODTREE_LINES.Add()
                                            oPRODTREE_LINES.ParentItem = oRecordSet.Fields.Item("u_itemcode").Value
                                            oPRODTREE_LINES.ItemCode = oRSet.Fields.Item("ChldItem").Value
                                            oPRODTREE_LINES.Quantity = oRSet.Fields.Item("Qty").Value
                                            oPRODTREE_LINES.Warehouse = ToWhs
                                            If Trim(oRSet.Fields.Item("IssMthd").Value) = "B" Then
                                                oPRODTREE_LINES.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                oPRODTREE_LINES.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            oCompany.StartTransaction()
                                            ErrAdd = oPRODTREE.Update
                                            If ErrAdd <> 0 Then
                                                oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                BubbleEvent = False
                                                Exit Try
                                            Else
                                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                            End If
                                        End If
                                        If Trim(oRSet.Fields.Item("Status").Value) = "DELETE" And Trim(oRSet.Fields.Item("Deleted").Value) = "NO" Then
                                            oRecordSet.DoQuery("Select Distinct B.u_itemcode From [@GEN_PROD_PRCS] A INNER JOIN [@GEN_PROD_PRCS_D0] B ON A.CODE = B.CODE Where A.u_itemcode = '" + ItemCode + "' And B.u_process = '" + Trim(oRSet.Fields.Item("Process").Value) + "'")
                                            oPRODTREE.GetByKey(Trim(oRecordSet.Fields.Item("u_itemcode").Value))
                                            oPRODTREE_LINES = oPRODTREE.Items
                                            RS.DoQuery("Select Distinct ChildNum As 'ChildNum' From ITT1 Where Father = '" + Trim(oRecordSet.Fields.Item("u_itemcode").Value) + "' And Code = '" + Trim(oRSet.Fields.Item("ChldItem").Value) + "'")
                                            oPRODTREE_LINES.SetCurrentLine(Trim(RS.Fields.Item("ChildNum").Value))
                                            oPRODTREE_LINES.Delete()
                                            oCompany.StartTransaction()
                                            ErrDel = oPRODTREE.Update
                                            If ErrDel <> 0 Then
                                                oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                BubbleEvent = False
                                                Exit Try
                                            Else
                                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                            End If
                                        End If
                                        oRSet.MoveNext()
                                    Next
                                    If ErrUpd = 0 And ErrAdd = 0 And ErrDel = 0 Then
                                        oApplication.StatusBar.SetText("BOMs update successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    End If
                                End If
                                oRecordSet.DoQuery("Delete From TMP_CST_BOM Where MacId = '" + MAC_ID + "'")
                                oRecordSet.DoQuery("Update [@GEN_CUST_BOM_D0] Set u_status = 'ACTIVE' Where DocEntry = '" + DocEntry + "' And u_status = 'CHANGE'")
                                oRecordSet.DoQuery("Update [@GEN_CUST_BOM_D0] Set u_status = 'ACTIVE' Where DocEntry = '" + DocEntry + "' And u_status = 'NEW'")
                                oRecordSet.DoQuery("Update [@GEN_CUST_BOM_D0] Set u_status = 'ACTIVE',u_deleted = 'YES' Where DocEntry = '" + DocEntry + "' And u_status = 'DELETE'")
                                oRecordSet.DoQuery("Update [@GEN_CUST_BOM] Set u_status = 'ACTIVE' Where DocEntry = '" + DocEntry + "'")
                            Catch ex As Exception
                                oRecordSet.DoQuery("Delete From TMP_CST_BOM Where MacId = '" + MAC_ID + "'")
                                oApplication.StatusBar.SetText(ex.Message)
                            End Try
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
                        If oCFL.UniqueID = "CSTBOM" Then
                            If Trim(objForm.Items.Item("itemcode").Specific.value) = "" Then
                                oApplication.StatusBar.SetText("Please select style", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            Me.FilterBOMDoc(FormUID)
                        End If
                    Else
                        If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_PROD_PRCS")
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_PROD_PRCS_D0")
                            If oCFL.UniqueID = "ITCFL" Then
                                oDBs_Head.SetValue("u_itemcode", 0, oDT.GetValue("ItemCode", 0))
                                oDBs_Head.SetValue("u_itemname", 0, oDT.GetValue("ItemName", 0))
                                objMatrix = objForm.Items.Item("mtx").Specific
                                objMatrix.Clear()
                                objMatrix.AddRow(1)
                                Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
                            End If
                            If oCFL.UniqueID = "W1CFL" Then
                                oDBs_Head.SetValue("u_stwhs", 0, oDT.GetValue("WhsCode", 0))
                            End If
                            If oCFL.UniqueID = "W2CFL" Then
                                oDBs_Head.SetValue("u_inwhs", 0, oDT.GetValue("WhsCode", 0))
                            End If
                            If oCFL.UniqueID = "CSTBOM" Then
                                oDBs_Head.SetValue("u_cstbom", 0, oDT.GetValue("DocNum", 0))
                                oDBs_Head.SetValue("u_soref", 0, oDT.GetValue("U_soref", 0))
                            End If
                            If oCFL.UniqueID = "W3CFL" Then
                                oDBs_Head.SetValue("u_outwhs", 0, oDT.GetValue("WhsCode", 0))
                            End If
                            If oCFL.UniqueID = "PRCCFL" Then
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
                                    oDBs_Detail.SetValue("u_process", oDBs_Detail.Offset, oDT.GetValue("Name", i))
                                    objMatrix.SetLineData(pVal.Row + i)
                                    objForm.EnableMenu("1293", True)
                                Next
                                If Flag = True Then
                                    objMatrix.AddRow(1, objMatrix.VisualRowCount)
                                    Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
                                End If
                            End If
                            If oCFL.UniqueID = "RITCFL" Then
                                objMatrix = objForm.Items.Item("mtx").Specific
                                oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                For i As Integer = 0 To oDT.Rows.Count - 1
                                    Dim cflSelectedcount As Integer = oDT.Rows.Count
                                    oDBs_Detail.Offset = pVal.Row - 1 + i
                                    oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                                    oDBs_Detail.SetValue("u_process", oDBs_Detail.Offset, objMatrix.Columns.Item("process").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, oDT.GetValue("ItemCode", i))
                                    oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, oDT.GetValue("ItemName", i))
                                    oDBs_Detail.SetValue("u_sfgcode", oDBs_Detail.Offset, objMatrix.Columns.Item("sfgcode").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_sfgname", oDBs_Detail.Offset, objMatrix.Columns.Item("sfgname").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_sfgqty", oDBs_Detail.Offset, objMatrix.Columns.Item("sfgqty").Cells.Item(pVal.Row).Specific.value)
                                    objMatrix.SetLineData(pVal.Row + i)
                                Next
                            End If
                            If oCFL.UniqueID = "RSITCFL" Then
                                objMatrix = objForm.Items.Item("mtx").Specific
                                oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                For i As Integer = 0 To oDT.Rows.Count - 1
                                    Dim cflSelectedcount As Integer = oDT.Rows.Count
                                    oDBs_Detail.Offset = pVal.Row - 1 + i
                                    oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                                    oDBs_Detail.SetValue("u_process", oDBs_Detail.Offset, objMatrix.Columns.Item("process").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, objMatrix.Columns.Item("itemname").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_sfgcode", oDBs_Detail.Offset, oDT.GetValue("ItemCode", i))
                                    oDBs_Detail.SetValue("u_sfgname", oDBs_Detail.Offset, oDT.GetValue("ItemName", i))
                                    oDBs_Detail.SetValue("u_sfgqty", oDBs_Detail.Offset, objMatrix.Columns.Item("sfgqty").Cells.Item(pVal.Row).Specific.value)
                                    objMatrix.SetLineData(pVal.Row + i)
                                Next
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
                    Case "GEN_PROD_PRCS"
                        Me.CreateForm(objForm.UniqueID)
                    Case "1282"
                        If objForm.TypeEx = "GEN_PROD_PRCS" Then
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select Max(Convert(Integer,Code)) + 1 AS 'Count' From [@GEN_PROD_PRCS]")
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_PROD_PRCS")
                            oDBs_Head.SetValue("Code", 0, oRSet.Fields.Item("Count").Value)
                        End If
                    Case "1281"
                        If objForm.TypeEx = "GEN_PROD_PRCS" Then
                            objForm.EnableMenu("1282", True)
                        End If
                    Case "1288", "1289", "1290", "1291"
                        If objForm.TypeEx = "GEN_PROD_PRCS" Then
                            objForm.EnableMenu("1282", True)
                        End If
                    Case "1293"
                        If objForm.TypeEx = "GEN_PROD_PRCS" Then
                            If ITEM_ID.Equals("mtx") = True Then
                                objMatrix = objForm.Items.Item("mtx").Specific
                                oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_PROD_PRCS")
                                oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_PROD_PRCS_D0")
                                For Row As Integer = 1 To objMatrix.VisualRowCount
                                    objMatrix.GetLineData(Row)
                                    oDBs_Detail.Offset = Row - 1
                                    oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, Row)
                                    oDBs_Detail.SetValue("u_process", oDBs_Detail.Offset, objMatrix.Columns.Item("process").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, objMatrix.Columns.Item("itemcode").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, objMatrix.Columns.Item("itemname").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_sfgcode", oDBs_Detail.Offset, objMatrix.Columns.Item("sfgcode").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_sfgname", oDBs_Detail.Offset, objMatrix.Columns.Item("sfgname").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_sfgqty", oDBs_Detail.Offset, objMatrix.Columns.Item("sfgqty").Cells.Item(Row).Specific.value)
                                    objMatrix.SetLineData(Row)
                                Next
                                objMatrix.FlushToDataSource()
                                oDBs_Detail.RemoveRecord(oDBs_Detail.Size - 1)
                                objMatrix.LoadFromDataSource()
                            End If
                        End If
                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            RowCount = eventInfo.Row
            If eventInfo.Row > 0 Then
                ITEM_ID = eventInfo.ItemUID
                Dim oForm As SAPbouiCOM.Form = oApplication.Forms.Item(eventInfo.FormUID)
                objMatrix = oForm.Items.Item("mtx").Specific
                If objMatrix.VisualRowCount > 1 Then
                    oForm.EnableMenu("1293", True)
                Else
                    oForm.EnableMenu("1293", False)
                End If
            Else
                ITEM_ID = ""
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = True Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        objMatrix = objForm.Items.Item("mtx").Specific
                        If objMatrix.VisualRowCount <> 0 Then
                            objMatrix.DeleteRow(objMatrix.VisualRowCount)
                            objMatrix.FlushToDataSource()
                        End If
                    ElseIf BusinessObjectInfo.ActionSuccess = True Then
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Then
                            objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                            objMatrix = objForm.Items.Item("mtx").Specific
                            objMatrix.AddRow()
                            objMatrix.FlushToDataSource()
                            Me.SetNewLine(BusinessObjectInfo.FormUID, objMatrix.VisualRowCount, objMatrix)
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        objMatrix = objForm.Items.Item("mtx").Specific
                        objMatrix.AddRow()
                        objMatrix.FlushToDataSource()
                        Me.SetNewLine(BusinessObjectInfo.FormUID, objMatrix.VisualRowCount, objMatrix)
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
