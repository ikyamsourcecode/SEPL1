Public Class ClsPTN

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
    Dim DocNo As String
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            oUtilities.SAPXML("PTN.xml")
            objForm = oApplication.Forms.GetForm("GEN_PTN", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_PTN")
            objForm.EnableMenu("1288", True)
            objForm.EnableMenu("1289", True)
            objForm.EnableMenu("1290", True)
            objForm.EnableMenu("1291", True)
            objForm.DataBrowser.BrowseBy = "docnum"
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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_PTN")
            oUtilities.GetSeries(FormUID, "series", "GEN_PTN")
            oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("series").Specific.Selected.Value, "GEN_PTN"))
            oDBs_Head.SetValue("U_docdate", 0, DateTime.Today.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_compdate", 0, DateTime.Today.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_status", 0, "Open")
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterProdOrders(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("PRDCFL")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            oCon = oCons.Add()
            oCon.Alias = "Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            oCon.CondVal = "L"
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add()
            oCon.Alias = "Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            oCon.CondVal = "C"
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add()
            oCon.Alias = "OriginNum"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = Trim(objForm.Items.Item("sono").Specific.value)
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add
            oCon.Alias = "Type"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            oCon.CondVal = "D"
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_PTN")
            If Trim(oDBs_Head.GetValue("u_sono", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter Sales Order No", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            If Trim(oDBs_Head.GetValue("u_prdno", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter production order no", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            If Trim(objForm.Items.Item("unit").Specific.value) = "" Then
                oApplication.StatusBar.SetText("Please enter Unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            If Trim(oDBs_Head.GetValue("u_compdate", 0)) = "" Or Trim(oDBs_Head.GetValue("u_docdate", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter document date and completion date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            If oDBs_Head.GetValue("u_accpqty", 0) > 0 Then
                If Trim(oDBs_Head.GetValue("u_accpwhs", 0)) = "" Then
                    oApplication.StatusBar.SetText("please select accepted whs", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                    Exit Function
                End If
            End If
            If Trim(oDBs_Head.GetValue("u_unit", 0)) = "" Then
                oApplication.StatusBar.SetText("Please select unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            If oDBs_Head.GetValue("u_rejqty", 0) > 0 Then
                If Trim(oDBs_Head.GetValue("u_rejwhs", 0)) = "" Then
                    oApplication.StatusBar.SetText("please select rejected whs", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                    Exit Function
                End If
            End If
            If oDBs_Head.GetValue("u_rewqty", 0) > 0 Then
                If Trim(oDBs_Head.GetValue("u_rewwhs", 0)) = "" Then
                    oApplication.StatusBar.SetText("please select reworked whs", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                    Exit Function
                End If
            End If
            If Trim(oDBs_Head.GetValue("u_accpwhs", 0)) <> "" And Trim(oDBs_Head.GetValue("u_rejwhs", 0)) <> "" Then
                If Trim(oDBs_Head.GetValue("u_accpwhs", 0)) = Trim(oDBs_Head.GetValue("u_rejwhs", 0)) Then
                    oApplication.StatusBar.SetText("Accepted and Rejected warehouse cannot be the same", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                    Exit Function
                End If
            End If
            If Trim(oDBs_Head.GetValue("u_accpwhs", 0)) <> "" And Trim(oDBs_Head.GetValue("u_rewwhs", 0)) <> "" Then
                If Trim(oDBs_Head.GetValue("u_accpwhs", 0)) = Trim(oDBs_Head.GetValue("u_rewwhs", 0)) Then
                    oApplication.StatusBar.SetText("Accepted and Rework warehouse cannot be the same", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                    Exit Function
                End If
            End If
            If Trim(oDBs_Head.GetValue("u_rejwhs", 0)) <> "" And Trim(oDBs_Head.GetValue("u_rewwhs", 0)) <> "" Then
                If Trim(oDBs_Head.GetValue("u_rejwhs", 0)) = Trim(oDBs_Head.GetValue("u_rewwhs", 0)) Then
                    oApplication.StatusBar.SetText("Rejected and Rework warehouse cannot be the same", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                    Exit Function
                End If
            End If
            If CDbl(oDBs_Head.GetValue("u_compqty", 0)) > CDbl(oDBs_Head.GetValue("u_prdoqty", 0)) Then
                oApplication.StatusBar.SetText("Completed Quantity cannot be more than Production Order Open Quantity", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            If CDbl(oDBs_Head.GetValue("u_compqty", 0)) <> (CDbl(oDBs_Head.GetValue("u_accpqty", 0)) + CDbl(oDBs_Head.GetValue("u_rejqty", 0)) + CDbl(oDBs_Head.GetValue("u_rewqty", 0))) Then
                oApplication.StatusBar.SetText("Sum of accepted, rejected and reworked should be equal to completed quantity", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            Dim oProductionOrders As SAPbobsCOM.ProductionOrders
            oProductionOrders = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
            oProductionOrders.GetByKey(oDBs_Head.GetValue("u_prdentry", 0))
            If oProductionOrders.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned Then
                oProductionOrders.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased
                Dim ErrFlag As Integer = oProductionOrders.Update()
                If ErrFlag <> 0 Then
                    oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                    Exit Function
                End If
            End If
            DocNo = objForm.Items.Item("docnum").Specific.value
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
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        Dim DocNum As String
                        DocNum = objForm.Items.Item("docnum").Specific.value
                        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery("Select DocNum From OIGE Where u_ptnno = '" + Trim(objForm.Items.Item("docnum").Specific.value) + "'")
                        If oRecordSet.RecordCount > 0 Then
                            objForm.Close()
                            oUtilities.SAPXML("PTN.xml")
                            objForm = oApplication.Forms.GetForm("GEN_PTN", oApplication.Forms.ActiveForm.TypeCount)
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("docnum").Specific.value = DocNum
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Exit Select
                        End If
                        oRecordSet.DoQuery("Select DocNum From OIGN Where u_ptnno = '" + Trim(objForm.Items.Item("docnum").Specific.value) + "'")
                        If oRecordSet.RecordCount > 0 Then
                            objForm.Close()
                            oUtilities.SAPXML("PTN.xml")
                            objForm = oApplication.Forms.GetForm("GEN_PTN", oApplication.Forms.ActiveForm.TypeCount)
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("docnum").Specific.value = DocNum
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Exit Select
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation(FormUID) = False Then BubbleEvent = False
                    ElseIf pVal.ItemUID = "1" And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim UserID As String

                        oRSet.DoQuery("Select UserID From  OUSR Where User_Code = '" + oCompany.UserName.ToString.Trim + "'")
                        UserID = oRSet.Fields.Item("UserID").Value
                        oRSet.DoQuery("Select Max(DocEntry) As 'DocEntry' From [@GEN_PTN] Where UserSign = '" + UserID + "'")
                        DocNo = oRSet.Fields.Item("DocEntry").Value
                        oRSet.DoQuery("Select DocNum from [@GEN_PTN] where DocEntry= '" + DocNo + "' and UserSign = '" + UserID + "'")
                        Dim _str_DocNum As String
                        _str_DocNum = oRSet.Fields.Item("DocNum").Value
                        oApplication.MessageBox("The PTN No of the last added document: - " + _str_DocNum + "")
                        objForm.Close()
                        oUtilities.SAPXML("PTN.xml")
                        objForm = oApplication.Forms.GetForm("GEN_PTN", oApplication.Forms.ActiveForm.TypeCount)
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        objForm.Items.Item("docnum").Specific.value = _str_DocNum
                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objForm.Items.Item("series").DisplayDesc = True
                        'Me.SetDefault(FormUID)
                    End If
                    If pVal.ItemUID = "btnbc" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        RS.DoQuery("Select DocNum From OIGE Where u_sono = '" + Trim(objForm.Items.Item("sono").Specific.value) + "' And u_ptnno = '" + Trim(objForm.Items.Item("docnum").Specific.value) + "'")
                        If RS.RecordCount > 0 Then
                            oApplication.StatusBar.SetText("Consumption already done for this document", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                    If pVal.ItemUID = "btnbc" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim RSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim ProdNo As String = objForm.Items.Item("prdno").Specific.value
                        Dim SONO As String = objForm.Items.Item("sono").Specific.value
                        Dim PTNNo As String = objForm.Items.Item("docnum").Specific.value
                        Dim Unit As String = objForm.Items.Item("unit").Specific.Value
                        Dim Process As String = objForm.Items.Item("process").Specific.value
                        DocNo = objForm.Items.Item("docnum").Specific.value
                        RSet.DoQuery("Select B.u_inwhs From [@GEN_UNIT_MST] A Inner Join [@GEN_UNIT_MST_D0] B On A.Code = B.Code Where A.Name = '" + Unit + "' And B.u_process = '" + Process + "'")
                        If RSet.RecordCount = 0 Then
                            RSet.DoQuery("Select u_inwhs From [@GEN_PROD_PRCS] Where u_itemCode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "'")
                        End If
                        oRecordSet.DoQuery("Select Distinct A.PlannedQty,B.ItemCode,B.warehouse,B.BaseQty ,IsNull(B.LineNum,0) AS 'LN' From OWOR A Inner Join WOR1 B On A.DocEntry = B.DocEntry WHere A.DocNum = '" + Trim(objForm.Items.Item("prdno").Specific.value) + "' And B.IssueType = 'M'")
                        oApplication.ActivateMenuItem("4371")
                        Dim IPForm As SAPbouiCOM.Form
                        Dim IPMatrix As SAPbouiCOM.Matrix
                        IPForm = oApplication.Forms.GetForm("65213", oApplication.Forms.ActiveForm.TypeCount)
                        IPMatrix = IPForm.Items.Item("13").Specific
                        IPForm.Items.Item("sono").Specific.value = SONO
                        IPForm.Items.Item("ptnno").Specific.Value = PTNNo
                        IPMatrix.Columns.Item("U_totavlbl").Editable = True
                        For i As Integer = 1 To oRecordSet.RecordCount
                            Try
                                IPMatrix.Columns.Item("61").Cells.Item(i).Specific.value = ProdNo
                                IPMatrix.Columns.Item("60").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("LN").Value + 1
                                IPMatrix.Columns.Item("9").Cells.Item(i).Specific.Value = CDbl(objForm.Items.Item("compqty").Specific.value) * oRecordSet.Fields.Item("BaseQty").Value
                                IPMatrix.Columns.Item("15").Cells.Item(i).Specific.value = RSet.Fields.Item("u_inwhs").Value
                                Dim ITQty, RetITQty, IssQty As Double
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet.DoQuery("Select IsNull(Sum(B.Quantity),0) As 'ITQty' From OWTR A Inner Join WTR1 B On A.DocEntry = B.DocEntry And A.u_sono = '" + SONO + "' And B.ItemCode = '" + Trim(oRecordSet.Fields.Item("Itemcode").Value) + "' And B.WhsCode = '" + Trim(RSet.Fields.Item("u_inwhs").Value) + "'")
                                ITQty = oRSet.Fields.Item("ITQty").Value
                                oRSet.DoQuery("Select IsNull(Sum(B.Quantity),0) As 'RetITQty' From OWTR A Inner Join WTR1 B On A.DocEntry = B.DocEntry And A.u_sono = '" + SONO + "' And B.ItemCode = '" + Trim(oRecordSet.Fields.Item("Itemcode").Value) + "' And A.Filler = '" + Trim(RSet.Fields.Item("u_inwhs").Value) + "'")
                                RetITQty = oRSet.Fields.Item("RetITQty").Value
                                oRSet.DoQuery("Select IsNull(Sum(B.Quantity),0) As 'IssQty' From OIGE A Inner Join IGE1 B On A.DocEntry = B.DocEntry And A.u_sono = '" + SONO + "' And B.ItemCode = '" + Trim(oRecordSet.Fields.Item("Itemcode").Value) + "' And B.WhsCode = '" + Trim(RSet.Fields.Item("u_inwhs").Value) + "' And B.BaseRef = '" + ProdNo + "'")
                                IssQty = oRSet.Fields.Item("IssQty").Value
                                IPMatrix.Columns.Item("U_totavlbl").Cells.Item(i).Specific.value = (ITQty - RetITQty - IssQty)
                                oRSet.DoQuery("Select OnHand From OITW Where WhsCode = '" + RSet.Fields.Item("u_inwhs").Value + "' And ItemCode = '" + Trim(IPMatrix.Columns.Item("1").Cells.Item(i).Specific.value) + "'")
                                IPMatrix.Columns.Item("U_qty").Cells.Item(i).Specific.value = oRSet.Fields.Item("OnHand").Value
                                oRecordSet.MoveNext()
                                IPMatrix.Columns.Item("61").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Catch ex As Exception
                                oApplication.StatusBar.SetText(ex.Message)
                            End Try
                        Next
                        IPMatrix.Columns.Item("U_totavlbl").Editable = False
                    End If
                    If pVal.ItemUID = "btncnf" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        RS.DoQuery("Select DocNum From OIGN Where u_sono = '" + Trim(objForm.Items.Item("sono").Specific.value) + "' And u_ptnno = '" + Trim(objForm.Items.Item("docnum").Specific.value) + "'")
                        If RS.RecordCount > 0 Then
                            oApplication.StatusBar.SetText("Confirmation already done for this document", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                    If pVal.ItemUID = "btncnf" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim RSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim ProdNo As String = objForm.Items.Item("prdno").Specific.value
                        Dim SONO As String = objForm.Items.Item("sono").Specific.value
                        Dim PTNNo As String = objForm.Items.Item("docnum").Specific.value
                        Dim Unit As String = objForm.Items.Item("unit").Specific.Value
                        Dim Process As String = objForm.Items.Item("process").Specific.value
                        RSet.DoQuery("Select B.u_outwhs From [@GEN_UNIT_MST] A Inner Join [@GEN_UNIT_MST_D0] B On A.Code = B.Code Where A.Name = '" + Unit + "' And B.u_process = '" + Process + "'")
                        If RSet.RecordCount = 0 Then
                            RSet.DoQuery("Select u_outwhs From [@GEN_PROD_PRCS] Where u_itemCode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "'")
                        End If
                        oRecordSet.DoQuery("Select ItemCode,Warehouse,PlannedQty From OWOR Where DocNum = '" + ProdNo + "' And ItemCode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "'  And ItemCode In (Select ItemCode From OITM Where IssueMthd = 'M')")
                        oApplication.ActivateMenuItem("4370")
                        Dim RPForm As SAPbouiCOM.Form
                        Dim RPMatrix As SAPbouiCOM.Matrix
                        RPForm = oApplication.Forms.GetForm("65214", oApplication.Forms.ActiveForm.TypeCount)
                        RPMatrix = RPForm.Items.Item("13").Specific
                        RPForm.Items.Item("sono").Specific.value = SONO
                        RPForm.Items.Item("ptnno").Specific.Value = PTNNo

                        'Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'oRSet.DoQuery("select U_unit from owor where docnum='" + ProdNo + "'")
                        RPForm.Items.Item("unit").Specific.value = Unit
                        Dim oCombo As SAPbouiCOM.ComboBox
                        If objForm.Items.Item("accpqty").Specific.value > 0 And objForm.Items.Item("rejqty").Specific.value > 0 And objForm.Items.Item("rewqty").Specific.value > 0 Then
                            Try
                                RPMatrix.Columns.Item("61").Cells.Item(1).Specific.value = ProdNo
                                oCombo = RPMatrix.Columns.Item("69").Cells.Item(1).Specific
                                oCombo.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                RPMatrix.Columns.Item("9").Cells.Item(1).Specific.value = objForm.Items.Item("accpqty").Specific.value
                                RPMatrix.Columns.Item("15").Cells.Item(1).Specific.value = objForm.Items.Item("accpwhs").Specific.value
                                RPMatrix.Columns.Item("61").Cells.Item(2).Specific.value = ProdNo
                                oCombo = RPMatrix.Columns.Item("69").Cells.Item(2).Specific
                                oCombo.Select("R", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                RPMatrix.Columns.Item("9").Cells.Item(2).Specific.value = objForm.Items.Item("rejqty").Specific.value
                                RPMatrix.Columns.Item("15").Cells.Item(2).Specific.value = objForm.Items.Item("rejwhs").Specific.value
                                RPMatrix.Columns.Item("61").Cells.Item(3).Specific.value = ProdNo
                                oCombo = RPMatrix.Columns.Item("69").Cells.Item(3).Specific
                                oCombo.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                RPMatrix.Columns.Item("9").Cells.Item(3).Specific.value = objForm.Items.Item("rewqty").Specific.value
                                RPMatrix.Columns.Item("15").Cells.Item(3).Specific.value = objForm.Items.Item("rewwhs").Specific.value
                            Catch ex As Exception
                                oApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                        If objForm.Items.Item("accpqty").Specific.value = 0 And objForm.Items.Item("rejqty").Specific.value > 0 And objForm.Items.Item("rewqty").Specific.value > 0 Then
                            Try
                                RPMatrix.Columns.Item("61").Cells.Item(1).Specific.value = ProdNo
                                oCombo = RPMatrix.Columns.Item("69").Cells.Item(1).Specific
                                oCombo.Select("R", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                RPMatrix.Columns.Item("9").Cells.Item(1).Specific.value = objForm.Items.Item("rejqty").Specific.value
                                RPMatrix.Columns.Item("15").Cells.Item(1).Specific.value = objForm.Items.Item("rejwhs").Specific.value
                                RPMatrix.Columns.Item("61").Cells.Item(2).Specific.value = ProdNo
                                oCombo = RPMatrix.Columns.Item("69").Cells.Item(2).Specific
                                oCombo.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                RPMatrix.Columns.Item("9").Cells.Item(2).Specific.value = objForm.Items.Item("rewqty").Specific.value
                                RPMatrix.Columns.Item("15").Cells.Item(2).Specific.value = objForm.Items.Item("rewwhs").Specific.value
                            Catch ex As Exception
                                oApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                        If objForm.Items.Item("accpqty").Specific.value > 0 And objForm.Items.Item("rejqty").Specific.value = 0 And objForm.Items.Item("rewqty").Specific.value > 0 Then
                            Try
                                RPMatrix.Columns.Item("61").Cells.Item(1).Specific.value = ProdNo
                                oCombo = RPMatrix.Columns.Item("69").Cells.Item(1).Specific
                                oCombo.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                RPMatrix.Columns.Item("9").Cells.Item(1).Specific.value = objForm.Items.Item("accpqty").Specific.value
                                RPMatrix.Columns.Item("15").Cells.Item(1).Specific.value = objForm.Items.Item("accpwhs").Specific.value
                                RPMatrix.Columns.Item("61").Cells.Item(2).Specific.value = ProdNo
                                oCombo = RPMatrix.Columns.Item("69").Cells.Item(2).Specific
                                oCombo.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                RPMatrix.Columns.Item("9").Cells.Item(2).Specific.value = objForm.Items.Item("rewqty").Specific.value
                                RPMatrix.Columns.Item("15").Cells.Item(2).Specific.value = objForm.Items.Item("rewwhs").Specific.value
                            Catch ex As Exception
                                oApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                        If objForm.Items.Item("accpqty").Specific.value > 0 And objForm.Items.Item("rejqty").Specific.value > 0 And objForm.Items.Item("rewqty").Specific.value = 0 Then
                            Try
                                RPMatrix.Columns.Item("61").Cells.Item(1).Specific.value = ProdNo
                                oCombo = RPMatrix.Columns.Item("69").Cells.Item(1).Specific
                                oCombo.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                RPMatrix.Columns.Item("9").Cells.Item(1).Specific.value = objForm.Items.Item("accpqty").Specific.value
                                RPMatrix.Columns.Item("15").Cells.Item(1).Specific.value = objForm.Items.Item("accpwhs").Specific.value
                                RPMatrix.Columns.Item("61").Cells.Item(2).Specific.value = ProdNo
                                oCombo = RPMatrix.Columns.Item("69").Cells.Item(2).Specific
                                oCombo.Select("R", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                RPMatrix.Columns.Item("9").Cells.Item(2).Specific.value = objForm.Items.Item("rejqty").Specific.value
                                RPMatrix.Columns.Item("15").Cells.Item(2).Specific.value = objForm.Items.Item("rejwhs").Specific.value
                            Catch ex As Exception
                                oApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                        If objForm.Items.Item("accpqty").Specific.value = 0 And objForm.Items.Item("rejqty").Specific.value = 0 And objForm.Items.Item("rewqty").Specific.value > 0 Then
                            Try
                                RPMatrix.Columns.Item("61").Cells.Item(1).Specific.value = ProdNo
                                oCombo = RPMatrix.Columns.Item("69").Cells.Item(1).Specific
                                oCombo.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                RPMatrix.Columns.Item("9").Cells.Item(1).Specific.value = objForm.Items.Item("rewqty").Specific.value
                                RPMatrix.Columns.Item("15").Cells.Item(1).Specific.value = objForm.Items.Item("rewwhs").Specific.value
                            Catch ex As Exception
                                oApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                        If objForm.Items.Item("accpqty").Specific.value > 0 And objForm.Items.Item("rejqty").Specific.value = 0 And objForm.Items.Item("rewqty").Specific.value = 0 Then
                            Try
                                RPMatrix.Columns.Item("61").Cells.Item(1).Specific.value = ProdNo
                                oCombo = RPMatrix.Columns.Item("69").Cells.Item(1).Specific
                                oCombo.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                RPMatrix.Columns.Item("9").Cells.Item(1).Specific.value = objForm.Items.Item("accpqty").Specific.value
                                RPMatrix.Columns.Item("15").Cells.Item(1).Specific.value = objForm.Items.Item("accpwhs").Specific.value
                            Catch ex As Exception
                                oApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                        If objForm.Items.Item("accpqty").Specific.value = 0 And objForm.Items.Item("rejqty").Specific.value > 0 And objForm.Items.Item("rewqty").Specific.value = 0 Then
                            Try
                                RPMatrix.Columns.Item("61").Cells.Item(1).Specific.value = ProdNo
                                oCombo = RPMatrix.Columns.Item("69").Cells.Item(1).Specific
                                oCombo.Select("R", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                RPMatrix.Columns.Item("9").Cells.Item(1).Specific.value = objForm.Items.Item("rejqty").Specific.value
                                RPMatrix.Columns.Item("15").Cells.Item(1).Specific.value = objForm.Items.Item("rejwhs").Specific.value
                            Catch ex As Exception
                                oApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If (pVal.ItemUID = "prdno" Or pVal.ItemUID = "itemcode" Or pVal.ItemUID = "series" Or pVal.ItemUID = "docdate" Or pVal.ItemUID = "compqty" Or pVal.ItemUID = "accpqty" Or pVal.ItemUID = "rejqty" Or pVal.ItemUID = "rewqty" Or pVal.ItemUID = "accpwhs" Or pVal.ItemUID = "rejwhs" Or pVal.ItemUID = "rewwhs") _
                       And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If objForm.DataSources.DBDataSources.Item("@GEN_PTN").GetValue("u_sono", 0) = "" Then
                            oApplication.StatusBar.SetText("Please select Sales Order no", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
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
                        If oCFL.UniqueID = "PRDCFL" Then
                            If oDBs_Head.GetValue("u_sono", 0) = "" Then
                                oApplication.StatusBar.SetText("Please select sales order no", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            Me.FilterProdOrders(FormUID)
                        End If
                        If oCFL.UniqueID = "UNTCFL" Then
                            If oDBs_Head.GetValue("u_prdno", 0) = "" Then
                                oApplication.StatusBar.SetText("Please select the production order", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    Else
                        If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_PTN")
                            If oCFL.UniqueID = "SOCFL" Then
                                oDBs_Head.SetValue("u_sono", 0, oDT.GetValue("DocNum", 0))
                                oDBs_Head.SetValue("u_soentry", 0, oDT.GetValue("DocEntry", 0))
                                oDBs_Head.SetValue("u_soref", 0, oDT.GetValue("NumAtCard", 0))
                                oDBs_Head.SetValue("u_prdno", 0, "")
                                oDBs_Head.SetValue("u_prdentry", 0, "")
                                oDBs_Head.SetValue("u_itemcode", 0, "")
                                oDBs_Head.SetValue("u_itemname", 0, "")
                                oDBs_Head.SetValue("u_prdqty", 0, "")
                                oDBs_Head.SetValue("u_prdoqty", 0, "")
                                oDBs_Head.SetValue("u_compqty", 0, "")
                                oDBs_Head.SetValue("u_accpqty", 0, "")
                                oDBs_Head.SetValue("u_accpwhs", 0, "")
                                oDBs_Head.SetValue("u_rejqty", 0, "")
                                oDBs_Head.SetValue("u_rejwhs", 0, "")
                                oDBs_Head.SetValue("u_rewqty", 0, "")
                                oDBs_Head.SetValue("u_rewqty", 0, "")
                                oDBs_Head.SetValue("u_unit", 0, "")
                            End If
                            If oCFL.UniqueID = "PRDCFL" Then
                                oDBs_Head.SetValue("u_prdno", 0, oDT.GetValue("DocNum", 0))
                                oDBs_Head.SetValue("u_prdentry", 0, oDT.GetValue("DocEntry", 0))
                                oDBs_Head.SetValue("u_itemcode", 0, oDT.GetValue("ItemCode", 0))
                                Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("Select ItemName From OITM Where ItemCode = '" + Trim(oDT.GetValue("ItemCode", 0)) + "'")
                                oDBs_Head.SetValue("u_itemname", 0, oRecordSet.Fields.Item("ItemName").Value)
                                oDBs_Head.SetValue("u_prdqty", 0, oDT.GetValue("PlannedQty", 0))
                                oDBs_Head.SetValue("u_prdoqty", 0, (oDT.GetValue("PlannedQty", 0) - oDT.GetValue("CmpltQty", 0)))
                                oDBs_Head.SetValue("u_process", 0, oDT.GetValue("U_process", 0))
                                oDBs_Head.SetValue("u_unit", 0, oDT.GetValue("U_unit", 0))
                                oRecordSet.DoQuery("Select B.u_outwhs From [@GEN_UNIT_MST] A Inner Join [@GEN_UNIT_MST_D0] B On A.Code = B.Code Where A.Name = '" + Trim(oDT.GetValue("U_unit", 0)) + "' And B.u_process = '" + Trim(oDT.GetValue("U_process", 0)) + "'")
                                If oRecordSet.RecordCount = 0 Then
                                    oRecordSet.DoQuery("Select u_outwhs From [@GEN_PROD_PRCS] Where u_itemCode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "'")
                                End If
                                oDBs_Head.SetValue("u_accpwhs", 0, oRecordSet.Fields.Item("u_outwhs").Value)
                            End If
                            If oCFL.UniqueID = "ITCFL" Then
                                oDBs_Head.SetValue("u_itemcode", 0, oDT.GetValue("ItemCode", 0))
                                oDBs_Head.SetValue("u_itemname", 0, oDT.GetValue("ItemName", 0))
                            End If
                            If oCFL.UniqueID = "W1CFL" Then
                                oDBs_Head.SetValue("u_accpwhs", 0, oDT.GetValue("WhsCode", 0))
                            End If
                            If oCFL.UniqueID = "W2CFL" Then
                                oDBs_Head.SetValue("u_rejwhs", 0, oDT.GetValue("WhsCode", 0))
                            End If
                            If oCFL.UniqueID = "W3CFL" Then
                                oDBs_Head.SetValue("u_rewwhs", 0, oDT.GetValue("WhsCode", 0))
                            End If
                            If oCFL.UniqueID = "UNTCFL" Then
                                oDBs_Head.SetValue("u_unit", 0, oDT.GetValue("Name", 0))
                                oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("Select B.u_outwhs From [@GEN_UNIT_MST] A Inner Join [@GEN_UNIT_MST_D0] B On A.Code = B.Code Where A.Name = '" + Trim(oDT.GetValue("Name", 0)) + "' And B.u_process = '" + Trim(objForm.Items.Item("process").Specific.value) + "'")
                                oDBs_Head.SetValue("u_accpwhs", 0, oRecordSet.Fields.Item("u_outwhs").Value)
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
                    Case "GEN_PTN"
                        Me.CreateForm(objForm.UniqueID)
                    Case "1282"
                        If objForm.TypeEx = "GEN_PTN" Then
                            Me.SetDefault(objForm.UniqueID)
                            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        End If
                    Case "1281"
                        If objForm.TypeEx = "GEN_PTN" Then
                            objForm.EnableMenu("1282", True)
                            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        End If
                    Case "1288", "1289", "1290", "1291"
                        If objForm.TypeEx = "GEN_PTN" Then
                            objForm.EnableMenu("1282", True)
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
                    If BusinessObjectInfo.ActionSuccess = True Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_PTN")
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                        objForm.Items.Item("series").DisplayDesc = True
                        If Trim(oDBs_Head.GetValue("canceled", 0)) = "Y" Then
                            objForm.Items.Item("btncnf").Enabled = False
                            objForm.Items.Item("btnbc").Enabled = False
                        Else
                            objForm.Items.Item("btncnf").Enabled = True
                            objForm.Items.Item("btnbc").Enabled = True
                            If Trim(oDBs_Head.GetValue("u_status", 0)) = "Open" Then
                                objForm.Items.Item("btncnf").Enabled = False
                                objForm.Items.Item("btnbc").Enabled = True
                                objForm.EnableMenu("1284", True)
                            End If
                            If Trim(oDBs_Head.GetValue("u_status", 0)) = "Consumed" Then
                                objForm.Items.Item("btncnf").Enabled = True
                                objForm.Items.Item("btnbc").Enabled = False
                                objForm.EnableMenu("1284", False)
                            End If
                            If Trim(oDBs_Head.GetValue("u_status", 0)) = "Confirmed" Then
                                objForm.Items.Item("btncnf").Enabled = False
                                objForm.Items.Item("btnbc").Enabled = False
                                objForm.EnableMenu("1284", False)
                            End If
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.ActionSuccess = True Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
