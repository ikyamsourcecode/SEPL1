Public Class ClsSubContract_Return

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
#End Region


    Sub CreateForm()
        Try
            oUtilities.SAPXML("SubContractingReturn.xml")
            objForm = oApplication.Forms.GetForm("GEN_SCRET", oApplication.Forms.ActiveForm.TypeCount)
            objForm.Items.Item("cardcode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("c_series").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("t_docdt").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("t_docno").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET_D0")

            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Select SlpCode,SlpName from OSLP")
            Dim objcombobuyer As SAPbouiCOM.ComboBox
            objcombobuyer = objForm.Items.Item("Buyer").Specific
            For i As Integer = 1 To oRS.RecordCount
                objcombobuyer.ValidValues.Add(Trim(oRS.Fields.Item("SlpCode").Value), Trim(oRS.Fields.Item("SlpName").Value))
                oRS.MoveNext()
            Next

            objForm.Select()
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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET_D0")
            oUtilities.GetSeries(FormUID, "c_series", "GEN_SC_RET")
            oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("c_series").Specific.Selected.Value, "GEN_SC_RET"))
            oDBs_Head.SetValue("U_Status", 0, "Open")
            oDBs_Head.SetValue("U_DcDat", 0, DateTime.Today.ToString("yyyyMMdd"))
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Select ISNULL(T0.firstName,'') + ' ' + ISNULL(T0.lastName,'') Owner,empid from OHEM T0 INNER JOIN OUSR T1 ON T0.UserID=T1.USERID Where USER_CODE='" & oCompany.UserName & "'")
            If oRS.RecordCount > 0 Then
                oDBs_Head.SetValue("U_Owner", 0, Trim(oRS.Fields.Item("Owner").Value))
                oDBs_Head.SetValue("U_OwnerCod", 0, Trim(oRS.Fields.Item("empid").Value))
            End If


            oDBs_Head.SetValue("U_RefNo", 0, "")
            oDBs_Head.SetValue("U_SCNo", 0, "")
            oDBs_Head.SetValue("U_SCDat", 0, "")
            oDBs_Head.SetValue("U_ItemNo", 0, "")
            oDBs_Head.SetValue("U_Qty", 0, 0)
            oDBs_Head.SetValue("U_Total", 0, 0)
            oDBs_Head.SetValue("U_Buyer", 0, "")

            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            objMatrix.Clear()
            objMatrix.AddRow()
            objMatrix.FlushToDataSource()
            Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET_D0")
            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
            oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_Qty", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_FWhs", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, "")
            objMatrix.SetLineData(Row)
            objMatrix.FlushToDataSource()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "c_series" And pVal.BeforeAction = False Then
                        oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("c_series").Specific.Selected.Value, "GEN_SC_RET"))
                    End If
                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    If pVal.ItemUID = "37" And pVal.BeforeAction = False And pVal.CharPressed = 9 And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        If objMatrix.VisualRowCount > 0 Then
                            objMatrix.Columns.Item("itemno").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Exit Sub
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If pVal.ItemUID = "37" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        'objForm = oApplication.Forms.Item(FormUID)
                        'Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'oRS.DoQuery("Select T1.U_ItemCode,ISNULL(T1.U_DCQty,0)- ISNULL(T1.U_RetQty,0) Quantity from [@GEN_SUB_CONTRACT] T0 INNER JOIN [@GEN_SUB_CONTRACT_D1] T1 on T0.DocEntry=T1.DocEntry Where T0.DocNum= '" & Trim(oDBs_Head.GetValue("U_SCNo", 0)) & "' and T1.U_Father='" & Trim(objForm.Items.Item("ItemNo").Specific.Value) & "'")
                        'If oRS.RecordCount > 0 Then
                        '    If CDbl(objForm.Items.Item("37").Specific.Value) > CDbl(oRS.Fields.Item("Quantity").Value) Then
                        '        oApplication.StatusBar.SetText("Item Quantity should not be exceed " & CDbl(oRS.Fields.Item("Quantity").Value), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '        BubbleEvent = False
                        '    End If
                        'End If
                    ElseIf pVal.ItemUID = "t_docdt" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Trim(objForm.Items.Item("t_docdt").Specific.Value).Equals("") = False Then
                            If (DateDiff(DateInterval.Day, DateTime.ParseExact(Trim(objForm.Items.Item("t_docdt").Specific.Value), "yyyyMMdd", Nothing), DateTime.Today)) <> 0 Then
                                oApplication.StatusBar.SetText("Document date varies from system date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        End If
                    ElseIf pVal.ItemUID = "t_deldt" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Trim(objForm.Items.Item("t_deldt").Specific.Value).Equals("") = False Then
                            If (DateDiff(DateInterval.Day, DateTime.ParseExact(Trim(objForm.Items.Item("t_docdt").Specific.Value), "yyyyMMdd", Nothing), DateTime.ParseExact(Trim(objForm.Items.Item("t_deldt").Specific.Value), "yyyyMMdd", Nothing))) < 0 Then
                                oApplication.StatusBar.SetText("Delivery date is before Document date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation(FormUID) = False Then BubbleEvent = False
                    ElseIf pVal.ItemUID = "1" And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If pVal.ItemUID = "37" And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Try
                            objForm = oApplication.Forms.Item(FormUID)
                            objForm.Freeze(True)
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET")
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET_D0")
                            For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                                oDBs_Detail.Offset = Row - 1
                                oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
                                oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemno").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("desc").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_fwhs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("fwhs").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("stock").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("twhs").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("BOMQty").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("BOMQty").Cells.Item(Row).Specific.Value) * CDbl(objForm.Items.Item("37").Specific.Value))
                                oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("price").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("price").Cells.Item(Row).Specific.Value) * (CDbl(objMatrix.Columns.Item("BOMQty").Cells.Item(Row).Specific.Value) * CDbl(objForm.Items.Item("37").Specific.Value)))
                                oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("rmrk").Cells.Item(Row).Specific.Value))
                                objMatrix.SetLineData(Row)
                            Next
                            Me.CalculateTotal(FormUID)
                            objForm.Freeze(False)
                        Catch ex As Exception
                            oApplication.StatusBar.SetText(ex.Message)
                            objForm.Freeze(False)
                        End Try
                    ElseIf pVal.ItemUID = "ItemMatrix" And (pVal.ColUID = "issqty" Or pVal.ColUID = "price") And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        oDBs_Detail.Offset = pVal.Row - 1
                        oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                        oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemno").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("desc").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_fwhs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("fwhs").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("stock").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("twhs").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("BOMQty").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("issqty").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("price").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("price").Cells.Item(pVal.Row).Specific.Value) * CDbl(objMatrix.Columns.Item("issqty").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("rmrk").Cells.Item(pVal.Row).Specific.Value))
                        objMatrix.SetLineData(pVal.Row)
                        Me.CalculateTotal(FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET")
                        oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET_D0")
                        If oCFL.UniqueID = "CFL_Vendor" Then
                            oDBs_Head.SetValue("U_CardCode", 0, oDT.GetValue("CardCode", 0))
                            oDBs_Head.SetValue("U_CardName", 0, oDT.GetValue("CardName", 0))
                            oDBs_Head.SetValue("U_SCDocNo", 0, "")
                            oDBs_Head.SetValue("U_SCNo", 0, "")
                            oDBs_Head.SetValue("U_SCDat", 0, "")
                            oDBs_Head.SetValue("U_RefNo", 0, "")
                            oDBs_Head.SetValue("U_Buyer", 0, "")
                            oDBs_Head.SetValue("U_ContPer", 0, "")
                            Me.FilterSC(FormUID, oDT.GetValue("CardCode", 0))
                        ElseIf oCFL.UniqueID = "CFL_SCNo" Then
                            oDBs_Head.SetValue("U_SCDocNo", 0, oDT.GetValue("DocEntry", 0))
                            oDBs_Head.SetValue("U_SCNo", 0, oDT.GetValue("DocNum", 0))
                            oDBs_Head.SetValue("U_SCDat", 0, Format(oDT.GetValue("U_DocDate", 0), "yyyyMMdd"))
                            oDBs_Head.SetValue("U_RefNo", 0, oDT.GetValue("U_VendRef", 0))
                            oDBs_Head.SetValue("U_Buyer", 0, oDT.GetValue("U_Buyer", 0))
                            oDBs_Head.SetValue("U_ContPer", 0, oDT.GetValue("U_ContPer", 0))
                            Me.FilterItemBOM(FormUID)
                        ElseIf oCFL.UniqueID = "CFL_BOMITM" Then
                            oDBs_Head.SetValue("U_ItemNo", 0, oDT.GetValue("ItemCode", 0))
                            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRS.DoQuery("Select (ISNULL(U_DCQty,0)-ISNULL(U_RetQty,0)) Quantity from [@GEN_SUB_CONTRACT_D0] Where DocEntry='" & Trim(oDBs_Head.GetValue("U_SCDocNo", 0)) & "' and U_ItemCode='" & oDT.GetValue("ItemCode", 0) & "'")
                            oDBs_Head.SetValue("U_Qty", 0, CDbl(oRS.Fields.Item("Quantity").Value))
                            Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
                            Me.CalculateTotal(FormUID)
                            Me.FilterChildItems(FormUID)
                            'oRS.DoQuery("select U_VendWhs VendWhs from [@GEN_SUB_CONTRACT] Where DocEntry='" & Trim(oDBs_Head.GetValue("U_SCDocNo", 0)) & "'")
                            'Dim VendorWhs As String = Trim(oRS.Fields.Item("VendWhs").Value)
                            'oRS.DoQuery("Select T1.ItemCode,T1.ItemName,T0.Quantity BOMQty,T2.AvgPrice Price,T1.InvntryUoM UoM,T2.OnHand InStock,T2.WhsCode from ITT1 T0 INNER JOIN OITM T1 ON T0.Code=T1.ItemCode INNER JOIN OITW T2 ON T2.ItemCode=T1.ItemCode and T2.WhsCode='" & VendorWhs & "' Where T0.Father='" & oDT.GetValue("ItemCode", 0) & "'")
                            'objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            'objMatrix.Clear()
                            'For Row As Integer = 1 To oRS.RecordCount
                            '    objMatrix.AddRow()
                            '    objMatrix.FlushToDataSource()
                            '    oDBs_Detail.Offset = Row - 1
                            '    oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
                            '    oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, Trim(oRS.Fields.Item("ItemCode").Value))
                            '    oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, Trim(oRS.Fields.Item("ItemName").Value))
                            '    oDBs_Detail.SetValue("U_fwhs", oDBs_Detail.Offset, Trim(oRS.Fields.Item("WhsCode").Value))
                            '    oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("InStock").Value))
                            '    oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, "")
                            '    oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, Trim(oRS.Fields.Item("UoM").Value))
                            '    oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("BOMQty").Value))
                            '    oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("BOMQty").Value) * CDbl(objForm.Items.Item("37").Specific.Value))
                            '    oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("Price").Value))
                            '    oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("Price").Value) * (CDbl(oRS.Fields.Item("BOMQty").Value) * CDbl(objForm.Items.Item("37").Specific.Value)))
                            '    oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, "")
                            '    objMatrix.SetLineData(Row)
                            '    oRS.MoveNext()
                            'Next
                            'objMatrix.FlushToDataSource()
                            'objMatrix.AddRow()
                            'objMatrix.FlushToDataSource()
                            'objMatrix.AutoResizeColumns()


                        ElseIf oCFL.UniqueID = "CFL_Owner" Then
                            oDBs_Head.SetValue("U_OwnerCod", 0, oDT.GetValue("empID", 0))
                            oDBs_Head.SetValue("U_Owner", 0, oDT.GetValue("firstName", 0) + " " + oDT.GetValue("lastName", 0))
                        ElseIf oCFL.UniqueID = "CFL_twhs" Then
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            oDBs_Detail.Offset = pVal.Row - 1
                            oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                            oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemno").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("desc").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_fwhs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("fwhs").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("stock").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, Trim(oDT.GetValue("WhsCode", 0)))
                            oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("BOMQty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("issqty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("price").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("total").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("rmrk").Cells.Item(pVal.Row).Specific.Value))
                            objMatrix.SetLineData(pVal.Row)
                        ElseIf oCFL.UniqueID = "CFL_Item" Then
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRS.DoQuery("select U_VendWhs VendWhs from [@GEN_SUB_CONTRACT] Where DocEntry='" & Trim(oDBs_Head.GetValue("U_SCDocNo", 0)) & "'")
                            Dim VendorWhs As String = Trim(oRS.Fields.Item("VendWhs").Value)
                            Dim OrginRow As Integer = objMatrix.VisualRowCount
                            For i As Integer = 0 To oDT.Rows.Count - 1
                                Dim cflSelectedcount As Integer = oDT.Rows.Count
                                If i < cflSelectedcount - 1 Then
                                    objMatrix.AddRow(1, pVal.Row)
                                    oDBs_Detail.InsertRecord(pVal.Row + i - 1)
                                End If
                                oDBs_Detail.Offset = pVal.Row - 1 + i
                                oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                                oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, oDT.GetValue("ItemCode", i))
                                oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, oDT.GetValue("ItemName", i))
                                oDBs_Detail.SetValue("U_fwhs", oDBs_Detail.Offset, oRS.Fields.Item("VendWhs").Value)
                                Dim WHs As String = oRS.Fields.Item("VendWhs").Value
                                oRS.DoQuery("Select OnHand,AvgPrice from OITW Where ItemCode = '" + Trim(oDT.GetValue("ItemCode", i)) + "' And WhsCode = '" + WHs + "'")
                                oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, oRS.Fields.Item("OnHand").Value)
                                oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, "")
                                oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, 0)
                                oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, 0)
                                oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, oRS.Fields.Item("AvgPrice").Value)
                                oRS.DoQuery("Select DfltWh From OITM WHere ItemCOde = '" + Trim(oDT.GetValue("ItemCode", i)) + "'")
                                oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, oRS.Fields.Item("DfltWh").Value)
                                oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, 0)
                                oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, "")
                                objMatrix.SetLineData(pVal.Row + i)
                            Next
                            objMatrix.FlushToDataSource()
                            If OrginRow = pVal.Row Then
                                objMatrix.AddRow()
                                objMatrix.FlushToDataSource()
                                Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
                            End If
                            For Row As Integer = 1 To objMatrix.VisualRowCount
                                objMatrix.Columns.Item("SNo").Cells.Item(Row).Specific.Value = Row
                            Next
                            Me.CalculateTotal(FormUID)
                        ElseIf oCFL.UniqueID = "CFL_fwhs" Then
                            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRS.DoQuery("Select OnHand,AvgPrice from OITW Where ItemCode='" & Trim(objMatrix.Columns.Item("itemno").Cells.Item(pVal.Row).Specific.Value) & "' and WhsCode='" & Trim(oDT.GetValue("WhsCode", 0)) & "'")
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            oDBs_Detail.Offset = pVal.Row - 1
                            oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                            oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemno").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("desc").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_fwhs", oDBs_Detail.Offset, Trim(oDT.GetValue("WhsCode", 0)))
                            oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("OnHand").Value))
                            oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("twhs").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("BOMQty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("issqty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("AvgPrice").Value))
                            oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("AvgPrice").Value) * CDbl(objMatrix.Columns.Item("issqty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("rmrk").Cells.Item(pVal.Row).Specific.Value))
                            objMatrix.SetLineData(pVal.Row)
                            Me.CalculateTotal(FormUID)
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
                Select Case pVal.MenuUID
                    Case "6913"
                        If pVal.BeforeAction = True Then
                            objForm = oApplication.Forms.ActiveForm
                            If objForm.TypeEx = "GEN_SCForm" Then
                                Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRs.DoQuery("Select ISNull(SuperUser,'') SuperUser from OUSR where USER_CODE = '" & oCompany.UserName & "'")
                                If oRs.Fields.Item(0).Value = "Y" Then
                                    BubbleEvent = True
                                Else
                                    BubbleEvent = False
                                End If
                            End If
                        End If
                End Select
            ElseIf pVal.BeforeAction = False Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "SC_RETURN"
                        If pVal.BeforeAction = False Then
                            Me.CreateForm()
                        End If
                    Case "1282"
                        If objForm.TypeEx = "GEN_SCRET" Then
                            Me.SetDefault(objForm.UniqueID)
                        End If
                    Case "1281"
                        If objForm.TypeEx = "GEN_SCRET" Then
                            objForm.EnableMenu("1282", True)
                            objForm.Items.Item("t_docno").Click()
                        End If
                    Case "Close"
                        If objForm.TypeEx = "GEN_SCRET" Then
                            If oApplication.MessageBox("Do you want to close?", 2, "Ok", "Cancel") = 1 Then
                                Dim ORS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                ORS.DoQuery("UPDATE [@GEN_SC_DC] SET U_Status='Closed' Where DocNum='" & oDBs_Head.GetValue("DocNum", 0) & "'")
                                oDBs_Head.SetValue("U_Status", 0, "Closed")
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                objForm.Items.Item("1").Enabled = True
                            End If
                        End If
                    Case "1293"
                        If objForm.TypeEx = "GEN_SCRET" Then
                            objForm.Freeze(True)
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET")
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET_D0")
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            For Row As Integer = 1 To objMatrix.VisualRowCount
                                objMatrix.GetLineData(Row)
                                oDBs_Detail.Offset = Row - 1
                                oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
                                oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemno").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("desc").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_fwhs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("fwhs").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("stock").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("twhs").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("BOMQty").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("issqty").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("price").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("total").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("rmrk").Cells.Item(Row).Specific.Value))
                                objMatrix.SetLineData(Row)
                            Next
                            objMatrix.FlushToDataSource()
                            oDBs_Detail.RemoveRecord(oDBs_Detail.Size - 1)
                            objMatrix.LoadFromDataSource()
                            objForm.Freeze(False)
                        End If
                End Select
            End If

        Catch ex As Exception

        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Select Case BusinessObjectInfo.EventType
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET")
                oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET_D0")
                If BusinessObjectInfo.BeforeAction = True Then
                    objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                        oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("c_series").Specific.Selected.Value, "GEN_SC_RET"))
                        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRS.DoQuery("UPDATE OWTR SET U_Type='SubCont_Return',U_DocNum='" & Trim(oDBs_Head.GetValue("DocNum", 0)) & "' Where DocEntry='" & Trim(oDBs_Head.GetValue("U_InvTrNo", 0)) & "'")
                        oRS.DoQuery("UPDATE [@GEN_SUB_CONTRACT_D0] SET U_RetQty=ISNULL(U_RetQty,0)+" & CDbl(objForm.Items.Item("37").Specific.Value) & " Where DocEntry='" & Trim(oDBs_Head.GetValue("U_SCDocNo", 0)) & "' and U_ItemCode='" & Trim(objForm.Items.Item("ItemNo").Specific.value) & "'")
                    End If

                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                    objMatrix.LoadFromDataSource()
                    If objMatrix.VisualRowCount <> 0 Then
                        oDBs_Detail.RemoveRecord(objMatrix.VisualRowCount - 1)
                        objMatrix.LoadFromDataSource()
                    End If
                ElseIf BusinessObjectInfo.ActionSuccess = True Then
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Then
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        objMatrix.AddRow()
                        objMatrix.FlushToDataSource()
                        Me.SetNewLine(BusinessObjectInfo.FormUID, objMatrix.VisualRowCount)
                    End If
                End If
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                If BusinessObjectInfo.ActionSuccess = True Then
                    objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    objForm.EnableMenu("1282", True)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET")
                    oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET_D0")
                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                    objMatrix.AddRow()
                    objMatrix.FlushToDataSource()
                    Me.SetNewLine(BusinessObjectInfo.FormUID, objMatrix.VisualRowCount)
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                    objForm.Items.Item("1").Enabled = True
                End If
        End Select
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET_D0")

            If Trim(objForm.Items.Item("cardcode").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Vendor Code should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Trim(objForm.Items.Item("SCNo").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Subcontracting PO.No. should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Trim(objForm.Items.Item("ItemNo").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Item No. should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Trim(objForm.Items.Item("t_deldt").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Delivery Date should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Trim(objForm.Items.Item("t_docdt").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Document Date should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf (DateDiff(DateInterval.Day, DateTime.ParseExact(Trim(objForm.Items.Item("t_docdt").Specific.Value), "yyyyMMdd", Nothing), DateTime.ParseExact(Trim(objForm.Items.Item("t_deldt").Specific.Value), "yyyyMMdd", Nothing))) < 0 Then
                oApplication.StatusBar.SetText("Delivery date is before Document date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                'ElseIf CDbl(objForm.Items.Item("37").Specific.Value) <= 0 Then
                '    oApplication.StatusBar.SetText("Item Quantity should be greater than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Return False
            End If

            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            If objMatrix.VisualRowCount = 1 Then
                oApplication.StatusBar.SetText("No items defined", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                    If Trim(objMatrix.Columns.Item("itemno").Cells.Item(Row).Specific.Value).Equals("") = True Then
                        oApplication.StatusBar.SetText("Row [ " & Row & " ] - ItemCode should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf Trim(objMatrix.Columns.Item("fwhs").Cells.Item(Row).Specific.Value).Equals("") = True Then
                        oApplication.StatusBar.SetText("Row [ " & Row & " ] - From Warehouse should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf Trim(objMatrix.Columns.Item("twhs").Cells.Item(Row).Specific.Value).Equals("") = True Then
                        oApplication.StatusBar.SetText("Row [ " & Row & " ] - To Warehouse should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf Trim(objMatrix.Columns.Item("fwhs").Cells.Item(Row).Specific.Value).Equals(Trim(objMatrix.Columns.Item("twhs").Cells.Item(Row).Specific.Value)) = True Then
                        oApplication.StatusBar.SetText("Row [ " & Row & " ] - From and To Warehouse should not be same", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf CDbl(objMatrix.Columns.Item("issqty").Cells.Item(Row).Specific.Value) <= 0 Then
                        oApplication.StatusBar.SetText("Row [ " & Row & " ] - Issue Quantity should greater than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf CDbl(objMatrix.Columns.Item("issqty").Cells.Item(Row).Specific.Value) > CDbl(objMatrix.Columns.Item("stock").Cells.Item(Row).Specific.Value) Then
                        oApplication.StatusBar.SetText("Row [ " & Row & " ] - Issue Quantity should greater than InStock", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                Next
            End If

            Dim TransferNo As Integer = PostStockTransfer(FormUID)
            If TransferNo <> 0 Then
                objForm.Items.Item("InvTrNo").Specific.Value = TransferNo
            End If

            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
            Return False
        End Try
    End Function

    Sub CalculateTotal(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET_D0")
            Dim TotalAmount As Double = 0
            For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                TotalAmount = TotalAmount + CDbl(objMatrix.Columns.Item("total").Cells.Item(Row).Specific.Value)
            Next
            oDBs_Head.SetValue("U_Total", 0, TotalAmount)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterSC(ByVal FormUID As String, ByVal CardCode As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("CFL_SCNo")
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_CardCode"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = Trim(objForm.Items.Item("cardcode").Specific.Value)
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add()
            oCon.Alias = "U_Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Open"
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
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

    Sub FilterItemBOM(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("CFL_BOMITM")
            oCFL.SetConditions(emptyConds)
            If Trim(objForm.Items.Item("SCNo").Specific.value) <> "" Then
                objForm = oApplication.Forms.Item(FormUID)
                oCFL.SetConditions(emptyConds)
                oCons = oCFL.GetConditions()
                Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRS.DoQuery("Select Distinct C.U_ItemCode from [@GEN_SUB_CONTRACT] a Inner join [@GEN_SUB_CONTRACT_D1] b on a.docentry=b.docentry Inner Join [@GEN_SUB_CONTRACT_D0] C On B.DocEntry = C.DocEntry ANd C.u_ItemCode = B.u_Father where  a.DocNum= '" & Trim(objForm.Items.Item("SCNo").Specific.Value) & "'")
                For i As Integer = 0 To oRS.RecordCount - 1
                    If i > 0 Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRS.Fields.Item("U_ItemCode").Value
                    oRS.MoveNext()
                Next
                If oRS.RecordCount = 0 Then
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "-1"
                End If
                oCFL.SetConditions(oCons)
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterChildItems(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("CFL_Item")
            oCFL.SetConditions(emptyConds)
            If Trim(objForm.Items.Item("SCNo").Specific.value) <> "" Then
                objForm = oApplication.Forms.Item(FormUID)
                oCFL.SetConditions(emptyConds)
                oCons = oCFL.GetConditions()
                Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRS.DoQuery("Select Distinct B.U_Code from [@GEN_SUB_CONTRACT] a Inner join [@GEN_SUB_CONTRACT_D1] b on a.docentry=b.docentry Inner Join [@GEN_SUB_CONTRACT_D0] C On B.DocEntry = C.DocEntry ANd C.u_ItemCode = B.u_Father where  a.DocNum= '" & Trim(objForm.Items.Item("SCNo").Specific.Value) & "' And B.U_Father = '" + Trim(objForm.Items.Item("ItemNo").Specific.value) + "'")
                For i As Integer = 0 To oRS.RecordCount - 1
                    If i > 0 Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRS.Fields.Item("U_Code").Value
                    oRS.MoveNext()
                Next
                If oRS.RecordCount = 0 Then
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "-1"
                End If
                oCFL.SetConditions(oCons)
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function PostStockTransfer(ByVal FormUID As String) As Integer
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_RET_D0")
            Dim oStockTransfer As SAPbobsCOM.StockTransfer = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
            oCompany.StartTransaction()
            oStockTransfer.DocDate = DateTime.ParseExact(objForm.Items.Item("t_docdt").Specific.value, "yyyyMMdd", Nothing)
            'oStockTransfer.UserFields.Fields.Item("U_Type").Value = "SubCont_Return"
            'oStockTransfer.UserFields.Fields.Item("U_DocNum").Value = Trim(objForm.Items.Item("t_docno").Specific.Value)
            For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                If Row > 1 Then oStockTransfer.Lines.Add()
                oStockTransfer.FromWarehouse = Trim(objMatrix.Columns.Item("fwhs").Cells.Item(Row).Specific.Value)
                oStockTransfer.Lines.ItemCode = Trim(objMatrix.Columns.Item("itemno").Cells.Item(Row).Specific.Value)
                oStockTransfer.Lines.WarehouseCode = Trim(objMatrix.Columns.Item("twhs").Cells.Item(Row).Specific.Value)
                oStockTransfer.Lines.Quantity = Trim(objMatrix.Columns.Item("issqty").Cells.Item(Row).Specific.Value)
                oStockTransfer.Lines.SetCurrentLine(Row - 1)
            Next
            If oStockTransfer.Add = 0 Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                Dim DocEntry As String = Trim(oCompany.GetNewObjectKey)
                Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRS.DoQuery("Select DocNum from OWTR Where DocEntry='" & DocEntry & "'")
                oDBs_Head.SetValue("U_InvTrDNo", 0, Trim(oRS.Fields.Item("DocNum").Value))
                Return CInt(DocEntry)
            Else
                oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription & " - [ Error Code : " & oCompany.GetLastErrorCode & " ]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                Return 0
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
            Return 0
        End Try
    End Function


End Class
