Public Class ClsGEN_COST_SHEET

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim objMatrix1 As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail1 As SAPbouiCOM.DBDataSource
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
    Dim PrevDocNo As String
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            oUtilities.SAPXML("GEN_COST_SHEET.xml")
            objForm = oApplication.Forms.GetForm("GEN_COST_SHEET", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET_D0")
            oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET_D1")
            objForm.EnableMenu("1282", False)
            objForm.EnableMenu("1288", True)
            objForm.EnableMenu("1289", True)
            objForm.EnableMenu("1290", True)
            objForm.EnableMenu("1291", True)
            objForm.DataBrowser.BrowseBy = "docnum"
            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("final").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSet.DoQuery("Select USER_CODE From OUSR Where USER_CODE = '" + oCompany.UserName.ToString + "' And IsNull(u_cstsht,'N') = 'Y'")
            If oRSet.RecordCount > 0 Then
                objForm.Items.Item("final").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            End If
            objForm.PaneLevel = "1"
            SetDefault(objForm.UniqueID)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterBP(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("RBPCFL")
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)
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
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "DocStatus"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "O"
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterSOItems(ByVal FormUID As String, ByVal SONO As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("ITCFL")
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Distinct B.ItemCode From ORDR A Inner Join RDR1 B On A.DocEntry = B.DocEntry Where A.DocNum = '" + SONO + "'")
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

    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            objForm.EnableMenu("1282", False)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET_D0")
            oUtilities.GetSeries(FormUID, "series", "GEN_COST_SHEET")
            oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("series").Specific.Selected.Value, "GEN_COST_SHEET"))
            oDBs_Head.SetValue("U_docdate", 0, DateTime.Today.ToString("yyyyMMdd"))
            objMatrix = objForm.Items.Item("mtx1").Specific
            objMatrix.AddRow(1, objMatrix.VisualRowCount)
            Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
            objMatrix.AutoResizeColumns()
            objMatrix1 = objForm.Items.Item("mtx2").Specific
            objMatrix1.AddRow(1, objMatrix1.VisualRowCount)
            Me.SetNewLine1(FormUID, objMatrix1.VisualRowCount, objMatrix1)
            objMatrix1.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET_D0")
            objMatrix = oMatrix
            objMatrix.FlushToDataSource()
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, objMatrix.VisualRowCount)
            oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_itmtype", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_import", oDBs_Detail.Offset, "No")
            oDBs_Detail.SetValue("u_doccur", oDBs_Detail.Offset, "INR")
            oDBs_Detail.SetValue("u_docrate", oDBs_Detail.Offset, "1")
            oDBs_Detail.SetValue("u_qty", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_uom", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_rate", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_rowtotal", oDBs_Detail.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine1(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET_D1")
            objMatrix = oMatrix
            objMatrix.FlushToDataSource()
            oDBs_Detail1.Offset = Row - 1
            oDBs_Detail1.SetValue("LineID", oDBs_Detail1.Offset, objMatrix.VisualRowCount)
            oDBs_Detail1.SetValue("u_prcs", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("u_prcsname", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("u_rate", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("u_rowtotal", oDBs_Detail1.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET_D0")
            oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET_D1")
            If Trim(oDBs_Head.GetValue("u_doccur", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter document currency", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            If Trim(oDBs_Head.GetValue("u_itemcode", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter Style code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            If Trim(oDBs_Head.GetValue("u_docrate", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter document rate", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If oDBs_Head.GetValue("u_mtotal", 0) = 0 Then
                oApplication.StatusBar.SetText("Please enter value for items in rows and click refresh", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            objMatrix = objForm.Items.Item("mtx1").Specific
            If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(1).Specific.value) = "" Then
                oApplication.StatusBar.SetText("Please enter items", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            objMatrix1 = objForm.Items.Item("mtx2").Specific
            If Trim(objMatrix1.Columns.Item("prcs").Cells.Item(1).Specific.value) = "" Then
                oApplication.StatusBar.SetText("Please enter process", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSet.DoQuery("Select DocNum From [@GEN_COST_SHEET] Where u_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' ANd DocNum <> '" + Trim(objForm.Items.Item("docnum").Specific.value) + "' And Isnull(u_final,'N') = 'Y'")
            If oRSet.RecordCount > 0 Then
                oApplication.StatusBar.SetText("Cost sheet is finalized for this style", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
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
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation(FormUID) = False Then BubbleEvent = False
                    ElseIf pVal.ItemUID = "1" And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(FormUID)
                    End If
                    If pVal.ItemUID = "flditem" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 1
                        objForm.Items.Item("flditem").AffectsFormMode = False
                    End If
                    If pVal.ItemUID = "fldexp" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 2
                        objForm.Items.Item("fldexp").AffectsFormMode = False
                    End If
                    If pVal.ItemUID = "rfrsh" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If pVal.BeforeAction = False Then
                            objMatrix = objForm.Items.Item("mtx1").Specific
                            objMatrix1 = objForm.Items.Item("mtx2").Specific
                            Dim MatVal As Double
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                MatVal = MatVal + objMatrix.Columns.Item("rowtotal").Cells.Item(i).Specific.value
                            Next
                            oDBs_Head.SetValue("u_mtotal", 0, MatVal)
                            Dim ExpVal As Double
                            For i As Integer = 1 To objMatrix1.VisualRowCount
                                ExpVal = ExpVal + objMatrix1.Columns.Item("rowtotal").Cells.Item(i).Specific.value
                            Next
                            Dim Efficiency, Cost, TotalMacCost As Double
                            Efficiency = (480 / objForm.Items.Item("sam").Specific.value) * (objForm.Items.Item("effcy").Specific.value / 100)
                            Cost = objForm.Items.Item("maccost").Specific.value / Efficiency
                            TotalMacCost = Cost
                            oDBs_Head.SetValue("u_etotal", 0, ExpVal)
                            oDBs_Head.SetValue("u_ototal", 0, TotalMacCost)
                            oDBs_Head.SetValue("u_total", 0, MatVal + ExpVal + TotalMacCost)
                            Dim PrfVal, WasVal As Double
                            WasVal = (objForm.Items.Item("total").Specific.value * objForm.Items.Item("wasper").Specific.value) / 100
                            oDBs_Head.SetValue("u_wasval", 0, WasVal)
                            PrfVal = ((objForm.Items.Item("total").Specific.value + WasVal) * objForm.Items.Item("prfper").Specific.value) / 100
                            oDBs_Head.SetValue("u_prfval", 0, PrfVal)
                            oDBs_Head.SetValue("u_gtotal", 0, MatVal + ExpVal + PrfVal + WasVal + TotalMacCost)
                            oDBs_Head.SetValue("u_costinr", 0, (MatVal + ExpVal + PrfVal + WasVal + TotalMacCost))
                            oDBs_Head.SetValue("u_costusd", 0, ((MatVal + ExpVal + PrfVal + WasVal)) / objForm.Items.Item("docrate").Specific.value)
                        Else
                            objMatrix = objForm.Items.Item("mtx1").Specific
                            objMatrix1 = objForm.Items.Item("mtx2").Specific
                            If objMatrix.VisualRowCount = 0 Then
                                oApplication.StatusBar.SetText("Please select items with BOM", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            If objForm.Items.Item("qty").Specific.value <= 0 Then
                                oApplication.StatusBar.SetText("Please enter quantity", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            If objForm.Items.Item("sam").Specific.value <= 0 Then
                                oApplication.StatusBar.SetText("Please enter SAM", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            If objForm.Items.Item("effcy").Specific.value <= 0 Then
                                oApplication.StatusBar.SetText("Please enter Efficiency", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            If objForm.Items.Item("maccost").Specific.value <= 0 Then
                                oApplication.StatusBar.SetText("Please enter Machine Cost", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            If objForm.Items.Item("prfper").Specific.value = 0 Then
                                oApplication.StatusBar.SetText("Please enter profit percentage and click refresh again", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                            If objForm.Items.Item("wasper").Specific.value = 0 Then
                                oApplication.StatusBar.SetText("Please enter wastage percentage and click refresh again", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If (pVal.ItemUID = "mtx1" Or pVal.ItemUID = "mtx2") And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                        If Trim(objForm.Items.Item("itemcode").Specific.value) = "" Or objForm.Items.Item("qty").Specific.value = 0 Then
                            oApplication.StatusBar.SetText("Please select style code and quantity", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                    If (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = True Then
                        If pVal.ItemUID <> "final" And pVal.ItemUID <> "2" And pVal.ItemUID <> "1" And pVal.ItemUID <> "flditem" And pVal.ItemUID <> "fldexp" Then
                            BubbleEvent = False
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    If pVal.CharPressed <> 13 And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = True Then
                        BubbleEvent = False
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If pVal.ItemUID = "mtx2" And pVal.ColUID = "rate" And pVal.BeforeAction = False Then
                        objMatrix1 = objForm.Items.Item("mtx2").Specific
                        objMatrix1.Columns.Item("rowtotal").Cells.Item(pVal.Row).Specific.value = objMatrix1.Columns.Item("rate").Cells.Item(pVal.Row).Specific.value
                    End If
                    If pVal.ItemUID = "mtx1" And pVal.ColUID = "doccur" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = False Then
                        objMatrix = objForm.Items.Item("mtx1").Specific
                        If Trim(objMatrix.Columns.Item("doccur").Cells.Item(pVal.Row).Specific.value) <> "" Then
                            If Trim(objMatrix.Columns.Item("doccur").Cells.Item(pVal.Row).Specific.value) = "INR" Or Trim(objMatrix.Columns.Item("doccur").Cells.Item(pVal.Row).Specific.value) = "inr" Then
                                objMatrix.Columns.Item("docrate").Cells.Item(pVal.Row).Specific.value = 1
                            Else
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet.DoQuery("Select Rate From ORTT Where RateDate = '" + Trim(objForm.Items.Item("docdate").Specific.value) + "' And Currency = '" + Trim(objMatrix.Columns.Item("doccur").Cells.Item(pVal.Row).Specific.value) + "'")
                                If oRSet.RecordCount = 0 Then
                                    oApplication.StatusBar.SetText("Please enter exchange rate for this currency for the date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                End If
                                objMatrix.Columns.Item("docrate").Cells.Item(pVal.Row).Specific.value = CDbl(oRSet.Fields.Item("Rate").Value)
                            End If
                        End If
                    End If
                    If pVal.ItemUID = "doccur" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = False Then
                        If Trim(objForm.Items.Item("doccur").Specific.value) <> "" Then
                            If Trim(objForm.Items.Item("doccur").Specific.value) = "INR" Or Trim(objForm.Items.Item("doccur").Specific.value) = "inr" Then
                                objForm.Items.Item("docrate").Specific.value = 1
                            Else
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet.DoQuery("Select Rate From ORTT Where RateDate = '" + Trim(objForm.Items.Item("docdate").Specific.value) + "' And Currency = '" + Trim(objForm.Items.Item("doccur").Specific.value) + "'")
                                If oRSet.RecordCount = 0 Then
                                    oApplication.StatusBar.SetText("Please enter exchange rate for this currency for the date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                End If
                                objForm.Items.Item("docrate").Specific.value = CDbl(oRSet.Fields.Item("Rate").Value)
                            End If
                        End If
                    End If
                    If pVal.ItemUID = "mtx1" And (pVal.ColUID = "rate" Or pVal.ColUID = "qty") And pVal.BeforeAction = False Then
                        objMatrix = objForm.Items.Item("mtx1").Specific
                        objMatrix.Columns.Item("rowtotal").Cells.Item(pVal.Row).Specific.value = objMatrix.Columns.Item("rate").Cells.Item(pVal.Row).Specific.value * objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value
                    End If
                    If pVal.ItemUID = "wasper" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = False Then
                        oDBs_Head.SetValue("u_wasval", 0, (objForm.Items.Item("total").Specific.value * objForm.Items.Item("wasper").Specific.value) / 100)
                    End If
                    If pVal.ItemUID = "prfper" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = False Then
                        oDBs_Head.SetValue("u_prfval", 0, ((CDbl(objForm.Items.Item("total").Specific.value) + CDbl(objForm.Items.Item("wasval").Specific.value)) * objForm.Items.Item("prfper").Specific.value) / 100)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If pVal.ItemUID = "itemcode" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRSet.DoQuery("Select DocNum From [@GEN_COST_SHEET] Where u_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "'")
                        If oRSet.RecordCount > 0 Then
                            oApplication.StatusBar.SetText("Cost Sheet already defined for this item, please go to the document and duplicate it", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                        If oCFL.UniqueID = "RBPCFL" Then
                            Me.FilterBP(FormUID)
                        End If
                        If oCFL.UniqueID = "SOCFL" Then
                            Me.FilterSO(FormUID)
                        End If
                        If oCFL.UniqueID = "ITCFL" Then
                            If Trim(objForm.Items.Item("sono").Specific.value) = "" Then
                                oApplication.StatusBar.SetText("Please select sales order", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            Me.FilterSOItems(FormUID, objForm.Items.Item("sono").Specific.value)
                        End If
                    Else
                        If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET")
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET_D0")
                            oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET_D1")
                            If oCFL.UniqueID = "SOCFL" Then
                                oDBs_Head.SetValue("u_sono", 0, oDT.GetValue("DocNum", 0))
                                oDBs_Head.SetValue("u_soref", 0, oDT.GetValue("NumAtCard", 0))
                                oDBs_Head.SetValue("u_cardcode", 0, oDT.GetValue("CardCode", 0))
                                oDBs_Head.SetValue("u_cardname", 0, oDT.GetValue("CardName", 0))
                                oDBs_Head.SetValue("u_doccur", 0, oDT.GetValue("DocCur", 0))
                                oDBs_Head.SetValue("u_docrate", 0, oDT.GetValue("DocRate", 0))
                                oDBs_Head.SetValue("u_itemcode", 0, "")
                                oDBs_Head.SetValue("u_itemname", 0, "")
                                oDBs_Head.SetValue("u_qty", 0, "")
                                objMatrix = objForm.Items.Item("mtx1").Specific
                                objMatrix.Clear()
                            End If
                            If oCFL.UniqueID = "ITCFL" Then
                                objMatrix = objForm.Items.Item("mtx1").Specific
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet.DoQuery("Select DocNum From [@GEN_COST_SHEET] Where u_sono = '" + Trim(objForm.Items.Item("sono").Specific.value) + "' And u_itemcode = '" + Trim(oDT.GetValue("ItemCode", 0)) + "'")
                                If oRSet.RecordCount > 0 Then
                                    oApplication.StatusBar.SetText("Cost Sheet already entered for this sales order and style combination, please duplicate from existing cost sheet", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                oDBs_Head.SetValue("u_itemcode", 0, oDT.GetValue("ItemCode", 0))
                                oDBs_Head.SetValue("u_itemname", 0, oDT.GetValue("ItemName", 0))
                                oRSet.DoQuery("Select Sum(B.Quantity) AS 'Quantity' From ORDR A Inner Join RDR1 B On A.DocEntry  = B.DocEntry Where A.DocNum = '" + Trim(objForm.Items.Item("sono").Specific.value) + "' And B.ItemCode = '" + Trim(oDT.GetValue("ItemCode", 0)) + "'")
                                oDBs_Head.SetValue("u_qty", 0, oRSet.Fields.Item("Quantity").Value)
                                Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("Select Distinct B.LineId,B.u_itemcode,B.u_itemname,B.u_qty,B.u_uom,C.AvgPrice,C.OnHand,C.CardCode From [@GEN_CUST_BOM] A INNER JOIN [@GEN_CUST_BOM_D0] B ON A.DocEntry = B.DocEntry Inner Join OITM C on B.u_itemcode = C.ItemCode Where A.u_sono = '" + Trim(objForm.Items.Item("sono").Specific.value) + "' And A.u_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' Order By B.LineId")
                                objMatrix.Clear()
                                For i As Integer = 1 To oRecordSet.RecordCount
                                    objMatrix.AddRow(1, objMatrix.VisualRowCount)
                                    Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
                                    oDBs_Detail.SetValue("u_itemcode", i - 1, oRecordSet.Fields.Item("u_itemcode").Value)
                                    oDBs_Detail.SetValue("u_itemname", i - 1, oRecordSet.Fields.Item("u_itemname").Value)
                                    oDBs_Detail.SetValue("u_qty", i - 1, oRecordSet.Fields.Item("u_qty").Value * objForm.Items.Item("qty").Specific.value)
                                    oDBs_Detail.SetValue("u_uom", i - 1, oRecordSet.Fields.Item("u_uom").Value)
                                    oDBs_Detail.SetValue("u_rate", i - 1, oRecordSet.Fields.Item("AvgPrice").Value)
                                    oDBs_Detail.SetValue("u_onhand", i - 1, oRecordSet.Fields.Item("OnHand").Value)
                                    oDBs_Detail.SetValue("u_prefvend", i - 1, oRecordSet.Fields.Item("CardCode").Value)
                                    objMatrix.SetLineData(i)
                                    oRecordSet.MoveNext()
                                Next
                            End If
                            If oCFL.UniqueID = "RBPCFL" Then
                                objMatrix = objForm.Items.Item("mtx1").Specific
                                For i As Integer = 0 To oDT.Rows.Count - 1
                                    Dim cflSelectedcount As Integer = oDT.Rows.Count
                                    oDBs_Detail.Offset = pVal.Row - 1 + i
                                    oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                                    oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, objMatrix.Columns.Item("itemname").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Detail.SetValue("u_qty", oDBs_Detail.Offset, objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Detail.SetValue("u_uom", oDBs_Detail.Offset, objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Detail.SetValue("u_rate", oDBs_Detail.Offset, objMatrix.Columns.Item("rate").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Detail.SetValue("u_rowtotal", oDBs_Detail.Offset, objMatrix.Columns.Item("rowtotal").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Detail.SetValue("u_prefvend", oDBs_Detail.Offset, oDT.GetValue("CardCode", i))
                                    objMatrix.SetLineData(pVal.Row + i)
                                Next
                            End If
                            If oCFL.UniqueID = "RITCFL" Then
                                objMatrix = objForm.Items.Item("mtx1").Specific
                                oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim Total As Double = 0
                                oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET_D0")
                                Dim Flag As Boolean = False
                                Dim errflag As Boolean = False
                                If objMatrix.VisualRowCount = 1 Or pVal.Row = objMatrix.VisualRowCount Then
                                    Flag = True
                                End If
                                For i As Integer = 0 To oDT.Rows.Count - 1
                                    Dim cflSelectedcount As Integer = oDT.Rows.Count
                                    If i < cflSelectedcount - 1 Then
                                        objMatrix.AddRow(1, pVal.Row)
                                        oDBs_Detail.InsertRecord(pVal.RoSw + i - 1)
                                    End If
                                    oDBs_Detail.Offset = pVal.Row - 1 + i
                                    oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                                    oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, oDT.GetValue("ItmsGrpCod", i))
                                    oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, oDT.GetValue("ItmsGrpNam", i))
                                    objMatrix.SetLineData(pVal.Row + i)
                                    objForm.EnableMenu("1293", True)
                                Next
                                If Flag = True Then
                                    objMatrix.AddRow(1, objMatrix.VisualRowCount)
                                    Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
                                End If
                            End If
                            If oCFL.UniqueID = "PRCSCFL" Then
                                objMatrix1 = objForm.Items.Item("mtx2").Specific
                                oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim Total As Double = 0
                                oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET_D1")
                                Dim Flag As Boolean = False
                                Dim errflag As Boolean = False
                                If objMatrix1.VisualRowCount = 1 Or pVal.Row = objMatrix1.VisualRowCount Then
                                    Flag = True
                                End If
                                For i As Integer = 0 To oDT.Rows.Count - 1
                                    Dim cflSelectedcount As Integer = oDT.Rows.Count
                                    If i < cflSelectedcount - 1 Then
                                        objMatrix1.AddRow(1, pVal.Row)
                                        oDBs_Detail1.InsertRecord(pVal.Row + i - 1)
                                    End If
                                    oDBs_Detail1.Offset = pVal.Row - 1 + i
                                    oDBs_Detail1.SetValue("LineID", oDBs_Detail1.Offset, i + pVal.Row)
                                    oDBs_Detail1.SetValue("u_prcs", oDBs_Detail1.Offset, oDT.GetValue("Code", i))
                                    oDBs_Detail1.SetValue("u_prcsname", oDBs_Detail1.Offset, oDT.GetValue("Name", i))
                                    objMatrix1.SetLineData(pVal.Row + i)
                                    objForm.EnableMenu("1293", True)
                                Next
                                If Flag = True Then
                                    objMatrix1.AddRow(1, objMatrix1.VisualRowCount)
                                    Me.SetNewLine1(FormUID, objMatrix1.VisualRowCount, objMatrix1)
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
            If pVal.BeforeAction = True Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "1287"
                        If objForm.TypeEx = "GEN_COST_SHEET" Then
                            PrevDocNo = objForm.Items.Item("docnum").Specific.value
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select DocNum From [@GEN_COST_SHEET] Where u_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' And Isnull(u_final,'N') = 'Y'")
                            If oRSet.RecordCount > 0 Then
                                oApplication.StatusBar.SetText("Cost Sheet already finalized for this style combination, you cannot duplicate anymore", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            oRSet.DoQuery("Select USER_CODE From OUSR Where USER_CODE = '" + oCompany.UserName.ToString + "' And IsNull(u_cstsht,'N') = 'Y'")
                            If oRSet.RecordCount > 0 Then
                                objForm.Items.Item("final").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            End If
                        End If
                End Select
            End If
            If pVal.BeforeAction = False Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "GEN_COST_SHEET"
                        Me.CreateForm(objForm.UniqueID)
                    Case "1287"
                        If objForm.TypeEx = "GEN_COST_SHEET" Then
                            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("series").Specific.Selected.Value, "GEN_COST_SHEET"))
                            objMatrix = objForm.Items.Item("mtx1").Specific
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Exec SEPL_Cost_Sheet '" + PrevDocNo + "','" + Trim(objForm.Items.Item("sono").Specific.value) + "','" + Trim(objForm.Items.Item("itemcode").Specific.value) + "'")
                            objMatrix.Clear()
                            For i As Integer = 1 To oRSet.RecordCount
                                objMatrix.AddRow(1, objMatrix.VisualRowCount)
                                Me.SetNewLine(objForm.UniqueID, objMatrix.VisualRowCount, objMatrix)
                                oDBs_Detail.SetValue("u_itemcode", i - 1, oRSet.Fields.Item("ItemCode").Value)
                                oDBs_Detail.SetValue("u_itemname", i - 1, oRSet.Fields.Item("ItemName").Value)
                                oDBs_Detail.SetValue("u_qty", i - 1, oRSet.Fields.Item("Qty").Value * objForm.Items.Item("qty").Specific.value)
                                oDBs_Detail.SetValue("u_uom", i - 1, oRSet.Fields.Item("UOM").Value)
                                oDBs_Detail.SetValue("u_rate", i - 1, oRSet.Fields.Item("Rate").Value)
                                oDBs_Detail.SetValue("u_rowtotal", i - 1, oRSet.Fields.Item("Rate").Value * oRSet.Fields.Item("Qty").Value * objForm.Items.Item("qty").Specific.value)
                                oDBs_Detail.SetValue("u_onhand", i - 1, oRSet.Fields.Item("OnHand").Value)
                                oDBs_Detail.SetValue("u_prefvend", i - 1, oRSet.Fields.Item("PrefVend").Value)
                                objMatrix.SetLineData(i)
                                oRSet.MoveNext()
                            Next
                        End If
                    Case "1282"
                        If objForm.TypeEx = "GEN_COST_SHEET" Then
                            Me.SetDefault(objForm.UniqueID)
                            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select USER_CODE From OUSR Where USER_CODE = '" + oCompany.UserName.ToString + "' And IsNull(u_cstsht,'N') = 'Y'")
                            If oRSet.RecordCount > 0 Then
                                objForm.Items.Item("final").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            End If
                        End If
                    Case "1281"
                        If objForm.TypeEx = "GEN_COST_SHEET" Then
                            objForm.EnableMenu("1282", True)
                            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        End If
                    Case "1288", "1289", "1290", "1291"
                        If objForm.TypeEx = "GEN_COST_SHEET" Then
                            objForm.EnableMenu("1282", True)
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select USER_CODE From OUSR Where USER_CODE = '" + oCompany.UserName.ToString + "' And IsNull(u_cstsht,'N') = 'Y'")
                            If oRSet.RecordCount > 0 Then
                                objForm.Items.Item("final").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            Else
                                objForm.Items.Item("final").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            End If
                            objForm.EnableMenu("1287", True)
                        End If
                    Case "1293"
                        If objForm.TypeEx = "GEN_COST_SHEET" Then
                            If ITEM_ID.Equals("mtx1") = True Then
                                Dim Total As Double
                                objMatrix = objForm.Items.Item("mtx1").Specific
                                oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET_D0")
                                For Row As Integer = 1 To objMatrix.VisualRowCount
                                    objMatrix.GetLineData(Row)
                                    oDBs_Detail.Offset = Row - 1
                                    oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, Row)
                                    oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, objMatrix.Columns.Item("itemcode").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, objMatrix.Columns.Item("itemname").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_qty", oDBs_Detail.Offset, objMatrix.Columns.Item("qty").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_uom", oDBs_Detail.Offset, objMatrix.Columns.Item("uom").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_rate", oDBs_Detail.Offset, objMatrix.Columns.Item("rate").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_rowtotal", oDBs_Detail.Offset, objMatrix.Columns.Item("rowtotal").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_onhand", oDBs_Detail.Offset, objMatrix.Columns.Item("onhand").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_prefvend", oDBs_Detail.Offset, objMatrix.Columns.Item("prefvend").Cells.Item(Row).Specific.value)
                                    objMatrix.SetLineData(Row)
                                Next
                                objMatrix.FlushToDataSource()
                                oDBs_Detail.RemoveRecord(oDBs_Detail.Size - 1)
                                objMatrix.LoadFromDataSource()
                            End If
                            If ITEM_ID.Equals("mtx2") = True Then
                                Dim Total As Double
                                objMatrix1 = objForm.Items.Item("mtx2").Specific
                                oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET_D1")
                                For Row As Integer = 1 To objMatrix1.VisualRowCount
                                    objMatrix1.GetLineData(Row)
                                    oDBs_Detail1.Offset = Row - 1
                                    oDBs_Detail1.SetValue("LineID", oDBs_Detail1.Offset, Row)
                                    oDBs_Detail1.SetValue("u_pcode", oDBs_Detail1.Offset, objMatrix1.Columns.Item("pcode").Cells.Item(Row).Specific.value)
                                    oDBs_Detail1.SetValue("u_pname", oDBs_Detail1.Offset, objMatrix1.Columns.Item("pname").Cells.Item(Row).Specific.value)
                                    oDBs_Detail1.SetValue("u_pc", oDBs_Detail1.Offset, objMatrix1.Columns.Item("pc").Cells.Item(Row).Specific.value)
                                    oDBs_Detail1.SetValue("u_ldc", oDBs_Detail1.Offset, objMatrix1.Columns.Item("ldc").Cells.Item(Row).Specific.value)
                                    oDBs_Detail1.SetValue("u_ac", oDBs_Detail1.Offset, objMatrix1.Columns.Item("ac").Cells.Item(Row).Specific.value)
                                    oDBs_Detail1.SetValue("u_rc", oDBs_Detail1.Offset, objMatrix1.Columns.Item("rc").Cells.Item(Row).Specific.value)
                                    objMatrix1.SetLineData(Row)
                                Next
                                objMatrix1.FlushToDataSource()
                                oDBs_Detail1.RemoveRecord(oDBs_Detail1.Size - 1)
                                objMatrix1.LoadFromDataSource()
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
                objMatrix = oForm.Items.Item("mtx1").Specific
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
                        objMatrix1 = objForm.Items.Item("mtx2").Specific
                        If objMatrix1.VisualRowCount <> 0 Then
                            objMatrix1.DeleteRow(objMatrix1.VisualRowCount)
                            objMatrix1.FlushToDataSource()
                        End If
                    ElseIf BusinessObjectInfo.ActionSuccess = True Then
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Then
                            objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                            objMatrix1 = objForm.Items.Item("mtx2").Specific
                            objMatrix1.AddRow()
                            objMatrix1.FlushToDataSource()
                            Me.SetNewLine1(BusinessObjectInfo.FormUID, objMatrix1.VisualRowCount, objMatrix1)
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        objMatrix1 = objForm.Items.Item("mtx2").Specific
                        objMatrix1.AddRow()
                        objMatrix1.FlushToDataSource()
                        Me.SetNewLine1(BusinessObjectInfo.FormUID, objMatrix1.VisualRowCount, objMatrix1)
                        objForm.Items.Item("mtx1").AffectsFormMode = False
                        objForm.Items.Item("mtx2").AffectsFormMode = False
                        objForm.Items.Item("fldexp").AffectsFormMode = False
                        objForm.Items.Item("flditem").AffectsFormMode = False
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRSet.DoQuery("Select USER_CODE From OUSR Where USER_CODE = '" + oCompany.UserName.ToString + "' And IsNull(u_cstsht,'N') = 'Y'")
                        If oRSet.RecordCount > 0 Then
                            objForm.Items.Item("final").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        End If
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
