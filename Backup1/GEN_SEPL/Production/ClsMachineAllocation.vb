Public Class ClsMachineAllocation

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim objSubForm As SAPbouiCOM.Form
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim objSubMatrix As SAPbouiCOM.Matrix
    Dim objSubMatrix1 As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oSDBs_Head As SAPbouiCOM.DBDataSource
    Dim oSDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oSDBs_Detail1 As SAPbouiCOM.DBDataSource
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
    Dim UpdMode As Boolean = False
    Dim DocStatus As String
    Dim ModalForm As Boolean = False
    Dim PARENT_FORM As String
    Dim BaseRow As Integer
    Dim TrgtCode As String
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            oUtilities.SAPXML("GEN_MACH_ALLOC.xml")
            objForm = oApplication.Forms.GetForm("GEN_MACH_ALLOC", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MACH_ALLOC")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_MACH_ALLOC_D0")
            objForm.DataBrowser.BrowseBy = "code"
            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            objForm.Items.Item("name").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objMatrix = objForm.Items.Item("mtx").Specific
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_MACH_ALLOC_D0")
            objMatrix = oMatrix
            objMatrix.FlushToDataSource()
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, objMatrix.VisualRowCount)
            oDBs_Detail.SetValue("u_sono", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_qty", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_deldate", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_stdate", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_eddate", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_unit", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_nom", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_lineno", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_trgtcode", oDBs_Detail.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try
            objMatrix = objForm.Items.Item("mtx").Specific
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MACH_ALLOC")
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If Trim(objMatrix.Columns.Item("sono").Cells.Item(1).Specific.value) = "" Then
                oApplication.StatusBar.SetText("Please enter rows", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Dim CHILD_FORM As String = "GEN_CAP_PLAN@" & pVal.FormUID
            Dim ModalForm As Boolean = False
            For i As Integer = 0 To oApplication.Forms.Count - 1
                If oApplication.Forms.Item(i).UniqueID = (CHILD_FORM) Then
                    objSubForm = oApplication.Forms.Item("GEN_CAP_PLAN@" & pVal.FormUID)
                    ModalForm = True
                    Exit For
                End If
            Next
            If ModalForm = False Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                        objForm = oApplication.Forms.Item(pVal.FormUID)
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            If Me.Validation(FormUID) = False Then BubbleEvent = False
                        ElseIf pVal.ItemUID = "1" And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select Count(*) + 1 AS 'Count' From [@GEN_MACH_ALLOC]")
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MACH_ALLOC")
                            oDBs_Head.SetValue("Code", 0, oRSet.Fields.Item("Count").Value)
                            objMatrix = objForm.Items.Item("mtx").Specific
                            objMatrix.AddRow()
                            Me.SetNewLine(objForm.UniqueID, objMatrix.VisualRowCount, objMatrix)
                        End If
                        If pVal.ItemUID = "2" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select Code From [@GEN_CAP_PLAN] Where u_basecode = '" + Trim(objForm.Items.Item("code").Specific.value) + "'")
                            Dim DelCode As String
                            DelCode = oRSet.Fields.Item("Code").Value
                            oRSet.DoQuery("Delete From [@GEN_CAP_PLAN_D1] Where Code = '" + DelCode + "'")
                            oRSet.DoQuery("Delete From [@GEN_CAP_PLAN_D0] Where Code = '" + DelCode + "'")
                            oRSet.DoQuery("Delete From [@GEN_CAP_PLAN] Where Code = '" + DelCode + "'")
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select Code From [@GEN_CAP_PLAN] Where u_basecode = '" + Trim(objForm.Items.Item("code").Specific.value) + "'")
                            Dim DelCode As String
                            DelCode = oRSet.Fields.Item("Code").Value
                            oRSet.DoQuery("Delete From [@GEN_CAP_PLAN_D1] Where Code = '" + DelCode + "'")
                            oRSet.DoQuery("Delete From [@GEN_CAP_PLAN_D0] Where Code = '" + DelCode + "'")
                            oRSet.DoQuery("Delete From [@GEN_CAP_PLAN] Where Code = '" + DelCode + "'")
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        objMatrix = objForm.Items.Item("mtx").Specific
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MACH_ALLOC")
                        If pVal.ItemUID = "mtx" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            If Trim(oDBs_Head.GetValue("u_manual", 0)) <> "Y" Then
                                BubbleEvent = False
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        objMatrix = objForm.Items.Item("mtx").Specific
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MACH_ALLOC")
                        If pVal.ItemUID = "mtx" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            If Trim(oDBs_Head.GetValue("u_manual", 0)) <> "Y" Then
                                BubbleEvent = False
                            End If
                        End If
                        If pVal.ItemUID = "mtx" And pVal.ColUID = "trgtcode" And pVal.BeforeAction = True And pVal.CharPressed <> 9 And pVal.CharPressed <> 13 Then
                            BubbleEvent = False
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                        If pVal.ItemUID = "mtx" And pVal.Row > 0 And pVal.ColUID = "lineid" Then
                            objMatrix = objForm.Items.Item("mtx").Specific
                            If pVal.BeforeAction = True Then
                                If Trim(objMatrix.Columns.Item("sono").Cells.Item(pVal.Row).Specific.value) = "" Or Trim(objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.value) = "" Or Trim(objMatrix.Columns.Item("unit").Cells.Item(pVal.Row).Specific.value) = "" Or Trim(objMatrix.Columns.Item("lineno").Cells.Item(pVal.Row).Specific.value) = "" Or Trim(objMatrix.Columns.Item("nom").Cells.Item(pVal.Row).Specific.value) = "" Or Trim(objMatrix.Columns.Item("stdate").Cells.Item(pVal.Row).Specific.value) = "" Or Trim(objMatrix.Columns.Item("eddate").Cells.Item(pVal.Row).Specific.value) = "" Then
                                    oApplication.StatusBar.SetText("Please enter all details", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                            Else
                                Me.CreateFormFind(FormUID, objForm.Items.Item("code").Specific.value, objMatrix.Columns.Item("trgtcode").Cells.Item(pVal.Row).Specific.value, pVal.Row, objMatrix.Columns.Item("sono").Cells.Item(pVal.Row).Specific.value, objMatrix.Columns.Item("lineid").Cells.Item(pVal.Row).Specific.value, objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.value, objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value, objMatrix.Columns.Item("unit").Cells.Item(pVal.Row).Specific.value, objMatrix.Columns.Item("lineno").Cells.Item(pVal.Row).Specific.value, objMatrix.Columns.Item("nom").Cells.Item(pVal.Row).Specific.value, objMatrix.Columns.Item("stdate").Cells.Item(pVal.Row).Specific.Value, objMatrix.Columns.Item("eddate").Cells.Item(pVal.Row).Specific.value)
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
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MACH_ALLOC")
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_MACH_ALLOC_D0")
                            If oCFL.UniqueID = "SOCFL" Then
                                Me.FilterSO(FormUID)
                            End If
                            If oCFL.UniqueID = "ITCFL" Then
                                If Trim(objMatrix.Columns.Item("sono").Cells.Item(pVal.Row).Specific.value) <> "" Then
                                    Me.FilterSOItems(FormUID, Trim(objMatrix.Columns.Item("sono").Cells.Item(pVal.Row).Specific.value))
                                Else
                                    oApplication.StatusBar.SetText("Please select Sales Order No", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                End If
                            End If
                            If oCFL.UniqueID = "UNTCFL" Then
                                If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.value) = "" Then
                                    oApplication.StatusBar.SetText("Please select Style", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                End If
                            End If
                            If oCFL.UniqueID = "LNCFL" Then
                                If Trim(objMatrix.Columns.Item("unit").Cells.Item(pVal.Row).Specific.value) = "" Then
                                    oApplication.StatusBar.SetText("Please enter Unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                End If
                            End If
                            If oCFL.UniqueID = "LNTPCFL" Then
                                If Trim(objMatrix.Columns.Item("lineno").Cells.Item(pVal.Row).Specific.value) = "" Then
                                    oApplication.StatusBar.SetText("Please select line", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                End If
                            End If
                        End If
                        If pVal.BeforeAction = False Then
                            If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MACH_ALLOC")
                                oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_MACH_ALLOC_D0")
                                If oCFL.UniqueID = "SOCFL" Then
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
                                        oDBs_Detail.SetValue("u_sono", oDBs_Detail.Offset, oDT.GetValue("DocNum", i))
                                        objMatrix.SetLineData(pVal.Row + i)
                                        objForm.EnableMenu("1293", True)
                                    Next
                                    If Flag = True Then
                                        objMatrix.AddRow(1, objMatrix.VisualRowCount)
                                        Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
                                    End If
                                End If
                                If oCFL.UniqueID = "ITCFL" Then
                                    objMatrix = objForm.Items.Item("mtx").Specific
                                    oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    For i As Integer = 0 To oDT.Rows.Count - 1
                                        Dim cflSelectedcount As Integer = oDT.Rows.Count
                                        oDBs_Detail.Offset = pVal.Row - 1 + i
                                        oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                                        oDBs_Detail.SetValue("u_sono", oDBs_Detail.Offset, objMatrix.Columns.Item("sono").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, oDT.GetValue("ItemCode", i))
                                        oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, oDT.GetValue("ItemName", i))
                                        oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRSet.DoQuery("Select Sum(B.Quantity) AS 'Quantity' From ORDR A Inner Join RDR1 B On A.DocEntry = B.DocEntry Where A.DocNum = '" + Trim(objMatrix.Columns.Item("sono").Cells.Item(pVal.Row).Specific.Value) + "' And B.ItemCode = '" + Trim(oDT.GetValue("ItemCode", i)) + "'")
                                        oDBs_Detail.SetValue("u_qty", oDBs_Detail.Offset, oRSet.Fields.Item("Quantity").Value)
                                        oRSet.DoQuery("Select B.ShipDate AS 'DelDate' From ORDR A Inner Join RDR1 B On A.DocEntry = B.DocEntry Where A.DocNum = '" + Trim(objMatrix.Columns.Item("sono").Cells.Item(pVal.Row).Specific.Value) + "' And B.ItemCode = '" + Trim(oDT.GetValue("ItemCode", i)) + "' Order By B.ShipDate Desc")
                                        oDBs_Detail.SetValue("u_deldate", oDBs_Detail.Offset, Format(oRSet.Fields.Item("DelDate").Value, "yyyyMMdd"))
                                        oDBs_Detail.SetValue("u_stdate", oDBs_Detail.Offset, "")
                                        oDBs_Detail.SetValue("u_eddate", oDBs_Detail.Offset, "")
                                        oDBs_Detail.SetValue("u_unit", oDBs_Detail.Offset, "")
                                        oDBs_Detail.SetValue("u_nom", oDBs_Detail.Offset, "")
                                        oDBs_Detail.SetValue("u_lineno", oDBs_Detail.Offset, "")
                                        oDBs_Detail.SetValue("u_trgtcode", oDBs_Detail.Offset, "")
                                        objMatrix.SetLineData(pVal.Row + i)
                                    Next
                                End If
                                If oCFL.UniqueID = "UNTCFL" Then
                                    objMatrix = objForm.Items.Item("mtx").Specific
                                    oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    For i As Integer = 0 To oDT.Rows.Count - 1
                                        Dim cflSelectedcount As Integer = oDT.Rows.Count
                                        oDBs_Detail.Offset = pVal.Row - 1 + i
                                        oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                                        oDBs_Detail.SetValue("u_sono", oDBs_Detail.Offset, objMatrix.Columns.Item("sono").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, objMatrix.Columns.Item("itemname").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_qty", oDBs_Detail.Offset, objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_deldate", oDBs_Detail.Offset, objMatrix.Columns.Item("deldate").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_stdate", oDBs_Detail.Offset, objMatrix.Columns.Item("stdate").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_eddate", oDBs_Detail.Offset, objMatrix.Columns.Item("eddate").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_unit", oDBs_Detail.Offset, oDT.GetValue("Name", i))
                                        oDBs_Detail.SetValue("u_nom", oDBs_Detail.Offset, objMatrix.Columns.Item("nom").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_lineno", oDBs_Detail.Offset, objMatrix.Columns.Item("lineno").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_trgtcode", oDBs_Detail.Offset, objMatrix.Columns.Item("trgtcode").Cells.Item(pVal.Row).Specific.value)
                                        objMatrix.SetLineData(pVal.Row + i)
                                    Next
                                End If
                                If oCFL.UniqueID = "LNCFL" Then
                                    objMatrix = objForm.Items.Item("mtx").Specific
                                    oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    For i As Integer = 0 To oDT.Rows.Count - 1
                                        Dim cflSelectedcount As Integer = oDT.Rows.Count
                                        oDBs_Detail.Offset = pVal.Row - 1 + i
                                        oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                                        oDBs_Detail.SetValue("u_sono", oDBs_Detail.Offset, objMatrix.Columns.Item("sono").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, objMatrix.Columns.Item("itemname").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_qty", oDBs_Detail.Offset, objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_deldate", oDBs_Detail.Offset, objMatrix.Columns.Item("deldate").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_stdate", oDBs_Detail.Offset, objMatrix.Columns.Item("stdate").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_eddate", oDBs_Detail.Offset, objMatrix.Columns.Item("eddate").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_unit", oDBs_Detail.Offset, objMatrix.Columns.Item("unit").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_nom", oDBs_Detail.Offset, objMatrix.Columns.Item("nom").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_lineno", oDBs_Detail.Offset, oDT.GetValue("Name", i))
                                        oDBs_Detail.SetValue("u_trgtcode", oDBs_Detail.Offset, objMatrix.Columns.Item("trgtcode").Cells.Item(pVal.Row).Specific.value)
                                        objMatrix.SetLineData(pVal.Row + i)
                                    Next
                                End If
                                If oCFL.UniqueID = "LNTPCFL" Then
                                    objMatrix = objForm.Items.Item("mtx").Specific
                                    oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    For i As Integer = 0 To oDT.Rows.Count - 1
                                        Dim cflSelectedcount As Integer = oDT.Rows.Count
                                        oDBs_Detail.Offset = pVal.Row - 1 + i
                                        oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                                        oDBs_Detail.SetValue("u_sono", oDBs_Detail.Offset, objMatrix.Columns.Item("sono").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, objMatrix.Columns.Item("itemname").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_qty", oDBs_Detail.Offset, objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_deldate", oDBs_Detail.Offset, objMatrix.Columns.Item("deldate").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_stdate", oDBs_Detail.Offset, objMatrix.Columns.Item("stdate").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_eddate", oDBs_Detail.Offset, objMatrix.Columns.Item("eddate").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_unit", oDBs_Detail.Offset, objMatrix.Columns.Item("unit").Cells.Item(pVal.Row).Specific.value)
                                        oDBs_Detail.SetValue("u_lineno", oDBs_Detail.Offset, objMatrix.Columns.Item("lineno").Cells.Item(pVal.Row).Specific.value)
                                        oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRSet.DoQuery("Select B.u_nom From [@GEN_MACH_POOL] A Inner Join [@GEN_MACH_POOL_D0] B  On A.Code = B.Code Where A.Name = '" + Trim(objMatrix.Columns.Item("unit").Cells.Item(pVal.Row).Specific.value) + "' And B.u_line = '" + Trim(objMatrix.Columns.Item("lineno").Cells.Item(pVal.Row).Specific.value) + "' And B.u_type = '" + Trim(oDT.GetValue("Code", i)) + "'")
                                        oDBs_Detail.SetValue("u_nom", oDBs_Detail.Offset, oRSet.Fields.Item("u_nom").Value)
                                        oDBs_Detail.SetValue("u_trgtcode", oDBs_Detail.Offset, objMatrix.Columns.Item("trgtcode").Cells.Item(pVal.Row).Specific.value)
                                        objMatrix.SetLineData(pVal.Row + i)
                                    Next
                                End If
                            End If
                        End If
                End Select
            ElseIf pVal.BeforeAction = True And ModalForm = True And pVal.FormUID = (objSubForm.UniqueID.Substring(objSubForm.UniqueID.IndexOf("@") + 1)) Then
                objSubForm = oApplication.Forms.Item("GEN_CAP_PLAN@" & pVal.FormUID)
                objSubForm.Select()
                BubbleEvent = False
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ChildForm_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objSubForm = oApplication.Forms.Item(pVal.FormUID)
                    PARENT_FORM = (pVal.FormUID.Substring(objSubForm.UniqueID.IndexOf("@") + 1))
                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    If pVal.ItemUID = "code" And pVal.CharPressed <> 9 And pVal.CharPressed <> 13 And pVal.BeforeAction = True Then
                        BubbleEvent = False
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "flddate" And pVal.BeforeAction = False Then
                        objSubForm.PaneLevel = 1
                    End If
                    If pVal.ItemUID = "fldmach" And pVal.BeforeAction = False Then
                        objSubForm.PaneLevel = 2
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objSubForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    If pVal.BeforeAction = True Then
                        If oCFL.UniqueID = "LNTPCFL" Then
                            Me.FilterTypeCode(FormUID, Trim(objSubForm.Items.Item("unit").Specific.value), Trim(objSubForm.Items.Item("stdate").Specific.value), Trim(objSubForm.Items.Item("eddate").Specific.value))
                        End If
                    End If
                    If pVal.BeforeAction = False Then
                        If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            oSDBs_Detail1 = objSubForm.DataSources.DBDataSources.Item("@GEN_CAP_PLAN_D1")
                            If oCFL.UniqueID = "LNTPCFL" Then
                                objSubMatrix1 = objSubForm.Items.Item("mtx2").Specific
                                oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim Flag As Boolean = False
                                Dim errflag As Boolean = False
                                If objSubMatrix1.VisualRowCount = 1 Or pVal.Row = objSubMatrix1.VisualRowCount Then
                                    Flag = True
                                End If
                                For i As Integer = 0 To oDT.Rows.Count - 1
                                    Dim cflSelectedcount As Integer = oDT.Rows.Count
                                    If i < cflSelectedcount - 1 Then
                                        objSubMatrix1.AddRow(1, pVal.Row)
                                        oSDBs_Detail1.InsertRecord(pVal.Row + i - 1)
                                    End If
                                    oSDBs_Detail1.Offset = pVal.Row - 1 + i
                                    oSDBs_Detail1.SetValue("LineID", oSDBs_Detail1.Offset, i + pVal.Row)
                                    oSDBs_Detail1.SetValue("u_type", oSDBs_Detail1.Offset, oDT.GetValue("Code", i))
                                    oSDBs_Detail1.SetValue("u_typename", oSDBs_Detail1.Offset, oDT.GetValue("Name", i))
                                    oSDBs_Detail1.SetValue("u_nom", oSDBs_Detail1.Offset, objSubMatrix1.Columns.Item("nom").Cells.Item(pVal.Row).Specific.value)
                                    objSubMatrix1.SetLineData(pVal.Row + i)
                                    objSubForm.EnableMenu("1293", True)
                                Next
                                If Flag = True Then
                                    objSubMatrix1.AddRow(1, objSubMatrix1.VisualRowCount)
                                    Me.SetNewLineChild1(objSubForm.UniqueID, objSubMatrix1.VisualRowCount, objSubMatrix1)
                                End If
                            End If
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        objSubForm.Close()
                        objForm = oApplication.Forms.GetForm("GEN_MACH_ALLOC", oApplication.Forms.ActiveForm.TypeCount)
                        Dim DBSource As SAPbouiCOM.DBDataSource
                        DBSource = objForm.DataSources.DBDataSources.Item("@GEN_MACH_ALLOC_D0")
                        objMatrix = objForm.Items.Item("mtx").Specific
                        objMatrix.Columns.Item("trgtcode").Cells.Item(BaseRow).Specific.value = CStr(TrgtCode)
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        objSubMatrix = objSubForm.Items.Item("mtx").Specific
                        If objSubMatrix.VisualRowCount < 1 Then
                            oApplication.StatusBar.SetText("No dates free in this range", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        objSubMatrix1 = objSubForm.Items.Item("mtx2").Specific
                        If Trim(objSubMatrix1.Columns.Item("type").Cells.Item(1).Specific.value) = "" Then
                            oApplication.StatusBar.SetText("Please enter type of machines", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        Dim NOM As Integer
                        For K As Integer = 1 To objSubMatrix1.VisualRowCount
                            If Trim(objSubMatrix1.Columns.Item("nom").Cells.Item(K).Specific.value) <> "" Then
                                NOM = NOM + CInt(objSubMatrix1.Columns.Item("nom").Cells.Item(K).Specific.value)
                            End If
                        Next
                        If CInt(objSubForm.Items.Item("nom").Specific.value) <> NOM Then
                            oApplication.StatusBar.SetText("No of Machines should match with the total no of machines", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                    If pVal.ItemUID = "2" And pVal.BeforeAction = False Then
                        ModalForm = False
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        ModalForm = False
                    End If
                    If pVal.ItemUID = "btnfill" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = False Then
                        objSubMatrix = objSubForm.Items.Item("mtx").Specific
                        Dim Row As Integer
                        Dim Per, Qty As Double
                        For k As Integer = 1 To objSubMatrix.VisualRowCount
                            If objSubMatrix.IsRowSelected(k) = True Then
                                Row = k
                                Per = objSubMatrix.Columns.Item("per").Cells.Item(k).Specific.Value
                                Qty = objSubMatrix.Columns.Item("qty").Cells.Item(k).Specific.Value
                                Exit For
                            End If
                        Next
                        If Row = Nothing Then
                            oApplication.StatusBar.SetText("please select a row", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        oDBs_Detail = objSubForm.DataSources.DBDataSources.Item("@GEN_CAP_PLAN_D0")
                        objSubForm.Freeze(True)
                        For i As Integer = Row To objSubMatrix.VisualRowCount
                            objSubMatrix.Columns.Item("per").Cells.Item(i).Specific.value = Per
                            objSubMatrix.Columns.Item("qty").Cells.Item(i).Specific.value = Qty
                        Next
                        objSubForm.Freeze(False)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If pVal.ItemUID = "mtx" And pVal.ColUID = "per" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = False Then
                        objSubMatrix = objSubForm.Items.Item("mtx").Specific
                        If objSubMatrix.Columns.Item("per").Cells.Item(pVal.Row).Specific.value > 0 Then
                            objSubMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value = CInt(objSubForm.Items.Item("trgtop").Specific.value * objSubMatrix.Columns.Item("per").Cells.Item(pVal.Row).Specific.value / 100)
                        End If
                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "GEN_MACH_ALLOC"
                        Me.CreateForm(objForm.UniqueID)
                    Case "1282"
                        If objForm.TypeEx = "GEN_MACH_ALLOC" Then
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select Count(*) + 1 AS 'Count' From [@GEN_MACH_ALLOC]")
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MACH_ALLOC")
                            oDBs_Head.SetValue("Code", 0, oRSet.Fields.Item("Count").Value)
                            objMatrix = objForm.Items.Item("mtx").Specific
                            objMatrix.AddRow()
                            Me.SetNewLine(objForm.UniqueID, objMatrix.VisualRowCount, objMatrix)
                        End If
                    Case "1281"
                        If objForm.TypeEx = "GEN_MACH_ALLOC" Then
                            objForm.EnableMenu("1282", True)
                        End If
                    Case "1288", "1289", "1290", "1291"
                        If objForm.TypeEx = "GEN_MACH_ALLOC" Then
                            objForm.EnableMenu("1282", True)
                        End If
                    Case "1293"
                        If objForm.TypeEx = "GEN_MACH_ALLOC" Then
                            If ITEM_ID.Equals("mtx") = True Then
                                objMatrix = objForm.Items.Item("mtx").Specific
                                oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_MACH_ALLOC")
                                oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_MACH_ALLOC_D0")
                                For Row As Integer = 1 To objMatrix.VisualRowCount
                                    objMatrix.GetLineData(Row)
                                    oDBs_Detail.Offset = Row - 1
                                    oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, Row)
                                    oDBs_Detail.SetValue("u_sono", oDBs_Detail.Offset, objMatrix.Columns.Item("sono").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, objMatrix.Columns.Item("itemcode").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, objMatrix.Columns.Item("itemname").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_qty", oDBs_Detail.Offset, objMatrix.Columns.Item("qty").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_deldate", oDBs_Detail.Offset, objMatrix.Columns.Item("deldate").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_stdate", oDBs_Detail.Offset, objMatrix.Columns.Item("stdate").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_eddate", oDBs_Detail.Offset, objMatrix.Columns.Item("eddate").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_unit", oDBs_Detail.Offset, objMatrix.Columns.Item("unit").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_nom", oDBs_Detail.Offset, objMatrix.Columns.Item("nom").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_lineno", oDBs_Detail.Offset, objMatrix.Columns.Item("lineno").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_trgtcode", oDBs_Detail.Offset, objMatrix.Columns.Item("trgtcode").Cells.Item(Row).Specific.value)
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

    Sub CreateFormFind(ByVal FormUID As String, ByVal BaseCode As String, ByVal Code As String, ByVal RowNo As Integer, ByVal SONO As String, ByVal Ln As String, ByVal ITEMCODE As String, ByVal QTY As Double, ByVal UNIT As String, ByVal LINE As String, ByVal NOM As Integer, ByVal STDate As String, ByVal EdDate As String)
        Try
            PARENT_FORM = FormUID
            Dim CHILD_FORM As String = "GEN_CAP_PLAN@" & FormUID
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
                oUtilities.SAPXML("GEN_CAP_PLAN.xml", CHILD_FORM)
                objSubForm = oApplication.Forms.Item(CHILD_FORM)
                objSubForm.Select()
            End If
            ModalForm = True
            objSubForm = oApplication.Forms.Item(FormUID)
            objSubForm = oApplication.Forms.GetForm("GEN_CAP_PLAN", oApplication.Forms.ActiveForm.TypeCount)
            oSDBs_Head = objSubForm.DataSources.DBDataSources.Item("@GEN_CAP_PLAN")
            oSDBs_Detail = objSubForm.DataSources.DBDataSources.Item("@GEN_CAP_PLAN_D0")
            oSDBs_Detail1 = objSubForm.DataSources.DBDataSources.Item("@GEN_CAP_PLAN_D1")
            objSubMatrix = objSubForm.Items.Item("mtx").Specific
            objSubMatrix1 = objSubForm.Items.Item("mtx2").Specific

            BaseRow = RowNo
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Code From [@GEN_CAP_PLAN] Where Code = '" + Code + "'")
            If oRecordSet.RecordCount > 0 Then
                oSDBs_Head.SetValue("u_sono", 0, SONO)
                oSDBs_Head.SetValue("u_itemcode", 0, ITEMCODE)
                oSDBs_Head.SetValue("u_unit", 0, UNIT)
                oSDBs_Head.SetValue("u_line", 0, LINE)
                oSDBs_Head.SetValue("u_ln", 0, Ln)
                oSDBs_Head.SetValue("Code", 0, Code)
                oSDBs_Head.SetValue("u_nom", 0, NOM)
                oSDBs_Head.SetValue("u_basecode", 0, BaseCode)
                objSubForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                objSubForm.Items.Item("flddate").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                objSubForm.EnableMenu("1282", False)
                objSubForm.EnableMenu("1281", False)
                objSubForm.EnableMenu("1288", False)
                objSubForm.EnableMenu("1289", False)
                objSubForm.EnableMenu("1290", False)
                objSubForm.EnableMenu("1291", False)
                'objForm.Items.Item("type").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                'objForm.Items.Item("rptno").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Else
                objSubForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                oSDBs_Head.SetValue("u_sono", 0, SONO)
                oSDBs_Head.SetValue("u_itemcode", 0, ITEMCODE)
                oSDBs_Head.SetValue("u_unit", 0, UNIT)
                oSDBs_Head.SetValue("u_qty", 0, QTY)
                oSDBs_Head.SetValue("u_nom", 0, NOM)
                oSDBs_Head.SetValue("u_line", 0, LINE)
                oSDBs_Head.SetValue("u_stdate", 0, STDate)
                oSDBs_Head.SetValue("u_eddate", 0, EdDate)
                oSDBs_Head.SetValue("u_basecode", 0, BaseCode)
                oSDBs_Head.SetValue("u_ln", 0, Ln)
                objSubForm.Items.Item("flddate").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRSet.DoQuery("Select Count(*) + 1 AS 'Count' From [@GEN_CAP_PLAN]")
                oSDBs_Head.SetValue("Code", 0, oRSet.Fields.Item("Count").Value)
                TrgtCode = oRSet.Fields.Item("Count").Value
                objSubForm.EnableMenu("1288", False)
                objSubForm.EnableMenu("1281", False)
                objSubForm.EnableMenu("1289", False)
                objSubForm.EnableMenu("1290", False)
                objSubForm.EnableMenu("1291", False)
                oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRSet.DoQuery("Exec Planning_DateList '" + STDate + "','" + EdDate + "','" + UNIT + "','" + LINE + "'")
                For i As Integer = 1 To oRSet.RecordCount
                    objSubMatrix.AddRow()
                    Me.SetNewLineChild(objSubForm.UniqueID, objSubMatrix.VisualRowCount, objSubMatrix)
                    oSDBs_Detail.Offset = i - 1
                    oSDBs_Detail.SetValue("u_cdate", oSDBs_Detail.Offset, Format(oRSet.Fields.Item("WorkingDays").Value, "yyyyMMdd"))
                    objSubMatrix.SetLineData(i)
                    oRSet.MoveNext()
                Next
                objSubMatrix1.AddRow()
                Me.SetNewLineChild1(objSubForm.UniqueID, objSubMatrix1.VisualRowCount, objSubMatrix1)
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLineChild(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objSubForm = oApplication.Forms.Item(FormUID)
            oSDBs_Detail = objSubForm.DataSources.DBDataSources.Item("@GEN_CAP_PLAN_D0")
            objMatrix = oMatrix
            objSubForm.Freeze(True)
            objMatrix.FlushToDataSource()
            oSDBs_Detail.Offset = Row - 1
            oSDBs_Detail.SetValue("LineId", oSDBs_Detail.Offset, objMatrix.VisualRowCount)
            oSDBs_Detail.SetValue("u_cdate", oSDBs_Detail.Offset, "")
            oSDBs_Detail.SetValue("u_avlbl", oSDBs_Detail.Offset, "")
            oSDBs_Detail.SetValue("u_per", oSDBs_Detail.Offset, "")
            oSDBs_Detail.SetValue("u_qty", oSDBs_Detail.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objSubForm.Freeze(False)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLineChild1(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objSubForm = oApplication.Forms.Item(FormUID)
            oSDBs_Detail1 = objSubForm.DataSources.DBDataSources.Item("@GEN_CAP_PLAN_D1")
            objMatrix = oMatrix
            objSubForm.Freeze(True)
            objMatrix.FlushToDataSource()
            oSDBs_Detail1.Offset = Row - 1
            oSDBs_Detail1.SetValue("LineId", oSDBs_Detail1.Offset, objMatrix.VisualRowCount)
            oSDBs_Detail1.SetValue("u_type", oSDBs_Detail1.Offset, "")
            oSDBs_Detail1.SetValue("u_typename", oSDBs_Detail1.Offset, "")
            oSDBs_Detail1.SetValue("u_nom", oSDBs_Detail1.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objSubForm.Freeze(False)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterTypeCode(ByVal FormUID As String, ByVal Unit As String, ByVal StDate As String, ByVal EdDate As String)
        Try
            objSubForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objSubForm.ChooseFromLists.Item("LNTPCFL")
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Distinct A.Code From [@GEN_LINE_TYPE] A Where A.Code in (Select Y.u_type From [@GEN_MACH_POOL] X INNER JOIN [@GEN_MACH_POOL_D0] Y ON X.CODE = Y.CODE AND X.NAME = '" + Unit + "') And A.Code Not In ( Select R.u_type From [@GEN_CAP_PLAN_D1] R INNER JOIN [@GEN_CAP_PLAN_D0] S ON R.CODE = S.CODE Where S.u_cdate between '" + StDate + "' And '" + EdDate + "' ) ")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            If oRecordSet.RecordCount = 0 Then
                oCon = oCons.Add()
                oCon.Alias = "Code"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "-1"
            End If
            For IntICount As Integer = 0 To oRecordSet.RecordCount - 1
                If IntICount = (oRecordSet.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "Code"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("Code").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "Code"
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

    Sub Child_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = True Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        objMatrix = objForm.Items.Item("mtx2").Specific
                        If objMatrix.VisualRowCount <> 0 Then
                            objMatrix.DeleteRow(objMatrix.VisualRowCount)
                            objMatrix.FlushToDataSource()
                        End If
                    ElseIf BusinessObjectInfo.ActionSuccess = True Then
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Then
                            objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                            objMatrix = objForm.Items.Item("mtx2").Specific
                            objMatrix.AddRow()
                            objMatrix.FlushToDataSource()
                            Me.SetNewLineChild1(BusinessObjectInfo.FormUID, objMatrix.VisualRowCount, objMatrix)
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        objMatrix = objForm.Items.Item("mtx2").Specific
                        objMatrix.AddRow()
                        objMatrix.FlushToDataSource()
                        Me.SetNewLineChild1(BusinessObjectInfo.FormUID, objMatrix.VisualRowCount, objMatrix)
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
