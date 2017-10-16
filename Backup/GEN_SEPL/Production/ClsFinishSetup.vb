Public Class ClsFinishSetup

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
            oUtilities.SAPXML("GEN_FIN_SETUP.xml")
            objForm = oApplication.Forms.GetForm("GEN_FIN_SETUP", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_FIN_SETUP")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_FIN_SETUP_D0")
            objForm.DataBrowser.BrowseBy = "itemcode"
            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            objForm.Items.Item("itemcode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("prodno").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("cardcode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_FIN_SETUP_D0")
            objMatrix = oMatrix
            objMatrix.FlushToDataSource()
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, objMatrix.VisualRowCount)
            oDBs_Detail.SetValue("u_prcs", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_prcsname", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_reqdno", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_cappm", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_trgtop", oDBs_Detail.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try
            objMatrix = objForm.Items.Item("mtx").Specific
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_FIN_SETUP")
            If Trim(oDBs_Head.GetValue("u_itemcode", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter style code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If Trim(oDBs_Head.GetValue("u_itemcode", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter style code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If Trim(oDBs_Head.GetValue("u_prodno", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter Production Order", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSet.DoQuery("Select Code From [@GEN_FIN_SETUP] Where u_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' And u_prodno = '" + Trim(objForm.Items.Item("prodno").Specific.value) + "' And Code <> '" + Trim(objForm.Items.Item("code").Specific.value) + "'")
            If oRSet.RecordCount > 0 Then
                oApplication.StatusBar.SetText("Finishing already defined for this production order", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Sub FilterProductionOrders(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("PRDCFL")
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Distinct A.DocNum From OWOR A Where A.ItemCode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' And A.u_process = 'Finishing'")
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
            If oRecordSet.RecordCount > 0 Then
                oCFL.SetConditions(oCons)
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
                        If Me.Validation(FormUID) = False Then BubbleEvent = False
                    ElseIf pVal.ItemUID = "1" And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRSet.DoQuery("Select Count(*) + 1 AS 'Count' From [@GEN_FIN_SETUP]")
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_FIN_SETUP")
                        oDBs_Head.SetValue("Code", 0, oRSet.Fields.Item("Count").Value)
                        objMatrix = objForm.Items.Item("mtx").Specific
                        objMatrix.AddRow()
                        Me.SetNewLine(objForm.UniqueID, objMatrix.VisualRowCount, objMatrix)
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
                            If Trim(oDBs_Head.GetValue("u_itemcode", 0)) = "" Then
                                oApplication.StatusBar.SetText("Please enter item code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("Select Distinct A.DocNum From OWOR A Where A.ItemCode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' And A.u_process = 'Finishing'")
                            If oRecordSet.RecordCount = 0 Then
                                oApplication.StatusBar.SetText("No Production Orders for this item", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            Else
                                Me.FilterProductionOrders(FormUID)
                            End If
                        End If
                    End If
                    If pVal.BeforeAction = False Then
                        If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_FIN_SETUP")
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_FIN_SETUP_D0")
                            If oCFL.UniqueID = "ITCFL" Then
                                oDBs_Head.SetValue("u_itemcode", 0, oDT.GetValue("ItemCode", 0))
                                oDBs_Head.SetValue("u_itemname", 0, oDT.GetValue("ItemName", 0))
                            End If
                            If oCFL.UniqueID = "PRDCFL" Then
                                oDBs_Head.SetValue("u_prodno", 0, oDT.GetValue("DocNum", 0))
                                oDBs_Head.SetValue("u_cardcode", 0, oDT.GetValue("CardCode", 0))
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet.DoQuery("Select CardName From OCRD Where CardCode = '" + Trim(oDT.GetValue("CardCode", 0)) + "'")
                                oDBs_Head.SetValue("u_cardname", 0, oRSet.Fields.Item("CardName").Value)
                            End If
                            If oCFL.UniqueID = "PRCSCFL" Then
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
                                    oDBs_Detail.SetValue("u_prcs", oDBs_Detail.Offset, oDT.GetValue("Code", i))
                                    oDBs_Detail.SetValue("u_prcsname", oDBs_Detail.Offset, oDT.GetValue("Name", i))
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
                    Case "GEN_FIN_SETUP"
                        Me.CreateForm(objForm.UniqueID)
                    Case "1282"
                        If objForm.TypeEx = "GEN_FIN_SETUP" Then
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select Count(*) + 1 AS 'Count' From [@GEN_FIN_SETUP]")
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_FIN_SETUP")
                            oDBs_Head.SetValue("Code", 0, oRSet.Fields.Item("Count").Value)
                            objMatrix = objForm.Items.Item("mtx").Specific
                            objMatrix.AddRow()
                            Me.SetNewLine(objForm.UniqueID, objMatrix.VisualRowCount, objMatrix)
                        End If
                    Case "1281"
                        If objForm.TypeEx = "GEN_FIN_SETUP" Then
                            objForm.EnableMenu("1282", True)
                        End If
                    Case "1288", "1289", "1290", "1291"
                        If objForm.TypeEx = "GEN_FIN_SETUP" Then
                            objForm.EnableMenu("1282", True)
                        End If
                    Case "1293"
                        If objForm.TypeEx = "GEN_FIN_SETUP" Then
                            If ITEM_ID.Equals("mtx") = True Then
                                objMatrix = objForm.Items.Item("mtx").Specific
                                oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_FIN_SETUP")
                                oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_FIN_SETUP_D0")
                                For Row As Integer = 1 To objMatrix.VisualRowCount
                                    objMatrix.GetLineData(Row)
                                    oDBs_Detail.Offset = Row - 1
                                    oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, Row)
                                    oDBs_Detail.SetValue("u_prcs", oDBs_Detail.Offset, objMatrix.Columns.Item("prcs").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_prcsname", oDBs_Detail.Offset, objMatrix.Columns.Item("prcsname").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_reqdno", oDBs_Detail.Offset, objMatrix.Columns.Item("reqdno").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_cappm", oDBs_Detail.Offset, objMatrix.Columns.Item("cappm").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_trgtop", oDBs_Detail.Offset, objMatrix.Columns.Item("trgtop").Cells.Item(Row).Specific.value)
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
