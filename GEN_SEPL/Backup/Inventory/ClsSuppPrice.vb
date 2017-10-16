﻿Public Class ClsSuppPrice


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
    Public sDocNum As String
    Public sRptName As String

#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            oUtilities.SAPXML("Supp_Price.xml")
            objForm = oApplication.Forms.GetForm("GEN_SUPP_PRICE", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUPP_PRICE")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SUPP_PRICE_D0")
            objForm.EnableMenu("1288", True)
            objForm.EnableMenu("1289", True)
            objForm.EnableMenu("1290", True)
            objForm.EnableMenu("1291", True)
            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            '    FilterSO(FormUID)
            'FilterItems(FormUID)
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SUPP_PRICE_D0")
            objMatrix = oMatrix
            objMatrix.FlushToDataSource()
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, objMatrix.VisualRowCount)
            oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_price", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_ansp", oDBs_Detail.Offset, "")
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
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("CustCFL")
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add()
            oCon.Alias = "u_pc"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "PC2"
            'Vijeesh_24_5_2012
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
            oCon = oCons.Add()
            oCon.Alias = "u_pc"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "PC6"
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            'Vijeesh_24_5_2012
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
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("ItCFL")
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "u_pc"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "PC2"
            ''Vijeesh_24_5_2012
            'oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
            'oCon = oCons.Add()
            'oCon.Alias = "u_pc"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "PC6"
            ''Vijeesh_24_5_2012
            'oCFL.SetConditions(oCons)
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Distinct U_unit From OCRD Where CardCode = '" + Trim(objForm.Items.Item("cardcode").Specific.value) + "'")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            For IntICount As Integer = 0 To oRecordSet.RecordCount - 1
                If IntICount = (oRecordSet.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "U_unit"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("U_unit").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "U_unit"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("U_unit").Value
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                oRecordSet.MoveNext()
            Next
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUPP_PRICE")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SUPP_PRICE_D0")
            If Trim(oDBs_Head.GetValue("u_cardcode", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter Customer Code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            If Trim(oDBs_Head.GetValue("u_doccur", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter Currency", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
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
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    End If
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If pVal.ItemUID = "cardcode" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                        Dim oRSet As SAPbobsCOM.Recordset
                        oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRSet.DoQuery("Select u_cardcode From [@GEN_SUPP_PRICE] Where u_cardcode = '" + Trim(objForm.Items.Item("cardcode").Specific.value) + "'")
                        If oRSet.RecordCount > 0 Then
                            oApplication.StatusBar.SetText("Supplementary Price already added for this customer", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                    'Vijeesh
                    If pVal.BeforeAction = True Then
                        If oCFL.UniqueID = "ItCFL" Then
                            If Trim(objForm.Items.Item("cardcode").Specific.value) <> "" Then
                                FilterItems(FormUID)
                            Else
                                oApplication.StatusBar.SetText("Please Select Customer", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If
                    'Vijeesh
                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If oCFL.UniqueID = "CustCFL" Then
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUPP_PRICE")
                            oDBs_Head.SetValue("u_cardcode", 0, oDT.GetValue("CardCode", 0))
                            oDBs_Head.SetValue("u_cardname", 0, oDT.GetValue("CardName", 0))
                        End If
                        If oCFL.UniqueID = "ItCFL" Then
                            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            objMatrix = objForm.Items.Item("mtx").Specific
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SUPP_PRICE_D0")
                            Dim Flag As Boolean = False
                            Dim erflag As Boolean = False
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
                                Dim j As Integer
                                j = 1
                                While j <= objMatrix.RowCount
                                    If objMatrix.Columns.Item("itemcode").Cells.Item(j).Specific.Value = oDT.GetValue("ItemCode", i) Then
                                        erflag = True
                                    End If
                                    j = j + 1
                                End While
                                If erflag Then
                                    oApplication.StatusBar.SetText("Items cannot be duplicated", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                ElseIf erflag = False Then
                                    oDBs_Detail.SetValue("U_itemcode", oDBs_Detail.Offset, oDT.GetValue("ItemCode", i))
                                    oDBs_Detail.SetValue("U_itemname", oDBs_Detail.Offset, oDT.GetValue("ItemName", i))
                                    objMatrix.SetLineData(pVal.Row + i)
                                End If
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

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "GEN_SUPP_PRICE"
                        Me.CreateForm(objForm.UniqueID)
                    Case "1282"
                        If objForm.TypeEx = "GEN_SUPP_PRICE" Then
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRSet.DoQuery("Select Count(*) + 1 AS 'Count' From [@GEN_SUPP_PRICE]")
                            oRSet.DoQuery("SELECT isnull(MAX(CAST( Code AS int)),0) +1 AS 'Count'  FROM [@GEN_SUPP_PRICE]")
                            objForm.Items.Item("code").Specific.value = oRSet.Fields.Item("Count").Value
                            objForm.Items.Item("code").Specific.value = oRSet.Fields.Item("Count").Value
                            objMatrix = objForm.Items.Item("mtx").Specific
                            objMatrix.AddRow(1, objMatrix.VisualRowCount)
                            Me.SetNewLine(objForm.UniqueID, objMatrix.VisualRowCount, objMatrix)
                        End If
                    Case "1281"
                        If objForm.TypeEx = "GEN_SUPP_PRICE" Then
                            objForm.EnableMenu("1282", True)
                        End If
                    Case "1288", "1289", "1290", "1291"
                        If objForm.TypeEx = "GEN_SUPP_PRICE" Then
                            objForm.EnableMenu("1282", True)
                        End If
                    Case "1293"
                        If objForm.TypeEx = "GEN_SUPP_PRICE" Then
                            If ITEM_ID.Equals("mtx") = True Then
                                objMatrix = objForm.Items.Item("mtx").Specific
                                oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUPP_PRICE")
                                oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SUPP_PRICE_D0")
                                For Row As Integer = 1 To objMatrix.VisualRowCount
                                    objMatrix.GetLineData(Row)
                                    oDBs_Detail.Offset = Row - 1
                                    oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, Row)
                                    oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, objMatrix.Columns.Item("itemcode").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, objMatrix.Columns.Item("itemname").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_price", oDBs_Detail.Offset, objMatrix.Columns.Item("price").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_ansp", oDBs_Detail.Offset, objMatrix.Columns.Item("ansp").Cells.Item(Row).Specific.value)
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
