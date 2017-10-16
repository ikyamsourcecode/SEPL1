Public Class ClsGRPOLC

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm, GRPOForm As SAPbouiCOM.Form
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim objMatrix, GRPOMatrix As SAPbouiCOM.Matrix
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
            oUtilities.SAPXML("GEN_LCosts.xml")
            objForm = oApplication.Forms.GetForm("GEN_LCOSTS", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_LCOSTS")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_LCOSTS_D0")
            objForm.DataBrowser.BrowseBy = "code"
            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            objForm.Items.Item("name").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_GRPO_LCOSTS_D0")
            objMatrix = oMatrix
            objMatrix.FlushToDataSource()
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, objMatrix.VisualRowCount)
            oDBs_Detail.SetValue("u_lname", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_glacct", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_rate", oDBs_Detail.Offset, "")

            oDBs_Detail.SetValue("u_qty", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_value", oDBs_Detail.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try
            'objMatrix = objForm.Items.Item("mtx").Specific
            'oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_UNIT_MST")
            'If Trim(oDBs_Head.GetValue("Name", 0)) = "" Then
            '    oApplication.StatusBar.SetText("Please enter Unit Name", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If
            'Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRSet.DoQuery("Select Name From [@GEN_UNIT_MST] Where Name = '" + Trim(oDBs_Head.GetValue("Name", 0)) + "' And Code <> '" + Trim(objForm.Items.Item("code").Specific.value) + "'")
            'If oRSet.RecordCount > 0 Then
            '    oApplication.StatusBar.SetText("Unit name already exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If
            'If CheckForWild(Trim(objForm.Items.Item("name").Specific.value)) = False Then
            '    oApplication.StatusBar.SetText("Unit Name cannot have special characters", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            '    Exit Function
            'End If
            'For I As Integer = 1 To objMatrix.VisualRowCount
            '    If Trim(objMatrix.Columns.Item("process").Cells.Item(I).Specific.value) <> "" Then
            '        If Trim(objMatrix.Columns.Item("inwhs").Cells.Item(I).Specific.value) = "" Or Trim(objMatrix.Columns.Item("outwhs").Cells.Item(I).Specific.value) = "" Or Trim(objMatrix.Columns.Item("stwhs").Cells.Item(I).Specific.value) = "" Then
            '            oApplication.StatusBar.SetText("Please enter in, out and stored warehouses", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '            Return False
            '        End If
            '    End If
            'Next
            'Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                    objMatrix = objForm.Items.Item("mtx").Specific
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.ItemUID = "1" And pVal.ActionSuccess = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        Dim oRSet1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRSet1.DoQuery("SELECT top 1 u_grpono from [@GEN_GRPO_LCOSTS] order by u_grpono desc ")
                        objForm.Items.Item("grpono").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        objForm.Items.Item("grpono").Specific.value = oRSet1.Fields.Item(0).Value
                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                    'And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormModede.fm_UPDATE_MODE) Then
                    '    ' If Me.Validation(FormUID) = False Then BubbleEvent = False
                    'ElseIf pVal.ItemUID = "1" And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    '    Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '    oRSet.DoQuery("Select Count(*) + 1 AS 'Count' From [@GEN_UNIT_MST]")
                    '    oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_UNIT_MST")
                    '    oDBs_Head.SetValue("Code", 0, oRSet.Fields.Item("Count").Value)
                    '    objMatrix = objForm.Items.Item("mtx").Specific
                    '    objMatrix.AddRow()
                    '    Me.SetNewLine(objForm.UniqueID, objMatrix.VisualRowCount, objMatrix)
                    'End If
                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                    objMatrix = objForm.Items.Item("mtx").Specific
                    'GRPOForm = oApplication.Forms.GetForm("OPDN", oApplication.Forms.ActiveForm.TypeCount)
                    ' GRPOForm = oApplication.Forms.Item(pVal.FormUID)
                    'GRPOMatrix = GRPOForm.Items.Item("38").Specific
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_GRPO_LCOSTS")
                    oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_GRPO_LCOSTS_D0")
                    If pVal.ItemUID = "mtx" And pVal.ColUID = "lname" And pVal.ActionSuccess = True Then
                        Dim Flag As Boolean = False
                        Dim errflag As Boolean = False
                        If objMatrix.VisualRowCount = 1 Or pVal.Row = objMatrix.VisualRowCount Then
                            Flag = True
                        End If
                        If Flag = True Then
                            objMatrix.AddRow(1, objMatrix.VisualRowCount)
                            Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
                        End If
                        For i As Integer = 1 To objMatrix.VisualRowCount
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("SELECT ALCNAME,LaCAllcAcc,U_USER,OACT.FormatCode FROM OALC INNER JOIN OACT ON OACT.AcctCode=OALC.LaCAllcAcc where AlcName='" + objMatrix.Columns.Item("lname").Cells.Item(i).Specific.value + "'")
                            'oDBs_Detail.SetValue("u_glacct", oDBs_Detail.Offset, oRSet.Fields.Item(3).Value)
                            objMatrix.Columns.Item("rate").Cells.Item(i).Specific.value = oRSet.Fields.Item(2).Value
                            objMatrix.Columns.Item("glacct").Cells.Item(i).Specific.value = oRSet.Fields.Item(3).Value
                            oDBs_Detail.SetValue("u_rate", oDBs_Detail.Offset, oRSet.Fields.Item("U_USER").Value)

                        Next
                        'objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        'Dim oRSet1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'oRSet1.DoQuery("SELECT top 1 u_grpono from [@GEN_GRPO_LCOSTS] order by u_grpono desc ")
                        'objForm.Items.Item("docnum").Specific.value = oRSet1.Fields.Item(0).Value
                        'objForm.Items.Item("1").Click()
                        'Dim Total As Integer = 0
                        'For i As Integer = 0 To GRPOMatrix.VisualRowCount
                        '    Total = Total + GRPOMatrix.Columns.Item("11").Cells.Item(i).Specific.value
                        'Next
                        '   oDBs_Detail.SetValue("u_qty", oDBs_Detail.Offset, oRecordSet.Fields.Item("U_USER").Value)
                        'Me.SetNewLine(objForm.UniqueID, objMatrix.VisualRowCount, objMatrix)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    'Dim oCFL As SAPbouiCOM.ChooseFromList
                    'Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    'Dim CFL_Id As String
                    'CFL_Id = CFLEvent.ChooseFromListUID
                    'oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    'Dim oDT As SAPbouiCOM.DataTable
                    'oDT = CFLEvent.SelectedObjects
                    'If pVal.BeforeAction = False Then
                    '    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    '        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_LCOSTS")
                    '        oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_LCOSTS_D0")
                    '        If oCFL.UniqueID = "GLCFL" Then
                    '            objMatrix = objForm.Items.Item("mtx").Specific
                    '            oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            Dim Flag As Boolean = False
                    '            Dim errflag As Boolean = False
                    '            If objMatrix.VisualRowCount = 1 Or pVal.Row = objMatrix.VisualRowCount Then
                    '                Flag = True
                    '            End If
                    '            For i As Integer = 0 To oDT.Rows.Count - 1
                    '                Dim cflSelectedcount As Integer = oDT.Rows.Count
                    '                If i < cflSelectedcount - 1 Then
                    '                    objMatrix.AddRow(1, pVal.Row)
                    '                    oDBs_Detail.InsertRecord(pVal.Row + i - 1)
                    '                End If
                    '                oDBs_Detail.Offset = pVal.Row - 1 + i
                    '                oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                    '                oDBs_Detail.SetValue("u_lcode", oDBs_Detail.Offset, objMatrix.Columns.Item("lcode").Cells.Item(pVal.Row).Specific.value)
                    '                oDBs_Detail.SetValue("u_lname", oDBs_Detail.Offset, objMatrix.Columns.Item("lname").Cells.Item(pVal.Row).Specific.value)
                    '                oDBs_Detail.SetValue("u_rate", oDBs_Detail.Offset, objMatrix.Columns.Item("rate").Cells.Item(pVal.Row).Specific.value)
                    '                'oDBs_Detail.SetValue("u_glacct", oDBs_Detail.Offset, objMatrix.Columns.Item("stwhs").Cells.Item(Row).Specific.value)
                    '                oDBs_Detail.SetValue("u_glacct", oDBs_Detail.Offset, oDT.GetValue("FormatCode", i))
                    '                objMatrix.SetLineData(pVal.Row + i)
                    '                objForm.EnableMenu("1293", True)
                    '            Next
                    '            If Flag = True Then
                    '                objMatrix.AddRow(1, objMatrix.VisualRowCount)
                    '                Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
                    '            End If
                    '        End If
                    '                If oCFL.UniqueID = "WHCFL1" Then
                    '                    objMatrix = objForm.Items.Item("mtx").Specific
                    '                    oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                    oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                    For i As Integer = 0 To oDT.Rows.Count - 1
                    '                        Dim cflSelectedcount As Integer = oDT.Rows.Count
                    '                        oDBs_Detail.Offset = pVal.Row - 1 + i
                    '                        oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                    '                        oDBs_Detail.SetValue("u_process", oDBs_Detail.Offset, objMatrix.Columns.Item("process").Cells.Item(pVal.Row).Specific.Value)
                    '                        oDBs_Detail.SetValue("u_inwhs", oDBs_Detail.Offset, oDT.GetValue("WhsCode", i))
                    '                        oDBs_Detail.SetValue("u_outwhs", oDBs_Detail.Offset, objMatrix.Columns.Item("outwhs").Cells.Item(pVal.Row).Specific.Value)
                    '                        oDBs_Detail.SetValue("u_stwhs", oDBs_Detail.Offset, objMatrix.Columns.Item("stwhs").Cells.Item(pVal.Row).Specific.Value)
                    '                        objMatrix.SetLineData(pVal.Row + i)
                    '                    Next
                    '                End If
                    '                If oCFL.UniqueID = "WHCFL2" Then
                    '                    objMatrix = objForm.Items.Item("mtx").Specific
                    '                    oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                    oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                    For i As Integer = 0 To oDT.Rows.Count - 1
                    '                        Dim cflSelectedcount As Integer = oDT.Rows.Count
                    '                        oDBs_Detail.Offset = pVal.Row - 1 + i
                    '                        oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                    '                        oDBs_Detail.SetValue("u_process", oDBs_Detail.Offset, objMatrix.Columns.Item("process").Cells.Item(pVal.Row).Specific.Value)
                    '                        oDBs_Detail.SetValue("u_inwhs", oDBs_Detail.Offset, objMatrix.Columns.Item("inwhs").Cells.Item(pVal.Row).Specific.Value)
                    '                        oDBs_Detail.SetValue("u_outwhs", oDBs_Detail.Offset, oDT.GetValue("WhsCode", i))
                    '                        oDBs_Detail.SetValue("u_stwhs", oDBs_Detail.Offset, objMatrix.Columns.Item("stwhs").Cells.Item(pVal.Row).Specific.Value)
                    '                        objMatrix.SetLineData(pVal.Row + i)
                    '                    Next
                    '                End If
                    '                If oCFL.UniqueID = "WHCFL3" Then
                    '                    objMatrix = objForm.Items.Item("mtx").Specific
                    '                    oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                    oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                    For i As Integer = 0 To oDT.Rows.Count - 1
                    '                        Dim cflSelectedcount As Integer = oDT.Rows.Count
                    '                        oDBs_Detail.Offset = pVal.Row - 1 + i
                    '                        oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                    '                        oDBs_Detail.SetValue("u_process", oDBs_Detail.Offset, objMatrix.Columns.Item("process").Cells.Item(pVal.Row).Specific.Value)
                    '                        oDBs_Detail.SetValue("u_inwhs", oDBs_Detail.Offset, objMatrix.Columns.Item("inwhs").Cells.Item(pVal.Row).Specific.Value)
                    '                        oDBs_Detail.SetValue("u_outwhs", oDBs_Detail.Offset, objMatrix.Columns.Item("outwhs").Cells.Item(pVal.Row).Specific.Value)
                    '                        oDBs_Detail.SetValue("u_stwhs", oDBs_Detail.Offset, oDT.GetValue("WhsCode", i))
                    '                        objMatrix.SetLineData(pVal.Row + i)
                    '                    Next
                    ' End If
                    'End If
                    'End If
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
                    Case "GEN_GRPO_LCOSTS"
                        Me.CreateForm(objForm.UniqueID)
                        '    Case "1282"
                        '        If objForm.TypeEx = "GEN_LCOSTS" Then
                        '            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '            oRSet.DoQuery("Select Count(*) + 1 AS 'Count' From [@GEN_LCOSTS]")
                        '            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_LCOSTS")
                        '            oDBs_Head.SetValue("Code", 0, oRSet.Fields.Item("Count").Value)
                        '            objMatrix = objForm.Items.Item("mtx").Specific
                        '            objMatrix.AddRow()
                        '            Me.SetNewLine(objForm.UniqueID, objMatrix.VisualRowCount, objMatrix)
                        '        End If
                        '    Case "1281"
                        '        If objForm.TypeEx = "GEN_LCOSTS" Then
                        '            objForm.EnableMenu("1282", True)
                        '        End If
                        '    Case "1288", "1289", "1290", "1291"
                        '        If objForm.TypeEx = "GEN_LCOSTS" Then
                        '            objForm.EnableMenu("1282", True)
                        '        End If
                        '    Case "1293"
                        '        If objForm.TypeEx = "GEN_LCOSTS" Then
                        '            If ITEM_ID.Equals("mtx") = True Then
                        '                objMatrix = objForm.Items.Item("mtx").Specific
                        '                oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_LCOSTS")
                        '                oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_LCOSTS_D0")
                        '                For Row As Integer = 1 To objMatrix.VisualRowCount
                        '                    objMatrix.GetLineData(Row)
                        '                    oDBs_Detail.Offset = Row - 1
                        '                    oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, Row)
                        '                    oDBs_Detail.SetValue("u_lcode", oDBs_Detail.Offset, objMatrix.Columns.Item("process").Cells.Item(Row).Specific.value)
                        '                    oDBs_Detail.SetValue("u_lname", oDBs_Detail.Offset, objMatrix.Columns.Item("inwhs").Cells.Item(Row).Specific.value)
                        '                    oDBs_Detail.SetValue("u_rate", oDBs_Detail.Offset, objMatrix.Columns.Item("outwhs").Cells.Item(Row).Specific.value)
                        '                    oDBs_Detail.SetValue("u_glacct", oDBs_Detail.Offset, objMatrix.Columns.Item("stwhs").Cells.Item(Row).Specific.value)
                        '                    objMatrix.SetLineData(Row)
                        '                Next
                        '                objMatrix.FlushToDataSource()
                        '                oDBs_Detail.RemoveRecord(oDBs_Detail.Size - 1)
                        '                objMatrix.LoadFromDataSource()
                        '            End If
                        '        End If
                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            'RowCount = eventInfo.Row
            'If eventInfo.Row > 0 Then
            '    ITEM_ID = eventInfo.ItemUID
            '    Dim oForm As SAPbouiCOM.Form = oApplication.Forms.Item(eventInfo.FormUID)
            '    objMatrix = oForm.Items.Item("mtx").Specific
            '    If objMatrix.VisualRowCount > 1 Then
            '        oForm.EnableMenu("1293", True)
            '    Else
            '        oForm.EnableMenu("1293", False)
            '    End If
            'Else
            '    ITEM_ID = ""
            'End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            'Select Case BusinessObjectInfo.EventType
            '    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
            '        If BusinessObjectInfo.BeforeAction = True Then
            '            objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
            '            objMatrix = objForm.Items.Item("mtx").Specific
            '            If objMatrix.VisualRowCount <> 0 Then
            '                objMatrix.DeleteRow(objMatrix.VisualRowCount)
            '                objMatrix.FlushToDataSource()
            '            End If
            '        ElseIf BusinessObjectInfo.ActionSuccess = True Then
            '            If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Then
            '                objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
            '                objMatrix = objForm.Items.Item("mtx").Specific
            '                objMatrix.AddRow()
            '                objMatrix.FlushToDataSource()
            '                Me.SetNewLine(BusinessObjectInfo.FormUID, objMatrix.VisualRowCount, objMatrix)
            '            End If
            '        End If
            '    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
            '        If BusinessObjectInfo.BeforeAction = False Then
            '            objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
            '            objMatrix = objForm.Items.Item("mtx").Specific
            '            objMatrix.AddRow()
            '            objMatrix.FlushToDataSource()
            '            Me.SetNewLine(BusinessObjectInfo.FormUID, objMatrix.VisualRowCount, objMatrix)
            '        End If
            'End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Function CheckForWild(ByVal Code As String) As Boolean
        Dim iCount As Integer = 0
        Dim c As Char
        For iCount = 0 To Code.Length - 1
            c = Code.Substring(iCount, 1)
            Select Case c
                Case "!" : Return False
                Case "@" : Return False
                Case "#" : Return False
                Case "$" : Return False
                Case "%" : Return False
                Case "^" : Return False
                Case "&" : Return False
                Case "*" : Return False
                Case "(" : Return False
                Case ")" : Return False
                Case "_" : Return False
                Case "+" : Return False
                Case "=" : Return False
                Case "[" : Return False
                Case "]" : Return False
                Case "{" : Return False
                Case "}" : Return False
                Case "\" : Return False
                Case "|" : Return False
                Case ";" : Return False
                Case ":" : Return False
                Case """" : Return False
                Case "'" : Return False
                Case "<" : Return False
                Case ">" : Return False
                Case "?" : Return False
                Case "/" : Return False
            End Select
        Next
        Return True
    End Function


    '    End Class


End Class
