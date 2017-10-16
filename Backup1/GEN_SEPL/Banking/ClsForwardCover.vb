Imports System
Public Class ClsForwardCover

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
    Dim doc As String
    Dim status, unit As String
    Dim preamt As String
    Dim tdt As String
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            oUtilities.SAPXML("ForwardCover.xml")
            objForm = oApplication.Forms.GetForm("GEN_FWD_COVER", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_FWD_COVER")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_FWD_COVER_D0")
            objForm.EnableMenu("1288", True)
            objForm.EnableMenu("1289", True)
            objForm.EnableMenu("1290", True)
            objForm.EnableMenu("1291", True)
            objForm.DataBrowser.BrowseBy = "docnum"
            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_FWD_COVER")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_FWD_COVER_D0")
            oUtilities.GetSeries(FormUID, "series", "GEN_FWD_COVER")
            oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("series").Specific.Selected.Value, "GEN_FWD_COVER"))
            oDBs_Head.SetValue("U_docdate", 0, DateTime.Today.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_status", 0, "Open")
            'objMatrix = objForm.Items.Item("mtx").Specific
            'objMatrix.AddRow(1, objMatrix.VisualRowCount)
            'Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    'Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
    '    Try
    '        objForm = oApplication.Forms.Item(FormUID)
    '        oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_FWD_COVER_D0")
    '        objMatrix = oMatrix
    '        objMatrix.FlushToDataSource()
    '        oDBs_Detail.Offset = Row - 1
    '        oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, objMatrix.VisualRowCount)
    '        oDBs_Detail.SetValue("u_docdate", oDBs_Detail.Offset, "")
    '        oDBs_Detail.SetValue("u_amount", oDBs_Detail.Offset, "")
    '        oDBs_Detail.SetValue("u_remarks", oDBs_Detail.Offset, "")
    '        oDBs_Detail.SetValue("u_pc", oDBs_Detail.Offset, "")
    '        objMatrix.SetLineData(objMatrix.VisualRowCount)
    '    Catch ex As Exception
    '        oApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_FWD_COVER")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_FWD_COVER_D0")
            If Trim(oDBs_Head.GetValue("u_fdate", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter from date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            If Trim(oDBs_Head.GetValue("u_tdate", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter to date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSet.DoQuery("Select DateDiff(dd,'" + oDBs_Head.GetValue("u_fdate", 0) + "','" + oDBs_Head.GetValue("u_tdate", 0) + "') As 'Diff'")
            If oRSet.Fields.Item("Diff").Value < 0 Then
                oApplication.StatusBar.SetText("To date cannot be less than from date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            If Trim(oDBs_Head.GetValue("u_doccur", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter document currency", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            If Trim(oDBs_Head.GetValue("u_amount", 0)) <= 0 Then
                oApplication.StatusBar.SetText("Please enter valid amount", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            If Trim(oDBs_Head.GetValue("u_contrno", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter contract no.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            oRSet.DoQuery("Select * from [@GEN_FWD_COVER] Where u_contrno = '" + Trim(oDBs_Head.GetValue("u_contrno", 0)) + "' and u_pc = '" + Trim(oDBs_Head.GetValue("u_contrno", 0)) + "'")
            If oRSet.RecordCount > 1 Then
                oApplication.StatusBar.SetText("Duplicate contract no. should not be allowed", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            'Dim SubTot As Double
            'objMatrix = objForm.Items.Item("mtx").Specific
            'For i As Integer = 1 To objMatrix.VisualRowCount
            '    SubTot = SubTot + objMatrix.Columns.Item("amount").Cells.Item(i).Specific.value
            'Next
            'If SubTot > objForm.Items.Item("amount").Specific.value Then
            '    oApplication.StatusBar.SetText("Sum of amounts exceeds document level amount", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            '    Exit Function
            'End If
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
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        doc = objForm.Items.Item("docnum").Specific.Value
                        oRSet.DoQuery("Select u_status,u_unit,u_amount from [@GEN_FWD_COVER] Where DocNum = '" + doc + "'")
                        status = oRSet.Fields.Item(0).Value.ToString.Trim '
                        preamt = oRSet.Fields.Item(2).Value.ToString.Trim
                        unit = oRSet.Fields.Item(1).Value.ToString.Trim
                        If Me.Validation(FormUID) = False Then BubbleEvent = False
                    ElseIf pVal.ItemUID = "1" And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(FormUID)
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
                        If oCFL.UniqueID = "Unit1" Then
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_FWD_COVER")
                            oDBs_Head.SetValue("u_unit", 0, oDT.GetValue("Name", 0))
                        End If
                        'If oCFL.UniqueID = "PCROWS" Then
                        '    objMatrix = objForm.Items.Item("mtx").Specific
                        '    oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '    oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '    Dim Total As Double = 0
                        '    oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_FWD_COVER_D0")
                        '    Dim Flag As Boolean = False
                        '    Dim errflag As Boolean = False
                        '    If objMatrix.VisualRowCount = 1 Or pVal.Row = objMatrix.VisualRowCount Then
                        '        Flag = True
                        '    End If
                        '    For i As Integer = 0 To oDT.Rows.Count - 1
                        '        Dim cflSelectedcount As Integer = oDT.Rows.Count
                        '        If i < cflSelectedcount - 1 Then
                        '            objMatrix.AddRow(1, pVal.Row)
                        '            oDBs_Detail.InsertRecord(pVal.Row + i - 1)
                        '        End If
                        '        oDBs_Detail.Offset = pVal.Row - 1 + i
                        '        oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                        '        oDBs_Detail.SetValue("u_pc", oDBs_Detail.Offset, oDT.GetValue("Code", i))
                        '        objMatrix.SetLineData(pVal.Row + i)
                        '        objForm.EnableMenu("1293", True)
                        '    Next
                        '    If Flag = True Then
                        '        objMatrix.AddRow(1, objMatrix.VisualRowCount)
                        '        Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
                        '    End If
                        'End If
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
                    Case "GEN_FWD_COVER"
                        Me.CreateForm(objForm.UniqueID)
                    Case "1282"
                        If objForm.TypeEx = "GEN_FWD_COVER" Then
                            Me.SetDefault(objForm.UniqueID)
                            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        End If
                    Case "1281"
                        If objForm.TypeEx = "GEN_FWD_COVER" Then
                            objForm.EnableMenu("1282", True)
                            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        End If
                    Case "1288", "1289", "1290", "1291"
                        If objForm.TypeEx = "GEN_FWD_COVER" Then
                            objForm.EnableMenu("1282", True)
                        End If
                        'Case "1293"
                        '    If objForm.TypeEx = "GEN_FWD_COVER" Then
                        '        If ITEM_ID.Equals("mtx") = True Then
                        '            Dim Total As Double
                        '            objMatrix = objForm.Items.Item("mtx").Specific
                        '            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_FWD_COVER")
                        '            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_FWD_COVER_D0")
                        '            For Row As Integer = 1 To objMatrix.VisualRowCount
                        '                objMatrix.GetLineData(Row)
                        '                oDBs_Detail.Offset = Row - 1
                        '                oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, Row)
                        '                oDBs_Detail.SetValue("u_pc", oDBs_Detail.Offset, objMatrix.Columns.Item("unit").Cells.Item(Row).Specific.value)
                        '                oDBs_Detail.SetValue("u_docdate", oDBs_Detail.Offset, objMatrix.Columns.Item("docdate").Cells.Item(Row).Specific.value)
                        '                oDBs_Detail.SetValue("u_amount", oDBs_Detail.Offset, objMatrix.Columns.Item("amount").Cells.Item(Row).Specific.value)
                        '                oDBs_Detail.SetValue("u_remarks", oDBs_Detail.Offset, objMatrix.Columns.Item("remarks").Cells.Item(Row).Specific.value)
                        '                objMatrix.SetLineData(Row)
                        '            Next
                        '            objMatrix.FlushToDataSource()
                        '            oDBs_Detail.RemoveRecord(oDBs_Detail.Size - 1)
                        '            objMatrix.LoadFromDataSource()
                        '        End If
                        '    End If
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
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_FWD_COVER")
                        If Trim(oDBs_Head.GetValue("U_status", 0)) = "Encash" Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                        End If
                        If Trim(oDBs_Head.GetValue("U_status", 0)) = "Cancelled" Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                        End If
                        If Trim(oDBs_Head.GetValue("U_status", 0)) = "Open" Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.BeforeAction = False Then
                        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim year As Integer
                        Dim actualFC As Double
                        Dim blendrate As Double
                        Dim actual As Double
                        Dim month As String
                        Dim fdate, tdate As Date
                        Dim DocEntry, UNIT As String
                        Dim UserSign As String
                        Dim curr As String
                        Dim cnt As Integer
                        Dim cd As String
                        oRSet.DoQuery("Select UserID From OUSR Where User_Code = '" + oCompany.UserName.ToString.Trim + "'")
                        UserSign = oRSet.Fields.Item("UserID").Value
                        oRSet.DoQuery("Select Max(DocEntry) As 'DocEntry' From [@GEN_FWD_COVER] Where UserSign = '" + UserSign + "'")
                        DocEntry = oRSet.Fields.Item("DocEntry").Value
                        oRSet.DoQuery("Select U_docdate,U_fdate,U_tdate,U_doccur,U_amount,U_docrate,U_unit,U_status from [@GEN_FWD_COVER] Where DocEntry = '" + DocEntry + "'")
                        fdate = oRSet.Fields.Item(1).Value
                        tdate = oRSet.Fields.Item(2).Value
                        month = tdate.Month
                        year = tdate.Year
                        curr = oRSet.Fields.Item(3).Value
                        UNIT = oRSet.Fields.Item(6).Value
                        oRecordSet.DoQuery("Select u_actfc,u_act,u_blend,u_unit from [@UBG_FWD_REM] Where u_unit= '" + UNIT + "' and u_month = '" + month + "'")
                        actualFC = oRecordSet.Fields.Item(0).Value + oRSet.Fields.Item(4).Value
                        actual = oRecordSet.Fields.Item(1).Value + oRSet.Fields.Item(4).Value * oRSet.Fields.Item(5).Value
                        blendrate = (actual / actualFC)
                        cnt = oRecordSet.RecordCount
                        cd = MonthName(month, True) & Right(UNIT, 1) & "-" & Right(year, 2) & Left(curr, 1).ToUpper
                        If cnt = 0 Then
                            oRecordSet.DoQuery("INSERT INTO [@UBG_FWD_REM] (Code,Name,u_year,u_month,u_curr,u_unit,u_actfc,u_act,u_balfc,u_bal,u_blend,u_status) VALUES('" + cd.ToString + "','" + cd.ToString + "','" + year.ToString + "','" + month.ToString + "','" + curr.ToString + "','" + UNIT.ToString + "','" + actualFC.ToString + "','" + actual.ToString + "','" + actualFC.ToString + "','" + actual.ToString + "','" + blendrate.ToString + "','Open')")
                        Else
                            oRecordSet.DoQuery("UPDATE [@UBG_FWD_REM] SET u_year = '" + year.ToString + "',u_month = '" + month.ToString + "',u_curr = '" + curr.ToString + "',u_unit = '" + UNIT.ToString + "',u_actfc = '" + actualFC.ToString + "',u_act = '" + actual.ToString + "',u_balfc = '" + actualFC.ToString + "',u_bal ='" + actual.ToString + "',u_blend = '" + blendrate.ToString + "',u_status = 'Open' Where Code ='" + cd.ToString + "'")
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    'If BusinessObjectInfo.BeforeAction = True Then
                    '    objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    '    doc = objForm.Items.Item("docnum").Specific.Value
                    '    status = objForm.Items.Item("status").Specific.Value
                    '    preamt = objForm.Items.Item("amount").Specific.Value
                    '    unit = objForm.Items.Item("unit").Specific.Value
                    'End If
                    If BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.BeforeAction = False Then
                        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim curamt, genamt As Double
                        Dim postamtfc, postamt, cur As String
                        Dim rate, blrate As String
                        Dim actamtfc, actamt As String
                        Dim stat As String
                        Dim ref As String
                        Dim sdate As DateTime
                        cur = objForm.Items.Item("doccur").Specific.Value
                        curamt = objForm.Items.Item("amount").Specific.Value
                        rate = objForm.Items.Item("docrate").Specific.Value
                        stat = objForm.Items.Item("status").Specific.Value.ToString.Trim
                        tdt = objForm.Items.Item("tdate").Specific.Value
                        sdate = DateTime.ParseExact(tdt, "yyyyMMdd", Nothing)
                        ref = MonthName(sdate.Month, True) & Right(unit, 1) & "-" & Right(sdate.Year, 2) & Left(cur, 1).ToUpper
                        If stat = "Cancelled" Then
                            oRecordSet.DoQuery("Select u_balfc,u_bal,u_actfc,u_act from [@UBG_FWD_REM] Where Code = '" + ref + "'")
                            postamtfc = oRecordSet.Fields.Item(0).Value - curamt
                            postamt = oRecordSet.Fields.Item(1).Value - (curamt * rate)
                            actamtfc = oRecordSet.Fields.Item(2).Value - curamt
                            actamt = oRecordSet.Fields.Item(3).Value - (curamt * rate)
                            blrate = postamt / postamtfc
                            If postamtfc = 0 Then
                                oRecordSet.DoQuery("Update [@UBG_FWD_REM] SET u_status = 'Cancelled' Where Code = '" + ref + "'")
                            Else
                                oRecordSet.DoQuery("Update [@UBG_FWD_REM] SET u_actfc = '" + actamtfc.ToString + "',u_act = '" + actamt.ToString + "',u_balfc = '" + postamtfc.ToString + "',u_bal ='" + postamt.ToString + "',u_blend = '" + blrate.ToString + "' Where Code = '" + ref + "'")
                            End If

                        ElseIf stat = "Open" Then
                            If preamt > curamt Then
                                oRecordSet.DoQuery("Select u_balfc,u_bal,u_actfc,u_act from [@UBG_FWD_REM] Where Code = '" + ref + "'")
                                genamt = preamt - curamt
                                postamtfc = oRecordSet.Fields.Item(0).Value - genamt
                                postamt = oRecordSet.Fields.Item(1).Value - (genamt * rate)
                                actamtfc = oRecordSet.Fields.Item(2).Value - genamt
                                actamt = oRecordSet.Fields.Item(3).Value - (genamt * rate)
                                blrate = postamt / postamtfc
                                oRecordSet.DoQuery("Update [@UBG_FWD_REM] SET u_actfc = '" + actamtfc.ToString + "',u_act = '" + actamt.ToString + "',u_balfc = '" + postamtfc.ToString + "',u_bal ='" + postamt.ToString + "',u_blend = '" + blrate.ToString + "' Where Code = '" + ref + "'")
                            ElseIf preamt < curamt Then
                                oRecordSet.DoQuery("Select u_balfc from [@UBG_FWD_REM] Where Code = '" + ref + "'")
                                genamt = preamt - curamt
                                postamtfc = oRecordSet.Fields.Item(0).Value + genamt
                                postamt = oRecordSet.Fields.Item(1).Value + (genamt * rate)
                                actamtfc = oRecordSet.Fields.Item(2).Value + genamt
                                actamt = oRecordSet.Fields.Item(3).Value + (genamt * rate)
                                blrate = postamt / postamtfc
                                oRecordSet.DoQuery("Update [@UBG_FWD_REM] SET u_actfc = '" + actamtfc.ToString + "',u_act = '" + actamt.ToString + "',u_balfc = '" + postamtfc.ToString + "',u_bal ='" + postamt.ToString + "',u_blend = '" + blrate.ToString + "' Where Code = '" + ref + "'")
                            ElseIf preamt = curamt Then
                                oRecordSet.DoQuery("Update [@UBG_FWD_REM] SET u_status = 'Open' Where Code = '" + ref + "'")
                            End If
                        End If

                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
