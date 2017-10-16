Public Class ClsEcgc
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
            oUtilities.SAPXML("ECGCMaster.xml")
            objForm = oApplication.Forms.GetForm("GEN_ECGC", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_ECGC")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_ECGC_D0")
            objForm.DataBrowser.BrowseBy = "12"
            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            'objForm.Items.Item("name").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_ECGC_D0")
            objMatrix = oMatrix
            objMatrix.FlushToDataSource()
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, objMatrix.VisualRowCount)
            oDBs_Detail.SetValue("u_cardcode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_cardname", oDBs_Detail.Offset, "")
            
            objMatrix.SetLineData(objMatrix.VisualRowCount)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            If oApplication.Forms.Item(eventInfo.FormUID).Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oApplication.Forms.Item(eventInfo.FormUID).Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                Dim MenuItem As SAPbouiCOM.MenuItem
                Dim Menu As SAPbouiCOM.Menus
                Dim MenuParam As SAPbouiCOM.MenuCreationParams
                MenuParam = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                MenuParam.Type = SAPbouiCOM.BoMenuType.mt_STRING
                MenuParam.UniqueID = "Close"
                MenuParam.String = "Close"
                MenuParam.Enabled = True
                MenuItem = oApplication.Menus.Item("1280")
                Menu = MenuItem.SubMenus
                If MenuItem.SubMenus.Exists("Close") = False Then Menu.AddEx(MenuParam)
            Else
                ROW_ID = eventInfo.Row
                If oApplication.Menus.Exists("Close") = True Then oApplication.Menus.RemoveEx("Close")
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
        'Try
        '    If oApplication.Forms.Item(eventInfo.FormUID).Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oApplication.Forms.Item(eventInfo.FormUID).Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
        '        Dim MenuItem1 As SAPbouiCOM.MenuItem
        '        Dim Menu1 As SAPbouiCOM.Menus
        '        Dim MenuParam1 As SAPbouiCOM.MenuCreationParams
        '        MenuParam1 = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        '        MenuParam1.Type = SAPbouiCOM.BoMenuType.mt_STRING
        '        MenuParam1.UniqueID = "Cancel"
        '        MenuParam1.String = "Cancel"
        '        MenuParam1.Enabled = True
        '        MenuItem1 = oApplication.Menus.Item("1280")
        '        Menu1 = MenuItem1.SubMenus
        '        If MenuItem1.SubMenus.Exists("Cancel") = False Then
        '            Menu1.AddEx(MenuParam1)
        '        End If
        '    Else
        '        ROW_ID = eventInfo.Row
        '        If oApplication.Menus1.Exists("Cancel") = True Then oApplication.Menus1.RemoveEx("Cancel")
        '    End If
        'Catch ex As Exception
        '    oApplication.StatusBar.SetText(ex.Message)
        'End Try
        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
              ROW_ID = eventInfo.Row
                If eventInfo.Row > 0 Then
                    ITEM_ID = eventInfo.ItemUID
            Dim objMatrix As SAPbouiCOM.Matrix

            objMatrix = objForm.Items.Item("11").Specific
            If ITEM_ID.Equals("11") = True Then
                If objMatrix.VisualRowCount >= 1 Then
                    objForm.EnableMenu("1293", True)
                Else
                    objForm.EnableMenu("1293", False)
                End If
            ElseIf ITEM_ID.Equals("11") = True Then
               
            End If
                Else
            ITEM_ID = ""
        End If


    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_ECGC")
                        oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_ECGC_D0")
                        objMatrix = objForm.Items.Item("11").Specific
                        Dim chk As Int16 = 0
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        For i As Integer = 1 To objMatrix.VisualRowCount
                            If Trim(objMatrix.Columns.Item("V_1").Cells.Item(i).Specific.value) <> "" Then
                                Dim strItemCode As String = objMatrix.Columns.Item("V_1").Cells.Item(i).Specific.value
                                oRSet.DoQuery("select T0.Code  from(Select code from [dbo].[@GEN_ECGC] union all select U_cardcode from [dbo].[@GEN_ECGC_D0])T0 where T0.Code<>'' and T0.Code ='" + strItemCode.Trim() + "'")
                                If oRSet.RecordCount > 0 Then
                                    chk = 1
                                End If
                            End If
                        Next
                        If chk = 1 Then
                            BubbleEvent = False
                            oApplication.StatusBar.SetText("CardCode Already exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        End If
                        If oDBs_Head.GetValue("U_Ecgc", 0).ToString() = "0.0000" Then
                            BubbleEvent = False
                            oApplication.StatusBar.SetText("Customer ECGC should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        End If
                    ElseIf pVal.ItemUID = "1" And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    'If pVal.BeforeAction = True Then
                    '    Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '    oRecordSet.DoQuery("Select u_unit From [@GEN_USR_UNIT] Where u_user = '" + oCompany.UserName.ToString.Trim + "'")

                    '    Dim objForm1 As SAPbouiCOM.Form
                    '    objForm1 = oApplication.Forms.Item(FormUID)
                    '    Dim oCFL1 As SAPbouiCOM.ChooseFromList
                    '    Dim CFLEvent1 As SAPbouiCOM.IChooseFromListEvent = pVal
                    '    Dim CFL_Id1 As String
                    '    CFL_Id1 = CFLEvent1.ChooseFromListUID
                    '    oCFL1 = objForm1.ChooseFromLists.Item(CFL_Id1)
                    '    If oCFL1.UniqueID = "CFL_2" Then
                    '        Me.FilterCustomers(FormUID)
                    '    End If
                    'End If

                    Dim objForm As SAPbouiCOM.Form
                    objForm = oApplication.Forms.Item(FormUID)
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    'Dim pp As Int16 = oDT.Rows.Count
                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_ECGC")
                        oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_ECGC_D0")
                        If oCFL.UniqueID = "CFL_2" Then
                            oDBs_Head.SetValue("U_cardcode", 0, oDT.GetValue("CardCode", 0))
                            oDBs_Head.SetValue("Code", 0, oDT.GetValue("CardCode", 0))
                            oDBs_Head.SetValue("U_cardname", 0, oDT.GetValue("CardName", 0))

                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select u_unit,Currency From OCRD Where CardCode = '" + Trim(oDT.GetValue("CardCode", 0)) + "'")
                            Dim unit As String = oRSet.Fields.Item("u_unit").Value
                            'objForm.Items.Item("unit").Specific.value = unit
                            oDBs_Head.SetValue("U_Unit", 0, unit)
                            objMatrix = objForm.Items.Item("11").Specific
                            objMatrix.Clear()
                            objMatrix.AddRow()
                            objMatrix.FlushToDataSource()
                            Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
                        End If
                        If oCFL.UniqueID = "CFL_3" Then
                            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim oRecSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            objMatrix = objForm.Items.Item("11").Specific
                            Dim OrginRow As Integer = objMatrix.VisualRowCount
                            For i As Integer = 0 To oDT.Rows.Count - 1
                                Dim cflSelectedcount As Integer = oDT.Rows.Count
                                If i < cflSelectedcount - 1 Then
                                    objMatrix.AddRow(1, pVal.Row)
                                    oDBs_Detail.InsertRecord(pVal.Row + i - 1)
                                End If
                                Dim cr As String

                                oDBs_Detail.Offset = pVal.Row - 1 + i
                                oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, i + pVal.Row)
                                oDBs_Detail.SetValue("U_cardcode", oDBs_Detail.Offset, oDT.GetValue("CardCode", i))
                                oDBs_Detail.SetValue("U_cardname", oDBs_Detail.Offset, oDT.GetValue("CardName", i))
                                oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, oDT.GetValue("U_unit", i))
                                objMatrix.SetLineData(pVal.Row + i)
                            Next
                            objMatrix.FlushToDataSource()

                            If objMatrix.VisualRowCount = pVal.Row Then
                                objMatrix.AddRow()
                                objMatrix.FlushToDataSource()
                                Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
                            End If


                            objMatrix.AutoResizeColumns()

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
                    Case "GEN_ECGC"
                        Me.CreateForm(objForm.UniqueID)
                    Case "1282"
                        If objForm.TypeEx = "GEN_ECGC" Then
                           
                        End If
                    Case "1281"
                        If objForm.TypeEx = "GEN_ECGC" Then
                            objForm.EnableMenu("1282", True)
                        End If
                    Case "1288", "1289", "1290", "1291"
                        If objForm.TypeEx = "GEN_ECGC" Then
                            objForm.EnableMenu("1282", True)
                        End If
                    Case "1293"
                        If objForm.TypeEx = "GEN_ECGC" Then
                            If ITEM_ID.Equals("11") = True Then
                                objMatrix = objForm.Items.Item("11").Specific
                                oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_ECGC")
                                oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_ECGC_D0")
                                For Row As Integer = 1 To objMatrix.VisualRowCount
                                    objMatrix.GetLineData(Row)
                                    oDBs_Detail.Offset = Row - 1
                                    oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, Row)
                                    oDBs_Detail.SetValue("U_cardcode", oDBs_Detail.Offset, objMatrix.Columns.Item("V_1").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("U_cardname", oDBs_Detail.Offset, objMatrix.Columns.Item("V_0").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, objMatrix.Columns.Item("V_2").Cells.Item(Row).Specific.value)

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

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_ECGC")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_ECGC_D0")
            objMatrix = objForm.Items.Item("11").Specific
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
            oDBs_Detail.SetValue("U_cardcode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_cardname", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, "")
            objMatrix.SetLineData(Row)
            objMatrix.FlushToDataSource()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterCustomers(ByVal FormUID As String)
        Try
            Dim objForm As SAPbouiCOM.Form
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("CFL_2")
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Distinct U_unit From [@GEN_USR_UNIT] Where U_user = '" + oCompany.UserName.ToString.Trim + "'")
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
End Class
