Public Class ClsInventoryTransfer

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oItem As SAPbouiCOM.Item
    Dim oTempItem As SAPbouiCOM.Item
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim oRS As SAPbobsCOM.Recordset
    Dim oRSet As SAPbobsCOM.Recordset
    Dim User_Code As String
    Dim DocEntry As String
    Dim PurType As String
    Public MRNo As String
    Dim TransNo As String
    Dim NewPrice As Double
    Dim DocNO As String
    Dim DOCNUM As String = ""
    Dim FrmWhs As String
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            objForm.Title = "Material Issue Note"
            oTempItem = objForm.Items.Item("16")
            oItem = objForm.Items.Add("type", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OWTR", "u_type")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("mrnno", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OWTR", "u_mrnno")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("subconno", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OWTR", "u_subconno")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("isstyp", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OWTR", "u_isstyp")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("subretno", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OWTR", "u_subretno")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("sono", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OWTR", "u_sono")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("itemcode", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OWTR", "u_itemcode")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("sfgcode", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OWTR", "u_sfgcode")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("grnno", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OWTR", "u_grnno")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("scpono", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OWTR", "u_DocNum")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("scpotp", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OWTR", "u_Type")
            oItem.Visible = False
            oItem.LinkTo = "16"
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetChooseFromList(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = True
            oCFLCreationParams.ObjectType = "GEN_MREQ"
            oCFLCreationParams.UniqueID = "MREQ"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetConditionToMREQ(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("MREQ")
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Open"
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                    If pVal.BeforeAction = False Then
                        If Trim(DOCNUM).Equals("") = False Then
                            objForm = oApplication.Forms.Item(FormUID)
                            objForm.Freeze(True)
                            Me.LoadItems(FormUID, DOCNUM)
                            DOCNUM = ""
                            objForm.Freeze(False)
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    If pVal.BeforeAction = True Then
                        Me.CreateForm(pVal.FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        DocNO = objForm.Items.Item("11").Specific.Value
                        FrmWhs = objForm.Items.Item("18").Specific.value
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRSet.DoQuery("Select WhsCode From OWHS Where WhsCode = '" + FrmWhs + "' And IsNull(u_inspwhs,'NO') = 'YES'")
                        If oRSet.RecordCount > 0 Then
                            If Trim(objForm.Items.Item("grnno").Specific.value) = "" Then
                                oApplication.StatusBar.SetText("Copy from GRPO to move items from inspection warehouse", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                        If Trim(objForm.Items.Item("grnno").Specific.value) <> "" Then

                            Dim oMatrix As SAPbouiCOM.Matrix
                            Dim ld_TotalQty As Decimal

                            oMatrix = objForm.Items.Item("23").Specific
                            For i As Integer = 1 To oMatrix.VisualRowCount - 1
                                If oMatrix.Columns.Item("U_grpostat").Cells.Item(i).Specific.selected Is Nothing = True Then
                                    oApplication.StatusBar.SetText("GRPO Status can not be blank.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                If Trim(oMatrix.Columns.Item("U_grpostat").Cells.Item(i).Specific.selected.Value) = "" Then
                                    oApplication.StatusBar.SetText("GRPO Status can not be blank.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                            Next
                        End If

                        'Vijeesh
                        If Trim(objForm.Items.Item("mrnno").Specific.value) <> "" Then
                            Dim oMatrix As SAPbouiCOM.Matrix
                            oMatrix = objForm.Items.Item("23").Specific
                            For i As Integer = 1 To oMatrix.VisualRowCount - 1
                                If oMatrix.Columns.Item("1").Cells.Item(i).Specific.value <> "" Then
                                    If (oMatrix.Columns.Item("10").Cells.Item(i).Specific.value > (oMatrix.Columns.Item("U_rqstqty").Cells.Item(i).Specific.value - oMatrix.Columns.Item("U_issued").Cells.Item(i).Specific.value)) Then
                                        oApplication.StatusBar.SetText("Issueing Quantity Exceeds the MRN Quantity of Item Code --> " + oMatrix.Columns.Item("1").Cells.Item(i).Specific.value + " ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Next
                        End If
                        'Vijeesh

                    End If
                    ''Added By Vivek for validation

                    'Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                    '    If pVal.Before_Action = True Then
                    '        Dim oMatrix As SAPbouiCOM.Matrix
                    '        Dim ld_TotalQty As Decimal
                    '        If pVal.CharPressed = "9" And pVal.ColUID = "10" Then


                    '            oMatrix = objForm.Items.Item("23").Specific
                    '            For i As Integer = 1 To oMatrix.VisualRowCount
                    '                If oMatrix.Columns.Item("U_grnlnid").Cells.Item(pVal.Row).Specific.value = oMatrix.Columns.Item("U_grnlnid").Cells.Item(i).Specific.value Then
                    '                    ld_TotalQty += oMatrix.Columns.Item("10").Cells.Item(i).Specific.value
                    '                    If ld_TotalQty > oMatrix.Columns.Item("U_BAL_QTY").Cells.Item(i).Specific.value Then
                    '                        oApplication.StatusBar.SetText("Quantity exceeds Balance quantity.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '                        BubbleEvent = False
                    '                        Exit Sub
                    '                    End If
                    '                End If
                    '            Next
                    '        End If
                    '    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    objForm = oApplication.Forms.Item(FormUID)
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If oCFL.UniqueID = "MREQ" Then
                            If oDT.Rows.Count > 1 Then
                                oApplication.StatusBar.SetText("Cannot select more than one document", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            DOCNUM = Trim(oDT.GetValue("DocNum", 0))
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
                        If objForm.TypeEx = "940" Then
                            BubbleEvent = False
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
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.ActionSuccess = True Then
                        Dim oITForm As SAPbouiCOM.Form = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        Dim oMatrix As SAPbouiCOM.Matrix
                        oMatrix = oITForm.Items.Item("23").Specific
                        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim MRNEntry As String
                        If Trim(oITForm.Items.Item("mrnno").Specific.Value) <> "" Then
                            If Trim(oITForm.Items.Item("isstyp").Specific.value) = "I" Then
                                oRecordSet.DoQuery("Select DocEntry From [@GEN_MREQ] Where DocNum = '" + Trim(oITForm.Items.Item("mrnno").Specific.value) + "'")
                                MRNEntry = oRecordSet.Fields.Item("DocEntry").Value
                                For i As Integer = 1 To oMatrix.VisualRowCount
                                    oRecordSet.DoQuery("Update [@GEN_MREQ_D0] Set u_issued = u_issued + Convert(Money,'" + Trim(oMatrix.Columns.Item("10").Cells.Item(i).Specific.value) + "') Where DocEntry = '" + MRNEntry + "' And LineId = '" + Trim(oMatrix.Columns.Item("U_mrnlid").Cells.Item(i).Specific.Value) + "'")
                                    oRecordSet.DoQuery("Update [@GEN_MREQ_D0] Set u_stat = 'Closed' Where DocEntry = '" + MRNEntry + "' And LineId = '" + Trim(oMatrix.Columns.Item("U_mrnlid").Cells.Item(i).Specific.Value) + "' And u_issued >= u_rqstqty")
                                Next
                                oRecordSet.DoQuery("Update [@GEN_MREQ] Set u_status = 'Closed' Where DocEntry = '" + MRNEntry + "' And DocEntry Not in (Select DocEntry From [@GEN_MREQ_D0] Where DocEntry = '" + MRNEntry + "' And IsNull(u_stat,'N') != 'Closed')")
                            End If
                            If Trim(oITForm.Items.Item("isstyp").Specific.value) = "R" Then
                                oRecordSet.DoQuery("Select DocEntry From [@GEN_MREQ] Where DocNum = '" + Trim(oITForm.Items.Item("mrnno").Specific.value) + "'")
                                MRNEntry = oRecordSet.Fields.Item("DocEntry").Value
                                For i As Integer = 1 To oMatrix.VisualRowCount
                                    oRecordSet.DoQuery("Update [@GEN_MREQ_D0] Set u_returned = IsNull(u_returned,0) + Convert(Money,'" + Trim(oMatrix.Columns.Item("10").Cells.Item(i).Specific.value) + "') Where DocEntry = '" + MRNEntry + "' And LineId = '" + Trim(oMatrix.Columns.Item("U_mrnlid").Cells.Item(i).Specific.Value) + "'")
                                Next
                            End If
                        End If
                        If Trim(oITForm.Items.Item("grnno").Specific.Value) <> "" Then
                            oRecordSet.DoQuery("Select DocEntry From OPDN Where DocNum = '" + Trim(oITForm.Items.Item("grnno").Specific.value) + "'")
                            MRNEntry = oRecordSet.Fields.Item("DocEntry").Value
                            For i As Integer = 1 To oMatrix.VisualRowCount
                                If Trim(oMatrix.Columns.Item("1").Cells.Item(i).Specific.value) <> "" Then
                                    oRecordSet.DoQuery("Update PDN1 Set u_accqty = Isnull(u_accqty,0) + Convert(Money,'" + Trim(oMatrix.Columns.Item("10").Cells.Item(i).Specific.value) + "') Where DocEntry = '" + MRNEntry + "' And LineNum = '" + Trim(oMatrix.Columns.Item("U_grnlnid").Cells.Item(i).Specific.Value) + "' And '" + Trim(oMatrix.Columns.Item("U_grpostat").Cells.Item(i).Specific.selected.Value) + "' = 'Accepted'")
                                    oRecordSet.DoQuery("Update PDN1 Set u_rejqty = Isnull(u_rejqty,0) + Convert(Money,'" + Trim(oMatrix.Columns.Item("10").Cells.Item(i).Specific.value) + "') Where DocEntry = '" + MRNEntry + "' And LineNum = '" + Trim(oMatrix.Columns.Item("U_grnlnid").Cells.Item(i).Specific.Value) + "' And '" + Trim(oMatrix.Columns.Item("U_grpostat").Cells.Item(i).Specific.selected.Value) + "' = 'Rejected'")
                                    oRecordSet.DoQuery("Update PDN1 Set u_shqty = Isnull(u_shqty,0) + Convert(Money,'" + Trim(oMatrix.Columns.Item("10").Cells.Item(i).Specific.value) + "') Where DocEntry = '" + MRNEntry + "' And LineNum = '" + Trim(oMatrix.Columns.Item("U_grnlnid").Cells.Item(i).Specific.Value) + "' And '" + Trim(oMatrix.Columns.Item("U_grpostat").Cells.Item(i).Specific.selected.Value) + "' = 'Shortage'")
                                    oRecordSet.DoQuery("Update PDN1 Set u_exqty = Isnull(u_exqty,0) + Convert(Money,'" + Trim(oMatrix.Columns.Item("10").Cells.Item(i).Specific.value) + "') Where DocEntry = '" + MRNEntry + "' And LineNum = '" + Trim(oMatrix.Columns.Item("U_grnlnid").Cells.Item(i).Specific.Value) + "' And '" + Trim(oMatrix.Columns.Item("U_grpostat").Cells.Item(i).Specific.selected.Value) + "' = 'Excess'")
                                    oRecordSet.DoQuery("Update PDN1 Set u_openqty = Isnull(u_openqty,0) + Convert(Money,'" + Trim(oMatrix.Columns.Item("10").Cells.Item(i).Specific.value) + "') Where DocEntry = '" + MRNEntry + "' And LineNum = '" + Trim(oMatrix.Columns.Item("U_grnlnid").Cells.Item(i).Specific.Value) + "'")
                                    oRecordSet.DoQuery("Update PDN1 Set u_insstat = 'Closed' Where DocEntry = '" + MRNEntry + "' And LineNum = '" + Trim(oMatrix.Columns.Item("U_grnlnid").Cells.Item(i).Specific.Value) + "' And u_openqty >= (quantity * NumperMsr)")
                                End If
                            Next
                            oRecordSet.DoQuery("Update OPDN Set u_insstat = 'Closed' Where DocEntry = '" + MRNEntry + "' And DocEntry Not in (Select DocEntry From [PDN1] Where DocEntry = '" + MRNEntry + "' And IsNull(u_insstat,'Open') != 'Closed')")
                        End If
                        If Trim(oITForm.Items.Item("scpono").Specific.Value) <> "" Then
                            Dim SCEntry, SCNo As String
                            oRecordSet.DoQuery("Select U_SCNo,DocEntry From [@GEN_SC_DC] Where DocNum = '" + Trim(oITForm.Items.Item("scpono").Specific.value) + "'")
                            MRNEntry = oRecordSet.Fields.Item("DocEntry").Value
                            SCNo = oRecordSet.Fields.Item("U_SCNo").Value
                            oRecordSet.DoQuery("Select DocEntry From [@GEN_SUB_CONTRACT] Where DocNum = '" + SCNo + "'")
                            SCEntry = oRecordSet.Fields.Item("DocEntry").Value
                            For i As Integer = 1 To oMatrix.VisualRowCount
                                oRecordSet.DoQuery("Update [@GEN_SC_DC_D0] Set u_CompQty = IsNull(u_CompQty,0) + Convert(Money,'" + Trim(oMatrix.Columns.Item("10").Cells.Item(i).Specific.value) + "') Where DocEntry = '" + MRNEntry + "' And LineId = '" + Trim(oMatrix.Columns.Item("U_scpoln").Cells.Item(i).Specific.Value) + "'")
                                oRecordSet.DoQuery("Update [@GEN_SUB_CONTRACT_D1] Set u_DCQty = IsNull(u_DCQty,0) + Convert(Money,'" + Trim(oMatrix.Columns.Item("10").Cells.Item(i).Specific.value) + "') Where DocEntry = '" + SCEntry + "' And u_Code = '" + Trim(oMatrix.Columns.Item("1").Cells.Item(i).Specific.value) + "'")
                            Next
                        End If

                        'oApplication.Forms.Item("GEN_MREQ").Close()
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadItems(ByVal FormUID As String, ByVal MreqNo As String)
        Try
            Dim ITForm As SAPbouiCOM.Form
            Dim ITMatrix As SAPbouiCOM.Matrix
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select A.u_type,A.u_sono,A.u_itemcode,A.u_sfgcode,A.u_wipwhs,A.DocNum,B.u_itemcode as 'RItemCode' ,B.u_rqstqty- B.u_issued as 'rqstqty',B.u_rqstqty,B.u_issued,B.u_whs,B.lineid From [@GEN_MREQ] A INNER JOIN [@GEN_MREQ_D0] B On A.DocEntry = B.DocEntry And A.Docnum = '" + MreqNo + "'")
            ITForm = oApplication.Forms.Item(FormUID)
            Try
                ITForm.Freeze(True)
                ITMatrix = ITForm.Items.Item("23").Specific
                ITForm.Items.Item("18").Specific.Value = oRecordSet.Fields.Item("u_whs").Value
                ITForm.Items.Item("sono").Specific.value = oRecordSet.Fields.Item("u_sono").Value
                ITForm.Items.Item("type").Specific.value = oRecordSet.Fields.Item("u_type").Value
                ITForm.Items.Item("mrnno").Specific.value = oRecordSet.Fields.Item("DocNum").Value
                ITForm.Items.Item("itemcode").Specific.value = oRecordSet.Fields.Item("u_itemcode").Value
                ITForm.Items.Item("sfgcode").Specific.value = oRecordSet.Fields.Item("u_sfgcode").Value
                ITMatrix.Columns.Item("U_mrnlid").Editable = True
                ITMatrix.Columns.Item("U_rqstqty").Editable = True
                ITMatrix.Columns.Item("U_issued").Editable = True
                For i As Integer = 1 To oRecordSet.RecordCount
                    ITMatrix.Columns.Item("1").Cells.Item(i).Specific.value = oRecordSet.Fields.Item("RItemCode").Value
                    ITMatrix.Columns.Item("5").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("u_wipwhs").Value
                    ITMatrix.Columns.Item("10").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("rqstqty").Value
                    ITMatrix.Columns.Item("U_BAL_QTY").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("rqstqty").Value
                    ITMatrix.Columns.Item("U_mrnlid").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("lineid").Value
                    ITMatrix.Columns.Item("U_rqstqty").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("u_rqstqty").Value
                    ITMatrix.Columns.Item("U_issued").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("u_issued").Value
                    ITMatrix.Columns.Item("1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    oRecordSet.MoveNext()
                Next
                ITMatrix.Columns.Item("U_mrnlid").Editable = False
                ITMatrix.Columns.Item("U_rqstqty").Editable = False
                ITMatrix.Columns.Item("U_issued").Editable = False
                ITForm.Freeze(False)
            Catch ex As Exception
                ITForm.Freeze(False)
                oApplication.StatusBar.SetText(ex.Message)
            End Try

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

End Class
