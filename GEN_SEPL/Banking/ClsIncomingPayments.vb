Public Class ClsIncomingPayments


#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm, objSubForm As SAPbouiCOM.Form
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim objItem, TempItem As SAPbouiCOM.Item
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim optnBtn, optnBtn1, optnBtn2 As SAPbouiCOM.OptionBtn
    Dim PARENT_FORM As String
    Dim oRS As SAPbobsCOM.Recordset
    Dim oRSet As SAPbobsCOM.Recordset
    Dim ROW_ID As Integer = 0
    Dim ITEM_ID As String
    Dim RowCount As Integer
    Dim enableflag As Boolean = False
    Dim ModalForm As Boolean = False
    Dim docno As String
    Dim docdt As Date
    Dim pfc, docur As String
    Dim PAYNUM As String
    Dim DocDate, RefDate, JVRem As String
    Dim CurrSpot As Double
    Dim refs As String
    Dim yr, mnth As String
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            TempItem = objForm.Items.Item("96")
            objItem = objForm.Items.Add("sbc", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = TempItem.Top + 30
            objItem.Left = TempItem.Left
            objItem.Height = TempItem.Height
            objItem.Width = TempItem.Width
            objItem.LinkTo = "96"
            objItem.Specific.Caption = "Other Charges - JV"
            TempItem = objForm.Items.Item("95")
            objItem = objForm.Items.Add("ebc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = TempItem.Top + 30
            objItem.Left = TempItem.Left
            objItem.Height = TempItem.Height
            objItem.Width = TempItem.Width
            objItem.Specific.Databind.setbound(True, "ORCT", "u_bcjv")
            objItem.Specific.taborder = TempItem.Specific.taborder - 1
            objItem.LinkTo = "95"
            objItem = objForm.Items.Add("lbc", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            Dim LB As SAPbouiCOM.LinkedButton
            LB = objItem.Specific
            LB.LinkedObjectType = "30"
            objItem.Top = TempItem.Top + 30
            objItem.Left = TempItem.Left - 20
            objItem.Width = 20
            objItem.Height = 14
            objItem.LinkTo = "ebc"
            TempItem = objForm.Items.Item("2")
            objItem = objForm.Items.Add("btn", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Top = TempItem.Top
            objItem.Left = TempItem.Left + TempItem.Width + 5
            objItem.Height = TempItem.Height
            objItem.Width = TempItem.Width + 30
            objItem.Specific.caption = "Other Charges"
            objItem.LinkTo = "2"
            ''Rajkumar ----- Forward Cover------- 26.08.14
            'TempItem = objForm.Items.Item("151")
            'objItem = objForm.Items.Add("rts", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'objItem.Top = TempItem.Top + 30
            'objItem.Left = TempItem.Left
            'objItem.Height = TempItem.Height
            'objItem.Width = TempItem.Width
            'objItem.LinkTo = "151"
            'objItem.Specific.Caption = "Rate Type"
            'TempItem = objForm.Items.Item("152")
            'objItem = objForm.Items.Add("rte", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            'objItem.Top = TempItem.Top + 30
            'objItem.Left = TempItem.Left
            'objItem.Height = TempItem.Height
            'objItem.Width = TempItem.Width
            'objItem.Specific.Databind.setbound(True, "ORCT", "U_Frgn")
            'objItem.Specific.taborder = TempItem.Specific.taborder - 1
            'objItem.LinkTo = "152"
            'objForm.Items.Item("btn").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("ebc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            TempItem = objForm.Items.Item("sbc")
            objItem = objForm.Items.Add("cps", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = TempItem.Top + 45
            objItem.Left = TempItem.Left
            objItem.Height = TempItem.Height
            objItem.Width = TempItem.Width
            objItem.LinkTo = "sbc"
            objItem.Specific.Caption = "Unit"
            TempItem = objForm.Items.Item("ebc")
            objItem = objForm.Items.Add("cpc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = TempItem.Top + 45
            objItem.Left = TempItem.Left
            objItem.Height = TempItem.Height
            objItem.Width = TempItem.Width
            objItem.Specific.Databind.setbound(True, "ORCT", "u_unit")
            objItem.Specific.taborder = TempItem.Specific.taborder - 1
            objItem.LinkTo = "ebc"

            'Rajkumar ----- Forward Cover------- 26.08.14
            TempItem = objForm.Items.Item("151")
            objItem = objForm.Items.Add("rts", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = TempItem.Top + 45
            objItem.Left = TempItem.Left
            objItem.Height = TempItem.Height
            objItem.Width = TempItem.Width
            objItem.LinkTo = "151"
            objItem.Specific.Caption = "Rate Type"
            TempItem = objForm.Items.Item("152")
            objItem = objForm.Items.Add("rte", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            objItem.Top = TempItem.Top + 45
            objItem.Left = TempItem.Left
            objItem.Height = TempItem.Height
            objItem.Width = TempItem.Width
            objItem.Specific.Databind.setbound(True, "ORCT", "U_Frgn")
            objItem.Specific.taborder = TempItem.Specific.taborder - 1
            objItem.LinkTo = "152"

            TempItem = objForm.Items.Item("rts")
            objItem = objForm.Items.Add("cntrnos", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = TempItem.Top + 15
            objItem.Left = TempItem.Left
            objItem.Height = TempItem.Height
            objItem.Width = TempItem.Width
            objItem.LinkTo = "rts"
            objItem.Specific.Caption = "Contract No."
            TempItem = objForm.Items.Item("rte")
            objItem = objForm.Items.Add("cntrnoe", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = TempItem.Top + 15
            objItem.Left = TempItem.Left
            objItem.Height = TempItem.Height
            objItem.Width = TempItem.Width
            objItem.Specific.Databind.setbound(True, "ORCT", "u_contrno")
            objItem.Specific.taborder = TempItem.Specific.taborder - 1
            objItem.LinkTo = "rts"
            Me.SetChooseFromList(FormUID)
            objItem.Specific.ChooseFromListUID = "FRDCFL"
            objItem.Specific.ChooseFromListAlias = "U_contrno"

            TempItem = objForm.Items.Item("cntrnos")
            objItem = objForm.Items.Add("docrts", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = TempItem.Top + 15
            objItem.Left = TempItem.Left
            objItem.Height = TempItem.Height
            objItem.Width = TempItem.Width
            objItem.LinkTo = "cntrnos"
            objItem.Specific.Caption = "Doc Rate"
            TempItem = objForm.Items.Item("cntrnoe")
            objItem = objForm.Items.Add("docrte", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = TempItem.Top + 15
            objItem.Left = TempItem.Left
            objItem.Height = TempItem.Height
            objItem.Width = TempItem.Width
            objItem.Specific.Databind.setbound(True, "ORCT", "u_docrate")
            objItem.Specific.taborder = TempItem.Specific.taborder - 1
            objItem.LinkTo = "cntrnoe"

            TempItem = objForm.Items.Item("docrts")
            objItem = objForm.Items.Add("balamts", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = TempItem.Top + 15
            objItem.Left = TempItem.Left
            objItem.Height = TempItem.Height
            objItem.Width = TempItem.Width
            objItem.LinkTo = "docrts"
            objItem.Specific.Caption = "Balance Amt"

            TempItem = objForm.Items.Item("docrte")
            objItem = objForm.Items.Add("balamte", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = TempItem.Top + 15
            objItem.Left = TempItem.Left
            objItem.Height = TempItem.Height
            objItem.Width = TempItem.Width
            objItem.Specific.Databind.setbound(True, "ORCT", "u_balamt")
            objItem.Specific.taborder = TempItem.Specific.taborder - 1
            objItem.LinkTo = "docrte"

            objForm.Items.Item("btn").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("ebc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("docrte").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("balamte").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("rts").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("rte").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'objForm.Items.Item("rts").Visible = False
            'objForm.Items.Item("rte").Visible = False

            visualbehaviour("Spot")
            optnBtn = objForm.Items.Item("58").Specific
            If optnBtn.Selected = True Then
                visualbehaviour(objForm.Items.Item("rte").Specific.Value)
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                objForm = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "1288", "1289", "1290", "1291"
                        If objForm.TypeEx = "170" Then
                            If objForm.Items.Item("rte").Specific.Value.ToString.Trim = "FWD Cover" Then
                                objForm.Items.Item("cntrnos").Visible = False
                                objForm.Items.Item("cntrnoe").Visible = False
                                objForm.Items.Item("docrts").Visible = False
                                objForm.Items.Item("balamts").Visible = False
                                objForm.Items.Item("docrte").Visible = False
                                objForm.Items.Item("balamte").Visible = False
                            Else
                                objForm.Items.Item("cntrnos").Visible = True
                                objForm.Items.Item("cntrnoe").Visible = True
                                objForm.Items.Item("docrts").Visible = True
                                objForm.Items.Item("balamts").Visible = True
                                objForm.Items.Item("docrte").Visible = True
                                objForm.Items.Item("balamte").Visible = True
                            End If
                        End If
                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Dim CHILD_FORM As String = "ORCT_JV@" & pVal.FormUID
            Dim ModalForm As Boolean = False
            For i As Integer = 0 To oApplication.Forms.Count - 1
                If oApplication.Forms.Item(i).UniqueID = (CHILD_FORM) Then
                    objSubForm = oApplication.Forms.Item("ORCT_JV@" & pVal.FormUID)
                    ModalForm = True
                    Exit For
                End If
            Next
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                    objMatrix = objForm.Items.Item("71").Specific
                    PARENT_FORM = (pVal.FormUID.Substring(objForm.UniqueID.IndexOf("@") + 1))
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    If pVal.BeforeAction = True Then
                        If pVal.FormTypeCount = 1 Then
                            Me.CreateForm(FormUID)
                        Else
                            BubbleEvent = False
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    If pVal.ItemUID = "21" And (pVal.BeforeAction = False And pVal.ActionSuccess = True) And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        'If pVal.CharPressed = 53 Or pVal.CharPressed = 54 Or pVal.CharPressed = 55 Then
                        If pVal.CharPressed = 9 Then
                            objCombo = objForm.Items.Item("rte").Specific
                            If objCombo.Selected.Value = "Spot" Then
                                Dim j, actual, minadjust, maxadjust As Integer
                                j = CurrSpot * 10 / 100
                                actual = objForm.Items.Item("21").Specific.Value
                                maxadjust = CurrSpot + j
                                minadjust = CurrSpot - j
                                If Not (actual <= maxadjust And actual >= minadjust) Then
                                    objForm.Items.Item("21").Specific.Value = CurrSpot
                                    oApplication.StatusBar.SetText("Currency rate exceed 10 percent from system rate", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                            End If
                            'End If
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "rte" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then 'And pVal.InnerEvent = True
                        If pVal.ActionSuccess = True Then
                            objForm = oApplication.Forms.Item(FormUID)
                            objCombo = objForm.Items.Item("rte").Specific
                            optnBtn = objForm.Items.Item("58").Specific
                            docdt = DateTime.ParseExact(objForm.Items.Item("10").Specific.Value, "yyyyMMdd", Nothing)
                            pfc = objForm.Items.Item("cpc").Specific.value
                            docur = objForm.Items.Item("41").Specific.Value
                            If optnBtn.Selected = True Then
                                If objCombo.Selected.Value = "FWD Cover" Then
                                    visualbehaviour("FWD Cover")
                                Else
                                    visualbehaviour("Spot")
                                    objForm.Items.Item("docrte").Specific.value = ""
                                    objForm.Items.Item("balamte").Specific.value = ""
                                    objForm.Items.Item("cntrnoe").Specific.value = ""
                                End If
                            Else
                                If objCombo.Selected.Value = "FWD Cover" Then
                                    If objForm.Items.Item("5").Specific.Value <> "" Then
                                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        If objForm.Items.Item("cpc").Specific.value <> "" Then
                                            refs = MonthName(docdt.Month, True) & Right(pfc, 1) & "-" & Right(docdt.Year, 2) & Left(docur, 1).ToUpper
                                            'oRSet.DoQuery("Select SUM(U_amount)FCAmount,SUM(U_docrate * U_amount) LCAmount,SUM(U_docrate * U_amount) / SUM(U_amount) 'Average Rate',U_doccur from [@GEN_FWD_COVER] Where U_unit = '" & objForm.Items.Item("cpc").Specific.value & "' and U_status = 'Open' Group By U_doccur")
                                            oRSet.DoQuery("Select U_actfc,U_act,U_blend,U_curr,u_bal from [@UBG_FWD_REM] Where Code = '" & refs & "'")
                                            'For cr As Integer = 1 To oRSet.RecordCount
                                            If oRSet.RecordCount > 0 Then
                                                If objForm.Items.Item("41").Specific.Value.ToString = oRSet.Fields.Item(3).Value.ToString Then
                                                    objForm.Items.Item("21").Enabled = True
                                                    objForm.Items.Item("21").Specific.Value = oRSet.Fields.Item(2).Value
                                                    objForm.Items.Item("balamte").Specific.value = oRSet.Fields.Item(4).Value
                                                    objForm.Items.Item("26").Specific.Value = "Extingush Rate"
                                                    objForm.Items.Item("21").Enabled = False
                                                End If
                                            Else
                                                oApplication.StatusBar.SetText("No forward cover for this unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                BubbleEvent = False
                                            End If
                                            '    oRSet.MoveNext()
                                            'Next
                                        End If
                                    End If
                                Else
                                    If objForm.Items.Item("5").Specific.Value <> "" Then
                                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        If objForm.Items.Item("cpc").Specific.value <> "" Then
                                            oRSet.DoQuery("Select Rate,Currency from ORTT Where LEFT(RateDate,11) = LEFT(GETDATE(),11) ")
                                            For cr As Integer = 1 To oRSet.RecordCount
                                                If objForm.Items.Item("41").Specific.Value.ToString = oRSet.Fields.Item(1).Value Then
                                                    objForm.Items.Item("21").Enabled = True
                                                    objForm.Items.Item("21").Specific.Value = oRSet.Fields.Item(0).Value
                                                    CurrSpot = oRSet.Fields.Item(0).Value
                                                    objForm.Items.Item("26").Specific.Value = "Spot Rate"
                                                    Exit For
                                                End If
                                                oRSet.MoveNext()
                                            Next
                                        End If
                                    End If
                                    visualbehaviour("Spot")
                                    objForm.Items.Item("docrte").Specific.value = ""
                                    objForm.Items.Item("balamte").Specific.value = ""
                                    objForm.Items.Item("cntrnoe").Specific.value = ""
                                End If
                            End If
                        ElseIf pVal.BeforeAction = True And pVal.InnerEvent = False Then
                            objForm = oApplication.Forms.Item(FormUID)
                            objCombo = objForm.Items.Item("rte").Specific
                            If objForm.Items.Item("cpc").Specific.Value = "" Then
                                If objCombo.Selected.Value = "Spot" Then
                                    'objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                    oApplication.StatusBar.SetText("Please select unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                End If
                            End If
                        End If
                    ElseIf pVal.ItemUID = "74" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If pVal.ActionSuccess = True Then
                            Dim doccur As String = objForm.Items.Item("74").Specific.value
                            If doccur <> "INR" Then
                                objForm.Items.Item("rts").Visible = True
                                objForm.Items.Item("rte").Visible = True
                            Else
                                objForm.Items.Item("rts").Visible = False
                                objForm.Items.Item("rte").Visible = False
                            End If
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If pVal.ItemUID = "cpc" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        objForm.Items.Item("docrte").Specific.value = ""
                        objForm.Items.Item("balamte").Specific.value = ""
                        objForm.Items.Item("cntrnoe").Specific.value = ""
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    If pVal.BeforeAction = True Then
                        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery("Select u_unit From [@GEN_USR_UNIT] Where u_user = '" + oCompany.UserName.ToString.Trim + "'")
                        If oRecordSet.RecordCount = 0 Then
                            oApplication.StatusBar.SetText("No Unit assigned to User", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        objForm = oApplication.Forms.Item(FormUID)
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                        Dim CFL_Id As String
                        CFL_Id = CFLEvent.ChooseFromListUID
                        oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                        Dim oDT As SAPbouiCOM.DataTable
                        oDT = CFLEvent.SelectedObjects
                        If oCFL.UniqueID = "FRDCFL" Then
                            Me.SetConditionToForward(FormUID)
                        End If
                    End If
                    If pVal.BeforeAction = False Then
                        Dim objForm As SAPbouiCOM.Form
                        objForm = oApplication.Forms.Item(FormUID)
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                        Dim CFL_Id As String
                        CFL_Id = CFLEvent.ChooseFromListUID
                        oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                        Dim oDT As SAPbouiCOM.DataTable
                        oDT = CFLEvent.SelectedObjects
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("ORCT")
                        If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            If oCFL.UniqueID = "14" Or oCFL.UniqueID = "15" Then
                                objForm.Refresh()
                                Dim doccur As String = Trim(oDT.GetValue("Currency", 0))
                                If doccur <> "INR" Then
                                    objForm.Items.Item("rts").Visible = True
                                    objForm.Items.Item("rte").Visible = True
                                Else
                                    objForm.Items.Item("rts").Visible = False
                                    objForm.Items.Item("rte").Visible = False
                                End If
                                'oDBs_Head.SetValue("U_unit", 0, Trim(oDT.GetValue("U_unit", 0)))
                                'objForm.Items.Item("cpc").Specific.value = Trim(oDT.GetValue("U_unit", 0))
                                'objForm.Refresh()
                                objForm.Update()

                            End If
                            If oCFL.UniqueID = "FRDCFL" Then
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                'oDBs_Head.SetValue("U_contrno", 0, oDT.GetValue("U_contrno", 0))
                                If objForm.Items.Item("cpc").Specific.value = "" Then
                                    oRSet.DoQuery("Select U_docrate,U_amount,U_doccur from [@GEN_FWD_COVER] Where U_contrno = '" & oDT.GetValue("U_contrno", 0) & "'") ' and ('" & objForm.Items.Item("10").Specific.value & "' >= U_fdate and '" & objForm.Items.Item("10").Specific.value & "' <= U_tdate)
                                Else
                                    oRSet.DoQuery("Select U_docrate,U_amount,U_doccur from [@GEN_FWD_COVER] Where U_unit = '" & objForm.Items.Item("cpc").Specific.value & "' and U_contrno = '" & oDT.GetValue("U_contrno", 0) & "' ") 'and ('" & objForm.Items.Item("10").Specific.value & "' >= U_fdate and '" & objForm.Items.Item("10").Specific.value & "' <= U_tdate)
                                End If
                                Try
                                    Dim cont As String = oDT.GetValue("U_contrno", 0).trim
                                    'objForm.Items.Item("26").Specific.Value = "Extinguih Rate"
                                    'oDBs_Head.SetValue("U_contrno", 0, cont)
                                    objForm.Items.Item("cntrnoe").Specific.value = cont
                                Catch ex As Exception
                                End Try
                                objForm.Items.Item("docrte").Specific.value = oRSet.Fields.Item("U_docrate").Value
                                'If objForm.Items.Item("74").Specific.value <> oRSet.Fields.Item("U_doccur").Value Then
                                Dim objcombo1 As SAPbouiCOM.ComboBox = objForm.Items.Item("74").Specific
                                For u As Integer = 0 To objcombo1.ValidValues.Count - 1
                                    If objcombo1.ValidValues.Item(u).Value = oRSet.Fields.Item("U_doccur").Value Then
                                        objcombo1.Select(u, SAPbouiCOM.BoSearchKey.psk_Index)
                                    End If
                                Next
                                objForm.Items.Item("72").Enabled = True
                                objForm.Items.Item("72").Specific.value = oRSet.Fields.Item("U_docrate").Value
                                objForm.Items.Item("26").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                objForm.Items.Item("72").Enabled = False
                                'End If
                                objForm.Items.Item("balamte").Specific.value = oRSet.Fields.Item("U_amount").Value * oRSet.Fields.Item("U_docrate").Value
                            End If
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Dim USER_NAME As String = oCompany.UserName
                    If pVal.ItemUID = "2" Then
                        objForm.Close()
                    End If
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("ORCT")
                        PAYNUM = objForm.Items.Item("3").Specific.value
                        DocDate = objForm.Items.Item("10").Specific.value
                        RefDate = objForm.Items.Item("90").Specific.value
                        JVRem = objForm.Items.Item("59").Specific.value
                        Dim DocType As String = oDBs_Head.GetValue("DocType", 0).Trim().ToString()
                        If DocType = "A" Then
                            If USER_NAME <> "manager" Then
                                Dim GLAccount As String
                                For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                                    GLAccount = objMatrix.Columns.Item("1").Cells.Item(Row).Specific.value
                                    Dim GLacc As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    GLacc.DoQuery("Select (substring('" + GLAccount + "', 1, len('" + GLAccount + "')-3)+RIGHT('" + GLAccount + "',2))")
                                    Dim ManualJE As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    ManualJE.DoQuery("Select count(Code) From [@GEN_M_JE] where code='" + GLacc.Fields.Item(0).Value + "'")
                                    If ManualJE.Fields.Item(0).Value > 0 Then
                                        oApplication.StatusBar.SetText("You Are Not Permitted To Perform This Action", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                Next
                            End If
                        End If
                        Dim oPC As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oSeries As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oPC.DoQuery("Select u_unit from [@GEN_USR_unit] where u_user='" + oCompany.UserName + "'")
                        oSeries.DoQuery("Select (case when SeriesName like '%U1%' then 'UNIT1' when SeriesName like '%U2%' then 'UNIT2' when SeriesName like '%U3%' then 'UNIT3' when SeriesName like '%LU%' then 'LG-UNIT1' end) from NNM1 where series='" + Trim(objForm.Items.Item("87").Specific.value) + "'")
                        For i As Integer = 1 To oPC.RecordCount
                            If i <> oPC.RecordCount Then
                                If oSeries.Fields.Item(0).Value <> oPC.Fields.Item(0).Value Then
                                    oPC.MoveNext()
                                End If
                            Else
                                If oSeries.Fields.Item(0).Value <> oPC.Fields.Item(0).Value Then
                                    oApplication.StatusBar.SetText("Please Select Correct Series", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                        Next
                    End If
                    If pVal.ItemUID = "2" And pVal.BeforeAction = False Then
                        ModalForm = False
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        ModalForm = False
                    End If
                    'If pVal.ItemUID = "2" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    '    Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                    '    objForm = oApplication.Forms.Item(FormUID)
                    '    docno = objForm.Items.Item("3").Specific.value
                    '    Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '    Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '    RS.DoQuery("Delete From [@ORCT_JV] Where u_docno = '" + Trim(objForm.Items.Item("3").Specific.value) + "' And u_macid = '" + MAC_ID + "'")
                    'End If
                    If pVal.ItemUID = "btn" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        docno = objForm.Items.Item("3").Specific.value
                        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                        Me.Open_OtherCharges_Form(pVal.FormUID, Trim(objForm.Items.Item("3").Specific.value), MAC_ID)
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        PAYNUM = objForm.Items.Item("3").Specific.value
                        DocDate = objForm.Items.Item("10").Specific.value
                        RefDate = objForm.Items.Item("90").Specific.value
                        JVRem = objForm.Items.Item("59").Specific.value
                        If objForm.Items.Item("rte").Specific.Value.ToString.Trim = "FWD Cover" Then
                            If Me.Validation_ForwardCover(pVal.FormUID) = False Then
                                BubbleEvent = False
                            End If
                        End If
                    End If
                    'Dim objForm As SAPbouiCOM.Form

                    objForm = oApplication.Forms.Item(FormUID)
                    optnBtn = objForm.Items.Item("58").Specific
                    optnBtn1 = objForm.Items.Item("56").Specific
                    optnBtn2 = objForm.Items.Item("57").Specific
                    If (pVal.ItemUID = "58" Or pVal.ItemUID = "56" Or pVal.ItemUID = "57") And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If optnBtn.Selected = True Then
                            objForm.Items.Item("rts").Visible = False
                            objForm.Items.Item("rte").Visible = False
                            visualbehaviour("Spot")
                            objForm.Items.Item("docrte").Specific.value = ""
                            objForm.Items.Item("balamte").Specific.value = ""
                            objForm.Items.Item("cntrnoe").Specific.value = ""
                        ElseIf optnBtn1.Selected = True Then
                            objForm.Items.Item("rts").Visible = False
                            objForm.Items.Item("rte").Visible = False
                            visualbehaviour("Spot")
                            objForm.Items.Item("docrte").Specific.value = ""
                            objForm.Items.Item("balamte").Specific.value = ""
                            objForm.Items.Item("cntrnoe").Specific.value = ""
                        ElseIf optnBtn2.Selected = True Then
                            objForm.Items.Item("rts").Visible = False
                            objForm.Items.Item("rte").Visible = False
                            visualbehaviour("Spot")
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                    If pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        RS.DoQuery("Delete From [@ORCT_JV] Where u_docno = '" + Trim(objForm.Items.Item("3").Specific.value) + "' And u_macid = '" + MAC_ID + "'")
                    End If
            End Select
        Catch ex As Exception
            '    oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then
                        objForm.Items.Item("btn").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.BeforeAction = False Then
                        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim DocEntry As String
                        Dim UserSign As String
                        Dim DocNum As String
                        oRSet.DoQuery("Select UserID From OUSR Where User_Code = '" + oCompany.UserName.ToString.Trim + "'")
                        UserSign = oRSet.Fields.Item("UserID").Value
                        oRSet.DoQuery("Select Max(DocEntry) As 'DocEntry' From ORCT Where UserSign = '" + UserSign + "'")
                        DocEntry = oRSet.Fields.Item("DocEntry").Value
                        oRSet.DoQuery("Select DocNum From ORCT Where DocEntry = '" + DocEntry + "'")
                        DocNum = oRSet.Fields.Item("DocNum").Value
                        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                        oRSet.DoQuery("Update [@ORCT_JV] Set u_docno = '" + DocNum + "' Where u_docno = '" + PAYNUM + "' And u_macid = '" + MAC_ID + "'")
                        oRSet.DoQuery("Select u_acctcode,u_debit as 'Debit',u_credit As 'Credit' From [@ORCT_JV] Where u_docno = '" + DocNum + "' And u_macid = '" + MAC_ID + "' ANd u_acctcode <> ''")
                        Dim oJE As SAPbobsCOM.JournalEntries = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                        oJE.TaxDate = DateTime.ParseExact(RefDate, "yyyyMMdd", Nothing)
                        oJE.ReferenceDate = DateTime.ParseExact(DocDate, "yyyyMMdd", Nothing)
                        oJE.Memo = JVRem
                        oJE.TransactionCode = "101"
                        oJE.Reference = DocNum
                        For k As Integer = 1 To oRSet.RecordCount
                            'Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("Select AcctCode From OACT WHere FormatCode = '" + Trim(oRSet.Fields.Item("u_acctcode").Value) + "'")
                            If oRSet.Fields.Item("Debit").Value > 0 Then
                                oJE.Lines.AccountCode = oRecordSet.Fields.Item("AcctCode").Value
                                oJE.Lines.Debit = oRSet.Fields.Item("Debit").Value
                                oJE.Lines.Credit = 0
                            End If
                            If oRSet.Fields.Item("Credit").Value > 0 Then
                                oJE.Lines.AccountCode = oRecordSet.Fields.Item("AcctCode").Value
                                oJE.Lines.Credit = oRSet.Fields.Item("Credit").Value
                                oJE.Lines.Debit = 0
                            End If
                            oJE.Lines.Add()
                            oJE.Lines.SetCurrentLine(oJE.Lines.Count - 1)
                            oRSet.MoveNext()
                        Next
                        Dim errflg As Integer = oJE.Add()
                        If errflg <> 0 Then
                            oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription.ToString)
                        Else
                            Dim Key As String = oCompany.GetNewObjectKey
                            oRSet.DoQuery("Update ORCT Set u_bcjv = '" + Key + "' Where DocNum = '" + DocNum + "'")
                            oRSet.DoQuery("Delete From [@ORCT_JV] Where u_docno = '" + DocNum + "' And u_macid = '" + MAC_ID + "'")
                        End If
                        oRSet.DoQuery("Select U_Frgn,DocTotal,U_balamt,U_contrno,DocType,u_unit,DocRate,DocDate,DocCurr From ORCT Where DocNum = '" + DocNum + "'")
                        Dim actual, post As Double
                        Dim actualfc, curr As String
                        Dim balance, bal As String
                        Dim contract, unit As String
                        Dim postdate As Date
                        Dim refcode As String
                        curr = oRSet.Fields.Item(8).Value
                        actual = oRSet.Fields.Item(2).Value
                        post = oRSet.Fields.Item(1).Value
                        balance = actual - post
                        actualfc = balance / oRSet.Fields.Item(6).Value
                        contract = oRSet.Fields.Item(3).Value.ToString.Trim
                        unit = oRSet.Fields.Item(5).Value.ToString.Trim
                        postdate = oRSet.Fields.Item(7).Value
                        refcode = MonthName(postdate.Month, True) & Right(unit, 1) & "-" & Right(postdate.Year, 2) & Left(curr, 1)
                        If oRSet.Fields.Item(0).Value = "FWD Cover" Then
                            If oRSet.Fields.Item(4).Value = "A" Then
                                oRecordSet.DoQuery("Update [@GEN_FWD_COVER] Set U_amount = '" + actualfc + "' Where U_contrno = '" + contract + "'")
                                If balance = 0 Then
                                    oRecordSet.DoQuery("Update [@GEN_FWD_COVER] Set U_amount = '" + actualfc + "',U_status = 'Encash' Where U_contrno = '" + contract + "'")
                                End If
                            ElseIf oRSet.Fields.Item(4).Value = "C" Then
                                bal = balance / oRSet.Fields.Item(6).Value
                                oRecordSet.DoQuery("Update [@UBG_FWD_REM] Set U_balfc = '" + bal + "',U_bal = '" + balance + "' Where Code = '" + refcode + "'")
                                If balance = 0 Then
                                    oRecordSet.DoQuery("Update [@UBG_FWD_REM] Set U_balfc = '" + bal + "',U_bal = '" + balance + "',U_status = 'Encash' Where Code = '" + refcode + "'")
                                End If
                            End If
                        End If
                        visualbehaviour("Spot")
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub Open_OtherCharges_Form(ByVal FormUID As String, ByVal docno As String, ByVal macid As String)
        Try
            PARENT_FORM = FormUID
            Dim RS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oMatrix As SAPbouiCOM.Matrix
            Dim CHILD_FORM As String = "ORCT_JV@" & FormUID
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
                oUtilities.SAPXML("OTHERCHRGS.xml", CHILD_FORM)
                objSubForm = oApplication.Forms.Item(CHILD_FORM)
                objSubForm.Select()
            End If
            oMatrix = objSubForm.Items.Item("mtx").Specific
            RS1.DoQuery("Select Distinct code,u_acctcode,u_acctname,u_debit,u_credit From [@ORCT_JV] Where u_docno = '" + docno + "' and u_macid = '" + macid + "' And Isnull(u_acctcode,'') <> ''")
            If RS1.RecordCount > 0 Then
                oMatrix.AddRow(1)
                Me.SetNewLine(objSubForm.UniqueID, oMatrix.VisualRowCount, oMatrix)
                For k As Integer = 1 To RS1.RecordCount
                    ' objSubForm = oApplication.Forms.Item(FormUID)
                    oDBs_Detail = objSubForm.DataSources.DBDataSources.Item("@ORCT_JV")
                    oDBs_Detail.Offset = k - 1
                    oDBs_Detail.SetValue("u_acctcode", oDBs_Detail.Offset, RS1.Fields.Item("u_acctcode").Value)
                    oDBs_Detail.SetValue("u_acctname", oDBs_Detail.Offset, RS1.Fields.Item("u_acctname").Value)
                    oDBs_Detail.SetValue("u_debit", oDBs_Detail.Offset, RS1.Fields.Item("u_debit").Value)
                    oDBs_Detail.SetValue("u_credit", oDBs_Detail.Offset, RS1.Fields.Item("u_credit").Value)
                    oMatrix.SetLineData(oMatrix.VisualRowCount)
                    RS1.MoveNext()
                    oMatrix.AddRow(1, oMatrix.VisualRowCount)
                    Me.SetNewLine(objSubForm.UniqueID, oMatrix.VisualRowCount, oMatrix)
                Next
                Dim totcr, totdr As Double
                For i As Integer = 1 To oMatrix.VisualRowCount
                    totcr = totcr + oMatrix.Columns.Item("credit").Cells.Item(i).Specific.value
                    totdr = totdr + oMatrix.Columns.Item("debit").Cells.Item(i).Specific.value
                Next
                objSubForm.Items.Item("drtot").Specific.value = totcr
                objSubForm.Items.Item("crtot").Specific.value = totdr
            End If
            If RS1.RecordCount = 0 Then
                oMatrix.AddRow(1)
                Me.SetNewLine(objSubForm.UniqueID, oMatrix.VisualRowCount, oMatrix)
            End If
            Me.SetConditionToSO(objSubForm.UniqueID)
            ModalForm = True
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
            oCFLCreationParams.ObjectType = "GEN_FWD_COVER"
            oCFLCreationParams.UniqueID = "FRDCFL"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
   
    Sub SetConditionToForward(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("FRDCFL")
            docdt = DateTime.ParseExact(objForm.Items.Item("10").Specific.Value, "yyyyMMdd", Nothing)
            mnth = docdt.Month
            yr = docdt.Year
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            'Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("12")
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If Trim(objForm.Items.Item("cpc").Specific.value) = "" Then
                oRecordSet.DoQuery("Select U_contrno From [@GEN_FWD_COVER] Where U_status = 'Open' and  ('" & objForm.Items.Item("10").Specific.value & "' >= U_fdate and '" & objForm.Items.Item("10").Specific.value & "' <= U_tdate)") '
            Else
                oRecordSet.DoQuery("Select U_contrno From [@GEN_FWD_COVER] Where  U_unit = '" + Trim(objForm.Items.Item("cpc").Specific.value) + "' and U_status <> 'Cancelled' and MONTH(U_fdate) = '" + mnth + "' and YEAR(U_fdate) = '" + yr + "'") 'and ('" & objForm.Items.Item("10").Specific.value & "' >= U_fdate and '" & objForm.Items.Item("10").Specific.value & "' <= U_tdate)
            End If

            For i As Integer = 0 To oRecordSet.RecordCount - 1
                If i > 0 Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                oCon = oCons.Add()
                oCon.Alias = "U_contrno"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = oRecordSet.Fields.Item("U_contrno").Value
                oRecordSet.MoveNext()
            Next
            If oRecordSet.RecordCount = 0 Then
                oCon = oCons.Add()
                oCon.Alias = "U_contrno"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "-1"
            End If
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub SetConditionToSO(ByVal FormUID As String)
        Try
            objSubForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objSubForm.ChooseFromLists.Item("ACTCFL")
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objSubForm = oApplication.Forms.Item(FormUID)
            oDBs_Detail = objSubForm.DataSources.DBDataSources.Item("@ORCT_JV")
            objMatrix = oMatrix
            objMatrix.FlushToDataSource()
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("u_acctcode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_acctname", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_debit", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_credit", oDBs_Detail.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent_OtherCharges(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objSubForm = oApplication.Forms.Item(pVal.FormUID)
                    PARENT_FORM = (pVal.FormUID.Substring(objSubForm.UniqueID.IndexOf("@") + 1))
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If pVal.BeforeAction = False Then
                        oDBs_Detail = objSubForm.DataSources.DBDataSources.Item("@ORCT_JV")
                        Dim oMatrix As SAPbouiCOM.Matrix = objSubForm.Items.Item("mtx").Specific
                        Dim totdebit As Double
                        Dim totcredit As Double
                        Try
                            If pVal.ItemUID = "mtx" And (pVal.ColUID = "debit" Or pVal.ColUID = "credit") Then
                                For i As Integer = 1 To oMatrix.VisualRowCount
                                    If oMatrix.Columns.Item("debit").Cells.Item(i).Specific.value <> "" Then
                                        totdebit = totdebit + oMatrix.Columns.Item("debit").Cells.Item(i).Specific.value
                                    End If
                                    If oMatrix.Columns.Item("credit").Cells.Item(i).Specific.value <> "" Then
                                        totcredit = totcredit + oMatrix.Columns.Item("credit").Cells.Item(i).Specific.value
                                    End If
                                Next
                                objSubForm.Items.Item("drtot").Specific.value = totdebit
                                objSubForm.Items.Item("crtot").Specific.value = totcredit
                            End If
                        Catch ex As Exception
                            oApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "btn" And pVal.BeforeAction = True Then
                        If Me.Validation_OtherCharges(pVal.FormUID) = False Then
                            BubbleEvent = False
                        End If
                    End If
                    If pVal.ItemUID = "btn" And pVal.BeforeAction = False Then
                        Dim oMatrix As SAPbouiCOM.Matrix = objSubForm.Items.Item("mtx").Specific
                        Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' Dim Code As String = RS.Fields.Item("Code").Value
                        oRSet.DoQuery("Delete From [@ORCT_JV] Where u_docno = '" + docno + "' And u_macid = '" + MAC_ID + "'")
                        For i As Integer = 1 To oMatrix.VisualRowCount
                            If oMatrix.Columns.Item("acctcode").Cells.Item(i).Specific.Value <> "" Then
                                RS.DoQuery("SELECT isnull(MAX(CAST( Code AS int)),0) +1 AS Code  FROM [@ORCT_JV]")
                                Dim dbt, cdt As String
                                dbt = oMatrix.Columns.Item("debit").Cells.Item(i).Specific.value
                                cdt = oMatrix.Columns.Item("credit").Cells.Item(i).Specific.value
                                oRSet.DoQuery("Insert Into [@ORCT_JV] (Code,Name,u_docno,u_acctcode,u_acctname,u_debit,u_credit,u_macid) Values('" + Trim(RS.Fields.Item("Code").Value) + "','" + Trim(RS.Fields.Item("Code").Value) + "','" + docno + "','" + objMatrix.Columns.Item("acctcode").Cells.Item(i).Specific.value.ToString.Trim + "','" + objMatrix.Columns.Item("acctname").Cells.Item(i).Specific.value.ToString.Trim + "','" + dbt + "','" + cdt + "','" + MAC_ID + "')")
                            End If
                        Next
                        objSubForm.Close()
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objSubForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objSubForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If oCFL.UniqueID = "ACTCFL" Then
                            objMatrix = objSubForm.Items.Item("mtx").Specific
                            oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim Total As Double = 0
                            oDBs_Detail = objSubForm.DataSources.DBDataSources.Item("@ORCT_JV")
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
                                oDBs_Detail.SetValue("u_acctcode", oDBs_Detail.Offset, oDT.GetValue("FormatCode", 0))
                                oDBs_Detail.SetValue("u_acctname", oDBs_Detail.Offset, oDT.GetValue("AcctName", 0))
                                objMatrix.SetLineData(pVal.Row + i)
                                objSubForm.EnableMenu("1293", True)
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

    Function Validation_OtherCharges(ByVal FormUID As String) As Boolean
        Try
            Dim errflag As Boolean = False
            Dim CreditTotal, DebitTotal As Double
            objSubForm = oApplication.Forms.Item(FormUID)
            Dim oMatrix As SAPbouiCOM.Matrix
            oMatrix = objSubForm.Items.Item("mtx").Specific
            For i As Integer = 1 To oMatrix.VisualRowCount
                If Trim(oMatrix.Columns.Item("acctcode").Cells.Item(i).Specific.value) <> "" Then
                    If CDbl(oMatrix.Columns.Item("debit").Cells.Item(i).Specific.value) > 0 And CDbl(oMatrix.Columns.Item("credit").Cells.Item(i).Specific.value) > 0 Then
                        oApplication.StatusBar.SetText("Cannot enter debit and credit in the same row", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                        Exit Function
                    End If
                    CreditTotal = CreditTotal + oMatrix.Columns.Item("credit").Cells.Item(i).Specific.value
                    DebitTotal = DebitTotal + oMatrix.Columns.Item("debit").Cells.Item(i).Specific.value
                End If
            Next
            If DebitTotal <> CreditTotal Then
                oApplication.StatusBar.SetText("Credit and Debit does not match", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function
    Function Validation_ForwardCover(ByVal FormUID As String) As Boolean
        Try
            Dim PaidTotal, ForwardBal As Double
            PaidTotal = objForm.Items.Item("77").Specific.Value.ToString.Substring(3)
            ForwardBal = objForm.Items.Item("balamte").Specific.Value
            If PaidTotal > ForwardBal Then
                oApplication.StatusBar.SetText("No balance amount for forward cover", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function
    Sub visualbehaviour(ByVal selection As String)
        If selection.Trim = "Spot" Then
            objForm.Items.Item("cntrnos").Visible = False
            objForm.Items.Item("cntrnoe").Visible = False
            objForm.Items.Item("docrts").Visible = False
            objForm.Items.Item("docrte").Visible = False
            objForm.Items.Item("balamts").Visible = False
            objForm.Items.Item("balamte").Visible = False
            objForm.Items.Item("72").Enabled = True
        ElseIf selection.Trim = "FWD Cover" Then
            objForm.Items.Item("cntrnos").Visible = True
            objForm.Items.Item("cntrnoe").Visible = True
            objForm.Items.Item("docrts").Visible = True
            objForm.Items.Item("docrte").Visible = True
            objForm.Items.Item("balamts").Visible = True
            objForm.Items.Item("balamte").Visible = True

        End If
    End Sub
End Class
