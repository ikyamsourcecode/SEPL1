Public Class ClsPurchaseOrder

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm, objSubForm, objSForm, objform1 As SAPbouiCOM.Form
    Dim objItem, objOldItem, TempItem, oItem, SaleItem, TypeItem As SAPbouiCOM.Item
    Dim objMatrix, objSubMatrix, objBatchMatrix As SAPbouiCOM.Matrix
    Dim SZDBHead As SAPbouiCOM.DBDataSource
    Dim SZDBDetail As SAPbouiCOM.DBDataSource
    Dim SMDBHead As SAPbouiCOM.DBDataSource
    Dim SMDBDetail As SAPbouiCOM.DBDataSource
    Dim oCombox As SAPbouiCOM.ComboBox
    Dim oDBs_Head, oDBs_Type_Head As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim RS1, RS2 As SAPbobsCOM.Recordset
    Dim ModalForm As Boolean = False
    Dim ChildModalForm As Boolean = False
    Dim PARENT_FORM As String
    Dim RowNo As Integer
    Dim orderno, hwid, mat As String
    Dim sorderno, shwid, sitemcode, saleno As String
    Dim Mode As Integer
    Dim TotQty As Double
    Dim GSONO, GMACID, GITEMCODE As String
    Dim RowID As Integer
    Dim DeleteItemCode As String
    Dim objBtnCmb As SAPbouiCOM.ButtonCombo
    Dim DOCNUM As String = ""
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("38").Specific
            objMatrix.Columns.Item("3").Editable = False


            objOldItem = objForm.Items.Item("86")
            objItem = objForm.Items.Add("sseason", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width
            objItem.Specific.Caption = "Season"
            objItem.LinkTo = "86"
            objOldItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("season", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width
            objItem.Specific.DataBind.setbound(True, "OPOR", "u_season")
            objItem.Specific.TabOrder = objOldItem.Specific.TabOrder + 1
            ''added by vivek
            Dim oSeason As SAPbouiCOM.ComboBox
            oSeason = objItem.Specific

            oSeason.ValidValues.Add("Summer", "Summer")
            oSeason.ValidValues.Add("Fall Summer", "Fall Summer")
            oSeason.ValidValues.Add("Autum", "Autum")
            oSeason.ValidValues.Add("Winter", "Winter")
            oSeason.ValidValues.Add("Fall Winter", "Fall Winter")
            oSeason.ValidValues.Add("Spring", "Spring")
            oSeason.ValidValues.Add("General", "General")



            objItem.LinkTo = "46"
            objOldItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("opono", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width
            objItem.Specific.DataBind.setbound(True, "OPOR", "u_opono")
            objItem.LinkTo = "46"
            objItem.Visible = False
            objItem.Specific.TabOrder = objOldItem.Specific.TabOrder + 1
            objItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm = oApplication.Forms.Item(FormUID)
            TempItem = objForm.Items.Item("86")
            objItem = objForm.Items.Add("spc", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = TempItem.Top + TempItem.Height + 20
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.Caption = "Unit"
            objItem.LinkTo = "86"
            TempItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("cpc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = TempItem.Top + TempItem.Height + 20
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.databind.setbound(True, "OPOR", "u_unit")
            objItem.Specific.TabOrder = TempItem.Specific.TabOrder + 1
            objItem.LinkTo = "46"
            objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            SaleItem = objForm.Items.Item("86")
            objItem = objForm.Items.Add("sale", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = SaleItem.Top + SaleItem.Height + 35
            objItem.Left = SaleItem.Left
            objItem.Width = SaleItem.Width
            objItem.Height = SaleItem.Height
            objItem.Specific.Caption = "Sale Order No."
            objItem.LinkTo = "86"
            SaleItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("sales", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = SaleItem.Top + SaleItem.Height + 35
            objItem.Left = SaleItem.Left
            objItem.Width = SaleItem.Width
            objItem.Height = SaleItem.Height
            objItem.Specific.databind.setbound(True, "OPOR", "u_sono")
            objItem.Specific.TabOrder = SaleItem.Specific.TabOrder + 1
            objItem.LinkTo = "46"
            objForm.Items.Item("sales").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("sales").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'Me.SetChooseFromList1(FormUID)

            TypeItem = objForm.Items.Item("15")
            objItem = objForm.Items.Add("type", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = TypeItem.Top + TypeItem.Height + 20
            objItem.Left = TypeItem.Left
            objItem.Width = TypeItem.Width
            objItem.Height = TypeItem.Height
            objItem.Specific.Caption = "Purchase Type"
            objItem.LinkTo = "15"
            TypeItem = objForm.Items.Item("14")
            objItem = objForm.Items.Add("types", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            objItem.Top = TypeItem.Top + TypeItem.Height + 20
            objItem.Left = TypeItem.Left
            objItem.Width = TypeItem.Width
            objItem.Height = TypeItem.Height
            objItem.Specific.databind.setbound(True, "OPOR", "U_Types")
            objItem.Specific.TabOrder = TypeItem.Specific.TabOrder + 1
            objItem.LinkTo = "14"
            objForm.Items.Item("types").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("types").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            Dim oPurType As SAPbouiCOM.ComboBox
            oPurType = objItem.Specific

            oPurType.ValidValues.Add("Regular", "Regular")
            oPurType.ValidValues.Add("Consumable", "Consumable")

            TempItem = objForm.Items.Item("2")
            objItem = objForm.Items.Add("btn", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Top = TempItem.Top
            objItem.Left = TempItem.Left + TempItem.Width + 10
            objItem.Width = TempItem.Width + 20
            objItem.Height = TempItem.Height
            objItem.Specific.caption = "Copy From BOM"
            Me.SetChooseFromList(FormUID)
            objItem.Specific.ChooseFromListUID = "BOMCFL"
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetChooseFromList(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition

            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = True
            oCFLCreationParams.ObjectType = "GEN_CUST_BOM"
            oCFLCreationParams.UniqueID = "BOMCFL"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetChooseFromList1(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition

            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = True
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CardCode"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetCondition(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("BOMCFL")
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_closed"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            oCon.CondVal = "Y"
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
            oCon = oCons.Add()
            oCon.Alias = "U_closed"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_IS_NULL
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Function Validation(ByVal FormUID As String) As Boolean
        Try
            objForm = oApplication.Forms.Item(FormUID)
            objform1 = oApplication.Forms.GetFormByTypeAndCount("-142", 1)
            If Trim(objForm.Items.Item("season").Specific.value) = "" Then
                oApplication.StatusBar.SetText("Please enter season in PO", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If Trim(objForm.Items.Item("types").Specific.value) = "Regular" Then
                If Trim(objForm.Items.Item("sales").Specific.value) = "" Then
                    oApplication.StatusBar.SetText("Please enter sale order no in PO", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
            Return False
        End Try
    End Function

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
                        Me.CreateForm(FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If pVal.ItemUID = "cpc" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = True Then
                        If Trim(objForm.Items.Item("cpc").Specific.value) <> "" Then
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("Select u_unit From [@GEN_USR_UNIT] Where u_user = '" + Trim(oCompany.UserName.ToString) + "' And u_unit = '" + Trim(objForm.Items.Item("cpc").Specific.value) + "'")
                            If oRecordSet.RecordCount = 0 Then
                                oApplication.StatusBar.SetText("Please select the correct Unit for the user", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If


                    '------------------------------------------ Vijeesh ----------------------------------------------'

                    'If pVal.ItemUID = "38" And (pVal.ColUID = "1") And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.InnerEvent = False Then
                    '    Try
                    '        objForm.Freeze(True)
                    '        objMatrix = objForm.Items.Item("38").Specific
                    '        If Trim(objForm.Items.Item("cpc").Specific.value) <> "" Then
                    '            If Trim(objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.value) <> "" Then
                    '                For i As Integer = 1 To objMatrix.RowCount - 1
                    '                    If Trim(objForm.Items.Item("cpc").Specific.value) = "UNIT1" Then
                    '                        objMatrix.Columns.Item("24").Cells.Item(i).Specific.value = "INSP-1"
                    '                    ElseIf Trim(objForm.Items.Item("cpc").Specific.value) = "UNIT2" Then
                    '                        objMatrix.Columns.Item("24").Cells.Item(i).Specific.value = "INSP-2"
                    '                    End If
                    '                Next
                    '            End If
                    '        End If
                    '        objForm.Freeze(False)
                    '    Catch ex As Exception
                    '    End Try
                    'End If

                    If pVal.ItemUID = "38" And (pVal.ColUID = "24") And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.InnerEvent = False Then
                        Try
                            objForm.Freeze(True)
                            objMatrix = objForm.Items.Item("38").Specific
                            If Trim(objForm.Items.Item("cpc").Specific.value) <> "" Then
                                If Trim(objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.value) <> "" Then
                                    If Trim(objForm.Items.Item("cpc").Specific.value) = "UNIT1" Then
                                        objMatrix.Columns.Item("24").Cells.Item(pVal.Row).Specific.value = "INSP-1"
                                    ElseIf Trim(objForm.Items.Item("cpc").Specific.value) = "UNIT2" Then
                                        objMatrix.Columns.Item("24").Cells.Item(pVal.Row).Specific.value = "INSP-2"
                                    End If
                                End If
                            End If
                            objForm.Freeze(False)
                        Catch ex As Exception
                        End Try
                    End If

                    '--------------------------------------------------------------------------------------------------------'

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    If pVal.BeforeAction = True Then
                        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery("Select u_unit From [@GEN_USR_UNIT] Where u_user = '" + oCompany.UserName.ToString.Trim + "'")
                        If oRecordSet.RecordCount = 0 Then
                            oApplication.StatusBar.SetText("No PC assigned to User", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        Dim objForm As SAPbouiCOM.Form
                        objForm = oApplication.Forms.Item(FormUID)
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                        Dim CFL_Id As String
                        CFL_Id = CFLEvent.ChooseFromListUID
                        oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                        If oCFL.UniqueID = "BOMCFL" Then
                            Me.SetCondition(FormUID)
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
                        If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            If oCFL.UniqueID = "2" Then
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet.DoQuery("Select u_unit From OCRD Where CardCode = '" + Trim(oDT.GetValue("CardCode", 0)) + "'")
                                objForm.Items.Item("cpc").Specific.value = oRSet.Fields.Item("u_unit").Value
                                objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            End If
                            If oCFL.UniqueID = "3" Then
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet.DoQuery("Select u_unit From OCRD Where CardCode = '" + Trim(oDT.GetValue("CardCode", 0)) + "'")
                                objForm.Items.Item("cpc").Specific.value = oRSet.Fields.Item("u_unit").Value
                                objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            End If
                            If oCFL.UniqueID = "BOMCFL" Then
                                For i As Integer = 0 To oDT.Rows.Count - 1
                                    DOCNUM = DOCNUM & "'" & Trim(oDT.GetValue("DocNum", i)) & "',"
                                Next
                                DOCNUM = DOCNUM.Substring(0, Len(DOCNUM) - 1)
                            ElseIf oCFL.UniqueID = "2" Or oCFL.UniqueID = "3" Then
                                objForm.Items.Item("btn").Enabled = True
                            End If
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "types" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = False Then
                        objMatrix = objForm.Items.Item("38").Specific
                        If Trim(objForm.Items.Item("types").Specific.value) = "Regular" Then
                            objForm.Items.Item("sales").Enabled = True
                            If objMatrix.VisualRowCount > 0 Then
                                objMatrix.Clear()
                                objMatrix.AddRow()
                                mat = "Y"
                            Else
                                mat = "N"
                            End If
                        Else
                            objForm.Items.Item("sales").Enabled = False
                            If Trim(objForm.Items.Item("sales").Specific.value) <> "" Then
                                objForm.Items.Item("sales").Specific.Value = ""
                                objMatrix.Clear()
                                objMatrix.AddRow()
                                mat = "Y"
                            Else
                                mat = "N"
                            End If
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "sales" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = True Then
                        saleno = objForm.Items.Item("sales").Specific.value
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If pVal.ItemUID = "38" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And (pVal.ColUID = "U_qty" Or pVal.ColUID = "U_tol") And pVal.BeforeAction = False Then
                        Try
                            objMatrix = objForm.Items.Item("38").Specific
                            If Trim(objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.value) <> "" Then
                                objMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific.value = CDbl(objMatrix.Columns.Item("U_qty").Cells.Item(pVal.Row).Specific.value + (objMatrix.Columns.Item("U_tol").Cells.Item(pVal.Row).Specific.value * objMatrix.Columns.Item("U_qty").Cells.Item(pVal.Row).Specific.value) / 100)
                            End If
                        Catch ex As Exception
                            oApplication.StatusBar.SetText(ex.Message)
                        End Try
                    ElseIf pVal.ItemUID = "sales" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = False Then
                        Try
                            objMatrix = objForm.Items.Item("38").Specific
                            If Trim(objForm.Items.Item("types").Specific.value) = "Regular" Then
                                If Trim(objForm.Items.Item("sales").Specific.value) <> "" Then
                                    If ((saleno <> objForm.Items.Item("sales").Specific.value) And saleno <> Nothing And objMatrix.VisualRowCount <> 0) Or objMatrix.VisualRowCount <> 0 Then
                                        objMatrix.Clear()
                                        objMatrix.AddRow()
                                        mat = "Y"
                                    Else
                                        mat = "N"
                                    End If
                                    Me.FilterItem(FormUID)
                                Else
                                    BubbleEvent = False
                                    oApplication.SetStatusBarMessage("Sales order no. is not given")
                                End If
                            Else
                                Me.FilterItem1(FormUID)
                            End If
                        Catch ex As Exception
                            oApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If


                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = True Then
                        If Trim(objForm.Items.Item("cpc").Specific.value) = "" Then
                            oApplication.StatusBar.SetText("Please select Unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                    If pVal.ItemUID = "38" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And (pVal.ColUID = "11") And pVal.ActionSuccess = True Then
                        Try
                            objMatrix = objForm.Items.Item("38").Specific
                            '        ' If Trim(objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.value) <> "" Then

                            objMatrix.Columns.Item("U_qty").Cells.Item(pVal.Row).Specific.value = CDbl(objMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific.value) ' + (objMatrix.Columns.Item("U_tol").Cells.Item(pVal.Row).Specific.value * objMatrix.Columns.Item("U_qty").Cells.Item(pVal.Row).Specific.value) / 100)
                            '        ' End If
                        Catch ex As Exception
                            oApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If
                    'Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                    '   Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    'If pVal.ItemUID = "38" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And (pVal.ColUID = "11" Or pVal.ColUID = "24") And (pVal.CharPressed <> 9 And pVal.CharPressed <> 13) Then
                    '    BubbleEvent = False
                    'End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                        'oItem = objForm.Items.Item("sales").Specific

                        If Me.Validation(FormUID) = False Then
                            BubbleEvent = False
                        End If

                        '-----------------------'
                        Try
                            objForm.Freeze(True)
                            objMatrix = objForm.Items.Item("38").Specific
                            If Trim(objForm.Items.Item("cpc").Specific.value) <> "" Then
                                If mat = "N" Then
                                    For i As Integer = 1 To objMatrix.RowCount - 1
                                        If Trim(objForm.Items.Item("cpc").Specific.value) = "UNIT1" Then
                                            objMatrix.Columns.Item("24").Cells.Item(i).Specific.value = "INSP-1"
                                        ElseIf Trim(objForm.Items.Item("cpc").Specific.value) = "UNIT2" Then
                                            objMatrix.Columns.Item("24").Cells.Item(i).Specific.value = "INSP-2"
                                        End If
                                    Next
                                End If
                            End If
                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            oApplication.StatusBar.SetText(ex.Message)
                        End Try
                        '-----------------------'
                    End If
                    If pVal.ItemUID = "btn" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                        If Trim(objForm.Items.Item("4").Specific.value) = "" Then
                            oApplication.StatusBar.SetText("Please select Vendor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
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
                    Case "1288", "1289", "1290", "1291"
                        If objForm.TypeEx = "142" Then
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
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    End If
            End Select
        Catch ex As Exception

        End Try
    End Sub

    Sub LoadItems(ByVal FormUID As String, ByVal MreqNo As String)
        Try
            Dim ITForm As SAPbouiCOM.Form
            Dim ITMatrix As SAPbouiCOM.Matrix
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select B.u_itemcode,B.u_totqty,A.DocNum,B.LineId From [@GEN_CUST_BOM] A Inner Join [@GEN_CUST_BOM_D0] B On A.DocEntry = B.DocEntry And B.u_cardcode = '" + Trim(objForm.Items.Item("4").Specific.value) + "' And IsNull(A.u_closed,'N') = 'N' And A.u_sono = '" + DOCNUM + "'")
            ITForm = oApplication.Forms.Item(FormUID)
            Try
                ITForm.Freeze(True)
                ITMatrix = ITForm.Items.Item("38").Specific
                ITMatrix.Clear()
                ITMatrix.AddRow(1)
                If oRecordSet.RecordCount = 0 Then
                    ITForm.Freeze(False)
                    Exit Sub
                End If
                ITMatrix.Columns.Item("U_bomno").Editable = True
                ITMatrix.Columns.Item("U_bomlnid").Editable = True
                For i As Integer = 1 To oRecordSet.RecordCount
                    ITMatrix.Columns.Item("1").Cells.Item(i).Specific.value = oRecordSet.Fields.Item("u_itemcode").Value
                    ITMatrix.Columns.Item("U_qty").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("u_totqty").Value
                    ITMatrix.Columns.Item("U_bomno").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("DocNum").Value
                    ITMatrix.Columns.Item("U_bomlnid").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("LineId").Value
                    ITMatrix.Columns.Item("1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    oRecordSet.MoveNext()
                Next
                ITMatrix.Columns.Item("U_bomno").Editable = False
                ITMatrix.Columns.Item("U_bomlnid").Editable = False
                ITForm.Freeze(False)
            Catch ex As Exception
                ITForm.Freeze(False)
                oApplication.StatusBar.SetText(ex.Message)
            End Try

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub FilterItem(ByVal FormUID As String)
        Try

            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("6")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            Dim oRSets As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim h As String = Trim(objForm.Items.Item("sales").Specific.value)
            oRSets.DoQuery("Select * from OITM Where ItemCode IN (Select distinct B.U_itemcode from [@GEN_CUST_BOM] A inner join [@GEN_CUST_BOM_D0] B on A.DocEntry = B.DocEntry Where A.U_sono ='" & h & "')")
            'oRSets.DoQuery("Select B.U_itemcode from [@GEN_CUST_BOM] A inner join [@GEN_CUST_BOM_D0] B on A.DocEntry = B.DocEntry Where A.U_sono ='" & h & "'")

            Dim orsf As Integer = oRSets.RecordCount
            If orsf = 0 Then
                oCon = oCons.Add()
                oCon.Alias = "ItemCode"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = ""
            End If
            For IntICount As Integer = 0 To oRSets.RecordCount - 1
                If IntICount = (oRSets.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRSets.Fields.Item("ItemCode").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRSets.Fields.Item("ItemCode").Value
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                oRSets.MoveNext()
            Next
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub FilterItem1(ByVal FormUID As String)
        Try

            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("6")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            Dim oRSets As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim h As String = Trim(objForm.Items.Item("sales").Specific.value)
            oRSets.DoQuery("Select * from OITM ")

            Dim orsf As Integer = oRSets.RecordCount
            If orsf = 0 Then
                oCon = oCons.Add()
                oCon.Alias = "ItemCode"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = ""
            End If
            For IntICount As Integer = 0 To oRSets.RecordCount - 1
                If IntICount = (oRSets.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRSets.Fields.Item("ItemCode").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRSets.Fields.Item("ItemCode").Value
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                oRSets.MoveNext()
            Next
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
End Class
