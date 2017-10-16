Public Class ClsDelivery

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm, objSubForm, objSForm, objform1, objForm2 As SAPbouiCOM.Form
    Dim objItem, objOldItem As SAPbouiCOM.Item
    Dim TempItem As SAPbouiCOM.Item
    Dim objMatrix, objSubMatrix, objBatchMatrix As SAPbouiCOM.Matrix
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim SZDBHead As SAPbouiCOM.DBDataSource
    Dim SZDBDetail As SAPbouiCOM.DBDataSource
    Dim SMDBHead As SAPbouiCOM.DBDataSource
    Dim SMDBDetail As SAPbouiCOM.DBDataSource
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim RS1, RS2 As SAPbobsCOM.Recordset
    Dim ModalForm As Boolean = False
    Dim ChildModalForm As Boolean = False
    Dim PARENT_FORM As String
    Dim RowNo As Integer
    Dim orderno, hwid As String
    Dim sorderno, shwid, sitemcode As String
    Dim Mode As Integer
    Dim TotQty As Double
    Dim GSONO, GMACID, GITEMCODE, GASRTCODE As String
    Dim RowID As Integer
    Dim DeleteItemCode As String
    Dim InvNo As String
    Dim FrghtFlag As Boolean = False
    Dim PreInvNo As String
    Dim DOCNUM As String = ""
    Dim SELDOC As String = ""
    Dim Invtype As String = ""

    Dim Count As Double = 0
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            TempItem = objForm.Items.Item("86")
            objItem = objForm.Items.Add("spc", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = TempItem.Top + TempItem.Height + 5
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.Caption = "Unit"
            objItem.LinkTo = "86"
            TempItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("cpc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = TempItem.Top + TempItem.Height + 5
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.databind.setbound(True, "ODLN", "u_unit")
            objItem.Specific.TabOrder = TempItem.Specific.TabOrder + 1
            objItem.LinkTo = "46"

            TempItem = objForm.Items.Item("86")
            objItem = objForm.Items.Add("typeinv", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = TempItem.Top + TempItem.Height + 25
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.Caption = "Type of Invoice"
            objItem.LinkTo = "86"
            TempItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("tinv", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            objItem.Top = TempItem.Top + TempItem.Height + 25
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.databind.setbound(True, "ODLN", "u_tinv")
            objItem.Specific.TabOrder = TempItem.Specific.TabOrder + 1
            objItem.LinkTo = "46"

            'objOldItem = objForm.Items.Item("2")
            'objItem = objForm.Items.Add("btntq", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'objItem.Top = objOldItem.Top
            'objItem.Left = objOldItem.Left + objOldItem.Width + 5
            'objItem.Height = objOldItem.Height
            'objItem.Width = objOldItem.Width + 50
            'objItem.Specific.caption = "ShowTotalQuantity"
            'objItem.LinkTo = "2"
            ' objOldItem = objForm.Items.Item("btntq")


            TempItem = objForm.Items.Item("21")
            objItem = objForm.Items.Add("TQty", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = objForm.Items.Item("30").Top
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.Caption = "Total Quantity"
            objItem.LinkTo = "21"
            TempItem = objForm.Items.Item("222")
            objItem = objForm.Items.Add("tqty", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = objForm.Items.Item("29").Top
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.databind.setbound(True, "ODLN", "U_season")
            objItem.Specific.TabOrder = TempItem.Specific.TabOrder + 1
            objItem.LinkTo = "222"




            objOldItem = objForm.Items.Item("10000330")
            objItem = objForm.Items.Add("copyfr", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Width = objOldItem.Width
            objItem.Top = objOldItem.Top
            objItem.Left = objOldItem.Left
            objItem.Height = objOldItem.Height
            objItem.Specific.caption = "Copy From Pre-Ship"
            objItem.LinkTo = "10000330"
            objOldItem.Visible = False
            Me.SetChooseFromList(FormUID)


            objItem.Specific.ChooseFromListUID = "DPICFL"
            objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(FormUID)
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
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    If pVal.BeforeAction = True Then
                        Me.CreateForm(FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    If pVal.BeforeAction = True Then
                        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery("Select u_unit From [@GEN_USR_UNIT] Where u_user = '" + oCompany.UserName.ToString.Trim + "'")
                        If oRecordSet.RecordCount = 0 Then
                            oApplication.StatusBar.SetText("No PC assigned to User", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
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
                            If oCFL.UniqueID = "DPICFL" Then
                                For i As Integer = 0 To oDT.Rows.Count - 1
                                    DOCNUM = DOCNUM & "'" & Trim(oDT.GetValue("DocNum", i)) & "',"
                                Next
                                DOCNUM = DOCNUM.Substring(0, Len(DOCNUM) - 1)
                            End If
                        End If
                        If oCFL.UniqueID = "DPICFL" Then
                            Me.FilterSC(FormUID)
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = True Then
                        If Trim(objForm.Items.Item("cpc").Specific.value) = "" Then
                            oApplication.StatusBar.SetText("Please select Unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                    If pVal.ItemUID = "btntq" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("ODLN")
                        Count = 0
                        ' oDBs_Detail = objForm.DataSources.DBDataSources.Item("INV1")
                        Dim objMatrix As SAPbouiCOM.Matrix
                        objMatrix = objForm.Items.Item("38").Specific

                        For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                            Count = Count + Convert.ToDouble(objMatrix.Columns.Item("11").Cells.Item(Row).Specific.value)

                        Next
                        Dim a As String = oDBs_Head.GetValue("U_season", 0)
                        objForm.Items.Item("tqty").Specific.value = Count.ToString()
                        'SAPbouiCOM.EditText oEdit=SAPbouiCOM.
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "tinv" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = False Then
                        oCombo = objForm.Items.Item("tinv").Specific
                        Dim objMatrix As SAPbouiCOM.Matrix
                        objMatrix = objForm.Items.Item("38").Specific
                        Dim invoiceType As String = Trim(objForm.Items.Item("tinv").Specific.value)
                        Dim Unit As String = Trim(objForm.Items.Item("cpc").Specific.value)
                        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery("Select formatcode from oact inner join [@GEN_GL_ACCOUNT] T0 on oact.formatcode=T0.u_glacc where T0.U_tinv ='" & invoiceType & "' and T0.U_pc ='" & Unit & "'")

                        If objMatrix.Columns.Item("1").Cells.Item(1).Specific.value <> "" And oRecordSet.Fields.Count > 0 Then

                            For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                                objMatrix.Columns.Item("159").Cells.Item(Row).Specific.value = oRecordSet.Fields.Item("formatcode").Value
                            Next
                        End If
                        If Trim(oCombo.Selected.Value = "Direct") Then
                            objForm.Items.Item("copyfr").Visible = True
                            Invtype = " Direct"

                            'objForm.Items.Item("pre").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            'objForm.Items.Item("preship").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        Else
                            objForm.Items.Item("copyfr").Visible = False
                            objForm.Items.Item("10000330").Enabled = True
                            objForm.Items.Item("10000330").Visible = True
                            Invtype = ""
                            'objForm.Items.Item("pre").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            'objForm.Items.Item("preship").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If pVal.ItemUID = "38" And pVal.ColUID = "11" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                       
                    End If

                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    If pVal.FormTypeEx = "9999" And pVal.ItemUID = "7" And pVal.ColUID <> "0" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE And pVal.BeforeAction = True Then
                        objform1 = oApplication.Forms.ActiveForm
                        If objform1.Title = "List of Pre-Shipment" Then
                            Dim selmatrix As SAPbouiCOM.Matrix
                            selmatrix = objform1.Items.Item("7").Specific
                            SELDOC = ""
                            For sel As Integer = 1 To selmatrix.VisualRowCount
                                If selmatrix.IsRowSelected(sel) = True Then
                                    SELDOC = SELDOC & "'" & selmatrix.Columns.Item("DocNum").Cells.Item(sel).Specific.Value & "',"
                                End If
                            Next
                            SELDOC = SELDOC.Substring(0, Len(SELDOC) - 1)
                            Dim selRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            selRec.DoQuery("Select A.U_SONo From [@PRE_SHIPMENT_D0] A inner join [@PRE_SHIPMENT] B on A.DocEntry = B.DocEntry  Where B.DocNum IN (" & SELDOC & ") and A.U_SONo <> '' Group By A.U_SONo")
                            Dim num As Integer = selRec.RecordCount
                            objform1.Close()
                            objForm = oApplication.Forms.ActiveForm
                            If objForm.Title = "Pre-Shipment Invoice" Then
                                objForm = oApplication.Forms.GetForm("PRE_SHIPMENT", 1)
                                selRec.DoQuery("Select DocNum From [@PRE_SHIPMENT] Where DocNum IN (" & SELDOC & ")")
                                objForm.Items.Item("preno").Specific.Value = selRec.Fields.Item(0).Value
                                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Exit Sub
                            End If
                            objForm = oApplication.Forms.GetFormByTypeAndCount("140", 1)
                            oCombo = objForm.Items.Item("10000330").Specific
                            oCombo.Select("Sales Orders", SAPbouiCOM.BoSearchKey.psk_ByValue)

                            objform1 = oApplication.Forms.ActiveForm
                            objform1.Title = "List of Pre-Shipment"
                            Dim selmx As SAPbouiCOM.Matrix
                            selmx = objform1.Items.Item("7").Specific
                            Dim rowindex As Integer = 0
                            selmx.ClearSelections()
                            For selNo As Integer = 1 To selRec.RecordCount
                                If selmx.VisualRowCount = 1 Then
                                    objform1.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    Exit Sub
                                End If
                                For maxsel As Integer = 1 To selmx.VisualRowCount
                                    If selmx.Columns.Item("DocNum").Cells.Item(maxsel).Specific.Value = selRec.Fields.Item(0).Value Then
                                        Try
                                            Dim snoo As String = selRec.Fields.Item(0).Value
                                            selmx.Columns.Item("DocNum").Cells.Item(maxsel).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 4096)
                                            selRec.MoveNext()
                                            rowindex = rowindex + 1
                                            If rowindex = selRec.RecordCount Then
                                                objform1.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                Exit Sub
                                            End If
                                            Exit For
                                        Catch ex As Exception
                                            oApplication.SetStatusBarMessage(ex.Message)
                                        End Try
                                    End If
                                Next
                            Next
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
                        If objForm.TypeEx = "140" Then
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

    Sub ItemEvent_Pre(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Select Case pVal.EventType
            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                If pVal.FormTypeEx = "9999" And pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE And pVal.BeforeAction = True Then
                    objform1 = oApplication.Forms.ActiveForm
                    Try
                        If objform1.Title = "List of Pre-Shipment" Then
                            objform1.Freeze(True)
                            Dim selmatrix As SAPbouiCOM.Matrix
                            selmatrix = objform1.Items.Item("7").Specific
                            Dim row As Integer = selmatrix.VisualRowCount
                            SELDOC = ""
                            For sel As Integer = 1 To selmatrix.VisualRowCount
                                If selmatrix.IsRowSelected(sel) = True Then
                                    SELDOC = SELDOC & "'" & selmatrix.Columns.Item("DocNum").Cells.Item(sel).Specific.Value & "',"
                                End If
                            Next
                            SELDOC = SELDOC.Substring(0, Len(SELDOC) - 1)
                            objform1.Close()
                            Dim selRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            selRec.DoQuery("Select A.U_SONo From [@PRE_SHIPMENT_D0] A inner join [@PRE_SHIPMENT] B on A.DocEntry = B.DocEntry  Where B.DocNum IN (" & SELDOC & ") and A.U_SONo <> '' Group By A.U_SONo")
                            objForm = oApplication.Forms.ActiveForm
                            If objForm.Title = "Pre-Shipment Invoice" Then
                                objForm = oApplication.Forms.GetForm("PRE_SHIPMENT", 1)
                                selRec.DoQuery("Select DocNum From [@PRE_SHIPMENT] Where DocNum IN (" & SELDOC & ")")
                                objForm.Items.Item("preno").Specific.Value = selRec.Fields.Item(0).Value
                                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Exit Sub
                            End If
                            objForm = oApplication.Forms.GetFormByTypeAndCount("140", 1)
                            oCombo = objForm.Items.Item("10000330").Specific
                            oCombo.Select("Sales Orders", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            objform1 = oApplication.Forms.ActiveForm
                            objform1.Title = "List of Pre-Shipment"
                            Dim selmx As SAPbouiCOM.Matrix
                            selmx = objform1.Items.Item("7").Specific
                            Dim rowindex As Integer = 0
                            selmx.ClearSelections()
                            For selNo As Integer = 1 To selRec.RecordCount
                                If selmx.VisualRowCount = 1 Then
                                    objform1.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    Exit Sub
                                End If

                                For maxsel As Integer = 1 To selmx.VisualRowCount
                                    If selmx.Columns.Item("DocNum").Cells.Item(maxsel).Specific.Value = selRec.Fields.Item(0).Value Then
                                        Try
                                            Dim snoo As String = selRec.Fields.Item(0).Value
                                            selmx.Columns.Item("DocNum").Cells.Item(maxsel).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 4096)
                                            selRec.MoveNext()
                                            rowindex = rowindex + 1
                                            If rowindex = selRec.RecordCount Then
                                                objform1.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                Exit Sub
                                            End If
                                            Exit For
                                        Catch ex As Exception
                                            oApplication.SetStatusBarMessage(ex.Message)
                                        End Try
                                    End If
                                Next
                            Next
                            objform1.Freeze(False)
                        End If
                    Catch ex As Exception
                        objform1.Freeze(False)
                    End Try
                End If
                If pVal.FormTypeEx = "425" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ItemUID = "49" And pVal.BeforeAction = True Then
                    objForm2 = oApplication.Forms.ActiveForm
                    Try
                        objForm2.Freeze(True)
                        If Invtype = "Direct" Then
                            LoadItems_pre(FormUID, SELDOC, pVal.Row)
                        End If
                        objForm2.Freeze(False)
                    Catch ex As Exception
                        objForm2.Freeze(False)
                    End Try

                End If
                If pVal.FormTypeEx = "425" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ItemUID = "43" And pVal.ActionSuccess = True Then
                    objForm2 = oApplication.Forms.ActiveForm
                    If Invtype = "Direct" Then
                        objForm2.Items.Item("49").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                End If
            Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                If pVal.FormTypeEx = "9999" And pVal.ItemUID = "7" And pVal.ColUID <> "0" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE And pVal.BeforeAction = True Then
                    objform1 = oApplication.Forms.ActiveForm
                    If objform1.Title = "List of Pre-Shipment" Then
                        Dim selmatrix As SAPbouiCOM.Matrix
                        selmatrix = objform1.Items.Item("7").Specific
                        SELDOC = ""
                        For sel As Integer = 1 To selmatrix.VisualRowCount
                            If selmatrix.IsRowSelected(sel) = True Then
                                SELDOC = SELDOC & "'" & selmatrix.Columns.Item("DocNum").Cells.Item(sel).Specific.Value & "',"
                            End If
                        Next
                        SELDOC = SELDOC.Substring(0, Len(SELDOC) - 1)
                        Dim selRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        selRec.DoQuery("Select A.U_SONo From [@PRE_SHIPMENT_D0] A inner join [@PRE_SHIPMENT] B on A.DocEntry = B.DocEntry  Where B.DocNum IN (" & SELDOC & ") and A.U_SONo <> '' Group By A.U_SONo")
                        Dim num As Integer = selRec.RecordCount
                        objform1.Close()
                        objForm = oApplication.Forms.ActiveForm
                        If objForm.Title = "Pre-Shipment Invoice" Then
                            objForm = oApplication.Forms.GetForm("PRE_SHIPMENT", 1)
                            selRec.DoQuery("Select DocNum From [@PRE_SHIPMENT] Where DocNum IN (" & SELDOC & ")")
                            objForm.Items.Item("preno").Specific.Value = selRec.Fields.Item(0).Value
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Exit Sub
                        End If
                        objForm = oApplication.Forms.GetFormByTypeAndCount("133", 1)
                        oCombo = objForm.Items.Item("10000330").Specific
                        oCombo.Select("Sales Orders", SAPbouiCOM.BoSearchKey.psk_ByValue)

                        objform1 = oApplication.Forms.ActiveForm
                        objform1.Title = "List of Pre-Shipment"
                        Dim selmx As SAPbouiCOM.Matrix
                        selmx = objform1.Items.Item("7").Specific
                        Dim rowindex As Integer = 0
                        selmx.ClearSelections()
                        For selNo As Integer = 1 To selRec.RecordCount
                            If selmx.VisualRowCount = 1 Then
                                objform1.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Exit Sub
                            End If
                            For maxsel As Integer = 1 To selmx.VisualRowCount
                                If selmx.Columns.Item("DocNum").Cells.Item(maxsel).Specific.Value = selRec.Fields.Item(0).Value Then
                                    Try
                                        Dim snoo As String = selRec.Fields.Item(0).Value
                                        selmx.Columns.Item("DocNum").Cells.Item(maxsel).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 4096)
                                        selRec.MoveNext()
                                        rowindex = rowindex + 1
                                        If rowindex = selRec.RecordCount Then
                                            objform1.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            Exit Sub
                                        End If
                                        Exit For
                                    Catch ex As Exception
                                        oApplication.SetStatusBarMessage(ex.Message)
                                    End Try
                                End If
                            Next
                        Next
                    End If
                End If
        End Select
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
            oCFLCreationParams.ObjectType = "PRE_SHIPMENT"
            oCFLCreationParams.UniqueID = "DPICFL"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterSC(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("DPICFL")
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_CardCode"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = Trim(objForm.Items.Item("cardcode").Specific.Value)
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add()
            oCon.Alias = "U_Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Open"
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadItems_pre(ByVal FormUID As String, ByVal preNo As String, ByVal ROW As Integer)
        Try
            objForm2 = oApplication.Forms.GetFormByTypeAndCount("425", 1)
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select B.U_ItemCode,B.U_ItemName,B.U_Quantity,B.u_UOM,B.U_Price,B.U_TaxCode,A.DocNum,B.LineId,B.U_TotalLC,B.U_Whse,A.DocEntry,A.U_CustRef,A.U_DocCur,A.U_DelDate,A.U_Buyer,A.U_Owner,B.U_SONo From [@PRE_SHIPMENT] A Inner Join [@PRE_SHIPMENT_D0] B On A.DocEntry = B.DocEntry  And IsNull(A.U_Status,'Open') = 'Open'  And A.DocNum IN (" + preNo + ") and B.U_ItemCode <> '' and B.U_Quantity > 0 Order By  U_SONo ASC")
            Try
                objForm2.Freeze(True)
                Dim objmtx As SAPbouiCOM.Matrix = objForm2.Items.Item("3").Specific
                objForm2.Items.Item("3").Enabled = True
                objmtx.ClearSelections()
                Dim count As Integer = 0
                Dim totcount As Integer = 0
                For i As Integer = 1 To objmtx.VisualRowCount
                    If objmtx.Columns.Item("1").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("U_SONo").Value.ToString.Trim And objmtx.Columns.Item("2").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("U_ItemCode").Value.ToString.Trim Then
                        objmtx.Columns.Item("0").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular, SAPbouiCOM.BoModifiersEnum.mt_CTRL)
                        objmtx.Columns.Item("5").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("U_Quantity").Value.ToString.Trim
                        objmtx.Columns.Item("10").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("U_Price").Value.ToString.Trim
                        objmtx.Columns.Item("U_preno").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("DocNum").Value.ToString.Trim
                        'objForm2.Refresh()
                        count = count + 1
                        oRecordSet.MoveNext()
                    Else
                        totcount = totcount + 1
                    End If
                    If objmtx.VisualRowCount = count Then
                        Exit For
                    End If
                Next
                objForm2.Freeze(False)
            Catch ex As Exception
                objForm2.Freeze(False)
                oApplication.StatusBar.SetText(ex.Message)
            End Try

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
End Class
