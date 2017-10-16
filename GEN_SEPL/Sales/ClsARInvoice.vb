Public Class ClsARInvoice

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
    Dim oDBs_Head, oDBs_Detail As SAPbouiCOM.DBDataSource
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
    Dim DeleteFlag As Boolean = False
    Dim PreInvNo As String
    Dim loadcount As Integer = 0
    'Rajkumar
    Dim Invtype As String = ""
    Dim SONo As String
    Dim SALDOC As String = ""
    Dim DOCNUM As String = ""
    Dim SELDOC As String = ""
    Dim FormMode As String
    Dim BASENUM As String = ""
    Dim DbkVal, DbkPer, Dbk, ANSP, INS, COMM, TRANS As Double
    Dim ANSPCur As String
    Dim DocDate, TaxDate, DueDate As String
    Dim Fs_Source, Fs_PortOfLoading, Fs_PlaceOfSupplier As String
    Dim Fd_NoOfContainer As Decimal
    Dim COM, Curr As Double
    Dim LineTotalSum As Double
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
            objItem.Specific.databind.setbound(True, "OINV", "u_unit")
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
            objItem.Specific.databind.setbound(True, "OINV", "u_tinv")
            objItem.Specific.TabOrder = TempItem.Specific.TabOrder + 1
            objItem.LinkTo = "46"


            
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
            objItem.Specific.databind.setbound(True, "OINV", "U_season")
            objItem.Specific.TabOrder = TempItem.Specific.TabOrder + 1
            objItem.LinkTo = "222"

            'TempItem = objForm.Items.Item("15")
            'objItem = objForm.Items.Add("pre", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'objItem.Top = TempItem.Top + TempItem.Height + 25
            'objItem.Left = TempItem.Left
            'objItem.Width = TempItem.Width
            'objItem.Height = TempItem.Height
            'objItem.Specific.Caption = "Preshipment No."
            'objItem.LinkTo = "15"
            'TempItem = objForm.Items.Item("14")
            'objItem = objForm.Items.Add("preship", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'objItem.Top = TempItem.Top + TempItem.Height + 25
            'objItem.Left = TempItem.Left
            'objItem.Width = TempItem.Width
            'objItem.Height = TempItem.Height
            'objItem.Specific.databind.setbound(True, "OINV", "u_preship")
            'objItem.Specific.TabOrder = TempItem.Specific.TabOrder + 1
            'objItem.LinkTo = "14"

            'objForm.Items.Item("preship").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("preship").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'Rajkumar
            objForm = oApplication.Forms.Item(FormUID)
            objForm.Title = "Post Shipment Invoice"
            objOldItem = objForm.Items.Item("2")
            objItem = objForm.Items.Add("btnac", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Top = objOldItem.Top
            objItem.Left = objOldItem.Left + objOldItem.Width + 5
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width + 50
            objItem.Specific.caption = "Accruals & Expenses"
            objItem.LinkTo = "2"

            objOldItem = objForm.Items.Item("btnac")
            objItem = objForm.Items.Add("btnbc", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Top = objOldItem.Top
            objItem.Left = objOldItem.Left + objOldItem.Width + 5
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width + 10
            objItem.Specific.caption = "Book Consumption"
            objItem.LinkTo = "btnac"

            'objOldItem = objForm.Items.Item("btnac")
            'objItem = objForm.Items.Add("btntq", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'objItem.Top = objOldItem.Top
            'objItem.Left = objOldItem.Left + objOldItem.Width + 5
            'objItem.Height = objOldItem.Height
            'objItem.Width = objOldItem.Width + 50
            'objItem.Specific.caption = "ShowTotalQuantity"
            'objItem.LinkTo = "btnac"
            'Changed To Delivery LakkshmiKnatth

            'objOldItem = objForm.Items.Item("10000330")
            'objItem = objForm.Items.Add("copyfr", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'objItem.Width = objOldItem.Width
            'objItem.Top = objOldItem.Top
            'objItem.Left = objOldItem.Left
            'objItem.Height = objOldItem.Height
            'objItem.Specific.caption = "Copy From Pre-Ship"
            'objItem.LinkTo = "10000330"
            'objOldItem.Visible = False
            'Me.SetChooseFromList(FormUID)
            'objItem.Specific.ChooseFromListUID = "DPICFL"

            'Changed To Delivery
            'Rajkumar'
            objItem = objForm.Items.Item("10")
            objItem.BackColor = RGB(0, 0, 0)
            objItem.ForeColor = RGB(255, 255, 255)

            Dim oRs2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs2.DoQuery("Delete From [@GEN_ACCRUALS]")
            'Rajkumar'

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try

            objForm = oApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("38").Specific
            If objForm.Items.Item("3").Specific.value = "I" And (Trim(objForm.Items.Item("tinv").Specific.value) = "" Or Trim(objForm.Items.Item("tinv").Specific.value) = "-") Then
                oApplication.StatusBar.SetText("Please Select Type Of Invoice", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            Return True
        Catch ex As Exception

        End Try
    End Function

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            'Rajkumar
            Dim CHILD_FORM As String = "ACCRUALS@" & pVal.FormUID
            Dim ModalForm As Boolean = False
            For i As Integer = 0 To oApplication.Forms.Count - 1
                If oApplication.Forms.Item(i).UniqueID = (CHILD_FORM) Then
                    objSubForm = oApplication.Forms.Item("ACCRUALS@" & pVal.FormUID)
                    ModalForm = True
                    Exit For
                End If
            Next
            'Rajkumar

            If ModalForm = False Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                        objForm = oApplication.Forms.Item(FormUID)
                        If FrghtFlag = True Then
                            LoadFreight(FormUID)
                        End If
                        If pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.InnerEvent = False Then
                            If pVal.ItemUID = "38" And pVal.ColUID = "11" Then
                                objMatrix.Columns.Item("11").Cells.Item(1).Specific.value = "40"
                            End If
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
                       
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        If pVal.BeforeAction = True Then
                            Me.CreateForm(FormUID)
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        If pVal.ItemUID = "39" And pVal.ColUID = "94" And pVal.BeforeAction = False Then
                            Dim USER_NAME As String = oCompany.UserName
                            If USER_NAME <> "manager" And objForm.Items.Item("3").Specific.Value = "S" Then
                                Dim GLAccount As String
                                objMatrix = objForm.Items.Item("39").Specific
                                For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                                    GLAccount = objMatrix.Columns.Item("94").Cells.Item(Row).Specific.value
                                    If GLAccount <> "" Then


                                        Dim str As String = "Select (substring('" + GLAccount + "', 1, len('" + GLAccount + "')-3)+RIGHT('" + GLAccount + "',2))"
                                        Dim GLacc As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        GLacc.DoQuery("Select (substring('" + GLAccount + "', 1, len('" + GLAccount + "')-3)+RIGHT('" + GLAccount + "',2))")
                                        Dim ManualJE As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        ManualJE.DoQuery("Select count(Code) From [@GEN_M_JE] where code='" + GLacc.Fields.Item(0).Value + "'")
                                        If ManualJE.Fields.Item(0).Value > 0 Then
                                            oApplication.StatusBar.SetText("You Are Not Permitted To Perform This Action", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            objMatrix.Columns.Item("94").Cells.Item(Row).Specific.value = ""
                                            Exit Sub
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        'If pVal.ItemUID = "39" And pVal.ColUID = "94" And pVal.BeforeAction = True Then
                        '    Dim USER_NAME As String = oCompany.UserName
                        '    If USER_NAME = "manager" And objForm.Items.Item("3").Specific.Value = "S" Then
                        '        Dim GLAccount As String
                        '        objMatrix = objForm.Items.Item("39").Specific
                        '        For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                        '            GLAccount = objMatrix.Columns.Item("94").Cells.Item(Row).Specific.value
                        '            Dim str As String = "Select (substring('" + GLAccount + "', 1, len('" + GLAccount + "')-3)+RIGHT('" + GLAccount + "',2))"
                        '            Dim GLacc As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '            GLacc.DoQuery("Select (substring('" + GLAccount + "', 1, len('" + GLAccount + "')-3)+RIGHT('" + GLAccount + "',2))")
                        '            Dim ManualJE As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '            ManualJE.DoQuery("Select count(Code) From [@GEN_M_JE] where code='" + GLacc.Fields.Item(0).Value + "'")
                        '            If ManualJE.Fields.Item(0).Value > 0 Then
                        '                oApplication.StatusBar.SetText("You Are Not Permitted To Perform This Action", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '                BubbleEvent = False
                        '                Exit Sub
                        '            End If
                        '        Next
                        '    End If
                        'End If
                        If pVal.BeforeAction = True Then
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("Select u_unit From [@GEN_USR_UNIT] Where u_user = '" + oCompany.UserName.ToString.Trim + "'")
                            If oRecordSet.RecordCount = 0 Then
                                oApplication.StatusBar.SetText("No PC assigned to User", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If

                            'Rajkumar
                            objForm = oApplication.Forms.Item(FormUID)
                            Dim oCFL As SAPbouiCOM.ChooseFromList
                            Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                            Dim CFL_Id As String
                            CFL_Id = CFLEvent.ChooseFromListUID
                            oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                            Dim oDT As SAPbouiCOM.DataTable
                            oDT = CFLEvent.SelectedObjects
                            If oCFL.UniqueID = "DPICFL" Then
                                Me.FilterSC(FormUID)
                            End If
                            'Rajkumar
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
                                If oCFL.UniqueID = "DPICFL" Then
                                    For i As Integer = 0 To oDT.Rows.Count - 1
                                        DOCNUM = DOCNUM & "'" & Trim(oDT.GetValue("DocNum", i)) & "',"
                                    Next
                                    DOCNUM = DOCNUM.Substring(0, Len(DOCNUM) - 1)
                                End If
                                If oCFL.UniqueID = "3" Then
                                    Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRSet.DoQuery("Select u_unit From OCRD Where CardCode = '" + Trim(oDT.GetValue("CardCode", 0)) + "'")
                                    objForm.Items.Item("cpc").Specific.value = oRSet.Fields.Item("u_unit").Value
                                    objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                End If
                                'If oCFL.UniqueID = "9" Then
                                '    Dim GlAccountName As String = Trim(oDT.GetValue("AcctName", 0))
                                '    Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                '    oRSet.DoQuery("select Segment_0+Segment_1  as Code from OACT where AcctName ='" + GlAccountName + "'")
                                '    Dim GlAccount As String = oRSet.Fields.Item("Code").Value

                                '    'Dim str As String = "Select (substring('" + GlAccount + "', 1, len('" + GlAccount + "')-3)+RIGHT('" + GlAccount + "',2))"
                                '    'Dim GLacc As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                '    'GLacc.DoQuery("Select (substring('" + GlAccount + "', 1, len('" + GlAccount + "')-3)+RIGHT('" + GlAccount + "',2))")
                                '    Dim ManualJE As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                '    ManualJE.DoQuery("Select count(Code) From [@GEN_M_JE] where code='" + GlAccount + "'")
                                '    If ManualJE.Fields.Item(0).Value > 0 Then
                                '        oApplication.StatusBar.SetText("You Are Not Permitted To Perform This Action", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                '        BubbleEvent = False


                                '    End If

                                'End If
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        Dim USER_NAME As String = oCompany.UserName
                        objMatrix = objForm.Items.Item("38").Specific
                        objSubMatrix = objForm.Items.Item("39").Specific

                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = True Then
                            SONo = objForm.Items.Item("8").Specific.Value
                            DocDate = objForm.Items.Item("10").Specific.value
                            TaxDate = objForm.Items.Item("10").Specific.value
                            DueDate = objForm.Items.Item("10").Specific.value

                            'User Authorisations
                            Dim oUnit As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oUnit.DoQuery("Select u_unit from [@GEN_USER_AUTH] where u_user='" + oCompany.UserName + "'")
                            For i As Integer = 1 To oUnit.RecordCount
                                If i <> oUnit.RecordCount Then
                                    If objForm.Items.Item("cpc").Specific.value <> oUnit.Fields.Item(0).Value Then
                                        oUnit.MoveNext()
                                    End If
                                Else
                                    If objForm.Items.Item("cpc").Specific.value <> oUnit.Fields.Item(0).Value Then
                                        oApplication.StatusBar.SetText("Please Select Correct UNIT", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Next

                            If Trim(objForm.Items.Item("cpc").Specific.value) = "" Then
                                oApplication.StatusBar.SetText("Please select Unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                            loadcount = 0
                            If Me.Validation(FormUID) = False Then
                                BubbleEvent = False
                                'Exit Select
                            End If

                            'Rajkumar
                            PreInvNo = objForm.Items.Item("8").Specific.value
                            If oApplication.MessageBox("Whether You Have Booked The Accruals And Filled All the Other Information For This Particular Invoice?", 2, "Yes", "No") = 2 Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                            'Rajkumar

                            If Trim(objForm.Items.Item("3").Specific.value) = "S" Then
                                If USER_NAME <> "manager" Then
                                    Dim GLAccount As String
                                    For Row As Integer = 1 To objSubMatrix.VisualRowCount - 1
                                        GLAccount = objSubMatrix.Columns.Item("94").Cells.Item(Row).Specific.value
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
                            For j As Integer = 1 To objMatrix.VisualRowCount - 1
                                If Trim(objForm.Items.Item("tinv").Specific.value) = "Local" Then
                                    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRs.DoQuery("Select u_locgl from OWHS WHERE whscode='" + Trim(objMatrix.Columns.Item("24").Cells.Item(j).Specific.value) + "'")
                                    oRs1.DoQuery("Select formatcode from oact inner join [@GEN_GL_ACCOUNT] T0 on oact.formatcode=T0.u_glacc where formatcode='" + oRs.Fields.Item(0).Value + "' ")
                                    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                                        objMatrix.Columns.Item("159").Cells.Item(i).Specific.value = oRs1.Fields.Item(0).Value
                                    Next
                                    Exit For
                                End If
                                If Trim(objForm.Items.Item("tinv").Specific.value) = "Deemed" Then
                                    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRs.DoQuery("Select u_deemgl from OWHS WHERE whscode='" + Trim(objMatrix.Columns.Item("24").Cells.Item(j).Specific.value) + "'")
                                    oRs1.DoQuery("Select formatcode from oact inner join [@GEN_GL_ACCOUNT] T0 on oact.formatcode=T0.u_glacc where formatcode='" + oRs.Fields.Item(0).Value + "' ")
                                    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                                        objMatrix.Columns.Item("159").Cells.Item(i).Specific.value = oRs1.Fields.Item(0).Value
                                    Next
                                    Exit For
                                End If
                                If Trim(objForm.Items.Item("tinv").Specific.value) = "Direct" Then
                                    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRs.DoQuery("Select u_dirgl from OWHS WHERE whscode='" + Trim(objMatrix.Columns.Item("24").Cells.Item(j).Specific.value) + "'")
                                    oRs1.DoQuery("Select formatcode from oact inner join [@GEN_GL_ACCOUNT] T0 on oact.formatcode=T0.u_glacc where formatcode='" + oRs.Fields.Item(0).Value + "' ")
                                    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                                        objMatrix.Columns.Item("159").Cells.Item(i).Specific.value = oRs1.Fields.Item(0).Value
                                    Next
                                    Exit For
                                End If
                            Next

                            'Rajkumar
                            'If Trim(objForm.Items.Item("tinv").Specific.value) = "Local" Then
                            '    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            '    Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            '    oRs.DoQuery("Select u_locgl from OWHS WHERE whscode='" + Trim(objForm.Items.Item("cpc").Specific.value) + "'")
                            '    oRs1.DoQuery("Select formatcode from oact inner join [@GEN_GL_ACCOUNT] T0 on oact.formatcode=T0.u_glacc where formatcode='" + oRs.Fields.Item(0).Value + "' ")
                            '    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                            '        objMatrix.Columns.Item("159").Cells.Item(i).Specific.value = oRs1.Fields.Item(0).Value
                            '    Next
                            'End If
                            'If Trim(objForm.Items.Item("cpc").Specific.value) = "UNIT2" Then
                            '    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            '    Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            '    oRs.DoQuery("Select u_glacc from [@GEN_GL_ACCOUNT] WHERE u_pc='" + Trim(objForm.Items.Item("cpc").Specific.value) + "' and u_tinv='" + Trim(objForm.Items.Item("tinv").Specific.value) + "'")
                            '    oRs1.DoQuery("Select formatcode from oact inner join [@GEN_GL_ACCOUNT] T0 on oact.formatcode=T0.u_glacc where formatcode='" + oRs.Fields.Item(0).Value + "' ")
                            '    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                            '        objMatrix.Columns.Item("159").Cells.Item(i).Specific.value = oRs1.Fields.Item(0).Value
                            '    Next
                            'End If
                            'If Trim(objForm.Items.Item("cpc").Specific.value) = "UNIT3" Then
                            '    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            '    Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            '    oRs.DoQuery("Select u_glacc from [@GEN_GL_ACCOUNT] WHERE u_pc='" + Trim(objForm.Items.Item("cpc").Specific.value) + "' and u_tinv='" + Trim(objForm.Items.Item("tinv").Specific.value) + "'")
                            '    oRs1.DoQuery("Select formatcode from oact inner join [@GEN_GL_ACCOUNT] T0 on oact.formatcode=T0.u_glacc where formatcode='" + oRs.Fields.Item(0).Value + "' ")
                            '    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                            '        objMatrix.Columns.Item("159").Cells.Item(i).Specific.value = oRs1.Fields.Item(0).Value
                            '    Next
                            'End If

                        End If

                        'Rajkumar
                        '            If pVal.ItemUID = "btnsp" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) And pVal.BeforeAction = False Then
                        '                objMatrix = objForm.Items.Item("38").Specific
                        '                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '                For i As Integer = 1 To objMatrix.VisualRowCount
                        '                    oRSet.DoQuery("Select Distinct B.u_price From [@GEN_SUPP_PRICE] A INNER JOIN [@GEN_SUPP_PRICE_D0] B ON A.CODE = B.CODE Where A.u_cardcode = '" + Trim(objForm.Items.Item("4").Specific.value) + "' And B.u_itemcode = '" + Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) + "'")
                        '                    If oRSet.RecordCount > 0 Then
                        '                        objMatrix.Columns.Item("U_sprice").Cells.Item(i).Specific.value = oRSet.Fields.Item("u_price").Value
                        '                    End If
                        '                Next
                        '            End If
                        'If pVal.ItemUID = "btfrt" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = False Then
                        '    objForm.Freeze(True)
                        '    oApplication.SetStatusBarMessage("Please wait...", BoMessageTime.bmt_Medium, False)
                        '    LoadFreight_Pre(FormUID)
                        '    oApplication.SetStatusBarMessage("Update the freight charges", BoMessageTime.bmt_Short, False)
                        '    objForm.Freeze(False)
                        'End If

                        If pVal.ItemUID = "btnbc" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            If pVal.BeforeAction = False Then
                                Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("Select B.ItemCode,B.WhsCode,C.Code, B.Quantity * C.Quantity + ((B.Quantity * C.Quantity * IsNull(C.u_tol,0))/100) As 'Qty' From OINV A Inner Join INV1 B On A.DocEntry = B.DocEntry Inner Join OITM D On B.ItemCode = D.ItemCode Inner Join ITT1 C On B.ItemCode = C.Father Where A.DocNum = '" + Trim(objForm.Items.Item("8").Specific.value) + "' And IsNUll(D.u_bc,'N') = 'Y' ")
                                If oRecordSet.RecordCount > 0 Then
                                    oApplication.ActivateMenuItem("3079")
                                    oRecordSet.MoveFirst()
                                    Dim GIFOrm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                                    GIFOrm.Items.Item("einv").Specific.Value = objForm.Items.Item("8").Specific.value
                                    Dim GIMatrix As SAPbouiCOM.Matrix = GIFOrm.Items.Item("13").Specific
                                    For i As Integer = 1 To oRecordSet.RecordCount
                                        GIMatrix.Columns.Item("1").Cells.Item(i).Specific.value = oRecordSet.Fields.Item("Code").Value
                                        GIMatrix.Columns.Item("9").Cells.Item(i).Specific.value = Math.Round(oRecordSet.Fields.Item("Qty").Value, 0)
                                        GIMatrix.Columns.Item("15").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("WhsCode").Value
                                        oRecordSet.MoveNext()
                                    Next
                                End If
                            End If
                            If pVal.BeforeAction = True Then
                                Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("Select DocNum From OIGE Where u_invno = '" + Trim(objForm.Items.Item("8").Specific.value) + "'")
                                If oRecordSet.RecordCount > 0 Then
                                    oApplication.StatusBar.SetText("Comsumption already done for this invoice", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                        End If
                        If pVal.ItemUID = "2" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            oRecordSet.DoQuery("Delete From [@GEN_ACCRUALS] Where u_invno = '" + Trim(objForm.Items.Item("8").Specific.Value) + "' And u_macid = '" + MAC_ID + "'")
                        End If


                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess = True Then
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            'invoice_ref(pVal.FormUID, Trim(objForm.Items.Item("88").Specific.Value))
                            oRecordSet.DoQuery("Delete From [@GEN_ACCRUALS] Where u_invno = '" + PreInvNo + "' And u_macid = '" + MAC_ID + "'")
                            oRecordSet.DoQuery("Select B.ItemCode,B.WhsCode,C.Code, B.Quantity * C.Quantity + ((B.Quantity * C.Quantity * IsNull(C.u_tol,0))/100) As 'Qty' From OINV A Inner Join INV1 B On A.DocEntry = B.DocEntry Inner Join OITM D On B.ItemCode = D.ItemCode Inner Join ITT1 C On B.ItemCode = C.Father Where A.DocNum = '" + PreInvNo + "' And IsNUll(D.u_bc,'N') = 'Y' ")
                            If oRecordSet.RecordCount > 0 Then
                                oApplication.ActivateMenuItem("3079")
                                oRecordSet.MoveFirst()
                                Dim GIFOrm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                                GIFOrm.Items.Item("einv").Specific.Value = PreInvNo
                                Dim GIMatrix As SAPbouiCOM.Matrix = GIFOrm.Items.Item("13").Specific
                                For i As Integer = 1 To oRecordSet.RecordCount
                                    GIMatrix.Columns.Item("1").Cells.Item(i).Specific.value = oRecordSet.Fields.Item("Code").Value
                                    GIMatrix.Columns.Item("9").Cells.Item(i).Specific.value = Math.Round(oRecordSet.Fields.Item("Qty").Value, 0)
                                    GIMatrix.Columns.Item("15").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("WhsCode").Value
                                    oRecordSet.MoveNext()
                                Next
                            End If
                            loadcount = 0
                        End If
                        If pVal.ItemUID = "btntq" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("OINV")
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
                        If pVal.ItemUID = "btnac" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = False Then
                            InvNo = objForm.Items.Item("8").Specific.Value
                            objform1 = oApplication.Forms.GetFormByTypeAndCount("133", 1)
                            objMatrix = objForm.Items.Item("38").Specific
                            Dbk = 0
                            ANSP = 0
                            COMM = 0
                            TRANS = 0
                            LineTotalSum = 0
                            Dim oDBDataSource As SAPbouiCOM.DBDataSource

                            oDBDataSource = objForm.DataSources.DBDataSources.Item("OINV")
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("INV1")
                            ''Added By Rajkumar
                            Fs_PortOfLoading = oDBDataSource.GetValue("U_PortLoad", 0).Trim
                            Fd_NoOfContainer = oDBDataSource.GetValue("U_NO_OF_CN", 0).Trim
                            Fs_PlaceOfSupplier = oDBDataSource.GetValue("U_SUPP_PLC1", 0).Trim
                            Fs_Source = objMatrix.Columns.Item("U_Source").Cells.Item(1).Specific.Selected.Value
                            Curr = oDBDataSource.GetValue("DocRate", 0).Trim

                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            For i As Integer = 1 To objMatrix.VisualRowCount
                                If Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) <> "" And Trim(objMatrix.Columns.Item("14").Cells.Item(i).Specific.value) <> "" Then
                                    oRSet.DoQuery("Select Distinct IsNull(u_expval,0) 'u_expval',isnull(u_expper,0) 'u_expper',ISNULL(u_cap,0) 'u_cap' From OITM Where ItemCode = '" + Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) + "'")
                                    Dbk = Dbk + (oRSet.Fields.Item("u_expval").Value * objMatrix.Columns.Item("11").Cells.Item(i).Specific.value)
                                    If (oRSet.Fields.Item("u_expper").Value * objMatrix.Columns.Item("11").Cells.Item(i).Specific.value * objMatrix.Columns.Item("14").Cells.Item(i).Specific.value.ToString.Substring(3) * Curr) / 100 > oRSet.Fields.Item("u_cap").Value * objMatrix.Columns.Item("11").Cells.Item(i).Specific.value Then
                                        Dbk = oRSet.Fields.Item("u_cap").Value * objMatrix.Columns.Item("11").Cells.Item(i).Specific.value
                                    Else
                                        Dbk = ((oRSet.Fields.Item("u_expper").Value) * objMatrix.Columns.Item("11").Cells.Item(i).Specific.value * Curr * objMatrix.Columns.Item("14").Cells.Item(i).Specific.value.ToString.Substring(3)) / 100

                                    End If
                                    objMatrix.Columns.Item("U_dbkval").Cells.Item(i).Specific.value = Dbk
                                    ''(oRSet.Fields.Item("u_expval").Value * objMatrix.Columns.Item("11").Cells.Item(i).Specific.value)
                                    oRSet.DoQuery("Select B.u_ansp,A.u_doccur,B.u_comm From [@GEN_SUPP_PRICE] A INNER JOIN [@GEN_SUPP_PRICE_D0] B On A.Code = B.Code Where B.u_ItemCode = '" + Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) + "' And A.u_cardcode = '" + Trim(objForm.Items.Item("4").Specific.value) + "'")
                                    ANSP = ANSP + (objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value * oRSet.Fields.Item("u_ansp").Value)
                                    ANSPCur = oRSet.Fields.Item("u_doccur").Value
                                    COMM = COMM + (objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value * oRSet.Fields.Item("u_comm").Value)
                                    objMatrix.Columns.Item("U_anspval").Cells.Item(i).Specific.value = (objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value * oRSet.Fields.Item("u_ansp").Value)
                                    objMatrix.Columns.Item("U_comm").Cells.Item(i).Specific.value = (objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value * oRSet.Fields.Item("u_comm").Value)
                                End If
                            Next
                            ''oRSet.DoQuery("Select A.U_TotFrgn from [@PRE_SHIPMENT_D3] A left outer  join [@PRE_SHIPMENT] B ON A.DocEntry = B.DocEntry Where B.DocNum IN (" + SALDOC + ") and A.U_ExpnCode = '15'")
                            ''TRANS = oRSet.Fields.Item(0).Value

                            objMatrix = objForm.Items.Item("38").Specific
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                If Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) <> "" And Trim(objMatrix.Columns.Item("14").Cells.Item(i).Specific.value) <> "" Then
                                    LineTotalSum = LineTotalSum + (CDbl(objForm.Items.Item("64").Specific.value) * CDbl(objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value.ToString.Substring(4)) * objMatrix.Columns.Item("11").Cells.Item(i).Specific.value)
                                End If
                            Next
                            'oRSet.DoQuery("Select IsNull(PrintHeadr,0) 'Per1',IsNull(Manager,0) 'Per2' From OADM")
                            'INS = (LineTotalSum * CDbl(oRSet.Fields.Item("Per1").Value) * CDbl(oRSet.Fields.Item("Per2").Value)) / 100
                            oRSet.DoQuery("Select IsNull(u_comper,0) AS 'COMPER' From OCRD Where CardCode = '" + Trim(objForm.Items.Item("4").Specific.value) + "'")
                            If CDbl(oRSet.Fields.Item("COMPER").Value) > 0 Then
                                COM = LineTotalSum * CDbl(oRSet.Fields.Item("COMPER").Value) / 100
                            End If
                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            'Dim Route As String = objform1.Items.Item("U_MulRote").Specific.Value
                            Dim Route As String = ""
                            Dim TMPSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            TMPSet.DoQuery("Delete From OINV_BASENUM WHere macid = '" + MAC_ID + "'")
                            For i As Integer = 1 To objMatrix.VisualRowCount - 1
                                TMPSet.DoQuery("Insert Into OINV_BASENUM(invno,basenum,macid) Values('" + InvNo + "','" + Trim(objMatrix.Columns.Item("44").Cells.Item(i).Specific.value) + "','" + MAC_ID + "')")
                            Next
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                FormMode = "A"
                            End If
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                FormMode = "O"
                            End If
                            If Fs_PlaceOfSupplier = "" Or Fs_PortOfLoading = "" Then
                                oApplication.SetStatusBarMessage("Source and Port of Loading should not be empty", SAPbouiCOM.BoMessageTime.bmt_Short)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            Me.Open_Accruals_Form(pVal.FormUID, InvNo, MAC_ID, FormMode, Route)
                        End If
                        'Rajkumar
                End Select
            ElseIf pVal.BeforeAction = True And (ModalForm = True) And (pVal.FormUID = (objSubForm.UniqueID.Substring(objSubForm.UniqueID.IndexOf("@") + 1))) Then
                objSubForm = oApplication.Forms.Item("ACCRUALS@" & pVal.FormUID)
                objSubForm.Select()
                BubbleEvent = False
            End If
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
                        If objForm.TypeEx = "133" Then
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
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.BeforeAction = False Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        Dim recupdt As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim recsel As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("OINV")
                        oDBs_Detail = objForm.DataSources.DBDataSources.Item("INV1")
                        objMatrix = objForm.Items.Item("38").Specific
                        Dim chk As Integer = 0


                        'If Trim(objForm.Items.Item("3").Specific.Selected.Value) = "I" Then
                        '    oRSet.DoQuery("Select UserID From OUSR Where User_Code = '" + oCompany.UserName.ToString.Trim + "'")
                        '    UserSign = oRSet.Fields.Item("UserID").Value
                        '    oRSet.DoQuery("Select Max(DocEntry) As 'DocEntry' From OINV Where UserSign = '" + UserSign + "'")
                        '    DocNum = oRSet.Fields.Item("DocEntry").Value
                        '    oRSet.DoQuery("Select DocEntry,u_unit,DocCur,DocRate From OINV Where DocEntry = '" + DocNum + "'")
                        '    DocEntry = oRSet.Fields.Item("DocEntry").Value
                        '    PC = oRSet.Fields.Item("u_unit").Value
                        '    oRecordSet.DoQuery("Select ItemCode,u_preship,LineNum from OINV A inner join INV1 B on A.DocEntry = B.DocEntry Where A.DocEntry = '" + DocNum + "'")
                        '    Dim PreNum As Integer = oRecordSet.Fields.Item(1).Value
                        '    oRSet.DoQuery("Select B.U_ItemCode,A.DocNum,B.LineId,U_BaseLine From [@PRE_SHIPMENT] A Inner Join [@PRE_SHIPMENT_D0] B On A.DocEntry = B.DocEntry    And A.DocNum = " + PreNum + "' and B.U_ItemCode <> ''")
                        '    'Rajkumar
                        '    If oRSet.RecordCount = oRecordSet.RecordCount Then
                        '        Invoice.UserFields.Fields.Item("U_oinvno").Value = DocNum
                        '        Invoice.UserFields.Fields.Item("U_unit").Value = PC
                        '        Invoice.Lines.ItemCode = oRecordSet.Fields.Item("U_ItemCode").Value.ToString.Trim
                        '        Invoice.Lines.BaseLine = oRecordSet.Fields.Item("LineId").Value
                        '        Invoice.Lines.BaseType = "17"
                        '        Invoice.Lines.Add()
                        '        Invoice.Lines.SetCurrentLine(1)
                        '        If Invoice.Update() <> 0 Then
                        '            If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        '            oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription)
                        '            Exit Select
                        '        Else
                        '            If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        '            oRSet.DoQuery("Select DocNum From OINV Where DocEntry = '" + oCompany.GetNewObjectKey.ToString + "'")
                        '            Dim strprint As String
                        '            strprint = "Document Nos : " & objForm.Items.Item("8").Specific.value & " & " & oRSet.Fields.Item("DocNum").Value & " are created"
                        '            oApplication.MessageBox(strprint)
                        '        End If
                        '    End If
                        '    Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                        '    Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '    oRSet.DoQuery("Select UserId From OUSR Where User_Code = '" + oCompany.UserName.Trim + "'")
                        '    oRS.DoQuery("Select Max(DocEntry) AS 'DocEntry' From OINV Where UserSign = '" + Trim(oRSet.Fields.Item("UserID").Value) + "'")
                        '    oRSet.DoQuery("Select DocNum From  OINV Where DocEntry = '" + Trim(oRS.Fields.Item("DocEntry").Value) + "'")
                        'End If

                        'Added By Rajkumar
                        For j As Integer = 0 To objMatrix.VisualRowCount - 2
                            'Dim PreDoc As String = (oDBs_Detail.GetValue("U_preno", j).Trim).Substring(0, Len(oDBs_Detail.GetValue("U_preno", j).Trim) - 1)
                            Dim PreDoc As String = "Select B.U_Quantity,B.U_Status,A.DocEntry,A.DocNum from [@PRE_SHIPMENT] A inner join [@PRE_SHIPMENT_D0] B on A.DocEntry = B.DocEntry Where A.Docnum = '" & oDBs_Head.GetValue("u_preship", 0).Trim & "' and B.U_Quantity = '" + oDBs_Detail.GetValue("Quantity", j).Trim + "' and A.U_CustCode = '" + oDBs_Head.GetValue("CardCode", 0).Trim + "' and B.U_ItemCode = '" + oDBs_Detail.GetValue("ItemCode", j).Trim + "'"
                            Dim st As String = "Select B.U_Quantity,B.U_Status,A.DocEntry,A.DocNum from [@PRE_SHIPMENT] A inner join [@PRE_SHIPMENT_D0] B on A.DocEntry = B.DocEntry Where A.U_CustRef = '" & oDBs_Head.GetValue("u_preship", 0).Trim & "' and B.U_Quantity = '" + oDBs_Detail.GetValue("Quantity", j).Trim + "' and A.U_CustCode = '" + oDBs_Head.GetValue("CardCode", 0).Trim + "' and B.U_ItemCode = '" + oDBs_Detail.GetValue("ItemCode", j).Trim + "'"
                            Dim st1 As String = oDBs_Detail.GetValue("u_preno", j).Trim
                            recsel.DoQuery("Select B.U_Quantity,B.U_Status,A.DocEntry,A.DocNum from [@PRE_SHIPMENT] A inner join [@PRE_SHIPMENT_D0] B on A.DocEntry = B.DocEntry Where A.DocNum = '" & oDBs_Detail.GetValue("u_preno", j).Trim & "' and B.U_Quantity = '" + oDBs_Detail.GetValue("Quantity", j).Trim + "' and A.U_CustCode = '" + oDBs_Head.GetValue("CardCode", 0).Trim + "' and B.U_ItemCode = '" + oDBs_Detail.GetValue("ItemCode", j).Trim + "'")
                            Dim salqty, preqty, balqty As Double
                            Dim prno As String = recsel.Fields.Item("DocEntry").Value
                            Dim Prenos As String = recsel.Fields.Item("DocNum").Value
                            salqty = recsel.Fields.Item(0).Value
                            preqty = oDBs_Detail.GetValue("Quantity", j).Trim
                            balqty = salqty - preqty
                            If balqty = 0 Then
                                recupdt.DoQuery("Update [@PRE_SHIPMENT_D0] Set U_Status = 'Closed',U_invqty = '" + oDBs_Detail.GetValue("Quantity", j).Trim + "' Where DocEntry = '" + prno + "' and U_ItemCode = '" + oDBs_Detail.GetValue("ItemCode", j).Trim + "'")
                            Else
                                recupdt.DoQuery("Update [@PRE_SHIPMENT_D0] Set U_Status = 'Open',U_invqty = '" + oDBs_Detail.GetValue("Quantity", j).Trim + "' Where DocEntry = '" + prno + "' and U_ItemCode = '" + oDBs_Detail.GetValue("ItemCode", j).Trim + "'")
                            End If
                            'recsel.DoQuery("Select B.U_Quantity,B.U_Status,A.DocEntry,A.DocNum from [@PRE_SHIPMENT] A inner join [@PRE_SHIPMENT_D0] B on A.DocEntry = B.DocEntry Where A.U_CustRef = '" & oDBs_Head.GetValue("u_preship", 0).Trim & "' and B.U_Quantity = '" + oDBs_Detail.GetValue("Quantity", j).Trim + "' and A.U_CustCode = '" + oDBs_Head.GetValue("CardCode", 0).Trim + "' and B.U_ItemCode = '" + oDBs_Detail.GetValue("ItemCode", j).Trim + "'")

                            recsel.DoQuery("Select B.U_Quantity,B.U_Status,A.DocEntry,A.DocNum from [@PRE_SHIPMENT] A inner join [@PRE_SHIPMENT_D0] B on A.DocEntry = B.DocEntry Where A.Docnum = '" & oDBs_Detail.GetValue("u_preno", j).Trim & "' and B.U_Quantity = CONVERT(decimal, REPLACE('" + oDBs_Detail.GetValue("Quantity", j).Trim + "',',','')) and A.U_CustCode = '" + oDBs_Head.GetValue("CardCode", 0).Trim + "' and B.U_ItemCode = '" + oDBs_Detail.GetValue("ItemCode", j).Trim + "'")

                            If recsel.Fields.Item("U_Status").Value = "Closed" Then
                                chk = chk + 1
                                Dim ii As Integer = objMatrix.VisualRowCount - 1
                            End If
                            If chk = objMatrix.VisualRowCount - 1 Then
                                recupdt.DoQuery("Update [@PRE_SHIPMENT] Set U_Status = 'Closed',U_postno = '" + oDBs_Head.GetValue("DocNum", 0).Trim + "' Where DocNum = '" + Prenos + "'")
                            End If
                        Next
                    End If
            End Select
        Catch ex As Exception

        End Try
    End Sub
#Region "Rajkumar"
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
    Sub LoadFreight(ByVal FormUID As String)
        If loadcount <> 0 Then
            Exit Sub
        End If
        loadcount += 1
        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oForm As SAPbouiCOM.Form = oApplication.Forms.Item(FormUID)
        oForm.Items.Item("91").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
        Dim FrgtForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
        Dim FrgtMatrix As SAPbouiCOM.Matrix = FrgtForm.Items.Item("3").Specific
        For k As Integer = 1 To FrgtMatrix.VisualRowCount
            oRSet.DoQuery("Insert Into Freight_Order_Pre(RowNo,ExpnsCode,MacID) Values('" + k.ToString.Trim + "','" + Trim(FrgtMatrix.Columns.Item("1").Cells.Item(k).Specific.value) + "','" + MAC_ID + "')")
        Next
        Try
            'Dim FrgtMatrix As SAPbouiCOM.Matrix = FrgtForm.Items.Item("3").Specific
            FrgtForm.Freeze(True)
            For k As Integer = 1 To FrgtMatrix.VisualRowCount
                FrgtMatrix.Columns.Item("3").Cells.Item(k).Specific.Value = 0
            Next
            oRecordSet.DoQuery("Select u_fcode,u_fname,u_tax,u_amount,u_posfrgt,u_postax,u_negfrgt,u_negtax From [@GEN_ACCRUALS] Where u_invno = '" + InvNo + "' And u_macid = '" + MAC_ID + "'")
            For i As Integer = 1 To oRecordSet.RecordCount
                oRSet.DoQuery("Select Top 1 RowNo From Freight_Order_Pre Where ExpnsCode = '" + Trim(oRecordSet.Fields.Item("u_fcode").Value) + "' And Macid = '" + MAC_ID + "'")
                If oRSet.RecordCount > 0 Then
                    Dim RowNo As Integer = oRSet.Fields.Item("RowNo").Value
                    If CDbl(FrgtMatrix.Columns.Item("3").Cells.Item(RowNo).Specific.Value.ToString.Substring(3)) = 0 Then
                        FrgtMatrix.Columns.Item("3").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_amount").Value
                        FrgtMatrix.Columns.Item("17").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_tax").Value
                    End If
                End If
                oRSet.DoQuery("Select Top 1 RowNo From Freight_Order_Pre Where ExpnsCode = '" + Trim(oRecordSet.Fields.Item("u_posfrgt").Value) + "' And Macid = '" + MAC_ID + "'")
                If oRSet.RecordCount > 0 Then
                    Dim RowNo As Integer = oRSet.Fields.Item("RowNo").Value
                    If CDbl(FrgtMatrix.Columns.Item("3").Cells.Item(RowNo).Specific.Value.ToString.Substring(3)) = 0 Then
                        FrgtMatrix.Columns.Item("3").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_amount").Value
                        FrgtMatrix.Columns.Item("17").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_postax").Value
                    End If
                End If
                oRSet.DoQuery("Select Top 1 RowNo From Freight_Order_Pre Where ExpnsCode = '" + Trim(oRecordSet.Fields.Item("u_negfrgt").Value) + "' And Macid = '" + MAC_ID + "'")
                If oRSet.RecordCount > 0 Then
                    Dim RowNo As Integer = oRSet.Fields.Item("RowNo").Value
                    If CDbl(FrgtMatrix.Columns.Item("3").Cells.Item(RowNo).Specific.Value.ToString.Substring(3)) = 0 Then
                        FrgtMatrix.Columns.Item("3").Cells.Item(RowNo).Specific.Value = "-" & oRecordSet.Fields.Item("u_amount").Value
                        FrgtMatrix.Columns.Item("17").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_negtax").Value
                    End If
                End If
                oRecordSet.MoveNext()
            Next
            FrgtForm.Freeze(False)
            FrghtFlag = False
        Catch ex As Exception
            FrgtForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub LoadFreight_Pre(ByVal FormUID As String)
        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oForm As SAPbouiCOM.Form = oApplication.Forms.GetFormByTypeAndCount("133", 1)
        oForm.Items.Item("91").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
        Dim FrgtForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
        Dim FrgtMatrix As SAPbouiCOM.Matrix = FrgtForm.Items.Item("3").Specific
        For k As Integer = 1 To FrgtMatrix.VisualRowCount
            oRSet.DoQuery("Insert Into Freight_Order_Pre(RowNo,ExpnsCode,MacID) Values('" + k.ToString.Trim + "','" + Trim(FrgtMatrix.Columns.Item("1").Cells.Item(k).Specific.value) + "','" + MAC_ID + "')")
        Next
        Try
            'Dim FrgtMatrix As SAPbouiCOM.Matrix = FrgtForm.Items.Item("3").Specific
            FrgtForm.Freeze(True)
            For k As Integer = 1 To FrgtMatrix.VisualRowCount
                FrgtMatrix.Columns.Item("3").Cells.Item(k).Specific.Value = 0
            Next
            oRecordSet.DoQuery("Select u_ExpnCode,u_ExpnName,u_TaxCode,u_TotFrgn From [@PRE_SHIPMENT_D3] Where U_ExpnCode <> '' and DocEntry IN (" & SELDOC & ")") 'DocEnrty = '" + InvNo + "'") ' And u_macid = '" + MAC_ID + "'")
            For i As Integer = 1 To oRecordSet.RecordCount
                oRSet.DoQuery("Select Top 1 RowNo From Freight_Order_Pre Where ExpnsCode = '" + Trim(oRecordSet.Fields.Item("u_ExpnCode").Value) + "' And Macid = '" + MAC_ID + "'")
                If oRSet.RecordCount > 0 Then
                    Dim RowNo As Integer = oRSet.Fields.Item("RowNo").Value
                    If CDbl(FrgtMatrix.Columns.Item("3").Cells.Item(RowNo).Specific.Value.ToString.Substring(3)) = 0 Then
                        FrgtMatrix.Columns.Item("3").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_TotFrgn").Value
                        FrgtMatrix.Columns.Item("17").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_TaxCode").Value
                    End If
                End If
                oRSet.DoQuery("Select Top 1 RowNo From Freight_Order_Pre Where ExpnsCode = '" + Trim(oRecordSet.Fields.Item("u_ExpnCode").Value) + "' And Macid = '" + MAC_ID + "'")
                If oRSet.RecordCount > 0 Then
                    Dim RowNo As Integer = oRSet.Fields.Item("RowNo").Value
                    If CDbl(FrgtMatrix.Columns.Item("3").Cells.Item(RowNo).Specific.Value.ToString.Substring(3)) = 0 Then
                        FrgtMatrix.Columns.Item("3").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_TotFrgn").Value
                        FrgtMatrix.Columns.Item("17").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_TaxCode").Value
                    End If
                End If
                oRSet.DoQuery("Select Top 1 RowNo From Freight_Order_Pre Where ExpnsCode = '" + Trim(oRecordSet.Fields.Item("u_ExpnCode").Value) + "' And Macid = '" + MAC_ID + "'")
                If oRSet.RecordCount > 0 Then
                    Dim RowNo As Integer = oRSet.Fields.Item("RowNo").Value
                    If CDbl(FrgtMatrix.Columns.Item("3").Cells.Item(RowNo).Specific.Value.ToString.Substring(3)) = 0 Then
                        FrgtMatrix.Columns.Item("3").Cells.Item(RowNo).Specific.Value = "-" & oRecordSet.Fields.Item("u_TotFrgn").Value
                        FrgtMatrix.Columns.Item("17").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_TaxCode").Value
                    End If
                End If
                oRecordSet.MoveNext()
            Next
            FrgtForm.Freeze(False)
            FrghtFlag = False
        Catch ex As Exception
            FrgtForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub Open_Accruals_Form(ByVal FormUID As String, ByVal InvoiceNo As String, ByVal MACID As String, ByVal Mode As String, ByVal Route As String)
        Try
            PARENT_FORM = FormUID
            Dim RS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim CHILD_FORM As String = "ACCRUALS@" & FormUID
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
                oUtilities.SAPXML("Accruals.xml", CHILD_FORM)
                objSubForm = oApplication.Forms.Item(CHILD_FORM)
                objSubForm.Select()
            End If
            ChildModalForm = True
            Dim ogrid As SAPbouiCOM.Grid
            ogrid = objSubForm.Items.Item("grd").Specific
            objSubForm.DataSources.DataTables.Add("MyDataTable")
            If Mode = "A" Then
                'Dim oRs2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRs2.DoQuery("Delete From [@GEN_ACCRUALS] ")
                RS1.DoQuery("Select u_invno From [@GEN_ACCRUALS] Where u_invno = '" + InvoiceNo + "'")
                If RS1.RecordCount > 0 Then
                    objSubForm.DataSources.DataTables.Item(0).ExecuteQuery("Select ExpnsCode As 'Freight Code',ExpnsName As 'Freight Name', U_appltax As 'Tax Applicable',IsNull((Select u_amount From [@GEN_ACCRUALS] Where u_fcode = ExpnsCode And u_invno = '" + InvoiceNo + "' And u_macid = '" + MACID + "'),0) As 'Freight Amount' From OEXD Where IsNull(u_incl,'NO') = 'YES'")
                    ogrid.DataTable = objSubForm.DataSources.DataTables.Item("MyDataTable")
                Else
                    objSubForm.DataSources.DataTables.Item(0).ExecuteQuery("Exec UBG_Acrruals_Order_BaseNum '" + InvoiceNo + "','" + MACID + "'")
                    ogrid.DataTable = objSubForm.DataSources.DataTables.Item("MyDataTable")
                    If ANSP > 0 Then
                        ogrid.DataTable.Columns.Item("Freight Amount").Cells.Item(0).Value = "INR " + CStr(ANSP)
                    End If
                    If INS > 0 Then
                        ogrid.DataTable.Columns.Item("Freight Amount").Cells.Item(4).Value = "INR " + CStr(INS)
                    End If
                    If COMM > 0 Then
                        ogrid.DataTable.Columns.Item("Freight Amount").Cells.Item(2).Value = "USD " + CStr(COMM)
                    End If
                    If Dbk > 0 Then
                        ogrid.DataTable.Columns.Item("Freight Amount").Cells.Item(3).Value = "INR " + CStr(Dbk)
                    End If
                    If TRANS > 0 Then 'And Route = "Single" Then
                        ogrid.DataTable.Columns.Item("Freight Amount").Cells.Item(8).Value = "INR " + CStr(TRANS)
                    End If
                End If

                'If Fs_Source <> "" And Fs_PortOfLoading <> "" Then

                '    oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                '    oRecordSet.DoQuery("SELECT (T0.[U_AMOUNT]/T0.[U_NO_OF_CON]) as Amount FROM [dbo].[@GEN_FREIGHT]  T0 WHERE T0.[U_PR_OF_LOD]  = '" + Fs_PortOfLoading + "'" + " and T0.[U_SOURCE] = '" + Fs_Source + "'")
                '    If oRecordSet.RecordCount > 0 Then
                '        ogrid.DataTable.Columns.Item("Freight Amount").Cells.Item(2).Value = oRecordSet.Fields.Item("Amount").Value * Fd_NoOfContainer
                '    End If
                'End If

            End If
            If Mode = "O" Then
                RS1.DoQuery("Select B.ExpnsCode,B.LineTotal From OINV A Inner Join INV3 B On A.DocEntry = B.DocEntry Inner Join OEXD C On B.ExpnsCode = C.ExpnsCode And C.u_incl = 'YES' Where A.DocNum = '" + InvoiceNo + "'")
                If RS1.RecordCount > 0 Then
                    objSubForm.DataSources.DataTables.Item(0).ExecuteQuery("Exec UBG_Acrruals_Invoice '" + InvoiceNo + "'")
                    ogrid.DataTable = objSubForm.DataSources.DataTables.Item("MyDataTable")
                Else
                    objSubForm.DataSources.DataTables.Item(0).ExecuteQuery("Exec UBG_Acrruals_Order_BaseNum '" + InvoiceNo + "','" + MACID + "'")
                    ogrid.DataTable = objSubForm.DataSources.DataTables.Item("MyDataTable")
                End If
            End If
            RS1.DoQuery("Delete From OINV_BASENUM Where macid = '" + MACID + "'")
            ' ogrid.DataTable = objSubForm.DataSources.DataTables.Item("MyDataTable")
            For i As Integer = 0 To ogrid.Columns.Count - 1
                ogrid.Columns.Item(i).Editable = False
            Next
            ogrid.Columns.Item(ogrid.Columns.Count - 1).Editable = True
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
            oCon.Alias = "U_CustCode"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = Trim(objForm.Items.Item("4").Specific.Value)
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
    'Rajkumar
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
    Sub ItemEvent_Accrual_Form(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objSubForm = oApplication.Forms.Item(FormUID)
                    PARENT_FORM = (pVal.FormUID.Substring(objSubForm.UniqueID.IndexOf("@") + 1))
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If pVal.ItemUID = "grd" And pVal.ColUID = "Freight Amount" And pVal.BeforeAction = False Then
                        Dim LineTotalSum As Double
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("OINV")
                        If Trim(oDBs_Head.GetValue("U_CF", 0)) = "CIF" And pVal.Row = 1 Then
                            Dim Freight As String
                            Dim FrgtVal As Double
                            Dim grd As SAPbouiCOM.Grid = objSubForm.Items.Item("grd").Specific
                            Freight = grd.DataTable.Columns.Item("Freight Amount").Cells.Item(1).Value
                            'PRTVal = grd.DataTable.Columns.Item("Freight Amount").Cells.Item(
                            If Freight <> "" Then
                                Dim FreightCur As String = Freight.Substring(0, 3)
                                If FreightCur <> "INR" And FreightCur <> "inr" Then
                                    oRSet.DoQuery("Select Rate From ORTT Where Currency = '" + FreightCur + "' ANd RateDate = '" + Trim(objForm.Items.Item("10").Specific.value) + "'")
                                    FrgtVal = oRSet.Fields.Item("Rate").Value * CDbl(Freight.Substring(3))
                                Else
                                    FrgtVal = CDbl(Freight.Substring(3))
                                End If
                            End If
                            objMatrix = objForm.Items.Item("38").Specific
                            LineTotalSum = 0
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                If Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) <> "" And Trim(objMatrix.Columns.Item("14").Cells.Item(i).Specific.value) <> "" Then
                                    LineTotalSum = LineTotalSum + (CDbl(objForm.Items.Item("64").Specific.value) * CDbl(objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value.ToString.Substring(4)) * objMatrix.Columns.Item("11").Cells.Item(i).Specific.value)
                                End If
                            Next
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("OINV")
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@INV1")
                            Dbk = 0
                            ANSP = 0
                            COMM = 0
                            TRANS = 0

                            For i As Integer = 1 To objMatrix.VisualRowCount
                                If Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) <> "" And Trim(objMatrix.Columns.Item("14").Cells.Item(i).Specific.value) <> "" Then
                                    oRSet.DoQuery("Select Distinct IsNull(u_expval,0) 'u_expval',isnull(u_expper,0) 'u_expper' From OITM Where ItemCode = '" + Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) + "'")
                                    Dbk = Dbk + (oRSet.Fields.Item("u_expval").Value * objMatrix.Columns.Item("11").Cells.Item(i).Specific.value)
                                    oRSet.DoQuery("Select B.u_ansp,A.u_doccur From [@GEN_SUPP_PRICE] A INNER JOIN [@GEN_SUPP_PRICE_D0] B On A.Code = B.Code Where B.u_ItemCode = '" + Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) + "' And A.u_cardcode = '" + Trim(objForm.Items.Item("4").Specific.value) + "'")
                                    ANSP = ANSP + (objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value * oRSet.Fields.Item("u_ansp").Value)
                                    ANSPCur = oRSet.Fields.Item("u_doccur").Value
                                    COMM = COMM + (objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value * oRSet.Fields.Item("u_comm").Value)
                                    oRSet.DoQuery("SELECT T0.[U_AMOUNT] FROM [dbo].[@GEN_FREIGHT]  T0 WHERE T0.[U_PR_OF_LOD] = '" + Fs_PortOfLoading + "' and  T0.[U_SuppMan] = '" + Fs_PlaceOfSupplier + "'")
                                    TRANS = oRSet.Fields.Item("U_AMOUNT").Value * Fd_NoOfContainer
                                End If
                            Next

                            grd.DataTable.Columns.Item("Freight Amount").Cells.Item(0).Value = "INR" & CStr(CDbl(Dbk))
                        End If
                        If Trim(oDBs_Head.GetValue("U_CF", 0)) = "CIF" And pVal.Row = 6 Then
                            Dim Freight As String
                            Dim FrgtVal As Double
                            Dim grd As SAPbouiCOM.Grid = objSubForm.Items.Item("grd").Specific
                            Freight = grd.DataTable.Columns.Item("Freight Amount").Cells.Item(6).Value
                            'PRTVal = grd.DataTable.Columns.Item("Freight Amount").Cells.Item(
                            If Freight <> "" Then
                                Dim FreightCur As String = Freight.Substring(0, 3)
                                If FreightCur <> "INR" And FreightCur <> "inr" Then
                                    oRSet.DoQuery("Select Rate From ORTT Where Currency = '" + FreightCur + "' ANd RateDate = '" + Trim(objForm.Items.Item("10").Specific.value) + "'")
                                    FrgtVal = oRSet.Fields.Item("Rate").Value * CDbl(Freight.Substring(3))
                                Else
                                    FrgtVal = CDbl(Freight.Substring(3))
                                End If
                            End If
                            objMatrix = objForm.Items.Item("38").Specific
                            LineTotalSum = 0
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                If Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) <> "" And Trim(objMatrix.Columns.Item("14").Cells.Item(i).Specific.value) <> "" Then
                                    LineTotalSum = LineTotalSum + (CDbl(objForm.Items.Item("64").Specific.value) * CDbl(objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value.ToString.Substring(4)) * objMatrix.Columns.Item("11").Cells.Item(i).Specific.value)
                                End If
                            Next
                            oRSet.DoQuery("Select IsNull(PrintHeadr,0) 'Per1',IsNull(Manager,0) 'Per2' From OADM")
                            INS = (LineTotalSum + FrgtVal) * CDbl(oRSet.Fields.Item("Per1").Value) * CDbl(oRSet.Fields.Item("Per2").Value) / 100
                            grd.DataTable.Columns.Item("Freight Amount").Cells.Item(7).Value = "INR" & CStr(INS)
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "bkac" And pVal.BeforeAction = False Then
                        loadcount = 0
                        Dim grd As SAPbouiCOM.Grid = objSubForm.Items.Item("grd").Specific
                        Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' Dim Code As String = RS.Fields.Item("Code").Value
                        oRSet.DoQuery("Delete From [@GEN_ACCRUALS] Where u_invno = '" + InvNo + "' And u_macid = '" + MAC_ID + "'")
                        For i As Integer = 0 To grd.Rows.Count - 1
                            If grd.DataTable.Columns.Item("Freight Amount").Cells.Item(i).Value <> "0" And grd.DataTable.Columns.Item("Freight Amount").Cells.Item(i).Value <> "" Then
                                RS.DoQuery("Select Convert(VarChar,Count(*) + 1) AS 'Code' From [@GEN_ACCRUALS]")
                                Dim posfcode, postax, negfcode, negtax As String
                                Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("Select IsNull(u_pfreight,'') As 'pfreight' From OEXD Where ExpnsCode = '" + Trim(grd.DataTable.Columns.Item("Freight Code").Cells.Item(i).Value) + "'")
                                posfcode = oRecordSet.Fields.Item("pfreight").Value
                                oRecordSet.DoQuery("Select ExpnsCode From OEXD Where ExpnsName = '" + posfcode + "'")
                                posfcode = oRecordSet.Fields.Item("ExpnsCode").Value
                                oRecordSet.DoQuery("Select IsNull(U_appltax,'') AS 'ptax' From OEXD Where ExpnsCode = '" + posfcode + "'")
                                postax = oRecordSet.Fields.Item("ptax").Value
                                oRecordSet.DoQuery("Select IsNull(u_nfreight,'') As 'nfreight' From OEXD Where ExpnsCode = '" + Trim(grd.DataTable.Columns.Item("Freight Code").Cells.Item(i).Value) + "'")
                                negfcode = oRecordSet.Fields.Item("nfreight").Value
                                oRecordSet.DoQuery("Select ExpnsCode From OEXD Where ExpnsName = '" + negfcode + "'")
                                negfcode = oRecordSet.Fields.Item("ExpnsCode").Value
                                oRecordSet.DoQuery("Select IsNull(U_appltax,'') AS 'ntax' From OEXD Where ExpnsCode = '" + negfcode + "'")
                                negtax = oRecordSet.Fields.Item("ntax").Value
                                oRSet.DoQuery("Insert Into [@GEN_ACCRUALS] (Code,Name,u_invno,u_macid,u_fcode,u_fname,u_tax,u_amount,u_posfrgt,u_postax,u_negfrgt,u_negtax) Values('" + Trim(RS.Fields.Item("Code").Value) + "','" + Trim(RS.Fields.Item("Code").Value) + "','" + InvNo + "','" + MAC_ID + "','" + grd.DataTable.Columns.Item("Freight Code").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("Freight Name").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("Tax Applicable").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("Freight Amount").Cells.Item(i).Value.ToString.Trim + "','" + posfcode + "','" + postax + "','" + negfcode + "','" + negtax + "') ")
                                oRSet.DoQuery("Delete From Freight_Order_Pre Where MacID = '" + MAC_ID + "'")
                                'oRSet.DoQuery("Insert Into Freight_Order (RowNo,ExpnsCode,MacID) SELECT ROW_NUMBER() OVER (ORDER BY ExpnsName) AS Row, ExpnsCode,'" + MAC_ID + "' FROM OEXD")
                            End If
                        Next
                        FrghtFlag = True
                        objSubForm.Close()
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub LoadItems(ByVal FormUID As String, ByVal preNo As String, ByVal ROW As Integer)
        Try
            objForm = oApplication.Forms.GetFormByTypeAndCount("133", 1)
            'objform1 = oApplication.Forms.GetFormByTypeAndCount("-133", 1)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("OINV")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("INV1")
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select B.U_ItemCode,B.U_ItemName,B.U_Quantity,B.u_UOM,B.U_Price,B.U_TaxCode,A.DocNum,B.LineId,B.U_TotalLC,B.U_Whse,A.DocEntry,A.U_CustRef,A.U_DocCur,A.U_DelDate,A.U_Buyer,A.U_Owner,B.U_SONo,B.U_BaseLine From [@PRE_SHIPMENT] A Inner Join [@PRE_SHIPMENT_D0] B On A.DocEntry = B.DocEntry  And IsNull(A.U_Status,'Open') = 'Open'  And A.DocNum IN (" + preNo + ") and B.U_ItemCode <> '' and B.U_Quantity > 0")
            Try
                objForm.Freeze(True)
                objForm.Items.Item("38").Enabled = True
                DeleteFlag = False
                Dim norecord As Integer = 0
                For i As Integer = 1 To objMatrix.VisualRowCount - 2

                    If oRecordSet.RecordCount > norecord Then
                        If objMatrix.Columns.Item("44").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("U_SONo").Value And objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("U_ItemCode").Value Then 'And objMatrix.Columns.Item("46").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("U_BaseLine").Value Then
                            objMatrix.Columns.Item("U_preno").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("DocNum").Value
                            If oRecordSet.Fields.Item("U_Quantity").Value > 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value = (oRecordSet.Fields.Item("U_Quantity").Value) '- oRecordSet.Fields.Item("U_postqty").Value
                                objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("U_Price").Value
                            End If
                            oRecordSet.MoveNext()
                            DeleteFlag = True
                            norecord = norecord + 1
                        End If

                    End If
                    'If DeleteFlag = False Then
                    '    'objMatrix.DeleteRow(i)
                    '    'i = i - 1
                    '    objMatrix.ClearRowData(i)
                    'End If
                    DeleteFlag = False
                Next
                oApplication.StatusBar.SetText("Updated Complete")
                'objForm.Items.Item("16").Specific.Value = "Based On PreShipment No." & "" & DOCNUM
                objForm.Freeze(False)
            Catch ex As Exception
                objForm.Freeze(False)
                oApplication.StatusBar.SetText(ex.Message)
            End Try

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
    Sub invoice_ref(ByRef FormUID As String, ByRef Series As String)
        objForm = oApplication.Forms.Item(FormUID)
        Dim recupdt As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim recsel As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oDBs_Head = objForm.DataSources.DBDataSources.Item("OINV")
        oDBs_Detail = objForm.DataSources.DBDataSources.Item("INV1")
        objMatrix = objForm.Items.Item("38").Specific
        Dim chk As Integer = 0
        Dim Invoice As SAPbobsCOM.Documents
        Dim Invoice_Lines As SAPbobsCOM.Document_Lines
        Dim PC As String
        Dim DocNum As String
        Dim UserSign As String
        Invoice = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        Invoice_Lines = Invoice.Lines
        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oRSets As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim DocEntry As String
        If Trim(objForm.Items.Item("3").Specific.Selected.Value) = "I" Then
            oRSet.DoQuery("Select UserID From OUSR Where User_Code = '" + oCompany.UserName.ToString.Trim + "'")
            UserSign = oRSet.Fields.Item("UserID").Value
            oRSet.DoQuery("Select Max(DocEntry) As 'DocEntry' From OINV Where UserSign = '" + UserSign + "' and series='" + Series + "'")
            DocNum = oRSet.Fields.Item("DocEntry").Value
            oRSet.DoQuery("Select DocEntry,u_unit,DocCur,DocRate From OINV Where DocEntry = '" + DocNum + "'")
            DocEntry = oRSet.Fields.Item("DocEntry").Value
            PC = oRSet.Fields.Item("u_unit").Value
            oRecordSet.DoQuery("Select ItemCode,isnull(u_preship,0),LineNum from OINV A inner join INV1 B on A.DocEntry = B.DocEntry Where A.DocEntry = '" + DocNum + "'")
            Dim PreNum As String = oRecordSet.Fields.Item(1).Value
            Dim str_qry As String = "Select B.U_ItemCode,A.DocNum,B.U_BaseLine From [@PRE_SHIPMENT] A Inner Join [@PRE_SHIPMENT_D0] B On A.DocEntry = B.DocEntry Where A.DocNum = '" & PreNum & "' and B.U_ItemCode <> ''"
            oRSets.DoQuery(str_qry)
            'Rajkumar
            If oRSet.RecordCount = oRecordSet.RecordCount Then
                Invoice.UserFields.Fields.Item("U_oinvno").Value = DocNum
                Invoice.UserFields.Fields.Item("U_unit").Value = PC
                Invoice.Lines.ItemCode = oRecordSet.Fields.Item("ItemCode").Value.ToString.Trim
                Invoice.Lines.BaseLine = oRSets.Fields.Item("U_BaseLine").Value
                Invoice.Lines.BaseType = "17"
                Invoice.Lines.Add()
                Invoice.Lines.SetCurrentLine(1)
                Dim res As Integer = Invoice.Update()
                If res <> 0 Then
                    If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription)
                    Exit Sub
                Else
                    If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    oRSet.DoQuery("Select DocNum From OINV Where DocEntry = '" + oCompany.GetNewObjectKey.ToString + "'")
                    Dim strprint As String
                    strprint = "Document Nos : " & objForm.Items.Item("8").Specific.value & " & " & oRSet.Fields.Item("DocNum").Value & " are created"
                    oApplication.MessageBox(strprint)
                End If
            End If
            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSet.DoQuery("Select UserId From OUSR Where User_Code = '" + oCompany.UserName.Trim + "'")
            oRS.DoQuery("Select Max(DocEntry) AS 'DocEntry' From OINV Where UserSign = '" + Trim(oRSet.Fields.Item("UserID").Value) + "'")
            oRSet.DoQuery("Select DocNum From  OINV Where DocEntry = '" + Trim(oRS.Fields.Item("DocEntry").Value) + "'")


        End If
    End Sub
#End Region

End Class
