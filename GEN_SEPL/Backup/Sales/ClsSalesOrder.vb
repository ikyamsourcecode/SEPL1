Public Class ClsSalesOrder

#Region "        Declaration        "
    Dim Count As Double = 0
    Dim oUtilities As New ClsUtilities
    Dim objForm, objSubForm, objSForm, objPreForm As SAPbouiCOM.Form
    Dim objItem, objOldItem, objItem1 As SAPbouiCOM.Item
    Dim objMatrix, objSubMatrix, objBatchMatrix As SAPbouiCOM.Matrix
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim SZDBHead As SAPbouiCOM.DBDataSource
    Dim SZDBDetail As SAPbouiCOM.DBDataSource
    Dim SMDBHead As SAPbouiCOM.DBDataSource
    Dim SMDBDetail As SAPbouiCOM.DBDataSource
    Dim oDBs_Head, oDBs_Detail, oDBs_Pre_Head, oDBs_Pre_Detail As SAPbouiCOM.DBDataSource
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
    Dim DeleteItemCode As String
    Dim InvNo As String
    Dim FrghtFlag As Boolean = False
    Dim DbkVal, DbkPer, Dbk, ANSP, COMM, INS, LineTotalSum As Double
    Dim RowID As Integer
    Dim SONO As String
    Dim BASENUM As String
    Dim FormMode As String
    Dim COM As Double
    Dim ANSPCur As String
    Dim DOCNUM As Integer
    Dim loadcount As Integer = 0
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            objOldItem = objForm.Items.Item("86")
            objItem = objForm.Items.Add("spc", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = objOldItem.Top + objOldItem.Height + 20
            objItem.Left = objOldItem.Left
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.Caption = "Unit"
            objItem.LinkTo = "86"
            objOldItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("cpc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = objOldItem.Top + objOldItem.Height + 20
            objItem.Left = objOldItem.Left
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.databind.setbound(True, "ORDR", "u_unit")
            objItem.Specific.TabOrder = objOldItem.Specific.TabOrder + 1
            objItem.LinkTo = "46"
            objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objOldItem = objForm.Items.Item("2")
            objItem = objForm.Items.Add("btnsize", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width + 20
            objItem.Left = objOldItem.Left + objOldItem.Width + 5
            objItem.Top = objOldItem.Top
            objItem.Specific.caption = "Allocate Size"
            objItem.LinkTo = "2"
            objOldItem = objForm.Items.Item("86")
            objItem = objForm.Items.Add("sseason", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Specific.Caption = "Season"
            objItem.LinkTo = "86"
            objOldItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("tseason", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Specific.databind.setbound(True, "ORDR", "u_season")
            objItem.LinkTo = "46"
            objItem.Specific.TabOrder = objOldItem.Specific.TabOrder + 1
            objItem.DisplayDesc = True
            ObjOldItem = objForm.Items.Item("15")
            objItem = objForm.Items.Add("sdoccur", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = ObjOldItem.Top + ObjOldItem.Height + 15
            objItem.Left = ObjOldItem.Left
            objItem.Width = ObjOldItem.Width
            objItem.Height = ObjOldItem.Height
            objItem.Specific.Caption = "Buyer Currency"
            objItem.LinkTo = "15"
            ObjOldItem = objForm.Items.Item("15")
            objItem = objForm.Items.Add("sdocrate", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = ObjOldItem.Top + ObjOldItem.Height + 30
            objItem.Left = ObjOldItem.Left
            objItem.Width = ObjOldItem.Width
            objItem.Height = ObjOldItem.Height
            objItem.Specific.Caption = "Document Rate"
            objItem.LinkTo = "15"
            ObjOldItem = objForm.Items.Item("14")
            objItem = objForm.Items.Add("doccur", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = ObjOldItem.Top + ObjOldItem.Height + 15
            objItem.Left = ObjOldItem.Left
            objItem.Width = ObjOldItem.Width
            objItem.Height = ObjOldItem.Height
            objItem.Specific.databind.setbound(True, "ORDR", "u_doccur")
            objItem.Specific.TabOrder = ObjOldItem.Specific.TabOrder + 1
            objItem.LinkTo = "14"
            ObjOldItem = objForm.Items.Item("14")
            objItem = objForm.Items.Add("docrate", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = ObjOldItem.Top + ObjOldItem.Height + 30
            objItem.Left = ObjOldItem.Left
            objItem.Width = ObjOldItem.Width
            objItem.Height = ObjOldItem.Height
            objItem.Specific.databind.setbound(True, "ORDR", "u_docrate")
            objItem.Specific.TabOrder = ObjOldItem.Specific.TabOrder + 1
            objItem.LinkTo = "14"
            objForm.Items.Item("doccur").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("docrate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objOldItem = objForm.Items.Item("btnsize")
            objItem = objForm.Items.Add("btn", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Top = ObjOldItem.Top
            objItem.Left = ObjOldItem.Left + ObjOldItem.Width + 10
            objItem.Width = objOldItem.Width + 30
            objItem.Height = objOldItem.Height
            objItem.Specific.caption = "Generate FC value"

            'Export 
            objForm = oApplication.Forms.Item(FormUID)
            objItem = objForm.Items.Item("btn")
            objItem1 = objForm.Items.Add("btnac", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem1.Top = objItem.Top
            objItem1.Left = objItem.Left + objItem.Width + 5
            objItem1.Height = objItem.Height
            objItem1.Width = objItem.Width + 50
            objItem1.Specific.caption = "Accruals & Expenses"
            objItem1.LinkTo = "2"
            objItem1 = objForm.Items.Item("btnac")

            'objOldItem = objForm.Items.Item("btnac")
            'objItem = objForm.Items.Add("btntq", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'objItem.Top = objOldItem.Top
            'objItem.Left = objOldItem.Left + objOldItem.Width + 5
            'objItem.Height = objOldItem.Height
            'objItem.Width = objOldItem.Width + 50
            'objItem.Specific.caption = "ShowTotalQuantity"
            'objItem.LinkTo = "2"
            ' objOldItem = objForm.Items.Item("btntq")


            objOldItem = objForm.Items.Item("21")
            objItem = objForm.Items.Add("TQty", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = objForm.Items.Item("30").Top
            objItem.Left = objOldItem.Left
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.Caption = "Total Quantity"
            objItem.LinkTo = "21"
            objOldItem = objForm.Items.Item("222")
            objItem = objForm.Items.Add("tqty", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = objForm.Items.Item("29").Top
            objItem.Left = objOldItem.Left
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.databind.setbound(True, "ORDR", "U_season")
            objItem.Specific.TabOrder = objOldItem.Specific.TabOrder + 1
            objItem.LinkTo = "222"


            objOldItem = objForm.Items.Item("10000329")
            objItem = objForm.Items.Add("copyto", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Width = objOldItem.Width
            objItem.Top = objOldItem.Top
            objItem.Left = objOldItem.Left
            objItem.Height = objOldItem.Height
            objItem.Specific.caption = "Copy To Pre-Ship"
            objItem.LinkTo = "10000329"
            objOldItem.Visible = False
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
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
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("ORDR")
                        ''    If Trim(oDBs_Head.GetValue("U_CF", 0)) = "CIF" And pVal.Row = 1 Then
                        ''        Dim Freight As String
                        ''        Dim FrgtVal As Double
                        ''        Dim grd As SAPbouiCOM.Grid = objSubForm.Items.Item("grd").Specific
                        ''        Freight = grd.DataTable.Columns.Item("Freight Amount").Cells.Item(1).Value
                        ''        'PRTVal = grd.DataTable.Columns.Item("Freight Amount").Cells.Item(
                        ''        If Freight <> "" Then
                        ''            Dim FreightCur As String = Freight.Substring(0, 3)
                        ''            If FreightCur <> "INR" And FreightCur <> "inr" Then
                        ''                oRSet.DoQuery("Select Rate From ORTT Where Currency = '" + FreightCur + "' ANd RateDate = '" + Trim(objForm.Items.Item("10").Specific.value) + "'")
                        ''                FrgtVal = oRSet.Fields.Item("Rate").Value * CDbl(Freight.Substring(3))
                        ''            Else
                        ''                FrgtVal = CDbl(Freight.Substring(3))
                        ''            End If
                        ''        End If
                        ''        objMatrix = objForm.Items.Item("38").Specific
                        ''        LineTotalSum = 0
                        ''        For i As Integer = 1 To objMatrix.VisualRowCount
                        ''            If Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) <> "" And Trim(objMatrix.Columns.Item("14").Cells.Item(i).Specific.value) <> "" Then
                        ''                LineTotalSum = LineTotalSum + (CDbl(objForm.Items.Item("64").Specific.value) * CDbl(objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value.ToString.Substring(4)) * objMatrix.Columns.Item("11").Cells.Item(i).Specific.value)
                        ''            End If
                        ''        Next
                        ''        oDBs_Head = objForm.DataSources.DBDataSources.Item("ORDR")
                        ''        Dbk = 0
                        ''        ANSP = 0
                        ''        COMM = 0
                        ''        If Trim(objForm.Items.Item("cpc").Specific.Value) = "LG-UNIT1" Then
                        ''            For i As Integer = 1 To objMatrix.VisualRowCount
                        ''                If Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) <> "" And Trim(objMatrix.Columns.Item("14").Cells.Item(i).Specific.value) <> "" Then
                        ''                    oRSet.DoQuery("Select Distinct IsNull(u_expval,0) 'u_expval',isnull(u_expper,0) 'u_expper' From OITM Where ItemCode = '" + Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) + "'")
                        ''                    Dbk = Dbk + (oRSet.Fields.Item("u_expval").Value * objMatrix.Columns.Item("11").Cells.Item(i).Specific.value)
                        ''                    'oRSet.DoQuery("Select B.u_ansp,A.u_doccur,B.u_comm From [@GEN_SUPP_PRICE] A INNER JOIN [@GEN_SUPP_PRICE_D0] B On A.Code = B.Code Where B.u_ItemCode = '" + Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) + "' And A.u_cardcode = '" + Trim(objForm.Items.Item("4").Specific.value) + "'")
                        ''                    'ANSP = ANSP + (objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value * oRSet.Fields.Item("u_ansp").Value)
                        ''                    'ANSPCur = oRSet.Fields.Item("u_doccur").Value
                        ''                    'COMM = COMM + (objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value * oRSet.Fields.Item("u_comm").Value)
                        ''                End If
                        ''            Next
                        ''        End If
                        ''        grd.DataTable.Columns.Item("Freight Amount").Cells.Item(0).Value = "INR" & CStr(CDbl(Dbk))
                        ''    End If
                        ''    If Trim(oDBs_Head.GetValue("U_CF", 0)) = "CIF" And pVal.Row = 6 Then
                        ''        Dim Freight As String
                        ''        Dim FrgtVal As Double
                        ''        Dim grd As SAPbouiCOM.Grid = objSubForm.Items.Item("grd").Specific
                        ''        Freight = grd.DataTable.Columns.Item("Freight Amount").Cells.Item(6).Value
                        ''        'PRTVal = grd.DataTable.Columns.Item("Freight Amount").Cells.Item(
                        ''        If Freight <> "" Then
                        ''            Dim FreightCur As String = Freight.Substring(0, 3)
                        ''            If FreightCur <> "INR" And FreightCur <> "inr" Then
                        ''                oRSet.DoQuery("Select Rate From ORTT Where Currency = '" + FreightCur + "' ANd RateDate = '" + Trim(objForm.Items.Item("10").Specific.value) + "'")
                        ''                FrgtVal = oRSet.Fields.Item("Rate").Value * CDbl(Freight.Substring(3))
                        ''            Else
                        ''                FrgtVal = CDbl(Freight.Substring(3))
                        ''            End If
                        ''        End If
                        ''        objMatrix = objForm.Items.Item("38").Specific
                        ''        LineTotalSum = 0
                        ''        For i As Integer = 1 To objMatrix.VisualRowCount
                        ''            If Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) <> "" And Trim(objMatrix.Columns.Item("14").Cells.Item(i).Specific.value) <> "" Then
                        ''                LineTotalSum = LineTotalSum + (CDbl(objForm.Items.Item("64").Specific.value) * CDbl(objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value.ToString.Substring(4)) * objMatrix.Columns.Item("11").Cells.Item(i).Specific.value)
                        ''            End If
                        ''        Next
                        ''        oRSet.DoQuery("Select IsNull(PrintHeadr,0) 'Per1',IsNull(Manager,0) 'Per2' From OADM")
                        ''        INS = (LineTotalSum + FrgtVal) * CDbl(oRSet.Fields.Item("Per1").Value) * CDbl(oRSet.Fields.Item("Per2").Value) / 100
                        ''        grd.DataTable.Columns.Item("Freight Amount").Cells.Item(7).Value = "INR" & CStr(INS)
                        ''    End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "bkac" And pVal.BeforeAction = False Then
                        Dim grd As SAPbouiCOM.Grid = objSubForm.Items.Item("grd").Specific
                        Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' Dim Code As String = RS.Fields.Item("Code").Value
                        oRSet.DoQuery("Delete From [@GEN_ACCRUALS_RDR] Where u_invno = '" + InvNo + "' And u_macid = '" + MAC_ID + "'")
                        For i As Integer = 0 To grd.Rows.Count - 1
                            If grd.DataTable.Columns.Item("Freight Amount").Cells.Item(i).Value <> "0" And grd.DataTable.Columns.Item("Freight Amount").Cells.Item(i).Value <> "" Then
                                RS.DoQuery("Select Convert(VarChar,Count(*) + 1) AS 'Code' From [@GEN_ACCRUALS_RDR]")
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
                                oRSet.DoQuery("Insert Into [@GEN_ACCRUALS_RDR] (Code,Name,u_invno,u_macid,u_fcode,u_fname,u_tax,u_amount,u_posfrgt,u_postax,u_negfrgt,u_negtax) Values('" + Trim(RS.Fields.Item("Code").Value) + "','" + Trim(RS.Fields.Item("Code").Value) + "','" + InvNo + "','" + MAC_ID + "','" + grd.DataTable.Columns.Item("Freight Code").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("Freight Name").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("Tax Applicable").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("Freight Amount").Cells.Item(i).Value.ToString.Trim + "','" + posfcode + "','" + postax + "','" + negfcode + "','" + negtax + "') ")
                                oRSet.DoQuery("Delete From Freight_Order_Rdr Where MacID = '" + MAC_ID + "'")
                                'oRSet.DoQuery("Insert Into Freight_Order_Rdr (RowNo,ExpnsCode,MacID) SELECT ROW_NUMBER() OVER (ORDER BY ExpnsName) AS Row, ExpnsCode,'" + MAC_ID + "' FROM OEXD")
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

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Dim CHILD_FORM As String = "SOORD@" & pVal.FormUID
            Dim ModalForm As Boolean = False
            For i As Integer = 0 To oApplication.Forms.Count - 1
                If oApplication.Forms.Item(i).UniqueID = (CHILD_FORM) Then
                    objSubForm = oApplication.Forms.Item("SOORD@" & pVal.FormUID)
                    ModalForm = True
                    Exit For
                End If
            Next
            If ModalForm = False Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                        objForm = oApplication.Forms.Item(pVal.FormUID)
                        If FrghtFlag = True Then
                            LoadFreight(FormUID)
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        If pVal.BeforeAction = True Then
                            If pVal.FormTypeCount = 1 Then
                                Me.CreateForm(FormUID)
                            Else
                                BubbleEvent = False
                            End If
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
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        If (pVal.ItemUID = "10" Or pVal.ItemUID = "doccur") And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = False Then
                            Try
                                If Trim(objForm.Items.Item("doccur").Specific.value) <> "" Then
                                    Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRSet.DoQuery("Select Rate From ORTT Where RateDate = '" + Trim(objForm.Items.Item("10").Specific.value) + "' And Currency = '" + Trim(objForm.Items.Item("doccur").Specific.value) + "'")
                                    objForm.Items.Item("docrate").Specific.value = oRSet.Fields.Item("Rate").Value
                                Else
                                    objForm.Items.Item("docrate").Specific.value = ""
                                End If
                            Catch ex As Exception
                                oApplication.StatusBar.SetText(ex.Message)
                            End Try
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
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        If pVal.ItemUID = "38" And pVal.BeforeAction = True Then
                            objMatrix = objForm.Items.Item("38").Specific
                            If pVal.Row > 0 And pVal.Row <= objMatrix.VisualRowCount Then
                                RowID = pVal.Row
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                        If pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            RS.DoQuery("Delete From [@GEN_SZ_ORDR] Where u_sono = '" + Trim(objForm.Items.Item("8").Specific.value) + "' And u_macid = '" + MAC_ID + "'")
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS
                        If pVal.ItemUID = "tqty" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = False Then
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("ORDR")
                            Count = 0
                            ' oDBs_Detail = objForm.DataSources.DBDataSources.Item("INV1")
                            Dim objMatrix As SAPbouiCOM.Matrix
                            objMatrix = objForm.Items.Item("38").Specific

                            For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                                Count = Count + Convert.ToDouble(objMatrix.Columns.Item("11").Cells.Item(Row).Specific.value)

                            Next
                            Dim a As String = oDBs_Head.GetValue("U_season", 0)
                            objForm.Items.Item("tqty").Specific.value = Count.ToString()
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                            SONO = objForm.Items.Item("8").Specific.Value
                        End If
                        If pVal.ItemUID = "btntq" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("ORDR")
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
                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = True Then
                            If Trim(objForm.Items.Item("cpc").Specific.value) = "" Then
                                oApplication.StatusBar.SetText("Please select Unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                        If pVal.ItemUID = "btn" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            objMatrix = objForm.Items.Item("38").Specific
                            If pVal.BeforeAction = True Then
                                If Trim(objForm.Items.Item("doccur").Specific.value) = "" Or Trim(objForm.Items.Item("docrate").Specific.value) = 0 Then
                                    oApplication.StatusBar.SetText("Please select appropriate buyer currency and rate", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                End If
                            Else
                                For i As Integer = 1 To objMatrix.VisualRowCount
                                    If Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) <> "" And Trim(objMatrix.Columns.Item("14").Cells.Item(i).Specific.value) <> "" Then
                                        objMatrix.Columns.Item("U_pricefc").Cells.Item(i).Specific.value = CDbl(objMatrix.Columns.Item("14").Cells.Item(i).Specific.value.ToString.Substring(3)) / objForm.Items.Item("docrate").Specific.value
                                        objMatrix.Columns.Item("U_totalfc").Cells.Item(i).Specific.value = (CDbl(objMatrix.Columns.Item("14").Cells.Item(i).Specific.value.ToString.Substring(3)) / objForm.Items.Item("docrate").Specific.value) * objMatrix.Columns.Item("11").Cells.Item(i).Specific.value
                                    End If
                                Next
                            End If
                        End If
                        If pVal.ItemUID = "btnac" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.BeforeAction = False Then
                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            InvNo = objForm.Items.Item("8").Specific.Value
                            objMatrix = objForm.Items.Item("38").Specific
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("ORDR")
                            Dbk = 0
                            ANSP = 0
                            COMM = 0
                            LineTotalSum = 0
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'If Trim(objForm.Items.Item("cpc").Specific.Value) = "LG-UNIT1" Or Trim(objForm.Items.Item("cpc").Specific.Value) = "UNIT1" Then
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                If Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) <> "" And Trim(objMatrix.Columns.Item("14").Cells.Item(i).Specific.value) <> "" Then
                                    oRSet.DoQuery("Select Distinct IsNull(u_expval,0) 'u_expval',isnull(u_expper,0) 'u_expper' From OITM Where ItemCode = '" + Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) + "'")
                                    Dbk = Dbk + (oRSet.Fields.Item("u_expval").Value * objMatrix.Columns.Item("11").Cells.Item(i).Specific.value)
                                    oRSet.DoQuery("Select B.u_ansp,A.u_doccur,B.u_comm From [@GEN_SUPP_PRICE] A INNER JOIN [@GEN_SUPP_PRICE_D0] B On A.Code = B.Code Where B.u_ItemCode = '" + Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) + "' And A.u_cardcode = '" + Trim(objForm.Items.Item("4").Specific.value) + "'")
                                    ANSP = ANSP + (objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value * oRSet.Fields.Item("u_ansp").Value)
                                    ANSPCur = oRSet.Fields.Item("u_doccur").Value
                                    COMM = COMM + (objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value * oRSet.Fields.Item("u_comm").Value)
                                End If
                            Next
                            'End If
                            objMatrix = objForm.Items.Item("38").Specific
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                If Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) <> "" And Trim(objMatrix.Columns.Item("14").Cells.Item(i).Specific.value) <> "" Then
                                    LineTotalSum = LineTotalSum + (CDbl(objForm.Items.Item("64").Specific.value) * CDbl(objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value.ToString.Substring(4)) * objMatrix.Columns.Item("11").Cells.Item(i).Specific.value)
                                End If
                            Next
                            Dim h As String = objForm.Items.Item("64").Specific.value
                            oRSet.DoQuery("Select IsNull(PrintHeadr,0) 'Per1',IsNull(Manager,0) 'Per2' From OADM")
                            INS = (LineTotalSum * 0) / 100
                            'INS = (LineTotalSum * CDbl(oRSet.Fields.Item("Per1").Value) * CDbl(oRSet.Fields.Item("Per2").Value)) / 100
                            oRSet.DoQuery("Select IsNull(u_comper,0) AS 'COMPER' From OCRD Where CardCode = '" + Trim(objForm.Items.Item("4").Specific.value) + "'")
                            If CDbl(oRSet.Fields.Item("COMPER").Value) > 0 Then
                                COM = LineTotalSum * CDbl(oRSet.Fields.Item("COMPER").Value) / 100
                            End If
                            Dim TMPSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            TMPSet.DoQuery("Delete From ORDR_BASENUM WHere macid = '" + MAC_ID + "'")
                            For i As Integer = 1 To objMatrix.VisualRowCount - 1
                                TMPSet.DoQuery("Insert Into ORDR_BASENUM(invno,basenum,macid) Values('" + InvNo + "','" + Trim(objMatrix.Columns.Item("44").Cells.Item(i).Specific.value) + "','" + MAC_ID + "')")
                                'Dim h As String = Trim(objMatrix.Columns.Item("44").Cells.Item(i).Specific.value)
                            Next
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                FormMode = "A"
                            End If
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                FormMode = "O"
                            End If
                            Me.Open_Accruals_Form(pVal.FormUID, InvNo, MAC_ID, FormMode, BASENUM)
                        ElseIf pVal.ItemUID = "copyto" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.BeforeAction = False Then
                            DOCNUM = objForm.Items.Item("8").Specific.value
                            oApplication.ActivateMenuItem("PRE_SHIPMENT")
                            objPreForm = oApplication.Forms.ActiveForm
                            LoadItems(FormUID)
                        End If
                        If pVal.ItemUID = "btnsize" Then
                            If pVal.BeforeAction = True Then
                                objMatrix = objForm.Items.Item("38").Specific
                                If objMatrix.VisualRowCount < 1 Then
                                    oApplication.StatusBar.SetText("Please select items", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                Else
                                    Mode = pVal.FormMode
                                    Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                                    Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim RS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    RS1.DoQuery("Delete From TMP_ORDR_ITEMS Where ordrno = '" + Trim(objForm.Items.Item("8").Specific.Value) + "' And macid = '" + MAC_ID + "'")
                                    'RS.DoQuery("Select DocEntry From [@GEN_SIZE_ORDR] Where u_ordrno = '" + Trim(objForm.Items.Item("8").Specific.value) + "' And u_macid = '" + MAC_ID + "'")
                                    ''RS1.DoQuery("Delete From [@GEN_SIZE_ORDR_D0] Where DocEntry = '" + Trim(RS.Fields.Item("DocEntry").Value) + "'")
                                    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                                        If Trim(objMatrix.Columns.Item("U_asrtcode").Cells.Item(i).Specific.value) = "" Then
                                            oApplication.StatusBar.SetText("Please select Assorted code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    Next
                                    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                                        oRecordSet.DoQuery("Insert Into TMP_ORDR_ITEMS(ordrno,itemcode,asrtcode,qty,macid) Values('" + Trim(objForm.Items.Item("8").Specific.value) + "','" + Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) + "','" + Trim(objMatrix.Columns.Item("U_asrtcode").Cells.Item(i).Specific.value) + "','" + Trim(objMatrix.Columns.Item("11").Cells.Item(i).Specific.value) + "','" + MAC_ID + "')")
                                    Next
                                    If objForm.Items.Item("81").Specific.selected.value = "1" Or objForm.Items.Item("81").Specific.selected.value = "2" Then
                                        Me.Open_Order_Allocation_Form(pVal.FormUID, Trim(objForm.Items.Item("8").Specific.value), MAC_ID)
                                    End If
                                End If
                            End If
                        End If
                        If pVal.ItemUID = "2" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            RS.DoQuery("Delete From [@GEN_SZ_ORDR] Where u_sono = '" + Trim(objForm.Items.Item("8").Specific.value) + "' And u_macid = '" + MAC_ID + "'")
                        End If
                End Select
            ElseIf pVal.BeforeAction = True And ModalForm = True And pVal.FormUID = (objSubForm.UniqueID.Substring(objSubForm.UniqueID.IndexOf("@") + 1)) Then
                objSubForm = oApplication.Forms.Item("SOORD@" & pVal.FormUID)
                objSubForm.Select()
                BubbleEvent = False
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
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRSet.DoQuery("Select UserId From OUSR Where User_Code = '" + oCompany.UserName.Trim + "        '")
                        oRS.DoQuery("Select Max(DocEntry) AS 'DocEntry' From ORDR Where UserSign = '" + Trim(oRSet.Fields.Item("UserID").Value) + "'")
                        oRSet.DoQuery("Select DocNum From  ORDR Where DocEntry = '" + Trim(oRS.Fields.Item("DocEntry").Value) + "'")
                        oRS.DoQuery("Update [@GEN_SZ_ORDR] Set u_sono = '" + Trim(oRSet.Fields.Item("DocNum").Value) + "' Where u_sono = '" + SONO + "' And u_macid = '" + MAC_ID + "'")
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation_SalesOrder_Allocation(ByVal FormUID As String) As Boolean
        Try
            objSubForm = oApplication.Forms.Item(FormUID)
            Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim ogrd As SAPbouiCOM.Grid
            ogrd = objSubForm.Items.Item("grd").Specific
            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
            For i As Integer = 0 To ogrd.Rows.Count - 1
                RS.DoQuery("Select IsNull(Sum(Convert(money,u_qty)),0) AS 'Qty' From [@GEN_SZ_ORDR] Where u_sono = '" + ogrd.DataTable.Columns.Item("OrderNo").Cells.Item(i).Value.ToString.Trim + "' And u_itemcode = '" + ogrd.DataTable.Columns.Item("ItemCode").Cells.Item(i).Value.ToString.Trim + "' And u_macid = '" + MAC_ID + "' And u_asrtcode = '" + ogrd.DataTable.Columns.Item("AssortedCode").Cells.Item(i).Value.ToString.Trim + "'")
                If CDbl(ogrd.DataTable.Columns.Item("Qty").Cells.Item(i).Value) <> CDbl(RS.Fields.Item("Qty").Value) Then
                    oApplication.StatusBar.SetText("Please enter sizes for the items", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                    Exit Function
                End If
            Next
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Sub ItemEvent_SalesOrder_Allocation(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Dim CHILD_FORM As String = "GEN_SZ_ORDR@" & pVal.FormUID
            Dim ChildModalForm As Boolean = False
            For i As Integer = 0 To oApplication.Forms.Count - 1
                If oApplication.Forms.Item(i).UniqueID = (CHILD_FORM) Then
                    objSubForm = oApplication.Forms.Item("GEN_SZ_ORDR@" & pVal.FormUID)
                    ChildModalForm = True
                    Exit For
                End If
            Next
            If ChildModalForm = False Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                        objSubForm = oApplication.Forms.Item(pVal.FormUID)
                        PARENT_FORM = (pVal.FormUID.Substring(objSubForm.UniqueID.IndexOf("@") + 1))
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                            If Me.Validation_SalesOrder_Allocation(pVal.FormUID) = False Then
                                BubbleEvent = False
                            End If
                        End If
                        If pVal.ItemUID = "2" And pVal.BeforeAction = False Then
                            ModalForm = False
                        End If
                        If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            ModalForm = False
                            objSubForm.Close()
                            objForm = oApplication.Forms.ActiveForm
                            'Dim oMatrix As SAPbouiCOM.Matrix
                            'oMatrix = objForm.Items.Item("38").Specific
                            'oMatrix.Columns.Item("U_sizemtx").Editable = True
                            'For i As Integer = 1 To oMatrix.VisualRowCount - 1
                            '    oMatrix.Columns.Item("U_sizemtx").Cells.Item(i).Specific.value = "Yes"
                            'Next
                            'oMatrix.Columns.Item("U_sizemtx").Editable = False
                        End If
                        If pVal.ItemUID = "chs" And pVal.BeforeAction = True Then
                            Dim flg As Boolean = False
                            Dim slflag As Boolean = False
                            Dim ogrd As SAPbouiCOM.Grid
                            ogrd = objSubForm.Items.Item("grd").Specific
                            For i As Integer = 0 To ogrd.Rows.Count - 1
                                If ogrd.Rows.IsSelected(i) = True Then
                                    flg = True
                                End If
                            Next
                            If flg = False Then
                                oApplication.StatusBar.SetText("Please select items", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                        If pVal.ItemUID = "chs" And pVal.BeforeAction = False Then
                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            Dim ogrd As SAPbouiCOM.Grid
                            ogrd = objSubForm.Items.Item("grd").Specific
                            For i As Integer = 0 To ogrd.Rows.Count - 1
                                If ogrd.Rows.IsSelected(i) = True Then
                                    Me.Open_Size_Matrix_Form(pVal.FormUID, CStr(ogrd.DataTable.Columns.Item("OrderNo").Cells.Item(i).Value).Trim(), CStr(ogrd.DataTable.Columns.Item("ItemCode").Cells.Item(i).Value).Trim(), CStr(ogrd.DataTable.Columns.Item("Qty").Cells.Item(i).Value).Trim(), MAC_ID, CStr(ogrd.DataTable.Columns.Item("AssortedCode").Cells.Item(i).Value).Trim())
                                End If
                            Next
                        End If
                End Select
            ElseIf pVal.BeforeAction = True And ChildModalForm = True And pVal.FormUID = (objSubForm.UniqueID.Substring(objSubForm.UniqueID.IndexOf("@") + 1)) Then
                objSubForm = oApplication.Forms.Item("GEN_SZ_ORDR@" & pVal.FormUID)
                objSubForm.Select()
                BubbleEvent = False
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FormDataEvent_SalesOrder_Allocation(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.ActionSuccess = True Then
                        'objSubForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        'objSubForm.EnableMenu("1281", True)
                        'objSubForm.Items.Item("ordrno").Specific.Value = orderno
                        'objSubForm.Items.Item("macid").Specific.value = hwid
                        'objSubForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        'objSubForm.EnableMenu("1281", False)
                    End If
            End Select

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation_SalesOrder_SizeMatrix(ByVal FormUID As String) As Boolean
        Try
            Dim errflag As Boolean = False
            Dim Total As Double
            objSForm = oApplication.Forms.Item(FormUID)
            Dim ogrd As SAPbouiCOM.Grid
            ogrd = objSForm.Items.Item("grd").Specific
            For i As Integer = 0 To ogrd.Rows.Count - 1
                Total = Total + ogrd.DataTable.Columns.Item("Qty").Cells.Item(i).Value
            Next
            If Total <> TotQty Then
                oApplication.StatusBar.SetText("Total quantity should be equal to the amount in sales order", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Sub ItemEvent_SalesOrder_SizeMatrix(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objSForm = oApplication.Forms.Item(pVal.FormUID)
                    PARENT_FORM = (pVal.FormUID.Substring(objSForm.UniqueID.IndexOf("@") + 1))
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation_SalesOrder_SizeMatrix(pVal.FormUID) = False Then
                            BubbleEvent = False
                        End If
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        'Dim grd As SAPbouiCOM.Grid = objSForm.Items.Item("grd").Specific
                        'Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'Dim Code As String = RS.Fields.Item("Code").Value
                        'For i As Integer = 0 To grd.Rows.Count - 1
                        '    If grd.DataTable.Columns.Item("Qty").Cells.Item(i).Value > 0 Then
                        '        RS.DoQuery("Select Count(*) + 1 AS 'Code' From [@GEN_SZ_ORDR]")
                        '        oRSet.DoQuery("Insert Into [@GEN_SZ_ORDR] (Code,u_sono,u_itemcode,u_macid,u_size,u_qty) Values('" + Trim(RS.Fields.Item("Code").Value) + "','" + grd.DataTable.Columns.Item("SalesOrderNO").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("ItemCode").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("Size").Cells.Item(i).Value.ToString.Trim + "'," + grd.DataTable.Columns.Item("Qty").Cells.Item(i).Value + ") ")
                        '    End If
                        'Next
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Dim grd As SAPbouiCOM.Grid = objSForm.Items.Item("grd").Specific
                        Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' Dim Code As String = RS.Fields.Item("Code").Value
                        oRSet.DoQuery("Delete From [@GEN_SZ_ORDR] Where u_sono = '" + GSONO + "' And u_macid = '" + GMACID + "' And u_itemcode = '" + GITEMCODE + "' And u_asrtcode = '" + GASRTCODE + "'")
                        For i As Integer = 0 To grd.Rows.Count - 1
                            If grd.DataTable.Columns.Item("Qty").Cells.Item(i).Value > 0 Then
                                RS.DoQuery("Select Convert(VarChar,Count(*) + 1) AS 'Code' From [@GEN_SZ_ORDR]")
                                oRSet.DoQuery("Insert Into [@GEN_SZ_ORDR] (Code,Name,u_sono,u_itemcode,u_asrtcode,u_macid,u_size,u_qty,u_cutqty) Values('" + Trim(RS.Fields.Item("Code").Value) + "','" + Trim(RS.Fields.Item("Code").Value) + "','" + grd.DataTable.Columns.Item("SalesOrderNO").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("ItemCode").Cells.Item(i).Value.ToString.Trim + "','" + GASRTCODE + "','" + MAC_ID + "','" + grd.DataTable.Columns.Item("Size").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("Qty").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("CutQty").Cells.Item(i).Value.ToString.Trim + "') ")
                            End If
                        Next
                        objSForm.Close()
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FormDataEvent_SalesOrder_SizeMatrix(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.ActionSuccess = True Then
                        objSForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        objSForm.EnableMenu("1281", True)
                        Dim DBSource As SAPbouiCOM.DBDataSource
                        DBSource = objSForm.DataSources.DBDataSources.Item("@GEN_SIZE_MX")
                        DBSource.SetValue("U_ordrno", 0, sorderno)
                        DBSource.SetValue("U_macid", 0, shwid)
                        DBSource.SetValue("U_itemcode", 0, sitemcode)
                        objSForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objSForm.EnableMenu("1281", False)
                    End If
            End Select

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub Open_Order_Allocation_Form(ByVal FormUID As String, ByVal ordrno As String, ByVal macid As String)
        Try
            PARENT_FORM = FormUID
            Dim RS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim CHILD_FORM As String = "SOORD@" & FormUID
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
                oUtilities.SAPXML("SOORD.xml", CHILD_FORM)
                objSubForm = oApplication.Forms.Item(CHILD_FORM)
                objSubForm.Select()
            End If
            ChildModalForm = True
            Dim ogrid As SAPbouiCOM.Grid
            ogrid = objSubForm.Items.Item("grd").Specific
            objSubForm.DataSources.DataTables.Add("MyDataTable")
            objSubForm.DataSources.DataTables.Item(0).ExecuteQuery("Select Distinct A.OrdrNo 'OrderNo',A.itemcode 'ItemCode',B.ItemName 'ItemName',A.asrtcode 'AssortedCode',Sum(Convert(Money,qty)) 'Qty' From tmp_ordr_items A Inner Join OITM B On A.ItemCOde = B.ItemCode  Where A.ordrno = '" + ordrno + "' And A.macid = '" + macid + "' Group By A.OrdrNo,A.ItemCode,B.ItemName,A.asrtcode")
            ogrid.DataTable = objSubForm.DataSources.DataTables.Item("MyDataTable")
            For i As Integer = 0 To ogrid.Columns.Count - 1
                ogrid.Columns.Item(i).Editable = False
            Next
            ogrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub Open_Size_Matrix_Form(ByVal FormUID As String, ByVal ordrno As String, ByVal itemno As String, ByVal quantity As String, ByVal macid As String, ByVal asrtcode As String)
        Try
            PARENT_FORM = FormUID
            Dim CHILD_FORM As String = "GEN_SZ_ORDR@" & FormUID
            Dim oBool As Boolean = False
            For i As Integer = 0 To oApplication.Forms.Count - 1
                If oApplication.Forms.Item(i).UniqueID = (CHILD_FORM) Then
                    objSForm = oApplication.Forms.Item(CHILD_FORM)
                    objSForm.Select()
                    oBool = True
                    Exit For
                End If
            Next
            If oBool = False Then
                oUtilities.SAPXML("SizeMatrix.xml", CHILD_FORM)
                objSForm = oApplication.Forms.Item(CHILD_FORM)
                objSForm.Select()
            End If
            ChildModalForm = True
            TotQty = quantity
            GSONO = ordrno
            GMACID = macid
            GITEMCODE = itemno
            GASRTCODE = asrtcode
            Dim ogrid As SAPbouiCOM.Grid
            ogrid = objSForm.Items.Item("grd").Specific
            objSForm.DataSources.DataTables.Add("MyDataTable")
            objSForm.DataSources.DataTables.Item(0).ExecuteQuery("Exec Allocate_Size '" + ordrno + "','" + itemno + "','" + quantity + "','" + macid + "','" + asrtcode + "'")
            ogrid.DataTable = objSForm.DataSources.DataTables.Item("MyDataTable")
            For i As Integer = 0 To ogrid.Columns.Count - 1
                ogrid.Columns.Item(i).Editable = False
            Next
            ogrid.Columns.Item(ogrid.Columns.Count - 1).Editable = True
            ogrid.Columns.Item(ogrid.Columns.Count - 2).Editable = True
            ogrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefaultSEM(ByVal FormUID As String)
        Try
            objSubForm = oApplication.Forms.Item(FormUID)
            objSubForm.Freeze(True)
            If objSubForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                objSubForm.EnableMenu("1282", False)
            End If
            objSubForm.PaneLevel = 1
            objSubMatrix = objSubForm.Items.Item("OrdrMatrix").Specific
            objSubMatrix.Clear()
            objSubMatrix.FlushToDataSource()
            objSubMatrix.Clear()
            objSubMatrix.AddRow()
            objSubMatrix.FlushToDataSource()
            Me.SetNewLineSEM(objSubForm.UniqueID, objSubMatrix.VisualRowCount, objSubMatrix)
            objSubForm.Freeze(False)
        Catch ex As Exception
            objSubForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLineSEM(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objSubForm = oApplication.Forms.Item(FormUID)
            SZDBDetail = objSubForm.DataSources.DBDataSources.Item("@ORDR_ITEMS")
            objMatrix = oMatrix
            objSubForm.Freeze(True)
            objMatrix.FlushToDataSource()
            SZDBDetail.Offset = Row - 1
            SZDBDetail.SetValue("u_sono", SZDBDetail.Offset, "")
            SZDBDetail.SetValue("u_itemcode", SZDBDetail.Offset, "")
            SZDBDetail.SetValue("u_itemname", SZDBDetail.Offset, "")
            SZDBDetail.SetValue("u_qty", SZDBDetail.Offset, "")
            SZDBDetail.SetValue("u_asrtcode", SZDBDetail.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objSubForm.Freeze(False)
        Catch ex As Exception
            objSubForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLineSZ(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objSForm = oApplication.Forms.Item(FormUID)
            SMDBDetail = objSForm.DataSources.DBDataSources.Item("@GEN_SIZE_MX_D0")
            objMatrix = oMatrix
            objSForm.Freeze(True)
            objMatrix.FlushToDataSource()
            SMDBDetail.Offset = Row - 1
            SMDBDetail.SetValue("LineId", SMDBDetail.Offset, objMatrix.VisualRowCount)
            SMDBDetail.SetValue("u_size", SMDBDetail.Offset, "")
            SMDBDetail.SetValue("u_qty", SMDBDetail.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objSForm.Freeze(False)
        Catch ex As Exception
            objSForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = True Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "1293"
                        If objForm.TypeEx = "139" Then
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            objMatrix = objForm.Items.Item("38").Specific
                            DeleteItemCode = objMatrix.Columns.Item("1").Cells.Item(RowID).Specific.Value
                            oRecordSet.DoQuery("Delete From [@GEN_SZ_ORDR] Where u_itemcode = '" + DeleteItemCode + "' And u_sono = '" + Trim(objForm.Items.Item("8").Specific.value) + "'")
                        End If
                    Case "1288", "1289", "1290", "1291"
                        If objForm.TypeEx = "139" Then
                            BubbleEvent = False
                        End If
                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            RowID = eventInfo.Row
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub LoadFreight(ByVal FormUID As String)
        
        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oForm As SAPbouiCOM.Form = oApplication.Forms.Item(FormUID)
        oForm.Items.Item("91").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
        Dim FrgtForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
        Dim FrgtMatrix As SAPbouiCOM.Matrix = FrgtForm.Items.Item("3").Specific
        For k As Integer = 1 To FrgtMatrix.VisualRowCount
            oRSet.DoQuery("Insert Into Freight_Order_Rdr(RowNo,ExpnsCode,MacID) Values('" + k.ToString.Trim + "','" + Trim(FrgtMatrix.Columns.Item("1").Cells.Item(k).Specific.value) + "','" + MAC_ID + "')")
        Next
        Try
            FrgtForm.Freeze(True)
            For k As Integer = 1 To FrgtMatrix.VisualRowCount
                FrgtMatrix.Columns.Item("3").Cells.Item(k).Specific.Value = 0
            Next
            oRecordSet.DoQuery("Select u_fcode,u_fname,u_tax,u_amount,u_posfrgt,u_postax,u_negfrgt,u_negtax From [@GEN_ACCRUALS_RDR] Where u_invno = '" + InvNo + "' And u_macid = '" + MAC_ID + "'")
            For i As Integer = 1 To oRecordSet.RecordCount
                oRSet.DoQuery("Select Top 1 RowNo From Freight_Order_Rdr Where ExpnsCode = '" + Trim(oRecordSet.Fields.Item("u_fcode").Value) + "' And Macid = '" + MAC_ID + "'")
                If oRSet.RecordCount > 0 Then
                    Dim RowNo As Integer = oRSet.Fields.Item("RowNo").Value
                    If CDbl(FrgtMatrix.Columns.Item("3").Cells.Item(RowNo).Specific.Value.ToString.Substring(3)) = 0 Then
                        FrgtMatrix.Columns.Item("3").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_amount").Value
                        FrgtMatrix.Columns.Item("17").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_tax").Value
                    End If
                End If
                oRSet.DoQuery("Select Top 1 RowNo From Freight_Order_Rdr Where ExpnsCode = '" + Trim(oRecordSet.Fields.Item("u_posfrgt").Value) + "' And Macid = '" + MAC_ID + "'")
                If oRSet.RecordCount > 0 Then
                    RowNo = oRSet.Fields.Item("RowNo").Value
                    If CDbl(FrgtMatrix.Columns.Item("3").Cells.Item(RowNo).Specific.Value.ToString.Substring(3)) = 0 Then
                        FrgtMatrix.Columns.Item("3").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_amount").Value
                        FrgtMatrix.Columns.Item("17").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_postax").Value
                    End If
                End If
                oRSet.DoQuery("Select Top 1 RowNo From Freight_Order_Rdr Where ExpnsCode = '" + Trim(oRecordSet.Fields.Item("u_negfrgt").Value) + "' And Macid = '" + MAC_ID + "'")
                If oRSet.RecordCount > 0 Then
                    RowNo = oRSet.Fields.Item("RowNo").Value
                    If CDbl(FrgtMatrix.Columns.Item("3").Cells.Item(RowNo).Specific.Value.ToString.Substring(3)) = 0 Then
                        FrgtMatrix.Columns.Item("3").Cells.Item(RowNo).Specific.Value = "-" & oRecordSet.Fields.Item("u_amount").Value
                        FrgtMatrix.Columns.Item("17").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_negtax").Value
                    End If
                End If
                oRecordSet.MoveNext()
            Next


            'For i As Integer = 1 To FrgtMatrix.VisualRowCount
            '    oRecordSet.DoQuery("Select u_fcode,u_fname,u_tax,u_amount,u_posfcode,u_postax,u_negfcode,u_negtax From [@GEN_ACCRUALS] Where u_invno = '" + InvNo + "' And u_macid = '" + MAC_ID + "' And u_fcode = '" + Trim(FrgtMatrix.Columns.Item("1").Cells.Item(i).Specific.Value) + "'")
            '    If oRecordSet.RecordCount > 0 Then
            '        FrgtMatrix.Columns.Item("3").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("u_amount").Value
            '        FrgtMatrix.Columns.Item("17").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("u_tax").Value
            '    End If
            'Next

            'oRSet.DoQuery("Select (Select ExpnsCode From OEXD Where ExpnsName = B.u_nfreight) As 'FrgtCode',B.u_nfreight As 'FrgtName',A.u_fcode,(Select U_appltax From OEXD Where ExpnsName = B.u_nfreight) As 'Tax',A.u_amount From [@GEN_ACCRUALS] A INNER JOIN OEXD B ON A.u_fcode = B.ExpnsCode And A.u_invno = '" + InvNo + "' And A.u_macid = '" + MAC_ID + "'")
            'While Not oRSet.EoF
            '    For i As Integer = 1 To FrgtMatrix.VisualRowCount
            '        If Trim(FrgtMatrix.Columns.Item("1").Cells.Item(i).Specific.Value) = Trim(oRSet.Fields.Item("FrgtCode").Value) Then
            '            FrgtMatrix.Columns.Item("3").Cells.Item(i).Specific.Value = -oRSet.Fields.Item("u_amount").Value
            '            FrgtMatrix.Columns.Item("17").Cells.Item(i).Specific.Value = oRSet.Fields.Item("Tax").Value
            '        End If
            '    Next
            '    oRSet.MoveNext()
            'End While
            FrgtForm.Freeze(False)
            FrghtFlag = False
        Catch ex As Exception
            FrgtForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub LoadItems(ByVal FormUID As String)
        Try
            'Dim ITForm As SAPbouiCOM.Form
            Dim ITMatrix As SAPbouiCOM.Matrix
            objPreForm = oApplication.Forms.GetForm("PRE_SHIPMENT", 1)
            'objPreForm = oApplication.Forms.Item(FormUID)
            oDBs_Pre_Head = objPreForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
            oDBs_Pre_Detail = objPreForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D0")
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select B.ItemCode,B.Dscription,B.Quantity,B.unitMsr,B.Price,B.TaxCode,A.DocNum,B.LineNum,B.LineTotal,B.WhsCode,A.DocEntry,B.U_preqty,A.Address,A.NumAtCard,A.DocCur,A.DocRate,A.DocDueDate,A.SlpCode,A.OwnerCode,B.Currency,A.U_unit,A.GroupNum,A.SlpCode,A.OwnerCode,TotalSumSy From ORDR A Inner Join RDR1 B On A.DocEntry = B.DocEntry And A.CardCode = '" + objForm.Items.Item("4").Specific.value.ToString.Trim + "' And IsNull(A.DocStatus,'O') = 'O' And IsNull(B.U_prests,'Open') = 'Open' And A.DocNum IN (" + DOCNUM.ToString.Trim + ")")
            'ITForm = oApplication.Forms.Item(FormUID)
            Try
                objPreForm.Freeze(True)
                ITMatrix = objPreForm.Items.Item("ItemMatrix").Specific
                ITMatrix.Clear()
                ITMatrix.AddRow(1)
                If oRecordSet.RecordCount = 0 Then
                    objPreForm.Freeze(False)
                    Exit Sub
                End If
                'ITMatrix.Columns.Item("ItemMatrix").Editable = True
                'ITForm.Items.Item("U_SaleNo").Specific.Value = DOCNUM
                oDBs_Pre_Head.SetValue("U_CustCode", 0, objForm.Items.Item("4").Specific.Value)
                oDBs_Pre_Head.SetValue("U_CustName", 0, objForm.Items.Item("54").Specific.Value)
                oDBs_Pre_Head.SetValue("U_SaleNo", 0, DOCNUM)
                oDBs_Pre_Head.SetValue("U_CustRef", 0, oRecordSet.Fields.Item("NumAtCard").Value)
                oDBs_Pre_Head.SetValue("U_DocCur", 0, oRecordSet.Fields.Item("DocCur").Value)
                'oDBs_Pre_Head.SetValue("U_CF", 0, oRecordSet.Fields.Item("").Value)
                oDBs_Pre_Head.SetValue("U_Addr", 0, oRecordSet.Fields.Item("Address").Value)
                oDBs_Pre_Head.SetValue("U_JourRem", 0, "PreShipment -" & "" & objForm.Items.Item("4").Specific.Value)
                'oDBs_Pre_Head.SetValue("U_CF", 0, oRecordSet.Fields.Item("DocCur").Value)
                oDBs_Pre_Head.SetValue("U_Unit", 0, oRecordSet.Fields.Item("U_unit").Value)
                objPreForm.Items.Item("unit").Enabled = False

                Dim dats As DateTime = oRecordSet.Fields.Item("DocDueDate").Value
                oDBs_Pre_Head.SetValue("U_DelDate", 0, dats.ToString("yyyyMMdd"))
                Dim Pay, Sale, Own As String
                Pay = oRecordSet.Fields.Item("GroupNum").Value
                Sale = oRecordSet.Fields.Item("OwnerCode").Value
                Own = oRecordSet.Fields.Item("SlpCode").Value

                'oDBs_Head.SetValue("U_Buyer", 0, DOCNUM)
                For i As Integer = 1 To oRecordSet.RecordCount
                    Dim OrginRow As Integer = ITMatrix.VisualRowCount
                    Dim rowcount As Integer = oRecordSet.RecordCount
                    'If i < rowcount - 1 Then
                    '    objMatrix.AddRow(1, OrginRow)
                    '    oDBs_Pre_Detail.InsertRecord(OrginRow + i - 1)
                    'End If
                    ITMatrix.FlushToDataSource()
                    oDBs_Pre_Detail.Offset = i - 1
                    oDBs_Pre_Detail.SetValue("LineId", oDBs_Pre_Detail.Offset, i)
                    oDBs_Pre_Detail.SetValue("U_ItemCode", oDBs_Pre_Detail.Offset, oRecordSet.Fields.Item("ItemCode").Value.ToString.Trim)
                    oDBs_Pre_Detail.SetValue("U_ItemName", oDBs_Pre_Detail.Offset, oRecordSet.Fields.Item("Dscription").Value.ToString.Trim)
                    oDBs_Pre_Detail.SetValue("U_Quantity", oDBs_Pre_Detail.Offset, (oRecordSet.Fields.Item("Quantity").Value - oRecordSet.Fields.Item("U_preqty").Value))
                    oDBs_Pre_Detail.SetValue("U_Price", oDBs_Pre_Detail.Offset, oRecordSet.Fields.Item("Price").Value)
                    oDBs_Pre_Detail.SetValue("U_Price_A", oDBs_Pre_Detail.Offset, oRecordSet.Fields.Item("DocCur").Value.ToString + " " + oRecordSet.Fields.Item("Price").Value.ToString)
                    oDBs_Pre_Detail.SetValue("U_UOM", oDBs_Pre_Detail.Offset, oRecordSet.Fields.Item("unitMsr").Value)
                    oDBs_Pre_Detail.SetValue("U_DocCur", oDBs_Pre_Detail.Offset, oRecordSet.Fields.Item("DocCur").Value.ToString)
                    oDBs_Pre_Detail.SetValue("U_TaxCode", oDBs_Pre_Detail.Offset, oRecordSet.Fields.Item("TaxCode").Value)
                    oDBs_Pre_Detail.SetValue("U_Whse", oDBs_Pre_Detail.Offset, oRecordSet.Fields.Item("WhsCode").Value)
                    oDBs_Pre_Detail.SetValue("U_TotalLC", oDBs_Pre_Detail.Offset, oRecordSet.Fields.Item("TotalSumSy").Value)
                    oDBs_Pre_Detail.SetValue("U_Total_A", oDBs_Pre_Detail.Offset, oRecordSet.Fields.Item("DocCur").Value.ToString + " " + oRecordSet.Fields.Item("TotalSumSy").Value.ToString)
                    oDBs_Pre_Detail.SetValue("U_SONo", oDBs_Pre_Detail.Offset, oRecordSet.Fields.Item("DocNum").Value)
                    oDBs_Pre_Detail.SetValue("U_BaseLine", oDBs_Pre_Detail.Offset, oRecordSet.Fields.Item("LineNum").Value.ToString)
                    oDBs_Pre_Detail.SetValue("U_BaseRef", oDBs_Pre_Detail.Offset, oRecordSet.Fields.Item("DocEntry").Value.ToString)
                    'oDBs_Pre_Detail.SetValue("U_Curr", oDBs_Pre_Detail.Offset, oRecordSet.Fields.Item("Currency").Value.ToString)
                    'ITMatrix.Columns.Item("qty").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("Quantity").Value
                    ''ITMatrix.Columns.Item("total").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("Total").Value
                    'ITMatrix.Columns.Item("saleno").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("DocNum").Value
                    'ITMatrix.Columns.Item("itemcode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    'objMatrix.AddRow(i)
                    oRecordSet.MoveNext()
                    ITMatrix.LoadFromDataSource()
                    If OrginRow = ITMatrix.VisualRowCount Then
                        ITMatrix.AddRow()
                        ITMatrix.FlushToDataSource()
                        Me.SetNewLine(objPreForm.UniqueID, ITMatrix.VisualRowCount)
                    End If
                Next

                oRecordSet.DoQuery("SELECT T0.[PymntGroup] FROM OCTG T0 WHERE T0.[GroupNum] = '" + Pay + "'")
                oDBs_Pre_Head.SetValue("U_PayTrms", 0, oRecordSet.Fields.Item("PymntGroup").Value)
                oRecordSet.DoQuery("SELECT T0.[SlpName] FROM OSLP T0 WHERE T0.[SlpCode] = '" + Sale + "'")
                oDBs_Pre_Head.SetValue("U_Buyer", 0, oRecordSet.Fields.Item("SlpName").Value)
                oRecordSet.DoQuery("SELECT (T0.[lastName] + ' ' + T0.[firstName]) 'OwnerName' FROM OHEM T0 WHERE T0.[empID]  = '" + Own + "'")
                oDBs_Pre_Head.SetValue("U_Owner", 0, oRecordSet.Fields.Item("OwnerName").Value)
                LoadUDF(objPreForm.UniqueID)
                oDBs_Pre_Head.SetValue("U_Remarks", 0, "Based On Sales Order No." & "" & objForm.Items.Item("8").Specific.Value)
                Me.CalculateTotal(objPreForm.UniqueID)
                'ITMatrix.Columns.Item("ItemMatrix").Editable = False
                objPreForm.Freeze(False)
            Catch ex As Exception
                objPreForm.Freeze(False)
                oApplication.StatusBar.SetText(ex.Message)
            End Try

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub LoadUDF(ByVal FormUID As String)
        Try
            objPreForm = oApplication.Forms.Item(FormUID)
            oDBs_Pre_Head = objPreForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select A.[U_DelTerm], A.[U_PortLoad], A.[U_CntyFin], A.[U_PorDisch], A.[U_FinDest], A.[U_GrssWt], A.[U_NetWt], A.[U_contno], A.[U_SUPP_PLC1], A.[U_NO_OF_CN]  From ORDR A Inner Join RDR1 B On A.DocEntry = B.DocEntry And A.CardCode = '" + objForm.Items.Item("4").Specific.value.ToString.Trim + "' And IsNull(A.DocStatus,'O') = 'O' And A.DocNum IN (" + DOCNUM.ToString.Trim + ")")
            oDBs_Pre_Head.SetValue("U_DelTerm", 0, oRecordSet.Fields.Item("U_DelTerm").Value)
            oDBs_Pre_Head.SetValue("U_PortLoad", 0, oRecordSet.Fields.Item("U_PortLoad").Value)
            oDBs_Pre_Head.SetValue("U_CntyFin", 0, oRecordSet.Fields.Item("U_CntyFin").Value)
            oDBs_Pre_Head.SetValue("U_PorDisch", 0, oRecordSet.Fields.Item("U_PorDisch").Value)
            oDBs_Pre_Head.SetValue("U_FinDest", 0, oRecordSet.Fields.Item("U_FinDest").Value)
            oDBs_Pre_Head.SetValue("U_GrssWt", 0, oRecordSet.Fields.Item("U_GrssWt").Value)
            oDBs_Pre_Head.SetValue("U_NetWt", 0, oRecordSet.Fields.Item("U_NetWt").Value)
            oDBs_Pre_Head.SetValue("U_contno", 0, oRecordSet.Fields.Item("U_contno").Value)
            'oDBs_Pre_Head.SetValue("U_SUPP_PLC1", 0, oRecordSet.Fields.Item("U_SUPP_PLC1").Value)
            'oDBs_Pre_Head.SetValue("U_NO_OF_CN", 0, oRecordSet.Fields.Item("U_NO_OF_CN").Value)
            'oDBs_Pre_Head.SetValue("U_DelTerm", 0, oRecordSet.Fields.Item("U_DelTerm").Value)
            'oDBs_Pre_Head.SetValue("U_PortLoad", 0, oRecordSet.Fields.Item("U_PortLoad").Value)
            'oDBs_Pre_Head.SetValue("U_DelTerm", 0, oRecordSet.Fields.Item("U_DelTerm").Value)
            'oDBs_Pre_Head.SetValue("U_PortLoad", 0, oRecordSet.Fields.Item("U_PortLoad").Value)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub CalculateTotal(ByVal FormUID As String)
        Try
            objPreForm = oApplication.Forms.Item(FormUID)
            oDBs_Pre_Head = objPreForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
            objMatrix = objPreForm.Items.Item("ItemMatrix").Specific
            Dim TotalLC = 0, totalTax As Double = 0
            For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                TotalLC = TotalLC + CDbl(objMatrix.Columns.Item("total").Cells.Item(Row).Specific.Value)
                totalTax = totalTax + CDbl(objMatrix.Columns.Item("taxamt").Cells.Item(Row).Specific.Value)
            Next
            oDBs_Pre_Head.SetValue("U_TotBefTa", 0, TotalLC)
            oDBs_Pre_Head.SetValue("U_Tax", 0, totalTax)
            oDBs_Pre_Head.SetValue("U_Total", 0, TotalLC + totalTax + objPreForm.Items.Item("roundpr").Specific.Value + objPreForm.Items.Item("freight").Specific.Value)

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer)
        Try
            objPreForm = oApplication.Forms.Item(FormUID)
            oDBs_Pre_Head = objPreForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
            oDBs_Pre_Detail = objPreForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D0")
            objMatrix = objPreForm.Items.Item("ItemMatrix").Specific
            oDBs_Pre_Detail.Offset = Row - 1
            oDBs_Pre_Detail.SetValue("LineId", oDBs_Pre_Detail.Offset, Row)
            oDBs_Pre_Detail.SetValue("U_ItemCode", oDBs_Pre_Detail.Offset, "")
            oDBs_Pre_Detail.SetValue("U_ItemName", oDBs_Pre_Detail.Offset, "")
            oDBs_Pre_Detail.SetValue("U_Quantity", oDBs_Pre_Detail.Offset, 0)
            oDBs_Pre_Detail.SetValue("U_Price", oDBs_Pre_Detail.Offset, 0)
            oDBs_Pre_Detail.SetValue("U_Price_A", oDBs_Pre_Detail.Offset, "")
            oDBs_Pre_Detail.SetValue("U_DocCur", oDBs_Pre_Detail.Offset, oDBs_Pre_Head.GetValue("U_DocCur", 0).Trim)
            oDBs_Pre_Detail.SetValue("U_TotalLC", oDBs_Pre_Detail.Offset, 0)
            oDBs_Pre_Detail.SetValue("U_Total_A", oDBs_Pre_Detail.Offset, "")
            oDBs_Pre_Detail.SetValue("U_TaxCode", oDBs_Pre_Detail.Offset, "")
            oDBs_Pre_Detail.SetValue("U_UOM", oDBs_Pre_Detail.Offset, "")
            oDBs_Pre_Detail.SetValue("U_TaxAmt", oDBs_Pre_Detail.Offset, 0)
            oDBs_Pre_Detail.SetValue("U_Whse", oDBs_Pre_Detail.Offset, "")
            oDBs_Pre_Detail.SetValue("U_Whse", oDBs_Pre_Detail.Offset, "")
            oDBs_Pre_Detail.SetValue("U_BaseLine", oDBs_Pre_Detail.Offset, "")
            oDBs_Pre_Detail.SetValue("U_BaseRef", oDBs_Pre_Detail.Offset, "")
            oDBs_Pre_Detail.SetValue("U_Note", oDBs_Pre_Detail.Offset, "")
            oDBs_Pre_Detail.SetValue("U_Remarks", oDBs_Pre_Detail.Offset, "")
            objMatrix.SetLineData(Row)
            objMatrix.FlushToDataSource()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub Open_Accruals_Form(ByVal FormUID As String, ByVal InvoiceNo As String, ByVal MACID As String, ByVal Mode As String, ByVal BASENO As String)
        Try
            PARENT_FORM = FormUID
            Dim RS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim RS2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim CHILD_FORM As String = "ACCRUALS_ORDER@" & FormUID
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
                oUtilities.SAPXML("Accruals_Order.xml", CHILD_FORM)
                objSubForm = oApplication.Forms.Item(CHILD_FORM)
                objSubForm.Select()
            End If
            ChildModalForm = True
            Dim ogrid As SAPbouiCOM.Grid
            ogrid = objSubForm.Items.Item("grd").Specific
            objSubForm.DataSources.DataTables.Add("MyDataTable")
            If Mode = "A" Then
                Dim oRs2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRs2.DoQuery("Delete From [@GEN_ACCRUALS_RDR] ")
                RS1.DoQuery("Select u_invno From [@GEN_ACCRUALS_RDR] Where u_invno = '" + InvoiceNo + "'")
                If RS1.RecordCount > 0 Then
                    objSubForm.DataSources.DataTables.Item(0).ExecuteQuery("Select ExpnsCode As 'Freight Code',ExpnsName As 'Freight Name', U_appltax As 'Tax Applicable',IsNull((Select u_amount From [@GEN_ACCRUALS_RDR] Where u_fcode = ExpnsCode And u_invno = '" + InvoiceNo + "' And u_macid = '" + MACID + "'),0) As 'Freight Amount' From OEXD Where IsNull(u_incl,'NO') = 'YES'")
                    ogrid.DataTable = objSubForm.DataSources.DataTables.Item("MyDataTable")
                Else
                    objSubForm.DataSources.DataTables.Item(0).ExecuteQuery("Exec UBG_Acrruals_Quotation_BaseNum '" + InvoiceNo + "','" + MACID + "'")
                    ogrid.DataTable = objSubForm.DataSources.DataTables.Item("MyDataTable")
                    If Dbk > 0 Then
                        ogrid.DataTable.Columns.Item("Freight Amount").Cells.Item(0).Value = "INR " + CStr(Dbk)
                    End If
                    If ANSP > 0 Then
                        ogrid.DataTable.Columns.Item("Freight Amount").Cells.Item(1).Value = "INR " + CStr(ANSP)
                    End If
                    If INS > 0 Then
                        ogrid.DataTable.Columns.Item("Freight Amount").Cells.Item(3).Value = "INR " + CStr(INS)
                    End If
                    If COMM > 0 Then
                        ogrid.DataTable.Columns.Item("Freight Amount").Cells.Item(2).Value = "USD " + CStr(COMM)
                    End If
                    'If TRANS > 0 Then
                    '    ogrid.DataTable.Columns.Item("Freight Amount").Cells.Item(4).Value = "USD " + CStr(COMM)
                    'End If

                End If
            End If
            If Mode = "O" Then
                RS1.DoQuery("Select B.ExpnsCode,B.LineTotal From ORDR A Inner Join RDR3 B On A.DocEntry = B.DocEntry Inner Join OEXD C On B.ExpnsCode = C.ExpnsCode And C.u_incl = 'YES' Where A.DocNum = '" + InvoiceNo + "'")
                If RS1.RecordCount > 0 Then
                    objSubForm.DataSources.DataTables.Item(0).ExecuteQuery("Exec UBG_Acrruals_Order '" + InvoiceNo + "'")
                    ogrid.DataTable = objSubForm.DataSources.DataTables.Item("MyDataTable")
                Else
                    objSubForm.DataSources.DataTables.Item(0).ExecuteQuery("Exec UBG_Acrruals_Quotation_BaseNum '" + InvoiceNo + "','" + MACID + "'")
                    ogrid.DataTable = objSubForm.DataSources.DataTables.Item("MyDataTable")
                End If
            End If
            RS2.DoQuery("Delete From ORDR_BASENUM Where macid = '" + MACID + "'")
            ' ogrid.DataTable = objSubForm.DataSources.DataTables.Item("MyDataTable")
            For i As Integer = 0 To ogrid.Columns.Count - 1
                ogrid.Columns.Item(i).Editable = False
            Next
            ogrid.Columns.Item(ogrid.Columns.Count - 1).Editable = True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
End Class
