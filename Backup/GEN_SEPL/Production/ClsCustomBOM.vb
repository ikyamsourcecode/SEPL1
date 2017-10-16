Imports System.Threading
Imports System.Globalization
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text

Public Class ClsCustomBOM

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
            oUtilities.SAPXML("GEN_CUST_BOM.xml")
            objForm = oApplication.Forms.GetForm("GEN_CUST_BOM", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_CUST_BOM")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_CUST_BOM_D0")
            objForm.EnableMenu("1288", True)
            objForm.EnableMenu("1289", True)
            objForm.EnableMenu("1290", True)
            objForm.EnableMenu("1291", True)
            objForm.EnableMenu("1294", True)
            objForm.EnableMenu("1287", True)

            objForm.DataBrowser.BrowseBy = "docnum"
            objForm.Items.Item("sono").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("itemcode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            objForm.EnableMenu("5890", True)
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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_CUST_BOM")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_CUST_BOM_D0")
            oUtilities.GetSeries(FormUID, "series", "GEN_CUST_BOM")
            oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("series").Specific.Selected.Value, "GEN_CUST_BOM"))
            oDBs_Head.SetValue("U_docdate", 0, DateTime.Today.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_status", 0, "NEW")
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSet.DoQuery("Select U_Name From OUSR WHere User_Code = '" + oCompany.UserName.Trim + "'")
            oDBs_Head.SetValue("U_empname", 0, oRSet.Fields.Item("U_Name").Value)
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_CUST_BOM_D0")
            objMatrix = oMatrix
            objMatrix.FlushToDataSource()
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, objMatrix.VisualRowCount)
            oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_size", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_qty", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_ordrqty", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_per", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_totqty", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_uom", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_process", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_cardcode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_cardname", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_issmthd", oDBs_Detail.Offset, "B")
            oDBs_Detail.SetValue("u_deleted", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_status", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_remarks", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_place", oDBs_Detail.Offset, "")
            'Vijeesh
            oDBs_Detail.SetValue("U_fremark", oDBs_Detail.Offset, "")
            'Vijeesh
            objMatrix.SetLineData(objMatrix.VisualRowCount)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterSOItems(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("ITCFL")
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Distinct B.ItemCode From ORDR A Inner Join RDR1 B On A.DocEntry = B.DocEntry Where A.DocNum = '" + Trim(objForm.Items.Item("sono").Specific.value) + "'")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            For IntICount As Integer = 0 To oRecordSet.RecordCount - 1
                If IntICount = (oRecordSet.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("ItemCode").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("ItemCode").Value
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
        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim RS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim ToWhs As String
        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
        oRecordSet.DoQuery("Delete From TMP_CST_BOM Where MacId = '" + MAC_ID + "'")
        Try
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oPRODTREE As SAPbobsCOM.ProductTrees
            Dim oPRODTREE_LINES As SAPbobsCOM.ProductTrees_Lines
            'Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
            Dim ItemCode As String
            Dim SOQty As Double
            Dim ErrUpd, ErrAdd, ErrDel As Integer
            objMatrix = objForm.Items.Item("mtx").Specific
            If Trim(objForm.Items.Item("sono").Specific.value) = "" Then
                oApplication.StatusBar.SetText("Please enter SO No", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If Trim(objForm.Items.Item("itemcode").Specific.value) = "" Then
                oApplication.StatusBar.SetText("Please select item", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If Trim(objForm.Items.Item("unit").Specific.value) = "" Then
                oApplication.StatusBar.SetText("Please enter Unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(1).Specific.value) = "" Then
                oApplication.StatusBar.SetText("Please enter items", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_CUST_BOM")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_CUST_BOM_D0")
            If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(1).Specific.value) = "" Then
                oApplication.StatusBar.SetText("Please enter items", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If Trim(oDBs_Head.GetValue("u_sono", 0)) = "" Then
                oApplication.StatusBar.SetText("Please select Sales Order", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If Trim(oDBs_Head.GetValue("u_itemcode", 0)) = "" Then
                oApplication.StatusBar.SetText("Please select items", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If oDBs_Head.GetValue("u_qty", 0) <= 0 Then
                oApplication.StatusBar.SetText("Please enter quantity greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If oDBs_Head.GetValue("u_unit", 0) = "" Then
                oApplication.StatusBar.SetText("Please enter unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            For i As Integer = 1 To objMatrix.VisualRowCount
                If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) <> "" Then
                    If Trim(objMatrix.Columns.Item("process").Cells.Item(i).Specific.value) = "" Then
                        oApplication.StatusBar.SetText("Please enter process for all items", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            Next
            'oRecordSet.DoQuery("Delete From TMP_CST_BOM Where MacId = '" + MAC_ID + "'")
            'For i As Integer = 1 To objMatrix.VisualRowCount
            '    If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) <> "" Then
            '        oRSet.DoQuery("Insert Into TMP_CST_BOM(SONO,DOCNUM,ItemCode,SOQty,Unit,ChldItem,Qty,Process,IssMthd,Status,MacId,Deleted) Values('" + Trim(objForm.Items.Item("sono").Specific.value) + "','" + Trim(objForm.Items.Item("docnum").Specific.value) + "','" + Trim(objForm.Items.Item("itemcode").Specific.value) + "','" + Trim(objForm.Items.Item("qty").Specific.value) + "','" + Trim(objForm.Items.Item("unit").Specific.Value) + "','" + Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) + "','" + Trim(objMatrix.Columns.Item("qty").Cells.Item(i).Specific.value) + "','" + Trim(objMatrix.Columns.Item("process").Cells.Item(i).Specific.value) + "','" + Trim(objMatrix.Columns.Item("issmthd").Cells.Item(i).Specific.selected.value) + "','" + Trim(objMatrix.Columns.Item("status").Cells.Item(i).Specific.Selected.Value) + "','" + MAC_ID + "','" + Trim(objMatrix.Columns.Item("deleted").Cells.Item(i).Specific.Selected.value) + "')")
            '    End If
            'Next
            'Dim Counter As Integer
            'ItemCode = objForm.Items.Item("itemcode").Specific.value
            'SOQty = objForm.Items.Item("qty").Specific.value
            'If Trim(oDBs_Head.GetValue("u_status", 0)) = "NEW" Then
            '    oRecordSet.DoQuery("Select Distinct B.u_itemcode,B.u_process,A.Code From [@GEN_PROD_PRCS] A INNER JOIN [@GEN_PROD_PRCS_D0] B ON A.Code = B.Code Where A.u_itemcode = '" + ItemCode + "'")
            '    For k As Integer = 1 To oRecordSet.RecordCount
            '        oPRODTREE = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees)
            '        oRSet.DoQuery("Select Distinct SONO,DOCNUm,ItemCode,SOQty,Unit,ChldItem,Qty,Process,IssMthd,Status,MacId From TMP_CST_BOM Where ItemCode = '" + ItemCode + "' And Process = '" + Trim(oRecordSet.Fields.Item("u_process").Value) + "'")
            '        oPRODTREE.TreeCode = oRecordSet.Fields.Item("u_itemcode").Value
            '        oPRODTREE.TreeType = SAPbobsCOM.BoItemTreeTypes.iProductionTree
            '        oPRODTREE.Quantity = 1
            '        oPRODTREE_LINES = oPRODTREE.Items
            '        RS1.DoQuery("Select B.u_inwhs From [@GEN_UNIT_MST] A INNER JOIN [@GEN_UNIT_MST_D0] B ON A.CODE = B.CODE Where A.Name = '" + Trim(oRSet.Fields.Item("Unit").Value) + "' And B.u_process = '" + Trim(oRSet.Fields.Item("Process").Value) + "'")
            '        ToWhs = RS1.Fields.Item("u_inwhs").Value
            '        For J As Integer = 1 To oRSet.RecordCount
            '            oPRODTREE_LINES.ParentItem = oRecordSet.Fields.Item("u_itemcode").Value
            '            oPRODTREE_LINES.ItemCode = oRSet.Fields.Item("ChldItem").Value
            '            oPRODTREE_LINES.Quantity = oRSet.Fields.Item("Qty").Value
            '            oPRODTREE_LINES.Warehouse = ToWhs
            '            If Trim(oRSet.Fields.Item("IssMthd").Value) = "B" Then
            '                oPRODTREE_LINES.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Backflush
            '            Else
            '                oPRODTREE_LINES.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Manual
            '            End If
            '            oRSet.MoveNext()
            '            oPRODTREE_LINES.SetCurrentLine(J - 1)
            '            oPRODTREE_LINES.Add()
            '            Counter = J
            '        Next
            '        RS.DoQuery("Select Distinct B.u_sfgcode,B.u_sfgqty From [@GEN_PROD_PRCS] A INNER JOIN [@GEN_PROD_PRCS_D0] B ON A.Code = B.Code Where A.u_itemcode = '" + ItemCode + "' And B.u_itemcode = '" + Trim(oRecordSet.Fields.Item("u_itemcode").Value) + "' And (B.u_sfgcode <> '' Or B.u_sfgcode is not null) And A.u_itemcode = '" + ItemCode + "'")
            '        For L As Integer = 1 To RS.RecordCount
            '            If Trim(RS.Fields.Item("u_sfgcode").Value) <> "" Then
            '                oPRODTREE_LINES.ParentItem = oRecordSet.Fields.Item("u_itemcode").Value
            '                oPRODTREE_LINES.ItemCode = RS.Fields.Item("u_sfgcode").Value
            '                oPRODTREE_LINES.Quantity = RS.Fields.Item("u_sfgqty").Value
            '                oPRODTREE_LINES.Warehouse = ToWhs
            '                oPRODTREE_LINES.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Backflush
            '                oPRODTREE_LINES.SetCurrentLine(Counter)
            '                Counter = Counter + 1
            '                oPRODTREE_LINES.Add()
            '            End If
            '            RS.MoveNext()
            '        Next
            '        oCompany.StartTransaction()
            '        Dim Err As Integer = oPRODTREE.Add()
            '        If Err <> 0 Then
            '            oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            '            Return False
            '        Else
            '            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            '        End If
            '        oRecordSet.MoveNext()
            '    Next
            '    oRecordSet.DoQuery("Select Distinct B.u_itemcode,B.u_process From [@GEN_PROD_PRCS] A INNER JOIN [@GEN_PROD_PRCS_D0] B ON A.Code = B.Code Where A.u_itemcode = '" + ItemCode + "' And B.u_process = 'Finishing'")
            '    RS1.DoQuery("Select B.u_inwhs From [@GEN_UNIT_MST] A INNER JOIN [@GEN_UNIT_MST_D0] B ON A.CODE = B.CODE Where A.Name = '" + Trim(objForm.Items.Item("unit").Specific.value) + "' And B.u_process = '" + Trim(oRecordSet.Fields.Item("u_process").Value) + "'")
            '    ToWhs = RS1.Fields.Item("u_inwhs").Value
            '    oPRODTREE = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees)
            '    oPRODTREE.TreeCode = ItemCode
            '    oPRODTREE.TreeType = SAPbobsCOM.BoItemTreeTypes.iProductionTree
            '    oPRODTREE.Quantity = 1
            '    oPRODTREE_LINES = oPRODTREE.Items
            '    For N As Integer = 1 To oRecordSet.RecordCount
            '        oPRODTREE_LINES.ParentItem = ItemCode
            '        oPRODTREE_LINES.ItemCode = oRecordSet.Fields.Item("u_itemcode").Value
            '        oPRODTREE_LINES.Quantity = 1
            '        oPRODTREE_LINES.Warehouse = ToWhs
            '        oPRODTREE_LINES.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Manual
            '        oRecordSet.MoveNext()
            '        oPRODTREE_LINES.SetCurrentLine(N - 1)
            '        oPRODTREE_LINES.Add()
            '    Next
            '    oCompany.StartTransaction()
            '    Dim ErrFlg As Integer = oPRODTREE.Add()
            '    If ErrFlg <> 0 Then
            '        oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            '        Return False
            '    Else
            '        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            '    End If
            '    oRecordSet.DoQuery("Delete From TMP_CST_BOM Where MacId = '" + MAC_ID + "'")
            '    'oDBs_Head.SetValue("u_status", 0, "ACTIVE")
            '    For l As Integer = 1 To objMatrix.VisualRowCount
            '        If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(l).Specific.value) <> "" Then
            '            If Trim(objMatrix.Columns.Item("status").Cells.Item(l).Specific.Selected.Value) = "NEW" Then
            '                objMatrix.Columns.Item("status").Cells.Item(l).Specific.Select("ACTIVE", SAPbouiCOM.BoSearchKey.psk_ByValue)
            '            End If
            '        End If
            '    Next
            '    objForm.Items.Item("status").Specific.Select("ACTIVE", SAPbouiCOM.BoSearchKey.psk_ByValue)
            'End If
            'oPRODTREE = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees)
            'If Trim(oDBs_Head.GetValue("u_status", 0)) = "CHANGE" Then
            '    oRSet.DoQuery("Select Distinct SONO,DOCNUM,ItemCode,SOQty,Unit,ChldItem,Qty,Process,IssMthd,Status,MacId,Deleted From TMP_CST_BOM Where ItemCode = '" + ItemCode + "' Order by Status")
            '    For i As Integer = 1 To oRSet.RecordCount
            '        If Trim(oRSet.Fields.Item("Status").Value) = "CHANGE" Then
            '            oRecordSet.DoQuery("Select Distinct B.u_itemcode From [@GEN_PROD_PRCS] A INNER JOIN [@GEN_PROD_PRCS_D0] B ON A.CODE = B.CODE Where A.u_itemcode = '" + ItemCode + "' And B.u_process = '" + Trim(oRSet.Fields.Item("Process").Value) + "'")
            '            oPRODTREE.GetByKey(Trim(oRecordSet.Fields.Item("u_itemcode").Value))
            '            oPRODTREE_LINES = oPRODTREE.Items
            '            RS.DoQuery("Select Distinct ChildNum From ITT1 Where Father = '" + Trim(oRecordSet.Fields.Item("u_itemcode").Value) + "' And Code = '" + Trim(oRSet.Fields.Item("ChldItem").Value) + "'")
            '            oPRODTREE_LINES.SetCurrentLine(Trim(RS.Fields.Item("ChildNum").Value))
            '            oPRODTREE_LINES.Quantity = oRSet.Fields.Item("Qty").Value
            '            If Trim(oRSet.Fields.Item("IssMthd").Value) = "B" Then
            '                oPRODTREE_LINES.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Backflush
            '            Else
            '                oPRODTREE_LINES.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Manual
            '            End If
            '            oCompany.StartTransaction()
            '            ErrUpd = oPRODTREE.Update
            '            If ErrUpd <> 0 Then
            '                oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            '                Return False
            '            Else
            '                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            '            End If
            '        End If
            '        If Trim(oRSet.Fields.Item("Status").Value) = "NEW" Then
            '            oRecordSet.DoQuery("Select Distinct B.u_itemcode,B.u_process From [@GEN_PROD_PRCS] A INNER JOIN [@GEN_PROD_PRCS_D0] B ON A.CODE = B.CODE Where A.u_itemcode = '" + ItemCode + "' And B.u_process = '" + Trim(oRSet.Fields.Item("Process").Value) + "'")
            '            RS1.DoQuery("Select B.u_inwhs From [@GEN_UNIT_MST] A INNER JOIN [@GEN_UNIT_MST_D0] B ON A.CODE = B.CODE Where A.Name = '" + Trim(objForm.Items.Item("unit").Specific.value) + "' And B.u_process = '" + Trim(oRecordSet.Fields.Item("u_process").Value) + "'")
            '            ToWhs = RS1.Fields.Item("u_inwhs").Value
            '            oPRODTREE.GetByKey(Trim(oRecordSet.Fields.Item("u_itemcode").Value))
            '            oPRODTREE_LINES = oPRODTREE.Items
            '            RS.DoQuery("Select Distinct Max(ChildNum) As 'ChildNum' From ITT1 Where Father = '" + Trim(oRecordSet.Fields.Item("u_itemcode").Value) + "' And Code = '" + Trim(oRSet.Fields.Item("ChldItem").Value) + "'")
            '            oPRODTREE_LINES.SetCurrentLine(oPRODTREE.Items.Count - 1)
            '            oPRODTREE_LINES.Add()
            '            oPRODTREE_LINES.ParentItem = oRecordSet.Fields.Item("u_itemcode").Value
            '            oPRODTREE_LINES.ItemCode = oRSet.Fields.Item("ChldItem").Value
            '            oPRODTREE_LINES.Quantity = oRSet.Fields.Item("Qty").Value
            '            oPRODTREE_LINES.Warehouse = ToWhs
            '            If Trim(oRSet.Fields.Item("IssMthd").Value) = "B" Then
            '                oPRODTREE_LINES.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Backflush
            '            Else
            '                oPRODTREE_LINES.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Manual
            '            End If
            '            oCompany.StartTransaction()
            '            ErrAdd = oPRODTREE.Update
            '            If ErrAdd <> 0 Then
            '                oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            '                Return False
            '            Else
            '                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            '            End If
            '        End If
            '        If Trim(oRSet.Fields.Item("Status").Value) = "DELETE" And Trim(oRSet.Fields.Item("Deleted").Value) = "NO" Then
            '            oRecordSet.DoQuery("Select Distinct B.u_itemcode From [@GEN_PROD_PRCS] A INNER JOIN [@GEN_PROD_PRCS_D0] B ON A.CODE = B.CODE Where A.u_itemcode = '" + ItemCode + "' And B.u_process = '" + Trim(oRSet.Fields.Item("Process").Value) + "'")
            '            oPRODTREE.GetByKey(Trim(oRecordSet.Fields.Item("u_itemcode").Value))
            '            oPRODTREE_LINES = oPRODTREE.Items
            '            RS.DoQuery("Select Distinct ChildNum As 'ChildNum' From ITT1 Where Father = '" + Trim(oRecordSet.Fields.Item("u_itemcode").Value) + "' And Code = '" + Trim(oRSet.Fields.Item("ChldItem").Value) + "'")
            '            oPRODTREE_LINES.SetCurrentLine(Trim(RS.Fields.Item("ChildNum").Value))
            '            oPRODTREE_LINES.Delete()
            '            oCompany.StartTransaction()
            '            ErrDel = oPRODTREE.Update
            '            If ErrDel <> 0 Then
            '                oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            '                Return False
            '            Else
            '                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            '            End If
            '        End If
            '        oRSet.MoveNext()
            '    Next
            '    For l As Integer = 1 To objMatrix.VisualRowCount
            '        If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(l).Specific.value) <> "" Then
            '            If Trim(objMatrix.Columns.Item("status").Cells.Item(l).Specific.Selected.Value) = "CHANGE" And ErrUpd = 0 Then
            '                objMatrix.Columns.Item("status").Cells.Item(l).Specific.Select("ACTIVE", SAPbouiCOM.BoSearchKey.psk_ByValue)
            '            End If
            '            If Trim(objMatrix.Columns.Item("status").Cells.Item(l).Specific.Selected.Value) = "NEW" And ErrAdd = 0 Then
            '                objMatrix.Columns.Item("status").Cells.Item(l).Specific.Select("ACTIVE", SAPbouiCOM.BoSearchKey.psk_ByValue)
            '            End If
            '            If Trim(objMatrix.Columns.Item("status").Cells.Item(l).Specific.Selected.Value) = "DELETE" And ErrDel = 0 Then
            '                objMatrix.Columns.Item("status").Cells.Item(l).Specific.Select("ACTIVE", SAPbouiCOM.BoSearchKey.psk_ByValue)
            '                objMatrix.Columns.Item("deleted").Cells.Item(l).Specific.Select("YES", SAPbouiCOM.BoSearchKey.psk_ByValue)
            '            End If
            '        End If
            '    Next
            '    oRecordSet.DoQuery("Delete From TMP_CST_BOM Where MacId = '" + MAC_ID + "'")
            '    objForm.Items.Item("status").Specific.Select("ACTIVE", SAPbouiCOM.BoSearchKey.psk_ByValue)
            'End If
            Return True
        Catch ex As Exception
            oRecordSet.DoQuery("Delete From TMP_CST_BOM Where MacId = '" + MAC_ID + "'")
            oApplication.StatusBar.SetText(ex.Message)
            Return False
        End Try
    End Function

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    'If pVal.ItemUID = "itemcode" And pVal.BeforeAction = True Then
                    '    Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '    RS.DoQuery("Select DocNum From [@GEN_CUST_BOM] Where u_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' And Docnum <> '" + Trim(objForm.Items.Item("docnum").Specific.value) + "'")
                    '    If RS.RecordCount > 0 Then
                    '        oApplication.StatusBar.SetText("BOM already created for this FG", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        BubbleEvent = False
                    '        Exit Sub
                    '    End If
                    'End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If pVal.ItemUID = "mtx" And pVal.ColUID = "per" And pVal.BeforeAction = False And pVal.Row > 0 Then
                        objMatrix = objForm.Items.Item("mtx").Specific
                        If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.value) <> "" Then
                            If objMatrix.Columns.Item("per").Cells.Item(pVal.Row).Specific.value >= 0 Then
                                objMatrix.Columns.Item("totqty").Cells.Item(pVal.Row).Specific.value = objMatrix.Columns.Item("ordrqty").Cells.Item(pVal.Row).Specific.value + (objMatrix.Columns.Item("ordrqty").Cells.Item(pVal.Row).Specific.value * objMatrix.Columns.Item("per").Cells.Item(pVal.Row).Specific.value) / 100
                            End If
                        End If
                    End If
                    If (pVal.ItemUID = "mtx" And pVal.ColUID = "qty") Or (pVal.ItemUID = "qty") And pVal.BeforeAction = False And pVal.Row > 0 Then
                        objMatrix = objForm.Items.Item("mtx").Specific
                        If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.value) <> "" Then
                            If objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value > 0 Then
                                objMatrix.Columns.Item("ordrqty").Cells.Item(pVal.Row).Specific.value = objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value * objForm.Items.Item("qty").Specific.value
                                objMatrix.Columns.Item("totqty").Cells.Item(pVal.Row).Specific.value = objMatrix.Columns.Item("ordrqty").Cells.Item(pVal.Row).Specific.value + (objMatrix.Columns.Item("ordrqty").Cells.Item(pVal.Row).Specific.value * objMatrix.Columns.Item("per").Cells.Item(pVal.Row).Specific.value) / 100
                            End If
                        End If
                        objMatrix.Columns.Item("process").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                    'If pVal.ItemUID = "mtx" And pVal.ColUID = "ordrqty" And pVal.BeforeAction = False Then
                    '    objMatrix = objForm.Items.Item("mtx").Specific
                    '    If objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value <= 0 Then
                    '        objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value = objMatrix.Columns.Item("ordrqty").Cells.Item(pVal.Row).Specific.value / objForm.Items.Item("qty").Specific.value
                    '    End If
                    'End If
                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    'If (pVal.CharPressed <> 9 And pVal.CharPressed <> 13) And pVal.ColUID = "ordrqty" And pVal.BeforeAction = True And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    '    BubbleEvent = False
                    'End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "btn" And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        objMatrix = objForm.Items.Item("mtx").Specific
                        For i As Integer = 1 To objMatrix.VisualRowCount
                            If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) <> "" And objMatrix.Columns.Item("qty").Cells.Item(i).Specific.value <= 0 Then
                                objMatrix.Columns.Item("qty").Cells.Item(i).Specific.value = objMatrix.Columns.Item("ordrqty").Cells.Item(i).Specific.value / objForm.Items.Item("qty").Specific.value
                            End If
                        Next
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            UpdMode = True
                        Else
                            UpdMode = False
                        End If
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
                    If pVal.BeforeAction = True Then
                        If oCFL.UniqueID = "ITCFL" Then
                            Me.FilterSOItems(FormUID)
                        End If
                    Else
                        If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_CUST_BOM")
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_CUST_BOM_D0")
                            If oCFL.UniqueID = "SOCFL" Then
                                oDBs_Head.SetValue("u_sono", 0, oDT.GetValue("DocNum", 0))
                                oDBs_Head.SetValue("u_soref", 0, oDT.GetValue("NumAtCard", 0))
                                'objMatrix = objForm.Items.Item("mtx").Specific
                                'objMatrix.Clear()
                            End If
                            If oCFL.UniqueID = "ITCFL" Then
                                oDBs_Head.SetValue("u_itemcode", 0, oDT.GetValue("ItemCode", 0))
                                oDBs_Head.SetValue("u_itemname", 0, oDT.GetValue("ItemName", 0))
                                oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("Select Sum(B.Quantity) As 'Qty' From ORDR A Inner Join RDR1 B On A.DocEntry = B.DocEntry And A.DocNum = '" + Trim(objForm.Items.Item("sono").Specific.value) + "' And B.ItemCode = '" + oDT.GetValue("ItemCode", 0) + "'")
                                oDBs_Head.SetValue("u_qty", 0, oRecordSet.Fields.Item("Qty").Value)
                                objMatrix = objForm.Items.Item("mtx").Specific
                                If oApplication.MessageBox("Do you want refresh the row items.", 2, "Yes", "No") = "1" Then
                                    objMatrix.Clear()
                                    objMatrix.AddRow(1)
                                    Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
                                End If
                            End If
                            If oCFL.UniqueID = "UNITCFL" Then
                                oDBs_Head.SetValue("u_unit", 0, oDT.GetValue("Name", 0))
                            End If
                            If oCFL.UniqueID = "RITCFL" Then
                                objMatrix = objForm.Items.Item("mtx").Specific
                                oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
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
                                    oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                                    oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, oDT.GetValue("ItemCode", i))
                                    oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, oDT.GetValue("ItemName", i))
                                    oDBs_Detail.SetValue("u_size", oDBs_Detail.Offset, objMatrix.Columns.Item("size").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_qty", oDBs_Detail.Offset, objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_ordrqty", oDBs_Detail.Offset, objMatrix.Columns.Item("ordrqty").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_per", oDBs_Detail.Offset, objMatrix.Columns.Item("per").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_totqty", oDBs_Detail.Offset, objMatrix.Columns.Item("totqty").Cells.Item(pVal.Row).Specific.value)
                                    oRSet.DoQuery("Select Invntryuom From OITM Where ItemCode = '" + Trim(oDT.GetValue("ItemCode", i)) + "'")
                                    oDBs_Detail.SetValue("u_uom", oDBs_Detail.Offset, oRSet.Fields.Item("InvntryUOM").Value)
                                    oDBs_Detail.SetValue("u_process", oDBs_Detail.Offset, objMatrix.Columns.Item("process").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_issmthd", oDBs_Detail.Offset, oDT.GetValue("IssueMthd", i))
                                    ''oDBs_Detail.SetValue("u_issmthd", oDBs_Detail.Offset, objMatrix.Columns.Item("issmthd").Cells.Item(pVal.Row).Specific.Selected.Value)
                                    oDBs_Detail.SetValue("u_cardcode", oDBs_Detail.Offset, objMatrix.Columns.Item("cardcode").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_cardname", oDBs_Detail.Offset, objMatrix.Columns.Item("cardname").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_status", oDBs_Detail.Offset, "NEW")
                                    oDBs_Detail.SetValue("u_deleted", oDBs_Detail.Offset, "NO")
                                    oDBs_Detail.SetValue("u_remarks", oDBs_Detail.Offset, objMatrix.Columns.Item("remarks").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_place", oDBs_Detail.Offset, objMatrix.Columns.Item("place").Cells.Item(pVal.Row).Specific.value)
                                    'Vijeesh
                                    oDBs_Detail.SetValue("U_fremark", oDBs_Detail.Offset, objMatrix.Columns.Item("fremark").Cells.Item(pVal.Row).Specific.value)

                                    objMatrix.SetLineData(pVal.Row + i)
                                    objForm.EnableMenu("1293", True)
                                Next
                                If Flag = True Then
                                    objMatrix.AddRow(1, objMatrix.VisualRowCount)
                                    Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
                                End If
                            End If
                            If oCFL.UniqueID = "PRCCFL" Then
                                objMatrix = objForm.Items.Item("mtx").Specific
                                oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                For i As Integer = 0 To oDT.Rows.Count - 1
                                    Dim cflSelectedcount As Integer = oDT.Rows.Count
                                    oDBs_Detail.Offset = pVal.Row - 1 + i
                                    oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                                    oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, objMatrix.Columns.Item("itemname").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_size", oDBs_Detail.Offset, objMatrix.Columns.Item("size").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_qty", oDBs_Detail.Offset, objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_ordrqty", oDBs_Detail.Offset, objMatrix.Columns.Item("ordrqty").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_per", oDBs_Detail.Offset, objMatrix.Columns.Item("per").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_totqty", oDBs_Detail.Offset, objMatrix.Columns.Item("totqty").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_uom", oDBs_Detail.Offset, objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_process", oDBs_Detail.Offset, oDT.GetValue("Name", i))
                                    oDBs_Detail.SetValue("u_issmthd", oDBs_Detail.Offset, objMatrix.Columns.Item("issmthd").Cells.Item(pVal.Row).Specific.Selected.Value)
                                    oDBs_Detail.SetValue("u_cardcode", oDBs_Detail.Offset, objMatrix.Columns.Item("cardcode").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_cardname", oDBs_Detail.Offset, objMatrix.Columns.Item("cardname").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_remarks", oDBs_Detail.Offset, objMatrix.Columns.Item("remarks").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_place", oDBs_Detail.Offset, objMatrix.Columns.Item("place").Cells.Item(pVal.Row).Specific.value)
                                    'Vijeesh
                                    oDBs_Detail.SetValue("U_fremark", oDBs_Detail.Offset, objMatrix.Columns.Item("fremark").Cells.Item(pVal.Row).Specific.value)

                                    oDBs_Detail.SetValue("u_deleted", oDBs_Detail.Offset, objMatrix.Columns.Item("deleted").Cells.Item(pVal.Row).Specific.selected.value)
                                    oDBs_Detail.SetValue("u_status", oDBs_Detail.Offset, objMatrix.Columns.Item("status").Cells.Item(pVal.Row).Specific.selected.value)
                                    objMatrix.SetLineData(pVal.Row + i)
                                    objForm.EnableMenu("1293", True)
                                Next
                            End If
                            If oCFL.UniqueID = "VENDCFL" Then
                                objMatrix = objForm.Items.Item("mtx").Specific
                                oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                For i As Integer = 0 To oDT.Rows.Count - 1
                                    Dim cflSelectedcount As Integer = oDT.Rows.Count
                                    oDBs_Detail.Offset = pVal.Row - 1 + i
                                    oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                                    oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, objMatrix.Columns.Item("itemname").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_size", oDBs_Detail.Offset, objMatrix.Columns.Item("size").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_qty", oDBs_Detail.Offset, objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_ordrqty", oDBs_Detail.Offset, objMatrix.Columns.Item("ordrqty").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_per", oDBs_Detail.Offset, objMatrix.Columns.Item("per").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_totqty", oDBs_Detail.Offset, objMatrix.Columns.Item("totqty").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_uom", oDBs_Detail.Offset, objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_process", oDBs_Detail.Offset, objMatrix.Columns.Item("process").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_issmthd", oDBs_Detail.Offset, objMatrix.Columns.Item("issmthd").Cells.Item(pVal.Row).Specific.Selected.Value)
                                    oDBs_Detail.SetValue("u_cardcode", oDBs_Detail.Offset, oDT.GetValue("CardCode", i))
                                    oDBs_Detail.SetValue("u_cardname", oDBs_Detail.Offset, oDT.GetValue("CardName", i))
                                    oDBs_Detail.SetValue("u_remarks", oDBs_Detail.Offset, objMatrix.Columns.Item("remarks").Cells.Item(pVal.Row).Specific.value)
                                    oDBs_Detail.SetValue("u_place", oDBs_Detail.Offset, objMatrix.Columns.Item("place").Cells.Item(pVal.Row).Specific.value)
                                    'Vijeesh
                                    oDBs_Detail.SetValue("U_fremark", oDBs_Detail.Offset, objMatrix.Columns.Item("fremark").Cells.Item(pVal.Row).Specific.value)

                                    oDBs_Detail.SetValue("u_deleted", oDBs_Detail.Offset, objMatrix.Columns.Item("deleted").Cells.Item(pVal.Row).Specific.selected.value)
                                    oDBs_Detail.SetValue("u_status", oDBs_Detail.Offset, objMatrix.Columns.Item("status").Cells.Item(pVal.Row).Specific.selected.value)
                                    objMatrix.SetLineData(pVal.Row + i)
                                    objForm.EnableMenu("1293", True)
                                Next
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
                    Case "GEN_CUST_BOM"
                        Me.CreateForm(objForm.UniqueID)
                    Case "1282"
                        If objForm.TypeEx = "GEN_CUST_BOM" Then
                            Me.SetDefault(objForm.UniqueID)
                            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        End If
                    Case "1281"
                        If objForm.TypeEx = "GEN_CUST_BOM" Then
                            objForm.EnableMenu("1282", True)
                            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        End If
                    Case "1288", "1289", "1290", "1291"
                        If objForm.TypeEx = "GEN_CUST_BOM" Then
                            objForm.EnableMenu("1282", True)
                        End If
                    Case "1287"
                        If objForm.TypeEx = "GEN_CUST_BOM" Then
                            Me.SetDefault(objForm.UniqueID)
                            objMatrix = objForm.Items.Item("mtx").Specific
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_CUST_BOM")
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_CUST_BOM_D0")

                            oDBs_Head.SetValue("U_itemcode", 0, "")
                            oDBs_Head.SetValue("U_itemname", 0, "")
                            oDBs_Head.SetValue("U_qty", 0, 1)
                            For Row As Integer = 1 To objMatrix.RowCount - 1
                                objMatrix.GetLineData(Row)
                                oDBs_Detail.Offset = Row - 1
                                If oDBs_Detail.GetValue("u_itemcode", oDBs_Detail.Offset) <> "" Then
                                    oDBs_Detail.SetValue("U_status", oDBs_Detail.Offset, "NEW")
                                    objMatrix.SetLineData(Row)
                                End If
                            Next
                        End If
                    Case "1293"
                        If objForm.TypeEx = "GEN_CUST_BOM" Then
                            If ITEM_ID.Equals("mtx") = True Then
                                objMatrix = objForm.Items.Item("mtx").Specific
                                oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_CUST_BOM")
                                oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_CUST_BOM_D0")
                                For Row As Integer = 1 To objMatrix.VisualRowCount
                                    objMatrix.GetLineData(Row)
                                    oDBs_Detail.Offset = Row - 1
                                    oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, Row)
                                    oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, objMatrix.Columns.Item("itemcode").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, objMatrix.Columns.Item("itemname").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_size", oDBs_Detail.Offset, objMatrix.Columns.Item("size").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_qty", oDBs_Detail.Offset, objMatrix.Columns.Item("qty").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_ordrqty", oDBs_Detail.Offset, objMatrix.Columns.Item("ordrqty").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_per", oDBs_Detail.Offset, objMatrix.Columns.Item("per").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_totqty", oDBs_Detail.Offset, objMatrix.Columns.Item("totqty").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_uom", oDBs_Detail.Offset, objMatrix.Columns.Item("uom").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_process", oDBs_Detail.Offset, objMatrix.Columns.Item("process").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_cardcode", oDBs_Detail.Offset, objMatrix.Columns.Item("cardcode").Cells.Item(Row).Specific.Value)
                                    oDBs_Detail.SetValue("u_cardname", oDBs_Detail.Offset, objMatrix.Columns.Item("cardname").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_remarks", oDBs_Detail.Offset, objMatrix.Columns.Item("remarks").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("u_place", oDBs_Detail.Offset, objMatrix.Columns.Item("place").Cells.Item(Row).Specific.value)
                                    'Vijeesh
                                    oDBs_Detail.SetValue("U_fremark", oDBs_Detail.Offset, objMatrix.Columns.Item("fremark").Cells.Item(Row).Specific.value)

                                    oDBs_Detail.SetValue("u_issmthd", oDBs_Detail.Offset, objMatrix.Columns.Item("issmthd").Cells.Item(Row).Specific.selected.value)
                                    ' oDBs_Detail.SetValue("u_status", oDBs_Detail.Offset, objMatrix.Columns.Item("status").Cells.Item(Row).Specific.Selected.Value)
                                    objMatrix.SetLineData(Row)
                                Next
                                objMatrix.FlushToDataSource()
                                oDBs_Detail.RemoveRecord(oDBs_Detail.Size - 1)
                                objMatrix.LoadFromDataSource()
                            End If
                        End If
                End Select
            ElseIf pVal.BeforeAction = True Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "519"
                        Try
                            If objForm.TypeEx = "GEN_CUST_BOM" Then
                                BubbleEvent = False
                                sDocNum = objForm.Items.Item("docnum").Specific.Value
                                sRptName = "CustomBOM.rpt"
                                Me.Report1()
                            End If
                        Catch ex As Exception

                        End Try
                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            RowCount = eventInfo.Row
            Dim oForm As SAPbouiCOM.Form = oApplication.Forms.Item(eventInfo.FormUID)
            If eventInfo.Row > 0 Then
                ITEM_ID = eventInfo.ItemUID
                objMatrix = oForm.Items.Item("mtx").Specific
                If objMatrix.VisualRowCount > 1 Then
                    oForm.EnableMenu("1293", True)
                Else
                    oForm.EnableMenu("1293", False)
                End If
            Else
                ITEM_ID = ""
            End If
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                oForm.EnableMenu("1283", False)
                'eventInfo.RemoveFromContent("1283")
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

    Private Sub Report1()
        Dim oThread As New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Report1Thread))
        oThread.SetApartmentState(System.Threading.ApartmentState.STA)
        oThread.Start()
    End Sub

    Private Sub Report1Thread()
        Try
            Dim oCRForm As New Crystal_Form
            oCRForm.ShowDialog()
        Catch ex As Exception
            oApplication.MessageBox(ex.Message.ToString)
        End Try
    End Sub

End Class
