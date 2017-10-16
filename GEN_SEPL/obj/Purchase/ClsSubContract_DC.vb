'//Created by PRIYA on 01/06/2011

Imports System.Threading
Imports System.Globalization
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text

Public Class ClsSubContract_DC

#Region "        Declaration        "

    Dim oUtilities As New ClsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oDBs_Head1 As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail1 As SAPbouiCOM.DBDataSource
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim docentry As Integer
    Dim ROW_ID As Integer = 0
#End Region


    Sub CreateForm()
        Try
            oUtilities.SAPXML("SubContractingDC.xml")
            objForm = oApplication.Forms.GetForm("GEN_SCDC", oApplication.Forms.ActiveForm.TypeCount)
            objForm.Items.Item("cardcode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("c_series").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("t_docdt").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("t_docno").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Issue").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC_D0")

            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Select SlpCode,SlpName from OSLP")
            objCombo = objForm.Items.Item("Buyer").Specific
            For i As Integer = 1 To oRS.RecordCount
                objCombo.ValidValues.Add(Trim(oRS.Fields.Item("SlpCode").Value), Trim(oRS.Fields.Item("SlpName").Value))
                oRS.MoveNext()
            Next
            Me.FilterWarehouse(objForm.UniqueID)
            objForm.Select()
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
            objForm.Items.Item("Issue").Enabled = False
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC_D0")
            oUtilities.GetSeries(FormUID, "c_series", "GEN_SC_DC")
            oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("c_series").Specific.Selected.Value, "GEN_SC_DC"))
            oDBs_Head.SetValue("U_Status", 0, "Open")
            oDBs_Head.SetValue("U_DcDat", 0, DateTime.Today.ToString("yyyyMMdd"))
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oRS.RecordCount > 0 Then
                oRS.DoQuery("Select ISNULL(T0.firstName,'') + ' ' + ISNULL(T0.lastName,'') Owner,empid from OHEM T0 INNER JOIN OUSR T1 ON T0.UserID=T1.USERID Where USER_CODE='" & oCompany.UserName & "'")
                oDBs_Head.SetValue("U_Owner", 0, Trim(oRS.Fields.Item("Owner").Value))
                oDBs_Head.SetValue("U_OwnerCod", 0, Trim(oRS.Fields.Item("empid").Value))
            End If
            oDBs_Head.SetValue("U_Qty", 0, 1)
            objCombo = objForm.Items.Item("Buyer").Specific
            If objCombo.ValidValues.Count > 0 Then objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            'objMatrix = objForm.Items.Item("ItemMatrix").Specific
            'objMatrix.Clear()
            'objMatrix.AddRow()
            'objMatrix.FlushToDataSource()
            'Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
            objForm.Items.Item("cardcode").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            objForm.Freeze(False)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC_D0")
            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
            oDBs_Detail.SetValue("U_IsCheck", oDBs_Detail.Offset, "N")
            oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_Qty", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_CompQty", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_FWhs", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, "")
            objMatrix.SetLineData(Row)
            objMatrix.FlushToDataSource()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "c_series" And pVal.BeforeAction = False Then
                        oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("c_series").Specific.Selected.Value, "GEN_SC_DC"))
                    End If
                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    If pVal.ItemUID = "t_Qty" And pVal.BeforeAction = False And pVal.CharPressed = 9 And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        If objMatrix.VisualRowCount > 0 Then
                            objMatrix.Columns.Item("itemno").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Exit Sub
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation(FormUID) = False Then BubbleEvent = False
                    ElseIf pVal.ItemUID = "1" And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(FormUID)
                    End If
                    If pVal.ItemUID = "Issue" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        If pVal.BeforeAction = True Then
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select B.u_ItemNo , B.u_FWhs , B.u_TWhs , IsNull(B.u_IssQty,0) - IsNull(B.u_CompQty,0) AS 'OpenQty' From [@GEN_SC_DC] A Inner Join [@GEN_SC_DC_D0] B On A.DocEntry = B.DocEntry Where B.u_Fwhs in (Select u_whs From [@GEN_WHS_USR] Where u_User = '" + oCompany.UserName.ToString + "') And B.u_IssQty > B.u_CompQty And A.DocNum = '" + Trim(objForm.Items.Item("t_docno").Specific.value) + "'")
                            If oRSet.RecordCount = 0 Then
                                BubbleEvent = False
                            End If
                        Else
                            Dim oRSWHcount As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSWHcount.DoQuery("Select Distinct B.u_Fwhs From [@GEN_SC_DC] A Inner Join [@GEN_SC_DC_D0] B On A.DocEntry = B.DocEntry Where A.DocNum = '" + Trim(objForm.Items.Item("t_docno").Specific.value) + "' And B.u_Fwhs in (Select Distinct u_whs From [@GEN_WHS_USR] Where u_user = '" + oCompany.UserName.Trim + "') Group By B.u_Fwhs")
                            For cnt As Integer = 1 To oRSWHcount.RecordCount
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet.DoQuery("Select A.DocNum,B.LineID,B.u_ItemNo,B.u_FWhs,B.u_TWhs,IsNull(B.u_IssQty,0) - IsNull(B.u_CompQty,0) AS 'OpenQty' From [@GEN_SC_DC] A Inner Join [@GEN_SC_DC_D0] B On A.DocEntry = B.DocEntry Where B.u_Fwhs = '" + Trim(oRSWHcount.Fields.Item("u_Fwhs").Value) + "' And B.u_IssQty > B.u_CompQty And A.DocNum = '" + Trim(objForm.Items.Item("t_docno").Specific.value) + "'")
                                oApplication.ActivateMenuItem("3080")
                                Dim ITForm As SAPbouiCOM.Form
                                Dim ITMatrix As SAPbouiCOM.Matrix
                                ITForm = oApplication.Forms.ActiveForm
                                ITMatrix = ITForm.Items.Item("23").Specific
                                ITForm.Items.Item("18").Specific.value = oRSet.Fields.Item("u_FWhs").Value
                                ITForm.Items.Item("scpono").Specific.value = oRSet.Fields.Item("DocNum").Value
                                ITForm.Items.Item("scpotp").Specific.value = "SubCon_DC"
                                ITMatrix.Columns.Item("U_scpono").Editable = True
                                ITMatrix.Columns.Item("U_scpoln").Editable = True
                                For i As Integer = 1 To oRSet.RecordCount
                                    Try
                                        ITMatrix.Columns.Item("1").Cells.Item(i).Specific.value = oRSet.Fields.Item("u_ItemNo").Value
                                        ITMatrix.Columns.Item("5").Cells.Item(i).Specific.value = oRSet.Fields.Item("u_TWhs").Value
                                        ITMatrix.Columns.Item("10").Cells.Item(i).Specific.value = oRSet.Fields.Item("OpenQty").Value
                                        ITMatrix.Columns.Item("U_scpono").Cells.Item(i).Specific.value = oRSet.Fields.Item("DocNum").Value
                                        ITMatrix.Columns.Item("U_scpoln").Cells.Item(i).Specific.value = oRSet.Fields.Item("LineId").Value
                                        ITMatrix.Columns.Item("10").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    Catch ex As Exception
                                        oApplication.StatusBar.SetText(ex.Message)
                                    End Try
                                    oRSet.MoveNext()
                                Next
                                ITMatrix.Columns.Item("U_scpono").Editable = False
                                ITMatrix.Columns.Item("U_scpoln").Editable = False
                                oRSWHcount.MoveNext()
                            Next
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If pVal.ItemUID = "t_docdt" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Trim(objForm.Items.Item("t_docdt").Specific.Value).Equals("") = False Then
                            If (DateDiff(DateInterval.Day, DateTime.ParseExact(Trim(objForm.Items.Item("t_docdt").Specific.Value), "yyyyMMdd", Nothing), DateTime.Today)) <> 0 Then
                                oApplication.StatusBar.SetText("Document date varies from system date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        End If
                    ElseIf pVal.ItemUID = "t_deldt" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Trim(objForm.Items.Item("t_deldt").Specific.Value).Equals("") = False Then
                            If (DateDiff(DateInterval.Day, DateTime.ParseExact(Trim(objForm.Items.Item("t_docdt").Specific.Value), "yyyyMMdd", Nothing), DateTime.ParseExact(Trim(objForm.Items.Item("t_deldt").Specific.Value), "yyyyMMdd", Nothing))) < 0 Then
                                oApplication.StatusBar.SetText("Delivery date is before Document date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If pVal.ItemUID = "t_Qty" And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Try
                            objForm = oApplication.Forms.Item(FormUID)
                            objForm.Freeze(True)
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC")
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC_D0")
                            For Row As Integer = 1 To objMatrix.VisualRowCount
                                oDBs_Detail.Offset = Row - 1
                                oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
                                If objMatrix.Columns.Item("IsCheck").Cells.Item(Row).Specific.Checked = True Then
                                    oDBs_Detail.SetValue("U_IsCheck", oDBs_Detail.Offset, "Y")
                                Else
                                    oDBs_Detail.SetValue("U_IsCheck", oDBs_Detail.Offset, "N")
                                End If
                                oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemno").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("desc").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_fwhs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("fwhs").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("stock").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("twhs").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("BOMQty").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Qty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Qty").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("BOMQty").Cells.Item(Row).Specific.Value) * CDbl(objForm.Items.Item("t_Qty").Specific.Value))
                                oDBs_Detail.SetValue("U_CompQty", oDBs_Detail.Offset, 0)
                                oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("price").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("price").Cells.Item(Row).Specific.Value) * CDbl(objMatrix.Columns.Item("BOMQty").Cells.Item(Row).Specific.Value) * CDbl(objForm.Items.Item("t_Qty").Specific.Value))
                                oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("rmrk").Cells.Item(Row).Specific.Value))
                                objMatrix.SetLineData(Row)
                            Next
                            Me.CalculateTotal(FormUID)
                            objForm.Freeze(False)
                        Catch ex As Exception
                            oApplication.StatusBar.SetText(ex.Message)
                            objForm.Freeze(False)
                        End Try
                    ElseIf pVal.ItemUID = "ItemMatrix" And (pVal.ColUID = "issqty" Or pVal.ColUID = "price") And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        oDBs_Detail.Offset = pVal.Row - 1
                        oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                        If objMatrix.Columns.Item("IsCheck").Cells.Item(pVal.Row).Specific.Checked = True Then
                            oDBs_Detail.SetValue("U_IsCheck", oDBs_Detail.Offset, "Y")
                        Else
                            oDBs_Detail.SetValue("U_IsCheck", oDBs_Detail.Offset, "N")
                        End If
                        oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemno").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("desc").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_fwhs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("fwhs").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("stock").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("twhs").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("BOMQty").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Qty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("issqty").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("price").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("price").Cells.Item(pVal.Row).Specific.Value) * CDbl(objMatrix.Columns.Item("issqty").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("rmrk").Cells.Item(pVal.Row).Specific.Value))
                        objMatrix.SetLineData(pVal.Row)
                        Me.CalculateTotal(FormUID)
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
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC")
                        oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC_D0")
                        If oCFL.UniqueID = "CFL_Vendor" Then
                            oDBs_Head.SetValue("U_CardCode", 0, oDT.GetValue("CardCode", 0))
                            oDBs_Head.SetValue("U_CardName", 0, oDT.GetValue("CardName", 0))
                            oDBs_Head.SetValue("U_SCNo", 0, "")
                            oDBs_Head.SetValue("U_SCDat", 0, "")
                            oDBs_Head.SetValue("U_RefNo", 0, "")
                            oDBs_Head.SetValue("U_Buyer", 0, "")
                            oDBs_Head.SetValue("U_ContPer", 0, "")
                            Me.FilterSC(FormUID)

                            ''Contact Person
                            'objCombo = objForm.Items.Item("contper").Specific
                            'If objCombo.ValidValues.Count > 0 Then
                            '    For i As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
                            '        objCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                            '    Next
                            'End If
                            'objCombo.ValidValues.Add("", "")
                            'Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRS.DoQuery("Select CntctCode,Name from OCPR where CardCode='" & oDT.GetValue("CardCode", 0) & "'")
                            'If oRS.RecordCount > 0 Then
                            '    For i As Integer = 1 To oRS.RecordCount
                            '        objCombo.ValidValues.Add(Trim(oRS.Fields.Item("CntctCode").Value), Trim(oRS.Fields.Item("Name").Value))
                            '        oRS.MoveNext()
                            '    Next
                            'End If
                        ElseIf oCFL.UniqueID = "CFL_SCNo" Then
                            oDBs_Head.SetValue("U_SCDocNo", 0, oDT.GetValue("DocEntry", 0))
                            oDBs_Head.SetValue("U_SCNo", 0, oDT.GetValue("DocNum", 0))
                            oDBs_Head.SetValue("U_SCDat", 0, CDate(oDT.GetValue("U_DocDate", 0)).ToString("yyyyMMdd"))
                            oDBs_Head.SetValue("U_RefNo", 0, oDT.GetValue("U_VendRef", 0))
                            oDBs_Head.SetValue("U_Buyer", 0, oDT.GetValue("U_Buyer", 0))
                            oDBs_Head.SetValue("U_ContPer", 0, oDT.GetValue("U_ContPer", 0))
                            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'Vijeesh
                            oRS.DoQuery("Select DocEntry From [@GEN_SUB_CONTRACT] Where DocEntry = '" + Trim(oDBs_Head.GetValue("U_SCDocNo", 0)) + "' And (ISnull(u_manual,'N') = 'Y' Or ISnull(U_manwobom,'N') = 'Y') ")
                            If oRS.RecordCount > 0 Then
                                objForm.Items.Item("t_Qty").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            Else
                                objForm.Items.Item("t_Qty").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            End If

                            'oRS.DoQuery("Select DocEntry From [@GEN_SUB_CONTRACT] Where DocEntry = '" + Trim(oDBs_Head.GetValue("U_SCDocNo", 0)) + "' And ISnull(u_manual,'N') = 'Y'")
                            'If oRS.RecordCount > 0 Then
                            '    Me.FilterManualItems(FormUID)
                            'Else
                            Me.FilterItemBOM(FormUID)
                            'End If
                        ElseIf oCFL.UniqueID = "CFL_SO" Then
                            oDBs_Head.SetValue("U_SONo", 0, oDT.GetValue("DocEntry", 0))
                            oDBs_Head.SetValue("U_SODNo", 0, oDT.GetValue("DocNum", 0))
                            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ElseIf oCFL.UniqueID = "CFL_BOMITM" Then
                            oDBs_Head.SetValue("U_ItemNo", 0, oDT.GetValue("ItemCode", 0))
                            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim oRS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRS.DoQuery("select Isnull(U_manual,'N') 'Manual',Isnull(U_cstbom,'N') 'CSTBOM',Isnull(U_manwobom,'N')'MWOBOM',U_VendWhs VendWhs from [@GEN_SUB_CONTRACT] Where DocEntry='" & Trim(oDBs_Head.GetValue("U_SCDocNo", 0)) & "'")
                            Dim Manual As String = Trim(oRS.Fields.Item("Manual").Value)
                            Dim CSTBOM As String = Trim(oRS.Fields.Item("CSTBOM").Value)
                            Dim Manual_wo_BOM As String = Trim(oRS.Fields.Item("MWOBOM").Value)
                            Dim VendorWhs As String = Trim(oRS.Fields.Item("VendWhs").Value)
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            objMatrix.Clear()
                            If Manual = "N" And CSTBOM = "N" And Manual_wo_BOM = "N" Then
                                'RS.DoQuery("Select T1.ItemCode,T1.ItemName,T0.Quantity BOMQty,T2.AvgPrice Price,T1.InvntryUoM UoM,T2.OnHand InStock,T2.WhsCode from ITT1 T0 INNER JOIN OITM T1 ON T0.Code=T1.ItemCode INNER JOIN OITW T2 ON T2.ItemCode=T1.ItemCode and T2.WhsCode='" & VendorWhs & "' Where T0.Father='" & oDT.GetValue("ItemCode", 0) & "'")
                                'oRS.DoQuery("Select T1.ItemCode,T1.ItemName,T3.Quantity BOMQty,ISNULL(T0.U_POQty,0) POQty,(CASE WHEN ISNULL(T0.U_POQty,0)-ISNULL(T0.U_DCQty,0)>0 THEN ISNULL(T0.U_POQty,0)-ISNULL(T0.U_DCQty,0) ELSE 0 END) Quantity,T2.AvgPrice Price,T1.InvntryUoM UoM,T2.OnHand InStock,T2.WhsCode,T3.Warehouse FromWhsCode from [@GEN_SUB_CONTRACT_D1] T0 INNER JOIN OITM T1 ON T1.ItemCode=T0.U_Code and T1.EvalSystem<>'S' INNER JOIN OITW T2 ON T2.ItemCode=T1.ItemCode and T2.WhsCode='" & VendorWhs & "' INNER JOIN ITT1 T3 ON T3.Father='" & oDT.GetValue("ItemCode", 0) & "' and T3.Code=T1.ItemCode Where T0.U_Father='" & oDT.GetValue("ItemCode", 0) & "' and T0.DocEntry='" & Trim(oDBs_Head.GetValue("U_SCDocNo", 0)) & "'")
                                Dim _strQry As String = "Select T1.ItemCode,T1.ItemName,T0.U_BOMQty BOMQty,ISNULL(T0.U_POQty,0) POQty,(CASE WHEN ISNULL(T0.U_POQty,0)-ISNULL(T0.U_DCQty,0)>0 THEN ISNULL(T0.U_POQty,0)-ISNULL(T0.U_DCQty,0) ELSE 0 END) Quantity,T2.AvgPrice Price,T1.InvntryUoM UoM,T2.OnHand InStock,T2.WhsCode,T0.U_FWhs FromWhsCode from [@GEN_SUB_CONTRACT_D1] T0 INNER JOIN OITM T1 ON T1.ItemCode=T0.U_Code and T1.EvalSystem<>'S' INNER JOIN OITW T2 ON T2.ItemCode=T1.ItemCode and T2.WhsCode='" & VendorWhs & "' /*INNER JOIN ITT1 T3 ON T3.Father='" & oDT.GetValue("ItemCode", 0) & "' and T3.Code=T1.ItemCode*/ Where T0.U_Father='" & oDT.GetValue("ItemCode", 0) & "' and T0.DocEntry='" & Trim(oDBs_Head.GetValue("U_SCDocNo", 0)) & "'"
                                oRS.DoQuery(_strQry)
                                For Row As Integer = 1 To oRS.RecordCount
                                    objMatrix.AddRow()
                                    'Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
                                    objMatrix.FlushToDataSource()
                                    oDBs_Detail.Offset = Row - 1
                                    oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
                                    oDBs_Detail.SetValue("U_IsCheck", oDBs_Detail.Offset, "N")
                                    oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, Trim(oRS.Fields.Item("ItemCode").Value))
                                    oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, Trim(oRS.Fields.Item("ItemName").Value))
                                    oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, Trim(oRS.Fields.Item("WhsCode").Value))
                                    oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, Trim(oRS.Fields.Item("UoM").Value))
                                    oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("BOMQty").Value))
                                    oDBs_Detail.SetValue("U_Qty", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("POQty").Value))
                                    oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("BOMQty").Value) * CDbl(objForm.Items.Item("t_Qty").Specific.Value))
                                    oDBs_Detail.SetValue("U_CompQty", oDBs_Detail.Offset, 0)
                                    oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, "")
                                    oDBs_Detail.SetValue("U_fwhs", oDBs_Detail.Offset, Trim(oRS.Fields.Item("FromWhsCode").Value))
                                    oRS1.DoQuery("Select OnHand,AvgPrice from OITW Where ItemCode='" & Trim(oRS.Fields.Item("ItemCode").Value) & "' and WhsCode='" & Trim(oRS.Fields.Item("FromWhsCode").Value) & "'")
                                    If oRS1.RecordCount > 0 Then
                                        oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, CDbl(oRS1.Fields.Item("OnHand").Value))
                                        oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(oRS1.Fields.Item("AvgPrice").Value))
                                        oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, CDbl(oRS1.Fields.Item("AvgPrice").Value) * CDbl(oRS.Fields.Item("Quantity").Value))
                                    Else
                                        oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, 0)
                                        oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, 0)
                                        oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, 0)
                                    End If
                                    objMatrix.SetLineData(Row)
                                    oRS.MoveNext()
                                Next
                            End If
                            If CSTBOM = "Y" And Manual = "N" And Manual_wo_BOM = "N" Then
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet.DoQuery("Select Top 1 DocEntry As 'DocEntry' From [@GEN_CUST_BOM] Where u_itemcode = '" + Trim(oDT.GetValue("ItemCode", 0)) + "' Order By u_docdate desc")
                                Dim docentry As String = oRSet.Fields.Item("DocEntry").Value
                                'RS.DoQuery("Select T1.ItemCode,T1.ItemName,T0.Quantity BOMQty,T2.AvgPrice Price,T1.InvntryUoM UoM,T2.OnHand InStock,T2.WhsCode from ITT1 T0 INNER JOIN OITM T1 ON T0.Code=T1.ItemCode INNER JOIN OITW T2 ON T2.ItemCode=T1.ItemCode and T2.WhsCode='" & VendorWhs & "' Where T0.Father='" & oDT.GetValue("ItemCode", 0) & "'")

                                ''oRS.DoQuery("Select Distinct T3.u_unit,T4.u_process,T4.u_itemcode,T4.u_itemname,T4.u_qty BOMQty,ISNULL(T0.U_POQty,0) POQty,(CASE WHEN ISNULL(T0.U_POQty,0)-ISNULL(T0.U_DCQty,0)>0 THEN ISNULL(T0.U_POQty,0)-ISNULL(T0.U_DCQty,0) ELSE 0 END) Quantity,T2.AvgPrice Price,T1.InvntryUoM UoM,T2.OnHand InStock,T2.WhsCode from [@GEN_SUB_CONTRACT_D1] T0 INNER JOIN OITM T1 ON T1.ItemCode=T0.U_Code and T1.EvalSystem<>'S' INNER JOIN OITW T2  ON T2.ItemCode=T1.ItemCode and T2.WhsCode='" & VendorWhs & "' INNER JOIN [@GEN_CUST_BOM] T3 ON T3.u_itemcode = T0.u_Father Inner Join [@GEN_CUST_BOM_D0] T4 On T3.DocEntry = T4.DocEntry And T0.u_Code = T4.u_itemcode Where T0.U_Father='" & Trim(oDT.GetValue("ItemCode", 0)) & "' And T3.u_itemcode = T0.U_Father and T0.DocEntry='" & Trim(oDBs_Head.GetValue("U_SCDocNo", 0)) & "' And T3.DocEntry = '" + docentry + "'")

                                Dim _str_Query As String = "Select Distinct T0.U_Code u_itemcode ,T1.itemname u_itemname,T0.U_BOMQty BOMQty," _
                                                            & "ISNULL(T0.U_POQty,0) POQty, " _
                                                            & "(CASE WHEN ISNULL(T0.U_POQty,0)-ISNULL(T0.U_DCQty,0)>0 THEN ISNULL(T0.U_POQty,0)-ISNULL(T0.U_DCQty,0) ELSE 0 END) Quantity," _
                                                            & "T2.AvgPrice Price,T1.InvntryUoM UoM,T2.OnHand InStock,T2.WhsCode,T0.U_FWhs " _
                                                            & "from [@GEN_SUB_CONTRACT_D1] T0 " _
                                                            & "INNER JOIN OITM T1 ON T1.ItemCode=T0.U_Code and T1.EvalSystem<>'S' " _
                                                            & "INNER JOIN OITW T2  ON T2.ItemCode=T1.ItemCode and T2.WhsCode='" & VendorWhs & "' " _
                                                            & "Where T0.U_Father='" & Trim(oDT.GetValue("ItemCode", 0)) & "' and T0.DocEntry='" & Trim(oDBs_Head.GetValue("U_SCDocNo", 0)) & "'"
                                oRS.DoQuery(_str_Query)

                                For Row As Integer = 1 To oRS.RecordCount
                                    objMatrix.AddRow()
                                    'Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
                                    'oRSet.DoQuery("Select B.u_stwhs From [@GEN_UNIT_MST] A Inner Join [@GEN_UNIT_MST_D0] B On A.Code = B.Code ANd A.Name = '" + Trim(oRS.Fields.Item("u_unit").Value) + "' Where B.u_process = '" + Trim(oRS.Fields.Item("u_process").Value) + "'")
                                    objMatrix.FlushToDataSource()
                                    oDBs_Detail.Offset = Row - 1
                                    oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
                                    oDBs_Detail.SetValue("U_IsCheck", oDBs_Detail.Offset, "N")
                                    oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, Trim(oRS.Fields.Item("u_itemcode").Value))
                                    oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, Trim(oRS.Fields.Item("u_itemname").Value))
                                    oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, Trim(oRS.Fields.Item("WhsCode").Value))
                                    oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, Trim(oRS.Fields.Item("UoM").Value))
                                    oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("BOMQty").Value))
                                    oDBs_Detail.SetValue("U_Qty", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("POQty").Value))
                                    oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("BOMQty").Value) * CDbl(objForm.Items.Item("t_Qty").Specific.Value))
                                    oDBs_Detail.SetValue("U_CompQty", oDBs_Detail.Offset, 0)
                                    oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, "")
                                    oDBs_Detail.SetValue("U_fwhs", oDBs_Detail.Offset, Trim(oRS.Fields.Item("U_FWhs").Value))
                                    'oDBs_Detail.SetValue("U_fwhs", oDBs_Detail.Offset, Trim(oRSet.Fields.Item("u_stwhs").Value))
                                    'oRS1.DoQuery("Select OnHand,AvgPrice from OITW Where ItemCode='" & Trim(oRS.Fields.Item("u_itemcode").Value) & "' and WhsCode='" & Trim(oRSet.Fields.Item("u_stwhs").Value) & "'")
                                    oRS1.DoQuery("Select OnHand,AvgPrice from OITW Where ItemCode='" & Trim(oRS.Fields.Item("u_itemcode").Value) & "' and WhsCode='" & Trim(oRS.Fields.Item("U_FWhs").Value) & "'")
                                    If oRS1.RecordCount > 0 Then
                                        oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, CDbl(oRS1.Fields.Item("OnHand").Value))
                                        oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(oRS1.Fields.Item("AvgPrice").Value))
                                        oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, CDbl(oRS1.Fields.Item("AvgPrice").Value) * CDbl(oRS.Fields.Item("Quantity").Value))
                                    Else
                                        oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, 0)
                                        oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, 0)
                                        oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, 0)
                                    End If
                                    objMatrix.SetLineData(Row)
                                    oRS.MoveNext()
                                Next
                            End If

                            If Manual = "Y" And Manual_wo_BOM = "N" Then
                                'Vijeesh
                                'objMatrix.AddRow()
                                'objMatrix.FlushToDataSource()
                                'objMatrix.AutoResizeColumns()
                                'Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
                                Dim strQuery As String = "Select T1.ItemCode,T1.ItemName,ISNULL(T0.U_POQty,0) POQty," _
                                            & " (CASE WHEN ISNULL(T0.U_POQty,0)-ISNULL(T0.U_DCQty,0)>0 THEN ISNULL(T0.U_POQty,0)-ISNULL(T0.U_DCQty,0) ELSE 0 END) Quantity," _
                                            & " T2.AvgPrice Price,T1.InvntryUoM UoM,T2.OnHand InStock,T2.WhsCode,T3.U_Whs FromWhsCode" _
                                            & " from [@GEN_SUB_CONTRACT_D1] T0 INNER JOIN OITM T1 " _
                                            & " ON T1.ItemCode=T0.U_Code and T1.EvalSystem<>'S' " _
                                            & " INNER JOIN OITW T2 ON T2.ItemCode=T1.ItemCode" _
                                            & " INNER JOIN [@GEN_SUB_CONTRACT_D0]T3 on T0.DocEntry = T3.DocEntry and T0.LineId = T3.LineId " _
                                            & " and T2.WhsCode='" & VendorWhs & "' " _
                                            & " and T0.DocEntry='" & Trim(oDBs_Head.GetValue("U_SCDocNo", 0)) & "' and T0.U_Code ='" & oDT.GetValue("ItemCode", 0) & "'"""



                                oRS.DoQuery("Select T1.ItemCode,T1.ItemName,ISNULL(T0.U_POQty,0) POQty," _
                                            & " (CASE WHEN ISNULL(T0.U_POQty,0)-ISNULL(T0.U_DCQty,0)>0 THEN ISNULL(T0.U_POQty,0)-ISNULL(T0.U_DCQty,0) ELSE 0 END) Quantity," _
                                            & " T2.AvgPrice Price,T1.InvntryUoM UoM,T2.OnHand InStock,T2.WhsCode,T3.U_Whs FromWhsCode" _
                                            & " from [@GEN_SUB_CONTRACT_D1] T0 INNER JOIN OITM T1 " _
                                            & " ON T1.ItemCode=T0.U_Code and T1.EvalSystem<>'S' " _
                                            & " INNER JOIN OITW T2 ON T2.ItemCode=T1.ItemCode" _
                                            & " INNER JOIN [@GEN_SUB_CONTRACT_D0]T3 on T0.DocEntry = T3.DocEntry and T0.LineId = T3.LineId " _
                                            & " and T2.WhsCode='" & VendorWhs & "' " _
                                            & " and T0.DocEntry='" & Trim(oDBs_Head.GetValue("U_SCDocNo", 0)) & "' and T0.U_Code ='" & oDT.GetValue("ItemCode", 0) & "'")
                                For Row As Integer = 1 To oRS.RecordCount
                                    objMatrix.AddRow()
                                    objMatrix.FlushToDataSource()
                                    oDBs_Detail.Offset = Row - 1
                                    oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
                                    oDBs_Detail.SetValue("U_IsCheck", oDBs_Detail.Offset, "N")
                                    oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, Trim(oRS.Fields.Item("ItemCode").Value))
                                    oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, Trim(oRS.Fields.Item("ItemName").Value))
                                    oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, Trim(oRS.Fields.Item("WhsCode").Value))
                                    oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, Trim(oRS.Fields.Item("UoM").Value))
                                    'oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("POQty").Value))
                                    oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, 1)
                                    oDBs_Detail.SetValue("U_Qty", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("POQty").Value))
                                    oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("POQty").Value) * CDbl(objForm.Items.Item("t_Qty").Specific.Value))
                                    oDBs_Detail.SetValue("U_CompQty", oDBs_Detail.Offset, 0)
                                    oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, "")
                                    oDBs_Detail.SetValue("U_fwhs", oDBs_Detail.Offset, Trim(oRS.Fields.Item("FromWhsCode").Value))
                                    oRS1.DoQuery("Select OnHand,AvgPrice from OITW Where ItemCode='" & Trim(oRS.Fields.Item("ItemCode").Value) & "' and WhsCode='" & Trim(oRS.Fields.Item("FromWhsCode").Value) & "'")
                                    If oRS1.RecordCount > 0 Then
                                        oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, CDbl(oRS1.Fields.Item("OnHand").Value))
                                        oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(oRS1.Fields.Item("AvgPrice").Value))
                                        oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, CDbl(oRS1.Fields.Item("AvgPrice").Value) * CDbl(oRS.Fields.Item("Quantity").Value))
                                    Else
                                        oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, 0)
                                        oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, 0)
                                        oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, 0)
                                    End If
                                    objMatrix.SetLineData(Row)
                                    oRS.MoveNext()
                                Next
                            End If
                            'Vijeesh

                            If Manual_wo_BOM = "Y" Then
                                'Vijeesh
                                'objMatrix.AddRow()
                                'objMatrix.FlushToDataSource()
                                'objMatrix.AutoResizeColumns()
                                'Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
                                oRS.DoQuery("Select T1.ItemCode,T1.ItemName,ISNULL(T0.U_POQty,0) POQty,ISNULL(T3.U_Quantity,0)Qty," _
                                            & " (CASE WHEN ISNULL(T0.U_POQty,0)-ISNULL(T0.U_DCQty,0)>0 THEN ISNULL(T0.U_POQty,0)-ISNULL(T0.U_DCQty,0) ELSE 0 END) Quantity," _
                                            & " T2.AvgPrice Price,T1.InvntryUoM UoM,T2.OnHand InStock,T2.WhsCode,T3.U_Whs FromWhsCode" _
                                            & " from [@GEN_SUB_CONTRACT_D1] T0 INNER JOIN OITM T1 " _
                                            & " ON T1.ItemCode=T0.U_Code and T1.EvalSystem<>'S' " _
                                            & " INNER JOIN OITW T2 ON T2.ItemCode=T1.ItemCode" _
                                            & " INNER JOIN [@GEN_SUB_CONTRACT_D0]T3 on T0.DocEntry = T3.DocEntry " _
                                            & " and T2.WhsCode='" & VendorWhs & "' " _
                                            & " and T0.DocEntry='" & Trim(oDBs_Head.GetValue("U_SCDocNo", 0)) & "' and T0.U_Father ='" & oDT.GetValue("ItemCode", 0) & "'")
                                For Row As Integer = 1 To oRS.RecordCount
                                    objMatrix.AddRow()
                                    objMatrix.FlushToDataSource()
                                    oDBs_Detail.Offset = Row - 1
                                    oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
                                    oDBs_Detail.SetValue("U_IsCheck", oDBs_Detail.Offset, "N")
                                    oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, Trim(oRS.Fields.Item("ItemCode").Value))
                                    oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, Trim(oRS.Fields.Item("ItemName").Value))
                                    oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, Trim(oRS.Fields.Item("WhsCode").Value))
                                    oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, Trim(oRS.Fields.Item("UoM").Value))
                                    oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("POQty").Value) / CDbl(oRS.Fields.Item("Qty").Value))
                                    'oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, 1)
                                    oDBs_Detail.SetValue("U_Qty", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("POQty").Value))
                                    oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("POQty").Value) * CDbl(objForm.Items.Item("t_Qty").Specific.Value))
                                    oDBs_Detail.SetValue("U_CompQty", oDBs_Detail.Offset, 0)
                                    oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, "")
                                    oDBs_Detail.SetValue("U_fwhs", oDBs_Detail.Offset, Trim(oRS.Fields.Item("FromWhsCode").Value))
                                    oRS1.DoQuery("Select OnHand,AvgPrice from OITW Where ItemCode='" & Trim(oRS.Fields.Item("ItemCode").Value) & "' and WhsCode='" & Trim(oRS.Fields.Item("FromWhsCode").Value) & "'")
                                    If oRS1.RecordCount > 0 Then
                                        oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, CDbl(oRS1.Fields.Item("OnHand").Value))
                                        oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(oRS1.Fields.Item("AvgPrice").Value))
                                        oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, CDbl(oRS1.Fields.Item("AvgPrice").Value) * CDbl(oRS.Fields.Item("Quantity").Value))
                                    Else
                                        oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, 0)
                                        oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, 0)
                                        oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, 0)
                                    End If
                                    objMatrix.SetLineData(Row)
                                    oRS.MoveNext()
                                Next
                            End If

                            'Vijeesh
                            Me.CalculateTotal(FormUID)
                        ElseIf oCFL.UniqueID = "CFL_Owner" Then
                            oDBs_Head.SetValue("U_OwnerCod", 0, oDT.GetValue("empID", 0))
                            oDBs_Head.SetValue("U_Owner", 0, oDT.GetValue("firstName", 0) + " " + oDT.GetValue("lastName", 0))
                        ElseIf oCFL.UniqueID = "CFL_twhs" Then
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            oDBs_Detail.Offset = pVal.Row - 1
                            oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                            If objMatrix.Columns.Item("IsCheck").Cells.Item(pVal.Row).Specific.Checked = True Then
                                oDBs_Detail.SetValue("U_IsCheck", oDBs_Detail.Offset, "Y")
                            Else
                                oDBs_Detail.SetValue("U_IsCheck", oDBs_Detail.Offset, "N")
                            End If
                            oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemno").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("desc").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_fwhs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("fwhs").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("stock").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, Trim(oDT.GetValue("WhsCode", 0)))
                            oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("BOMQty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Qty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("issqty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_CompQty", oDBs_Detail.Offset, 0)
                            oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("price").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("total").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("rmrk").Cells.Item(pVal.Row).Specific.Value))
                            objMatrix.SetLineData(pVal.Row)
                        ElseIf oCFL.UniqueID = "CFL_Item" Then
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRS.DoQuery("select U_VendWhs VendWhs from [@GEN_SUB_CONTRACT] Where DocEntry='" & Trim(oDBs_Head.GetValue("U_SCDocNo", 0)) & "'")
                            Dim VendorWhs As String = Trim(oRS.Fields.Item("VendWhs").Value)
                            Dim OrginRow As Integer = objMatrix.VisualRowCount
                            For i As Integer = 0 To oDT.Rows.Count - 1
                                Dim cflSelectedcount As Integer = oDT.Rows.Count
                                If i < cflSelectedcount - 1 Then
                                    objMatrix.AddRow(1, pVal.Row)
                                    oDBs_Detail.InsertRecord(pVal.Row + i - 1)
                                End If
                                oDBs_Detail.Offset = pVal.Row - 1 + i
                                oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                                oDBs_Detail.SetValue("U_IsCheck", oDBs_Detail.Offset, "N")
                                oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, oDT.GetValue("ItemCode", i))
                                oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, oDT.GetValue("ItemName", i))
                                oDBs_Detail.SetValue("U_fwhs", oDBs_Detail.Offset, "")
                                oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, 0)
                                oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, VendorWhs)
                                oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, oDT.GetValue("InvntryUom", i))
                                oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, 0)
                                oDBs_Detail.SetValue("U_Qty", oDBs_Detail.Offset, 0)
                                oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, 1)
                                oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, 0)
                                oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, oDT.GetValue("LastPurPrc", i))
                                oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, oDT.GetValue("LastPurPrc", i))
                                oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, "")
                                objMatrix.SetLineData(pVal.Row + i)
                            Next
                            objMatrix.FlushToDataSource()
                            If OrginRow = pVal.Row Then
                                objMatrix.AddRow()
                                objMatrix.FlushToDataSource()
                                Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
                            End If
                            For Row As Integer = 1 To objMatrix.VisualRowCount
                                objMatrix.Columns.Item("SNo").Cells.Item(Row).Specific.Value = Row
                            Next
                            Me.CalculateTotal(FormUID)
                        ElseIf oCFL.UniqueID = "CFL_fwhs" Then
                            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRS.DoQuery("Select OnHand,AvgPrice from OITW Where ItemCode='" & Trim(objMatrix.Columns.Item("itemno").Cells.Item(pVal.Row).Specific.Value) & "' and WhsCode='" & Trim(oDT.GetValue("WhsCode", 0)) & "'")
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            oDBs_Detail.Offset = pVal.Row - 1
                            oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                            If objMatrix.Columns.Item("IsCheck").Cells.Item(pVal.Row).Specific.Checked = True Then
                                oDBs_Detail.SetValue("U_IsCheck", oDBs_Detail.Offset, "Y")
                            Else
                                oDBs_Detail.SetValue("U_IsCheck", oDBs_Detail.Offset, "N")
                            End If
                            oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemno").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("desc").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_fwhs", oDBs_Detail.Offset, Trim(oDT.GetValue("WhsCode", 0)))
                            oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("OnHand").Value))
                            oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("twhs").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("BOMQty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Qty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("issqty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_CompQty", oDBs_Detail.Offset, 0)
                            oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("AvgPrice").Value))
                            oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("AvgPrice").Value) * CDbl(objMatrix.Columns.Item("issqty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("rmrk").Cells.Item(pVal.Row).Specific.Value))
                            objMatrix.SetLineData(pVal.Row)
                            Me.CalculateTotal(FormUID)
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
                    Case "SC_DC"
                        Me.CreateForm()
                    Case "1282"
                        If objForm.TypeEx = "GEN_SCDC" Then
                            Me.SetDefault(objForm.UniqueID)
                        End If
                    Case "1281"
                        If objForm.TypeEx = "GEN_SCDC" Then
                            objForm.EnableMenu("1282", True)
                            objForm.Items.Item("Issue").Enabled = False
                            objForm.Items.Item("t_docno").Click()
                        End If
                    Case "Close"
                        If objForm.TypeEx = "GEN_SCDC" Then
                            If oApplication.MessageBox("Do you want to close?", 2, "Ok", "Cancel") = 1 Then
                                Dim ORS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                ORS.DoQuery("UPDATE [@GEN_SC_DC] SET U_Status='Closed' Where DocNum='" & oDBs_Head.GetValue("DocNum", 0) & "'")
                                oDBs_Head.SetValue("U_Status", 0, "Closed")
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                objForm.Items.Item("1").Enabled = True
                            End If
                        End If
                    Case "1293"
                        If objForm.TypeEx = "GEN_SCDC" Then
                            objForm.Freeze(True)
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC")
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC_D0")
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            For Row As Integer = 1 To objMatrix.VisualRowCount
                                objMatrix.GetLineData(Row)
                                oDBs_Detail.Offset = Row - 1
                                oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
                                If objMatrix.Columns.Item("IsCheck").Cells.Item(Row).Specific.Checked = True Then
                                    oDBs_Detail.SetValue("U_IsCheck", oDBs_Detail.Offset, "Y")
                                Else
                                    oDBs_Detail.SetValue("U_IsCheck", oDBs_Detail.Offset, "N")
                                End If
                                oDBs_Detail.SetValue("U_ItemNo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemno").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Desc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("desc").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_fwhs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("fwhs").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Stock", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("stock").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_TWhs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("twhs").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Unit", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_BOMQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("BOMQty").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Qty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Qty").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_IssQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("issqty").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_CompQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("compqty").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("price").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Total", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("total").Cells.Item(Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Rmrk", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("rmrk").Cells.Item(Row).Specific.Value))
                                objMatrix.SetLineData(Row)
                            Next
                            objMatrix.FlushToDataSource()
                            oDBs_Detail.RemoveRecord(oDBs_Detail.Size - 1)
                            objMatrix.LoadFromDataSource()
                            objForm.Freeze(False)
                        End If
                End Select

            ElseIf pVal.BeforeAction = True Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "519"
                        Try
                            If objForm.TypeEx = "GEN_SCDC" Then
                                BubbleEvent = False
                                sDocNum = objForm.Items.Item("t_docno").Specific.Value
                                sRptName = "GEN_SEPL_DC.rpt"
                                Me.PrintSC_DCReport()
                            End If
                        Catch ex As Exception

                        End Try
                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    'Private Sub Report1()
    '    Dim oThread As New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Report1Thread))
    '    oThread.SetApartmentState(System.Threading.ApartmentState.STA)
    '    oThread.Start()
    'End Sub

    'Private Sub Report1Thread()
    '    Try
    '        Dim oCRForm As New Crystal_Form
    '        oCRForm.ShowDialog()
    '    Catch ex As Exception
    '        oApplication.MessageBox(ex.Message.ToString)
    '    End Try
    'End Sub

    'Private Sub Report1()
    '    Dim oThread As New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Report1Thread))
    '    oThread.SetApartmentState(System.Threading.ApartmentState.STA)
    '    oThread.Start()
    'End Sub

    'Private Sub Report1Thread()
    '    Try
    '        Dim oCRForm As New Crystal_Form
    '        oCRForm.ShowDialog()
    '    Catch ex As Exception
    '        oApplication.MessageBox(ex.Message.ToString)
    '    End Try
    'End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Select Case BusinessObjectInfo.EventType
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                'objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                'oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC")
                'oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC_D0")
                'If BusinessObjectInfo.BeforeAction = True Then
                '    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                '        oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("c_series").Specific.Selected.Value, "GEN_SC_DC"))
                '        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                '        oRS.DoQuery("UPDATE OWTR SET U_Type='SubCont_DC',U_DocNum='" & Trim(oDBs_Head.GetValue("DocNum", 0)) & "' Where DocEntry='" & Trim(oDBs_Head.GetValue("U_InvTrNo", 0)) & "'")
                '        'oRS.DoQuery("UPDATE [@GEN_SUB_CONTRACT_D0] SET U_DCQty=ISNULL(U_DCQty,0)+" & CDbl(objForm.Items.Item("37").Specific.Value) & " Where DocEntry='" & Trim(oDBs_Head.GetValue("U_SCDocNo", 0)) & "' and U_ItemCode='" & Trim(objForm.Items.Item("ItemNo").Specific.value) & "'")
                '        For Row As Integer = 1 To objMatrix.VisualRowCount
                '            If objMatrix.Columns.Item("IsCheck").Cells.Item(Row).Specific.Checked = True Then
                '                oRS.DoQuery("UPDATE [@GEN_SUB_CONTRACT_D1] SET U_DCQty=ISNULL(U_DCQty,0)+" & CDbl(objMatrix.Columns.Item("issqty").Cells.Item(Row).Specific.Value) & " Where DocEntry='" & Trim(oDBs_Head.GetValue("U_SCDocNo", 0)) & "' and U_Father='" & Trim(objForm.Items.Item("ItemNo").Specific.value) & "' and U_Code='" & Trim(objMatrix.Columns.Item("itemno").Cells.Item(Row).Specific.Value) & "'")
                '            End If
                '        Next
                '    End If

                'objMatrix = objForm.Items.Item("ItemMatrix").Specific
                'objMatrix.LoadFromDataSource()
                'If objMatrix.VisualRowCount <> 0 Then
                '    oDBs_Detail.RemoveRecord(objMatrix.VisualRowCount - 1)
                '    objMatrix.LoadFromDataSource()
                'End If
                If BusinessObjectInfo.ActionSuccess = True Then
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Then
                        'objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        'objMatrix.AddRow()
                        'objMatrix.FlushToDataSource()
                        'Me.SetNewLine(BusinessObjectInfo.FormUID, objMatrix.VisualRowCount)
                    End If
                End If
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                If BusinessObjectInfo.ActionSuccess = True Then
                    objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    objForm.EnableMenu("1282", True)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC")
                    oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC_D0")
                    'objMatrix = objForm.Items.Item("ItemMatrix").Specific
                    'objMatrix.AddRow()
                    'objMatrix.FlushToDataSource()
                    'Me.SetNewLine(BusinessObjectInfo.FormUID, objMatrix.VisualRowCount)
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                    objForm.Items.Item("1").Enabled = True
                    objForm.Items.Item("Issue").Enabled = True
                End If
        End Select
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try

            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC_D0")
            Dim oCheck As SAPbouiCOM.CheckBox
            Dim Checked As Integer
            For I As Integer = 1 To objMatrix.VisualRowCount
                oCheck = objMatrix.Columns.Item("IsCheck").Cells.Item(I).Specific
                If oCheck.Checked = False Then
                    Checked = Checked + 1
                    If Checked = objMatrix.VisualRowCount Then
                        oApplication.StatusBar.SetText("Select Items")
                        Return False
                    End If
                End If
            Next
            Dim DeletedRowNo As Integer = 0
            Dim InitialRow As Integer = objMatrix.VisualRowCount
            Checked = 0
            For I As Integer = 1 To InitialRow - 1
                oCheck = objMatrix.Columns.Item("IsCheck").Cells.Item(I).Specific
                If oCheck.Checked = True Then
                    Checked = Checked + 1
                End If
                If oCheck.Checked = False Then
                    objMatrix.DeleteRow(I)
                    DeletedRowNo = DeletedRowNo + 1
                    I = I - 1
                End If
                If InitialRow = Checked + DeletedRowNo Then
                    Exit For
                End If
            Next
            For I As Integer = 1 To objMatrix.VisualRowCount
                oCheck = objMatrix.Columns.Item("IsCheck").Cells.Item(I).Specific
                If oCheck.Checked = False Then
                    objMatrix.DeleteRow(I)
                End If
            Next
            For I As Integer = 1 To objMatrix.VisualRowCount
                objMatrix.Columns.Item("SNo").Cells.Item(I).Specific.Value = I
            Next

            If Trim(objForm.Items.Item("cardcode").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Vendor Code should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Trim(objForm.Items.Item("SCNo").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Subcontracting PO.No. should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Trim(objForm.Items.Item("ItemNo").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Item No. should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Trim(objForm.Items.Item("t_deldt").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Delivery Date should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Trim(objForm.Items.Item("t_docdt").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Document Date should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf (DateDiff(DateInterval.Day, DateTime.ParseExact(Trim(objForm.Items.Item("t_docdt").Specific.Value), "yyyyMMdd", Nothing), DateTime.ParseExact(Trim(objForm.Items.Item("t_deldt").Specific.Value), "yyyyMMdd", Nothing))) < 0 Then
                oApplication.StatusBar.SetText("Delivery date is before Document date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                'ElseIf (objForm.Items.Item("t_SODNo").Specific.Value).Equals("") = True Then
                '    oApplication.StatusBar.SetText("Sales Order No. should be mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Return False
            End If

            'Vijeesh
            If Trim(objForm.Items.Item("SCNo").Specific.Value).Equals("") = False Then
                Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRS.DoQuery("Select U_approve from [@GEN_SUB_CONTRACT] where DocNum ='" + Trim(objForm.Items.Item("SCNo").Specific.Value) + "'")
                If oRS.Fields.Item("U_approve").Value.ToString() = "N" Then
                    oApplication.StatusBar.SetText("Cannot Proceed DC, Approval Needed for Purchase Order No:-> " + Trim(objForm.Items.Item("SCNo").Specific.Value) + "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
            'Vijeesh

            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            Dim WhsCode As String = ""
            If objMatrix.VisualRowCount < 1 Then
                oApplication.StatusBar.SetText("No items defined", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                    oDBs_Detail.Offset = Row - 1
                    If objMatrix.Columns.Item("IsCheck").Cells.Item(Row).Specific.Checked = True Then
                        oDBs_Detail.SetValue("U_IsCheck", oDBs_Detail.Offset, "Y")
                        If Trim(objMatrix.Columns.Item("itemno").Cells.Item(Row).Specific.Value).Equals("") = True Then
                            oApplication.StatusBar.SetText("Row [ " & Row & " ] - ItemCode should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        ElseIf Trim(objMatrix.Columns.Item("fwhs").Cells.Item(Row).Specific.Value).Equals("") = True Then
                            oApplication.StatusBar.SetText("Row [ " & Row & " ] - From Warehouse should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        ElseIf Trim(objMatrix.Columns.Item("twhs").Cells.Item(Row).Specific.Value).Equals("") = True Then
                            oApplication.StatusBar.SetText("Row [ " & Row & " ] - To Warehouse should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        ElseIf Trim(objMatrix.Columns.Item("fwhs").Cells.Item(Row).Specific.Value).Equals(Trim(objMatrix.Columns.Item("twhs").Cells.Item(Row).Specific.Value)) = True Then
                            oApplication.StatusBar.SetText("Row [ " & Row & " ] - From and To Warehouse should not be same", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        ElseIf CDbl(objMatrix.Columns.Item("issqty").Cells.Item(Row).Specific.Value) <= 0 Then
                            oApplication.StatusBar.SetText("Row [ " & Row & " ] - Issue Quantity should greater than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                            'ElseIf CDbl(objMatrix.Columns.Item("issqty").Cells.Item(Row).Specific.Value) > CDbl(objMatrix.Columns.Item("stock").Cells.Item(Row).Specific.Value) Then
                            '    oApplication.StatusBar.SetText("Row [ " & Row & " ] - Issue Quantity should greater than InStock", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            '    Return False
                        End If

                        If WhsCode = "" Then
                            WhsCode = Trim(objMatrix.Columns.Item("fwhs").Cells.Item(Row).Specific.Value)
                        End If
                    Else
                        oDBs_Detail.SetValue("U_IsCheck", oDBs_Detail.Offset, "N")
                    End If
                Next
            End If

            Dim cnt As Integer = 0
            For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                If objMatrix.Columns.Item("IsCheck").Cells.Item(Row).Specific.Checked = True Then
                    If Trim(objMatrix.Columns.Item("fwhs").Cells.Item(Row).Specific.Value).Equals(WhsCode) = False Then
                        cnt = cnt + 1
                    End If
                End If
            Next


            'If cnt > 0 Then
            '    oApplication.StatusBar.SetText("From Warehouse should be same as for all rows", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

            'Dim initSize As Integer = oDBs_Detail.Size
            'Dim currSize As Integer = oDBs_Detail.Size
            'For i As Integer = 0 To initSize - 1
            '    oDBs_Detail.Offset = i - (initSize - currSize)
            '    If Trim(oDBs_Detail.GetValue("U_IsCheck", oDBs_Detail.Offset)).Equals("Y") = False Then
            '        oDBs_Detail.RemoveRecord(oDBs_Detail.Offset)
            '    End If
            '    currSize = oDBs_Detail.Size
            'Next
            'objMatrix.LoadFromDataSource()
            'For i As Integer = 0 To oDBs_Detail.Size - 1
            '    oDBs_Detail.Offset = i
            '    oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + 1)
            'Next
            'objMatrix.LoadFromDataSource()
            'If objMatrix.VisualRowCount = 0 Then
            '    oApplication.StatusBar.SetText("No items defined", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If
            'If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            '    Dim TransferNo As Integer = PostStockTransfer(FormUID)
            '    If TransferNo <> 0 Then
            '        objForm.Items.Item("InvTrNo").Specific.Value = TransferNo
            '    Else
            '        Return False
            '    End If
            'End If

            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
            Return False
        End Try
    End Function

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
    End Sub

    Sub FilterItemBOM(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("CFL_BOMITM")
            oCFL.SetConditions(emptyConds)
            If Trim(objForm.Items.Item("SCNo").Specific.value) <> "" Then
                objForm = oApplication.Forms.Item(FormUID)
                oCFL.SetConditions(emptyConds)
                oCons = oCFL.GetConditions()
                Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRS.DoQuery("Select b.U_ItemCode from [@GEN_SUB_CONTRACT] a join [@GEN_SUB_CONTRACT_D0] b on a.docentry=b.docentry where ISNULL(b.U_Quantity,0)>ISNULL(b.U_DCQty,0) and a.DocEntry= '" & Trim(objForm.Items.Item("SCDocNo").Specific.Value) & "'")
                oRS.DoQuery("Select DISTINCT T0.U_ItemCode  from [@GEN_SUB_CONTRACT_D0] T0 INNER JOIN [@GEN_SUB_CONTRACT_D1] T1 ON T0.DocEntry=T1.DocEntry and T0.LineID=T1.U_LineID Where T0.DocEntry='" & Trim(objForm.Items.Item("SCDocNo").Specific.Value) & "'")
                For i As Integer = 0 To oRS.RecordCount - 1
                    If i > 0 Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRS.Fields.Item("U_ItemCode").Value
                    oRS.MoveNext()
                Next
                If oRS.RecordCount = 0 Then
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "-1"
                End If
                oCFL.SetConditions(oCons)
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterManualItems(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("CFL_Item")
            oCFL.SetConditions(emptyConds)
            If Trim(objForm.Items.Item("SCNo").Specific.value) <> "" Then
                objForm = oApplication.Forms.Item(FormUID)
                oCFL.SetConditions(emptyConds)
                oCons = oCFL.GetConditions()
                Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRS.DoQuery("Select b.U_ItemCode from [@GEN_SUB_CONTRACT] a join [@GEN_SUB_CONTRACT_D0] b on a.docentry=b.docentry where ISNULL(b.U_Quantity,0)>ISNULL(b.U_DCQty,0) and a.DocEntry= '" & Trim(objForm.Items.Item("SCDocNo").Specific.Value) & "'")
                oRS.DoQuery("Select DISTINCT T0.U_ItemCode  from [@GEN_SUB_CONTRACT_D0] T0  Where T0.DocEntry='" & Trim(objForm.Items.Item("SCDocNo").Specific.Value) & "'")
                For i As Integer = 0 To oRS.RecordCount - 1
                    If i > 0 Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRS.Fields.Item("U_ItemCode").Value
                    oRS.MoveNext()
                Next
                If oRS.RecordCount = 0 Then
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "-1"
                End If
                oCFL.SetConditions(oCons)
            End If
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
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("CFL_SCNo")
            oCFL.SetConditions(emptyConds)
            objForm = oApplication.Forms.Item(FormUID)
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Select DISTINCT T0.DocEntry from [@GEN_SUB_CONTRACT] T0 Where T0.U_CardCode='" & Trim(objForm.Items.Item("cardcode").Specific.Value) & "'  and T0.U_Status='Open'")
            For i As Integer = 0 To oRS.RecordCount - 1
                If i > 0 Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                oCon = oCons.Add()
                oCon.Alias = "DocEntry"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = oRS.Fields.Item("DocEntry").Value
                oRS.MoveNext()
            Next
            If oRS.RecordCount = 0 Then
                oCon = oCons.Add()
                oCon.Alias = "DocEntry"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "-1"
            End If
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterWarehouse(ByVal FormUID As String)
        Try

            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("CFL_fwhs")
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Select DISTINCT U_whs WhsCode from [@GEN_WHS_USR] Where U_user='" & oCompany.UserName & "'")
            For i As Integer = 0 To oRS.RecordCount - 1
                If i > 0 Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                oCon = oCons.Add()
                oCon.Alias = "WhsCode"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = oRS.Fields.Item("WhsCode").Value
                oRS.MoveNext()
            Next
            If oRS.RecordCount = 0 Then
                oCon = oCons.Add()
                oCon.Alias = "WhsCode"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "-1"
            End If
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CalculateTotal(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC_D0")
            Dim TotalAmount As Double = 0
            For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                TotalAmount = TotalAmount + CDbl(objMatrix.Columns.Item("total").Cells.Item(Row).Specific.Value)
            Next
            oDBs_Head.SetValue("U_Total", 0, TotalAmount)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function PostStockTransfer(ByVal FormUID As String) As Integer
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_DC_D0")
            Dim oStockTransfer As SAPbobsCOM.StockTransfer = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
            oCompany.StartTransaction()
            oStockTransfer.DocDate = DateTime.Today
            For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                If objMatrix.Columns.Item("IsCheck").Cells.Item(Row).Specific.Checked = True Then
                    If oStockTransfer.Lines.Count > 1 Then oStockTransfer.Lines.Add()
                    oStockTransfer.FromWarehouse = Trim(objMatrix.Columns.Item("fwhs").Cells.Item(Row).Specific.Value)
                    oStockTransfer.Lines.ItemCode = Trim(objMatrix.Columns.Item("itemno").Cells.Item(Row).Specific.Value)
                    oStockTransfer.Lines.WarehouseCode = Trim(objMatrix.Columns.Item("twhs").Cells.Item(Row).Specific.Value)
                    oStockTransfer.Lines.Quantity = Trim(objMatrix.Columns.Item("issqty").Cells.Item(Row).Specific.Value)
                    oStockTransfer.Lines.SetCurrentLine(oStockTransfer.Lines.Count - 1)
                End If
            Next
            If oStockTransfer.Add = 0 Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                Dim DocEntry As String = Trim(oCompany.GetNewObjectKey)
                Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRS.DoQuery("Select DocNum from OWTR Where DocEntry='" & DocEntry & "'")
                oDBs_Head.SetValue("U_InvTrDNo", 0, Trim(oRS.Fields.Item("DocNum").Value))
                Return CInt(DocEntry)
            Else
                oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription & " - [ Error Code : " & oCompany.GetLastErrorCode & " ]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                Return 0
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
            Return 0
        End Try
    End Function

    Sub PrintSC_DCReport()
        Try
            Dim oFile As New StreamReader(Application.StartupPath & "\DBLogin.ini", False)
            Dim s As String = ""
            Dim i As Integer = 1
            Dim Company = "", UserName = "", Password As String = ""
            s = oFile.ReadLine()
            While s <> ""
                Select Case i
                    Case 1
                        Company = s.Trim
                    Case 2
                        UserName = s.Trim
                    Case 3
                        Password = s.Trim
                End Select
                i = i + 1
                s = oFile.ReadLine
            End While
            Dim strcon As New SqlConnection("user id=" & UserName & ";data source=" & Company & ";pwd=" & Password & ";initial catalog=" & oCompany.CompanyDB & ";")
            strcon.Open()
            objForm = oApplication.Forms.ActiveForm
            Dim cmd As New SqlCommand("GEN_SEPL_PRC_DC", strcon)
            cmd.Connection = strcon
            cmd.CommandType = CommandType.StoredProcedure
            Dim oParameter As New SqlParameter("@DocNum", SqlDbType.NVarChar)
            oParameter.Value = Trim(objForm.Items.Item("t_docno").Specific.Value)
            Dim dsReport As DataSet = Helper.SqlHelper.ExecuteDataset(strcon, CommandType.StoredProcedure, "GEN_SEPL_PRC_DC", oParameter)
            dsReport.WriteXml(System.IO.Path.GetTempPath() & "GEN_SEPL_DC.xml", System.Data.XmlWriteMode.WriteSchema)
            oUtilities.ShowReport("GEN_SEPL_DC.rpt", "GEN_SEPL_DC.xml")
            strcon.Close()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


End Class

