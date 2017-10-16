'//Created by PRIYA on 25/05/2011


Imports System.Threading
Imports System.Globalization
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text

Public Class ClsSubContract_GRPO

#Region "        Declaration        "

    Dim oUtilities As New ClsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim objItem, objOldItem, TempItem As SAPbouiCOM.Item
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oDBs_DetailRM As SAPbouiCOM.DBDataSource
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim ROW_ID As Integer = 0
    Dim ITEM As String = ""
    Dim whs As String
    Dim taxcode As String
    Dim DelItemCode, DelLineID As String
#End Region


    Sub CreateForm()
        Try
            oUtilities.SAPXML("SubContracting_GRPO.xml")
            objForm = oApplication.Forms.GetForm("GEN_SCGRPO", oApplication.Forms.ActiveForm.TypeCount)
            objForm.Items.Item("cardcode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("c_series").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("t_docdt").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("t_docno").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("MIssue").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D0")
            objForm.Select()
            objForm.Items.Item("TabFG").AffectsFormMode = False
            objForm.Items.Item("TabRM").AffectsFormMode = False

            TempItem = objForm.Items.Item("59")
            objItem = objForm.Items.Add("gate", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = TempItem.Top + TempItem.Height + 20
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.Caption = "GateEntry"
            objItem.LinkTo = "59"
            objItem.Visible = False
            'TempItem = objForm.Items.Item("GRDocNo")
            objItem = objForm.Items.Add("gateno", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = TempItem.Top + TempItem.Height + 20
            objItem.Left = TempItem.Left + 80
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.databind.setbound(True, "@GEN_SC_GRPO", "U_GateEntrNo")
            objItem.Visible = False
            'objItem.Specific.TabOrder = TempItem.Specific.TabOrder + 1
            ' objItem.LinkTo = "GRDocNo"

            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Select SlpCode,SlpName from OSLP")
            objCombo = objForm.Items.Item("buyer").Specific
            For i As Integer = 1 To oRS.RecordCount
                objCombo.ValidValues.Add(Trim(oRS.Fields.Item("SlpCode").Value), Trim(oRS.Fields.Item("SlpName").Value))
                oRS.MoveNext()
            Next

            oRS.DoQuery("Select GroupNum,pymntGroup from OCTG")
            objCombo = objForm.Items.Item("paytrms").Specific
            objCombo.ValidValues.Add("", "")
            For i As Integer = 1 To oRS.RecordCount
                objCombo.ValidValues.Add(Trim(oRS.Fields.Item("GroupNum").Value), Trim(oRS.Fields.Item("pymntGroup").Value))
                oRS.MoveNext()
            Next

            SetDefault(objForm.UniqueID)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim folderDN As SAPbouiCOM.Folder
            folderDN = objForm.Items.Item("TabFG").Specific
            folderDN.Select()
            objForm.Freeze(True)
            objForm.EnableMenu("1282", False)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D0")
            oUtilities.GetSeries(FormUID, "c_series", "GEN_SC_GRPO")
            oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("c_series").Specific.Selected.Value, "GEN_SC_GRPO"))
            oDBs_Head.SetValue("U_Status", 0, "Open")
            oDBs_Head.SetValue("U_PostDate", 0, DateTime.Today.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_DocDate", 0, DateTime.Today.ToString("yyyyMMdd"))
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Select ISNULL(T0.firstName,'') + ' ' + ISNULL(T0.lastName,'') Owner,empid from OHEM T0 INNER JOIN OUSR T1 ON T0.UserID=T1.USERID Where USER_CODE='" & oCompany.UserName & "'")
            If oRS.RecordCount > 0 Then
                oDBs_Head.SetValue("U_Owner", 0, Trim(oRS.Fields.Item("Owner").Value))
                oDBs_Head.SetValue("U_OwnerCod", 0, Trim(oRS.Fields.Item("empid").Value))
            End If
            objCombo = objForm.Items.Item("buyer").Specific
            If objCombo.ValidValues.Count > 0 Then objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            objMatrix.Clear()
            objMatrix = objForm.Items.Item("RMMatrix").Specific
            objMatrix.Clear()

            'objMatrix.AddRow()
            'objMatrix.FlushToDataSource()
            'Me.SetNewLineFG(FormUID, objMatrix.VisualRowCount)

            'Dim objcombo As SAPbouiCOM.ButtonCombo
            'objcombo = objForm.Items.Item("copyto").Specific
            'objcombo.ValidValues.Add("Return", "Return")
            objForm.Items.Item("cardcode").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            objForm.Freeze(False)
        Catch ex As Exception
            'oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try
    End Sub

    Sub SetNewLineRM(ByVal FormUID As String, ByVal Row As Integer) 'for RM tab
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO")
            oDBs_DetailRM = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D1")
            Dim objMatrixRM As SAPbouiCOM.Matrix = objForm.Items.Item("RMMatrix").Specific
            oDBs_DetailRM.Offset = Row - 1
            oDBs_DetailRM.SetValue("LineId", oDBs_DetailRM.Offset, Row)
            oDBs_DetailRM.SetValue("U_Parent", oDBs_DetailRM.Offset, "")
            oDBs_DetailRM.SetValue("U_Line", oDBs_DetailRM.Offset, 0)
            oDBs_DetailRM.SetValue("U_ItemCode", oDBs_DetailRM.Offset, "")
            oDBs_DetailRM.SetValue("U_ItemName", oDBs_DetailRM.Offset, "")
            oDBs_DetailRM.SetValue("U_BOMQty", oDBs_DetailRM.Offset, 0)
            oDBs_DetailRM.SetValue("U_RecdQty", oDBs_DetailRM.Offset, 0)
            oDBs_DetailRM.SetValue("U_ItemQty", oDBs_DetailRM.Offset, 0)
            oDBs_DetailRM.SetValue("U_Whs", oDBs_DetailRM.Offset, "")
            oDBs_DetailRM.SetValue("U_WhsQty", oDBs_DetailRM.Offset, 0)
            oDBs_DetailRM.SetValue("U_ItemCost", oDBs_DetailRM.Offset, 0)
            oDBs_DetailRM.SetValue("U_TotCost", oDBs_DetailRM.Offset, 0)
            objMatrixRM.SetLineData(Row)
            objMatrixRM.FlushToDataSource()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLineFG(ByVal FormUID As String, ByVal Row As Integer) 'For FG Tab
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D0")
            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
            oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_RecdQty", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_POPrice", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_SerPrice", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_TaxRate", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_Whs", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_Line", oDBs_Detail.Offset, "")
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
                        oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("c_series").Specific.Selected.Value, "GEN_SC_GRPO"))
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        '---> Vijeesh
                        Me.Refresh_RawMaterial(FormUID)
                        If Me.Validation(FormUID) = False Then BubbleEvent = False
                    ElseIf pVal.ItemUID = "1" And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        Dim Docnum As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Docnum.DoQuery("Select docnum from [@GEN_SC_GRPO] Where DocEntry=(Select Max(DocEntry) From [@GEN_SC_GRPO])")
                        objForm.Items.Item("t_docno").Specific.value = Docnum.Fields.Item(0).Value
                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        oCompany.StartTransaction()
                        Dim iGoodsIssue As Integer = PostGoodsIssue(FormUID)
                        Dim iGoodsReceipt As Integer = PostGoodsReceipt(FormUID)
                        If iGoodsReceipt <> 0 And iGoodsIssue <> 0 Then
                            objForm.Items.Item("GINO").Specific.Value = iGoodsIssue
                            objForm.Items.Item("GRNO").Specific.Value = iGoodsReceipt
                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        Else
                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            ' Return False
                        End If
                        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRS.DoQuery("UPDATE OIGE SET U_Type='SubCont_GRPO',U_DocNum='" & Trim(oDBs_Head.GetValue("DocNum", 0)) & "' Where DocEntry='" & Trim(oDBs_Head.GetValue("U_GINO", 0)) & "'")
                        oRS.DoQuery("UPDATE OIGN SET U_Type='SubCont_GRPO',U_DocNum='" & Trim(oDBs_Head.GetValue("DocNum", 0)) & "' Where DocEntry='" & Trim(oDBs_Head.GetValue("U_GRNO", 0)) & "'")
                        Dim DocEntry As String = Trim(oCompany.GetNewObjectKey)
                        Dim oRS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRS.DoQuery("Select DocNum from OIGE Where DocEntry=(Select Max(DocEntry) From OIGE)")
                        oDBs_Head.SetValue("U_GIDocNO", 0, Trim(oRS.Fields.Item("DocNum").Value))
                        oRS1.DoQuery("Update [@GEN_SC_GRPO] set U_GIDocNO='" + Trim(oRS.Fields.Item("DocNum").Value) + "' Where DocEntry=(Select Max(DocEntry) From [@GEN_SC_GRPO])")

                        oRS.DoQuery("Select DocNum from OIGN Where DocEntry=(Select Max(DocEntry) From OIGN)")
                        oDBs_Head.SetValue("U_GRDocNO", 0, Trim(oRS.Fields.Item("DocNum").Value))
                        oRS1.DoQuery("Update [@GEN_SC_GRPO] set U_GRDocNO='" + Trim(oRS.Fields.Item("DocNum").Value) + "' Where DocEntry=(Select Max(DocEntry) From [@GEN_SC_GRPO])")

                        '  Me.SetDefault(FormUID)
                    End If


                    If pVal.ItemUID = "MIssue" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE) Then
                        Dim objMatrixRM As SAPbouiCOM.Matrix
                        objMatrixRM = objForm.Items.Item("RMMatrix").Specific
                        'For i As Integer = 1 To objMatrixRM.VisualRowCount
                        '    Dim Instock As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '    Instock.DoQuery("Select sum(B.OnHand) From OITM A Inner join OITW B on A.ItemCode=B.ItemCode where A.Itemcode='" + objMatrixRM.Columns.Item("ItemCode").Cells.Item(i).Specific.value + "' and B.Whscode='" + objMatrixRM.Columns.Item("Whs").Cells.Item(i).Specific.value + "'")
                        '    If Instock.Fields.Item(0).Value < objMatrixRM.Columns.Item("ItemQty").Cells.Item(i).Specific.value Then
                        '        oApplication.StatusBar.SetText("Quantity falls into negative inventory for item '" + objMatrixRM.Columns.Item("ItemCode").Cells.Item(i).Specific.value + "' in line '" + objMatrixRM.Columns.Item("SNo").Cells.Item(i).Specific.value + "'")
                        '        'Return False
                        '        Exit Sub
                        '    End If
                        'Next
                        Dim Check As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        If (objForm.Items.Item("GIDocNO").Specific.value = "" Or objForm.Items.Item("GRDocNO").Specific.value = "") Then
                            oCompany.StartTransaction()
                            Dim iGoodsIssue As Integer = PostGoodsIssue(FormUID)
                            Dim iGoodsReceipt As Integer = PostGoodsReceipt(FormUID)
                            If iGoodsReceipt <> 0 And iGoodsIssue <> 0 Then
                                objForm.Items.Item("GINO").Specific.Value = iGoodsIssue
                                objForm.Items.Item("GRNO").Specific.Value = iGoodsReceipt
                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            Else
                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                ' Return False
                            End If
                            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRS.DoQuery("UPDATE OIGE SET U_Type='SubCont_GRPO',U_DocNum='" & Trim(oDBs_Head.GetValue("DocNum", 0)) & "' Where DocEntry='" & Trim(oDBs_Head.GetValue("U_GINO", 0)) & "'")
                            oRS.DoQuery("UPDATE OIGN SET U_Type='SubCont_GRPO',U_DocNum='" & Trim(oDBs_Head.GetValue("DocNum", 0)) & "' Where DocEntry='" & Trim(oDBs_Head.GetValue("U_GRNO", 0)) & "'")
                            Dim DocEntry As String = Trim(oCompany.GetNewObjectKey)
                            Dim oRS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRS.DoQuery("Select DocNum from OIGE Where DocEntry=(Select Max(DocEntry) From OIGE)")
                            oDBs_Head.SetValue("U_GIDocNO", 0, Trim(oRS.Fields.Item("DocNum").Value))
                            oRS1.DoQuery("Update [@GEN_SC_GRPO] set U_GIDocNO='" + Trim(oRS.Fields.Item("DocNum").Value) + "' Where DocNum='" + objForm.Items.Item("t_docno").Specific.value + "'")

                            oRS.DoQuery("Select DocNum from OIGN Where DocEntry=(Select Max(DocEntry) From OIGN)")
                            oDBs_Head.SetValue("U_GRDocNO", 0, Trim(oRS.Fields.Item("DocNum").Value))
                            oRS1.DoQuery("Update [@GEN_SC_GRPO] set U_GRDocNO='" + Trim(oRS.Fields.Item("DocNum").Value) + "' Where DocNum='" + objForm.Items.Item("t_docno").Specific.value + "'")

                            '  Me.SetDefault(FormUID)
                        End If
                    End If



                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    If pVal.ItemUID = "t_docdt" And pVal.BeforeAction = False And pVal.CharPressed = 9 And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        objForm = oApplication.Forms.Item(pVal.FormUID)
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        If objMatrix.VisualRowCount > 0 Then
                            objMatrix.Columns.Item("ItemCode").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Exit Sub
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If pVal.ItemUID = "t_postdt" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        objForm = oApplication.Forms.Item(pVal.FormUID)
                        If Trim(objForm.Items.Item("t_postdt").Specific.Value).Equals("") = False Then
                            If (DateDiff(DateInterval.Day, DateTime.ParseExact(Trim(objForm.Items.Item("t_postdt").Specific.Value), "yyyyMMdd", Nothing), DateTime.Today)) < 0 Then
                                oApplication.StatusBar.SetText("Posting date varies from system date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        End If
                    ElseIf pVal.ItemUID = "t_deldt" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        objForm = oApplication.Forms.Item(pVal.FormUID)
                        If Trim(objForm.Items.Item("t_deldt").Specific.Value).Equals("") = False Then
                            If (DateDiff(DateInterval.Day, DateTime.ParseExact(Trim(objForm.Items.Item("t_postdt").Specific.Value), "yyyyMMdd", Nothing), DateTime.ParseExact(Trim(objForm.Items.Item("t_deldt").Specific.Value), "yyyyMMdd", Nothing))) < 0 Then
                                oApplication.StatusBar.SetText("Delivery date is before posting date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        End If
                    ElseIf pVal.ItemUID = "ItemMatrix" And pVal.ColUID = "recdqty" And pVal.Row > 0 And pVal.BeforeAction = True Then
                        objForm = oApplication.Forms.Item(pVal.FormUID)
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        If CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(pVal.Row).Specific.value) > CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.value) Then
                            oApplication.StatusBar.SetText("Quantity cannot be greater than " & CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.value), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                        End If
                    ElseIf pVal.ItemUID = "ItemMatrix" And pVal.ColUID = "ItemCode" And pVal.BeforeAction = True Then
                        objForm = oApplication.Forms.Item(pVal.FormUID)
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        If Trim(objForm.Items.Item("PONo").Specific.Value).Equals("") = True Then
                            oApplication.StatusBar.SetText("Subcontracting PO.No. is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                        End If
                    ElseIf pVal.ItemUID = "RMMatrix" And pVal.ColUID = "ItemQty" And pVal.BeforeAction = True Then
                        objForm = oApplication.Forms.Item(pVal.FormUID)
                        Dim objMatrixRM As SAPbouiCOM.Matrix = objForm.Items.Item("RMMatrix").Specific
                        If CDbl(objMatrixRM.Columns.Item("ItemQty").Cells.Item(pVal.Row).Specific.Value) < (CDbl(objMatrixRM.Columns.Item("BOMQty").Cells.Item(pVal.Row).Specific.Value) * CDbl(objMatrixRM.Columns.Item("RecdQty").Cells.Item(pVal.Row).Specific.Value)) Then
                            'oApplication.StatusBar.SetText("Item Qty. should not be less than " & (CDbl(objMatrixRM.Columns.Item("BOMQty").Cells.Item(pVal.Row).Specific.Value) * CDbl(objMatrixRM.Columns.Item("RecdQty").Cells.Item(pVal.Row).Specific.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            'BubbleEvent = False
                        ElseIf CDbl(objMatrixRM.Columns.Item("ItemQty").Cells.Item(pVal.Row).Specific.Value) > (CDbl(objMatrixRM.Columns.Item("WhsQty").Cells.Item(pVal.Row).Specific.Value)) Then
                            oApplication.StatusBar.SetText("Item Qty. should not be greater than Warehouse Quantity ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If pVal.ItemUID = "ItemMatrix" And pVal.ColUID = "UnitPrice" And pVal.Row > 0 Then
                        objForm = oApplication.Forms.Item(pVal.FormUID)
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D0")
                        oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                        oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemDesc").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_RecdQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_POPrice", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("POPrice").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_SerPrice", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("SerPrice").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("TaxCode").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_TaxRate", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, (CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(pVal.Row).Specific.Value) * CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(pVal.Row).Specific.Value)) * CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(pVal.Row).Specific.Value) / 100)
                        oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(pVal.Row).Specific.Value) * CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Whs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Whs").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("remarks").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Line", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Line").Cells.Item(pVal.Row).Specific.Value))
                        objMatrix.SetLineData(pVal.Row)
                        Me.CalculateTotal(FormUID)
                    ElseIf pVal.ItemUID = "ItemMatrix" And pVal.ColUID = "recdqty" And pVal.Row > 0 And pVal.BeforeAction = False And pVal.ActionSuccess = True Then
                        'ElseIf pVal.ItemUID = "ItemMatrix" And pVal.ColUID = "recdqty" And pVal.Row > 0 And pVal.BeforeAction = True Then
                        objForm = oApplication.Forms.Item(pVal.FormUID)
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        Dim RMMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("RMMatrix").Specific
                        oDBs_DetailRM = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D1")
                        'Do Validation on the Rawmaterial Matrix
                        Dim UnitPrice As Double = 0
                        For Row As Integer = 1 To RMMatrix.VisualRowCount
                            oDBs_DetailRM.Offset = Row - 1
                            If Trim(oDBs_DetailRM.GetValue("U_Line", oDBs_DetailRM.Offset)).Equals(Trim(objMatrix.Columns.Item("Line").Cells.Item(pVal.Row).Specific.value)) = True Then
                                oDBs_DetailRM.SetValue("LineId", oDBs_DetailRM.Offset, Trim(RMMatrix.Columns.Item("SNo").Cells.Item(Row).Specific.Value))
                                oDBs_DetailRM.SetValue("U_Parent", oDBs_DetailRM.Offset, Trim(RMMatrix.Columns.Item("Parent").Cells.Item(Row).Specific.Value))
                                oDBs_DetailRM.SetValue("U_Line", oDBs_DetailRM.Offset, Trim(RMMatrix.Columns.Item("Line").Cells.Item(Row).Specific.Value))
                                oDBs_DetailRM.SetValue("U_ItemCode", oDBs_DetailRM.Offset, Trim(RMMatrix.Columns.Item("ItemCode").Cells.Item(Row).Specific.Value))
                                oDBs_DetailRM.SetValue("U_ItemName", oDBs_DetailRM.Offset, Trim(RMMatrix.Columns.Item("ItemName").Cells.Item(Row).Specific.Value))
                                oDBs_DetailRM.SetValue("U_Whs", oDBs_DetailRM.Offset, Trim(RMMatrix.Columns.Item("Whs").Cells.Item(Row).Specific.Value))
                                oDBs_DetailRM.SetValue("U_BOMQty", oDBs_DetailRM.Offset, CDbl(RMMatrix.Columns.Item("BOMQty").Cells.Item(Row).Specific.Value))
                                oDBs_DetailRM.SetValue("U_WhsQty", oDBs_DetailRM.Offset, CDbl(RMMatrix.Columns.Item("WhsQty").Cells.Item(Row).Specific.Value))
                                oDBs_DetailRM.SetValue("U_RecdQty", oDBs_DetailRM.Offset, CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_DetailRM.SetValue("U_ItemQty", oDBs_DetailRM.Offset, CDbl(RMMatrix.Columns.Item("BOMQty").Cells.Item(Row).Specific.Value) * CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_DetailRM.SetValue("U_ItemCost", oDBs_DetailRM.Offset, CDbl(RMMatrix.Columns.Item("ItemCost").Cells.Item(Row).Specific.Value))
                                oDBs_DetailRM.SetValue("U_TotCost", oDBs_DetailRM.Offset, (CDbl(RMMatrix.Columns.Item("BOMQty").Cells.Item(Row).Specific.Value) * CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(pVal.Row).Specific.Value) * CDbl(RMMatrix.Columns.Item("ItemCost").Cells.Item(Row).Specific.Value)))
                                RMMatrix.SetLineData(Row)
                            End If
                        Next
                        For i As Integer = 1 To RMMatrix.VisualRowCount
                            oDBs_DetailRM.Offset = i - 1
                            If Trim(oDBs_DetailRM.GetValue("U_Line", oDBs_DetailRM.Offset)) = Trim(objMatrix.Columns.Item("Line").Cells.Item(pVal.Row).Specific.value) Then
                                UnitPrice = UnitPrice + CDbl(RMMatrix.Columns.Item("TotCost").Cells.Item(i).Specific.value)
                            End If
                        Next
                        'Setting Line Items to Parent Matrix
                        oDBs_Detail.Offset = pVal.Row - 1
                        oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                        oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemDesc").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_RecdQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_POPrice", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("POPrice").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_SerPrice", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(pVal.Row).Specific.Value) * CDbl(objMatrix.Columns.Item("POPrice").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("POPrice").Cells.Item(pVal.Row).Specific.Value) + (UnitPrice / CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(pVal.Row).Specific.value)))
                        oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("TaxCode").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_TaxRate", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, (CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(pVal.Row).Specific.Value) * (CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(pVal.Row).Specific.Value))))
                        oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, (CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(pVal.Row).Specific.Value) * (CDbl(objMatrix.Columns.Item("POPrice").Cells.Item(pVal.Row).Specific.Value) + UnitPrice)) * CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(pVal.Row).Specific.Value) / 100)
                        oDBs_Detail.SetValue("U_Whs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Whs").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("remarks").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Line", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Line").Cells.Item(pVal.Row).Specific.Value))
                        objMatrix.SetLineData(pVal.Row)
                        Me.CalculateTotal(FormUID)
                    ElseIf pVal.ItemUID = "RMMatrix" And pVal.ColUID = "ItemQty" And pVal.Row > 0 And pVal.BeforeAction = False Then
                        objForm = oApplication.Forms.Item(pVal.FormUID)
                        'RAW Material Item Quantity Change
                        Dim RMMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("RMMatrix").Specific
                        oDBs_DetailRM = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D1")
                        oDBs_DetailRM.Offset = pVal.Row - 1
                        oDBs_DetailRM.SetValue("LineId", oDBs_DetailRM.Offset, Trim(RMMatrix.Columns.Item("SNo").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_DetailRM.SetValue("U_Parent", oDBs_DetailRM.Offset, Trim(RMMatrix.Columns.Item("Parent").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_DetailRM.SetValue("U_Line", oDBs_DetailRM.Offset, Trim(RMMatrix.Columns.Item("Line").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_DetailRM.SetValue("U_ItemCode", oDBs_DetailRM.Offset, Trim(RMMatrix.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_DetailRM.SetValue("U_ItemName", oDBs_DetailRM.Offset, Trim(RMMatrix.Columns.Item("ItemName").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_DetailRM.SetValue("U_Whs", oDBs_DetailRM.Offset, Trim(RMMatrix.Columns.Item("Whs").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_DetailRM.SetValue("U_BOMQty", oDBs_DetailRM.Offset, CDbl(RMMatrix.Columns.Item("BOMQty").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_DetailRM.SetValue("U_WhsQty", oDBs_DetailRM.Offset, CDbl(RMMatrix.Columns.Item("WhsQty").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_DetailRM.SetValue("U_RecdQty", oDBs_DetailRM.Offset, CDbl(RMMatrix.Columns.Item("RecdQty").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_DetailRM.SetValue("U_ItemQty", oDBs_DetailRM.Offset, CDbl(RMMatrix.Columns.Item("ItemQty").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_DetailRM.SetValue("U_ItemCost", oDBs_DetailRM.Offset, CDbl(RMMatrix.Columns.Item("ItemCost").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_DetailRM.SetValue("U_TotCost", oDBs_DetailRM.Offset, (CDbl(RMMatrix.Columns.Item("ItemCost").Cells.Item(pVal.Row).Specific.Value) * (CDbl(RMMatrix.Columns.Item("ItemQty").Cells.Item(pVal.Row).Specific.Value))))
                        RMMatrix.SetLineData(pVal.Row)
                        Dim ROW As Integer = CInt(RMMatrix.Columns.Item("Line").Cells.Item(pVal.Row).Specific.Value)

                        Dim UnitPrice As Double = 0
                        For i As Integer = 1 To RMMatrix.VisualRowCount
                            If Trim(RMMatrix.Columns.Item("Line").Cells.Item(i).Specific.Value) = ROW Then
                                UnitPrice = UnitPrice + CDbl(RMMatrix.Columns.Item("TotCost").Cells.Item(i).Specific.Value)
                            End If
                        Next

                        'Setting Line Items to Parent Matrix
                        For k As Integer = 1 To objMatrix.VisualRowCount
                            If Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(k).Specific.value) <> "" Then
                                oDBs_Detail.Offset = k - 1
                                If Trim(objMatrix.Columns.Item("Line").Cells.Item(k).Specific.Value) = ROW.ToString Then
                                    oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, ROW)
                                    oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(ROW).Specific.Value))
                                    oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemDesc").Cells.Item(ROW).Specific.Value))
                                    oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(ROW).Specific.Value))
                                    oDBs_Detail.SetValue("U_RecdQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(ROW).Specific.Value))
                                    oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(ROW).Specific.Value))
                                    oDBs_Detail.SetValue("U_POPrice", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("POPrice").Cells.Item(ROW).Specific.Value))
                                    oDBs_Detail.SetValue("U_SerPrice", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(ROW).Specific.Value) * CDbl(objMatrix.Columns.Item("POPrice").Cells.Item(ROW).Specific.Value))
                                    oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("POPrice").Cells.Item(ROW).Specific.Value) + (UnitPrice / CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(ROW).Specific.value)))
                                    oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("TaxCode").Cells.Item(ROW).Specific.Value))
                                    oDBs_Detail.SetValue("U_TaxRate", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(ROW).Specific.Value))
                                    oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, (CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(ROW).Specific.Value) * (CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(ROW).Specific.Value)))) '+ UnitPrice
                                    oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, (CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(ROW).Specific.Value) * (CDbl(objMatrix.Columns.Item("POPrice").Cells.Item(ROW).Specific.Value) + UnitPrice)) * CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(ROW).Specific.Value) / 100)
                                    oDBs_Detail.SetValue("U_Whs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Whs").Cells.Item(ROW).Specific.Value))
                                    oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("remarks").Cells.Item(ROW).Specific.Value))
                                    oDBs_Detail.SetValue("U_Line", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Line").Cells.Item(ROW).Specific.Value))
                                    objMatrix.SetLineData(ROW)
                                End If
                            End If
                        Next
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
                    If pVal.BeforeAction = True Then
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO")
                        If oCFL.UniqueID = "SCCFL" Then
                            If Trim(oDBs_Head.GetValue("U_PONo", 0)) = "" Then
                                oApplication.StatusBar.SetText("Please select Sub Contractor Order", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            Me.FilterDC(FormUID)
                        End If
                        If oCFL.UniqueID = "ITEM_CFL" Then
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select DocNum From [@GEN_SUB_CONTRACT] Where DocNum = '" + Trim(objForm.Items.Item("PONo").Specific.value) + "' And IsNull(u_manual,'N') = 'Y'")
                            If oRSet.RecordCount > 0 Then
                                If Trim(oDBs_Head.GetValue("U_scdcno", 0)) = "" Then
                                    oApplication.StatusBar.SetText("Please enter Sun Contractor DC No", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO")
                        oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D0")
                        oDBs_DetailRM = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D1")
                        If oCFL.UniqueID = "VendCFL" Then
                            oDBs_Head.SetValue("U_CardCode", 0, oDT.GetValue("CardCode", 0))
                            oDBs_Head.SetValue("U_CardName", 0, oDT.GetValue("CardName", 0))
                            oDBs_Head.SetValue("U_DelDate", 0, DateTime.Today.ToString("yyyyMMdd"))
                            oDBs_Head.SetValue("U_JourRem", 0, "SubContract GRN - " + oDT.GetValue("CardCode", 0))
                            oDBs_Head.SetValue("U_PONo", 0, "")
                            oDBs_Head.SetValue("U_PODate", 0, "")
                            oDBs_Head.SetValue("U_VendRef", 0, "")
                            oDBs_Head.SetValue("U_VendWhs", 0, "")
                            oDBs_Head.SetValue("U_Buyer", 0, "")
                            oDBs_Head.SetValue("U_PayTrms", 0, "")
                            oDBs_Head.SetValue("U_JourRem", 0, "")
                            oDBs_Head.SetValue("U_TotBefTa", 0, 0)
                            oDBs_Head.SetValue("U_Total", 0, 0)
                            oDBs_Head.SetValue("U_Tax", 0, 0)
                            oDBs_Head.SetValue("U_Remarks", 0, "")
                            oDBs_Head.SetValue("U_DCNo", 0, "")
                            oDBs_Head.SetValue("U_DCDate", 0, "")
                            oDBs_Head.SetValue("U_PONo", 0, "")
                            oDBs_Head.SetValue("U_PODate", 0, "")
                            oDBs_Head.SetValue("U_GINO", 0, "")
                            oDBs_Head.SetValue("U_GRNO", 0, "")
                            oDBs_Head.SetValue("U_PayNum", 0, "")
                            oDBs_Head.SetValue("DocEntry", 0, "")
                            Me.FilterSC(FormUID)
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific '//Added for checking
                            objMatrix.Clear()
                            Dim objMatrixRM As SAPbouiCOM.Matrix = objForm.Items.Item("RMMatrix").Specific '//Added for checking
                            objMatrixRM.Clear()
                        ElseIf oCFL.UniqueID = "CFL_SO" Then
                            oDBs_Head.SetValue("U_SONo", 0, oDT.GetValue("DocEntry", 0))
                            oDBs_Head.SetValue("U_SODNo", 0, oDT.GetValue("DocNum", 0))
                        ElseIf oCFL.UniqueID = "CFL_SCNo" Then
                            oDBs_Head.SetValue("U_PODocNo", 0, oDT.GetValue("DocEntry", 0)) ' PO DocEntry
                            oDBs_Head.SetValue("U_PONo", 0, oDT.GetValue("DocNum", 0))
                            oDBs_Head.SetValue("U_PODate", 0, Format(oDT.GetValue("U_DocDate", 0), "yyyyMMdd"))
                            oDBs_Head.SetValue("U_VendRef", 0, oDT.GetValue("U_VendRef", 0))
                            oDBs_Head.SetValue("U_Buyer", 0, oDT.GetValue("U_Buyer", 0))
                            oDBs_Head.SetValue("U_ContPer", 0, oDT.GetValue("U_ContPer", 0))
                            oDBs_Head.SetValue("U_VendWhs", 0, oDT.GetValue("U_VendWhs", 0))
                            Me.FilterFG(FormUID)
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            objMatrix.Clear()
                            objMatrix.AddRow()
                            objMatrix.FlushToDataSource()
                            Me.SetNewLineFG(FormUID, objMatrix.VisualRowCount)
                            Me.CalculateTotal(FormUID)

                            Dim objMatrixRM As SAPbouiCOM.Matrix = objForm.Items.Item("RMMatrix").Specific
                            objMatrixRM.Clear()
                            oDBs_DetailRM.Clear()

                            'Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRS.DoQuery("select b.LineID,b.U_ItemCode ItemCode,b.U_ItemDesc ItemDesc,ISNULL(b.U_Quantity,0)-ISNULL(b.U_GRNQty,0) Quantity,b.U_UOM UoM,b.U_Quantity POQty " _
                            '& " ,b.U_Price Price,b.U_TotalLC TotalLC,b.U_TaxRate TaxRate,b.U_TaxAmt TaxAmount,b.U_TaxCode TaxCode,b.U_Whs WhsCode,b.U_Remarks Remarks from  " _
                            '& " [@GEN_SUB_CONTRACT] a join  [@GEN_SUB_CONTRACT_D0] b on a.docentry=b.docentry   where a.DocEntry='" & Trim(oDT.GetValue("DocEntry", 0)) & "' and (ISNULL(b.U_Quantity,0)-ISNULL(b.U_GRNQty,0))>0")

                            'objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            'objMatrix.Clear()
                            'For Row As Integer = 1 To oRS.RecordCount
                            '    objMatrix.AddRow()
                            '    objMatrix.FlushToDataSource()
                            '    oDBs_Detail.Offset = Row - 1
                            '    oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
                            '    oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(oRS.Fields.Item("ItemCode").Value))
                            '    oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, Trim(oRS.Fields.Item("ItemDesc").Value))
                            '    oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("Quantity").Value))
                            '    oDBs_Detail.SetValue("U_RecdQty", oDBs_Detail.Offset, 0)
                            '    oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(oRS.Fields.Item("UoM").Value))
                            '    oDBs_Detail.SetValue("U_POPrice", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("Price").Value))
                            '    oDBs_Detail.SetValue("U_SerPrice", oDBs_Detail.Offset, 0)
                            '    oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, 0)
                            '    oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(oRS.Fields.Item("TaxCode").Value))
                            '    oDBs_Detail.SetValue("U_TaxRate", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("TaxRate").Value))
                            '    oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, 0)
                            '    oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, 0)
                            '    oDBs_Detail.SetValue("U_Whs", oDBs_Detail.Offset, Trim(oRS.Fields.Item("WhsCode").Value))
                            '    oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(oRS.Fields.Item("Remarks").Value))
                            '    oDBs_Detail.SetValue("U_Line", oDBs_Detail.Offset, Trim(oRS.Fields.Item("LineID").Value))
                            '    objMatrix.SetLineData(Row)
                            '    oRS.MoveNext()
                            'Next
                            'objMatrix.FlushToDataSource()
                            'objMatrix.AddRow()
                            'objMatrix.FlushToDataSource()
                            'objMatrix.AutoResizeColumns()
                            'Me.SetNewLineFG(FormUID, objMatrix.VisualRowCount)
                            'Me.CalculateTotal(FormUID)

                            ''//For the RM tab
                            ''oRS.DoQuery("Select c.code,d.itemname,c.quantity,(c.quantity*b.U_quantity) as TotQty,c.Father,a.U_vendwhs,(select z.onhand from oitw z where z.whscode=a.U_vendwhs and z.ItemCode=c.code) as whsqty,(select z.AvgPrice from oitw z where z.whscode=a.U_vendwhs and z.ItemCode=c.code) as Itemcost from   [@GEN_SUB_CONTRACT] a join  [@GEN_SUB_CONTRACT_D0] b on a.docentry=b.docentry  join  ITT1 c on c.Father=b.U_ItemCode join OITM d on c.Code=d.ItemCode where a.DocNum='" + objForm.Items.Item("PONo").Specific.value + "' order by b.LineId")
                            'oRS.DoQuery("Select c.Code,d.ItemName,c.Quantity,(c.Quantity*b.U_Quantity) as TotQty,c.Father,a.U_Vendwhs VendWhs,e.OnHand  WhsQty,e.AvgPrice ItemCost,b.LineID from  " _
                            '& " [@GEN_SUB_CONTRACT] a join  [@GEN_SUB_CONTRACT_D0] b on a.docentry=b.docentry  inner join ITT1 c on c.Father=b.U_ItemCode " _
                            '& " inner join OITM d on c.Code=d.ItemCode inner join OITW e on e.whscode=a.U_VendWhs and e.ItemCode=c.Code where a.DocEntry='" & Trim(oDT.GetValue("DocEntry", 0)) & "' and d.EvalSystem<>'S' order by b.LineId")

                            'Dim objMatrixRM As SAPbouiCOM.Matrix = objForm.Items.Item("RMMatrix").Specific
                            'objMatrixRM.Clear()
                            'For Row As Integer = 1 To oRS.RecordCount
                            '    objMatrixRM.AddRow()
                            '    objMatrixRM.FlushToDataSource()
                            '    oDBs_DetailRM.Offset = Row - 1
                            '    oDBs_DetailRM.SetValue("LineId", oDBs_DetailRM.Offset, Row)
                            '    oDBs_DetailRM.SetValue("U_Parent", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("Father").Value))
                            '    oDBs_DetailRM.SetValue("U_Line", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("LineID").Value))
                            '    oDBs_DetailRM.SetValue("U_ItemCode", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("Code").Value))
                            '    oDBs_DetailRM.SetValue("U_ItemName", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("ItemName").Value))
                            '    oDBs_DetailRM.SetValue("U_BOMQty", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("Quantity").Value))
                            '    oDBs_DetailRM.SetValue("U_RecdQty", oDBs_DetailRM.Offset, 0)
                            '    oDBs_DetailRM.SetValue("U_ItemQty", oDBs_DetailRM.Offset, 0)
                            '    oDBs_DetailRM.SetValue("U_Whs", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("VendWhs").Value))
                            '    oDBs_DetailRM.SetValue("U_WhsQty", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("WhsQty").Value))
                            '    oDBs_DetailRM.SetValue("U_ItemCost", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("ItemCost").Value))
                            '    oDBs_DetailRM.SetValue("U_TotCost", oDBs_DetailRM.Offset, 0)
                            '    objMatrixRM.SetLineData(Row)
                            '    oRS.MoveNext()
                            'Next
                            'objMatrixRM.FlushToDataSource()
                            'objMatrixRM.AutoResizeColumns()

                        ElseIf oCFL.UniqueID = "CFL_Owner" Then
                            oDBs_Head.SetValue("U_OwnerCod", 0, oDT.GetValue("empID", 0))
                            oDBs_Head.SetValue("U_Owner", 0, oDT.GetValue("firstName", 0) + " " + oDT.GetValue("lastName", 0))
                        ElseIf oCFL.UniqueID = "SCCFL" Then
                            oDBs_Head.SetValue("U_scdcno", 0, oDT.GetValue("DocNum", 0))
                        ElseIf oCFL.UniqueID = "CFL_WHS1" Then
                            oDBs_Head.SetValue("U_VendWhs", 0, oDT.GetValue("WhsCode", 0))
                        ElseIf oCFL.UniqueID = "CFL_Whs" Then
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D0")
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            oDBs_Detail.Offset = pVal.Row - 1
                            oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                            oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemDesc").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_RecdQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_POPrice", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("POPrice").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_SerPrice", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("SerPrice").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("TaxCode").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_TaxRate", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxamt").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Total").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Whs", oDBs_Detail.Offset, Trim(oDT.GetValue("WhsCode", 0)))
                            oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("remarks").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Line", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Line").Cells.Item(pVal.Row).Specific.Value))
                            objMatrix.SetLineData(pVal.Row)
                            Me.CalculateTotal(FormUID)
                        ElseIf oCFL.UniqueID = "CFL_Tax" Then
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            oDBs_Detail.Offset = pVal.Row - 1
                            oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                            oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemDesc").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_RecdQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_POPrice", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("POPrice").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_SerPrice", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("SerPrice").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(oDT.GetValue("Code", 0)))
                            oDBs_Detail.SetValue("U_TaxRate", oDBs_Detail.Offset, CDbl(oDT.GetValue("Rate", 0)))
                            oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, (CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(pVal.Row).Specific.Value) * CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(pVal.Row).Specific.Value)) * CDbl(oDT.GetValue("Rate", 0)) / 100)
                            oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, (CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(pVal.Row).Specific.Value) * CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(pVal.Row).Specific.Value)))
                            oDBs_Detail.SetValue("U_Whs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Whs").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("remarks").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Line", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Line").Cells.Item(pVal.Row).Specific.Value))
                            objMatrix.SetLineData(pVal.Row)
                            Me.CalculateTotal(FormUID)
                        ElseIf oCFL.UniqueID = "ITEM_CFL" Then
                            Dim LineNo As Integer
                            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRS.DoQuery("select b.LineID,b.U_ItemCode ItemCode,b.U_ItemDesc ItemDesc,ISNULL(b.U_Quantity,0)-ISNULL(b.U_GRNQty,0) Quantity,b.U_UOM UoM,b.U_Quantity POQty " _
                             & " ,b.U_Price Price,b.U_TotalLC TotalLC,b.U_TaxRate TaxRate,b.U_TaxAmt TaxAmount,b.U_TaxCode TaxCode,b.U_Whs WhsCode,b.U_Remarks Remarks from  " _
                            & " [@GEN_SUB_CONTRACT] a join  [@GEN_SUB_CONTRACT_D0] b on a.docentry=b.docentry   where a.DocEntry='" & Trim(oDBs_Head.GetValue("U_PODocNo", 0)) & "' and b.U_ItemCode='" & Trim(oDT.GetValue("ItemCode", 0)) & "' and (ISNULL(b.U_Quantity,0)-ISNULL(b.U_GRNQty,0))>0")


                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            oDBs_Detail.Offset = pVal.Row - 1
                            oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                            oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(oRS.Fields.Item("ItemCode").Value))
                            oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, Trim(oRS.Fields.Item("ItemDesc").Value))
                            oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("Quantity").Value))
                            oDBs_Detail.SetValue("U_RecdQty", oDBs_Detail.Offset, 0)
                            oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(oRS.Fields.Item("UoM").Value))
                            oDBs_Detail.SetValue("U_POPrice", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("Price").Value))
                            oDBs_Detail.SetValue("U_SerPrice", oDBs_Detail.Offset, 0)
                            oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, 0)
                            oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(oRS.Fields.Item("TaxCode").Value))
                            oDBs_Detail.SetValue("U_TaxRate", oDBs_Detail.Offset, CDbl(oRS.Fields.Item("TaxRate").Value))
                            oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, 0)
                            oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, 0)
                            oDBs_Detail.SetValue("U_Whs", oDBs_Detail.Offset, Trim(oRS.Fields.Item("WhsCode").Value))
                            oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(oRS.Fields.Item("Remarks").Value))
                            oDBs_Detail.SetValue("U_Line", oDBs_Detail.Offset, Trim(oRS.Fields.Item("LineID").Value))
                            LineNo = oRS.Fields.Item("LineID").Value
                            objMatrix.SetLineData(pVal.Row)
                            objMatrix.FlushToDataSource()
                            objMatrix.AutoResizeColumns()
                            If objMatrix.VisualRowCount = pVal.Row Then
                                objMatrix.AddRow()
                                objMatrix.FlushToDataSource()
                                Me.SetNewLineFG(FormUID, objMatrix.VisualRowCount)
                            End If
                            Me.CalculateTotal(FormUID)
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select DocNum From [@GEN_SUB_CONTRACT] Where DocNum = '" + Trim(objForm.Items.Item("PONo").Specific.value) + "' And IsNull(u_manual,'N') = 'N' And IsNull(u_cstbom,'N') = 'N' And IsNull(U_manwobom,'N') = 'N'")
                            If oRSet.RecordCount > 0 Then
                                '//For the RM tab
                                'oRS.DoQuery("Select c.code,d.itemname,c.quantity,(c.quantity*b.U_quantity) as TotQty,c.Father,a.U_vendwhs,(select z.onhand from oitw z where z.whscode=a.U_vendwhs and z.ItemCode=c.code) as whsqty,(select z.AvgPrice from oitw z where z.whscode=a.U_vendwhs and z.ItemCode=c.code) as Itemcost from   [@GEN_SUB_CONTRACT] a join  [@GEN_SUB_CONTRACT_D0] b on a.docentry=b.docentry  join  ITT1 c on c.Father=b.U_ItemCode join OITM d on c.Code=d.ItemCode where a.DocNum='" + objForm.Items.Item("PONo").Specific.value + "' order by b.LineId")

                                'oRS.DoQuery("Select c.Code,d.ItemName,c.Quantity,(c.Quantity*b.U_Quantity) as TotQty,c.Father,a.U_Vendwhs VendWhs,e.OnHand  WhsQty,e.AvgPrice ItemCost,b.LineID from  " _
                                '& " [@GEN_SUB_CONTRACT] a join  [@GEN_SUB_CONTRACT_D0] b on a.docentry=b.docentry  inner join ITT1 c on c.Father=b.U_ItemCode " _
                                '& " inner join OITM d on c.Code=d.ItemCode inner join OITW e on e.whscode=a.U_VendWhs and e.ItemCode=c.Code where a.DocEntry='" & Trim(oDBs_Head.GetValue("U_PODocNo", 0)) & "' and b.U_ItemCode='" & Trim(oDT.GetValue("ItemCode", 0)) & "' and d.EvalSystem<>'S' order by b.LineId")

                                Dim _strQry As String = "Select c.U_Code Code,d.ItemName,c.U_BOMQty Quantity ,(c.U_BOMQty*b.U_Quantity) as TotQty, " _
                                                    & "c.U_Father Father,a.U_Vendwhs VendWhs,e.OnHand  WhsQty,e.AvgPrice ItemCost,b.LineID  " _
                                                    & "from   [@GEN_SUB_CONTRACT] a join  [@GEN_SUB_CONTRACT_D0] b on a.docentry=b.docentry " _
                                                    & "inner join [@GEN_SUB_CONTRACT_D1]c on c.DocEntry = b.DocEntry " _
                                                    & "inner join OITM d on c.U_Code =d.ItemCode " _
                                                    & "inner join OITW e on e.whscode=a.U_VendWhs and e.ItemCode=c.U_Code " _
                                                    & "where a.DocEntry='" & Trim(oDBs_Head.GetValue("U_PODocNo", 0)) & "' and b.U_ItemCode='" & Trim(oDT.GetValue("ItemCode", 0)) & "'  " _
                                                    & "and d.EvalSystem<>'S' order by b.LineId"
                                oRS.DoQuery(_strQry)

                                Dim objMatrixRM As SAPbouiCOM.Matrix = objForm.Items.Item("RMMatrix").Specific
                                objMatrixRM.FlushToDataSource()
                                objMatrixRM.Clear()
                                oDBs_DetailRM.Clear()
                                For Row As Integer = 0 To oRS.RecordCount - 1
                                    oDBs_DetailRM.InsertRecord(oDBs_DetailRM.Size)
                                    oDBs_DetailRM.Offset = oDBs_DetailRM.Size - 1
                                    oDBs_DetailRM.SetValue("LineId", oDBs_DetailRM.Offset, oDBs_DetailRM.Size)
                                    oDBs_DetailRM.SetValue("U_Parent", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("Father").Value))
                                    oDBs_DetailRM.SetValue("U_Line", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("LineID").Value))
                                    oDBs_DetailRM.SetValue("U_ItemCode", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("Code").Value))
                                    oDBs_DetailRM.SetValue("U_ItemName", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("ItemName").Value))
                                    oDBs_DetailRM.SetValue("U_BOMQty", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("Quantity").Value))
                                    oDBs_DetailRM.SetValue("U_RecdQty", oDBs_DetailRM.Offset, 0)
                                    oDBs_DetailRM.SetValue("U_ItemQty", oDBs_DetailRM.Offset, 0)
                                    oDBs_DetailRM.SetValue("U_Whs", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("VendWhs").Value))
                                    oDBs_DetailRM.SetValue("U_WhsQty", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("WhsQty").Value))
                                    oDBs_DetailRM.SetValue("U_ItemCost", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("ItemCost").Value))
                                    oDBs_DetailRM.SetValue("U_TotCost", oDBs_DetailRM.Offset, 0)
                                    oRS.MoveNext()
                                Next
                                objMatrixRM.LoadFromDataSource()
                                objMatrixRM.AutoResizeColumns()
                            End If
                            oRSet.DoQuery("Select DocNum From [@GEN_SUB_CONTRACT] Where DocNum = '" + Trim(objForm.Items.Item("PONo").Specific.value) + "' And IsNull(u_manual,'N') = 'Y'")
                            If oRSet.RecordCount > 0 Then
                                oRS.DoQuery("Select A.U_ItemNo,B.U_ItemNo AS 'LITEM',B.u_desc,B.U_TWhs,B.U_IssQty,B.U_BOMQty,C.OnHand,C.AvgPrice From [@GEN_SC_DC] A INNER JOIN [@GEN_SC_DC_D0] B ON A.DocEntry = B.DocEntry Inner Join OITW C On B.u_TWhs = C.WhsCode And B.u_ItemNo = C.ItemCode Where DocNum = '" + Trim(objForm.Items.Item("scdcno").Specific.value) + "'")
                                Dim objMatrixRM As SAPbouiCOM.Matrix = objForm.Items.Item("RMMatrix").Specific
                                objMatrixRM.Clear()
                                For Row As Integer = 1 To oRS.RecordCount
                                    If objMatrix.VisualRowCount - 1 = pVal.Row Then
                                        oDBs_DetailRM.InsertRecord(oDBs_DetailRM.Size)
                                    End If
                                    oDBs_DetailRM.Offset = oDBs_DetailRM.Size - 1
                                    oDBs_DetailRM.SetValue("LineId", oDBs_DetailRM.Offset, oDBs_DetailRM.Size)
                                    oDBs_DetailRM.SetValue("U_Parent", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("U_ItemNo").Value))
                                    oDBs_DetailRM.SetValue("U_Line", oDBs_DetailRM.Offset, LineNo)
                                    oDBs_DetailRM.SetValue("U_ItemCode", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("LITEM").Value))
                                    oDBs_DetailRM.SetValue("U_ItemName", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("u_desc").Value))
                                    oDBs_DetailRM.SetValue("U_BOMQty", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("U_BOMQty").Value))
                                    oDBs_DetailRM.SetValue("U_RecdQty", oDBs_DetailRM.Offset, 0)
                                    oDBs_DetailRM.SetValue("U_ItemQty", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("u_IssQty").Value))
                                    oDBs_DetailRM.SetValue("U_Whs", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("U_TWhs").Value))
                                    oDBs_DetailRM.SetValue("U_WhsQty", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("OnHand").Value))
                                    oDBs_DetailRM.SetValue("U_ItemCost", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("AvgPrice").Value))
                                    oDBs_DetailRM.SetValue("U_TotCost", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("u_IssQty").Value) * CDbl(oRS.Fields.Item("AvgPrice").Value))
                                    oRS.MoveNext()
                                Next
                                objMatrixRM.LoadFromDataSource()
                                objMatrixRM.AutoResizeColumns()
                            End If
                            oRSet.DoQuery("Select DocNum From [@GEN_SUB_CONTRACT] Where DocNum = '" + Trim(objForm.Items.Item("PONo").Specific.value) + "' And IsNull(u_cstbom,'N') = 'Y'")
                            If oRSet.RecordCount > 0 Then
                                oRS.DoQuery("Select A.U_ItemNo,B.U_ItemNo AS 'LITEM',B.u_desc,B.u_BOMQty,B.U_TWhs,B.U_IssQty,C.OnHand,C.AvgPrice From [@GEN_SC_DC] A INNER JOIN [@GEN_SC_DC_D0] B ON A.DocEntry = B.DocEntry Inner Join OITW C On B.u_TWhs = C.WhsCode And B.u_ItemNo = C.ItemCode Where DocNum = '" + Trim(objForm.Items.Item("scdcno").Specific.value) + "'")
                                Dim objMatrixRM As SAPbouiCOM.Matrix = objForm.Items.Item("RMMatrix").Specific
                                objMatrixRM.Clear()
                                For Row As Integer = 1 To oRS.RecordCount
                                    oDBs_DetailRM.InsertRecord(oDBs_DetailRM.Size)
                                    oDBs_DetailRM.Offset = oDBs_DetailRM.Size - 1
                                    oDBs_DetailRM.SetValue("LineId", oDBs_DetailRM.Offset, oDBs_DetailRM.Size)
                                    oDBs_DetailRM.SetValue("U_Parent", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("U_ItemNo").Value))
                                    oDBs_DetailRM.SetValue("U_Line", oDBs_DetailRM.Offset, LineNo)
                                    oDBs_DetailRM.SetValue("U_ItemCode", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("LITEM").Value))
                                    oDBs_DetailRM.SetValue("U_ItemName", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("u_desc").Value))
                                    oDBs_DetailRM.SetValue("U_BOMQty", oDBs_DetailRM.Offset, oRS.Fields.Item("u_BOMQty").Value)
                                    oDBs_DetailRM.SetValue("U_RecdQty", oDBs_DetailRM.Offset, 0)
                                    oDBs_DetailRM.SetValue("U_ItemQty", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("u_IssQty").Value))
                                    oDBs_DetailRM.SetValue("U_Whs", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("U_TWhs").Value))
                                    oDBs_DetailRM.SetValue("U_WhsQty", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("OnHand").Value))
                                    oDBs_DetailRM.SetValue("U_ItemCost", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("AvgPrice").Value))
                                    oDBs_DetailRM.SetValue("U_TotCost", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("u_IssQty").Value) * CDbl(oRS.Fields.Item("AvgPrice").Value))
                                    oRS.MoveNext()
                                Next
                                objMatrixRM.LoadFromDataSource()
                                objMatrixRM.AutoResizeColumns()
                            End If
                            oRSet.DoQuery("Select DocNum From [@GEN_SUB_CONTRACT] Where DocNum = '" + Trim(objForm.Items.Item("PONo").Specific.value) + "' And IsNull(U_manwobom,'N') = 'Y'")
                            If oRSet.RecordCount > 0 Then
                                oRS.DoQuery("Select A.U_ItemNo,B.U_ItemNo AS 'LITEM',B.u_desc,B.u_BOMQty,B.LineId,B.U_TWhs,B.U_IssQty,C.OnHand,C.AvgPrice From [@GEN_SC_DC] A INNER JOIN [@GEN_SC_DC_D0] B ON A.DocEntry = B.DocEntry Inner Join OITW C On B.u_TWhs = C.WhsCode And B.u_ItemNo = C.ItemCode Where DocNum = '" + Trim(objForm.Items.Item("scdcno").Specific.value) + "'")
                                Dim objMatrixRM As SAPbouiCOM.Matrix = objForm.Items.Item("RMMatrix").Specific
                                objMatrixRM.Clear()
                                For Row As Integer = 1 To oRS.RecordCount
                                    oDBs_DetailRM.InsertRecord(oDBs_DetailRM.Size)
                                    oDBs_DetailRM.Offset = oDBs_DetailRM.Size - 1
                                    oDBs_DetailRM.SetValue("LineId", oDBs_DetailRM.Offset, oDBs_DetailRM.Size)
                                    oDBs_DetailRM.SetValue("U_Parent", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("U_ItemNo").Value))
                                    oDBs_DetailRM.SetValue("U_Line", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("LineId").Value))
                                    oDBs_DetailRM.SetValue("U_ItemCode", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("LITEM").Value))
                                    oDBs_DetailRM.SetValue("U_ItemName", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("u_desc").Value))
                                    oDBs_DetailRM.SetValue("U_BOMQty", oDBs_DetailRM.Offset, oRS.Fields.Item("u_BOMQty").Value)
                                    oDBs_DetailRM.SetValue("U_RecdQty", oDBs_DetailRM.Offset, 0)
                                    oDBs_DetailRM.SetValue("U_ItemQty", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("u_IssQty").Value))
                                    oDBs_DetailRM.SetValue("U_Whs", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("U_TWhs").Value))
                                    oDBs_DetailRM.SetValue("U_WhsQty", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("OnHand").Value))
                                    oDBs_DetailRM.SetValue("U_ItemCost", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("AvgPrice").Value))
                                    oDBs_DetailRM.SetValue("U_TotCost", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("u_IssQty").Value) * CDbl(oRS.Fields.Item("AvgPrice").Value))
                                    oRS.MoveNext()
                                Next
                                objMatrixRM.LoadFromDataSource()
                                objMatrixRM.AutoResizeColumns()
                            End If
                        End If
                    End If
            End Select

            If pVal.ItemUID = "TabFG" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.BeforeAction = False Then
                objForm = oApplication.Forms.Item(FormUID)
                objForm.PaneLevel = 1
            End If

            If pVal.ItemUID = "TabRM" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.BeforeAction = False Then
                objForm = oApplication.Forms.Item(FormUID)
                objForm.PaneLevel = 2
            End If
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


    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = True Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "519"
                        Try
                            If objForm.TypeEx = "GEN_SCGRPO" Then
                                BubbleEvent = False
                                sDocNum = objForm.Items.Item("t_docno").Specific.Value
                                sRptName = "SubContract_GRPO.rpt"
                                Me.Report1()
                            End If
                        Catch ex As Exception

                        End Try
                End Select
            ElseIf pVal.BeforeAction = False Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "SC_GRPO"
                        If pVal.BeforeAction = False Then
                            Me.CreateForm()
                        End If
                    Case "1282"
                        If objForm.TypeEx = "GEN_SCGRPO" Then
                            Me.SetDefault(objForm.UniqueID)
                        End If
                    Case "1281"
                        If objForm.TypeEx = "GEN_SCGRPO" Then
                            objForm.EnableMenu("1282", True)
                            objForm.Items.Item("t_docno").Click()
                        End If
                    Case "Close"
                        If objForm.TypeEx = "GEN_SCGRPO" Then
                            If oApplication.MessageBox("Do you want to close?", 2, "Ok", "Cancel") = 1 Then
                                Dim ORS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                ORS.DoQuery("UPDATE [@GEN_SC_GRPO] SET U_Status='Closed' Where DocNum='" & oDBs_Head.GetValue("DocNum", 0) & "'")
                                oDBs_Head.SetValue("U_Status", 0, "Closed")
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                objForm.Items.Item("1").Enabled = True
                            End If
                        End If
                    Case "1293"
                        If objForm.TypeEx = "GEN_SCGRPO" Then
                            Try
                                If ITEM.Equals("ItemMatrix") = True Then
                                    objForm.Freeze(True)
                                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO")
                                    oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D0")
                                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                                    For Row As Integer = 1 To objMatrix.VisualRowCount
                                        objMatrix.GetLineData(Row)
                                        oDBs_Detail.Offset = Row - 1
                                        oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
                                        oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemDesc").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Quantity").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Total").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_TaxRate", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxamt").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("TaxCode").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_Whs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Whs").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("remarks").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_RecdQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_POPrice", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("POPrice").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_SerPrice", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("SerPrice").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_Line", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Line").Cells.Item(Row).Specific.Value))
                                        objMatrix.SetLineData(Row)
                                    Next
                                    objMatrix.FlushToDataSource()
                                    oDBs_Detail.RemoveRecord(oDBs_Detail.Size - 1)
                                    objMatrix.LoadFromDataSource()

                                    '//Removing corresponding rows from RMMatrix
                                    oDBs_DetailRM = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D1")
                                    Dim objMatrixRM As SAPbouiCOM.Matrix = objForm.Items.Item("RMMatrix").Specific
                                    For Row As Integer = 0 To oDBs_DetailRM.Size - 1
                                        If oDBs_DetailRM.Size >= Row + 1 Then
                                            oDBs_DetailRM.Offset = Row
                                            oDBs_DetailRM.SetValue("LineId", oDBs_DetailRM.Offset, oDBs_DetailRM.Offset + 1)
                                            If CInt(oDBs_DetailRM.GetValue("U_Line", oDBs_DetailRM.Offset)) = DelLineID Then
                                                'objMatrixRM.FlushToDataSource()
                                                oDBs_DetailRM.RemoveRecord(Row)
                                                'objMatrixRM.LoadFromDataSource()
                                                Row = Row - 1
                                            End If
                                        End If
                                    Next
                                    objMatrixRM.LoadFromDataSource()
                                    For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                                        For i As Integer = 0 To oDBs_DetailRM.Size - 1
                                            oDBs_DetailRM.Offset = i
                                            If Trim(oDBs_DetailRM.GetValue("U_Parent", oDBs_DetailRM.Offset)) = DelItemCode Then
                                                oDBs_DetailRM.SetValue("U_Line", oDBs_DetailRM.Offset, Row)
                                            End If
                                        Next
                                    Next
                                    objMatrixRM.LoadFromDataSource()
                                    Me.CalculateTotal(objForm.UniqueID)
                                    objForm.Freeze(False)
                                ElseIf ITEM.Equals("RMMatrix") = True Then
                                    Dim objMatrixRM As SAPbouiCOM.Matrix = objForm.Items.Item("RMMatrix").Specific
                                    oDBs_DetailRM = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D1")
                                    For Row As Integer = 1 To objMatrixRM.VisualRowCount
                                        objMatrixRM.GetLineData(Row)
                                        oDBs_DetailRM.Offset = Row - 1
                                        oDBs_DetailRM.SetValue("LineId", oDBs_DetailRM.Offset, Row)
                                        oDBs_DetailRM.SetValue("U_Parent", oDBs_DetailRM.Offset, Trim(objMatrixRM.Columns.Item("Parent").Cells.Item(Row).Specific.Value))
                                        oDBs_DetailRM.SetValue("U_Line", oDBs_DetailRM.Offset, Trim(objMatrixRM.Columns.Item("Line").Cells.Item(Row).Specific.Value))
                                        oDBs_DetailRM.SetValue("U_ItemCode", oDBs_DetailRM.Offset, Trim(objMatrixRM.Columns.Item("ItemCode").Cells.Item(Row).Specific.Value))
                                        oDBs_DetailRM.SetValue("U_ItemName", oDBs_DetailRM.Offset, Trim(objMatrixRM.Columns.Item("ItemName").Cells.Item(Row).Specific.Value))
                                        oDBs_DetailRM.SetValue("U_Whs", oDBs_DetailRM.Offset, Trim(objMatrixRM.Columns.Item("Whs").Cells.Item(Row).Specific.Value))
                                        oDBs_DetailRM.SetValue("U_WhsQty", oDBs_DetailRM.Offset, CDbl(objMatrixRM.Columns.Item("WhsQty").Cells.Item(Row).Specific.Value))
                                        oDBs_DetailRM.SetValue("U_BOMQty", oDBs_DetailRM.Offset, CDbl(objMatrixRM.Columns.Item("BOMQty").Cells.Item(Row).Specific.Value))
                                        oDBs_DetailRM.SetValue("U_RecdQty", oDBs_DetailRM.Offset, CDbl(objMatrixRM.Columns.Item("RecdQty").Cells.Item(Row).Specific.Value))
                                        oDBs_DetailRM.SetValue("U_ItemQty", oDBs_DetailRM.Offset, CDbl(objMatrixRM.Columns.Item("ItemQty").Cells.Item(Row).Specific.Value))
                                        oDBs_DetailRM.SetValue("U_ItemCost", oDBs_DetailRM.Offset, CDbl(objMatrixRM.Columns.Item("ItemCost").Cells.Item(Row).Specific.Value))
                                        oDBs_DetailRM.SetValue("U_TotCost", oDBs_DetailRM.Offset, CDbl(objMatrixRM.Columns.Item("TotCost").Cells.Item(Row).Specific.Value))
                                        objMatrixRM.SetLineData(Row)
                                    Next
                                    objMatrixRM.FlushToDataSource()
                                    oDBs_DetailRM.RemoveRecord(oDBs_DetailRM.Size - 1)
                                    objMatrixRM.LoadFromDataSource()

                                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                                    For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                                        Dim UnitPrice As Double = 0
                                        For i As Integer = 1 To objMatrixRM.VisualRowCount
                                            If CInt(objMatrixRM.Columns.Item("Line").Cells.Item(i).Specific.Value) = Row Then
                                                UnitPrice = UnitPrice + (CDbl(objMatrixRM.Columns.Item("TotCost").Cells.Item(i).Specific.Value))
                                            End If
                                        Next
                                        oDBs_Detail.Offset = Row - 1
                                        oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
                                        oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemDesc").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_RecdQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_POPrice", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("POPrice").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_SerPrice", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(Row).Specific.Value) * CDbl(objMatrix.Columns.Item("POPrice").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("POPrice").Cells.Item(Row).Specific.Value) + (UnitPrice / CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(Row).Specific.value)))
                                        oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("TaxCode").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_TaxRate", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, (CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(Row).Specific.Value) * (CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(Row).Specific.Value)))) '+ UnitPrice
                                        oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, (CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(Row).Specific.Value) * (CDbl(objMatrix.Columns.Item("POPrice").Cells.Item(Row).Specific.Value) + UnitPrice)) * CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(Row).Specific.Value) / 100)
                                        oDBs_Detail.SetValue("U_Whs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Whs").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("remarks").Cells.Item(Row).Specific.Value))
                                        oDBs_Detail.SetValue("U_Line", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Line").Cells.Item(Row).Specific.Value))
                                        objMatrix.SetLineData(Row)
                                    Next
                                    objMatrixRM.LoadFromDataSource()
                                    Me.CalculateTotal(objForm.UniqueID)
                                    objForm.Freeze(False)
                                End If
                            Catch ex As Exception
                                objForm.Freeze(False)
                                oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try
                        End If
                End Select
            End If
        Catch ex As Exception

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

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Select Case BusinessObjectInfo.EventType
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                If BusinessObjectInfo.BeforeAction = True Then
                    objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO")
                    oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D0")
                    Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                        oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("c_series").Specific.Selected.Value, "GEN_SC_GRPO"))
                        oRS.DoQuery("UPDATE OIGE SET U_Type='SubCont_GRPO',U_DocNum='" & Trim(oDBs_Head.GetValue("DocNum", 0)) & "' Where DocEntry='" & Trim(oDBs_Head.GetValue("U_GINO", 0)) & "'")
                        oRS.DoQuery("UPDATE OIGN SET U_Type='SubCont_GRPO',U_DocNum='" & Trim(oDBs_Head.GetValue("DocNum", 0)) & "' Where DocEntry='" & Trim(oDBs_Head.GetValue("U_GRNO", 0)) & "'")
                        'GRN Quantity updated in PO
                        For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                            oRS.DoQuery("UPDATE [@GEN_SUB_CONTRACT_D0]  set U_GRNQty=ISNULL(U_GRNQty,0)+" & CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(Row).Specific.Value) & " where DocEntry='" & Trim(oDBs_Head.GetValue("U_PODocNo", 0)) & "' and U_ItemCode='" & Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(Row).Specific.Value) & "'")
                        Next
                        oRS.DoQuery("Select count(*) from [@GEN_SUB_CONTRACT_D0] Where DocEntry='" & Trim(oDBs_Head.GetValue("U_PODocNo", 0)) & "' and ISNULL(U_Quantity,0)>ISNULL(U_GRNQty,0)")
                        If CInt(oRS.Fields.Item(0).Value) = 0 Then
                            oRS.DoQuery("Update [@GEN_SUB_CONTRACT]  set U_Status='Closed' where DocEntry= '" & Trim(oDBs_Head.GetValue("U_PODocNo", 0)) & "'")
                        End If
                    End If


                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                    objMatrix.LoadFromDataSource()
                    If objMatrix.VisualRowCount <> 0 Then
                        oDBs_Detail.RemoveRecord(objMatrix.VisualRowCount - 1)
                        objMatrix.LoadFromDataSource()
                    End If
                ElseIf BusinessObjectInfo.ActionSuccess = True Then
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Then
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        objMatrix.AddRow()
                        objMatrix.FlushToDataSource()
                        Me.SetNewLineFG(BusinessObjectInfo.FormUID, objMatrix.VisualRowCount)
                    End If
                End If
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                If BusinessObjectInfo.ActionSuccess = True Then
                    objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    objForm.EnableMenu("1282", True)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO")
                    oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D0")
                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                    objMatrix.AddRow()
                    objMatrix.FlushToDataSource()
                    Me.SetNewLineFG(BusinessObjectInfo.FormUID, objMatrix.VisualRowCount)
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                    objForm.Items.Item("1").Enabled = True
                    objForm.Items.Item("TabFG").Enabled = True
                    objForm.Items.Item("TabRM").Enabled = True
                    If (objForm.Items.Item("GIDocNO").Specific.value = "" Or objForm.Items.Item("GRDocNO").Specific.value = "") Then
                        objForm.Items.Item("MIssue").Enabled = True
                    End If

                End If
        End Select
    End Sub

    Function PostGoodsIssue(ByVal FormUID As String) As Integer
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D1")
            Dim objMartixRM As SAPbouiCOM.Matrix
            objMartixRM = objForm.Items.Item("RMMatrix").Specific
            Dim oIssue As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
            Dim VendWhs As String = Trim(objForm.Items.Item("vendwhs").Specific.value)

            For Row As Integer = 1 To objMartixRM.VisualRowCount
                If Row > 1 Then oIssue.Lines.Add()
                oIssue.DocDate = DateTime.ParseExact(objForm.Items.Item("t_postdt").Specific.value, "yyyyMMdd", Nothing)
                oIssue.Lines.ItemCode = Trim(objMartixRM.Columns.Item("ItemCode").Cells.Item(Row).Specific.Value)
                oIssue.Lines.Quantity = Trim(objMartixRM.Columns.Item("ItemQty").Cells.Item(Row).Specific.Value) 'oRSRM.Fields.Item("U_ItemQty").Value
                oIssue.Lines.WarehouseCode = VendWhs
                oIssue.Lines.SetCurrentLine(Row - 1)
            Next
            Dim iEr As Integer
            iEr = oIssue.Add
            If iEr <> 0 Then
                oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return 0
            Else
                Dim DocEntry As String = Trim(oCompany.GetNewObjectKey)
                Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRS.DoQuery("Select DocNum from OIGE Where DocEntry='" & DocEntry & "'")
                oDBs_Head.SetValue("U_GIDocNO", 0, Trim(oRS.Fields.Item("DocNum").Value))
                Return CInt(DocEntry)
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
            Return 0
        End Try
    End Function

    Function PostGoodsReceipt(ByVal FormUID As String) As Integer
        Try

            objForm = oApplication.Forms.Item(FormUID)
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D0")
            Dim objMartixFG As SAPbouiCOM.Matrix
            Dim oReceipt As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
            objMartixFG = objForm.Items.Item("ItemMatrix").Specific
            oReceipt.DocDate = DateTime.ParseExact(objForm.Items.Item("t_postdt").Specific.value, "yyyyMMdd", Nothing)
            For i As Integer = 1 To objMartixFG.VisualRowCount - 1
                If i > 1 Then oReceipt.Lines.Add()
                oReceipt.Lines.ItemCode = Trim(objMartixFG.Columns.Item("ItemCode").Cells.Item(i).Specific.Value)
                oReceipt.Lines.Quantity = CDbl(objMartixFG.Columns.Item("recdqty").Cells.Item(i).Specific.Value)
                oReceipt.Lines.WarehouseCode = Trim(objMartixFG.Columns.Item("Whs").Cells.Item(i).Specific.Value)
                oReceipt.Lines.Price = Trim(objMartixFG.Columns.Item("UnitPrice").Cells.Item(i).Specific.Value)
                oReceipt.Lines.UnitPrice = Trim(objMartixFG.Columns.Item("UnitPrice").Cells.Item(i).Specific.Value)
                oReceipt.Lines.SetCurrentLine(i - 1)

                'BOM Items having negative quantity
                'oRS.DoQuery("Select T1.ItemCode,-T0.Quantity Quantity,T1.AvgPrice,(Select top 1 (case when DfltWH is null then (Select DfltWhs from oadm) else DfltWH end ) from OITM where itemcode=T0.Code)  DfltWH from ITT1 T0 INNER JOIN OITM T1 ON T0.Code=T1.ItemCode Where T0.Father='" & Trim(objMartixFG.Columns.Item("ItemCode").Cells.Item(i).Specific.Value) & "' and T1.EvalSystem='S'")
                'For j As Integer = 1 To oRS.RecordCount
                '    oReceipt.Lines.Add()
                '    oReceipt.Lines.ItemCode = Trim(oRS.Fields.Item("ItemCode").Value)
                '    oReceipt.Lines.Quantity = CDbl(oRS.Fields.Item("Quantity").Value) * CDbl(objMartixFG.Columns.Item("recdqty").Cells.Item(i).Specific.Value)
                '    oReceipt.Lines.WarehouseCode = Trim(oRS.Fields.Item("DfltWH").Value)
                '    oReceipt.Lines.Price = Trim(oRS.Fields.Item("AvgPrice").Value)
                '    oReceipt.Lines.UnitPrice = Trim(oRS.Fields.Item("AvgPrice").Value)
                '    oReceipt.Lines.SetCurrentLine(oReceipt.Lines.Count - 1)
                'Next
            Next

            Dim iEr As Integer
            iEr = oReceipt.Add
            If iEr <> 0 Then
                oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return 0
            Else
                Dim DocEntry As String = Trim(oCompany.GetNewObjectKey)
                oRS.DoQuery("Select DocNum from OIGN Where DocEntry='" & DocEntry & "'")
                oDBs_Head.SetValue("U_GRDocNO", 0, Trim(oRS.Fields.Item("DocNum").Value))
                Return CInt(DocEntry)
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
            Return 0
        End Try
    End Function

    Function Validation(ByVal FormUID As String) As Boolean
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D0")

            If Trim(objForm.Items.Item("cardcode").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("CardCode should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Trim(objForm.Items.Item("PONo").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Subcontracting PO.No. should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Trim(objForm.Items.Item("t_postdt").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Posting Date should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Trim(objForm.Items.Item("t_deldt").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Delivery Date should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Trim(objForm.Items.Item("t_docdt").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Document Date should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Trim(objForm.Items.Item("vendwhs").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Vendor Warehouse should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf (DateDiff(DateInterval.Day, DateTime.ParseExact(Trim(objForm.Items.Item("t_postdt").Specific.Value), "yyyyMMdd", Nothing), DateTime.ParseExact(Trim(objForm.Items.Item("t_deldt").Specific.Value), "yyyyMMdd", Nothing))) < 0 Then
                oApplication.StatusBar.SetText("Delivery date is before posting date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Trim(objForm.Items.Item("gateno").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Please Enter GateEntry Number", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Dim objMatrixRM As SAPbouiCOM.Matrix
            objMatrixRM = objForm.Items.Item("RMMatrix").Specific
            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            If objMatrix.VisualRowCount = 1 Then
                oApplication.StatusBar.SetText("No items defined", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                    If Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(Row).Specific.Value).Equals("") = True Then
                        oApplication.StatusBar.SetText("Row [ " & Row & " ] - ItemCode should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(Row).Specific.Value) <= 0 Then
                        oApplication.StatusBar.SetText("Row [ " & Row & " ] - Quantity should be entered", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(Row).Specific.Value) > CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(Row).Specific.Value) Then
                        oApplication.StatusBar.SetText("Row [ " & Row & " ] - Quantity should not be greater than " & CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(Row).Specific.Value), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf CDbl(objMatrix.Columns.Item("SerPrice").Cells.Item(Row).Specific.Value) <= 0 Then
                        oApplication.StatusBar.SetText("Row [ " & Row & " ] - Service Price should be entered", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(Row).Specific.Value) <= 0 Then
                        oApplication.StatusBar.SetText("Row [ " & Row & " ] - Item Cost should be entered", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf Trim(objMatrix.Columns.Item("Whs").Cells.Item(Row).Specific.Value) = objForm.Items.Item("vendwhs").Specific.value Then
                        oApplication.StatusBar.SetText("Row [ " & Row & " ] - Row level Warehouse cannot be same as Vendor Warehouse", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    Else
                        Dim cnt As Integer = 0
                        'FG is repeat or not?
                        For i As Integer = 1 To objMatrix.VisualRowCount - 1
                            If Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(Row).Specific.Value).Equals(Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(i).Specific.Value)) = True Then
                                cnt = cnt + 1
                            End If
                        Next
                        If cnt > 1 Then
                            oApplication.StatusBar.SetText("Row [ " & Row & " ] - ItemCode should not be repeat", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If


                        'Raw Material is there or not ? 
                        cnt = 0
                        For i As Integer = 1 To objMatrixRM.VisualRowCount
                            If CInt(objMatrixRM.Columns.Item("Line").Cells.Item(i).Specific.Value) = CInt(objMatrix.Columns.Item("Line").Cells.Item(Row).Specific.value) Then
                                cnt = cnt + 1
                            End If
                        Next
                        If cnt = 0 Then
                            oApplication.StatusBar.SetText("Row [ " & Row & " ] - ItemCode does not have raw material.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next

                For i As Integer = 1 To objMatrix.VisualRowCount
                    If Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(i).Specific.value) <> "" Then
                        Dim UnitPrice As Double
                        For k As Integer = 1 To objMatrixRM.VisualRowCount
                            If Trim(objMatrixRM.Columns.Item("Line").Cells.Item(k).Specific.value) = Trim(objMatrix.Columns.Item("Line").Cells.Item(i).Specific.value) Then
                                UnitPrice = UnitPrice + objMatrixRM.Columns.Item("TotCost").Cells.Item(k).Specific.value
                            End If
                        Next
                        If Math.Ceiling(CDbl(((UnitPrice / objMatrix.Columns.Item("recdqty").Cells.Item(i).Specific.value) + objMatrix.Columns.Item("POPrice").Cells.Item(i).Specific.value) * objMatrix.Columns.Item("recdqty").Cells.Item(i).Specific.value) + 10) < Math.Ceiling(CDbl(objMatrix.Columns.Item("Total").Cells.Item(i).Specific.value)) Then
                            oApplication.StatusBar.SetText("Difference Between Item Cost of Raw Materials and SFG Item cost is greater than 10", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            ' oApplication.MessageBox("Difference Between Item Cost of Raw Materials and SFG Item cost is greater tahn 100-Do you Still Want to continue?", 2, "Yes", "No") = 2 Then
                            ' BubbleEvent = False
                            'Return False
                            'End If
                            Return False
                        End If
                    End If
                Next

                Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                For Row As Integer = 1 To objMatrixRM.VisualRowCount
                    If Trim(objMatrixRM.Columns.Item("ItemCode").Cells.Item(Row).Specific.Value).Equals("") = True Then
                        oApplication.StatusBar.SetText("Row [ " & Row & " ] - ItemCode should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf CDbl(objMatrixRM.Columns.Item("ItemQty").Cells.Item(Row).Specific.Value) <= 0 Then
                        oApplication.StatusBar.SetText("Row [ " & Row & " ] - Planned Quantity should not be zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    Else
                        oRS.DoQuery("Select OnHand from OITW Where ItemCode='" & Trim(objMatrixRM.Columns.Item("ItemCode").Cells.Item(Row).Specific.Value) & "' and WhsCode='" & Trim(objMatrixRM.Columns.Item("Whs").Cells.Item(Row).Specific.Value) & "'")
                        If oRS.RecordCount > 0 Then
                            objMatrixRM.Columns.Item("WhsQty").Cells.Item(Row).Specific.Value = CDbl(oRS.Fields.Item("OnHand").Value)
                            Dim RMQty As Double = 0
                            For i As Integer = 1 To objMatrixRM.VisualRowCount
                                If Trim(objMatrixRM.Columns.Item("ItemCode").Cells.Item(Row).Specific.Value).Equals(Trim(objMatrixRM.Columns.Item("ItemCode").Cells.Item(i).Specific.Value)) = True Then
                                    RMQty = RMQty + CDbl(objMatrixRM.Columns.Item("ItemQty").Cells.Item(Row).Specific.Value)
                                    'RMQty = RMQty + CDbl(objMatrixRM.Columns.Item("ItemQty").Cells.Item(i).Specific.Value)
                                End If
                            Next
                            If RMQty > CDbl(oRS.Fields.Item("OnHand").Value) Then
                                oApplication.StatusBar.SetText("Row [ " & Row & " ] - Planned Quantity for Raw material is not having enough stock - [Planned Qty - " & RMQty & ", InStock - " & CDbl(oRS.Fields.Item("OnHand").Value) & " ]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        End If
                    End If
                Next
            End If

            'oCompany.StartTransaction()
            'Dim iGoodsIssue As Integer = PostGoodsIssue(FormUID)
            'Dim iGoodsReceipt As Integer = PostGoodsReceipt(FormUID)
            'If iGoodsReceipt <> 0 And iGoodsIssue <> 0 Then
            '    objForm.Items.Item("GINO").Specific.Value = iGoodsIssue
            '    objForm.Items.Item("GRNO").Specific.Value = iGoodsReceipt
            '    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            'Else
            '    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            '    Return False
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
                If eventInfo.ItemUID = "ItemMatrix" Then
                    Dim oFOrm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                    Dim oMatrix As SAPbouiCOM.Matrix = oFOrm.Items.Item("ItemMatrix").Specific
                    ITEM = eventInfo.ItemUID
                    ROW_ID = eventInfo.Row
                    DelItemCode = oMatrix.Columns.Item("ItemCode").Cells.Item(ROW_ID).Specific.Value
                    DelLineID = oMatrix.Columns.Item("Line").Cells.Item(ROW_ID).Specific.value
                    If oApplication.Menus.Exists("Close") = True Then oApplication.Menus.RemoveEx("Close")
                End If
                If eventInfo.ItemUID = "RMMatrix" Then
                    Dim oFOrm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                    Dim oMatrix As SAPbouiCOM.Matrix = oFOrm.Items.Item("RMMatrix").Specific
                    ITEM = eventInfo.ItemUID
                    ROW_ID = eventInfo.Row
                    If oApplication.Menus.Exists("Close") = True Then oApplication.Menus.RemoveEx("Close")
                End If

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

    Sub FilterDC(ByVal FormUID As String)
        Try

            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("SCCFL")
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Select DocNum from [@GEN_SC_DC] Where u_SCNo ='" & Trim(objForm.Items.Item("PONo").Specific.Value) & "'")
            For i As Integer = 0 To oRS.RecordCount - 1
                If i > 0 Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                oCon = oCons.Add()
                oCon.Alias = "DocNum"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = oRS.Fields.Item("DocNum").Value
                oRS.MoveNext()
            Next
            If oRS.RecordCount = 0 Then
                oCon = oCons.Add()
                oCon.Alias = "DocNum"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "-1"
            End If
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterFG(ByVal FormUID As String)
        Try

            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("ITEM_CFL")
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Select U_ItemCode ItemCode from [@GEN_SUB_CONTRACT_D0] Where DocEntry='" & Trim(objForm.Items.Item("PODocNo").Specific.Value) & "' and ISNULL(U_Quantity,0)-ISNULL(U_GRNQty,0)>0")
            For i As Integer = 0 To oRS.RecordCount - 1
                If i > 0 Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                oCon = oCons.Add()
                oCon.Alias = "ItemCode"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = oRS.Fields.Item("ItemCode").Value
                oRS.MoveNext()
            Next
            If oRS.RecordCount = 0 Then
                oCon = oCons.Add()
                oCon.Alias = "ItemCode"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "-1"
            End If
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CalculateTotal(ByVal FormUID As String)
4:      Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO")
            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            Dim TotalLC = 0, totalTax As Double = 0
            For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                TotalLC = TotalLC + CDbl(objMatrix.Columns.Item("SerPrice").Cells.Item(Row).Specific.Value)
                totalTax = totalTax + CDbl(objMatrix.Columns.Item("taxamt").Cells.Item(Row).Specific.Value)
            Next
            oDBs_Head.SetValue("U_TotBefTa", 0, TotalLC)
            oDBs_Head.SetValue("U_Tax", 0, totalTax)
            oDBs_Head.SetValue("U_Total", 0, TotalLC + totalTax)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub PrintSC_GRPOReport()
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
            Dim cmd As New SqlCommand("Subcontract_GRPO", strcon)
            cmd.Connection = strcon
            cmd.CommandType = CommandType.StoredProcedure
            Dim oParameter As New SqlParameter("@docNum", SqlDbType.NVarChar)
            oParameter.Value = Trim(objForm.Items.Item("t_docno").Specific.Value)
            Dim dsReport As DataSet = Helper.SqlHelper.ExecuteDataset(strcon, CommandType.StoredProcedure, "Subcontract_GRPO", oParameter)
            dsReport.WriteXml(System.IO.Path.GetTempPath() & "Subcontract_GRPO.xml", System.Data.XmlWriteMode.WriteSchema)
            oUtilities.ShowReport("Subcontract_GRPO.rpt", "Subcontract_GRPO.xml")
            strcon.Close()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    '---> Vijeesh

    Function Generate_Postings(ByVal FormUID As String) As Boolean
        Try
            oCompany.StartTransaction()
            Dim iGoodsIssue As Integer = PostGoodsIssue(FormUID)
            Dim iGoodsReceipt As Integer = PostGoodsReceipt(FormUID)
            If iGoodsReceipt <> 0 And iGoodsIssue <> 0 Then
                objForm.Items.Item("GINO").Specific.Value = iGoodsIssue
                objForm.Items.Item("GRNO").Specific.Value = iGoodsReceipt
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            Else
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                Return False
            End If

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Sub Refresh_RawMaterial(ByVal FormUID As String)
        objForm = oApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("ItemMatrix").Specific
        Me.CalculateTotal(FormUID)
        objMatrix = objForm.Items.Item("ItemMatrix").Specific
        Dim RMMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("RMMatrix").Specific
        oDBs_DetailRM = objForm.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D1")
        'Do Validation on the Rawmaterial Matrix
        For i As Integer = 1 To objMatrix.VisualRowCount
            Dim UnitPrice As Double = 0
            For Row As Integer = 1 To RMMatrix.VisualRowCount
                oDBs_DetailRM.Offset = Row - 1
                If Trim(oDBs_DetailRM.GetValue("U_Line", oDBs_DetailRM.Offset)).Equals(Trim(objMatrix.Columns.Item("Line").Cells.Item(i).Specific.value)) = True Then
                    oDBs_DetailRM.SetValue("LineId", oDBs_DetailRM.Offset, Trim(RMMatrix.Columns.Item("SNo").Cells.Item(Row).Specific.Value))
                    oDBs_DetailRM.SetValue("U_Parent", oDBs_DetailRM.Offset, Trim(RMMatrix.Columns.Item("Parent").Cells.Item(Row).Specific.Value))
                    oDBs_DetailRM.SetValue("U_Line", oDBs_DetailRM.Offset, Trim(RMMatrix.Columns.Item("Line").Cells.Item(Row).Specific.Value))
                    oDBs_DetailRM.SetValue("U_ItemCode", oDBs_DetailRM.Offset, Trim(RMMatrix.Columns.Item("ItemCode").Cells.Item(Row).Specific.Value))
                    oDBs_DetailRM.SetValue("U_ItemName", oDBs_DetailRM.Offset, Trim(RMMatrix.Columns.Item("ItemName").Cells.Item(Row).Specific.Value))
                    oDBs_DetailRM.SetValue("U_Whs", oDBs_DetailRM.Offset, Trim(RMMatrix.Columns.Item("Whs").Cells.Item(Row).Specific.Value))
                    oDBs_DetailRM.SetValue("U_BOMQty", oDBs_DetailRM.Offset, CDbl(RMMatrix.Columns.Item("BOMQty").Cells.Item(Row).Specific.Value))
                    oDBs_DetailRM.SetValue("U_WhsQty", oDBs_DetailRM.Offset, CDbl(RMMatrix.Columns.Item("WhsQty").Cells.Item(Row).Specific.Value))
                    oDBs_DetailRM.SetValue("U_RecdQty", oDBs_DetailRM.Offset, CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(i).Specific.Value))
                    oDBs_DetailRM.SetValue("U_ItemQty", oDBs_DetailRM.Offset, CDbl(RMMatrix.Columns.Item("BOMQty").Cells.Item(Row).Specific.Value) * CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(i).Specific.Value))
                    oDBs_DetailRM.SetValue("U_ItemCost", oDBs_DetailRM.Offset, CDbl(RMMatrix.Columns.Item("ItemCost").Cells.Item(Row).Specific.Value))
                    oDBs_DetailRM.SetValue("U_TotCost", oDBs_DetailRM.Offset, (CDbl(RMMatrix.Columns.Item("BOMQty").Cells.Item(Row).Specific.Value) * CDbl(objMatrix.Columns.Item("recdqty").Cells.Item(i).Specific.Value) * CDbl(RMMatrix.Columns.Item("ItemCost").Cells.Item(Row).Specific.Value)))
                    RMMatrix.SetLineData(Row)
                End If
            Next
            For j As Integer = 1 To RMMatrix.VisualRowCount
                oDBs_DetailRM.Offset = j - 1
                If Trim(oDBs_DetailRM.GetValue("U_Line", oDBs_DetailRM.Offset)) = Trim(objMatrix.Columns.Item("Line").Cells.Item(i).Specific.value) Then
                    UnitPrice = UnitPrice + CDbl(RMMatrix.Columns.Item("TotCost").Cells.Item(j).Specific.value)
                End If
            Next
        Next
        Me.CalculateTotal(FormUID)
    End Sub

End Class



