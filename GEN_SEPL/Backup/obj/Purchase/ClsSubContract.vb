Imports System.Threading
Imports System.Globalization
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text

Public Class ClsSubContract

#Region "        Declaration        "

    Dim oUtilities As New ClsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim objMatrix, objMatrix1 As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail1 As SAPbouiCOM.DBDataSource
    Dim oDBs_DetailRM As SAPbouiCOM.DBDataSource
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim objCheckBox As SAPbouiCOM.CheckBox
    Dim ITEM_ID As String
    'Public sDocNum As String
    'Public sRptName As String

    Dim ROW_ID As Integer = 0
#End Region

    Sub CreateForm()
        Try
            oUtilities.SAPXML("SubContracting.xml")
            objForm = oApplication.Forms.GetForm("GEN_SCForm", oApplication.Forms.ActiveForm.TypeCount)
            objForm.Items.Item("cardcode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("c_series").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("t_docdt").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("t_docno").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT_D0")
            
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Select SlpCode,SlpName from OSLP")
            objCombo = objForm.Items.Item("buyer").Specific
            For i As Integer = 1 To oRS.RecordCount
                objCombo.ValidValues.Add(Trim(oRS.Fields.Item("SlpCode").Value), Trim(oRS.Fields.Item("SlpName").Value))
                oRS.MoveNext()
            Next

            oRS.DoQuery("Select GroupNum,PymntGroup from OCTG")
            objCombo = objForm.Items.Item("paytrms").Specific
            objCombo.ValidValues.Add("", "")
            For i As Integer = 1 To oRS.RecordCount
                objCombo.ValidValues.Add(Trim(oRS.Fields.Item("GroupNum").Value), Trim(oRS.Fields.Item("PymntGroup").Value))
                oRS.MoveNext()
            Next
            objForm.Items.Item("manual").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("cstbom").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("mwobom").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oDBs_Head.SetValue("u_approve", 0, "Y")
            objForm.Items.Item("flditm").AffectsFormMode = False
            objForm.Items.Item("fldrm").AffectsFormMode = False
            objForm.Select()
            objForm.PaneLevel = 1
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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT_D0")
            oUtilities.GetSeries(FormUID, "c_series", "GEN_SUB_CONTRACT")
            oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("c_series").Specific.Selected.Value, "GEN_SUB_CONTRACT"))
            oDBs_Head.SetValue("U_Status", 0, "Open")
            oDBs_Head.SetValue("U_PostDate", 0, DateTime.Today.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_DocDate", 0, DateTime.Today.ToString("yyyyMMdd"))
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Select ISNULL(T0.firstName,'') + ' ' + ISNULL(T0.lastName,'') Owner,empid from OHEM T0 INNER JOIN OUSR T1 ON T0.UserID=T1.USERID Where U_NAME='" & oCompany.UserName & "'")
            If oRS.RecordCount > 0 Then
                oDBs_Head.SetValue("U_Owner", 0, Trim(oRS.Fields.Item("Owner").Value))
                oDBs_Head.SetValue("U_OwnerCod", 0, Trim(oRS.Fields.Item("empid").Value))
            End If
            objCombo = objForm.Items.Item("buyer").Specific
            If objCombo.ValidValues.Count > 0 Then objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            objMatrix.Clear()
            objMatrix.AddRow()
            objMatrix.FlushToDataSource()
            Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
            Dim objMatrixRM As SAPbouiCOM.Matrix
            objMatrixRM = objForm.Items.Item("RMMatrix").Specific
            objMatrixRM.Columns.Item("POQty").Editable = False
            'Dim objcombo As SAPbouiCOM.ButtonCombo
            'objcombo = objForm.Items.Item("copy").Specific
            'objcombo.ValidValues.Add("DC", "DC")
            'objcombo.ValidValues.Add("GRN", "GRN")
            'Dim oCombo As SAPbouiCOM.ButtonCombo
            'oCombo = objForm.Items.Item("copy").Specific
            'oCombo.Caption = "Copy To"
           
            objForm.Items.Item("cardcode").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            objForm.Freeze(False)
        Catch ex As Exception
            'oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT_D0")
            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
            oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_TaxRate", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_DCQty", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_GRNQty", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_Whs", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_SONo", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_SODNo", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_RetQty", oDBs_Detail.Offset, 0)
            objMatrix.SetLineData(Row)
            objMatrix.FlushToDataSource()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine_Raw(ByVal FormUID As String, ByVal Row As Integer)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim objMatrixRM As SAPbouiCOM.Matrix
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT_D1")
            objMatrixRM = objForm.Items.Item("RMMatrix").Specific
            oDBs_Detail.Offset = Row - 1
            oDBs_DetailRM.SetValue("LineID", oDBs_DetailRM.Offset, "")
            oDBs_DetailRM.SetValue("U_LineID", oDBs_DetailRM.Offset, "")
            oDBs_DetailRM.SetValue("U_Father", oDBs_DetailRM.Offset, "")
            oDBs_DetailRM.SetValue("U_Code", oDBs_DetailRM.Offset, "")
            oDBs_DetailRM.SetValue("U_POQty", oDBs_DetailRM.Offset, 0)
            'oDBs_DetailRM.SetValue("U_BOMQty", oDBs_DetailRM.Offset, 0)
            oDBs_DetailRM.SetValue("U_DCQty", oDBs_DetailRM.Offset, 0)
            oDBs_DetailRM.SetValue("U_RetQty", oDBs_DetailRM.Offset, 0)
            objMatrixRM.SetLineData(Row)
            objMatrixRM.FlushToDataSource()
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
                        oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("c_series").Specific.Selected.Value, "GEN_SUB_CONTRACT"))
                       
                       
                    End If
                    If pVal.ItemUID = "buyer" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = False Then
                        Dim BuyerCode As String = oDBs_Head.GetValue("U_Buyer", 0).ToString().Trim()
                        If BuyerCode = "148" Then
                            objForm.Items.Item("59").Visible = True
                            objForm.Items.Item("62").Visible = True
                            objForm.Items.Item("60").Visible = True
                            objForm.Items.Item("63").Visible = True
                            objForm.Items.Item("61").Visible = True
                        Else
                            objForm.Items.Item("59").Visible = False
                            objForm.Items.Item("62").Visible = False
                            objForm.Items.Item("60").Visible = False
                            objForm.Items.Item("63").Visible = False
                            objForm.Items.Item("61").Visible = False
                        End If
                    End If


                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                    Dim objMatrixRM As SAPbouiCOM.Matrix
                    objMatrixRM = objForm.Items.Item("RMMatrix").Specific
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation(FormUID) = False Then BubbleEvent = False
                    ElseIf pVal.ItemUID = "1" And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(FormUID)
                        objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        oDBs_Head.SetValue("u_approve", 0, "Y")
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        Dim BuyerCode As String = oDBs_Head.GetValue("U_Buyer", 0).ToString().Trim()
                        
                            objForm.Items.Item("59").Visible = False
                            objForm.Items.Item("62").Visible = False
                            objForm.Items.Item("60").Visible = False
                            objForm.Items.Item("63").Visible = False
                            objForm.Items.Item("61").Visible = False

                        End If

                    If pVal.ItemUID = "59" And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRSet1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim OldWhs As String = oDBs_Head.GetValue("U_Owhs", 0).ToString().Trim()
                        Dim NewWhs As String = oDBs_Head.GetValue("U_Nwhs", 0).ToString().Trim()
                        'Dim DocnNUm As String = oDBs_Head.GetValue("DocNum", 0).ToString().Trim()
                        'oRSet.DoQuery("SELECT T0.[DocEntry] FROM [dbo].[@GEN_SUB_CONTRACT]  T0 WHERE T0.[DocNum] ='" + DocnNUm + "'")
                        'Dim DocEntry As String = oRSet.Fields.Item("DocEntry").Value.ToString().Trim()

                        'If OldWhs <> "" And NewWhs <> "" Then
                        '    oRSet1.DoQuery("update T0 set T0.U_FWhs='" + NewWhs + "' from [dbo].[@GEN_SUB_CONTRACT_D1]  T0  where T0.docentry='" + DocEntry + "' and T0.U_FWhs='" + OldWhs + "'")

                        'Else

                        'End If
                        Dim objMatrixRM1 As SAPbouiCOM.Matrix = objForm.Items.Item("RMMatrix").Specific
                        'objMatrixRM = 
                        For i As Integer = 1 To objMatrixRM1.VisualRowCount
                            If objMatrixRM1.VisualRowCount > 0 Then
                                Dim oEdit1 As SAPbouiCOM.EditText = objMatrixRM.Columns.Item("Whs").Cells.Item(i).Specific
                                Dim whs As String = oEdit1.Value
                                If whs = OldWhs Then
                                    oEdit1.Value = NewWhs
                                End If

                                'oDBs_Detail.SetValue("U_Whs", i - 1, NewWhs)

                            End If


                        Next
                        objForm.Items.Item("buyer").Enabled = False
                    End If

                    If pVal.ItemUID = "fldrm" Then
                        objForm = oApplication.Forms.Item(pVal.FormUID)
                        objForm.PaneLevel = 2
                    ElseIf pVal.ItemUID = "flditm" Then
                        objForm = oApplication.Forms.Item(pVal.FormUID)
                        objForm.PaneLevel = 1
                    End If

                    '---> Vijeesh
                    If pVal.ItemUID = "manual" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = True Then
                        objCheckBox = objForm.Items.Item("manual").Specific
                        If objCheckBox.Checked = True Then
                            objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            oDBs_Head.SetValue("u_approve", 0, "Y")
                            objMatrixRM.Columns.Item("POQty").Editable = False
                        End If
                    End If
                    If pVal.ItemUID = "cstbom" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = True Then
                        If Trim(oDBs_Head.GetValue("u_cstbom", 0)) = "Y" Then
                            objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            oDBs_Head.SetValue("u_approve", 0, "Y")
                            objMatrixRM.Columns.Item("POQty").Editable = False
                        End If
                    End If
                    If pVal.ItemUID = "mwobom" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = True Then
                        If Trim(oDBs_Head.GetValue("U_manwobom", 0)) = "Y" Then

                            objMatrixRM.Columns.Item("POQty").Editable = True
                            oDBs_Head.SetValue("u_approve", 0, "N")
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select USER_CODE From [OUSR] Where USER_CODE = '" + oCompany.UserName.ToString.Trim + "' And IsNull(u_approve,'N') = 'Y'")
                            If oRSet.RecordCount > 0 Then
                                objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            Else
                                objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            End If
                        Else
                            objMatrixRM.Columns.Item("POQty").Editable = False
                            objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            oDBs_Head.SetValue("u_approve", 0, "Y")
                        End If
                    End If
                    '---> Vijeesh


                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "manual" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = True Then
                        If Trim(oDBs_Head.GetValue("u_cstbom", 0)) = "Y" Then
                            oApplication.StatusBar.SetText("Please uncheck Custom BOM", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                        End If
                        If Trim(oDBs_Head.GetValue("U_manwobom", 0)) = "Y" Then
                            oApplication.StatusBar.SetText("Please uncheck Manual With Out BOM", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                        End If

                    End If
                    If pVal.ItemUID = "cstbom" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = True Then
                        If Trim(oDBs_Head.GetValue("u_manual", 0)) = "Y" Then
                            oApplication.StatusBar.SetText("Please uncheck manual", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                        End If
                        If Trim(oDBs_Head.GetValue("U_manwobom", 0)) = "Y" Then
                            oApplication.StatusBar.SetText("Please uncheck Manual With Out BOM", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                        End If
                    End If
                    If pVal.ItemUID = "mwobom" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = True Then
                        If Trim(oDBs_Head.GetValue("u_cstbom", 0)) = "Y" Then
                            oApplication.StatusBar.SetText("Please uncheck Custom BOM", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                        End If
                        If Trim(oDBs_Head.GetValue("u_manual", 0)) = "Y" Then
                            oApplication.StatusBar.SetText("Please uncheck Manual", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                        End If
                    End If

                    'Vijeesh
                    If (pVal.ItemUID = "cstbom" Or pVal.ItemUID = "mwobom" Or pVal.ItemUID = "manual") And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = False Then
                        Dim objMatrixRM As SAPbouiCOM.Matrix
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        objMatrixRM = objForm.Items.Item("RMMatrix").Specific
                        objMatrix.Clear()
                        objMatrix.AddRow()
                        objMatrix.FlushToDataSource()
                        Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
                        objMatrixRM.Clear()
                    End If
                    'Vijeesh

                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    If pVal.ItemUID = "t_docdt" And pVal.BeforeAction = False And pVal.CharPressed = 9 And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        If objMatrix.VisualRowCount > 0 Then
                            objMatrix.Columns.Item("ItemCode").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Exit Sub
                        End If
                    ElseIf pVal.ItemUID = "RMMatrix" And pVal.BeforeAction = False And pVal.CharPressed = 9 And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.ColUID = "LineID" Then
                        Try
                            objMatrix1 = objForm.Items.Item("RMMatrix").Specific
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT")
                            oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT_D1")
                            Dim Line As Integer = objMatrix1.Columns.Item("LineID").Cells.Item(pVal.Row).Specific.Value
                            If Line <= objMatrix.VisualRowCount Then
                                'Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                'oRS.DoQuery("Select B.Lineid,B.u_itemcode,B.u_qty,A.u_unit,B.u_process from [@GEN_CUST_BOM] A inner join [@GEN_CUST_BOM_D0] B on A.DocEntry = B.DocEntry Where A.u_itemcode = '" & Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(Line).Specific.value) & "' ANd A.DocEntry in (Select Top 1 DocEntry From [@GEN_CUST_BOM] Where u_itemcode = '" + Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(Line).Specific.Value) + "' Order By u_docdate desc)")
                                'oDBs_Detail1.InsertRecord(pVal.Row - 1)
                                oDBs_Detail1.Offset = pVal.Row - 1
                                oDBs_Detail1.SetValue("U_LineID", oDBs_Detail1.Offset, Line)
                                oDBs_Detail1.SetValue("U_Father", oDBs_Detail1.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(Line).Specific.value))
                                objMatrix1.LoadFromDataSource()
                                Me.FilterItem(FormUID, Line)
                            Else
                                objMatrix1.Columns.Item("LineID").Cells.Item(pVal.Row).Specific.Value = ""
                                objMatrix1.Columns.Item("Father").Cells.Item(pVal.Row).Specific.Value = ""
                            End If
                        Catch ex As Exception
                        End Try
                    End If

                    '---> Vijeesh
                    'If pVal.ItemUID = "RMMatrix" And pVal.ColUID = "BOMQty" And pVal.BeforeAction = True And pVal.CharPressed <> 9 And pVal.CharPressed <> 13 And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                    '    If Trim(oDBs_Head.GetValue("U_manwobom", 0)) <> "Y" Then
                    '        BubbleEvent = False
                    '    End If
                    'End If


                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                    Dim objMatrixRM As SAPbouiCOM.Matrix
                    objMatrixRM = objForm.Items.Item("RMMatrix").Specific
                    If pVal.ItemUID = "t_postdt" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Trim(objForm.Items.Item("t_postdt").Specific.Value).Equals("") = False Then
                            If (DateDiff(DateInterval.Day, DateTime.ParseExact(Trim(objForm.Items.Item("t_postdt").Specific.Value), "yyyyMMdd", Nothing), DateTime.Today)) <> 0 Then
                                oApplication.StatusBar.SetText("Posting date varies from system date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        End If
                    ElseIf pVal.ItemUID = "t_deldt" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Trim(objForm.Items.Item("t_deldt").Specific.Value).Equals("") = False Then
                            If (DateDiff(DateInterval.Day, DateTime.ParseExact(Trim(objForm.Items.Item("t_postdt").Specific.Value), "yyyyMMdd", Nothing), DateTime.ParseExact(Trim(objForm.Items.Item("t_deldt").Specific.Value), "yyyyMMdd", Nothing))) < 0 Then
                                oApplication.StatusBar.SetText("Delivery date is before posting date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        End If
                    End If
                    'If pVal.ItemUID = "RMMatrix" And pVal.ColUID = "BOMQty" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.InnerEvent = False Then
                    '    Try
                    '        If oDBs_Head.GetValue("U_manwobom", 0).Trim() = "Y" Then
                    '            objMatrixRM.Columns.Item("POQty").Cells.Item(pVal.Row).Specific.value = (CDbl(Trim(objMatrixRM.Columns.Item("BOMQty").Cells.Item(pVal.Row).Specific.Value)) * CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(1).Specific.Value))
                    '        End If
                    '    Catch ex As Exception
                    '    End Try
                    'End If
                    'If pVal.ItemUID = "RMMatrix" And pVal.ColUID = "POQty" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.InnerEvent = False Then
                    'If pVal.ItemUID = "RMMatrix" And pVal.ColUID = "POQty" And pVal.BeforeAction = True Then
                    '    If oDBs_Head.GetValue("U_manwobom", 0).Trim() = "Y" Then
                    '        objMatrixRM.Columns.Item("BOMQty").Cells.Item(pVal.Row).Specific.value = (CDbl(Trim(objMatrixRM.Columns.Item("POQty").Cells.Item(pVal.Row).Specific.Value)) / CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(1).Specific.Value))
                    '    End If
                    'End If


                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT_D0")
                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                    Dim objMatrixRM As SAPbouiCOM.Matrix
                    objMatrixRM = objForm.Items.Item("RMMatrix").Specific
                    If pVal.ItemUID = "ItemMatrix" And (pVal.ColUID = "UnitPrice" Or pVal.ColUID = "Quantity") And pVal.Row > 0 And pVal.BeforeAction = False Then
                        oDBs_Detail.Offset = pVal.Row - 1
                        oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                        oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemDesc").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_DCQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("DCQty").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_GRNQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("GRNQty").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_RetQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("RetQty").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("TaxCode").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_TaxRate", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, (CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(pVal.Row).Specific.value) * CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.value)) * CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(pVal.Row).Specific.value) / 100)
                        oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(pVal.Row).Specific.value) * CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.value))
                        oDBs_Detail.SetValue("U_Whs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Whs").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_SONo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("SONo").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_SODNo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("SODNo").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("remarks").Cells.Item(pVal.Row).Specific.Value))
                        objMatrix.SetLineData(pVal.Row)
                        Me.CalculateTotal(FormUID)
                        For i As Integer = 1 To objMatrixRM.VisualRowCount
                            If Trim(objMatrixRM.Columns.Item("LineID").Cells.Item(i).Specific.Value) = Trim(objMatrix.Columns.Item("SNo").Cells.Item(pVal.Row).Specific.value) Then
                                If Trim(oDBs_Head.GetValue("u_cstbom", 0)) <> "Y" And Trim(oDBs_Head.GetValue("u_manual", 0)) <> "Y" And Trim(oDBs_Head.GetValue("U_manwobom", 0)) <> "Y" Then
                                    'Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    'oRSet.DoQuery("Select Code,Quantity From ITT1 Where Father = '" + Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.value) + "' And Code = '" + Trim(objMatrixRM.Columns.Item("Code").Cells.Item(i).Specific.value) + "'")
                                    'objMatrixRM.Columns.Item("POQty").Cells.Item(i).Specific.Value = CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.value) * CDbl(oRSet.Fields.Item("Quantity").Value)

                                    '---> Vijeesh
                                    objMatrixRM.Columns.Item("POQty").Cells.Item(i).Specific.Value = CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.value) * CDbl(objMatrixRM.Columns.Item("BOMQty").Cells.Item(i).Specific.value)
                                End If
                                If Trim(oDBs_Head.GetValue("u_cstbom", 0)) = "Y" Then
                                    'Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    'Dim DocEntry As String
                                    'oRSet.DoQuery("Select Top 1 DocEntry As 'DocEntry' From [@GEN_CUST_BOM] Where u_itemcode = '" + Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.value) + "' Order By u_docdate desc")
                                    'DocEntry = oRSet.Fields.Item("DocEntry").Value
                                    'oRSet.DoQuery("Select B.Lineid,B.u_itemcode,B.u_qty from [@GEN_CUST_BOM_D0] B  Where B.u_itemcode = '" & Trim(objMatrixRM.Columns.Item("Code").Cells.Item(i).Specific.value) & "' ANd B.DocEntry = '" + DocEntry + "'")
                                    'objMatrixRM.Columns.Item("POQty").Cells.Item(i).Specific.Value = CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.value) * CDbl(oRSet.Fields.Item("u_qty").Value)

                                    '---> Vijeesh
                                    objMatrixRM.Columns.Item("POQty").Cells.Item(i).Specific.Value = CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.value) * CDbl(objMatrixRM.Columns.Item("BOMQty").Cells.Item(i).Specific.value)
                                End If
                                '---> Vijeesh
                                If Trim(oDBs_Head.GetValue("u_manual", 0)) = "Y" Then
                                    objMatrixRM.Columns.Item("POQty").Cells.Item(i).Specific.Value = CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.value)
                                End If
                                If Trim(oDBs_Head.GetValue("U_manwobom", 0)) = "Y" Then
                                    objMatrixRM.Columns.Item("POQty").Cells.Item(i).Specific.Value = CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.value)
                                End If
                                '---> Vijeesh
                            End If
                        Next
                    End If

                    'If pVal.ItemUID = "RMMatrix" And pVal.ColUID = "BOMQty" And pVal.BeforeAction = False And pVal.ActionSuccess = True Then
                    '    If oDBs_Head.GetValue("U_manwobom", 0).Trim() = "Y" Then
                    '        objMatrixRM.Columns.Item("POQty").Cells.Item(pVal.Row).Specific.value = (CDbl(Trim(objMatrixRM.Columns.Item("BOMQty").Cells.Item(pVal.Row).Specific.Value)) * CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(1).Specific.Value))
                    '    End If
                    'End If
                    'If pVal.ItemUID = "RMMatrix" And pVal.ColUID = "POQty" And pVal.BeforeAction = False And pVal.ActionSuccess = True Then
                    '    If oDBs_Head.GetValue("U_manwobom", 0).Trim() = "Y" Then
                    '        objMatrixRM.Columns.Item("BOMQty").Cells.Item(pVal.Row).Specific.value = (CDbl(Trim(objMatrixRM.Columns.Item("POQty").Cells.Item(pVal.Row).Specific.Value)) / CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(1).Specific.Value))
                    '    End If
                    'End If

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
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT")
                        oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT_D0")
                        oDBs_DetailRM = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT_D1")
                        If oCFL.UniqueID = "VendCFL" Then
                            oDBs_Head.SetValue("U_CardCode", 0, oDT.GetValue("CardCode", 0))
                            oDBs_Head.SetValue("U_CardName", 0, oDT.GetValue("CardName", 0))
                            oDBs_Head.SetValue("U_JourRem", 0, "SubContract Order - " + oDT.GetValue("CardCode", 0))
                            oDBs_Head.SetValue("U_contper", 0, "")
                            oDBs_Head.SetValue("U_vendref", 0, "")
                            oDBs_Head.SetValue("U_vendwhs", 0, "")
                            oDBs_Head.SetValue("U_ContPer", 0, "")
                            oDBs_Head.SetValue("U_VendWhs", 0, oDT.GetValue("U_WhsCode", 0))
                        ElseIf oCFL.UniqueID = "CFL_Owner" Then
                            oDT = CFLEvent.SelectedObjects
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT")
                            oDBs_Head.SetValue("U_OwnerCod", 0, oDT.GetValue("empID", 0))
                            oDBs_Head.SetValue("U_Owner", 0, oDT.GetValue("firstName", 0) + " " + oDT.GetValue("lastName", 0))
                        ElseIf oCFL.UniqueID = "CFL_WHS1" Then
                            oDBs_Head.SetValue("U_VendWhs", 0, oDT.GetValue("WhsCode", 0))
                        ElseIf oCFL.UniqueID = "CFL_Whs" Then
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            oDBs_Detail.Offset = pVal.Row - 1
                            oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                            oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemDesc").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_DCQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("DCQty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_GRNQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("GRNQty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_RetQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("RetQty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("TaxCode").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_TaxRate", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxamt").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Total").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Whs", oDBs_Detail.Offset, Trim(oDT.GetValue("WhsCode", 0)))
                            oDBs_Detail.SetValue("U_SONo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("SONo").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_SODNo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("SODNo").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("remarks").Cells.Item(pVal.Row).Specific.Value))
                            objMatrix.SetLineData(pVal.Row)
                        ElseIf oCFL.UniqueID = "CFL_SO" Then
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            oDBs_Detail.Offset = pVal.Row - 1
                            oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                            oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemDesc").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_DCQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("DCQty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_GRNQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("GRNQty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_RetQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("RetQty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("TaxCode").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_TaxRate", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxamt").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Total").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Whs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Whs").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_SONo", oDBs_Detail.Offset, Trim(oDT.GetValue("DocEntry", 0)))
                            oDBs_Detail.SetValue("U_SODNo", oDBs_Detail.Offset, Trim(oDT.GetValue("DocNum", 0)))
                            oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("remarks").Cells.Item(pVal.Row).Specific.Value))
                            objMatrix.SetLineData(pVal.Row)
                        ElseIf oCFL.UniqueID = "CFL_Tax" Then
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            oDBs_Detail.Offset = pVal.Row - 1
                            oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                            oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemDesc").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_DCQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("DCQty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_GRNQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("GRNQty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_RetQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("RetQty").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, oDT.GetValue("Code", 0))
                            oDBs_Detail.SetValue("U_TaxRate", oDBs_Detail.Offset, CDbl(oDT.GetValue("Rate", 0)))
                            oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, (CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.Value) * CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(pVal.Row).Specific.Value)) * CDbl(oDT.GetValue("Rate", 0)) / 100)
                            oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, (CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.Value) * CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(pVal.Row).Specific.Value)))
                            oDBs_Detail.SetValue("U_Whs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Whs").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_SONo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("SONo").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_SODNo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("SODNo").Cells.Item(pVal.Row).Specific.Value))
                            oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("remarks").Cells.Item(pVal.Row).Specific.Value))
                            objMatrix.SetLineData(pVal.Row)
                            Me.CalculateTotal(FormUID)
                        ElseIf oCFL.UniqueID = "ITEM_CFL1" Then
                            Dim objMatrixRM As SAPbouiCOM.Matrix
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            objMatrixRM = objForm.Items.Item("RMMatrix").Specific
                            If Trim(oDBs_Head.GetValue("U_manwobom", 0)) = "Y" Then
                                Dim OrginRow1 As Integer = objMatrixRM.VisualRowCount
                                For i As Integer = 0 To oDT.Rows.Count - 1
                                    Dim cflSelectedcount1 As Integer = oDT.Rows.Count
                                    If i < cflSelectedcount1 - 1 Then
                                        objMatrixRM.AddRow(1, pVal.Row)
                                        oDBs_Detail.InsertRecord(pVal.Row + i - 1)
                                    End If
                                    oDBs_DetailRM.Offset = pVal.Row - 1 + i
                                    oDBs_DetailRM.SetValue("U_Code", oDBs_DetailRM.Offset, oDT.GetValue("ItemCode", 0))
                                    oDBs_DetailRM.SetValue("U_Father", oDBs_DetailRM.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(1).Specific.value))
                                    oDBs_DetailRM.SetValue("U_POQty", oDBs_DetailRM.Offset, Trim(objMatrix.Columns.Item("Quantity").Cells.Item(1).Specific.value))
                                    oDBs_DetailRM.SetValue("LineID", oDBs_DetailRM.Offset, pVal.Row)
                                    oDBs_DetailRM.SetValue("U_LineID", oDBs_DetailRM.Offset, 1)
                                    objMatrixRM.SetLineData(pVal.Row + i)
                                Next
                                objMatrixRM.FlushToDataSource()
                                If OrginRow1 = pVal.Row Then
                                    objMatrixRM.AddRow()
                                    objMatrixRM.FlushToDataSource()
                                    Me.SetNewLine_Raw(FormUID, objMatrixRM.VisualRowCount)
                                End If
                            Else
                                objMatrixRM.Columns.Item("Code").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Dim oEdit As SAPbouiCOM.EditText = objMatrixRM.Columns.Item("Code").Cells.Item(pVal.Row).Specific
                                Dim oEdit1 As SAPbouiCOM.EditText = objMatrixRM.Columns.Item("Whs").Cells.Item(pVal.Row).Specific
                                Dim oEdit2 As SAPbouiCOM.EditText = objMatrixRM.Columns.Item("POQty").Cells.Item(pVal.Row).Specific
                                Dim oEdit3 As SAPbouiCOM.EditText = objMatrixRM.Columns.Item("BOMQty").Cells.Item(pVal.Row).Specific
                                'oDBs_DetailRM.SetValue("U_Code", pVal.Row, oDT.GetValue("ItemCode", 0))
                                Try
                                    oEdit.Value = oDT.GetValue("ItemCode", 0).ToString.Trim
                                Catch ex As Exception
                                End Try
                                Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                If Trim(oDBs_Head.GetValue("u_cstbom", 0)) = "Y" Then
                                    oRS.DoQuery("Select B.Lineid,B.u_itemcode,B.u_qty,A.u_unit,B.u_process from [@GEN_CUST_BOM] A inner join [@GEN_CUST_BOM_D0] B on A.DocEntry = B.DocEntry Where A.u_itemcode = '" & Trim(objMatrixRM.Columns.Item("Father").Cells.Item(pVal.Row).Specific.value) & "' ANd A.DocEntry in (Select Top 1 DocEntry From [@GEN_CUST_BOM] Where u_itemcode = '" + Trim(objMatrixRM.Columns.Item("Father").Cells.Item(pVal.Row).Specific.value) + "' Order By u_docdate desc)")
                                    oEdit2.Value = CDbl(oRS.Fields.Item("u_qty").Value) * CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(1).Specific.value)
                                    oEdit3.Value = CDbl(oRS.Fields.Item("u_qty").Value)
                                    oRS.DoQuery("Select B.u_stwhs From [@GEN_UNIT_MST] A Inner Join [@GEN_UNIT_MST_D0] B On A.Code = B.Code ANd A.Name = '" + Trim(oRS.Fields.Item("u_unit").Value) + "' Where B.u_process = '" + Trim(oRS.Fields.Item("u_process").Value) + "'")
                                    oEdit1.Value = oRS.Fields.Item(0).Value
                                Else
                                    oRS.DoQuery("Select B.Warehouse,B.Quantity from OITT A inner join ITT1 B on A.Code = B.father Where B.father = '" + Trim(objMatrixRM.Columns.Item("Father").Cells.Item(pVal.Row).Specific.value) + "' and B.Code = '" + Trim(objMatrixRM.Columns.Item("Code").Cells.Item(pVal.Row).Specific.value) + "'")
                                    Dim rw As Integer = Trim(objMatrixRM.Columns.Item("LineID").Cells.Item(pVal.Row).Specific.Value)
                                    oEdit2.Value = CDbl(oRS.Fields.Item("Quantity").Value) * CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(rw).Specific.value)
                                    oEdit3.Value = CDbl(oRS.Fields.Item("Quantity").Value)
                                    oEdit1.Value = oRS.Fields.Item(0).Value
                                End If
                                'objMatrixRM.SetLineData(pVal.Row + 1)
                            End If

                        ElseIf oCFL.UniqueID = "ITEM_CFL" Then
                            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim oRecSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'Dim oRSTax As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRSTax.DoQuery("Select DftApCode from otcd where TcdType='MI'")
                            'Dim DefTaxCode As String
                            'DefTaxCode = oRSTax.Fields.Item(0).Value
                            'Dim oRSTaxRate As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRSTaxRate.DoQuery("Select Code,Rate from OSTC where Code ='" + DefTaxCode + "'")
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            Dim OrginRow As Integer = objMatrix.VisualRowCount
                            For i As Integer = 0 To oDT.Rows.Count - 1
                                Dim cflSelectedcount As Integer = oDT.Rows.Count
                                If i < cflSelectedcount - 1 Then
                                    objMatrix.AddRow(1, pVal.Row)
                                    oDBs_Detail.InsertRecord(pVal.Row + i - 1)
                                End If
                                oRS.DoQuery("Select top 1 (case when DfltWH is null then (Select DfltWhs from oadm) else DfltWH end ) DfltWH from OITM where itemcode='" + oDT.GetValue("ItemCode", i) + "'")
                                oDBs_Detail.Offset = pVal.Row - 1 + i
                                oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, i + pVal.Row)
                                oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, oDT.GetValue("ItemCode", i))
                                oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, oDT.GetValue("ItemName", i))
                                oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, 1)
                                oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, oDT.GetValue("InvntryUom", i))
                                oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, oDT.GetValue("LastPurPrc", i))
                                oDBs_Detail.SetValue("U_DCQty", oDBs_Detail.Offset, 0)
                                oDBs_Detail.SetValue("U_GRNQty", oDBs_Detail.Offset, 0)
                                oDBs_Detail.SetValue("U_RetQty", oDBs_Detail.Offset, 0)
                                oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, "")
                                oDBs_Detail.SetValue("U_TaxRate", oDBs_Detail.Offset, 0)
                                oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, 0)
                                oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, 1 * oDT.GetValue("LastPurPrc", i))
                                oDBs_Detail.SetValue("U_Whs", oDBs_Detail.Offset, oRS.Fields.Item("DfltWH").Value)
                                oDBs_Detail.SetValue("U_SONo", oDBs_Detail.Offset, "")
                                oDBs_Detail.SetValue("U_SODNo", oDBs_Detail.Offset, "")
                                oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, "")
                                objMatrix.SetLineData(pVal.Row + i)
                            Next
                            objMatrix.FlushToDataSource()
                            'If OrginRow = pVal.Row Then
                            '    objMatrix.AddRow()
                            '    objMatrix.FlushToDataSource()
                            '    Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
                            'End If

                            '---> Vijeesh'
                            If OrginRow = pVal.Row And Trim(oDBs_Head.GetValue("U_manwobom", 0)) <> "Y" Then
                                objMatrix.AddRow()
                                objMatrix.FlushToDataSource()
                                Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
                            End If

                            For Row As Integer = 1 To objMatrix.VisualRowCount
                                objMatrix.Columns.Item("SNo").Cells.Item(Row).Specific.Value = Row
                            Next
                            objMatrix.AutoResizeColumns()
                            Me.CalculateTotal(FormUID)
                            Dim objMatrixRM As SAPbouiCOM.Matrix
                            objMatrixRM = objForm.Items.Item("RMMatrix").Specific
                            If Trim(oDBs_Head.GetValue("u_cstbom", 0)) = "Y" Then
                                oDBs_DetailRM.Clear()
                                For i As Integer = 1 To objMatrix.VisualRowCount
                                    If Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(i).Specific.value) <> "" Then
                                        oRS.DoQuery("Select B.Lineid,B.u_itemcode,B.u_qty,A.u_unit,B.u_process from [@GEN_CUST_BOM] A inner join [@GEN_CUST_BOM_D0] B on A.DocEntry = B.DocEntry Where A.u_itemcode = '" & Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(i).Specific.value) & "' ANd A.DocEntry in (Select Top 1 DocEntry From [@GEN_CUST_BOM] Where u_itemcode = '" + Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(i).Specific.value) + "' Order By u_docdate desc)")
                                        Dim j As Integer = oRS.RecordCount
                                        For k As Integer = 1 To oRS.RecordCount
                                            '---> Vijeesh
                                            oRecSet.DoQuery("Select B.u_stwhs From [@GEN_UNIT_MST] A Inner Join [@GEN_UNIT_MST_D0] B On A.Code = B.Code ANd A.Name = '" + Trim(oRS.Fields.Item("u_unit").Value) + "' Where B.u_process = '" + Trim(oRS.Fields.Item("u_process").Value) + "'")
                                            Dim j1 As Integer = oRecSet.RecordCount
                                            oDBs_DetailRM.InsertRecord(oDBs_DetailRM.Size)
                                            oDBs_DetailRM.Offset = oDBs_DetailRM.Size - 1
                                            oDBs_DetailRM.SetValue("LineID", oDBs_DetailRM.Offset, oDBs_DetailRM.Size)
                                            oDBs_DetailRM.SetValue("U_LineID", oDBs_DetailRM.Offset, Trim(objMatrix.Columns.Item("SNo").Cells.Item(i).Specific.value))
                                            oDBs_DetailRM.SetValue("U_Father", oDBs_DetailRM.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(i).Specific.value))
                                            oDBs_DetailRM.SetValue("U_Code", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("u_itemcode").Value))
                                            oDBs_DetailRM.SetValue("U_POQty", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("u_qty").Value) * CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(i).Specific.value))
                                            oDBs_DetailRM.SetValue("U_BOMQty", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("u_qty").Value))
                                            oDBs_DetailRM.SetValue("U_FWhs", oDBs_DetailRM.Offset, oRecSet.Fields.Item("u_stwhs").Value)
                                            oDBs_DetailRM.SetValue("U_DCQty", oDBs_DetailRM.Offset, 0)
                                            oDBs_DetailRM.SetValue("U_RetQty", oDBs_DetailRM.Offset, 0)
                                            oRS.MoveNext()
                                        Next
                                        objMatrixRM.LoadFromDataSource()
                                        'Me.LoadRMs(FormUID, objMatrix.Columns.Item("ItemCode").Cells.Item(i).Specific.value, objMatrix.Columns.Item("SNo").Cells.Item(i).Specific.value, objMatrix.Columns.Item("Quantity").Cells.Item(i).Specific.value)
                                    End If
                                Next
                            End If
                            ' Dim objMatrixRM As SAPbouiCOM.Matrix
                            'objMatrixRM = objForm.Items.Item("RMMatrix").Specific
                            If Trim(oDBs_Head.GetValue("u_cstbom", 0)) <> "Y" And Trim(oDBs_Head.GetValue("u_manual", 0)) <> "Y" And Trim(oDBs_Head.GetValue("U_manwobom", 0)) <> "Y" Then
                                oDBs_DetailRM.Clear()
                                For i As Integer = 1 To objMatrix.VisualRowCount
                                    If Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(i).Specific.value) <> "" Then
                                        'oRS.DoQuery("Select Code,Quantity from ITT1 Where Father='" & Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(i).Specific.value) & "'")
                                        oRS.DoQuery("Select Code,Quantity,Warehouse from ITT1 Where Father='" & Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(i).Specific.value) & "'")
                                        For k As Integer = 1 To oRS.RecordCount
                                            oDBs_DetailRM.InsertRecord(oDBs_DetailRM.Size)
                                            oDBs_DetailRM.Offset = oDBs_DetailRM.Size - 1
                                            oDBs_DetailRM.SetValue("LineID", oDBs_DetailRM.Offset, oDBs_DetailRM.Size)
                                            oDBs_DetailRM.SetValue("U_LineID", oDBs_DetailRM.Offset, Trim(objMatrix.Columns.Item("SNo").Cells.Item(i).Specific.value))
                                            oDBs_DetailRM.SetValue("U_Father", oDBs_DetailRM.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(i).Specific.value))
                                            oDBs_DetailRM.SetValue("U_Code", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("Code").Value))
                                            oDBs_DetailRM.SetValue("U_POQty", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("Quantity").Value) * CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(i).Specific.value))
                                            oDBs_DetailRM.SetValue("U_BOMQty", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("Quantity").Value))
                                            oDBs_DetailRM.SetValue("U_FWhs", oDBs_DetailRM.Offset, oRS.Fields.Item("Warehouse").Value)
                                            oDBs_DetailRM.SetValue("U_DCQty", oDBs_DetailRM.Offset, 0)
                                            oDBs_DetailRM.SetValue("U_RetQty", oDBs_DetailRM.Offset, 0)
                                            oRS.MoveNext()
                                        Next
                                        objMatrixRM.LoadFromDataSource()
                                        'Me.LoadRMs(FormUID, objMatrix.Columns.Item("ItemCode").Cells.Item(i).Specific.value, objMatrix.Columns.Item("SNo").Cells.Item(i).Specific.value, objMatrix.Columns.Item("Quantity").Cells.Item(i).Specific.value)
                                    End If
                                Next
                            End If
                            If Trim(oDBs_Head.GetValue("u_manual", 0)) = "Y" Then
                                oDBs_DetailRM.Clear()
                                For i As Integer = 1 To objMatrix.VisualRowCount
                                    If Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(i).Specific.value) <> "" Then
                                        oDBs_DetailRM.InsertRecord(oDBs_DetailRM.Size)
                                        oDBs_DetailRM.Offset = oDBs_DetailRM.Size - 1
                                        oDBs_DetailRM.SetValue("LineID", oDBs_DetailRM.Offset, oDBs_DetailRM.Size)
                                        oDBs_DetailRM.SetValue("U_LineID", oDBs_DetailRM.Offset, Trim(objMatrix.Columns.Item("SNo").Cells.Item(i).Specific.value))
                                        oDBs_DetailRM.SetValue("U_Father", oDBs_DetailRM.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(i).Specific.value))
                                        oDBs_DetailRM.SetValue("U_Code", oDBs_DetailRM.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(i).Specific.value))
                                        oDBs_DetailRM.SetValue("U_POQty", oDBs_DetailRM.Offset, CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(i).Specific.value))
                                        oDBs_DetailRM.SetValue("U_DCQty", oDBs_DetailRM.Offset, 0)
                                        oDBs_DetailRM.SetValue("U_RetQty", oDBs_DetailRM.Offset, 0)
                                        objMatrixRM.LoadFromDataSource()
                                        'Me.LoadRMs(FormUID, objMatrix.Columns.Item("ItemCode").Cells.Item(i).Specific.value, objMatrix.Columns.Item("SNo").Cells.Item(i).Specific.value, objMatrix.Columns.Item("Quantity").Cells.Item(i).Specific.value)
                                    End If
                                Next
                            End If
                            If Trim(oDBs_Head.GetValue("U_manwobom", 0)) = "Y" Then
                                If objMatrixRM.VisualRowCount < 1 Then
                                    oDBs_DetailRM.Clear()
                                    objMatrixRM.AddRow(1, 1)
                                    oDBs_DetailRM.InsertRecord(0)
                                End If
                            End If
                        End If
                    End If
            End Select
            'If pVal.ItemUID = "copy" And pVal.FormMode = 2 And pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.BeforeAction = False And pVal.ItemChanged = True Then
            '    If pVal.ItemUID = "copy" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.BeforeAction = False And pVal.ItemChanged = True Then
            '        objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            '        Dim oCombo As SAPbouiCOM.ButtonCombo
            '        oCombo = objForm.Items.Item("copy").Specific
            '        If oCombo.Selected.Description = "GRN" Then
            '            oCombo.Caption = "Copy To"
            '            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT")
            '            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT_D0")

            '            oRS.DoQuery("Select U_CardCode,U_CardName,U_ContPer,U_VendRef,U_Owner,U_Buyer,U_PayTrms,U_JourRem,U_TotBefTa,U_Total,U_Tax,U_OwnerCod,U_PayCode, DocNum,U_DocDate from [@GEN_SUB_CONTRACT]  where DocNum='" + objForm.Items.Item("t_docno").Specific.value + "'")
            '            oApplication.ActivateMenuItem("SC_GRPO")
            '            Dim formA As SAPbouiCOM.Form = oApplication.Forms.GetForm("GEN_SCGRPO", oApplication.Forms.ActiveForm.TypeCount)
            '            Dim folderDN As SAPbouiCOM.Folder
            '            folderDN = formA.Items.Item("TabFG").Specific
            '            folderDN.Select()
            '            objMatrix = formA.Items.Item("ItemMatrix").Specific

            '            oDBs_Head1 = formA.DataSources.DBDataSources.Item("@GEN_SC_GRPO")
            '            oDBs_Detail1 = formA.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D0")
            '            oDBs_DetailRM = formA.DataSources.DBDataSources.Item("@GEN_SC_GRPO_D1")
            '            oDBs_Head1.SetValue("U_CardCode", 0, oRS.Fields.Item(0).Value)
            '            oDBs_Head1.SetValue("U_CardName", 0, oRS.Fields.Item(1).Value)
            '            oDBs_Head1.SetValue("U_ContPer", 0, oRS.Fields.Item(2).Value)
            '            oDBs_Head1.SetValue("U_VendRef", 0, oRS.Fields.Item(3).Value)
            '            Dim oRS1 As SAPbobsCOM.Recordset
            '            oRS1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '            oRS1.DoQuery("Select ISNULL(T0.firstName,'') + ' ' + ISNULL(T0.lastName,'') Owner,empid from OHEM T0 INNER JOIN OUSR T1 ON T0.UserID=T1.USERID Where USER_CODE='" & oCompany.UserName & "'")
            '            oDBs_Head1.SetValue("U_Owner", 0, Trim(oRS1.Fields.Item("Owner").Value))
            '            oDBs_Head1.SetValue("U_OwnerCod", 0, Trim(oRS1.Fields.Item("empid").Value))
            '            oDBs_Head1.SetValue("U_Buyer", 0, oRS.Fields.Item(5).Value)
            '            oDBs_Head1.SetValue("U_PayTrms", 0, oRS.Fields.Item(6).Value)
            '            oDBs_Head1.SetValue("U_JourRem", 0, oRS.Fields.Item(7).Value)
            '            oDBs_Head1.SetValue("U_TotBefTa", 0, oRS.Fields.Item(8).Value)
            '            oDBs_Head1.SetValue("U_Total", 0, oRS.Fields.Item(9).Value)
            '            oDBs_Head1.SetValue("U_Tax", 0, oRS.Fields.Item(10).Value)
            '            oDBs_Head1.SetValue("U_PayCode", 0, oRS.Fields.Item(11).Value)
            '            oDBs_Head1.SetValue("U_PONo", 0, oRS.Fields.Item("DocNum").Value)
            '            oDBs_Head1.SetValue("U_PODate", 0, objForm.Items.Item("t_docdt").Specific.value)

            '            oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '            oRS.DoQuery("select b.U_ItemCode,b.U_ItemDesc,b.U_Quantity,b.U_Price,b.U_TotalLC,b.U_TaxRate,b.U_TaxAmt,b.U_TaxCode,b.U_Whs,b.U_Remarks,b.U_UOM from   [@GEN_SUB_CONTRACT] a join  [@GEN_SUB_CONTRACT_D0] b on a.docentry=b.docentry   where a.DocNum='" + objForm.Items.Item("t_docno").Specific.value + "'")

            '            For i As Integer = 0 To oDBs_Detail.Size - 1
            '                objMatrix.AddRow(1, pVal.Row)
            '                'oDBs_Detail1.Offset = pVal.Row - 1 + i
            '                oDBs_Detail1.SetValue("LineID", i, i + 1)
            '                oDBs_Detail1.SetValue("U_ItemCode", i, oDBs_Detail.GetValue(5, 0))
            '                oDBs_Detail1.SetValue("U_ItemDesc", i, oDBs_Detail.GetValue(6, 0))
            '                oDBs_Detail1.SetValue("U_Quantity", i, oDBs_Detail.GetValue(7, 0))
            '                oDBs_Detail1.SetValue("U_Price", i, oDBs_Detail.GetValue(8, 0))
            '                oDBs_Detail1.SetValue("U_TotalLC", i, oDBs_Detail.GetValue(9, 0))
            '                oDBs_Detail1.SetValue("U_TaxRate", i, oDBs_Detail.GetValue(10, 0))
            '                oDBs_Detail1.SetValue("U_TaxAmt", i, oDBs_Detail.GetValue(11, 0))
            '                oDBs_Detail1.SetValue("U_TaxCode", i, oDBs_Detail.GetValue(12, 0))
            '                oDBs_Detail1.SetValue("U_Whs", i, oDBs_Detail.GetValue(13, 0))
            '                oDBs_Detail1.SetValue("U_Remarks", i, oDBs_Detail.GetValue(14, 0))
            '                oDBs_Detail1.SetValue("U_UOM", i, oDBs_Detail.GetValue(15, 0))
            '            Next
            '            objMatrix.FlushToDataSource()
            '            objMatrix.LoadFromDataSource()
            '        End If
            '    End If
            'End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "SC_MENU"
                        If pVal.BeforeAction = False Then
                            Me.CreateForm()
                        End If
                    Case "1282"
                        If objForm.TypeEx = "GEN_SCForm" Then
                            Me.SetDefault(objForm.UniqueID)
                            objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            oDBs_Head.SetValue("u_approve", 0, "Y")
                        End If
                    Case "1281"
                        If objForm.TypeEx = "GEN_SCForm" Then
                            objForm.EnableMenu("1282", True)
                            objForm.Items.Item("t_docno").Click()
                        End If
                    Case "Close"
                        If objForm.TypeEx = "GEN_SCForm" Then
                            If oApplication.MessageBox("Do you want to close?", 2, "Ok", "Cancel") = 1 Then
                                Dim ORS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                ORS.DoQuery("UPDATE [@GEN_SUB_CONTRACT] SET U_Status='Closed' Where DocNum='" & oDBs_Head.GetValue("DocNum", 0) & "'")
                                oDBs_Head.SetValue("U_Status", 0, "Closed")
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                objForm.Items.Item("1").Enabled = True
                            End If
                        End If
                    Case "1293"
                        If objForm.TypeEx = "GEN_SCForm" Then
                            If ITEM_ID.Equals("ItemMatrix") = True Then
                                objForm.Freeze(True)
                                oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT")
                                oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT_D0")
                                objMatrix = objForm.Items.Item("ItemMatrix").Specific
                                For Row As Integer = 1 To objMatrix.VisualRowCount
                                    objMatrix.GetLineData(Row)
                                    oDBs_Detail.Offset = Row - 1
                                    oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
                                    oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemDesc").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_DCQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("DCQty").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_GRNQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("GRNQty").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_RetQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("RetQty").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("TaxCode").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_TaxRate", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxamt").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Total").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_Whs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Whs").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_SONo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("SONo").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_SODNo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("SODNo").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("remarks").Cells.Item(Row).Specific.Value))
                                    objMatrix.SetLineData(Row)
                                Next
                                objMatrix.FlushToDataSource()
                                oDBs_Detail.RemoveRecord(oDBs_Detail.Size - 1)
                                objMatrix.LoadFromDataSource()
                                objForm.Freeze(False)
                            ElseIf ITEM_ID.Equals("RMMatrix") = True Then
                                objForm.Freeze(True)
                                oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT")
                                oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT_D1")
                                objMatrix1 = objForm.Items.Item("RMMatrix").Specific
                                For Row As Integer = 1 To objMatrix1.VisualRowCount
                                    If objMatrix1.IsRowSelected(Row) Then
                                        objMatrix1.FlushToDataSource()
                                        oDBs_Detail1.RemoveRecord(oDBs_Detail1.Size - 1)
                                        objMatrix1.LoadFromDataSource()
                                    End If
                                Next
                                For Row As Integer = 1 To objMatrix1.VisualRowCount
                                    objMatrix1.GetLineData(Row)
                                    oDBs_Detail1.Offset = Row - 1
                                    oDBs_Detail1.Offset = objMatrix1.VisualRowCount - 1
                                    oDBs_Detail1.SetValue("LineId", oDBs_Detail1.Offset, Row)
                                    objMatrix1.SetLineData(Row)
                                Next
                                objForm.Freeze(False)
                            End If

                        End If
                    Case "1292"
                        If objForm.TypeEx = "GEN_SCForm" Then
                            Try
                                If ITEM_ID.Equals("RMMatrix") = True Then
                                    objForm.Freeze(True)
                                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT")
                                    oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT_D1")
                                    objMatrix = objForm.Items.Item("RMMatrix").Specific
                                    objMatrix.AddRow()
                                    objMatrix.FlushToDataSource()
                                    oDBs_Detail1.Offset = objMatrix.VisualRowCount - 1
                                    oDBs_Detail1.SetValue("LineId", oDBs_Detail1.Offset, objMatrix.VisualRowCount)
                                    oDBs_Detail1.SetValue("U_LineID", oDBs_Detail1.Offset, "")
                                    oDBs_Detail1.SetValue("U_Father", oDBs_Detail1.Offset, "")
                                    oDBs_Detail1.SetValue("U_Code", oDBs_Detail1.Offset, "")
                                    oDBs_Detail1.SetValue("U_POQty", oDBs_Detail1.Offset, "0.00")
                                    oDBs_Detail1.SetValue("U_DCQty", oDBs_Detail1.Offset, "0.00")
                                    oDBs_Detail1.SetValue("U_BOMQty", oDBs_Detail1.Offset, "0.00")
                                    oDBs_Detail1.SetValue("U_RetQty", oDBs_Detail1.Offset, "0.00")
                                    oDBs_Detail1.SetValue("U_FWhs", oDBs_Detail1.Offset, "")
                                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                                    objForm.Freeze(False)
                                End If
                            Catch ex As Exception
                                objForm.Freeze(False)
                            End Try
                        End If
                End Select

            ElseIf pVal.BeforeAction = True Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "519"
                        Try
                            If objForm.TypeEx = "GEN_SCForm" Then
                                'BubbleEvent = False
                                sDocNum = objForm.Items.Item("t_docno").Specific.Value
                                sRptName = "SCPO.rpt"
                                Me.Report1()
                                'Me.PrintSCRep()
                            End If
                        Catch ex As Exception

                        End Try
                End Select
            End If
            'ElseIf pVal.BeforeAction = True Then
            '    Select Case pVal.MenuUID
            '        Case "519"
            '            Try
            '                If objForm.TypeEx = "GEN_SCForm" Then
            '                    BubbleEvent = False
            '                    sDocNum = objForm.Items.Item("t_docno").Specific.Value
            '                    sRptName = "SubContract.rpt"
            '                    Me.Report1()
            '                End If
            '            Catch ex As Exception
            '            End Try
            '    End Select
            'End If
        Catch ex As Exception

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

    Sub PrintSCRep()
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
            Dim cmd As New SqlCommand("GEN_SEPL_PRC_SC_PO", strcon)
            cmd.Connection = strcon
            cmd.CommandType = CommandType.StoredProcedure
            Dim oParameter As New SqlParameter("@DocNum", SqlDbType.NVarChar)
            oParameter.Value = Trim(objForm.Items.Item("t_docno").Specific.Value)
            Dim dsReport As DataSet = Helper.SqlHelper.ExecuteDataset(strcon, CommandType.StoredProcedure, "GEN_SEPL_PRC_SC_PO", oParameter)
            dsReport.WriteXml(System.IO.Path.GetTempPath() & "GEN_SEPL_SC_PO.xml", System.Data.XmlWriteMode.WriteSchema)
            oUtilities.ShowReport("GEN_SEPL_SC_PO.rpt", "GEN_SEPL_SC_PO.xml")
            strcon.Close()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Select Case BusinessObjectInfo.EventType
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                Dim objRMMatrix As SAPbouiCOM.Matrix
                If BusinessObjectInfo.BeforeAction = True Then
                    objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT")
                    oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT_D0")
                    oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT_D1")
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                        oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("c_series").Specific.Selected.Value, "GEN_SUB_CONTRACT"))
                    End If
                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                    objRMMatrix = objForm.Items.Item("RMMatrix").Specific
                    objMatrix.LoadFromDataSource()
                    objRMMatrix.LoadFromDataSource()
                    If objMatrix.VisualRowCount <> 0 And Trim(oDBs_Head.GetValue("U_manwobom", 0)) <> "Y" Then
                        oDBs_Detail.RemoveRecord(objMatrix.VisualRowCount - 1)
                        objMatrix.LoadFromDataSource()
                    End If
                    If objRMMatrix.VisualRowCount <> 0 And Trim(oDBs_Head.GetValue("U_manwobom", 0)) = "Y" And BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                        oDBs_Detail1.RemoveRecord(objRMMatrix.VisualRowCount - 1)
                        objRMMatrix.LoadFromDataSource()
                    End If
                    '---> Vijeesh
                    If objMatrix.VisualRowCount > 1 And Trim(oDBs_Head.GetValue("U_manwobom", 0)) = "Y" And BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                        oDBs_Detail.RemoveRecord(objMatrix.VisualRowCount - (objMatrix.VisualRowCount - 1))
                        objMatrix.LoadFromDataSource()
                    End If
                ElseIf BusinessObjectInfo.ActionSuccess = True Then
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Then
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        '---> Vijeesh
                        If Trim(oDBs_Head.GetValue("U_manwobom", 0)) <> "Y" Then
                            objMatrix.AddRow()
                            objMatrix.FlushToDataSource()
                            Me.SetNewLine(BusinessObjectInfo.FormUID, objMatrix.VisualRowCount)
                        End If
                    End If
                End If
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                If BusinessObjectInfo.ActionSuccess = True Then
                    objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    objForm.EnableMenu("1282", True)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT")
                    oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT_D0")
                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                    '---> Vijeesh'
                    If Trim(oDBs_Head.GetValue("U_manwobom", 0)) <> "Y" Then
                        objMatrix.AddRow()
                        objMatrix.FlushToDataSource()
                        Me.SetNewLine(BusinessObjectInfo.FormUID, objMatrix.VisualRowCount)
                    End If
                    '---> Vijeesh
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                    If Trim(oDBs_Head.GetValue("U_Status", 0)).Equals("Closed") = True Then
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                    Else
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                        '---> Vijeesh
                        If Trim(oDBs_Head.GetValue("U_manwobom", 0)) = "Y" Then
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select USER_CODE From [OUSR] Where USER_CODE = '" + oCompany.UserName.ToString.Trim + "' And IsNull(u_approve,'N') = 'Y'")
                            If oRSet.RecordCount > 0 Then
                                objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            Else
                                objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            End If
                        Else
                            objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            oDBs_Head.SetValue("u_approve", 0, "Y")
                        End If
                        '---> Vijeesh

                    End If
                    Dim BuyerCode As String = oDBs_Head.GetValue("U_Buyer", 0).ToString().Trim()
                   
                        objForm.Items.Item("buyer").Enabled = True


                    objForm.Items.Item("1").Enabled = True
                End If
        End Select
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT_D0")

            If Trim(objForm.Items.Item("cardcode").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("CardCode should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Trim(objForm.Items.Item("vendwhs").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Vendor Warehouse should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
            ElseIf (DateDiff(DateInterval.Day, DateTime.ParseExact(Trim(objForm.Items.Item("t_postdt").Specific.Value), "yyyyMMdd", Nothing), DateTime.ParseExact(Trim(objForm.Items.Item("t_deldt").Specific.Value), "yyyyMMdd", Nothing))) < 0 Then
                oApplication.StatusBar.SetText("Delivery date is before posting date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            If objMatrix.VisualRowCount = 1 Then
                If Trim(oDBs_Head.GetValue("U_manwobom", 0)) <> "Y" Then
                    oApplication.StatusBar.SetText("No items defined", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                Else
                    If Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(objMatrix.VisualRowCount).Specific.Value).Equals("") = True Then
                        oApplication.StatusBar.SetText("ItemCode should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(objMatrix.VisualRowCount).Specific.Value) <= 0 Then
                        oApplication.StatusBar.SetText("Quantity should be greater than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf Trim(objMatrix.Columns.Item("TaxCode").Cells.Item(objMatrix.VisualRowCount).Specific.Value).Equals("") = True Then
                        oApplication.StatusBar.SetText("TaxCode should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf Trim(objMatrix.Columns.Item("Whs").Cells.Item(objMatrix.VisualRowCount).Specific.Value) = objForm.Items.Item("vendwhs").Specific.value Then
                        oApplication.StatusBar.SetText("Row [ " & objMatrix.VisualRowCount & " ] - Row level Warehouse cannot be same as Vendor Warehouse", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf Trim(objMatrix.Columns.Item("Whs").Cells.Item(objMatrix.VisualRowCount).Specific.Value).Equals("") = True Then
                        oApplication.StatusBar.SetText("Row [ " & objMatrix.VisualRowCount & " ] - Row level Warehouse cannot be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            Else
                For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                    If Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(Row).Specific.Value).Equals("") = True Then
                        oApplication.StatusBar.SetText("ItemCode should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(Row).Specific.Value) <= 0 Then
                        oApplication.StatusBar.SetText("Quantity should be greater than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf Trim(objMatrix.Columns.Item("TaxCode").Cells.Item(Row).Specific.Value).Equals("") = True Then
                        oApplication.StatusBar.SetText("TaxCode should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf Trim(objMatrix.Columns.Item("Whs").Cells.Item(Row).Specific.Value) = objForm.Items.Item("vendwhs").Specific.value Then
                        oApplication.StatusBar.SetText("Row [ " & Row & " ] - Row level Warehouse cannot be same as Vendor Warehouse", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf Trim(objMatrix.Columns.Item("Whs").Cells.Item(Row).Specific.Value).Equals("") = True Then
                        oApplication.StatusBar.SetText("Row [ " & Row & " ] - Row level Warehouse cannot be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                Next
            End If
            Dim objRMMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("RMMatrix").Specific
            If objRMMatrix.VisualRowCount = 1 Then
                If Trim(oDBs_Head.GetValue("U_manwobom", 0)) = "Y" Then
                    If Trim(objRMMatrix.Columns.Item("Code").Cells.Item(objRMMatrix.VisualRowCount).Specific.Value).Equals("") = True Then
                        oApplication.StatusBar.SetText("No Raw Materials defined", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            ElseIf objRMMatrix.VisualRowCount < 1 Then
                oApplication.StatusBar.SetText("No Raw Materials defined", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                For i As Integer = 1 To objRMMatrix.VisualRowCount - 1
                    If Trim(objRMMatrix.Columns.Item("Code").Cells.Item(i).Specific.Value).Equals("") = True Then
                        oApplication.StatusBar.SetText("Child ItemCode should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                    If Trim(objRMMatrix.Columns.Item("Father").Cells.Item(i).Specific.Value).Equals("") = True Then
                        oApplication.StatusBar.SetText("Father ItemCode should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                    'If Trim(objRMMatrix.Columns.Item("Whs").Cells.Item(i).Specific.Value).Equals("INSP") = True Then
                    '    oApplication.StatusBar.SetText("Child ItemCode should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    Return False
                    'End If

                    'If Trim(oDBs_Head.GetValue("U_manwobom", 0)) = "Y" Then
                    '    If CDbl(objRMMatrix.Columns.Item("BOMQty").Cells.Item(i).Specific.Value) <= 0 Then
                    '        oApplication.StatusBar.SetText("BOM Quantity should be greater than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        Return False
                    '    End If
                    'End If
                Next
            End If

            'Me.LoadRMs(FormUID)
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
        
        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRS.DoQuery("Select USER_CODE From [OUSR] Where USER_CODE = '" + oCompany.UserName.ToString.Trim + "' And U_POEdit = 'YES'")
        If oRS.RecordCount > 0 Then
            Try
                ROW_ID = eventInfo.Row
                If eventInfo.Row > 0 Then
                    ITEM_ID = eventInfo.ItemUID
                    Dim objMatrixRM As SAPbouiCOM.Matrix
                    objMatrixRM = objForm.Items.Item("RMMatrix").Specific
                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                    If ITEM_ID.Equals("RMMatrix") = True Then
                        If objMatrixRM.VisualRowCount > 1 Then
                            objForm.EnableMenu("1293", True)
                            objForm.EnableMenu("1292", True)
                        Else
                            objForm.EnableMenu("1293", False)
                            objForm.EnableMenu("1292", True)
                        End If
                    ElseIf ITEM_ID.Equals("ItemMatrix") = True Then
                        If objMatrix.VisualRowCount >= 1 Then
                            objForm.EnableMenu("1293", True)
                        Else
                            objForm.EnableMenu("1293", False)
                        End If
                    End If
                Else
                    ITEM_ID = ""
                End If
            Catch ex As Exception
                oApplication.StatusBar.SetText(ex.Message)
            End Try
        Else
            objForm.EnableMenu("1293", False)
            objForm.EnableMenu("1292", False)
        End If
    End Sub

    Sub CalculateTotal(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT")
            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            Dim TotalLC = 0, totalTax As Double = 0
            'Vijeesh
            If Trim(oDBs_Head.GetValue("U_manwobom", 0)) <> "Y" Then
                For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                    TotalLC = TotalLC + CDbl(objMatrix.Columns.Item("Total").Cells.Item(Row).Specific.Value)
                    totalTax = totalTax + CDbl(objMatrix.Columns.Item("taxamt").Cells.Item(Row).Specific.Value)
                Next
            Else
                TotalLC = TotalLC + CDbl(objMatrix.Columns.Item("Total").Cells.Item(objMatrix.VisualRowCount).Specific.Value)
                totalTax = totalTax + CDbl(objMatrix.Columns.Item("taxamt").Cells.Item(objMatrix.VisualRowCount).Specific.Value)
            End If
            'Vijeesh
            oDBs_Head.SetValue("U_TotBefTa", 0, TotalLC)
            oDBs_Head.SetValue("U_Tax", 0, totalTax)
            oDBs_Head.SetValue("U_Total", 0, TotalLC + totalTax)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadRMs(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim objMatrixRM As SAPbouiCOM.Matrix
            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            objMatrixRM = objForm.Items.Item("RMMatrix").Specific
            oDBs_DetailRM = objForm.DataSources.DBDataSources.Item("@GEN_SUB_CONTRACT_D1")
            oDBs_DetailRM.Clear()
            For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                oRS.DoQuery("Select Code,Quantity from ITT1 Where Father='" & Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(Row).Specific.Value) & "'")
                For i As Integer = 1 To oRS.RecordCount
                    oDBs_DetailRM.InsertRecord(oDBs_DetailRM.Size)
                    oDBs_DetailRM.Offset = oDBs_DetailRM.Size - 1
                    oDBs_DetailRM.SetValue("LineID", oDBs_DetailRM.Offset, oDBs_DetailRM.Size)
                    oDBs_DetailRM.SetValue("U_LineID", oDBs_DetailRM.Offset, Row)
                    oDBs_DetailRM.SetValue("U_Father", oDBs_DetailRM.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(Row).Specific.Value))
                    oDBs_DetailRM.SetValue("U_Code", oDBs_DetailRM.Offset, Trim(oRS.Fields.Item("Code").Value))
                    oDBs_DetailRM.SetValue("U_POQty", oDBs_DetailRM.Offset, CDbl(oRS.Fields.Item("Quantity").Value) * CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(Row).Specific.Value))
                    oDBs_DetailRM.SetValue("U_DCQty", oDBs_DetailRM.Offset, 0)
                    oDBs_DetailRM.SetValue("U_RetQty", oDBs_DetailRM.Offset, 0)
                    oRS.MoveNext()
                Next
            Next
            objMatrixRM.LoadFromDataSource()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    'Sub PrintSCReport()
    '    Try
    '        Dim oFile As New StreamReader(Application.StartupPath & "\DBLogin.ini", False)
    '        Dim s As String = ""
    '        Dim i As Integer = 1
    '        Dim Company = "", UserName = "", Password As String = ""
    '        s = oFile.ReadLine()
    '        While s <> ""
    '            Select Case i
    '                Case 1
    '                    Company = s.Trim
    '                Case 2
    '                    UserName = s.Trim
    '                Case 3
    '                    Password = s.Trim
    '            End Select
    '            i = i + 1
    '            s = oFile.ReadLine
    '        End While
    '        Dim strcon As New SqlConnection("user id=" & UserName & ";data source=" & Company & ";pwd=" & Password & ";initial catalog=" & oCompany.CompanyDB & ";")
    '        strcon.Open()
    '        objForm = oApplication.Forms.ActiveForm
    '        Dim cmd As New SqlCommand("Subcontract", strcon)
    '        cmd.Connection = strcon
    '        cmd.CommandType = CommandType.StoredProcedure
    '        Dim oParameter As New SqlParameter("@docNum", SqlDbType.NVarChar)
    '        oParameter.Value = Trim(objForm.Items.Item("t_docno").Specific.Value)
    '        Dim dsReport As DataSet = Helper.SqlHelper.ExecuteDataset(strcon, CommandType.StoredProcedure, "Subcontract", oParameter)
    '        dsReport.WriteXml(System.IO.Path.GetTempPath() & "Subcontract.xml", System.Data.XmlWriteMode.WriteSchema)
    '        oUtilities.ShowReport("SubContract.rpt", "Subcontract.xml")
    '        strcon.Close()
    '    Catch ex As Exception
    '        oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try
    'End Sub
    Sub FilterItem(ByVal FormUID As String, ByVal Line As Integer)
        Try

            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("ITEM_CFL1")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            Dim oRSets As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSets.DoQuery("Select B.u_itemcode from [@GEN_CUST_BOM] A inner join [@GEN_CUST_BOM_D0] B on A.DocEntry = B.DocEntry Where A.u_itemcode = '" & Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(Line).Specific.value) & "' ANd A.DocEntry in (Select Top 1 DocEntry From [@GEN_CUST_BOM] Where u_itemcode = '" + Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(Line).Specific.Value) + "' Order By u_docdate desc)")
            'oRSets.DoQuery("Select Code from ITT1 Where Father='" & Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(Line).Specific.value) & "'")

            Dim orsf As Integer = oRSets.RecordCount
            For IntICount As Integer = 0 To oRSets.RecordCount - 1
                If IntICount = (oRSets.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRSets.Fields.Item("u_itemcode").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRSets.Fields.Item("u_itemcode").Value
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
