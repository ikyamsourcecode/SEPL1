Imports System.Threading
Imports System.Globalization
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.Math
Public Class ClsPreShipment

#Region "        Declaration        "

    Dim oUtilities As New ClsUtilities
    Dim objForm, objSForm, objSubForm, FrgtForm, objFreightForm, exdForm As SAPbouiCOM.Form
    Dim objMatrix, objMatrix1 As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail1 As SAPbouiCOM.DBDataSource
    Dim oDBs_DetailRM As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail_Freight As SAPbouiCOM.DBDataSource

    Dim objCombo, objCombo1 As SAPbouiCOM.ComboBox
    Dim objCheckBox As SAPbouiCOM.CheckBox
    Dim ModalForm As Boolean = False
    Dim ChildModalForm As Boolean = False
    Dim PARENT_FORM As String
    Dim ITEM_ID, CUST_NO As String
    'Dim sDocNum As Integer
    Dim Curr As String
    'Public sDocNum As String
    Dim InvNo As String
    Dim FrghtFlag As Boolean = False
    Dim DbkVal, DbkPer, Dbk, ANSP, COMM, INS, LineTotalSum, TRNSP As Double
    Public tot, tots As String
    Dim DOCNUM As String = ""
    Dim SALDOC As String = ""
    Dim ROW_ID As Integer = 0
    Dim BASENUM As String
    Dim FormMode As String
    Dim COM As Double
    Dim ANSPCur As String
    Dim RowNo As Integer
    Dim flgexd As Boolean
    Dim loadcount As Integer = 0
#End Region

    Sub CreateForm()
        Try
            oUtilities.SAPXML("PreShipment.xml")
            objForm = oApplication.Forms.GetForm("PRE_SHIPMENT", oApplication.Forms.ActiveForm.TypeCount)
            objForm.Items.Item("custname").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("series").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("docdate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("preno").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("preno").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("doccur").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("custcode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("roundpr").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("totbef").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("freight").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("tax").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D0")
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Select SlpCode,SlpName from OSLP ORDER BY SlpCode")
            objCombo = objForm.Items.Item("buyer").Specific
            For i As Integer = 1 To oRS.RecordCount
                objCombo.ValidValues.Add(Trim(oRS.Fields.Item("SlpCode").Value), Trim(oRS.Fields.Item("SlpName").Value))
                oRS.MoveNext()
            Next
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSet.DoQuery("Select Code,Name From [@COSTFREIGHT]")
            objCombo = objForm.Items.Item("cfr").Specific
            For i As Integer = 1 To oRSet.RecordCount
                objCombo.ValidValues.Add(oRSet.Fields.Item("Code").Value, oRSet.Fields.Item("Name").Value)
                oRSet.MoveNext()
            Next
            'oRS.DoQuery("Select GroupNum,PymntGroup from OCTG")
            'objCombo = objForm.Items.Item("payterms").Specific
            'objCombo.ValidValues.Add("", "")
            'For i As Integer = 1 To oRS.RecordCount
            '    objCombo.ValidValues.Add(Trim(oRS.Fields.Item("GroupNum").Value), Trim(oRS.Fields.Item("PymntGroup").Value))
            '    oRS.MoveNext()
            'Next
            'objForm.Items.Item("manual").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("cstbom").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("mwobom").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'oDBs_Head.SetValue("u_approve", 0, "Y")
            'objForm.Items.Item("flditm").AffectsFormMode = False
            objForm.Items.Item("content").AffectsFormMode = False
            objForm.Select()
            objForm.PaneLevel = 1
            'oApplication.ActivateMenuItem("6913")
            objForm.EnableMenu("5890", True)
            'oApplication.ActivateMenuItem("5890")
            'objForm.ResetMenuStatus()
            'objForm.
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
            objForm.DataBrowser.BrowseBy = "preno"
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D0")
            oUtilities.GetSeries(FormUID, "series", "PRE_SHIPMENT")
            oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("series").Specific.Selected.Value, "PRE_SHIPMENT"))
            oDBs_Head.SetValue("U_Status", 0, "Open")
            oDBs_Head.SetValue("U_PosDate", 0, DateTime.Today.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_DocDate", 0, DateTime.Today.ToString("yyyyMMdd"))
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Select ISNULL(T0.firstName,'') + ' ' + ISNULL(T0.lastName,'') Owner,empid from OHEM T0 INNER JOIN OUSR T1 ON T0.UserID=T1.USERID Where U_NAME='" & oCompany.UserName & "'")
            If oRS.RecordCount > 0 Then
                oDBs_Head.SetValue("U_Owner", 0, Trim(oRS.Fields.Item("Owner").Value))
                oDBs_Head.SetValue("U_OwnCode", 0, Trim(oRS.Fields.Item("empid").Value))
            End If
            objCombo = objForm.Items.Item("buyer").Specific
            If objCombo.ValidValues.Count > 0 Then objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            ''objMatrix.Clear()
            ''objMatrix.AddRow()
            ''objMatrix.FlushToDataSource()
            ''Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
            ''oApplication.ActivateMenuItem("6913")
            'Dim objMatrixRM As SAPbouiCOM.Matrix
            'objMatrixRM = objForm.Items.Item("ItemMatrix").Specific
            'objMatrixRM.Columns.Item("POQty").Editable = False
            'Dim objcombo As SAPbouiCOM.ButtonCombo
            'objcombo = objForm.Items.Item("copy").Specific
            'objcombo.ValidValues.Add("DC", "DC")
            'objcombo.ValidValues.Add("GRN", "GRN")
            'Dim oCombo As SAPbouiCOM.ButtonCombo
            'oCombo = objForm.Items.Item("copy").Specific
            'oCombo.Caption = "Copy To"
            objForm.EnableMenu("5907", True)
            objForm.Items.Item("custcode").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            objForm.Freeze(False)
        Catch ex As Exception
            'oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D0")
            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
            oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_ItemName", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_Price_A", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_DocCur", oDBs_Detail.Offset, oDBs_Head.GetValue("U_DocCur", 0).Trim)
            oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_Total_A", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("U_Whse", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_SONo", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_Note", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, "")
            objMatrix.SetLineData(Row)
            objMatrix.FlushToDataSource()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Dim CHILD_FORM As String = "ACCRUALS_PRE@" & pVal.FormUID
            Dim ModalForm As Boolean = False
            For i As Integer = 0 To oApplication.Forms.Count - 1
                If oApplication.Forms.Item(i).UniqueID = (CHILD_FORM) Then
                    objSubForm = oApplication.Forms.Item("ACCRUALS_PRE@" & pVal.FormUID)
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
                        If pVal.BeforeAction = False Then
                            If Trim(DOCNUM).Equals("") = False Then
                                objForm = oApplication.Forms.Item(FormUID)
                                objForm.Freeze(True)
                                Me.LoadItems(FormUID, DOCNUM)
                                DOCNUM = ""
                                objForm.Freeze(False)
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If pVal.ItemUID = "series" And pVal.BeforeAction = False Then
                            oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("series").Specific.Selected.Value, "PRE_SHIPMENT"))
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                        If pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'RS.DoQuery("Delete From [@GEN_SZ_ORDR] Where u_sono = '" + Trim(objForm.Items.Item("8").Specific.value) + "' And u_macid = '" + MAC_ID + "'")
                        ElseIf pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            If Me.Validation_Close(FormUID) = False Then BubbleEvent = False
                            'oApplication.StatusBar.SetText("Please dont close this button", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            'BubbleEvent = False
                            'Exit Sub
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        objForm = oApplication.Forms.Item(pVal.FormUID)
                        If pVal.ItemUID = "1" And pVal.ActionSuccess = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            If Me.Validation(FormUID) = False Then BubbleEvent = False
                            Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim oRecordSet2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D0")
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
                            Dim oRecordSet1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim ss As String = "select U_Ecgc from OCRD where cadcode='" + oDBs_Head.GetValue("U_CustCode", 0).ToString().Trim() + "'"
                            oRecordSet1.DoQuery("select isnull(U_Ecgc,0) as ecgc from OCRD where cardcode='" + oDBs_Head.GetValue("U_CustCode", 0).ToString().Trim() + "'")
                            If oRecordSet1.Fields.Item("ecgc").Value.ToString().Trim() = "YES" Then
                                Dim strQry1 As String = "select Distinct T0.U_CustCode from [dbo].[@PRE_SHIPMENT] T0 where T0.U_CustCode in(select code from [dbo].[@GEN_ECGC] union all select U_cardcode from [dbo].[@GEN_ECGC_D0] ) and T0.U_CustCode='" + oDBs_Head.GetValue("U_CustCode", 0).ToString().Trim() + "'"
                                RS.DoQuery(strQry1)
                                If RS.RecordCount > 0 Then
                                    Dim strQuery As String = "Select T0.Code,SUM(T0.Balance)+(select balance from ocrd where cardcode=T0.code) as TotalOS,T0.U_Ecgc   from(SELECT Distinct T0.code,T1.U_cardcode,T0.U_ecgc,T2.balance  FROM [dbo].[@GEN_ECGC]  T0 inner join [dbo].[@GEN_ECGC_D0]  T1 on T1.code=T0.code inner join OCRD T2 on T2.cardcode=T1.U_cardcode WHERE T1.[U_cardcode]<>'' and (T0.Code ='" + RS.Fields.Item("U_CustCode").Value.ToString() + "' or T1.U_cardcode ='" + RS.Fields.Item("U_CustCode").Value.ToString() + "'))T0 group by T0.Code,T0.U_Ecgc"
                                    oRecordSet.DoQuery(strQuery)
                                    Dim TotalOS As Double = CDbl(oRecordSet.Fields.Item("TotalOS").Value)
                                    Dim SONo As String = oDBs_Detail.GetValue("U_SONo", 0).ToString().Trim()
                                    Dim strQuery12 As String = "SELECT Distinct T0.[DocNum],T0.[DocRate] FROM ORDR T0 where docnum='" + SONo + "'"

                                    oRecordSet2.DoQuery("SELECT Distinct  T0.[DocNum],T0.[DocRate] FROM ORDR T0   where docnum='" + SONo + "'")
                                    Dim U_Total As Double = oDBs_Head.GetValue("U_Total", 0).Trim() * Convert.ToDouble(oRecordSet2.Fields.Item("DocRate").Value.ToString().Trim())


                                    Dim U_Ecgc As Double = CDbl(oRecordSet.Fields.Item("U_Ecgc").Value)

                                    If TotalOS + U_Total > U_Ecgc Then
                                        BubbleEvent = False
                                        oApplication.StatusBar.SetText("Customer Exceeds the ECGC Limit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                    End If
                                End If
                            End If
                            ' Dim strQuery As String = "Select T0.Code,SUM(T0 .Balance)+(select balance from ocrd where cardcode='2V000841') as TotalOS,T0.U_Ecgc   from(SELECT Distinct T0.code,T1.U_cardcode,T0.U_ecgc,T2.balance  FROM [dbo].[@GEN_ECGC]  T0 inner join [dbo].[@GEN_ECGC_D0]  T1 on T1.code=T0.code inner join OCRD T2 on T2.cardcode=T1.U_cardcode WHERE T1.[U_cardcode]<>'' and (T0.Code ='C000059' or T1.U_cardcode ='C000059'))T0 group by T0.Code,T0.U_Ecgc"

                            'If Me.Validation_Close(FormUID) = False Then BubbleEvent = False
                        ElseIf pVal.ItemUID = "1" And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Try
                                Dim oRecset, orecset1 As SAPbobsCOM.Recordset
                                oRecset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim docno As String = oDBs_Head.GetValue("DocNum", 0).Trim
                                orecset1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecset.DoQuery("SELECT T0.[Code], T0.[U_MacId], T0.[U_FreCode], T0.[U_FreName], T0.[U_TaxCode], T0.[U_PreNo], T0.[U_TaxCode], T0.[U_Amt], T0.[U_TotTax] FROM [dbo].[@UBG_PRE_FRET_D0]  T0 Where T0.[U_PreNo] = '" + docno + "'")
                                Dim docent As String
                                docent = oDBs_Head.GetValue("DocEntry", 0).Trim
                                For fgrt As Integer = 1 To oRecset.RecordCount
                                    Dim doc, line, expcode, expname, frgt, taxcode, macid As String
                                    doc = oRecset.Fields.Item("U_PreNo").Value
                                    orecset1.DoQuery("Select * from [@PRE_SHIPMENT_D3] Where DocEntry = '" + docent + "'")
                                    If oRecset.RecordCount = orecset1.RecordCount Then
                                        Exit For
                                    End If
                                    line = oRecset.Fields.Item("Code").Value
                                    expcode = oRecset.Fields.Item("U_FreCode").Value
                                    expname = oRecset.Fields.Item("U_FreName").Value
                                    frgt = oRecset.Fields.Item("U_Amt").Value.ToString.Substring(3).Replace(" ", "")
                                    taxcode = oRecset.Fields.Item("U_TaxCode").Value
                                    macid = oRecset.Fields.Item("U_MacId").Value
                                    orecset1.DoQuery("Insert Into [@PRE_SHIPMENT_D3](DocEntry,U_DocNum,LineId,U_ExpnCode,U_ExpnName,U_LineTot,U_TotFrgn,U_TaxCode,U_TaxType,U_Curr) VALUES('" + docent.ToString.Trim + "','" + doc.ToString.Trim + "','" + line.ToString.Trim + "','" + expcode + "','" + expname + "','0','" + frgt + "','" + taxcode + "','','" + Curr + "')")
                                    oRecset.MoveNext() '.ToString.Trim
                                Next
                                loadcount = 0
                                Me.SetDefault(FormUID)
                            Catch ex As Exception
                                BubbleEvent = False
                            End Try
                        ElseIf pVal.ItemUID = "2" And pVal.ActionSuccess = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Dim delrec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            delrec.DoQuery("DELETE [@UBG_PRE_FRET_D0]")
                        ElseIf pVal.ItemUID = "freightlk" And pVal.ActionSuccess = True Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                objForm = oApplication.Forms.Item(FormUID)
                                Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                                Dim rd As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rd.DoQuery("Select * from [@UBG_PRE_FRET_D0]")
                                If rd.RecordCount = 0 Then
                                    Me.OpenFreight(FormUID)
                                Else
                                    Me.Open_Pre_Freight_Matrix_Form(FormUID, objForm.Items.Item("preno").Specific.Value, MAC_ID)
                                End If
                            ElseIf pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                objForm = oApplication.Forms.Item(FormUID)
                                Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                                Me.Open_Pre_Freight_Matrix_Form(FormUID, objForm.Items.Item("preno").Specific.Value, MAC_ID)
                            End If
                        ElseIf pVal.ItemUID = "btnac" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.BeforeAction = False Then
                            Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            InvNo = objForm.Items.Item("preno").Specific.Value
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
                            Dbk = 0
                            ANSP = 0
                            COMM = 0
                            TRNSP = 0
                            LineTotalSum = 0
                            'SALDOC = ""
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            For i As Integer = 0 To objMatrix.VisualRowCount - 1
                                SALDOC = SALDOC & "'" & oDBs_Detail.GetValue("U_SONo", i).Trim & "',"
                            Next
                            SALDOC = SALDOC.Substring(0, Len(SALDOC) - 1)
                            oRSet.DoQuery("Select A.ExpnsCode,C.ExpnsName,A.TaxCode,A.LineTotal,C.U_nfreight,C.U_pfreight,C.U_appltax from RDR3 A left outer  join ORDR B ON A.DocEntry = B.DocEntry left outer join OEXD C on C.ExpnsCode = A.ExpnsCode Where B.DocNum IN (" + SALDOC + ") and C.U_incl = 'YES'")
                            If oRSet.RecordCount = 0 Then
                                For i As Integer = 1 To objMatrix.VisualRowCount
                                    If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) <> "" And Trim(objMatrix.Columns.Item("unitprice").Cells.Item(i).Specific.value) <> "" Then
                                        oRSet.DoQuery("Select Distinct IsNull(u_expval,0) 'u_expval',isnull(u_expper,0) 'u_expper' From OITM Where ItemCode = '" + Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) + "'")
                                        Dbk = Dbk + (oRSet.Fields.Item("u_expval").Value * objMatrix.Columns.Item("qty").Cells.Item(i).Specific.value)
                                        oRSet.DoQuery("Select B.u_ansp,A.u_doccur,B.u_comm From [@GEN_SUPP_PRICE] A INNER JOIN [@GEN_SUPP_PRICE_D0] B On A.Code = B.Code Where B.u_ItemCode = '" + Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) + "' And A.u_cardcode = '" + Trim(objForm.Items.Item("custcode").Specific.value) + "'")
                                        ANSP = ANSP + (objMatrix.Columns.Item("qty").Cells.Item(i).Specific.Value * oRSet.Fields.Item("u_ansp").Value)
                                        ANSPCur = oRSet.Fields.Item("u_doccur").Value
                                        COMM = COMM + (objMatrix.Columns.Item("qty").Cells.Item(i).Specific.Value * oRSet.Fields.Item("u_comm").Value)
                                    End If
                                Next
                            Else
                                Dim flg As Boolean
                                For j As Integer = 1 To objMatrix.VisualRowCount
                                    If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(j).Specific.value) <> "" And Trim(objMatrix.Columns.Item("unitprice").Cells.Item(j).Specific.value) <> "" Then
                                        flg = True
                                    End If
                                Next
                                If flg = True Then
                                    ''Dim presqty As Integer = 0
                                    ''Dim salesqty As Integer = 0
                                    ' ''For i As Integer = 1 To objMatrix.VisualRowCount
                                    ' ''    oRSet.DoQuery("Select Quantity from RDR1 A left outer  join ORDR B ON A.DocEntry = B.DocEntry Where B.DocNum = '" + Trim(objMatrix.Columns.Item("saleno").Cells.Item(i).Specific.value) + "' and ItemCode = '" + Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) + "'")
                                    ' ''    presqty = presqty + objMatrix.Columns.Item("qty").Cells.Item(i).Specific.value
                                    ' ''    salesqty = salesqty + oRSet.Fields.Item(0).Value
                                    ' ''Next
                                    ' ''If presqty = salesqty Then
                                    ' ''    oRSet.DoQuery("Select LineTotal from RDR3 A left outer  join ORDR B ON A.DocEntry = B.DocEntry Where B.DocNum IN (" + SALDOC + ") and ExpnsCode = '11'")
                                    ' ''    Dbk = oRSet.Fields.Item(0).Value
                                    ' ''    oRSet.DoQuery("Select LineTotal from RDR3 A left outer  join ORDR B ON A.DocEntry = B.DocEntry Where B.DocNum IN (" + SALDOC + ") and ExpnsCode = '12'")
                                    ' ''    ANSP = oRSet.Fields.Item(0).Value
                                    ' ''    'oRSet.DoQuery("Select B.u_ansp,A.u_doccur,B.u_comm From [@GEN_SUPP_PRICE] A INNER JOIN [@GEN_SUPP_PRICE_D0] B On A.Code = B.Code Where B.u_ItemCode = '" + Trim(objMatrix.Columns.Item("itemcode").Cells.Item(j).Specific.value) + "' And A.u_cardcode = '" + Trim(objForm.Items.Item("custcode").Specific.value) + "'")
                                    ' ''    'ANSPCur = oRSet.Fields.Item("u_doccur").Value
                                    ' ''    oRSet.DoQuery("Select LineTotal from RDR3 A left outer  join ORDR B ON A.DocEntry = B.DocEntry Where B.DocNum IN (" + SALDOC + ") and ExpnsCode = '13'")
                                    ' ''    COMM = oRSet.Fields.Item(0).Value
                                    ' ''    'oRSet.DoQuery("Select LineTotal from RDR3 A left outer  join ORDR B ON A.DocEntry = B.DocEntry Where B.DocNum IN (" + SALDOC + ") and ExpnsCode = '15'")
                                    ' ''    'TRNSP = oRSet.Fields.Item(0).Value
                                    ' ''End If
                                    For i As Integer = 1 To objMatrix.VisualRowCount
                                        If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) <> "" And Trim(objMatrix.Columns.Item("unitprice").Cells.Item(i).Specific.value) <> "" Then
                                            oRSet.DoQuery("Select Distinct IsNull(u_expval,0) 'u_expval',isnull(u_expper,0) 'u_expper' From OITM Where ItemCode = '" + Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) + "'")
                                            Dbk = Dbk + (oRSet.Fields.Item("u_expval").Value * objMatrix.Columns.Item("qty").Cells.Item(i).Specific.value)
                                            oRSet.DoQuery("Select B.u_ansp,A.u_doccur,B.u_comm From [@GEN_SUPP_PRICE] A INNER JOIN [@GEN_SUPP_PRICE_D0] B On A.Code = B.Code Where B.u_ItemCode = '" + Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) + "' And A.u_cardcode = '" + Trim(objForm.Items.Item("custcode").Specific.value) + "'")
                                            ANSP = ANSP + (objMatrix.Columns.Item("qty").Cells.Item(i).Specific.Value * oRSet.Fields.Item("u_ansp").Value)
                                            ANSPCur = oRSet.Fields.Item("u_doccur").Value
                                            COMM = COMM + (objMatrix.Columns.Item("qty").Cells.Item(i).Specific.Value * oRSet.Fields.Item("u_comm").Value)
                                        End If
                                    Next
                                    oRSet.DoQuery("Select LineTotal from RDR3 A left outer  join ORDR B ON A.DocEntry = B.DocEntry Where B.DocNum IN (" + SALDOC + ") and ExpnsCode = '15'")
                                    TRNSP = oRSet.Fields.Item(0).Value
                                End If
                            End If
                            'If Trim(objForm.Items.Item("unit").Specific.Value) = "LG-UNIT1" Or Trim(objForm.Items.Item("unit").Specific.Value) = "UNIT1" Then
                            SALDOC = ""
                            'End If
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) <> "" And Trim(objMatrix.Columns.Item("unitprice").Cells.Item(i).Specific.value) <> "" Then
                                    LineTotalSum = LineTotalSum + (1 * CDbl(objMatrix.Columns.Item("unitprice").Cells.Item(i).Specific.Value) * objMatrix.Columns.Item("qty").Cells.Item(i).Specific.value)
                                End If
                            Next
                            ''oRSet.DoQuery("Select IsNull(PrintHeadr,0) 'Per1',IsNull(Manager,0) 'Per2' From OADM")
                            ''INS = (LineTotalSum * CDbl(oRSet.Fields.Item("Per1").Value) * CDbl(oRSet.Fields.Item("Per2").Value)) / 100
                            oRSet.DoQuery("Select IsNull(u_comper,0) AS 'COMPER' From OCRD Where CardCode = '" + Trim(objForm.Items.Item("custcode").Specific.value) + "'")
                            If CDbl(oRSet.Fields.Item("COMPER").Value) > 0 Then
                                COM = LineTotalSum * CDbl(oRSet.Fields.Item("COMPER").Value) / 100
                            End If
                            Dim TMPSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            TMPSet.DoQuery("Delete From PRE_BASENUM WHere macid = '" + MAC_ID + "'")
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                TMPSet.DoQuery("Insert Into PRE_BASENUM(invno,basenum,macid) Values('" + InvNo + "','" + "" + "','" + MAC_ID + "')")
                            Next
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                FormMode = "A"
                            End If
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                FormMode = "O"
                            End If
                            Me.Open_Accruals_Form(pVal.FormUID, InvNo, MAC_ID, FormMode, BASENUM)
                        End If
                        'Rajkumar
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        If pVal.ItemUID = "docdate" And pVal.BeforeAction = False And pVal.CharPressed = 9 And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            If objMatrix.VisualRowCount > 0 Then
                                objMatrix.Columns.Item("itemcode").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Exit Sub
                            End If
                            'ElseIf pVal.ItemUID = "ItemMatrix" And pVal.ColUID = "itemcode" And pVal.BeforeAction = False And pVal.CharPressed = 9 And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            '    objMatrix = objForm.Items.Item("ItemMatrix").Specific

                        End If
                    Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        If pVal.ItemUID = "postdate" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            If Trim(objForm.Items.Item("postdate").Specific.Value).Equals("") = False Then
                                If (DateDiff(DateInterval.Day, DateTime.ParseExact(Trim(objForm.Items.Item("postdate").Specific.Value), "yyyyMMdd", Nothing), DateTime.Today)) <> 0 Then
                                    oApplication.StatusBar.SetText("Posting date varies from system date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                End If
                            End If
                        ElseIf pVal.ItemUID = "deldate" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            If Trim(objForm.Items.Item("deldate").Specific.Value).Equals("") = False Then
                                If (DateDiff(DateInterval.Day, DateTime.ParseExact(Trim(objForm.Items.Item("postdate").Specific.Value), "yyyyMMdd", Nothing), DateTime.ParseExact(Trim(objForm.Items.Item("deldate").Specific.Value), "yyyyMMdd", Nothing))) < 0 Then
                                    oApplication.StatusBar.SetText("Delivery date is before posting date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                End If
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        objForm = oApplication.Forms.Item(pVal.FormUID)
                        If pVal.ItemUID = "rounding" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            objCheckBox = objForm.Items.Item("rounding").Specific
                            If objCheckBox.Checked = False Then
                                objForm.Items.Item("roundpr").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            Else
                                objForm.Items.Item("roundpr").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            End If

                        End If
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        oDBs_Detail = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D0")
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        If pVal.ItemUID = "ItemMatrix" And pVal.Row > 0 And pVal.BeforeAction = False Then
                            If (pVal.ColUID = "qty" Or pVal.ColUID = "total") Then
                                oDBs_Detail.Offset = pVal.Row - 1
                                oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                                oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_ItemName", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemdesc").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Price_A", oDBs_Detail.Offset, oDBs_Head.GetValue("U_DocCur", 0).Trim & " " & objMatrix.Columns.Item("unitprice").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Detail.SetValue("U_Total_A", oDBs_Detail.Offset, oDBs_Head.GetValue("U_DocCur", 0).Trim & " " & objMatrix.Columns.Item("total").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl((objMatrix.Columns.Item("total").Cells.Item(pVal.Row).Specific.Value) / (objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.Value)))
                                oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("total").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("taxcode").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_DocCur", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("doccur").Cells.Item(pVal.Row).Specific.Value))
                                'oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, (CDbl(objMatrix.Columns.Item("unitprice").Cells.Item(pVal.Row).Specific.value) * CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value)) * CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(pVal.Row).Specific.value) / 100)
                                'oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("unitprice").Cells.Item(pVal.Row).Specific.value) * CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value)) 
                                oDBs_Detail.SetValue("U_Whse", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("whse").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_SONo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("saleno").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("rem").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Discount", oDBs_Detail.Offset, "0")
                                objMatrix.SetLineData(pVal.Row)
                            ElseIf pVal.ColUID = "unitprice" Then
                                oDBs_Detail.Offset = pVal.Row - 1
                                oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                                oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_ItemName", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemdesc").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("unitprice").Cells.Item(pVal.Row).Specific.value))
                                'oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("total").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("taxcode").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_DocCur", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("doccur").Cells.Item(pVal.Row).Specific.Value))
                                'oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, (CDbl(objMatrix.Columns.Item("unitprice").Cells.Item(pVal.Row).Specific.value) * CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value)) * CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(pVal.Row).Specific.value) / 100)
                                oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("unitprice").Cells.Item(pVal.Row).Specific.value) * CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value))
                                oDBs_Detail.SetValue("U_Whse", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("whse").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_SONo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("saleno").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("rem").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Discount", oDBs_Detail.Offset, "0")
                                objMatrix.SetLineData(pVal.Row)
                            ElseIf pVal.ColUID = "price_X" Then
                                Dim untprc As String = Trim(objMatrix.Columns.Item("price_X").Cells.Item(pVal.Row).Specific.Value)
                                Dim n As Integer = splitchar(untprc)
                                Dim prc As Double = 0
                                If untprc.ToString.Length >= n Then
                                    prc = untprc.ToString.Substring(n)
                                ElseIf untprc.ToString.Length = 0 Then
                                    Exit Sub
                                Else
                                    prc = untprc
                                End If
                                oDBs_Detail.Offset = pVal.Row - 1
                                oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                                oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_ItemName", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemdesc").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Price_A", oDBs_Detail.Offset, oDBs_Head.GetValue("U_DocCur", 0).Trim & " " & prc.ToString("0.00"))
                                oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, Round(prc, 3))
                                'oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("total").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("taxcode").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_DocCur", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("doccur").Cells.Item(pVal.Row).Specific.Value))
                                'oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, (CDbl(objMatrix.Columns.Item("unitprice").Cells.Item(pVal.Row).Specific.value) * CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value)) * CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(pVal.Row).Specific.value) / 100)
                                oDBs_Detail.SetValue("U_Total_A", oDBs_Detail.Offset, oDBs_Head.GetValue("U_DocCur", 0).Trim & " " & (prc * CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value)).ToString("0.00"))
                                oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, Round(prc, 3) * CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value))
                                oDBs_Detail.SetValue("U_Whse", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("whse").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_SONo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("saleno").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("rem").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Discount", oDBs_Detail.Offset, "0")
                                objMatrix.SetLineData(pVal.Row)

                            ElseIf pVal.ColUID = "V_0" Then
                                'Dim UnitPrice As String = Trim(objMatrix.Columns.Item("Price_X").Cells.Item(pVal.Row).Specific.Value)
                                Dim untprc As String = Trim(objMatrix.Columns.Item("price_X").Cells.Item(pVal.Row).Specific.Value)
                                Dim totprc As String = Trim(objMatrix.Columns.Item("total_X").Cells.Item(pVal.Row).Specific.Value)
                                Dim DiscPer As String = Trim(objMatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                                Dim Qty1 As Double = CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.Value)
                                Dim n1 As Integer = splitchar(untprc)
                                Dim tot As Double
                                Dim Disc As Double
                                Dim d1 As Double
                                If untprc.ToString.Length > n1 Then
                                    tot = untprc.ToString.Substring(n1)
                                    d1 = tot * Qty1
                                    Disc = d1 * (DiscPer * 0.01)
                                    'qt = tot/
                                ElseIf untprc.ToString.Length = 0 Then
                                    Exit Sub
                                Else
                                    tot = untprc
                                    d1 = tot * Qty1
                                    Disc = d1 * (DiscPer * 0.01)
                                End If
                                oDBs_Detail.Offset = pVal.Row - 1
                                oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                                oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_ItemName", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemdesc").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                                'oDBs_Detail.SetValue("U_Price_A", oDBs_Detail.Offset, oDBs_Head.GetValue("U_DocCur", 0).Trim & " " & (tot / CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.Value)).ToString("0.00"))
                                oDBs_Detail.SetValue("U_Price_A", oDBs_Detail.Offset, oDBs_Head.GetValue("U_DocCur", 0).Trim & " " & untprc)

                                'oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, tot / CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, untprc)

                                'oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("total").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("taxcode").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_DocCur", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("doccur").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Discount", oDBs_Detail.Offset, DiscPer)
                                oDBs_Detail.SetValue("U_Total_A", oDBs_Detail.Offset, oDBs_Head.GetValue("U_DocCur", 0).Trim & " " & d1 - Disc)
                                oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, Round(d1 - Disc, 3))
                                oDBs_Detail.SetValue("U_Whse", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("whse").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_SONo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("saleno").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("rem").Cells.Item(pVal.Row).Specific.Value))
                                objMatrix.SetLineData(pVal.Row)

                            ElseIf pVal.ColUID = "total_X" Then
                                Dim totprc As String = Trim(objMatrix.Columns.Item("total_X").Cells.Item(pVal.Row).Specific.Value)
                                Dim n1 As Integer = splitchar(totprc)
                                Dim tot As Double
                                If totprc.ToString.Length > n1 Then
                                    tot = totprc.ToString.Substring(n1)
                                    'qt = tot/
                                ElseIf totprc.ToString.Length = 0 Then
                                    Exit Sub
                                Else
                                    tot = totprc
                                End If
                                oDBs_Detail.Offset = pVal.Row - 1
                                oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                                oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_ItemName", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemdesc").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Price_A", oDBs_Detail.Offset, oDBs_Head.GetValue("U_DocCur", 0).Trim & " " & (tot / CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.Value)).ToString("0.00"))
                                oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, tot / CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.Value))
                                'oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("total").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("taxcode").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_DocCur", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("doccur").Cells.Item(pVal.Row).Specific.Value))
                                'oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, (CDbl(objMatrix.Columns.Item("unitprice").Cells.Item(pVal.Row).Specific.value) * CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.value)) * CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(pVal.Row).Specific.value) / 100)
                                oDBs_Detail.SetValue("U_Total_A", oDBs_Detail.Offset, oDBs_Head.GetValue("U_DocCur", 0).Trim & " " & tot.ToString("0.00"))
                                oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, tot)
                                oDBs_Detail.SetValue("U_Whse", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("whse").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_SONo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("saleno").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("rem").Cells.Item(pVal.Row).Specific.Value))
                                oDBs_Detail.SetValue("U_Discount", oDBs_Detail.Offset, "0")
                                objMatrix.SetLineData(pVal.Row)
                            End If



                            Me.CalculateTotal(FormUID)
                            'For i As Integer = 1 To objMatrixRM.VisualRowCount
                            '    If Trim(objMatrixRM.Columns.Item("LineID").Cells.Item(i).Specific.Value) = Trim(objMatrix.Columns.Item("SNo").Cells.Item(pVal.Row).Specific.value) Then

                            '    End If
                            'Next
                        ElseIf pVal.ItemUID = "roundpr" And pVal.BeforeAction = False Then
                            Me.CalculateTotal(FormUID)
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        If pVal.BeforeAction = True Then
                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("Select u_unit From [@GEN_USR_UNIT] Where u_user = '" + oCompany.UserName.ToString.Trim + "'")
                            If oRecordSet.RecordCount = 0 Then
                                oApplication.StatusBar.SetText("No UNIT assigned to User", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                            If oCFL.UniqueID = "CFL_SO" Then
                                Me.SetFilterSO(FormUID, CUST_NO)
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
                            If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                                If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
                                oDBs_Detail = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D0")
                                If oCFL.UniqueID = "CFL_CUST" Then
                                    oDBs_Head.SetValue("U_CustCode", 0, oDT.GetValue("CardCode", 0))
                                    oDBs_Head.SetValue("U_CustName", 0, oDT.GetValue("CardName", 0))
                                    Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRSet.DoQuery("Select u_unit,Currency From OCRD Where CardCode = '" + Trim(oDT.GetValue("CardCode", 0)) + "'")
                                    Dim unit As String = oRSet.Fields.Item("u_unit").Value
                                    'objForm.Items.Item("unit").Specific.value = unit
                                    oDBs_Head.SetValue("U_UNIT", 0, unit)
                                    objForm.Items.Item("unit").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                    oDBs_Head.SetValue("U_JourRem", 0, "PreShipment - " + oDT.GetValue("CardCode", 0))
                                    oDBs_Head.SetValue("U_ConPer", 0, "")
                                    CUST_NO = oDBs_Head.GetValue("U_CustCode", 0).Trim
                                    Curr = oRSet.Fields.Item("Currency").Value
                                    If Curr = "##" Then
                                        Dim oRS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRS1.DoQuery("Select CurrCode,CurrName from OCRN ")
                                        objForm.Items.Item("doccur").Enabled = True
                                        objCombo1 = objForm.Items.Item("doccur").Specific
                                        For i As Integer = 1 To oRS1.RecordCount
                                            objCombo1.ValidValues.Add(Trim(oRS1.Fields.Item("CurrCode").Value), Trim(oRS1.Fields.Item("CurrName").Value))
                                            oRS1.MoveNext()
                                        Next
                                        oDBs_Head.SetValue("U_DocCur", 0, "INR")
                                    Else
                                        oDBs_Head.SetValue("U_DocCur", 0, Curr)
                                    End If
                                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                                    objMatrix.Clear()
                                    objMatrix.AddRow()
                                    objMatrix.FlushToDataSource()
                                    Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
                                ElseIf oCFL.UniqueID = "CFL_OWN" Then
                                    oDT = CFLEvent.SelectedObjects
                                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
                                    'oDBs_Head.SetValue("U_OwnerCod", 0, oDT.GetValue("empID", 0))
                                    oDBs_Head.SetValue("U_Owner", 0, oDT.GetValue("firstName", 0) + " " + oDT.GetValue("lastName", 0))
                                ElseIf oCFL.UniqueID = "CFL_PAY" Then
                                    oDBs_Head.SetValue("U_PayTrms", 0, oDT.GetValue("PymntGroup", 0))
                                ElseIf oCFL.UniqueID = "CFL_SO" Then
                                    For i As Integer = 0 To oDT.Rows.Count - 1
                                        DOCNUM = DOCNUM & "'" & Trim(oDT.GetValue("DocNum", i)) & "',"
                                        'DOCNUM = Trim(oDT.GetValue("DocNum", i))
                                    Next
                                    DOCNUM = DOCNUM.Substring(0, Len(DOCNUM) - 1)

                                ElseIf oCFL.UniqueID = "CFL_WHSE" Then
                                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                                    oDBs_Detail.Offset = pVal.Row - 1
                                    oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                                    oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemDesc").Cells.Item(pVal.Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(pVal.Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("TaxCode").Cells.Item(pVal.Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxamt").Cells.Item(pVal.Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Total").Cells.Item(pVal.Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_Whse", oDBs_Detail.Offset, Trim(oDT.GetValue("WhsCode", 0)))
                                    oDBs_Detail.SetValue("U_DocCur", oDBs_Detail.Offset, Curr)
                                    oDBs_Detail.SetValue("U_SONo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("SONo").Cells.Item(pVal.Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("remarks").Cells.Item(pVal.Row).Specific.Value))
                                    objMatrix.SetLineData(pVal.Row)
                                    'ElseIf oCFL.UniqueID = "CFL_SO" Then
                                    '    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                                    '    oDBs_Detail.Offset = pVal.Row - 1
                                    '    oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                                    '    oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value))
                                    '    oDBs_Detail.SetValue("U_ItemDesc", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("ItemDesc").Cells.Item(pVal.Row).Specific.Value))
                                    '    oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.Value))
                                    '    oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                                    '    oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("UnitPrice").Cells.Item(pVal.Row).Specific.Value))
                                    '    oDBs_Detail.SetValue("U_DCQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("DCQty").Cells.Item(pVal.Row).Specific.Value))
                                    '    oDBs_Detail.SetValue("U_GRNQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("GRNQty").Cells.Item(pVal.Row).Specific.Value))
                                    '    oDBs_Detail.SetValue("U_RetQty", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("RetQty").Cells.Item(pVal.Row).Specific.Value))
                                    '    oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("TaxCode").Cells.Item(pVal.Row).Specific.Value))
                                    '    oDBs_Detail.SetValue("U_TaxRate", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxrate").Cells.Item(pVal.Row).Specific.Value))
                                    '    oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxamt").Cells.Item(pVal.Row).Specific.Value))
                                    '    oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("Total").Cells.Item(pVal.Row).Specific.Value))
                                    '    oDBs_Detail.SetValue("U_Whs", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("Whs").Cells.Item(pVal.Row).Specific.Value))
                                    '    oDBs_Detail.SetValue("U_SONo", oDBs_Detail.Offset, Trim(oDT.GetValue("DocEntry", 0)))
                                    '    oDBs_Detail.SetValue("U_SODNo", oDBs_Detail.Offset, Trim(oDT.GetValue("DocNum", 0)))
                                    '    oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("remarks").Cells.Item(pVal.Row).Specific.Value))
                                    '    objMatrix.SetLineData(pVal.Row)
                                ElseIf oCFL.UniqueID = "CFL_TAX" Then
                                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                                    oDBs_Detail.Offset = pVal.Row - 1
                                    oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, pVal.Row)
                                    oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_ItemName", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemdesc").Cells.Item(pVal.Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(pVal.Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("unitprice").Cells.Item(pVal.Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, oDT.GetValue("Code", 0))
                                    oDBs_Detail.SetValue("U_DocCur", oDBs_Detail.Offset, Curr)
                                    oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, (CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.Value) * CDbl(objMatrix.Columns.Item("unitprice").Cells.Item(pVal.Row).Specific.Value)) * CDbl(oDT.GetValue("Rate", 0)) / 100)
                                    oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, (CDbl(objMatrix.Columns.Item("qty").Cells.Item(pVal.Row).Specific.Value) * CDbl(objMatrix.Columns.Item("unitprice").Cells.Item(pVal.Row).Specific.Value)))
                                    oDBs_Detail.SetValue("U_Whse", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("whse").Cells.Item(pVal.Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_SONo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("saleno").Cells.Item(pVal.Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("rem").Cells.Item(pVal.Row).Specific.Value))
                                    objMatrix.SetLineData(pVal.Row)
                                    Me.CalculateTotal(FormUID)

                                ElseIf oCFL.UniqueID = "CFL_ITEM" Then
                                    Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim oRecSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                                    Dim OrginRow As Integer = objMatrix.VisualRowCount
                                    For i As Integer = 0 To oDT.Rows.Count - 1
                                        Dim cflSelectedcount As Integer = oDT.Rows.Count
                                        If i < cflSelectedcount - 1 Then
                                            objMatrix.AddRow(1, pVal.Row)
                                            oDBs_Detail.InsertRecord(pVal.Row + i - 1)
                                        End If
                                        Dim cr As String
                                        If oDBs_Head.GetValue("U_DocCur", 0).Trim = "##" Then
                                            cr = "INR"
                                        Else
                                            cr = oDBs_Head.GetValue("U_DocCur", 0).Trim
                                        End If
                                        oRS.DoQuery("Select top 1 (case when DfltWH is null then (Select DfltWhs from oadm) else DfltWH end ) DfltWH from OITM where itemcode='" + oDT.GetValue("ItemCode", i) + "'")
                                        oDBs_Detail.Offset = pVal.Row - 1 + i
                                        oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, i + pVal.Row)
                                        oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, oDT.GetValue("ItemCode", i))
                                        oDBs_Detail.SetValue("U_ItemName", oDBs_Detail.Offset, oDT.GetValue("ItemName", i))
                                        oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, 1)
                                        oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, oDT.GetValue("InvntryUom", i))
                                        oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, oDT.GetValue("LastPurPrc", i))
                                        oDBs_Detail.SetValue("U_Price_A", oDBs_Detail.Offset, cr & " " & oDT.GetValue("LastPurPrc", i).ToString("0.00"))
                                        oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, "")
                                        oDBs_Detail.SetValue("U_DocCur", oDBs_Detail.Offset, Curr)
                                        oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, 0)
                                        oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, 1 * oDT.GetValue("LastPurPrc", i))
                                        oDBs_Detail.SetValue("U_Total_A", oDBs_Detail.Offset, cr & " " & (1 * oDT.GetValue("LastPurPrc", i)).ToString("0.00"))
                                        oDBs_Detail.SetValue("U_Whse", oDBs_Detail.Offset, oRS.Fields.Item("DfltWH").Value)
                                        oDBs_Detail.SetValue("U_SONo", oDBs_Detail.Offset, "")

                                        oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, "")
                                        objMatrix.SetLineData(pVal.Row + i)
                                    Next
                                    objMatrix.FlushToDataSource()
                                    'If OrginRow = pVal.Row Then
                                    '    objMatrix.AddRow()
                                    '    objMatrix.FlushToDataSource()
                                    '    Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
                                    'End If

                                    '---> Rajkumar'
                                    If objMatrix.VisualRowCount = pVal.Row Then
                                        objMatrix.AddRow()
                                        objMatrix.FlushToDataSource()
                                        Me.SetNewLine(FormUID, objMatrix.VisualRowCount)
                                    End If

                                    For Row As Integer = 1 To objMatrix.VisualRowCount
                                        objMatrix.Columns.Item("sno").Cells.Item(Row).Specific.Value = Row
                                    Next
                                    objMatrix.AutoResizeColumns()
                                    Me.CalculateTotal(FormUID)
                                End If
                            End If
                        End If
                End Select
            End If

            'If pVal.ItemUID = "copy" And pVal.FormMode = 2 And pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.BeforeAction = False And pVal.ItemChanged = True Then
            '    If pVal.ItemUID = "copy" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.BeforeAction = False And pVal.ItemChanged = True Then
            '        objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            '        Dim oCombo As SAPbouiCOM.ButtonCombo
            '        oCombo = objForm.Items.Item("copy").Specific
            '        If oCombo.Selected.Description = "GRN" Then
            '            oCombo.Caption = "Copy To"
            '            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '            oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
            '            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D0")

            '            oRS.DoQuery("Select U_CardCode,U_CardName,U_ContPer,U_VendRef,U_Owner,U_Buyer,U_PayTrms,U_JourRem,U_TotBefTa,U_Total,U_Tax,U_OwnerCod,U_PayCode, DocNum,U_DocDate from [@PRE_SHIPMENT]  where DocNum='" + objForm.Items.Item("t_docno").Specific.value + "'")
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
            '            oRS.DoQuery("select b.U_ItemCode,b.U_ItemDesc,b.U_Quantity,b.U_Price,b.U_TotalLC,b.U_TaxRate,b.U_TaxAmt,b.U_TaxCode,b.U_Whs,b.U_Remarks,b.U_UOM from   [@PRE_SHIPMENT] a join  [@PRE_SHIPMENT_D0] b on a.docentry=b.docentry   where a.DocNum='" + objForm.Items.Item("t_docno").Specific.value + "'")

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
    Sub ItemEvent_exd(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Select Case pVal.EventType
            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                If (pVal.FormTypeEx = "866") And (pVal.ItemUID = "1") And flgexd = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.ActionSuccess = True Then

                    Dim exd As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim currtime As String = DateTime.Now.ToString("yyyyMMdd")
                    exd.DoQuery("Select Rate from ORTT Where RateDate = '" + currtime + "'")
                    If exd.RecordCount = 0 Then
                        flgexd = False
                        Exit Sub
                    End If
                    exdForm = oApplication.Forms.GetForm("866", 1)
                    exdForm.Close()
                    Me.CreateForm()
                    flgexd = False
                End If
        End Select
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "PRE_SHIPMENT"
                        Dim exd As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim currtime As String = DateTime.Now.ToString("yyyyMMdd")
                        'Dim flgexd As Boolean
                        'oApplication.ActivateMenuItem("1284")
                        flgexd = True
                        exd.DoQuery("Select Rate from ORTT Where RateDate = '" + currtime + "'")
                        If exd.RecordCount = 0 Then
                            oApplication.ActivateMenuItem("3333")
                            flgexd = True
                            Exit Sub
                        End If
                        If pVal.BeforeAction = False Then
                            Me.CreateForm()
                            'objForm.Refresh()
                        End If
                    Case "1282"
                        If objForm.TypeEx = "PRE_SHIPMENT" Then
                            Me.SetDefault(objForm.UniqueID)
                            'objForm.Items.Item("approve").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            'oDBs_Head.SetValue("u_approve", 0, "Y")
                        End If
                    Case "1281"
                        If objForm.TypeEx = "PRE_SHIPMENT" Then
                            objForm.EnableMenu("1282", True)
                            objForm.Items.Item("preno").Click()
                        End If
                    Case "Close"
                        If objForm.TypeEx = "PRE_SHIPMENT" Then
                            If oApplication.MessageBox("Do you want to close?", 2, "Ok", "Cancel") = 1 Then
                                Dim ORS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                ORS.DoQuery("UPDATE [@PRE_SHIPMENT] SET U_Status='Closed' Where DocNum='" & oDBs_Head.GetValue("DocNum", 0) & "'")
                                oDBs_Head.SetValue("U_Status", 0, "Closed")
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                objForm.Items.Item("1").Enabled = True
                            End If
                        End If
                    Case "Cancel"
                        If objForm.TypeEx = "PRE_SHIPMENT" Then
                            If oApplication.MessageBox("Do you want to cancel this document?", 2, "Ok", "Cancel") = 1 Then
                                Dim ORS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                ORS.DoQuery("UPDATE [@PRE_SHIPMENT] SET U_Status='Cancelled' Where DocNum='" & oDBs_Head.GetValue("DocNum", 0) & "'")
                                oDBs_Head.SetValue("U_Status", 0, "Cancelled")
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                objForm.Items.Item("1").Enabled = True
                            End If
                        End If
                    Case "1293"
                        If objForm.TypeEx = "PRE_SHIPMENT" Then
                            If ITEM_ID.Equals("ItemMatrix") = True Then
                                objForm.Freeze(True)
                                oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
                                oDBs_Detail = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D0")
                                objMatrix = objForm.Items.Item("ItemMatrix").Specific
                                For Row As Integer = 1 To objMatrix.VisualRowCount
                                    objMatrix.GetLineData(Row)
                                    oDBs_Detail.Offset = Row - 1
                                    oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
                                    oDBs_Detail.SetValue("U_ItemCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemcode").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_ItemName", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("itemdesc").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_Quantity", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("qty").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_UOM", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("uom").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_Price", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("unitprice").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_Price_A", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("price_X").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_TaxCode", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("taxcode").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_TaxAmt", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("taxamt").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_TotalLC", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("total").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_Total_A", oDBs_Detail.Offset, CDbl(objMatrix.Columns.Item("total_X").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_Whse", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("whse").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_SONo", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("saleno").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_BaseRef", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("baseref").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_BaseLine", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("baseline").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_Note", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("note").Cells.Item(Row).Specific.Value))
                                    oDBs_Detail.SetValue("U_Remarks", oDBs_Detail.Offset, Trim(objMatrix.Columns.Item("rem").Cells.Item(Row).Specific.Value))
                                    objMatrix.SetLineData(Row)
                                Next
                                objMatrix.FlushToDataSource()
                                oDBs_Detail.RemoveRecord(oDBs_Detail.Size - 1)
                                objMatrix.LoadFromDataSource()
                                objForm.Freeze(False)
                            End If
                        End If
                    Case "1292"
                        If objForm.TypeEx = "PRE_SHIPMENT" Then
                            Try
                                If ITEM_ID.Equals("ItemMatrix") = True Then
                                    objForm.Freeze(True)
                                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
                                    oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D0")
                                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
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
                            If objForm.TypeEx = "PRE_SHIPMENT" Then
                                'BubbleEvent = False
                                sDocNum = objForm.Items.Item("preno").Specific.Value
                                sRptName = "preship.rpt"
                                Me.Report1()
                                '  Me.PrintSCRep()
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
            'objForm.Freeze(False)
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
            Dim cmd As New SqlCommand("UBG_SEPL_PRE_SHIPMENT", strcon) 'UBG_SEPL_PRE_SHIPMENT
            cmd.Connection = strcon
            cmd.CommandType = CommandType.StoredProcedure
            Dim oParameter As New SqlParameter("@DocNum", SqlDbType.NVarChar)
            oParameter.Value = Trim(objForm.Items.Item("preno").Specific.Value)
            Dim dsReport As DataSet = Helper.SqlHelper.ExecuteDataset(strcon, CommandType.StoredProcedure, "UBG_SEPL_PRE_SHIPMENT", oParameter) 'UBG_SEPL_PRE_SHIPMENT
            dsReport.WriteXml(System.IO.Path.GetTempPath() & "preship.xml", System.Data.XmlWriteMode.WriteSchema) 'preship.xml
            oUtilities.ShowReport("preship.rpt", "preship.xml") '"preship1.rpt", "preship.xml"
            strcon.Close()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Select Case BusinessObjectInfo.EventType
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                Dim objMatrix As SAPbouiCOM.Matrix
                If BusinessObjectInfo.BeforeAction = True Then
                    objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
                    oDBs_Detail = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D0")
                    'oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D3")
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                        oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("series").Specific.Selected.Value, "PRE_SHIPMENT"))
                    End If
                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                    Dim recupdt As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim recsel As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim orecint As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim chk As Integer = 0
                    Dim Line As Integer
                    Dim itm As String
                    Dim n As Boolean
                    Dim Noo As String = oDBs_Head.GetValue("DocNum", 0)
                    Dim SalesOrder As SAPbobsCOM.Documents
                    SalesOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                    recsel.DoQuery("DELETE From Temp_sales")
                    For j As Integer = 0 To objMatrix.VisualRowCount - 2
                        Dim s1 As String = "Select A.Quantity,A.U_prests,B.DocEntry,B.DocNum,A.U_preno,A.U_preqty,A.LineNum,A.ItemCode From RDR1 A inner join ORDR B on A.DocEntry = B.DocEntry Where B.DocNum = '" + oDBs_Detail.GetValue("U_SONo", j).Trim + "'  and ItemCode = '" + oDBs_Detail.GetValue("U_ItemCode", j).Trim + "'and LineNum = '" + oDBs_Detail.GetValue("U_BaseLine", j).Trim + "' and ISNULL(A.U_prests, 'Open')= 'Open'"

                        recsel.DoQuery("Select A.Quantity,A.U_prests,B.DocEntry,B.DocNum,A.U_preno,A.U_preqty,A.LineNum,A.ItemCode From RDR1 A inner join ORDR B on A.DocEntry = B.DocEntry Where B.DocNum = '" + oDBs_Detail.GetValue("U_SONo", j).Trim + "'  and ItemCode = '" + oDBs_Detail.GetValue("U_ItemCode", j).Trim + "'and LineNum = '" + oDBs_Detail.GetValue("U_BaseLine", j).Trim + "' and ISNULL(A.U_prests, 'Open')= 'Open'") '
                        Dim salqty, preqty, balqty As Double
                        Dim PRENO As String = ""
                        Dim QTY As Double = 0
                        salqty = recsel.Fields.Item(0).Value
                        Dim salno As String = recsel.Fields.Item("DocEntry").Value
                        Dim salesno As String = recsel.Fields.Item("DocNum").Value
                        Dim presno As String = Trim(oDBs_Head.GetValue("DocNum", 0))
                        PRENO = recsel.Fields.Item("U_preno").Value.ToString.Trim & "'" & Trim(oDBs_Head.GetValue("DocNum", 0)) & "',"
                        'For rc As Integer = 1 To recsel.RecordCount
                        QTY = recsel.Fields.Item("U_preqty").Value + oDBs_Detail.GetValue("U_Quantity", j).Trim
                        'PRENO = PRENO.Substring(0, Len(PRENO) - 1)
                        Line = recsel.Fields.Item("LineNum").Value
                        itm = recsel.Fields.Item("ItemCode").Value
                        preqty = QTY
                        balqty = salqty - preqty
                        If balqty = 0 Then
                            orecint.DoQuery("INSERT INTO Temp_Sales (code,line,saleno,saleent,preno,saleqty,preqty,prestat) VALUES ('" & j & "','" & Line & "','" & salesno & "','" & salno & "','" & presno & "','" & salqty & "','" & preqty & "','Closed')")
                            n = True
                        ElseIf balqty > 0 Then
                            orecint.DoQuery("INSERT INTO Temp_Sales (code,line,saleno,saleent,preno,saleqty,preqty,prestat) VALUES ('" & j & "','" & Line & "','" & salesno & "','" & salno & "','" & presno & "','" & salqty & "','" & preqty & "','Open')")
                            n = True
                        ElseIf balqty < 0 Then
                            n = False
                            oApplication.SetStatusBarMessage("Quantity falls negative", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            BubbleEvent = False
                        End If
                    Next
                    If n = True Then
                        orecint.DoQuery("Select code,line,saleno,saleent,preno,saleqty,preqty,prestat from Temp_sales")
                        For k As Integer = 1 To orecint.RecordCount

                            recupdt.DoQuery("UPDATE RDR1 SET U_preno='" & Noo & "', U_preqty = '" & orecint.Fields.Item(6).Value & "',U_prests = '" & orecint.Fields.Item(7).Value & "' Where DocEntry = '" & orecint.Fields.Item(3).Value & "' and LineNum = '" & orecint.Fields.Item(1).Value & "'")
                            orecint.MoveNext()
                        Next
                    Else
                        oApplication.SetStatusBarMessage("Preshipment not added", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    End If
                ElseIf BusinessObjectInfo.ActionSuccess = True Then
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Then
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                    End If
                End If
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                Dim objMatrix As SAPbouiCOM.Matrix
                If BusinessObjectInfo.BeforeAction = True Then
                    objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
                    oDBs_Detail = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D0")
                    'oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D3")
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                        oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("series").Specific.Selected.Value, "PRE_SHIPMENT"))
                    End If
                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                    Dim recupdt As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim recsel As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim recsel_pre As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim orecint As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim chk As Integer = 0
                    Dim Line As Integer
                    Dim itm As String
                    Dim n As Boolean
                    Dim SalesOrder As SAPbobsCOM.Documents
                    SalesOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                    For j As Integer = 0 To objMatrix.VisualRowCount - 1
                        recsel.DoQuery("Select A.Quantity,A.U_prests,B.DocEntry,B.DocNum,A.U_preno,A.U_preqty,A.LineNum,A.ItemCode From RDR1 A inner join ORDR B on A.DocEntry = B.DocEntry Where B.DocNum = '" + oDBs_Detail.GetValue("U_SONo", j).Trim + "' and ItemCode = '" + oDBs_Detail.GetValue("U_ItemCode", j).Trim + "' and ISNULL(A.U_prests, 'Open')= 'Open'")
                        recsel_pre.DoQuery("Select A.U_Quantity From [@PRE_SHIPMENT_D0] A inner join [@PRE_SHIPMENT] B on A.DocEntry = B.DocEntry Where B.DocNum = '" + oDBs_Head.GetValue("DocNum", j).Trim + "' and U_ItemCode = '" + oDBs_Detail.GetValue("U_ItemCode", j).Trim + "'")
                        Dim salqty, preqty, balqty As Double
                        Dim PRENO As String = ""
                        Dim QTY As Double = 0
                        salqty = recsel.Fields.Item(0).Value
                        Dim salno As String = recsel.Fields.Item("DocEntry").Value
                        Dim salesno As String = recsel.Fields.Item("DocNum").Value
                        Dim presno As String = Trim(oDBs_Head.GetValue("DocNum", 0))
                        'PRENO = recsel.Fields.Item("U_preno").Value.ToString.Trim & "'" & Trim(oDBs_Head.GetValue("DocNum", 0)) & "',"
                        Dim pre_qty As Integer = oDBs_Detail.GetValue("U_Quantity", j).Trim - recsel_pre.Fields.Item("U_Quantity").Value
                        QTY = recsel.Fields.Item("U_preqty").Value + pre_qty
                        'PRENO = PRENO.Substring(0, Len(PRENO) - 1)
                        Line = recsel.Fields.Item("LineNum").Value
                        itm = recsel.Fields.Item("ItemCode").Value
                        preqty = QTY
                        balqty = salqty - preqty
                        If balqty = 0 Then
                            orecint.DoQuery("INSERT INTO Temp_Sales (code,line,saleno,saleent,preno,saleqty,preqty,prestat) VALUES ('" & j & "','" & Line & "','" & salesno & "','" & salno & "','" & presno & "','" & salqty & "','" & preqty & "','Closed')")
                            n = True
                        ElseIf balqty > 0 Then
                            orecint.DoQuery("INSERT INTO Temp_Sales (code,line,saleno,saleent,preno,saleqty,preqty,prestat) VALUES ('" & j & "','" & Line & "','" & salesno & "','" & salno & "','" & presno & "','" & salqty & "','" & preqty & "','Open')")
                            n = True
                        ElseIf balqty < 0 Then
                            n = False
                            oApplication.SetStatusBarMessage("Quantity falls negative", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            BubbleEvent = False
                        End If
                    Next
                    If n = True Then
                        orecint.DoQuery("Select code,line,saleno,saleent,preno,saleqty,preqty,prestat from Temp_sales")
                        For k As Integer = 1 To orecint.RecordCount
                            recupdt.DoQuery("UPDATE RDR1 SET U_preqty = '" & orecint.Fields.Item(6).Value & "',U_prests = '" & orecint.Fields.Item(7).Value & "' Where DocEntry = '" & orecint.Fields.Item(3).Value & "' and LineNum = '" & orecint.Fields.Item(1).Value & "'")
                            orecint.MoveNext()
                        Next
                    Else
                        oApplication.SetStatusBarMessage("Preshipment not added", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    End If
                End If
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                If BusinessObjectInfo.ActionSuccess = True Then
                    objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    objForm.EnableMenu("1282", True)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
                    oDBs_Detail = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D0")
                    objMatrix = objForm.Items.Item("ItemMatrix").Specific

                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                    If Trim(oDBs_Head.GetValue("U_Status", 0)).Equals("Closed") = True Then
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                    Else
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                    End If
                    objForm.Items.Item("1").Enabled = True
                End If
        End Select
    End Sub
    Function Validation_Close(ByVal FormUID As String) As Boolean
        Try
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRSet2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRSet3 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objForm = oApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            For i As Integer = 1 To objMatrix.VisualRowCount
                Dim SOQty, PreShpQty As Double
                oRSet.DoQuery("Select B.Quantity From ORDR A Inner Join RDR1 B On A.DocEntry = B.DocEntry Where A.DocNum = '" + Trim(objMatrix.Columns.Item("saleno").Cells.Item(i).Specific.value) + "' And B.LineNum = '" + Trim(objMatrix.Columns.Item("baseline").Cells.Item(i).Specific.Value) + "'")
                If oRSet.RecordCount > 0 Then
                    SOQty = oRSet.Fields.Item("Quantity").Value
                Else
                    oRSet2.DoQuery("Select B.Quantity From ORDR A Inner Join RDR1 B ON A.DocEntry = B.DocEntry Where A.DocNum = '" + Trim(objMatrix.Columns.Item("saleno").Cells.Item(i).Specific.value) + "' And B.LineNum = '" + Trim(objMatrix.Columns.Item("baseline").Cells.Item(i).Specific.value) + "'")
                    SOQty = oRSet2.Fields.Item("Quantity").Value
                End If
                If Trim(objMatrix.Columns.Item("saleno").Cells.Item(i).Specific.value) <> "" Then
                    oRS.DoQuery("Select IsNull(Sum(B.U_Quantity),0) AS 'Quantity' From [@PRE_SHIPMENT] A Inner Join [@PRE_SHIPMENT_D0] B On A.DocEntry = B.DocEntry Where B.U_SONo = '" + Trim(objMatrix.Columns.Item("saleno").Cells.Item(i).Specific.value) + "' And B.U_BaseLine = '" + Trim(objMatrix.Columns.Item("baseline").Cells.Item(i).Specific.value) + "' And A.DocNum <> '" + Trim(objForm.Items.Item("preno").Specific.value) + "'")
                    PreShpQty = oRS.Fields.Item("Quantity").Value
                    oRS.DoQuery("Select IsNull(Sum(B.U_Quantity),0) AS 'Quantity' From [@PRE_SHIPMENT] A Inner Join [@PRE_SHIPMENT_D0] B On A.DocEntry = B.DocEntry Where B.U_SONo = '" + Trim(objMatrix.Columns.Item("saleno").Cells.Item(i).Specific.value) + "' And B.U_BaseLine = '" + Trim(objMatrix.Columns.Item("baseline").Cells.Item(i).Specific.value) + "' And A.DocNum <> '" + Trim(objForm.Items.Item("preno").Specific.value) + "' And IsNull(B.U_SoNo,'') = ''")
                    PreShpQty = PreShpQty + oRS.Fields.Item("Quantity").Value
                End If
                oRS.DoQuery("Select IsNull(Sum(B.U_Quantity),0) AS 'Quantity' From [@PRE_SHIPMENT] A Inner Join [@PRE_SHIPMENT_D0] B On A.DocEntry = B.DocEntry Where B.U_SONo = '" + Trim(objMatrix.Columns.Item("saleno").Cells.Item(i).Specific.value) + "' And B.U_BaseLine = '" + Trim(objMatrix.Columns.Item("baseline").Cells.Item(i).Specific.value) + "' And A.DocNum <> '" + Trim(objForm.Items.Item("preno").Specific.value) + "' And '" + Trim(objMatrix.Columns.Item("baseline").Cells.Item(i).Specific.value) + "' = ''")
                PreShpQty = PreShpQty + oRS.Fields.Item("Quantity").Value
                oRSet3.DoQuery("Select IsNull(u_tol,0) As 'Tolerance' From OITM Where ItemCode = '" + Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) + "'")
                If CDbl(PreShpQty + objMatrix.Columns.Item("qty").Cells.Item(i).Specific.value) > (SOQty + (SOQty * 1)) Then
                    oApplication.StatusBar.SetText("Please enter quantity less than Sales Order Quantity", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    objMatrix.Columns.Item("qty").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function
    Function Validation(ByVal FormUID As String) As Boolean
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D0")

            If Trim(objForm.Items.Item("custcode").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("CardCode should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Trim(objForm.Items.Item("postdate").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Posting Date should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Trim(objForm.Items.Item("deldate").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Delivery Date should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf Trim(objForm.Items.Item("docdate").Specific.Value).Equals("") = True Then
                oApplication.StatusBar.SetText("Document Date should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf (DateDiff(DateInterval.Day, DateTime.ParseExact(Trim(objForm.Items.Item("postdate").Specific.Value), "yyyyMMdd", Nothing), DateTime.ParseExact(Trim(objForm.Items.Item("deldate").Specific.Value), "yyyyMMdd", Nothing))) < 0 Then
                oApplication.StatusBar.SetText("Delivery date is before posting date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            If objMatrix.VisualRowCount = 1 Then
                For Row As Integer = 1 To objMatrix.VisualRowCount
                    If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(Row).Specific.Value).Equals("") = True Then
                        oApplication.StatusBar.SetText("ItemCode should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf CDbl(objMatrix.Columns.Item("qty").Cells.Item(Row).Specific.Value) <= 0 Then
                        oApplication.StatusBar.SetText("Quantity should be greater than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf Trim(objMatrix.Columns.Item("taxcode").Cells.Item(Row).Specific.Value).Equals("") = True Then
                        oApplication.StatusBar.SetText("TaxCode should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf Trim(objMatrix.Columns.Item("whse").Cells.Item(Row).Specific.Value).Equals("") = True Then
                        oApplication.StatusBar.SetText("Row [ " & Row & " ] - Row level Warehouse cannot be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
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
        'Try
        '    If oApplication.Forms.Item(eventInfo.FormUID).Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oApplication.Forms.Item(eventInfo.FormUID).Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
        '        Dim MenuItem1 As SAPbouiCOM.MenuItem
        '        Dim Menu1 As SAPbouiCOM.Menus
        '        Dim MenuParam1 As SAPbouiCOM.MenuCreationParams
        '        MenuParam1 = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        '        MenuParam1.Type = SAPbouiCOM.BoMenuType.mt_STRING
        '        MenuParam1.UniqueID = "Cancel"
        '        MenuParam1.String = "Cancel"
        '        MenuParam1.Enabled = True
        '        MenuItem1 = oApplication.Menus.Item("1280")
        '        Menu1 = MenuItem1.SubMenus
        '        If MenuItem1.SubMenus.Exists("Cancel") = False Then
        '            Menu1.AddEx(MenuParam1)
        '        End If
        '    Else
        '        ROW_ID = eventInfo.Row
        '        If oApplication.Menus1.Exists("Cancel") = True Then oApplication.Menus1.RemoveEx("Cancel")
        '    End If
        'Catch ex As Exception
        '    oApplication.StatusBar.SetText(ex.Message)
        'End Try
        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRS.DoQuery("Select USER_CODE From [OUSR] Where USER_CODE = '" + oCompany.UserName.ToString.Trim + "' And U_POEdit = 'YES'")
        If oRS.RecordCount > 0 Then
            Try
                ROW_ID = eventInfo.Row
                If eventInfo.Row > 0 Then
                    ITEM_ID = eventInfo.ItemUID
                    Dim objMatrixRM As SAPbouiCOM.Matrix
                    objMatrixRM = objForm.Items.Item("ItemMatrix").Specific
                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                    If ITEM_ID.Equals("ItemMatrix") = True Then
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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            Dim TotalLC = 0, totalTax As Double = 0
            'Dim price As Double = 0
            For Row As Integer = 1 To objMatrix.VisualRowCount
                TotalLC = TotalLC + CDbl(objMatrix.Columns.Item("total").Cells.Item(Row).Specific.Value)
                totalTax = totalTax + CDbl(objMatrix.Columns.Item("taxamt").Cells.Item(Row).Specific.Value)
                'price = objMatrix.Columns.Item("qty").Cells.Item(Row).Specific.Value / objMatrix.Columns.Item("total").Cells.Item(Row).Specific.Value
            Next
            oDBs_Head.SetValue("U_TotBefTa", 0, TotalLC)
            oDBs_Head.SetValue("U_Tax", 0, totalTax)
            oDBs_Head.SetValue("U_Total", 0, TotalLC + totalTax + objForm.Items.Item("roundpr").Specific.Value + objForm.Items.Item("freight").Specific.Value)

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
            objMatrixRM = objForm.Items.Item("ItemMatrix").Specific
            oDBs_DetailRM = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D1")
            oDBs_DetailRM.Clear()
            For Row As Integer = 1 To objMatrix.VisualRowCount
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
    'Sub Open_Pre_Freight_Gid_Form(ByVal FormUID As String, ByVal invno As Integer, ByVal macid As String)
    '    Try
    '        PARENT_FORM = FormUID
    '        Dim CHILD_FORM As String = "UBG_PRE_SHIPMENT@" & FormUID
    '        Dim oBool As Boolean = False
    '        For i As Integer = 0 To oApplication.Forms.Count - 1
    '            If oApplication.Forms.Item(i).UniqueID = (CHILD_FORM) Then
    '                objSForm = oApplication.Forms.Item(CHILD_FORM)
    '                objSForm.Select()
    '                oBool = True
    '                Exit For
    '            End If
    '        Next
    '        If oBool = False Then
    '            oUtilities.SAPXML("PreFreight.xml", CHILD_FORM)
    '            objSForm = oApplication.Forms.Item(CHILD_FORM)
    '            objSForm.Select()
    '        End If
    '        ChildModalForm = True
    '        Dim ogrid As SAPbouiCOM.Grid
    '        ogrid = objSForm.Items.Item("grd").Specific
    '        objSForm.DataSources.DataTables.Add("MyDataTable")
    '        'objSForm.DataSources.DataTables.Item(0).ExecuteQuery("Select TOP 5 * From OEXD")

    '        objSForm.DataSources.DataTables.Item(0).ExecuteQuery("Select ExpnsCode As 'Freight Code',ExpnsName As 'Freight Name', U_appltax As 'Tax Applicable',IsNull((Select u_amount From [@GEN_ACCRUALS_PRE] Where u_fcode = ExpnsCode And u_invno = '" + invno.ToString.Trim + "'") '" + invno + "'") ' And u_macid = '" + macid + "'")
    '        ogrid.DataTable = objSForm.DataSources.DataTables.Item("MyDataTable")
    '        'For i As Integer = 0 To ogrid.Columns.Count - 1
    '        '    ogrid.Columns.Item(i).Editable = False
    '        'Next
    '        'ogrid.Columns.Item(ogrid.Columns.Count - 1).Editable = True
    '        ogrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    '    Catch ex As Exception
    '        oApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub
    Sub Open_Pre_Freight_Gid_Form(ByVal FormUID As String, ByVal invno As Integer, ByVal macid As String)
        Try
            PARENT_FORM = FormUID
            Dim CHILD_FORM As String = "UBG_PRE_SHIPMENT@" & FormUID
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
                oUtilities.SAPXML("PreFreight.xml", CHILD_FORM)
                objSForm = oApplication.Forms.Item(CHILD_FORM)
                objSForm.Select()
            End If
            ChildModalForm = True
            Dim ogrid As SAPbouiCOM.Grid
            ogrid = objSForm.Items.Item("grd").Specific
            objSForm.DataSources.DataTables.Add("MyDataTable")
            'objSForm.DataSources.DataTables.Item(0).ExecuteQuery("Select TOP 5 * From OEXD")

            objSForm.DataSources.DataTables.Item(0).ExecuteQuery("Select ExpnsCode As 'Freight Code',ExpnsName As 'Freight Name', U_appltax As 'Tax Applicable',IsNull((Select u_amount From [@GEN_ACCRUALS_PRE] Where u_fcode = ExpnsCode And u_invno = '" + invno.ToString.Trim + "'") '" + invno + "'") ' And u_macid = '" + macid + "'")
            ogrid.DataTable = objSForm.DataSources.DataTables.Item("MyDataTable")
            'For i As Integer = 0 To ogrid.Columns.Count - 1
            '    ogrid.Columns.Item(i).Editable = False
            'Next
            'ogrid.Columns.Item(ogrid.Columns.Count - 1).Editable = True
            ogrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub Open_Pre_Freight_Matrix_Form(ByVal FormUID As String, ByVal invno As Integer, ByVal macid As String)
        Try
            PARENT_FORM = FormUID
            Dim CHILD_FORM As String = "UBG_PRE_SHIPMENT@" & FormUID
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
                oUtilities.SAPXML("PreFreight.xml", CHILD_FORM)
                objSForm = oApplication.Forms.Item(CHILD_FORM)
                objSForm.Select()
            End If
            ChildModalForm = True
            Dim ogrid As SAPbouiCOM.Grid
            ogrid = objSForm.Items.Item("grd").Specific
            objSForm.DataSources.DataTables.Add("MyDataTable")
            'objSForm.DataSources.DataTables.Item(0).ExecuteQuery("Select TOP 5 * From OEXD")
            objSForm.DataSources.DataTables.Item(0).ExecuteQuery("Select U_FreCode 'ExpnsCode',U_FreName 'ExpnsName',U_TaxCode 'TaxCode',U_Amt 'Amount',U_TotTax 'Total Tax Amount' From [@UBG_PRE_FRET_D0] Where u_PreNo = '" + invno.ToString.Trim + "'and u_MacId = '" + macid + "'")
            ogrid.DataTable = objSForm.DataSources.DataTables.Item("MyDataTable")
            For i As Integer = 0 To ogrid.Columns.Count - 1
                ogrid.Columns.Item(i).Editable = False
            Next
            ogrid.Columns.Item(ogrid.Columns.Count - 1).Editable = True
            ogrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub LoadItems(ByVal FormUID As String, ByVal MreqNo As String)
        Try
            Dim ITForm As SAPbouiCOM.Form
            Dim ITMatrix As SAPbouiCOM.Matrix
            'oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")

            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select B.ItemCode,B.Dscription,B.Quantity,B.unitMsr,B.Price,B.TaxCode,A.DocNum,B.LineNum,B.LineTotal,B.WhsCode,A.DocEntry,B.U_preqty,A.Address,A.NumAtCard,A.DocCur,A.DocRate,A.DocDueDate,A.SlpCode,A.OwnerCode,B.Currency,A.GroupNum,A.SlpCode,A.OwnerCode,TotalSumSy From ORDR A Inner Join RDR1 B On A.DocEntry = B.DocEntry And A.CardCode = '" + CUST_NO + "' And IsNull(A.DocStatus,'O') = 'O' and ISNULL(B.U_prests, 'Open')= 'Open' And A.DocNum IN (" + DOCNUM + ")")
            ITForm = oApplication.Forms.Item(FormUID)
            oDBs_Detail = ITForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT_D0")
            Try
                ITForm.Freeze(True)
                ITMatrix = ITForm.Items.Item("ItemMatrix").Specific
                ITMatrix.Clear()
                ITMatrix.AddRow(1)
                If oRecordSet.RecordCount = 0 Then
                    ITForm.Freeze(False)
                    Exit Sub
                End If
                oDBs_Head.SetValue("U_SaleNo", 0, DOCNUM)
                oDBs_Head.SetValue("U_CustRef", 0, oRecordSet.Fields.Item("NumAtCard").Value)
                oDBs_Head.SetValue("U_DocCur", 0, oRecordSet.Fields.Item("DocCur").Value)
                oDBs_Head.SetValue("U_Addr", 0, oRecordSet.Fields.Item("Address").Value)
                oDBs_Head.SetValue("U_Buyer", 0, oRecordSet.Fields.Item("SlpCode").Value)
                oDBs_Head.SetValue("U_Owner", 0, oRecordSet.Fields.Item("OwnerCode").Value)
                Dim dats As DateTime = oRecordSet.Fields.Item("DocDueDate").Value
                oDBs_Head.SetValue("U_DelDate", 0, dats.ToString("yyyyMMdd"))
                Dim Pay, Sale, Own As String
                Pay = oRecordSet.Fields.Item("GroupNum").Value
                Sale = oRecordSet.Fields.Item("OwnerCode").Value
                Own = oRecordSet.Fields.Item("SlpCode").Value
                Dim OrginRow As Integer = ITMatrix.VisualRowCount
                Dim rowcount As Integer = oRecordSet.RecordCount

                For i As Integer = 1 To oRecordSet.RecordCount

                   
                    ITMatrix.FlushToDataSource()

                    Dim ROWID As Integer
                    ROWID = i - 1
                   

                    oDBs_Detail.SetValue("LineId", ROWID, i)
                    oDBs_Detail.SetValue("U_ItemCode", ROWID, oRecordSet.Fields.Item("ItemCode").Value.ToString.Trim)
                    oDBs_Detail.SetValue("U_ItemName", ROWID, oRecordSet.Fields.Item("Dscription").Value.ToString.Trim)
                    oDBs_Detail.SetValue("U_Quantity", ROWID, (oRecordSet.Fields.Item("Quantity").Value - oRecordSet.Fields.Item("U_preqty").Value))
                    oDBs_Detail.SetValue("U_Price", ROWID, oRecordSet.Fields.Item("Price").Value)
                    oDBs_Detail.SetValue("U_Price_A", ROWID, Curr + " " + oRecordSet.Fields.Item("Price").Value.ToString)
                    oDBs_Detail.SetValue("U_UOM", ROWID, oRecordSet.Fields.Item("unitMsr").Value)
                    oDBs_Detail.SetValue("U_DocCur", ROWID, oRecordSet.Fields.Item("Currency").Value.ToString)
                    oDBs_Detail.SetValue("U_TaxCode", ROWID, oRecordSet.Fields.Item("TaxCode").Value)
                    oDBs_Detail.SetValue("U_Whse", ROWID, oRecordSet.Fields.Item("WhsCode").Value)
                    oDBs_Detail.SetValue("U_TotalLC", ROWID, oRecordSet.Fields.Item("TotalSumSy").Value)
                    oDBs_Detail.SetValue("U_Total_A", ROWID, Curr + " " + oRecordSet.Fields.Item("TotalSumSy").Value.ToString)
                    oDBs_Detail.SetValue("U_SONo", ROWID, oRecordSet.Fields.Item("DocNum").Value)
                    oDBs_Detail.SetValue("U_BaseLine", ROWID, oRecordSet.Fields.Item("LineNum").Value.ToString)
                    oDBs_Detail.SetValue("U_BaseRef", ROWID, oRecordSet.Fields.Item("DocEntry").Value.ToString)
                    ITMatrix.SetLineData(i)
                    oRecordSet.MoveNext()
                    ITMatrix.LoadFromDataSource()
                    'If rowcount <> ITMatrix.VisualRowCount Then
                    ITMatrix.AddRow()
                    ITMatrix.FlushToDataSource()
                    Me.SetNewLine(FormUID, ITMatrix.VisualRowCount)
                    'End If
                Next
                oRecordSet.DoQuery("SELECT T0.[PymntGroup] FROM OCTG T0 WHERE T0.[GroupNum] = '" + Pay + "'")
                oDBs_Head.SetValue("U_PayTrms", 0, oRecordSet.Fields.Item("PymntGroup").Value)
                oRecordSet.DoQuery("SELECT T0.[SlpName] FROM OSLP T0 WHERE T0.[SlpCode] = '" + Sale + "'")
                oDBs_Head.SetValue("U_Buyer", 0, oRecordSet.Fields.Item("SlpName").Value)
                oRecordSet.DoQuery("SELECT (T0.[lastName] + ' ' + T0.[firstName]) 'OwnerName' FROM OHEM T0 WHERE T0.[empID]  = '" + Own + "'")
                oDBs_Head.SetValue("U_Owner", 0, oRecordSet.Fields.Item("OwnerName").Value)
                'LoadUDF(FormUID)
                oDBs_Head.SetValue("U_Remarks", 0, "Based On Sales Order No." & "" & DOCNUM)
                Me.CalculateTotal(objForm.UniqueID)
                'ITMatrix.Columns.Item("ItemMatrix").Editable = False
                ITForm.Freeze(False)
            Catch ex As Exception
                ITForm.Freeze(False)
                oApplication.StatusBar.SetText(ex.Message)
            End Try

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub LoadUDF(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select A.[U_DelTerm], A.[U_PortLoad], A.[U_CntyFin], A.[U_PorDisch], A.[U_FinDest], A.[U_GrssWt], A.[U_NetWt], A.[U_contno], A.[U_SUPP_PLC1], A.[U_NO_OF_CN]  From ORDR A Inner Join RDR1 B On A.DocEntry = B.DocEntry And A.CardCode = '" + CUST_NO + "' And IsNull(A.DocStatus,'O') = 'O' And A.DocNum IN (" + DOCNUM + ")")
            oDBs_Head.SetValue("U_DelTerm", 0, oRecordSet.Fields.Item("U_DelTerm").Value)
            oDBs_Head.SetValue("U_PortLoad", 0, oRecordSet.Fields.Item("U_PortLoad").Value)
            oDBs_Head.SetValue("U_CntyFin", 0, oRecordSet.Fields.Item("U_CntyFin").Value)
            oDBs_Head.SetValue("U_PorDisch", 0, oRecordSet.Fields.Item("U_PorDisch").Value)
            oDBs_Head.SetValue("U_FinDest", 0, oRecordSet.Fields.Item("U_FinDest").Value)
            oDBs_Head.SetValue("U_GrssWt", 0, oRecordSet.Fields.Item("U_GrssWt").Value)
            oDBs_Head.SetValue("U_NetWt", 0, oRecordSet.Fields.Item("U_NetWt").Value)
            oDBs_Head.SetValue("U_contno", 0, oRecordSet.Fields.Item("U_contno").Value)
            'oDBs_Head.SetValue("U_SUPP_PLC1", 0, oRecordSet.Fields.Item("U_SUPP_PLC1").Value)
            'oDBs_Head.SetValue("U_NO_OF_CN", 0, oRecordSet.Fields.Item("U_NO_OF_CN").Value)

            'oDBs_Head.SetValue("U_DelTerm", 0, oRecordSet.Fields.Item("U_DelTerm").Value)
            'oDBs_Head.SetValue("U_PortLoad", 0, oRecordSet.Fields.Item("U_PortLoad").Value)
            'oDBs_Head.SetValue("U_DelTerm", 0, oRecordSet.Fields.Item("U_DelTerm").Value)
            'oDBs_Head.SetValue("U_PortLoad", 0, oRecordSet.Fields.Item("U_PortLoad").Value)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub SetFilterSO(ByVal FormUID As String, ByVal Custno As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("CFL_SO")
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Select DocNum from ORDR Where CardCode ='" & Custno & "' and (U_prests <> 'Closed' or U_prests IS NULL) and DocStatus = 'O' ")
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
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
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
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            LineTotalSum = 0
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) <> "" And Trim(objMatrix.Columns.Item("unitprice").Cells.Item(i).Specific.value) <> "" Then
                                    LineTotalSum = LineTotalSum + (1 * CDbl(objMatrix.Columns.Item("unitprice").Cells.Item(i).Specific.Value) * objMatrix.Columns.Item("qty").Cells.Item(i).Specific.value)
                                End If
                            Next
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@PRE_SHIPMENT")
                            Dbk = 0
                            ANSP = 0
                            COMM = 0
                            If Trim(objForm.Items.Item("unit").Specific.Value) = "UNIT1" Then
                                For i As Integer = 1 To objMatrix.VisualRowCount
                                    If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) <> "" And Trim(objMatrix.Columns.Item("unitprice").Cells.Item(i).Specific.value) <> "" Then
                                        oRSet.DoQuery("Select Distinct IsNull(u_expval,0) 'u_expval',isnull(u_expper,0) 'u_expper' From OITM Where ItemCode = '" + Trim(objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value) + "'")
                                        Dbk = Dbk + (oRSet.Fields.Item("u_expval").Value * objMatrix.Columns.Item("qty").Cells.Item(i).Specific.value)
                                        'oRSet.DoQuery("Select B.u_ansp,A.u_doccur,B.u_comm From [@GEN_SUPP_PRICE] A INNER JOIN [@GEN_SUPP_PRICE_D0] B On A.Code = B.Code Where B.u_ItemCode = '" + Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) + "' And A.u_cardcode = '" + Trim(objForm.Items.Item("4").Specific.value) + "'")
                                        'ANSP = ANSP + (objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value * oRSet.Fields.Item("u_ansp").Value)
                                        'ANSPCur = oRSet.Fields.Item("u_doccur").Value
                                        'COMM = COMM + (objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value * oRSet.Fields.Item("u_comm").Value)
                                    End If
                                Next
                            End If
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
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
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
                        oRSet.DoQuery("Delete From [@GEN_ACCRUALS_PRE] Where u_invno = '" + InvNo + "' And u_macid = '" + MAC_ID + "'")
                        For i As Integer = 0 To grd.Rows.Count - 1
                            If grd.DataTable.Columns.Item("Freight Amount").Cells.Item(i).Value <> "0" And grd.DataTable.Columns.Item("Freight Amount").Cells.Item(i).Value <> "" Then
                                RS.DoQuery("Select Convert(VarChar,Count(*) + 1) AS 'Code' From [@GEN_ACCRUALS_PRE]")
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
                                oRSet.DoQuery("Insert Into [@GEN_ACCRUALS_PRE] (Code,Name,u_invno,u_macid,u_fcode,u_fname,u_tax,u_amount,u_posfrgt,u_postax,u_negfrgt,u_negtax) Values('" + Trim(RS.Fields.Item("Code").Value) + "','" + Trim(RS.Fields.Item("Code").Value) + "','" + InvNo + "','" + MAC_ID + "','" + grd.DataTable.Columns.Item("Freight Code").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("Freight Name").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("Tax Applicable").Cells.Item(i).Value.ToString.Trim + "','" + grd.DataTable.Columns.Item("Freight Amount").Cells.Item(i).Value.ToString.Trim + "','" + posfcode + "','" + postax + "','" + negfcode + "','" + negtax + "') ")
                                oRSet.DoQuery("Delete From Freight_Order_Pre Where MacID = '" + MAC_ID + "'")
                                'oRSet.DoQuery("Insert Into Freight_Order_Rdr (RowNo,ExpnsCode,MacID) SELECT ROW_NUMBER() OVER (ORDER BY ExpnsName) AS Row, ExpnsCode,'" + MAC_ID + "' FROM OEXD")
                            End If
                        Next
                        FrghtFlag = True
                        'Me.CalculateTotal(FormUID)
                        objSubForm.Close()
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub ItemEvent_Freight_Form(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objFreightForm = oApplication.Forms.Item(FormUID)
                    PARENT_FORM = (pVal.FormUID.Substring(objFreightForm.UniqueID.IndexOf("@") + 1))
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
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
                        If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            oDBs_Detail_Freight = objForm.DataSources.DBDataSources.Item("@UBG_PRE_FRET_D0")
                            If oCFL.UniqueID = "CFL_TAX" Then
                                Dim FrgtMatrix As SAPbouiCOM.Matrix = FrgtForm.Items.Item("freight").Specific
                                For fr As Integer = 1 To FrgtMatrix.VisualRowCount
                                    If pVal.Row = fr Then
                                        oDBs_Detail_Freight.SetValue("U_TaxCode", fr, oDT.GetValue("Code", 0))
                                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRSet.DoQuery("SELECT Rate  FROM OSTA T0 Where Code =  '" + Trim(oDT.GetValue("Code", 0)) + "'")
                                        If Right(oDT.GetValue("Code", 0).trim, 1) = "N" Then
                                            Dim postive As String = FrgtMatrix.Columns.Item("amt").Cells.Item(fr).Specific.Value.ToString.Substring(4)
                                            Dim taxprice As String = oRSet.Fields.Item("Rate").Value * postive
                                            'oDBs_Detail_Freight.SetValue("U_TotTax", fr, taxprice)
                                            FrgtMatrix.Columns.Item("tottaxamt").Cells.Item(fr).Specific.Value = oDBs_Head.GetValue("U_DocCur", 0).Trim & -taxprice
                                        Else
                                            Dim postive As String = FrgtMatrix.Columns.Item("amt").Cells.Item(fr).Specific.Value.ToString.Substring(3)
                                            Dim taxprice As String = oRSet.Fields.Item("Rate").Value * postive
                                            'oDBs_Detail_Freight.SetValue("U_TotTax", fr, taxprice)
                                            FrgtMatrix.Columns.Item("tottaxamt").Cells.Item(fr).Specific.Value = oDBs_Head.GetValue("U_DocCur", 0).Trim & taxprice
                                        End If
                                        FrgtMatrix.Columns.Item("rem").Cells.Item(fr).Specific.Value = ""
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" Then
                        objFreightForm = oApplication.Forms.Item(FormUID)
                        Dim frieght_matrix As SAPbouiCOM.Matrix
                        frieght_matrix = objFreightForm.Items.Item("freight").Specific
                        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                        Dim frtRecords As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim frtRecord As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        If pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            Dim rnd As Integer = 0
                            frtRecords.DoQuery("Delete [@UBG_PRE_FRET_D0]")
                            For frt As Integer = 1 To frieght_matrix.VisualRowCount
                                If frieght_matrix.Columns.Item("amt").Cells.Item(frt).Specific.Value.ToString.Substring(3) <> "0.00" Then
                                    rnd = rnd + 1
                                    frtRecord.DoQuery("Insert into [@UBG_PRE_FRET_D0](Code,LineId,U_MacId,U_PreNo,U_FreCode,U_FreName,U_TaxCode,U_TotTax,U_Amt,U_Status) Values('" + rnd.ToString.Trim + "','" + rnd.ToString.Trim + "','" + MAC_ID + "','" + InvNo + "','" + Trim(frieght_matrix.Columns.Item("fretcode").Cells.Item(frt).Specific.value) + "','" + Trim(frieght_matrix.Columns.Item("fretname").Cells.Item(frt).Specific.value) + "','" + Trim(frieght_matrix.Columns.Item("taxcode").Cells.Item(frt).Specific.value) + "','" + Trim(frieght_matrix.Columns.Item("tottaxamt").Cells.Item(frt).Specific.value) + "','" + Trim(frieght_matrix.Columns.Item("amt").Cells.Item(frt).Specific.value) + "','" + Trim(frieght_matrix.Columns.Item("status").Cells.Item(frt).Specific.value) + "')")
                                End If
                            Next
                        ElseIf pVal.ActionSuccess = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                            frtRecords.DoQuery("Select SUBSTRING(U_Amt,4,LEN(U_amt)) from [@UBG_PRE_FRET_D0]")
                            objFreightForm.Close()
                            Dim frtotal As Double = 0
                            For frts As Integer = 1 To frtRecords.RecordCount
                                If frtRecords.Fields.Item(0).Value.ToString.Contains("-") = True Then
                                    frtotal = frtotal - frtRecords.Fields.Item(0).Value.ToString.Substring(1)
                                Else
                                    frtotal = frtotal + frtRecords.Fields.Item(0).Value
                                End If
                                frtRecords.MoveNext()
                            Next
                            objForm = oApplication.Forms.GetForm("PRE_SHIPMENT", oApplication.Forms.ActiveForm.TypeCount)
                            objForm.Items.Item("freight").Specific.Value = frtotal
                            Me.CalculateTotal(objForm.UniqueID)
                        End If
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
            BubbleEvent = False
        End Try
    End Sub
    Sub Open_Accruals_Form(ByVal FormUID As String, ByVal InvoiceNo As String, ByVal MACID As String, ByVal Mode As String, ByVal BASENO As String)
        Try
            PARENT_FORM = FormUID
            Dim RS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim RS2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim CHILD_FORM As String = "ACCRUALS_PRE@" & FormUID
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
                oUtilities.SAPXML("Accruals_Pre.xml", CHILD_FORM)
                objSubForm = oApplication.Forms.Item(CHILD_FORM)
                objSubForm.Select()
            End If
            ChildModalForm = True
            Dim ogrid As SAPbouiCOM.Grid
            ogrid = objSubForm.Items.Item("grd").Specific
            objSubForm.DataSources.DataTables.Add("MyDataTable")
            If Mode = "A" Then
                Dim oRs2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRs2.DoQuery("Delete From [@GEN_ACCRUALS_PRE] ")
                RS1.DoQuery("Select u_invno From [@GEN_ACCRUALS_PRE] Where u_invno = '" + InvoiceNo + "'")
                If RS1.RecordCount > 0 Then
                    objSubForm.DataSources.DataTables.Item(0).ExecuteQuery("Select ExpnsCode As 'Freight Code',ExpnsName As 'Freight Name', U_appltax As 'Tax Applicable',IsNull((Select u_amount From [@GEN_ACCRUALS_PRE] Where u_fcode = ExpnsCode And u_invno = '" + InvoiceNo + "' And u_macid = '" + MACID + "'),0) As 'Freight Amount' From OEXD Where IsNull(u_incl,'NO') = 'YES'")
                    ogrid.DataTable = objSubForm.DataSources.DataTables.Item("MyDataTable")
                Else
                    objSubForm.DataSources.DataTables.Item(0).ExecuteQuery("Exec UBG_Acrruals_Pre_BaseNum '" + InvoiceNo + "','" + MACID + "'")
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
                    If TRNSP > 0 Then
                        ogrid.DataTable.Columns.Item("Freight Amount").Cells.Item(4).Value = "INR " + CStr(TRNSP)
                    End If
                End If
            End If
            If Mode = "O" Then
                RS1.DoQuery("Select B.ExpnsCode,B.LineTotal From [@PRE_SHIPMENT] A Inner Join [@PRE_SHIPMENT_D3] B On A.DocEntry = B.DocEntry Inner Join OEXD C On B.ExpnsCode = C.ExpnsCode And C.u_incl = 'YES' Where A.DocNum = '" + InvoiceNo + "'")
                If RS1.RecordCount > 0 Then
                    objSubForm.DataSources.DataTables.Item(0).ExecuteQuery("Exec UBG_Acrruals_Pre '" + InvoiceNo + "'")
                    ogrid.DataTable = objSubForm.DataSources.DataTables.Item("MyDataTable")
                Else
                    'objSubForm.DataSources.DataTables.Item(0).ExecuteQuery("Exec UBG_Acrruals_Quotation_BaseNum '" + InvoiceNo + "','" + MACID + "'")
                    'ogrid.DataTable = objSubForm.DataSources.DataTables.Item("MyDataTable")
                End If
            End If
            RS2.DoQuery("Delete From PRE_BASENUM Where macid = '" + MACID + "'")
            ' ogrid.DataTable = objSubForm.DataSources.DataTables.Item("MyDataTable")
            For i As Integer = 0 To ogrid.Columns.Count - 1
                ogrid.Columns.Item(i).Editable = False
            Next
            ogrid.Columns.Item(ogrid.Columns.Count - 1).Editable = True
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
        'oForm.Items.Item("freightlk").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        Me.OpenFreight(FormUID)
        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
        Dim FrgtForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm

        Dim FrgtMatrix As SAPbouiCOM.Matrix = FrgtForm.Items.Item("freight").Specific
        'FrgtMatrix.AutoResizeColumns()
        For k As Integer = 1 To FrgtMatrix.VisualRowCount
            oRSet.DoQuery("Insert Into Freight_Order_Pre(RowNo,ExpnsCode,MacID) Values('" + k.ToString.Trim + "','" + Trim(FrgtMatrix.Columns.Item("fretcode").Cells.Item(k).Specific.value) + "','" + MAC_ID + "')")
        Next
        Try
            FrgtForm.Freeze(True)
            For k As Integer = 1 To FrgtMatrix.VisualRowCount
                FrgtMatrix.Columns.Item("amt").Cells.Item(k).Specific.Value = oDBs_Head.GetValue("U_DocCur", 0).Trim & "0.00"
            Next
            oRecordSet.DoQuery("Select u_fcode,u_fname,u_tax,u_amount,u_posfrgt,u_postax,u_negfrgt,u_negtax From [@GEN_ACCRUALS_PRE] Where u_invno = '" + InvNo + "' And u_macid = '" + MAC_ID + "'")
            For i As Integer = 1 To oRecordSet.RecordCount
                Dim itmcurr As String = oRecordSet.Fields.Item("u_amount").Value.ToString.Substring(0, 3)
                If itmcurr = oDBs_Head.GetValue("U_DocCur", 0).Trim Then
                    oRSet.DoQuery("Select Top 1 RowNo From Freight_Order_Pre Where ExpnsCode = '" + Trim(oRecordSet.Fields.Item("u_fcode").Value) + "' And Macid = '" + MAC_ID + "'")
                    If oRSet.RecordCount > 0 Then
                        Dim RowNo As Integer = oRSet.Fields.Item("RowNo").Value
                        If CDbl(FrgtMatrix.Columns.Item("amt").Cells.Item(RowNo).Specific.Value.ToString.Substring(3)) = 0 Then
                            FrgtMatrix.Columns.Item("amt").Cells.Item(RowNo).Specific.Value = oDBs_Head.GetValue("U_DocCur", 0).Trim & oRecordSet.Fields.Item("u_amount").Value.ToString.Substring(3)
                            FrgtMatrix.Columns.Item("taxcode").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_tax").Value
                        End If
                    End If
                    oRSet.DoQuery("Select Top 1 RowNo From Freight_Order_Pre Where ExpnsCode = '" + Trim(oRecordSet.Fields.Item("u_posfrgt").Value) + "' And Macid = '" + MAC_ID + "'")
                    If oRSet.RecordCount > 0 Then
                        RowNo = oRSet.Fields.Item("RowNo").Value
                        If CDbl(FrgtMatrix.Columns.Item("amt").Cells.Item(RowNo).Specific.Value.ToString.Substring(3)) = 0 Then
                            FrgtMatrix.Columns.Item("amt").Cells.Item(RowNo).Specific.Value = oDBs_Head.GetValue("U_DocCur", 0).Trim & oRecordSet.Fields.Item("u_amount").Value.ToString.Substring(3)
                            FrgtMatrix.Columns.Item("taxcode").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_postax").Value
                        End If
                    End If
                    oRSet.DoQuery("Select Top 1 RowNo From Freight_Order_Pre Where ExpnsCode = '" + Trim(oRecordSet.Fields.Item("u_negfrgt").Value) + "' And Macid = '" + MAC_ID + "'")
                    If oRSet.RecordCount > 0 Then
                        RowNo = oRSet.Fields.Item("RowNo").Value
                        If CDbl(FrgtMatrix.Columns.Item("amt").Cells.Item(RowNo).Specific.Value.ToString.Substring(3)) = 0 Then
                            FrgtMatrix.Columns.Item("amt").Cells.Item(RowNo).Specific.Value = oDBs_Head.GetValue("U_DocCur", 0).Trim & "-" & oRecordSet.Fields.Item("u_amount").Value.ToString.Substring(3)
                            FrgtMatrix.Columns.Item("taxcode").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_negtax").Value
                        End If
                    End If
                Else
                    Dim currtime As String = DateTime.Now.ToString("yyyyMMdd")
                    oRSet.DoQuery("Select Rate from ORTT Where RateDate = '" + currtime + "' and Currency = '" + itmcurr + "'")
                    Dim rate As Double = oRSet.Fields.Item(0).Value
                    oRSet.DoQuery("Select Rate from ORTT Where RateDate = '" + currtime + "' and Currency = '" + oDBs_Head.GetValue("U_DocCur", 0).Trim + "'")
                    Dim rate1 As Double = oRSet.Fields.Item(0).Value
                    Dim Ratevalue As String
                    If rate1 > rate Then
                        Ratevalue = 1 / rate1
                    Else
                        Ratevalue = rate
                    End If
                    Dim amt As Double = oRecordSet.Fields.Item("u_amount").Value.ToString.Substring(3)
                    oRSet.DoQuery("Select Top 1 RowNo From Freight_Order_Pre Where ExpnsCode = '" + Trim(oRecordSet.Fields.Item("u_fcode").Value) + "' And Macid = '" + MAC_ID + "'")
                    If oRSet.RecordCount > 0 Then
                        Dim RowNo As Integer = oRSet.Fields.Item("RowNo").Value
                        If CDbl(FrgtMatrix.Columns.Item("amt").Cells.Item(RowNo).Specific.Value.ToString.Substring(3)) = 0 Then
                            FrgtMatrix.Columns.Item("amt").Cells.Item(RowNo).Specific.Value = oDBs_Head.GetValue("U_DocCur", 0).Trim & (amt * Ratevalue) '.ToString.ToCharArray(0, 5)
                            FrgtMatrix.Columns.Item("taxcode").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_tax").Value
                        End If
                    End If
                    oRSet.DoQuery("Select Top 1 RowNo From Freight_Order_Pre Where ExpnsCode = '" + Trim(oRecordSet.Fields.Item("u_posfrgt").Value) + "' And Macid = '" + MAC_ID + "'")
                    If oRSet.RecordCount > 0 Then
                        RowNo = oRSet.Fields.Item("RowNo").Value
                        If CDbl(FrgtMatrix.Columns.Item("amt").Cells.Item(RowNo).Specific.Value.ToString.Substring(3)) = 0 Then
                            FrgtMatrix.Columns.Item("amt").Cells.Item(RowNo).Specific.Value = oDBs_Head.GetValue("U_DocCur", 0).Trim & (amt * Ratevalue) '.ToString.ToCharArray(0, 5)
                            FrgtMatrix.Columns.Item("taxcode").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_postax").Value
                        End If
                    End If
                    oRSet.DoQuery("Select Top 1 RowNo From Freight_Order_Pre Where ExpnsCode = '" + Trim(oRecordSet.Fields.Item("u_negfrgt").Value) + "' And Macid = '" + MAC_ID + "'")
                    If oRSet.RecordCount > 0 Then
                        RowNo = oRSet.Fields.Item("RowNo").Value
                        If CDbl(FrgtMatrix.Columns.Item("amt").Cells.Item(RowNo).Specific.Value.ToString.Substring(3)) = 0 Then
                            FrgtMatrix.Columns.Item("amt").Cells.Item(RowNo).Specific.Value = oDBs_Head.GetValue("U_DocCur", 0).Trim & "-" & (amt * Ratevalue) '.ToString.ToCharArray(0, 5)
                            FrgtMatrix.Columns.Item("taxcode").Cells.Item(RowNo).Specific.Value = oRecordSet.Fields.Item("u_negtax").Value
                        End If
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
    Function splitchar(ByRef strs As String)
        'Dim str As String
        'Dim strArr() As String
        Dim count As Integer = 0
        'str = strs
        'strArr = str.Split("")
        For v As Integer = 0 To strs.Length - 1
            If Not IsNumeric(strs.Substring(v, 1)) Then
                If (strs.Substring(v, 1)) = "." Then
                    Exit For
                End If
                count = count + 1
            End If
        Next
        Return count
    End Function
    'Sub LoadFreight(ByVal FormUID As String)
    '    Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    Dim oForm As SAPbouiCOM.Form = oApplication.Forms.Item(FormUID)
    '    Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
    '    Try
    '        'oRecordSet.DoQuery("Select U_MacId,U_FreCode,U_FreName,U_TaxCode,U_Amt,U_TotTax,U_Status,U_PreNo,U_Rem From [@UBG_PRE_FRET_D0] Where u_PreNo = '" + InvNo + "' And u_MacId = '" + MAC_ID + "'")
    '        oRecordSet.DoQuery("Select u_fcode,u_fname,u_tax,u_amount,u_posfrgt,u_postax,u_negfrgt,u_negtax From [@GEN_ACCRUALS_RDR] Where u_invno = '" + InvNo + "' And u_macid = '" + MAC_ID + "'")
    '        For i As Integer = 1 To oRecordSet.RecordCount
    '            'oRecordSet.Fields.Item("")
    '            Dim fcode, fname, tax, no As String
    '            Dim amt As Double
    '            fcode = oRecordSet.Fields.Item("u_fcode").Value
    '            fname = oRecordSet.Fields.Item("u_fname").Value
    '            tax = oRecordSet.Fields.Item("u_tax").Value
    '            amt = oRecordSet.Fields.Item("u_amount").Value.ToString.Substring(3)

    '            'oRSet.DoQuery("Insert Into [@UBG_PRE_FRET](Code,DocEntry) VALUES('" + InvNo + "','" + i + "')")
    '            'oRSet.DoQuery("Insert Into [@UBG_PRE_FRET_D0](Code,U_MacId,U_FreCode,U_FreName,U_TaxCode,U_Amt,U_TotTax,U_Status,U_PreNo,U_Rem) VALUES('" + InvNo + "','" + MAC_ID + "','" + fcode + "','" + fname + "','" + tax + "','" + amt + "','','O','','')")
    '            oRSet.DoQuery("Insert Into [@PRE_SHIPMENT_D3](DocEntry,LineId,U_ExpnCode,U_ExpnName,U_LineTot,U_TotFrgn,U_TaxCode,U_TaxType,U_Curr,U_MacId) VALUES('" + InvNo + "','" + i + "','" + fcode + "','" + fname + "','" + amt + "','" + amt + "','" + tax + "','Y','" + Curr + "','" + MAC_ID + "',)")
    '            oRecordSet.MoveNext()
    '        Next
    '        FrghtFlag = False
    '    Catch ex As Exception
    '        'FrgtForm.Freeze(False)
    '        oApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub
    Sub OpenFreight(ByVal FormUID As String)
        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oForm As SAPbouiCOM.Form = oApplication.Forms.Item(FormUID)
        Dim MAC_ID As String = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
        Try
            PARENT_FORM = FormUID
            Dim CHILD_FORM As String = "PRE_FREIGHT@" & FormUID
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
                oUtilities.SAPXML("Freight_Pre.xml", CHILD_FORM)
                objFreightForm = oApplication.Forms.Item(CHILD_FORM)
                objFreightForm.Select()
            End If
            ChildModalForm = True
            FrgtForm = oApplication.Forms.ActiveForm
            FrgtForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            Dim FrgtMatrix As SAPbouiCOM.Matrix = FrgtForm.Items.Item("freight").Specific
            FrgtMatrix.AutoResizeColumns()
            FrgtMatrix.Clear()
            FrgtForm.Items.Item("code").Specific.Value = 1
            'oRecordSet.DoQuery("Select u_fcode,u_fname,u_tax,u_amount,u_posfrgt,u_postax,u_negfrgt,u_negtax From [@GEN_ACCRUALS_PRE] Where u_invno = '" + InvNo + "' And u_macid = '" + MAC_ID + "'")
            oRecordSet.DoQuery("Select ExpnsCode,ExpnsName,(CASE WHEN DistrbMthd = 'Q' THEN 'Quantity' ELSE 'None' END) DistrbMthd From OEXD Order By ExpnsName Asc")
            FrgtForm.Freeze(True)
            For i As Integer = 1 To oRecordSet.RecordCount
                FrgtMatrix.AddRow()
                FrgtMatrix.Columns.Item("sno").Cells.Item(i).Specific.Value = i
                FrgtMatrix.Columns.Item("preno").Cells.Item(i).Specific.Value = oForm.Items.Item("preno").Specific.Value
                FrgtMatrix.Columns.Item("fretcode").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("ExpnsCode").Value
                FrgtMatrix.Columns.Item("fretname").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("ExpnsName").Value
                FrgtMatrix.Columns.Item("distmthd").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("DistrbMthd").Value
                FrgtMatrix.Columns.Item("amt").Cells.Item(i).Specific.Value = oDBs_Head.GetValue("U_DocCur", 0).Trim & "0.00"
                FrgtMatrix.Columns.Item("tottaxamt").Cells.Item(i).Specific.Value = oDBs_Head.GetValue("U_DocCur", 0).Trim & "0.00"
                FrgtMatrix.Columns.Item("status").Cells.Item(i).Specific.Value = "O"
                oRecordSet.MoveNext()
            Next
            FrghtFlag = False
            FrgtForm.Freeze(False)
        Catch ex As Exception
            FrgtForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


End Class
