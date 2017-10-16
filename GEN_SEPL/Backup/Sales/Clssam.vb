Imports System.Threading
Imports System.Globalization
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Public Class Clssam

#Region "        Declaration        "

    Dim oUtilities As New ClsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail1 As SAPbouiCOM.DBDataSource
    Dim oDBs_DetailRM As SAPbouiCOM.DBDataSource
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim objCheckBox As SAPbouiCOM.CheckBox
    Dim ITEM_ID As String
    Dim RowCount As Integer
    'Public sDocNum As String
    'Public sRptName As String

    Dim ROW_ID As Integer = 0
#End Region

    Sub CreateForm(ByVal Formuid As String)
        Try
            oUtilities.SAPXML("GEN_SAM.xml")
            objForm = oApplication.Forms.GetForm("GEN_SAM", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SAM")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SAM_D0")
            '  objMatrix = objForm.Items.Item("ItemMatrix").Specific

            'objForm.EnableMenu("1282", True)
            objForm.EnableMenu("1288", True)
            objForm.EnableMenu("1289", True)
            objForm.EnableMenu("1290", True)
            objForm.EnableMenu("1291", True)
            objForm.EnableMenu("1293", True)
            objForm.EnableMenu("774", True)
            objForm.DataBrowser.BrowseBy = "DocNum"
            
            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            objForm.Items.Item("18").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRS.DoQuery("select Code from OFPR A0 inner join OACP A1 on A0.Category=A1.PeriodCat where A1.Year= DATEPART(Year,GETDATE())")
            'oRS.DoQuery("select Code from OFPR A0 inner join OACP A1 on A0.Category=A1.PeriodCat where A1.Year= (CASE WHEN (DATEPART(MONTH,GETDATE())>4 and DATEPART(MONTH,GETDATE())<=12) THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END)")
            'objCombo = objForm.Items.Item("period").Specific
            'objCombo.ValidValues.Add("", "")
            'For i As Integer = 1 To oRS.RecordCount
            '    objCombo.ValidValues.Add(Trim(oRS.Fields.Item("code").Value), Trim(oRS.Fields.Item("code").Value))
            '    oRS.MoveNext()
            'Next

            oRS.DoQuery("select A1.Year from  OACP A1")
            objCombo = objForm.Items.Item("year").Specific

            For i As Integer = 1 To oRS.RecordCount
                objCombo.ValidValues.Add(Trim(oRS.Fields.Item("year").Value), Trim(oRS.Fields.Item("year").Value))
                oRS.MoveNext()
            Next

            'oRS.DoQuery("select Distinct u_unit from ORDR")
            'objCombo = objForm.Items.Item("unit").Specific
            'objCombo.ValidValues.Add("", "")
            'For i As Integer = 1 To oRS.RecordCount
            '    objCombo.ValidValues.Add(Trim(oRS.Fields.Item("u_unit").Value), Trim(oRS.Fields.Item("u_unit").Value))
            '    oRS.MoveNext()
            'Next            
            SetDefault(objForm.UniqueID)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Dim FormUID As String = objForm.UniqueID
                Select Case pVal.MenuUID
                    Case "CR_SAM"
                        If pVal.BeforeAction = False Then
                            Me.CreateForm(FormUID)
                        End If
                    Case "1282"
                        If objForm.TypeEx = "GEN_SAM" Then
                            Me.SetDefault(FormUID)
                        End If
                    Case "1293"
                        If objForm.TypeEx = "GEN_SAM" Then
                            'ITEM_ID = "ItemMatrix"

                            If ITEM_ID.Equals("ItemMatrix") = True Then
                                objForm.Freeze(True)
                                objMatrix = objForm.Items.Item("ItemMatrix").Specific
                                oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SAM")
                                oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SAM_D0")
                                For Row As Integer = 1 To objMatrix.VisualRowCount
                                    objMatrix.GetLineData(Row)
                                    'oDBs_Detail.Offset = Row - 1
                                    oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
                                    oDBs_Detail.SetValue("U_sono", oDBs_Detail.Offset, objMatrix.Columns.Item("sono").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("U_itemcode", oDBs_Detail.Offset, objMatrix.Columns.Item("itemcode").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("U_itemname", oDBs_Detail.Offset, objMatrix.Columns.Item("itemname").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("U_csam", oDBs_Detail.Offset, objMatrix.Columns.Item("csam").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("U_ssam", oDBs_Detail.Offset, objMatrix.Columns.Item("ssam").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("U_fsam", oDBs_Detail.Offset, objMatrix.Columns.Item("fsam").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("U_wstsam", oDBs_Detail.Offset, objMatrix.Columns.Item("wstsam").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("U_wstqty", oDBs_Detail.Offset, objMatrix.Columns.Item("wstqty").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("U_wfiqty", oDBs_Detail.Offset, objMatrix.Columns.Item("wfiqty").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("U_sprice", oDBs_Detail.Offset, objMatrix.Columns.Item("sprice").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("U_ccap", oDBs_Detail.Offset, objMatrix.Columns.Item("ccap").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("U_scap", oDBs_Detail.Offset, objMatrix.Columns.Item("scap").Cells.Item(Row).Specific.value)
                                    oDBs_Detail.SetValue("U_fcap", oDBs_Detail.Offset, objMatrix.Columns.Item("fcap").Cells.Item(Row).Specific.Value)

                                    ' oDBs_Detail.SetValue("u_status", oDBs_Detail.Offset, objMatrix.Columns.Item("status").Cells.Item(Row).Specific.Selected.Value)
                                    objMatrix.SetLineData(Row)
                                Next
                                objMatrix.FlushToDataSource()
                                oDBs_Detail.RemoveRecord(oDBs_Detail.Size - 1)
                                objMatrix.LoadFromDataSource()
                                objForm.Freeze(False)
                                '  objMatrix.AddRow()
                            End If
                        End If
                    Case "1281"
                        If objForm.TypeEx = "GEN_SAM" Then
                            objForm.EnableMenu("1282", True)
                            objForm.Items.Item("DocNum").Enabled = True
                            ' objMatrix.AutoResizeColumns()
                            'objForm.Items.Item("Issue").Enabled = False
                            'objForm.Items.Item("t_docno").Click()
                        End If
                    Case "1288", "1289", "1290", "1291"
                        If objForm.TypeEx = "GEN_SAM" Then
                            objForm.EnableMenu("1282", True)
                            objMatrix = objForm.Items.Item("ItemMatrix").Specific
                            '  objMatrix.AddRow()
                        End If
                End Select
            End If

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                objForm.EnableMenu("1282", False)
            End If
            '  objForm.EnableMenu("1282", True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SAM")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SAM_D0")
            'Dim strQuery As String = "SELECT  T0.[NextNumber] FROM NNM1 T0 WHERE T0.[ObjectCode] ='GEN_SAM' and  T0.[Series] ='266'"
            'Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRS.DoQuery(strQuery)
            'Dim nxt As Double = CDbl(oRS.Fields.Item("NextNumber").Value.ToString())
            'Dim a As String = objForm.BusinessObject.GetNextSerialNumber("Primary", "GEN_SAM")
            'oDBs_Head.SetValue("DocNum", 0, nxt)
            oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber("Primary", "GEN_SAM"))
            oDBs_Head.SetValue("U_cutper", 0, 70)
            oDBs_Head.SetValue("U_stitper", 0, 80)
            'oDBs_Head.SetValue("u_docnum", 0, oUtilities.keygencode("@GEN_SAM"))
            'oDBs_Head.SetValue("U_docnum", 0, oUtilities.keygencode("@GEN_SERV_BUDGET"))
            oDBs_Head.SetValue("U_date", 0, DateTime.Today.ToString("yyyyMMdd"))
            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try


            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                    'objMatrix = objForm.Items.Item("ItemMatrix").Specific
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SAM")
                    If pVal.ItemUID = "year" And pVal.ActionSuccess = True Then
                        Dim oYear As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oYear.DoQuery("select Code from OFPR A0 inner join OACP A1 on A0.Category=A1.PeriodCat where A1.Year= '" + objForm.Items.Item("year").Specific.value + "' ")
                        objCombo = objForm.Items.Item("period").Specific
                        If objCombo.ValidValues.Count > 0 Then
                            'If objCombo.ValidValues.Count > 0 Then
                            For i As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
                                objCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                            Next
                            'End If
                        End If


                        For i As Integer = 1 To oYear.RecordCount
                            objCombo.ValidValues.Add(Trim(oYear.Fields.Item("code").Value), Trim(oYear.Fields.Item("code").Value))
                            oYear.MoveNext()
                        Next
                        oDBs_Head.SetValue("u_period", 0, "")
                        If objMatrix.VisualRowCount > 0 Then
                            objMatrix.Clear()
                        End If
                    End If
                   
                    If pVal.ItemUID = "period" And pVal.BeforeAction = False Then
                        Dim Period As String
                        Period = objForm.Items.Item("period").Specific.value
                        Dim oPeriod As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oPeriod.DoQuery("select DateName(month,F_RefDate) from OFPR where Code='" + Period + "'")
                        objForm.Items.Item("month").Specific.value = oPeriod.Fields.Item(0).Value
                        If objMatrix.VisualRowCount > 0 Then
                            objMatrix.Clear()
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    If pVal.ItemUID = "ItemMatrix" And (pVal.ColUID = "ccapper") And pVal.ActionSuccess = True And pVal.CharPressed = "9" Then
                        If objMatrix.Columns.Item("ccapper").Cells.Item(pVal.Row).Specific.value > objMatrix.Columns.Item("scapper").Cells.Item(pVal.Row).Specific.value Then
                            oApplication.StatusBar.SetText("Cutting percentage is not greater than stitching percentage", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SAM")
                            objMatrix.Columns.Item("ccapper").Cells.Item(pVal.Row).Specific.value = (oDBs_Head.GetValue("U_cutper", 0).Trim)
                        End If
                        objMatrix.Columns.Item("ccap").Cells.Item(pVal.Row).Specific.value = ((objMatrix.Columns.Item("ccapper").Cells.Item(pVal.Row).Specific.value * objMatrix.Columns.Item("sprice").Cells.Item(pVal.Row).Specific.value) / 100)
                    ElseIf pVal.ItemUID = "ItemMatrix" And (pVal.ColUID = "sprice") And pVal.ActionSuccess = True And pVal.CharPressed = "9" Then
                        objMatrix.Columns.Item("ccap").Cells.Item(pVal.Row).Specific.value = ((objMatrix.Columns.Item("ccapper").Cells.Item(pVal.Row).Specific.value * objMatrix.Columns.Item("sprice").Cells.Item(pVal.Row).Specific.value) / 100)
                        objMatrix.Columns.Item("scap").Cells.Item(pVal.Row).Specific.value = ((objMatrix.Columns.Item("scapper").Cells.Item(pVal.Row).Specific.value * objMatrix.Columns.Item("sprice").Cells.Item(pVal.Row).Specific.value) / 100)
                        objMatrix.Columns.Item("fcap").Cells.Item(pVal.Row).Specific.value = ((objMatrix.Columns.Item("sprice").Cells.Item(pVal.Row).Specific.value))
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objMatrix = objForm.Items.Item("ItemMatrix").Specific
                    If pVal.ItemUID = "ItemMatrix" And (pVal.ColUID = "scapper") And pVal.ActionSuccess = True Then '(pVal.ColUID = "scapper" Or pVal.ColUID = "ccapper")
                        If objMatrix.Columns.Item("ccapper").Cells.Item(pVal.Row).Specific.value > objMatrix.Columns.Item("scapper").Cells.Item(pVal.Row).Specific.value Then
                            oApplication.StatusBar.SetText("Cutting percentage is not greater than stitching percentage", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SAM")
                            If pVal.ColUID = "scapper" Then
                                objMatrix.Columns.Item("scapper").Cells.Item(pVal.Row).Specific.value = (oDBs_Head.GetValue("U_stitper", 0).Trim)
                            Else
                                objMatrix.Columns.Item("ccapper").Cells.Item(pVal.Row).Specific.value = (oDBs_Head.GetValue("U_cutper", 0).Trim)
                            End If
                            Exit Sub
                        End If
                        objForm.Freeze(True)
                        objMatrix.Columns.Item("ccap").Cells.Item(pVal.Row).Specific.value = ((objMatrix.Columns.Item("ccapper").Cells.Item(pVal.Row).Specific.value * objMatrix.Columns.Item("sprice").Cells.Item(pVal.Row).Specific.value) / 100)
                        objMatrix.Columns.Item("scap").Cells.Item(pVal.Row).Specific.value = ((objMatrix.Columns.Item("scapper").Cells.Item(pVal.Row).Specific.value * objMatrix.Columns.Item("sprice").Cells.Item(pVal.Row).Specific.value) / 100)
                        objMatrix.Columns.Item("fcap").Cells.Item(pVal.Row).Specific.value = ((objMatrix.Columns.Item("sprice").Cells.Item(pVal.Row).Specific.value))
                        objForm.Freeze(False)
                    End If
                    If (pVal.ItemUID = "18" Or pVal.ItemUID = "20") And pVal.ActionSuccess = True Then
                        If objForm.Items.Item("18").Specific.Value > objForm.Items.Item("20").Specific.Value Or objForm.Items.Item("18").Specific.Value = "" Or objForm.Items.Item("20").Specific.Value = "" Then
                            oApplication.StatusBar.SetText("Cutting percentage is not greater than stitching percentage", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            If pVal.ItemUID = "18" Then
                                objForm.Items.Item("18").Specific.Value = 70
                            Else
                                objForm.Items.Item("20").Specific.Value = 80
                            End If
                            Exit Sub
                        End If
                    End If
                    If pVal.ItemUID = "15" And pVal.ActionSuccess = True And pVal.BeforeAction = False Then
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SAM")
                        Dim date2 As String = oDBs_Head.GetValue("U_Date", 0).ToString()
                        If date2 <> "" Then
                            Dim yy As String = date2.Substring(0, 4)
                            Dim mm As String = date2.Substring(4, 2)
                            Dim dd As String = date2.Substring(6, 2)
                            If mm <= "03" And dd <= "31" And yy <= "2016" Then
                                Dim strQuery As String = "SELECT  T0.[NextNumber] FROM NNM1 T0 WHERE T0.[ObjectCode] ='GEN_SAM' and  T0.[Series] ='266'"
                                Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRS.DoQuery(strQuery)
                                Dim nxt As Double = CDbl(oRS.Fields.Item("NextNumber").Value.ToString())
                                Dim a As String = objForm.BusinessObject.GetNextSerialNumber("Primary", "GEN_SAM")
                                oDBs_Head.SetValue("DocNum", 0, nxt)
                            Else
                                Dim strQuery As String = "SELECT  T0.[NextNumber] FROM NNM1 T0 WHERE T0.[ObjectCode] ='GEN_SAM' and  T0.[Series] ='904'"
                                Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRS.DoQuery(strQuery)
                                Dim nxt As Double = CDbl(oRS.Fields.Item("NextNumber").Value.ToString())
                                Dim a As String = objForm.BusinessObject.GetNextSerialNumber("Primary", "GEN_SAM")
                                oDBs_Head.SetValue("DocNum", 0, nxt)
                            End If
                            'Dim dt1 As DateTime = DateTime.ParseExact(date2, "yyyymmdd", CultureInfo.InvariantCulture)

                        End If

                        'Dim date1 As DateTime = Convert.ToDateTime(oDBs_Head.GetValue("U_Date", 0))
                        If date2 <= "31/3/2016" Then

                        End If
                    End If


                    'ElseIf (pVal.ItemUID = "18" Or pVal.ItemUID = "20") And pVal.ActionSuccess = True Then
                    '    objForm.Freeze(True)
                    '    If pVal.ItemUID = "18" Then
                    '        For k As Integer = 1 To objMatrix.VisualRowCount
                    '            objMatrix.Columns.Item("ccapper").Cells.Item(k).Specific.value = (oDBs_Head.GetValue("U_cutper", 0).Trim)
                    '        Next
                    '    ElseIf pVal.ItemUID = "20" Then
                    '        For k As Integer = 1 To objMatrix.VisualRowCount
                    '            objMatrix.Columns.Item("scapper").Cells.Item(k).Specific.value = (oDBs_Head.GetValue("U_stitper", 0).Trim)
                    '        Next
                    '    End If
                    '    objForm.Freeze(False)
                    '    Exit Sub
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If (pVal.ItemUID = "1" Or pVal.ItemUID = "report") And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim oValidate As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oValidate.DoQuery("Select u_period from [@GEN_SAM] where u_period='" + objForm.Items.Item("period").Specific.value + "' and u_unit='" + objForm.Items.Item("unit").Specific.value + "'")
                        If oValidate.RecordCount > 0 Then
                            oApplication.StatusBar.SetText("SAM Data was already entered for this Posting Period", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If

                        For i As Integer = 1 To objMatrix.VisualRowCount
                            If objMatrix.Columns.Item("sprice").Cells.Item(i).Specific.value = 0 Then
                                oApplication.StatusBar.SetText("SalePrice is missing in Row- " & i & "", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        Next
                    End If
                    If (pVal.ItemUID = "1") And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        If objMatrix.RowCount = 0 Then
                            BubbleEvent = False
                            Throw New Exception("Cannot add without any Data")
                        End If

                    End If
                    If pVal.ItemUID = "ItemMatrix" And pVal.ColUID = "sono" And pVal.BeforeAction = True Then
                        Me.FilterSO(FormUID)
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        For i As Integer = 1 To objMatrix.VisualRowCount
                            If objMatrix.Columns.Item("sprice").Cells.Item(i).Specific.value = 0 Then
                                oApplication.StatusBar.SetText("SalePrice is missing in Row- " & i & "", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        Next
                    End If
                    If pVal.ItemUID = "1" And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(FormUID)
                    End If

                    If pVal.ItemUID = "report" And pVal.ActionSuccess = True And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRSet1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRS.DoQuery("Select Convert(Varchar,F_RefDate,112),Convert(Varchar,T_RefDate,112) from OFPR where code='" + Trim(objForm.Items.Item("period").Specific.value) + "'")
                        'oRSet.DoQuery("Select A0.docnum,A1.Itemcode,A1.Dscription,A1.Linetotal From ORDR A0 Inner join RDR1 A1 on A0.Docentry=A1.Docentry Where A0.u_unit='" + Trim(objForm.Items.Item("unit").Specific.value) + "' and A0.Docdate between '" + Trim(oRS.Fields.Item(0).Value) + "' and '" + Trim(oRS.Fields.Item(1).Value) + "' and A0.DocStatus='O' order by A0.Docnum")
                        Dim _str_qry As String = "SELECT *,'1' 'CSam','0' 'SSam','1' 'FSam','0' 'wstsam' FROM (SELECT DISTINCT A2.OriginNum,A3.DocNum,A3.DocDate,(CASE WHEN A1.WhsCode LIKE '%-1' THEN 'UNIT1' WHEN A1.WhsCode LIKE '%-2' THEN 'UNIT2' WHEN A1.WhsCode LIKE '%-3' THEN 'UNIT3' ELSE NULL END)'UNIT',(CASE WHEN a1.ItemCode LIKE '%-%' THEN substring(A1.ItemCode, 0, charindex('-',A1.ItemCode)) ELSE A1.ItemCode END)'ITEM',(SELECT ItemName FROM OITM WHERE ItemCode=(CASE WHEN a1.ItemCode LIKE '%-%' THEN substring(A1.ItemCode, 0, charindex('-',A1.ItemCode)) ELSE A1.ItemCode END))'ITEMNAME',(SELECT distinct  Price FROM ORDR T0 INNER JOIN RDR1 T1 ON T0.DocEntry=T1.DocEntry WHERE T0.DocNum=A3.DocNum AND T1.ItemCode=(CASE WHEN a1.ItemCode LIKE '%-%' THEN substring(A1.ItemCode, 0, charindex('-',A1.ItemCode)) ELSE A1.ItemCode END))'Price' FROM OIGN A0 INNER JOIN IGN1 A1 ON A0.DocEntry=A1.DocEntry INNER JOIN OWOR A2 ON A2.DocNum=A1.BaseRef INNER JOIN ORDR A3 ON A2.OriginNum=A3.DocNum INNER JOIN RDR1 A4 ON A3.DocEntry=A4.DocEntry WHERE A0.DocDate between '" + Trim(oRS.Fields.Item(0).Value) + "' and '" + Trim(oRS.Fields.Item(1).Value) + "' AND (A4.ItemCode IN (SELECT DISTINCT (CASE WHEN T1.ItemCode LIKE '%-%' THEN substring(T1.ItemCode, 0, charindex('-',T1.ItemCode)) ELSE T1.ItemCode END) FROM OIGN T0 INNER JOIN IGN1 T1 ON T0.DocEntry=T1.DocEntry WHERE T0.DocDate between '" + Trim(oRS.Fields.Item(0).Value) + "' and '" + Trim(oRS.Fields.Item(1).Value) + "')) )A WHERE UNIT='" + Trim(objForm.Items.Item("unit").Specific.value) + "'"
                        oRSet.DoQuery(_str_qry)

                        objMatrix.Clear()
                        oApplication.SetStatusBarMessage("Processing Data", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                        objForm.Freeze(True)
                        Dim R11 As Int16 = oRSet.RecordCount
                        For i As Integer = 1 To oRSet.RecordCount
                            objMatrix.AddRow()
                            objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.value = i
                            objMatrix.Columns.Item("sono").Cells.Item(i).Specific.value = oRSet.Fields.Item(0).Value
                            'objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value = Trim(oRSet.Fields.Item(4).Value)
                            oRSet1.DoQuery("select itemcode,Itemname from OITM where itemcode='" + Trim(oRSet.Fields.Item(4).Value) + "'")
                            If oRSet1.RecordCount > 0 Then
                                objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value = Trim(oRSet1.Fields.Item(0).Value)
                                objMatrix.Columns.Item("itemname").Cells.Item(i).Specific.value = Trim(oRSet1.Fields.Item(1).Value)
                            Else
                                Dim k As Int16 = i
                                Dim w As String = "select min(itemcode) a,Itemname from oitm where ItemCode like '%" + Trim(oRSet.Fields.Item(4).Value) + "%' group by ItemName"

                                oRSet1.DoQuery(w)
                                objMatrix.Columns.Item("itemcode").Cells.Item(i).Specific.value = Trim(oRSet1.Fields.Item(0).Value)
                                objMatrix.Columns.Item("itemname").Cells.Item(i).Specific.value = Trim(oRSet1.Fields.Item(1).Value)
                            End If

                            'objMatrix.Columns.Item("itemname").Cells.Item(i).Specific.value = Trim(oRSet.Fields.Item(5).Value)
                            objMatrix.Columns.Item("ccapper").Cells.Item(i).Specific.value = (oDBs_Head.GetValue("U_cutper", 0).Trim)
                            objMatrix.Columns.Item("scapper").Cells.Item(i).Specific.value = (oDBs_Head.GetValue("U_stitper", 0).Trim)
                            objMatrix.Columns.Item("sprice").Cells.Item(i).Specific.value = Trim(oRSet.Fields.Item(6).Value)
                            objMatrix.Columns.Item("csam").Cells.Item(i).Specific.value = Trim(oRSet.Fields.Item(7).Value)
                            objMatrix.Columns.Item("ssam").Cells.Item(i).Specific.value = Trim(oRSet.Fields.Item(8).Value)
                            objMatrix.Columns.Item("fsam").Cells.Item(i).Specific.value = Trim(oRSet.Fields.Item(9).Value)
                            objMatrix.Columns.Item("wstsam").Cells.Item(i).Specific.value = Trim(oRSet.Fields.Item(10).Value)
                            oRSet.MoveNext()
                            Dim R1 As Int16 = i
                        Next
                        objMatrix.AutoResizeColumns()
                        objForm.Freeze(False)
                        objForm.Items.Item("18").Enabled = False
                        objForm.Items.Item("20").Enabled = False
                    End If
                    If pVal.ItemUID = "report" Or pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                        If objForm.Items.Item("year").Specific.value = "" Then
                            oApplication.StatusBar.SetText("Please Select Year", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            BubbleEvent = False
                            Exit Sub
                        ElseIf objForm.Items.Item("period").Specific.value = "" Then
                            oApplication.StatusBar.SetText("Please Select Period", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        ElseIf objForm.Items.Item("unit").Specific.value = "" Then
                            oApplication.StatusBar.SetText("Please Select Unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If pVal.ItemUID = "period" And pVal.BeforeAction = True Then
                        If objForm.Items.Item("year").Specific.value = "" Then
                            oApplication.StatusBar.SetText("Please Select Year", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                            'ElseIf objForm.Items.Item("period").Specific.value = "" Then
                            '    oApplication.StatusBar.SetText("Please Select Peiod", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            'ElseIf objForm.Items.Item("unit").Specific.value = "" Then
                            '    oApplication.StatusBar.SetText("Please Select Unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        End If
                        If pVal.ItemUID = "ItemMatrix" And pVal.ColUID = "itemcode" And pVal.BeforeAction = True Then
                            Me.FilterItem(FormUID, pVal.Row)
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    If pVal.ItemUID = "ItemMatrix" And pVal.ColUID = "sono" And pVal.BeforeAction = True Then
                        objForm = oApplication.Forms.Item(pVal.FormUID)
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        If Trim(objMatrix.Columns.Item("sono").Cells.Item(objMatrix.VisualRowCount).Specific.Value).Equals("") <> True Then
                            objMatrix.AddRow()
                            objMatrix.FlushToDataSource()
                            SetNewLine(objForm.UniqueID, objMatrix.VisualRowCount)
                        End If
                    End If
                    If pVal.ItemUID = "ItemMatrix" And pVal.ColUID = "itemcode" And pVal.BeforeAction = True Then
                        Me.FilterItem(FormUID, pVal.Row)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects


                    If pVal.BeforeAction = False Then


                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SAM")
                        oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SAM_D0")
                        If oCFL.UniqueID = "CFL_SO" Then
                            objForm.Items.Item("15").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            'objMatrix.Columns.Item("NWID").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Try
                                '  objMatrix.Columns.Item("sono").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                objMatrix.Columns.Item("sono").Cells.Item(pVal.Row).Specific.Value = oDT.GetValue("DocNum", 0)
                            Catch ex As Exception
                            End Try

                        End If
                        If oCFL.UniqueID = "CFL_ITEM" Then



                            'objMatrix.Columns.Item("NWID").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            objForm.Items.Item("15").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            'If oDT.GetValue("ItemCode", 0).Equals("") <> True Then
                            Try

                                oRS.DoQuery("select price from ordr inner join rdr1 on ordr.DocEntry =RDR1.DocEntry where DocNum='" & objMatrix.Columns.Item("sono").Cells.Item(pVal.Row).Specific.Value & "'  and  ItemCode ='" & oDT.GetValue("ItemCode", 0) & "' ")

                            Catch ex As Exception
                                Exit Sub
                            End Try
                            'End If
                            Try
                                objMatrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific.Value = oDT.GetValue("ItemCode", 0)
                                objMatrix.Columns.Item("itemname").Cells.Item(pVal.Row).Specific.Value = oDT.GetValue("ItemName", 0)
                                objMatrix.Columns.Item("sprice").Cells.Item(pVal.Row).Specific.Value = oRS.Fields.Item("price").Value
                            Catch ex As Exception
                                objMatrix.Columns.Item("itemname").Cells.Item(pVal.Row).Specific.Value = oDT.GetValue("ItemName", 0)
                                objMatrix.Columns.Item("sprice").Cells.Item(pVal.Row).Specific.Value = oRS.Fields.Item("price").Value
                            End Try

                        End If

                    End If

                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "ItemMatrix" And pVal.ColUID = "itemcode" And pVal.BeforeAction = True Then
                        Me.FilterItem(FormUID, pVal.Row)
                    End If


            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer) ', ByRef objMatrix As SAPbouiCOM.Matrix) 'For FG Tab
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SAM")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SAM_D0")
            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, objMatrix.VisualRowCount)
            oDBs_Detail.SetValue("u_sono", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_itemcode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_itemname", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_csam", oDBs_Detail.Offset, 1)
            oDBs_Detail.SetValue("u_ssam", oDBs_Detail.Offset, 1)
            oDBs_Detail.SetValue("u_fsam", oDBs_Detail.Offset, 1)
            oDBs_Detail.SetValue("u_wstsam", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("u_wstqty", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("u_wfiqty", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("u_sprice", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("u_ccap", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("u_scap", oDBs_Detail.Offset, 0)
            oDBs_Detail.SetValue("u_fcap", oDBs_Detail.Offset, 0)

            objMatrix.SetLineData(Row)
            objMatrix.FlushToDataSource()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            RowCount = eventInfo.Row
            Dim oForm As SAPbouiCOM.Form = oApplication.Forms.Item(eventInfo.FormUID)
            ITEM_ID = eventInfo.ItemUID
            If eventInfo.Row > 0 Then
                ITEM_ID = eventInfo.ItemUID
                objMatrix = oForm.Items.Item("ItemMatrix").Specific
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
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE, SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = True Then

                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SAM")
                        oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SAM_D0")
                        objMatrix = objForm.Items.Item("ItemMatrix").Specific
                        If oDBs_Detail.GetValue("u_sono", oDBs_Detail.Size - 1).Equals("") = True Or oDBs_Detail.GetValue("u_sono", oDBs_Detail.Size - 1).Equals("''") = True Then
                            oDBs_Detail.RemoveRecord(oDBs_Detail.Size - 1)
                        End If

                    End If

            End Select
        Catch ex As Exception

        End Try
    End Sub
    Sub FilterItem(ByVal FormUID As String, ByVal row As Integer)
        Try

            objForm = oApplication.Forms.ActiveForm
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("CFL_ITEM")
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            objMatrix = objForm.Items.Item("ItemMatrix").Specific
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("select ItemCode,Dscription from ORDR  inner join rdr1  on ordr.DocEntry=rdr1.DocEntry where DocNum = '" & Trim(objMatrix.Columns.Item("sono").Cells.Item(row).Specific.value) & "' ")
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
    Sub FilterSO(ByVal FormUID As String)
        Try

            objForm = oApplication.Forms.ActiveForm
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("CFL_SO")
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim _str_qry As String = "select ordr.DocNum from   ORDR " _
                                        & "inner join OFPR on OFPR.AbsEntry=ORDR.FinncPriod " _
                                        & "inner join OWOR on ORDR.DocNum =OWOR.OriginNum " _
                                        & "inner join IGN1 on IGN1.BaseEntry  =OWOR.DocEntry " _
                                       & "where Code = '" & Trim(objForm.Items.Item("period").Specific.Value) & "' and DocStatus ='O'  and " _
                                       & "ordr.U_Unit='" & objForm.Items.Item("unit").Specific.value & "'"
            oRS.DoQuery(_str_qry)
            For i As Integer = 0 To oRS.RecordCount - 1
                If i > 0 Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                oCon = oCons.Add()
                oCon.Alias = "Docnum"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = oRS.Fields.Item("Docnum").Value
                oRS.MoveNext()
            Next
            If oRS.RecordCount = 0 Then
                oCon = oCons.Add()
                oCon.Alias = "Docnum"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "-1"
            End If
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
End Class
