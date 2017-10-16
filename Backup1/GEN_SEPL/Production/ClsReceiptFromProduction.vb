Imports System.IO
Imports System.Text
Imports System.Drawing.Printing
Imports System.Reflection
Imports System.Data.SqlClient
Imports System.Data.SqlClient.SqlDataReader
Imports System.Globalization
Public Class ClsReceiptFromProduction


#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim DateTimeFormatInfo As New System.Globalization.DateTimeFormatInfo()
    Dim objForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oItem As SAPbouiCOM.Item
    Dim oTempItem As SAPbouiCOM.Item
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim oRS As SAPbobsCOM.Recordset
    Dim oRSet As SAPbobsCOM.Recordset
    Dim User_Code As String
    Dim DocEntry As String
    Dim PurType As String
    Public MRNo As String
    Dim TransNo As String
    Dim NewPrice As Double
    Dim DocNO As String
    Dim PTNNo, PrdNo As String
    Dim SONO As String
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("13").Specific
            objForm.Title = "Confirmation of Production"
            oTempItem = objForm.Items.Item("21")
            oItem = objForm.Items.Add("sono", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OIGN", "u_sono")
            oItem.Visible = False
            oItem.LinkTo = "21"
            oItem = objForm.Items.Add("ptnno", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OIGN", "u_ptnno")
            oItem.Visible = False
            oItem.LinkTo = "16"

            oItem = objForm.Items.Add("unit", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OIGN", "U_unit")
            oItem.Visible = False
            oItem.LinkTo = "16"

            Dim BForm As SAPbouiCOM.Form
            Dim BMatrix As SAPbouiCOM.Matrix
            BForm = oApplication.Forms.Item(FormUID)
            BMatrix = BForm.Items.Item("13").Specific
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSet.DoQuery("select U_unit from owor where docnum='" + BMatrix.Columns.Item("61").Cells.Item(1).Specific.value + "'")
            BForm.Items.Item("unit").Specific.value = oRSet.Fields.Item(0).Value
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    If pVal.BeforeAction = True Then
                        Me.CreateForm(pVal.FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objMatrix = objForm.Items.Item("13").Specific
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        DocNO = objForm.Items.Item("7").Specific.Value
                        PTNNo = objForm.Items.Item("ptnno").Specific.value
                        SONO = objForm.Items.Item("sono").Specific.value
                        PrdNo = objMatrix.Columns.Item("61").Cells.Item(1).Specific.value

                    End If
                    Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim oRs2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim oRs3 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim oRs4 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim BForm As SAPbouiCOM.Form
                    Dim BMatrix As SAPbouiCOM.Matrix
                    BForm = oApplication.Forms.Item(FormUID)
                    BMatrix = BForm.Items.Item("13").Specific

                    oRs1.DoQuery("select Count(B.ItemCode) from OWOR A inner join WOR1 B on A.DocEntry=B.DocEntry where A.DocNum=(Select BaseRef From IGN1 where DocEntry=(Select Top 1 DocEntry From IGN1 order by DocEntry Desc)) and B.ItemCode like '%U1'")
                    oRs2.DoQuery("select Count(B.ItemCode) from OWOR A inner join WOR1 B on A.DocEntry=B.DocEntry where A.DocNum=(Select BaseRef From IGN1 where DocEntry=(Select Top 1 DocEntry From IGN1 order by DocEntry Desc)) and B.ItemCode like '%U2'")
                    oRs3.DoQuery("select Count(B.ItemCode) from OWOR A inner join WOR1 B on A.DocEntry=B.DocEntry where A.DocNum=(Select BaseRef From IGN1 where DocEntry=(Select Top 1 DocEntry From IGN1 order by DocEntry Desc)) and B.ItemCode like '%U3'")



                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If pVal.ItemUID = "1" And pVal.ActionSuccess = True Then

                            ' Dim DateTimeFormatInfo As New System.Globalization.DateTimeFormatInfo()
                            'DateTimeFormatInfo.ShortDatePattern = DateFormat

                            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            ' Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            ' Dim oRs2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("select A.RefDate,A.BaseRef,A.DueDate,A.TaxDate,A.Memo,A.Ref1,B.ShortName,C.AcctCode,C.FormatCode,B.Debit,B.Credit from OJDT A inner join JDT1 B on A.TransId=B.TransId inner join OACT C on B.ShortName=C.AcctCode where A.TransId=(select top 1 TransId from OJDT order by TransId desc) and (FormatCode='40002201' or FormatCode='40002101' or FormatCode='40002202' or FormatCode='40002102' or FormatCode='40002203' or FormatCode='40002103') ")
                            oRs4.DoQuery("Select u_unit from OIGN where DocNum='" + oRecordSet.Fields.Item("BaseRef").Value + "'")


                            If ((oRs1.Fields.Item(0).Value > 0 And oRs4.Fields.Item(0).Value <> "UNIT1") Or (oRs2.Fields.Item(0).Value > 0 And oRs4.Fields.Item(0).Value <> "UNIT2") Or (oRs3.Fields.Item(0).Value > 0 And oRs4.Fields.Item(0).Value <> "UNIT3")) Then


                                Dim ITForm As SAPbouiCOM.Form
                                Dim ITMatrix As SAPbouiCOM.Matrix

                                oApplication.ActivateMenuItem("1540")
                                ITForm = oApplication.Forms.GetForm("392", oApplication.Forms.ActiveForm.TypeCount)
                                ITForm.Select()
                                ITMatrix = ITForm.Items.Item("76").Specific
                                'DateTime.ParseExact(Trim(dk.SelectSingleNode("//RefDate").InnerText), "yyyyMMdd", Nothing)
                                'DateTime.ParseExact(Trim(objForm.Items.Item("t_docdt").Specific.Value), "yyyyMMdd", Nothing)
                                Dim ReferenceDate As DateTime
                                ReferenceDate = oRecordSet.Fields.Item("RefDate").Value
                                ' Dim dt As Date = "12/27/2012"
                                ITForm.Items.Item("6").Specific.Value = DateTime.Parse(ReferenceDate).ToString("yyyyMMdd")
                                '.ToString("yyyyMMdd")
                                ITForm.Items.Item("102").Specific.value = DateTime.Parse(oRecordSet.Fields.Item("DueDate").Value).ToString("yyyyMMdd")
                                ITForm.Items.Item("97").Specific.value = DateTime.Parse(oRecordSet.Fields.Item("TaxDate").Value).ToString("yyyyMMdd")
                                ' ITForm.Items.Item("25").Specific.value = oRecordSet.Fields.Item("BaseRef").Value
                                ITForm.Items.Item("10").Specific.value = oRecordSet.Fields.Item("Memo").Value
                                ITForm.Items.Item("7").Specific.value = oRecordSet.Fields.Item("Ref1").Value
                                ITForm.Items.Item("8").Specific.value = PrdNo
                                'ITMatrix.Columns.Item("U_subconln").Editable = True
                                ITForm.Freeze(True)
                                For i As Integer = 1 To oRecordSet.RecordCount
                                    ITMatrix.Columns.Item("1").Cells.Item(i).Specific.value = oRecordSet.Fields.Item("FormatCode").Value
                                    ITMatrix.Columns.Item("5").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("Credit").Value
                                    'ITMatrix.Columns.Item("6").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("Debit").Value
                                    ITMatrix.Columns.Item("10").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("Ref1").Value
                                    'ITMatrix.Columns.Item("1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    If i = oRecordSet.RecordCount Then
                                        Dim unit As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        Dim unit1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        Dim unit2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        Dim unit3 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        ' unit1.DoQuery("select A.RefDate,A.BaseRef,A.DueDate,A.TaxDate,A.Memo,A.Ref1,B.ShortName,C.AcctCode,C.FormatCode,B.Debit,B.Credit from OJDT A inner join JDT1 B on A.TransId=B.TransId inner join OACT C on B.ShortName=C.AcctCode where A.TransId=(select top 1 TransId from OJDT order by TransId desc) and (FormatCode='40002201' or FormatCode='40002101') ")
                                        unit.DoQuery("Select U_unit from OIGN where DocEntry=(Select Top 1 Docentry from OIGN order By DocEntry Desc)")
                                        unit1.DoQuery("select A.FormatCode,A.Credit,A.Debit,C.ItemCode,C.u_acct1,C.u_acct2,C.u_acct3 from (select A.RefDate,A.BaseRef,A.DueDate,A.TaxDate,A.Memo,A.Ref1,B.ShortName,C.AcctCode,C.FormatCode,B.Debit,B.Credit from OJDT A inner join JDT1 B on A.TransId=B.TransId inner join OACT C on B.ShortName=C.AcctCode where A.TransId=(select top 1 TransId from OJDT order by TransId desc) and (FormatCode='40002201' or FormatCode='40002101' or FormatCode='40002202' or FormatCode='40002102' or FormatCode='40002203' or FormatCode='40002103' ) )A inner join (select ItemCode, u_acct1,u_acct2,u_acct3 from OITM where ItemCode in (select itemcode from WOR1 where DocEntry=(select DocEntry from OWOR where DocNum=(select baseref from IGN1 where DocEntry=(select top 1 DocEntry from OIGN order by DocEntry desc))) and (ItemCode like 'MTRL%' or ItemCode like 'Process%')))C on( A.FormatCode=C.U_acct1 or A.FormatCode=C.u_acct2 or A.FormatCode=C.u_acct3) ")
                                        ' unit2.DoQuery("select A.FormatCode,A.Credit,A.Debit,C.ItemCode,C.u_acct1 from (select A.RefDate,A.BaseRef,A.DueDate,A.TaxDate,A.Memo,A.Ref1,B.ShortName,C.AcctCode,C.FormatCode,B.Debit,B.Credit from OJDT A inner join JDT1 B on A.TransId=B.TransId inner join OACT C on B.ShortName=C.AcctCode where A.TransId=(select top 1 TransId from OJDT order by TransId desc) and (FormatCode='40002201' or FormatCode='40002101') )A inner join (select ItemCode, u_acct1,u_acct2,u_acct3 from OITM where ItemCode in (select itemcode from WOR1 where DocEntry=(select DocEntry from OWOR where DocNum=(select baseref from IGN1 where DocEntry=(select top 1 DocEntry from OIGN order by DocEntry desc))) and (ItemCode like 'MTRL%' or ItemCode like 'Process%')))C on A.FormatCode=C.U_acct2 ")
                                        'unit3.DoQuery("select A.FormatCode,A.Credit,A.Debit,C.ItemCode,C.u_acct1 from (select A.RefDate,A.BaseRef,A.DueDate,A.TaxDate,A.Memo,A.Ref1,B.ShortName,C.AcctCode,C.FormatCode,B.Debit,B.Credit from OJDT A inner join JDT1 B on A.TransId=B.TransId inner join OACT C on B.ShortName=C.AcctCode where A.TransId=(select top 1 TransId from OJDT order by TransId desc) and (FormatCode='40002201' or FormatCode='40002101') )A inner join (select ItemCode, u_acct1,u_acct2,u_acct3 from OITM where ItemCode in (select itemcode from WOR1 where DocEntry=(select DocEntry from OWOR where DocNum=(select baseref from IGN1 where DocEntry=(select top 1 DocEntry from OIGN order by DocEntry desc))) and (ItemCode like 'MTRL%' or ItemCode like 'Process%')))C on A.FormatCode=C.U_acct3 ")

                                        For J As Integer = i + 1 To (oRecordSet.RecordCount + oRecordSet.RecordCount)
                                            If unit.Fields.Item(0).Value = "UNIT1" Then
                                                ITMatrix.Columns.Item("1").Cells.Item(J).Specific.value = unit1.Fields.Item("u_acct1").Value
                                                ITMatrix.Columns.Item("6").Cells.Item(J).Specific.value = unit1.Fields.Item("Credit").Value
                                            ElseIf unit.Fields.Item(0).Value = "UNIT2" Then
                                                ITMatrix.Columns.Item("1").Cells.Item(J).Specific.value = unit1.Fields.Item("u_acct2").Value
                                                ITMatrix.Columns.Item("6").Cells.Item(J).Specific.value = unit1.Fields.Item("Credit").Value
                                            ElseIf unit.Fields.Item(0).Value = "UNIT3" Then
                                                ITMatrix.Columns.Item("1").Cells.Item(J).Specific.value = unit1.Fields.Item("u_acct3").Value
                                                ITMatrix.Columns.Item("6").Cells.Item(J).Specific.value = unit1.Fields.Item("Credit").Value
                                            End If
                                            unit1.MoveNext()
                                        Next
                                    End If
                                    oRecordSet.MoveNext()
                                Next
                                'ITMatrix.Columns.Item("1").Cells.Item(oRecordSet.RecordCount + 1).Specific.value = oRecordSet.Fields.Item("FormatCode").Value
                                ' objForm.Items.Item("2").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                                ITForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                ITForm.Freeze(False)
                                ITForm.Close()
                            End If
                        End If
                    End If
                    ' End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.ActionSuccess = True Then
                        Dim oITForm As SAPbouiCOM.Form = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        Dim oMatrix As SAPbouiCOM.Matrix
                        oMatrix = oITForm.Items.Item("13").Specific
                        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        If PTNNo.Trim.ToString <> "" Then
                            oRecordSet.DoQuery("Update [@GEN_PTN] Set u_status = 'Confirmed' Where Docnum = '" + PTNNo + "' And u_sono = '" + SONO + "'")
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
                    Case "1287"
                        If objForm.TypeEx = "65214" Then
                            BubbleEvent = False
                        End If
                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
