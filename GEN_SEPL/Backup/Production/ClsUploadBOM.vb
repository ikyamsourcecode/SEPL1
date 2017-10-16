Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Threading
Imports System.Globalization
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports SAPbobsCOM
Imports SAPbouiCOM
Imports System
Imports System.Security.Permissions
Imports System.Windows.Forms
Imports System.management

Public Class ClsUploadBOM


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

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.ItemUID = "btnload" And pVal.BeforeAction = True Then

                        'Dim oItem As SAPbouiCOM.Item
                        'oItem = objForm.Items.Item("dlg")
                        'oItem.Visible = True
                        'oItem.FontSize = 17
                        'oItem.TextStyle = 2
                        'Me.UploadFromExcel(objForm)

                        objForm = oApplication.Forms.Item(FormUID)
                        Dim oEditText As SAPbouiCOM.EditText = objForm.Items.Item("eFileName").Specific
                        'oEditText.Value = "C:\CustomBOM\BOM_EXCEL.xlsx"
                        oEditText.Value = "V:\CustomBOM\BOM_EXCEL.xlsm"

                        'Me.BrowseFileDialog(objForm)

                    End If

                    If pVal.ItemUID = "btnbom" And pVal.BeforeAction = True Then
                        objForm = oApplication.Forms.Item(FormUID)
                        Dim oItem As SAPbouiCOM.Item
                        oItem = objForm.Items.Item("dlg")
                        oItem.Visible = True
                        oItem.FontSize = 17
                        oItem.TextStyle = 2

                        Me.UploadFromExcel(objForm)

                        'Me.Add_BOM(FormUID, pVal)
                    End If

                    If pVal.ItemUID = "btnbom" And pVal.BeforeAction = False And pVal.ActionSuccess = True Then
                        objForm = oApplication.Forms.Item(FormUID)
                        Dim oItem As SAPbouiCOM.Item
                        Dim oEditText As SAPbouiCOM.EditText
                        oItem = objForm.Items.Item("dlg")
                        oItem.Visible = False
                        oEditText = objForm.Items.Item("eFileName").Specific
                        oEditText.Value = ""
                    End If
                    If pVal.ItemUID = "btnload" And pVal.BeforeAction = False And pVal.ActionSuccess = True Then
                        'Dim oItem As SAPbouiCOM.Item
                        'oItem = objForm.Items.Item("dlg")
                        'oItem.Visible = False
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

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
                    Case "UPLBOM"
                        Me.CreateForm()
                    Case "1282"

                    Case "1281"

                    Case "1288", "1289", "1290", "1291"

                    Case "1293"

                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = True Then

                    ElseIf BusinessObjectInfo.ActionSuccess = True Then

                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then

                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CreateForm()
        Try
            oUtilities.SAPXML("UploadExcel1.xml")
            'oUtilities.SAPXML("UploadExcel.xml")
            objForm = oApplication.Forms.ActiveForm
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Catch ex As Exception

        End Try
    End Sub

    Public Sub UploadFromExcel(ByRef objForm As SAPbouiCOM.Form)
        Try

            Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRset1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim _str_RefNo As String = ""
            Dim _str_ItemCode As String = ""
            Dim _str_Unit As String = ""
            Dim _str_SaleOrdNo As String = ""
            Dim _str_Combo As String = ""
            Dim _str_ComboHeader As String = ""
            Dim _str_ItemCodeL As String = ""
            Dim _str_CnsQty As String = ""
            Dim _str_Wastage As String = ""
            Dim _str_GTotal As String = ""
            Dim _str_Process As String = ""
            Dim _str_VendorCode As String = ""
            Dim _str_VendorName As String = ""
            Dim _str_Placement As String = ""
            Dim _str_FRemarks As String = ""
            Dim _str_Remarks As String = ""
            Dim _str_Description As String = ""
            Dim _str_Color As String = ""
            Dim _str_Size As String = ""
            Dim _str_User As String = ""
            Dim _str_StyleCode As String = ""
            Dim _str_StyleCodeL As String = ""
            Dim _str_SaleOrderNoL As String = ""
            Dim _str_DocEntry As String = ""


            Dim _str_TotalOrdQty As String = ""
            Dim _str_TotalOrdQtyL As String = ""
            Dim v_Sucess As Boolean = True
            Dim oFile As FileStream
            Dim _dbl_TotQty As Double = 0.0
            Dim _dbl_GTotal As Double = 0.0
            Dim l As Integer
            Dim _str_FileName As String = ""
            Dim sPath As String = ""
            Dim _Flag As Boolean = False

            Dim oEditText As SAPbouiCOM.EditText = objForm.Items.Item("eFileName").Specific

            Dim oApp As New Excel.Application
            Dim oWBa As Excel.Workbook
            Dim oWS As Excel.Worksheet

            If oEditText.Value <> "" Then
                sPath = oEditText.Value
                Dim sIndex As String() = sPath.Split(New Char() {"\"c})
                Dim _int_Length As Integer = sIndex.Length - 1
                _str_FileName = sIndex(_int_Length)
            Else
                oApplication.StatusBar.SetText(" Please Browse an Excel File for Uploading ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            'If _str_FileName = "BOM_EXCEL.xlsx" Then
            Try
                oFile = New FileStream(sPath, FileMode.Open)
                'oFile = New FileStream("D:\BOM_EXCEL\BOM_customisation.xlsx", FileMode.Open)
            Catch ex As Exception
                oApplication.MessageBox("Please Verify the Path Of the File / Close the File Before Attempting to Upload...")
                Exit Sub
            End Try
            oFile.Close()
            oWBa = oApp.Workbooks.Open(sPath)
            oWS = DirectCast(oWBa.Worksheets(1), Excel.Worksheet)

            _str_User = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name

            '2012-06-13
            If oWS.Cells(1, 1).value <> Nothing Then
                If oWS.Cells(1, 1).value.ToString() = "START" Then
                    If oWS.Cells(3, 2).value <> Nothing Then
                        Try
                            If oWS.Cells(4, 2).value <> Nothing Then
                                _str_SaleOrdNo = oWS.Cells(4, 2).value.ToString()
                            End If
                            If oWS.Cells(3, 2).value <> Nothing Then
                                _str_StyleCode = oWS.Cells(3, 2).value.ToString()
                            End If
                            oRs.DoQuery("Select DocDate,NumAtCard,CardCode,DocEntry from ORDR where DocNum ='" + _str_SaleOrdNo + "'")
                            _str_RefNo = oRs.Fields.Item("NumAtCard").Value.ToString()
                            _str_DocEntry = oRs.Fields.Item("DocEntry").Value.ToString()
                            If oWS.Cells(8, 2).value <> Nothing Then
                                _str_Unit = oWS.Cells(8, 2).value.ToString()
                            Else
                                _str_Unit = ""
                            End If

                            For i As Integer = 6 To 25
                                If oWS.Cells(i, 6).value <> Nothing Then
                                    _str_ItemCode = oWS.Cells(3, 2).value.ToString() + oWS.Cells(i, 6).value.ToString()
                                    _str_Combo = oWS.Cells(i, 6).value.ToString()

                                    For j As Integer = 29 To 48
                                        If oWS.Cells(j, 6).value <> Nothing Then
                                            If oWS.Cells(i, 6).value.ToString() = oWS.Cells(j, 6).value.ToString() Then
                                                If oWS.Cells(j, 28).value <> Nothing Then
                                                    _str_TotalOrdQty = oWS.Cells(j, 28).value.ToString()
                                                Else
                                                    _str_TotalOrdQty = "0"
                                                End If
                                            End If
                                        End If
                                    Next

                                    '---------------------'
                                    'oRset.DoQuery("Select COUNT(ItemCode)Count from OITM where ItemCode ='" + _str_ItemCode + "'")
                                    oRset.DoQuery("Select COUNT(ItemCode) Count  from RDR1 where DocEntry ='" + _str_DocEntry + "' and ItemCode ='" + _str_ItemCode + "'")
                                    If oRset.Fields.Item("Count").Value.ToString > 0 Then
                                        oRs.DoQuery(" Exec Upload_BOM '" & _str_SaleOrdNo & "','" & _str_StyleCode & "','" & _str_RefNo & "','" & _str_ItemCode & "','" & _str_Combo & "','" & _str_TotalOrdQty & "','" & _str_Unit & "','" & _str_User & "'")
                                    Else
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oWS)
                                        oWBa.Close(True)
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oWBa)
                                        oApp.Workbooks.Close()
                                        oApp.Quit()
                                        oApp.DisplayAlerts = True
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oApp)
                                        oApplication.MessageBox("The Item Code -> " + _str_ItemCode + " does not Contains in the Sale Order No. -> " + _str_SaleOrdNo + " , Please Check the Excel File and Upload Again..")
                                        oRset.DoQuery("Delete from Temp_BOM where SaleOrdNo ='" + _str_SaleOrdNo + "' and StyleCode ='" + _str_StyleCode + "' and Userid ='" + _str_User + "' and Unit ='" + _str_Unit + "'")
                                        Exit Sub
                                    End If
                                End If
                            Next

                            oRs.DoQuery("Select SaleOrdNo,ItemCode ,TotOrdQty ,Combo,StyleCode  from Temp_BOM where StatusH='Open' and Userid = '" + _str_User + "' and StyleCode ='" + _str_StyleCode + "' and SaleOrdNo ='" + _str_SaleOrdNo + "' ")

                            For k As Integer = 1 To oRs.RecordCount
                                _str_ComboHeader = oRs.Fields.Item("Combo").Value.ToString()
                                _str_TotalOrdQty = oRs.Fields.Item("TotOrdQty").Value.ToString()
                                _str_SaleOrderNoL = oRs.Fields.Item("SaleOrdNo").Value.ToString()
                                _str_StyleCodeL = oRs.Fields.Item("StyleCode").Value.ToString()

                                l = 51

                                While oWS.Cells(l, 1).value.ToString() <> "END"

                                    If oWS.Cells(l, 1).value <> Nothing Then
                                        If (oWS.Cells(l, 1).value.ToString() = _str_ComboHeader Or oWS.Cells(l, 1).value.ToString() = "ALL") Then

                                            If oWS.Cells(l, 1).value <> Nothing Then
                                                _str_Combo = _str_ComboHeader
                                            End If
                                            If oWS.Cells(l, 4).value <> Nothing Then
                                                _str_Color = oWS.Cells(l, 4).value.ToString()
                                            End If
                                            If oWS.Cells(l, 5).value <> Nothing Then
                                                _str_CnsQty = oWS.Cells(l, 5).value.ToString()
                                            End If


                                            If (oWS.Cells(l, 6).value.ToString()) = "NO" Then


                                                If oWS.Cells(l, 6).value <> Nothing Then
                                                    _str_Size = oWS.Cells(l, 6).value.ToString()
                                                End If
                                                If oWS.Cells(l, 2).value <> Nothing Then
                                                    _str_ItemCodeL = oWS.Cells(l, 2).value.ToString()
                                                End If
                                                If oWS.Cells(l, 3).value <> Nothing Then
                                                    _str_Description = oWS.Cells(l, 3).value.ToString()
                                                End If
                                                If oWS.Cells(l, 27).value <> Nothing Then
                                                    _str_TotalOrdQtyL = oWS.Cells(l, 27).value.ToString()
                                                Else
                                                    _str_TotalOrdQtyL = "0"
                                                End If
                                                If oWS.Cells(l, 28).value <> Nothing Then
                                                    _str_Wastage = oWS.Cells(l, 28).value.ToString()
                                                Else
                                                    _str_Wastage = "0"
                                                End If
                                                If oWS.Cells(l, 29).value <> Nothing Then
                                                    _str_GTotal = oWS.Cells(l, 29).value.ToString()
                                                Else
                                                    _str_GTotal = "0"
                                                End If
                                                If oWS.Cells(l, 30).value <> Nothing Then
                                                    _str_Process = oWS.Cells(l, 30).value.ToString()
                                                Else
                                                    _str_Process = ""
                                                End If
                                                If oWS.Cells(l, 31).value <> Nothing Then
                                                    _str_VendorCode = oWS.Cells(l, 31).value.ToString()
                                                    oRecordSet.DoQuery("Select CardName From OCRD Where CardCode = '" + _str_VendorCode + "'")
                                                    If oRecordSet.RecordCount > 0 Then
                                                        _str_VendorName = oRecordSet.Fields.Item(0).Value.ToString()
                                                    Else
                                                        _str_VendorCode = ""
                                                        _str_VendorName = ""
                                                    End If
                                                Else
                                                    _str_VendorCode = ""
                                                    _str_VendorName = ""
                                                End If
                                                If oWS.Cells(l, 32).value <> Nothing Then
                                                    _str_Placement = oWS.Cells(l, 32).value.ToString()
                                                Else
                                                    _str_Placement = ""
                                                End If
                                                If oWS.Cells(l, 33).value <> Nothing Then
                                                    _str_FRemarks = oWS.Cells(l, 33).value.ToString()
                                                Else
                                                    _str_FRemarks = ""
                                                End If
                                                If oWS.Cells(l, 34).value <> Nothing Then
                                                    _str_Remarks = oWS.Cells(l, 34).value.ToString()
                                                Else
                                                    _str_Remarks = ""
                                                End If

                                                oRset.DoQuery("Select COUNT(ItemCode)Count from OITM where ItemCode ='" + _str_ItemCodeL + "'")
                                                If oRset.Fields.Item("Count").Value.ToString > 0 Then
                                                    '
                                                    oRset1.DoQuery("Select count(Name)Count from [@GEN_PROCESS_MST] where Name ='" + _str_Process + "'")
                                                    If oRset1.Fields.Item("Count").Value.ToString() > 0 Then
                                                        oRecordSet.DoQuery(" Exec Upload_BOM_Line '" & _str_Combo & "','" & _str_ItemCodeL & "','" & _str_SaleOrderNoL & "','" & _str_StyleCodeL & "','" & _str_Color & "','" & _str_CnsQty & "','" & _str_Size & "','" & _str_TotalOrdQtyL & "', '" & _str_Wastage & "','" & _str_GTotal & "','" & _str_Process & "','" & _str_VendorCode & "','" & _str_VendorName & "', ' " & _str_Placement & "','" & _str_FRemarks & "','" & _str_Remarks & "','" & _str_User & "'")
                                                    Else
                                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oWS)
                                                        oWBa.Close(True)
                                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oWBa)
                                                        oApp.Workbooks.Close()
                                                        oApp.Quit()
                                                        oApp.DisplayAlerts = True
                                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oApp)
                                                        oApplication.MessageBox("Invalid Process Name -> " + _str_Process + " is Invalid for ItemCode - > " + _str_ItemCodeL + ", Please Check the Excel File and Upload Again..")
                                                        oRset.DoQuery("Delete from Temp_BOM where SaleOrdNo ='" + _str_SaleOrdNo + "' and StyleCode ='" + _str_StyleCode + "' and Userid ='" + _str_User + "' and Unit ='" + _str_Unit + "'")
                                                        oRset.DoQuery("Delete from Temp_BOM_Line where SaleOrdNoL = '" + _str_SaleOrderNoL + "' and StyleCodeL ='" + _str_StyleCodeL + "' and Userid ='" + _str_User + "'")
                                                        Exit Sub
                                                    End If
                                                Else
                                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oWS)
                                                    oWBa.Close(True)
                                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oWBa)
                                                    oApp.Workbooks.Close()
                                                    oApp.Quit()
                                                    oApp.DisplayAlerts = True
                                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oApp)
                                                    oApplication.MessageBox("The Item Code -> " + _str_ItemCodeL + " is Invalid, Please Check the Excel File and Upload Again..")
                                                    oRset.DoQuery("Delete from Temp_BOM where SaleOrdNo ='" + _str_SaleOrdNo + "' and StyleCode ='" + _str_StyleCode + "' and Userid ='" + _str_User + "' and Unit ='" + _str_Unit + "'")
                                                    oRset.DoQuery("Delete from Temp_BOM_Line where SaleOrdNoL = '" + _str_SaleOrderNoL + "' and StyleCodeL ='" + _str_StyleCodeL + "' and Userid ='" + _str_User + "'")
                                                    Exit Sub
                                                End If

                                            ElseIf (oWS.Cells(l, 6).value.ToString()) = "YES" Then

                                                For m As Integer = 7 To 26

                                                    If oWS.Cells(50, m).value <> Nothing Then
                                                        _str_Size = oWS.Cells(50, m).value.ToString()
                                                    End If
                                                    If oWS.Cells(l, m).value <> Nothing Then
                                                        _str_CnsQty = oWS.Cells(l, m).value.ToString()
                                                    End If
                                                    If oWS.Cells(l + 1, m).value <> Nothing Then
                                                        _str_ItemCodeL = oWS.Cells(l + 1, m).value.ToString()
                                                    End If
                                                    If oWS.Cells(l, 3).value <> Nothing Then
                                                        _str_Description = oWS.Cells(l, 3).value.ToString()
                                                    End If
                                                    'If oWS.Cells(l + 3, 27).value <> Nothing Then
                                                    If oWS.Cells(l + 3, m).value <> Nothing Then
                                                        _str_TotalOrdQtyL = oWS.Cells(l + 3, m).value.ToString()
                                                    Else
                                                        _str_TotalOrdQtyL = "0"
                                                    End If
                                                    If oWS.Cells(l + 3, 28).value <> Nothing Then
                                                        _str_Wastage = oWS.Cells(l + 3, 28).value.ToString()
                                                    Else
                                                        _str_Wastage = "0"
                                                    End If
                                                    If oWS.Cells(l + 3, 29).value <> Nothing Then
                                                        _str_GTotal = oWS.Cells(l + 3, 29).value.ToString()
                                                    Else
                                                        _str_GTotal = "0"
                                                    End If
                                                    If oWS.Cells(l + 3, 30).value <> Nothing Then
                                                        _str_Process = oWS.Cells(l + 3, 30).value.ToString()
                                                    Else
                                                        _str_Process = ""
                                                    End If
                                                    If oWS.Cells(l + 3, 31).value <> Nothing Then
                                                        _str_VendorCode = oWS.Cells(l + 3, 31).value.ToString()
                                                        oRecordSet.DoQuery("Select CardName From OCRD Where CardCode = '" + _str_VendorCode + "'")
                                                        If oRecordSet.RecordCount > 0 Then
                                                            _str_VendorName = oRecordSet.Fields.Item(0).Value.ToString()
                                                        Else
                                                            _str_VendorCode = ""
                                                            _str_VendorName = ""
                                                        End If
                                                    Else
                                                        _str_VendorCode = ""
                                                        _str_VendorName = ""
                                                    End If
                                                    If oWS.Cells(l + 3, 32).value <> Nothing Then
                                                        _str_Placement = oWS.Cells(l + 3, 32).value.ToString()
                                                    Else
                                                        _str_Placement = ""
                                                    End If
                                                    If oWS.Cells(l + 3, 33).value <> Nothing Then
                                                        _str_FRemarks = oWS.Cells(l + 3, 33).value.ToString()
                                                    Else
                                                        _str_FRemarks = ""
                                                    End If
                                                    If oWS.Cells(l + 3, 34).value <> Nothing Then
                                                        _str_Remarks = oWS.Cells(l + 3, 34).value.ToString()
                                                    Else
                                                        _str_Remarks = ""
                                                    End If

                                                    If oWS.Cells(l + 1, m).value <> Nothing Then
                                                        oRset.DoQuery("Select COUNT(ItemCode)Count from OITM where ItemCode ='" + _str_ItemCodeL + "'")
                                                        If oRset.Fields.Item("Count").Value.ToString > 0 Then
                                                            '
                                                            oRset1.DoQuery("Select count(Name)Count from [@GEN_PROCESS_MST] where Name ='" + _str_Process + "'")
                                                            If oRset1.Fields.Item("Count").Value.ToString() > 0 Then
                                                                oRecordSet.DoQuery(" Exec Upload_BOM_Line '" & _str_Combo & "','" & _str_ItemCodeL & "','" & _str_SaleOrderNoL & "','" & _str_StyleCodeL & "','" & _str_Color & "','" & _str_CnsQty & "','" & _str_Size & "','" & _str_TotalOrdQtyL & "', '" & _str_Wastage & "','" & _str_GTotal & "','" & _str_Process & "','" & _str_VendorCode & "','" & _str_VendorName & "', ' " & _str_Placement & "','" & _str_FRemarks & "','" & _str_Remarks & "','" & _str_User & "'")
                                                            Else
                                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oWS)
                                                                oWBa.Close(True)
                                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oWBa)
                                                                oApp.Workbooks.Close()
                                                                oApp.Quit()
                                                                oApp.DisplayAlerts = True
                                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oApp)
                                                                oApplication.MessageBox("Invalid Process Name -> " + _str_Process + " is Invalid for ItemCode - > " + _str_ItemCodeL + ", Please Check the Excel File and Upload Again..")
                                                                oRset.DoQuery("Delete from Temp_BOM where SaleOrdNo ='" + _str_SaleOrdNo + "' and StyleCode ='" + _str_StyleCode + "' and Userid ='" + _str_User + "' and Unit ='" + _str_Unit + "'")
                                                                oRset.DoQuery("Delete from Temp_BOM_Line where SaleOrdNoL = '" + _str_SaleOrderNoL + "' and StyleCodeL ='" + _str_StyleCodeL + "' and Userid ='" + _str_User + "'")
                                                                Exit Sub
                                                            End If
                                                        Else
                                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oWS)
                                                            oWBa.Close(True)
                                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oWBa)
                                                            oApp.Workbooks.Close()
                                                            oApp.Quit()
                                                            oApp.DisplayAlerts = True
                                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oApp)
                                                            oApplication.MessageBox("The Item Code -> " + _str_ItemCodeL + " is Invalid, Please Check the Excel File and Upload Again..")
                                                            oRset.DoQuery("Delete from Temp_BOM where SaleOrdNo ='" + _str_SaleOrdNo + "' and StyleCode ='" + _str_StyleCode + "' and Userid ='" + _str_User + "' and Unit ='" + _str_Unit + "'")
                                                            oRset.DoQuery("Delete from Temp_BOM_Line where SaleOrdNoL = '" + _str_SaleOrderNoL + "' and StyleCodeL ='" + _str_StyleCodeL + "' and Userid ='" + _str_User + "'")
                                                            Exit Sub
                                                        End If
                                                    End If
                                                Next

                                            End If
                                        End If
                                    End If

                                    l += 1
                                End While

                                oRs.MoveNext()
                            Next
                            If oApp.Visible = False Then
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oWS)
                                oWBa.Close(True)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oWBa)
                                oApp.Workbooks.Close()
                                oApp.Quit()
                                oApp.DisplayAlerts = True
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oApp)

                            End If
                        Catch ex As Exception
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oWS)
                            oWBa.Close(True)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oWBa)
                            oApp.Workbooks.Close()
                            oApp.Quit()
                            oApp.DisplayAlerts = True
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oApp)
                            oApplication.MessageBox("Error Occurs While Uploading An Excel..")
                            _Flag = True
                            oRset.DoQuery("Delete from Temp_BOM where SaleOrdNo ='" + _str_SaleOrdNo + "' and StyleCode ='" + _str_StyleCode + "' and Userid ='" + _str_User + "' and Unit ='" + _str_Unit + "'")
                            oRset.DoQuery("Delete from Temp_BOM_Line where SaleOrdNoL = '" + _str_SaleOrderNoL + "' and StyleCodeL ='" + _str_StyleCodeL + "' and Userid ='" + _str_User + "'")
                        End Try





                        '----------------------- Function To Create Bill OF Materials ----------------------------'

                        If _Flag = False Then
                            Try
                                'objForm = oApplication.Forms.Item(FormUID)
                                Dim _str_ParentName As String = ""
                                Dim _str_ChildName As String = ""
                                Dim _str_Uom As String
                                Dim ifp As IFormatProvider = New System.Globalization.CultureInfo("en-US", True)

                                Dim oRecordSet11 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim oRecordSet12 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRs.DoQuery("Select SaleOrdNo,StyleCode,PurOrdNo [RefNo],ItemCode[Parent],Unit,Combo ,TotOrdQty[Quantity] from Temp_BOM where StatusH ='Open' and Userid='" + _str_User + "' and StyleCode ='" + _str_StyleCode + "' and SaleOrdNo ='" + _str_SaleOrdNo + "'")

                                If oCompany.InTransaction = True Then
                                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                End If
                                oCompany.StartTransaction()

                                If oRs.RecordCount > 0 Then

                                    For i As Integer = 1 To oRs.RecordCount

                                        'oRecordSet11.DoQuery("Select COUNT(DocEntry)Count from [@GEN_CUST_BOM] where U_sono ='" + oRs.Fields.Item("SaleOrdNo").Value.ToString() + "' and U_itemcode ='" + oRs.Fields.Item("Parent").Value.ToString() + "'")
                                        oRecordSet11.DoQuery("Select COUNT(DocEntry)Count from [@GEN_CUST_BOM] where U_sono ='" + oRs.Fields.Item("SaleOrdNo").Value.ToString() + "' and U_itemcode ='" + oRs.Fields.Item("Parent").Value.ToString() + "' and ISNULL(U_closed,'N') <>'Y'")

                                        If oRecordSet11.Fields.Item("Count").Value.ToString() = "0" Then

                                            Dim oGeneralService As SAPbobsCOM.GeneralService
                                            Dim oGeneralData As SAPbobsCOM.GeneralData
                                            Dim oSons As SAPbobsCOM.GeneralDataCollection
                                            Dim oSon As SAPbobsCOM.GeneralData

                                            Dim sCmp As SAPbobsCOM.CompanyService
                                            sCmp = oCompany.GetCompanyService

                                            'Get a handle to the SM_MOR UDO
                                            oGeneralService = sCmp.GetGeneralService("GEN_CUST_BOM")

                                            'Specify data for main UDO
                                            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                                            'Specify data for child UDO
                                            oSons = oGeneralData.Child("GEN_CUST_BOM_D0")

                                            oGeneralData.SetProperty("U_sono", oRs.Fields.Item("SaleOrdNo").Value.ToString())
                                            oGeneralData.SetProperty("U_soref", oRs.Fields.Item("RefNo").Value.ToString())
                                            oGeneralData.SetProperty("U_unit", oRs.Fields.Item("Unit").Value.ToString())
                                            oGeneralData.SetProperty("U_itemcode", oRs.Fields.Item("Parent").Value.ToString())
                                            oRset.DoQuery("Select ItemName from OITM Where ItemCode='" + oRs.Fields.Item("Parent").Value.ToString() + "'")
                                            _str_ParentName = oRset.Fields.Item("ItemName").Value.ToString()
                                            oGeneralData.SetProperty("U_itemname", _str_ParentName)
                                            oRset.DoQuery("Select CONVERT( varchar(25), DocDate, 112)DocDate from ORDR Where DocNum = '" + oRs.Fields.Item("SaleOrdNo").Value.ToString() + "'")
                                            oGeneralData.SetProperty("U_docdate", DateTime.ParseExact(oRset.Fields.Item("DocDate").Value.ToString(), "yyyyMMdd", ifp))
                                            oRset.DoQuery("Select DfltSeries  from ONNM where ObjectCode ='GEN_CUST_BOM'")
                                            oGeneralData.SetProperty("Series", oRset.Fields.Item("DfltSeries").Value.ToString())
                                            oGeneralData.SetProperty("U_status", "NEW")
                                            '2012-06-12
                                            If (oRs.Fields.Item("Quantity").Value.ToString() <> "0") Then
                                                oGeneralData.SetProperty("U_qty", oRs.Fields.Item("Quantity").Value.ToString())
                                            Else
                                                oApplication.MessageBox("Total Quantity Is Zero for the ItemCode-> " + oRs.Fields.Item("Parent").Value.ToString() + ", Please Check the Excel and Upload Again...")
                                                If oCompany.InTransaction = True Then
                                                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                End If
                                                oRecordSet12.DoQuery("Delete Temp_BOM where Userid = '" + _str_User + "' and SaleOrdNo = '" + oRs.Fields.Item("SaleOrdNo").Value.ToString() + "'and StyleCode = '" + oRs.Fields.Item("StyleCode").Value.ToString() + "'")
                                                oRecordSet12.DoQuery("Delete Temp_BOM_Line where Userid='" + _str_User + "' and SaleOrdNoL = '" + oRs.Fields.Item("SaleOrdNo").Value.ToString() + "' and StyleCodeL = '" + oRs.Fields.Item("StyleCode").Value.ToString() + "' ")
                                                Exit Sub
                                            End If

                                            oRecordSet.DoQuery("Select B.SaleOrdNo,B.ItemCode[Parent],B.Combo,B.PurOrdNo [RefNo],B.TotOrdQty [QtyH],B.Unit, " _
                                                        & "BL.ItemCodeL [Child],BL.SaleOrdNoL,BL.StyleCodeL,BL.Dscrptn,BL.CnsQty,BL.Color,BL.TotOrdQtyL,BL.Wastage,BL.GTotal, " _
                                                        & "BL.VendorCode,BL.VendorName ,BL.Process ,BL.Placement,BL.FRemarks,BL.Remarks,BL.Size from Temp_BOM B " _
                                                        & "inner join Temp_BOM_Line BL on B.Combo = BL.Combo " _
                                                        & "where B.StatusH ='Open' and BL.Combo ='" + oRs.Fields.Item("Combo").Value.ToString() + "' and BL.SaleOrdNoL ='" + oRs.Fields.Item("SaleOrdNo").Value.ToString() + "' " _
                                                        & " and BL.StyleCodeL = '" + oRs.Fields.Item("StyleCode").Value.ToString() + "'")
                                            '2012-06-13
                                            If oRecordSet.RecordCount > 0 Then

                                                For j As Integer = 1 To oRecordSet.RecordCount

                                                    oSon = oSons.Add()
                                                    oSon.SetProperty("U_itemcode", oRecordSet.Fields.Item("Child").Value.ToString())
                                                    oRset.DoQuery("Select ItemName from OITM Where ItemCode='" + oRecordSet.Fields.Item("Child").Value.ToString() + "'")
                                                    _str_ChildName = oRset.Fields.Item("ItemName").Value.ToString()
                                                    oSon.SetProperty("U_itemname", _str_ChildName)
                                                    oSon.SetProperty("U_qty", oRecordSet.Fields.Item("CnsQty").Value.ToString())
                                                    oSon.SetProperty("U_size", oRecordSet.Fields.Item("Size").Value.ToString())
                                                    oSon.SetProperty("U_ordrqty", oRecordSet.Fields.Item("TotOrdQtyL").Value.ToString())
                                                    If oRecordSet.Fields.Item("Wastage").Value.ToString() <> "" Then
                                                        oSon.SetProperty("U_per", oRecordSet.Fields.Item("Wastage").Value.ToString())
                                                    End If
                                                    oSon.SetProperty("U_totqty", oRecordSet.Fields.Item("GTotal").Value.ToString())
                                                    oRset.DoQuery("Select InvntryUom from OITM where ItemCode = '" + oRecordSet.Fields.Item("Child").Value.ToString() + "'")
                                                    _str_Uom = oRset.Fields.Item("InvntryUom").Value.ToString()
                                                    oSon.SetProperty("U_uom", _str_Uom)
                                                    If oRecordSet.Fields.Item("VendorCode").Value.ToString() <> "" Then
                                                        oSon.SetProperty("U_cardcode", oRecordSet.Fields.Item("VendorCode").Value.ToString())
                                                    End If
                                                    If oRecordSet.Fields.Item("VendorName").Value.ToString() <> "" Then
                                                        oSon.SetProperty("U_cardname", oRecordSet.Fields.Item("VendorName").Value.ToString())
                                                    End If
                                                    oSon.SetProperty("U_issmthd", "M")
                                                    oSon.SetProperty("U_status", "NEW")
                                                    oSon.SetProperty("U_remarks", oRecordSet.Fields.Item("Remarks").Value.ToString())
                                                    oSon.SetProperty("U_place", oRecordSet.Fields.Item("Placement").Value.ToString())
                                                    oSon.SetProperty("U_process", oRecordSet.Fields.Item("Process").Value.ToString())
                                                    '2012-06-17
                                                    oSon.SetProperty("U_fremark", oRecordSet.Fields.Item("FRemarks").Value.ToString())

                                                    oRecordSet.MoveNext()
                                                Next
                                                oGeneralService.Add(oGeneralData)

                                                oRset.DoQuery("Delete Temp_BOM where Userid = '" + _str_User + "' and SaleOrdNo = '" + oRs.Fields.Item("SaleOrdNo").Value.ToString() + "' and ItemCode ='" + oRs.Fields.Item("Parent").Value.ToString() + "' and Combo ='" + oRs.Fields.Item("Combo").Value.ToString() + "' and StyleCode = '" + oRs.Fields.Item("StyleCode").Value.ToString() + "'")
                                                oRset.DoQuery("Delete Temp_BOM_Line where Userid='" + _str_User + "' and Combo ='" + oRs.Fields.Item("Combo").Value.ToString() + "' and SaleOrdNoL = '" + oRs.Fields.Item("SaleOrdNo").Value.ToString() + "' and StyleCodeL = '" + oRs.Fields.Item("StyleCode").Value.ToString() + "' ")
                                            Else
                                                oApplication.MessageBox("Components Is Missing for Combo -> " + oRs.Fields.Item("Combo").Value.ToString() + ", Please Check the Excel and Upload Again...")
                                                If oCompany.InTransaction = True Then
                                                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                End If
                                                oRecordSet12.DoQuery("Delete Temp_BOM where Userid = '" + _str_User + "' and SaleOrdNo = '" + oRs.Fields.Item("SaleOrdNo").Value.ToString() + "'and StyleCode = '" + oRs.Fields.Item("StyleCode").Value.ToString() + "'")
                                                oRecordSet12.DoQuery("Delete Temp_BOM_Line where Userid='" + _str_User + "' and SaleOrdNoL = '" + oRs.Fields.Item("SaleOrdNo").Value.ToString() + "' and StyleCodeL = '" + oRs.Fields.Item("StyleCode").Value.ToString() + "' ")
                                                Exit Sub
                                            End If
                                        Else
                                            oApplication.MessageBox("Custom BOM for Style Code -> " + oRs.Fields.Item("Parent").Value.ToString() + " and Sale Order No -> " + oRs.Fields.Item("SaleOrdNo").Value.ToString() + " is Already Exist,Please Check the Excel and Upload Again...")
                                            If oCompany.InTransaction = True Then
                                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                            End If
                                            oRecordSet12.DoQuery("Delete Temp_BOM where Userid = '" + _str_User + "' and SaleOrdNo = '" + oRs.Fields.Item("SaleOrdNo").Value.ToString() + "'and StyleCode = '" + oRs.Fields.Item("StyleCode").Value.ToString() + "'")
                                            oRecordSet12.DoQuery("Delete Temp_BOM_Line where Userid='" + _str_User + "' and SaleOrdNoL = '" + oRs.Fields.Item("SaleOrdNo").Value.ToString() + "' and StyleCodeL = '" + oRs.Fields.Item("StyleCode").Value.ToString() + "' ")
                                            Exit Sub
                                        End If
                                        oRs.MoveNext()
                                    Next
                                    If oCompany.InTransaction = True Then
                                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                    End If

                                    oApplication.StatusBar.SetText("Custom BOM Created Successfully", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success) 'oApplication.St("Upload Completed...")
                                End If

                            Catch ex As Exception
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oWS)
                                oWBa.Close(True)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oWBa)
                                oApp.Workbooks.Close()
                                oApp.Quit()
                                oApp.DisplayAlerts = True
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oApp)
                                If oCompany.InTransaction Then
                                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                End If
                            End Try
                        End If
                        '---------------------------------------------------------------------------------------'
                    Else
                        oApplication.StatusBar.SetText("Style Code Not Found....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oWS)
                        oWBa.Close(True)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oWBa)
                        oApp.Workbooks.Close()
                        oApp.Quit()
                        oApp.DisplayAlerts = True
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oApp)
                        Exit Sub
                    End If
                Else
                    oApplication.StatusBar.SetText("Format Not Found....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oWS)
                    oWBa.Close(True)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oWBa)
                    oApp.Workbooks.Close()
                    oApp.Quit()
                    oApp.DisplayAlerts = True
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oApp)
                    Exit Sub
                End If
            Else
                oApplication.StatusBar.SetText("Format Not Found....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oWS)
                oWBa.Close(True)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oWBa)
                oApp.Workbooks.Close()
                oApp.Quit()
                oApp.DisplayAlerts = True
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oApp)
                Exit Sub
            End If
        Catch ex As Exception
        End Try
    End Sub

    'Public Sub Add_BOM(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent)
    '    Try
    '        objForm = oApplication.Forms.Item(FormUID)
    '        Dim _str_ParentName As String = ""
    '        Dim _str_Uom As String
    '        Dim ifp As IFormatProvider = New System.Globalization.CultureInfo("en-US", True)

    '        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


    '        oRs.DoQuery("Select SaleOrdNo,PurOrdNo [RefNo],ItemCode[Parent],Unit,Combo ,TotOrdQty[Quantity] from Temp_BOM where StatusH ='Open' and Userid='" + oCompany.UserName.ToString() + "'")

    '        If oRs.RecordCount > 0 Then

    '            For i As Integer = 1 To oRs.RecordCount

    '                Dim oGeneralService As SAPbobsCOM.GeneralService
    '                Dim oGeneralData As SAPbobsCOM.GeneralData
    '                Dim oSons As SAPbobsCOM.GeneralDataCollection
    '                Dim oSon As SAPbobsCOM.GeneralData

    '                Dim sCmp As SAPbobsCOM.CompanyService
    '                sCmp = oCompany.GetCompanyService

    '                'Get a handle to the SM_MOR UDO
    '                oGeneralService = sCmp.GetGeneralService("GEN_CUST_BOM")

    '                'Specify data for main UDO
    '                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

    '                'Specify data for child UDO
    '                oSons = oGeneralData.Child("GEN_CUST_BOM_D0")

    '                oGeneralData.SetProperty("U_sono", oRs.Fields.Item("SaleOrdNo").Value.ToString())
    '                oGeneralData.SetProperty("U_soref", oRs.Fields.Item("RefNo").Value.ToString())
    '                oGeneralData.SetProperty("U_unit", oRs.Fields.Item("Unit").Value.ToString())
    '                oGeneralData.SetProperty("U_itemcode", oRs.Fields.Item("Parent").Value.ToString())
    '                oRSet.DoQuery("Select ItemName from OITM Where ItemCode='" + oRs.Fields.Item("Parent").Value.ToString() + "'")
    '                _str_ParentName = oRSet.Fields.Item("ItemName").Value.ToString()
    '                oGeneralData.SetProperty("U_itemname", _str_ParentName)
    '                oRSet.DoQuery("Select CONVERT( varchar(25), DocDate, 112)DocDate from ORDR Where DocNum = '" + oRs.Fields.Item("SaleOrdNo").Value.ToString() + "'")
    '                oGeneralData.SetProperty("U_docdate", DateTime.ParseExact(oRSet.Fields.Item("DocDate").Value.ToString(), "yyyyMMdd", ifp))
    '                oGeneralData.SetProperty("Series", "37")
    '                oGeneralData.SetProperty("U_status", "NEW")
    '                oGeneralData.SetProperty("U_qty", oRs.Fields.Item("Quantity").Value.ToString())

    '                oRecordSet.DoQuery("Select B.SaleOrdNo,B.ItemCode[Parent],B.Combo,B.PurOrdNo [RefNo],B.TotOrdQty [QtyH],B.Unit, " _
    '                            & "BL.ItemCodeL [Child],BL.Dscrptn,BL.CnsQty,BL.Color,BL.TotOrdQtyL,BL.Wastage,BL.GTotal, " _
    '                            & "BL.VendorCode,BL.VendorName ,BL.Process ,BL.Placement,BL.FRemarks,BL.Remarks,BL.Size from Temp_BOM B " _
    '                            & "inner join Temp_BOM_Line BL on B.Combo = BL.Combo " _
    '                            & "where B.StatusH ='Open' and BL.Combo ='" + oRs.Fields.Item("Combo").Value.ToString() + "'")


    '                For j As Integer = 1 To oRecordSet.RecordCount

    '                    oSon = oSons.Add()
    '                    oSon.SetProperty("U_itemcode", oRecordSet.Fields.Item("Child").Value.ToString())
    '                    oSon.SetProperty("U_itemname", oRecordSet.Fields.Item("Dscrptn").Value.ToString())
    '                    oSon.SetProperty("U_qty", oRecordSet.Fields.Item("CnsQty").Value.ToString())
    '                    oSon.SetProperty("U_size", oRecordSet.Fields.Item("Size").Value.ToString())
    '                    oSon.SetProperty("U_ordrqty", oRecordSet.Fields.Item("TotOrdQtyL").Value.ToString())
    '                    If oRecordSet.Fields.Item("Wastage").Value.ToString() <> "" Then
    '                        oSon.SetProperty("U_per", oRecordSet.Fields.Item("Wastage").Value.ToString())
    '                    End If
    '                    oSon.SetProperty("U_totqty", oRecordSet.Fields.Item("GTotal").Value.ToString())
    '                    oRSet.DoQuery("Select InvntryUom from OITM where ItemCode = '" + oRecordSet.Fields.Item("Child").Value.ToString() + "'")
    '                    _str_Uom = oRSet.Fields.Item("InvntryUom").Value.ToString()
    '                    oSon.SetProperty("U_uom", _str_Uom)
    '                    If oRecordSet.Fields.Item("VendorCode").Value.ToString() <> "" Then
    '                        oSon.SetProperty("U_cardcode", oRecordSet.Fields.Item("VendorCode").Value.ToString())
    '                    End If
    '                    If oRecordSet.Fields.Item("VendorName").Value.ToString() <> "" Then
    '                        oSon.SetProperty("U_cardname", oRecordSet.Fields.Item("VendorName").Value.ToString())
    '                    End If
    '                    oSon.SetProperty("U_issmthd", "M")
    '                    oSon.SetProperty("U_status", "NEW")
    '                    oSon.SetProperty("U_remarks", oRecordSet.Fields.Item("Remarks").Value.ToString())
    '                    oSon.SetProperty("U_place", oRecordSet.Fields.Item("Placement").Value.ToString())

    '                    oRecordSet.MoveNext()
    '                Next
    '                oGeneralService.Add(oGeneralData)

    '                oRSet.DoQuery("Update Temp_BOM set StatusH ='Closed' where Userid = '" + oCompany.UserName.ToString() + "' and SaleOrdNo = '" + oRs.Fields.Item("SaleOrdNo").Value.ToString() + "' and ItemCode ='" + oRs.Fields.Item("Parent").Value.ToString() + "' and Combo ='" + oRs.Fields.Item("Combo").Value.ToString() + "'")
    '                oRSet.DoQuery("Update Temp_BOM_Line set StatusL ='Closed' where Userid='" + oCompany.UserName.ToString() + "' and Combo ='" + oRs.Fields.Item("Combo").Value.ToString() + "'")

    '                oRs.MoveNext()
    '            Next
    '        Else
    '            oApplication.StatusBar.SetText(" There is No Open BOM's For User -> " + oCompany.UserName.ToString() + "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        End If

    '                Catch ex As Exception
    '    End Try

    'End Sub



#Region "Browse"


    Public Class WindowWrapper

        Implements System.Windows.Forms.IWin32Window
        Private _hwnd As IntPtr

        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

    End Class



    Sub BrowseFileDialog(ByVal objForm As SAPbouiCOM.Form)
        Try
            Dim ShowFolderBrowserThread As Threading.Thread

            ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowser)
            If ShowFolderBrowserThread.ThreadState = ThreadState.Unstarted Then
                ShowFolderBrowserThread.SetApartmentState(ApartmentState.STA)
                ShowFolderBrowserThread.Start()
            ElseIf ShowFolderBrowserThread.ThreadState = ThreadState.Stopped Then
                ShowFolderBrowserThread.Start()
                ShowFolderBrowserThread.Join()
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Public Sub ShowFolderBrowser()

        Dim f As New FileIOPermission(PermissionState.None)
        objForm = oApplication.Forms.GetForm("UPLBOM", oApplication.Forms.ActiveForm.TypeCount)

        Dim BankFileName As String
        f.AllLocalFiles = FileIOPermissionAccess.AllAccess
        Dim MyProcs() As System.Diagnostics.Process
        BankFileName = ""
        Dim OpenFile As New OpenFileDialog

        Try
            OpenFile.Multiselect = False
            OpenFile.Filter = "Excel files (*.xls)|*.xlsx"
            Dim filterindex As Integer = 0
            Try
                filterindex = 0
            Catch ex As Exception
            End Try

            OpenFile.FilterIndex = filterindex

            OpenFile.RestoreDirectory = True
            MyProcs = Process.GetProcessesByName("SAP Business One")

            If MyProcs.Length = 1 Then
                For i As Integer = 0 To MyProcs.Length - 1

                    Dim MyWindow As New WindowWrapper(MyProcs(i).MainWindowHandle)
                    'Dim ret As DialogResult = OpenFile.ShowDialog()
                    Dim ret As DialogResult = OpenFile.ShowDialog(MyWindow)

                    If ret = DialogResult.OK Then
                        Dim FileName As String = OpenFile.FileName
                        objForm.Items.Item("eFileName").Specific.value = FileName
                        OpenFile.Dispose()
                    Else
                        System.Windows.Forms.Application.ExitThread()
                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
            BankFileName = ""
        Finally
            OpenFile.Dispose()
        End Try

    End Sub


#End Region

End Class
