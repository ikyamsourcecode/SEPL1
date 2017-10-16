Imports System.IO
Imports System.Text
Imports System.Drawing.Printing
Imports System.Reflection
Imports System.Data.SqlClient
Imports System.Data.SqlClient.SqlDataReader
Imports System.Globalization
Imports System.Runtime.InteropServices.COMException
Public Class ClsGRPO

#Region "        Declaration        "
    Dim MAC_ID As String
    Dim oUtilities As New ClsUtilities
    Dim objForm, objSubForm, objSForm, objRevForm As SAPbouiCOM.Form
    Dim objItem, objOldItem, TempItem As SAPbouiCOM.Item
    Dim objMatrix, objSubMatrix, objBatchMatrix, objLCMatrix As SAPbouiCOM.Matrix
    Dim SZDBHead As SAPbouiCOM.DBDataSource
    Dim SZDBDetail As SAPbouiCOM.DBDataSource
    Dim SMDBHead As SAPbouiCOM.DBDataSource
    Dim SMDBDetail As SAPbouiCOM.DBDataSource
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
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
    Dim GSONO, GMACID, GITEMCODE As String
    Dim RowID As Integer
    Dim DeleteItemCode, DOCNUM, GRPODOCNUM, GRPODOCENTRY As String
    Dim Quantity As Decimal = 0
    Dim UnitPrice As Decimal = 0
    Dim objBtnCmb As SAPbouiCOM.ButtonCombo
    Dim GRPONO_Cancel As String
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)        
            objOldItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("insstat", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Specific.databind.setbound(True, "OPDN", "u_insstat")
            objItem.LinkTo = "46"
            objItem.Specific.TabOrder = objOldItem.Specific.TabOrder + 1
            objItem.DisplayDesc = True
            objItem.Visible = False
            objOldItem = objForm.Items.Item("2")
            objItem = objForm.Items.Add("btnit", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Left = objOldItem.Left + objOldItem.Width + 5
            objItem.Width = objOldItem.Width + 25
            objItem.Height = objOldItem.Height
            objItem.Top = objOldItem.Top
            objItem.Specific.caption = "Move to Main Wh"
            objItem.LinkTo = "2"
            objItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objOldItem = objForm.Items.Item("2")
            objItem = objForm.Items.Add("btnlc", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Left = objOldItem.Left + objOldItem.Width + 100
            objItem.Width = objOldItem.Width + 25
            objItem.Height = objOldItem.Height
            objItem.Top = objOldItem.Top
            objItem.Specific.caption = "LC"
            objItem.LinkTo = "2"

            objOldItem = objForm.Items.Item("2")
            objItem = objForm.Items.Add("btnmr", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Left = objOldItem.Left + objOldItem.Width + 200
            objItem.Width = objOldItem.Width + 25
            objItem.Height = objOldItem.Height
            objItem.Top = objOldItem.Top
            objItem.Specific.caption = "Manual-Revaluation"
            objItem.LinkTo = "2"
            objItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objOldItem = objForm.Items.Item("10000329")
            objItem = objForm.Items.Add("btncmb", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO)
            objItem.Left = objOldItem.Left
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Top = objOldItem.Top
            objItem.Specific.caption = "Copy To"
            objBtnCmb = objItem.Specific
            objBtnCmb.ValidValues.Add("1", "A/P Invoice")
            objBtnCmb.ValidValues.Add("2", "Goods Return")
            objItem.LinkTo = "10000329"
            objOldItem.Visible = False
            objItem.AffectsFormMode = False
            objForm = oApplication.Forms.Item(FormUID)
            TempItem = objForm.Items.Item("86")
            objItem = objForm.Items.Add("spc", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = TempItem.Top + TempItem.Height + 20
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.Caption = "Unit"
            objItem.LinkTo = "86"
            TempItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("cpc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = TempItem.Top + TempItem.Height + 20
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.databind.setbound(True, "OPDN", "u_unit")
            objItem.Specific.TabOrder = TempItem.Specific.TabOrder + 1
            objItem.LinkTo = "46"

            objForm = oApplication.Forms.Item(FormUID)
            TempItem = objForm.Items.Item("86")
            objItem = objForm.Items.Add("vendcode", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Top = TempItem.Top + TempItem.Height + 20
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.Caption = "VendorCode"
            objItem.LinkTo = "86"
            objItem.Visible = False
            TempItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("cardcode", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = TempItem.Top + TempItem.Height + 20
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.databind.setbound(True, "OPDN", "u_cardcode")
            objItem.Specific.TabOrder = TempItem.Specific.TabOrder + 1
            objItem.Visible = False
            objItem.LinkTo = "46"

            TempItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("macid", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Top = TempItem.Top + TempItem.Height + 40
            objItem.Left = TempItem.Left
            objItem.Width = TempItem.Width
            objItem.Height = TempItem.Height
            objItem.Specific.databind.setbound(True, "OPDN", "u_macid")
            objItem.Specific.TabOrder = TempItem.Specific.TabOrder + 1
            objItem.Visible = False
            objItem.LinkTo = "46"
            MAC_ID = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
            objForm.Items.Item("macid").Specific.value = MAC_ID
            
            objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("cpc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'Dim reval As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'Dim reval1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'reval.DoQuery("Select Docnum,Ref2 from omrv where Docdate>='20130401'")
            'For i As Integer = 1 To reval.RecordCount
            '    reval1.DoQuery("Update OPDN Set u_revalno='" + reval.Fields.Item(1).Value + "'")
            'Next
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub Open_LC_Form(ByVal FormUID As String) ', ByVal InvoiceNo As String, ByVal MACID As String, ByVal Mode As String, ByVal BASENO As String)
        Try
            PARENT_FORM = FormUID
            Dim RS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim RS2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim CHILD_FORM As String = "GEN_GRPO_LCOSTS@" & FormUID
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
                oUtilities.SAPXML("GEN_GRPO_LCosts.xml", CHILD_FORM)
                objSubForm = oApplication.Forms.Item(CHILD_FORM)
                objSubForm.Select()
            End If
            ChildModalForm = True
            Dim objLCMatrix As SAPbouiCOM.Matrix
            objLCMatrix = objSubForm.Items.Item("mtx").Specific
            objSubForm.DataSources.DataTables.Add("MyDataTable")
      
            Me.LoadItems(FormUID, DOCNUM)
            MAC_ID = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT top 1 U_lname,U_glacct,U_rate,U_qty,U_value,T0.code FROM [@GEN_GRPO_LCOSTS] T0  INNER JOIN [@GEN_GRPO_LCOSTS_D0] T1 ON T0.Code=T1.Code WHERE T0.U_grpono='" + GRPODOCNUM + "' and T0.u_macid='" + MAC_ID + "' order by T0.code desc")
            If oRecordSet.RecordCount = 0 Then
                'Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRSet.DoQuery("Select Count(*) + 1 AS 'Count' From [@GEN_GRPO_LCOSTS]")
                'oDBs_Head = objSubForm.DataSources.DBDataSources.Item("@GEN_GRPO_LCOSTS")
                'oDBs_Head.SetValue("Code", 0, oRSet.Fields.Item("Count").Value)
                objMatrix = objSubForm.Items.Item("mtx").Specific
                objMatrix.AddRow()
                'objMatrix.SetLineData(pVal.Row)
                'ITMatrix.FlushToDataSource()
                Me.SetNewLine(objSubForm.UniqueID, objMatrix.VisualRowCount, objMatrix)
            End If
            'For i As Integer = 0 To ogrid.Columns.Count - 1
            '    ogrid.Columns.Item(i).Editable = False
            'Next
            'ogrid.Columns.Item(ogrid.Columns.Count - 1).Editable = True
        Catch ex As Exception
            ' oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CreateSubForm(ByVal FormUID As String)
        Try
            oUtilities.SAPXML("GEN_GRPO_LCosts.xml")
            objForm = oApplication.Forms.GetForm("GEN_GRPO_LCOSTS", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_GRPO_LCOSTS")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_GRPO_LCOSTS_D0")
            objForm.DataBrowser.BrowseBy = "code"
            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            'objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            'objForm.Items.Item("name").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)


            'Me.SetNewLine(objForm.UniqueID, objMatrix.VisualRowCount, objMatrix)
            Me.LoadItems(FormUID, DOCNUM)
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT top 1 U_lname,U_glacct,U_rate,U_qty,U_value,T0.code FROM [@GEN_GRPO_LCOSTS] T0  INNER JOIN [@GEN_GRPO_LCOSTS_D0] T1 ON T0.Code=T1.Code WHERE T0.U_grpono='" + GRPODOCNUM + "' and T0.U_macid='" + MAC_ID + "' order by T0.code desc")
            If oRecordSet.RecordCount = 0 Then
                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRSet.DoQuery("Select code+ 1 AS 'Count' From [@GEN_GRPO_LCOSTS] where code=(Select top 1 code from [@GEN_GRPO_LCOSTS] order by DocEntry desc)")
                'oRSet.DoQuery("Select Count(*) + 1 AS 'Count' From [@GEN_GRPO_LCOSTS]")
                oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_GRPO_LCOSTS")
                oDBs_Head.SetValue("Code", 0, oRSet.Fields.Item("Count").Value)
                objMatrix = objForm.Items.Item("mtx").Specific
                objMatrix.AddRow()
                'objMatrix.SetLineData(pVal.Row)
                'ITMatrix.FlushToDataSource()
                Me.SetNewLine(objForm.UniqueID, objMatrix.VisualRowCount, objMatrix)

            End If

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadItems(ByVal FormUID As String, ByVal DOCNUM As String)
        Try
            Dim ITForm As SAPbouiCOM.Form
            MAC_ID = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
            Dim ITMatrix As SAPbouiCOM.Matrix
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT top 1 U_lname,U_glacct,U_rate,U_qty,U_value,T0.code FROM [@GEN_GRPO_LCOSTS] T0  INNER JOIN [@GEN_GRPO_LCOSTS_D0] T1 ON T0.Code=T1.Code WHERE T0.U_grpono='" + GRPODOCNUM + "' And T0.U_macid='" + MAC_ID + "' order by T0.code desc")
            ' ITForm = oApplication.Forms.Item(FormUID)
            ITForm = oApplication.Forms.GetForm("GEN_GRPO_LCOSTS", oApplication.Forms.ActiveForm.TypeCount)
            ITMatrix = ITForm.Items.Item("mtx").Specific
            Try
                If oRecordSet.RecordCount > 0 Then
                    ITForm.Items.Item("grpono").Specific.value = GRPODOCNUM
                    ITForm.Items.Item("macid").Visible = True
                    ITForm.Items.Item("macid").Specific.value = MAC_ID
                    'ITForm.Items.Item("grpono").editable = False
                    ITForm.Items.Item("1").Click()
                    ITForm.Items.Item("grpono").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    ITForm.Items.Item("macid").Visible = False
                    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRs.DoQuery("Select docnum from OPDN Where Docnum='" + GRPODOCNUM + "' And u_macid='" + MAC_ID + "' and u_lcadd='Y' and u_reval='YES'")
                    If oRs.RecordCount > 0 Then
                        ITMatrix.Columns.Item("lname").Editable = False
                        ITMatrix.Columns.Item("rate").Editable = False
                        ITMatrix.Columns.Item("glacct").Editable = False
                        ITMatrix.Columns.Item("qty").Editable = False
                        ITMatrix.Columns.Item("value").Editable = False
                        ITMatrix.Columns.Item("macid").Editable = False
                    Else
                        ITMatrix.Columns.Item("lname").Editable = True
                        ITMatrix.Columns.Item("rate").Editable = True
                        ITMatrix.Columns.Item("glacct").Editable = True
                        ITMatrix.Columns.Item("qty").Editable = True
                        ITMatrix.Columns.Item("value").Editable = True
                        ITMatrix.Columns.Item("macid").Editable = True
                    End If

                    'ITMatrix.SetLineData(ITMatrix.VisualRowCount)
                    'ITMatrix.FlushToDataSource()
                    'ITMatrix.AddRow(1, ITMatrix.VisualRowCount)
                    'Me.SetNewLine(ITForm.UniqueID, ITMatrix.VisualRowCount, ITMatrix)

                    'Dim Flag As Boolean = False
                    'Dim errflag As Boolean = False
                    'If objMatrix.VisualRowCount = 1 Or pVal.Row = objMatrix.VisualRowCount Then
                    '    Flag = True
                    'End If
                    'If Flag = True Then
                    '    objMatrix.AddRow(1, objMatrix.VisualRowCount)
                    '    Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
                    'End If
                Else
                    ITForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                    ITForm.Items.Item("grpono").Specific.value = GRPODOCNUM
                    objSubForm.Items.Item("grpono").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    objSubForm.Items.Item("macid").Specific.value = MAC_ID
                    objSubForm.Items.Item("macid").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    objLCMatrix.Columns.Item("lname").Editable = True
                    objLCMatrix.Columns.Item("rate").Editable = True
                    objLCMatrix.Columns.Item("glacct").Editable = True
                    objLCMatrix.Columns.Item("qty").Editable = True
                    objLCMatrix.Columns.Item("value").Editable = True
                    objLCMatrix.Columns.Item("macid").Editable = True
                    'ITForm.Items.Item("grpoqty").Specific.value = Quantity
                    'Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'oRSet.DoQuery("Select Count(*) + 1 AS 'Count' From [@GEN_GRPO_LCOSTS]")
                    'oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_GRPO_LCOSTS")
                    'oDBs_Head.SetValue("Code", 0, oRSet.Fields.Item("Count").Value)
                    'objMatrix = objForm.Items.Item("mtx").Specific
                    'objMatrix.AddRow()
                    ''objMatrix.SetLineData(pVal.Row)
                    ''ITMatrix.FlushToDataSource()
                    'Me.SetNewLine(objForm.UniqueID, objMatrix.VisualRowCount, objMatrix)
                End If
            Catch ex As Exception
                oApplication.StatusBar.SetText(ex.Message)

            End Try
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Sub Open_LC_Form_OK(ByVal FormUID As String) ', ByVal InvoiceNo As String, ByVal MACID As String, ByVal Mode As String, ByVal BASENO As String)
        Try
            PARENT_FORM = FormUID
            Dim RS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim RS2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim CHILD_FORM As String = "GEN_GRPO_LCOSTS@" & FormUID
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
                oUtilities.SAPXML("GEN_GRPO_LCosts.xml", CHILD_FORM)
                objSubForm = oApplication.Forms.Item(CHILD_FORM)
                objSubForm.Select()
            End If
            ChildModalForm = True
            Dim objLCMatrix As SAPbouiCOM.Matrix
            objLCMatrix = objSubForm.Items.Item("mtx").Specific
            objSubForm.DataSources.DataTables.Add("MyDataTable")

            Me.LoadItems_OK(FormUID, DOCNUM)
            MAC_ID = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT top 1 U_lname,U_glacct,U_rate,U_qty,U_value,T0.code FROM [@GEN_GRPO_LCOSTS] T0  INNER JOIN [@GEN_GRPO_LCOSTS_D0] T1 ON T0.Code=T1.Code WHERE T0.U_grpono='" + GRPODOCNUM + "' order by T0.code desc")
            If oRecordSet.RecordCount = 0 Then
                'Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRSet.DoQuery("Select Count(*) + 1 AS 'Count' From [@GEN_GRPO_LCOSTS]")
                'oDBs_Head = objSubForm.DataSources.DBDataSources.Item("@GEN_GRPO_LCOSTS")
                'oDBs_Head.SetValue("Code", 0, oRSet.Fields.Item("Count").Value)
                objMatrix = objSubForm.Items.Item("mtx").Specific
                objMatrix.AddRow()
                'objMatrix.SetLineData(pVal.Row)
                'ITMatrix.FlushToDataSource()
                Me.SetNewLine(objSubForm.UniqueID, objMatrix.VisualRowCount, objMatrix)
            End If
            'For i As Integer = 0 To ogrid.Columns.Count - 1
            '    ogrid.Columns.Item(i).Editable = False
            'Next
            'ogrid.Columns.Item(ogrid.Columns.Count - 1).Editable = True
        Catch ex As Exception
            ' oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CreateSubForm_OK(ByVal FormUID As String)
        Try
            oUtilities.SAPXML("GEN_GRPO_LCosts.xml")
            objForm = oApplication.Forms.GetForm("GEN_GRPO_LCOSTS", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_GRPO_LCOSTS")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_GRPO_LCOSTS_D0")
            objForm.DataBrowser.BrowseBy = "code"
            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            'objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            'objForm.Items.Item("name").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)


            'Me.SetNewLine(objForm.UniqueID, objMatrix.VisualRowCount, objMatrix)
            Me.LoadItems(FormUID, DOCNUM)
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT top 1 U_lname,U_glacct,U_rate,U_qty,U_value,T0.code FROM [@GEN_GRPO_LCOSTS] T0  INNER JOIN [@GEN_GRPO_LCOSTS_D0] T1 ON T0.Code=T1.Code WHERE T0.U_grpono='" + GRPODOCNUM + "' order by T0.code desc")
            If oRecordSet.RecordCount = 0 Then
                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRSet.DoQuery("Select code+ 1 AS 'Count' From [@GEN_GRPO_LCOSTS] where code=(Select top 1 code from [@GEN_GRPO_LCOSTS] order by DocEntry desc)")
                '                oRSet.DoQuery("Select Count(*) + 1 AS 'Count' From [@GEN_GRPO_LCOSTS]")
                oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_GRPO_LCOSTS")
                oDBs_Head.SetValue("Code", 0, oRSet.Fields.Item("Count").Value)
                objMatrix = objForm.Items.Item("mtx").Specific
                objMatrix.AddRow()
                'objMatrix.SetLineData(pVal.Row)
                'ITMatrix.FlushToDataSource()
                Me.SetNewLine(objForm.UniqueID, objMatrix.VisualRowCount, objMatrix)

            End If

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadItems_OK(ByVal FormUID As String, ByVal DOCNUM As String)
        Try
            Dim ITForm As SAPbouiCOM.Form
            MAC_ID = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
            Dim ITMatrix As SAPbouiCOM.Matrix
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT top 1 U_lname,U_glacct,U_rate,U_qty,U_value,T0.code FROM [@GEN_GRPO_LCOSTS] T0  INNER JOIN [@GEN_GRPO_LCOSTS_D0] T1 ON T0.Code=T1.Code WHERE T0.U_grpono='" + GRPODOCNUM + "'  order by T0.code desc")
            ' ITForm = oApplication.Forms.Item(FormUID)
            ITForm = oApplication.Forms.GetForm("GEN_GRPO_LCOSTS", oApplication.Forms.ActiveForm.TypeCount)
            ITMatrix = ITForm.Items.Item("mtx").Specific
            Try
                If oRecordSet.RecordCount > 0 Then
                    ITForm.Items.Item("grpono").Specific.value = GRPODOCNUM
                    'ITForm.Items.Item("macid").Visible = True
                    'ITForm.Items.Item("macid").Specific.value = MAC_ID
                    'ITForm.Items.Item("grpono").editable = False
                    ITForm.Items.Item("1").Click()
                    ITForm.Items.Item("grpono").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRs.DoQuery("Select docnum from OPDN Where Docnum='" + GRPODOCNUM + "' And u_macid='" + MAC_ID + "' and u_lcadd='Y' and u_reval='YES'")
                    ITForm.Items.Item("macid").Visible = False
                    If oRs.RecordCount > 0 Then
                        ITMatrix.Columns.Item("lname").Editable = False
                        ITMatrix.Columns.Item("rate").Editable = False
                        ITMatrix.Columns.Item("glacct").Editable = False
                        ITMatrix.Columns.Item("qty").Editable = False
                        ITMatrix.Columns.Item("value").Editable = False
                        ITMatrix.Columns.Item("macid").Editable = False
                    Else
                        ITMatrix.Columns.Item("lname").Editable = True
                        ITMatrix.Columns.Item("rate").Editable = True
                        ITMatrix.Columns.Item("glacct").Editable = True
                        ITMatrix.Columns.Item("qty").Editable = True
                        ITMatrix.Columns.Item("value").Editable = True
                        ITMatrix.Columns.Item("macid").Editable = True
                    End If

                    'ITMatrix.SetLineData(ITMatrix.VisualRowCount)
                    'ITMatrix.FlushToDataSource()
                    'ITMatrix.AddRow(1, ITMatrix.VisualRowCount)
                    'Me.SetNewLine(ITForm.UniqueID, ITMatrix.VisualRowCount, ITMatrix)

                    'Dim Flag As Boolean = False
                    'Dim errflag As Boolean = False
                    'If objMatrix.VisualRowCount = 1 Or pVal.Row = objMatrix.VisualRowCount Then
                    '    Flag = True
                    'End If
                    'If Flag = True Then
                    '    objMatrix.AddRow(1, objMatrix.VisualRowCount)
                    '    Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
                    'End If
                Else
                    ITForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                    ITForm.Items.Item("grpono").Specific.value = GRPODOCNUM
                    objSubForm.Items.Item("grpono").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    objSubForm.Items.Item("macid").Specific.value = MAC_ID
                    objSubForm.Items.Item("macid").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    objLCMatrix.Columns.Item("lname").Editable = True
                    objLCMatrix.Columns.Item("rate").Editable = True
                    objLCMatrix.Columns.Item("glacct").Editable = True
                    objLCMatrix.Columns.Item("qty").Editable = True
                    objLCMatrix.Columns.Item("value").Editable = True
                    objLCMatrix.Columns.Item("macid").Editable = True
                    'ITForm.Items.Item("grpoqty").Specific.value = Quantity
                    'Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'oRSet.DoQuery("Select Count(*) + 1 AS 'Count' From [@GEN_GRPO_LCOSTS]")
                    'oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_GRPO_LCOSTS")
                    'oDBs_Head.SetValue("Code", 0, oRSet.Fields.Item("Count").Value)
                    'objMatrix = objForm.Items.Item("mtx").Specific
                    'objMatrix.AddRow()
                    ''objMatrix.SetLineData(pVal.Row)
                    ''ITMatrix.FlushToDataSource()
                    'Me.SetNewLine(objForm.UniqueID, objMatrix.VisualRowCount, objMatrix)
                End If
            Catch ex As Exception
                oApplication.StatusBar.SetText(ex.Message)

            End Try
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim ITFORM As SAPbouiCOM.Form
            ITFORM = oApplication.Forms.GetForm("GEN_GRPO_LCOSTS", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Detail = ITFORM.DataSources.DBDataSources.Item("@GEN_GRPO_LCOSTS_D0")
            objMatrix = oMatrix
            objMatrix.FlushToDataSource()
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, objMatrix.VisualRowCount)
            'oDBs_Detail.SetValue("u_lcode", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_lname", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_glacct", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_rate", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_qty", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_value", oDBs_Detail.Offset, "")
            oDBs_Detail.SetValue("u_macid", oDBs_Detail.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try

            Dim CHILD_FORM As String = "GEN_GRPO_LCOSTS@" & pVal.FormUID
            Dim ModalForm As Boolean = False
            For i As Integer = 0 To oApplication.Forms.Count - 1
                If oApplication.Forms.Item(i).UniqueID = (CHILD_FORM) Then
                    objSubForm = oApplication.Forms.Item("GEN_GRPO_LCOSTS@" & pVal.FormUID)
                    ModalForm = True
                    Exit For
                End If
            Next
            If ModalForm = False Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                        objForm = oApplication.Forms.Item(pVal.FormUID)
                        objMatrix = objForm.Items.Item("38").Specific

                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        If pVal.BeforeAction = True Then
                            If pVal.FormTypeCount = 1 Then
                                Me.CreateForm(FormUID)
                            Else
                                BubbleEvent = False
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        If pVal.FormUID = "GEN_GRPO_LCOSTS" Then

                        End If

                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If pVal.ItemUID = "btncmb" And pVal.BeforeAction = False Then
                            Dim oCombo As SAPbouiCOM.ButtonCombo
                            Dim TRGTCombo As SAPbouiCOM.ComboBox
                            oCombo = objForm.Items.Item("btncmb").Specific
                            TRGTCombo = objForm.Items.Item("10000329").Specific
                            If Trim(oCombo.Selected.Value) = "1" Then
                                If Trim(objForm.Items.Item("insstat").Specific.selected.value) <> "Closed" Then
                                    oApplication.StatusBar.SetText("The inspection has to be done before you can create A/P Invoice", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                Else
                                    If objForm.Items.Item("10000329").Enabled = True Then
                                        TRGTCombo.Select("A/P Invoice", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                    End If
                                End If
                            End If
                            If Trim(oCombo.Selected.Value) = "2" Then
                                If objForm.Items.Item("10000329").Enabled = True Then
                                    TRGTCombo.Select("G. Return", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                End If
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                        If pVal.Before_Action = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            GRPONO_Cancel = objForm.Items.Item("8").Specific.value
                        End If
                        If pVal.ActionSuccess = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Dim GRPO_Doc As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim LC_Check As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim LC_Check1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            GRPO_Doc.DoQuery("Select Docnum from OPDN Where Docnum='" & GRPONO_Cancel & "'")
                            LC_Check1.DoQuery("Select u_grpono from [@GEN_GRPO_LCOSTS] where u_grpono='" + GRPONO_Cancel + "' and u_macid='" + MAC_ID + "'")
                            If GRPO_Doc.RecordCount = 0 Then
                                If LC_Check1.RecordCount > 0 Then
                                    LC_Check.DoQuery("Delete From [@GEN_GRPO_LCOSTS_D0] where code=(select code from [@GEN_GRPO_LCOSTS] where u_grpono='" & GRPONO_Cancel & "' and u_macid='" + MAC_ID + "')")
                                    LC_Check.DoQuery("Delete From [@GEN_GRPO_LCOSTS] where u_grpono='" & GRPONO_Cancel & "' and u_macid='" + MAC_ID + "'")
                                End If
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
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = True Then
                            If Trim(objForm.Items.Item("cpc").Specific.value) = "" Then
                                oApplication.StatusBar.SetText("Please select Unit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If

                        If pVal.ItemUID = "2" And pVal.Before_Action = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            GRPONO_Cancel = objForm.Items.Item("8").Specific.value
                        End If
                        If pVal.ItemUID = "2" And pVal.ActionSuccess = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Dim GRPO_Doc As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim LC_Check As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim LC_Check1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            GRPO_Doc.DoQuery("Select Docnum from OPDN Where Docnum='" & GRPONO_Cancel & "'")
                            LC_Check1.DoQuery("Select u_grpono from [@GEN_GRPO_LCOSTS] where u_grpono='" + GRPONO_Cancel + "' and u_macid='" + MAC_ID + "'")
                            If GRPO_Doc.RecordCount = 0 Then
                                If LC_Check1.RecordCount > 0 Then
                                    LC_Check.DoQuery("Delete From [@GEN_GRPO_LCOSTS_D0] where code=(select code from [@GEN_GRPO_LCOSTS] where u_grpono='" & GRPONO_Cancel & "' and u_macid='" + MAC_ID + "')")
                                    LC_Check.DoQuery("Delete From [@GEN_GRPO_LCOSTS] where u_grpono='" & GRPONO_Cancel & "' and u_macid='" + MAC_ID + "'")
                                End If
                            End If
                        End If
                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                            objMatrix = objForm.Items.Item("38").Specific
                            Dim ErrFlag As Boolean = False
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                If Trim(objMatrix.Columns.Item("1").Cells.Item(i).Specific.value) <> "" Then
                                    oRSet.DoQuery("Select OpenQty From POR1 WHere DocEntry = '" + Trim(objMatrix.Columns.Item("45").Cells.Item(i).Specific.value) + "' And LineNum = '" + Trim(objMatrix.Columns.Item("46").Cells.Item(i).Specific.Value) + "'")
                                    If oRSet.RecordCount > 0 Then
                                        If CDbl(objMatrix.Columns.Item("11").Cells.Item(i).Specific.value) < CDbl(oRSet.Fields.Item("OpenQty").Value) Then
                                            ErrFlag = True
                                        End If
                                    End If
                                End If
                            Next
                            If ErrFlag = True Then
                                If oApplication.MessageBox("GRPO quantity is less than PO quantity. Do you still want to continue? ", 2, "Yes", "No") = 2 Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                            If Me.Validation(FormUID) = False Then
                                BubbleEvent = False
                                Exit Sub

                            End If
                            If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                                objForm.Items.Item("macid").Specific.value = MAC_ID
                                If Trim(objForm.Items.Item("cardcode").Specific.value) = "" Then
                                    oApplication.StatusBar.SetText("Please Select VendorCode In UserDefinedFields", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                            End If
                            'End If

                        ElseIf pVal.ItemUID = "1" And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Dim orset1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            orset1.DoQuery("Update OPDN Set u_lcadd='Y' where OPDN.DocEntry=(Select Top 1 docentry from opdn where u_macid='" + MAC_ID + "' order by docentry desc)")
                            Dim macid As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim macid1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            macid.DoQuery("Select docnum,u_macid from OPDN where DocEntry=(Select Top 1 docentry from opdn where u_macid='" + MAC_ID + "' order by docentry desc)")
                            macid1.DoQuery("Update [@GEN_GRPO_LCOSTS] Set u_grpono='" + macid.Fields.Item(0).Value.ToString + "' where code=(Select top 1 code from [@GEN_GRPO_LCOSTS] Where u_macid='" + MAC_ID + "' order by DocEntry desc)")
                            Me.InventoryRevaluation()

                            'orset1.DoQuery("Update OPDN Set u_lcadd='Y' where OPDN.Docnum='" + objForm.Items.Item("grpono").Specific.value + "' and u_macid='" + MAC_ID + "'")

                            'Dim macid As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'Dim macid1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'macid.DoQuery("Select docnum,u_macid from OPDN where DocEntry=(Select Top 1 docentry from opdn where u_macid='" + MAC_ID + "' order by docentry desc)")
                            'macid1.DoQuery("Update [@GEN_GRPO_LCOSTS] Set u_grpono='" + macid.Fields.Item(0).Value.ToString + "' where code=(Select top 1 code from [@GEN_GRPO_LCOSTS] Where u_macid='" + MAC_ID + "' order by code desc)")


                            'objForm.Items.Item("btnlc").Enabled = False

                            ' objSForm = oApplication.Forms.Item("OPCH")
                            'objSForm.Items.Item("U_season").Specific.value = "Done"
                        End If
                        If pVal.ItemUID = "btnmr" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            Dim LCAdd As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            GRPODOCNUM = objForm.Items.Item("8").Specific.Value

                            Dim Docentry As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Docentry.DoQuery("Select Docentry from OPDN Where Docnum='" + GRPODOCNUM + "'")
                            GRPODOCENTRY = Docentry.Fields.Item(0).Value
                            LCAdd.DoQuery("Select u_reval,u_lcadd from OPDN Where Docnum='" + GRPODOCNUM + "'")
                            If LCAdd.Fields.Item(0).Value = "Yes" Then
                                oApplication.StatusBar.SetText("Revaluation Already Done For This Document", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                'GoTo Z
                                BubbleEvent = False
                            ElseIf LCAdd.Fields.Item(1).Value = "N" Then
                                oApplication.StatusBar.SetText("Please Select Landed Cost By Clicking LC Button", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                            End If
                        End If
                        If pVal.ItemUID = "btnmr" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE And pVal.ActionSuccess = True Then
                            Me.InventoryRevaluation_Manual()
                            Dim macid As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim macid1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            macid.DoQuery("Select docnum,u_macid from OPDN where DocEntry=(Select Top 1 docentry from opdn where u_macid='" + MAC_ID + "' order by docentry desc)")
                            '  macid1.DoQuery("Update [@GEN_GRPO_LCOSTS] Set u_grpono='" + macid.Fields.Item(0).Value + "' where code=(Select top 1 code from [@GEN_GRPO_LCOSTS] Where u_macid='" + MAC_ID + "' order by code desc)")
                        End If
                        'Z:
                        If pVal.ItemUID = "btnit" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            Dim LCAdd As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            GRPODOCNUM = objForm.Items.Item("8").Specific.Value
                            LCAdd.DoQuery("Select u_reval,u_lcadd from OPDN Where Docnum='" + GRPODOCNUM + "'")
                            If LCAdd.Fields.Item(1).Value = "N" Then
                                oApplication.StatusBar.SetText("Plaese Select Landed Cost By Clicking on LC Button", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                            ElseIf (LCAdd.Fields.Item(0).Value = "" Or LCAdd.Fields.Item(0).Value = "NO" Or LCAdd.Fields.Item(0).Value = "Partial") Then
                                oApplication.StatusBar.SetText("Plaese Click on Manual Revaluation", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                            Else
                                Dim whs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                For Row As Integer = 1 To objMatrix.VisualRowCount
                                    whs.DoQuery("select itemcode,Quantity,DocNum,opdn.DocDate from OPDN inner join PDN1 on opdn.DocEntry=pdn1.DocEntry inner join ostc on pdn1.taxcode=ostc.code where ItemCode='" + objMatrix.Columns.Item("1").Cells.Item(Row).Specific.value + "' and Whscode='" + objMatrix.Columns.Item("24").Cells.Item(Row).Specific.value + "' and opdn.U_insstat='Open' and ostc.Rate<>'0.00' and (u_reval='NO' or u_LCAdd='N') And OPDN.Docdate>='20130401'  and OPDN.Docnum<>'" + GRPODOCNUM + "'")                                

                                    If whs.RecordCount > 0 Then
                                        For i As Integer = 1 To whs.RecordCount
                                            oApplication.StatusBar.SetText("Please Select LC And Make 'Manual Revaluation' for pending GRN No '" & whs.Fields.Item("DocNum").Value & "'", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                        Next
                                    End If
                                Next
                            End If
                        End If
                        If pVal.ItemUID = "btnlc" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) And pVal.Before_Action = True Then
                            For Price As Integer = 1 To objMatrix.VisualRowCount - 1
                                If objMatrix.Columns.Item("14").Cells.Item(Price).Specific.value = "" Then
                                    oApplication.StatusBar.SetText("Please give UnitPrice for all the items", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Next
                            GRPODOCNUM = objForm.Items.Item("8").Specific.Value
                            objForm.Items.Item("macid").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("macid").Specific.value = MAC_ID
                            objForm.Items.Item("macid").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            Dim A As Integer = 0
                            Quantity = 0
                            UnitPrice = 0
                            For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                                Quantity = Quantity + objMatrix.Columns.Item("11").Cells.Item(Row).Specific.value
                            Next
                            Dim Temp As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Temp.DoQuery("Delete From LC_Percentage_value")
                            For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                                'UnitPrice = UnitPrice + CDbl(objMatrix.Columns.Item("14").Cells.Item(Row).Specific.value.ToString.Substring(3))
                                Temp.DoQuery("Insert into LC_Percentage_value Values ('" + objMatrix.Columns.Item("11").Cells.Item(Row).Specific.value + "','" + (objMatrix.Columns.Item("14").Cells.Item(Row).Specific.value.ToString.Substring(4)) + "','" + MAC_ID + "')")

                            Next
                        End If
                        If pVal.ItemUID = "btnlc" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.Before_Action = True Then
                            'GRPODOCNUM = objForm.Items.Item("8").Specific.Value
                            GRPODOCNUM = objForm.Items.Item("8").Specific.Value
                            'objForm.Items.Item("macid").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            'objForm.Items.Item("macid").Specific.value = MAC_ID
                            'objForm.Items.Item("macid").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            'If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            '    objForm.Items.Item(1).Click()
                            'End If
                            Dim A As Integer = 0
                            Quantity = 0
                            UnitPrice = 0
                            For Row As Integer = 1 To objMatrix.VisualRowCount
                                Quantity = Quantity + objMatrix.Columns.Item("11").Cells.Item(Row).Specific.value
                            Next
                            Dim Temp As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
5:                          Temp.DoQuery("Delete From LC_Percentage_value")
                            For Row As Integer = 1 To objMatrix.VisualRowCount
                                'UnitPrice = UnitPrice + CDbl(objMatrix.Columns.Item("14").Cells.Item(Row).Specific.value.ToString.Substring(3))
                                Temp.DoQuery("Insert into LC_Percentage_value Values ('" + objMatrix.Columns.Item("11").Cells.Item(Row).Specific.value + "','" + (objMatrix.Columns.Item("14").Cells.Item(Row).Specific.value.ToString.Substring(4)) + "','" + MAC_ID + "') ")

                            Next
                        End If
                        If pVal.ItemUID = "btnlc" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) And pVal.Action_Success = True Then
                            '  Me.CreateSubForm(FormUID)
                            Me.Open_LC_Form(FormUID)
                            'Dim objLCosts As ClsLCosts
                            'objLCosts.MenuEvent(pVal, BubbleEvent)
                        End If
                        If pVal.ItemUID = "btnlc" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE And pVal.Action_Success = True Then
                            '  Me.CreateSubForm(FormUID)
                            Me.Open_LC_Form_OK(FormUID)
                            'Dim objLCosts As ClsLCosts
                            'objLCosts.MenuEvent(pVal, BubbleEvent)
                        End If
                        If pVal.ItemUID = "btnit" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            If pVal.BeforeAction = True Then
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet.DoQuery("Select Distinct A.DocNum,B.ItemCode,B.WhsCode,B.Quantity From OPDN A Inner Join PDN1 B On A.DocEntry = B.DocEntry Where A.DocNum = '" + Trim(objForm.Items.Item("8").Specific.value) + "' And IsNull(B.u_insstat,'Open') = 'Open'")
                                If oRSet.RecordCount = 0 Then
                                    oApplication.StatusBar.SetText("Items already moved to main warehouse for this GRN", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                            If pVal.BeforeAction = False Then
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim RS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet.DoQuery("Select Distinct A.DocNum,B.ItemCode,B.WhsCode,(B.Quantity * B.NumPerMsr) - IsNull(B.u_openqty,0) AS 'Qty',B.LineNum From OPDN A Inner Join PDN1 B On A.DocEntry = B.DocEntry Where A.DocNum = '" + Trim(objForm.Items.Item("8").Specific.value) + "' And IsNull(B.u_insstat,'Open') = 'Open'")
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                If oRSet.RecordCount > 0 Then
                                    oApplication.ActivateMenuItem("3080")
                                    Dim ITForm As SAPbouiCOM.Form
                                    Dim ITMatrix As SAPbouiCOM.Matrix
                                    ITForm = oApplication.Forms.GetForm("940", oApplication.Forms.ActiveForm.TypeCount)
                                    ITMatrix = ITForm.Items.Item("23").Specific
                                    ITForm.Items.Item("grnno").Specific.value = oRSet.Fields.Item("DocNum").Value
                                    ITForm.Items.Item("18").Specific.value = oRSet.Fields.Item("WhsCode").Value
                                    ITMatrix.Columns.Item("U_grnno").Editable = True
                                    ITMatrix.Columns.Item("U_grnlnid").Editable = True
                                    For i As Integer = 1 To oRSet.RecordCount
                                        Try
                                            ITMatrix.Columns.Item("1").Cells.Item(i).Specific.value = oRSet.Fields.Item("ItemCode").Value
                                            RS.DoQuery("Select DfltWh From OITM Where ItemCode = '" + Trim(oRSet.Fields.Item("ItemCode").Value) + "'")
                                            If RS.RecordCount > 0 Then
                                                ITMatrix.Columns.Item("5").Cells.Item(i).Specific.value = RS.Fields.Item("DfltWh").Value
                                            End If
                                            If oRSet.Fields.Item("Qty").Value > 0 Then
                                                ITMatrix.Columns.Item("U_BAL_QTY").Editable = True
                                                ITMatrix.Columns.Item("U_BAL_QTY").Cells.Item(i).Specific.value = oRSet.Fields.Item("Qty").Value
                                                ITMatrix.Columns.Item("10").Cells.Item(i).Specific.value = oRSet.Fields.Item("Qty").Value

                                                ITMatrix.Columns.Item("U_BAL_QTY").Editable = False
                                            Else
                                                ITMatrix.Columns.Item("10").Cells.Item(i).Specific.value = 1
                                            End If
                                            ITMatrix.Columns.Item("U_grnno").Cells.Item(i).Specific.value = oRSet.Fields.Item("DocNum").Value
                                            ITMatrix.Columns.Item("U_grnlnid").Cells.Item(i).Specific.value = oRSet.Fields.Item("LineNum").Value
                                            ITMatrix.Columns.Item("1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            oRSet.MoveNext()
                                        Catch ex As Exception
                                            oApplication.StatusBar.SetText(ex.Message)
                                        End Try
                                    Next
                                    ITMatrix.Columns.Item("U_grnno").Editable = False
                                    ITMatrix.Columns.Item("U_grnlnid").Editable = False
                                End If
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        Dim LC As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim LC_Change As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim Value_Update As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim Value_Update1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        If pVal.ItemUID = "38" And ((pVal.ColUID = "11") Or (pVal.ColUID = "14")) And pVal.ActionSuccess = True Then
                            '  For i As Integer = 0 To objMatrix.VisualRowCount
                            '  If objMatrix.Columns.Item("11").Cells.Item(i).Specific.value Then
                            LC.DoQuery("Select (U_grpono),Code from [@GEN_GRPO_LCOSTS] where U_grpono='" + objForm.Items.Item("8").Specific.value + "' and u_macid='" + MAC_ID + "'")
                            Dim Quantity_Change As Integer = 0
                            For Row As Integer = 1 To objMatrix.VisualRowCount
                                Quantity_Change = Quantity_Change + objMatrix.Columns.Item("11").Cells.Item(Row).Specific.value
                            Next
                            Value_Update.DoQuery("Select * from [@GEN_GRPO_LCOSTS_D0] where Code= '" + LC.Fields.Item(1).Value + "' and u_glacct <> '' and u_macid='" + MAC_ID + "'")
                            Dim Temp As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Temp.DoQuery("Delete From LC_Percentage_value")

                            If LC.RecordCount > 0 Then
                                For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                                    'UnitPrice = UnitPrice + CDbl(objMatrix.Columns.Item("14").Cells.Item(Row).Specific.value.ToString.Substring(3))
                                    If objMatrix.Columns.Item("11").Cells.Item(Row).Specific.value <> "" And objMatrix.Columns.Item("14").Cells.Item(Row).Specific.value <> "" Then
                                        Temp.DoQuery("Insert into LC_Percentage_value Values ('" + objMatrix.Columns.Item("11").Cells.Item(Row).Specific.value + "','" + objMatrix.Columns.Item("11").Cells.Item(Row).Specific.value.ToString.Substring(3) + "','" + MAC_ID + "') ")
                                    End If

                                Next
                                For A As Integer = 1 To Value_Update.RecordCount
                                    LC_Change.DoQuery("Update [@GEN_GRPO_LCOSTS_D0] Set u_qty='" & Quantity_Change & "' Where Code='" + LC.Fields.Item(1).Value + "' and u_macid='" + MAC_ID + "' and LineId='" & A & "'")
                                    Dim LCPer As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    LCPer.DoQuery("Select U_type from OALC Where OALC.AlcName='" + Value_Update.Fields.Item("U_lname").Value + "'")
                                    Value_Update1.DoQuery("Select u_rate from [@GEN_GRPO_LCOSTS_D0] where Code='" + LC.Fields.Item(1).Value + "' and LineId='" & A & "'")
                                    Dim Value As Decimal
                                    Value = 0
                                    If LCPer.Fields.Item(0).Value = "Percent" Then
                                        Dim LC_Percent As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        LC_Percent.DoQuery("Select Price,Quantity from LC_Percentage_Value")
                                        For Row As Integer = 1 To LC_Percent.RecordCount
                                            UnitPrice = 0
                                            UnitPrice = UnitPrice + (((Value_Update1.Fields.Item(0).Value) / 100) * (LC_Percent.Fields.Item(0).Value))
                                            Value = Value + (UnitPrice * (LC_Percent.Fields.Item(1).Value))
                                            LC_Percent.MoveNext()
                                        Next
                                        LC_Change.DoQuery("Update [@GEN_GRPO_LCOSTS_D0] Set u_value='" & Value & "' Where Code='" + LC.Fields.Item(1).Value + "' and u_macid='" + MAC_ID + "' and LineId='" & A & "'")
                                    Else
                                        Value = Quantity_Change * (Value_Update1.Fields.Item(0).Value)
                                        LC_Change.DoQuery("Update [@GEN_GRPO_LCOSTS_D0] Set u_value='" & Value & "' Where Code='" + LC.Fields.Item(1).Value + "' and u_macid='" + MAC_ID + "' and LineId='" & A & "'")
                                    End If
                                Next
                            End If
                        End If
                End Select
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
                        If objForm.TypeEx = "143" Then
                            BubbleEvent = False
                        End If
                        'objSubForm = oApplication.Forms.GetForm("GEN_LCOSTS", oApplication.Forms.ActiveForm.TypeCount)
                    Case "1281"
                        'If pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        'GRPONO_Cancel = objForm.Items.Item("8").Specific.value
                        'End If
                        'If pVal.ActionSuccess = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim GRPO_Doc As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim LC_Check As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim LC_Check1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        GRPO_Doc.DoQuery("Select Docnum from OPDN Where Docnum='" & GRPODOCNUM & "'")
                        LC_Check1.DoQuery("Select u_grpono from [@GEN_GRPO_LCOSTS] where u_grpono='" + GRPODOCNUM + "' and u_macid='" + MAC_ID + "'")
                        If GRPO_Doc.RecordCount = 0 Then
                            If LC_Check1.RecordCount > 0 Then
                                LC_Check.DoQuery("Delete From [@GEN_GRPO_LCOSTS_D0] where code=(select code from [@GEN_GRPO_LCOSTS] where u_grpono='" & GRPODOCNUM & "'and u_macid='" + MAC_ID + "')")
                                LC_Check.DoQuery("Delete From [@GEN_GRPO_LCOSTS] where u_grpono='" & GRPODOCNUM & "'and u_macid='" + MAC_ID + "'")
                            End If
                        End If
                        'End If



                    Case "1282"
                        'Dim FormUID As String
                        If objForm.TypeEx = "143" Then
                            objForm.Items.Item("macid").Specific.value = MAC_ID
                            'TempItem = objForm.Items.Item("46")
                            'objItem = objForm.Items.Add("macid", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            'objItem.Top = TempItem.Top + TempItem.Height + 40
                            'objItem.Left = TempItem.Left
                            'objItem.Width = TempItem.Width
                            'objItem.Height = TempItem.Height
                            'objItem.Specific.databind.setbound(True, "OPDN", "u_macid")
                            'objItem.Specific.TabOrder = TempItem.Specific.TabOrder + 1
                            'objItem.Visible = True
                            'objItem.LinkTo = "46"
                            MAC_ID = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
                            objForm.Items.Item("macid").Specific.value = MAC_ID
                            'Me.CreateForm(FormUID)

                        End If
                        If objForm.TypeEx = "GEN_GRPO_LCOSTS" Then
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select Count(*) + 1 AS 'Count' From [@GEN_GRPO_LCOSTS]")
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_GRPO_LCOSTS")
                            oDBs_Head.SetValue("Code", 0, oRSet.Fields.Item("Count").Value)
                            objMatrix = objForm.Items.Item("mtx").Specific
                            objMatrix.AddRow()
                            Me.SetNewLine(objForm.UniqueID, objMatrix.VisualRowCount, objMatrix)
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
                        Dim oRset As SAPbobsCOM.Recordset
                        oRset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'oRset.DoQuery("Select u_reval from OPDN where Docnum='" + objForm.Items.Item("8").Specific.value + "'")
                        'If (oRset.Fields.Item(0).Value = "") And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        'objForm.Items.Item("btnmr").Enabled = False
                        'objForm.Items.Item("btnit").Enabled = False
                        'objForm.Items.Item("btnlc").Enabled = True
                        'End If
                    End If
                    If BusinessObjectInfo.ActionSuccess = True Then
                        Dim oRset As SAPbobsCOM.Recordset
                        oRset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRset.DoQuery("Select u_reval,u_lcadd from OPDN where Docnum='" + objForm.Items.Item("8").Specific.value + "'")
                        If (oRset.Fields.Item(0).Value = "Yes" Or oRset.Fields.Item(0).Value = "NO-LC") And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            objForm.Items.Item("btnit").Enabled = True
                            objForm.Items.Item("btnlc").Enabled = True
                            objForm.Items.Item("btnmr").Enabled = False
                        ElseIf (oRset.Fields.Item(0).Value = "") And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            objForm.Items.Item("btnit").Enabled = False
                            objForm.Items.Item("btnlc").Enabled = True
                            objForm.Items.Item("btnmr").Enabled = False
                        ElseIf (oRset.Fields.Item(0).Value = "" Or oRset.Fields.Item(0).Value = "No" Or oRset.Fields.Item(0).Value = "Partial") And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            'objForm.Items.Item("btnit").Enabled = False                            
                            objForm.Items.Item("btnit").Enabled = False
                            objForm.Items.Item("btnlc").Enabled = True
                            objForm.Items.Item("btnmr").Enabled = True
                        End If
                    End If
            End Select
        Catch ex As Exception
        End Try
    End Sub

    Sub ItemEvent_LC(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try


            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objSubForm = oApplication.Forms.Item(FormUID)
                    objSubForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
                    objLCMatrix = objSubForm.Items.Item("mtx").Specific
                    ' objLCMatrix.state = SAPbouiCOM.BoFormStateEnum.fs_Maximized
                    PARENT_FORM = (pVal.FormUID.Substring(objSubForm.UniqueID.IndexOf("@") + 1))
                    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRs.DoQuery("Select * from OPDN Where Docnum='" + objSubForm.Items.Item("grpono").Specific.value + "'")
                    '  objSubForm.Items.Item("grpono").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    'objLCMatrix.Columns.Item("lname").Cells.Item(1
                    'objSubForm.Items.Item("grpono").Enabled = False
                    If oRs.RecordCount > 0 Then
                        'objSubForm.Items.Item("1").GetAutoManagedAttribute.
                        'objItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        'objSubForm.Items.Item("grpono").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        ''objLCMatrix.Columns.Item("glacct"). = True
                        'objLCMatrix.Columns.Item("lname").Editable = False
                        'objLCMatrix.Columns.Item("rate").Editable = False
                        'objLCMatrix.Columns.Item("glacct").Editable = False
                        'objLCMatrix.Columns.Item("qty").Editable = False
                        'objLCMatrix.Columns.Item("value").Editable = False
                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.ActionSuccess = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        Dim oRSet1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRSet1.DoQuery("SELECT top 1 u_grpono,u_macid from [@GEN_GRPO_LCOSTS] WHERE u_macid='" + MAC_ID + "' order by u_grpono desc ")
                        '    objSubForm.Items.Item("grpono").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        objForm.Items.Item("grpono").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        objForm.Items.Item("macid").Visible = True
                        objForm.Items.Item("macid").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        objForm.Items.Item("grpono").Specific.value = oRSet1.Fields.Item(0).Value
                        ' oRSet1.DoQuery("Update OPDN Set u_lcadd='Y' where OPDN.Docnum='" + objForm.Items.Item("grpono").Specific.value + "' and u_macid='" + MAC_ID + "'")
                        objForm.Items.Item("macid").Specific.value = oRSet1.Fields.Item(1).Value
                        oRSet1.DoQuery("Update OPDN Set u_lcadd='Y' where OPDN.Docnum='" + objForm.Items.Item("grpono").Specific.value + "' and u_macid='" + MAC_ID + "'")
                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objForm.Items.Item("grpono").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        objForm.Items.Item("macid").Visible = False

                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then

                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Dim LCCode As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            LCCode.DoQuery("SELECT u_grpono FROM [@GEN_GRPO_LCOSTS] T0  WHERE T0.U_grpono='" + GRPODOCNUM + "' and T0.u_macid='" + MAC_ID + "' order by T0.code desc")
                            If LCCode.RecordCount = 0 Then
                                Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSet.DoQuery("Select code+ 1 AS 'Count' From [@GEN_GRPO_LCOSTS] where code=(Select top 1 code from [@GEN_GRPO_LCOSTS] order by DocEntry desc)")
                                oDBs_Head = objSubForm.DataSources.DBDataSources.Item("@GEN_GRPO_LCOSTS")
                                oDBs_Head.SetValue("Code", 0, oRSet.Fields.Item("Count").Value)
                                'objMatrix = objSubForm.Items.Item("mtx").Specific
                                'objMatrix.AddRow()
                                'objMatrix.SetLineData(pVal.Row)
                                'ITMatrix.FlushToDataSource()
                                'Me.SetNewLine(objSubForm.UniqueID, objMatrix.VisualRowCount, objMatrix)
                            End If
                        End If

                        If objLCMatrix.Columns.Item("glacct").Cells.Item(1).Specific.value = "" Then
                            oApplication.StatusBar.SetText("Please Select Landed Cost In Row Level", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                        End If

                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    ' objForm = oApplication.Forms.Item("OPDN")
                    'objMatrix=objForm.Items.Item("38").Specific
                    objSubForm = oApplication.Forms.Item(pVal.FormUID)
                    objLCMatrix = objSubForm.Items.Item("mtx").Specific
                    oDBs_Head = objSubForm.DataSources.DBDataSources.Item("@GEN_GRPO_LCOSTS")
                    oDBs_Detail = objSubForm.DataSources.DBDataSources.Item("@GEN_GRPO_LCOSTS_D0")
                    'oDBs_Head = objForm.DataSources.DBDataSources.Item("OPDN")
                    'oDBs_Detail = objForm.DataSources.DBDataSources.Item("PDN1")
                    ' If objLCMatrix.Columns.Item("lname").Cells.Item(1).Specific.value <> "" Then
                    ''  Dim Lname As String
                    'Lname = objLCMatrix.Columns.Item("glacct").Cells.Item(1).Specific.value
                    Dim LCPer As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    LCPer.DoQuery("Select U_type from OALC Where OALC.AlcName='" + (objLCMatrix.Columns.Item("lname").Cells.Item(1).Specific.value.ToString) + "'")
                    If LCPer.Fields.Item(0).Value <> "Percent" Then
                        If pVal.ItemUID = "mtx" And (pVal.ColUID = "qty" Or pVal.ColUID = "rate") And pVal.ActionSuccess = True Then
                            For i As Integer = 1 To objLCMatrix.VisualRowCount - 1
                                Dim Qty As Decimal = 0
                                If objLCMatrix.Columns.Item("qty").Cells.Item(i).Specific.value = "" Then
                                    Qty = Quantity
                                Else
                                    Qty = objLCMatrix.Columns.Item("qty").Cells.Item(i).Specific.value
                                End If
                                objLCMatrix.Columns.Item("value").Cells.Item(i).Specific.value = CDbl((Qty) * (objLCMatrix.Columns.Item("rate").Cells.Item(i).Specific.value))
                            Next
                        End If
                    End If
                    '   End If

                    If pVal.ItemUID = "mtx" And pVal.ColUID = "lname" And pVal.ActionSuccess = True Then
                        Dim Flag As Boolean = False
                        Dim errflag As Boolean = False
                        If objLCMatrix.VisualRowCount = 1 Or pVal.Row = objLCMatrix.VisualRowCount Then
                            Flag = True
                        End If
                        If Flag = True Then
                            objLCMatrix.AddRow(1, objLCMatrix.VisualRowCount)
                            Me.SetNewLine(FormUID, objLCMatrix.VisualRowCount, objLCMatrix)
                        End If
                        For i As Integer = 1 To objLCMatrix.VisualRowCount - 1
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("SELECT ALCNAME,LaCAllcAcc,U_rate,OACT.FormatCode FROM OALC INNER JOIN OACT ON OACT.AcctCode=OALC.LaCAllcAcc where AlcName='" + objLCMatrix.Columns.Item("lname").Cells.Item(i).Specific.value + "'")
                            '  Dim LCPer As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            LCPer.DoQuery("Select U_type from OALC Where OALC.AlcName='" + objLCMatrix.Columns.Item("lname").Cells.Item(i).Specific.value + "'")

                            objLCMatrix.Columns.Item("rate").Cells.Item(i).Specific.value = oRSet.Fields.Item(2).Value
                            'End If

                            objLCMatrix.Columns.Item("glacct").Cells.Item(i).Specific.value = oRSet.Fields.Item(3).Value
                            objLCMatrix.Columns.Item("qty").Cells.Item(i).Specific.value = Quantity
                            Dim Value As Decimal
                            Value = 0
                            If LCPer.Fields.Item(0).Value = "Percent" Then
                                Dim LC_Percent As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                LC_Percent.DoQuery("Select Price,Quantity from LC_Percentage_Value")
                                For Row As Integer = 1 To LC_Percent.RecordCount
                                    UnitPrice = 0
                                    UnitPrice = UnitPrice + (((oRSet.Fields.Item(2).Value) / 100) * CDbl(LC_Percent.Fields.Item(0).Value))
                                    Value = Value + CDbl(UnitPrice * (LC_Percent.Fields.Item(1).Value))
                                    LC_Percent.MoveNext()
                                Next
                                objLCMatrix.Columns.Item("value").Cells.Item(i).Specific.value = Value
                            Else
                                objLCMatrix.Columns.Item("value").Cells.Item(i).Specific.value = CDbl((objLCMatrix.Columns.Item("qty").Cells.Item(i).Specific.value) * (objLCMatrix.Columns.Item("rate").Cells.Item(i).Specific.value))
                            End If
                            objLCMatrix.Columns.Item("macid").Cells.Item(i).Specific.value = MAC_ID
                        Next

                    End If
            End Select
        Catch ex As Exception
            'oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub InventoryRevaluation()
        Try
            'Declarations
            Dim oSR As SAPbobsCOM.MaterialRevaluation
            'oForm = oApplication.Forms.Item("143")
            'Dim CompanyService As SAPbobsCOM.CompanyService
            Dim oSR_Lines As SAPbobsCOM.MaterialRevaluation_lines
            'Dim RS As SAPbobsCOM.Recordset

            oSR = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
            Dim oRecordSet1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRecordSet2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'CompanyService = oCompany.GetCompanyService
            'RS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '  '  oRecordSet.DoQuery("SELECT T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], T1.[LineNum], T1.[ItemCode],T1.[Quantity], T1.[WhsCode], T1.[TaxCode] FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry Where T0.DocEntry = (select Max(DocEntry) from OPDN)")
            oRecordSet.DoQuery("SELECT T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode],  T1.[ItemCode],sum(T1.[Quantity])'Quantity', T1.[WhsCode], T1.[TaxCode],sum(T1.LineTotal)'LineTotal',(select sum(distinct(PDN4.TaxSum)) from PDN1 inner join PDN4 on pdn1.DocEntry=pdn4.DocEntry And PDN1.LineNum = PDN4.LineNum where pdn4.NonDdctPrc='100' and pdn1.ItemCode=T1.ItemCode and pdn1.DocEntry=(Select MAX(Docentry)from OPDN))'Vatsum' FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry Where T0.DocEntry = (select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "' group by T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode],  T1.[ItemCode], T1.[WhsCode], T1.[TaxCode]")
            oRecordSet.MoveFirst()
            oRecordSet1.DoQuery("delete from TempReval where UserName = '" & oCompany.UserName & "' ")
            For I As Int16 = 0 To oRecordSet.RecordCount - 1
                '  ' oRecordSet1.DoQuery("SELECT T2.StaCode,T2.TaxSum,T2.LineNum,T1.ItemCode FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry Inner join PDN4 T2 ON T1.DocEntry = T2.DocEntry And T1.LineNum = T2.LineNum Where T0.DocEntry = (select Max(DocEntry) from OPDN) And T1.LineNum = '" & oRecordSet.Fields.Item("LineNum").Value & "' And T2.ExpnsCode = -1 And T2.NonDdctPrc='100'")
                '  ' For J As Int16 = 0 To oRecordSet1.RecordCount - 1
                ' oRecordSet2.DoQuery("Select U_Allow From [@TAXREV] Where Code = '" & oRecordSet1.Fields.Item("StaCode").Value & "'")
                ' If oRecordSet2.Fields.Item("U_Allow").Value = "YES" Then
                oRecordSet2.DoQuery("Insert into TempReval Values('" & oRecordSet.Fields.Item("ItemCode").Value & "','" & oRecordSet.Fields.Item("Quantity").Value & "','" & (oRecordSet.Fields.Item("Vatsum").Value) & "','" & oRecordSet.Fields.Item("WhsCode").Value & "','" & oRecordSet.Fields.Item("DocDate").Value & "','" & oRecordSet.Fields.Item("DocNum").Value & "','" & oApplication.Company.UserName & "')")
                'End If
                '  'oRecordSet1.MoveNext()
                '  ' Next
                oRecordSet.MoveNext()
            Next
            oRecordSet.DoQuery("Select Sum(Taxvalue),ItemCode,Quantity,WhsCode,Date,DocNum from TempReval Where UserName = '" & oApplication.Company.UserName & "'  Group By ItemCode,WhsCode,Date,DocNum,UserName,Quantity")
            oSR.RevalType = "P"
            oSR.Reference2 = oRecordSet.Fields.Item("DocNum").Value
            oSR.Comments = "ItemCost is revaluated for exciseduty or Cst tax amount(Excise Duty+CST/Instock of item in above Warehouse)"
            'oSR.Comments = "ItemCost is revaluated for LC(LC/Instock of item in above Warehouse)"
            'oSR.u_macid = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name
            ' oSR.UserFields.Fields.Item("MAC_ID").Value = oApplication.AppId & "_" & oCompany.UserName & "_" & My.Computer.Name

            ''****line Added by vivek for curecting date conversion***********

            Dim lSBObob As SAPbobsCOM.SBObob
            Dim lRecordset As SAPbobsCOM.Recordset
            Dim ld_CurrentItemCost As Decimal
            Dim ld_NewItemCost As Decimal

            ''lSBObob = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            ''lRecordset = lSBObob.Format_StringToDate(oRecordSet.Fields.Item("Date").Value)

            oSR.DocDate = oRecordSet.Fields.Item("Date").Value
            ''oSR.DocDate = lRecordset.Fields.Item(0).Value

            ''*********************************

            For I As Int16 = 0 To oRecordSet.RecordCount - 1
                oRecordSet1.DoQuery("Select AvgPrice,OnHand  From OITW where ItemCode = '" & oRecordSet.Fields.Item("ItemCode").Value & "' And WhsCode = '" & oRecordSet.Fields.Item("WhsCode").Value & "'")


                If oRecordSet1.RecordCount > 0 Then
                    ld_CurrentItemCost = Val(oRecordSet1.Fields.Item("AvgPrice").Value)
                    ld_NewItemCost = Math.Round(CDbl(oRecordSet1.Fields.Item("AvgPrice").Value) + (CDbl(oRecordSet.Fields.Item(0).Value) / CDbl(oRecordSet1.Fields.Item(1).Value)), 4)
                Else
                    ld_CurrentItemCost = 0
                    ld_NewItemCost = 0
                End If

                ''oApplication.MessageBox("current cost : " + CStr(ld_CurrentItemCost) + "new cost : " + CStr(ld_NewItemCost))
                If ld_CurrentItemCost <> ld_NewItemCost Then
                    oSR_Lines = oSR.Lines
                    oSR_Lines.SetCurrentLine(I)
                    oSR_Lines.ItemCode = oRecordSet.Fields.Item("ItemCode").Value
                    oSR_Lines.WarehouseCode = oRecordSet.Fields.Item("WhsCode").Value

                    oSR_Lines.Price = ld_NewItemCost
                    ' oSR_Lines.RevaluationIncrementAccount = "401020"
                    'oSR_Lines.RevaluationDecrementAccount = "401020"
                    ' oForm.Items.Item()
                    oSR_Lines.Add()
                End If
                oRecordSet.MoveNext()
            Next
            If oSR.Add <> 0 Then
                Dim oRset As SAPbobsCOM.Recordset
                oRset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRset.DoQuery("select count(stacode) from PDN4 inner join OPDN on opdn.DocEntry=pdn4.DocEntry where OPDN.DocEntry=(Select Max(Docentry) from opdn) and NonDdctPrc='100'")
                If oRset.Fields.Item(0).Value = 0 Then
                    '  Dim oLC As SAPbobsCOM.Recordset
                    ' oLC = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'oLC.DoQuery("Select T1.u_rate,T1.u_glacct,T1.u_value from [@GEN_GRPO_LCOSTS] T0 Inner Join [@GEN_GRPO_LCOSTS_D0] T1 on T0.Code=T1.Code Where T0.Docentry=(Select Max(DocEntry) from [@GEN_GRPO_LCOSTS] Where u_macid='" + MAC_ID + "') and T0.u_macid='" + MAC_ID + "'  And T1.u_glacct<>''")
                    Dim NA As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    NA.DoQuery("Select count(T1.u_rate) from [@GEN_GRPO_LCOSTS] T0 inner join [@GEN_GRPO_LCOSTS_D0] T1 on T0.Code=T1.Code Where T1.u_rate>'0.00' and T0.Docentry=(Select Max(DocEntry) from [@GEN_GRPO_LCOSTS] Where u_macid='" + MAC_ID + "') and T0.u_macid='" + MAC_ID + "'  ")
                    If NA.Fields.Item(0).Value > 0 Then

                        'Select Rate From GEN_GRPO_LCOSTS
                        'If Rate>0 then else error.
                        oRecordSet1.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                        Dim oRs As SAPbobsCOM.Recordset
                        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRs.DoQuery("Select T1.u_rate,T1.u_glacct,T1.u_value from [@GEN_GRPO_LCOSTS] T0 Inner Join [@GEN_GRPO_LCOSTS_D0] T1 on T0.Code=T1.Code Where T0.Docentry=(Select Max(DocEntry) from [@GEN_GRPO_LCOSTS] Where u_macid='" + MAC_ID + "') and T0.u_macid='" + MAC_ID + "' And T1.u_glacct<>''")
                        Dim oRs2 As SAPbobsCOM.Recordset
                        oRs2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRs2.DoQuery("Select count(T1.u_rate) from [@GEN_GRPO_LCOSTS] T0 Inner Join [@GEN_GRPO_LCOSTS_D0] T1 on T0.Code=T1.Code Where T0.u_grpono='" + GRPODOCNUM + "' And T1.u_glacct<>'' and T0.u_macid='" + MAC_ID + "'")
                        If oRs2.Fields.Item(0).Value > 0 Then
                            If oRs.Fields.Item(0).Value > "0.00" Then


                                Dim LCPer As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                LCPer.DoQuery("Select U_type from OALC Where OALC.AlcName=(Select u_lname from [@GEN_GRPO_LCOSTS] T0 inner join [@GEN_GRPO_LCOSTS_D0] T1 on T0.Code=T1.Code Where T0.u_grpono='" + GRPODOCNUM + "' And T0.u_macid='" + MAC_ID + "' and U_lname<> '')")
                                If LCPer.Fields.Item(0).Value = "Percent" Then
                                    For J As Int16 = 0 To oRs.RecordCount - 1
                                        oRecordSet.DoQuery("SELECT T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], SUm(T1.[Vatsum]+T1.[LineTotal])'VatSum', T1.[ItemCode],Sum(T1.[Quantity])'Quantity',sum(T1.[Price])'price', T1.[WhsCode], T1.[TaxCode] FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry Where T0.DocEntry = (select Max(DocEntry) from OPDN) and u_macid='" + MAC_ID + "' group by T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], T1.[ItemCode], T1.[WhsCode], T1.[TaxCode]")
                                        oRecordSet.MoveFirst()

                                        For I As Int16 = 0 To oRecordSet.RecordCount - 1
                                            '  oRecordSet1.DoQuery("Select u_rate,u_expact from [@GEN_EXP_TAX]")
                                            'Dim Expenses_Tax As Double
                                            ' Expenses_Tax = (CDbl(oRecordSet.Fields.Item("LineTotal").Value) * (CDbl(oRecordSet1.Fields.Item("u_rate").Value) / 100))
                                            oRecordSet2.DoQuery("Insert into TempExpReval Values('" & oRecordSet.Fields.Item("ItemCode").Value & "','" & oRecordSet.Fields.Item("Quantity").Value & "','" & (CDbl(oRecordSet.Fields.Item("VatSum").Value)) & "','" & oRecordSet.Fields.Item("WhsCode").Value & "','" & oRecordSet.Fields.Item("DocDate").Value & "','" & oRecordSet.Fields.Item("DocNum").Value & "','" + oRs.Fields.Item("u_glacct").Value + "','" & oApplication.Company.UserName & "','" & oRecordSet.Fields.Item("Price").Value & "')")
                                            oRecordSet.MoveNext()
                                        Next
                                        oRset.DoQuery("Select Sum(Taxvalue),ItemCode,Quantity,WhsCode,Date,DocNum,Sum(price)'price' from TempExpReval Where UserName = '" & oApplication.Company.UserName & "'  Group By ItemCode,WhsCode,Date,DocNum,UserName,Quantity")

                                        Dim oSR1 As SAPbobsCOM.MaterialRevaluation
                                        Dim oSR_Lines1 As SAPbobsCOM.MaterialRevaluation_lines
                                        oSR1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
                                        oSR1.RevalType = "P"
                                        'oSR1.Reference1 = oRecordSet.Fields.Item("DocNum").Value
                                        oSR.Reference2 = oRset.Fields.Item("DocNum").Value
                                        'SR.Comments = "ItemCost is revaluated for exciseduty or Cst tax amount(Excise Duty+CST/Instock of item in above Warehouse)"
                                        oSR.Comments = "ItemCost is revaluated for LC(LC/Instock of item in above Warehouse)"

                                        ' Dim lSBObob As SAPbobsCOM.SBObob
                                        'Dim lRecordset As SAPbobsCOM.Recordset
                                        Dim ld_CurrentItemCost1 As Decimal
                                        Dim ld_NewItemCost1 As Decimal

                                        oSR1.DocDate = oRset.Fields.Item("Date").Value

                                        ''*********************************
                                        For I As Int16 = 0 To oRset.RecordCount - 1
                                            'oRecordSet1.DoQuery("select price from MRV1 inner join OMRV on OMRV.docentry=MRV1.docentry where Itemcode='" + oRecordSet.Fields.Item("ItemCode").Value + "' and DocNum=(Select MAX(docnum)from omrv)")
                                            oRecordSet1.DoQuery("Select AvgPrice,OnHand From OITW where ItemCode = '" & oRset.Fields.Item("ItemCode").Value & "' And WhsCode = '" & oRset.Fields.Item("WhsCode").Value & "'")
                                            If oRecordSet1.RecordCount > 0 Then
                                                ld_CurrentItemCost1 = Val(oRecordSet1.Fields.Item("AvgPrice").Value)
                                                ld_NewItemCost1 = Math.Round(CDbl((oRecordSet1.Fields.Item("AvgPrice").Value) + ((((oRs.Fields.Item(0).Value / 100) * (oRset.Fields.Item("Price").Value)) * (oRset.Fields.Item("Quantity").Value)) / oRecordSet1.Fields.Item(1).Value)), 4)
                                            Else
                                                ld_CurrentItemCost1 = 0
                                                ld_NewItemCost1 = 0
                                            End If
                                            If ld_CurrentItemCost1 <> ld_NewItemCost1 Then
                                                oSR_Lines1 = oSR.Lines
                                                oSR_Lines1.SetCurrentLine(I)
                                                oSR_Lines1.ItemCode = oRset.Fields.Item("ItemCode").Value
                                                oSR_Lines1.WarehouseCode = oRset.Fields.Item("WhsCode").Value

                                                oSR_Lines1.Price = ld_NewItemCost1
                                                oRecordSet2.DoQuery("select acctcode from OACT where FormatCode='" + oRs.Fields.Item(1).Value + "'")
                                                oSR_Lines1.RevaluationIncrementAccount = oRecordSet2.Fields.Item(0).Value
                                                oSR_Lines1.RevaluationDecrementAccount = oRecordSet2.Fields.Item(0).Value
                                                oSR_Lines1.Add()
                                            End If
                                            oRset.MoveNext()
                                        Next
                                        If oSR.Add <> 0 Then
                                            Dim Str As String = oCompany.GetLastErrorDescription
                                            oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription)
                                            ' Else
                                            oRecordSet2.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                                            GoTo A
                                        End If
                                        oRecordSet2.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                                        oRs.MoveNext()
                                        ' End If
                                    Next
                                    ' Dim Str As String = oCompany.GetLastErrorDescription
                                    oApplication.StatusBar.SetText("Inventory Revaluation was successfully posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    oRecordSet.DoQuery("Update OPDN set u_reval='Yes' where DocEntry=(Select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "'")
                                    Dim oRs3 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim oRs4 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRs3.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocEntry=(Select Max(DocEntry) from OPDN) ")
                                    oRs4.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocEntry=(Select Max(DocEntry) from OPDN) order by omrv.DocNum desc")
                                    If oRs3.Fields.Item(0).Value = oRs4.Fields.Item(0).Value Then
                                        oRecordSet.DoQuery("Update OPDN set u_revalno=('" & oRs3.Fields.Item(0).Value & "')where DocEntry=(Select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "'")
                                    Else

                                        oRecordSet.DoQuery("Update OPDN set u_revalno=('" & oRs3.Fields.Item(0).Value & "' +' To '+ '" & oRs4.Fields.Item(0).Value & "')where DocEntry=(Select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "' ")
                                    End If
                                    Dim _str_DocNum1 As String
                                    Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRs1.DoQuery("select DocNum,u_unit from OPDN Where DocEntry=(Select Max(DocEntry) from opdn)")
                                    _str_DocNum1 = oRs1.Fields.Item(0).Value
                                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    objForm.Items.Item("8").Specific.value = _str_DocNum1
                                    objForm.Items.Item("cpc").Specific.value = oRs1.Fields.Item(1).Value
                                    objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Else
                                    For J As Int16 = 0 To oRs.RecordCount - 1
                                        oRecordSet.DoQuery("SELECT T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], SUm(T1.[Vatsum]+T1.[LineTotal])'VatSum', T1.[ItemCode],Sum(T1.[Quantity])'Quantity', T1.[WhsCode], T1.[TaxCode] FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry Where T0.DocEntry = (select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "' group by T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], T1.[ItemCode], T1.[WhsCode], T1.[TaxCode]")
                                        oRecordSet.MoveFirst()


                                        For I As Int16 = 0 To oRecordSet.RecordCount - 1
                                            '  oRecordSet1.DoQuery("Select u_rate,u_expact from [@GEN_EXP_TAX]")
                                            'Dim Expenses_Tax As Double
                                            ' Expenses_Tax = (CDbl(oRecordSet.Fields.Item("LineTotal").Value) * (CDbl(oRecordSet1.Fields.Item("u_rate").Value) / 100))
                                            oRecordSet2.DoQuery("Insert into TempExpReval Values('" & oRecordSet.Fields.Item("ItemCode").Value & "','" & oRecordSet.Fields.Item("Quantity").Value & "','" & (CDbl(oRecordSet.Fields.Item("VatSum").Value)) & "','" & oRecordSet.Fields.Item("WhsCode").Value & "','" & oRecordSet.Fields.Item("DocDate").Value & "','" & oRecordSet.Fields.Item("DocNum").Value & "','" + oRs.Fields.Item("u_glacct").Value + "','" & oApplication.Company.UserName & "','0.00')")
                                            oRecordSet.MoveNext()
                                        Next
                                        oRset.DoQuery("Select Sum(Taxvalue),ItemCode,Quantity,WhsCode,Date,DocNum from TempExpReval Where UserName = '" & oApplication.Company.UserName & "'  Group By ItemCode,WhsCode,Date,DocNum,UserName,Quantity")

                                        Dim oSR1 As SAPbobsCOM.MaterialRevaluation
                                        Dim oSR_Lines1 As SAPbobsCOM.MaterialRevaluation_lines
                                        oSR1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
                                        oSR1.RevalType = "P"
                                        'oSR1.Reference1 = oRecordSet.Fields.Item("DocNum").Value
                                        oSR.Reference2 = oRset.Fields.Item("DocNum").Value
                                        'oSR.Comments = "ItemCost is revaluated for exciseduty or Cst tax amount(Excise Duty+CST/Instock of item in above Warehouse)"
                                        oSR.Comments = "ItemCost is revaluated for LC(LC/Instock of item in above Warehouse)"
                                        'oSR.Comments = "ItemCost is revaluated for exciseduty or Cst tax amount(Excise Duty+CST/Instock of item in above Warehouse)"

                                        ' Dim lSBObob As SAPbobsCOM.SBObob
                                        'Dim lRecordset As SAPbobsCOM.Recordset
                                        Dim ld_CurrentItemCost1 As Decimal
                                        Dim ld_NewItemCost1 As Decimal

                                        oSR1.DocDate = oRset.Fields.Item("Date").Value

                                        ''*********************************
                                        For I As Int16 = 0 To oRset.RecordCount - 1
                                            'oRecordSet1.DoQuery("select price from MRV1 inner join OMRV on OMRV.docentry=MRV1.docentry where Itemcode='" + oRecordSet.Fields.Item("ItemCode").Value + "' and DocNum=(Select MAX(docnum)from omrv)")
                                            oRecordSet1.DoQuery("Select AvgPrice,OnHand  From OITW where ItemCode = '" & oRset.Fields.Item("ItemCode").Value & "' And WhsCode = '" & oRset.Fields.Item("WhsCode").Value & "'")
                                            If oRecordSet1.RecordCount > 0 Then
                                                ld_CurrentItemCost1 = Val(oRecordSet1.Fields.Item("AvgPrice").Value)
                                                ld_NewItemCost1 = Math.Round(CDbl(oRecordSet1.Fields.Item("AvgPrice").Value) + ((oRs.Fields.Item(0).Value * oRset.Fields.Item(2).Value) / oRecordSet1.Fields.Item(1).Value), 4)

                                            Else
                                                ld_CurrentItemCost1 = 0
                                                ld_NewItemCost1 = 0
                                            End If
                                            If ld_CurrentItemCost1 <> ld_NewItemCost1 Then
                                                oSR_Lines1 = oSR.Lines
                                                oSR_Lines1.SetCurrentLine(I)
                                                oSR_Lines1.ItemCode = oRset.Fields.Item("ItemCode").Value
                                                oSR_Lines1.WarehouseCode = oRset.Fields.Item("WhsCode").Value

                                                oSR_Lines1.Price = ld_NewItemCost1
                                                oRecordSet2.DoQuery("select acctcode from OACT where FormatCode='" + oRs.Fields.Item(1).Value + "'")
                                                oSR_Lines1.RevaluationIncrementAccount = oRecordSet2.Fields.Item(0).Value
                                                oSR_Lines1.RevaluationDecrementAccount = oRecordSet2.Fields.Item(0).Value
                                                oSR_Lines1.Add()
                                            End If
                                            oRset.MoveNext()
                                        Next
                                        If oSR.Add <> 0 Then
                                            Dim Str As String = oCompany.GetLastErrorDescription
                                            oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription)
                                            ' Else
                                            oRecordSet2.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                                            GoTo A
                                        End If
                                        oRecordSet2.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                                        oRs.MoveNext()
                                        ' End If
                                    Next
                                    ' Dim Str As String = oCompany.GetLastErrorDescription
                                    oApplication.StatusBar.SetText("Inventory Revaluation was successfully posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    oRecordSet.DoQuery("Update OPDN set u_reval='Yes' where DocEntry=(Select Max(DocEntry) from OPDN)And u_macid='" + MAC_ID + "' ")
                                    Dim oRs3 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim oRs4 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRs3.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocEntry=(Select Max(DocEntry) from OPDN) ")
                                    oRs4.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocEntry=(Select Max(DocEntry) from OPDN)  order by omrv.DocNum desc")
                                    If oRs3.Fields.Item(0).Value = oRs4.Fields.Item(0).Value Then
                                        oRecordSet.DoQuery("Update OPDN set u_revalno=('" & oRs3.Fields.Item(0).Value & "')where DocEntry=(Select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "'")
                                    Else
                                        oRecordSet.DoQuery("Update OPDN set u_revalno=('" & oRs3.Fields.Item(0).Value & "' +' To '+ '" & oRs4.Fields.Item(0).Value & "')where DocEntry=(Select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "' ")
                                    End If
                                    Dim _str_DocNum1 As String
                                    Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRs1.DoQuery("select DocNum,u_unit from OPDN Where DocEntry=(Select Max(DocEntry) from opdn)")
                                    _str_DocNum1 = oRs1.Fields.Item(0).Value
                                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    objForm.Items.Item("8").Specific.value = _str_DocNum1
                                    objForm.Items.Item("cpc").Specific.value = oRs1.Fields.Item(1).Value
                                    objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                End If
                            Else
                                oRecordSet.DoQuery("Update OPDN set u_reval='NO-LC' where DocEntry=(Select Max(DocEntry) from OPDN)And u_macid='" + MAC_ID + "' ")
                                Dim _str_DocNum1 As String
                                Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRs1.DoQuery("select DocNum,u_unit from OPDN Where DocEntry=(Select Max(DocEntry) from opdn)")
                                _str_DocNum1 = oRs1.Fields.Item(0).Value
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                objForm.Items.Item("8").Specific.value = _str_DocNum1
                                objForm.Items.Item("cpc").Specific.value = oRs1.Fields.Item(1).Value
                                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                            End If
                        Else
                            oRecordSet.DoQuery("Update OPDN set u_reval='NO-LC' where DocEntry=(Select Max(DocEntry) from OPDN)And u_macid='" + MAC_ID + "' ")
                            Dim _str_DocNum1 As String
                            Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRs1.DoQuery("select DocNum,u_unit from OPDN Where DocEntry=(Select Max(DocEntry) from opdn)")
                            _str_DocNum1 = oRs1.Fields.Item(0).Value
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("8").Specific.value = _str_DocNum1
                            objForm.Items.Item("cpc").Specific.value = oRs1.Fields.Item(1).Value
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                            '  oApplication.StatusBar.SetText("Select LandedCost by Clicking LC Button", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If

                        Else
                        oRecordSet.DoQuery("Update OPDN set u_reval='NO-LC' where DocEntry=(Select Max(DocEntry) from OPDN)And u_macid='" + MAC_ID + "' ")
                        Dim _str_DocNum1 As String
                        Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRs1.DoQuery("select DocNum,u_unit from OPDN Where DocEntry=(Select Max(DocEntry) from opdn)")
                        _str_DocNum1 = oRs1.Fields.Item(0).Value
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objForm.Items.Item("8").Specific.value = _str_DocNum1
                        objForm.Items.Item("cpc").Specific.value = oRs1.Fields.Item(1).Value
                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        'oApplication.StatusBar.SetText("Select LandedCost by Clicking LC Button", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If


A:
                        ' End If
                    Else
                        Dim Str As String = oCompany.GetLastErrorDescription
                        oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription)
                    End If
                Else
                    Dim NA As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    NA.DoQuery("Select count(T1.u_rate) from [@GEN_GRPO_LCOSTS] T0 inner join [@GEN_GRPO_LCOSTS_D0] T1 on T0.Code=T1.Code Where T1.u_rate>'0.00' and T0.Docentry=(Select Max(DocEntry) from [@GEN_GRPO_LCOSTS] Where u_macid='" + MAC_ID + "') and T0.u_macid='" + MAC_ID + "'  ")
                    If NA.Fields.Item(0).Value > 0 Then

                        'End If
                        oRecordSet1.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                        Dim oRs As SAPbobsCOM.Recordset
                        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRs.DoQuery("Select T1.u_rate,T1.u_glacct,u_value from [@GEN_GRPO_LCOSTS] T0 Inner Join [@GEN_GRPO_LCOSTS_D0] T1 on T0.Code=T1.Code Where T0.Docentry=(Select Max(DocEntry) from [@GEN_GRPO_LCOSTS] Where u_macid='" + MAC_ID + "') and T0.u_macid='" + MAC_ID + "' And T1.u_glacct<>''")
                        Dim LCPer As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        LCPer.DoQuery("Select U_type from OALC Where OALC.AlcName=(Select top 1(u_lname) from [@GEN_GRPO_LCOSTS] T0 inner join [@GEN_GRPO_LCOSTS_D0] T1 on T0.Code=T1.Code Where T0.u_grpono='" + GRPODOCNUM + "' And T0.u_macid='" + MAC_ID + "' and U_rate>0)")
                        If LCPer.Fields.Item(0).Value = "Percent" Then
                            For J As Int16 = 0 To oRs.RecordCount - 1
                                oRecordSet.DoQuery("SELECT T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], SUm(T1.[Vatsum]+T1.[LineTotal])'VatSum', T1.[ItemCode],Sum(T1.[Quantity])'Quantity',sum(T1.[Price])'price', T1.[WhsCode], T1.[TaxCode] FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry Where T0.DocEntry = (select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "' group by T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], T1.[ItemCode], T1.[WhsCode], T1.[TaxCode]")
                                oRecordSet.MoveFirst()
                                For I As Int16 = 0 To oRecordSet.RecordCount - 1
                                    '  oRecordSet1.DoQuery("Select u_rate,u_expact from [@GEN_EXP_TAX]")
                                    'Dim Expenses_Tax As Double
                                    ' Expenses_Tax = (CDbl(oRecordSet.Fields.Item("LineTotal").Value) * (CDbl(oRecordSet1.Fields.Item("u_rate").Value) / 100))
                                    oRecordSet2.DoQuery("Insert into TempExpReval Values('" & oRecordSet.Fields.Item("ItemCode").Value & "','" & oRecordSet.Fields.Item("Quantity").Value & "','" & (CDbl(oRecordSet.Fields.Item("VatSum").Value)) & "','" & oRecordSet.Fields.Item("WhsCode").Value & "','" & oRecordSet.Fields.Item("DocDate").Value & "','" & oRecordSet.Fields.Item("DocNum").Value & "','" + oRs.Fields.Item("u_glacct").Value + "','" & oApplication.Company.UserName & "','" & oRecordSet.Fields.Item("Price").Value & "')")
                                    oRecordSet.MoveNext()
                                Next
                                oRset.DoQuery("Select Sum(Taxvalue),ItemCode,Quantity,WhsCode,Date,DocNum,sum(price)'price' from TempExpReval Where UserName = '" & oApplication.Company.UserName & "'  Group By ItemCode,WhsCode,Date,DocNum,UserName,Quantity")

                                Dim oSR1 As SAPbobsCOM.MaterialRevaluation
                                Dim oSR_Lines1 As SAPbobsCOM.MaterialRevaluation_lines
                                oSR1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
                                oSR1.RevalType = "P"
                                'oSR1.Reference1 = oRecordSet.Fields.Item("DocNum").Value
                                oSR.Reference2 = oRset.Fields.Item("DocNum").Value
                                'oSR.Comments = "ItemCost is revaluated for exciseduty or Cst tax amount(Excise Duty+CST/Instock of item in above Warehouse)"
                                oSR.Comments = "ItemCost is revaluated for LandedCost (LandedCost/Instock of item in above Warehouse)"

                                ' Dim lSBObob As SAPbobsCOM.SBObob
                                'Dim lRecordset As SAPbobsCOM.Recordset
                                Dim ld_CurrentItemCost1 As Decimal
                                Dim ld_NewItemCost1 As Decimal

                                oSR1.DocDate = oRset.Fields.Item("Date").Value

                                ''*********************************
                                For I As Int16 = 0 To oRecordSet.RecordCount - 1
                                    'oRecordSet1.DoQuery("select price from MRV1 inner join OMRV on OMRV.docentry=MRV1.docentry where Itemcode='" + oRecordSet.Fields.Item("ItemCode").Value + "' and DocNum=(Select MAX(docnum)from omrv)")
                                    oRecordSet1.DoQuery("Select AvgPrice,OnHand  From OITW where ItemCode = '" & oRset.Fields.Item("ItemCode").Value & "' And WhsCode = '" & oRset.Fields.Item("WhsCode").Value & "'")
                                    If oRecordSet1.RecordCount > 0 Then
                                        ld_CurrentItemCost1 = Val(oRecordSet1.Fields.Item("AvgPrice").Value)
                                        ld_NewItemCost1 = Math.Round(CDbl((oRecordSet1.Fields.Item("AvgPrice").Value) + ((((oRs.Fields.Item(0).Value / 100) * (oRset.Fields.Item("Price").Value)) * (oRset.Fields.Item("Quantity").Value)) / oRecordSet1.Fields.Item(1).Value)), 4) ' / oRset.Fields.Item(2).Value)))
                                    Else
                                        ld_CurrentItemCost1 = 0
                                        ld_NewItemCost1 = 0
                                    End If
                                    If ld_CurrentItemCost1 <> ld_NewItemCost1 Then
                                        oSR_Lines1 = oSR.Lines
                                        oSR_Lines1.SetCurrentLine(I)
                                        oSR_Lines1.ItemCode = oRset.Fields.Item("ItemCode").Value
                                        oSR_Lines1.WarehouseCode = oRset.Fields.Item("WhsCode").Value

                                        oSR_Lines1.Price = ld_NewItemCost1
                                        oRecordSet2.DoQuery("select acctcode from OACT where FormatCode='" + oRs.Fields.Item(1).Value + "'")
                                        oSR_Lines1.RevaluationIncrementAccount = oRecordSet2.Fields.Item(0).Value
                                        oSR_Lines1.RevaluationDecrementAccount = oRecordSet2.Fields.Item(0).Value
                                        oSR_Lines1.Add()
                                    End If
                                    oRset.MoveNext()
                                Next
                                If oSR.Add <> 0 Then
                                    Dim Str As String = oCompany.GetLastErrorDescription
                                    oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription)
                                    ' Else
                                    oRecordSet2.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                                    GoTo A
                                End If
                                oRecordSet2.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                                oRs.MoveNext()
                                ' End If
                            Next
                            ' Dim Str As String = oCompany.GetLastErrorDescription
                            oApplication.StatusBar.SetText("Inventory Revaluation was successfully posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            oRecordSet.DoQuery("Update OPDN set u_reval='Yes' where DocEntry=(Select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "'")
                            Dim oRs3 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim oRs4 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRs3.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocEntry=(Select Max(DocEntry) from OPDN) ")
                            oRs4.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocEntry=(Select Max(DocEntry) from OPDN)  order by omrv.DocNum desc")
                            If oRs3.Fields.Item(0).Value = oRs4.Fields.Item(0).Value Then
                                oRecordSet.DoQuery("Update OPDN set u_revalno=('" & oRs3.Fields.Item(0).Value & "') where DocEntry=(Select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "'")
                            Else

                                oRecordSet.DoQuery("Update OPDN set u_revalno=('" & oRs3.Fields.Item(0).Value & "' +' To '+ '" & oRs4.Fields.Item(0).Value & "')where DocEntry=(Select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "' ")
                            End If
                            Dim _str_DocNum1 As String
                            Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRs1.DoQuery("select DocNum,u_unit from OPDN Where DocEntry=(Select Max(DocEntry) from opdn)")
                            _str_DocNum1 = oRs1.Fields.Item(0).Value
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("8").Specific.value = _str_DocNum1
                            objForm.Items.Item("cpc").Specific.value = oRs1.Fields.Item(1).Value
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Else

                            For J As Int16 = 0 To oRs.RecordCount - 1
                                oRecordSet.DoQuery("SELECT T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], SUm(T1.[Vatsum]+T1.[LineTotal])'VatSum', T1.[ItemCode],Sum(T1.[Quantity])'Quantity', T1.[WhsCode], T1.[TaxCode] FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry Where T0.DocEntry = (select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "' group by T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], T1.[ItemCode], T1.[WhsCode], T1.[TaxCode]")
                                'oRecordSet.DoQuery("SELECT T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode],  T1.[ItemCode],sum(T1.[Quantity])'Quantity', T1.[WhsCode], T1.[TaxCode],sum(T1.LineTotal)'LineTotal',(select sum(distinct(PDN4.TaxSum)) from PDN1 inner join PDN4 on pdn1.DocEntry=pdn4.DocEntry And PDN1.LineNum = PDN4.LineNum where pdn4.NonDdctPrc='100' and pdn1.ItemCode=T1.ItemCode and pdn1.DocEntry=(Select MAX(Docentry)from OPDN))'Vatsum' FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry Where T0.DocEntry = (select Max(DocEntry) from OPDN) group by T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode],  T1.[ItemCode], T1.[WhsCode], T1.[TaxCode]"))
                                oRecordSet.MoveFirst()


                                For I As Int16 = 0 To oRecordSet.RecordCount - 1
                                    '  oRecordSet1.DoQuery("Select u_rate,u_expact from [@GEN_EXP_TAX]")
                                    'Dim Expenses_Tax As Double
                                    ' Expenses_Tax = (CDbl(oRecordSet.Fields.Item("LineTotal").Value) * (CDbl(oRecordSet1.Fields.Item("u_rate").Value) / 100))
                                    oRecordSet2.DoQuery("Insert into TempExpReval Values('" & oRecordSet.Fields.Item("ItemCode").Value & "','" & oRecordSet.Fields.Item("Quantity").Value & "','" & (CDbl(oRecordSet.Fields.Item("VatSum").Value)) & "','" & oRecordSet.Fields.Item("WhsCode").Value & "','" & oRecordSet.Fields.Item("DocDate").Value & "','" & oRecordSet.Fields.Item("DocNum").Value & "','" + oRs.Fields.Item("u_glacct").Value + "','" & oApplication.Company.UserName & "','0.00')")
                                    oRecordSet.MoveNext()
                                Next
                                oRecordSet.DoQuery("Select Sum(Taxvalue),ItemCode,Quantity,WhsCode,Date,DocNum from TempExpReval Where UserName = '" & oApplication.Company.UserName & "'  Group By ItemCode,WhsCode,Date,DocNum,UserName,Quantity")

                                Dim oSR1 As SAPbobsCOM.MaterialRevaluation
                                Dim oSR_Lines1 As SAPbobsCOM.MaterialRevaluation_lines
                                oSR1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
                                oSR1.RevalType = "P"
                                oSR1.Reference2 = oRecordSet.Fields.Item("DocNum").Value
                                oSR.Comments = "ItemCost is revaluated for LandedCost (LandedCost/Instock of item in above Warehouse)"

                                ' Dim lSBObob As SAPbobsCOM.SBObob
                                'Dim lRecordset As SAPbobsCOM.Recordset
                                Dim ld_CurrentItemCost1 As Decimal
                                Dim ld_NewItemCost1 As Decimal

                                oSR1.DocDate = oRecordSet.Fields.Item("Date").Value

                                ''*********************************
                                For I As Int16 = 0 To oRecordSet.RecordCount - 1
                                    oRecordSet1.DoQuery("select price from MRV1 inner join OMRV on OMRV.docentry=MRV1.docentry where Itemcode='" + oRecordSet.Fields.Item("ItemCode").Value + "' and DocNum=(Select MAX(docnum)from omrv)")
                                    oRecordSet1.DoQuery("Select AvgPrice,OnHand From OITW where ItemCode = '" & oRecordSet.Fields.Item("ItemCode").Value & "' And WhsCode = '" & oRecordSet.Fields.Item("WhsCode").Value & "'")
                                    If oRecordSet1.RecordCount > 0 Then
                                        ld_CurrentItemCost1 = Val(oRecordSet1.Fields.Item("AvgPrice").Value)
                                        ld_NewItemCost1 = Math.Round(CDbl(oRecordSet1.Fields.Item("AvgPrice").Value) + ((oRs.Fields.Item(0).Value * oRecordSet.Fields.Item(2).Value) / oRecordSet1.Fields.Item(1).Value), 4)

                                    Else
                                        ld_CurrentItemCost1 = 0
                                        ld_NewItemCost1 = 0
                                    End If
                                    If ld_CurrentItemCost1 <> ld_NewItemCost1 Then
                                        oSR_Lines1 = oSR.Lines
                                        oSR_Lines1.SetCurrentLine(I)
                                        oSR_Lines1.ItemCode = oRecordSet.Fields.Item("ItemCode").Value
                                        oSR_Lines1.WarehouseCode = oRecordSet.Fields.Item("WhsCode").Value

                                        oSR_Lines1.Price = ld_NewItemCost1
                                        oRecordSet2.DoQuery("select acctcode from OACT where FormatCode='" + oRs.Fields.Item(1).Value + "'")
                                        oSR_Lines1.RevaluationIncrementAccount = oRecordSet2.Fields.Item(0).Value
                                        'oSR_Lines1.RevaluationDecrementAccount = "401015-01"
                                        oSR_Lines1.RevaluationDecrementAccount = oRecordSet2.Fields.Item(0).Value
                                        oSR_Lines1.Add()
                                    End If
                                    oRecordSet.MoveNext()
                                Next
                                If oSR.Add <> 0 Then
                                    Dim Str As String = oCompany.GetLastErrorDescription
                                    oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription)
                                    ' Else
                                    oRecordSet2.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                                    'where DocEntry=(Select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "'
                                    oRecordSet.DoQuery("Update OPDN set u_reval='Partial' where DocEntry=(Select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "'")
                                    Dim oRsp3 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim oRsp4 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsp3.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocEntry=(Select Max(DocEntry) from OPDN) ")
                                    oRsp4.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocEntry=(Select Max(DocEntry) from OPDN)  order by omrv.DocNum desc")
                                    If oRsp3.Fields.Item(0).Value = oRsp4.Fields.Item(0).Value Then
                                        oRecordSet.DoQuery("Update OPDN set u_revalno='" & oRsp3.Fields.Item(0).Value & "' where DocEntry=(Select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "'")
                                    Else

                                        oRecordSet.DoQuery("Update OPDN set u_revalno=('" & oRsp3.Fields.Item(0).Value & "' +' To '+ '" & oRsp4.Fields.Item(0).Value & "') where DocEntry=(Select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "' ")
                                    End If
                                    Dim _str_DocNum1 As String
                                    Dim oRsp1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsp1.DoQuery("select Docnum,u_unit from OPDN Where DocEntry=(Select Max(DocEntry) from OPDN)")
                                    _str_DocNum1 = oRsp1.Fields.Item(0).Value
                                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    objForm.Items.Item("8").Specific.value = _str_DocNum1
                                    objForm.Items.Item("cpc").Specific.value = oRsp1.Fields.Item(1).Value
                                    objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    GoTo D
                                End If
                                oRecordSet2.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                                oRs.MoveNext()
                            Next
                            oApplication.StatusBar.SetText("Inventory Revaluation Was successfully Posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            oRecordSet.DoQuery("Update OPDN set u_reval='Yes' where DocEntry=(Select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "'")
                            Dim oRs3 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim oRs4 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRs3.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocEntry=(Select Max(DocEntry) from OPDN) ")
                            oRs4.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocEntry=(Select Max(DocEntry) from OPDN)  order by omrv.DocNum desc")
                            If oRs3.Fields.Item(0).Value = oRs4.Fields.Item(0).Value Then
                                oRecordSet.DoQuery("Update OPDN set u_revalno='" & oRs3.Fields.Item(0).Value & "' where DocEntry=(Select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "'")
                            Else
                                oRecordSet.DoQuery("Update OPDN set u_revalno=('" & oRs3.Fields.Item(0).Value & "' +' To '+ '" & oRs4.Fields.Item(0).Value & "') where DocEntry=(Select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "' ")
                            End If
                            Dim _str_DocNum As String
                            Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRs1.DoQuery("select Docnum,u_unit from OPDN Where DocEntry=(Select Max(DocEntry) from OPDN)")
                            _str_DocNum = oRs1.Fields.Item(0).Value
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("8").Specific.value = _str_DocNum
                            objForm.Items.Item("cpc").Specific.value = oRs1.Fields.Item(1).Value
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If

                    Else
                        oApplication.StatusBar.SetText("Inventory Revaluation Was Posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        oRecordSet.DoQuery("Update OPDN set u_reval='NO-LC' where DocEntry=(Select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "'")
                        Dim oRs3 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRs4 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRs3.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocEntry=(Select Max(DocEntry) from OPDN) ")
                        oRs4.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocEntry=(Select Max(DocEntry) from OPDN)  order by omrv.DocNum desc")
                        If oRs3.Fields.Item(0).Value = oRs4.Fields.Item(0).Value Then
                            oRecordSet.DoQuery("Update OPDN set u_revalno='" & oRs3.Fields.Item(0).Value & "' where DocEntry=(Select Max(DocEntry) from OPDN) And u_macid='" + MAC_ID + "'")
                        Else
                            oRecordSet.DoQuery("Update OPDN set u_revalno=('" & oRs3.Fields.Item(0).Value & "' +' To '+ '" & oRs4.Fields.Item(0).Value & "') where DocEntry=(Select Max(DocEntry) from OPDN) u_macid='" + MAC_ID + "'  ")
                        End If

                        Dim _str_DocNum As String
                        Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRs1.DoQuery("select Docnum,u_unit from OPDN Where DocEntry=(Select Max(DocEntry) from OPDN)")
                        _str_DocNum = oRs1.Fields.Item(0).Value
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objForm.Items.Item("8").Specific.value = _str_DocNum
                        objForm.Items.Item("cpc").Specific.value = oRs1.Fields.Item(1).Value
                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
D:
                End If
                objForm.Refresh()
                ' If oSR.Add = 0 Then

                ' End If

                oRecordSet1.DoQuery("delete from TempReval where UserName = '" & oApplication.Company.UserName & "'")
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
        '  End If
    End Sub

    Sub InventoryRevaluation_Manual()
        Try
            'Declarations
            Dim oSR As SAPbobsCOM.MaterialRevaluation
            'oForm = oApplication.Forms.Item("143")
            'Dim CompanyService As SAPbobsCOM.CompanyService
            Dim oSR_Lines As SAPbobsCOM.MaterialRevaluation_lines
            'Dim RS As SAPbobsCOM.Recordset

            oSR = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
            Dim oMr As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRecordSet1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRecordSet2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oMr.DoQuery("Select count(Docnum) From OMRV Where OMRV.Ref2='" + objForm.Items.Item("8").Specific.value + "'")
            If oMr.Fields.Item(0).Value = 0 Then


                'CompanyService = oCompany.GetCompanyService
                'RS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                '  '  oRecordSet.DoQuery("SELECT T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], T1.[LineNum], T1.[ItemCode],T1.[Quantity], T1.[WhsCode], T1.[TaxCode] FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry Where T0.DocEntry = (select Max(DocEntry) from OPDN)")
                oRecordSet.DoQuery("SELECT T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode],  T1.[ItemCode],sum(T1.[Quantity])'Quantity', T1.[WhsCode], T1.[TaxCode],sum(T1.LineTotal)'LineTotal',(select sum(distinct(PDN4.TaxSum)) from PDN1 inner join PDN4 on pdn1.DocEntry=pdn4.DocEntry And PDN1.LineNum = PDN4.LineNum where pdn4.NonDdctPrc='100' and pdn1.ItemCode=T1.ItemCode and pdn1.DocEntry='" + GRPODOCENTRY + "')'Vatsum' FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry Where T0.DocNum='" + GRPODOCNUM + "' group by T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode],  T1.[ItemCode], T1.[WhsCode], T1.[TaxCode]")
                oRecordSet.MoveFirst()
                oRecordSet1.DoQuery("delete from TempReval where UserName = '" & oCompany.UserName & "' ")
                For I As Int16 = 0 To oRecordSet.RecordCount - 1
                    '  ' oRecordSet1.DoQuery("SELECT T2.StaCode,T2.TaxSum,T2.LineNum,T1.ItemCode FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry Inner join PDN4 T2 ON T1.DocEntry = T2.DocEntry And T1.LineNum = T2.LineNum Where T0.DocEntry = (select Max(DocEntry) from OPDN) And T1.LineNum = '" & oRecordSet.Fields.Item("LineNum").Value & "' And T2.ExpnsCode = -1 And T2.NonDdctPrc='100'")
                    '  ' For J As Int16 = 0 To oRecordSet1.RecordCount - 1
                    ' oRecordSet2.DoQuery("Select U_Allow From [@TAXREV] Where Code = '" & oRecordSet1.Fields.Item("StaCode").Value & "'")
                    ' If oRecordSet2.Fields.Item("U_Allow").Value = "YES" Then
                    oRecordSet2.DoQuery("Insert into TempReval Values('" & oRecordSet.Fields.Item("ItemCode").Value.ToString & "','" & oRecordSet.Fields.Item("Quantity").Value & "','" & (oRecordSet.Fields.Item("Vatsum").Value) & "','" & oRecordSet.Fields.Item("WhsCode").Value.ToString & "','" & oRecordSet.Fields.Item("DocDate").Value & "','" & oRecordSet.Fields.Item("DocNum").Value.ToString & "','" & oApplication.Company.UserName & "')")
                    'End If
                    '  'oRecordSet1.MoveNext()
                    '  ' Next
                    oRecordSet.MoveNext()
                Next
                oRecordSet.DoQuery("Select Sum(Taxvalue),ItemCode,Quantity,WhsCode,Date,DocNum from TempReval Where UserName = '" & oApplication.Company.UserName & "'  Group By ItemCode,WhsCode,Date,DocNum,UserName,Quantity")
                oSR.RevalType = "P"
                oSR.Reference2 = oRecordSet.Fields.Item("DocNum").Value

                ''****line Added by vivek for curecting date conversion***********

                Dim lSBObob As SAPbobsCOM.SBObob
                Dim lRecordset As SAPbobsCOM.Recordset
                Dim ld_CurrentItemCost As Decimal
                Dim ld_NewItemCost As Decimal

                ''lSBObob = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
                ''lRecordset = lSBObob.Format_StringToDate(oRecordSet.Fields.Item("Date").Value)

                oSR.DocDate = oRecordSet.Fields.Item("Date").Value
                ''oSR.DocDate = lRecordset.Fields.Item(0).Value

                ''*********************************

                For I As Int16 = 0 To oRecordSet.RecordCount - 1
                    oRecordSet1.DoQuery("Select AvgPrice,OnHand  From OITW where ItemCode = '" & oRecordSet.Fields.Item("ItemCode").Value & "' And WhsCode = '" & oRecordSet.Fields.Item("WhsCode").Value & "'")


                    If oRecordSet1.RecordCount > 0 Then
                        ld_CurrentItemCost = Val(oRecordSet1.Fields.Item("AvgPrice").Value)
                        ld_NewItemCost = Math.Round(CDbl(oRecordSet1.Fields.Item("AvgPrice").Value) + (CDbl(oRecordSet.Fields.Item(0).Value) / CDbl(oRecordSet1.Fields.Item(1).Value)), 4)

                    Else
                        ld_CurrentItemCost = 0
                        ld_NewItemCost = 0
                    End If

                    ''oApplication.MessageBox("current cost : " + CStr(ld_CurrentItemCost) + "new cost : " + CStr(ld_NewItemCost))
                    If ld_CurrentItemCost <> ld_NewItemCost Then
                        oSR_Lines = oSR.Lines
                        oSR_Lines.SetCurrentLine(I)
                        oSR_Lines.ItemCode = oRecordSet.Fields.Item("ItemCode").Value
                        oSR_Lines.WarehouseCode = oRecordSet.Fields.Item("WhsCode").Value

                        oSR_Lines.Price = ld_NewItemCost
                        ' oSR_Lines.RevaluationIncrementAccount = "401020"
                        'oSR_Lines.RevaluationDecrementAccount = "401020"
                        ' oForm.Items.Item()
                        oSR_Lines.Add()
                    End If
                    oRecordSet.MoveNext()
                Next
                If oSR.Add <> 0 Then
                    Dim oManual As SAPbobsCOM.Recordset
                    oManual = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oManual.DoQuery("select count(stacode) from PDN4 inner join OPDN on opdn.DocEntry=pdn4.DocEntry where OPDN.DocEntry=(Select Max(Docentry) from opdn) and NonDdctPrc='100'") ''Where U_unit = 'Unit1'
                    If oManual.Fields.Item(0).Value = 0 Then


                        oRecordSet1.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                        Dim oManual1 As SAPbobsCOM.Recordset
                        oManual1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oManual1.DoQuery("Select T1.u_rate,T1.u_glacct,T1.u_value from [@GEN_GRPO_LCOSTS] T0 Inner Join [@GEN_GRPO_LCOSTS_D0] T1 on T0.Code=T1.Code Where T0.u_grpono='" + GRPODOCNUM + "' And T1.u_glacct<>''")
                        Dim oManual2 As SAPbobsCOM.Recordset
                        oManual2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oManual2.DoQuery("Select count(T1.u_rate) from [@GEN_GRPO_LCOSTS] T0 Inner Join [@GEN_GRPO_LCOSTS_D0] T1 on T0.Code=T1.Code Where T0.u_grpono='" + GRPODOCNUM + "' And T1.u_glacct<>''")
                        If oManual2.Fields.Item(0).Value > 0 Then
                            If oManual1.Fields.Item(0).Value > "0.00" Then


                                Dim LCPer As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                LCPer.DoQuery("Select U_type from OALC Where OALC.AlcName=(Select u_lname from [@GEN_GRPO_LCOSTS] T0 inner join [@GEN_GRPO_LCOSTS_D0] T1 on T0.Code=T1.Code Where T0.u_grpono='" + GRPODOCNUM + "' and U_lname<> '')")
                                If LCPer.Fields.Item(0).Value = "Percent" Then
                                    For J As Int16 = 0 To oManual1.RecordCount - 1
                                        oRecordSet.DoQuery("SELECT T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], SUm(T1.[Vatsum]+T1.[LineTotal])'VatSum', T1.[ItemCode],Sum(T1.[Quantity])'Quantity',sum(T1.[Price])'price', T1.[WhsCode], T1.[TaxCode] FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry Where T0.DocNum='" + GRPODOCNUM + "' group by T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], T1.[ItemCode], T1.[WhsCode], T1.[TaxCode]")
                                        oRecordSet.MoveFirst()

                                        For I As Int16 = 0 To oRecordSet.RecordCount - 1
                                            ' oRecordSet1.DoQuery("Select u_rate,u_expact from [@GEN_EXP_TAX]")
                                            ' Dim Expenses_Tax As Double
                                            ' Expenses_Tax = (CDbl(oRecordSet.Fields.Item("LineTotal").Value) * (CDbl(oRecordSet1.Fields.Item("u_rate").Value) / 100))
                                            oRecordSet2.DoQuery("Insert into TempExpReval Values('" & oRecordSet.Fields.Item("ItemCode").Value & "','" & oRecordSet.Fields.Item("Quantity").Value & "','" & (CDbl(oRecordSet.Fields.Item("VatSum").Value)) & "','" & oRecordSet.Fields.Item("WhsCode").Value & "','" & oRecordSet.Fields.Item("DocDate").Value & "','" & oRecordSet.Fields.Item("DocNum").Value & "','" + oManual1.Fields.Item("u_glacct").Value + "','" & oApplication.Company.UserName & "','" & oRecordSet.Fields.Item("Price").Value & "')")
                                            oRecordSet.MoveNext()
                                        Next
                                        oManual.DoQuery("Select Sum(Taxvalue),ItemCode,Quantity,WhsCode,Date,DocNum,Sum(price)'price' from TempExpReval Where UserName = '" & oApplication.Company.UserName & "'  Group By ItemCode,WhsCode,Date,DocNum,UserName,Quantity")

                                        Dim oSR1 As SAPbobsCOM.MaterialRevaluation
                                        Dim oSR_Lines1 As SAPbobsCOM.MaterialRevaluation_lines
                                        oSR1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
                                        oSR1.RevalType = "P"
                                        'oSR1.Reference1 = oRecordSet.Fields.Item("DocNum").Value
                                        oSR.Reference2 = oManual.Fields.Item("DocNum").Value

                                        ' Dim lSBObob As SAPbobsCOM.SBObob
                                        'Dim lRecordset As SAPbobsCOM.Recordset
                                        Dim ld_CurrentItemCost1 As Decimal
                                        Dim ld_NewItemCost1 As Decimal

                                        oSR1.DocDate = oManual.Fields.Item("Date").Value

                                        ''*********************************
                                        For I As Int16 = 0 To oManual.RecordCount - 1
                                            'oRecordSet1.DoQuery("select price from MRV1 inner join OMRV on OMRV.docentry=MRV1.docentry where Itemcode='" + oRecordSet.Fields.Item("ItemCode").Value + "' and DocNum=(Select MAX(docnum)from omrv)")
                                            oRecordSet1.DoQuery("Select AvgPrice,OnHand From OITW where ItemCode = '" & oManual.Fields.Item("ItemCode").Value & "' And WhsCode = '" & oManual.Fields.Item("WhsCode").Value & "'")
                                            If oRecordSet1.RecordCount > 0 Then
                                                ld_CurrentItemCost1 = Val(oRecordSet1.Fields.Item("AvgPrice").Value)
                                                ld_NewItemCost1 = Math.Round(CDbl((oRecordSet1.Fields.Item("AvgPrice").Value) + ((((oManual1.Fields.Item(0).Value / 100) * (oManual.Fields.Item("Price").Value)) * (oManual.Fields.Item("Quantity").Value)) / oRecordSet1.Fields.Item(1).Value)), 4)
                                            Else
                                                ld_CurrentItemCost1 = 0
                                                ld_NewItemCost1 = 0
                                            End If
                                            If ld_CurrentItemCost1 <> ld_NewItemCost1 Then
                                                oSR_Lines1 = oSR.Lines
                                                oSR_Lines1.SetCurrentLine(I)
                                                oSR_Lines1.ItemCode = oManual.Fields.Item("ItemCode").Value
                                                oSR_Lines1.WarehouseCode = oManual.Fields.Item("WhsCode").Value

                                                oSR_Lines1.Price = ld_NewItemCost1
                                                oRecordSet2.DoQuery("select acctcode from OACT where FormatCode='" + oManual1.Fields.Item(1).Value + "'")
                                                oSR_Lines1.RevaluationIncrementAccount = oRecordSet2.Fields.Item(0).Value
                                                oSR_Lines1.RevaluationDecrementAccount = oRecordSet2.Fields.Item(0).Value
                                                oSR_Lines1.Add()
                                            End If
                                            oManual.MoveNext()
                                        Next
                                        If oSR.Add <> 0 Then
                                            Dim Str As String = oCompany.GetLastErrorDescription
                                            oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription)
                                            ' Else
                                            oRecordSet2.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                                            GoTo A
                                        End If
                                        oRecordSet2.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                                        oManual1.MoveNext()
                                        ' End If
                                    Next
                                    ' Dim Str As String = oCompany.GetLastErrorDescription
                                    oApplication.StatusBar.SetText("Inventory Revaluation was successfully posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    oRecordSet.DoQuery("Update OPDN set u_reval='Yes' where DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                                    Dim oRs3 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim oRs4 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRs3.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2  where OPDN.DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                                    oRs4.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocNum='" + objForm.Items.Item("8").Specific.value + "' order by omrv.DocNum desc")
                                    If oRs3.Fields.Item(0).Value = oRs4.Fields.Item(0).Value Then
                                        oRecordSet.DoQuery("Update OPDN set u_revalno=('" & oRs3.Fields.Item(0).Value & "')where DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                                    Else
                                        oRecordSet.DoQuery("Update OPDN set u_revalno=('" & oRs3.Fields.Item(0).Value & "' +' To '+ '" & oRs4.Fields.Item(0).Value & "') where DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                                    End If
                                    Dim _str_DocNum1 As String
                                    Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRs1.DoQuery("select DocNum,u_unit from OPDN Where DocEntry=(Select Max(DocEntry) from opdn)")
                                    _str_DocNum1 = oRs1.Fields.Item(0).Value
                                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    objForm.Items.Item("8").Specific.value = _str_DocNum1
                                    objForm.Items.Item("cpc").Specific.value = oRs1.Fields.Item(1).Value
                                    objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Else

                                    For J As Int16 = 0 To oManual1.RecordCount - 1
                                        oRecordSet.DoQuery("SELECT T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], SUm(T1.[Vatsum]+T1.[LineTotal])'VatSum', T1.[ItemCode],Sum(T1.[Quantity])'Quantity', T1.[WhsCode], T1.[TaxCode] FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry Where T0.DocNum='" + GRPODOCNUM + "' group by T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], T1.[ItemCode], T1.[WhsCode], T1.[TaxCode]")
                                        oRecordSet.MoveFirst()


                                        For I As Int16 = 0 To oRecordSet.RecordCount - 1
                                            '  oRecordSet1.DoQuery("Select u_rate,u_expact from [@GEN_EXP_TAX]")
                                            'Dim Expenses_Tax As Double
                                            ' Expenses_Tax = (CDbl(oRecordSet.Fields.Item("LineTotal").Value) * (CDbl(oRecordSet1.Fields.Item("u_rate").Value) / 100))
                                            oRecordSet2.DoQuery("Insert into TempExpReval Values('" & oRecordSet.Fields.Item("ItemCode").Value & "','" & oRecordSet.Fields.Item("Quantity").Value & "','" & (CDbl(oRecordSet.Fields.Item("VatSum").Value)) & "','" & oRecordSet.Fields.Item("WhsCode").Value & "','" & oRecordSet.Fields.Item("DocDate").Value & "','" & oRecordSet.Fields.Item("DocNum").Value & "','" + oManual1.Fields.Item("u_glacct").Value + "','" & oApplication.Company.UserName & "','0.00')")
                                            oRecordSet.MoveNext()
                                        Next
                                        oManual.DoQuery("Select Sum(Taxvalue),ItemCode,Quantity,WhsCode,Date,DocNum from TempExpReval Where UserName = '" & oApplication.Company.UserName & "'  Group By ItemCode,WhsCode,Date,DocNum,UserName,Quantity")

                                        Dim oSR1 As SAPbobsCOM.MaterialRevaluation
                                        Dim oSR_Lines1 As SAPbobsCOM.MaterialRevaluation_lines
                                        oSR1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
                                        oSR1.RevalType = "P"
                                        'oSR1.Reference1 = oRecordSet.Fields.Item("DocNum").Value
                                        oSR.Reference2 = oManual.Fields.Item("DocNum").Value

                                        ' Dim lSBObob As SAPbobsCOM.SBObob
                                        'Dim lRecordset As SAPbobsCOM.Recordset
                                        Dim ld_CurrentItemCost1 As Decimal
                                        Dim ld_NewItemCost1 As Decimal

                                        oSR1.DocDate = oManual.Fields.Item("Date").Value
                                        'oSR1.TaxDate = oManual.Fields.Item("Date").Value

                                        ''*********************************
                                        For I As Int16 = 0 To oRecordSet.RecordCount - 1
                                            'oRecordSet1.DoQuery("select price from MRV1 inner join OMRV on OMRV.docentry=MRV1.docentry where Itemcode='" + oRecordSet.Fields.Item("ItemCode").Value + "' and DocNum=(Select MAX(docnum)from omrv)")
                                            oRecordSet1.DoQuery("Select AvgPrice,OnHand  From OITW where ItemCode = '" & oManual.Fields.Item("ItemCode").Value & "' And WhsCode = '" & oManual.Fields.Item("WhsCode").Value & "'")
                                            If oRecordSet1.RecordCount > 0 Then
                                                ld_CurrentItemCost1 = Val(oRecordSet1.Fields.Item("AvgPrice").Value)
                                                ld_NewItemCost1 = Math.Round(CDbl(oRecordSet1.Fields.Item("AvgPrice").Value) + ((oManual1.Fields.Item(0).Value * oManual.Fields.Item(2).Value) / oRecordSet1.Fields.Item(1).Value), 4)

                                            Else
                                                ld_CurrentItemCost1 = 0
                                                ld_NewItemCost1 = 0
                                            End If
                                            If ld_CurrentItemCost1 <> ld_NewItemCost1 Then
                                                oSR_Lines1 = oSR.Lines
                                                oSR_Lines1.SetCurrentLine(I)
                                                oSR_Lines1.ItemCode = oManual.Fields.Item("ItemCode").Value
                                                oSR_Lines1.WarehouseCode = oManual.Fields.Item("WhsCode").Value

                                                oSR_Lines1.Price = ld_NewItemCost1
                                                oRecordSet2.DoQuery("select acctcode from OACT where FormatCode='" + oManual1.Fields.Item(1).Value + "'")
                                                oSR_Lines1.RevaluationIncrementAccount = oRecordSet2.Fields.Item(0).Value
                                                oSR_Lines1.RevaluationDecrementAccount = oRecordSet2.Fields.Item(0).Value
                                                oSR_Lines1.Add()
                                            End If
                                            oManual.MoveNext()
                                        Next
                                        If oSR.Add <> 0 Then
                                            Dim Str As String = oCompany.GetLastErrorDescription
                                            oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription)
                                            ' Else
                                            oRecordSet2.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                                            GoTo A

                                        End If
                                        oRecordSet2.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                                        oManual1.MoveNext()
                                        ' End If
                                    Next
                                    ' Dim Str As String = oCompany.GetLastErrorDescription
                                    oApplication.StatusBar.SetText("Inventory Revaluation was successfully posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    oRecordSet.DoQuery("Update OPDN set u_reval='Yes' where DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                                    Dim oRs3 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim oRs4 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRs3.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                                    oRs4.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocNum='" + objForm.Items.Item("8").Specific.value + "' order by omrv.DocNum desc")
                                    If oRs3.Fields.Item(0).Value = oRs4.Fields.Item(0).Value Then
                                        oRecordSet.DoQuery("Update OPDN set u_revalno=('" & oRs3.Fields.Item(0).Value & "')where DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                                    Else

                                        oRecordSet.DoQuery("Update OPDN set u_revalno=('" & oRs3.Fields.Item(0).Value & "' +' To '+ '" & oRs4.Fields.Item(0).Value & "')where DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                                    End If
                                    Dim _str_DocNum1 As String
                                    Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRs1.DoQuery("select DocNum,u_unit from OPDN Where DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                                    _str_DocNum1 = oRs1.Fields.Item(0).Value
                                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    objForm.Items.Item("8").Specific.value = _str_DocNum1
                                    objForm.Items.Item("cpc").Specific.value = oRs1.Fields.Item(1).Value
                                    objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                End If
                            Else
                                oRecordSet.DoQuery("Update OPDN set u_reval='NO-LC' where DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                                Dim _str_DocNum1 As String
                                Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRs1.DoQuery("select DocNum,u_unit from OPDN Where DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                                _str_DocNum1 = oRs1.Fields.Item(0).Value
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                objForm.Items.Item("8").Specific.value = _str_DocNum1
                                objForm.Items.Item("cpc").Specific.value = oRs1.Fields.Item(1).Value
                                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If

                        Else
                            'oRecordSet.DoQuery("Update OPDN set u_reval='NA' where DocEntry=(Select Max(DocEntry) from OPDN)")
                            'Dim _str_DocNum1 As String
                            'Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRs1.DoQuery("select Docnum,U_Unit from OPDN where docentry=(Select Max(DocEntry) from OPDN)")
                            '_str_DocNum1 = oRs1.Fields.Item(0).Value
                            'objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            'objForm.Items.Item("8").Specific.value = _str_DocNum1
                            'objForm.Items.Item("cpc").Specific.value = oRs1.Fields.Item(1).Value
                            'objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oApplication.StatusBar.SetText("Select LandedCost by Clicking LC Button", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
A:
                        'End If
                    Else
                        Dim Str As String = oCompany.GetLastErrorDescription
                        oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription)
                    End If
                    'End If
                Else

                    oRecordSet1.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                    Dim oRs As SAPbobsCOM.Recordset
                    oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRs.DoQuery("Select T1.u_rate,T1.u_glacct,T1.u_value from [@GEN_GRPO_LCOSTS] T0 Inner Join [@GEN_GRPO_LCOSTS_D0] T1 on T0.Code=T1.Code Where T0.u_grpono='" + objForm.Items.Item("8").Specific.value + "' And T1.u_glacct<>''")
                    Dim oRs2 As SAPbobsCOM.Recordset
                    Dim oRset As SAPbobsCOM.Recordset
                    oRs2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRs2.DoQuery("Select count(T1.u_rate) from [@GEN_GRPO_LCOSTS] T0 Inner Join [@GEN_GRPO_LCOSTS_D0] T1 on T0.Code=T1.Code Where T0.u_grpono='" + GRPODOCNUM + "' And T1.u_glacct<>''")
                    If oRs2.Fields.Item(0).Value > 0 Then
                        Dim LCPer As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        LCPer.DoQuery("Select U_type from OALC Where OALC.AlcName=(Select u_lname from [@GEN_GRPO_LCOSTS] T0 inner join [@GEN_GRPO_LCOSTS_D0] T1 on T0.Code=T1.Code Where T0.u_grpono='" + GRPODOCNUM + "' and U_lname<> '')")
                        If Trim(LCPer.Fields.Item(0).Value) = "Percent" Then
                            For J As Int16 = 0 To oRs.RecordCount - 1
                                oRecordSet.DoQuery("SELECT T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], SUm(T1.[Vatsum]+T1.[LineTotal])'VatSum', T1.[ItemCode],Sum(T1.[Quantity])'Quantity',sum(T1.[Price])'price', T1.[WhsCode], T1.[TaxCode] FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry Where T0.DocEntry = '" + GRPODOCENTRY + "' group by T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], T1.[ItemCode], T1.[WhsCode], T1.[TaxCode]")
                                oRecordSet.MoveFirst()

                                For I As Int16 = 0 To oRecordSet.RecordCount - 1
                                    '  oRecordSet1.DoQuery("Select u_rate,u_expact from [@GEN_EXP_TAX]")
                                    'Dim Expenses_Tax As Double
                                    ' Expenses_Tax = (CDbl(oRecordSet.Fields.Item("LineTotal").Value) * (CDbl(oRecordSet1.Fields.Item("u_rate").Value) / 100))
                                    oRecordSet2.DoQuery("Insert into TempExpReval Values('" & oRecordSet.Fields.Item("ItemCode").Value & "','" & oRecordSet.Fields.Item("Quantity").Value & "','" & (CDbl(oRecordSet.Fields.Item("VatSum").Value)) & "','" & oRecordSet.Fields.Item("WhsCode").Value & "','" & oRecordSet.Fields.Item("DocDate").Value & "','" & oRecordSet.Fields.Item("DocNum").Value & "','" + oRs.Fields.Item("u_glacct").Value + "','" & oApplication.Company.UserName & "','" & oRecordSet.Fields.Item("Price").Value & "')")
                                    oRecordSet.MoveNext()
                                Next
                                oRset.DoQuery("Select Sum(Taxvalue),ItemCode,Quantity,WhsCode,Date,DocNum,Sum(price)'price' from TempExpReval Where UserName = '" & oApplication.Company.UserName & "'  Group By ItemCode,WhsCode,Date,DocNum,UserName,Quantity")

                                Dim oSR1 As SAPbobsCOM.MaterialRevaluation
                                Dim oSR_Lines1 As SAPbobsCOM.MaterialRevaluation_lines
                                oSR1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
                                oSR1.RevalType = "P"
                                'oSR1.Reference1 = oRecordSet.Fields.Item("DocNum").Value
                                oSR.Reference2 = oRset.Fields.Item("DocNum").Value

                                ' Dim lSBObob As SAPbobsCOM.SBObob
                                'Dim lRecordset As SAPbobsCOM.Recordset
                                Dim ld_CurrentItemCost1 As Decimal
                                Dim ld_NewItemCost1 As Decimal

                                oSR1.DocDate = oRset.Fields.Item("Date").Value

                                ''*********************************
                                For I As Int16 = 0 To oRset.RecordCount - 1
                                    'oRecordSet1.DoQuery("select price from MRV1 inner join OMRV on OMRV.docentry=MRV1.docentry where Itemcode='" + oRecordSet.Fields.Item("ItemCode").Value + "' and DocNum=(Select MAX(docnum)from omrv)")
                                    oRecordSet1.DoQuery("Select AvgPrice,OnHand From OITW where ItemCode = '" & oRset.Fields.Item("ItemCode").Value & "' And WhsCode = '" & oRset.Fields.Item("WhsCode").Value & "'")
                                    If oRecordSet1.RecordCount > 0 Then
                                        ld_CurrentItemCost1 = Val(oRecordSet1.Fields.Item("AvgPrice").Value)
                                        ld_NewItemCost1 = Math.Round(CDbl((oRecordSet1.Fields.Item("AvgPrice").Value) + ((((oRs.Fields.Item(0).Value / 100) * (oRset.Fields.Item("Price").Value)) * (oRset.Fields.Item("Quantity").Value)) / oRecordSet1.Fields.Item(1).Value)), 4)
                                    Else
                                        ld_CurrentItemCost1 = 0
                                        ld_NewItemCost1 = 0
                                    End If
                                    If ld_CurrentItemCost1 <> ld_NewItemCost1 Then
                                        oSR_Lines1 = oSR.Lines
                                        oSR_Lines1.SetCurrentLine(I)
                                        oSR_Lines1.ItemCode = oRset.Fields.Item("ItemCode").Value
                                        oSR_Lines1.WarehouseCode = oRset.Fields.Item("WhsCode").Value

                                        oSR_Lines1.Price = ld_NewItemCost1
                                        oRecordSet2.DoQuery("select acctcode from OACT where FormatCode='" + oRs.Fields.Item(1).Value + "'")
                                        oSR_Lines1.RevaluationIncrementAccount = oRecordSet2.Fields.Item(0).Value
                                        oSR_Lines1.RevaluationDecrementAccount = oRecordSet2.Fields.Item(0).Value
                                        oSR_Lines1.Add()
                                    End If
                                    oRset.MoveNext()
                                Next
                                If oSR.Add <> 0 Then
                                    Dim Str As String = oCompany.GetLastErrorDescription
                                    oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription)
                                    ' Else
                                    oRecordSet2.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                                    GoTo A
                                End If
                                oRecordSet2.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                                oRs.MoveNext()
                                ' End If
                            Next
                            ' Dim Str As String = oCompany.GetLastErrorDescription
                            oApplication.StatusBar.SetText("Inventory Revaluation Was successfully Posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            oRecordSet.DoQuery("Update OPDN set u_reval='Yes' where DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                            Dim oRs3 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim oRs4 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRs3.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                            oRs4.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocNum='" + objForm.Items.Item("8").Specific.value + "' order by omrv.DocNum desc")
                            If oRs3.Fields.Item(0).Value = oRs4.Fields.Item(0).Value Then
                                oRecordSet.DoQuery("Update OPDN set u_revalno='" & oRs3.Fields.Item(0).Value & "' where DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                            Else

                                oRecordSet.DoQuery("Update OPDN set u_revalno=('" & oRs3.Fields.Item(0).Value & "' +' To '+ '" & oRs4.Fields.Item(0).Value & "') where DocNum='" + objForm.Items.Item("8").Specific.value + "' ")
                            End If
                            Dim _str_DocNum As String
                            Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRs1.DoQuery("select Docnum,u_unit from OPDN Where DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                            _str_DocNum = oRs1.Fields.Item(0).Value
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("8").Specific.value = _str_DocNum
                            objForm.Items.Item("cpc").Specific.value = oRs1.Fields.Item(1).Value
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Else
                            For J As Int16 = 0 To oRs.RecordCount - 1
                                oRecordSet.DoQuery("SELECT T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], SUm(T1.[Vatsum]+T1.[LineTotal])'VatSum', T1.[ItemCode],Sum(T1.[Quantity])'Quantity', T1.[WhsCode], T1.[TaxCode] FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry Where T0.DocNum = '" + objForm.Items.Item("8").Specific.value + "' group by T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], T1.[ItemCode], T1.[WhsCode], T1.[TaxCode]")
                                'oRecordSet.DoQuery("SELECT T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode],  T1.[ItemCode],sum(T1.[Quantity])'Quantity', T1.[WhsCode], T1.[TaxCode],sum(T1.LineTotal)'LineTotal',(select sum(distinct(PDN4.TaxSum)) from PDN1 inner join PDN4 on pdn1.DocEntry=pdn4.DocEntry And PDN1.LineNum = PDN4.LineNum where pdn4.NonDdctPrc='100' and pdn1.ItemCode=T1.ItemCode and pdn1.DocEntry=(Select MAX(Docentry)from OPDN))'Vatsum' FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry Where T0.DocEntry = (select Max(DocEntry) from OPDN) group by T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode],  T1.[ItemCode], T1.[WhsCode], T1.[TaxCode]"))
                                oRecordSet.MoveFirst()


                                For I As Int16 = 0 To oRecordSet.RecordCount - 1
                                    '  oRecordSet1.DoQuery("Select u_rate,u_expact from [@GEN_EXP_TAX]")
                                    'Dim Expenses_Tax As Double
                                    ' Expenses_Tax = (CDbl(oRecordSet.Fields.Item("LineTotal").Value) * (CDbl(oRecordSet1.Fields.Item("u_rate").Value) / 100))
                                    oRecordSet2.DoQuery("Insert into TempExpReval Values('" & oRecordSet.Fields.Item("ItemCode").Value & "','" & oRecordSet.Fields.Item("Quantity").Value & "','" & (CDbl(oRecordSet.Fields.Item("VatSum").Value)) & "','" & oRecordSet.Fields.Item("WhsCode").Value & "','" & oRecordSet.Fields.Item("DocDate").Value & "','" & oRecordSet.Fields.Item("DocNum").Value & "','" + oRs.Fields.Item("u_glacct").Value + "','" & oApplication.Company.UserName & "','0.00')")
                                    oRecordSet.MoveNext()
                                Next
                                oRecordSet.DoQuery("Select Sum(Taxvalue),ItemCode,Quantity,WhsCode,Date,DocNum from TempExpReval Where UserName = '" & oApplication.Company.UserName & "'  Group By ItemCode,WhsCode,Date,DocNum,UserName,Quantity")

                                Dim oSR1 As SAPbobsCOM.MaterialRevaluation
                                Dim oSR_Lines1 As SAPbobsCOM.MaterialRevaluation_lines
                                oSR1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
                                oSR1.RevalType = "P"
                                oSR1.Reference2 = oRecordSet.Fields.Item("DocNum").Value

                                ' Dim lSBObob As SAPbobsCOM.SBObob
                                'Dim lRecordset As SAPbobsCOM.Recordset
                                Dim ld_CurrentItemCost1 As Decimal
                                Dim ld_NewItemCost1 As Decimal

                                oSR1.DocDate = oRecordSet.Fields.Item("Date").Value

                                ''*********************************
                                For I As Int16 = 0 To oRecordSet.RecordCount - 1
                                    oRecordSet1.DoQuery("select price from MRV1 inner join OMRV on OMRV.docentry=MRV1.docentry where Itemcode='" + oRecordSet.Fields.Item("ItemCode").Value + "' and DocNum=(Select MAX(docnum)from omrv)")
                                    oRecordSet1.DoQuery("Select AvgPrice,OnHand From OITW where ItemCode = '" & oRecordSet.Fields.Item("ItemCode").Value & "' And WhsCode = '" & oRecordSet.Fields.Item("WhsCode").Value & "'")

                                    If oRecordSet1.RecordCount > 0 Then
                                        ld_CurrentItemCost1 = Val(oRecordSet1.Fields.Item("AvgPrice").Value)
                                        ld_NewItemCost1 = Math.Round(CDbl(oRecordSet1.Fields.Item("AvgPrice").Value) + ((oRs.Fields.Item(0).Value * oRecordSet.Fields.Item(2).Value) / oRecordSet1.Fields.Item(1).Value), 4)

                                    Else
                                        ld_CurrentItemCost1 = 0
                                        ld_NewItemCost1 = 0
                                    End If
                                    If ld_CurrentItemCost1 <> ld_NewItemCost1 Then
                                        oSR_Lines1 = oSR.Lines
                                        oSR_Lines1.SetCurrentLine(I)
                                        oSR_Lines1.ItemCode = oRecordSet.Fields.Item("ItemCode").Value
                                        oSR_Lines1.WarehouseCode = oRecordSet.Fields.Item("WhsCode").Value

                                        oSR_Lines1.Price = ld_NewItemCost1
                                        oRecordSet2.DoQuery("select acctcode from OACT where FormatCode='" + oRs.Fields.Item(1).Value + "'")
                                        oSR_Lines1.RevaluationIncrementAccount = oRecordSet2.Fields.Item(0).Value
                                        oSR_Lines1.RevaluationDecrementAccount = oRecordSet2.Fields.Item(0).Value
                                        oSR_Lines1.Add()
                                    End If
                                    oRecordSet.MoveNext()
                                Next
                                If oSR.Add <> 0 Then
                                    Dim Str As String = oCompany.GetLastErrorDescription
                                    oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription)
                                    ' Else
                                    oRecordSet2.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                                    GoTo D

                                End If
                                oRecordSet2.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                                oRs.MoveNext()
                            Next
                            oApplication.StatusBar.SetText("Inventory Revaluation Was successfully Posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            oRecordSet.DoQuery("Update OPDN set u_reval='Yes' where DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                            Dim oRs3 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim oRs4 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRs3.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                            oRs4.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocNum='" + objForm.Items.Item("8").Specific.value + "' order by omrv.DocNum desc")
                            If oRs3.Fields.Item(0).Value = oRs4.Fields.Item(0).Value Then
                                oRecordSet.DoQuery("Update OPDN set u_revalno='" & oRs3.Fields.Item(0).Value & "' where DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                            Else

                                oRecordSet.DoQuery("Update OPDN set u_revalno=('" & oRs3.Fields.Item(0).Value & "' +' To '+ '" & oRs4.Fields.Item(0).Value & "') where DocNum='" + objForm.Items.Item("8").Specific.value + "' ")
                            End If
                            Dim _str_DocNum As String
                            Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRs1.DoQuery("select Docnum,u_unit from OPDN Where DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                            _str_DocNum = oRs1.Fields.Item(0).Value
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("8").Specific.value = _str_DocNum
                            objForm.Items.Item("cpc").Specific.value = oRs1.Fields.Item(1).Value
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
D:
                        End If
                    Else
                        oRecordSet1.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                        ' Dim oRs As SAPbobsCOM.Recordset
                        ' oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRs.DoQuery("Select T1.u_rate,T1.u_glacct,T1.u_value from [@GEN_GRPO_LCOSTS] T0 Inner Join [@GEN_GRPO_LCOSTS_D0] T1 on T0.Code=T1.Code Where T0.u_grpono='" + objForm.Items.Item("8").Specific.value + "' And T1.u_glacct<>''")
                        For J As Int16 = 0 To oRs.RecordCount - 1
                            oRecordSet.DoQuery("SELECT T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], SUm(T1.[Vatsum]+T1.[LineTotal])'VatSum', T1.[ItemCode],Sum(T1.[Quantity])'Quantity', T1.[WhsCode], T1.[TaxCode] FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry Where T0.DocNum = '" + objForm.Items.Item("8").Specific.value + "' group by T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode], T1.[ItemCode], T1.[WhsCode], T1.[TaxCode]")
                            'oRecordSet.DoQuery("SELECT T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode],  T1.[ItemCode],sum(T1.[Quantity])'Quantity', T1.[WhsCode], T1.[TaxCode],sum(T1.LineTotal)'LineTotal',(select sum(distinct(PDN4.TaxSum)) from PDN1 inner join PDN4 on pdn1.DocEntry=pdn4.DocEntry And PDN1.LineNum = PDN4.LineNum where pdn4.NonDdctPrc='100' and pdn1.ItemCode=T1.ItemCode and pdn1.DocEntry=(Select MAX(Docentry)from OPDN))'Vatsum' FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry Where T0.DocEntry = (select Max(DocEntry) from OPDN) group by T0.DocEntry,T0.[DocNum],T0.[DocDate],T0.[CardCode],  T1.[ItemCode], T1.[WhsCode], T1.[TaxCode]"))
                            oRecordSet.MoveFirst()


                            For I As Int16 = 0 To oRecordSet.RecordCount - 1
                                '  oRecordSet1.DoQuery("Select u_rate,u_expact from [@GEN_EXP_TAX]")
                                'Dim Expenses_Tax As Double
                                ' Expenses_Tax = (CDbl(oRecordSet.Fields.Item("LineTotal").Value) * (CDbl(oRecordSet1.Fields.Item("u_rate").Value) / 100))
                                oRecordSet2.DoQuery("Insert into TempExpReval Values('" & oRecordSet.Fields.Item("ItemCode").Value & "','" & oRecordSet.Fields.Item("Quantity").Value & "','" & (CDbl(oRecordSet.Fields.Item("VatSum").Value)) & "','" & oRecordSet.Fields.Item("WhsCode").Value & "','" & oRecordSet.Fields.Item("DocDate").Value & "','" & oRecordSet.Fields.Item("DocNum").Value & "','" + oRs.Fields.Item("u_glacct").Value + "','" & oApplication.Company.UserName & "','0.00')")
                                oRecordSet.MoveNext()
                            Next
                            oRecordSet.DoQuery("Select Sum(Taxvalue),ItemCode,Quantity,WhsCode,Date,DocNum from TempExpReval Where UserName = '" & oApplication.Company.UserName & "'  Group By ItemCode,WhsCode,Date,DocNum,UserName,Quantity")

                            Dim oSR2 As SAPbobsCOM.MaterialRevaluation
                            Dim oSR_Lines1 As SAPbobsCOM.MaterialRevaluation_lines
                            oSR2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
                            oSR2.RevalType = "P"
                            oSR2.Reference2 = oRecordSet.Fields.Item("DocNum").Value

                            ' Dim lSBObob As SAPbobsCOM.SBObob
                            'Dim lRecordset As SAPbobsCOM.Recordset
                            Dim ld_CurrentItemCost1 As Decimal
                            Dim ld_NewItemCost1 As Decimal

                            oSR2.DocDate = oRecordSet.Fields.Item("Date").Value

                            ''*********************************
                            For I As Int16 = 0 To oRecordSet.RecordCount - 1
                                oRecordSet1.DoQuery("select price from MRV1 inner join OMRV on OMRV.docentry=MRV1.docentry where Itemcode='" + oRecordSet.Fields.Item("ItemCode").Value + "' and DocNum=(Select MAX(docnum)from omrv)")
                                oRecordSet1.DoQuery("Select AvgPrice,OnHand From OITW where ItemCode = '" & oRecordSet.Fields.Item("ItemCode").Value & "' And WhsCode = '" & oRecordSet.Fields.Item("WhsCode").Value & "'")
                                If oRecordSet1.RecordCount > 0 Then
                                    ld_CurrentItemCost1 = Val(oRecordSet1.Fields.Item("AvgPrice").Value)
                                    ld_NewItemCost1 = Math.Round(CDbl(oRecordSet1.Fields.Item("AvgPrice").Value) + ((oRs.Fields.Item(0).Value * oRecordSet.Fields.Item(2).Value) / oRecordSet1.Fields.Item(1).Value), 4)

                                Else
                                    ld_CurrentItemCost1 = 0
                                    ld_NewItemCost1 = 0
                                End If
                                If ld_CurrentItemCost1 <> ld_NewItemCost1 Then
                                    oSR_Lines1 = oSR.Lines
                                    oSR_Lines1.SetCurrentLine(I)
                                    oSR_Lines1.ItemCode = oRecordSet.Fields.Item("ItemCode").Value
                                    oSR_Lines1.WarehouseCode = oRecordSet.Fields.Item("WhsCode").Value

                                    oSR_Lines1.Price = ld_NewItemCost1
                                    oRecordSet2.DoQuery("select acctcode from OACT where FormatCode='" + oRs.Fields.Item(1).Value + "'")
                                    oSR_Lines1.RevaluationIncrementAccount = oRecordSet2.Fields.Item(0).Value
                                    oSR_Lines1.RevaluationDecrementAccount = oRecordSet2.Fields.Item(0).Value
                                    oSR_Lines1.Add()
                                End If
                                oRecordSet.MoveNext()
                            Next
                            If oSR.Add <> 0 Then
                                Dim Str As String = oCompany.GetLastErrorDescription
                                oApplication.StatusBar.SetText(oCompany.GetLastErrorDescription)
                                ' Else
                                oRecordSet2.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                                GoTo C

                            End If
                            oRecordSet2.DoQuery("delete from TempExpReval where UserName = '" & oApplication.Company.UserName & "' ")
                            oRs.MoveNext()
                        Next
                        oApplication.StatusBar.SetText("Inventory Revaluation Was successfully Posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        Dim oMr1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oMr1.DoQuery("Update OMRV Set Ref2='" + objForm.Items.Item("8").Specific.value + "' where OMRV.Docentry=(Select Max(DocEntry) From OMRV)")
                        oRecordSet.DoQuery("Update OPDN set u_reval='Yes' where DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                        Dim oRs3 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRs4 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRs3.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                        oRs4.DoQuery("select top 1(omrv.DocNum) from OMRV inner join OPDN on opdn.DocNum=OMRV.Ref2 where OPDN.DocNum='" + objForm.Items.Item("8").Specific.value + "' order by omrv.DocNum desc")
                        If oRs3.Fields.Item(0).Value = oRs4.Fields.Item(0).Value Then
                            oRecordSet.DoQuery("Update OPDN set u_revalno='" & oRs3.Fields.Item(0).Value & "' where DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                        Else

                            oRecordSet.DoQuery("Update OPDN set u_revalno=('" & oRs3.Fields.Item(0).Value & "' +' To '+ '" & oRs4.Fields.Item(0).Value & "') where DocNum='" + objForm.Items.Item("8").Specific.value + "' ")
                        End If
                        Dim _str_DocNum As String
                        Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRs1.DoQuery("select Docnum,u_unit from OPDN Where DocNum='" + objForm.Items.Item("8").Specific.value + "'")
                        _str_DocNum = oRs1.Fields.Item(0).Value
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objForm.Items.Item("8").Specific.value = _str_DocNum
                        objForm.Items.Item("cpc").Specific.value = oRs1.Fields.Item(1).Value
                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
C:
                    End If
                End If


                objForm.Refresh()
                ' If oSR.Add = 0 Then

            End If

            oRecordSet1.DoQuery("delete from TempReval where UserName = '" & oApplication.Company.UserName & "'")
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
        '  End If
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try

            objForm = oApplication.Forms.Item(FormUID)
            objForm.Refresh()
            objMatrix = objForm.Items.Item("38").Specific
            Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim whs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ' Whs.DoQuery("Select Itemcode from OPDN")
            oRs.DoQuery("select COUNT(*) from [@GEN_GRPO_LCOSTS] where U_grpono='" + objForm.Items.Item("8").Specific.value + "' and u_macid='" + MAC_ID + "'")
            If oRs.Fields.Item(0).Value = "0" Then ' Or oRs.Fields.Item(0).Value = "" Then
                oApplication.StatusBar.SetText("Please Select LandedCost By Clicking LC Button", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                whs.DoQuery("select itemcode,Quantity,DocNum,opdn.DocDate from OPDN inner join PDN1 on opdn.DocEntry = pdn1.DocEntry inner join ostc on pdn1.taxcode=ostc.code where ItemCode='" + objMatrix.Columns.Item("1").Cells.Item(Row).Specific.value + "' and Whscode='" + objMatrix.Columns.Item("24").Cells.Item(Row).Specific.value + "' and opdn.U_lcadd='N'  And OPDN.Docdate>='20130401' and OPDN.U_insstat='Open'")
                If Trim(objMatrix.Columns.Item("160").Cells.Item(Row).Specific.Value).Equals("") = True Then
                    oApplication.StatusBar.SetText("TaxCode should not be left blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                    ' Exit Function
                ElseIf whs.RecordCount > 0 Then
                    ' Return True
                    'End If
                    Dim LCAdd As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    GRPODOCNUM = objForm.Items.Item("8").Specific.Value
                    LCAdd.DoQuery("Select u_reval,u_lcadd from OPDN Where Docnum='" & whs.Fields.Item("DocNum").Value & "'")
                    If LCAdd.Fields.Item(1).Value = "N" Then
                        oApplication.StatusBar.SetText("Please Select Landed Cost By Clicking on LC Button For GRN No.'" & whs.Fields.Item("DocNum").Value & "'", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf (LCAdd.Fields.Item(0).Value = "" Or LCAdd.Fields.Item(0).Value = "NO" Or LCAdd.Fields.Item(0).Value = "Partial") Then
                        oApplication.StatusBar.SetText("Please Click on Manual Revaluation For GRN No.'" & whs.Fields.Item("DocNum").Value & "'", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                        'Else
                        '    Dim whs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '    For Row As Integer = 1 To objMatrix.VisualRowCount
                        '        whs.DoQuery("select itemcode,Quantity,DocNum,opdn.DocDate from OPDN inner join PDN1 on opdn.DocEntry=pdn1.DocEntry inner join ostc on pdn1.taxcode=ostc.code where ItemCode='" + objMatrix.Columns.Item("1").Cells.Item(Row).Specific.value + "' and Whscode='" + objMatrix.Columns.Item("24").Cells.Item(Row).Specific.value + "' and opdn.U_insstat='Open' and ostc.Rate<>'0.00' and (u_reval='NO' or u_LCAdd='N')  and OPDN.Docnum<>'" + GRPODOCNUM + "'")
                        '        If whs.RecordCount > 0 Then
                        '            oApplication.StatusBar.SetText("Please Select LC And Make 'Manual Revaluation' for pending GRN No '" & whs.Fields.Item("DocNum").Value & "'", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '            BubbleEvent = False
                        '        End If
                        '    Next
                    End If
                    'oApplication.StatusBar.SetText("Please Select LC For GRN No '" & whs.Fields.Item("DocNum").Value & "' Items to Main Warehouse", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                    'Next
                    'Else
                    '   Return True
                End If
            Next
            Return True
            ' End Select
        Catch ex As Exception

        End Try
    End Function

End Class

