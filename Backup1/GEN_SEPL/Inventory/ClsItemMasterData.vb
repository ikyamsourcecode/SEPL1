Public Class ClsItemMasterData

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim objItem, objOldItem As SAPbouiCOM.Item
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            objOldItem = objForm.Items.Item("46")
            objItem = objForm.Items.Add("scolor", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.caption = "Color"
            objItem.LinkTo = "46"
            objOldItem = objForm.Items.Item("47")
            objItem = objForm.Items.Add("color", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.databind.setbound(True, "OITM", "u_color")
            objItem.LinkTo = "47"
            objItem = objForm.Items.Add("colornm", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left + 80
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.databind.setbound(True, "OITM", "u_colornm")
            objItem.LinkTo = "47"

            objOldItem = objForm.Items.Item("scolor")
            objItem = objForm.Items.Add("scust", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.caption = "Customer"
            objItem.LinkTo = "scolor"
            objOldItem = objForm.Items.Item("color")
            objItem = objForm.Items.Add("cust", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.databind.setbound(True, "OITM", "u_cust")
            objItem.LinkTo = "color"

            objItem = objForm.Items.Add("custnm", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left + 80
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.databind.setbound(True, "OITM", "u_custnm")
            objItem.LinkTo = "color"

            objOldItem = objForm.Items.Item("scust")
            objItem = objForm.Items.Add("stype", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.caption = "Item Type"
            objItem.LinkTo = "scust"
            objOldItem = objForm.Items.Item("cust")
            objItem = objForm.Items.Add("type", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.databind.setbound(True, "OITM", "u_type")
            objItem.LinkTo = "cust"
            objItem = objForm.Items.Add("typenm", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left + 80
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.databind.setbound(True, "OITM", "u_typenm")
            objItem.LinkTo = "cust"

            objOldItem = objForm.Items.Item("stype")
            objItem = objForm.Items.Add("ssize", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.caption = "Size"
            objItem.LinkTo = "stype"
            objOldItem = objForm.Items.Item("type")
            objItem = objForm.Items.Add("size", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.databind.setbound(True, "OITM", "u_size")
            objItem.LinkTo = "type"
            objItem = objForm.Items.Add("sizenm", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left + 80
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.databind.setbound(True, "OITM", "u_sizenm")
            objItem.LinkTo = "type"

            objOldItem = objForm.Items.Item("ssize")
            objItem = objForm.Items.Add("ssubtype", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.caption = "Sub Type"
            objItem.LinkTo = "ssize"
            objOldItem = objForm.Items.Item("size")
            objItem = objForm.Items.Add("subtype", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.databind.setbound(True, "OITM", "u_subtype")
            objItem.LinkTo = "size"
            objItem = objForm.Items.Add("subtpnm", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left + 80
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.databind.setbound(True, "OITM", "u_subtpnm")
            objItem.LinkTo = "size"

            objOldItem = objForm.Items.Item("ssubtype")
            objItem = objForm.Items.Add("sstyle", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.caption = "Style"
            objItem.LinkTo = "ssubtype"
            objOldItem = objForm.Items.Item("subtype")
            objItem = objForm.Items.Add("style", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.databind.setbound(True, "OITM", "u_style")
            objItem.LinkTo = "subtype"
            objItem = objForm.Items.Add("stylenm", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left + 80
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.databind.setbound(True, "OITM", "u_stylenm")
            objItem.LinkTo = "subtype"

            objOldItem = objForm.Items.Item("sstyle")
            objItem = objForm.Items.Add("sqlty", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.caption = "Quality"
            objItem.LinkTo = "sstyle"
            objOldItem = objForm.Items.Item("style")
            objItem = objForm.Items.Add("qlty", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.databind.setbound(True, "OITM", "u_qlty")
            objItem.LinkTo = "style"
            objItem = objForm.Items.Add("qltynm", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.FromPane = objOldItem.FromPane
            objItem.ToPane = objOldItem.ToPane
            objItem.Left = objOldItem.Left + 80
            objItem.Top = objOldItem.Top + objOldItem.Height + 1
            objItem.Width = objOldItem.Width
            objItem.Height = objOldItem.Height
            objItem.Specific.databind.setbound(True, "OITM", "u_qltynm")
            objItem.LinkTo = "style"
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
                        Me.CreateForm(FormUID)
                    End If

                    'Vijeesh
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    'If pVal.BeforeAction = True Then
                    '    If pVal.ItemUID = "248" Then
                    '        Dim oForm As SAPbouiCOM.Form = oApplication.Forms.Item(pVal.FormUID)
                    '        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '        oRecordSet.DoQuery("Select isnull(U_costap,'N')U_costap from OUSR where USER_CODE='" + oCompany.UserName.ToString.Trim + "'")
                    '        If oRecordSet.Fields.Item("U_costap").Value.ToString() = "Y" Then
                    '            BubbleEvent = False
                    '        End If
                    '    End If
                    'End If
                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    'If pVal.BeforeAction = True Then
                    '    If pVal.ItemUID = "64" And pVal.CharPressed <> 9 Then
                    '        Dim oForm As SAPbouiCOM.Form = oApplication.Forms.Item(pVal.FormUID)
                    '        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '        oRecordSet.DoQuery("Select isnull(U_costap,'N')U_costap from OUSR where USER_CODE='" + oCompany.UserName.ToString.Trim + "'")
                    '        If oRecordSet.Fields.Item("U_costap").Value.ToString() = "Y" Then
                    '            BubbleEvent = False
                    '        End If
                    '    End If
                    'End If
                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS
                    'If pVal.BeforeAction = False And pVal.ActionSuccess = True Then
                    '    If pVal.ItemUID = "64" Then
                    '        Dim oForm As SAPbouiCOM.Form = oApplication.Forms.Item(pVal.FormUID)
                    '        Try
                    '            Dim objItem As SAPbouiCOM.Item
                    '            Dim objEditText As SAPbouiCOM.EditText
                    '            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            oRecordSet.DoQuery("Select isnull(U_costap,'N')U_costap from OUSR where USER_CODE='" + oCompany.UserName.ToString.Trim + "'")
                    '            Dim _str_User As String = oRecordSet.Fields.Item("U_costap").Value.ToString()
                    '            objItem = oForm.Items.Item("64")
                    '            objEditText = oForm.Items.Item("251").Specific
                    '            oForm.Freeze(True)
                    '            oRecordSet.DoQuery("Select InvntItem,EvalSystem from OITM where ItemCode='" + oForm.Items.Item("5").Specific.Value + "'")
                    '            If (oRecordSet.Fields.Item("InvntItem").Value.ToString = "N" And oRecordSet.Fields.Item("EvalSystem").Value.ToString = "S" And _str_User = "Y") Then
                    '                If objItem.Visible = True Then
                    '                    objItem.BackColor = RGB(255, 255, 255)
                    '                    objItem.ForeColor = RGB(255, 255, 255)
                    '                    objEditText.Active = True
                    '                    objItem.Enabled = False
                    '                    objItem.Visible = False
                    '                End If
                    '                BubbleEvent = False
                    '            End If
                    '            oForm.Freeze(False)
                    '        Catch ex As Exception
                    '            oForm.Freeze(False)
                    '        End Try
                    '    End If
                    'End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                   

                    '    
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    'If pVal.BeforeAction = False And pVal.ActrionSuccess = True Then
                    '    If pVal.ItemUID = "26" Then
                    '        Dim oForm As SAPbouiCOM.Form = oApplication.Forms.Item(pVal.FormUID)
                    '        Try
                    '            Dim objItem As SAPbouiCOM.Item
                    '            Dim objEditText As SAPbouiCOM.EditText
                    '            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            oRecordSet.DoQuery("Select isnull(U_costap,'N')U_costap from OUSR where USER_CODE='" + oCompany.UserName.ToString.Trim + "'")
                    '            Dim _str_User As String = oRecordSet.Fields.Item("U_costap").Value.ToString()
                    '            objItem = oForm.Items.Item("64")
                    '            objEditText = oForm.Items.Item("251").Specific
                    '            oForm.Freeze(True)
                    '            oRecordSet.DoQuery("Select InvntItem,EvalSystem from OITM where ItemCode='" + oForm.Items.Item("5").Specific.Value + "'")
                    '            If (oRecordSet.Fields.Item("InvntItem").Value.ToString = "N" And oRecordSet.Fields.Item("EvalSystem").Value.ToString = "S" And _str_User = "Y") Then
                    '                objItem.Visible = False
                    '                BubbleEvent = False
                    '            End If
                    '            oForm.Freeze(False)
                    '        Catch ex As Exception
                    '            oForm.Freeze(False)
                    '        End Try
                    '    End If
                    'End If
                    'Vijeesh
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub ItemEvent_udf(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Select Case pVal.EventType
            Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                objForm = oApplication.Forms.Item(pVal.FormUID)
            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                If pVal.BeforeAction = False Then
                    If pVal.ItemUID = "U_tariff" Or pVal.ItemUID = "U_blend" Then
                        If objForm.Items.Item("U_tariff").Specific.Value <> "" And objForm.Items.Item("U_blend").Specific.Value <> "" Then
                            Dim oTariff As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oTariff.DoQuery("Select A.Code,B.U_tarif,B.U_desc,B.U_dbk,B.U_cap from [@UBG_DBK_LST] A inner join [@UBG_DBK_LST_D0] B on A.Code = B.Code Where B.U_tarif = '" & objForm.Items.Item("U_tariff").Specific.Value.ToString.Trim & "' and A.Code = '" & objForm.Items.Item("U_blend").Specific.Value.ToString.Trim & "'")
                            objForm.Items.Item("U_expper").Specific.Value = oTariff.Fields.Item(3).Value
                            objForm.Items.Item("U_cap").Specific.Value = oTariff.Fields.Item(4).Value
                        End If
                    End If '
                End If
            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                If pVal.BeforeAction = True Then
                    If pVal.ItemUID = "U_cap" Or pVal.ItemUID = "U_expper" Then
                        oApplication.SetStatusBarMessage("Non-editable item", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If

        End Select
           
    End Sub
    'Vijeesh'
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    'If BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.BeforeAction = False Then
                    '    Dim oForm As SAPbouiCOM.Form = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    '    Dim objItem As SAPbouiCOM.Item
                    '    Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '    oRecordSet.DoQuery("Select isnull(U_costap,'N')U_costap from OUSR where USER_CODE='" + oCompany.UserName.ToString.Trim + "'")
                    '    Dim _str_User As String = oRecordSet.Fields.Item("U_costap").Value.ToString()
                    '    oRecordSet.DoQuery("Select InvntItem,EvalSystem from OITM where ItemCode='" + oForm.Items.Item("5").Specific.Value + "'")
                    '    objItem = oForm.Items.Item("64")
                    '    If (oRecordSet.Fields.Item("InvntItem").Value.ToString = "N" And oRecordSet.Fields.Item("EvalSystem").Value.ToString = "S" And _str_User = "Y") Then
                    '        'objItem.BackColor = RGB(200, 212, 248)
                    '        'objItem.ForeColor = RGB(200, 212, 248)
                    '        'objItem.Enabled = False
                    '        objItem.Visible = False
                    '    End If
                    'End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            'Dim oForm As SAPbouiCOM.Form = oApplication.Forms.Item(eventInfo.FormUID)
            'If eventInfo.ItemUID = "64" And eventInfo.BeforeAction = True Then
            '    Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '    oRecordSet.DoQuery("Select isnull(U_costap,'N')U_costap from OUSR where USER_CODE='" + oCompany.UserName.ToString.Trim + "'")
            '    Dim _str_User As String = oRecordSet.Fields.Item("U_costap").Value.ToString()
            '    If _str_User = "Y" Then
            '        eventInfo.RemoveFromContent("772")
            '    End If
            'End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
