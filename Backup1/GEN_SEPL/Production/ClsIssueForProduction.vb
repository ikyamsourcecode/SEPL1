Public Class ClsIssueForProduction

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
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
    Dim PTNNo As String
    Dim SONO As String
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            objForm.Title = "Book Consumption"
            oTempItem = objForm.Items.Item("21")
            oItem = objForm.Items.Add("sono", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OIGE", "u_sono")
            oItem.Visible = False
            oItem.LinkTo = "21"
            oItem = objForm.Items.Add("ptnno", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OIGE", "u_ptnno")
            oItem.Visible = False
            oItem.LinkTo = "16"
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
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        DocNO = objForm.Items.Item("7").Specific.Value
                        PTNNo = objForm.Items.Item("ptnno").Specific.value
                        SONO = objForm.Items.Item("sono").Specific.value
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If pVal.ItemUID = "13" And pVal.ColUID = "15" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = False Then
                        objMatrix = objForm.Items.Item("13").Specific
                        oRSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRSet.DoQuery("Select OnHand From OITW Where WhsCode = '" + Trim(objMatrix.Columns.Item("15").Cells.Item(pVal.Row).Specific.value) + "' And ItemCode = '" + Trim(objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.value) + "'")
                        objMatrix.Columns.Item("U_qty").Cells.Item(pVal.Row).Specific.value = oRSet.Fields.Item("OnHand").Value
                    End If
                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    If pVal.ItemUID = "13" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        'Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'oRSet.DoQuery("Select IsNull(U_alwis,'N') As 'AllowIssue' From OUSR Where User_Code = '" + oCompany.UserName.ToString.Trim + "'")
                        'If Trim(oRSet.Fields.Item("AllowIssue").Value) = "N" Then
                        '    BubbleEvent = False
                        'End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "13" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        'Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'oRSet.DoQuery("Select IsNull(U_alwis,'N') As 'AllowIssue' From OUSR Where User_Code = '" + oCompany.UserName.ToString.Trim + "'")
                        'If Trim(oRSet.Fields.Item("AllowIssue").Value) = "N" Then
                        '    BubbleEvent = False
                        'End If
                    End If
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
                            oRecordSet.DoQuery("Update [@GEN_PTN] Set u_status = 'Consumed' Where Docnum = '" + PTNNo + "' And u_sono = '" + SONO + "'")
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
                    Case "1293"
                        If objForm.TypeEx = "65213" Then
                            BubbleEvent = False
                        End If
                    Case "1287"
                        If objForm.TypeEx = "65213" Then
                            BubbleEvent = False
                        End If
                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
