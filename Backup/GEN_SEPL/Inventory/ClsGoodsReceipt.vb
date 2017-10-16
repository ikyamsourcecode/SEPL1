Public Class ClsGoodsReceipt

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
            'objForm.Title = "Material Issue Note"
            oTempItem = objForm.Items.Item("16")
            oItem = objForm.Items.Add("type", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OIGN", "u_type")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("mrnno", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OIGN", "u_mrnno")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("subconno", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OIGN", "u_subconno")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("isstyp", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OIGN", "u_isstyp")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("subretno", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OIGN", "u_subretno")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("sono", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OIGN", "u_sono")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("itemcode", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OIGN", "u_itemcode")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("sfgcode", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OIGN", "u_sfgcode")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("grnno", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OIGN", "u_grnno")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("scpono", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OIGN", "u_DocNum")
            oItem.Visible = False
            oItem.LinkTo = "16"
            oItem = objForm.Items.Add("scpotp", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 15
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = 14
            oItem.Specific.databind.setbound(True, "OIGN", "u_Type")
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

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                Case SAPbouiCOM.BoEventTypes.et_CLICK

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
                        Dim GIForm As SAPbouiCOM.Form = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        Dim oMatrix As SAPbouiCOM.Matrix
                        oMatrix = GIForm.Items.Item("13").Specific
                        Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim MRNEntry As String
                        If Trim(GIForm.Items.Item("mrnno").Specific.Value) <> "" Then
                            'If Trim(GIForm.Items.Item("isstyp").Specific.value) = "I" Then
                            '    oRecordSet.DoQuery("Select DocEntry From [@GEN_MREQ] Where DocNum = '" + Trim(GIForm.Items.Item("mrnno").Specific.value) + "'")
                            '    MRNEntry = oRecordSet.Fields.Item("DocEntry").Value
                            '    For i As Integer = 1 To oMatrix.VisualRowCount
                            '        oRecordSet.DoQuery("Update [@GEN_MREQ_D0] Set u_issued = u_issued + Convert(Money,'" + Trim(oMatrix.Columns.Item("9").Cells.Item(i).Specific.value) + "') Where DocEntry = '" + MRNEntry + "' And LineId = '" + Trim(oMatrix.Columns.Item("U_mrnlid").Cells.Item(i).Specific.Value) + "'")
                            '        oRecordSet.DoQuery("Update [@GEN_MREQ_D0] Set u_stat = 'Closed' Where DocEntry = '" + MRNEntry + "' And LineId = '" + Trim(oMatrix.Columns.Item("U_mrnlid").Cells.Item(i).Specific.Value) + "' And u_issued >= u_rqstqty")
                            '    Next
                            '    oRecordSet.DoQuery("Update [@GEN_MREQ] Set u_status = 'Closed' Where DocEntry = '" + MRNEntry + "' And DocEntry Not in (Select DocEntry From [@GEN_MREQ_D0] Where DocEntry = '" + MRNEntry + "' And IsNull(u_stat,'N') != 'Closed')")
                            'End If
                            If Trim(GIForm.Items.Item("isstyp").Specific.value) = "R" Then
                                oRecordSet.DoQuery("Select DocEntry From [@GEN_MREQ] Where DocNum = '" + Trim(GIForm.Items.Item("mrnno").Specific.value) + "'")
                                MRNEntry = oRecordSet.Fields.Item("DocEntry").Value
                                For i As Integer = 1 To oMatrix.VisualRowCount
                                    oRecordSet.DoQuery("Update [@GEN_MREQ_D0] Set u_returned = IsNull(u_returned,0) + Convert(Money,'" + Trim(oMatrix.Columns.Item("9").Cells.Item(i).Specific.value) + "') Where DocEntry = '" + MRNEntry + "' And LineId = '" + Trim(oMatrix.Columns.Item("U_mrnlid").Cells.Item(i).Specific.Value) + "'")
                                Next
                            End If
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
                        If objForm.TypeEx = "721" Then
                            'BubbleEvent = False
                        End If
                    Case "1287"
                        If objForm.TypeEx = "721" Then
                            'BubbleEvent = False
                        End If
                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
