Public Class ClsJE


#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim objMatrix As SAPbouiCOM.Matrix
#End Region



    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("76").Specific
            Select Case pVal.EventType
                '    Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                '        objForm = oApplication.Forms.Item(pVal.FormUID)
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Dim USER_NAME As String = oCompany.UserName
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                        '      If Trim(objForm.Items.Item("3").Specific.value) = "S" Then
                        If USER_NAME <> "manager" Then
                            Dim GLAccount As String
                            For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                                GLAccount = objMatrix.Columns.Item("1").Cells.Item(Row).Specific.value
                                Dim GLacc As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                GLacc.DoQuery("Select (substring('" + GLAccount + "', 1, len('" + GLAccount + "')-3)+RIGHT('" + GLAccount + "',2))")
                                Dim ManualJE As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                ManualJE.DoQuery("Select count(Code) From [@GEN_M_JE] where code='" + GLacc.Fields.Item(0).Value + "'")
                                If ManualJE.Fields.Item(0).Value > 0 Then
                                    oApplication.StatusBar.SetText("You Are Not Permitted To Perform This Action", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Next
                        End If
                        For Row As Integer = 1 To objMatrix.VisualRowCount - 1
                            Dim GAccnt As String
                            GAccnt = objMatrix.Columns.Item("1").Cells.Item(Row).Specific.value
                            Dim Gacc As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim Gacc_COA As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim CA As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Gacc.DoQuery("Select (substring('" + GAccnt + "', 1, len('" + GAccnt + "')-3)+RIGHT('" + GAccnt + "',2))")
                            Gacc_COA.DoQuery("Select U_ccentre From OACT where FormatCode='" + Gacc.Fields.Item(0).Value + "'")
                            CA.DoQuery("Select formatcode from OACT where Formatcode='" + Gacc.Fields.Item(0).Value + "'")
                            If CA.RecordCount > 0 Then
                                If Gacc_COA.Fields.Item(0).Value = "N" Or Gacc_COA.Fields.Item(0).Value = "" Then
                                    If objMatrix.Columns.Item("10002014").Cells.Item(Row).Specific.Value = "" Then
                                        Dim Rowval As Integer = Convert.ToInt32(Int(Row))
                                        oApplication.StatusBar.SetText("Please select CostCentre In Row - " & Row & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Next
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        '            If Me.Validation(FormUID) = False Then BubbleEvent = False
                        '        ElseIf pVal.ItemUID = "1" And pVal.ActionSuccess = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        '            Me.SetDefault(FormUID)
                    End If
            End Select
        Catch ex As Exception
            ' oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

  




End Class
