Public Class ClsUser

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm, objSubForm, objSForm As SAPbouiCOM.Form
    Dim objItem, objOldItem As SAPbouiCOM.Item
    Dim objMatrix, objSubMatrix, objBatchMatrix As SAPbouiCOM.Matrix
    Dim oChk As SAPbouiCOM.CheckBox
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            objOldItem = objForm.Items.Item("12")
            objItem = objForm.Items.Add("alwis", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width + 100
            objItem.Left = objOldItem.Left + 120
            objItem.Top = objOldItem.Top
            objItem.Specific.Caption = "MRN Approver"
            objItem.Specific.databind.setbound(True, "OUSR", "U_approve")
            objItem.LinkTo = "12"

            objOldItem = objForm.Items.Item("alwis")
            objItem = objForm.Items.Add("cstsht", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width + 100
            objItem.Left = objOldItem.Left + 100
            objItem.Top = objOldItem.Top
            objItem.Specific.Caption = "Cost Sheet Approver"
            objItem.Specific.databind.setbound(True, "OUSR", "U_cstsht")
            objItem.LinkTo = "alwis"

            'Vijeesh'
            'objOldItem = objForm.Items.Item("10000116")
            'objItem = objForm.Items.Add("cstapv", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            'objItem.Height = objOldItem.Height
            'objItem.Width = objOldItem.Width + 100
            'objItem.Left = objOldItem.Left + 100
            'objItem.Top = objOldItem.Top
            'objItem.Specific.Caption = "Cost Approve"
            'objItem.Specific.databind.setbound(True, "OUSR", "U_costap")
            'objItem.LinkTo = "10000116"
            'Vijeesh'

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
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
