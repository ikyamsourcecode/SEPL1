Public Class ClsCOA


#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm, objSubForm, objSForm As SAPbouiCOM.Form
    Dim objItem, objOldItem, TempItem As SAPbouiCOM.Item
    Dim objMatrix, objSubMatrix, objBatchMatrix As SAPbouiCOM.Matrix
    Dim SZDBHead As SAPbouiCOM.DBDataSource
    Dim SZDBDetail As SAPbouiCOM.DBDataSource
    Dim SMDBHead As SAPbouiCOM.DBDataSource
    Dim SMDBDetail As SAPbouiCOM.DBDataSource
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
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
    Dim DeleteItemCode As String
    Dim oBool As Boolean = False
    Dim DOCNUM As String = ""
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            objOldItem = objForm.Items.Item("49")
            objItem = objForm.Items.Add("ccentre", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            objItem.Height = objOldItem.Height
            objItem.Width = objOldItem.Width + 20
            objItem.Left = objOldItem.Left + 120
            objItem.Top = objOldItem.Top
            objItem.Specific.Caption = "Cost Center Not Required"
            objItem.Specific.databind.setbound(True, "OACT", "U_ccentre")
            objItem.LinkTo = "49"
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
