Public Class ClsProductionOrder
#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oEdit As SAPbouiCOM.EditText
    Dim oItem As SAPbouiCOM.Item
    Dim oTempItem As SAPbouiCOM.Item
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oCFL As SAPbouiCOM.ChooseFromList
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

            oTempItem = objForm.Items.Item("2")
            oItem = objForm.Items.Add("Ref", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Top = oTempItem.Top
            oItem.Left = oTempItem.Left + oTempItem.Width + 5
            oItem.Height = oTempItem.Height
            oItem.Width = oTempItem.Width + 50
            oItem.Specific.caption = "Refresh Location"

           


        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("OWOR")
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    If pVal.BeforeAction = True Then
                        Me.CreateForm(pVal.FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.ItemUID = "Ref" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Dim ITMatrix As SAPbouiCOM.Matrix
                        ITMatrix = objForm.Items.Item("37").Specific
                        oDBs_Detail = objForm.DataSources.DBDataSources.Item("WOR1")

                        For i As Integer = 1 To ITMatrix.VisualRowCount - 1
                            Dim whs As String = oDBs_Detail.GetValue("warehouse", i - 1)
                            ITMatrix.Columns.Item("10").Cells.Item(i).Specific.Value = ""
                            ITMatrix.Columns.Item("10").Cells.Item(i).Specific.Value = whs.Trim()

                        Next

                    End If
     
               
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
