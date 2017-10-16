Imports System
Public Class ClsBP

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
    Dim oChk As SAPbouiCOM.CheckBox
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
    Dim VendCode As String
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oTempItem = objForm.Items.Item("12")
            oItem = objForm.Items.Add("sadv", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Top = oTempItem.Top + oTempItem.Height + 5
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = oTempItem.Height
            oItem.Specific.Caption = "Advance Opening Balance"
            oItem.LinkTo = "12"
            oTempItem = objForm.Items.Item("11")
            oItem = objForm.Items.Add("eadv", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = oTempItem.Top + oTempItem.Height + 5
            oItem.Width = oTempItem.Width
            oItem.Left = oTempItem.Left
            oItem.Height = oTempItem.Height
            oItem.Specific.databind.setbound(True, "OCRD", "u_dpmadv")
            oItem.RightJustified = True
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.LinkTo = "11"
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
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("OCRD")
                        Dim Subcontractor As String = oDBs_Head.GetValue("U_subcon", 0).ToString().Trim()
                        If Subcontractor = "Y" Or Subcontractor = "y" Then
                            If oApplication.MessageBox("WHETHER THIS LOCATION IS COVERED FOR INSURANCE PURPOSE ?", 2, "Yes", "No") = 2 Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class