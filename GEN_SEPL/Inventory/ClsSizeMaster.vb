Public Class ClsSizeMaster

#Region "        Declaration        "

    Dim oUtilities As New ClsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim ROW_ID As Integer = 0
#End Region

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                    objForm.Title = "Size Master"
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If pVal.BeforeAction = True Then
                        objMatrix = objForm.Items.Item("3").Specific
                        If Me.CheckForWild(Trim(objMatrix.Columns.Item("Code").Cells.Item(pVal.Row).Specific.value)) = False Then
                            oApplication.StatusBar.SetText("Special characters are not allowed in code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Function CheckForWild(ByVal Code As String) As Boolean
        Dim iCount As Integer = 0
        Dim c As Char
        For iCount = 0 To Code.Length - 1
            c = Code.Substring(iCount, 1)
            Select Case c
                Case "!" : Return False
                Case "@" : Return False
                Case "#" : Return False
                Case "$" : Return False
                Case "%" : Return False
                Case "^" : Return False
                Case "&" : Return False
                Case "*" : Return False
                Case "(" : Return False
                Case ")" : Return False
                Case "_" : Return False
                Case "+" : Return False
                Case "=" : Return False
                Case "[" : Return False
                Case "]" : Return False
                Case "{" : Return False
                Case "}" : Return False
                Case "\" : Return False
                Case "|" : Return False
                Case ";" : Return False
                Case ":" : Return False
                Case """" : Return False
                Case "'" : Return False
                Case "<" : Return False
                Case ">" : Return False
                Case "?" : Return False
                Case "/" : Return False
            End Select
        Next
        Return True
    End Function

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "GEN_SIZE_MST" '//Opening OPERATIONS SCREEN BY Activating Menu
                        Dim Rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Rs.DoQuery("Select 47616 + count(*) MenuItem from OUDO where CanDefForm = 'Y' and code <= 'GEN_SIZE_MST'")
                        oApplication.ActivateMenuItem(Rs.Fields.Item("MenuItem").Value)
                End Select
            End If
        Catch ex As Exception

        End Try
    End Sub

End Class
