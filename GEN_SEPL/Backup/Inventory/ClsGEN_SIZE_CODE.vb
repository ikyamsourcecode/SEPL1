Public Class ClsGEN_SIZE_CODE

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset
#End Region

    Sub CreateForm()
        Try
            oUtilities.SAPXML("GEN_SIZE_CODE.xml")
            objForm = oApplication.Forms.GetForm("GEN_SIZE_CODE", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SIZE_CODE")
            objForm.EnableMenu("1281", True)
            objForm.EnableMenu("1288", True)
            objForm.EnableMenu("1289", True)
            objForm.EnableMenu("1290", True)
            objForm.EnableMenu("1291", True)
            objForm.DataBrowser.BrowseBy = "name"
            objForm.Items.Item("name").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Select()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                objForm.EnableMenu("1282", False)
            End If
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSet.DoQuery("Select Count(*) + 1 'Count' From [@GEN_SIZE_CODE]")
            objForm.Items.Item("code").Specific.value = oRSet.Fields.Item("Count").Value
            objForm.Items.Item("name").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try
            objForm = oApplication.Forms.Item(FormUID)
            If Trim(objForm.Items.Item("name").Specific.Value) = "" Then
                oApplication.StatusBar.SetText("Please enter Size Code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Name From [@GEN_SIZE_CODE] Where Name = '" + Trim(objForm.Items.Item("name").Specific.Value) + "' And Code <> '" + Trim(objForm.Items.Item("code").Specific.Value) + "' And u_type = '" + Trim(objForm.Items.Item("type").Specific.value) + "'")
            If oRecordSet.RecordCount > 0 Then
                oApplication.StatusBar.SetText("Size Code already exists in the database", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            If CheckForWild(Trim(objForm.Items.Item("name").Specific.value)) = False Then
                oApplication.StatusBar.SetText("Size Code cannot have special characters", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSet.DoQuery("Select IsNull(Max(Convert(Int,Code)),0) + 1 'Count' From [@GEN_SIZE_CODE]")
            oDBs_Head.SetValue("Code", 0, oRSet.Fields.Item("Count").Value)
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation(FormUID) = False Then BubbleEvent = False
                    ElseIf pVal.ItemUID = "1" And pVal.ActionSuccess = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        Me.SetDefault(FormUID)
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                Dim oForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "GEN_SIZE_CODE"
                        Me.CreateForm()
                    Case "1282"
                        If oForm.TypeEx = "GEN_SIZE_CODE" Then
                            Me.SetDefault(oForm.UniqueID)
                        End If
                    Case "1281"
                        If oForm.TypeEx = "GEN_SIZE_CODE" Then
                            Me.SetDefault(oForm.UniqueID)
                            oForm.EnableMenu("1282", True)
                        End If
                    Case "1288", "1289", "1290", "1291"
                        If oForm.TypeEx = "GEN_SIZE_CODE" Then
                            oForm.EnableMenu("1282", True)
                        End If
                End Select
            End If
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

End Class
