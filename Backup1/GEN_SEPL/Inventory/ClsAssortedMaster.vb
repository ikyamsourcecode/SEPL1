Public Class ClsAssortedMaster

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim ITEM_ID As String
#End Region

    Sub CreateForm()
        Try
            oUtilities.SAPXML("AssortedMaster.xml")
            objForm = oApplication.Forms.GetForm("GEN_ASSORTMENT", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_ASSORTMENT")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_ASSORTMENT_D0")
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
            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                objForm.EnableMenu("1282", False)
            End If
            Dim sCode As String
            sCode = oUtilities.keygencode("@GEN_ASSORTMENT")
            oDBs_Head.SetValue("Code", 0, sCode)
            objForm.Items.Item("name").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            objMatrix = objForm.Items.Item("SizeMatrix").Specific
            objMatrix.Clear()
            objMatrix.FlushToDataSource()
            objMatrix.Clear()
            objMatrix.AddRow()
            objMatrix.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, objMatrix.VisualRowCount, objMatrix)
        Catch ex As Exception
            objForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_ASSORTMENT_D0")
            objMatrix = oMatrix
            objMatrix.FlushToDataSource()
            oDBs_Detail.Offset = Row - 1
            oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, objMatrix.VisualRowCount)
            oDBs_Detail.SetValue("u_size", oDBs_Detail.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try
            objForm = oApplication.Forms.Item(FormUID)
            If CheckForWild(Trim(objForm.Items.Item("name").Specific.value)) = False Then
                oApplication.StatusBar.SetText("Assortment code cannot have special characters", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
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
                        Dim sCode As String
                        sCode = oUtilities.keygencode("@GEN_ASSORTMENT")
                        objForm.Items.Item("code").Specific.value = sCode
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If oCFL.UniqueID = "SCFL" Then
                            objMatrix = objForm.Items.Item("SizeMatrix").Specific
                            Dim Total As Double = 0
                            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_ASSORTMENT_D0")
                            Dim Flag As Boolean = False
                            Dim errflag As Boolean = False
                            If objMatrix.VisualRowCount = 1 Or pVal.Row = objMatrix.VisualRowCount Then
                                Flag = True
                            End If
                            For i As Integer = 0 To oDT.Rows.Count - 1
                                Dim cflSelectedcount As Integer = oDT.Rows.Count
                                If i < cflSelectedcount - 1 Then
                                    objMatrix.AddRow(1, pVal.Row)
                                    oDBs_Detail.InsertRecord(pVal.Row + i - 1)
                                End If
                                oDBs_Detail.Offset = pVal.Row - 1 + i
                                oDBs_Detail.SetValue("LineID", oDBs_Detail.Offset, i + pVal.Row)
                                oDBs_Detail.SetValue("U_size", oDBs_Detail.Offset, oDT.GetValue("Code", i))
                                objMatrix.SetLineData(pVal.Row + i)
                                objForm.EnableMenu("1293", True)
                            Next
                            If Flag = True Then
                                objMatrix.AddRow(1, objMatrix.VisualRowCount)
                                Me.SetNewLine(FormUID, objMatrix.VisualRowCount, objMatrix)
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
            If pVal.BeforeAction = False Then
                Dim oForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "GEN_ASSORTMENT"
                        Me.CreateForm()
                    Case "1282"
                        If oForm.TypeEx = "GEN_ASSORTMENT" Then
                            Me.SetDefault(oForm.UniqueID)
                            Dim sCode As String
                            sCode = oUtilities.keygencode("@GEN_ASSORTMENT")
                            oForm.Items.Item("code").Specific.value = sCode
                        End If
                    Case "1281"
                        If oForm.TypeEx = "GEN_ASSORTMENT" Then
                            Me.SetDefault(oForm.UniqueID)
                            oForm.EnableMenu("1282", True)
                        End If
                    Case "1288", "1289", "1290", "1291"
                        If oForm.TypeEx = "GEN_ASSORTMENT" Then
                            oForm.EnableMenu("1282", True)
                        End If
                    Case "1293"
                        If oForm.TypeEx = "GEN_ASSORTMENT" Then
                            If ITEM_ID.Equals("SizeMatrix") = True Then
                                objMatrix = oForm.Items.Item("SizeMatrix").Specific
                                oDBs_Detail = oForm.DataSources.DBDataSources.Item("@GEN_ASSORTMENT_D0")
                                For Row As Integer = 1 To objMatrix.VisualRowCount
                                    objMatrix.GetLineData(Row)
                                    oDBs_Detail.Offset = Row - 1
                                    oDBs_Detail.SetValue("LineId", oDBs_Detail.Offset, Row)
                                    oDBs_Detail.SetValue("u_size", oDBs_Detail.Offset, objMatrix.Columns.Item("size").Cells.Item(Row).Specific.value)
                                    objMatrix.SetLineData(Row)
                                Next
                                objMatrix.FlushToDataSource()
                                oDBs_Detail.RemoveRecord(oDBs_Detail.Size - 1)
                                objMatrix.LoadFromDataSource()
                            End If
                        End If

                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            If eventInfo.Row > 0 Then
                ITEM_ID = eventInfo.ItemUID
            Else
                ITEM_ID = ""
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.ActionSuccess = True Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        objMatrix = objForm.Items.Item("SizeMatrix").Specific
                        objMatrix.AddRow()
                        objMatrix.FlushToDataSource()
                        Me.SetNewLine(BusinessObjectInfo.FormUID, objMatrix.VisualRowCount, objMatrix)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = True Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        objMatrix = objForm.Items.Item("SizeMatrix").Specific
                        objMatrix.DeleteRow(objMatrix.VisualRowCount)
                        objMatrix.FlushToDataSource()
                    ElseIf BusinessObjectInfo.ActionSuccess = True Then
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Then
                            objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                            objMatrix = objForm.Items.Item("SizeMatrix").Specific
                            objMatrix.AddRow()
                            objMatrix.FlushToDataSource()
                            Me.SetNewLine(BusinessObjectInfo.FormUID, objMatrix.VisualRowCount, objMatrix)
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

End Class


