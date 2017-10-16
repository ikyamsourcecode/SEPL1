Public Class ClsSAMRevaluation

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim objEdit As SAPbouiCOM.EditText
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim objMatrix1 As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail1 As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim oRS As SAPbobsCOM.Recordset
    Dim oRSet As SAPbobsCOM.Recordset
    Dim ROW_ID As Integer = 0
    Dim ITEM_ID As String
    Dim RowCount As Integer
    Dim AlertWhs As String
    Dim AlertDocNum As String
    Dim enableflag As Boolean = False
    Dim PrevDocNo As String
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            oUtilities.SAPXML("SAM_REVALUATION.xml")
            objForm = oApplication.Forms.GetForm("SAM", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SAM_REV")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SAM_REV_D0")
            objForm.EnableMenu("1282", False)
            objForm.EnableMenu("1284", True)
            objForm.EnableMenu("1288", True)
            objForm.EnableMenu("1289", True)
            objForm.EnableMenu("1290", True)
            objForm.EnableMenu("1291", True)
            objForm.DataBrowser.BrowseBy = "docnum"
            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Unit").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoFormMode.fm_OK_MODE, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Year").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoFormMode.fm_OK_MODE, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("prd").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoFormMode.fm_OK_MODE, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("month").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoFormMode.fm_OK_MODE, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("gen").Enabled = False
            'objForm.Items.Item("exe").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoFormMode.fm_OK_MODE, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            SetDefault(objForm.UniqueID)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.ActiveForm()
            objForm.Freeze(True)
            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                objForm.EnableMenu("1282", False)
            End If

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SAM_REV")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SAM_REV_D0")
            ' oUtilities.GetSeries(FormUID, "series", "GEN_COST_SHEET")
            'oDBs_Head.SetValue("U_Stat", 0, "Open")
            oDBs_Head.SetValue("DocNum", 0, objForm.BusinessObject.GetNextSerialNumber("Primary", "GEN_SAM_REV"))
            oDBs_Head.SetValue("U_DocDate", 0, DateTime.Today.ToString("yyyyMMdd"))
            objMatrix = objForm.Items.Item("mtr").Specific
            'objMatrix.AddRow(1, objMatrix.VisualRowCount)
            'Me.SetNewLine1(FormUID, objMatrix.VisualRowCount,bjMatrix)docnum
            'objEdit = objForm.Items.Item("stat").Specific
            ' objEdit.Value = "Open
            objEdit = objForm.Items.Item("docnum").Specific
            objEdit.Value = objForm.BusinessObject.GetNextSerialNumber("Primary", "GEN_SAM_REV")
            oDBs_Head.SetValue("U_Stat", 0, "Open")
            objMatrix.AutoResizeColumns()
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("select Code from OFPR A0 inner join OACP A1 on A0.Category=A1.PeriodCat where A1.Year= DATEPART(Year,GETDATE())-1")
            oRS.DoQuery("select A1.Year from  OACP A1")
            objCombo = objForm.Items.Item("Year").Specific
            If objCombo.ValidValues.Count = 0 Then
                For i As Integer = 1 To oRS.RecordCount
                    objCombo.ValidValues.Add(Trim(oRS.Fields.Item("year").Value), Trim(oRS.Fields.Item("year").Value))
                    oRS.MoveNext()
                Next
            End If
            objForm.Items.Item("gen").Enabled = False
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

   
    Sub SetNewLine1(ByVal FormUID As String, ByVal Row As Integer, ByRef oMatrix As SAPbouiCOM.Matrix)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@GEN_SAM_REV_D0")
            objMatrix = oMatrix
            objMatrix.FlushToDataSource()
            oDBs_Detail1.Offset = Row - 1
            oDBs_Detail1.SetValue("LineID", oDBs_Detail1.Offset, objMatrix.VisualRowCount)
            oDBs_Detail1.SetValue("U_ItmCode", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_ItmName", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_CngSAM", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_SthSAM", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_FinSAM", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_Selprc", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_CapCtg", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_Capstg", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_CapFng", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_CTQty", oDBs_Detail1.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try
            objForm = oApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET")
            oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET_D0")
            oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@GEN_COST_SHEET_D1")
            If Trim(oDBs_Head.GetValue("u_doccur", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter document currency", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            If Trim(oDBs_Head.GetValue("u_itemcode", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter Style code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            If Trim(oDBs_Head.GetValue("u_docrate", 0)) = "" Then
                oApplication.StatusBar.SetText("Please enter document rate", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If oDBs_Head.GetValue("u_mtotal", 0) = 0 Then
                oApplication.StatusBar.SetText("Please enter value for items in rows and click refresh", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            objMatrix = objForm.Items.Item("mtx1").Specific
            If Trim(objMatrix.Columns.Item("itemcode").Cells.Item(1).Specific.value) = "" Then
                oApplication.StatusBar.SetText("Please enter items", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            objMatrix1 = objForm.Items.Item("mtx2").Specific
            If Trim(objMatrix1.Columns.Item("prcs").Cells.Item(1).Specific.value) = "" Then
                oApplication.StatusBar.SetText("Please enter process", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSet.DoQuery("Select DocNum From [@GEN_COST_SHEET] Where u_itemcode = '" + Trim(objForm.Items.Item("itemcode").Specific.value) + "' ANd DocNum <> '" + Trim(objForm.Items.Item("docnum").Specific.value) + "' And Isnull(u_final,'N') = 'Y'")
            If oRSet.RecordCount > 0 Then
                oApplication.StatusBar.SetText("Cost sheet is finalized for this style", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function
    Function Disable_form(ByVal FormUID As String) As Boolean
        objForm = oApplication.Forms.ActiveForm
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SAM_REV")
        oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SAM_REV_D0")
        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRS.DoQuery("select Ref2 from OMRV where Ref2 ='" & oDBs_Head.GetValue("DocEntry", 0) & "'  ")
        'If oDBs_Head.GetValue("Canceled", 0) <> "N" Then
        '    Return True
        'End If
        If oDBs_Head.GetValue("U_stat", 0).Trim() = "Generated" Then
            Return True
        End If
        '
        If oRS.RecordCount > 0 Then
            objForm.Items.Item("exe").Enabled = False
        End If
        oRS.DoQuery("select U_DocDate,U_ItmCode ,U_CTQty ,U_CTRev  from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_CTQty >0 and isnull(U_CTRefNo ,'')<>'' and isnull(U_select,'')<>'Y'")
        If oRS.RecordCount = 0 Then
            Return False
        End If
        oRS.DoQuery("select U_DocDate,U_ItmCode ,U_STQty ,U_STRCst  from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_STQty >0 and isnull(U_STRefNo ,'')<>'' and isnull(U_select,'')<>'Y'")
        If oRS.RecordCount = 0 Then
            Return False
        End If
        oRS.DoQuery("select U_DocDate,U_ItmCode ,U_FIQty ,U_FinFCst  from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_FIQty >0 and isnull(U_FIRefNo ,'')<>'' and isnull(U_select,'')<>'Y' ")
        If oRS.RecordCount = 0 Then
            Return False
        End If
        oRS.DoQuery("select U_DocDate,U_ItmCode ,U_WSTQty ,U_WSTRcst  from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_WSTQty >0 and isnull(U_WSTRefNo ,'')<>'' and isnull(U_select,'')<>'Y'")
        If oRS.RecordCount = 0 Then
            Return False
        End If
        oRS.DoQuery("select U_DocDate,U_ItmCode ,U_WFIqty ,U_WFIPval  from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_WFIqty >0 and isnull(U_WFIRefNo ,'')<>'' and isnull(U_select,'')<>'Y'")
        If oRS.RecordCount = 0 Then
            Return False
        End If
        oRS.DoQuery("select U_DocDate,U_ItmCode ,U_FGQTY,U_FGPQTY  from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_FGQTY >0 and isnull(U_FGRefNo ,'')<>''  and isnull(U_select,'')<>'Y'")
        If oRS.RecordCount = 0 Then
            Return False
        End If

        oRS.DoQuery("select U_DocDate,U_ItmCode   from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_SubCut >0 and isnull(U_SCutRefNo ,'')<>''  and isnull(U_select,'')<>'Y'")
        If oRS.RecordCount = 0 Then
            Return False
        End If
        oRS.DoQuery("select U_DocDate,U_ItmCode  from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_SubSit >0 and isnull(U_SSitRefno ,'')<>''  and isnull(U_select,'')<>'Y'")
        If oRS.RecordCount = 0 Then
            Return False
        End If
        oRS.DoQuery("select U_DocDate,U_ItmCode   from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_SubFin >0 and isnull(U_SFinRefno ,'')<>''  and isnull(U_select,'')<>'Y'")
        If oRS.RecordCount = 0 Then
            Return False
        End If
        oRS.DoQuery("select U_DocDate,U_ItmCode   from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_SubFG >0 and isnull(U_SFGRefno ,'')<>''  and isnull(U_select,'')<>'Y'")
        If oRS.RecordCount = 0 Then
            Return False
        End If

        Return True
    End Function

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "exe" And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                        fillmatrix()
                    End If
                    If pVal.ItemUID = "gen" And pVal.BeforeAction = False Then
                        InventoryRevaluation()
                    End If
                    If (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.BeforeAction = False Then
                        objForm.Items.Item("fcs").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        'objForm.Items.Item("Year").Enabled = False
                        'objForm.Items.Item("month").Enabled = False
                        'objForm.Items.Item("Unit").Enabled = False
                        'objForm.Items.Item("gen").Enabled = False
                        'objForm.Items.Item("prd").Enabled = False
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE) Then
                        'objForm.Items.Item("fcs").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        'objForm.Items.Item("Year").Enabled = True
                        'objForm.Items.Item("month").Enabled = True
                        'objForm.Items.Item("Unit").Enabled = True
                        'objForm.Items.Item("prd").Enabled = True

                    End If

                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE) Then
                    '    objForm.Items.Item("fcs").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '    objForm.Items.Item("Year").Enabled = False
                    '    objForm.Items.Item("month").Enabled = False
                    '    objForm.Items.Item("Unit").Enabled = False
                    '    objForm.Items.Item("prd").Enabled = False
                    'End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True Then
                        objForm = oApplication.Forms.ActiveForm()
                        objMatrix = objForm.Items.Item("mtr").Specific
                        If objMatrix.VisualRowCount = 0 And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            BubbleEvent = False
                            oApplication.StatusBar.SetText("Document cannot be added without any records")
                        End If

                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False Then
                        If (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                            If Disable_form(objForm.TypeEx) = True Then
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                            Else
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            End If


                        End If
                    End If
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess = True And pVal.BeforeAction = False Then

                        SetDefault(objForm.UniqueID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "Year" And pVal.ActionSuccess = True Then
                        objForm = oApplication.Forms.ActiveForm()
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SAM_REV")
                        Dim oYear As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oYear.DoQuery("select Code from OFPR A0 inner join OACP A1 on A0.Category=A1.PeriodCat where A1.Year= '" + objForm.Items.Item("Year").Specific.value + "' ")
                        objCombo = objForm.Items.Item("prd").Specific
                        If objCombo.ValidValues.Count > 0 Then
                            'If objCombo.ValidValues.Count > 0 Then
                            For i As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
                                objCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                            Next
                            'End If
                        End If
                        For i As Integer = 1 To oYear.RecordCount
                            objCombo.ValidValues.Add(Trim(oYear.Fields.Item("code").Value), Trim(oYear.Fields.Item("code").Value))
                            oYear.MoveNext()
                        Next
                        oDBs_Head.SetValue("U_period", 0, "")
                    End If


                    If pVal.ItemUID = "prd" And pVal.BeforeAction = False Then
                        Dim Period As String
                        Period = objForm.Items.Item("prd").Specific.value
                        Dim oPeriod As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oPeriod.DoQuery("select DateName(month,F_RefDate) from OFPR where Code='" + Period + "'")
                        ' objForm.Items.Item("month").Sp()
                        objForm.Items.Item("month").Specific.value = oPeriod.Fields.Item(0).Value
                    End If
                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS


                Case SAPbouiCOM.BoEventTypes.et_VALIDATE



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
                    Case "1287"
                        If objForm.TypeEx = "SAM" Then
                            SetDefault(objForm.UniqueID)
                        End If
                    Case "1284"
                        If objForm.TypeEx = "SAM" Then
                            If objForm.Items.Item("stat").Specific.value = "Generated" Then
                                Throw New Exception("Cannot Cancel because the status is Generated")
                            ElseIf objForm.Items.Item("stat").Specific.value = "Open" Then
                                Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SAM_REV")
                                oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SAM_REV_D0")
                                oRS.DoQuery(" update [@GEN_SAM_REV]  set U_Stat='Cancel' where DocEntry ='" & oDBs_Head.GetValue("DocEntry", 0) & "' ")


                            End If

                        End If
                End Select
               
            End If
            If pVal.BeforeAction = False Then
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm

                Select Case objForm.TypeEx
                    Case "SAM"
                        If objForm.Mode = 3 Then
                            objMatrix = objForm.Items.Item("mtr").Specific
                            If objMatrix.VisualRowCount = 0 Then
                                SetDefault(objForm.UDFFormUID)
                            End If
                        End If
                End Select

                Select Case pVal.MenuUID
                    Case "SAM"
                        Me.CreateForm(objForm.UniqueID)
                    Case "1287"
                        If objForm.TypeEx = "SAM" Then

                        End If
                    Case "1282"
                        If objForm.TypeEx = "SAM" Then
                            Me.SetDefault(objForm.UniqueID)
                        End If
                   
                    Case "1281"
                        If objForm.TypeEx = "SAM" Then
                            objForm.EnableMenu("1282", True)
                            objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        End If
                    Case "1288", "1289", "1290", "1291"
                        If objForm.TypeEx = "SAM" Then
                            objForm.EnableMenu("1282", True)
                            objForm.Items.Item("fcs").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            ' objForm.Items.Item("Year").Enabled = False
                            'objForm.Items.Item("month").Enabled = False
                            'objForm.Items.Item("Unit").Enabled = False
                            'objForm.Items.Item("prd").Enabled = False
                            If Disable_form(objForm.TypeEx) = True Then
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                            Else
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                objForm = oApplication.Forms.ActiveForm
                                oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SAM_REV")
                                oDBs_Detail = objForm.DataSources.DBDataSources.Item("@GEN_SAM_REV_D0")
                                Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRS.DoQuery("select Ref2 from OMRV where Ref2 ='" & oDBs_Head.GetValue("DocEntry", 0) & "'  ")
                                If oRS.RecordCount > 0 Then
                                    objForm.Items.Item("exe").Enabled = False
                                End If
                                objForm.Items.Item("fcs").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                ' objForm.Items.Item("Year").Enabled = False
                                'objForm.Items.Item("month").Enabled = False
                                'objForm.Items.Item("Unit").Enabled = False
                                'objForm.Items.Item("prd").Enabled = False
                            End If

                        End If
                    Case "1293"
                        If objForm.TypeEx = "SAM" Then

                        End If

                End Select


            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            RowCount = eventInfo.Row
            If eventInfo.Row > 0 Then
                ITEM_ID = eventInfo.ItemUID
                Dim oForm As SAPbouiCOM.Form = oApplication.Forms.Item(eventInfo.FormUID)
                objMatrix = oForm.Items.Item("mtr").Specific
                If objMatrix.VisualRowCount > 1 Then
                    oForm.EnableMenu("1293", True)
                Else
                    oForm.EnableMenu("1293", False)
                End If

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
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        SetDefault(objForm.UniqueID)
                        '
                        'End If
                    ElseIf BusinessObjectInfo.ActionSuccess = True Then
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Then
                            objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                            ' objMatrix1 = objForm.Items.Item("mtx2").Specific
                            'objMatrix1.AddRow()
                            ' objMatrix1.FlushToDataSource()
                            ' Me.SetNewLine1(BusinessObjectInfo.FormUID, objMatrix1.VisualRowCount, objMatrix1)
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then
                        'objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        'objMatrix1 = objForm.Items.Item("mtx2").Specific
                        'objMatrix1.AddRow()
                        'objMatrix1.FlushToDataSource()
                        'Me.SetNewLine1(BusinessObjectInfo.FormUID, objMatrix1.VisualRowCount, objMatrix1)
                        'objForm.Items.Item("mtx1").AffectsFormMode = False
                        'objForm.Items.Item("mtx2").AffectsFormMode = False
                        'objForm.Items.Item("fldexp").AffectsFormMode = False
                        'objForm.Items.Item("flditem").AffectsFormMode = False
                        'Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'oRSet.DoQuery("Select USER_CODE From OUSR Where USER_CODE = '" + oCompany.UserName.ToString + "' And IsNull(u_cstsht,'N') = 'Y'")
                        'If oRSet.RecordCount > 0 Then
                        '    objForm.Items.Item("final").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        'End If
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub fillmatrix()
        Try

            objForm = oApplication.Forms.ActiveForm()
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SAM_REV")
            oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@GEN_SAM_REV_D0")
            objForm.Freeze(True)
            objMatrix = objForm.Items.Item("mtr").Specific
           
            If oDBs_Head.GetValue("U_Unit", 0) = "" Then
                Throw New Exception("Please select the unit before you execute")
            End If
            If oDBs_Head.GetValue("U_Period", 0) = "" Then
                Throw New Exception("Please select the Period before you execute")
            End If
            If oDBs_Head.GetValue("U_Year", 0) = "" Then
                Throw New Exception("Please select the Year before you execute")
            End If


            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("select U_period ,docentry,U_Stat from [@GEN_SAM_REV] where U_period='" & oDBs_Head.GetValue("U_Period", 0).Trim() & "' and U_Unit= '" & oDBs_Head.GetValue("U_Unit", 0).Trim() & "' and U_Stat<>'Cancel'")
            If oRS.RecordCount > 0 Then
                oRS1.DoQuery("select T0.docentry ,T0.U_Stat from [@GEN_SAM_REV_D0]  inner join  [@GEN_SAM_REV] T0 on T0.docentry=[@GEN_SAM_REV_D0].Docentry where  (isnull(U_CTRefNo,'')<>'' or isnull(U_STRefNo,'')<>'' or isnull(U_WFIRefNo,'')<>''or isnull(U_WSTRefNo,'')<>''or isnull(U_FGRefNo ,'') <>'' or isnull(U_SCutRefNo,'')<>''  or  isnull(U_SSitRefno,'')<>'' or  isnull(U_SFinRefno,'')<>'' or  isnull(U_SFGRefno,'')<>'')  and T0.DocEntry ='" & oRS.Fields.Item("docentry").Value & "'")
                If oRS1.RecordCount > 0 Then
                    Throw New Exception("Already report has been generated against this period.CR No '" & oRS.Fields.Item("docentry").Value & "'")
                End If
                If oRS.Fields.Item("docentry").Value <> oRS1.Fields.Item("docentry").Value Then
                    Throw New Exception("Already report has been generated against this period.CR No '" & oRS.Fields.Item("docentry").Value & "'")
                End If
            End If
            If oRS.RecordCount = 0 Then
                oRS.DoQuery("select Convert(Varchar,F_RefDate,112)[F_RefDate],Convert(Varchar,T_RefDate,112)[T_RefDate] from OFPR  where Code='" & oDBs_Head.GetValue("U_period", 0) & "'")
                oApplication.SetStatusBarMessage("Processing Data", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                Dim _str_string As String = "Exec [GEN_SEPL_CostRun_Upgraded_Excel_WIP_Live] '" + oRS.Fields.Item("F_RefDate").Value + "','" + oRS.Fields.Item("T_RefDate").Value + "','','" + oDBs_Head.GetValue("U_Unit", 0).Trim + "','" + oDBs_Head.GetValue("U_period", 0).Trim + "'"
                oRS1.DoQuery(_str_string)
                objMatrix.Clear()
                oDBs_Detail1.Clear()
                objMatrix.FlushToDataSource()
                For i As Integer = 1 To oRS1.RecordCount
                    objMatrix.AddRow()
                    objMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.value = objMatrix.VisualRowCount
                    objMatrix.Columns.Item("itmCode").Cells.Item(i).Specific.value = oRS1.Fields.Item("Parent").Value
                    objMatrix.Columns.Item("ItmName").Cells.Item(i).Specific.value = oRS1.Fields.Item("ItemName").Value
                    objMatrix.Columns.Item("Selprc").Cells.Item(i).Specific.value = oRS1.Fields.Item("SalePrice").Value
                    objMatrix.Columns.Item("CTQty").Cells.Item(i).Specific.value = oRS1.Fields.Item("CT qty").Value
                    objMatrix.Columns.Item("OpnCT").Cells.Item(i).Specific.value = oRS1.Fields.Item("OPENING CUTTING COST").Value 'oRS1.Fields.Item("Parent").Value
                    objMatrix.Columns.Item("CutCst").Cells.Item(i).Specific.value = oRS1.Fields.Item("MC1 perqty").Value 'oRS1.Fields.Item("Parent").Value
                    objMatrix.Columns.Item("CutPCst").Cells.Item(i).Specific.value = oRS1.Fields.Item("PC1 perQty").Value 'oRS1.Fields.Item("ItemName").Value.ToString.Trim()
                    objMatrix.Columns.Item("CTFcst").Cells.Item(i).Specific.value = oRS1.Fields.Item("CUTTING COST FINAL").Value 'oRS1.Fields.Item("Parent").Value
                    objMatrix.Columns.Item("CTRev").Cells.Item(i).Specific.value = oRS1.Fields.Item("CUTTING REV COST").Value
                    objMatrix.Columns.Item("STQty").Cells.Item(i).Specific.value = oRS1.Fields.Item("CURRENT STQTY").Value
                    objMatrix.Columns.Item("OSTQty").Cells.Item(i).Specific.value = oRS1.Fields.Item("OPENING STCOST").Value
                    objMatrix.Columns.Item("STMCst").Cells.Item(i).Specific.value = oRS1.Fields.Item("CURRENT MC2").Value
                    objMatrix.Columns.Item("STPcst").Cells.Item(i).Specific.value = oRS1.Fields.Item("CURRENT PC2 PER QTY").Value
                    objMatrix.Columns.Item("STPFcst").Cells.Item(i).Specific.value = oRS1.Fields.Item("STITCHING COST FINAL").Value
                    objMatrix.Columns.Item("STRCst").Cells.Item(i).Specific.value = oRS1.Fields.Item("STICHING REV COST").Value
                    objMatrix.Columns.Item("WSTQty").Cells.Item(i).Specific.value = oRS1.Fields.Item("WST QTY").Value
                    objMatrix.Columns.Item("MCPWST").Cells.Item(i).Specific.value = oRS1.Fields.Item("MCPerqty-WST").Value
                    objMatrix.Columns.Item("WSTPQty").Cells.Item(i).Specific.value = oRS1.Fields.Item("WST Pcper qty").Value
                    If oRS1.Fields.Item("WST Rev Cost").Value > oRS1.Fields.Item("SalePrice").Value Then
                        objMatrix.Columns.Item("WSTRcst").Cells.Item(i).Specific.value = oRS1.Fields.Item("SalePrice").Value
                    Else
                        objMatrix.Columns.Item("WSTRcst").Cells.Item(i).Specific.value = oRS1.Fields.Item("WST Rev Cost").Value
                    End If
                    objMatrix.Columns.Item("WSTVal").Cells.Item(i).Specific.value = oRS1.Fields.Item("WST Value").Value
                    objMatrix.Columns.Item("FIQty").Cells.Item(i).Specific.value = oRS1.Fields.Item("CURRENT FIQTY").Value
                    objMatrix.Columns.Item("OFIcst").Cells.Item(i).Specific.value = oRS1.Fields.Item("OPENING FICOST").Value
                    objMatrix.Columns.Item("FinMCst").Cells.Item(i).Specific.value = oRS1.Fields.Item("CURRENT MC3").Value
                    objMatrix.Columns.Item("FinPCst").Cells.Item(i).Specific.value = oRS1.Fields.Item("CURRENT PC3 PER QTY").Value
                    objMatrix.Columns.Item("FinFCst").Cells.Item(i).Specific.value = oRS1.Fields.Item("FINISHING COST FINAL").Value
                    objMatrix.Columns.Item("FinRCst").Cells.Item(i).Specific.value = oRS1.Fields.Item("FINISHING REV FINAL").Value
                    objMatrix.Columns.Item("WFIqty").Cells.Item(i).Specific.value = oRS1.Fields.Item("WFI QTY").Value
                    objMatrix.Columns.Item("WFIPVal").Cells.Item(i).Specific.value = oRS1.Fields.Item("WFI Perqty").Value
                    objMatrix.Columns.Item("WFIVal").Cells.Item(i).Specific.value = oRS1.Fields.Item("WFI QTY").Value * oRS1.Fields.Item("WFI Perqty").Value
                    objMatrix.Columns.Item("FGQTY").Cells.Item(i).Specific.value = oRS1.Fields.Item("FG QTY").Value
                    objMatrix.Columns.Item("FGPQTY").Cells.Item(i).Specific.value = oRS1.Fields.Item("FG Perqty").Value
                    objMatrix.Columns.Item("FGVal").Cells.Item(i).Specific.value = oRS1.Fields.Item("FG QTY").Value * oRS1.Fields.Item("FG Perqty").Value

                    objMatrix.Columns.Item("MC1val").Cells.Item(i).Specific.value = oRS1.Fields.Item("MC1 value").Value
                    objMatrix.Columns.Item("PC1val").Cells.Item(i).Specific.value = oRS1.Fields.Item("PC1 VALUE").Value
                    objMatrix.Columns.Item("CCCst").Cells.Item(i).Specific.value = oRS1.Fields.Item("CURRENT CUTTING COST").Value
                    objMatrix.Columns.Item("CCFnl").Cells.Item(i).Specific.value = oRS1.Fields.Item("CUTTING COST FINAL").Value
                    objMatrix.Columns.Item("CCSprc").Cells.Item(i).Specific.value = oRS1.Fields.Item("Cutting Cost From Sales Price").Value
                    objMatrix.Columns.Item("CMPCST").Cells.Item(i).Specific.value = oRS1.Fields.Item("CURRENT MC+PC FOR ST").Value
                    objMatrix.Columns.Item("CSTQty").Cells.Item(i).Specific.value = oRS1.Fields.Item("STICHING TRIMS PER QTY").Value
                    objMatrix.Columns.Item("MC2Val").Cells.Item(i).Specific.value = oRS1.Fields.Item("MC2 VALUE").Value
                    objMatrix.Columns.Item("PC2Val").Cells.Item(i).Specific.value = oRS1.Fields.Item("PC2 VALUE").Value

                    objMatrix.Columns.Item("SSTQty").Cells.Item(i).Specific.value = oRS1.Fields.Item("Subcontract ST-Qty").Value
                    objMatrix.Columns.Item("STSChgs").Cells.Item(i).Specific.value = oRS1.Fields.Item("ST-Subcontract charges").Value
                    objMatrix.Columns.Item("STSPQty").Cells.Item(i).Specific.value = oRS1.Fields.Item("ST-Subcontract PerQty").Value
                    objMatrix.Columns.Item("CRTSTCst").Cells.Item(i).Specific.value = oRS1.Fields.Item("CURRENT ST COST").Value
                    objMatrix.Columns.Item("CRTSTVal").Cells.Item(i).Specific.value = oRS1.Fields.Item("CURRENT ST VALUE").Value
                    objMatrix.Columns.Item("STCSP").Cells.Item(i).Specific.value = oRS1.Fields.Item("Stitching  Cost From Sales Price").Value
                    objMatrix.Columns.Item("WSTSTrm").Cells.Item(i).Specific.value = oRS1.Fields.Item("WST-Stitch Trims").Value
                    objMatrix.Columns.Item("CMPFFI").Cells.Item(i).Specific.value = oRS1.Fields.Item("CURRENT MC+PC FOR FI").Value
                    objMatrix.Columns.Item("FINTQty").Cells.Item(i).Specific.value = oRS1.Fields.Item("FINISHING TRIMS PER QTY").Value
                    objMatrix.Columns.Item("SCFIQty").Cells.Item(i).Specific.value = oRS1.Fields.Item("Subcontract FI-Qty").Value
                    objMatrix.Columns.Item("FISCC").Cells.Item(i).Specific.value = oRS1.Fields.Item("FI-Subcontract charges").Value
                    objMatrix.Columns.Item("FISPQty").Cells.Item(i).Specific.value = oRS1.Fields.Item("FI-Subcontract PerQty").Value
                    objMatrix.Columns.Item("CurFICst").Cells.Item(i).Specific.value = oRS1.Fields.Item("CURRENT FI COST").Value
                    objMatrix.Columns.Item("CurFIVal").Cells.Item(i).Specific.value = oRS1.Fields.Item("CURRENT FI VALUE").Value
                    objMatrix.Columns.Item("FiCstSP").Cells.Item(i).Specific.value = oRS1.Fields.Item("Finishing Cost From Sales Price").Value
                    objMatrix.Columns.Item("WFinTrm").Cells.Item(i).Specific.value = oRS1.Fields.Item("WFI-Finsihing Trims").Value
                    objMatrix.Columns.Item("SubFGqty").Cells.Item(i).Specific.value = oRS1.Fields.Item("Subcontract FG-Qty").Value
                    objMatrix.Columns.Item("FGSChrgs").Cells.Item(i).Specific.value = oRS1.Fields.Item("FG-Subcontract charges").Value
                    objMatrix.Columns.Item("FGSpQty").Cells.Item(i).Specific.value = oRS1.Fields.Item("FG-Subcontract PerQty").Value

                    objMatrix.Columns.Item("SubCut").Cells.Item(i).Specific.value = oRS1.Fields.Item("SUB_CUTTING_QTY").Value
                    objMatrix.Columns.Item("SubSit").Cells.Item(i).Specific.value = oRS1.Fields.Item("SUB_STICHING_QTY").Value
                    objMatrix.Columns.Item("SubFin").Cells.Item(i).Specific.value = oRS1.Fields.Item("SUB_FINISH_QTY").Value
                    objMatrix.Columns.Item("SubFG").Cells.Item(i).Specific.value = oRS1.Fields.Item("SUB_FG_QTY").Value
                    ' FiCstSP

                    oRS1.MoveNext()
                Next
                'Dim h As Integer = 130 + oRS1.RecordCount
                'For j As Integer = oRS1.RecordCount + 1 To h
                '    objMatrix.AddRow()
                '    objMatrix.Columns.Item("V_-1").Cells.Item(j).Specific.value = objMatrix.VisualRowCount
                'Next
                oApplication.SetStatusBarMessage("Process over", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            Else
                Throw New Exception("Already report has been generated against this period.CR No '" & oRS.Fields.Item("docentry").Value & "'")
            End If
            objForm.Items.Item("gen").Enabled = False
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub InventoryRevaluation()
        Try
            objForm = oApplication.Forms.ActiveForm()
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_SAM_REV")
            oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@GEN_SAM_REV_D0")
            objForm.Freeze(True)
            objMatrix = objForm.Items.Item("mtr").Specific
            objMatrix.FlushToDataSource()
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRS1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRS2 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oSR As SAPbobsCOM.MaterialRevaluation
            Dim oSR_Lines As SAPbobsCOM.MaterialRevaluation_lines
            oSR = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
            Dim _str_whs_ As String = ""
            Dim _str_Account As String = ""
            Dim _str_flag = "Y"
            Dim _int_cnt As Integer = 0
            '******************Cutting*********************
            oRS.DoQuery("select U_DocDate,U_ItmCode ,U_CTQty ,U_CTRev  from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_CTQty >0 and isnull(U_CTRefNo,'')='' ")
            If oRS.RecordCount > 0 Then

                If oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT1" Then
                    _str_whs_ = "CT-1"
                    _str_Account = "_SYS00000001447"
                ElseIf oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT2" Then
                    _str_whs_ = "CT-2"
                    _str_Account = "_SYS00000001448"
                Else
                    'If oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT3" Then
                    _str_whs_ = "CT-3"
                    _str_Account = "_SYS00000001449"
                    'Else
                    '    _str_whs_ = "CTLG-1"
                    '    _str_Account = "_SYS00000001449"
                End If
                ' Dim _dcl_ctrev As Decimal = Decimal.Round(oRS2.Fields.Item("AvgPrice").Value, 2)
                For I As Int16 = 0 To oRS.RecordCount - 1
                    oRS2.DoQuery("select AvgPrice from OITW where ItemCode ='" & oRS.Fields.Item("U_ItmCode").Value + "-1" & "' and WhsCode ='" & _str_whs_ & "'")
                    If Decimal.Round(oRS2.Fields.Item("AvgPrice").Value, 1) <> Decimal.Round(oRS.Fields.Item("U_CTRev").Value, 1) Then
                        _str_flag = "N"
                        oSR_Lines = oSR.Lines
                        oSR_Lines.SetCurrentLine(_int_cnt)
                        oSR_Lines.ItemCode = oRS.Fields.Item("U_ItmCode").Value + "-1"
                        oSR_Lines.WarehouseCode = _str_whs_
                        oSR_Lines.RevaluationIncrementAccount = _str_Account
                        oSR_Lines.RevaluationDecrementAccount = _str_Account
                        oSR_Lines.Price = oRS.Fields.Item("U_CTRev").Value
                        oSR_Lines.Add()
                        oSR.DocDate = oRS.Fields.Item("U_DocDate").Value
                        oSR.TaxDate = oRS.Fields.Item("U_DocDate").Value
                        _int_cnt = _int_cnt + 1
                    End If
                    oRS.MoveNext()
                Next
                If _int_cnt <> 0 Then

                    oSR.Reference2 = oDBs_Head.GetValue("DocNum", 0)
                    oSR.Comments = "Cutting  stock is revaluated as per cost run report for the month of " + oDBs_Head.GetValue("U_Month", 0) + oDBs_Head.GetValue("U_Year", 0)
                    If oSR.Add <> 0 Then
                        _str_flag = "N"
                        'Dim Str As String = oCompany.GetLastErrorDescription
                        ' oApplication.MessageBox(oCompany.GetLastErrorDescription + "-" + "Cutting")
                        Throw New Exception(oCompany.GetLastErrorDescription + "-" + "Cutting")
                    Else
                        Dim docEntry As String = ""
                        oCompany.GetNewObjectCode(docEntry)
                        oRS1.DoQuery("update [@GEN_SAM_REV_D0] set U_CTRefNo='" & docEntry & "'  where DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_CTQty>0 ")
                        _str_flag = "Y"
                        oApplication.SetStatusBarMessage("Sucessfully Posted-Cutting", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    End If
                End If
                _str_flag = "Y"
            End If
            '******************Cutting*********************
            '******************Stiching*********************
            oSR = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
            oRS.DoQuery("select U_DocDate,U_ItmCode ,U_STQty ,U_STRCst  from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_STQty >0  and isnull(U_STRefNo,'')=''")
            If oRS.RecordCount > 0 Then
                If oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT1" Then
                    _str_whs_ = "ST-1"
                    _str_Account = "_SYS00000001447"
                ElseIf oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT2" Then
                    _str_whs_ = "ST-2"
                    _str_Account = "_SYS00000001448"
                Else
                    _str_whs_ = "ST-3"
                    _str_Account = "_SYS00000001449"
                End If
                _int_cnt = 0
                For I As Int16 = 0 To oRS.RecordCount - 1
                    oRS2.DoQuery("select AvgPrice from OITW where ItemCode ='" & oRS.Fields.Item("U_ItmCode").Value + "-2" & "' and WhsCode ='" & _str_whs_ & "'")
                    If Decimal.Round(oRS2.Fields.Item("AvgPrice").Value, 1) <> Decimal.Round(oRS.Fields.Item("U_STRCst").Value, 1) Then
                        _str_flag = "N"
                        oSR_Lines = oSR.Lines
                        oSR_Lines.SetCurrentLine(_int_cnt)
                        oSR_Lines.ItemCode = oRS.Fields.Item("U_ItmCode").Value + "-2"
                        oSR_Lines.WarehouseCode = _str_whs_
                        oSR_Lines.RevaluationIncrementAccount = _str_Account
                        oSR_Lines.RevaluationDecrementAccount = _str_Account
                        oSR_Lines.Price = oRS.Fields.Item("U_STRCst").Value
                        oSR_Lines.Add()
                        oSR.DocDate = oRS.Fields.Item("U_DocDate").Value
                        oSR.TaxDate = oRS.Fields.Item("U_DocDate").Value
                        _int_cnt = _int_cnt + 1
                    End If
                    oRS.MoveNext()
                Next
                If _int_cnt <> 0 Then
                    oSR.Reference2 = oDBs_Head.GetValue("DocNum", 0)
                    oSR.Comments = "Stiching  stock is revaluated as per cost run report for the month of " + oDBs_Head.GetValue("U_Month", 0) + oDBs_Head.GetValue("U_Year", 0)
                    If oSR.Add <> 0 Then
                        ' Dim Str As String = oCompany.GetLastErrorDescription
                        _str_flag = "N"
                        ' oApplication.MessageBox(oCompany.GetLastErrorDescription + "-" + "Stiching")
                        Throw New Exception(oCompany.GetLastErrorDescription + "-" + "Stiching")

                    Else
                        Dim docEntry As String = ""
                        oCompany.GetNewObjectCode(docEntry)
                        oRS1.DoQuery("update [@GEN_SAM_REV_D0] set U_STRefNo='" & docEntry & "'  where DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_STQty >0 ")
                        _str_flag = "Y"
                        oApplication.SetStatusBarMessage("Sucessfully Posted-Stiching", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    End If
                End If
                _str_flag = "Y"
            End If
            '******************Stiching*********************

            '******************FI *********************
            oSR = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
            oRS.DoQuery("select U_DocDate,U_ItmCode ,U_FIQty ,U_FinRCst  from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_FIQty >0  and isnull(U_FIRefNo,'')='' ")
            If oRS.RecordCount > 0 Then
                If oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT1" Then
                    _str_whs_ = "FI-1"
                    _str_Account = "_SYS00000001447"
                ElseIf oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT2" Then
                    _str_whs_ = "FI-2"
                    _str_Account = "_SYS00000001448"
                Else
                    _str_whs_ = "FI-3"
                    _str_Account = "_SYS00000001449"
                End If
                _int_cnt = 0
                For I As Int16 = 0 To oRS.RecordCount - 1
                    oRS2.DoQuery("select AvgPrice from OITW where ItemCode ='" & oRS.Fields.Item("U_ItmCode").Value + "-3" & "' and WhsCode ='" & _str_whs_ & "'")
                    If Decimal.Round(oRS2.Fields.Item("AvgPrice").Value, 1) <> Decimal.Round(oRS.Fields.Item("U_FinRCst").Value, 1) Then
                        _str_flag = "N"
                        oSR_Lines = oSR.Lines
                        oSR_Lines.SetCurrentLine(_int_cnt)
                        oSR_Lines.ItemCode = oRS.Fields.Item("U_ItmCode").Value + "-3"
                        oSR_Lines.WarehouseCode = _str_whs_
                        oSR_Lines.RevaluationIncrementAccount = _str_Account
                        oSR_Lines.RevaluationDecrementAccount = _str_Account
                        oSR_Lines.Price = oRS.Fields.Item("U_FinRCst").Value
                        oSR_Lines.Add()
                        oSR.DocDate = oRS.Fields.Item("U_DocDate").Value
                        oSR.TaxDate = oRS.Fields.Item("U_DocDate").Value
                        _int_cnt = _int_cnt + 1
                    End If
                    oRS.MoveNext()
                Next
                If _int_cnt <> 0 Then

                    oSR.Reference2 = oDBs_Head.GetValue("DocNum", 0)
                    oSR.Comments = "Finishing  stock is revaluated as per cost run report for the month of " + oDBs_Head.GetValue("U_Month", 0) + oDBs_Head.GetValue("U_Year", 0)
                    If oSR.Add <> 0 Then
                        'Dim Str As String = oCompany.GetLastErrorDescription
                        _str_flag = "N"
                        ' oApplication.MessageBox(oCompany.GetLastErrorDescription + "-" + "FI")
                        Throw New Exception(oCompany.GetLastErrorDescription + "-" + "FI")
                    Else
                        Dim docEntry As String = ""
                        _str_flag = "Y"
                        oCompany.GetNewObjectCode(docEntry)
                        oRS1.DoQuery("update [@GEN_SAM_REV_D0] set U_FIRefNo='" & docEntry & "'  where DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_FIQty >0 ")
                        oApplication.SetStatusBarMessage("Sucessfully Posted-FI", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    End If
                End If
            Else
                _str_flag = "Y"
            End If
            '******************FI *********************
            '******************WST*********************
            oSR = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
            oRS.DoQuery("select U_DocDate,U_ItmCode ,U_WSTQty ,U_WSTRcst  from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_WSTQty >0 and isnull(U_WSTRefNo,'')=''")
            If oRS.RecordCount > 0 Then

                If oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT1" Then
                    _str_whs_ = "WST-1"
                    _str_Account = "_SYS00000001447"
                ElseIf oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT2" Then
                    _str_whs_ = "WST-2"
                    _str_Account = "_SYS00000001448"
                Else
                    _str_whs_ = "WST-3"
                    _str_Account = "_SYS00000001449"
                End If
                Dim _int_val As Integer = 0
                ' Dim _dcl_ctrev As Decimal = Decimal.Round(oRS2.Fields.Item("AvgPrice").Value, 2)
                For I As Int16 = 0 To oRS.RecordCount - 1
                    oRS2.DoQuery("select AvgPrice from OITW where ItemCode ='" & oRS.Fields.Item("U_ItmCode").Value + "-1" & "' and WhsCode ='" & _str_whs_ & "'")
                    If Decimal.Round(oRS2.Fields.Item("AvgPrice").Value, 1) <> Decimal.Round(oRS.Fields.Item("U_WSTRcst").Value, 1) Then
                        _str_flag = "N"
                        oSR_Lines = oSR.Lines
                        oSR_Lines.SetCurrentLine(_int_val)
                        oSR_Lines.ItemCode = oRS.Fields.Item("U_ItmCode").Value + "-1"
                        oSR_Lines.WarehouseCode = _str_whs_
                        oSR_Lines.RevaluationIncrementAccount = _str_Account
                        oSR_Lines.RevaluationDecrementAccount = _str_Account
                        oSR_Lines.Price = oRS.Fields.Item("U_WSTRcst").Value
                        oSR_Lines.Add()
                        oSR.DocDate = oRS.Fields.Item("U_DocDate").Value
                        oSR.TaxDate = oRS.Fields.Item("U_DocDate").Value
                        _int_val = _int_val + 1
                    End If
                    oRS.MoveNext()
                Next
                If _int_val <> 0 Then

                    oSR.Reference2 = oDBs_Head.GetValue("DocNum", 0)
                    oSR.Comments = "Cutting  stock is revaluated as per cost run report for the month of " + oDBs_Head.GetValue("U_Month", 0) + oDBs_Head.GetValue("U_Year", 0)
                    If oSR.Add <> 0 Then
                        _str_flag = "N"
                        'Dim Str As String = oCompany.GetLastErrorDescription
                        ' oApplication.MessageBox(oCompany.GetLastErrorDescription + "-" + "Cutting")
                        Throw New Exception(oCompany.GetLastErrorDescription + "-" + "Cutting")
                    Else
                        Dim docEntry As String = ""
                        oCompany.GetNewObjectCode(docEntry)
                        oRS1.DoQuery("update [@GEN_SAM_REV_D0] set U_WSTRefNo='" & docEntry & "'  where DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_CTQty>0 ")
                        _str_flag = "Y"
                        oApplication.SetStatusBarMessage("Sucessfully Posted-WST", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    End If
                End If
                _str_flag = "Y"
            End If
            '******************WST*********************
            '******************WFI*********************
            oSR = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
            oRS.DoQuery("select U_DocDate,U_ItmCode ,U_WFIqty ,U_WFIPval  from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_WFIqty >0 and isnull(U_WFIRefNo,'')='' ")
            If oRS.RecordCount > 0 Then
                If oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT1" Then
                    _str_whs_ = "WFI-1"
                    _str_Account = "_SYS00000001447"
                ElseIf oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT2" Then
                    _str_whs_ = "WFI-2"
                    _str_Account = "_SYS00000001448"
                Else
                    _str_whs_ = "WFI-3"
                    _str_Account = "_SYS00000001449"
                End If
                _int_cnt = 0
                For I As Int16 = 0 To oRS.RecordCount - 1
                    oRS2.DoQuery("select AvgPrice from OITW where ItemCode ='" & oRS.Fields.Item("U_ItmCode").Value + "-2" & "' and WhsCode ='" & _str_whs_ & "'")
                    If Decimal.Round(oRS2.Fields.Item("AvgPrice").Value, 1) <> Decimal.Round(oRS.Fields.Item("U_WFIPval").Value, 1) Then
                        _str_flag = "N"
                        oSR_Lines = oSR.Lines
                        oSR_Lines.SetCurrentLine(_int_cnt)
                        oSR_Lines.ItemCode = oRS.Fields.Item("U_ItmCode").Value + "-2"
                        oSR_Lines.WarehouseCode = _str_whs_
                        oSR_Lines.RevaluationIncrementAccount = _str_Account
                        oSR_Lines.RevaluationDecrementAccount = _str_Account
                        oSR_Lines.Price = oRS.Fields.Item("U_WFIPval").Value
                        oSR_Lines.Add()
                        oSR.DocDate = oRS.Fields.Item("U_DocDate").Value
                        oSR.TaxDate = oRS.Fields.Item("U_DocDate").Value
                        _int_cnt = _int_cnt + 1
                    End If
                    oRS.MoveNext()

                Next
                If _int_cnt <> 0 Then
                    oSR.Reference2 = oDBs_Head.GetValue("DocNum", 0)
                    oSR.Comments = "Stiching  stock is revaluated as per cost run report for the month of " + oDBs_Head.GetValue("U_Month", 0) + oDBs_Head.GetValue("U_Year", 0)
                    If oSR.Add <> 0 Then
                        'Dim Str As String = oCompany.GetLastErrorDescription
                        _str_flag = "N"
                        'oApplication.MessageBox(oCompany.GetLastErrorDescription + "-" + "WFI")
                        Throw New Exception(oCompany.GetLastErrorDescription + "-" + "WFI")
                    Else
                        Dim docEntry As String = ""
                        oCompany.GetNewObjectCode(docEntry)
                        oRS1.DoQuery("update [@GEN_SAM_REV_D0] set U_WFIRefNo='" & docEntry & "'  where DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_WFIqty >0 ")
                        _str_flag = "Y"
                        oApplication.SetStatusBarMessage("Sucessfully Posted-WFI", SAPbouiCOM.BoMessageTime.bmt_Medium, False)

                    End If
                End If
            Else
                _str_flag = "Y"
            End If
            '******************WFI*********************

            '******************FG*********************
            oSR = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
            oRS.DoQuery("select U_DocDate,U_ItmCode ,U_FGQTY,U_FGPQTY  from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_FGQTY >0  and isnull(U_FGRefNo,'')=''")
            If oRS.RecordCount > 0 Then
                If oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT1" Then
                    _str_whs_ = "FG-1"
                    _str_Account = "_SYS00000001447"
                ElseIf oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT2" Then
                    _str_whs_ = "FG-2"
                    _str_Account = "_SYS00000001448"
                Else
                    _str_whs_ = "FG-3"
                    _str_Account = "_SYS00000001449"
                End If
                _int_cnt = 0
                For I As Int16 = 0 To oRS.RecordCount - 1
                    oRS2.DoQuery("select AvgPrice from OITW where ItemCode ='" & oRS.Fields.Item("U_ItmCode").Value & "' and WhsCode ='" & _str_whs_ & "'")
                    If Decimal.Round(oRS2.Fields.Item("AvgPrice").Value, 1) <> Decimal.Round(oRS.Fields.Item("U_FGPQTY").Value, 1) Then
                        _str_flag = "N"
                        oSR_Lines = oSR.Lines
                        oSR_Lines.SetCurrentLine(_int_cnt)
                        oSR_Lines.ItemCode = oRS.Fields.Item("U_ItmCode").Value
                        oSR_Lines.WarehouseCode = _str_whs_
                        oSR_Lines.RevaluationIncrementAccount = _str_Account
                        oSR_Lines.RevaluationDecrementAccount = _str_Account
                        oSR_Lines.Price = oRS.Fields.Item("U_FGPQTY").Value
                        oSR_Lines.Add()
                        oSR.DocDate = oRS.Fields.Item("U_DocDate").Value
                        oSR.TaxDate = oRS.Fields.Item("U_DocDate").Value
                        _int_cnt = _int_cnt + 1
                    End If
                    oRS.MoveNext()
                Next
                If _int_cnt <> 0 Then
                    oSR.Reference2 = oDBs_Head.GetValue("DocNum", 0)
                    oSR.Comments = "FG  stock is revaluated as per cost run report for the month of " + oDBs_Head.GetValue("U_Month", 0) + oDBs_Head.GetValue("U_Year", 0)
                    If oSR.Add <> 0 Then
                        'Dim Str As String = oCompany.GetLastErrorDescription
                        _str_flag = "N"
                        'oApplication.MessageBox(oCompany.GetLastErrorDescription + "-" + "FG")
                        Throw New Exception(oCompany.GetLastErrorDescription + "-" + "FG")
                    Else
                        _str_flag = "Y"
                        Dim docEntry As String = ""
                        oCompany.GetNewObjectCode(docEntry)
                        oRS1.DoQuery("update [@GEN_SAM_REV_D0] set U_FGRefNo='" & docEntry & "'  where DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_FGQTY >0 ")
                        oApplication.SetStatusBarMessage("Sucessfully Posted-FG", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    End If
                End If
            Else
                _str_flag = "Y"
            End If
            '******************FG*********************
            '******************Sub Contractor Cutting*********************
            oSR = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
            oRS.DoQuery("select U_DocDate,U_ItmCode ,U_SubCut ,U_CTRev  from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_SubCut >0 and isnull(U_SCutRefNo,'')='' ")
            Dim oRS3 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oRS.RecordCount > 0 Then

                If oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT1" Then
                    _str_Account = "_SYS00000001447"
                    _str_whs_ = "JW%"
                ElseIf oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT2" Then
                    _str_Account = "_SYS00000001448"
                    _str_whs_ = "2JW%"
                Else
                    _str_Account = "_SYS00000001449"
                    _str_whs_ = "3JW%"
                End If
                ' Dim _dcl_ctrev As Decimal = Decimal.Round(oRS2.Fields.Item("AvgPrice").Value, 2)
                _int_cnt = 0
                For j As Int16 = 0 To oRS.RecordCount - 1

                    oRS3.DoQuery("select OnHand,ItemCode,WhsCode from OITW where  OnHand>0 and ItemCode= '" & oRS.Fields.Item("U_ItmCode").Value.ToString.Trim() + "-1" & "' and  WhsCode like '" & _str_whs_ & "'")
                    For I As Int16 = 0 To oRS3.RecordCount - 1
                        oRS2.DoQuery("select AvgPrice from OITW where ItemCode ='" & oRS.Fields.Item("U_ItmCode").Value.ToString.Trim() + "-1" & "' and WhsCode ='" & oRS3.Fields.Item("WhsCode").Value & "'")
                        If Decimal.Round(oRS2.Fields.Item("AvgPrice").Value, 1) <> Decimal.Round(oRS.Fields.Item("U_CTRev").Value, 1) Then
                            _str_flag = "N"
                            oSR_Lines = oSR.Lines
                            oSR_Lines.SetCurrentLine(_int_cnt)
                            oSR_Lines.ItemCode = oRS.Fields.Item("U_ItmCode").Value.ToString.Trim() + "-1"
                            oSR_Lines.WarehouseCode = oRS3.Fields.Item("WhsCode").Value
                            oSR_Lines.RevaluationIncrementAccount = _str_Account
                            oSR_Lines.RevaluationDecrementAccount = _str_Account
                            oSR_Lines.Price = oRS.Fields.Item("U_CTRev").Value
                            oSR_Lines.Add()
                            oSR.DocDate = oRS.Fields.Item("U_DocDate").Value
                            oSR.TaxDate = oRS.Fields.Item("U_DocDate").Value
                            _int_cnt = _int_cnt + 1
                        End If
                        oRS3.MoveNext()
                    Next
                    oRS.MoveNext()
                Next
                If _int_cnt <> 0 Then

                    oSR.Reference2 = oDBs_Head.GetValue("DocNum", 0)
                    oSR.Comments = "Sub Contracting Cutting  stock is revaluated as per cost run report for the month of " + oDBs_Head.GetValue("U_Month", 0) + oDBs_Head.GetValue("U_Year", 0)
                    If oSR.Add <> 0 Then
                        _str_flag = "N"
                        'Dim Str As String = oCompany.GetLastErrorDescription
                        ' oApplication.MessageBox(oCompany.GetLastErrorDescription + "-" + "Cutting")
                        Throw New Exception(oCompany.GetLastErrorDescription + "-" + "Sub Contractor Cutting")
                    Else
                        Dim docEntry As String = ""
                        oCompany.GetNewObjectCode(docEntry)
                        oRS1.DoQuery("update [@GEN_SAM_REV_D0] set U_SCutRefNo='" & docEntry & "'  where DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_SubCut > 0 ")
                        _str_flag = "Y"
                        oApplication.SetStatusBarMessage("Sucessfully Posted-Sub Contractor Cutting", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    End If
                End If
            Else
                _str_flag = "Y"
            End If

            '******************Sub Contractor Cutting*********************

            '******************Sub Contractor Stiching*********************
            oSR = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
            _int_cnt = 0
            oRS.DoQuery("select U_DocDate,U_ItmCode ,U_SubSit ,U_STRCst  from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_SubSit >0 and isnull(U_SSitRefno,'')='' ")
            If oRS.RecordCount > 0 Then
                If oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT1" Then
                    _str_Account = "_SYS00000001447"
                    _str_whs_ = "JW%"
                ElseIf oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT2" Then
                    _str_Account = "_SYS00000001448"
                    _str_whs_ = "2JW%"
                Else
                    _str_Account = "_SYS00000001449"
                    _str_whs_ = "3JW%"
                End If
                ' Dim _dcl_ctrev As Decimal = Decimal.Round(oRS2.Fields.Item("AvgPrice").Value, 2)

                For j As Int16 = 0 To oRS.RecordCount - 1

                    oRS3.DoQuery("select OnHand,ItemCode,WhsCode from OITW where  OnHand>0 and ItemCode= '" & oRS.Fields.Item("U_ItmCode").Value.ToString.Trim() + "-2" & "' and  WhsCode like '" & _str_whs_ & "' ")
                    For I As Int16 = 0 To oRS3.RecordCount - 1
                        oRS2.DoQuery("select AvgPrice from OITW where ItemCode ='" & oRS.Fields.Item("U_ItmCode").Value.ToString.Trim() + "-2" & "' and WhsCode ='" & oRS3.Fields.Item("WhsCode").Value & "'")
                        If Decimal.Round(oRS2.Fields.Item("AvgPrice").Value, 1) <> Decimal.Round(oRS.Fields.Item("U_STRCst").Value, 1) Then
                            _str_flag = "N"
                            oSR_Lines = oSR.Lines
                            oSR_Lines.SetCurrentLine(_int_cnt)
                            oSR_Lines.ItemCode = oRS.Fields.Item("U_ItmCode").Value.ToString.Trim() + "-2"
                            oSR_Lines.WarehouseCode = oRS3.Fields.Item("WhsCode").Value
                            oSR_Lines.RevaluationIncrementAccount = _str_Account
                            oSR_Lines.RevaluationDecrementAccount = _str_Account
                            oSR_Lines.Price = oRS.Fields.Item("U_STRCst").Value
                            oSR_Lines.Add()
                            oSR.DocDate = oRS.Fields.Item("U_DocDate").Value
                            oSR.TaxDate = oRS.Fields.Item("U_DocDate").Value
                            _int_cnt = _int_cnt + 1
                        End If
                        oRS3.MoveNext()
                    Next
                    oRS.MoveNext()
                Next
                If _int_cnt <> 0 Then

                    oSR.Reference2 = oDBs_Head.GetValue("DocNum", 0)
                    oSR.Comments = "Sub Contracting Stiching  stock is revaluated as per cost run report for the month of " + oDBs_Head.GetValue("U_Month", 0) + oDBs_Head.GetValue("U_Year", 0)
                    If oSR.Add <> 0 Then
                        _str_flag = "N"
                        'Dim Str As String = oCompany.GetLastErrorDescription
                        ' oApplication.MessageBox(oCompany.GetLastErrorDescription + "-" + "Cutting")
                        Throw New Exception(oCompany.GetLastErrorDescription + "-" + "Sub Contracting Stiching ")
                    Else
                        Dim docEntry As String = ""
                        oCompany.GetNewObjectCode(docEntry)
                        oRS1.DoQuery("update [@GEN_SAM_REV_D0] set U_SSitRefno='" & docEntry & "'  where DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_SubSit > 0 ")
                        _str_flag = "Y"
                        oApplication.SetStatusBarMessage("Sucessfully Posted-Sub Contracting Stiching ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    End If
                End If
            Else
                _str_flag = "Y"
            End If
            '******************Sub Contractor Stiching*********************
            '******************Sub Contractor FI*********************
            oSR = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
            oRS.DoQuery("select U_DocDate,U_ItmCode ,U_SubFin ,U_FinRCst  from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_SubFin >0 and isnull(U_SFinRefno,'')='' ")
            If oRS.RecordCount > 0 Then
                If oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT1" Then
                    _str_Account = "_SYS00000001447"
                    _str_whs_ = "JW%"
                ElseIf oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT2" Then
                    _str_Account = "_SYS00000001448"
                    _str_whs_ = "2JW%"
                Else
                    _str_Account = "_SYS00000001449"
                    _str_whs_ = "3JW%"
                End If
                ' Dim _dcl_ctrev As Decimal = Decimal.Round(oRS2.Fields.Item("AvgPrice").Value, 2)
                _int_cnt = 0
                For j As Int16 = 0 To oRS.RecordCount - 1

                    oRS3.DoQuery("select OnHand,ItemCode,WhsCode from OITW where  OnHand>0 and ItemCode= '" & oRS.Fields.Item("U_ItmCode").Value.ToString.Trim() + "-3" & "' and  WhsCode  like '" & _str_whs_ & "' ")
                    For I As Int16 = 0 To oRS3.RecordCount - 1
                        oRS2.DoQuery("select AvgPrice from OITW where ItemCode ='" & oRS.Fields.Item("U_ItmCode").Value.ToString.Trim() + "-3" & "' and WhsCode ='" & oRS3.Fields.Item("WhsCode").Value & "'")
                        If Decimal.Round(oRS2.Fields.Item("AvgPrice").Value, 1) <> Decimal.Round(oRS.Fields.Item("U_FinRCst").Value, 1) Then
                            _str_flag = "N"
                            oSR_Lines = oSR.Lines
                            oSR_Lines.SetCurrentLine(_int_cnt)
                            oSR_Lines.ItemCode = oRS.Fields.Item("U_ItmCode").Value.ToString.Trim() + "-3"
                            oSR_Lines.WarehouseCode = oRS3.Fields.Item("WhsCode").Value
                            oSR_Lines.RevaluationIncrementAccount = _str_Account
                            oSR_Lines.RevaluationDecrementAccount = _str_Account
                            oSR_Lines.Price = oRS.Fields.Item("U_FinRCst").Value
                            oSR_Lines.Add()
                            oSR.DocDate = oRS.Fields.Item("U_DocDate").Value
                            oSR.TaxDate = oRS.Fields.Item("U_DocDate").Value
                            _int_cnt = _int_cnt + 1
                        End If
                        oRS3.MoveNext()
                    Next
                    oRS.MoveNext()
                Next
                If _int_cnt <> 0 Then

                    oSR.Reference2 = oDBs_Head.GetValue("DocNum", 0)
                    oSR.Comments = "Sub Contracting Finishing stock is revaluated as per cost run report for the month of " + oDBs_Head.GetValue("U_Month", 0) + oDBs_Head.GetValue("U_Year", 0)
                    If oSR.Add <> 0 Then
                        _str_flag = "N"
                        'Dim Str As String = oCompany.GetLastErrorDescription
                        ' oApplication.MessageBox(oCompany.GetLastErrorDescription + "-" + "Cutting")
                        Throw New Exception(oCompany.GetLastErrorDescription + "-" + "Sub Contracting Finishing Stock ")
                    Else
                        Dim docEntry As String = ""
                        oCompany.GetNewObjectCode(docEntry)
                        oRS1.DoQuery("update [@GEN_SAM_REV_D0] set U_SFinRefno='" & docEntry & "'  where DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_SubFin > 0 ")
                        _str_flag = "Y"
                        oApplication.SetStatusBarMessage("Sucessfully Posted-Sub Contracting Finishing stock ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    End If
                End If
            Else
                _str_flag = "Y"
            End If
            '******************Sub Contractor FI*********************
            '******************Sub Contractor FG*********************
            oSR = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
            oRS.DoQuery("select U_DocDate,U_ItmCode ,U_SubFG ,U_FGPQTY,U_FinRCst  from [@GEN_SAM_REV_D0] T0 inner join [@GEN_SAM_REV] T1 on T0.Docentry=T1.DocEntry  where T0.DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_SubFG >0 and isnull(U_SFGRefno,'')='' ")
            If oRS.RecordCount > 0 Then
                If oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT1" Then
                    _str_Account = "_SYS00000001447"
                    _str_whs_ = "JW%"
                ElseIf oDBs_Head.GetValue("U_Unit", 0).Trim() = "UNIT2" Then
                    _str_Account = "_SYS00000001448"
                    _str_whs_ = "2JW%"
                Else
                    _str_Account = "_SYS00000001449"
                    _str_whs_ = "3JW%"
                End If
                ' Dim _dcl_ctrev As Decimal = Decimal.Round(oRS2.Fields.Item("AvgPrice").Value, 2)
                _int_cnt = 0
                For j As Int16 = 0 To oRS.RecordCount - 1

                    oRS3.DoQuery("select OnHand,ItemCode,WhsCode from OITW where  OnHand>0 and ItemCode= '" & oRS.Fields.Item("U_ItmCode").Value.ToString.Trim() & "' and  WhsCode like  '" & _str_whs_ & "' ")
                    For I As Int16 = 0 To oRS3.RecordCount - 1
                        oRS2.DoQuery("select AvgPrice from OITW where ItemCode ='" & oRS3.Fields.Item("ItemCode").Value & "' and WhsCode ='" & oRS3.Fields.Item("WhsCode").Value & "'")
                        If Decimal.Round(oRS2.Fields.Item("AvgPrice").Value, 1) <> Decimal.Round(oRS.Fields.Item("U_FinRCst").Value, 1) Then
                            _str_flag = "N"
                            oSR_Lines = oSR.Lines
                            oSR_Lines.SetCurrentLine(_int_cnt)
                            oSR_Lines.ItemCode = oRS.Fields.Item("U_ItmCode").Value.ToString.Trim()
                            oSR_Lines.WarehouseCode = oRS3.Fields.Item("WhsCode").Value
                            oSR_Lines.RevaluationIncrementAccount = _str_Account
                            oSR_Lines.RevaluationDecrementAccount = _str_Account
                            oSR_Lines.Price = oRS.Fields.Item("U_FGPQTY").Value
                            oSR_Lines.Add()
                            oSR.DocDate = oRS.Fields.Item("U_DocDate").Value
                            oSR.TaxDate = oRS.Fields.Item("U_DocDate").Value
                            _int_cnt = _int_cnt + 1
                        End If
                        oRS3.MoveNext()
                    Next
                    oRS.MoveNext()
                Next
                If _int_cnt <> 0 Then

                    oSR.Reference2 = oDBs_Head.GetValue("DocNum", 0)
                    oSR.Comments = "Sub Contracting FG is revaluated as per cost run report for the month of " + oDBs_Head.GetValue("U_Month", 0) + oDBs_Head.GetValue("U_Year", 0)
                    If oSR.Add <> 0 Then
                        _str_flag = "N"
                        'Dim Str As String = oCompany.GetLastErrorDescription
                        ' oApplication.MessageBox(oCompany.GetLastErrorDescription + "-" + "Cutting")
                        Throw New Exception(oCompany.GetLastErrorDescription + "-" + "Sub Contracting FG ")
                    Else
                        Dim docEntry As String = ""
                        oCompany.GetNewObjectCode(docEntry)
                        oRS1.DoQuery("update [@GEN_SAM_REV_D0] set U_SFGRefno='" & docEntry & "'  where DocEntry='" & oDBs_Head.GetValue("DocEntry", 0) & "' and U_SubFG > 0 ")
                        _str_flag = "Y"
                        oApplication.SetStatusBarMessage("Sucessfully Posted-Sub Contracting FG ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    End If
                End If
            Else
                _str_flag = "Y"
            End If
            '******************Sub Contractor FG********************




            'oRS.DoQuery("select ")
            '  oRS.DoQuery("select T0.docentry ,T0.U_Stat from [@GEN_SAM_REV_D0]  inner join  [@GEN_SAM_REV] T0 on T0.docentry=[@GEN_SAM_REV_D0].Docentry where  (isnull(U_CTRefNo,'')<>'' and isnull(U_STRefNo,'')<>'' and isnull(U_WFIRefNo,'')<>''and isnull(U_WSTRefNo,'')<>''and isnull(U_FGRefNo ,'') <>'' )  and T0.DocEntry ='" & oDBs_Head.GetValue("DocEntry", 0) & "'")


            If _str_flag = "Y" Then
                oRS.DoQuery(" update [@GEN_SAM_REV]  set U_Stat='Generated' where DocEntry ='" & oDBs_Head.GetValue("DocEntry", 0) & "' ")
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
            End If
            oApplication.SetStatusBarMessage("Process Over", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
