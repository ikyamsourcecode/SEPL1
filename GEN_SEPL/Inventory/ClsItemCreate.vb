Public Class ClsItemCreate

#Region "        Declaration        "
    Dim oUtilities As New ClsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
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
    Dim UpdMode As Boolean = False
    Dim DocStatus As String
    Dim ItmTypFlg As Boolean = False
    Dim SubTypeFlg As Boolean = False
    Dim StyleFlg As Boolean = False
    Dim ColorFlg As Boolean = False
    Dim QualityFlg As Boolean = False
    Dim SizeFlg As Boolean = False
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            oUtilities.SAPXML("GEN_ITEM_CREATE.xml")
            objForm = oApplication.Forms.GetForm("GEN_ITEM_CREATE", oApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_ITEM_CREATE")
            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try
            objMatrix = objForm.Items.Item("mtx").Specific
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_ITEM_CREATE")

            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Sub FilterItemType(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("ITMTYPE")
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Code From [@GEN_ITM_TYPE] Where u_type = '" + Trim(objForm.Items.Item("itmmst").Specific.value) + "'")
            If oRecordSet.RecordCount = 0 Then
                ItmTypFlg = True
                Exit Sub
            End If
            ItmTypFlg = False
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            For IntICount As Integer = 0 To oRecordSet.RecordCount - 1
                If IntICount = (oRecordSet.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "Code"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("Code").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "Code"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("Code").Value
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                oRecordSet.MoveNext()
            Next
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterSubType(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("SUBTYPE")
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Code From [@GEN_SUB_TYPE] Where u_type = '" + Trim(objForm.Items.Item("itmmst").Specific.value) + "'")
            If oRecordSet.RecordCount = 0 Then
                SubTypeFlg = True
                Exit Sub
            End If
            SubTypeFlg = False
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            For IntICount As Integer = 0 To oRecordSet.RecordCount - 1
                If IntICount = (oRecordSet.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "Code"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("Code").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "Code"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("Code").Value
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                oRecordSet.MoveNext()
            Next
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterStyle(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("STLCODE")
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Code From [@GEN_STYLE_CODE] Where u_type = '" + Trim(objForm.Items.Item("itmmst").Specific.value) + "'")
            If oRecordSet.RecordCount = 0 Then
                StyleFlg = True
                Exit Sub
            End If
            StyleFlg = False
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            For IntICount As Integer = 0 To oRecordSet.RecordCount - 1
                If IntICount = (oRecordSet.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "Code"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("Code").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "Code"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("Code").Value
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                oRecordSet.MoveNext()
            Next
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterColor(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("COLOR")
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Code From [@GEN_COLOR_CODE] Where u_type = '" + Trim(objForm.Items.Item("itmmst").Specific.value) + "'")
            If oRecordSet.RecordCount = 0 Then
                ColorFlg = True
                Exit Sub
            End If
            ColorFlg = False
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            For IntICount As Integer = 0 To oRecordSet.RecordCount - 1
                If IntICount = (oRecordSet.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "Code"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("Code").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "Code"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("Code").Value
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                oRecordSet.MoveNext()
            Next
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterQuality(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("QLTY")
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Code From [@GEN_QLTY_CODE] Where u_type = '" + Trim(objForm.Items.Item("itmmst").Specific.value) + "'")
            If oRecordSet.RecordCount = 0 Then
                QualityFlg = True
                Exit Sub
            End If
            QualityFlg = False
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            For IntICount As Integer = 0 To oRecordSet.RecordCount - 1
                If IntICount = (oRecordSet.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "Code"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("Code").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "Code"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("Code").Value
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                oRecordSet.MoveNext()
            Next
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterSize(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("SIZE")
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Name From [@GEN_SIZE_CODE] Where u_type = '" + Trim(objForm.Items.Item("itmmst").Specific.value) + "'")
            If oRecordSet.RecordCount = 0 Then
                SizeFlg = True
                Exit Sub
            End If
            SizeFlg = False
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            For IntICount As Integer = 0 To oRecordSet.RecordCount - 1
                If IntICount = (oRecordSet.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "Name"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("Name").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "Name"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("Name").Value
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                oRecordSet.MoveNext()
            Next
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = oApplication.Forms.Item(pVal.FormUID)
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "copyto" Then
                        If pVal.BeforeAction = True Then
                            If Trim(objForm.Items.Item("itemcode").Specific.value) = "" Then
                                oApplication.StatusBar.SetText("Please generate Item Code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        Else
                            oApplication.ActivateMenuItem("3073")
                            Dim ItemForm As SAPbouiCOM.Form
                            ItemForm = oApplication.Forms.ActiveForm
                            ItemForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                            ItemForm.Items.Item("5").Specific.value = objForm.Items.Item("itemcode").Specific.value
                            ItemForm.Items.Item("color").Specific.value = objForm.Items.Item("color").Specific.value
                            ItemForm.Items.Item("colornm").Specific.value = objForm.Items.Item("fld6").Specific.value
                            ItemForm.Items.Item("cust").Specific.value = objForm.Items.Item("custcode").Specific.value
                            ItemForm.Items.Item("custnm").Specific.value = objForm.Items.Item("fld4").Specific.value
                            ItemForm.Items.Item("type").Specific.value = objForm.Items.Item("itmtype").Specific.value
                            ItemForm.Items.Item("typenm").Specific.value = objForm.Items.Item("fld2").Specific.value
                            ItemForm.Items.Item("size").Specific.value = objForm.Items.Item("size").Specific.value
                            ItemForm.Items.Item("sizenm").Specific.value = objForm.Items.Item("fld8").Specific.value
                            ItemForm.Items.Item("subtype").Specific.value = objForm.Items.Item("subtype").Specific.value
                            ItemForm.Items.Item("subtpnm").Specific.value = objForm.Items.Item("fld3").Specific.value
                            ItemForm.Items.Item("style").Specific.value = objForm.Items.Item("style").Specific.value
                            ItemForm.Items.Item("stylenm").Specific.value = objForm.Items.Item("fld5").Specific.value
                            ItemForm.Items.Item("qlty").Specific.value = objForm.Items.Item("quality").Specific.value
                            ItemForm.Items.Item("qltynm").Specific.value = objForm.Items.Item("fld7").Specific.value
                        End If
                    End If
                    If pVal.ItemUID = "btngen" Then
                        If pVal.BeforeAction = True Then
                            Dim Table, Field As String
                            Dim Flag As Boolean = False
                            Dim Length As Integer
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select u_field,u_length From [@GEN_PARAM_MST_D0] Where Code = '" + Trim(objForm.Items.Item("itmmst").Specific.value) + "' And IsNull(u_field,'') <> ''")
                            For i As Integer = 1 To oRSet.RecordCount
                                Field = oRSet.Fields.Item("u_field").Value
                                Length = oRSet.Fields.Item("u_length").Value
                                If Trim(objForm.Items.Item(Field).Specific.value) = "" Then
                                    oApplication.StatusBar.SetText("Please enter value in '" + Field + "'", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                If CStr(Trim(objForm.Items.Item(Field).Specific.value.ToString)).Length > Length And CStr(Trim(objForm.Items.Item(Field).Specific.value.ToString)).Length < 1 Then
                                    oApplication.StatusBar.SetText("Length of value in '" + Field + "' is invalid", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                oRSet.MoveNext()
                            Next
                        End If
                        If pVal.BeforeAction = False Then
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_ITEM_CREATE")
                            Dim Field As String
                            Dim ItemCode As String = ""
                            Dim Length As Integer
                            Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSet.DoQuery("Select u_field,u_length From [@GEN_PARAM_MST_D0] Where Code = '" + Trim(objForm.Items.Item("itmmst").Specific.value) + "' And IsNull(u_field,'') <> ''")
                            For i As Integer = 1 To oRSet.RecordCount
                                Field = oRSet.Fields.Item("u_field").Value
                                ItemCode = ItemCode + Trim(objForm.Items.Item(Field).Specific.Value)
                                oRSet.MoveNext()
                            Next
                            oDBs_Head.SetValue("u_itemcode", 0, ItemCode)
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If (pVal.ItemUID = "count" Or pVal.ItemUID = "width") And pVal.BeforeAction = True Then
                        If Trim(objForm.Items.Item("itmmst").Specific.value) <> "FABRIC" Then
                            BubbleEvent = False
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    If (pVal.ItemUID = "count" Or pVal.ItemUID = "width") And pVal.BeforeAction = True And (pVal.CharPressed <> 9 And pVal.CharPressed <> 13) Then
                        If Trim(objForm.Items.Item("itmmst").Specific.value) <> "FABRIC" Then
                            BubbleEvent = False
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim oRSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    If pVal.BeforeAction = True Then
                        If oCFL.UniqueID = "ITMTYPE" Or oCFL.UniqueID = "CUSTCODE" Or oCFL.UniqueID = "STLCODE" Or oCFL.UniqueID = "COLOR" Or oCFL.UniqueID = "QLTY" Or oCFL.UniqueID = "SIZE" Then
                            If Trim(oDBs_Head.GetValue("u_itmmst", 0)) = "" Then
                                oApplication.StatusBar.SetText("Please select Item Type Master", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                        If oCFL.UniqueID = "ITMTYPE" Then
                            Me.FilterItemType(FormUID)
                            If ItmTypFlg = True Then
                                oApplication.StatusBar.SetText("No item types for this kind of item", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                        If oCFL.UniqueID = "SUBTYPE" Then
                            Me.FilterSubType(FormUID)
                            If SubTypeFlg = True Then
                                oApplication.StatusBar.SetText("No sub item types for this kind of item", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                        If oCFL.UniqueID = "STLCODE" Then
                            Me.FilterStyle(FormUID)
                            If StyleFlg = True Then
                                oApplication.StatusBar.SetText("No styles for this kind of item", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                        If oCFL.UniqueID = "COLOR" Then
                            Me.FilterColor(FormUID)
                            If ColorFlg = True Then
                                oApplication.StatusBar.SetText("No colours for this kind of item", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                        If oCFL.UniqueID = "QLTY" Then
                            Me.FilterQuality(FormUID)
                            If QualityFlg = True Then
                                oApplication.StatusBar.SetText("No quality codes for this kind of item", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                        If oCFL.UniqueID = "SIZE" Then
                            Me.FilterSize(FormUID)
                            If SizeFlg = True Then
                                oApplication.StatusBar.SetText("No sizes for this kind of item", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If
                    If pVal.BeforeAction = False Then
                        If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@GEN_ITEM_CREATE")
                            If oCFL.UniqueID = "ITMMST" Then
                                oDBs_Head.SetValue("u_itmmst", 0, oDT.GetValue("Code", 0))
                                oDBs_Head.SetValue("u_fld1", 0, oDT.GetValue("Name", 0))
                            End If
                            If oCFL.UniqueID = "ITMTYPE" Then
                                oDBs_Head.SetValue("u_itmtype", 0, oDT.GetValue("Code", 0))
                                oDBs_Head.SetValue("u_fld2", 0, oDT.GetValue("Name", 0))
                            End If
                            If oCFL.UniqueID = "SUBTYPE" Then
                                oDBs_Head.SetValue("u_subtype", 0, oDT.GetValue("Name", 0))
                                oDBs_Head.SetValue("u_fld3", 0, oDT.GetValue("U_desc", 0))
                            End If
                            If oCFL.UniqueID = "CUSTCODE" Then
                                oDBs_Head.SetValue("u_custcode", 0, oDT.GetValue("Code", 0))
                                oDBs_Head.SetValue("u_fld4", 0, oDT.GetValue("Name", 0))
                            End If
                            If oCFL.UniqueID = "STLCODE" Then
                                oDBs_Head.SetValue("u_style", 0, oDT.GetValue("Name", 0))
                                oDBs_Head.SetValue("u_fld5", 0, oDT.GetValue("U_desc", 0))
                            End If
                            If oCFL.UniqueID = "COLOR" Then
                                oDBs_Head.SetValue("u_color", 0, oDT.GetValue("Name", 0))
                                oDBs_Head.SetValue("u_fld6", 0, oDT.GetValue("U_desc", 0))
                            End If
                            If oCFL.UniqueID = "QLTY" Then
                                oDBs_Head.SetValue("u_quality", 0, oDT.GetValue("Name", 0))
                                oDBs_Head.SetValue("u_fld7", 0, oDT.GetValue("U_desc", 0))
                            End If
                            If oCFL.UniqueID = "SIZE" Then
                                oDBs_Head.SetValue("u_size", 0, oDT.GetValue("Name", 0))
                                oDBs_Head.SetValue("u_fld8", 0, oDT.GetValue("U_size", 0))
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
                Dim objForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "GEN_ITEM_CREATE"
                        Me.CreateForm(objForm.UniqueID)
                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
