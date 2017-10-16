Public Class ClsWorOrdrGenWiz
    Dim objForm As SAPbouiCOM.Form
    Dim objUtilities As New ClsUtilities
    Dim oDBs_Detail As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim oMatrix As SAPbouiCOM.Matrix
    Dim oGrid As SAPbouiCOM.Grid
    Dim oGrid1 As SAPbouiCOM.Grid
    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = "GEN_WGW" Then
                If pVal.BeforeAction = True Then
                    Select Case pVal.EventType

                    End Select
                ElseIf pVal.BeforeAction = False Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("Delete From DI_API_ERROR_PP_ORDR")
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            objForm.Items.Item("S1G1").Enabled = True
                            objForm = oApplication.Forms.Item(FormUID)
                            If pVal.ItemUID = "S1B1" Then
                                oGrid = objForm.Items.Item("S2G2").Specific
                                If oGrid.Rows.Count = 0 Then
                                    Me.MoveToSecondScreen(objForm.UniqueID)
                                End If
                            End If
                            If pVal.ItemUID = "btnFind" Then
                                If pVal.BeforeAction = False Then
                                    If Trim(objForm.Items.Item("sorefno").Specific.value) <> "" Then
                                        oGrid = objForm.Items.Item("S1G1").Specific
                                        If objForm.Items.Item("S1G1").Visible = True Then
                                            If oGrid.Rows.Count > 0 Then
                                                For I As Integer = 0 To oGrid.Rows.Count - 1
                                                    If oGrid.DataTable.GetValue(2, I) = Trim(objForm.Items.Item("sorefno").Specific.value) Then
                                                        oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto
                                                        oGrid.Rows.SelectedRows.Add(I)
                                                        oGrid.DataTable.SetValue(0, I, "Y")
                                                    End If
                                                Next
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            If pVal.ItemUID = "S1B1" Then
                                oGrid = objForm.Items.Item("S2G2").Specific
                                Dim Flag As Int16 = 0
                                For I As Int16 = 0 To oGrid.Rows.Count - 1
                                    If I <= oGrid.Rows.Count - 1 Then
                                        If oGrid.DataTable.GetValue(0, I) <> "Y" Then
                                            oGrid.DataTable.Rows.Remove(I)
                                            I = I - 1
                                        End If
                                    End If
                                Next
                                For I As Int16 = 0 To oGrid.Rows.Count - 1
                                    If oGrid.DataTable.GetValue(6, I) > 0 And objForm.Items.Item("S1B1").Specific.Caption <> "Finish." Then
                                        objForm.Title = "Work Order Generation Wizard - Step 3 of 3"
                                        objForm.Items.Item("S1B1").Specific.Caption = "Finish"
                                    Else
                                        oApplication.StatusBar.SetText("Please enter quantity at Row : " & I + 1 & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        Exit For
                                    End If
                                Next
                                If objForm.Items.Item("S1B1").Specific.Caption = "Finish" Then
                                    For I As Int16 = 0 To oGrid.DataTable.Rows.Count - 1
                                        If oGrid.DataTable.GetValue(6, I) = 0 Then
                                            oApplication.StatusBar.SetText("Quantity cannot be zero for Item : " & oGrid.DataTable.GetValue(2, I) + "-" + oGrid.DataTable.GetValue(3, I) & " And Order No : " & oGrid.DataTable.GetValue(1, I) & "", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            Exit Sub
                                        End If
                                        If CStr(oGrid.DataTable.GetValue(10, I)) = String.Empty Then
                                            oApplication.StatusBar.SetText("PO-Date cannot be null for Item : " & oGrid.DataTable.GetValue(2, I) + "-" + oGrid.DataTable.GetValue(3, I) & " And Order No : " & oGrid.DataTable.GetValue(1, I) & "", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            Exit Sub
                                        End If
                                        'Vijeesh
                                        If CStr(oGrid.DataTable.GetValue(11, I)) = String.Empty Then
                                            oApplication.StatusBar.SetText("Unit cannot be null for Item : " & oGrid.DataTable.GetValue(2, I) + "-" + oGrid.DataTable.GetValue(3, I) & " And Order No : " & oGrid.DataTable.GetValue(1, I) & "", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            Exit Sub
                                        End If
                                        'Vijeesh
                                    Next
                                    For I As Int16 = 0 To oGrid.DataTable.Rows.Count - 1
                                        Me.CreateProductionOrder(objForm.UniqueID, oGrid.DataTable.GetValue(2, I), oGrid.DataTable.GetValue(6, I), oGrid.DataTable.GetValue(1, I), oGrid.DataTable.GetValue(10, I), oGrid.DataTable.GetValue(4, I), oGrid.DataTable.GetValue(11, I))
                                    Next
                                    objForm.Items.Item("S1B1").Specific.Caption = "Finish."
                                End If
                                If objForm.Items.Item("S1B1").Specific.Caption = "Finish." Then
                                    objForm.Items.Item("S3G3").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                    objForm.Items.Item("S2G2").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                    objForm.Items.Item("S1G1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                    Me.GetErrors(objForm.UniqueID)
                                End If
                            End If
                            If pVal.ItemUID = "S2B2" Then
                                objForm.Items.Item("S1B1").Specific.Caption = "Next"
                                Me.MoveBackToFirstScreen(objForm.UniqueID)
                            End If
                    End Select
                End If
            End If
        Catch ex As Exception
            oApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case "GEN_WGW"
                        objUtilities.SAPXML("WorkOrderWiz.xml")
                        objForm = oApplication.Forms.GetForm("GEN_WGW", 0)
                        Me.CreateFormItems(objForm.UniqueID)
                End Select
            End If
        Catch ex As Exception
            oApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.ActionSuccess = True Then
                        objForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    End If
            End Select
        Catch ex As Exception
            oApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    Sub GetErrors(ByVal FormUID As String)
        objForm = oApplication.Forms.Item(FormUID)
        Dim oGrid As SAPbouiCOM.Grid
        oGrid = objForm.Items.Item("S3G3").Specific
        objForm.DataSources.DataTables.Item(2).ExecuteQuery("Select A.OrderNo,A.PrdNo [Production Order],A.ItemCode,A.ErrCode [Result],A.Des [Result Description] From DI_API_ERROR_PP_ORDR A")
        oGrid.DataTable = objForm.DataSources.DataTables.Item("S1DT3")
        objForm.Items.Item("S1B1").Visible = False
        objForm.Items.Item("S2B2").Visible = False
    End Sub
    Sub CreateProductionOrder(ByVal FormUID As String, ByVal ItemCode As String, ByVal Qty As Double, ByVal OriginNum As Int64, ByVal PODate As Date, ByVal Process As String, ByVal Unit As String)
        objForm = oApplication.Forms.Item(FormUID)
        Dim oPO As SAPbobsCOM.ProductionOrders = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
        oPO.ProductionOrderOrigin = SAPbobsCOM.BoProductionOrderOriginEnum.bopooSalesOrder
        oPO.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard
        'oPO.PostingDate = DateTime.ParseExact(Trim(oApplication.Company.ServerDate), "yyyyMMdd", Nothing)
        'oPO.DueDate = DateTime.ParseExact(Trim(oApplication.Company.ServerDate), "yyyyMMdd", Nothing)
        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim RSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'oRecordSet.DoQuery("select GetDate()")
        Dim Date1 As Date = PODate
        'Dim Date1 As Date = "22/02/2011"
        oPO.PostingDate = Date1
        oPO.DueDate = Date1
        oPO.ItemNo = ItemCode
        oPO.PlannedQuantity = Qty
        oPO.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned
        'Vijeesh
        RSet.DoQuery("Select B.u_outwhs From [@GEN_UNIT_MST] A Inner Join [@GEN_UNIT_MST_D0] B On A.Code = B.Code Where A.Name = '" + Unit + "' And B.u_process = '" + Process + "'")
        If RSet.RecordCount = 0 Then
            RSet.DoQuery("Select u_outwhs From [@GEN_PROD_PRCS] Where u_itemCode = '" + ItemCode + "'")
        End If
        oPO.Warehouse = RSet.Fields.Item("u_outwhs").Value.ToString()
        'Vijeesh

        oRecordSet.DoQuery("Select DocEntry From ORDR Where DocNum = '" & OriginNum & "'")
        oPO.ProductionOrderOriginEntry = oRecordSet.Fields.Item(0).Value
        oPO.UserFields.Fields.Item("U_Created").Value = "Yes"
        oPO.UserFields.Fields.Item("U_unit").Value = Unit
        oPO.UserFields.Fields.Item("U_process").Value = Process
        ' oPO.UserFields.Fields.Item("U_type").Value = "Regular"
        Dim NewPoNo As Integer
        If oPO.Add() = 0 Then
            oRecordSet.DoQuery("Select DocNum From OWOR where DocEntry = '" & oCompany.GetNewObjectKey & "' ")
            NewPoNo = oRecordSet.Fields.Item(0).Value
            oRecordSet.DoQuery("Insert into DI_API_ERROR_PP_ORDR values ('" & OriginNum & "','" & NewPoNo & "','" & ItemCode & "','Success','Success')")
            oApplication.StatusBar.SetText("Production order created successfully (S.O:" & OriginNum & " Item : " & ItemCode & ")", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Else
            oRecordSet.DoQuery("Select DocNum From OWOR where DocEntry = '" & oCompany.GetNewObjectKey & "' ")
            NewPoNo = oRecordSet.Fields.Item(0).Value
            oRecordSet.DoQuery("Insert into DI_API_ERROR_PP_ORDR values ('" & OriginNum & "','" & NewPoNo & "','" & ItemCode & "','" & oCompany.GetLastErrorCode & "','Error')")
            oApplication.StatusBar.SetText("Error (" & oCompany.GetLastErrorDescription & ") (S.O:" & OriginNum & " Item : " & ItemCode & ")", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If
    End Sub
    Sub MoveToSecondScreen(ByVal FormUID As String)
        objForm = oApplication.Forms.Item(FormUID)
        objForm.Title = "Work Order Generation Wizard - Step 2 of 3"
        'objForm.Items.Item("S1S1").Visible = False
        objForm.Items.Item("S1G1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        'objForm.Items.Item("S1G1").Visible = False
        objForm.Items.Item("S2B2").Visible = True
        objForm.Items.Item("S2G2").Visible = True
        CreateWorkOrders(objForm.UniqueID)
    End Sub
    Sub CreateWorkOrders(ByVal FormUID As String)
        objForm = oApplication.Forms.Item(FormUID)
        objForm.Freeze(True)
        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oGrid = objForm.Items.Item("S1G1").Specific
        oGrid1 = objForm.Items.Item("S2G2").Specific
        For I As Int16 = 0 To oGrid.Rows.Count - 1
            If oGrid.DataTable.GetValue(0, I) = "Y" Then
                oRecordSet.DoQuery("Exec Recursive_BOM '" & oGrid.DataTable.GetValue("Item Code", I) & "','" & oGrid.DataTable.GetValue(1, I) & "','" & oGrid.DataTable.GetValue(5, I) & "'")
            End If
        Next
        objForm.DataSources.DataTables.Item(1).ExecuteQuery("select *,[Sal. Ordr. Qty] - isnull([Created Qty],0)  [Rem. Qty],GetDate() [PO Date],Convert(VarChar(30),'') [Unit] From WORK_ORDR_RESULT order by Cast([Doc Num] As Numeric(15)) ")
        oGrid1.DataTable = objForm.DataSources.DataTables.Item("S1DT2")
        oGrid1.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oGrid1.Columns.Item(1).Editable = False
        oGrid1.Columns.Item(2).Editable = False
        oGrid1.Columns.Item(3).Editable = False
        oGrid1.Columns.Item(4).Editable = False
        oGrid1.Columns.Item(5).Editable = False
        oGrid1.Columns.Item(7).Visible = False
        oGrid1.Columns.Item(6).Editable = True
        oGrid1.Columns.Item(7).Editable = False
        oGrid1.Columns.Item(8).Editable = False
        oGrid1.Columns.Item(9).Editable = False
        oGrid1.Columns.Item(2).Width = 200
        oGrid1.Columns.Item(3).Width = 100
        oGrid1.Columns.Item(1).Width = 100
        oGrid1.Columns.Item(0).Width = 35

        oRecordSet.DoQuery("Delete from WORK_ORDR_RESULT")
        objForm.Freeze(False)
    End Sub
    Sub MoveBackToFirstScreen(ByVal FormUID As String)
        objForm = oApplication.Forms.Item(FormUID)
        objForm.Title = "Work Order Generation Wizard - Step 1 of 3"
        'objForm.Items.Item("S1S1").Visible = True
        objForm.Items.Item("S1G1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        'objForm.Items.Item("S1G1").Visible = True
        objForm.Items.Item("S1B1").Visible = True
        objForm.Items.Item("S2B2").Visible = False
        objForm.Items.Item("S2G2").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        oGrid = objForm.Items.Item("S2G2").Specific
        oGrid.DataTable.Clear()
        oGrid = objForm.Items.Item("S1G1").Specific
        objForm.Items.Item("S1G1").Enabled = True
        oGrid.Columns.Item(0).Editable = True
        'objForm.Items.Item("S2G2").Visible = False
    End Sub
    Sub CreateFormItems(ByVal FormUID As String)
        Try
            objForm = oApplication.Forms.Item(FormUID)
            'objForm.Items.Item("S1S1").Visible = False
            objForm.DataSources.DataTables.Add("S1DT1")
            objForm.DataSources.DataTables.Add("S1DT2")
            objForm.DataSources.DataTables.Add("S1DT3")
            AddValuesToGrid(objForm.UniqueID)
            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try   
    End Sub
    Sub AddValuesToGrid(ByVal FormUID As String)
        objForm = oApplication.Forms.Item(FormUID)
        Dim oGrid As SAPbouiCOM.Grid
        oGrid = objForm.Items.Item("S1G1").Specific
        'oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objForm.DataSources.DataTables.Item(0).ExecuteQuery("Select 'N' [Select],T0.DocNum [Doc Number],T0.NumAtCard [Customer Ref No],T1.ItemCode As [Item Code],T1.Dscription As [Description],T1.OpenQty AS 'Quantity' From ORDR T0 Inner join RDR1 T1 ON T0.DocEntry = T1.DocEntry Where T0.DocStatus = 'O' And T1.ItemCode in (Select Code From OITT) Order by T0.DocNum")
        oGrid.DataTable = objForm.DataSources.DataTables.Item("S1DT1")
        oGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        objForm.Items.Item("S1G1").Enabled = True
        oGrid.Columns.Item(1).Editable = False
        oGrid.Columns.Item(2).Editable = False
        oGrid.Columns.Item(3).Editable = False
        oGrid.Columns.Item(4).Editable = False
    End Sub
   
End Class

