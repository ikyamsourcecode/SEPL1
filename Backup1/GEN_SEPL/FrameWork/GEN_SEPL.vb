Imports System.IO

Module GEN_SEPL
    Public WithEvents oApplication As SAPbouiCOM.Application
    Public oCompany As SAPbobsCOM.Company
    Dim objUtilities As New ClsUtilities
    Public SERVER As String
    Public DB_USERNAME As String
    Public DB_PASSWORD As String
    Public PTNFlag As Boolean = False
    Public sRptName As String ''''
    Public sDocNum As String ''''

    Structure DB_Login
        Public SERVER As String
        Public DB_USERNAME As String
        Public DB_PASSWORD As String
    End Structure

    Sub Main()
        objUtilities.StartUp()
        Dim oLogin As New CustomLogin.GetLoginInfo
        Dim Flag As Boolean = True
        oLogin.VerifyConnectionInfo(oCompany.Server, oCompany.CompanyDB, Flag)
        If Flag = True Then
            Dim DB_Property As DB_Login
            Dim LoginInfo As String() = oLogin.GetInfo()
            DB_Property.SERVER = LoginInfo(0)
            DB_Property.DB_USERNAME = LoginInfo(1)
            DB_Property.DB_PASSWORD = LoginInfo(2)
        End If
        Application.Run()
    End Sub

#Region "Class Declaration"
    Dim objGEN_CST_PRCS As New ClsGEN_CST_PRCS
    Dim objProductionProcess As New ClsProductionProcess
    Dim objCOA As New ClsCOA
    Dim objJE As New ClsJE
    Dim objsam As New Clssam
    Dim objIncomingPayments As New ClsIncomingPayments
    Dim objOutgoingPayments As New ClsOutgoingPayments
    Dim objCustomBOM As New ClsCustomBOM
    Dim objMaterialRequisition As New ClsMaterialRequisition
    Dim objPTN As New ClsPTN
    Dim objUnitMaster As New ClsUnitMaster

    Dim objGRPOLC As New ClsGRPOLC
    Dim objInventoryTransfer As New ClsInventoryTransfer
    Dim objSubContract As New ClsSubContract
    Dim objSubContract_GRPO As New ClsSubContract_GRPO
    Dim objSubContract_DC As New ClsSubContract_DC
    Dim objSubContract_Return As New ClsSubContract_Return
    Dim objWorOrdrGenWiz As New ClsWorOrdrGenWiz
    Dim objIssueForProduction As New ClsIssueForProduction
    Dim objReceiptFromProduction As New ClsReceiptFromProduction
    Dim objFinishSetup As New ClsFinishSetup
    Dim objFinishScreen As New ClsFinishScreen
    Dim objGRPO As New ClsGRPO
    Dim objAPInvoice As New ClsAPInvoice
    Dim objMachinePool As New ClsMachinePool
    Dim objMachineAllocation As New ClsMachineAllocation
    Dim objParamMst As New ClsParamMst
    Dim objItemCreate As New ClsItemCreate
    Dim objPurchaseOrder As New ClsPurchaseOrder
    Dim objStitchingScreen As New ClsStitchScreen
    Dim objAPCreditMemo As New ClsAPCreditMemo
    Dim objItemType As New ClsItemType
    Dim objItemMst As New ClsItemMst
    Dim objGEN_CUST_CODE As New ClsGEN_CUST_CODE
    Dim objGEN_SUB_TYPE As New ClsGEN_SUB_TYPE
    Dim objGEN_STYLE_CODE As New ClsGEN_STYLE_CODE
    Dim objGEN_COLOR_CODE As New ClsGEN_COLOR_CODE
    Dim objGEN_QLTY_CODE As New ClsGEN_QLTY_CODE
    Dim objGEN_SIZE_CODE As New ClsGEN_SIZE_CODE
    Dim objSalesReturn As New ClsSalesReturn
    Dim objSalesQuotation As New ClsSalesQuotation
    Dim objSalesOrder As New ClsSalesOrder
    Dim objDelivery As New ClsDelivery
    Dim objARInvoice As New ClsARInvoice
    Dim objARCreditMemo As New ClsARCreditMemo
    Dim objPurchaseReturn As New ClsPurchaseReturn
    Dim objUser As New ClsUser

    Dim objSAMRevaluation As New ClsSAMRevaluation
    Dim objGEN_COST_SHEET As New ClsGEN_COST_SHEET
    Dim objItemMasterData As New ClsItemMasterData
    Dim objAssortedMaster As New ClsAssortedMaster
    Dim objSizeMaster As New ClsSizeMaster
    'Vijeesh
    Dim objGoodsIssue As New ClsGoodsIssue
    Dim objGoodsReceipt As New ClsGoodsReceipt
    'Vijeesh
    Dim objClsUploadBOM As New ClsUploadBOM
    'Vijeesh

    'Rajkumar
    Dim objPreShipment As New ClsPreShipment
    Dim objSuppPrice As New ClsSuppPrice
    Dim objApportionAccural As New ClsApportionAccural

    'Rajkumar 18.08.14
    Dim objAppDbk As New ClsAppDbk
    Dim objForwardCover As New ClsForwardCover
#End Region

    Private Sub oApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles oApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                'objUtilities.SAPXML("RemoveMenu.xml")
                oCompany.Disconnect()
                End
        End Select
    End Sub

    Private Sub oApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles oApplication.FormDataEvent
        Try
            Select Case BusinessObjectInfo.FormTypeEx
                Case "GEN_ASSORTMENT"
                    objAssortedMaster.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "181"
                    objAPCreditMemo.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "170"
                    objIncomingPayments.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "SOORD"
                    objSalesOrder.FormDataEvent_SalesOrder_Allocation(BusinessObjectInfo, BubbleEvent)
                Case "GEN_SZ_ORDR"
                    objSalesOrder.FormDataEvent_SalesOrder_SizeMatrix(BusinessObjectInfo, BubbleEvent)
                Case "182"
                    objPurchaseReturn.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "143"
                    objGRPO.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                Case "142"
                    objPurchaseOrder.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "179"
                    objARCreditMemo.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "133"
                    objARInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "140"
                    objDelivery.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "139"
                    objSalesOrder.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "149"
                    objSalesQuotation.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "180"
                    objSalesReturn.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                    'Case "806"
                    '    objCOA.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_PROD_PRCS"
                    objProductionProcess.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_CUST_BOM"
                    objCustomBOM.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_MREQ"
                    objMaterialRequisition.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_PTN"
                    objPTN.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_UNIT_MST"
                    objUnitMaster.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                
                Case "GEN_GRPO_LCOSTS"
                    objGRPOLC.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "940"
                    objInventoryTransfer.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_SCForm"
                    objSubContract.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_SCGRPO"
                    objSubContract_GRPO.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_SCDC"
                    objSubContract_DC.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_SCRET"
                    objSubContract_Return.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_WGW"
                    objWorOrdrGenWiz.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "65213"
                    objIssueForProduction.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "65214"
                    objReceiptFromProduction.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_FIN_SETUP"
                    objFinishSetup.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_FIN_DESCR"
                    objFinishScreen.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_MACH_POOL"
                    objMachinePool.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_MACH_ALLOC"
                    objMachineAllocation.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_PARAM_MST"
                    objParamMst.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_CAP_PLAN"
                    objMachineAllocation.Child_FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_STH_DESCR"
                    objStitchingScreen.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "141"
                    objAPInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                    'Vijeesh
                Case "720"
                    objGoodsIssue.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "721"
                    objGoodsReceipt.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "150"
                    objItemMasterData.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "SAM"
                    objSAMRevaluation.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_SAM"
                    objsam.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                    'Rajkumar
                Case "PRE_SHIPMENT"
                    objPreShipment.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_SUPP_PRICE"
                    objSuppPrice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_FRM_APP_ACC"
                    objApportionAccural.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "GEN_FWD_COVER"
                    objForwardCover.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            End Select
        Catch ex As Exception
            oApplication.SetStatusBarMessage(ex.Message)
        End Try
    End Sub

    Private Sub oApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles oApplication.ItemEvent
        Try
            Select Case pVal.FormTypeEx
                Case "GEN_SIZE_MST"
                    objSizeMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_SAM"
                    objsam.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_ASSORTMENT"
                    objAssortedMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "SOORD"
                    objSalesOrder.ItemEvent_SalesOrder_Allocation(FormUID, pVal, BubbleEvent)
                Case "GEN_SZ_ORDR"
                    objSalesOrder.ItemEvent_SalesOrder_SizeMatrix(FormUID, pVal, BubbleEvent)
                Case "GEN_CST_PRCS"
                    objGEN_CST_PRCS.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "150"
                    objItemMasterData.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "-150"
                    objItemMasterData.ItemEvent_udf(FormUID, pVal, BubbleEvent)
                Case "GEN_COST_SHEET"
                    objGEN_COST_SHEET.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "20700"
                    objUser.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "62"
                    'objWhs.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "806"
                    objCOA.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "392"
                    objJE.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "170"
                    objIncomingPayments.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "426"
                    objOutgoingPayments.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "182"
                    objPurchaseReturn.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "142"
                    objPurchaseOrder.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "179"
                    objARCreditMemo.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "133"
                    objARInvoice.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "140"
                    objDelivery.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "139"
                    objSalesOrder.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "149"
                    objSalesQuotation.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "180"
                    objSalesReturn.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_SIZE_CODE"
                    objGEN_SIZE_CODE.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_QLTY_CODE"
                    objGEN_QLTY_CODE.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_COLOR_CODE"
                    objGEN_COLOR_CODE.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_STYLE_CODE"
                    objGEN_STYLE_CODE.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_SUB_TYPE"
                    objGEN_SUB_TYPE.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_CUST_CODE"
                    objGEN_CUST_CODE.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_ITM_TYPE"
                    objItemType.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_ITM_MST"
                    objItemMst.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_PROD_PRCS"
                    objProductionProcess.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_CUST_BOM"
                    objCustomBOM.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_MREQ"
                    objMaterialRequisition.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_PTN"
                    objPTN.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_UNIT_MST"
                    objUnitMaster.ItemEvent(FormUID, pVal, BubbleEvent)
               
                Case "GEN_GRPO_LCOSTS"
                    objGRPO.ItemEvent_LC(FormUID, pVal, BubbleEvent)
                Case "940"
                    objInventoryTransfer.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "-940"
                    objInventoryTransfer.ItemEvent_Intran(FormUID, pVal, BubbleEvent)
                Case "GEN_SCForm"
                    objSubContract.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_SCGRPO"
                    objSubContract_GRPO.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_SCDC"
                    objSubContract_DC.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_SCRET"
                    objSubContract_Return.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_WGW"
                    objWorOrdrGenWiz.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "65213"
                    objIssueForProduction.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "65214"
                    objReceiptFromProduction.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_FIN_SETUP"
                    objFinishSetup.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_FIN_DESCR"
                    objFinishScreen.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "143"
                    objGRPO.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "141"
                    objAPInvoice.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "181"
                    objAPCreditMemo.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_MACH_POOL"
                    objMachinePool.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_MACH_ALLOC"
                    objMachineAllocation.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_CAP_PLAN"
                    objMachineAllocation.ChildForm_ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_PARAM_MST"
                    objParamMst.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_ITEM_CREATE"
                    objItemCreate.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "142"
                    objPurchaseOrder.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_STH_DESCR"
                    objStitchingScreen.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Vijeesh
                Case "720"
                    objGoodsIssue.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "721"
                    objGoodsReceipt.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "UPLBOM"
                    objClsUploadBOM.ItemEvent(FormUID, pVal, BubbleEvent)

                    'Vijeesh
                Case "SAM"
                    objSAMRevaluation.ItemEvent(FormUID, pVal, BubbleEvent)

                    'Rajkumar
                Case "ACCRUALS_ORDER"
                    objSalesOrder.ItemEvent_Accrual_Form(FormUID, pVal, BubbleEvent)
                Case "ACCRUALS_PRE"
                    objPreShipment.ItemEvent_Accrual_Form(FormUID, pVal, BubbleEvent)
                Case "ACCRUALS"
                    objARInvoice.ItemEvent_Accrual_Form(FormUID, pVal, BubbleEvent)
                Case "PRE_SHIPMENT"
                    objPreShipment.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "9999"
                    objARInvoice.ItemEvent_Pre(FormUID, pVal, BubbleEvent)
                Case "425"
                    objARInvoice.ItemEvent_Pre(FormUID, pVal, BubbleEvent)
                Case "GEN_SUPP_PRICE"
                    objSuppPrice.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "PRE_FREIGHT"
                    objPreShipment.ItemEvent_Freight_Form(FormUID, pVal, BubbleEvent)
                Case "866"
                    objPreShipment.ItemEvent_exd(FormUID, pVal, BubbleEvent)
                Case "GEN_FRM_APP_ACC"
                    objApportionAccural.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Rajkumar 18.08.14
                Case "UBG_DBK_LST"
                    objAppDbk.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GEN_FWD_COVER"
                    objForwardCover.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ORCT_JV"
                    objIncomingPayments.ItemEvent_OtherCharges(FormUID, pVal, BubbleEvent)
                Case "OVPM_JV"
                    objOutgoingPayments.ItemEvent_OtherCharges(FormUID, pVal, BubbleEvent)
            End Select
            If pVal.FormTypeEx = "149" Or pVal.FormTypeEx = "139" Or pVal.FormTypeEx = "140" Or pVal.FormTypeEx = "180" Or pVal.FormTypeEx = "65308" Or pVal.FormTypeEx = "65300" Or pVal.FormTypeEx = "133" Or pVal.FormTypeEx = "179" Then
                Try
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD And pVal.BeforeAction = True Then
                        FilterCustomers(FormUID)
                        FilterCustomersName(FormUID)
                        'FilterItems(FormUID)
                    End If
                Catch ex As Exception
                    oApplication.StatusBar.SetText(ex.Message)
                End Try
            End If
           
            If pVal.FormTypeEx = "142" Or pVal.FormTypeEx = "143" Or pVal.FormTypeEx = "182" Or pVal.FormTypeEx = "65309" Or pVal.FormTypeEx = "65301" Or pVal.FormTypeEx = "141" Or pVal.FormTypeEx = "181" Then
                Try
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD And pVal.BeforeAction = True Then
                        Dim objForm As SAPbouiCOM.Form
                        objForm = oApplication.Forms.Item(FormUID)
                        FilterVendors(FormUID)
                        FilterVendorsName(FormUID)
                        'FilterItems(FormUID)
                    End If
                Catch ex As Exception
                    oApplication.StatusBar.SetText(ex.Message)
                End Try
            End If

            'If pVal.FormTypeEx = "170" Or pVal.FormTypeEx = "426" Then
            '    Try
            '        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD And pVal.BeforeAction = True Then
            '            Dim objForm As SAPbouiCOM.Form
            '            objForm = oApplication.Forms.Item(FormUID)
            '            FilterVendorsBanking(FormUID)
            '   End If
            'Catch ex As Exception
            '    oApplication.StatusBar.SetText(ex.Message)
            'End Try
            ' End If
        Catch ex As Exception
            oApplication.SetStatusBarMessage(ex.Message)
        End Try
    End Sub

    Sub FilterVendorsBanking(ByVal FormUID As String)
        Try
            Dim objForm As SAPbouiCOM.Form
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("14")
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Distinct U_unit From [@GEN_USR_UNIT] Where U_user = '" + oCompany.UserName.ToString.Trim + "'")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            For IntICount As Integer = 0 To oRecordSet.RecordCount - 1
                If IntICount = (oRecordSet.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "U_unit"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("U_unit").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "U_unit"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("U_unit").Value
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                oRecordSet.MoveNext()
            Next
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterCustomers(ByVal FormUID As String)
        Try
            Dim objForm As SAPbouiCOM.Form
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("2")
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Distinct U_unit From [@GEN_USR_UNIT] Where U_user = '" + oCompany.UserName.ToString.Trim + "'")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            For IntICount As Integer = 0 To oRecordSet.RecordCount - 1
                If IntICount = (oRecordSet.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "U_unit"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("U_unit").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "U_unit"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("U_unit").Value
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                oRecordSet.MoveNext()
            Next
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterCustomersName(ByVal FormUID As String)
        Try
            Dim objForm As SAPbouiCOM.Form
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("3")
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Distinct U_unit From [@GEN_USR_UNIT] Where U_user = '" + oCompany.UserName.ToString.Trim + "'")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            For IntICount As Integer = 0 To oRecordSet.RecordCount - 1
                If IntICount = (oRecordSet.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "U_unit"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("U_unit").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "U_unit"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("U_unit").Value
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                oRecordSet.MoveNext()
            Next
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterVendors(ByVal FormUID As String)
        Try
            Dim objForm As SAPbouiCOM.Form
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("2")
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Distinct U_unit From [@GEN_USR_UNIT] Where U_user = '" + oCompany.UserName.ToString.Trim + "'")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            For IntICount As Integer = 0 To oRecordSet.RecordCount - 1
                If IntICount = (oRecordSet.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "U_unit"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("U_unit").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "U_unit"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("U_unit").Value
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                oRecordSet.MoveNext()
            Next
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterVendorsName(ByVal FormUID As String)
        Try
            Dim objForm As SAPbouiCOM.Form
            objForm = oApplication.Forms.Item(FormUID)
            Dim emptyConds As New SAPbouiCOM.Conditions
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("3")
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select Distinct U_unit From [@GEN_USR_UNIT] Where U_user = '" + oCompany.UserName.ToString.Trim + "'")
            oCFL.SetConditions(Nothing)
            oCons = oCFL.GetConditions
            For IntICount As Integer = 0 To oRecordSet.RecordCount - 1
                If IntICount = (oRecordSet.RecordCount - 1) Then
                    oCon = oCons.Add()
                    oCon.Alias = "U_unit"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("U_unit").Value
                Else
                    oCon = oCons.Add()
                    oCon.Alias = "U_unit"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRecordSet.Fields.Item("U_unit").Value
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                oRecordSet.MoveNext()
            Next
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub oApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles oApplication.MenuEvent
        Try
            objSizeMaster.MenuEvent(pVal, BubbleEvent)
            ' objCOA.MenuEvent(pVal, BubbleEvent)
            objsam.MenuEvent(pVal, BubbleEvent)
            objAssortedMaster.MenuEvent(pVal, BubbleEvent)
            objIncomingPayments.MenuEvent(pVal, BubbleEvent)
            objGEN_CST_PRCS.MenuEvent(pVal, BubbleEvent)
            objGEN_COST_SHEET.MenuEvent(pVal, BubbleEvent)
            objAPCreditMemo.MenuEvent(pVal, BubbleEvent)
            objPurchaseReturn.MenuEvent(pVal, BubbleEvent)
            objARInvoice.MenuEvent(pVal, BubbleEvent)
            objGRPO.MenuEvent(pVal, BubbleEvent)
            objPurchaseOrder.MenuEvent(pVal, BubbleEvent)
            objARCreditMemo.MenuEvent(pVal, BubbleEvent)
            objARInvoice.MenuEvent(pVal, BubbleEvent)
            objDelivery.MenuEvent(pVal, BubbleEvent)
            objSalesOrder.MenuEvent(pVal, BubbleEvent)
            objSalesQuotation.MenuEvent(pVal, BubbleEvent)
            objSalesReturn.MenuEvent(pVal, BubbleEvent)
            objGEN_SIZE_CODE.MenuEvent(pVal, BubbleEvent)
            objGEN_QLTY_CODE.MenuEvent(pVal, BubbleEvent)
            objGEN_COLOR_CODE.MenuEvent(pVal, BubbleEvent)
            objGEN_STYLE_CODE.MenuEvent(pVal, BubbleEvent)
            objGEN_SUB_TYPE.MenuEvent(pVal, BubbleEvent)
            objGEN_CUST_CODE.MenuEvent(pVal, BubbleEvent)
            objItemType.MenuEvent(pVal, BubbleEvent)
            objItemMst.MenuEvent(pVal, BubbleEvent)
            objProductionProcess.MenuEvent(pVal, BubbleEvent)
            objCustomBOM.MenuEvent(pVal, BubbleEvent)
            objMaterialRequisition.MenuEvent(pVal, BubbleEvent)
            objPTN.MenuEvent(pVal, BubbleEvent)
            objUnitMaster.MenuEvent(pVal, BubbleEvent)

            objGRPOLC.MenuEvent(pVal, BubbleEvent)
            objInventoryTransfer.MenuEvent(pVal, BubbleEvent)
            objSubContract.MenuEvent(pVal, BubbleEvent)
            objSubContract_GRPO.MenuEvent(pVal, BubbleEvent)
            objSubContract_DC.MenuEvent(pVal, BubbleEvent)
            objSubContract_Return.MenuEvent(pVal, BubbleEvent)
            objWorOrdrGenWiz.MenuEvent(pVal, BubbleEvent)
            objReceiptFromProduction.MenuEvent(pVal, BubbleEvent)
            objIssueForProduction.MenuEvent(pVal, BubbleEvent)
            objFinishSetup.MenuEvent(pVal, BubbleEvent)
            objFinishScreen.MenuEvent(pVal, BubbleEvent)
            objStitchingScreen.MenuEvent(pVal, BubbleEvent)
            objForwardCover.MenuEvent(pVal, BubbleEvent)
            If pVal.MenuUID = "GEN_FIELD_ID" And pVal.BeforeAction = False Then
                Dim Rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rs.DoQuery("Select 47616 + count(*) MenuItem from OUDO where CanDefForm = 'Y' and code <= 'GEN_FIELD_ID'")
                oApplication.ActivateMenuItem(Rs.Fields.Item("MenuItem").Value)
            End If
            If pVal.MenuUID = "GEN_FIN_PRCS" And pVal.BeforeAction = False Then
                Dim Rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rs.DoQuery("Select 47616 + count(*) MenuItem from OUDO where CanDefForm = 'Y' and code <= 'GEN_FIN_PRCS'")
                oApplication.ActivateMenuItem(Rs.Fields.Item("MenuItem").Value)
            End If
            If pVal.MenuUID = "GEN_LINE_MST" And pVal.BeforeAction = False Then
                Dim Rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rs.DoQuery("Select 47616 + count(*) MenuItem from OUDO where CanDefForm = 'Y' and code <= 'GEN_LINE_MST'")
                oApplication.ActivateMenuItem(Rs.Fields.Item("MenuItem").Value)
            End If
            If pVal.MenuUID = "GEN_LINE_TYPE" And pVal.BeforeAction = False Then
                Dim Rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rs.DoQuery("Select 47616 + count(*) MenuItem from OUDO where CanDefForm = 'Y' and code <= 'GEN_LINE_TYPE'")
                oApplication.ActivateMenuItem(Rs.Fields.Item("MenuItem").Value)
            End If
            If pVal.MenuUID = "GEN_PROCESS_MST" And pVal.BeforeAction = False Then
                Dim Rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rs.DoQuery("Select 47616 + count(*) MenuItem from OUDO where CanDefForm = 'Y' and code <= 'GEN_PROCESS_MST'")
                oApplication.ActivateMenuItem(Rs.Fields.Item("MenuItem").Value)
            End If
            If pVal.MenuUID = "GEN_STH_PRCS" And pVal.BeforeAction = False Then
                Dim Rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rs.DoQuery("Select 47616 + count(*) MenuItem from OUDO where CanDefForm = 'Y' and code <= 'GEN_STH_PRCS'")
                oApplication.ActivateMenuItem(Rs.Fields.Item("MenuItem").Value)
            End If
            If pVal.MenuUID = "GEN_USR_UNIT" And pVal.BeforeAction = False Then
                Dim Rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rs.DoQuery("Select 47616 + count(*) MenuItem from OUDO where CanDefForm = 'Y' and code <= 'GEN_USR_UNIT'")
                oApplication.ActivateMenuItem(Rs.Fields.Item("MenuItem").Value)
            End If
            If pVal.MenuUID = "GEN_WHS_USR" And pVal.BeforeAction = False Then
                Dim Rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rs.DoQuery("Select 47616 + count(*) MenuItem from OUDO where CanDefForm = 'Y' and code <= 'GEN_WHS_USR'")
                oApplication.ActivateMenuItem(Rs.Fields.Item("MenuItem").Value)
            End If
            objMachinePool.MenuEvent(pVal, BubbleEvent)
            objMachineAllocation.MenuEvent(pVal, BubbleEvent)
            objParamMst.MenuEvent(pVal, BubbleEvent)
            objItemCreate.MenuEvent(pVal, BubbleEvent)
            'Vijeesh
            objGoodsIssue.MenuEvent(pVal, BubbleEvent)
            objGoodsReceipt.MenuEvent(pVal, BubbleEvent)
            objClsUploadBOM.MenuEvent(pVal, BubbleEvent)
            'Vijeesh
            objSAMRevaluation.MenuEvent(pVal, BubbleEvent)

            'Rajkumar
            objPreShipment.MenuEvent(pVal, BubbleEvent)
            objSuppPrice.MenuEvent(pVal, BubbleEvent)
            objApportionAccural.MenuEvent(pVal, BubbleEvent)

            'Rajkumar 18.08.14
            objAppDbk.MenuEvent(pVal, BubbleEvent)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub oApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles oApplication.RightClickEvent
        Try
            Dim oForm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
            Select Case oForm.TypeEx
                Case "GEN_ASSORTMENT"
                    objAssortedMaster.RightClickEvent(eventInfo, BubbleEvent)
                Case "GEN_COST_SHEET"
                    objGEN_COST_SHEET.RightClickEvent(eventInfo, BubbleEvent)
                Case "GEN_PROD_PRCS"
                    objProductionProcess.RightClickEvent(eventInfo, BubbleEvent)
                Case "GEN_CUST_BOM"
                    objCustomBOM.RightClickEvent(eventInfo, BubbleEvent)
                Case "GEN_MREQ"
                    objMaterialRequisition.RightClickEvent(eventInfo, BubbleEvent)
                Case "GEN_UNIT_MST"
                    objUnitMaster.RightClickEvent(eventInfo, BubbleEvent)
               
                Case "GEN_GRPO_LCOSTS"
                    objGRPOLC.RightClickEvent(eventInfo, BubbleEvent)
                Case "GEN_SCForm"
                    objSubContract.RightClickEvent(eventInfo, BubbleEvent)
                Case "GEN_SCGRPO"
                    objSubContract_GRPO.RightClickEvent(eventInfo, BubbleEvent)
                Case "GEN_SCDC"
                    objSubContract_DC.RightClickEvent(eventInfo, BubbleEvent)
                Case "GEN_FIN_SETUP"
                    objFinishSetup.RightClickEvent(eventInfo, BubbleEvent)
                Case "GEN_FIN_DESCR"
                    objFinishScreen.RightClickEvent(eventInfo, BubbleEvent)
                Case "GEN_MACH_POOL"
                    objMachinePool.RightClickEvent(eventInfo, BubbleEvent)
                Case "GEN_MACH_ALLOC"
                    objMachineAllocation.RightClickEvent(eventInfo, BubbleEvent)
                Case "GEN_PARAM_MST"
                    objParamMst.RightClickEvent(eventInfo, BubbleEvent)
                Case "GEN_STH_DESCR"
                    objStitchingScreen.RightClickEvent(eventInfo, BubbleEvent)
                    'Vijeesh
                Case "150"
                    objItemMasterData.RightClickEvent(eventInfo, BubbleEvent)
                    'Vijeesh
                Case "SAM"
                    objSAMRevaluation.RightClickEvent(eventInfo, BubbleEvent)
                Case "GEN_SAM"
                    objsam.RightClickEvent(eventInfo, BubbleEvent)

                    'Rajkumar
                Case "PRE_SHIPMENT"
                    objPreShipment.RightClickEvent(eventInfo, BubbleEvent)
                Case "GEN_SUPP_PRICE"
                    objSuppPrice.RightClickEvent(eventInfo, BubbleEvent)
                Case "GEN_FWD_COVER"
                    objForwardCover.RightClickEvent(eventInfo, BubbleEvent)
            End Select
        Catch ex As Exception
            oApplication.SetStatusBarMessage(ex.Message)
        End Try
    End Sub

End Module

