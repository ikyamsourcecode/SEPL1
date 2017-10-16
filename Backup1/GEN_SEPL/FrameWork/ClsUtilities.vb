Imports System.Reflection

Public Class ClsUtilities

    Dim v_RetVal, v_ErrCode As Integer
    Dim v_ErrMsg As String
    Dim DB_Restart As Boolean = False

    Sub StartUp()
        SetoApplication()
        If Not SetConnectionContext() = 0 Then
            oApplication.MessageBox("Failed setting a connection to DI API")
            End
        End If
        SAPXML("Menu.xml")
        Me.GEN_Tables()
        oApplication.StatusBar.SetText("Genisys Add-On Connected successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    End Sub

#Region "Company Connection"

    Private Sub SetoApplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String
        SboGuiApi = New SAPbouiCOM.SboGuiApi
        sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
        Try
            SboGuiApi.Connect(sConnectionString)
        Catch
            System.Windows.Forms.MessageBox.Show("No SAP Business One oApplication was found")
            End
        End Try
        oApplication = SboGuiApi.GetApplication()
    End Sub

    Private Function SetConnectionContext() As Integer
        oCompany = oApplication.Company.GetDICompany()
    End Function

#End Region

    Sub SAPXML(ByVal path As String, Optional ByVal CHILD_FORM As String = "")
        Try
            Dim xmldoc As New Xml.XmlDocument
            Dim Streaming As System.IO.Stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("GEN_SEPL." + path)
            Dim StreamRead As New System.IO.StreamReader(Streaming, True)
            xmldoc.LoadXml(StreamRead.ReadToEnd)
            StreamRead.Close()
            If Not xmldoc.SelectSingleNode("//form") Is Nothing Then
                If Trim(CHILD_FORM).Equals("") = True Then
                    Dim r As New Random
                    xmldoc.SelectSingleNode("//form").Attributes.GetNamedItem("uid").Value = xmldoc.SelectSingleNode("//form").Attributes.GetNamedItem("uid").Value & "_" & r.Next(100)
                Else
                    xmldoc.SelectSingleNode("//form").Attributes.GetNamedItem("uid").Value = CHILD_FORM
                End If
            End If
            oApplication.LoadBatchActions(xmldoc.InnerXml)
        Catch ex As Exception
            oApplication.MessageBox(ex.Message)
        End Try
    End Sub

    Function getNextSeriesVal(ByVal udoID As String) As Integer
        Try
            Dim seriesService As SAPbobsCOM.SeriesService
            Dim v_CompanyService As SAPbobsCOM.CompanyService
            Dim objectType As SAPbobsCOM.DocumentTypeParams
            Dim crmSeries As SAPbobsCOM.Series
            v_CompanyService = oCompany.GetCompanyService
            seriesService = v_CompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.SeriesService)
            objectType = seriesService.GetDataInterface(SAPbobsCOM.SeriesServiceDataInterfaces.ssdiDocumentTypeParams)
            objectType.Document = udoID
            crmSeries = seriesService.GetDefaultSeries(objectType)
            Return crmSeries.NextNumber
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Sub GetSeries(ByVal FormUID As String, ByVal ItemUID As String, ByVal ObjectType As String)
        Try
            Dim objForm As SAPbouiCOM.Form = oApplication.Forms.Item(FormUID)
            Dim objCombo As SAPbouiCOM.ComboBox = objForm.Items.Item(ItemUID).Specific
            Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Select Series,SeriesName from NNM1 Where ObjectCode='" & Trim(ObjectType) & "'")
            If objCombo.ValidValues.Count = 0 Then
                For Row As Integer = 1 To oRS.RecordCount
                    objCombo.ValidValues.Add(oRS.Fields.Item("Series").Value, oRS.Fields.Item("SeriesName").Value)
                    oRS.MoveNext()
                Next
            End If
            oRS.DoQuery("Select DfltSeries from ONNM Where ObjectCode='" & Trim(ObjectType) & "'")
            If objCombo.ValidValues.Count > 0 Then objCombo.Select(Trim(oRS.Fields.Item("DfltSeries").Value), SAPbouiCOM.BoSearchKey.psk_ByValue)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function keygencode(ByVal vtablename As String, Optional ByVal prefix As String = "DOC-") As String

        Dim str As String = ""
        Dim Query As String
        Try
            Query = "SELECT MAX(CAST(Code AS int)) AS code FROM [" + vtablename + "]"
            Dim v_recordset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            v_recordset.DoQuery(Query)
            v_recordset.MoveFirst()
            Dim code As Integer = v_recordset.Fields.Item("code").Value.ToString
            If code > 0 Then
                code += 1
                Dim docid As String = prefix
                If code.ToString.Length < 6 Then
                    For count As Integer = 0 To 5 - code.ToString.Length
                        docid += "0"
                    Next
                End If
                docid += code.ToString
                str = code
            Else
                str = "1"
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_recordset)
            v_recordset = Nothing
            GC.Collect()
        Catch ex As Exception
            oApplication.MessageBox(ex.Message)
        End Try
        keygencode = str
    End Function

    Sub GEN_Tables()
        Me.UDF()
        If DB_Restart = True Then
            oApplication.MessageBox("Please re-login to update database structure")
        End If
    End Sub

#Region "       -- UDF --          "

    Sub UDF()
        Try
            Me.SAP_TABLES_UDF()
            GEN_CST_PRCS()
            Me.GEN_SAM_REV()
            Me.GEN_SAM()
            User_Unit_Linkage()
            GEN_COST_SHEET()
            Me.GEN_UNIT_MST()
            Me.UBG_TARIF_MST()
            Me.UBG_GARMENT_MST()
            Me.UBG_DBK_LST()
            Me.Forward_Cover()
            Me.Forward_Rem()

            Me.GEN_GRPO_LCOSTS()

            Me.GEN_PROCESS_MST()
            Me.GEN_PROD_PROCESS()
            Me.GEN_CUST_BOM()
            Me.Assortment_Master()
            Me.Size_Master()
            Me.Size_Matrix_Order()
            Me.Ordr_Tmp()
            Me.GEN_FIN_PRCS()
            Me.GEN_STH_PRCS()
            Me.GEN_FIN_SETUP()
            Me.GEN_FIN_DESCR()            
            Me.GEN_STH_DESCR()
            Me.GEN_LINE_MST()
            Me.GEN_LINE_TYPE()
            Me.GEN_MACH_POOL()
            Me.GEN_MACH_ALLOC()
            Me.GEN_CAP_PLAN()
            Me.MaterialRequisitionNote()
            Me.Warehouse_User_Alert()
            Me.Production_Transfer_Note()
            Me.SubContracting()
            Me.SubContracting_GRPO()
            Me.SubContractDeliveryChallan()
            Me.ARInvoice_GLAccount()
            Me.Journal_Entry()

            Me.SubContractReturn()

            Me.GEN_ITM_TYPE()
            Me.GEN_ITM_MST()
            Me.GEN_PARAM_MST()
            Me.GEN_SUB_TYPE()
            Me.GEN_CUST_CODE()
            Me.GEN_STYLE_CODE()
            Me.GEN_COLOR_CODE()
            Me.GEN_SIZE_CODE()
            Me.GEN_FIELD_ID()
            Me.GEN_QLTY_CODE()
            Me.GEN_PARAM_MST()
            Me.GEN_ITEM_CREATE()            

            Me.Transportaion_Master()
            Me.Sales_Freight_Accruals()
            Me.Sales_Freight_Accruals_Order()
            Me.Sales_Freight_Accruals_Pre()
            Me.Pre_Freight_Matrix()
            Me.GEN_SPPL_PRICE()
            Me.No_Objects_Tables()
            Me.DocumentSeriesAuthorisation()
            Me.Sales_Matrix_Order()
            Me.ApportioningAccural()
            Me.Pre_Shipment()

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub Transportaion_Master()
        Me.AddTable("GEN_FREIGHT", "Gen->Transport Master", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        Me.AddColumns("@GEN_FREIGHT", "SOURCE", "Source", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_FREIGHT", "PR_OF_LOD", "Port Of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_FREIGHT", "NO_OF_CON", "No Of Container", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        Me.AddColumns("@GEN_FREIGHT", "AMOUNT", "Transp Amount", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Sum)
    End Sub
    Sub Sales_Freight_Accruals()
        Me.AddTable("GEN_ACCRUALS", "Gen->Accruals", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        Me.AddColumns("@GEN_ACCRUALS", "invno", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_ACCRUALS", "macid", "MAC ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_ACCRUALS", "fcode", "Freight Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_ACCRUALS", "fname", "Freight Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_ACCRUALS", "tax", "Tax Applicable", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_ACCRUALS", "amount", "Freight Amount", SAPbobsCOM.BoFieldTypes.db_Alpha, 60)
        Me.AddColumns("@GEN_ACCRUALS", "posfrgt", "Positive Freight", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_ACCRUALS", "postax", "Positive Tax", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_ACCRUALS", "negfrgt", "Negative Freight", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_ACCRUALS", "negtax", "Negative Tax", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
    End Sub
    Sub Sales_Freight_Accruals_Order()
        Me.AddTable("GEN_ACCRUALS_RDR", "Gen->Accruals Order", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        Me.AddColumns("@GEN_ACCRUALS_RDR", "invno", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_ACCRUALS_RDR", "macid", "MAC ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_ACCRUALS_RDR", "fcode", "Freight Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_ACCRUALS_RDR", "fname", "Freight Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_ACCRUALS_RDR", "tax", "Tax Applicable", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_ACCRUALS_RDR", "amount", "Freight Amount", SAPbobsCOM.BoFieldTypes.db_Alpha, 60)
        Me.AddColumns("@GEN_ACCRUALS_RDR", "posfrgt", "Positive Freight", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_ACCRUALS_RDR", "postax", "Positive Tax", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_ACCRUALS_RDR", "negfrgt", "Negative Freight", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_ACCRUALS_RDR", "negtax", "Negative Tax", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
    End Sub
    Sub Sales_Freight_Accruals_Pre()
        Me.AddTable("GEN_ACCRUALS_PRE", "Gen->Accruals Order", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        Me.AddColumns("@GEN_ACCRUALS_PRE", "invno", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_ACCRUALS_PRE", "macid", "MAC ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_ACCRUALS_PRE", "fcode", "Freight Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_ACCRUALS_PRE", "fname", "Freight Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_ACCRUALS_PRE", "tax", "Tax Applicable", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_ACCRUALS_PRE", "amount", "Freight Amount", SAPbobsCOM.BoFieldTypes.db_Alpha, 60)
        Me.AddColumns("@GEN_ACCRUALS_PRE", "posfrgt", "Positive Freight", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_ACCRUALS_PRE", "postax", "Positive Tax", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_ACCRUALS_PRE", "negfrgt", "Negative Freight", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_ACCRUALS_PRE", "negtax", "Negative Tax", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
    End Sub
    Sub ApportioningAccural()
        Me.AddTable("GEN_APP_ACC", "Gen->Apportioning Accural", SAPbobsCOM.BoUTBTableType.bott_Document)
        Me.AddTable("GEN_APP_ACC_D0", "Gen->Apport Accural Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        Me.AddColumns("@GEN_APP_ACC", "DOC_DATE", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@GEN_APP_ACC", "POS_DATE", "Posting Date", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@GEN_APP_ACC", "DATEFROM", "Date From", SAPbobsCOM.BoFieldTypes.db_Date, 30)
        Me.AddColumns("@GEN_APP_ACC", "DATE_TO", "Date To", SAPbobsCOM.BoFieldTypes.db_Date, 0)
        Me.AddColumns("@GEN_APP_ACC", "PC", "Profit Center", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
        Me.AddColumns("@GEN_APP_ACC", "JV_NO", "Journal Vouch No", SAPbobsCOM.BoFieldTypes.db_Alpha, 12)
        Me.AddColumns("@GEN_APP_ACC", "REMARKS", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254)

        Me.AddColumns("@GEN_APP_ACC_D0", "DOC_TYPE", "Doc Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
        Me.AddColumns("@GEN_APP_ACC_D0", "SELECT", "Select", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
        Me.AddColumns("@GEN_APP_ACC_D0", "PC", "Profit Center", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
        Me.AddColumns("@GEN_APP_ACC_D0", "INV_SER", "Invoice Series", SAPbobsCOM.BoFieldTypes.db_Alpha, 12, SAPbobsCOM.BoFldSubTypes.st_None)
        Me.AddColumns("@GEN_APP_ACC_D0", "INV_NO", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, 12, SAPbobsCOM.BoFldSubTypes.st_None)
        Me.AddColumns("@GEN_APP_ACC_D0", "INV_ENT", "Invoice DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, 12, SAPbobsCOM.BoFldSubTypes.st_None)
        Me.AddColumns("@GEN_APP_ACC_D0", "INV_DATE", "Invoice Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, )
        Me.AddColumns("@GEN_APP_ACC_D0", "UNIT_NO", "Unit No", SAPbobsCOM.BoFieldTypes.db_Alpha, 8, SAPbobsCOM.BoFldSubTypes.st_None)

        If Not Me.UDOExists("GEN_UDO_APP_ACC") Then
            Dim findAliasNDescription As String(,) = {{"DocEntry", "DocEntry"}}
            Me.registerUDO("GEN_UDO_APP_ACC", "GEN_APP_ACC", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "GEN_APP_ACC", "GEN_APP_ACC_D0", "", "", SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO)
            findAliasNDescription = Nothing
        End If
    End Sub
    Sub Pre_Freight_Matrix()
        Me.AddTable("UBG_PRE_FRET", "Freight Pre-Shipment", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        Me.AddTable("UBG_PRE_FRET_D0", "Pre-Freight Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)

        Me.AddColumns("@UBG_PRE_FRET_D0", "MacId", "MAC ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@UBG_PRE_FRET_D0", "FreCode", "Freight Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@UBG_PRE_FRET_D0", "FreName", "Freight Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@UBG_PRE_FRET_D0", "TaxCode", "Tax Applicable", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@UBG_PRE_FRET_D0", "Amt", "Freight Amount", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@UBG_PRE_FRET_D0", "TotTax", "Total Tax Amount", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        'Me.AddColumns("@UBG_PRE_FRET_D0", "Amt", "Freight Amount", SAPbobsCOM.BoFieldTypes.db_Float, 30, SAPbobsCOM.BoFldSubTypes.st_Sum)
        'Me.AddColumns("@UBG_PRE_FRET_D0", "TotTax", "Total Tax Amount", SAPbobsCOM.BoFieldTypes.db_Float, 30, SAPbobsCOM.BoFldSubTypes.st_Rate)
        Me.AddColumns("@UBG_PRE_FRET_D0", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
        Me.AddColumns("@UBG_PRE_FRET_D0", "PreNo", "Pre-Shipment No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@UBG_PRE_FRET_D0", "Rem", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

        'If Not Me.UDOExists("UBG_PRE_FRET") Then
        '    Dim findAliasNDescription = New String(,) {{"DocEntry", "DocEntry"}}
        '    Me.registerUDO("UBG_PRE_FRET", "Freight Charges", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "UBG_PRE_FRET", "UBG_PRE_FRET_D0", "")
        '    findAliasNDescription = Nothing
        'End If
    End Sub
    Sub Forward_Cover()
        Me.AddTable("GEN_FWD_COVER", "Gen->Forward Cover", SAPbobsCOM.BoUTBTableType.bott_Document)
        Me.AddTable("GEN_FWD_COVER_D0", "Gen->Forward Cover Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        Me.AddColumns("@GEN_FWD_COVER", "docdate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@GEN_FWD_COVER", "fdate", "From Date", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@GEN_FWD_COVER", "tdate", "To Date", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@GEN_FWD_COVER", "doccur", "Document Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_FWD_COVER", "contrno", "Contract No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_FWD_COVER", "amount", "Amount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
        Me.AddColumns("@GEN_FWD_COVER", "docrate", "Document Rate", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)
        Dim ValidValues = New String(,) {{"1", "EFC"}, {"2", "Remittance"}, {"3", "Re-Booking against Cancellation"}}
        Me.AddColumns("@GEN_FWD_COVER", "remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, Nothing)
        Me.AddColumns("@GEN_FWD_COVER", "sptrate", "Spot Rate", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)
        Me.AddColumns("@GEN_FWD_COVER", "unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
        ValidValues = New String(,) {{"Encash", "Encash"}, {"Cancelled", "Cancelled"}}
        Me.AddColumns("@GEN_FWD_COVER", "status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, Nothing)
        Me.AddColumns("@GEN_FWD_COVER_D0", "docdate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@GEN_FWD_COVER_D0", "amount", "Amount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
        Me.AddColumns("@GEN_FWD_COVER_D0", "remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254)
        Me.AddColumns("@GEN_FWD_COVER_D0", "pc", "PC", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        If Not Me.UDOExists("GEN_FWD_COVER") Then
            Dim findAliasNDescription As String(,) = {{"DocEntry", "DocEntry"}}
            Me.registerUDO("GEN_FWD_COVER", "GEN_FWD_COVER", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "GEN_FWD_COVER", "GEN_FWD_COVER_D0", "", "", SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO)
            findAliasNDescription = Nothing
        End If
    End Sub
    Sub Forward_Rem()
        Me.AddTable("UBG_FWD_REM", "UBG-> Forward Remarks", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        Me.AddColumns("@UBG_FWD_REM", "year", "Financial Year", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
        Me.AddColumns("@UBG_FWD_REM", "month", "Month", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
        Me.AddColumns("@UBG_FWD_REM", "curr", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
        Me.AddColumns("@UBG_FWD_REM", "status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        Me.AddColumns("@UBG_FWD_REM", "unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
        Me.AddColumns("@UBG_FWD_REM", "actfc", "Actual FC", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
        Me.AddColumns("@UBG_FWD_REM", "act", "Actual", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
        Me.AddColumns("@UBG_FWD_REM", "balfc", "Balance FC", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
        Me.AddColumns("@UBG_FWD_REM", "bal", "Balance", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
        Me.AddColumns("@UBG_FWD_REM", "blend", "Blend Rate", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)
    End Sub
    Sub GEN_SPPL_PRICE()
        Me.AddTable("GEN_SUPP_PRICE", "Gen->Supplementary Price", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        Me.AddTable("GEN_SUPP_PRICE_D0", "Gen->Supplementary Price Items", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)

        Me.AddColumns("@GEN_SUPP_PRICE", "cardcode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_SUPP_PRICE", "cardname", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_SUPP_PRICE", "doccur", "Document Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

        Me.AddColumns("@GEN_SUPP_PRICE_D0", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_SUPP_PRICE_D0", "itemname", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_SUPP_PRICE_D0", "price", "Supplementary Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
        Me.AddColumns("@GEN_SUPP_PRICE_D0", "ansp", "ANSP Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
        Me.AddColumns("@GEN_SUPP_PRICE_D0", "comm", "Commission Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
        If Not Me.UDOExists("GEN_SUPP_PRICE") Then
            Dim findAliasNDescription As String(,) = {{"Code", "Code"}}
            Me.registerUDO("GEN_SUPP_PRICE", "GEN_SUPP_PRICE", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_SUPP_PRICE", "GEN_SUPP_PRICE_D0", "", "", SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO)
            findAliasNDescription = Nothing
        End If
    End Sub
    Sub Pre_Shipment()
        Me.AddTable("PRE_SHIPMENT", "Pre-shipment", SAPbobsCOM.BoUTBTableType.bott_Document)
        Me.AddTable("PRE_SHIPMENT_D0", "Pre-Shipment Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        Me.AddTable("PRE_SHIPMENT_D3", "Pre-Shipment Freight", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        Me.AddColumns("@PRE_SHIPMENT", "PreNo", "Pre-Shipment No", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@PRE_SHIPMENT", "Series", "Series", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
        Me.AddColumns("@PRE_SHIPMENT", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
        Me.AddColumns("@PRE_SHIPMENT", "PosDate", "Posting Date", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@PRE_SHIPMENT", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@PRE_SHIPMENT", "DelDate", "Delivey Date", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@PRE_SHIPMENT", "CustCode", "Cust Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
        Me.AddColumns("@PRE_SHIPMENT", "CustName", "Cust Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 35)
        Me.AddColumns("@PRE_SHIPMENT", "ConPer", "Contact Person", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@PRE_SHIPMENT", "CustRef", "Customer Ref. No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@PRE_SHIPMENT", "Unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
        Me.AddColumns("@PRE_SHIPMENT", "Fold", "Folder", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
        Me.AddColumns("@PRE_SHIPMENT", "DocCur", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
        Me.AddColumns("@PRE_SHIPMENT", "CF", "Cost & Freight", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
        Me.AddColumns("@PRE_SHIPMENT", "SaleNo", "SaleNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)

        Me.AddColumns("@PRE_SHIPMENT", "Owner", "Owner", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@PRE_SHIPMENT", "OwnCode", "Owner Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@PRE_SHIPMENT", "Buyer", "Buyer", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)

        Me.AddColumns("@PRE_SHIPMENT", "PayTrms", "PayTrms", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@PRE_SHIPMENT", "JourRem", "JourRem", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)

        Me.AddColumns("@PRE_SHIPMENT", "TotBefTa", "TotBefTax", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Sum)
        Me.AddColumns("@PRE_SHIPMENT", "Tax", "Tax", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Sum)
        Me.AddColumns("@PRE_SHIPMENT", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Sum)
        Me.AddColumns("@PRE_SHIPMENT", "Freight", "Freight", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Sum)
        Me.AddColumns("@PRE_SHIPMENT", "Round", "Round", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Sum)
        Me.AddColumns("@PRE_SHIPMENT", "Rond", "Rounding", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)

        Me.AddColumns("@PRE_SHIPMENT", "TotBefTa_A", "TotBefTax", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@PRE_SHIPMENT", "Tax_A", "Tax", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@PRE_SHIPMENT", "Total_A", "Total", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@PRE_SHIPMENT", "Freight_A", "Freight", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@PRE_SHIPMENT", "Round_A", "Round", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

        Me.AddColumns("@PRE_SHIPMENT", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)


        Me.AddColumns("@PRE_SHIPMENT_D0", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@PRE_SHIPMENT_D0", "ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 35)
        Me.AddColumns("@PRE_SHIPMENT_D0", "Quantity", "Committed Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
        Me.AddColumns("@PRE_SHIPMENT_D0", "Price", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
        Me.AddColumns("@PRE_SHIPMENT_D0", "Price_A", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@PRE_SHIPMENT_D0", "TotalLC", "TotalLC", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
        Me.AddColumns("@PRE_SHIPMENT_D0", "Total_A", "TotalLC", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@PRE_SHIPMENT_D0", "TaxAmt", "TaxAmt", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
        Me.AddColumns("@PRE_SHIPMENT_D0", "DocCur", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
        Me.AddColumns("@PRE_SHIPMENT_D0", "TaxCode", "TaxCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@PRE_SHIPMENT_D0", "Whse", "Warehouse Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@PRE_SHIPMENT_D0", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@PRE_SHIPMENT_D0", "Note", "Note", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        Me.AddColumns("@PRE_SHIPMENT_D0", "UOM", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
        Me.AddColumns("@PRE_SHIPMENT_D0", "SONo", "Sales Order No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@PRE_SHIPMENT_D0", "BaseRef", "BaseRef", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@PRE_SHIPMENT_D0", "BaseLine", "Base LineNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@PRE_SHIPMENT_D0", "TrgetNo", "Target Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@PRE_SHIPMENT_D0", "TaxRate", "TaxRate", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
        Me.AddColumns("@PRE_SHIPMENT_D0", "postno", "Post Shipment No", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@PRE_SHIPMENT_D0", "Source", "Source", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@PRE_SHIPMENT_D0", "MonoCart", "MonoCarton", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)

        Dim ValidValues = New String(,) {{"Open", "Open"}, {"Closed", "Closed"}}
        Dim DefaultVal = New String(,) {{"Open", "Open"}}
        Me.AddColumns("@PRE_SHIPMENT_D0", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)
        Me.AddColumns("@PRE_SHIPMENT_D0", "invqty", "Invoice Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
        Me.AddColumns("@PRE_SHIPMENT", "Addr", "Address", SAPbobsCOM.BoFieldTypes.db_Memo)

        Me.AddColumns("@PRE_SHIPMENT_D3", "ExpnCode", "Freight Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@PRE_SHIPMENT_D3", "ExpnName", "Freight Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 35)
        Me.AddColumns("@PRE_SHIPMENT_D3", "LineTot", "Line Total", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Sum)
        Me.AddColumns("@PRE_SHIPMENT_D3", "TotFrgn", "Total Freight", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Sum)
        Me.AddColumns("@PRE_SHIPMENT_D3", "TotSumSy", "Total System Currency", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Sum)
        Me.AddColumns("@PRE_SHIPMENT_D3", "TaxCode", "Tax Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@PRE_SHIPMENT_D3", "TaxType", "Tax Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
        Me.AddColumns("@PRE_SHIPMENT_D3", "Curr", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
        Me.AddColumns("@PRE_SHIPMENT_D3", "DocNum", "DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 35)



        Me.AddColumns("@PRE_SHIPMENT", "pc", "Profit Center", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@PRE_SHIPMENT", "user", "User", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        'Me.AddColumns("@PRE_SHIPMENT", "oinvno", "Original Inv No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@PRE_SHIPMENT", "postno", "Post Shipment No", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)

        'Me.AddColumns("@PRE_SHIPMENT", "VesslNo", "Vessel No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        'Me.AddColumns("@PRE_SHIPMENT", "PortLoad", "Port Of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        'Me.AddColumns("@PRE_SHIPMENT", "CntyFin", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        'Me.AddColumns("@PRE_SHIPMENT", "PorDisch", "Port Of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        'Me.AddColumns("@PRE_SHIPMENT", "FinDest", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, 60)
        'Me.AddColumns("@PRE_SHIPMENT", "GrssWt", "Gross Weight", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)
        'Me.AddColumns("@PRE_SHIPMENT", "NetWt", "Nett Weight", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)
        'Me.AddColumns("@PRE_SHIPMENT", "CBM", "CBM", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)
        'Me.AddColumns("@PRE_SHIPMENT", "TotCart", "Total Cartons", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)
        'Me.AddColumns("@PRE_SHIPMENT", "HAWB", "HAWB No", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        'Me.AddColumns("@PRE_SHIPMENT", "MAWB", "MAWB No", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        'Me.AddColumns("@PRE_SHIPMENT", "SBNo", "SB No", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        'Me.AddColumns("@PRE_SHIPMENT", "HAWBDt", "HAWB Date", SAPbobsCOM.BoFieldTypes.db_Date)
        'Me.AddColumns("@PRE_SHIPMENT", "MAWBDt", "MAWB Date", SAPbobsCOM.BoFieldTypes.db_Date)
        'Me.AddColumns("@PRE_SHIPMENT", "SBNoDt", "SB No Dt", SAPbobsCOM.BoFieldTypes.db_Date)
        'Me.AddColumns("@PRE_SHIPMENT", "FligtDtl", "Flight Detail", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        'Me.AddColumns("@PRE_SHIPMENT", "contno", "Container No", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        'Me.AddColumns("@PRE_SHIPMENT", "blattno", "Blatt No", SAPbobsCOM.BoFieldTypes.db_Alpha, 60)
        'Me.AddColumns("@PRE_SHIPMENT", "delterm", "Delivery Terms", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)

        Dim ValidValues2 = New String(,) {{"Single", "Single"}, {"Multiple", "Multiple"}}
        Dim DefaultVal2 = New String(,) {{"Single", "Single"}}
        Me.AddColumns("@PRE_SHIPMENT", "MulRote", "Multiple Route", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)


        Dim ValidValues1 = New String(,) {{"Sea", "Sea"}, {"Air", "Air"}}
        Dim DefaultVal1 = New String(,) {{"Sea", "Sea"}}
        Me.AddColumns("@PRE_SHIPMENT", "MOS", "Mode of Shipment", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)


        'ValidValues = New String(,) {{"BILLABLE", "BILLABLE"}, {"FOC", "FOC"}}
        'DefaultVal = New String(,) {{"BILLABLE", "BILLABLE"}}
        'Me.AddColumns("PRE_SHIPMENT_D0", "FOC", "Billing Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)

        Me.AddColumns("@PRE_SHIPMENT", "invno", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@PRE_SHIPMENT", "oinvno", "Original Inv No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@PRE_SHIPMENT", "VesslNo", "Vessel No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        Me.AddColumns("@PRE_SHIPMENT", "CntyFin", "Country of Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@PRE_SHIPMENT", "PorDisch", "Port Of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@PRE_SHIPMENT", "FinDest", "Place of Delivery", SAPbobsCOM.BoFieldTypes.db_Alpha, 60)
        Me.AddColumns("@PRE_SHIPMENT", "HAWB", "HAWB No", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@PRE_SHIPMENT", "MAWB", "MAWB No", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@PRE_SHIPMENT", "SBNo", "SB No", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@PRE_SHIPMENT", "FligtDtl", "Flight Detail", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@PRE_SHIPMENT", "contno", "Container No", SAPbobsCOM.BoFieldTypes.db_Alpha, 150)
        Me.AddColumns("@PRE_SHIPMENT", "delterm", "Delivery Terms", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@PRE_SHIPMENT", "vehno", "Vehicle No", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@PRE_SHIPMENT", "stat", "DPR Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@PRE_SHIPMENT", "status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@PRE_SHIPMENT", "PortDelv", "Port of delivery", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
        Me.AddColumns("@PRE_SHIPMENT", "CF", "Cost & Frieght", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
        Me.AddColumns("@PRE_SHIPMENT", "NoKndPck", "No. & Kind of Packages", SAPbobsCOM.BoFieldTypes.db_Alpha, 254)
        Me.AddColumns("@PRE_SHIPMENT", "PreCarier", "Place Of Receipt by pre-Carier", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
        Me.AddColumns("@PRE_SHIPMENT", "PortLoad", "Port Of loading", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
        Me.AddColumns("@PRE_SHIPMENT", "bldno", "Bill Of Lading No", SAPbobsCOM.BoFieldTypes.db_Alpha, 70)
        Me.AddColumns("@PRE_SHIPMENT", "LCN", "LC No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@PRE_SHIPMENT", "NCA", "Name of Clearing Agent", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@PRE_SHIPMENT", "FMVessel", "F/M Vessel", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        Me.AddColumns("@PRE_SHIPMENT", "STA_GOOD", "State of Origin of Goods", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        Me.AddColumns("@PRE_SHIPMENT", "SUPP_PLC", "Place of supp Manuf", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@PRE_SHIPMENT", "TYP_COM", "Type Of Company", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        Me.AddColumns("@PRE_SHIPMENT", "trgdocno", "Target Doc No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@PRE_SHIPMENT", "SBdate", "Shipping bill Date", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@PRE_SHIPMENT", "BLdate", "BL Date", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@PRE_SHIPMENT", "GOODSDESC", "Goods Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@PRE_SHIPMENT", "STA_GOOD1", "State of Origin of Goods", SAPbobsCOM.BoFieldTypes.db_Alpha, 8)
        Me.AddColumns("@PRE_SHIPMENT", "SUPP_PLC1", "Place of Supp Manufacturer", SAPbobsCOM.BoFieldTypes.db_Alpha, 8)
        Me.AddColumns("@PRE_SHIPMENT", "MANF1", "Name of Manufacturer", SAPbobsCOM.BoFieldTypes.db_Alpha, 8)
        Me.AddColumns("@PRE_SHIPMENT", "Buyer", "Buyer Address", SAPbobsCOM.BoFieldTypes.db_Alpha, 254)
        Me.AddColumns("@PRE_SHIPMENT", "PayTerm", "Payment Term", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@PRE_SHIPMENT", "Linername", "Liner Details", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        'Me.AddColumns("@PRE_SHIPMENT", "MOS", "Mode of Shipment", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
        'Me.AddColumns("@PRE_SHIPMENT", "MulRote", "Multiple Route", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)

        Me.AddColumns("@PRE_SHIPMENT", "DEPB_AMT", "Amount of DEPB as per S/B", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
        Me.AddColumns("@PRE_SHIPMENT", "FOB_SB", "FOB in INR as per S/B", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
        Me.AddColumns("@PRE_SHIPMENT", "FOB_AMT", "FOB Value as per Invoice(INR)", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
        Me.AddColumns("@PRE_SHIPMENT", "NO_OF_CN", "No Of Container", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)


        Me.AddColumns("@PRE_SHIPMENT", "Blddate", "Bill Of Lading Date", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@PRE_SHIPMENT", "LCD", "LC Date", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@PRE_SHIPMENT", "LDS", "Last Date of Shipment", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@PRE_SHIPMENT", "LDN", "Last Date of Negotiation", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@PRE_SHIPMENT", "LET_DT", "LET Export Date", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@PRE_SHIPMENT", "HAWBDt", "HAWB Date", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@PRE_SHIPMENT", "MAWBDt", "MAWB Date", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@PRE_SHIPMENT", "SBNoDt", "SB No Dt", SAPbobsCOM.BoFieldTypes.db_Date)

        Me.AddColumns("@PRE_SHIPMENT", "GrssWt", "Gross Weight", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)
        Me.AddColumns("@PRE_SHIPMENT", "NetWt", "Nett Weight", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)
        Me.AddColumns("@PRE_SHIPMENT", "CBM", "CBM", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)
        Me.AddColumns("@PRE_SHIPMENT", "TotCart", "Total Cartons", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)
        Me.AddColumns("@PRE_SHIPMENT", "CUSTRATE", "Exchange Rate as (B/E)/SB", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)

        Me.AddColumns("@PRE_SHIPMENT", "Notify", "Notify", SAPbobsCOM.BoFieldTypes.db_Memo)

        If Not Me.UDOExists("PRE_SHIPMENT") Then
            Dim findAliasNDescription = New String(,) {{"DocEntry", "DocEntry"}}
            Me.registerUDO("PRE_SHIPMENT", "Pre-Shipment", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "PRE_SHIPMENT", "PRE_SHIPMENT_D0", "")
            findAliasNDescription = Nothing
        End If
    End Sub
    Sub SAP_TABLES_UDF()
        Try
            Me.AddColumns("RDR1", "asrtcodetemp", "Assorted Code Temp", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("UFD1", "season1", "season1", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("RDR1", "asrtcode", "Assorted Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("POR1", "bomno", "BOM No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("POR1", "bomlnid", "BOM Line No", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            Me.AddColumns("ORDR", "doccur", "Buyer Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            Me.AddColumns("ORDR", "docrate", "Document Rate", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)
            Me.AddColumns("RDR1", "pricefc", "Price FC", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("RDR1", "totalfc", "LineTotal FC", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("OITM", "color", "Colour", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OITM", "colornm", "Colour Desc", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OITM", "cust", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OACT", "ccentre", "Cost Centre Not Required", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)
            'Me.AddColumns("@GEN_MREQ_D0", "minchk", "MINCheckBox", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)
            Me.AddColumns("OALC", "rate", "Rate", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)

            Dim ValidValues1 = New String(,) {{"Rate", "Rate"}, {"Percent", "Percent"}}
            Dim DefaultVal1 = New String(,) {{"Rate", "Rate"}}
            'Me.AddColumns("OPDN", "insstat", "Inspection Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)
            Me.AddColumns("OALC", "type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues1, 1, DefaultVal1)

            Me.AddColumns("OITM", "Acct1", "Account-Unit1", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OITM", "Acct2", "Account-Unit2", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OITM", "Acct3", "Account-Unit3", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OITM", "Acct4", "Account-LGUnit1", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OIGN", "unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

            Me.AddColumns("OITM", "custnm", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OITM", "type", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            'Rajkumar 18.08.14
            Me.AddColumns("OITM", "tariff", "Tariff Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OITM", "blend", "Blend Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OITM", "cap", "CAP Value", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)

            Me.AddColumns("OITM", "typenm", "Item Type Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OITM", "size", "Size", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OITM", "sizenm", "Size Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OITM", "subtype", "Sub Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OITM", "subtpnm", "Sub Type Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OITM", "style", "Style", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OITM", "stylenm", "Style Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OITM", "qlty", "Quality", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OITM", "qltynm", "Quality Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OITM", "mtrgrp", "Master Group", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OITM", "MonoCart", "MonoCart", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OITM", "MT_NAME", "Master Catogery", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OITM", "descr", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OINV", "pino", "Proforma Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OPCH", "grpodoc", "Sub Contract GRPO No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            Me.AddColumns("OPCH", "grpono", "GRPO No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            Me.AddColumns("OINV", "pidate", "Proforma Invoice Date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("OINV", "styleref", "Style Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OINV", "buyer", "Buyer", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OINV", "frgtrem", "Freight Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OINV", "tinv", "Type Of Invoice", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OINV", "lrno", "LR No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OINV", "srlno", "SRL No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OINV", "pono", "PO NO", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OINV", "ubginvno", "UBG Inv No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OINV", "iecno", "IEC No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OINV", "quotdate", "Quotation Date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("OINV", "ourrefso", "Our Ref SO", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OPOR", "unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OCRD", "unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OPOR", "season", "Season", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OPOR", "opono", "Original PO NO", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("PDN1", "accqty", "Accepted Qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("PDN1", "rejqty", "Rejected Qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("PDN1", "shqty", "Shortage Qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("PDN1", "exqty", "Excess Qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("PDN1", "qty", "Qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("PDN1", "tol", "Tolerance", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("WTR1", "BAL_QTY", "Balance Qty.", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("OWTR", "isstyp", "Issue Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OWTR", "Type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("OPOR", "Types", "Purchase Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("OWTR", "twhs", "Destination WareHouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("OCRD", "WhsCode", "Warehouse Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OWTR", "DocNum", "DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("WTR1", "scpono", "SC PO No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("WTR1", "scpoln", "SC PO Ln", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            Me.AddColumns("PCH1", "SCGRNNo", "SubContract GRNNo.", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("PCH1", "SCGRNLine", "SubContract GRNLine", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("PCH1", "SCGRNQty", "SubContract GRN Qty.", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("OIGE", "sono", "Sales Order no", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OIGE", "ptnno", "PTN No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("IGE1", "totavlbl", "Total Availability", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("ITT1", "tol", "Tolerance", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage)
            Me.AddColumns("WTR1", "scpono", "SC PO No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("WTR1", "scpoln", "SC PO Ln", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            Me.AddColumns("OCRD", "SubAcct", "Subcontract G/L", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("PCH1", "price", "Item Price", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("PCH1", "itemcode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("PCH1", "grpode", "GRPO DocEntry.", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            'Me.AddColumns("PCH1", "SGRNNo", "SubContract GRNNo.", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            'Me.AddColumns("PCH1", "SGRNLine", "SubContract GRNLine", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            'Me.AddColumns("PCH1", "SGRNQty", "SubContract GRN Qty.", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("OVPM", "bcjv", "Bank Charges JV", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("ORCT", "frgn", "Foreign Rate", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("ORCT", "docrate", "DocumentRate", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("ORCT", "balamt", "Balance Amount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)
            Me.AddColumns("ORCT", "contrno", "Contract No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("ORCT", "unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("ORCT", "bcjv", "Bank Charges JV", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("RCT4", "invno", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("RCT4", "expcode", "Freight Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("JDT1", "invno", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("JDT1", "expcode", "Freight Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

            Me.AddColumns("ORPC", "invno", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OWTR", "type", "Doc Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OWTR", "grnno", "GRN No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OWTR", "mrnno", "MRN No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("WTR1", "grnno", "GRN No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("WTR1", "grnlnid", "GRN Line ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            Me.AddColumns("OWTR", "sono", "SO NO", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OWTR", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OWTR", "sfgcode", "SFG Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OWTR", "type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OWTR", "mrnno", "MRN No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OWTR", "isstyp", "Issue Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("WTR1", "mrnlid", "Material Line ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("WTR1", "rqstqty", "Material Req Requested qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("WTR1", "issued", "Material Req Issued qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)

            Me.AddColumns("OWTR", "grnno", "GRN No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

            Me.AddColumns("WTR1", "grnno", "GRN No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("WTR1", "grnlnid", "GRN Line ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            Dim ValidValues = New String(,) {{"Open", "Open"}, {"Closed", "Closed"}}
            Dim DefaultVal = New String(,) {{"Open", "Open"}}
            Me.AddColumns("OPDN", "insstat", "Inspection Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)
            ValidValues = New String(,) {{"Spot", "Spot"}, {"FWD Cover", "FWD Cover"}}
            DefaultVal = New String(,) {{"Spot", "Spot"}}
            Me.AddColumns("ORCT", "Frgn", "Rate Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)
            Me.AddColumns("OPDN", "reval", "Inventory Revaluation", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None)
            Me.AddColumns("OPDN", "revalno", "Inventory Revaluation Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OPDN", "macid", "MAC_ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            'Me.AddColumns("OMRV", "macid", "MAC_ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)

            Me.AddColumns("OPDN", "lcadd", "LC Added", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

            Me.AddColumns("OPDN", "cardcode", "Vendor Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OPDN", "lcno", "LC RefNo.", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("PDN1", "insstat", "Inspection Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)
            Me.AddColumns("PDN1", "openqty", "Open Qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)

            Me.AddColumns("OIGE", "ptnno", "PTN No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OIGE", "sono", "Sales Order No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("IGE1", "totavlbl", "Total Available", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)

            Me.AddColumns("PCH1", "invno", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("PCH1", "expcode", "Freight Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("PCH1", "PROVIDER", "Provider Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("PCH1", "SERTAXNO", "Service Tax Reg No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("PCH1", "INV_NO", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            Me.AddColumns("PCH1", "INV_DATE", "Invoice Date", SAPbobsCOM.BoFieldTypes.db_Date, 30)
            Me.AddColumns("PCH1", "SER_DESC", "Service Desc", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("PCH1", "SERACTNO", "Service Act No", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)

            ValidValues = New String(,) {{"Yes", "Yes"}, {"No", "No"}}
            DefaultVal = New String(,) {{"No", "No"}}
            Me.AddColumns("OWHS", "inspwhs", "Inspection Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 2, DefaultVal)
            Me.AddColumns("OWHS", "intwhs", "Intransit Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 2, DefaultVal)
            Me.AddColumns("OWHS", "locgl", "Local GL Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OWHS", "deemgl", "Deemed GL Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OWHS", "dirgl", "Direct GL Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OUSR", "approve", "Approval", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            Me.AddColumns("OUSR", "cstsht", "Cost Sheet", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            Me.AddColumns("OWHS", "shwhs", "Shortage Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 2, DefaultVal)
            Me.AddColumns("OWHS", "exwhs", "Excess Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 2, DefaultVal)
            Me.AddColumns("OWOR", "Created", "Created by Wiz.", SAPbobsCOM.BoFieldTypes.db_Alpha, 4, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 2, DefaultVal)
            Me.AddColumns("OWOR", "sorefno", "Sales Order Ref No", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)

            Me.AddColumns("OWTR", "subconno", "Sub Contract No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OWTR", "subretno", "Sub Contractor Return", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("WTR1", "subconln", "Sub Contract LID", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            Me.AddColumns("WTR1", "subretln", "Sub Contract Ret LID", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            Me.AddColumns("OWOR", "process", "Process", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OWOR", "unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OCRD", "subcon", "Sub Contractor", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)
            Me.AddColumns("OWTR", "subconno", "Sub Contract No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OWTR", "subretno", "Sub Contractor Return", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("WTR1", "subconln", "Sub Contract LID", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            Me.AddColumns("WTR1", "subretln", "Sub Contract Ret LID", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            ValidValues = New String(,) {{"Accepted", "Accepted"}, {"Rejected", "Rejected"}, {"Shortage", "Shortage"}, {"Excess", "Excess"}}
            DefaultVal = New String(,) {{"Accepted", "Accepted"}}
            Me.AddColumns("WTR1", "grpostat", "GRPO Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, )

            Me.AddColumns("OITM", "tol", "Tolerance", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)

            ValidValues = New String(,) {{"YES", "YES"}, {"NO", "NO"}}
            DefaultVal = New String(,) {{"NO", "NO"}}
            Me.AddColumns("OPOR", "approval", "Approval", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)

            'Vijeesh
            Me.AddColumns("OWHS", "UNIT", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)

            'Rajkumar

            'Dim ValidValues2 = New String(,) {{"Single", "Single"}, {"Multiple", "Multiple"}}
            'Dim DefaultVal2 = New String(,) {{"Single", "Single"}}
            'Me.AddColumns("OINV", "MulRote", "Multiple Route", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)


            'ValidValues1 = New String(,) {{"Sea", "Sea"}, {"Air", "Air"}}
            'DefaultVal1 = New String(,) {{"Sea", "Sea"}}
            'Me.AddColumns("OINV", "MOS", "Mode of Shipment", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)

            'Me.AddColumns("QUT1", "FOC", "Billing Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            'ValidValues = New String(,) {{"BILLABLE", "BILLABLE"}, {"FOC", "FOC"}}
            'DefaultVal = New String(,) {{"BILLABLE", "BILLABLE"}}
            Me.AddColumns("OINV", "invno", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OINV", "oinvno", "Original Inv No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OINV", "VesslNo", "Vessel No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            Me.AddColumns("OINV", "CntyFin", "Country of Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OINV", "PorDisch", "Port Of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OINV", "FinDest", "Place of Delivery", SAPbobsCOM.BoFieldTypes.db_Alpha, 60)
            Me.AddColumns("OINV", "HAWB", "HAWB No", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OINV", "MAWB", "MAWB No", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OINV", "SBNo", "SB No", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OINV", "FligtDtl", "Flight Detail", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OINV", "contno", "Container No", SAPbobsCOM.BoFieldTypes.db_Alpha, 150)
            Me.AddColumns("OINV", "delterm", "Delivery Terms", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OINV", "vehno", "Vehicle No", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OINV", "stat", "DPR Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OINV", "status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OINV", "PortDelv", "Port of delivery", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            Me.AddColumns("OINV", "CF", "Cost & Frieght", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            Me.AddColumns("OINV", "NoKndPck", "No. & Kind of Packages", SAPbobsCOM.BoFieldTypes.db_Alpha, 254)
            Me.AddColumns("OINV", "PreCarier", "Place Of Receipt by pre-Carier", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            Me.AddColumns("OINV", "PortLoad", "Port Of loading", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            Me.AddColumns("OINV", "bldno", "Bill Of Lading No", SAPbobsCOM.BoFieldTypes.db_Alpha, 70)
            Me.AddColumns("OINV", "LCN", "LC No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OINV", "NCA", "Name of Clearing Agent", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OINV", "FMVessel", "F/M Vessel", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            Me.AddColumns("OINV", "STA_GOOD", "State of Origin of Goods", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            Me.AddColumns("OINV", "SUPP_PLC", "Place of supp Manuf", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OINV", "TYP_COM", "Type Of Company", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            Me.AddColumns("OINV", "trgdocno", "Target Doc No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("OINV", "SBdate", "Shipping bill Date", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OINV", "BLdate", "BL Date", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OINV", "GOODSDESC", "Goods Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OINV", "STA_GOOD1", "State of Origin of Goods", SAPbobsCOM.BoFieldTypes.db_Alpha, 8)
            Me.AddColumns("OINV", "SUPP_PLC1", "Place of Supp Manufacturer", SAPbobsCOM.BoFieldTypes.db_Alpha, 8)
            Me.AddColumns("OINV", "MANF1", "Name of Manufacturer", SAPbobsCOM.BoFieldTypes.db_Alpha, 8)
            Me.AddColumns("OINV", "Buyer", "Buyer Address", SAPbobsCOM.BoFieldTypes.db_Alpha, 254)
            Me.AddColumns("OINV", "PayTerm", "Payment Term", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("OINV", "Linername", "Liner Details", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            'Me.AddColumns("OINV", "MOS", "Mode of Shipment", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            'Me.AddColumns("OINV", "MulRote", "Route", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            Me.AddColumns("INV1", "MonoCart", "MonoCart", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

            Me.AddColumns("OINV", "DEPB_AMT", "Amount of DEPB as per S/B", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("OINV", "FOB_SB", "FOB in INR as per S/B", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("OINV", "FOB_AMT", "FOB Value as per Invoice(INR)", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("OINV", "NO_OF_CN", "No Of Container", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)


            Me.AddColumns("OINV", "Blddate", "Bill Of Lading Date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("OINV", "LCD", "LC Date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("OINV", "LDS", "Last Date of Shipment", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("OINV", "LDN", "Last Date of Negotiation", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("OINV", "LET_DT", "LET Export Date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("OINV", "HAWBDt", "HAWB Date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("OINV", "MAWBDt", "MAWB Date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("OINV", "SBNoDt", "SB No Dt", SAPbobsCOM.BoFieldTypes.db_Date)

            Me.AddColumns("OINV", "Notify", "Notify", SAPbobsCOM.BoFieldTypes.db_Memo)

            Me.AddColumns("INV1", "sprice", "Supplementary Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("INV1", "dbkval", "DBK Value", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("INV1", "anspval", "ANSP Value", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("INV1", "comm", "Commission", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("INV1", "dprno", "DPR No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("INV1", "dprlno", "DPR Line Id", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            Me.AddColumns("INV1", "Source", "Source", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            Me.AddColumns("INV1", "preno", "Pre-Shipment No", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)


            ValidValues = New String(,) {{"Open", "Open"}, {"Closed", "Closed"}}
            DefaultVal = New String(,) {{"Open", "Open"}}
            Me.AddColumns("ORDR", "status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)
            Me.AddColumns("ORDR", "prests", "Pre Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)
            Me.AddColumns("ORDR", "preship", "Pre Shipment No", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)


            Me.AddColumns("RDR1", "prests", "Pre Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)
            Me.AddColumns("RDR1", "preqty", "Pre Shipment Qty", SAPbobsCOM.BoFieldTypes.db_Float, 20, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("RDR1", "poqty", "Open Qty", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("RDR1", "status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)
            Me.AddColumns("RDR1", "sostat", "SO Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)
            Me.AddColumns("RDR1", "invqty", "Inv Open Qty", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("ORDR", "NO_OF_CN", "No Of Container", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Sum)

            Me.AddColumns("ORDR", "CntyFin", "Country of Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            Me.AddColumns("ORDR", "PreCarrier", "Place of Pre Carrier", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            Me.AddColumns("ORDR", "CF", "Cost and Freight", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            Me.AddColumns("ORDR", "NoKndPck", "No. & Kind of Package", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)


            ValidValues = New String(,) {{"BILLABLE", "BILLABLE"}, {"FOC", "FOC"}}
            DefaultVal = New String(,) {{"BILLABLE", "BILLABLE"}}
            Me.AddColumns("RDR1", "FOC", "Billing Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)
            'Me.AddColumns("OUSR", "costap", "Cost Approval", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            'Vijeesh
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub
    Sub No_Objects_Tables()
        Me.AddTable("COSTFREIGHT", "UBG->Cost and Freight", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        Me.AddTable("MNF_MASTER", "Manufacturer Master", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        Me.AddTable("MNF_PLACE", "Manufacturer Place", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        Me.AddTable("PORT_LOADING", "Port of Loading", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        Me.AddTable("SOURCE", "Source", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        Me.AddTable("STATEORIGIN", "State of Origin Goods", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        Me.AddTable("PRE_CARRIER", "Place of Pre Carrier", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        Me.AddTable("GEN_SER_PROVIDER", "Service Provider", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        Me.AddColumns("@GEN_SER_PROVIDER", "PROVIDER", "Provider Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddTable("GEN_SERVICE", "Service Master", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        Me.AddColumns("@GEN_SERVICE", "SER_DESC", "Service Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_SERVICE", "SERACTNO", "Service Act No", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
    End Sub
    Sub Sales_Matrix_Order()
        Me.AddTable("UBG_SALES_ORDR", "UBG->Sales Order", SAPbobsCOM.BoUTBTableType.bott_NoObject)

        Me.AddColumns("@UBG_SALES_ORDR", "sono", "Sales Order No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@UBG_SALES_ORDR", "Date", "Date", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@UBG_SALES_ORDR", "macid", "MAC ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        Me.AddColumns("@UBG_SALES_ORDR", "CusName", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
    End Sub
    Sub DocumentSeriesAuthorisation()
        Try
            Me.AddTable("GEN_USER_AUTH", "DocumentSeriesAuthorisation", SAPbobsCOM.BoUTBTableType.bott_NoObject)

            Me.AddColumns("@GEN_USER_AUTH", "user", "UserCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("@GEN_USER_AUTH", "unit", "UNIT", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_CST_PRCS()
        Try
            Me.AddTable("GEN_CST_PRCS", "Gen->Cost Sheet Prcs", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            Dim ValidValues = New String(,) {{"Yes", "Yes"}, {"No", "No"}}
            Me.AddColumns("@GEN_CST_PRCS", "qtyinkg", "Quantity in KG", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)


            If Not Me.UDOExists("GEN_CST_PRCS") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}}
                Me.registerUDO("GEN_CST_PRCS", "GEN_CST_PRCS", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_CST_PRCS", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SubContracting()
        Try
            Me.AddTable("GEN_SUB_CONTRACT", "Gen ->SubContract", SAPbobsCOM.BoUTBTableType.bott_Document)

            Me.AddColumns("@GEN_SUB_CONTRACT", "cstbom", "Custom BOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)
            Me.AddColumns("@GEN_SUB_CONTRACT", "CardCode", "Vendor Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SUB_CONTRACT", "quotdate", "Quotation Date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_SUB_CONTRACT", "ourrefso", "Our Ref So", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_SUB_CONTRACT", "note", "Note", SAPbobsCOM.BoFieldTypes.db_Alpha, 250)
            Me.AddColumns("@GEN_SUB_CONTRACT", "manual", "Manual", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)
            Me.AddColumns("@GEN_SUB_CONTRACT", "CardName", "Vendor Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SUB_CONTRACT", "ContPer", "Contact Person", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SUB_CONTRACT", "VendRef", "Vendor Ref. No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SUB_CONTRACT", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            Me.AddColumns("@GEN_SUB_CONTRACT", "PostDate", "PostDate", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_SUB_CONTRACT", "DelDate", "DelDate", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_SUB_CONTRACT", "DocDate", "DocDate", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_SUB_CONTRACT", "Owner", "Owner", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SUB_CONTRACT", "OwnerCod", "OwnerCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SUB_CONTRACT", "VendWhs", "VendWhs", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SUB_CONTRACT", "Buyer", "Buyer", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SUB_CONTRACT", "PayTrms", "PayTrms", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SUB_CONTRACT", "PayCode", "PayCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SUB_CONTRACT", "JourRem", "JourRem", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SUB_CONTRACT", "TotBefTa", "TotBefTax", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SUB_CONTRACT", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SUB_CONTRACT", "Tax", "Tax", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SUB_CONTRACT", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)
            'Vijeesh
            Me.AddColumns("@GEN_SUB_CONTRACT", "manwobom", "Manual WithOut BOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)
            Me.AddColumns("@GEN_SUB_CONTRACT", "approve", "Approve", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)


            Me.AddTable("GEN_SUB_CONTRACT_D0", "Gen ->SubContract Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            Me.AddColumns("@GEN_SUB_CONTRACT_D0", "ItemCode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SUB_CONTRACT_D0", "ItemDesc", "ItemDesc", SAPbobsCOM.BoFieldTypes.db_Alpha, 250)
            Me.AddColumns("@GEN_SUB_CONTRACT_D0", "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SUB_CONTRACT_D0", "Price", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SUB_CONTRACT_D0", "TotalLC", "TotalLC", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SUB_CONTRACT_D0", "TaxRate", "TaxRate", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SUB_CONTRACT_D0", "TaxAmt", "TaxAmt", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SUB_CONTRACT_D0", "DCQty", "DCQty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SUB_CONTRACT_D0", "GRNQty", "GRNQty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SUB_CONTRACT_D0", "RetQty", "RetQty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SUB_CONTRACT_D0", "TaxCode", "TaxCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SUB_CONTRACT_D0", "Whs", "Whs", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SUB_CONTRACT_D0", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_SUB_CONTRACT_D0", "UOM", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_SUB_CONTRACT_D0", "SONo", "Sales Order No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SUB_CONTRACT_D0", "SODNo", "Sales Order Doc.No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)

            Me.AddTable("GEN_SUB_CONTRACT_D1", "Gen ->SubContract RMLines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            Me.AddColumns("@GEN_SUB_CONTRACT_D1", "LineID", "LineID", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SUB_CONTRACT_D1", "Father", "Father", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SUB_CONTRACT_D1", "Code", "Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SUB_CONTRACT_D1", "POQty", "PO Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SUB_CONTRACT_D1", "DCQty", "DC Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SUB_CONTRACT_D1", "RetQty", "Return Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            'Vijeesh
            Me.AddColumns("@GEN_SUB_CONTRACT_D1", "BOMQty", "BOM Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SUB_CONTRACT_D1", "FWhs", "From Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)


            If Not Me.UDOExists("GEN_SUB_CONTRACT") Then
                Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
                Me.registerUDO("GEN_SUB_CONTRACT", "GEN: Sub_Contract", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "GEN_SUB_CONTRACT", "GEN_SUB_CONTRACT_D0", "GEN_SUB_CONTRACT_D1", "")
                findAliasNDescription = Nothing
            End If

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SubContractDeliveryChallan()
        Try
            Me.AddTable("GEN_SC_DC", "Gen -> SubContractDelivery", SAPbobsCOM.BoUTBTableType.bott_Document)

            Me.AddColumns("@GEN_SC_DC", "CardCode", "Vendor Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_DC", "CardName", "Vendor Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_DC", "SCDocNo", "Sub Contract Doc.No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) ' ''''
            Me.AddColumns("@GEN_SC_DC", "SCNo", "Sub Contract No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) ' ''''
            Me.AddColumns("@GEN_SC_DC", "SCDat", "Sub Contract Date", SAPbobsCOM.BoFieldTypes.db_Date) '
            Me.AddColumns("@GEN_SC_DC", "RefNo", "Vendor Ref. No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_DC", "ContPer", "Contact Person", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_DC", "DCDat", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date) '
            Me.AddColumns("@GEN_SC_DC", "DelDat", "Delivery Date", SAPbobsCOM.BoFieldTypes.db_Date) '
            Me.AddColumns("@GEN_SC_DC", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 15) '
            Me.AddColumns("@GEN_SC_DC", "DCNo", "Doc Num", SAPbobsCOM.BoFieldTypes.db_Numeric) ' ''''
            Me.AddColumns("@GEN_SC_DC", "Buyer", "Buyer", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_DC", "Owner", "Owner", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_DC", "OwnerCod", "OwnerCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_DC", "Rmrk", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo) '
            Me.AddColumns("@GEN_SC_DC", "ItemNo", "BOM Item", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_DC", "Qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity) '
            Me.AddColumns("@GEN_SC_DC", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price) '
            Me.AddColumns("@GEN_SC_DC", "InvTrNo", "InvTrNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 20) '
            Me.AddColumns("@GEN_SC_DC", "InvTrDNo", "InvTr Doc.No", SAPbobsCOM.BoFieldTypes.db_Alpha, 20) '
            Me.AddColumns("@GEN_SC_DC", "SONo", "SalesOrderNo.", SAPbobsCOM.BoFieldTypes.db_Alpha, 20) '
            Me.AddColumns("@GEN_SC_DC", "SODNo", "Sales OrderDNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 20) '


            Me.AddTable("GEN_SC_DC_D0", "Gen ->SubCont. Delivery Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            Me.AddColumns("@GEN_SC_DC_D0", "IsCheck", "IsCheck", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_DC_D0", "ItemNo", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_DC_D0", "Desc", "Item Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 250) '
            Me.AddColumns("@GEN_SC_DC_D0", "FWhs", "From Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_DC_D0", "TWhs", "To Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_DC_D0", "Stock", "Stock Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity) '
            Me.AddColumns("@GEN_SC_DC_D0", "Unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_DC_D0", "Qty", "Actual Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity) '
            Me.AddColumns("@GEN_SC_DC_D0", "IssQty", "Issued Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity) '
            Me.AddColumns("@GEN_SC_DC_D0", "CompQty", "Completed Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SC_DC_D0", "Rmrk", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 100) '
            Me.AddColumns("@GEN_SC_DC_D0", "Price", "Price", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price) '
            Me.AddColumns("@GEN_SC_DC_D0", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price) '
            Me.AddColumns("@GEN_SC_DC_D0", "BOMQty", "BOMQty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity) '

            If Not Me.UDOExists("GEN_SC_DC") Then
                Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
                Me.registerUDO("GEN_SC_DC", "GEN: Sub_Contract_DC", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "GEN_SC_DC", "GEN_SC_DC_D0", "", "")
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    'Prasad 16-7-13
    Sub GEN_SAM_REV()
        Try
            Me.AddTable("GEN_SAM_REV", "Gen->SAM REVALUATION", SAPbobsCOM.BoUTBTableType.bott_Document)
            Me.AddColumns("@GEN_SAM_REV", "period", "period", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SAM_REV", "Month", "Month", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SAM_REV", "Unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SAM_REV", "DocDate", "DocDate", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_SAM_REV", "Year", "Year", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("@GEN_SAM_REV", "Stat", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
         

            Me.AddTable("GEN_SAM_REV_D0", "Gen->SAM REVALUATION Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            Me.AddColumns("@GEN_SAM_REV_D0", "ItmCode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("@GEN_SAM_REV_D0", "ItmName", "ItemName", SAPbobsCOM.BoFieldTypes.db_Alpha, 150)
            'Me.AddColumns("@GEN_SAM_REV_D0", "CngSAM", "CuttingSAM", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            'Me.AddColumns("@GEN_SAM_REV_D0", "SthSAM", "SitchSAM", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            'Me.AddColumns("@GEN_SAM_REV_D0", "FinSAM", "FinishSAM", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_REV_D0", "Selprc", "Selling price", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            ' Me.AddColumns("@GEN_SAM_REV_D0", "CapCtg", "CapCutting", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            ' Me.AddColumns("@GEN_SAM_REV_D0", "Capstg", "Capstiching", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            'Me.AddColumns("@GEN_SAM_REV_D0", "CapFng", "CapFinishing", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "CTQty", "CTQty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_REV_D0", "OpnCT", "OpenCTCost", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "CutCst", "Cutting Material Cost", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "CutPCst", "Cutting Process cost", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "CTFcst", "CT Cost final", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "CTRev", "CT Rev Cost", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "STQty", "ST Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "OpnSTQty", "Opening ST cost", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_REV_D0", "STMCst", "Stiching Material cost", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "STPcst", "Stiching Process cost", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "STPFcst", "ST Cost Final", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "STRCst", "ST Rev Cost", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "WSTQty", "WST Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_REV_D0", "MCPWST", "MCPerqty-WST", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_REV_D0", "WSTPQty", "WST Pcper qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_REV_D0", "WSTRcst", "WST Rev Cost", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "WSTVal", "WST Value", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "FIQty", "FI Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_REV_D0", "OpnFIcst", "Opening FI cost", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "FinMCst", "Finishing mat cost", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "FinPCst", "Finish Process cost", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "FinFCst", "FI Cost Final", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "FinRCst", "FI REV COST", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "WFIqty", "WFI QTY", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "WFIPval", "WFI Perqty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_REV_D0", "WFIVal", "WFI Value", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "FGQTY", "FG QTY", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_REV_D0", "FGPQTY", "FG Perqty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_REV_D0", "FGVal", "FG Value", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_REV_D0", "CTRefNo", "CTRefNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("@GEN_SAM_REV_D0", "STRefNo", "STRefNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("@GEN_SAM_REV_D0", "WSTRefNo", "WSTRefNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("@GEN_SAM_REV_D0", "FIRefNo", "FIRefNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("@GEN_SAM_REV_D0", "WFIRefNo", "WFIRefNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("@GEN_SAM_REV_D0", "FGRefNo", "FGRefNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("@GEN_SAM_REV_D0", "select", "select", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Dim ValidValues = New String(,) {{"Y", "Yes"}, {"N", "No"}}
            Me.AddColumns("@GEN_SAM_REV_D0", "select", "select", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues)
            Me.AddColumns("@GEN_SAM_REV_D0", "StSlct", "Stiching select", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues)
            Me.AddColumns("@GEN_SAM_REV_D0", "WstSlct", "Wst select", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues)
            Me.AddColumns("@GEN_SAM_REV_D0", "CTslct", "CTslecting", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues)
            Me.AddColumns("@GEN_SAM_REV_D0", "FIslct", "F'I selecting", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues)
            Me.AddColumns("@GEN_SAM_REV_D0", "WFIslct", "WFIselect", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues)

            Me.AddColumns("@GEN_SAM_REV_D0", "MC1val", "MC1 value", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "PC1val", "PC1 VALUE", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "CCCst", "CUR CUTNG COST", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "CCFnl", "CUT COST FINAL", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "CCSprc", "Cut Cost SalPrc", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "CMPCST", "CUR MC+PC FOR ST", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "CSTQty", "CURRENT STQTY", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_REV_D0", "MC2Val", "MC2 VALUE", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "PC2Val", "PC2 VALUE", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "SSTQty", "SubcontractSTQty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_REV_D0", "STSChgs", "STSubcontractCharges", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "STSPQty", "STSubcontractPerQty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "CRTSTCst", "CURRENTSTCOST", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "CRTSTVal", "CURRENTSTVALUE", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "STCSP", "StitchCostSalesPrice", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "WSTSTrm", "WSTStitchTrims", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "CMPFFI", "CUR MC PC FOR FI", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "FINTQty", "FINISH TRIMSPERQTY", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "SCFIQty", "SubcontractFIQty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_REV_D0", "FISCC", "FISubcontrchrges", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "FISPQty", "FISubcontrPerQty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "CurFICst", "CURRENTFICOST", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "CurFIVal", "CURRENTFIVALUE", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "FiCstSP", "FinCostSalesPrice", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "WFinTrm", "WFIFinsihTrims", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "SubFGqty", "SubcontrctFGQty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "FGSChrgs", "FGSubcontrctchrgs", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_REV_D0", "FGSpQty", "FGSubcntrctPerQty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)

            Me.AddColumns("@GEN_SAM_REV_D0", "SubCut", "SUB CUTTING", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_REV_D0", "SubSit", "Sub Stiching", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_REV_D0", "SubFin", "Sub Finishing", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_REV_D0", "SubFG", "Sub FG", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_REV_D0", "SCutRefNo", "SubCutRefNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("@GEN_SAM_REV_D0", "SSitRefno", "SubSitRefno", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("@GEN_SAM_REV_D0", "SFinRefno", "SubFinRefno", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("@GEN_SAM_REV_D0", "SFGRefno", "SubFGRefno", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)

            'Subcontract FG-Qty




            If Not Me.UDOExists("GEN_SAM_REV") Then
                Dim findAliasNDescription = New String(,) {{"DocEntry", "DocEntry"}}
                Me.registerUDO("GEN_SAM_REV", "GEN_SAM_REV", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "GEN_SAM_REV", "GEN_SAM_REV_D0", "", "")
                findAliasNDescription = Nothing
            End If

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
        'Prasad 16-7-13
    End Sub
    Sub GEN_SAM()
        Try
            Me.AddTable("GEN_SAM", "Gen->SAM", SAPbobsCOM.BoUTBTableType.bott_Document)
            Me.AddTable("GEN_SAM_D0", "Gen->SAM Master Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            Me.AddColumns("@GEN_SAM", "docnum", "Docnum", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_SAM", "period", "Posting Period", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_SAM", "month", "Month", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_SAM", "unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_SAM", "date", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_SAM", "year", "Year", SAPbobsCOM.BoFieldTypes.db_Numeric)
            Me.AddColumns("@GEN_SAM", "one", "one", SAPbobsCOM.BoFieldTypes.db_Numeric)
            Me.AddColumns("@GEN_SAM", "cutper", "Cut Percentage", SAPbobsCOM.BoFieldTypes.db_Numeric)
            Me.AddColumns("@GEN_SAM", "stitper", "Stitching Percentage", SAPbobsCOM.BoFieldTypes.db_Numeric)

            Me.AddColumns("@GEN_SAM_D0", "sono", "SaleOrderNo.", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_SAM_D0", "itemcode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_SAM_D0", "itemname", "ItemName", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_SAM_D0", "csam", "CuttingSAM", SAPbobsCOM.BoFieldTypes.db_Float, 1, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_D0", "ssam", "StitchingSAM", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_D0", "fsam", "FinsihingSAM", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_D0", "wstsam", "WST SAM", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_D0", "wstqty", "WST Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_D0", "wfiqty", "WFI Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SAM_D0", "sprice", "Selling Price", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_D0", "ccapper", "Cut Percentage", SAPbobsCOM.BoFieldTypes.db_Numeric)
            Me.AddColumns("@GEN_SAM_D0", "scapper", "Stitching Percentage", SAPbobsCOM.BoFieldTypes.db_Numeric)
            Me.AddColumns("@GEN_SAM_D0", "ccap", "CAP Cutting", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_D0", "scap", "CAP Stitching", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SAM_D0", "fcap", "CAP Finishing", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            If Not Me.UDOExists("GEN_SAM") Then
                Dim findAliasNDescription = New String(,) {{"DocEntry", "DocEntry"}}
                Me.registerUDO("GEN_SAM", "GEN_SAM", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "GEN_SAM", "GEN_SAM_D0", "", "")
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SubContracting_GRPO()
        Try
            Me.AddTable("GEN_SC_GRPO", "Gen ->SubContract_GRPO", SAPbobsCOM.BoUTBTableType.bott_Document)

            Me.AddColumns("@GEN_SC_GRPO", "CardCode", "Vendor Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_GRPO", "CardName", "Vendor Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_GRPO", "ContPer", "Contact Person", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_GRPO", "VendRef", "Vendor Ref. No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_GRPO", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            Me.AddColumns("@GEN_SC_GRPO", "PostDate", "PostDate", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_SC_GRPO", "DelDate", "DelDate", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_SC_GRPO", "DocDate", "DocDate", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_SC_GRPO", "Owner", "Owner", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_GRPO", "OwnerCod", "OwnerCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_GRPO", "VendWhs", "VendWhs", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_GRPO", "Buyer", "Buyer", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_GRPO", "PayTrms", "PayTrms", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_GRPO", "PayCode", "PayCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_GRPO", "PayNum", "Payment No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)

            Me.AddColumns("@GEN_SC_GRPO", "JourRem", "JourRem", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_GRPO", "TotBefTa", "TotBefTax", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SC_GRPO", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SC_GRPO", "Tax", "Tax", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SC_GRPO", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)
            Me.AddColumns("@GEN_SC_GRPO", "DCNo", "Delivery Challan No", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_GRPO", "DCDate", "DeliveryChallanDate", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_SC_GRPO", "PONo", "SC_PONo", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_GRPO", "PODocNo", "SC_PODocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_GRPO", "PODate", "SC_PODate", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_SC_GRPO", "GINO", "GoodsIssueNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("@GEN_SC_GRPO", "GRNO", "GoodsReceiptNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("@GEN_SC_GRPO", "GIDocNO", "GoodsIssue Doc.No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("@GEN_SC_GRPO", "GRDocNO", "GoodsReceipt Doc No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            Me.AddColumns("@GEN_SC_GRPO", "SONo", "SalesOrderNo.", SAPbobsCOM.BoFieldTypes.db_Alpha, 20) '
            Me.AddColumns("@GEN_SC_GRPO", "SODNo", "Sales OrderDNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 20) '
            Me.AddColumns("@GEN_SC_GRPO", "scdcno", "DC No", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)



            Me.AddTable("GEN_SC_GRPO_D0", "Gen ->SubContract_GRPO Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            Me.AddColumns("@GEN_SC_GRPO_D0", "ItemCode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_GRPO_D0", "ItemDesc", "ItemDesc", SAPbobsCOM.BoFieldTypes.db_Alpha, 250)
            Me.AddColumns("@GEN_SC_GRPO_D0", "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SC_GRPO_D0", "Price", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SC_GRPO_D0", "TotalLC", "TotalLC", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SC_GRPO_D0", "TaxRate", "TaxRate", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SC_GRPO_D0", "TaxAmt", "TaxAmt", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SC_GRPO_D0", "TaxCode", "TaxCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_GRPO_D0", "Whs", "Whs", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_GRPO_D0", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_SC_GRPO_D0", "UOM", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_SC_GRPO_D0", "RecdQty", "Received Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SC_GRPO_D0", "POPrice", "POPrice", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SC_GRPO_D0", "SerPrice", "SerPrice", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SC_GRPO_D0", "Line", "Line", SAPbobsCOM.BoFieldTypes.db_Numeric, 11)

            Me.AddTable("GEN_SC_GRPO_D1", "Gen ->SubContract_GRPO RM", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            Me.AddColumns("@GEN_SC_GRPO_D1", "ItemCode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_GRPO_D1", "ItemName", "ItemName", SAPbobsCOM.BoFieldTypes.db_Alpha, 250)
            Me.AddColumns("@GEN_SC_GRPO_D1", "ItemQty", "ItemQty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SC_GRPO_D1", "WhsQty", "WhsQty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SC_GRPO_D1", "ItemCost", "ItemCost", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_SC_GRPO_D1", "Parent", "Parent", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_SC_GRPO_D1", "Whs", "Whs", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_SC_GRPO_D1", "Line", "Line", SAPbobsCOM.BoFieldTypes.db_Numeric, 11)
            Me.AddColumns("@GEN_SC_GRPO_D1", "RecdQty", "RecdQty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SC_GRPO_D1", "BOMQty", "BOMQty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_SC_GRPO_D1", "TotCost", "TotalCost", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)


            If Not Me.UDOExists("GEN_SC_GRPO") Then
                Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
                Me.registerUDO("GEN_SC_GRPO", "GEN: Sub_Contract_GRPO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "GEN_SC_GRPO", "GEN_SC_GRPO_D0", "GEN_SC_GRPO_D1", "")
                findAliasNDescription = Nothing
            End If

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SubContractReturn()
        Try
            Me.AddTable("GEN_SC_RET", "Gen -> SubContractReturn", SAPbobsCOM.BoUTBTableType.bott_Document)

            Me.AddColumns("@GEN_SC_RET", "CardCode", "Vendor Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_RET", "CardName", "Vendor Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_RET", "SCDocNo", "Sub Cont.DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) ' ''''
            Me.AddColumns("@GEN_SC_RET", "SCNo", "Sub Contract No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) ' ''''
            Me.AddColumns("@GEN_SC_RET", "SCDat", "Sub Contract Date", SAPbobsCOM.BoFieldTypes.db_Date) '
            Me.AddColumns("@GEN_SC_RET", "RefNo", "Vendor Ref. No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_RET", "ContPer", "Contact Person", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_RET", "DCDat", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date) '
            Me.AddColumns("@GEN_SC_RET", "DelDat", "Delivery Date", SAPbobsCOM.BoFieldTypes.db_Date) '
            Me.AddColumns("@GEN_SC_RET", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 15) '
            Me.AddColumns("@GEN_SC_RET", "DCNo", "Doc Num", SAPbobsCOM.BoFieldTypes.db_Numeric) ' ''''
            Me.AddColumns("@GEN_SC_RET", "Buyer", "Buyer", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_RET", "Owner", "Owner", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_RET", "OwnerCod", "OwnerCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 25)
            Me.AddColumns("@GEN_SC_RET", "Rmrk", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo) '
            Me.AddColumns("@GEN_SC_RET", "ItemNo", "BOM Item", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_RET", "Qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity) '
            Me.AddColumns("@GEN_SC_RET", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price) '
            Me.AddColumns("@GEN_SC_RET", "InvTrNo", "InvTrNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 20) '
            Me.AddColumns("@GEN_SC_RET", "InvTrDNo", "InvTr Doc.No", SAPbobsCOM.BoFieldTypes.db_Alpha, 20) '

            Me.AddTable("GEN_SC_RET_D0", "Gen ->SubCont. Return Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            Me.AddColumns("@GEN_SC_RET_D0", "ItemNo", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_RET_D0", "Desc", "Item Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 250) '
            Me.AddColumns("@GEN_SC_RET_D0", "FWhs", "From Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_RET_D0", "TWhs", "To Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_RET_D0", "Stock", "Stock Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity) '
            Me.AddColumns("@GEN_SC_RET_D0", "Unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 25) '
            Me.AddColumns("@GEN_SC_RET_D0", "Qty", "Actual Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity) '
            Me.AddColumns("@GEN_SC_RET_D0", "IssQty", "Issued Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity) '
            Me.AddColumns("@GEN_SC_RET_D0", "Rmrk", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 100) '
            Me.AddColumns("@GEN_SC_RET_D0", "Price", "Price", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price) '
            Me.AddColumns("@GEN_SC_RET_D0", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price) '
            Me.AddColumns("@GEN_SC_RET_D0", "BOMQty", "BOMQty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity) '

            If Not Me.UDOExists("GEN_SC_RET") Then
                Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
                Me.registerUDO("GEN_SC_RET", "GEN: Sub_Contract_Return", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "GEN_SC_RET", "GEN_SC_RET_D0", "", "")
                findAliasNDescription = Nothing
            End If


        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_UNIT_MST()
        Try
            Me.AddTable("GEN_UNIT_MST", "Gen->Unit Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            Me.AddTable("GEN_UNIT_MST_D0", "Gen->Unit Master Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)

            Me.AddColumns("@GEN_UNIT_MST_D0", "process", "Process", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_UNIT_MST_D0", "inwhs", "Input Whs", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_UNIT_MST_D0", "outwhs", "Output Whs", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_UNIT_MST_D0", "stwhs", "Stored Whs", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            If Not Me.UDOExists("GEN_UNIT_MST") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}}
                Me.registerUDO("GEN_UNIT_MST", "GEN_UNIT_MST", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_UNIT_MST", "GEN_UNIT_MST_D0", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub UBG_TARIF_MST()
        Try
            Me.AddTable("UBG_TARIF_MST", "Goods Tariff Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            If Not Me.UDOExists("UBG_TARIF_MST") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}}
                Me.registerUDO("UBG_TARIF_MST", "UBG_TARIF_MST", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "UBG_TARIF_MST", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub UBG_GARMENT_MST()
        Try
            Me.AddTable("UBG_GARMENT_MST", "Garments Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            If Not Me.UDOExists("UBG_GARMENT_MST") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}}
                Me.registerUDO("UBG_GARMENT_MST", "UBG_GARMENT_MST", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "UBG_GARMENT_MST", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub UBG_DBK_LST()
        Try
            Me.AddTable("UBG_DBK_LST", "DBK/CAP Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            Me.AddTable("UBG_DBK_LST_D0", "DBK/CAP Master Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)

            Me.AddColumns("@UBG_DBK_LST", "gar", "Garments Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

            Me.AddColumns("@UBG_DBK_LST_D0", "tarif", "Tariff Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            Me.AddColumns("@UBG_DBK_LST_D0", "desc", "Description of Goods", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@UBG_DBK_LST_D0", "dbk", "DBK %", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Rate)
            Me.AddColumns("@UBG_DBK_LST_D0", "cap", "CAP/UNIT", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Rate)
            If Not Me.UDOExists("UBG_DBK_LST") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}}
                Me.registerUDO("UBG_DBK_LST", "UBG_DBK_LST", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "UBG_DBK_LST", "UBG_DBK_LST_D0", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub GEN_GRPO_LCOSTS()
        Try
            Me.AddTable("GEN_GRPO_LCOSTS", "Gen->Landed Costs", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            Me.AddTable("GEN_GRPO_LCOSTS_D0", "Gen->Landed Costs Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)

            Me.AddColumns("@GEN_GRPO_LCOSTS", "grpono", "GRPO No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_GRPO_LCOSTS", "macid", "MAC_ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)

            Me.AddColumns("@GEN_GRPO_LCOSTS_D0", "lname", "Landed Costs Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_GRPO_LCOSTS_D0", "glacct", "GL Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_GRPO_LCOSTS_D0", "rate", "Rate", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_GRPO_LCOSTS_D0", "qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_GRPO_LCOSTS_D0", "value", "Total Value", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_GRPO_LCOSTS_D0", "macid", "MAC-ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            If Not Me.UDOExists("GEN_GRPO_LCOSTS") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}, {"u_grpono", "u_grpono"}, {"u_macid", "u_macid"}}
                Me.registerUDO("GEN_GRPO_LCOSTS", "GEN_GRPO_LCOSTS", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_GRPO_LCOSTS", "GEN_GRPO_LCOSTS_D0", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
   
    Sub GEN_PROCESS_MST()
        Try
            Me.AddTable("GEN_PROCESS_MST", "GEN->Process Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            If Not Me.UDOExists("GEN_PROCESS_MST") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}}
                Me.registerUDO("GEN_PROCESS_MST", "GEN_PROCESS_MST", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_PROCESS_MST", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_PROD_PROCESS()
        Try
            Me.AddTable("GEN_PROD_PRCS", "Gen->Production Process", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            Me.AddTable("GEN_PROD_PRCS_D0", "Gen->Production Process Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)

            Me.AddColumns("@GEN_PROD_PRCS", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_PROD_PRCS", "itemname", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_PROD_PRCS", "stwhs", "Stored Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_PROD_PRCS", "inwhs", "In Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_PROD_PRCS", "outwhs", "Out Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_PROD_PRCS", "cstbom", "BOM No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_PROD_PRCS", "soref", "SO Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

            Me.AddColumns("@GEN_PROD_PRCS_D0", "process", "Process", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_PROD_PRCS_D0", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_PROD_PRCS_D0", "itemname", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_PROD_PRCS_D0", "sfgcode", "SFG Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_PROD_PRCS_D0", "sfgname", "SFG Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_PROD_PRCS_D0", "sfgqty", "SFG Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)

            If Not Me.UDOExists("GEN_PROD_PRCS") Then
                Dim findAliasNDescription = New String(,) {{"u_itemcode", "u_itemcode"}, {"u_itemname", "u_itemname"}}
                Me.registerUDO("GEN_PROD_PRCS", "GEN_PROD_PRCS", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_PROD_PRCS", "GEN_PROD_PRCS_D0", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_CUST_BOM()
        Try
            Me.AddTable("GEN_CUST_BOM", "Gen->Custom BOM", SAPbobsCOM.BoUTBTableType.bott_Document)
            Me.AddTable("GEN_CUST_BOM_D0", "Gen->Custom BOM Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            Me.AddColumns("@GEN_CUST_BOM", "sono", "Sales Order No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_CUST_BOM", "soref", "Sales Order Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_CUST_BOM", "closed", "Closed", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)
            Me.AddColumns("@GEN_CUST_BOM", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_CUST_BOM", "itemname", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_CUST_BOM", "docdate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_CUST_BOM", "empname", "Employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_CUST_BOM", "qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_CUST_BOM", "unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_CUST_BOM", "tprice", "Total Price", SAPbobsCOM.BoFieldTypes.db_Float, 20, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_CUST_BOM", "mprice", "Manual Price", SAPbobsCOM.BoFieldTypes.db_Float, 20, SAPbobsCOM.BoFldSubTypes.st_Price)
            Dim ValidValues = New String(,) {{"NEW", "NEW"}, {"ACTIVE", "ACTIVE"}, {"CHANGE", "CHANGE"}}
            Me.AddColumns("@GEN_CUST_BOM", "status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues)

            Me.AddColumns("@GEN_CUST_BOM_D0", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_CUST_BOM_D0", "itemname", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_CUST_BOM_D0", "size", "Size", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_CUST_BOM_D0", "qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_CUST_BOM_D0", "ordrqty", "Ordr Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_CUST_BOM_D0", "per", "Percentage of Wastage", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage)
            Me.AddColumns("@GEN_CUST_BOM_D0", "totqty", "Total Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_CUST_BOM_D0", "uom", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_CUST_BOM_D0", "process", "Process", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            '  Me.AddColumns("@GEN_CUST_BOM_D0", "price", "Price", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            ValidValues = New String(,) {{"B", "Backflush"}, {"M", "Manual"}}
            Dim DefaultVal = New String(,) {{"M", "Manual"}}
            Me.AddColumns("@GEN_CUST_BOM_D0", "issmthd", "Issue Method", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)
            ValidValues = New String(,) {{"NEW", "NEW"}, {"ACTIVE", "ACTIVE"}, {"CHANGE", "CHANGE"}, {"DELETE", "DELETE"}}
            Me.AddColumns("@GEN_CUST_BOM_D0", "status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues)
            Me.AddColumns("@GEN_CUST_BOM_D0", "cardcode", "Supplier", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_CUST_BOM_D0", "cardname", "Supplier Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            ValidValues = New String(,) {{"YES", "YES"}, {"NO", "NO"}}
            DefaultVal = New String(,) {{"NO", "NO"}}
            Me.AddColumns("@GEN_CUST_BOM_D0", "deleted", "Deleted", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)
            Me.AddColumns("@GEN_CUST_BOM_D0", "remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 150)
            Me.AddColumns("@GEN_CUST_BOM_D0", "place", "Placement", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            'Vijeesh
            Me.AddColumns("@GEN_CUST_BOM_D0", "fremark", "Factory Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            ' Me.AddColumns("@GEN_CUST_BOM_D0", "fremark1", "Factory Remarks1", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            Me.AddColumns("@GEN_CUST_BOM_D0", "iprice", "Item Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)

            'Vijeesh
            If Not Me.UDOExists("GEN_CUST_BOM") Then
                Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}, {"u_itemcode", "u_itemcode"}, {"u_itemname", "u_itemname"}}
                Me.registerUDO("GEN_CUST_BOM", "GEN_CUST_BOM", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "GEN_CUST_BOM", "GEN_CUST_BOM_D0", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub Assortment_Master()
        Me.AddTable("GEN_ASSORTMENT", "Gen-> Assortment Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        Me.AddTable("GEN_ASSORTMENT_D0", "Gen->Assortment Master Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        Me.AddColumns("@GEN_ASSORTMENT_D0", "size", "Size", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, SAPbobsCOM.BoFldSubTypes.st_None)
        If Not Me.UDOExists("GEN_ASSORTMENT") Then
            Dim findAliasNDescription = New String(,) {{"Name", "Name"}}
            Me.registerUDO("GEN_ASSORTMENT", "GEN_ASSORTMENT", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_ASSORTMENT", "GEN_ASSORTMENT_D0", "", "")
            findAliasNDescription = Nothing
        End If
    End Sub

    Sub Size_Master()
        Me.AddTable("GEN_SIZE_MST", "Gen->Size Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        If Not Me.UDOExists("GEN_SIZE_MST") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("GEN_SIZE_MST", "GEN_SIZE_MST", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_SIZE_MST", "", "", "", SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES)
            findAliasNDescription = Nothing
        End If
    End Sub

    Sub Size_Matrix_Order()
        Me.AddTable("GEN_SZ_ORDR", "Gen->Sales Order", SAPbobsCOM.BoUTBTableType.bott_NoObject)

        Me.AddColumns("@GEN_SZ_ORDR", "sono", "Sales Order No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_SZ_ORDR", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_SZ_ORDR", "asrtcode", "Assorted Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_SZ_ORDR", "macid", "MAC ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        Me.AddColumns("@GEN_SZ_ORDR", "size", "Size", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_SZ_ORDR", "qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_SZ_ORDR", "cutqty", "Cut Qty", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
    End Sub

    Sub Ordr_Tmp()
        Me.AddTable("ORDR_ITEMS", "Order Items", SAPbobsCOM.BoUTBTableType.bott_NoObject)

        Me.AddColumns("@ORDR_ITEMS", "sono", "SONo", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@ORDR_ITEMS", "macid", "MACID", SAPbobsCOM.BoFieldTypes.db_Alpha, 60)
        Me.AddColumns("@ORDR_ITEMS", "itemcode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@ORDR_ITEMS", "itemname", "ItemName", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@ORDR_ITEMS", "asrtcode", "Assorted Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@ORDR_ITEMS", "qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
    End Sub

    Sub GEN_FIN_PRCS()
        Try
            Me.AddTable("GEN_FIN_PRCS", "Gen->Finish Process", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            If Not Me.UDOExists("GEN_FIN_PRCS") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}}
                Me.registerUDO("GEN_FIN_PRCS", "GEN_FIN_PRCS", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_FIN_PRCS", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_STH_PRCS()
        Try
            Me.AddTable("GEN_STH_PRCS", "Gen->Stitching Process", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            If Not Me.UDOExists("GEN_STH_PRCS") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}}
                Me.registerUDO("GEN_STH_PRCS", "GEN_STH_PRCS", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_STH_PRCS", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_FIN_SETUP()
        Try
            Me.AddTable("GEN_FIN_SETUP", "Gen->Finish Setup", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            Me.AddTable("GEN_FIN_SETUP_D0", "Gen->Finish Setup Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)

            Me.AddColumns("@GEN_FIN_SETUP", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_FIN_SETUP", "itemname", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_FIN_SETUP", "prodno", "Production Order No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_FIN_SETUP", "cardcode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_FIN_SETUP", "cardname", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)

            Me.AddColumns("@GEN_FIN_SETUP_D0", "prcs", "Process", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_FIN_SETUP_D0", "prcsname", "Process Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_FIN_SETUP_D0", "reqdno", "Required No", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_FIN_SETUP_D0", "cappm", "Capacity PM", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_FIN_SETUP_D0", "trgtop", "Target Output", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)

            If Not Me.UDOExists("GEN_FIN_SETUP") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}, {"u_itemcode", "u_itemcode"}, {"u_prodno", "u_prodno"}, {"u_cardcode", "u_cardcode"}}
                Me.registerUDO("GEN_FIN_SETUP", "GEN_FIN_SETUP", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_FIN_SETUP", "GEN_FIN_SETUP_D0", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_FIN_DESCR()
        Try
            Me.AddTable("GEN_FIN_DESCR", "Gen->Finish Operation Screen", SAPbobsCOM.BoUTBTableType.bott_Document)
            Me.AddTable("GEN_FIN_DESCR_D0", "Gen->Finish Operation Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            Me.AddColumns("@GEN_FIN_DESCR", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_FIN_DESCR", "itemname", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_FIN_DESCR", "prodno", "Production Order No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_FIN_DESCR", "prdentry", "Production Order Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_FIN_DESCR", "cardcode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_FIN_DESCR", "cardname", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_FIN_DESCR", "docdate", "Document date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_FIN_DESCR", "empname", "Employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)


            Me.AddColumns("@GEN_FIN_DESCR_D0", "prcs", "Process", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_FIN_DESCR_D0", "prcsname", "Process Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_FIN_DESCR_D0", "hour1", "Hour 1", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_FIN_DESCR_D0", "hour2", "Hour 2", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_FIN_DESCR_D0", "hour3", "Hour 3", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_FIN_DESCR_D0", "hour4", "Hour 4", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_FIN_DESCR_D0", "hour5", "Hour 5", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_FIN_DESCR_D0", "hour6", "Hour 6", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_FIN_DESCR_D0", "hour7", "Hour 7", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_FIN_DESCR_D0", "hour8", "Hour 8", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_FIN_DESCR_D0", "ot", "Over Time", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_FIN_DESCR_D0", "total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)


            If Not Me.UDOExists("GEN_FIN_DESCR") Then
                Dim findAliasNDescription = New String(,) {{"DocEntry", "DocEntry"}, {"DocNum", "DocNum"}, {"u_itemcode", "u_itemcode"}, {"u_prodno", "u_prodno"}, {"u_cardcode", "u_cardcode"}}
                Me.registerUDO("GEN_FIN_DESCR", "GEN_FIN_DESCR", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "GEN_FIN_DESCR", "GEN_FIN_DESCR_D0", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_STH_DESCR()
        Try
            Me.AddTable("GEN_STH_DESCR", "Gen->Finish Operation Screen", SAPbobsCOM.BoUTBTableType.bott_Document)
            Me.AddTable("GEN_STH_DESCR_D0", "Gen->Finish Operation Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            Me.AddColumns("@GEN_STH_DESCR", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_STH_DESCR", "itemname", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_STH_DESCR", "prodno", "Production Order No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_STH_DESCR", "prdentry", "Production Order Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_STH_DESCR", "docdate", "Document date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_STH_DESCR", "empname", "Employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)


            Me.AddColumns("@GEN_STH_DESCR_D0", "line", "Line No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_STH_DESCR_D0", "linename", "Line Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_STH_DESCR_D0", "prcs", "Process", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_STH_DESCR_D0", "prcsname", "Process Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_STH_DESCR_D0", "hour1", "Hour 1", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_STH_DESCR_D0", "hour2", "Hour 2", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_STH_DESCR_D0", "hour3", "Hour 3", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_STH_DESCR_D0", "hour4", "Hour 4", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_STH_DESCR_D0", "hour5", "Hour 5", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_STH_DESCR_D0", "hour6", "Hour 6", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_STH_DESCR_D0", "hour7", "Hour 7", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_STH_DESCR_D0", "hour8", "Hour 8", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_STH_DESCR_D0", "ot", "Over Time", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_STH_DESCR_D0", "total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)


            If Not Me.UDOExists("GEN_STH_DESCR") Then
                Dim findAliasNDescription = New String(,) {{"DocEntry", "DocEntry"}, {"DocNum", "DocNum"}, {"u_itemcode", "u_itemcode"}, {"u_prodno", "u_prodno"}}
                Me.registerUDO("GEN_STH_DESCR", "GEN_STH_DESCR", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "GEN_STH_DESCR", "GEN_STH_DESCR_D0", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_LINE_MST()
        Try
            Me.AddTable("GEN_LINE_MST", "Gen->Line Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            If Not Me.UDOExists("GEN_LINE_MST") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}}
                Me.registerUDO("GEN_LINE_MST", "GEN_LINE_MST", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_LINE_MST", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_LINE_TYPE()
        Try
            Me.AddTable("GEN_LINE_TYPE", "Gen->Line Type", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            If Not Me.UDOExists("GEN_LINE_TYPE") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}}
                Me.registerUDO("GEN_LINE_TYPE", "GEN_LINE_TYPE", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_LINE_TYPE", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_MACH_POOL()
        Try
            Me.AddTable("GEN_MACH_POOL", "Gen->Machine Pool", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            Me.AddTable("GEN_MACH_POOL_D0", "Gen->Machine Pool Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)

            Me.AddColumns("@GEN_MACH_POOL_D0", "type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_MACH_POOL_D0", "typename", "Type Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_MACH_POOL_D0", "nom", "No Of Machines", SAPbobsCOM.BoFieldTypes.db_Numeric, 10)
            
            If Not Me.UDOExists("GEN_MACH_POOL") Then
                Dim findAliasNDescription = New String(,) {{"Name", "Name"}}
                Me.registerUDO("GEN_MACH_POOL", "GEN_MACH_POOL", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_MACH_POOL", "GEN_MACH_POOL_D0", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_OTP()
        Me.AddTable("GEN_OTP", "GEN_OTP->OTP Genration", SAPbobsCOM.BoUTBTableType.bott_Document)
        Me.AddColumns("@GEN_OTP", "PrdtNo", "Product No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        Me.AddColumns("@GEN_OTP", "PrdtDes", "Product Desc", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_OTP", "OTPdate", "OTPdate", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@GEN_OTP", "OTPuser", "OTPuser", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        Me.AddColumns("@GEN_OTP", "OTPNo", "OTPNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 8)
        Me.AddColumns("@GEN_OTP", "OTPRmk", "OTPRemarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 250)
        Me.AddColumns("@GEN_OTP", "BomCrdt", "BOMCreationDate", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@GEN_OTP", "BomupDt", "BOMUpdateDate", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@GEN_OTP", "BomCtUsr", "BomCreateUser", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        Me.AddColumns("@GEN_OTP", "BomRvn", "BomRevisionNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 8)

    End Sub

    Sub GEN_MACH_ALLOC()
        Try
            Me.AddTable("GEN_MACH_ALLOC", "Gen->Machine Allocation", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            Me.AddTable("GEN_MACH_ALLOC_D0", "Gen->Machine Allocation Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)

            Me.AddColumns("@GEN_MACH_ALLOC", "manual", "Manual", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)
            Me.AddColumns("@GEN_MACH_ALLOC", "stdate", "Start Date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_MACH_ALLOC", "eddate", "End Date", SAPbobsCOM.BoFieldTypes.db_Date)

            Me.AddColumns("@GEN_MACH_ALLOC_D0", "sono", "Sales Order No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_MACH_ALLOC_D0", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_MACH_ALLOC_D0", "itemname", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_MACH_ALLOC_D0", "qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_MACH_ALLOC_D0", "deldate", "Delivery Date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_MACH_ALLOC_D0", "stdate", "Start Date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_MACH_ALLOC_D0", "eddate", "End Date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_MACH_ALLOC_D0", "unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_MACH_ALLOC_D0", "lineno", "Line No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_MACH_ALLOC_D0", "nom", "No Of Machines", SAPbobsCOM.BoFieldTypes.db_Numeric, 10)
            Me.AddColumns("@GEN_MACH_ALLOC_D0", "trgtcode", "Target Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

            If Not Me.UDOExists("GEN_MACH_ALLOC") Then
                Dim findAliasNDescription = New String(,) {{"Name", "Name"}}
                Me.registerUDO("GEN_MACH_ALLOC", "GEN_MACH_ALLOC", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_MACH_ALLOC", "GEN_MACH_ALLOC_D0", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub User_Unit_Linkage()
        Me.AddTable("GEN_USR_UNIT", "Gen->User UNIT Link", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        Me.AddColumns("@GEN_USR_UNIT", "user", "User", SAPbobsCOM.BoFieldTypes.db_Alpha, 60)
        Me.AddColumns("@GEN_USR_UNIT", "unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        If Not Me.UDOExists("GEN_USR_UNIT") Then
            Dim findAliasNDescription As String(,) = {{"Code", "Code"}, {"Name", "Name"}, {"u_user", "u_user"}, {"u_unit", "u_unit"}}
            Me.registerUDO("GEN_USR_UNIT", "GEN_USR_UNIT", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_USR_UNIT", "", "", "", SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES)
            findAliasNDescription = Nothing
        End If
    End Sub
    Sub Journal_Entry()
        Me.AddTable("GEN_M_JE", "Gen->Manual JournalEntry", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        '  Me.AddColumns("@GEN_USR_UNIT", "user", "User", SAPbobsCOM.BoFieldTypes.db_Alpha, 60)
        ' Me.AddColumns("@GEN_USR_UNIT", "unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        If Not Me.UDOExists("GEN_M_JE") Then
            Dim findAliasNDescription As String(,) = {{"Code", "Code"}, {"Name", "Name"}} ', {"u_user", "u_user"}, {"u_unit", "u_unit"}}
            Me.registerUDO("GEN_M_JE", "GEN_M_JE", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_M_JE", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES)
            findAliasNDescription = Nothing
        End If
    End Sub

    'Sub Journal_Entry()
    '    Try
    '        Me.AddTable("GEN_M_JE", "Gen->Field ID", SAPbobsCOM.BoUTBTableType.bott_MasterData)

    '        If Not Me.UDOExists("GEN_FIELD_ID") Then
    '            Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}}
    '            Me.registerUDO("GEN_FIELD_ID", "GEN_FIELD_ID", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_FIELD_ID", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES)
    '            findAliasNDescription = Nothing
    '        End If
    '    Catch ex As Exception
    '        oApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub

    Sub GEN_CAP_PLAN()
        Try
            Me.AddTable("GEN_CAP_PLAN", "GEN->Capcity Plan", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            Me.AddTable("GEN_CAP_PLAN_D0", "GEN->Capacity Plan Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            Me.AddTable("GEN_CAP_PLAN_D1", "GEN->Capacity Type Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)

            Me.AddColumns("@GEN_CAP_PLAN", "sono", "SO NO", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_CAP_PLAN", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_CAP_PLAN", "qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_CAP_PLAN", "stdate", "Start Date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_CAP_PLAN", "eddate", "End Date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_CAP_PLAN", "unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_CAP_PLAN", "line", "Line", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_CAP_PLAN", "type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_CAP_PLAN", "nom", "Number of Machines", SAPbobsCOM.BoFieldTypes.db_Numeric, 10)
            Me.AddColumns("@GEN_CAP_PLAN", "typename", "Type Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_CAP_PLAN", "ln", "Base Line", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_CAP_PLAN", "basecode", "Base Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_CAP_PLAN", "reqdsam", "Required SAM", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_CAP_PLAN", "mtmop", "MTM Output", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_CAP_PLAN", "trgtop", "Target OutPut", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)

            Me.AddColumns("@GEN_CAP_PLAN_D0", "cdate", "Date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_CAP_PLAN_D0", "avlbl", "Available", SAPbobsCOM.BoFieldTypes.db_Numeric, 10)
            Me.AddColumns("@GEN_CAP_PLAN_D0", "per", "Percentage", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_CAP_PLAN_D0", "qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)

            Me.AddColumns("@GEN_CAP_PLAN_D1", "type", "Type Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_CAP_PLAN_D1", "typename", "Type Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_CAP_PLAN_D1", "nom", "No of Machines", SAPbobsCOM.BoFieldTypes.db_Numeric, 10)


            If Not Me.UDOExists("GEN_CAP_PLAN") Then
                Dim findAliasNDescription = New String(,) {{"u_sono", "u_sono"}, {"u_itemcode", "u_itemcode"}, {"u_unit", "u_unit"}, {"u_line", "u_line"}, {"u_type", "u_type"}, {"u_ln", "u_ln"}, {"Code", "Code"}, {"u_basecode", "u_basecode"}}
                Me.registerUDO("GEN_CAP_PLAN", "GEN_CAP_PLAN", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_CAP_PLAN", "GEN_CAP_PLAN_D0", "GEN_CAP_PLAN_D1", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MaterialRequisitionNote()
        Me.AddTable("GEN_MREQ", "Material Requisition", SAPbobsCOM.BoUTBTableType.bott_Document)
        Me.AddTable("GEN_MREQ_D0", "Material Requisition Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)



        Dim ValidValues = New String(,) {{"Open", "Open"}, {"Closed", "Closed"}}
        Me.AddColumns("@GEN_MREQ", "status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues)

        ValidValues = New String(,) {{"Open", "Open"}, {"Closed", "Closed"}}
        Me.AddColumns("@GEN_MREQ", "mstatus", "Manual Closing Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues)

        ValidValues = New String(,) {{"", ""}, {"Regular", "Regular"}, {"Excess", "Excess"}, {"Consumable", "Consumable"}, {"Sampling", "Sampling"}, {"Production Consumable", "Production Consumable"}}
        Me.AddColumns("@GEN_MREQ", "type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues)
        Me.AddColumns("@GEN_MREQ", "sono", "Sales Order No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_MREQ", "approve", "Approve", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)
        Me.AddColumns("@GEN_MREQ", "soentry", "Sales Order Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_MREQ", "soref", "Sales Order Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_MREQ", "docdate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@GEN_MREQ", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_MREQ", "itemname", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_MREQ", "ordrqty", "Ordered Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        Me.AddColumns("@GEN_MREQ", "excsqty", "Excess Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        Me.AddColumns("@GEN_MREQ", "sfgcode", "Semi Finished Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_MREQ", "sfgname", "Semi Finished Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_MREQ", "whs", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_MREQ", "wipwhs", "WIP Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_MREQ", "remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254)
        Me.AddColumns("@GEN_MREQ", "bomrem", "BOM Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None)
        Me.AddColumns("@GEN_MREQ", "empname", "Employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_MREQ", "unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_MREQ", "process", "Process", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_MREQ", "EMP_ID", "Buyer's Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
        Me.AddColumns("@GEN_MREQ", "EMP_NAME", "Buyer's Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)

        Me.AddColumns("@GEN_MREQ_D0", "chk", "CheckBox", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)
        Me.AddColumns("@GEN_MREQ_D0", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_MREQ_D0", "itemname", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_MREQ_D0", "rqstqty", "Requested Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        Me.AddColumns("@GEN_MREQ_D0", "tol", "Tolerance Percent", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage)
        Me.AddColumns("@GEN_MREQ_D0", "uom", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_MREQ_D0", "reqdqty", "Required Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        Me.AddColumns("@GEN_MREQ_D0", "wipavlbl", "WIP Availability", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        Me.AddColumns("@GEN_MREQ_D0", "totavlbl", "Total Availability", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        Me.AddColumns("@GEN_MREQ_D0", "issued", "Issued", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        Me.AddColumns("@GEN_MREQ_D0", "returned", "Returned", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        Me.AddColumns("@GEN_MREQ_D0", "totis", "Total Issues", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        Me.AddColumns("@GEN_MREQ_D0", "whs", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None)
        Me.AddColumns("@GEN_MREQ_D0", "remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
        Me.AddColumns("@GEN_MREQ_D0", "stat", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None)
        'Vijeesh
        Me.AddColumns("@GEN_MREQ_D0", "minchk", "MINCheckBox", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)
        'Vijeesh
        If Not Me.UDOExists("GEN_MREQ") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}, {"u_status", "u_status"}, {"u_type", "u_type"}, {"u_sono", "u_sono"}, {"u_itemcode", "u_itemcode"}, {"u_sfgcode", "u_sfgcode"}, {"u_docdate", "u_docdate"}}
            Me.registerUDO("GEN_MREQ", "GEN_MREQ", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "GEN_MREQ", "GEN_MREQ_D0", "", "")
            findAliasNDescription = Nothing
        End If
    End Sub

    Sub Warehouse_User_Alert()
        Me.AddTable("GEN_WHS_USR", "Gen->Warehouse User", SAPbobsCOM.BoUTBTableType.bott_MasterData)

        Me.AddColumns("@GEN_WHS_USR", "whs", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_WHS_USR", "user", "User code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Dim ValidValues = New String(,) {{"YES", "YES"}, {"NO", "NO"}}
        Me.AddColumns("@GEN_WHS_USR", "alert", "Send Alert", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues)
        If Not Me.UDOExists("GEN_WHS_USR") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"u_whs", "u_whs"}, {"u_user", "u_user"}, {"u_alert", "u_alert"}}
            Me.registerUDO("GEN_WHS_USR", "GEN_WHS_USR", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_WHS_USR", "", "", "", SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES)
            findAliasNDescription = Nothing
        End If
    End Sub
    Sub ARInvoice_GLAccount()
        Me.AddTable("GEN_GL_ACCOUNT", "Gen->APInvoice-GLAccount", SAPbobsCOM.BoUTBTableType.bott_NoObject)

        Me.AddColumns("@GEN_GL_ACCOUNT", "pc", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_GL_ACCOUNT", "glacc", "GLAccount", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_GL_ACCOUNT", "tinv", "Type Of Invoice", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        'Dim ValidValues = New String(,) {{"", ""}, {"", ""}}
        'Me.AddColumns("@GEN_GL_ACCOUNT", "tinv", "Type Of Invoice", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues)
        'If Not Me.UDOExists("GEN_GL_ACCOUNT") Then
        '    Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"u_pc", "u_pc"}, {"u_glacc", "u_glacc"}, {"u_tinv", "u_tinv"}}
        '    Me.registerUDO("GEN_GL_ACCOUNT", "GEN_GL_ACCOUNT", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "GEN_GL_ACCOUNT", "", "", "", SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES)
        '    findAliasNDescription = Nothing
        'End If
    End Sub

    Sub Production_Transfer_Note()
        Me.AddTable("GEN_PTN", "Gen->PTN", SAPbobsCOM.BoUTBTableType.bott_Document)

        Me.AddColumns("@GEN_PTN", "docdate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
        Dim ValidValues = New String(,) {{"Open", "Open"}, {"Consumed", "Consumed"}, {"Confirmed", "Confirmed"}}
        Me.AddColumns("@GEN_PTN", "status", "Document Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues)
        Me.AddColumns("@GEN_PTN", "sono", "Sales Order NO", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_PTN", "soentry", "Sales Order Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_PTN", "soref", "Sales Order Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_PTN", "prdno", "Production Order No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_PTN", "prdentry", "Production Order Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_PTN", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_PTN", "itemname", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
        Me.AddColumns("@GEN_PTN", "compdate", "Completion Date", SAPbobsCOM.BoFieldTypes.db_Date)
        Me.AddColumns("@GEN_PTN", "prdqty", "Production Qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        Me.AddColumns("@GEN_PTN", "prdoqty", "Production Open Qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        Me.AddColumns("@GEN_PTN", "compqty", "Completed Qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        Me.AddColumns("@GEN_PTN", "accpqty", "Accepted Qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        Me.AddColumns("@GEN_PTN", "accpwhs", "Accepted Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_PTN", "rejqty", "Rejected Qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        Me.AddColumns("@GEN_PTN", "rejwhs", "Rejected Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_PTN", "rewqty", "Rework Qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        Me.AddColumns("@GEN_PTN", "rewwhs", "Rework Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_PTN", "unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Me.AddColumns("@GEN_PTN", "process", "Process", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        If Not Me.UDOExists("GEN_PTN") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}, {"u_status", "u_status"}, {"u_sono", "u_sono"}, {"u_itemcode", "u_itemcode"}, {"u_docdate", "u_docdate"}}
            Me.registerUDO("GEN_PTN", "GEN_PTN", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "GEN_PTN", "", "", "")
            findAliasNDescription = Nothing
        End If

    End Sub

    Sub GEN_ITM_MST()
        Try
            Me.AddTable("GEN_ITM_MST", "Gen->Item Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            If Not Me.UDOExists("GEN_ITM_MST") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}}
                Me.registerUDO("GEN_ITM_MST", "GEN_ITM_MST", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_ITM_MST", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_ITM_TYPE()
        Try
            Me.AddTable("GEN_ITM_TYPE", "Gen->Item Type", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            Me.AddColumns("@GEN_ITM_TYPE", "type", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

            If Not Me.UDOExists("GEN_ITM_TYPE") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}, {"u_type", "u_type"}}
                Me.registerUDO("GEN_ITM_TYPE", "GEN_ITM_TYPE", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_ITM_TYPE", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_SUB_TYPE()
        Try
            Me.AddTable("GEN_SUB_TYPE", "Gen->Item Sub Type", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            Me.AddColumns("@GEN_SUB_TYPE", "desc", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_SUB_TYPE", "type", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

            If Not Me.UDOExists("GEN_SUB_TYPE") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}, {"u_type", "u_type"}, {"u_desc", "u_desc"}}
                Me.registerUDO("GEN_SUB_TYPE", "GEN_SUB_TYPE", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_SUB_TYPE", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_CUST_CODE()
        Try
            Me.AddTable("GEN_CUST_CODE", "Gen->Cust Code", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            If Not Me.UDOExists("GEN_CUST_CODE") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}}
                Me.registerUDO("GEN_CUST_CODE", "GEN_CUST_CODE", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_CUST_CODE", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_STYLE_CODE()
        Try
            Me.AddTable("GEN_STYLE_CODE", "Gen->Style Code", SAPbobsCOM.BoUTBTableType.bott_MasterData)


            Me.AddColumns("@GEN_STYLE_CODE", "desc", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_STYLE_CODE", "type", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            If Not Me.UDOExists("GEN_STYLE_CODE") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}, {"u_type", "u_type"}, {"u_desc", "u_desc"}}
                Me.registerUDO("GEN_STYLE_CODE", "GEN_STYLE_CODE", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_STYLE_CODE", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_COLOR_CODE()
        Try
            Me.AddTable("GEN_COLOR_CODE", "Gen->Color Code", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            Me.AddColumns("@GEN_COLOR_CODE", "desc", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_COLOR_CODE", "type", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            If Not Me.UDOExists("GEN_COLOR_CODE") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}, {"u_type", "u_type"}, {"u_desc", "u_desc"}}
                Me.registerUDO("GEN_COLOR_CODE", "GEN_COLOR_CODE", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_COLOR_CODE", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_ITEM_CREATE()
        Try
            Me.AddTable("GEN_ITEM_CREATE", "Gen->Create Item", SAPbobsCOM.BoUTBTableType.bott_NoObject)

            Me.AddColumns("@GEN_ITEM_CREATE", "itmmst", "Item Master", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_ITEM_CREATE", "fld1", "Field1", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_ITEM_CREATE", "itmtype", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_ITEM_CREATE", "fld2", "Field2", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_ITEM_CREATE", "subtype", "Sub Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_ITEM_CREATE", "fld3", "Field3", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_ITEM_CREATE", "custcode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_ITEM_CREATE", "fld4", "Field4", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_ITEM_CREATE", "style", "Style", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_ITEM_CREATE", "fld5", "Field5", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_ITEM_CREATE", "color", "Color", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_ITEM_CREATE", "fld6", "Field6", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_ITEM_CREATE", "count", "Count", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_ITEM_CREATE", "width", "Width", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_ITEM_CREATE", "size", "Size", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_ITEM_CREATE", "fld7", "Field7", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_ITEM_CREATE", "fld8", "Field8", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_ITEM_CREATE", "quality", "Quality", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_ITEM_CREATE", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_SIZE_CODE()
        Try
            Me.AddTable("GEN_SIZE_CODE", "Gen->Size Code", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            Me.AddColumns("@GEN_SIZE_CODE", "size", "Size", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_SIZE_CODE", "type", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

            If Not Me.UDOExists("GEN_SIZE_CODE") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}, {"U_size", "U_size"}, {"U_type", "U_type"}}
                Me.registerUDO("GEN_SIZE_CODE", "GEN_SIZE_CODE", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_SIZE_CODE", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_FIELD_ID()
        Try
            Me.AddTable("GEN_FIELD_ID", "Gen->Field ID", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            If Not Me.UDOExists("GEN_FIELD_ID") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}}
                Me.registerUDO("GEN_FIELD_ID", "GEN_FIELD_ID", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_FIELD_ID", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_QLTY_CODE()
        Try
            Me.AddTable("GEN_QLTY_CODE", "Gen->Quality Code", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            Me.AddColumns("@GEN_QLTY_CODE", "desc", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_QLTY_CODE", "type", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

            If Not Me.UDOExists("GEN_QLTY_CODE") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}, {"u_type", "u_type"}, {"u_desc", "u_desc"}}
                Me.registerUDO("GEN_QLTY_CODE", "GEN_QLTY_CODE", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_QLTY_CODE", "", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_PARAM_MST()
        Try
            Me.AddTable("GEN_PARAM_MST", "Gen->Param Order", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            Me.AddTable("GEN_PARAM_MST_D0", "Gen-> Param Order Line", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)

            Me.AddColumns("@GEN_PARAM_MST_D0", "field", "Field", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_PARAM_MST_D0", "length", "Length", SAPbobsCOM.BoFieldTypes.db_Numeric, 10)

            If Not Me.UDOExists("GEN_PARAM_MST") Then
                Dim findAliasNDescription = New String(,) {{"Code", "Code"}, {"Name", "Name"}}
                Me.registerUDO("GEN_PARAM_MST", "GEN_PARAM_MST", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "GEN_PARAM_MST", "GEN_PARAM_MST_D0", "", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GEN_COST_SHEET()
        Try
            Me.AddTable("GEN_COST_SHEET", "Gen->Cost Sheet", SAPbobsCOM.BoUTBTableType.bott_Document)
            Me.AddTable("GEN_COST_SHEET_D0", "Gen->Cost Sheet Items", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            Me.AddTable("GEN_COST_SHEET_D1", "GEN->Cost Sheet Exps", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            Me.AddColumns("@GEN_COST_SHEET", "docdate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            Me.AddColumns("@GEN_COST_SHEET", "total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_COST_SHEET", "ototal", "Over Head Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_COST_SHEET", "final", "Final", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)
            Me.AddColumns("@GEN_COST_SHEET", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_COST_SHEET", "cardcode", "Buyer", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_COST_SHEET", "qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_COST_SHEET", "doccur", "Document Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_COST_SHEET", "docrate", "Document Rate", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)
            Me.AddColumns("@GEN_COST_SHEET", "garwash", "Garment Wash", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_COST_SHEET", "fabtype", "Fabric Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_COST_SHEET", "fabfin", "Fabric Finish", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_COST_SHEET", "sam", "SAM", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)
            Me.AddColumns("@GEN_COST_SHEET", "effcy", "Efficiency", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)
            Me.AddColumns("@GEN_COST_SHEET", "maccost", "Machine Cost", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_COST_SHEET", "weight", "Weight", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_COST_SHEET", "wasper", "Wastage Percentage", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage)
            Me.AddColumns("@GEN_COST_SHEET", "wasval", "Wastage Value", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_COST_SHEET", "prfper", "Profit Percentage", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage)
            Me.AddColumns("@GEN_COST_SHEET", "prfval", "Profit Value", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_COST_SHEET", "mtotal", "Material Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_COST_SHEET", "etotal", "Exp Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_COST_SHEET", "gtotal", "Grand Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_COST_SHEET", "costinr", "Cost in INR", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_COST_SHEET", "costusd", "Cost in usd", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)

            Me.AddColumns("@GEN_COST_SHEET_D0", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_COST_SHEET_D0", "itemname", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_COST_SHEET_D0", "itmtype", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            'Dim ValidValues = New String(,) {{"Yes", "Yes"}, {"No", "No"}}
            'Dim DefaultVal = New String(,) {{"No", "No"}}
            'Me.AddColumns("@GEN_COST_SHEET_DO", "import", "Import", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", Nothing, ValidValues, 0, DefaultVal)
            'Me.AddColumns("@GEN_COST_SHEET_DO", "doccur", "Document Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            'Me.AddColumns("@GEN_COST_SHEET_DO", "docrate", "Document Rate", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_COST_SHEET_D0", "qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.AddColumns("@GEN_COST_SHEET_D0", "uom", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_COST_SHEET_D0", "rate", "Rate", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_COST_SHEET_D0", "rowtotal", "Row Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
           
            Me.AddColumns("@GEN_COST_SHEET_D1", "prcs", "Process", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            Me.AddColumns("@GEN_COST_SHEET_D1", "prcsname", "Process Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            Me.AddColumns("@GEN_COST_SHEET_D1", "rate", "Rate", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            Me.AddColumns("@GEN_COST_SHEET_D1", "rowtotal", "Row Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            If Not Me.UDOExists("GEN_COST_SHEET") Then
                Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}, {"u_itemcode", "u_itemcode"}, {"u_cardcode", "u_cardcode"}}
                Me.registerUDO("GEN_COST_SHEET", "GEN_COST_SHEET", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "GEN_COST_SHEET", "GEN_COST_SHEET_D0", "GEN_COST_SHEET_D1", "", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO)
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function keygen(ByVal vtablename As String, Optional ByVal prefix As String = "DOC-") As String

        Dim str As String = ""
        Dim Query As String
        Try
            Query = "SELECT MAX(CAST(Code AS int)) AS code FROM [" + vtablename + "]"
            Dim v_recordset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            v_recordset.DoQuery(Query)
            v_recordset.MoveFirst()
            Dim code As Integer = v_recordset.Fields.Item("code").Value.ToString
            If code > 0 Then
                code += 1
                Dim docid As String = prefix
                If code.ToString.Length < 6 Then
                    For count As Integer = 0 To 5 - code.ToString.Length
                        docid += "0"
                    Next
                End If
                docid += code.ToString
                str = code
            Else
                str = "1"
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_recordset)
            v_recordset = Nothing
            GC.Collect()
        Catch ex As Exception
            oApplication.MessageBox(ex.Message)
        End Try
        keygen = str
    End Function

#End Region

#Region "    -- DataBase Creation --      "

    Function TableExists(ByVal TableName As String) As Boolean
        Dim oTables As SAPbobsCOM.UserTablesMD
        Dim oFlag As Boolean
        oTables = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
        oFlag = oTables.GetByKey(TableName)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oTables)
        Return oFlag
    End Function

    Function ColumnExists(ByVal TableName As String, ByVal FieldID As String) As Boolean
        Try
            Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oFlag As Boolean = True
            rs.DoQuery("Select 1 from [CUFD] Where TableID='" & Trim(TableName) & "' and AliasID='" & Trim(FieldID) & "'")
            If rs.EoF Then oFlag = False
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
            rs = Nothing
            GC.Collect()
            Return oFlag
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Public Function UDOExists(ByVal code As String) As Boolean
        GC.Collect()
        Dim v_UDOMD As SAPbobsCOM.UserObjectsMD
        Dim v_ReturnCode As Boolean
        v_UDOMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        v_ReturnCode = v_UDOMD.GetByKey(code)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UDOMD)
        v_UDOMD = Nothing
        Return v_ReturnCode
    End Function

    Function AddTable(ByVal TableName As String, ByVal TableDescription As String, ByVal TableType As SAPbobsCOM.BoUTBTableType) As Boolean
        Try
            GC.Collect()
            If Not Me.TableExists(TableName) Then
                Dim v_UserTableMD As SAPbobsCOM.UserTablesMD
                oApplication.StatusBar.SetText("Creating Table " & TableName & " ...................", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                v_UserTableMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
                v_UserTableMD.TableName = TableName
                v_UserTableMD.TableDescription = TableDescription
                v_UserTableMD.TableType = TableType
                v_RetVal = v_UserTableMD.Add()
                If v_RetVal <> 0 Then
                    oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                    oApplication.StatusBar.SetText("Failed to Create Table " & TableName & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserTableMD)
                    v_UserTableMD = Nothing
                    GC.Collect()
                    Return False
                Else
                    oApplication.StatusBar.SetText("[@" & TableName & "] - " & TableDescription & " created successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserTableMD)
                    v_UserTableMD = Nothing
                    GC.Collect()
                    DB_Restart = True
                    Return True
                End If
            Else
                GC.Collect()
                Return False
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function AddColumns(ByVal TableName As String, ByVal Name As String, ByVal Description As String, ByVal Type As SAPbobsCOM.BoFieldTypes, Optional ByVal Size As Long = 0, Optional ByVal SubType As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None, Optional ByVal LinkedTable As String = "", Optional ByVal Token As Hashtable = Nothing, Optional ByVal ValidValues As String(,) = Nothing, Optional ByVal iCount As Integer = 0, Optional ByVal DefaultValues As String(,) = Nothing) As Boolean
        Try
            If Not Me.ColumnExists(TableName, Name) Then
                Dim v_UserField As SAPbobsCOM.UserFieldsMD
                v_UserField = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                v_UserField.TableName = TableName
                v_UserField.Name = Name
                v_UserField.Description = Description
                v_UserField.Type = Type
                If Type <> SAPbobsCOM.BoFieldTypes.db_Date Then
                    If Size <> 0 Then
                        If Type = SAPbobsCOM.BoFieldTypes.db_Numeric Then
                            v_UserField.EditSize = Size
                        Else
                            v_UserField.Size = Size
                        End If
                    End If
                End If
                If SubType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                    v_UserField.SubType = SubType
                End If
                If LinkedTable <> "" Then v_UserField.LinkedTable = LinkedTable

                If Not (ValidValues Is Nothing) Then
                    If ValidValues.GetLength(0) > 0 Then
                        For i As Integer = 0 To ValidValues.GetLength(0) - 1
                            v_UserField.ValidValues.SetCurrentLine(i)
                            v_UserField.ValidValues.Value = ValidValues(i, 0)
                            v_UserField.ValidValues.Description = ValidValues(i, 1)
                            v_UserField.ValidValues.Add()
                        Next
                        If Not (DefaultValues) Is Nothing Then
                            If DefaultValues.Length > 0 Then
                                v_UserField.DefaultValue = DefaultValues(0, 0)
                            Else
                                v_UserField.DefaultValue = ValidValues(1, 0)
                            End If
                        End If
                    End If
                End If
                v_RetVal = v_UserField.Add()
                If v_RetVal <> 0 Then
                    oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                    oApplication.StatusBar.SetText("Failed to add UserField " & Description & " - " & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                    v_UserField = Nothing
                    Return False
                Else
                    oApplication.StatusBar.SetText("[@" & TableName & "] - " & Description & " added successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                    v_UserField = Nothing
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function registerUDO(ByVal UDOCode As String, ByVal UDOName As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType, ByVal findAliasNDescription As String(,), ByVal parentTableName As String, Optional ByVal childTable1 As String = "", Optional ByVal childTable2 As String = "", Optional ByVal childTable3 As String = "", Optional ByVal LogOption As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal DefaultForm As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO) As Boolean
        Dim actionSuccess As Boolean = False
        Try
            registerUDO = False
            Dim v_udoMD As SAPbobsCOM.UserObjectsMD
            v_udoMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            v_udoMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanCreateDefaultForm = DefaultForm
            v_udoMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanLog = LogOption
            v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.Code = UDOCode
            v_udoMD.Name = UDOName
            v_udoMD.TableName = parentTableName
            If LogOption = SAPbobsCOM.BoYesNoEnum.tYES Then
                v_udoMD.LogTableName = "A" & parentTableName
            End If
            v_udoMD.ObjectType = UDOType
            For i As Int16 = 0 To findAliasNDescription.GetLength(0) - 1
                If i > 0 Then
                    v_udoMD.FindColumns.Add()
                    v_udoMD.FormColumns.Add()
                End If

                v_udoMD.FindColumns.ColumnAlias = findAliasNDescription(i, 0)
                v_udoMD.FindColumns.ColumnDescription = findAliasNDescription(i, 1)

                v_udoMD.FormColumns.FormColumnAlias = findAliasNDescription(i, 0)
                v_udoMD.FormColumns.FormColumnDescription = findAliasNDescription(i, 1)
            Next
            If childTable1 <> "" Then
                v_udoMD.ChildTables.TableName = childTable1
                v_udoMD.ChildTables.Add()
            End If
            If childTable2 <> "" Then
                v_udoMD.ChildTables.TableName = childTable2
                v_udoMD.ChildTables.Add()
            End If
            If childTable3 <> "" Then
                v_udoMD.ChildTables.TableName = childTable3
                v_udoMD.ChildTables.Add()
            End If

            If v_udoMD.Add() = 0 Then
                DB_Restart = True
                registerUDO = True
                oApplication.StatusBar.SetText("Successfully Registered UDO >" & UDOCode & ">" & UDOName & " >" & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                oApplication.StatusBar.SetText("Failed to Register UDO >" & UDOCode & ">" & UDOName & " >" & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                registerUDO = False
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD)
            v_udoMD = Nothing
            GC.Collect()
        Catch ex As Exception
        End Try
    End Function


#End Region

    Sub ShowReport(ByVal rptName As String, ByVal SourceXML As String, Optional ByVal Type As String = "", Optional ByVal ShowReport As Boolean = True, Optional ByVal PrintCount As Integer = 1, Optional ByVal v_Display As Boolean = False)
        Try

            Dim oSubReport As CrystalDecisions.CrystalReports.Engine.SubreportObject
            Dim rptSubReportDoc As CrystalDecisions.CrystalReports.Engine.ReportDocument
            Dim rptView As New CrystalDecisions.Windows.Forms.CrystalReportViewer
            Dim rptPath As String = System.Windows.Forms.Application.StartupPath & "\" & rptName
            Dim rptDoc As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            rptDoc.Load(rptPath)
            For Each oMainReportTable As CrystalDecisions.CrystalReports.Engine.Table In rptDoc.Database.Tables
                oMainReportTable.Location = System.IO.Path.GetTempPath() & SourceXML
            Next
            For Each rptSection As CrystalDecisions.CrystalReports.Engine.Section In rptDoc.ReportDefinition.Sections
                For Each rptObject As CrystalDecisions.CrystalReports.Engine.ReportObject In rptSection.ReportObjects
                    If rptObject.Kind = CrystalDecisions.Shared.ReportObjectKind.SubreportObject Then
                        oSubReport = rptObject
                        rptSubReportDoc = oSubReport.OpenSubreport(oSubReport.SubreportName)
                        For Each oSubTable As CrystalDecisions.CrystalReports.Engine.Table In rptSubReportDoc.Database.Tables
                            oSubTable.Location = System.IO.Path.GetTempPath() & SourceXML
                        Next
                    End If
                Next
            Next

            If ShowReport = True Then
                Dim rptForm1 As New Crystal_Form
                rptForm1.CrystalReportViewer1.ReportSource = rptDoc
                rptForm1.ShowDialog()
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
