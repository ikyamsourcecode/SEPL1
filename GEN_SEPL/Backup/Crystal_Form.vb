Imports GEN_SEPL.ClsSubContract  ''''AAAAA
Imports GEN_SEPL.ClsCustomBOM
Imports GEN_SEPL.ClsPreShipment
'Imports GEN_CRYSTAL.ClsPurchaseOrder
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO


Public Class Crystal_Form

    Protected WithEvents oCrystalViewer As CrystalDecisions.Windows.Forms.CrystalReportViewer

    Private Sub CrystalReportViewer1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles CrystalReportViewer1.Load
        Try
            Dim oPassword As String
            Dim oServerName As String
            Dim oUserName As String
            Dim oCrystalReport As New ReportDocument
            Dim oSubReport As New ReportDocument
            Dim myTable As CrystalDecisions.CrystalReports.Engine.Table
            Dim myLogin As CrystalDecisions.Shared.TableLogOnInfo
            Dim pFields As ParameterFields
            Dim pField As ParameterField
            Dim pFieldValue As ParameterDiscreteValue
            Dim pSubFields As ParameterFields
            Dim pSubField As ParameterField
            Dim pSubFieldValue As ParameterDiscreteValue
            Dim sFileReader As StreamReader

            'GIVE THE NAME OF THE RPT FILE. RPT FILE TO BE PLACED WITHIN THE PROJECT /BIN/DEBUG FOLDER
            oCrystalReport.Load(sRptName)

            ' DB.TXT CONTAINS THE PASSWORD FOR THE SQL SERVER DATABASE. TO BE STORED IN BIN/DEBUG FOLDER
            sFileReader = File.OpenText("DBLogin.ini")
            pFields = New ParameterFields
            pField = New ParameterField
            pFieldValue = New ParameterDiscreteValue
            pSubFields = New ParameterFields
            pSubField = New ParameterField
            pSubFieldValue = New ParameterDiscreteValue
            oServerName = sFileReader.ReadLine
            oUserName = sFileReader.ReadLine
            oPassword = sFileReader.ReadLine

            For Each myTable In oCrystalReport.Database.Tables
                myLogin = myTable.LogOnInfo
                myLogin.ConnectionInfo.DatabaseName = oCompany.CompanyDB
                myLogin.ConnectionInfo.ServerName = oServerName
                myLogin.ConnectionInfo.UserID = oUserName
                myLogin.ConnectionInfo.Password = oPassword
                myTable.ApplyLogOnInfo(myLogin)
            Next
            For Each oSubReport In oCrystalReport.Subreports
                For Each myTable In oSubReport.Database.Tables
                    myLogin = myTable.LogOnInfo
                    myLogin.ConnectionInfo.DatabaseName = oCompany.CompanyDB
                    myLogin.ConnectionInfo.ServerName = oServerName
                    myLogin.ConnectionInfo.UserID = oUserName
                    myLogin.ConnectionInfo.Password = oPassword
                    myTable.ApplyLogOnInfo(myLogin)
                Next
            Next
            oCrystalReport.SetParameterValue("@DocNum", sDocNum)
            CrystalReportViewer1.ReportSource = oCrystalReport
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Protected Overrides Sub Finalize()

        MyBase.Finalize()
        GC.Collect()

    End Sub

End Class
