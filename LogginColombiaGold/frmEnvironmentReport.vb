Imports System.IO
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports clsDHPlatform
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Configuration
Public Class frmEnvironmentReport
    Dim oPlatform As New clsDHPlatform
    Public Shared Platform As String
    Private Sub frmEnvironmentReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            txtYear.Value = Year(Date.Now).ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub btnExcel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcel2.Click
        Try
            Dim oXL As Excel.Application
            Dim oWB As Excel._Workbook
            Dim oSheet As Excel._Worksheet
            Dim oRng As Excel.Range

            oXL = New Excel.Application()
            oXL.Visible = True


            oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings("Ruta_Environment").ToString(), 0, False, 5, Type.Missing, Type.Missing, _
             False, Type.Missing, Type.Missing, True, False, Type.Missing, _
             False, False, False)

            oSheet = DirectCast(oWB.Sheets(1), Excel._Worksheet)


            oPlatform.sDateini = txtYear.Value
            oPlatform.sPlatform = Platform

            Dim dtEnvironmentReport As DataTable = oPlatform.getDH_EnvironmentReportProdGral

            Dim iInicial As Integer = 4
            For i As Integer = 0 To dtEnvironmentReport.Rows.Count - 1

                oSheet.Cells(iInicial, 1) = dtEnvironmentReport.Rows(i)("QUESTIONSUBGROUP").ToString()
                oSheet.Cells(iInicial, 2) = dtEnvironmentReport.Rows(i)("PERIODO1").ToString()
                oSheet.Cells(iInicial, 3) = dtEnvironmentReport.Rows(i)("PERIODO2").ToString()
                oSheet.Cells(iInicial, 4) = dtEnvironmentReport.Rows(i)("PERIODO3").ToString()
                oSheet.Cells(iInicial, 5) = dtEnvironmentReport.Rows(i)("PERIODO4").ToString()

                iInicial += 1
            Next
            iInicial = 23
            Dim dtEnvironmentReportGroup As DataTable = oPlatform.getDH_EnvironmentReportProdGralGroup

            For i As Integer = 0 To dtEnvironmentReportGroup.Rows.Count - 1

                oSheet.Cells(iInicial, 1) = dtEnvironmentReportGroup.Rows(i)("QUESTIONGROUP").ToString()
                oSheet.Cells(iInicial, 2) = dtEnvironmentReportGroup.Rows(i)("PERIODO1").ToString()
                oSheet.Cells(iInicial, 3) = dtEnvironmentReportGroup.Rows(i)("PERIODO2").ToString()
                oSheet.Cells(iInicial, 4) = dtEnvironmentReportGroup.Rows(i)("PERIODO3").ToString()
                oSheet.Cells(iInicial, 5) = dtEnvironmentReportGroup.Rows(i)("PERIODO4").ToString()

                iInicial += 1
            Next

            iInicial = 39
            Dim dtEnvironmentReportImpact As DataTable = oPlatform.getDH_EnvironmentReportProdGralImpact

            For i As Integer = 0 To dtEnvironmentReportImpact.Rows.Count - 1

                oSheet.Cells(iInicial, 1) = dtEnvironmentReportImpact.Rows(i)("Impact").ToString()
                oSheet.Cells(iInicial, 2) = dtEnvironmentReportImpact.Rows(i)("Leve").ToString()
                oSheet.Cells(iInicial, 3) = dtEnvironmentReportImpact.Rows(i)("Severo").ToString()
                oSheet.Cells(iInicial, 4) = dtEnvironmentReportImpact.Rows(i)("Critico").ToString()

                iInicial += 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try



    End Sub
End Class