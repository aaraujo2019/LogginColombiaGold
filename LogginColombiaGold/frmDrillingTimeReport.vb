Imports System.IO
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports clsDHPlatform
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Configuration
Public Class frmDrillingTimeReport
    Dim oPlatform As New clsDHPlatform
    Private Sub frmDrillingTimeReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnExcel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcel2.Click
        Try
            Dim oXL As Excel.Application
            Dim oWB As Excel._Workbook
            Dim oSheet As Excel._Worksheet
            Dim oRng As Excel.Range

            oXL = New Excel.Application()
            oXL.Visible = True


            oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings("Ruta_DrillTimeReport").ToString(), 0, False, 5, Type.Missing, Type.Missing, _
             False, Type.Missing, Type.Missing, True, False, Type.Missing, _
             False, False, False)

            oSheet = DirectCast(oWB.Sheets(1), Excel._Worksheet)

            oPlatform.sDateini = dtpDateIni.Value
            oPlatform.sDatefin = dtpDateFin.Value

            Dim dtPlatformReportDrillingTime As DataSet = oPlatform.getDH_DrillingTime
            oSheet.Cells(2, 1) = "Date From " & Format(dtpDateIni.Value, "MM" & "/" & "dd" & "/" & "yyyy") & " Date to " & Format(dtpDateFin.Value, "MM" & "/" & "dd" & "/" & "yyyy")
            oSheet.Cells(5, 2) = clsRf.sUser

            Dim iInicial As Integer = 9
            For i As Integer = 0 To dtPlatformReportDrillingTime.Tables(0).Rows.Count - 1

                oSheet.Cells(iInicial, 1) = dtPlatformReportDrillingTime.Tables(0).Rows(i)("Date").ToString()
                oSheet.Cells(iInicial, 2) = dtPlatformReportDrillingTime.Tables(0).Rows(i)("Turn").ToString()
                oSheet.Cells(iInicial, 3) = dtPlatformReportDrillingTime.Tables(0).Rows(i)("Rig").ToString()
                oSheet.Cells(iInicial, 4) = dtPlatformReportDrillingTime.Tables(0).Rows(i)("Contractor").ToString()
                oSheet.Cells(iInicial, 5) = dtPlatformReportDrillingTime.Tables(0).Rows(i)("Item").ToString()
                oSheet.Cells(iInicial, 6) = dtPlatformReportDrillingTime.Tables(0).Rows(i)("TimeReportDrill").ToString()
                oSheet.Cells(iInicial, 7) = dtPlatformReportDrillingTime.Tables(0).Rows(i)("TimeApprovedInter").ToString()
                oSheet.Cells(iInicial, 8) = dtPlatformReportDrillingTime.Tables(0).Rows(i)("ResTimecont").ToString()
                oSheet.Cells(iInicial, 9) = dtPlatformReportDrillingTime.Tables(0).Rows(i)("ResTimecomp").ToString()
                iInicial += 1
            Next

            'LostTools
            iInicial = 9
            For i As Integer = 0 To dtPlatformReportDrillingTime.Tables(1).Rows.Count - 1

                oSheet.Cells(iInicial, 11) = dtPlatformReportDrillingTime.Tables(1).Rows(i)("Description").ToString()
                oSheet.Cells(iInicial, 12) = dtPlatformReportDrillingTime.Tables(1).Rows(i)("Amount").ToString()
                oSheet.Cells(iInicial, 13) = dtPlatformReportDrillingTime.Tables(1).Rows(i)("PercentPay").ToString()
                oSheet.Cells(iInicial, 14) = dtPlatformReportDrillingTime.Tables(1).Rows(i)("PercentPayAdmon").ToString()
                iInicial += 1
            Next


            'Billable Additives
            iInicial = 9
            For i As Integer = 0 To dtPlatformReportDrillingTime.Tables(2).Rows.Count - 1
                oSheet.Cells(iInicial, 16) = dtPlatformReportDrillingTime.Tables(2).Rows(i)("Description").ToString()
                oSheet.Cells(iInicial, 17) = dtPlatformReportDrillingTime.Tables(2).Rows(i)("Amount").ToString()
                iInicial += 1
            Next




            oXL.Visible = True
            oXL.UserControl = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
End Class