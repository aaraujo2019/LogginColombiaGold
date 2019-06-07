Imports System.IO
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports clsDHPlatform
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Configuration
Public Class frmDrillingReport
    Dim oPlatform As New clsDHPlatform
    Private Sub frmDrillingReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnExcel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcel2.Click
        Try
            Dim oXL As Excel.Application
            Dim oWB As Excel._Workbook
            Dim oSheet As Excel._Worksheet
            Dim oRng As Excel.Range

            oXL = New Excel.Application()
            oXL.Visible = True


            oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings("Ruta_DrillProgressMxD").ToString(), 0, False, 5, Type.Missing, Type.Missing, _
             False, Type.Missing, Type.Missing, True, False, Type.Missing, _
             False, False, False)

            oSheet = DirectCast(oWB.Sheets(1), Excel._Worksheet)


            oPlatform.sDateini = dtpDateIni.Value
            oPlatform.sDatefin = dtpDateFin.Value
            Dim DiasMes As String
            ' MsgBox(DateTime.DaysInMonth(Now.Year, Now.Month))
            DiasMes = DateTime.DaysInMonth(Year(dtpDateIni.Value), Month(dtpDateIni.Value))
            ' MsgBox(DiasMes)
            Dim dtPlatformReport As DataTable = oPlatform.getDH_DrillingReportProdGral

            Dim iInicial As Integer = 4
            Dim cInicial As Integer
            Dim cnInicial As Integer

            For i As Integer = 0 To dtPlatformReport.Rows.Count - 1
                oSheet.Cells(iInicial, 1) = dtPlatformReport.Rows(i)("DATE").ToString()
                oSheet.Cells(iInicial, 2) = dtPlatformReport.Rows(i)("To").ToString()
                oSheet.Cells(iInicial, 3) = dtPlatformReport.Rows(i)("Prom").ToString() / DiasMes
                iInicial += 1
            Next

            Dim dtPlatformReportContractor As DataTable = oPlatform.getDH_DrillingReportContractor

            iInicial = 4
            For i As Integer = 0 To dtPlatformReportContractor.Rows.Count - 1

                oSheet.Cells(iInicial, 4) = dtPlatformReportContractor.Rows(i)("CONTRACTOR").ToString()
                oSheet.Cells(iInicial, 5) = dtPlatformReportContractor.Rows(i)("TOTAL").ToString()
                iInicial += 1
            Next
            oSheet = DirectCast(oWB.Sheets(2), Excel._Worksheet)
            Dim dtPlatformReportS2 As DataTable = oPlatform.getDH_DrillingReport
            iInicial = 4
            For i As Integer = 0 To dtPlatformReportS2.Rows.Count - 1

                oSheet.Cells(iInicial, 1) = dtPlatformReportS2.Rows(i)("HoleID").ToString()
                oSheet.Cells(iInicial, 2) = dtPlatformReportS2.Rows(i)("From").ToString()
                oSheet.Cells(iInicial, 3) = dtPlatformReportS2.Rows(i)("To").ToString()
                oSheet.Cells(iInicial, 4) = dtPlatformReportS2.Rows(i)("Machine").ToString()
                oSheet.Cells(iInicial, 5) = dtPlatformReportS2.Rows(i)("EstimatedTargetDrill").ToString()
                oSheet.Cells(iInicial, 6) = dtPlatformReportS2.Rows(i)("Contractor").ToString()
                oSheet.Cells(iInicial, 7) = dtPlatformReportS2.Rows(i)("Date").ToString()
                oSheet.Cells(iInicial, 8) = dtPlatformReportS2.Rows(i)("Executed").ToString()
                oSheet.Cells(iInicial, 9) = dtPlatformReportS2.Rows(i)("EstimatedTargetDrill").ToString() / DiasMes
                iInicial += 1
            Next

            Dim dtPlatformReportProdMac As DataTable = oPlatform.getDH_DrillingReportProduccion
            'Dim view As DataView = New DataView(dtPlatformReportProdMac)
            ''view.Sort
            ''Dim dt As DataTable = view.Table
            'view.Sort() = "Fecha ASC"
            'iInicial = 37

            'Dim dvDatos As New DataView
            'dvDatos = dtPlatformReportProdMac.DefaultView
            'dvDatos.Sort = "Fecha asc"
            'Dim dtg As New DataGridView
            'dtg.DataSource = dvDatos


            'MsgBox(dtg.Columns.Count)
            'MsgBox(dtg.Rows.Count)
            iInicial = 37
            'Dim C As Integer

            For C = 0 To dtPlatformReportProdMac.Columns.Count - 1
                oSheet.Cells(iInicial - 1, 22 + C) = dtPlatformReportProdMac.Columns(C).ColumnName.ToString
                For i As Integer = 0 To dtPlatformReportProdMac.Rows.Count - 1
                    oSheet.Cells(iInicial, 22 + C) = dtPlatformReportProdMac.Rows(i).Item(C).ToString
                    iInicial += 1
                Next
                iInicial = 37
            Next



            iInicial = 4
            For i As Integer = 0 To dtPlatformReportContractor.Rows.Count - 1

                oSheet.Cells(iInicial, 22) = dtPlatformReportContractor.Rows(i)("CONTRACTOR").ToString()
                oSheet.Cells(iInicial, 23) = dtPlatformReportContractor.Rows(i)("MACHINE").ToString()
                oSheet.Cells(iInicial, 24) = dtPlatformReportContractor.Rows(i)("TOTAL").ToString()
                iInicial += 1
            Next

            oSheet = DirectCast(oWB.Sheets(3), Excel._Worksheet)
            Dim dtPlatformReportRig As DataTable = oPlatform.getDH_DrillingReportRig


            If dtPlatformReportRig.Rows.Count > 0 Then
                oSheet.Cells(2, 1) = dtPlatformReportRig.Rows(0)("Contractor").ToString()
                oSheet.Cells(3, 1) = "Date From " & Format(dtpDateIni.Value, "MM" & "/" & "dd" & "/" & "yyyy") & " Date to " & Format(dtpDateFin.Value, "MM" & "/" & "dd" & "/" & "yyyy")
                'oSheet.Cells(5, 2) = dtPlatformReportRig.Rows(0)("Rig").ToString()
                oSheet.Cells(6, 2) = clsRf.sUser
            End If


            iInicial = 12
            For i As Integer = 0 To dtPlatformReportRig.Rows.Count - 1

                oSheet.Cells(iInicial, 1) = dtPlatformReportRig.Rows(i)("Item").ToString()
                oSheet.Cells(iInicial, 2) = dtPlatformReportRig.Rows(i)("ResTimeCont").ToString()
                oSheet.Cells(iInicial, 4) = dtPlatformReportRig.Rows(i)("ResTimeComp").ToString()
                oSheet.Cells(iInicial, 6) = dtPlatformReportRig.Rows(i)("Rig").ToString()
                oSheet.Cells(iInicial, 7) = dtPlatformReportRig.Rows(i)("Contractor").ToString()
                iInicial += 1
            Next

            oSheet = DirectCast(oWB.Sheets(4), Excel._Worksheet)
            Dim dtPlatformReportCompany As DataTable = oPlatform.getDH_DrillingReportCompany


            iInicial = 3
            For i As Integer = 0 To dtPlatformReportCompany.Rows.Count - 1

                oSheet.Cells(iInicial, 1) = dtPlatformReportCompany.Rows(i)("Name").ToString()
                oSheet.Cells(iInicial, 2) = dtPlatformReportCompany.Rows(i)("RigUsed").ToString()
                oSheet.Cells(iInicial, 3) = dtPlatformReportCompany.Rows(i)("HoleID").ToString()
                oSheet.Cells(iInicial, 4) = dtPlatformReportCompany.Rows(i)("Meters").ToString()
                oSheet.Cells(iInicial, 5) = dtPlatformReportCompany.Rows(i)("DateIni").ToString()
                oSheet.Cells(iInicial, 6) = dtPlatformReportCompany.Rows(i)("DateFinish").ToString()
                oSheet.Cells(iInicial, 7) = dtPlatformReportCompany.Rows(i)("Status").ToString()
                iInicial += 1
            Next



            oXL.Visible = True
            oXL.UserControl = True

            MsgBox("Successfully Report", MsgBoxStyle.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try



    End Sub
End Class