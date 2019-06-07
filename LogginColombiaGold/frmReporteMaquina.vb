Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports clsDHPlatform
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Configuration
Public Class frmReporteMaquina
    Dim oPlatform As New clsDHPlatform
    Private Sub FillRig()
        Try
            oPlatform.sRfRig = ""
            'cmbToRelocateD.ValueMember = ""
            Dim dtRig As DataTable = oPlatform.getRfRig()
            Dim drC As DataRow = dtRig.NewRow()
            drC(0) = "-1"
            drC(1) = "Select an option.."
            dtRig.Rows.Add(drC)
            dtRgNoDC.DataSource = dtRig
            dtRgNoDC.DisplayMember = "Description"
            dtRgNoDC.ValueMember = "RigID"
            dtRgNoDC.SelectedValue = "-1"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Public Sub llenarCmb()
        FillRig()
    End Sub
    Private Sub frmReporteMaquina_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        llenarCmb()
    End Sub
End Class