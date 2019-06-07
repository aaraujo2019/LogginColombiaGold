Imports System.IO
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports Microsoft.Office.Interop
Imports clsDHSurvey
Imports System.Configuration

Public Class frmSurvey
    Dim Accion As String
    Dim dt As New DataTable
    Dim Hole As String
    Dim oSurvey As New clsDHSurvey
    Dim MinValSurvey As String
    Dim sEditDr As String = "0"
    Dim sInsert As Boolean
    Dim sPermiteImport As Boolean
    Dim EOH As String


    Private Sub Validar()
        Try
            'Validar Que no exista
            sPermiteImport = True
            Dim fila As Integer


            If dg_Excel.Rows.Count > 0 Then

                oSurvey.sHoleID = dg_Excel.Rows(0).Cells(0).Value.ToString
                Dim dtSurvey As DataTable = oSurvey.getDHSurvey_ID

                fila = dg_Excel.Rows.Count - 1
                'MsgBox(fila)
                'MsgBox(dtSurvey.Rows(0)("EOH").ToString)
                'MsgBox(dg_Excel.Rows(fila).Cells(1).Value.ToString)
                If dtSurvey.Rows.Count > 0 Then
                    dg_Validacion.Rows.Add("Hold: " & dg_Excel.Rows(0).Cells(0).Value.ToString & " Exist in database, may continue and update the data")
                    sPermiteImport = True
                End If

                For i = 0 To dg_Excel.Rows.Count - 1
                    If dg_Excel.Rows(i).Cells(2).Value < 0 Or dg_Excel.Rows(i).Cells(2).Value > 360 Then
                        dg_Validacion.Rows.Add("Error: Azimuth range (0 to 360) " & dg_Excel.Rows(i).Cells(2).Value.ToString & ", can not continue ")
                        sPermiteImport = False
                    End If
                Next
                For i = 0 To dg_Excel.Rows.Count - 1
                    If dg_Excel.Rows(i).Cells(2).Value < -90 Or dg_Excel.Rows(i).Cells(3).Value > 90 Then
                        dg_Validacion.Rows.Add("Error: Dip range range (-90 and 90) " & dg_Excel.Rows(i).Cells(3).Value.ToString & ", can not continue ")
                        sPermiteImport = False
                    End If
                Next


                If dtSurvey.Rows.Count > 0 Then
                    If dg_Excel.Rows(fila).Cells(1).Value.ToString > dtSurvey.Rows(0)("EOH").ToString Then
                        dg_Validacion.Rows.Add("Error: Final Depth > EOH, can not continue ")
                        sPermiteImport = False
                    End If
                End If


                If dg_Excel.Rows(0).Cells(1).Value.ToString <> 0 Then
                    dg_Validacion.Rows.Add("Error: Inictial Depth <> 0, can not continue")
                    sPermiteImport = False
                End If

                MinValSurvey = ConfigurationSettings.AppSettings("MinValSurvey").ToString
                For i = 0 To dg_Excel.Rows.Count - 1
                    If i > 0 Then
                        ' MsgBox(dg_Excel.Rows(i).Cells(3).Value)
                        ' MsgBox(dg_Excel.Rows(i - 1).Cells(3).Value)
                        'MsgBox(dg_Excel.Rows(i).Cells(3).Value - dg_Excel.Rows(i - 1).Cells(3).Value)
                        If (Val(dg_Excel.Rows(i - 1).Cells(3).Value) - Val(dg_Excel.Rows(i).Cells(3).Value)) > MinValSurvey Then
                            dg_Validacion.Rows.Add("Hold: " & dg_Excel.Rows(i).Cells(0).Value.ToString & ", Depth: " & dg_Excel.Rows(i).Cells(1).Value.ToString & " , Azimuth: " & dg_Excel.Rows(i).Cells(2).Value.ToString & " , exceeds the maximum accepted in Dip(" & MinValSurvey & "), Warning!")
                            'sPermiteImport = False

                        End If
                    End If
                Next



                MsgBox("Check the tab Validarion Errors ", MsgBoxStyle.Information, "Information")



            End If


            

        Catch ex As Exception
            MsgBox("Error : " + ex.Message, MsgBoxStyle.Information, "Information")
        End Try

    End Sub

    Private Sub btn_AbrirExcel__Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_AbrirExcel_.Click
        Try
            Dim openFD As New OpenFileDialog()
            dg_Excel.DataSource = ""
            With openFD
                .Title = "Seleccionar archivos"
                .Filter = "Todos los archivos (*.xlsx)|*.xlsx|Todos los archivos(*.xls)|*.xls"
                .Multiselect = False
                .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    Txt_PathExcel_.Text = .FileName
                    'btn_ImportarExcelaGrid.Enabled = True
                End If
            End With

            Dim objExcel As New Excel.Application
            Dim hoja As Excel.Worksheet

            objExcel.Workbooks.Open(Txt_PathExcel_.Text)
            For Each hoja In objExcel.Sheets
                cmb_Validacion.Items.Add(hoja.Name)
            Next
            objExcel.Workbooks(1).Close()

            cmb_Validacion.Enabled = True
        Catch ex As Exception

        End Try


    End Sub

    Private Sub btn_ImportarExcelaGrid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ImportarExcelaGrid.Click
        Try
            Dim strconn As String
            strconn = "Provider=Microsoft.ACE.OLEDB.12.0; data source= " + Txt_PathExcel_.Text + ";Extended properties=""Excel 8.0;hdr=yes;imex=1"""
            Dim mconn As New System.Data.OleDb.OleDbConnection(strconn)
            Dim ad As New System.Data.OleDb.OleDbDataAdapter("Select * from [" & Accion & "$]", mconn)
            mconn.Open()
            ad.Fill(dt)
            mconn.Close()
            Me.dg_Excel.DataSource = dt
            btn_ValidarNulos.Enabled = True

            'AutoNumberRowsForGridView(dg_Excel)



        Catch ex As System.Data.OleDb.OleDbException
            MessageBox.Show(ex.Message)

        End Try
        'cmb_Validacion.Enabled = False
        btn_ValidarNulos.Enabled = True
        'btn_Importar_db.Enabled = True
    End Sub
    Public Sub AutoNumberRowsForGridView(ByVal dataGridView As DataGridView)
        Try
            'If dataGridView IsNot Nothing Then
            '    Return
            'End If
            Dim count As Integer = 1
            For Each row As DataGridViewRow In dataGridView.Rows
                row.HeaderCell.Value = String.Format("{0:0}", count)
                count += 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub
    Private Sub cmb_Validacion_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_Validacion.SelectedIndexChanged
        Try
            Accion = cmb_Validacion.SelectedItem.ToString
            btn_ImportarExcelaGrid.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub btn_ValidarNulos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ValidarNulos.Click
        Try
            dg_Validacion.Rows.Clear()
            Validar()
        Catch ex As Exception
            MsgBox("Error: " + ex.Message, MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Sub tabSurvey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tabSurvey.Click
        Try
            If Me.tabSurvey.SelectedIndex = 1 Then
                If sPermiteImport = True Then
                    btn_Importar_db.Enabled = True
                Else
                    btn_Importar_db.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub btn_Importar_db_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Importar_db.Click
        Try

            sInsert = False
            oSurvey.sHoleID = dg_Excel.Rows(0).Cells(0).Value.ToString
            oSurvey.DHSurveyDel()
            Dim sResps As String = ""

            For i = 0 To dg_Excel.Rows.Count - 1
                If sEditDr = "0" Then
                    oSurvey.sOpcion = "1"
                    oSurvey.sHoleID = dg_Excel.Rows(i).Cells(0).Value.ToString
                    oSurvey.sDepth = Decimal.Parse(dg_Excel.Rows(i).Cells(1).Value.ToString)
                    oSurvey.sAz = Decimal.Parse(dg_Excel.Rows(i).Cells(2).Value.ToString)
                    oSurvey.sDip = Decimal.Parse(dg_Excel.Rows(i).Cells(3).Value.ToString)
                    oSurvey.sMeaseredBy = dg_Excel.Rows(i).Cells(4).Value.ToString
                    oSurvey.sInstrument = dg_Excel.Rows(i).Cells(5).Value.ToString
                    oSurvey.sMethod = dg_Excel.Rows(i).Cells(6).Value.ToString
                    oSurvey.sTemp = dg_Excel.Rows(i).Cells(7).Value.ToString
                    oSurvey.sMagField = dg_Excel.Rows(i).Cells(8).Value.ToString
                    oSurvey.sGravFieald = dg_Excel.Rows(i).Cells(9).Value.ToString
                    oSurvey.sObservation = dg_Excel.Rows(i).Cells(10).Value.ToString
                    If dg_Excel.Rows(i).Cells(11).Value.ToString = "" Then
                        oSurvey.sDate = "01/01/1900"
                    Else
                        oSurvey.sDate = dg_Excel.Rows(i).Cells(11).Value.ToString
                    End If

                    sResps = oSurvey.DH_Survey_Add()
                    If sResps = "OK" Then
                        'MsgBox("Survey Insert.", MsgBoxStyle.Information)
                        'EOLG. 20121009. Para almacenar en el log de transacciones
                        Dim oRf As New clsRf
                        oRf.InsertTrans("DH_Survey", "Insert", clsRf.sUser.ToString(), _
                        "HoleId : " + dg_Excel.Rows(i).Cells(0).Value.ToString + ". " + _
                        "Depth: " + dg_Excel.Rows(i).Cells(1).Value.ToString + ". " + _
                        "Azimuth: " + dg_Excel.Rows(i).Cells(2).Value.ToString + ". " + _
                        "Dip: " + dg_Excel.Rows(i).Cells(3).Value.ToString + ". " + _
                        "Observation: " + dg_Excel.Rows(i).Cells(10).Value.ToString + ". " + _
                        "Event Date " + Date.Now())

                        sInsert = True
                    End If
                    sEditDr = "0"
                End If
            Next

            If sInsert = True Then
                MsgBox("Survey Insert.", MsgBoxStyle.Information)
            Else
                MsgBox("Survey Error: " + sResps, MsgBoxStyle.Information)
            End If
            'Me.Close()

        Catch ex As Exception
            MsgBox("Survey Error: " + ex.Message, MsgBoxStyle.Information)
        End Try
        
    End Sub
End Class