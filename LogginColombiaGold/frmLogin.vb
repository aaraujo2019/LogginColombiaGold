Imports System.Security.Cryptography
Imports System.Text
Imports System.Configuration
Public Class frmLogin
    Dim oRf As New clsRf
    Dim swVersion As Boolean

    Private Sub frmLogin_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If swVersion = True Then
            'MsgBox("Final")
            'Shell("C:\ColombiaGold\Drilling\Actualizar.bat", AppWinStyle.MaximizedFocus)
            Shell("Actualizar.bat", AppWinStyle.MaximizedFocus)
        End If

    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Formato As String
        Dim SeparadorDecimal As String
        Dim SeparadorMiles As String
        Dim SeparadorLista As String
        Dim sValida As Boolean
        Dim Mensaje As String
        sValida = True
        Formato = System.Globalization.CultureInfo.CurrentCulture.ToString
        SeparadorDecimal = System.Globalization.NumberFormatInfo.CurrentInfo.CurrencyDecimalSeparator
        SeparadorMiles = System.Globalization.NumberFormatInfo.CurrentInfo.CurrencyGroupSeparator

        'If Formato <> "en-US" Then
        '    sValida = False
        '    Mensaje = "- El formato debe ser : Inglés (Estados Unidos)"
        '    Mensaje = Mensaje & vbNewLine & "- Su formato actual es :" & Formato
        'End If
        'If SeparadorDecimal <> "." Then
        '    sValida = False
        '    Mensaje = Mensaje & vbNewLine & ""
        '    Mensaje = Mensaje & vbNewLine & "- El separador decimal debe ser : Punto (.)"
        '    Mensaje = Mensaje & vbNewLine & "- Su separador decimal actual es : " & SeparadorDecimal
        'End If

        'If SeparadorMiles <> " " Then
        '    sValida = False
        '    Mensaje = Mensaje & vbNewLine & ""
        '    Mensaje = Mensaje & vbNewLine & "- El separador de miles debe ser : Espacio ( )"
        '    Mensaje = Mensaje & vbNewLine & "- Su separador de miles actual es : " & SeparadorMiles
        'End If
        'If sValida = False Then
        '    Mensaje = Mensaje & vbNewLine & ""
        '    MsgBox(Mensaje & vbNewLine & "Debe corregir la configuracion antes de Continuar", MsgBoxStyle.Critical, "Error")
        '    End
        'End If


        'Solicita la actualización de la versión del proyecto. AAA
        'Try
        '    swVersion = False
        '    'ConfigurationSettings.AppSettings["IDProject"].ToString()
        '    oRf.iIdProject = Integer.Parse(ConfigurationSettings.AppSettings("IDProject").ToString())
        '    Dim dtVers As DataTable = oRf.getVersionProject()
        '    If Double.Parse(dtVers.Rows(0)("version").ToString()) > Double.Parse(ConfigurationSettings.AppSettings("Version").ToString()) Then
        '        swVersion = True

        '        MsgBox("Hay una nueva actualización en el sistema, se realizará la actualización automáticamente, despues de ello podrá ingresar al sistema", MsgBoxStyle.Information, "Actualización")
        '        'MsgBox(Application.StartupPath & Application.ExecutablePath)
        '        Me.Close()
        '        Kill(Application.ExecutablePath)

        '    End If
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        'End Try
    End Sub

    Private Sub btnCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelar.Click
        Close()
    End Sub

    Private Sub btnAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAceptar.Click
        Dim sPwd As String = GetSHA1(txtPwd.Text.ToString())
        Dim dtUser As New DataTable()
        dtUser = oRf.getUsersPortal(txtUser.Text.ToString())
        Try
            If dtUser.Rows.Count > 0 Then
                If Boolean.Parse(dtUser.Rows(0)("activo_User").ToString()) = False Then
                    MessageBox.Show("Disabled User", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                If dtUser.Rows(0)("login_User").ToString().ToUpper() = txtUser.Text.ToString().ToUpper And dtUser.Rows(0)("passwd_User").ToString().ToUpper() = sPwd.ToString().ToUpper() Then
                    clsRf.sUser = txtUser.Text.ToString()
                    clsRf.sIdentification = dtUser.Rows(0)("id_User").ToString()
                    clsRf.sIdGrupo = dtUser.Rows(0)("idGrupo_User").ToString()
                    Dim oSplash As New frmSplash()
                    oSplash.Show()
                    Me.Hide()
                Else
                    MessageBox.Show("Credentials failed", "Drilling", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                MessageBox.Show("Credentials failed", "Drilling", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            Throw New Exception("Error: " & ex.Message)
        End Try
    End Sub
    Public Shared Function GetSHA1(ByVal texto As String) As String
        Try
            Dim sha1 As SHA1 = SHA1CryptoServiceProvider.Create()
            Dim textOriginal As Byte() = ASCIIEncoding.Default.GetBytes(texto)
            Dim hash As Byte() = sha1.ComputeHash(textOriginal)
            Dim cadena As New StringBuilder()
            For Each i As Byte In hash
                cadena.AppendFormat("{0:x2}", i)

            Next
            Return cadena.ToString()

        Catch ex As Exception
            Throw New Exception(ex.Message)

        End Try


    End Function


    Private Sub groupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles groupBox1.Enter

    End Sub

    Private Sub txtPwd_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPwd.KeyDown
        If e.KeyValue = Keys.Enter Then
            'SendKeys.Send("{tab}")
            Dim sPwd As String = GetSHA1(txtPwd.Text.ToString())
            Dim dtUser As New DataTable()
            dtUser = oRf.getUsersPortal(txtUser.Text.ToString())
            Try
                If dtUser.Rows.Count > 0 Then

                    If Boolean.Parse(dtUser.Rows(0)("activo_User").ToString()) = False Then
                        MessageBox.Show("Disabled User", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Return

                    End If

                    If dtUser.Rows(0)("login_User").ToString().ToUpper() = txtUser.Text.ToString().ToUpper And dtUser.Rows(0)("passwd_User").ToString().ToUpper() = sPwd.ToString().ToUpper() Then

                        clsRf.sUser = txtUser.Text.ToString()
                        'clsRf.sIdentification = dtRfWorker.Rows[0]["Identification"].ToString();
                        clsRf.sIdentification = dtUser.Rows(0)("id_User").ToString()
                        clsRf.sIdGrupo = dtUser.Rows(0)("idGrupo_User").ToString()
                        'FrmPpal oPpal = new FrmPpal();
                        'oPpal.Show();
                        Dim oSplash As New frmSplash()
                        oSplash.Show()
                        Me.Hide()
                        'this.Dispose();

                    Else
                        MessageBox.Show("Credentials failed", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If

                Else
                    MessageBox.Show("Credentials failed", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error)

                End If

            Catch ex As Exception
                Throw New Exception("Error: " & ex.Message)

            End Try
        End If
    End Sub

    Private Sub txtPwd_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPwd.KeyPress
        'Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
        'KeyAscii = CShort(SoloNumeros(KeyAscii))
        'If KeyAscii = 0 Then
        '    'e.Handled = True

        'End If
    End Sub
    Function SoloNumeros(ByVal Keyascii As Short) As Short
        If InStr("1234567890.-,", Chr(Keyascii)) = 0 Then
            SoloNumeros = 0
        Else
            SoloNumeros = Keyascii
        End If
        Select Case Keyascii
            Case 8
                SoloNumeros = Keyascii
            Case 13
                SoloNumeros = Keyascii
        End Select
    End Function
  
End Class
