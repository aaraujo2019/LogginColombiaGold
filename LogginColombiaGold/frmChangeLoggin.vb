Imports System.Security.Cryptography
Imports System.Text
Public Class frmChangeLoggin
    Dim oRf As New clsRf()
    Private Sub btnAccept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAccept.Click
        Try
            'Valida que la contraseña nueva sea la correcta
            If txtNewPass.Text.ToString() <> txtRepPass.Text.ToString() Then
                MessageBox.Show("Different New Password and Repeat Password", "Logging", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                Return
            End If

            Dim sPwd As String = GetSHA1(txtNewPass.Text.ToString())
            Dim sPwdOld As String = GetSHA1(txtOldPass.Text.ToString())

            Dim sResp As String = oRf.UpdatePass(sPwdOld.ToString(), sPwd.ToString(), clsRf.sUser.ToString())


            MessageBox.Show(sResp)
        Catch ex As Exception
            MessageBox.Show("Error: " + ex.Message)
        End Try
    End Sub
    Public Shared Function GetSHA1(ByVal texto As [String]) As String
        Try
            Dim sha1 As SHA1 = SHA1CryptoServiceProvider.Create()
            Dim textOriginal As [Byte]() = ASCIIEncoding.[Default].GetBytes(texto)
            Dim hash As [Byte]() = sha1.ComputeHash(textOriginal)
            Dim cadena As New StringBuilder()
            For Each i As Byte In hash
                cadena.AppendFormat("{0:x2}", i)
            Next
            Return cadena.ToString()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class