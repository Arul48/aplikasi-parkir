Imports System.Data.OleDb

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        CNN = New OleDbConnection(KONEKSI)
        If CNN.State <> ConnectionState.Closed Then CNN.Close()
        CNN.Open()
        OLECMD = New OleDbCommand("SELECT * From login  WHERE username = '" & TextBox1.Text & _
                                   "' and password = '" & TextBox2.Text & "'", CNN)
        OLERDR = OLECMD.ExecuteReader
        If (OLERDR.Read()) Then
            Form5.Show()
            Me.Hide()
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox1.Focus()
        Else
            MsgBox("Username & Password Anda Salah!", MsgBoxStyle.OkOnly, _
                   "Login gagal")
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox1.Focus()
        End If

    End Sub
End Class
