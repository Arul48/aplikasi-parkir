Imports System.Data.OleDb

Public Class Form2

    Sub KodeOtomatis()
        CNN = New OleDbConnection(KONEKSI)
        If CNN.State <> ConnectionState.Closed Then CNN.Close()
        CNN.Open()
        OLECMD = New OleDbCommand("select * from Table1 order by [No Antrian] desc", CNN)
        OLERDR = OLECMD.ExecuteReader
        OLERDR.Read()

        If Not OLERDR.HasRows Then
            TextBox2.Text = "000001"
        Else
            TextBox2.Text = Val(Microsoft.VisualBasic.Mid(OLERDR.Item("No Antrian").ToString, 5, 3)) + 1

            If Len(TextBox2.Text) = 1 Then
                TextBox2.Text = "00000" & TextBox2.Text & ""
            ElseIf Len(TextBox2.Text) = 2 Then
                TextBox2.Text = "0000" & TextBox2.Text & ""
            ElseIf Len(TextBox2.Text) = 3 Then
                TextBox2.Text = "000" & TextBox2.Text & ""
            End If
        End If
        TextBox3.Focus()
    End Sub
    Sub bersih()
        TextBox3.Text = ""
        ComboBox1.Text = ""
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Then
            MsgBox("Isi data dengan benar!!", MsgBoxStyle.Exclamation, "Kesalahan")
            Exit Sub
        End If
        CNN = New OleDbConnection(KONEKSI)
        If CNN.State <> ConnectionState.Closed Then CNN.Close()
        CNN.Open()
        OLECMD = New OleDbCommand("insert into Table1 ([No Antrian],[No Plat Polisi],[Jenis Kendaraan],[Jam Masuk],[Operator],[Tanggal]) values ('" & _
                                  TextBox2.Text & "','" & TextBox3.Text & "','" & ComboBox1.Text & "','" & TextBox4.Text & "','" & TextBox6.Text & "','" & TextBox1.Text & "')", CNN)

        x = OLECMD.ExecuteNonQuery
        If x = 1 Then
            MsgBox("Data berhasil di simpan", MsgBoxStyle.Information, "Informasi")
            Call bersih()
            Call KodeOtomatis()
        Else
            MsgBox("Gagal menyimpan data", MsgBoxStyle.Exclamation, "Kesalahan")
        End If
    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call KodeOtomatis()
        TextBox4.Text = TimeOfDay
        TextBox1.Text = Format(Now, "dd/MM/yyyy")
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("Motor")
        ComboBox1.Items.Add("Mobil")
        TextBox5.Text = 4000
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        TextBox4.Text = Format(Now, "HH:mm:ss")
        TextBox1.Text = Format(Now, "dd/MM/yyyy")

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call bersih()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Close()
    End Sub
End Class