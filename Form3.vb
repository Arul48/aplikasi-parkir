Imports System.Data.OleDb

Public Class Form3
    Sub OpenDB()
        Dim KONEKSI = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\DB Zery 1.accdb"
        CNN = New OleDbConnection(KONEKSI)
        If CNN.State <> ConnectionState.Closed Then

        End If
        CNN.Open()
    End Sub
    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        CNN = New OleDbConnection(KONEKSI)
        If CNN.State <> ConnectionState.Closed Then CNN.Close()
        CNN.Open()
        OLECMD = New OleDbCommand("select * from Table1 where [No Antrian] = '" & TextBox2.Text & "'", CNN)
        OLERDR = OLECMD.ExecuteReader
        OLERDR.Read()

        If OLERDR.HasRows Then
            TextBox3.Text = OLERDR("No Plat Polisi")
            TextBox4.Text = OLERDR("Jenis Kendaraan")
            TextBox5.Text = OLERDR("Jam Masuk")
            TextBox1.Text = OLERDR("Tanggal")
            TextBox10.Text = OLERDR("Operator")
            TextBox6.Text = TimeOfDay
            TextBox8.Text = "4000"
        Else
            MsgBox("Data Tidak Ditemukan….!", MsgBoxStyle.Exclamation, "Perhatian")
        End If
    End Sub

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox8.Text = "Rp4000"
        TextBox2.Focus()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("Silahkan Cari Data!!", MsgBoxStyle.Exclamation, "Kesalahan")
            Exit Sub
        End If
        Dim detik, menit, jam, second As Integer
        detik = DateDiff(DateInterval.Second, CDate(TextBox5.Text), CDate(TextBox6.Text))
        menit = DateDiff(DateInterval.Minute, CDate(TextBox5.Text), CDate(TextBox6.Text))
        jam = DateDiff(DateInterval.Hour, CDate(TextBox5.Text), CDate(TextBox6.Text))
        jam = detik / 3600
        menit = (detik Mod 3600) / 60
        second = (detik Mod 3600) Mod 60

        TextBox7.Text = jam & ":" & menit & ":" & second
        If jam > 3 Then
            TextBox9.Text = 8000
            If jam > 6 Then
                TextBox9.Text = 16000
            End If
        Else
            TextBox9.Text = 4000
        End If
        Label12.Text = TextBox3.Text
        Label13.Text = TextBox7.Text
        Label14.Text = TextBox8.Text

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox2.Text = "" Then
            MsgBox("Isi data dengan benar!!", MsgBoxStyle.Exclamation, "Kesalahan")
            Exit Sub
        End If
        CNN = New OleDbConnection(KONEKSI)
        If CNN.State <> ConnectionState.Closed Then CNN.Close()
        CNN.Open()
        OLECMD = New OleDbCommand("Update Table1 Set [Jam Keluar]='" & TextBox6.Text & "' ,[Lama Parkir]='" & TextBox7.Text & "' ,[Biaya]='" & TextBox9.Text & "' Where [No Antrian]='" & TextBox2.Text & "'", CNN)
        x = OLECMD.ExecuteNonQuery
        If x = 1 Then
            MsgBox("Data berhasil di Simpan", MsgBoxStyle.Information, "Informasi")
            Call bersih()
            TextBox2.Focus()
        Else
            MsgBox("Gagal Menyimpan data", MsgBoxStyle.Exclamation, "Kesalahan")
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox8.Text = "" Or TextBox10.Text = "" Then
            MsgBox("Isi data dengan benar!!", MsgBoxStyle.Exclamation, "Kesalahan")
            Exit Sub
        End If

        If MsgBox("Ingin menghapus data ?", MsgBoxStyle.YesNo, "Konfirmasi") = MsgBoxResult.Yes Then
            CNN = New OleDbConnection(KONEKSI)
            If CNN.State <> ConnectionState.Closed Then CNN.Close()
            CNN.Open()
            OLECMD = New OleDbCommand("Delete from Table1 Where [No Antrian] ='" & TextBox2.Text & "'", CNN)
            x = OLECMD.ExecuteNonQuery
        End If
        If x = 1 Then
            MsgBox("Data berhasil di hapus", MsgBoxStyle.Information, "Informasi")
            Call bersih()
            TextBox2.Focus()
        Else
            MsgBox("Gagal menghapus data", MsgBoxStyle.Exclamation, "Kesalahan")
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call bersih()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Close()
    End Sub
End Class