Imports System.Data.OleDb

Public Class Form4
    Sub OpenDB()
        Dim KONEKSI = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\DB Zery 1.accdb"
        CNN = New OleDbConnection(KONEKSI)
        If CNN.State <> ConnectionState.Closed Then
            CNN.Open()
        End If
    End Sub
    Sub TampilTable1()
        OLEDA = New OleDbDataAdapter("select * from Table1", CNN)
        DS = New DataSet
        OLEDA.Fill(DS, "Table1")
        DataGridView1.DataSource = DS.Tables("Table1")
    End Sub
    Sub Hitung()
        Dim h As Integer
        h = DataGridView1.Rows.Count - 1
        TextBox2.Text = h

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call OpenDB()
        Call TampilTable1()
        Call Hitung()

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        CNN = New OleDbConnection(KONEKSI)
        If CNN.State <> ConnectionState.Closed Then CNN.Close()
        CNN.Open()
        OLECMD = New OleDbCommand("select * from Table1 where [No Antrian] like'%" & TextBox1.Text & "%'", CNN)
        OLERDR = OLECMD.ExecuteReader
        OLERDR.Read()

        If OLERDR.HasRows Then
            OLEDA = New OleDbDataAdapter("select * from Table1 where [No Antrian] like'%" & TextBox1.Text & "%'", CNN)
            DS = New DataSet
            OLEDA.Fill(DS, "Ketemu")
            DataGridView1.DataSource = DS.Tables("Ketemu")
            DataGridView1.ReadOnly = True
        End If
    End Sub
End Class