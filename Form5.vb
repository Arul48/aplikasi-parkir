Public Class Form5

    Private Sub ParkirMasukToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ParkirMasukToolStripMenuItem.Click
        Dim anak As New Form2
        anak.MdiParent = Me
        anak.Show()
    End Sub

    Private Sub ParkirKeluarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ParkirKeluarToolStripMenuItem.Click
        Dim anak As New Form3
        anak.MdiParent = Me
        anak.Show()

    End Sub

    Private Sub LihatDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LihatDataToolStripMenuItem.Click
        Dim anak As New Form4
        anak.MdiParent = Me
        anak.Show()

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label1.Text = Format(Now, "HH:mm:ss")
        Label2.Text = Format(Now, "dd/MM/yyyy")

    End Sub

    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = Format(Now, "HH:mm:ss")
        Label2.Text = Format(Now, "dd/MM/yyyy")

    End Sub

    Private Sub KeluarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KeluarToolStripMenuItem.Click
        End
    End Sub
End Class