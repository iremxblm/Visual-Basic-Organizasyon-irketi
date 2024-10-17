Imports System.Data.OleDb
Imports System.Data.SqlClient
Public Class Form10
    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        DataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically

        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        DataGridView2.EditMode = DataGridViewEditMode.EditProgrammatically

        DataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Listele("SELECT * FROM  frozenteklif'")
        Listele2("SELECT * FROM mickeyteklif")
    End Sub
    Private Sub Listele(ByVal SQL As String)

        Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='221903058.mdb'")

        Dim frozenteklif As New DataTable("frozenteklif")

        Dim adapter As New OleDbDataAdapter(SQL, baglanti)

        adapter.Fill(frozenteklif)

        DataGridView1.DataSource = frozenteklif



    End Sub
    Private Sub Listele2(ByVal SQL As String)

        Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='221903058.mdb'")

        Dim frozenteklif As New DataTable("mickeyteklif")

        Dim adapter As New OleDbDataAdapter(SQL, baglanti)

        adapter.Fill(frozenteklif)

        DataGridView2.DataSource = frozenteklif



    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub ANASAYFaToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form2.Show()
        Me.Hide()



    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PrintDocument1.Print()

    End Sub
    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs)

        Dim bm As New Bitmap(Me.DataGridView1.Width, Me.DataGridView1.Height)
        DataGridView1.DrawToBitmap(bm, New Rectangle(0, 0, Me.DataGridView1.Width, Me.DataGridView1.Height))
        e.Graphics.DrawImage(bm, 0, 0)

    End Sub

    Private Sub ANASAYFAToolStripMenuItem_Click_1(sender As Object, e As EventArgs)
        Form2.Show()
        Me.Hide()

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PrintDocument2.Print()

    End Sub
    Private Sub MenuStrip1_ItemClicked_1(sender As Object, e As ToolStripItemClickedEventArgs)

    End Sub
    Private Sub PrintDocument2_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs)

        Dim bm As New Bitmap(Me.DataGridView2.Width, Me.DataGridView2.Height)
        DataGridView2.DrawToBitmap(bm, New Rectangle(0, 0, Me.DataGridView2.Width, Me.DataGridView2.Height))
        e.Graphics.DrawImage(bm, 0, 0)

    End Sub
End Class
