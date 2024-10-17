Imports System.Data.OleDb
Imports System.Data.SqlClient
Public Class Form9
    Private Sub Form9_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        DataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically

        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Listele("SELECT * FROM  rezervasyon'")
    End Sub
    Private Sub Listele(ByVal SQL As String)

        Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='221903058.mdb'")

        Dim rezervasyon As New DataTable("rezervasyon")

        Dim adapter As New OleDbDataAdapter(SQL, baglanti)

        adapter.Fill(rezervasyon)

        DataGridView1.DataSource = rezervasyon


    End Sub





    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        PrintDocument1.Print()
    End Sub
    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim bm As New Bitmap(Me.DataGridView1.Width, Me.DataGridView1.Height)
        DataGridView1.DrawToBitmap(bm, New Rectangle(0, 0, Me.DataGridView1.Width, Me.DataGridView1.Height))
        e.Graphics.DrawImage(bm, 0, 0)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form2.Show()
        Me.Hide()

    End Sub
End Class