Imports System.Data.OleDb
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim con As New OleDbConnection
        Dim com As New OleDbCommand
        Dim adp As New OleDbDataAdapter
        Dim dataset As New DataSet
        con.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0; Data Source= 221903058.mdb")
        com.CommandText = ("Select * From yonetici Where y_kadi= ' " + TextBox1.Text + " ' AND y_sifre= ' " + TextBox2.Text + " ' ; ")
        con.Open()
        com.Connection = con
        adp.SelectCommand = com
        adp.Fill(dataset, "0")
        Dim a = dataset.Tables(0).Rows.Count
        If a = 0 Then
            MsgBox("Doğru Giriş !")
            Form2.Show()
            Me.Hide()
        Else
            MsgBox("Lütfen Doğru Giriş Yapınız!", MsgBoxStyle.Critical)
            TextBox1.Clear()
            TextBox2.Clear()

        End If

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class