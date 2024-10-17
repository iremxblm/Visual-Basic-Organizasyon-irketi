Imports System.Data.OleDb
Public Class Form8
    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox1.Text = Val(TextBox1.Text) + 120
            TextBox2.Text = "+"
        Else
            TextBox1.Text = Val(TextBox1.Text) - 120
            TextBox2.Text = ""
        End If

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox1.Text = Val(TextBox1.Text) + 50
            TextBox3.Text = "+"
        Else
            TextBox1.Text = Val(TextBox1.Text) - 50
            TextBox3.Text = ""
        End If
    End Sub
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            TextBox1.Text = Val(TextBox1.Text) + 20
            TextBox4.Text = "+"
        Else
            TextBox1.Text = Val(TextBox1.Text) - 20
            TextBox4.Text = ""
        End If


    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            TextBox1.Text = Val(TextBox1.Text) + 30
            TextBox5.Text = "+"
        Else
            TextBox1.Text = Val(TextBox1.Text) - 30
            TextBox5.Text = ""
        End If

    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            TextBox1.Text = Val(TextBox1.Text) + 50
            TextBox6.Text = "+"
        Else
            TextBox1.Text = Val(TextBox1.Text) - 50
            TextBox6.Text = ""
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            TextBox1.Text = Val(TextBox1.Text) + 40
            TextBox7.Text = "+"
        Else
            TextBox1.Text = Val(TextBox1.Text) - 40
            TextBox7.Text = ""
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form2.Show()
        Me.Hide()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form4.Show()
        Me.Hide()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim sql2 As New String("INSERT INTO mickeyteklif (pasta,pecete,davetiye,masaortusu,roadset,pinyata,ücret) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')")

        sql2 = String.Format(sql2, TextBox2.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text, TextBox6.Text, TextBox7.Text, TextBox1.Text)

        Dim baglanti2 As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='221903058.mdb'")

        Dim komutnesnesi2 As New OleDb.OleDbCommand(sql2, baglanti2)

        Dim sonuc2 As Integer

        baglanti2.Open()

        sonuc2 = komutnesnesi2.ExecuteNonQuery()

        If sonuc2 = 1 Then

            MsgBox("Yandaki Girdiğiniz Veriler Teklif Formuna Kayıt Olmuştur.", MsgBoxStyle.Exclamation, "Kayıt Uyarı")

        End If

    End Sub
End Class