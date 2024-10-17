Imports System.Data.OleDb
Public Class Form3
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        DataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically

        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Listele("SELECT * FROM  musteri'")


    End Sub
    Private Sub Listele(ByVal SQL As String)

        Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='221903058.mdb'")

        Dim musteri As New DataTable("musteri")

        Dim adapter As New OleDbDataAdapter(SQL, baglanti)

        adapter.Fill(musteri)

        DataGridView1.DataSource = musteri

    End Sub
    Private Sub Temizle()

        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim durum As String

        durum = MsgBox("mus_ad = " & TextBox1.Text & vbNewLine & "mus_soyad = " & TextBox2.Text & vbNewLine & "mus_adres = " & TextBox3.Text & vbNewLine & " mus_dogumtarıhı = " & DateTimePicker1.Text & vbNewLine & " mus_telno= " & TextBox4.Text & vbNewLine & "Yukarıdaki yazdığınız veriler kayıt edilsinmi '", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "Kayıt Uyarı")

        If durum = vbYes Then

            Dim sql As New String("INSERT INTO musteri (mus_ad,mus_soyad,mus_adres,mus_dogumtarıhı,mus_telno) values ('{0}','{1}', '{2} ','{3}','{4}')")

            sql = String.Format(sql, TextBox1.Text, TextBox2.Text, TextBox3.Text, DateTimePicker1.Text, TextBox4.Text)

            Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='221903058.mdb'")

            Dim komutnesnesi As New OleDb.OleDbCommand(sql, baglanti)

            Dim sonuc As Integer

            baglanti.Open()

            sonuc = komutnesnesi.ExecuteNonQuery()

            If sonuc = 1 Then

                MsgBox("Yandaki Girdiğiniz Veriler Veri Tabanına Kayıt Olmuştur.", MsgBoxStyle.Exclamation, "Kayıt Uyarı")

            End If


            Listele("SELECT * FROM musteri'")

            baglanti.Close()

            Temizle()


        Else
        End If

    End Sub
    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        TextBox1.Text = DataGridView1.CurrentRow.Cells(0).Value

        TextBox2.Text = DataGridView1.CurrentRow.Cells(1).Value

        TextBox3.Text = DataGridView1.CurrentRow.Cells(2).Value

        DateTimePicker1.Text = DataGridView1.CurrentRow.Cells(3).Value

        TextBox4.Text = DataGridView1.CurrentRow.Cells(4).Value

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        Dim sql As New String("UPDATE musteri SET mus_ad='{0}',mus_soyad='{1}',mus_adres='{2}',mus_dogumtarıhı='{3}',mus_telno='{4}' WHERE mus_ad='{5}' AND mus_soyad='{6}' AND mus_adres='{7}' AND mus_dogumtarıhı='{8}' AND mus_telno = '{9}'")

        sql = String.Format(sql, TextBox1.Text, TextBox2.Text, TextBox3.Text, DateTimePicker1.Text, TextBox4.Text, DataGridView1.CurrentRow.Cells(0).Value, DataGridView1.CurrentRow.Cells(1).Value, DataGridView1.CurrentRow.Cells(2).Value, DataGridView1.CurrentRow.Cells(3).Value, DataGridView1.CurrentRow.Cells(4).Value)

        Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='221903058.mdb'")

        Dim komutnesnesi As New OleDb.OleDbCommand(sql, baglanti)

        Dim sonuc As Integer

        baglanti.Open()

        sonuc = komutnesnesi.ExecuteNonQuery()

        If sonuc = 1 Then

            MsgBox("Değiştirmiş Olduğunuz Veriler Güncellenmiştir.", MsgBoxStyle.Exclamation, "Kayıt Uyarı")



        End If

        Listele("SELECT * FROM musteri'")

        baglanti.Close()

        Temizle()

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim sql As New String("DELETE FROM musteri WHERE mus_ad='{0}' AND mus_soyad='{1}' AND mus_adres='{2}' AND mus_dogumtarıhı='{3}' AND mus_telno = '{4}' ")

        sql = String.Format(sql, TextBox1.Text, TextBox2.Text, TextBox3.Text, DateTimePicker1.Text, TextBox4.Text)


        Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='221903058.mdb'")

        Dim komutnesnesi As New OleDb.OleDbCommand(sql, baglanti)

        Dim sonuc As Integer

        baglanti.Open()

        sonuc = komutnesnesi.ExecuteNonQuery()

        If sonuc = 1 Then

            MsgBox("Listeden Seçmiş Olduğunuz Veri Silinmiştir.", MsgBoxStyle.Exclamation, "Kayıt Uyarı")



        End If

        Listele("SELECT * FROM musteri'")

        baglanti.Close()

        Temizle()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form2.Show()
        Me.Hide()


    End Sub
End Class