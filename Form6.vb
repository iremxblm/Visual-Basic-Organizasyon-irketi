Imports System.Data.OleDb
Public Class Form6

    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically

        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Listele("SELECT * FROM  personel'")
    End Sub

    Private Sub Listele(ByVal SQL As String)

        Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='221903058.mdb'")

        Dim personel As New DataTable("personel")

        Dim adapter As New OleDbDataAdapter(SQL, baglanti)

        adapter.Fill(personel)

        DataGridView1.DataSource = personel

    End Sub

    Private Sub Temizle()

        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim durum As String

        durum = MsgBox("pers_tcno = " & TextBox1.Text & vbNewLine & "pers_ad= " & TextBox2.Text & vbNewLine & "pers_soyad = " & TextBox3.Text & vbNewLine & "pers_dt = " & DateTimePicker1.Text & vbNewLine & " pers_maas= " & TextBox4.Text & vbNewLine & "pers_gorevı = " & TextBox5.Text & vbNewLine & "Yukarıdaki yazdığınız veriler kayıt edilsinmi '", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "Kayıt Uyarı")

        If durum = vbYes Then

            Dim sql As New String("INSERT INTO personel (pers_tcno,pers_ad,pers_soyad,pers_dt,pers_maas,pers_gorevı) values ('{0}','{1}','{2} ','{3}','{4}','{5}')")

            sql = String.Format(sql, TextBox1.Text, TextBox2.Text, TextBox3.Text, DateTimePicker1.Text, TextBox4.Text, TextBox5.Text)

            Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='221903058.mdb'")

            Dim komutnesnesi As New OleDb.OleDbCommand(sql, baglanti)

            Dim sonuc As Integer

            baglanti.Open()

            sonuc = komutnesnesi.ExecuteNonQuery()

            If sonuc = 1 Then

                MsgBox("Yandaki Girdiğiniz Veriler Veri Tabanına Kayıt Olmuştur.", MsgBoxStyle.Exclamation, "Kayıt Uyarı")

            End If


            Listele("SELECT * FROM personel'")

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

        TextBox5.Text = DataGridView1.CurrentRow.Cells(5).Value

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        Dim sql As New String("UPDATE personel SET pers_tcno='{0}',pers_ad='{1}',pers_soyad='{2}',pers_dt='{3}',pers_maas='{4}',pers_gorevı='{5}' WHERE pers_tcno='{6}' AND pers_ad='{7}' AND pers_soyad='{8}' AND pers_dt='{9}' AND pers_maas = '{10}' AND  pers_gorevı='{11}' ")

        sql = String.Format(sql, TextBox1.Text, TextBox2.Text, TextBox3.Text, DateTimePicker1.Text, TextBox4.Text, TextBox5.Text, DataGridView1.CurrentRow.Cells(0).Value, DataGridView1.CurrentRow.Cells(1).Value, DataGridView1.CurrentRow.Cells(2).Value, DataGridView1.CurrentRow.Cells(3).Value, DataGridView1.CurrentRow.Cells(4).Value, DataGridView1.CurrentRow.Cells(5).Value)

        Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='221903058.mdb'")

        Dim komutnesnesi As New OleDb.OleDbCommand(sql, baglanti)

        Dim sonuc As Integer

        baglanti.Open()

        sonuc = komutnesnesi.ExecuteNonQuery()

        If sonuc = 1 Then

            MsgBox("Değiştirmiş Olduğunuz Veriler Güncellenmiştir.", MsgBoxStyle.Exclamation, "Kayıt Uyarı")



        End If

        Listele("SELECT * FROM personel'")

        baglanti.Close()

        Temizle()

    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim sql As New String("DELETE FROM personel WHERE pers_tcno='{0}' AND pers_ad='{1}' AND pers_soyad='{2}' AND pers_dt='{3}' AND pers_maas = '{4}' AND pers_gorevı='{5}'")

        sql = String.Format(sql, TextBox1.Text, TextBox2.Text, TextBox3.Text, DateTimePicker1.Text, TextBox4.Text, TextBox5.Text)


        Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='221903058.mdb'")

        Dim komutnesnesi As New OleDb.OleDbCommand(sql, baglanti)

        Dim sonuc As Integer

        baglanti.Open()

        sonuc = komutnesnesi.ExecuteNonQuery()

        If sonuc = 1 Then

            MsgBox("Listeden Seçmiş Olduğunuz Veri Silinmiştir.", MsgBoxStyle.Exclamation, "Kayıt Uyarı")



        End If

        Listele("SELECT * FROM personel'")

        baglanti.Close()

        Temizle()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form2.Show()
        Me.Hide()

    End Sub
End Class