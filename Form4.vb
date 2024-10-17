Imports System.Data.OleDb
Public Class Form4
    Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source= '221903058.mdb' ")
    Dim cmd As OleDbCommand
    Dim da As OleDbDataAdapter
    Dim dt As DataTable
    Dim sql As String
    Dim maxrow As Integer

    Private Sub load_cbo(sql As String, cbo As ComboBox)
        Try
            con.Open()
            cmd = New OleDb.OleDbCommand
            da = New OleDb.OleDbDataAdapter
            dt = New DataTable

            With cmd
                .Connection = con
                .CommandText = sql

            End With
            da.SelectCommand = cmd
            da.Fill(dt)

            cbo.DataSource = dt
            cbo.DisplayMember = dt.Columns(0).ToString




        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            con.Close()
            da.Dispose()
        End Try
    End Sub

    Private Sub load_ComboBox1(sql As String, ComboBox1 As ComboBox)
        Try
            con.Open()
            cmd = New OleDb.OleDbCommand
            da = New OleDb.OleDbDataAdapter
            dt = New DataTable

            With cmd
                .Connection = con
                .CommandText = sql

            End With
            da.SelectCommand = cmd
            da.Fill(dt)

            ComboBox1.DataSource = dt
            ComboBox1.DisplayMember = dt.Columns(1).ToString




        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            con.Close()
            da.Dispose()
        End Try
    End Sub
    Private Sub load_ComboBox2(sql As String, ComboBox2 As ComboBox)
        Try
            con.Open()
            cmd = New OleDb.OleDbCommand
            da = New OleDb.OleDbDataAdapter
            dt = New DataTable

            With cmd
                .Connection = con
                .CommandText = sql

            End With
            da.SelectCommand = cmd
            da.Fill(dt)

            ComboBox2.DataSource = dt
            ComboBox2.DisplayMember = dt.Columns(4).ToString




        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            con.Close()
            da.Dispose()
        End Try
    End Sub


    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load



        Listele("SELECT * FROM  rezervasyon'")

        sql = "Select * From musteri"

        load_cbo(sql, cbo)
        load_ComboBox1(sql, ComboBox1)
        load_ComboBox2(sql, ComboBox2)
        ComboBox3.Items.Add("FROZEN")
        ComboBox3.Items.Add("MİCKEY MOUSE")
    End Sub
    Private Sub Listele(ByVal SQL As String)

        Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='221903058.mdb'")

        Dim rezervasyon As New DataTable("rezervasyon")

        Dim adapter As New OleDbDataAdapter(SQL, baglanti)

        adapter.Fill(rezervasyon)



    End Sub
    Private Sub Temizle()

        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox7.Clear()
        TextBox6.Clear()


    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim durum As String

        durum = MsgBox("mus_adı = " & TextBox1.Text & vbNewLine & "mus_soyadı = " & TextBox2.Text & vbNewLine & "rezervasyon = " & TextBox3.Text & vbNewLine & " mus_telefonu = " & TextBox4.Text & vbNewLine & " paket_adı= " & TextBox5.Text & vbNewLine & "personel = " & TextBox6.Text & vbNewLine & "ücret = " & TextBox7.Text & vbNewLine & "Yukarıdaki yazdığınız veriler kayıt edilsinmi '", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "Kayıt Uyarı")

        If durum = vbYes Then

            Dim sql As New String("INSERT INTO rezervasyon (mus_adı,mus_soyadı,rezervasyon,mus_telefonu,paket_adı,personel,ücret) values ('{0}','{1}', '{2} ','{3}','{4}','{5}','{6}')")

            sql = String.Format(sql, TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text, TextBox6.Text, TextBox7.Text)

            Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='221903058.mdb'")

            Dim komutnesnesi As New OleDb.OleDbCommand(sql, baglanti)

            Dim sonuc As Integer

            baglanti.Open()

            sonuc = komutnesnesi.ExecuteNonQuery()

            If sonuc = 1 Then

                MsgBox("Yandaki Girdiğiniz Veriler Veri Tabanına Kayıt Olmuştur.", MsgBoxStyle.Exclamation, "Kayıt Uyarı")

            End If


            Listele("SELECT * FROM rezervasyon '")

            baglanti.Close()

            Temizle()


        Else
        End If
    End Sub

    Private Sub cbo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo.SelectedIndexChanged
        TextBox1.Text = cbo.Text
    End Sub



    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox2.Text = ComboBox1.Text
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        TextBox3.Text = DateTimePicker1.Text
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        TextBox4.Text = ComboBox2.Text
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        TextBox5.Text = ComboBox3.Text

    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox7.Text = Val(TextBox7.Text) + 120
        Else
            TextBox7.Text = Val(TextBox7.Text) - 120
        End If

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox7.Text = Val(TextBox7.Text) + 80
        Else
            TextBox7.Text = Val(TextBox7.Text) - 80
        End If
    End Sub
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            TextBox7.Text = Val(TextBox7.Text) + 60
        Else
            TextBox7.Text = Val(TextBox7.Text) - 60
        End If


    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            TextBox7.Text = Val(TextBox7.Text) + 30
        Else
            TextBox7.Text = Val(TextBox7.Text) - 30
        End If

    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            TextBox7.Text = Val(TextBox7.Text) + 40
        Else
            TextBox7.Text = Val(TextBox7.Text) - 40
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            TextBox7.Text = Val(TextBox7.Text) + 40
        Else
            TextBox7.Text = Val(TextBox7.Text) - 40
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            TextBox7.Text = Val(TextBox7.Text) + 10
        Else
            TextBox7.Text = Val(TextBox7.Text) - 10
        End If
    End Sub
    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.Checked = True Then
            TextBox7.Text = Val(TextBox7.Text) + 120
        Else
            TextBox7.Text = Val(TextBox7.Text) - 120
        End If

    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        If CheckBox9.Checked = True Then
            TextBox7.Text = Val(TextBox7.Text) + 50
        Else
            TextBox7.Text = Val(TextBox7.Text) - 50
        End If
    End Sub
    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked = True Then
            TextBox7.Text = Val(TextBox7.Text) + 20
        Else
            TextBox7.Text = Val(TextBox7.Text) - 20
        End If


    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        If CheckBox11.Checked = True Then
            TextBox7.Text = Val(TextBox7.Text) + 30
        Else
            TextBox7.Text = Val(TextBox7.Text) - 30
        End If

    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        If CheckBox12.Checked = True Then
            TextBox7.Text = Val(TextBox7.Text) + 50
        Else
            TextBox7.Text = Val(TextBox7.Text) - 50
        End If
    End Sub

    Private Sub CheckBox13_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox13.CheckedChanged
        If CheckBox13.Checked = True Then
            TextBox7.Text = Val(TextBox7.Text) + 40
        Else
            TextBox7.Text = Val(TextBox7.Text) - 40
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form2.Show()
        Me.Hide()

    End Sub
End Class
