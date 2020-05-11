Public Class form1
    Dim sqlnya As String
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        If CheckBox1.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox2.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox3.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox4.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox5.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox6.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox7.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox8.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox9.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox10.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox11.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox12.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox13.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox14.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox15.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox16.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox17.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox18.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox19.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox20.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        If CheckBox21.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        CheckBox3.Enabled = False
        CheckBox4.Enabled = False
        CheckBox5.Enabled = False
        CheckBox6.Enabled = False
        CheckBox7.Enabled = False
        CheckBox8.Enabled = False
        CheckBox9.Enabled = False
        CheckBox10.Enabled = False
        CheckBox11.Enabled = False
        CheckBox12.Enabled = False
        CheckBox13.Enabled = False
        CheckBox14.Enabled = False
        CheckBox15.Enabled = False
        CheckBox16.Enabled = False
        CheckBox17.Enabled = False
        CheckBox18.Enabled = False
        CheckBox19.Enabled = False
        CheckBox20.Enabled = False
        CheckBox21.Enabled = False

    End Sub

    Sub panggildata()
        konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM pts", conn)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "pts")
        DataGridView1.DataSource = DS.Tables("pts")
        DataGridView1.Enabled = True
    End Sub
    Sub jalan()
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        objcmd.Connection = conn
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = sqlnya
        objcmd.ExecuteNonQuery()
        objcmd.Dispose()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        RichTextBox1.Text = ""
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call panggildata()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        sqlnya = "insert into pts (Nama,Umur,TTL,Point) values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & RichTextBox1.Text & "')"
        Call jalan()
        MsgBox("Data Berhasil Tersimpan")
        Call panggildata()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        sqlnya = "UPDATE pts set Nama = '" & TextBox1.Text & "',Point = '" & RichTextBox1.Text & "',TTL = '" & TextBox3.Text & "'.'"
        Call jalan()
        MsgBox("Data Berhasil Terubah")
        Call panggildata()
    End Sub
    Private Sub DataGridView1_RowHeaderMouseClick(sender As Object, e As EventArgs) Handles DataGridView1.RowHeaderMouseClick
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        TextBox1.Text = DataGridView1.Item(0, i).Value
        TextBox2.Text = DataGridView1.Item(1, i).Value
        TextBox3.Text = DataGridView1.Item(2, i).Value
        RichTextBox1.Text = DataGridView1.Item(3, i).Value
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        sqlnya = "delete from pts where Nama = '" & TextBox1.Text & "'"
        Call jalan()
        MsgBox("Data Berhasil Dihapus")
        Call panggildata()
    End Sub
End Class