Imports System.Data.Odbc
Public Class FormMasterAdmin

    Sub KondisiAwal()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        ComboBox1.Items.Clear()
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        ComboBox1.Enabled = False

        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button1.Text = "Input"
        Button2.Text = "Edit"
        Button3.Text = "Hapus"
        Button4.Text = "Tutup"
        Call koneksi()
        da = New OdbcDataAdapter("select kodeadmin,namaadmin,leveladmin from tb_admin", con)
        ds = New DataSet
        da.Fill(ds, "tb_admin")
        DataGridView1.DataSource = ds.Tables("tb_admin")
        DataGridView1.ReadOnly = True
    End Sub

    Sub SiapIsi()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        ComboBox1.Enabled = True
        ComboBox1.Items.Add("ADMIN")
        ComboBox1.Items.Add("USER")
    End Sub

    Private Sub FormMasterAdmin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call KondisiAwal()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Button1.Text = "Input" Then
            Button1.Text = "Simpan"
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Text = "Batal"
            Call SiapIsi()
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
                MsgBox("silahkan isi")
            Else
                Call koneksi()
                Dim InputData As String = "insert into tb_admin values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & ComboBox1.Text & "')"
                cmd = New OdbcCommand(InputData, con)
                cmd.ExecuteNonQuery()
                MsgBox("input berhasil")
                Call KondisiAwal()

            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Button2.Text = "Edit" Then
            Button2.Text = "Simpan"
            Button1.Enabled = False
            Button3.Enabled = False
            Button4.Text = "Batal"
            Call SiapIsi()
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
                MsgBox("silahkan isi")
            Else
                Call koneksi()
                Dim UpdateData As String = "Update tb_admin set namaadmin='" & TextBox2.Text & "',passwordadmin='" & TextBox3.Text & "',leveladmin='" & ComboBox1.Text & "' where kodeadmin='" & TextBox1.Text & "'"
                cmd = New OdbcCommand(UpdateData, con)
                cmd.ExecuteNonQuery()
                MsgBox("update berhasil")
                Call KondisiAwal()

            End If
        End If
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call koneksi()
            cmd = New OdbcCommand("select * from tb_admin where kodeadmin='" & TextBox1.Text & "'", con)
            rd = cmd.ExecuteReader
            rd.Read()
            If Not rd.HasRows Then
                MsgBox("kode tidak ada")
            Else
                TextBox1.Text = rd.Item("kodeadmin")
                TextBox2.Text = rd.Item("namaadmin")
                TextBox3.Text = rd.Item("passwordadmin")
                ComboBox1.Text = rd.Item("leveladmin")
            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If Button4.Text = "Tutup" Then
            Me.Close()
        Else
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Button3.Text = "Hapus" Then
            Button3.Text = "Delete"
            Button1.Enabled = False
            Button2.Enabled = False
            Button4.Text = "Batal"
            Call SiapIsi()
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
                MsgBox("silahkan isi")
            Else
                Call koneksi()
                Dim HapusData As String = "delete from tb_admin where kodeadmin='" & TextBox1.Text & "'"
                cmd = New OdbcCommand(HapusData, con)
                cmd.ExecuteNonQuery()
                MsgBox("delete berhasil")
                Call KondisiAwal()

            End If
        End If
    End Sub
End Class