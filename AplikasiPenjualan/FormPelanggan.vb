Imports System.Data.Odbc
Public Class FormPelanggan
    Sub KondisiAwal()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False

        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button1.Text = "Input"
        Button2.Text = "Edit"
        Button3.Text = "Hapus"
        Button4.Text = "Tutup"
        Call koneksi()
        da = New OdbcDataAdapter("select kodepelanggan,namapelanggan,alamatpelanggan,telppelanggan from tb_pelanggan", con)
        ds = New DataSet
        da.Fill(ds, "tb_pelanggan")
        DataGridView1.DataSource = ds.Tables("tb_pelanggan")
        DataGridView1.ReadOnly = True
    End Sub

    Sub SiapIsi()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
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
            Call NO_otomatis()
            TextBox1.Enabled = False
            TextBox2.Focus()
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("silahkan isi")
            Else
                Call koneksi()
                Dim InputData As String = "insert into tb_pelanggan values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
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
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("silahkan isi")
            Else
                Call koneksi()
                Dim UpdateData As String = "Update tb_pelanggan set namapelanggan='" & TextBox2.Text & "',alamatpelanggan='" & TextBox3.Text & "',telppelanggan='" & TextBox4.Text & "' where kodepelanggan='" & TextBox1.Text & "'"
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
            cmd = New OdbcCommand("select * from tb_pelanggan where kodepelanggan='" & TextBox1.Text & "'", con)
            rd = cmd.ExecuteReader
            rd.Read()
            If Not rd.HasRows Then
                MsgBox("kode tidak ada")
            Else
                TextBox1.Text = rd.Item("kodepelanggan")
                TextBox2.Text = rd.Item("namapelanggan")
                TextBox3.Text = rd.Item("alamatpelanggan")
                TextBox4.Text = rd.Item("telppelanggan")
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
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("silahkan isi")
            Else
                Call koneksi()
                Dim HapusData As String = "delete from tb_pelanggan where kodepelanggan='" & TextBox1.Text & "'"
                cmd = New OdbcCommand(HapusData, con)
                cmd.ExecuteNonQuery()
                MsgBox("delete berhasil")
                Call KondisiAwal()

            End If
        End If
    End Sub
    Sub NO_otomatis()
        Call koneksi()
        cmd = New OdbcCommand("select * from tb_pelanggan where kodepelanggan in(select max(kodepelanggan) from tb_pelanggan)", con)
        Dim Urutankode As String
        Dim Hitung As Long
        Rd = Cmd.ExecuteReader
        Rd.Read()
        If Not Rd.HasRows Then
            Urutankode = "PLG" + "001"
        Else
            Hitung = Microsoft.VisualBasic.Right(Rd.GetString(0), 3) + 1
            Urutankode = "PLG" + Microsoft.VisualBasic.Right("000" & Hitung, 3)
        End If
        TextBox1.Text = Urutankode
    End Sub
End Class