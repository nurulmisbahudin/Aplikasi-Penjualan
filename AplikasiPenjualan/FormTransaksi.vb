Imports System.Data.Odbc
Public Class FormTransaksi
    Dim Tglmysql As String
    Sub KondisiAwal()
        LBLnama.Text = ""
        LBLalamat.Text = ""
        LBLtelp.Text = ""
        LBLtanggal.Text = Today
        LBLkembali.Text = ""
        LBLadmin.Text = FormMenuUtama.Stlabel4.Text
        Button4.Enabled = False
        TextBox2.Text = ""
        LBLnamabrg.Text = ""
        LBLhrg.Text = ""
        TextBox3.Enabled = False
        LBLitem.Text = ""
        Label9.Text = "0"
        TextBox1.Text = ""
        Call munculkodepelanggan()
        Call NO_otomatis()
        Call Buatkolom()

    End Sub


    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As EventArgs) Handles Timer1.Tick
        LBLjam.Text = TimeOfDay
    End Sub

    Private Sub Form_Transaksi_Penjualan_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Call KondisiAwal()
        Call munculkodepelanggan()
        Call NO_otomatis()
        Call Buatkolom()
    End Sub
    Sub munculkodepelanggan()
        Call koneksi()
        ComboBox1.Items.Clear()
        cmd = New OdbcCommand("select * from tb_pelanggan", con)
        rd = cmd.ExecuteReader
        Do While rd.Read
            ComboBox1.Items.Add(rd.Item(0))
        Loop
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call koneksi()
        cmd = New OdbcCommand("select * from tb_pelanggan where kodepelanggan ='" & ComboBox1.Text & "'", con)
        rd = cmd.ExecuteReader
        rd.Read()
        If rd.HasRows Then
            LBLnama.Text = rd!namapelanggan
            LBLalamat.Text = rd!alamatpelanggan
            LBLtelp.Text = rd!telppelanggan
        End If
    End Sub
    Sub NO_otomatis()
        Call koneksi()
        cmd = New OdbcCommand("select * from tb_jual where Nojual in(select max(Nojual) from tb_jual)", con)
        Dim Urutankode As String
        Dim Hitung As Long
        rd = cmd.ExecuteReader
        rd.Read()
        If Not rd.HasRows Then
            Urutankode = "J" + Format(Now, "yyMMdd") + "001"
        Else
            Hitung = Microsoft.VisualBasic.Right(rd.GetString(0), 9) + 1
            Urutankode = "J" + Format(Now, "yyMMdd") + Microsoft.VisualBasic.Right("000" & Hitung, 3)
        End If
        LBLnojual.Text = Urutankode
    End Sub
    Sub Buatkolom()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("kode", "kode")
        DataGridView1.Columns.Add("nama", "nama barang")
        DataGridView1.Columns.Add("harga", "harga")
        DataGridView1.Columns.Add("jumlah", "jumlah")
        DataGridView1.Columns.Add("subtotal", "subtotal")

    End Sub
    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Chr(13) Then
            Call koneksi()
            cmd = New OdbcCommand("Select * From tb_barang where kodebarang='" & TextBox2.Text & "'", con)
            rd = cmd.ExecuteReader
            rd.Read()
            If Not rd.HasRows Then
                MsgBox("kode barang tidak ada")
            Else
                TextBox2.Text = rd.Item("kodebarang")
                LBLnamabrg.Text = rd.Item("namabarang")
                LBLhrg.Text = rd.Item("hargabarang")
                Label18.Text = rd.Item("jumlahbarang")
                Button4.Enabled = True
                TextBox3.Enabled = True
                TextBox3.Focus()
            End If
        End If
    End Sub
    Private Sub Button4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button4.Click
        If LBLnamabrg.Text = "" Or TextBox3.Text = "" Then
            MsgBox("Silahkan masukan kode Dan jumlah barang")
        Else
            If Val(Label18.Text) < Val(TextBox3.Text) Then
                MsgBox("Barang kurang!")
            Else
                DataGridView1.Rows.Add(New String() {TextBox2.Text, LBLnamabrg.Text, LBLhrg.Text, TextBox3.Text, Val(LBLhrg.Text) * Val(TextBox3.Text)})
                Call subtotal()
                Call keuntungan()
                TextBox2.Text = ""
                LBLnamabrg.Text = ""
                LBLhrg.Text = ""
                TextBox3.Text = ""
                TextBox3.Enabled = False
                Call item()
            End If
        End If
    End Sub
    Sub subtotal()
        Dim hitung As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            hitung = hitung + DataGridView1.Rows(i).Cells(4).Value
            Label9.Text = hitung
        Next
    End Sub
    Sub keuntungan()
        Dim untung As String
        untung = Val(Label9.Text) * Val(100) / Val(20)
        Label19.Text = untung


    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            If Val(TextBox1.Text) < Val(Label9.Text) Then
                MsgBox("Pembayaran Kurang !")
            ElseIf Val(TextBox1.Text) = Val(Label9.Text) Then
                LBLkembali.Text = 0
            ElseIf Val(TextBox1.Text) > Val(Label9.Text) Then
                LBLkembali.Text = Val(TextBox1.Text) - Val(Label9.Text)
                Button1.Focus()
            End If
        End If

    End Sub
    Sub item()
        Dim hitungitem As Integer
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            hitungitem = hitungitem + DataGridView1.Rows(i).Cells(3).Value
            LBLitem.Text = hitungitem
        Next
    End Sub


    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles TextBox1.TextChanged
        LBLkembali.Text = ""
    End Sub
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        If LBLkembali.Text = "" Or LBLnama.Text = "" Or Label9.Text = "" Then
            MsgBox("Transaksi gagal,silahkan lakukan pembayaran terlebih dahulu")
        Else
            Tglmysql = Format(Today, "yyyy-MM-dd")
            Dim simpanjual As String = "insert into tb_jual values ('" & LBLnojual.Text & "','" & Tglmysql & "','" & LBLjam.Text & "','" & LBLitem.Text & "','" & Label9.Text & "','" & TextBox1.Text & "','" & LBLkembali.Text & "','" & ComboBox1.Text & "','" & FormMenuUtama.Stlabel2.Text & "')"
            cmd = New OdbcCommand(simpanjual, con)
            cmd.ExecuteNonQuery()

            For baris As Integer = 0 To DataGridView1.Rows.Count - 2
                Dim simpandetail As String = "insert into tb_detailjual values('" & LBLnojual.Text & "','" & DataGridView1.Rows(baris).Cells(0).Value & "','" & DataGridView1.Rows(baris).Cells(1).Value & "','" & DataGridView1.Rows(baris).Cells(2).Value & "','" & DataGridView1.Rows(baris).Cells(3).Value & "','" & DataGridView1.Rows(baris).Cells(4).Value & "')"
                cmd = New OdbcCommand(simpandetail, con)
                cmd.ExecuteNonQuery()

                cmd = New OdbcCommand("Select * From tb_barang where kodebarang = '" & DataGridView1.Rows(baris).Cells(0).Value & "'", con)
                rd = cmd.ExecuteReader
                rd.Read()
                Dim kurangstok As String = "update tb_barang set jumlahbarang  = '" & rd.Item("jumlahbarang") - DataGridView1.Rows(baris).Cells(3).Value & "' where kodebarang = '" & DataGridView1.Rows(baris).Cells(0).Value & "'"
                cmd = New OdbcCommand(kurangstok, con)
                cmd.ExecuteNonQuery()
            Next
            If MessageBox.Show("Apakah ingin cetak nota...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                AxCrystalReport1.SelectionFormula = "totext({tb_jual.nojual})='" & LBLnojual.Text & "'"
                AxCrystalReport1.ReportFileName = "Nota.rpt"
                AxCrystalReport1.WindowState = Crystal.WindowStateConstants.crptMaximized
                AxCrystalReport1.RetrieveDataFiles()
                AxCrystalReport1.Action = 1
                Call KondisiAwal()
            Else
                Call KondisiAwal()
                MsgBox("Transaksil Berhasil Disimpan")
            End If
        End If
    End Sub
    Private Sub Button3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button2.Click
        Call KondisiAwal()
    End Sub
End Class