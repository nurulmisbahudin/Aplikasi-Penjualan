Imports System.Data.Odbc
Public Class Form1
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Hide()
    End Sub
    Sub Terbuka()
        FormMenuUtama.LoginToolStripMenuItem.Enabled = False
        FormMenuUtama.LogoutToolStripMenuItem.Enabled = True
        FormMenuUtama.MasterToolStripMenuItem.Enabled = True
        FormMenuUtama.TransaksiToolStripMenuItem.Enabled = True
        FormMenuUtama.LaporanToolStripMenuItem.Enabled = True
        FormMenuUtama.UtilityToolStripMenuItem.Enabled = True
    End Sub
    Sub KondisiAwal()
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Kode Admin dan Password Tidak Boleh Kosong!")
        Else
            Call koneksi()
            cmd = New OdbcCommand("Select * From tb_admin where kodeadmin='" & TextBox1.Text & "' and passwordadmin = '" & TextBox2.Text & "'", con)
            rd = cmd.ExecuteReader
            rd.Read()
            If rd.HasRows Then
                Me.Close()
                Call Terbuka()
                Call Terbuka()
                FormMenuUtama.Stlabel2.Text = rd!kodeadmin
                FormMenuUtama.Stlabel4.Text = rd!namaadmin
                FormMenuUtama.Stlabel6.Text = rd!leveladmin

                If FormMenuUtama.Stlabel6.Text = "User" Then
                    FormMenuUtama.AdminToolStripMenuItem.Enabled = False
                ElseIf FormMenuUtama.Stlabel6.Text = "Admin" Then
                    FormMenuUtama.AdminToolStripMenuItem.Enabled = True
                ElseIf FormMenuUtama.Stlabel6.Text = "Manager" Then
                    FormMenuUtama.MasterToolStripMenuItem.Enabled = False
                    FormMenuUtama.TransaksiToolStripMenuItem.Enabled = False
                End If
            Else
                MsgBox("Kodeadmin atau password salah!")
            End If
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call KondisiAwal()
        TextBox1.Focus()
        'TextBox1.Text = "adm001"
        '.Text = "hahaha"
    End Sub
End Class
