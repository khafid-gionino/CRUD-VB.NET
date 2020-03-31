Imports System.Data.Odbc
Public Class Form1
    Dim CONN As OdbcConnection
    Dim CMD As OdbcCommand
    Dim DS As New DataSet
    Dim DA As OdbcDataAdapter
    Dim RD As OdbcDataReader
    Dim LokasiData As String
    Sub Koneksi()
        LokasiData = "Driver={MySQL ODBC 3.51 Driver};Database=cobavb_db;server=localhost;uid=root"
        CONN = New OdbcConnection(LokasiData)
        If CONN.State = ConnectionState.Closed Then
            CONN.Open()
        End If
    End Sub
    Sub TampilGrid()
        Call Koneksi()
        DA = New OdbcDataAdapter("select * From tbl_mahasiswa ", CONN)
        DS = New DataSet
        DA.Fill(DS, "tbl_mahasiswa")
        DataGridView1.DataSource = DS.Tables("tbl_mahasiswa")
        DataGridView1.ReadOnly = True
    End Sub
    Sub MunculCombo()
        ComboBox1.Items.Add("Informatika")
        ComboBox1.Items.Add("Ekonomi")
        ComboBox1.Items.Add("Hukum")
        ComboBox1.Items.Add("Management")
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call TampilGrid()
        Call MunculCombo()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("Silahkan Isi Semua Form")
        Else
            Call Koneksi()
            Dim simpan As String = "insert into tbl_mahasiswa values ('" & TextBox1.Text & "','" & TextBox2.Text & "',
                                    '" & TextBox3.Text & "','" & TextBox4.Text & "','" & ComboBox1.Text & "')"
            CMD = New OdbcCommand(simpan, CONN)
            CMD.ExecuteNonQuery()
            MsgBox("Input data berhasil",)
            Call TampilGrid()
            Call KosongkanData()
        End If
    End Sub
    Sub KosongkanData()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub
    Private Sub TextBox1_KeyPress1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 6
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            CMD = New OdbcCommand("Select * From tbl_mahasiswa  where NIMMHS='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                MsgBox("NIM Tidak Ada, Silahkan coba lagi!")
                TextBox1.Focus()
            Else
                TextBox2.Text = RD.Item("namamhs")
                TextBox3.Text = RD.Item("alamatmhs")
                TextBox4.Text = RD.Item("teleponmhs")
                ComboBox1.Text = RD.Item("jurusanmhs")
                TextBox2.Focus()
                End If
            End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call Koneksi()
        Dim edit As String = "update tbl_mahasiswa set namamhs='" & TextBox2.Text & "',alamatmhs='" & TextBox3.Text & "',
                            teleponmhs='" & TextBox4.Text & "',jurusanmhs='" & ComboBox1.Text & "' where nimmhs='" & TextBox1.Text & "'"
        CMD = New OdbcCommand(edit, CONN)
        CMD.ExecuteNonQuery()
        MsgBox("Data Berhasil diUpdate", )
        Call TampilGrid()
        Call KosongkanData()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("Silahkan Pilih Data yang akan di hapus dengan Masukan NIM dan ENTER")
        Else
            If MessageBox.Show("Yakin akan dihapus..?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Call Koneksi()
                Dim hapus As String = "delete From tbl_mahasiswa  where nimmhs='" & TextBox1.Text & "'"
                CMD = New OdbcCommand(hapus, CONN)
                CMD.ExecuteNonQuery()
                Call TampilGrid()
                Call KosongkanData()
                End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        End
    End Sub

End Class
