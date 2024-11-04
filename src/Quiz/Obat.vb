Imports System
Imports System.Data
Imports System.Data.OleDb

Public Class Obat
    Dim _koneksiString As String
    Dim _koneksi As New OleDbConnection
    Dim komandambil As New OleDbCommand
    Dim datatabelku As New DataTable
    Dim dataadapterku As New OleDbDataAdapter
    Dim x As String

    Private Sub Obat_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _koneksiString = "Provider=Microsoft.Jet.OleDb.4.0;" + "Data Source=D:\Campus\Semester V\Kecerdasan Komputasi\Quiz\database\Quiz.mdb;"
        _koneksi.ConnectionString = _koneksiString
        _koneksi.Open()

        komandambil.Connection = _koneksi
        komandambil.CommandType = CommandType.Text

        komandambil.CommandText = "SELECT * FROM Obat"
        dataadapterku.SelectCommand = komandambil
        dataadapterku.Fill(datatabelku)
        Bs_Coba.DataSource = datatabelku
        dgv_coba.DataSource = Bs_Coba
    End Sub

    Private Function GetHarga(ByVal jenis As String, ByVal satuan As String, ByVal golongan As String) As Integer
        Dim harga As Integer = 0
        Select Case jenis
            Case "Tablet"
                Select Case satuan
                    Case "Box"
                        Select Case golongan
                            Case "Obat Bebas" : harga = 5000
                            Case "Obat Keras" : harga = 10000
                            Case "Obat Generik" : harga = 7500
                            Case "Obat Khusus" : harga = 15000
                        End Select
                    Case "Botol"
                        Select Case golongan
                            Case "Obat Bebas" : harga = 6000
                            Case "Obat Keras" : harga = 12000
                            Case "Obat Generik" : harga = 8500
                            Case "Obat Khusus" : harga = 18000
                        End Select
                    Case "Strip"
                        Select Case golongan
                            Case "Obat Bebas" : harga = 4000
                            Case "Obat Keras" : harga = 9000
                            Case "Obat Generik" : harga = 6500
                            Case "Obat Khusus" : harga = 14000
                        End Select
                    Case "Pcs"
                        Select Case golongan
                            Case "Obat Bebas" : harga = 1500
                            Case "Obat Keras" : harga = 3000
                            Case "Obat Generik" : harga = 2000
                            Case "Obat Khusus" : harga = 5000
                        End Select
                End Select
            Case "Kapsul"
                Select Case satuan
                    Case "Box"
                        Select Case golongan
                            Case "Obat Bebas" : harga = 7000
                            Case "Obat Keras" : harga = 12000
                            Case "Obat Generik" : harga = 9000
                            Case "Obat Khusus" : harga = 16000
                        End Select
                    Case "Botol"
                        Select Case golongan
                            Case "Obat Bebas" : harga = 8000
                            Case "Obat Keras" : harga = 14000
                            Case "Obat Generik" : harga = 10000
                            Case "Obat Khusus" : harga = 20000
                        End Select
                    Case "Strip"
                        Select Case golongan
                            Case "Obat Bebas" : harga = 5000
                            Case "Obat Keras" : harga = 11000
                            Case "Obat Generik" : harga = 7500
                            Case "Obat Khusus" : harga = 15000
                        End Select
                    Case "Pcs"
                        Select Case golongan
                            Case "Obat Bebas" : harga = 2000
                            Case "Obat Keras" : harga = 4000
                            Case "Obat Generik" : harga = 3000
                            Case "Obat Khusus" : harga = 6000
                        End Select
                End Select
            Case "Sirup"
                Select Case satuan
                    Case "Box"
                        Select Case golongan
                            Case "Obat Bebas" : harga = 10000
                            Case "Obat Keras" : harga = 15000
                            Case "Obat Generik" : harga = 12000
                            Case "Obat Khusus" : harga = 18000
                        End Select
                    Case "Botol"
                        Select Case golongan
                            Case "Obat Bebas" : harga = 11000
                            Case "Obat Keras" : harga = 17000
                            Case "Obat Generik" : harga = 13000
                            Case "Obat Khusus" : harga = 22000
                        End Select
                    Case "Strip"
                        Select Case golongan
                            Case "Obat Bebas" : harga = 7000
                            Case "Obat Keras" : harga = 14000
                            Case "Obat Generik" : harga = 9000
                            Case "Obat Khusus" : harga = 16000
                        End Select
                    Case "Pcs"
                        Select Case golongan
                            Case "Obat Bebas" : harga = 3000
                            Case "Obat Keras" : harga = 5000
                            Case "Obat Generik" : harga = 4000
                            Case "Obat Khusus" : harga = 7000
                        End Select
                End Select
            Case "Salep"
                Select Case satuan
                    Case "Box"
                        Select Case golongan
                            Case "Obat Bebas" : harga = 8000
                            Case "Obat Keras" : harga = 14000
                            Case "Obat Generik" : harga = 10000
                            Case "Obat Khusus" : harga = 15000
                        End Select
                    Case "Botol"
                        Select Case golongan
                            Case "Obat Bebas" : harga = 9000
                            Case "Obat Keras" : harga = 16000
                            Case "Obat Generik" : harga = 12000
                            Case "Obat Khusus" : harga = 20000
                        End Select
                    Case "Strip"
                        Select Case golongan
                            Case "Obat Bebas" : harga = 5000
                            Case "Obat Keras" : harga = 11000
                            Case "Obat Generik" : harga = 8000
                            Case "Obat Khusus" : harga = 14000
                        End Select
                    Case "Pcs"
                        Select Case golongan
                            Case "Obat Bebas" : harga = 2000
                            Case "Obat Keras" : harga = 4000
                            Case "Obat Generik" : harga = 3000
                            Case "Obat Khusus" : harga = 6000
                        End Select
                End Select
        End Select

        Return harga
    End Function

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ComboBox1.SelectedIndexChanged, ComboBox2.SelectedIndexChanged, ComboBox3.SelectedIndexChanged
        Dim harga As Integer = GetHarga(ComboBox1.Text, ComboBox2.Text, ComboBox3.Text)
        TextBox3.Text = harga.ToString()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim cmdTambah As New OleDbCommand
        Dim tanya As String
        Dim x As DataRow
        cmdTambah.Connection = _koneksi
        cmdTambah.CommandText = "INSERT INTO " + "Obat ([Kode Obat], [Nama Obat], Jenis, Satuan, Golongan, [Harga Satuan])" +
            "VALUES ('" + TextBox1.Text + "','" + TextBox2.Text + "','" + ComboBox1.Text + "','" +
            ComboBox2.Text + "','" + ComboBox3.Text + "','" + TextBox3.Text + " ')"
        tanya = MessageBox.Show("Data Ini di Simpan ?", "info", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If tanya = vbYes Then
            cmdTambah.ExecuteNonQuery()
            x = datatabelku.NewRow
            x("Kode Obat") = TextBox1.Text
            x("Nama Obat") = TextBox2.Text
            x("Jenis") = ComboBox1.Text
            x("Satuan") = ComboBox2.Text
            x("Golongan") = ComboBox3.Text
            x("Harga Satuan") = TextBox3.Text
            datatabelku.Rows.Add(x)
            Bs_coba.DataSource = Nothing
            Bs_coba.DataSource = datatabelku

            dgv_coba.Refresh()
            Bs_coba.MoveFirst()
        End If
    End Sub

    Private Sub dgv_coba_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_coba.CellContentClick
        TextBox1.Text = dgv_coba.CurrentRow.Cells(0).Value.ToString()
        TextBox2.Text = dgv_coba.CurrentRow.Cells(1).Value.ToString()
        ComboBox1.Text = dgv_coba.CurrentRow.Cells(2).Value.ToString()
        ComboBox2.Text = dgv_coba.CurrentRow.Cells(3).Value.ToString()
        ComboBox3.Text = dgv_coba.CurrentRow.Cells(4).Value.ToString()
        TextBox3.Text = dgv_coba.CurrentRow.Cells(5).Value.ToString()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim cmdUpdate As New OleDbCommand
        cmdUpdate.Connection = _koneksi
        cmdUpdate.CommandType = CommandType.Text
        x = MessageBox.Show("Yakin Data Ingin di Perbarui?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If x = vbYes Then
            cmdUpdate.CommandText = "UPDATE Obat SET " +
                "[Nama Obat] = '" + TextBox2.Text + "', " +
                "Jenis = '" + ComboBox1.Text + "', " +
                "Satuan = '" + ComboBox2.Text + "', " +
                "Golongan = '" + ComboBox3.Text + "', " +
                "[Harga Satuan] = '" + TextBox3.Text + "' " +
                "WHERE [Kode Obat] = '" + TextBox1.Text + "'"
            cmdUpdate.ExecuteNonQuery()

            Dim rowToUpdate As DataRow = datatabelku.Select("[Kode Obat] = '" + TextBox1.Text + "'").FirstOrDefault()
            If rowToUpdate IsNot Nothing Then
                rowToUpdate("Kode Obat") = TextBox1.Text
                rowToUpdate("Nama Obat") = TextBox2.Text
                rowToUpdate("Jenis") = ComboBox1.Text
                rowToUpdate("Satuan") = ComboBox2.Text
                rowToUpdate("Golongan") = ComboBox3.Text
                rowToUpdate("Harga Satuan") = TextBox3.Text
            End If

            Bs_coba.DataSource = Nothing
            Bs_coba.DataSource = datatabelku
            dgv_coba.Refresh()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim cmdHapus As New OleDbCommand
        cmdHapus.Connection = _koneksi
        cmdHapus.CommandType = CommandType.Text
        x = MessageBox.Show("Yakin Data Akan di Hapus ?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If x = vbYes Then
            cmdHapus.CommandText = "DELETE FROM " + "Obat WHERE [Kode Obat]= '" + TextBox1.Text + "'"
            cmdHapus.ExecuteNonQuery()
        End If
        Bs_coba.RemoveCurrent()
        dgv_coba.Refresh()

        TextBox1.Clear()
        TextBox2.Clear()
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox1.Focus()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox1.Clear()
        TextBox2.Clear()
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        TextBox3.Clear()
        TextBox4.Clear()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        datatabelku.Clear()

        komandambil.Connection = _koneksi
        komandambil.CommandType = CommandType.Text
        komandambil.CommandText = "SELECT * FROM Obat WHERE [Kode Obat] LIKE '%" + TextBox4.Text + "%' " +
                                       "OR [Nama Obat] LIKE '%" + TextBox4.Text + "%' " +
                                       "OR [Jenis] LIKE '%" + TextBox4.Text + "%' " +
                                       "OR [Satuan] LIKE '%" + TextBox4.Text + "%' " +
                                       "OR [Golongan] LIKE '%" + TextBox4.Text + "%' " +
                                       "OR [Harga Satuan] LIKE '%" + TextBox4.Text + "%'"

        dataadapterku.SelectCommand = komandambil
        dataadapterku.Fill(datatabelku)

        Bs_coba.DataSource = datatabelku
        dgv_coba.DataSource = Bs_coba
        dgv_coba.Refresh()
    End Sub
End Class