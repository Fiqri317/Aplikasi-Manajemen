Imports System
Imports System.Data
Imports System.Data.OleDb

Public Class Dokter
    Dim _koneksiString As String
    Dim _koneksi As New OleDbConnection
    Dim komandambil As New OleDbCommand
    Dim datatabelku As New DataTable
    Dim dataadapterku As New OleDbDataAdapter
    Dim x As String

    Private Sub Dokter_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _koneksiString = "Provider=Microsoft.Jet.OleDb.4.0;" + "Data Source=D:\Campus\Semester V\Kecerdasan Komputasi\Quiz\database\Quiz.mdb;"
        _koneksi.ConnectionString = _koneksiString
        _koneksi.Open()

        komandambil.Connection = _koneksi
        komandambil.CommandType = CommandType.Text

        komandambil.CommandText = "SELECT * FROM Dokter"
        dataadapterku.SelectCommand = komandambil
        dataadapterku.Fill(datatabelku)
        Bs_Coba.DataSource = datatabelku
        dgv_coba.DataSource = Bs_Coba
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim cmdTambah As New OleDbCommand
        Dim tanya As String
        Dim x As DataRow
        cmdTambah.Connection = _koneksi
        cmdTambah.CommandText = "INSERT INTO " + "Dokter ([Kode Dokter], [Nama Dokter], Spesialis, [No Telepon], Alamat)" +
            "VALUES ('" + TextBox1.Text + "','" + TextBox2.Text + "','Spesialis" + TextBox3.Text + "','" +
            TextBox4.Text + "','" + TextBox5.Text + " ')"
        tanya = MessageBox.Show("Data Ini di Simpan ?", "info", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If tanya = vbYes Then
            cmdTambah.ExecuteNonQuery()
            x = datatabelku.NewRow
            x("Kode Dokter") = TextBox1.Text
            x("Nama Dokter") = TextBox2.Text
            x("Spesialis") = "Spesialis " + TextBox3.Text
            x("No Telepon") = TextBox4.Text
            x("Alamat") = TextBox5.Text
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
        TextBox3.Text = dgv_coba.CurrentRow.Cells(2).Value.ToString()
        TextBox4.Text = dgv_coba.CurrentRow.Cells(3).Value.ToString()
        TextBox5.Text = dgv_coba.CurrentRow.Cells(4).Value.ToString()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim cmdUpdate As New OleDbCommand
        cmdUpdate.Connection = _koneksi
        cmdUpdate.CommandType = CommandType.Text
        x = MessageBox.Show("Yakin Data Ingin di Perbarui?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If x = vbYes Then
            cmdUpdate.CommandText = "UPDATE Dokter SET " +
                "[Nama Dokter] = '" + TextBox2.Text + "', " +
                "Spesialis = '" + TextBox3.Text + "', " +
                "[No Telepon] = '" + TextBox4.Text + "', " +
                "Alamat = '" + TextBox5.Text + "' " +
                "WHERE [Kode Dokter] = '" + TextBox1.Text + "'"
            cmdUpdate.ExecuteNonQuery()

            Dim rowToUpdate As DataRow = datatabelku.Select("[Kode Dokter] = '" + TextBox1.Text + "'").FirstOrDefault()
            If rowToUpdate IsNot Nothing Then
                rowToUpdate("Kode Dokter") = TextBox1.Text
                rowToUpdate("Nama Dokter") = TextBox2.Text
                rowToUpdate("Spesialis") = TextBox3.Text
                rowToUpdate("No Telepon") = TextBox4.Text
                rowToUpdate("Alamat") = TextBox5.Text
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
            cmdHapus.CommandText = "DELETE FROM " + "Dokter WHERE [Kode Dokter]= '" + TextBox1.Text + "'"
            cmdHapus.ExecuteNonQuery()
        End If
        Bs_coba.RemoveCurrent()
        dgv_coba.Refresh()

        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox1.Focus()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        datatabelku.Clear()

        komandambil.Connection = _koneksi
        komandambil.CommandType = CommandType.Text
        komandambil.CommandText = "SELECT * FROM Dokter WHERE [Kode Dokter] LIKE '%" + TextBox6.Text + "%' " +
                                       "OR [Nama Dokter] LIKE '%" + TextBox6.Text + "%' " +
                                       "OR [Spesialis] LIKE '%" + TextBox6.Text + "%' " +
                                       "OR [No Telepon] LIKE '%" + TextBox6.Text + "%' " +
                                       "OR [Alamat] LIKE '%" + TextBox6.Text + "%'"

        dataadapterku.SelectCommand = komandambil
        dataadapterku.Fill(datatabelku)

        Bs_coba.DataSource = datatabelku
        dgv_coba.DataSource = Bs_coba
        dgv_coba.Refresh()
    End Sub
End Class