Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing.Printing

Public Class Transaksi_Pembelian
    Dim _koneksiString As String
    Dim _koneksi As New OleDbConnection
    Dim komandambil As New OleDbCommand
    Dim datatabelku As New DataTable
    Dim dataadapterku As New OleDbDataAdapter
    Dim x As String

    Private Sub dgv_coba_CellFormatting(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) Handles dgv_coba.CellFormatting
        If dgv_coba.Columns(e.ColumnIndex).Name = "Tanggal Faktur" Then
            If e.Value IsNot Nothing Then
                Dim tgl As DateTime = Convert.ToDateTime(e.Value)
                e.Value = tgl.ToString("dd MMMM yyyy", New System.Globalization.CultureInfo("id-ID"))
                e.FormattingApplied = True
            End If
        End If

        If dgv_coba.Columns(e.ColumnIndex).Name = "Tanggal Expired" Then
            If e.Value IsNot Nothing Then
                Dim tgl As DateTime = Convert.ToDateTime(e.Value)
                e.Value = tgl.ToString("dd MMMM yyyy", New System.Globalization.CultureInfo("id-ID"))
                e.FormattingApplied = True
            End If
        End If
    End Sub

    Private Sub Transaksi_Pembelian_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _koneksiString = "Provider=Microsoft.Jet.OleDb.4.0;" + "Data Source=D:\Campus\Semester V\Kecerdasan Komputasi\Quiz\database\Quiz.mdb;"
        _koneksi.ConnectionString = _koneksiString
        _koneksi.Open()


        LoadPembelian()
        LoadSupplier()
        LoadObat()
    End Sub

    Private Sub LoadPembelian()
        komandambil.Connection = _koneksi
        komandambil.CommandText = "SELECT * FROM Pembelian"
        dataadapterku.SelectCommand = komandambil
        dataadapterku.Fill(datatabelku)
        Bs_coba.DataSource = datatabelku
        dgv_coba.DataSource = Bs_coba
    End Sub

    Private Sub LoadSupplier()
        Dim cmd As New OleDbCommand("SELECT [Nama Supplier], Alamat FROM Supplier", _koneksi)
        Dim adapter As New OleDbDataAdapter(cmd)
        Dim datatabelSupplier As New DataTable()
        adapter.Fill(datatabelSupplier)

        ComboBox1.DisplayMember = "Nama Supplier"
        ComboBox1.ValueMember = "Alamat Supplier"
        ComboBox1.DataSource = datatabelSupplier
    End Sub

    Private Sub LoadObat()
        Dim cmd As New OleDbCommand("SELECT [Nama Obat], [Harga Satuan] FROM Obat", _koneksi)
        Dim adapter As New OleDbDataAdapter(cmd)
        Dim datatabelSupplier As New DataTable()
        adapter.Fill(datatabelSupplier)

        ComboBox2.DisplayMember = "Nama Obat"
        ComboBox2.ValueMember = "Harga Satuan"
        ComboBox2.DataSource = datatabelSupplier
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex <> -1 Then
            Dim selectedRow As DataRowView = CType(ComboBox1.SelectedItem, DataRowView)
            TextBox2.Text = selectedRow("Alamat").ToString()
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex <> -1 Then
            Dim selectedRow As DataRowView = CType(ComboBox2.SelectedItem, DataRowView)
            TextBox4.Text = selectedRow("Harga Satuan").ToString()
        End If
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        HitungTotalHarga()
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        HitungTotalHarga()
    End Sub

    Private Sub HitungTotalHarga()
        Dim hargaBeli As Decimal
        Dim banyak As Integer

        If Decimal.TryParse(TextBox4.Text, hargaBeli) AndAlso Integer.TryParse(TextBox5.Text, banyak) Then
            Dim totalHarga As Decimal = hargaBeli * banyak
            TextBox6.Text = totalHarga.ToString("N0")
        Else
            TextBox6.Text = ""
        End If
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim cmdTambah As New OleDbCommand
        Dim tanya As String
        Dim x As DataRow
        cmdTambah.Connection = _koneksi
        cmdTambah.CommandText = "INSERT INTO " + "Pembelian ([Nama Faktur], [Tanggal Faktur], [Nama Supplier], [Alamat Supplier], [Tanggal Expired], [Nama Obat], Stok, [Harga Beli], Banyak, [Total Harga])" +
            "VALUES ('" + TextBox1.Text + "','" + DateTimePicker1.Text + "','" + ComboBox1.Text + "','" +
            TextBox2.Text + "','" + DateTimePicker2.Text + "','" + ComboBox2.Text + "','" +
            TextBox3.Text + "','" + TextBox4.Text + "','" + TextBox5.Text + "','" + TextBox6.Text + " ')"
        tanya = MessageBox.Show("Data Ini di Simpan ?", "info", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If tanya = vbYes Then
            cmdTambah.ExecuteNonQuery()
            x = datatabelku.NewRow
            x("Nama Faktur") = TextBox1.Text
            x("Tanggal Faktur") = DateTimePicker1.Text
            x("Nama Supplier") = ComboBox1.Text
            x("Alamat Supplier") = TextBox2.Text
            x("Tanggal Expired") = DateTimePicker2.Text
            x("Nama Obat") = ComboBox2.Text
            x("Stok") = TextBox3.Text
            x("Harga Beli") = TextBox4.Text
            x("Banyak") = TextBox5.Text
            x("Total Harga") = TextBox6.Text
            datatabelku.Rows.Add(x)
            Bs_coba.DataSource = Nothing
            Bs_coba.DataSource = datatabelku

            dgv_coba.Refresh()
            Bs_coba.MoveFirst()
        End If
    End Sub

    Private Sub dgv_coba_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_coba.CellContentClick
        TextBox1.Text = dgv_coba.CurrentRow.Cells(0).Value.ToString()
        DateTimePicker1.Value = Convert.ToDateTime(dgv_coba.CurrentRow.Cells(1).Value)
        ComboBox1.Text = dgv_coba.CurrentRow.Cells(2).Value.ToString()
        TextBox2.Text = dgv_coba.CurrentRow.Cells(3).Value.ToString()
        DateTimePicker2.Value = Convert.ToDateTime(dgv_coba.CurrentRow.Cells(4).Value)
        ComboBox2.Text = dgv_coba.CurrentRow.Cells(5).Value.ToString()
        TextBox3.Text = dgv_coba.CurrentRow.Cells(6).Value.ToString()
        TextBox4.Text = dgv_coba.CurrentRow.Cells(7).Value.ToString()
        TextBox5.Text = dgv_coba.CurrentRow.Cells(8).Value.ToString()
        TextBox6.Text = dgv_coba.CurrentRow.Cells(9).Value.ToString()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim cmdUpdate As New OleDbCommand
        cmdUpdate.Connection = _koneksi
        cmdUpdate.CommandType = CommandType.Text
        x = MessageBox.Show("Yakin Data Ingin di Perbarui?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If x = vbYes Then
            cmdUpdate.CommandText = "UPDATE Pembelian SET " +
                "[Tanggal Faktur] = '" + DateTimePicker1.Text + "', " +
                "[Nama Supplier] = '" + ComboBox1.Text + "', " +
                "[Alamat Supplier] = '" + TextBox2.Text + "', " +
                "[Tanggal Expired] = '" + DateTimePicker2.Text + "', " +
                "[Nama Obat] = '" + ComboBox2.Text + "', " +
                "Stok = '" + TextBox3.Text + "', " +
                "[Harga Beli] = '" + TextBox4.Text + "', " +
                "Banyak = '" + TextBox5.Text + "', " +
                "[Total Harga] = '" + TextBox6.Text + "' " +
                "WHERE [Nama Faktur] = '" + TextBox1.Text + "'"
            cmdUpdate.ExecuteNonQuery()

            Dim rowToUpdate As DataRow = datatabelku.Select("[Nama Faktur] = '" + TextBox1.Text + "'").FirstOrDefault()
            If rowToUpdate IsNot Nothing Then
                rowToUpdate("Nama Faktur") = TextBox1.Text
                rowToUpdate("Tanggal Faktur") = DateTimePicker1.Text
                rowToUpdate("Nama Supplier") = ComboBox1.Text
                rowToUpdate("Alamat Supplier") = TextBox2.Text
                rowToUpdate("Tanggal Expired") = DateTimePicker2.Text
                rowToUpdate("Nama Obat") = ComboBox2.Text
                rowToUpdate("Stok") = TextBox3.Text
                rowToUpdate("Harga Beli") = TextBox4.Text
                rowToUpdate("Banyak") = TextBox5.Text
                rowToUpdate("Total Harga") = TextBox6.Text
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
            cmdHapus.CommandText = "DELETE FROM " + "Pembelian WHERE [Nama Faktur]= '" + TextBox1.Text + "'"
            cmdHapus.ExecuteNonQuery()
        End If
        Bs_coba.RemoveCurrent()
        dgv_coba.Refresh()

        TextBox1.Clear()
        DateTimePicker1.Value = DateTime.Now
        ComboBox1.SelectedIndex = -1
        TextBox2.Clear()
        DateTimePicker2.Value = DateTime.Now
        ComboBox2.SelectedIndex = -1
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        RichTextBoxStruk.Clear()
        TextBox1.Focus()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox1.Clear()
        DateTimePicker1.Value = DateTime.Now
        ComboBox1.SelectedIndex = -1
        TextBox2.Clear()
        DateTimePicker2.Value = DateTime.Now
        ComboBox2.SelectedIndex = -1
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        RichTextBoxStruk.Clear()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        datatabelku.Clear()

        komandambil.Connection = _koneksi
        komandambil.CommandType = CommandType.Text
        komandambil.CommandText = "SELECT * FROM Pembelian WHERE [Nama Faktur] LIKE '%" + TextBox7.Text + "%' " +
                                       "OR FORMAT([Tanggal Faktur], 'dd MMMM yyyy') LIKE '%" + TextBox7.Text + "%' " +
                                       "OR [Nama Supplier] LIKE '%" + TextBox7.Text + "%' " +
                                       "OR [Alamat Supplier] LIKE '%" + TextBox7.Text + "%' " +
                                       "OR FORMAT([Tanggal Expired], 'dd MMMM yyyy') LIKE '%" + TextBox7.Text + "%' " +
                                       "OR [Nama Obat] LIKE '%" + TextBox7.Text + "%' " +
                                       "OR [Stok] LIKE '%" + TextBox7.Text + "%' " +
                                       "OR [Harga Beli] LIKE '%" + TextBox7.Text + "%' " +
                                       "OR [Banyak] LIKE '%" + TextBox7.Text + "%' " +
                                       "OR [Total Harga] LIKE '%" + TextBox7.Text + "%'"

        dataadapterku.SelectCommand = komandambil
        dataadapterku.Fill(datatabelku)

        Bs_coba.DataSource = datatabelku
        dgv_coba.DataSource = Bs_coba
        dgv_coba.Refresh()
    End Sub

    Private Sub TampilkanStruk()
        RichTextBoxStruk.Clear()

        RichTextBoxStruk.AppendText("----------------------------- STRUK PEMBELIAN -----------------------------" + Environment.NewLine)
        RichTextBoxStruk.AppendText("Nama Faktur        : " + TextBox1.Text + Environment.NewLine)
        RichTextBoxStruk.AppendText("Tanggal Faktur     : " + DateTimePicker1.Text + Environment.NewLine)
        RichTextBoxStruk.AppendText("Nama Supplier      : " + ComboBox1.Text + Environment.NewLine)
        RichTextBoxStruk.AppendText("Alamat Supplier     : " + TextBox2.Text + Environment.NewLine)
        RichTextBoxStruk.AppendText("Tanggal Expired    : " + DateTimePicker2.Text + Environment.NewLine)
        RichTextBoxStruk.AppendText("----------------------------------------------------------------------------------------------" + Environment.NewLine)

        RichTextBoxStruk.AppendText("Nama Obat  : " + ComboBox2.Text + Environment.NewLine)
        RichTextBoxStruk.AppendText("Stok            : " + TextBox3.Text + Environment.NewLine)
        RichTextBoxStruk.AppendText("Harga Beli   : " + TextBox4.Text + Environment.NewLine)
        RichTextBoxStruk.AppendText("Banyak        : " + TextBox5.Text + Environment.NewLine)
        RichTextBoxStruk.AppendText("Total Harga : " + TextBox6.Text + Environment.NewLine)
        RichTextBoxStruk.AppendText("----------------------------------------------------------------------------------------------" + Environment.NewLine)

        RichTextBoxStruk.AppendText("Terima Kasih Telah Berbelanja !" + Environment.NewLine)
        RichTextBoxStruk.AppendText("-------------------------- SEMOGA BERMANFAAT --------------------------")
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        TampilkanStruk()
    End Sub
End Class
