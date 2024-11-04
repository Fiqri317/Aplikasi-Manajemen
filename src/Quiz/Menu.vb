Public Class Menu
    Private Sub SupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SupplierToolStripMenuItem.Click
        Supplier.Show()
    End Sub

    Private Sub KonsumenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KonsumenToolStripMenuItem.Click
        Konsumen.Show()
    End Sub

    Private Sub DokterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DokterToolStripMenuItem.Click
        Dokter.Show()
    End Sub

    Private Sub ObatToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ObatToolStripMenuItem.Click
        Obat.Show()
    End Sub

    Private Sub TransaksiPembelianToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TransaksiPembelianToolStripMenuItem.Click
        Transaksi_Pembelian.Show()
    End Sub

    Private Sub BackToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackToolStripMenuItem.Click
        Me.Close()
        Login.Show()
        Login.TextBox1.Clear()
        Login.TextBox2.Clear()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub
End Class