<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Menu
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.DataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SupplierToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.KonsumenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DokterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ObatToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TransaksiPembelianToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BackToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("High Tower Text", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(2, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(398, 32)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "APOTEK SEHAT BERSAMA"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Malgun Gothic", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(64, 35)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(257, 25)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Jl. Kesehatan No. 15, Medan"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.MenuStrip1.Dock = System.Windows.Forms.DockStyle.None
        Me.MenuStrip1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DataToolStripMenuItem, Me.TransaksiPembelianToolStripMenuItem, Me.BackToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(8, 71)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(355, 24)
        Me.MenuStrip1.TabIndex = 3
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'DataToolStripMenuItem
        '
        Me.DataToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SupplierToolStripMenuItem, Me.KonsumenToolStripMenuItem, Me.DokterToolStripMenuItem, Me.ObatToolStripMenuItem})
        Me.DataToolStripMenuItem.Name = "DataToolStripMenuItem"
        Me.DataToolStripMenuItem.Size = New System.Drawing.Size(46, 20)
        Me.DataToolStripMenuItem.Text = "Data"
        '
        'SupplierToolStripMenuItem
        '
        Me.SupplierToolStripMenuItem.Name = "SupplierToolStripMenuItem"
        Me.SupplierToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.SupplierToolStripMenuItem.Text = "Supplier"
        '
        'KonsumenToolStripMenuItem
        '
        Me.KonsumenToolStripMenuItem.Name = "KonsumenToolStripMenuItem"
        Me.KonsumenToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.KonsumenToolStripMenuItem.Text = "Konsumen"
        '
        'DokterToolStripMenuItem
        '
        Me.DokterToolStripMenuItem.Name = "DokterToolStripMenuItem"
        Me.DokterToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.DokterToolStripMenuItem.Text = "Dokter"
        '
        'ObatToolStripMenuItem
        '
        Me.ObatToolStripMenuItem.Name = "ObatToolStripMenuItem"
        Me.ObatToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.ObatToolStripMenuItem.Text = "Obat"
        '
        'TransaksiPembelianToolStripMenuItem
        '
        Me.TransaksiPembelianToolStripMenuItem.Name = "TransaksiPembelianToolStripMenuItem"
        Me.TransaksiPembelianToolStripMenuItem.Size = New System.Drawing.Size(128, 20)
        Me.TransaksiPembelianToolStripMenuItem.Text = "Transaksi Pembelian"
        '
        'BackToolStripMenuItem
        '
        Me.BackToolStripMenuItem.Name = "BackToolStripMenuItem"
        Me.BackToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.BackToolStripMenuItem.Text = "Back"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'Menu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(402, 203)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Name = "Menu"
        Me.Text = "Menu"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents DataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SupplierToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents KonsumenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DokterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ObatToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TransaksiPembelianToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BackToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
