<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormPrincipale
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormPrincipale))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.not_zero = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.OuvrirToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FermerToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.MeteoDetails = New System.Windows.Forms.ToolTip(Me.components)
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Controls.Add(Me.DateTimePicker1)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(964, 20)
        Me.Panel1.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.Label2.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(23, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(22, 20)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "-"
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.PictureBox1.BackgroundImage = Global.Organizer.My.Resources.Resources.parametress
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Left
        Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(23, 20)
        Me.PictureBox1.TabIndex = 11
        Me.PictureBox1.TabStop = False
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePicker1.CalendarForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.DateTimePicker1.CalendarMonthBackground = System.Drawing.SystemColors.Info
        Me.DateTimePicker1.Dock = System.Windows.Forms.DockStyle.Right
        Me.DateTimePicker1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePicker1.Location = New System.Drawing.Point(741, 0)
        Me.DateTimePicker1.Margin = New System.Windows.Forms.Padding(0)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.DateTimePicker1.Size = New System.Drawing.Size(223, 20)
        Me.DateTimePicker1.TabIndex = 10
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label1.Location = New System.Drawing.Point(0, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(964, 20)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Mois Annee"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'not_zero
        '
        Me.not_zero.ContextMenuStrip = Me.ContextMenuStrip1
        Me.not_zero.Icon = CType(resources.GetObject("not_zero.Icon"), System.Drawing.Icon)
        Me.not_zero.Text = "Organizer" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Alt+F11: Cacher" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Alt+F12: Ouvrir"
        Me.not_zero.Visible = True
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OuvrirToolStripMenuItem, Me.FermerToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(105, 48)
        '
        'OuvrirToolStripMenuItem
        '
        Me.OuvrirToolStripMenuItem.Name = "OuvrirToolStripMenuItem"
        Me.OuvrirToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F11), System.Windows.Forms.Keys)
        Me.OuvrirToolStripMenuItem.ShowShortcutKeys = False
        Me.OuvrirToolStripMenuItem.Size = New System.Drawing.Size(104, 22)
        Me.OuvrirToolStripMenuItem.Text = "Ouvrir"
        '
        'FermerToolStripMenuItem
        '
        Me.FermerToolStripMenuItem.Name = "FermerToolStripMenuItem"
        Me.FermerToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F12), System.Windows.Forms.Keys)
        Me.FermerToolStripMenuItem.ShowShortcutKeys = False
        Me.FermerToolStripMenuItem.Size = New System.Drawing.Size(104, 22)
        Me.FermerToolStripMenuItem.Text = "Fermer"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.[Single]
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 20)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 3
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 60.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(964, 195)
        Me.TableLayoutPanel1.TabIndex = 1
        '
        'MeteoDetails
        '
        Me.MeteoDetails.AutoPopDelay = 10000
        Me.MeteoDetails.InitialDelay = 500
        Me.MeteoDetails.IsBalloon = True
        Me.MeteoDetails.ReshowDelay = 100
        '
        'Timer1
        '
        '
        'Timer2
        '
        Me.Timer2.Interval = 86400000
        '
        'FormPrincipale
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.ClientSize = New System.Drawing.Size(964, 215)
        Me.ControlBox = False
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.Panel1)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimizeBox = False
        Me.Name = "FormPrincipale"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Organizer"
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents not_zero As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents OuvrirToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FermerToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents MeteoDetails As System.Windows.Forms.ToolTip
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Timer2 As System.Windows.Forms.Timer

End Class
