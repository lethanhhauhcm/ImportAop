<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.TxtFeedBack = New System.Windows.Forms.TextBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.barRun = New System.Windows.Forms.ToolStripMenuItem()
        Me.barRunAir = New System.Windows.Forms.ToolStripMenuItem()
        Me.barRunNonAir = New System.Windows.Forms.ToolStripMenuItem()
        Me.barRunAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.barAllOdbc = New System.Windows.Forms.ToolStripMenuItem()
        Me.barViewPendingQ = New System.Windows.Forms.ToolStripMenuItem()
        Me.barStop = New System.Windows.Forms.ToolStripMenuItem()
        Me.TestToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.barDeleteDupCust4HAN = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TxtFeedBack
        '
        Me.TxtFeedBack.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFeedBack.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtFeedBack.Location = New System.Drawing.Point(3, 29)
        Me.TxtFeedBack.Multiline = True
        Me.TxtFeedBack.Name = "TxtFeedBack"
        Me.TxtFeedBack.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TxtFeedBack.Size = New System.Drawing.Size(607, 229)
        Me.TxtFeedBack.TabIndex = 2
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.barRun, Me.barViewPendingQ, Me.barStop, Me.TestToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(611, 24)
        Me.MenuStrip1.TabIndex = 3
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'barRun
        '
        Me.barRun.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.barRunAir, Me.barRunNonAir, Me.barRunAll, Me.barAllOdbc})
        Me.barRun.Name = "barRun"
        Me.barRun.Size = New System.Drawing.Size(40, 20)
        Me.barRun.Text = "Run"
        '
        'barRunAir
        '
        Me.barRunAir.Name = "barRunAir"
        Me.barRunAir.Size = New System.Drawing.Size(180, 22)
        Me.barRunAir.Text = "Air"
        '
        'barRunNonAir
        '
        Me.barRunNonAir.Name = "barRunNonAir"
        Me.barRunNonAir.Size = New System.Drawing.Size(180, 22)
        Me.barRunNonAir.Text = "NonAir"
        '
        'barRunAll
        '
        Me.barRunAll.Name = "barRunAll"
        Me.barRunAll.Size = New System.Drawing.Size(180, 22)
        Me.barRunAll.Text = "All"
        '
        'barAllOdbc
        '
        Me.barAllOdbc.Name = "barAllOdbc"
        Me.barAllOdbc.Size = New System.Drawing.Size(180, 22)
        Me.barAllOdbc.Text = "AllOdbc"
        '
        'barViewPendingQ
        '
        Me.barViewPendingQ.Name = "barViewPendingQ"
        Me.barViewPendingQ.Size = New System.Drawing.Size(97, 20)
        Me.barViewPendingQ.Text = "ViewPendingQ"
        '
        'barStop
        '
        Me.barStop.Name = "barStop"
        Me.barStop.Size = New System.Drawing.Size(43, 20)
        Me.barStop.Text = "Stop"
        '
        'TestToolStripMenuItem
        '
        Me.TestToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.barDeleteDupCust4HAN})
        Me.TestToolStripMenuItem.Name = "TestToolStripMenuItem"
        Me.TestToolStripMenuItem.Size = New System.Drawing.Size(41, 20)
        Me.TestToolStripMenuItem.Text = "Test"
        '
        'barDeleteDupCust4HAN
        '
        Me.barDeleteDupCust4HAN.Name = "barDeleteDupCust4HAN"
        Me.barDeleteDupCust4HAN.Size = New System.Drawing.Size(185, 22)
        Me.barDeleteDupCust4HAN.Text = "DeleteDupCust4HAN"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(611, 261)
        Me.Controls.Add(Me.TxtFeedBack)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmMain"
        Me.Text = "ImportAOP"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TxtFeedBack As TextBox
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents barRun As ToolStripMenuItem
    Friend WithEvents barViewPendingQ As ToolStripMenuItem
    Friend WithEvents barRunAir As ToolStripMenuItem
    Friend WithEvents barRunNonAir As ToolStripMenuItem
    Friend WithEvents barRunAll As ToolStripMenuItem
    Friend WithEvents barStop As ToolStripMenuItem
    Friend WithEvents barAllOdbc As ToolStripMenuItem
    Friend WithEvents TestToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents barDeleteDupCust4HAN As ToolStripMenuItem
End Class
