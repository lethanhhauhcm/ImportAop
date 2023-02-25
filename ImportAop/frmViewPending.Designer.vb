<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmViewPending
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
        Me.dgrPendings = New System.Windows.Forms.DataGridView()
        Me.txtQuerry = New System.Windows.Forms.TextBox()
        Me.lbkSearch = New System.Windows.Forms.LinkLabel()
        CType(Me.dgrPendings, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgrPendings
        '
        Me.dgrPendings.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgrPendings.Location = New System.Drawing.Point(0, 27)
        Me.dgrPendings.Name = "dgrPendings"
        Me.dgrPendings.Size = New System.Drawing.Size(583, 458)
        Me.dgrPendings.TabIndex = 0
        '
        'txtQuerry
        '
        Me.txtQuerry.Location = New System.Drawing.Point(0, 491)
        Me.txtQuerry.Multiline = True
        Me.txtQuerry.Name = "txtQuerry"
        Me.txtQuerry.Size = New System.Drawing.Size(583, 69)
        Me.txtQuerry.TabIndex = 1
        '
        'lbkSearch
        '
        Me.lbkSearch.AutoSize = True
        Me.lbkSearch.Location = New System.Drawing.Point(443, 9)
        Me.lbkSearch.Name = "lbkSearch"
        Me.lbkSearch.Size = New System.Drawing.Size(41, 13)
        Me.lbkSearch.TabIndex = 2
        Me.lbkSearch.TabStop = True
        Me.lbkSearch.Text = "Search"
        '
        'frmViewPending
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(584, 561)
        Me.Controls.Add(Me.lbkSearch)
        Me.Controls.Add(Me.txtQuerry)
        Me.Controls.Add(Me.dgrPendings)
        Me.Name = "frmViewPending"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ViewPending"
        CType(Me.dgrPendings, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents dgrPendings As DataGridView
    Friend WithEvents txtQuerry As TextBox
    Friend WithEvents lbkSearch As LinkLabel
End Class
