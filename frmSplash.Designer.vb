<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSplash
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
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

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.PictureSplash = New System.Windows.Forms.PictureBox()
        Me.lblAppTitel = New System.Windows.Forms.Label()
        Me.lblCopyright = New System.Windows.Forms.Label()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.lblWWW = New System.Windows.Forms.LinkLabel()
        Me.lblAktion = New System.Windows.Forms.Label()
        CType(Me.PictureSplash, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureSplash
        '
        Me.PictureSplash.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureSplash.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureSplash.Location = New System.Drawing.Point(2, 2)
        Me.PictureSplash.Name = "PictureSplash"
        Me.PictureSplash.Size = New System.Drawing.Size(468, 270)
        Me.PictureSplash.TabIndex = 0
        Me.PictureSplash.TabStop = False
        '
        'lblAppTitel
        '
        Me.lblAppTitel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAppTitel.Location = New System.Drawing.Point(23, 158)
        Me.lblAppTitel.Name = "lblAppTitel"
        Me.lblAppTitel.Size = New System.Drawing.Size(236, 25)
        Me.lblAppTitel.TabIndex = 1
        Me.lblAppTitel.Text = "AppTitle"
        '
        'lblCopyright
        '
        Me.lblCopyright.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.lblCopyright.Location = New System.Drawing.Point(52, 250)
        Me.lblCopyright.Name = "lblCopyright"
        Me.lblCopyright.Size = New System.Drawing.Size(353, 16)
        Me.lblCopyright.TabIndex = 2
        Me.lblCopyright.Text = "Copyright"
        Me.lblCopyright.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblVersion
        '
        Me.lblVersion.Location = New System.Drawing.Point(23, 183)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(418, 18)
        Me.lblVersion.TabIndex = 3
        Me.lblVersion.Text = "Version"
        '
        'lblWWW
        '
        Me.lblWWW.Location = New System.Drawing.Point(23, 201)
        Me.lblWWW.Name = "lblWWW"
        Me.lblWWW.Size = New System.Drawing.Size(418, 18)
        Me.lblWWW.TabIndex = 4
        Me.lblWWW.TabStop = True
        Me.lblWWW.Text = "WWW"
        '
        'lblAktion
        '
        Me.lblAktion.Location = New System.Drawing.Point(23, 222)
        Me.lblAktion.Name = "lblAktion"
        Me.lblAktion.Size = New System.Drawing.Size(417, 16)
        Me.lblAktion.TabIndex = 5
        Me.lblAktion.Text = "Aktion"
        '
        'frmSplash
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(472, 273)
        Me.Controls.Add(Me.lblAktion)
        Me.Controls.Add(Me.lblWWW)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me.lblCopyright)
        Me.Controls.Add(Me.lblAppTitel)
        Me.Controls.Add(Me.PictureSplash)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSplash"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmSplash"
        Me.TopMost = True
        CType(Me.PictureSplash, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PictureSplash As System.Windows.Forms.PictureBox
    Friend WithEvents lblAppTitel As System.Windows.Forms.Label
    Friend WithEvents lblCopyright As System.Windows.Forms.Label
    Friend WithEvents lblVersion As System.Windows.Forms.Label
    Friend WithEvents lblWWW As System.Windows.Forms.LinkLabel
    Friend WithEvents lblAktion As System.Windows.Forms.Label
End Class
