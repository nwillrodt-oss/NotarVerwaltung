<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEdit
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
        Me.GrpBoxDaten = New System.Windows.Forms.GroupBox()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnEsc = New System.Windows.Forms.Button()
        Me.lblDirty = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'GrpBoxDaten
        '
        Me.GrpBoxDaten.Location = New System.Drawing.Point(5, 5)
        Me.GrpBoxDaten.Name = "GrpBoxDaten"
        Me.GrpBoxDaten.Size = New System.Drawing.Size(436, 74)
        Me.GrpBoxDaten.TabIndex = 0
        Me.GrpBoxDaten.TabStop = False
        Me.GrpBoxDaten.Text = "Daten"
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(131, 249)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(102, 23)
        Me.btnOK.TabIndex = 8
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(239, 249)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(102, 23)
        Me.btnSave.TabIndex = 7
        Me.btnSave.Text = "Übernehmen"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnEsc
        '
        Me.btnEsc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnEsc.Location = New System.Drawing.Point(347, 249)
        Me.btnEsc.Name = "btnEsc"
        Me.btnEsc.Size = New System.Drawing.Size(92, 23)
        Me.btnEsc.TabIndex = 6
        Me.btnEsc.Text = "Abbrechen"
        Me.btnEsc.UseVisualStyleBackColor = True
        '
        'lblDirty
        '
        Me.lblDirty.AutoSize = True
        Me.lblDirty.Location = New System.Drawing.Point(2, 251)
        Me.lblDirty.Name = "lblDirty"
        Me.lblDirty.Size = New System.Drawing.Size(62, 13)
        Me.lblDirty.TabIndex = 9
        Me.lblDirty.Text = "gespeichert"
        '
        'frmEdit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(442, 273)
        Me.Controls.Add(Me.lblDirty)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnEsc)
        Me.Controls.Add(Me.GrpBoxDaten)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(450, 150)
        Me.Name = "frmEdit"
        Me.Text = "frmEdit"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GrpBoxDaten As System.Windows.Forms.GroupBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnEsc As System.Windows.Forms.Button
    Friend WithEvents lblDirty As System.Windows.Forms.Label
End Class
