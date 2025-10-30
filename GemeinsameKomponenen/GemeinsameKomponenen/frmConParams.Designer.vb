<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConParams
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmConParams))
        Me.radSQL = New System.Windows.Forms.RadioButton()
        Me.radAcc = New System.Windows.Forms.RadioButton()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.lblServer = New System.Windows.Forms.Label()
        Me.lblDBName = New System.Windows.Forms.Label()
        Me.txtDBName = New System.Windows.Forms.TextBox()
        Me.lblDBUser = New System.Windows.Forms.Label()
        Me.txtDBUSer = New System.Windows.Forms.TextBox()
        Me.lblPWD = New System.Windows.Forms.Label()
        Me.txtPWD = New System.Windows.Forms.TextBox()
        Me.chkNT = New System.Windows.Forms.CheckBox()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnEsc = New System.Windows.Forms.Button()
        Me.btnDBPath = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'radSQL
        '
        Me.radSQL.AutoSize = True
        Me.radSQL.Location = New System.Drawing.Point(13, 4)
        Me.radSQL.Name = "radSQL"
        Me.radSQL.Size = New System.Drawing.Size(99, 17)
        Me.radSQL.TabIndex = 0
        Me.radSQL.TabStop = True
        Me.radSQL.Text = "MS SQL Server"
        Me.radSQL.UseVisualStyleBackColor = True
        '
        'radAcc
        '
        Me.radAcc.AutoSize = True
        Me.radAcc.Location = New System.Drawing.Point(118, 3)
        Me.radAcc.Name = "radAcc"
        Me.radAcc.Size = New System.Drawing.Size(60, 17)
        Me.radAcc.TabIndex = 1
        Me.radAcc.TabStop = True
        Me.radAcc.Text = "Access"
        Me.radAcc.UseVisualStyleBackColor = True
        '
        'txtServer
        '
        Me.txtServer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtServer.Location = New System.Drawing.Point(118, 27)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(148, 20)
        Me.txtServer.TabIndex = 2
        '
        'lblServer
        '
        Me.lblServer.AutoSize = True
        Me.lblServer.Location = New System.Drawing.Point(13, 29)
        Me.lblServer.Name = "lblServer"
        Me.lblServer.Size = New System.Drawing.Size(67, 13)
        Me.lblServer.TabIndex = 3
        Me.lblServer.Text = "Servername:"
        '
        'lblDBName
        '
        Me.lblDBName.AutoSize = True
        Me.lblDBName.Location = New System.Drawing.Point(13, 55)
        Me.lblDBName.Name = "lblDBName"
        Me.lblDBName.Size = New System.Drawing.Size(89, 13)
        Me.lblDBName.TabIndex = 5
        Me.lblDBName.Text = "Datenbankname:"
        '
        'txtDBName
        '
        Me.txtDBName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDBName.Location = New System.Drawing.Point(118, 53)
        Me.txtDBName.Name = "txtDBName"
        Me.txtDBName.Size = New System.Drawing.Size(148, 20)
        Me.txtDBName.TabIndex = 4
        '
        'lblDBUser
        '
        Me.lblDBUser.AutoSize = True
        Me.lblDBUser.Location = New System.Drawing.Point(13, 81)
        Me.lblDBUser.Name = "lblDBUser"
        Me.lblDBUser.Size = New System.Drawing.Size(78, 13)
        Me.lblDBUser.TabIndex = 7
        Me.lblDBUser.Text = "Benutzername:"
        '
        'txtDBUSer
        '
        Me.txtDBUSer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDBUSer.Location = New System.Drawing.Point(118, 79)
        Me.txtDBUSer.Name = "txtDBUSer"
        Me.txtDBUSer.Size = New System.Drawing.Size(148, 20)
        Me.txtDBUSer.TabIndex = 6
        '
        'lblPWD
        '
        Me.lblPWD.AutoSize = True
        Me.lblPWD.Location = New System.Drawing.Point(13, 107)
        Me.lblPWD.Name = "lblPWD"
        Me.lblPWD.Size = New System.Drawing.Size(53, 13)
        Me.lblPWD.TabIndex = 9
        Me.lblPWD.Text = "Passwort:"
        '
        'txtPWD
        '
        Me.txtPWD.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPWD.Location = New System.Drawing.Point(118, 105)
        Me.txtPWD.Name = "txtPWD"
        Me.txtPWD.Size = New System.Drawing.Size(148, 20)
        Me.txtPWD.TabIndex = 8
        '
        'chkNT
        '
        Me.chkNT.AutoSize = True
        Me.chkNT.Location = New System.Drawing.Point(16, 131)
        Me.chkNT.Name = "chkNT"
        Me.chkNT.Size = New System.Drawing.Size(117, 17)
        Me.chkNT.TabIndex = 10
        Me.chkNT.Text = "NT Authentifikation"
        Me.chkNT.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(89, 167)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(102, 23)
        Me.btnOK.TabIndex = 13
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnEsc
        '
        Me.btnEsc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnEsc.Location = New System.Drawing.Point(197, 167)
        Me.btnEsc.Name = "btnEsc"
        Me.btnEsc.Size = New System.Drawing.Size(92, 23)
        Me.btnEsc.TabIndex = 11
        Me.btnEsc.Text = "Abbrechen"
        Me.btnEsc.UseVisualStyleBackColor = True
        '
        'btnDBPath
        '
        Me.btnDBPath.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDBPath.Location = New System.Drawing.Point(267, 53)
        Me.btnDBPath.Name = "btnDBPath"
        Me.btnDBPath.Size = New System.Drawing.Size(21, 20)
        Me.btnDBPath.TabIndex = 14
        Me.btnDBPath.Text = "..."
        Me.btnDBPath.UseVisualStyleBackColor = True
        '
        'frmConParams
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 193)
        Me.Controls.Add(Me.btnDBPath)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnEsc)
        Me.Controls.Add(Me.chkNT)
        Me.Controls.Add(Me.lblPWD)
        Me.Controls.Add(Me.txtPWD)
        Me.Controls.Add(Me.lblDBUser)
        Me.Controls.Add(Me.txtDBUSer)
        Me.Controls.Add(Me.lblDBName)
        Me.Controls.Add(Me.txtDBName)
        Me.Controls.Add(Me.lblServer)
        Me.Controls.Add(Me.txtServer)
        Me.Controls.Add(Me.radAcc)
        Me.Controls.Add(Me.radSQL)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(300, 220)
        Me.Name = "frmConParams"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Datenbank Verbindung"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents radSQL As System.Windows.Forms.RadioButton
    Friend WithEvents radAcc As System.Windows.Forms.RadioButton
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents lblServer As System.Windows.Forms.Label
    Friend WithEvents lblDBName As System.Windows.Forms.Label
    Friend WithEvents txtDBName As System.Windows.Forms.TextBox
    Friend WithEvents lblDBUser As System.Windows.Forms.Label
    Friend WithEvents txtDBUSer As System.Windows.Forms.TextBox
    Friend WithEvents lblPWD As System.Windows.Forms.Label
    Friend WithEvents txtPWD As System.Windows.Forms.TextBox
    Friend WithEvents chkNT As System.Windows.Forms.CheckBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnEsc As System.Windows.Forms.Button
    Friend WithEvents btnDBPath As System.Windows.Forms.Button
End Class
