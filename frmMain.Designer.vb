<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form
    'Inherits frmBase

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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnInfo = New System.Windows.Forms.Button()
        Me.btnEnviroment = New System.Windows.Forms.Button()
        Me.btnReadMe = New System.Windows.Forms.Button()
        Me.btnOptions = New System.Windows.Forms.Button()
        Me.StatusStripMain = New System.Windows.Forms.StatusStrip()
        Me.SpliterMain = New System.Windows.Forms.SplitContainer()
        Me.TVMain = New System.Windows.Forms.TreeView()
        Me.ILTree = New System.Windows.Forms.ImageList(Me.components)
        Me.LVMain = New System.Windows.Forms.ListView()
        Me.MenuMain = New System.Windows.Forms.MenuStrip()
        Me.ILMain = New System.Windows.Forms.ImageList(Me.components)
        Me.TestSubToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMain = New System.Windows.Forms.ToolStrip()
        Me.btnShowNav = New System.Windows.Forms.Button()
        CType(Me.SpliterMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SpliterMain.Panel1.SuspendLayout()
        Me.SpliterMain.Panel2.SuspendLayout()
        Me.SpliterMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(665, 97)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(72, 28)
        Me.btnExit.TabIndex = 0
        Me.btnExit.Text = "Beenden"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnInfo
        '
        Me.btnInfo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnInfo.Location = New System.Drawing.Point(664, 135)
        Me.btnInfo.Name = "btnInfo"
        Me.btnInfo.Size = New System.Drawing.Size(72, 31)
        Me.btnInfo.TabIndex = 1
        Me.btnInfo.Text = "Info"
        Me.btnInfo.UseVisualStyleBackColor = True
        '
        'btnEnviroment
        '
        Me.btnEnviroment.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnEnviroment.Location = New System.Drawing.Point(667, 176)
        Me.btnEnviroment.Name = "btnEnviroment"
        Me.btnEnviroment.Size = New System.Drawing.Size(68, 27)
        Me.btnEnviroment.TabIndex = 2
        Me.btnEnviroment.Text = "Enviroment"
        Me.btnEnviroment.UseVisualStyleBackColor = True
        '
        'btnReadMe
        '
        Me.btnReadMe.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnReadMe.Location = New System.Drawing.Point(664, 213)
        Me.btnReadMe.Name = "btnReadMe"
        Me.btnReadMe.Size = New System.Drawing.Size(70, 31)
        Me.btnReadMe.TabIndex = 3
        Me.btnReadMe.Text = "ReadMe"
        Me.btnReadMe.UseVisualStyleBackColor = True
        '
        'btnOptions
        '
        Me.btnOptions.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOptions.Location = New System.Drawing.Point(665, 256)
        Me.btnOptions.Name = "btnOptions"
        Me.btnOptions.Size = New System.Drawing.Size(68, 28)
        Me.btnOptions.TabIndex = 4
        Me.btnOptions.Text = "Options Liste"
        Me.btnOptions.UseVisualStyleBackColor = True
        '
        'StatusStripMain
        '
        Me.StatusStripMain.Location = New System.Drawing.Point(0, 387)
        Me.StatusStripMain.Name = "StatusStripMain"
        Me.StatusStripMain.Size = New System.Drawing.Size(738, 22)
        Me.StatusStripMain.TabIndex = 5
        '
        'SpliterMain
        '
        Me.SpliterMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SpliterMain.Location = New System.Drawing.Point(0, 67)
        Me.SpliterMain.Name = "SpliterMain"
        '
        'SpliterMain.Panel1
        '
        Me.SpliterMain.Panel1.Controls.Add(Me.TVMain)
        '
        'SpliterMain.Panel2
        '
        Me.SpliterMain.Panel2.Controls.Add(Me.LVMain)
        Me.SpliterMain.Size = New System.Drawing.Size(659, 317)
        Me.SpliterMain.SplitterDistance = 223
        Me.SpliterMain.TabIndex = 6
        '
        'TVMain
        '
        Me.TVMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TVMain.ImageIndex = 0
        Me.TVMain.ImageList = Me.ILTree
        Me.TVMain.Location = New System.Drawing.Point(0, 0)
        Me.TVMain.Name = "TVMain"
        Me.TVMain.SelectedImageIndex = 0
        Me.TVMain.Size = New System.Drawing.Size(221, 317)
        Me.TVMain.TabIndex = 0
        '
        'ILTree
        '
        Me.ILTree.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ILTree.ImageSize = New System.Drawing.Size(16, 16)
        Me.ILTree.TransparentColor = System.Drawing.Color.Transparent
        '
        'LVMain
        '
        Me.LVMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LVMain.FullRowSelect = True
        Me.LVMain.Location = New System.Drawing.Point(0, 0)
        Me.LVMain.Name = "LVMain"
        Me.LVMain.Size = New System.Drawing.Size(430, 317)
        Me.LVMain.TabIndex = 0
        Me.LVMain.UseCompatibleStateImageBehavior = False
        Me.LVMain.View = System.Windows.Forms.View.Details
        '
        'MenuMain
        '
        Me.MenuMain.Location = New System.Drawing.Point(0, 0)
        Me.MenuMain.Name = "MenuMain"
        Me.MenuMain.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional
        Me.MenuMain.Size = New System.Drawing.Size(738, 24)
        Me.MenuMain.TabIndex = 7
        Me.MenuMain.Text = "MenuMain"
        '
        'ILMain
        '
        Me.ILMain.ImageStream = CType(resources.GetObject("ILMain.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ILMain.TransparentColor = System.Drawing.Color.Transparent
        Me.ILMain.Images.SetKeyName(0, "Forward.ico")
        Me.ILMain.Images.SetKeyName(1, "Back.ico")
        Me.ILMain.Images.SetKeyName(2, "Help Light.ico")
        Me.ILMain.Images.SetKeyName(3, "Info Light.ico")
        Me.ILMain.Images.SetKeyName(4, "Stop.png")
        Me.ILMain.Images.SetKeyName(5, "Notes.ico")
        Me.ILMain.Images.SetKeyName(6, "gear-steel.png")
        Me.ILMain.Images.SetKeyName(7, "boss-icon.png")
        Me.ILMain.Images.SetKeyName(8, "engineer-icon.png")
        Me.ILMain.Images.SetKeyName(9, "user-icon.png")
        Me.ILMain.Images.SetKeyName(10, "user-group-icon.png")
        Me.ILMain.Images.SetKeyName(11, "users-icon.png")
        Me.ILMain.Images.SetKeyName(12, "statistik.ico")
        Me.ILMain.Images.SetKeyName(13, "23.png")
        Me.ILMain.Images.SetKeyName(14, "22.png")
        '
        'TestSubToolStripMenuItem
        '
        Me.TestSubToolStripMenuItem.Name = "TestSubToolStripMenuItem"
        Me.TestSubToolStripMenuItem.Size = New System.Drawing.Size(32, 19)
        '
        'ToolStripMain
        '
        Me.ToolStripMain.AutoSize = False
        Me.ToolStripMain.CanOverflow = False
        Me.ToolStripMain.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.ToolStripMain.Location = New System.Drawing.Point(0, 24)
        Me.ToolStripMain.Name = "ToolStripMain"
        Me.ToolStripMain.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.ToolStripMain.Size = New System.Drawing.Size(738, 40)
        Me.ToolStripMain.TabIndex = 8
        Me.ToolStripMain.Text = "ToolStrip1"
        '
        'btnShowNav
        '
        Me.btnShowNav.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnShowNav.Location = New System.Drawing.Point(666, 290)
        Me.btnShowNav.Name = "btnShowNav"
        Me.btnShowNav.Size = New System.Drawing.Size(68, 28)
        Me.btnShowNav.TabIndex = 9
        Me.btnShowNav.Text = "Navigation"
        Me.btnShowNav.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(738, 409)
        Me.Controls.Add(Me.btnShowNav)
        Me.Controls.Add(Me.ToolStripMain)
        Me.Controls.Add(Me.SpliterMain)
        Me.Controls.Add(Me.StatusStripMain)
        Me.Controls.Add(Me.MenuMain)
        Me.Controls.Add(Me.btnOptions)
        Me.Controls.Add(Me.btnReadMe)
        Me.Controls.Add(Me.btnEnviroment)
        Me.Controls.Add(Me.btnInfo)
        Me.Controls.Add(Me.btnExit)
        Me.MainMenuStrip = Me.MenuMain
        Me.MinimumSize = New System.Drawing.Size(550, 350)
        Me.Name = "frmMain"
        Me.Text = "frmMain"
        Me.SpliterMain.Panel1.ResumeLayout(False)
        Me.SpliterMain.Panel2.ResumeLayout(False)
        CType(Me.SpliterMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SpliterMain.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnInfo As System.Windows.Forms.Button
    Friend WithEvents btnEnviroment As System.Windows.Forms.Button
    Friend WithEvents btnReadMe As System.Windows.Forms.Button
    Friend WithEvents btnOptions As System.Windows.Forms.Button
    Friend WithEvents StatusStripMain As System.Windows.Forms.StatusStrip
    Friend WithEvents SpliterMain As System.Windows.Forms.SplitContainer
    Friend WithEvents TVMain As System.Windows.Forms.TreeView
    Friend WithEvents LVMain As System.Windows.Forms.ListView
    Friend WithEvents MenuMain As System.Windows.Forms.MenuStrip
    Friend WithEvents ILMain As System.Windows.Forms.ImageList
    Friend WithEvents TestSubToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMain As System.Windows.Forms.ToolStrip
    Friend WithEvents ILTree As System.Windows.Forms.ImageList
    Friend WithEvents btnShowNav As System.Windows.Forms.Button

End Class
