VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame FrameFile 
      Caption         =   "Datei"
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   6495
      Begin VB.CheckBox chkFileForceExist 
         Caption         =   "Erzwingen"
         Height          =   255
         Left            =   3240
         TabIndex        =   15
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdFileDel 
         Caption         =   "Löschen"
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdFileCreate 
         Caption         =   "Anlegen"
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdFileExist 
         Caption         =   "Existenzprüfen"
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtFile 
         Height          =   375
         Left            =   75
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblResultExistFile 
         Caption         =   "-"
         Height          =   315
         Left            =   4800
         TabIndex        =   21
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblResultCreateFile 
         Caption         =   "-"
         Height          =   315
         Left            =   4800
         TabIndex        =   20
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblResultDelFile 
         Caption         =   "-"
         Height          =   315
         Left            =   4800
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame FrameOrdner 
      Caption         =   "txtTestDir"
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   6495
      Begin VB.CheckBox chkDirForceExist 
         Caption         =   "Erzwingen"
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtDir 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdDirExist 
         Caption         =   "Existenzprüfen"
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdDirCreate 
         Caption         =   "Anlegen"
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdDirDel 
         Caption         =   "Löschen"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblResultDelDir 
         Caption         =   "-"
         Height          =   315
         Left            =   4800
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblResultCreateDir 
         Caption         =   "-"
         Height          =   315
         Left            =   4800
         TabIndex        =   17
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblResultExistDir 
         Caption         =   "-"
         Height          =   315
         Left            =   4800
         TabIndex        =   16
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdTestDirExplorer 
      Caption         =   "Explorer"
      Height          =   315
      Left            =   4680
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdTestDir 
      Caption         =   "..."
      Height          =   315
      Left            =   4320
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txtTestDir 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label lblTestDir 
      Caption         =   "TestOrdner"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objTools As Object


Private Sub cmdDirCreate_Click()
Dim result As Boolean
    Call InitResults
    If txtDir = "" Then Exit Sub
    If Not objTools.CheckLastChar(txtTestDir, "\") Then txtTestDir = txtTestDir & "\"
    result = objTools.FolderExist(txtTestDir & txtDir, True)
    If result Then
        lblResultCreateDir.Caption = "Erfolgreich"
        lblResultCreateDir.ForeColor = &H8000&
    Else
        lblResultCreateDir.Caption = "Gescheitert"
        lblResultCreateDir.ForeColor = &HFF&
    End If
    
End Sub

Private Sub cmdDirExist_Click()
Dim result As Boolean
    Call InitResults
    If txtDir = "" Then Exit Sub
    If Not objTools.CheckLastChar(txtTestDir, "\") Then txtTestDir = txtTestDir & "\"
    result = objTools.FolderExist(txtTestDir & txtDir, CBool(chkDirForceExist.Value))
    If result Then
        lblResultExistDir.Caption = "Erfolgreich"
        lblResultExistDir.ForeColor = &H8000&
    Else
        lblResultExistDir.Caption = "Gescheitert"
        lblResultExistDir.ForeColor = &HFF&
    End If
End Sub

Private Sub cmdFileExist_Click()
Dim result As Boolean
    Call InitResults
    If txtFile = "" Then Exit Sub
    If Not objTools.CheckLastChar(txtTestDir, "\") Then txtTestDir = txtTestDir & "\"
    result = objTools.FileExist(txtTestDir & txtFile, CBool(chkFileForceExist.Value))
    If result Then
        lblResultExistFile.Caption = "Erfolgreich"
        lblResultExistFile.ForeColor = &H8000&
    Else
        lblResultExistFile.Caption = "Gescheitert"
        lblResultExistFile.ForeColor = &HFF&
    End If
End Sub

Private Sub cmdTestDir_Click()
    Call InitResults
    txtTestDir.Text = objTools.ShowDirSelect("", txtTestDir)
    If Not objTools.CheckLastChar(txtTestDir, "\") Then txtTestDir = txtTestDir & "\"
End Sub

Private Sub Form_Load()

    Set objTools = CreateObject("Tools.clsTools")
    If Not objTools Is Nothing Then
    
    End If
    Call InitResults
    
End Sub

Private Sub InitResults()
    lblResultExistDir.Caption = "-"
    lblResultExistDir.ForeColor = &H80000012
    lblResultCreateDir.Caption = "-"
    lblResultCreateDir.ForeColor = &H80000012
    lblResultDelDir.Caption = "-"
    lblResultDelDir.ForeColor = &H80000012
    lblResultDelFile.Caption = "-"
    lblResultDelFile.ForeColor = &H80000012
    lblResultExistFile.Caption = "-"
    lblResultExistFile.ForeColor = &H80000012
    lblResultCreateFile.Caption = "-"
    lblResultCreateFile.ForeColor = &H80000012
End Sub

