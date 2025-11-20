VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   3555
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer TimerSplash 
      Left            =   2640
      Top             =   1680
   End
   Begin VB.Label lblWWW 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmSplash.frx":9F73A
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   4
      Top             =   2520
      Width           =   5415
   End
   Begin VB.Label lblAction 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   5280
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Label lblAppTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   945
   End
   Begin VB.Image ImageSplash 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   3495
      Left            =   0
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   7155
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MODULNAME = "frmSplash"                               ' Modulname für Fehlerbehandlung
    
Public objObjectBag As Object                                       ' ObjectBag object

Private Sub Form_Load()
On Error Resume Next                                                ' Feherbehandlung deaktivieren
    Me.lblCopyright.Caption = objObjectBag.GetCopyright()           ' Copyright anzeigen
    Me.lblVersion.Caption = "Version " & objObjectBag.GetMajor() _
            & "." & objObjectBag.GetMinor() & "." _
            & objObjectBag.GetRevision()                            ' Version anzeigen
    Me.lblAppTitle.Caption = objObjectBag.GetAppTitle()             ' Anwendungstitel
    Me.lblWWW.Caption = objObjectBag.GetWWW                         ' Internet adresse
    If Me.lblWWW.Caption <> "" Then
        Me.lblWWW.Visible = True
    Else
        Me.lblWWW.Visible = False
    End If
    Err.Clear                                                       ' Evtl Error clearen
End Sub

Private Sub lblWWW_Click()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call objObjectBag.FolowLink(Me, lblWWW.Caption)                 ' Link Folgen
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub TimerSplash_Timer()
On Error Resume Next
    ' Erster Timer blendet frm Main ein
    ' Zweiter Timer blendet splash aus
    'If frmMDIMain.Visible Then
        TimerSplash.Enabled = False
'        If Me.Visible Then Me.Hide
    'Else
    '    Call OpenMainForm
    'End If
    Err.Clear
End Sub
