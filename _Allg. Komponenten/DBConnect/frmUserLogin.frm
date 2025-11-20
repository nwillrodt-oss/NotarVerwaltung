VERSION 5.00
Begin VB.Form frmUserLogin 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Anmeldung"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4590
   Icon            =   "frmUserLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4590
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdOK 
      Caption         =   "Anmelden"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdESC 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtMwgaPWD 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox txtMegaUser 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label lblLogintext 
      Caption         =   "Geben Sie Ihren Mega-Benutzernamen und Ihr Mega-Kennwort ein."
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblMegaPWD 
      Caption         =   "Kennwort"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblMegaUser 
      Caption         =   "Benutzername"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MODULNAME = "frmUserLogin"                            ' Modulname für Fehlerbehandlung

Public bCancel As Boolean                                           ' True wenn benutzer abbruch

Private Sub Form_Activate()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    'Call OnTop(Me)                                                  ' Formula nach "Vorne" bringen
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub Form_Load()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call CheckOKEnabled                                             ' Prüfen ob OK noch enabeld
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub cmdOK_Click()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    bCancel = False                                                 ' Kein abbruch
    Me.Hide                                                         ' Form ausblenden
    'Call MegaLogin(txtMegaUser.Text, txtMwgaPWD.Text)
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub cmdEsc_Click()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    bCancel = True                                                  ' Abbruch
    Me.Hide                                                         ' Form ausblenden
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub HandleKeyUpEnter()

    If txtMegaUser.Text = "" Then
        txtMegaUser.SetFocus
    ElseIf txtMegaUser.Text <> "" And txtMwgaPWD.Text = "" Then
        txtMwgaPWD.SetFocus
    ElseIf txtMegaUser.Text <> "" And txtMwgaPWD.Text <> "" Then
        Call CheckOKEnabled
        Call cmdOK_Click
    End If

End Sub

Private Sub CheckOKEnabled()
    
    If txtMegaUser <> "" And txtMwgaPWD <> "" Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
End Sub

Private Sub txtMegaUser_Change()
    Call CheckOKEnabled
End Sub

Private Sub txtMwgaPWD_Change()
    Call CheckOKEnabled
End Sub

Private Sub txtMwgaPWD_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 And Shift = 0 Then ' Enter
        Call HandleKeyUpEnter
        If cmdOK.Enabled Then cmdOK.SetFocus
    End If
End Sub

Private Sub txtMegaUser_KeyUp(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 And Shift = 0 Then ' Enter
        Call HandleKeyUpEnter
        ''txtMwgaPWD.SetFocus
    End If
End Sub
