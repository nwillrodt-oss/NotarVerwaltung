VERSION 5.00
Begin VB.Form frmUserLogin 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Anmeldung"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4260
   Icon            =   "frmUserLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdOK 
      Caption         =   "Anmelden"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdESC 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtMwgaPWD 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox txtMegaUser 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Geben Sie Ihren Mega-Benutzernamen und Ihr Mega-Kennwort ein."
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblMegaPWD 
      Caption         =   "Kennwort"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblMegaUser 
      Caption         =   "Benutzername"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bCancel As Boolean ' True wenn benutzer abbruch

Private Const MODULNAME = "frmUserLogin"

Private Sub cmdOK_Click()
    bCancel = False
    Me.Hide
    'Call MegaLogin(txtMegaUser.Text, txtMwgaPWD.Text)
End Sub

Private Sub Form_Load()
    Call CheckOKEnabled
End Sub

Private Sub cmdESC_Click()
    bCancel = True
    Me.Hide
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
