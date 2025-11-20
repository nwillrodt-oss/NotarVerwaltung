VERSION 5.00
Begin VB.Form frmUserChangePWD 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Kennwort ändern"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "frmUserChangePWD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   5100
   StartUpPosition =   1  'Fenstermitte
   Begin VB.TextBox txtPWDOld 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox txtPWDNew2 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox txtPWDNew1 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblPWDOld 
      Caption         =   "Altes Kennwort"
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblUsername 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblText 
      Caption         =   "Kennwort ändern für Benutzer"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblPWDNew2 
      Caption         =   "Kennwort wiederholung"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblPWDNew1 
      Caption         =   "Neues Kennwort"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "frmUserChangePWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULNAME = "frmChnagePWD"
Public bCancel As Boolean

Private Sub cmdEsc_Click()
    bCancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    Me.cmdOK.Enabled = False
    'Me.chkNextLoginChange.Value = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 And Shift = 0 Then   ' ESC
        ' Form Schliessen ohne speichern
        bCancel = True
        Me.Hide
    End If
End Sub


Private Sub txtPWDNew1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 And Shift = 0 Then   ' ESC
        ' Form Schliessen ohne speichern
        bCancel = True
        Me.Hide
    End If
    
    If KeyCode = 13 And Shift = 0 Then ' Enter
        txtPWDNew2.SetFocus     ' focus aus pwd wiederholug
    End If
    
End Sub

Private Sub txtPWDNew2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 And Shift = 0 Then ' Enter
        ' Prüfen of eingabe erfolgt
        If txtPWDNew2 = "" Then Exit Sub
       
    End If
    
End Sub

Private Function InsertOK() As Boolean
        ' Prüfen ob Pwd gleich
        If txtPWDNew1.Text = txtPWDNew2.Text Then
            cmdOK.SetFocus          ' focus auf Ok
            InsertOK = True
        Else
            ' Meldung PWD stimmen nicht überein
            
            ' txt eingabe löschen ?
        
            txtPWDNew1.SetFocus     ' Fokus auf pwd 1
        End If

End Function


Private Sub txtPWDNew2_Change()
    If txtPWDNew2.Text <> "" Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
End Sub

Private Sub txtPWDNew2_LostFocus()
    Call InsertOK
End Sub
