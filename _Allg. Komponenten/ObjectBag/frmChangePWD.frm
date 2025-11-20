VERSION 5.00
Begin VB.Form frmChangePWD 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Kennwort zurücksetzen"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   Icon            =   "frmChangePWD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4740
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CheckBox chkNextLoginChange 
      Caption         =   "Benutzer muß das Kennwort bei der nächsten Anmeldung ändern"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   4575
   End
   Begin VB.TextBox txtPWDNew2 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      ToolTipText     =   "Max. 8 Zeichen"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox txtPWDNew1 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Max. 8 Zeichen"
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblUsername 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblText 
      Caption         =   "Kennwort ändern für Benutzer"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblPWDNew2 
      Caption         =   "Kennwort wiederholung"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblPWDNew1 
      Caption         =   "Neues Kennwort"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "frmChangePWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MODULNAME = "frmChnagePWD"                            ' Modulname für Fehlerbehandlung

Private objError As Object                                          ' Error object
Private objObjectBag As Object

Public bCancel As Boolean                                           ' Abbruch duch den Benutzer

Public Function InitObjectBag(objOb As Object)

    Set objObjectBag = objOb
    
    If Not objObjectBag Is Nothing Then
        Set objError = objObjectBag.GetErrorObj
    End If
End Function
                                                                    ' *****************************************
                                                                    ' Button Events
Private Sub cmdEsc_Click()
On Error Resume Next
    bCancel = True                                                  ' Abbruch duch den Benutzer
    Me.Hide
End Sub

Private Sub cmdOK_Click()
On Error Resume Next
    If InsertOK Then Me.Hide                                        ' Form schliessen
End Sub

Private Sub Form_Load()
On Error Resume Next
    Me.cmdOK.Enabled = False                                        ' OK erstmal Disablen
    Me.chkNextLoginChange.Value = False                             '
End Sub
                                                                    ' *****************************************
                                                                    ' Key Events
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 27 And Shift = 0 Then                              ' ESC
        bCancel = True                                              ' Abbruch duch den Benutzer
        Me.Hide                                                     ' Form Schliessen ohne speichern
    End If
End Sub

Private Sub txtPWDNew1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 27 And Shift = 0 Then                              ' ESC
        bCancel = True                                              ' Abbruch duch den Benutzer
        Me.Hide                                                     ' Form Schliessen ohne speichern
    End If
    If KeyCode = 13 And Shift = 0 Then                              ' Enter
        txtPWDNew2.SetFocus                                         ' focus aus pwd wiederholug
    End If
End Sub

Private Sub txtPWDNew2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 27 And Shift = 0 Then                              ' ESC
        bCancel = True                                              ' Abbruch duch den Benutzer
        Me.Hide                                                     ' Form Schliessen ohne speichern
    End If
    If KeyCode = 13 And Shift = 0 Then                              ' Enter
        If txtPWDNew2 = "" Then Exit Sub                            ' Prüfen ob eingabe erfolgt
       
    End If
    
End Sub
                                                                    ' *****************************************
                                                                    ' Change Events
Private Sub txtPWDNew2_Change()
On Error Resume Next
    Call CheckOK
End Sub

Private Sub txtPWDNew1_Change()
On Error Resume Next
    Call CheckOK
End Sub

Private Sub txtPWDNew2_LostFocus()
On Error Resume Next
    If InsertOK Then Call CheckOK
End Sub

Private Sub txtPWDNew1_LostFocus()
On Error Resume Next
    If InsertOK Then Call CheckOK
End Sub
                                                                    ' *****************************************
                                                                    ' Hilfs Funktionen
Private Function InsertOK() As Boolean
        
    Dim szMsg As String                                             ' Message Text
    Dim szTitle As String                                           ' Meldungstitel
    
On Error Resume Next
    szMsg = "Die Kennwörter müssen übereinstimmen!"
    szTitle = "Kennwöter nicht gleich"
    If txtPWDNew1.Text = "" Then Exit Function                      ' Kein Kennwort -> fertig
    If txtPWDNew2.Text = "" Then Exit Function                      ' Kein Kennwort -> fertig
    
    If txtPWDNew1.Text = txtPWDNew2.Text Then                       ' Prüfen ob Pwd gleich
        cmdOK.SetFocus                                              ' focus auf Ok
        InsertOK = True                                             ' erfolg zurück
    Else
        Call objError.ShowErrMsg(szMsg, vbInformation, szTitle)     ' Meldung PWD stimmen nicht überein
        txtPWDNew1.Text = ""                                        ' txt eingabe löschen ?
        txtPWDNew2.Text = ""
        txtPWDNew1.SetFocus                                         ' Fokus auf pwd 1
    End If
End Function

Private Sub CheckOK()
On Error Resume Next
    If txtPWDNew1.Text = "" Then Exit Sub                           ' Kein Kennwort -> fertig
    If txtPWDNew2.Text = "" Then Exit Sub                           ' Kein Kennwort -> fertig
    cmdOK.Enabled = True                                            ' OK evtl. Enablen
End Sub

