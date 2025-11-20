VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame1 
      Caption         =   "Fehlermeldungen"
      Height          =   4695
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   10695
      Begin VB.TextBox Text10 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   4800
         TabIndex        =   35
         Text            =   "Laufende Nase"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   1680
         TabIndex        =   33
         Text            =   "Irgendein Modul halt"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Frame Frame3 
         Caption         =   "Meldungs Art"
         Height          =   615
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   5415
         Begin VB.OptionButton Option7 
            Caption         =   "Meldung"
            Height          =   255
            Left            =   2640
            TabIndex        =   31
            Top             =   240
            Width           =   2175
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Fehler"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   3720
         TabIndex        =   28
         Text            =   "Echt Kölnisch Wasser"
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   1680
         TabIndex        =   27
         Text            =   "4711"
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Meldung Ignorieren anzeigen"
         Height          =   615
         Left            =   8640
         TabIndex        =   24
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Caption         =   "Buttons"
         Height          =   1455
         Left            =   7200
         TabIndex        =   17
         Top             =   240
         Width           =   2655
         Begin VB.OptionButton Option1 
            Caption         =   "OK Only"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton Option2 
            Caption         =   "OK Cancel"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   480
            Width           =   2175
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Abort Retry Ignore"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   720
            Width           =   2055
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Yes No Cancel"
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Yes No"
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   1200
            Width           =   1575
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Warnung"
         Height          =   315
         Left            =   7080
         TabIndex        =   15
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Information"
         Height          =   315
         Left            =   7080
         TabIndex        =   14
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Kritisch"
         Height          =   315
         Left            =   7080
         TabIndex        =   13
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   8640
         TabIndex        =   12
         Top             =   3480
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Frage"
         Height          =   315
         Left            =   7080
         TabIndex        =   11
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Text            =   "Fehlertitel"
         Top             =   1680
         Width           =   5295
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000002&
         Height          =   1095
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   8
         Text            =   "frmTest.frx":0000
         Top             =   3000
         Width           =   5295
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000002&
         Height          =   855
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   5
         Text            =   "frmTest.frx":000B
         Top             =   2040
         Width           =   5295
      End
      Begin VB.Label Label11 
         Caption         =   "Funktionsname"
         Height          =   255
         Left            =   3480
         TabIndex        =   34
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Modulname"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Error Text"
         Height          =   255
         Left            =   2880
         TabIndex        =   26
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Error nr"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label7 
         Height          =   255
         Left            =   8640
         TabIndex        =   23
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Ergebnis"
         Height          =   255
         Left            =   8640
         TabIndex        =   16
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Fehlertitel"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Detailtext"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Fehlertext"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   1455
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Text            =   "C:\ErrTestProt.txt"
      Top             =   480
      Width           =   8895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Text            =   "C:\ErrTestProt.txt"
      Top             =   120
      Width           =   8895
   End
   Begin VB.Label Label2 
      Caption         =   "Fehlerprotokoll"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Protokolldatei"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MODULNAME = "frmTest"

Public objError  As Object                                          ' Glob Error object


Private Sub Form_Load()

    Set objError = CreateObject("Error.clsErrorHandler")
    Call Option6_Click
End Sub

Private Sub Command1_Click()
    ' warnung
    Dim bIgnore As Boolean
    bIgnore = CBool(Check1.Value)
    
    If Option6 = True Then
        Text3.Text = objError.ErrorHandler(Text6.Text, Text10.Text, _
        Text4.Text, "", Text8.Text)
    ElseIf Option7 = True Then
        Text3.Text = objError.ShowErrMsg(Text7.Text, vbExclamation + GetButtonsFromOptionsFrame(), "Warnung - " & Text9.Text, _
            True, Text8.Text, Me, bIgnore)
        MsgBox "Meldung in zukunft ignorieren: " & CStr(bIgnore)
    End If
End Sub

Private Sub Command2_Click()
    ' Information
    Text3.Text = objError.ShowErrMsg(Text7.Text, vbInformation + GetButtonsFromOptionsFrame(), "Information - " & Text9.Text, _
            True, Text8.Text, Me, CBool(Check1.Value))
End Sub

Private Sub Command3_Click()
    ' Kritisch
    Text3.Text = objError.ShowErrMsg(Text7.Text, vbCritical + GetButtonsFromOptionsFrame(), "Kritisch - " & Text9.Text, _
            True, Text8.Text, Me, CBool(Check1.Value))
End Sub

Private Sub Command4_Click()
    ' Frage
    Text3.Text = objError.ShowErrMsg(Text7.Text, vbQuestion + GetButtonsFromOptionsFrame(), "Frage - " & Text9.Text, _
            True, Text8.Text, Me, CBool(Check1.Value))
End Sub

Private Function GetButtonsFromOptionsFrame()

    Dim Button As Integer
    
    If Me.Option1.Value Then
        Button = vbOKOnly
    ElseIf Me.Option2.Value Then
        Button = vbOKCancel
    ElseIf Me.Option3.Value Then
        Button = vbAbortRetryIgnore
    ElseIf Me.Option4.Value Then
        Button = vbYesNoCancel
    ElseIf Me.Option5.Value Then
        Button = vbYesNo
    End If
    
    GetButtonsFromOptionsFrame = Button
End Function





Private Sub Option6_Click()
    Text7.Enabled = False
'    Text8.Enabled = False
    Text9.Enabled = False
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Text10.Enabled = True
End Sub

Private Sub Option7_Click()
    Text7.Enabled = True
'    Text8.Enabled = True
    Text9.Enabled = True
    Text4.Enabled = False
    Text5.Enabled = False
    Text6.Enabled = False
    Text10.Enabled = False
End Sub

Private Sub Text3_Change()
    If Text2.Text = "" Then
        Label7.Caption = ""
        Exit Sub
    End If
    Select Case CStr(Text3.Text)
    Case "1"
        Label7.Caption = "OK"
    Case "2"
        Label7.Caption = "Abbrechen"
    Case "3"
        Label7.Caption = "Abbruch"
    Case "4"
        Label7.Caption = "Wiederholen"
    Case "5"
        Label7.Caption = "Ignorieren"
    Case "6"
        Label7.Caption = "Ja"
    Case "7"
        Label7.Caption = "Nein"
    Case Else
        Label7.Caption = ""
    End Select
End Sub

