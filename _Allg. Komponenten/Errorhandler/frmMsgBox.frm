VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMsgBox 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Form1"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   6360
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CheckBox chkIgnoreErr 
      Caption         =   "Diese Meldung in Zukunft nicht mehr anzeigen"
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   5175
   End
   Begin VB.CommandButton cmdRetry 
      Caption         =   "&Retry"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ILErrorFrmIcons 
      Left            =   840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsgBox.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsgBox.frx":1CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsgBox.frx":39B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicError 
      BorderStyle     =   0  'Kein
      Height          =   735
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.ImageList ILError 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsgBox.frx":568E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsgBox.frx":7368
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsgBox.frx":7CEC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "Details >>"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Image ImageError 
      Height          =   855
      Left            =   120
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblDetails 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "details"
      Height          =   615
      Left            =   1080
      TabIndex        =   6
      Top             =   240
      Width           =   5175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMSG 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   855
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   5160
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                                         ' Variaben Deklaration erzwingen
Option Compare Text                                                     ' Sortierreihenfolge festlegen
Private Const MODULNAME = "frmMsgBox"                                   ' Modulname für Fehlerbehandlung

Const IMG_CRITICAL = 1                                                  ' Image Index für Critical
Const IMG_QUESTION = 2                                                  ' Image index für Question
Const IMG_EXCLAMATION = 3                                               ' Image index für Exclamation
Const IMG_INFO = 3                                                      ' Image Index für Information
Private DetailtextTop As Integer
Private bDetailsVisible As Boolean                                      ' Flag True wenn Detais Sichtbar

Public result  As Integer                                               ' Result analog msgresult
Public bIgnorErr As Boolean                                             ' Ob die Meldung in zukunft ignoriert werden soll
Private cmdOKResult As Integer                                          ' OK Ergebnis
Private cmdEscResult As Integer                                         ' Abbrechnen Ergebnis
Private cmdRetryResult As Integer                                       ' Nochmal Ergebnis

Public Function InitPicture(Picture As VbMsgBoxStyle)
' Initialisiert MSG Picture durch logischen Maske über den Picture Wert
    Dim test                                                            ' Ergebnis der Logischen Maske (AND)
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    test = vbCritical And Picture                                       ' Picture Wert mit 16 Maskieren
    If test = vbCritical Then                                           ' Wenn das Erg = vbCritical = 16
        ImageError.Picture = Me.ILError.ListImages(IMG_CRITICAL).Picture ' Critical Image setzen
        Me.Icon = Me.ILErrorFrmIcons.ListImages(IMG_CRITICAL).Picture   ' Critical Icon setzen
    End If
    test = vbQuestion And Picture                                       ' Picture Wert mit 32 Maskieren
    If test = vbQuestion Then                                           ' Wenn das Erg = vbQuestion = 32
        ImageError.Picture = Me.ILError.ListImages(IMG_QUESTION).Picture ' Question Image setzen
        Me.Icon = Me.ILErrorFrmIcons.ListImages(IMG_QUESTION).Picture   ' Question Icon setzen
    End If
    test = vbExclamation And Picture                                    ' Picture Wert mit 48 Maskieren
    If test = vbExclamation Then                                        ' Wenn das Erg = vbExclamation = 48
        ImageError.Picture = Me.ILError.ListImages(IMG_EXCLAMATION).Picture ' Exclamation Image setzen
        Me.Icon = Me.ILErrorFrmIcons.ListImages(IMG_EXCLAMATION).Picture ' Exclamation Icon setzen
    End If
    test = vbInformation And Picture                                    ' Picture Wert mit 64 Maskieren
    If test = vbInformation Then                                        ' Wenn das Erg = vbInformation = 64
        ImageError.Picture = Me.ILError.ListImages(IMG_INFO).Picture    ' Information Image setzen
        Me.Icon = Me.ILErrorFrmIcons.ListImages(IMG_INFO).Picture       ' Information Icon setzen
    End If
    Err.Clear                                                           ' Evtl Error Clearen
End Function

Public Function InitButtons(Buttons As VbMsgBoxStyle)
' Initialisiert MSG Buttons durch logischen Maske über den Buttons Wert
    Dim test                                                            ' Ergebnis der Logischen Maske (AND)
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    test = vbOKOnly And Buttons                                         ' Buttons Wert mit 0 Maskieren
    If test = vbOKOnly Then                                             ' Wenn das Erg = vbOKOnly = 0
        cmdOK.Visible = True                                            ' OK Button Sichtbar
        cmdOKResult = vbOK                                              ' OK Ergebnis Festlegen  =1
        cmdEsc.Visible = False                                          ' Abbrechen Ausblenden
        cmdRetry.Visible = False                                        ' Nochmal Ausblenden
    End If
    test = vbOKCancel And Buttons                                       ' Buttons Wert mit 1 Maskieren
     If test = vbOKCancel Then                                          ' Wenn das Erg = vbOKCancel = 1
        cmdOK.Visible = True                                            ' OK Button Sichtbar
        cmdOKResult = vbOK                                              ' OK Ergebnis Festlegen = 1
        cmdEsc.Visible = True                                           ' Esc Sichtbar
        cmdEscResult = vbCancel                                         ' Esc Ergebnis Festlegen  =2
        cmdRetry.Visible = False                                        ' Nochmal Ausblenden
    End If
    test = vbAbortRetryIgnore And Buttons                               ' Buttons Wert mit 2 Maskieren
    If test = vbAbortRetryIgnore Then                                   ' Wenn das Erg = vbAbortRetryIgnore = 2
        cmdOK.Caption = "Abbrechen"                                     ' OK Caption auf Abbrechen ändern
        cmdOKResult = vbAbort                                           ' OK Ergebnis Festlegen = 3
        cmdOK.Visible = True                                            ' OK Button Sichtbar
        cmdEsc.Caption = "Wiederholen"                                  ' Esc Caption auf Wiederholen ändern
        cmdEscResult = vbRetry                                          ' Esc Ergebnis Festlegen = 4
        cmdEsc.Visible = True                                           ' Esc Button Sichtbar
        cmdRetry.Caption = "Ignorieren"                                 ' Nochmal Caption auf Abbrechen ändern
        cmdRetryResult = vbIgnore                                       ' Nochmal Ergebnis Festlegen = 5
        cmdRetry.Visible = True                                         ' Nochmal Button Sichtbar
    End If
    test = vbYesNoCancel And Buttons                                    ' Buttons Wert mit 3 Maskieren
    If test = vbYesNoCancel Then                                        ' Wenn das Erg = vbYesNoCancel = 3
        cmdOK.Caption = "Ja"                                            ' OK Caption auf Ja ändern
        cmdOKResult = vbYes                                             ' OK Ergebnis Festlegen = 6
        cmdOK.Visible = True                                            ' OK Button Sichtbar
        cmdEsc.Caption = "Nein"                                         ' Esc Caption auf Nein ändern
        cmdEscResult = vbNo                                             ' Esc Ergebnis Festlegen = 7
        cmdEsc.Visible = True                                           ' Esc Button Sichtbar
        cmdRetry.Caption = "Abbrechen"                                  ' Esc Caption auf Abbrechen ändern
        cmdRetryResult = vbCancel                                       ' Nochmal Ergebnis Festlegen  =2
        cmdRetry.Visible = True                                         ' Nochmal Button Sichtbar
    End If
    test = vbYesNo And Buttons                                          ' Buttons Wert mit 4 Maskieren
     If test = vbYesNo Then                                             ' Wenn das Erg = vbYesNo = 4
        cmdOK.Caption = "Ja"                                            ' OK Caption auf Ja ändern
        cmdOKResult = vbYes                                             ' OK Ergebnis Festlegen = 6
        cmdOK.Visible = True                                            ' OK Button Sichtbar
        cmdEsc.Caption = "Nein"                                         ' Esc Caption auf Nein ändern
        cmdEscResult = vbNo                                             ' Esc Ergebnis Festlegen = 7
        cmdEsc.Visible = True                                           ' Esc Button Sichtbar
        cmdRetry.Visible = False                                        ' Nochmal Ausblenden
    End If
    test = vbRetryCancel And Buttons                                    ' Buttons Wert mit 5 Maskieren
    If test = vbRetryCancel Then                                        ' Wenn das Erg = vbRetryCancel = 5
        cmdOK.Caption = "Wiederholen"                                   ' OK Caption auf Wiederholen ändern
        cmdOKResult = vbRetry                                           ' OK Ergebnis Festlegen = 4
        cmdOK.Visible = True                                            ' OK Button Sichtbar
        cmdEsc.Caption = "Abbrechen"                                    ' Esc Caption auf Abbrechen ändern
        cmdEscResult = vbCancel                                         ' Esc Ergebnis Festlegen = 2
        cmdEsc.Visible = True                                           ' Esc Button Sichtbar
        cmdRetry.Visible = False                                        ' Nochmal Ausblenden
    End If
    test = vbDefaultButton1 And Buttons                                 ' Buttons Wert mit 0 Maskieren
    If test = vbDefaultButton1 Then                                     ' Wenn das Erg = vbDefaultButton1 = 0
        If cmdOK.Visible Then cmdOK.SetFocus                            ' Focus auf OK wenn Sichtbar
    End If
    test = vbDefaultButton2 And Buttons                                 ' Buttons Wert mit 256 Maskieren
    If test = vbDefaultButton2 Then                                     ' Wenn das Erg = vbDefaultButton2 = 256
        If cmdEsc.Visible Then cmdEsc.SetFocus                          ' Focus auf Esc wenn Sichtbar
    End If
'    test = vbDefaultButton3 And Buttons
'    test = vbDefaultButton4 And Buttons
    Err.Clear                                                           ' Evtl Error Clearen
End Function

Private Sub SetButtonFocus(AktCmb As CommandButton, Optional bLeft As Boolean)
    Dim szButtonName As String
    Dim szNextButton As String
    Dim ButtonNr As Integer
    Dim ButtonOffSet As Integer
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If Not bLeft Then
        ButtonOffSet = 1
    Else
        ButtonOffSet = -1
    End If
    szButtonName = AktCmb.Name                                          ' Name des Akt Controls (Button)
    Select Case szButtonName                                            ' Auswerten
    Case "cmdOK"                                                        ' OK
        ButtonNr = 1
    Case "cmdEsc"                                                       ' Abbrechen
        ButtonNr = 2
    Case "cmdRetry"                                                     ' Nochmal
        ButtonNr = 3
    Case "cmdDetails"                                                   ' Details
        ButtonNr = 4
    Case Else                                                           ' Sonst
    
    End Select
    ButtonNr = ButtonNr + ButtonOffSet
    If ButtonNr > 4 Then ButtonNr = 1                                   ' > 4 geht nicht also 1
    If ButtonNr = 0 Then ButtonNr = 4                                   ' Null geht auch nicht also 4
Jump:
    Select Case ButtonNr                                                ' Welchen Button behandeln wir
    Case 1                                                              ' 1: OK
        If cmdOK.Visible = False Or cmdOK.Enabled = False Then
            ButtonNr = ButtonNr + ButtonOffSet
            GoTo Jump                                                   ' und nochmal
        Else
            cmdOK.SetFocus                                              ' Fokus auf OK
        End If
    Case 2                                                              ' 2: Esc
        If cmdEsc.Visible = False Or cmdEsc.Enabled = False Then
            ButtonNr = ButtonNr + ButtonOffSet
            GoTo Jump                                                   ' und nochmal
        Else
            cmdEsc.SetFocus                                             ' Focus auf Abbrechen
        End If
    Case 3                                                              ' 3: Retry
        If cmdRetry.Visible = False Or cmdRetry.Enabled = False Then
            ButtonNr = ButtonNr + ButtonOffSet
            GoTo Jump                                                   ' und nochmal
        Else
            cmdRetry.SetFocus                                           ' Focus auf Nochmal
        End If
    Case 4                                                              ' 4: Details
        If cmdDetails.Visible = False Or cmdDetails.Enabled = False Then
            ButtonNr = ButtonNr + ButtonOffSet
            GoTo Jump                                                   ' und nochmal
        Else
            cmdDetails.SetFocus                                         ' Focus auf Detais
        End If
    End Select
    Err.Clear                                                           ' Evtl. Error clearen
End Sub

Private Sub HandleKeyDown(KeyCode As Integer, Shift As Integer, Optional cmdButton As CommandButton)
' Behandelt KeyDownEvents im Edit Form
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If KeyCode = 38 And Shift = 0 Then                                  ' Pfeil nach oben
        If Not cmdButton Is Nothing Then Call SetButtonFocus(cmdButton, True)
    End If
    If KeyCode = 39 And Shift = 0 Then                                  ' Pfeil nach Rechts
        If Not cmdButton Is Nothing Then Call SetButtonFocus(cmdButton, False)
    End If
    If KeyCode = 40 And Shift = 0 Then                                  ' Pfeil nach unten
        If Not cmdButton Is Nothing Then Call SetButtonFocus(cmdButton, False)
    End If
    If KeyCode = 27 And Shift = 0 Then                                  ' ESC
        'Unload frmEdit                                                  ' Form Schliessen ohne speichern
    End If
    If KeyCode = 83 And Shift = 2 Then                                  ' STGR + S
        'Call frmEdit.cmdUpdate_Click                                    ' Form Speichern
    End If
    'Call HandleKeyDownEdit(Me, KeyCode, Shift)                          ' Spezielle KeyDon Events dieses Forms
    'Call frmParent.HandleGlobalKeyCodes(KeyCode, Shift)                 ' Key Down Events der Anwendung
    Err.Clear                                                           ' Evtl. error clearen
End Sub

Private Sub chkIgnoreErr_Click()
    bIgnorErr = CBool(Me.chkIgnoreErr.Value)
End Sub
                                                                        ' *****************************************
                                                                        ' Button Events
Private Sub cmdDetails_Click()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    bDetailsVisible = Not bDetailsVisible                               ' Sichtbarkeit der Details umschalten
    If bDetailsVisible Then                                             ' Details Sichtbar
        lblDetails.Visible = bDetailsVisible                            ' Details einblenden
        Me.Height = Me.Height + (lblDetails.Height + 50)                ' Button Positionieren
        lblDetails.Top = DetailtextTop                                  ' Top Pos. Festlegen
        cmdDetails.Caption = "Details <<"                               ' Caption anpassen
    Else                                                                ' Sonst
        lblDetails.Visible = bDetailsVisible                            ' Details ausblenden
        Me.Height = Me.Height - (lblDetails.Height + 50)                ' Button Positionieren
        lblDetails.Top = DetailtextTop                                  ' Top Pos. Festlegen
        cmdDetails.Caption = "Details >>"                               ' Caption anpassen
    End If
    Err.Clear                                                           ' Evtl. error clearen
End Sub

Private Sub cmdEsc_Click()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    result = cmdEscResult                                               ' Esc Result zurück
    Me.Hide                                                             ' Form ausblenden
    Err.Clear                                                           ' Evtl. error clearen
End Sub

Private Sub cmdOK_Click()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    result = cmdOKResult                                                ' OK Result zurück
    Me.Hide                                                             ' Form ausblenden
    Unload Me                                                           ' Form entladen
    Err.Clear                                                           ' Evtl. error clearen
End Sub

Private Sub cmdRetry_Click()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    result = cmdRetryResult                                             ' Nochmal Result zurück
    Me.Hide                                                             ' Form ausblenden
    Err.Clear                                                           ' Evtl. error clearen
End Sub

                                                                        ' *****************************************
                                                                        ' Key Events

Private Sub cmdDetails_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(KeyCode, Shift)                                  ' KeyDown behandeln
End Sub

Private Sub cmdEsc_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(KeyCode, Shift)                                  ' KeyDown behandeln
End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(KeyCode, Shift)                                  ' KeyDown behandeln
End Sub

Private Sub cmdRetry_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(KeyCode, Shift)                                  ' KeyDown behandeln
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(KeyCode, Shift)                                  ' KeyDown behandeln
End Sub

Private Sub PicError_KeyDown(KeyCode As Integer, Shift As Integer)
    cmdOK.SetFocus                                                      ' Focus auf OK Button
    Call HandleKeyDown(KeyCode, Shift)                                  ' KeyDown behandeln
End Sub
                                                                        ' *****************************************
                                                                        ' Form Events
Private Sub Form_Activate()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    lblDetails.Visible = bDetailsVisible
    Err.Clear                                                           ' Evtl. error clearen
End Sub

Private Sub Form_Resize()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    DetailtextTop = Me.cmdOK.Top + Me.cmdOK.Height + 50
    Err.Clear                                                           ' Evtl. error clearen
End Sub



