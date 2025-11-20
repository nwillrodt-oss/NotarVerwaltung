VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "DeCrypter "
   ClientHeight    =   870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdExit 
      Caption         =   "Beenden"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdDeCrypt 
      Caption         =   "Entschlüsseln"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdCrypt 
      Caption         =   "Verschlüsseln"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtPhrase 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label lblPhrase 
      Caption         =   "Phrase: "
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                                     ' Variaben Deklaration erzwingen
Option Compare Text                                                 ' Sortierreihenfolge festlegen

Const MODULNAME = "frmMain"                                         ' Modulname für Fehlerbehandlung

Private Sub cmdCrypt_Click()
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If Me.txtPhrase = "" Then GoTo exithandler                      ' Nix zum VerEntschhlüsseln -> Fertig
    Me.txtPhrase = objTools.Crypt(Me.txtPhrase, True)               ' Phrase Verschlüsseln
exithandler:
On Error Resume Next                                                ' Hier keine Fehler Mehr
Exit Sub                                                            ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "cmdCrypt_Click", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Sub

Private Sub cmdDeCrypt_Click()
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If Me.txtPhrase = "" Then GoTo exithandler                      ' Nix zum VerEntschhlüsseln -> Fertig
    Me.txtPhrase = objTools.Crypt(Me.txtPhrase, False)              ' Phrase Entschlüsseln
exithandler:
On Error Resume Next                                                ' Hier keine Fehler Mehr
Exit Sub                                                            ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "cmdDeCrypt_Click", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Sub

Private Sub cmdExit_Click()
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    Call Unload(Me)                                                 ' Form schliessen
    Call AppExit                                                    ' Application beenden
exithandler:
On Error Resume Next                                                ' Hier keine Fehler Mehr
Exit Sub                                                            ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "cmdExit_Click", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Sub
