VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVorlagenSelect 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Vorlagen auswahl"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdESC 
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   5040
      Width           =   1455
   End
   Begin MSComctlLib.ListView LVVorlagen 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   8705
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmVorlagenSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                                     ' Variaben Deklaration erzwingen
Private Const MODULNAME = "modMain"                                 ' Modulname für Fehlerbehandlung

Private szVorlagenVerz As String                                    ' Aktuelles Vorlagen verz.
Public bCancel As Boolean                                           ' Abbruch bedingung
' MW 26.08.11 {
Private bDOTXPossible As Boolean                                    ' Dotx Vorlagen (Word 2007) zulässig
' MW 26.08.11 }
Private Sub Form_Load()
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    bCancel = False                                                 ' Abbruchbedingung initialisieren
    cmdOK.Enabled = False                                           ' OK Button erstmal deakt.
    szVorlagenVerz = objOptions.GetOptionByName(OPTION_TEMPLATES)   ' Vorlagen verz aus Optionen lesen
    bDOTXPossible = objOptions.GetOptionByName(OPTION_DOTX)         ' DotxVorlagen zuslassen?
    If szVorlagenVerz = "" Then szVorlagenVerz = objObjectBag.Getappdir & "Vorlagen"  ' Gegebenenfalls Default Verz. vorschlagen
    Call ListTemplates                                              ' Vorlagen auflisten
    Call CheckSelect                                                ' Prüfen og ein Eintrag ausgewählt
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Err.Clear                                                       ' Evtl. error clearen
Exit Sub                                                            ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "Form_Load", errNr, errDesc)   ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Sub

Private Sub cmdESC_Click()
On Error Resume Next                                                ' Hier keine Fehlerbehandlung
    bCancel = True                                                  ' Abbruchbed. setzen
    Me.Hide                                                         ' Form Ausblenden
    Err.Clear                                                       ' Evtl. error clearen
End Sub

Private Sub cmdOK_Click()
On Error Resume Next                                                ' Hier keine Fehlerbehandlung
    Me.Hide                                                         ' Form Ausblenden
    Err.Clear                                                       ' Evtl. error clearen
End Sub

Private Sub LVVorlagen_Click()
On Error Resume Next                                                ' Hier keine Fehlerbehandlung
    Call CheckSelect
    Err.Clear                                                       ' Evtl. error clearen
End Sub

Private Function CheckSelect() As Boolean
On Error Resume Next                                                ' Hier keine Fehlerbehandlung
    If Not LVVorlagen.SelectedItem Is Nothing Then CheckSelect = True   ' Prüfen og ein Eintrag ausgewählt
    cmdOK.Enabled = CheckSelect                                     ' Abhängig vom ergebnis OK Button einblenden
    Err.Clear                                                       ' Evtl. error clearen
End Function

Private Function ListTemplates()
    Dim objFso As Object                                            ' File System object
    Dim objFile As Object
    Dim objFolder As Object
    Dim szFilename As String                                        ' Akt Dateiname
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    LVVorlagen.ListItems.Clear                                      ' Liste Clearen
    LVVorlagen.ColumnHeaders.Add , , "Vorlagen", LVVorlagen.Width - 100
    Set objFso = CreateObject("Scripting.FileSystemObject")         ' File System object erstellen
    If objFso.FolderExists(szVorlagenVerz) Then                     ' Prüfen ob vorlagen verz. erxistiert
        ' Ausgangsverzeichnis
        Set objFolder = objFso.GetFolder(szVorlagenVerz)            ' Folder öffnen
        ' alle Dateien im Stammverzeichnis C: anzeigen
        For Each objFile In objFolder.Files                         ' Alle Dateien durchgehen
' MW 26.08.11 {
            szFilename = objFile.Name
'            If Not Left(objFile.Name, 2) = "~$" Then                ' Keine Temp dateien berücksichtigen
'                LVVorlagen.ListItems.Add , , objFile.Name
'            End If
            If Left(objFile.Name, 2) = "~$" Then GoTo NextFile      ' Keine Temp dateien berücksichtigen
            If Not bDOTXPossible Then                               ' Wenn keine DocxVorlagen zulässig
                If UCase(Right(objFile.Name, 4)) <> ".DOT" Then _
                        GoTo NextFile                               ' Nur dots berücksichtigen
            Else                                                    ' Sonst
                If UCase(Right(objFile.Name, 4)) <> ".DOT" _
                        And UCase(Right(objFile.Name, 5)) <> ".DOTX" _
                        Then GoTo NextFile                          ' Nur dots & dotxs  berücksichtigen
            End If
            LVVorlagen.ListItems.Add , , objFile.Name               ' Auflisten
NextFile:
' MW 26.08.11 }
        Next                                                        ' Nächste Datei
    End If
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Err.Clear                                                       ' Evtl. error clearen
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ListTemplates", errNr, errDesc)   ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function
