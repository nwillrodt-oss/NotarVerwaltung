VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Optionen"
   ClientHeight    =   2730
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6660
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6660
   StartUpPosition =   1  'Fenstermitte
   Begin VB.ComboBox cmbOption 
      Height          =   315
      Index           =   0
      Left            =   2280
      TabIndex        =   9
      Top             =   240
      Width           =   4215
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "..."
      Height          =   315
      Index           =   0
      Left            =   6120
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox CheckOption 
      Caption         =   "Check"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6375
   End
   Begin VB.TextBox txtOption 
      Height          =   315
      Index           =   0
      Left            =   2280
      TabIndex        =   5
      Top             =   480
      Width           =   3855
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      Height          =   315
      Index           =   0
      Left            =   6120
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.Frame FrameKategorie 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   6615
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Übernehmen"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblOption 
      Caption         =   "Label"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                                     ' Variaben Deklaration erzwingen
Option Compare Text                                                 ' Sortierreihenfolge festlegen

Private Const MODULNAME = "frmOptions"                              ' Modulname für Fehlerbehandlung

Public bInit As Boolean                                             ' Form wird initialisiert
Public objObjbag As Object                                          ' Objectbag object
Public objOptions As Object
Public MaxDisplayPathLen As Long                                    ' Max Länge für Pfadangaben

Private Sub Form_Load()
On Error Resume Next                                                ' Feherbehandlung deaktivieren
    Set objOptions = objObjectBag.GetOptionsObj                     ' Options Object aus ObjBag holen
    cmdUpdate.Enabled = False                                       ' OK Button ertmal disablen
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub CancelButton_Click()
On Error Resume Next                                                ' Feherbehandlung deaktivieren
    Call Unload(Me)                                                 ' Options Form entladen
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub cmdPath_Click(Index As Integer)
    Dim szTmp As String                                             ' Augewählter Pfad
'On Error Resume Next                                                ' Feherbehandlung deaktivieren
    szTmp = objObjbag.SelectFolder("Wählen sie einen Pfad aus:", txtOption(Index).Text) ' Ordner öffnen Dialog anzeigen
    If szTmp <> "" Then                                             ' Pfad ausgewählt kein abbruch
        cmdPath(Index).Tag = szTmp                                  ' Pfad in button Tag
        txtOption(Index).Text = objTools.GetShortPath(Me, CStr(szTmp), _
                                    MaxDisplayPathLen)              ' Pfad anzeige kürzen
        'txtOption(Index).Text = szTmp                               ' Pfad in Option setzen
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub cmdFile_Click(Index As Integer)
    Dim szTmp As String
    szTmp = objObjectBag.OpenFile("", Me, cmdFile(Index).Tag)       ' Datei Öffne Dialog anzeigen
    'Public Function OpenFile(Filter As String, AktForm As Object, Optional DefPath As String) As String
    If szTmp <> "" Then                                             ' Ist was zurück gekommen
        cmdFile(Index).Tag = szTmp                                  ' Pfad in button Tag
        txtOption(Index).Text = objTools.GetShortPath(Me, CStr(szTmp), _
                                    MaxDisplayPathLen)              ' Pfad anzeige kürzen
        'txtOption(Index).Text = szTmp
    End If
    'txtOption(Index).Text = objObjbag.OpenFile("", txtOption(Index).Text)
End Sub

Private Sub cmdUpdate_Click()                                       ' Speichert die änderungen
On Error Resume Next                                                ' Feherbehandlung deaktivieren
    Call SaveOptionsInArray(Me)                                     ' Options Werte in Optionsarray speichern
    'Call objOptions.SaveOptions                                     ' optionen in Registry Speichern
    Me.cmdUpdate.Enabled = False                                    ' Übernehmen disablen
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub OKButton_Click()                                        ' Speichert evtl. änderungen & schliesst form
On Error Resume Next                                                ' Feherbehandlung deaktivieren
    Call SaveOptionsInArray(Me)                                     ' Options Werte in Optionsarray speichern
    'Call objOptions.SaveOptions                                     ' optionen in Registry Speichern
    Call Unload(Me)                                                 ' Options Form entladen
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub cmbOption_Click(Index As Integer)
On Error Resume Next                                                ' Feherbehandlung deaktivieren
    If bInit Then Exit Sub                                          ' Form wird noch initialisiert -> fertig
    cmdUpdate.Enabled = True                                        ' Übernehmen enablen
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub txtOption_Change(Index As Integer)
On Error Resume Next                                                ' Feherbehandlung deaktivieren
    If bInit Then Exit Sub                                          ' Form wird noch initialisiert -> fertig
    cmdUpdate.Enabled = True                                        ' Übernehmen enablen
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub CheckOption_Click(Index As Integer)
On Error Resume Next                                                ' Feherbehandlung deaktivieren
    If bInit Then Exit Sub                                          ' Form wird noch initialisiert -> fertig
    cmdUpdate.Enabled = True                                        ' Übernehmen enablen
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub cmbOption_Change(Index As Integer)
On Error Resume Next                                                ' Feherbehandlung deaktivieren
    If bInit Then Exit Sub                                          ' Form wird noch initialisiert -> fertig
    cmdUpdate.Enabled = True                                        ' Übernehmen enablen
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Public Function SaveOptionsInArray(f As Form)
' Speichert alle Feld werte des Optionforms ins Option array
    Dim i As Integer                                                ' Counter
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    For i = 0 To f.CheckOption.Count - 1                            ' Alle Check Felder durchgehen
        If f.CheckOption(i).Tag <> "" Then                          ' Wenn Tag angegeben (optionsname)
            'Call objOptions.SetOptionByName(f.CheckOption(i).Tag, CBool(f.CheckOption(i).Value))
            Call OptionSetByName(f.CheckOption(i).Tag, CBool(f.CheckOption(i).Value)) ' Wert Speichern
        End If
    Next i                                                          ' Nächstes Check feld
    For i = 0 To f.txtOption.Count - 1                              ' Alle Text Felder durchgehen
        If f.txtOption(i).Tag <> "" Then                            ' Wenn Tag angegeben (optionsname)
            'Call objOptions.SetOptionByName(f.txtOption(i).Tag, f.txtOption(i).Text)
            If cmdFile(i).Visible Then
                Call OptionSetByName(f.txtOption(i).Tag, f.cmdFile(i).Tag) ' Wert Speichern
            ElseIf cmdPath(i).Visible Then
                Call OptionSetByName(f.txtOption(i).Tag, f.cmdPath(i).Tag) ' Wert Speichern
            Else
                Call OptionSetByName(f.txtOption(i).Tag, f.txtOption(i).Text) ' Wert Speichern
            End If
        End If
    Next i                                                          ' Nächstes text Feld
     For i = 0 To f.cmbOption.Count - 1                             ' Alle Combo Felder durchgehen
        If f.cmbOption(i).Tag <> "" Then                            ' Wenn Tag angegeben (optionsname)
            'Call objOptions.SetOptionByName(f.txtOption(i).Tag, f.txtOption(i).Text)
            Call OptionSetByName(f.cmbOption(i).Tag, f.cmbOption(i).Text) ' Wert Speichern
        End If
    Next i                                                          ' Nächstes Combo Feld
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr

Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "SaveOptionsInArray", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function
'Public Function SaveOptionsInArray(f As Form)
'' Speichert alle Feld werte des Optionforms ins Option array
'
'    Dim i As Integer                                                ' Counter
'    Dim szRegKey As String                                          ' Registry Schlüssel
'
'On Error GoTo ErrorHandler                                          ' Fehlerbehandlung aktivieren
'
'    ' Erst Control Values in Array
'    For i = 0 To f.CheckOption.Count - 1                            ' Alle Check Felder durchgehen
'        If f.CheckOption(i).Tag <> "" Then
'            Call OptionSetByName(f.CheckOption(i).Tag, CBool(f.CheckOption(i).Value))
'        End If
'    Next i
'
'    For i = 0 To f.txtOption.Count - 1                              ' Alle Text Felder durchgehen
'        If f.txtOption(i).Tag <> "" Then
'            Call OptionSetByName(f.txtOption(i).Tag, f.txtOption(i).Text)
'        End If
'    Next i
'
'exithandler:
'On Error Resume Next
'
'Exit Function
'ErrorHandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.ErrorHandler(MODULNAME, "SaveOptionsInArray", errNr, errDesc)
'    Resume exithandler
'End Function

