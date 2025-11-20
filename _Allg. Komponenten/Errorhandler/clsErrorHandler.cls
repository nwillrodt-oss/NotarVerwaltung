VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsErrorHandler"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit                                                         ' Variaben Deklaration erzwingen
Option Compare Text                                                     ' Sortierreihenfolge festlegen
Private Const MODULNAME = "clsErrorhandler"                             ' Modulname für Fehlerbehandlung

Private objObjectBag As Object                                          ' ObjectBag object
Private MainForm As Object                                              ' Ref. Aufs Mainform der anwendung
Private ProtPath As String                                              ' Protokoll Pfad
Private ErrPath As String                                               ' Fehler Protokoll Pfad
Private Computer As String                                              ' Akt Computername fürs Protokoll
Private User As String                                                  ' Akt Benutzername fürs Protokoll
Private szIgnoreMsg() As String                                         ' Array das zu Ignorierende Meldungen enthält
' Für Table Logging
Private bTabLogging As Boolean                                          ' Logging in Tabelle Aktiviert
Private szTabLog_Table As String                                        ' Tabellenname
Private szTabLog_PCField As String                                      ' Feld für Client namen
Private szTabLog_UserField As String                                    ' Feld für Usernamen
Private szTabLog_AppField As String                                     ' Feld Für Application Namen
Private szTabLog_MSGField As String                                     ' Feld für Fehlermeldung
Private szTAbLog_TimeField As String                                    ' Datum Uhrzeit

Private bMSGInit As Boolean                                             ' True wenn gerade eine Meldung abgearbeitet wird

Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
On Error Resume Next                                                    ' Hier keine Fehler mehr
    Set objObjectBag = Nothing                                          ' Object Bag Object killen
    Err.Clear                                                           ' Evtl. error clearen
End Sub

Public Function InitObjectBag(objOb As Object)
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    Set objObjectBag = objOb
    If Not objObjectBag Is Nothing Then
        'Set objError = objObjectBag.GetErrorObj
    End If
    Err.Clear                                                           ' Evtl. error clearen
End Function

Public Function InitTableLog()
' Initialisiert das Mittlogen von Fehlern in einer DB Tabelle

End Function
                                                                        ' *****************************************
                                                                        ' Properties
Public Property Let SetProtFileName(ByVal szProtFilename As String)
     ProtPath = szProtFilename                                          ' Setzt Pfad der Protokolldatei
End Property

Public Property Let SetErrFileName(ByVal szErrFilename As String)
      ErrPath = szErrFilename                                           ' Setzt Pfad der Fehlerprotokolldatei
End Property

Public Property Get GetErrFileName()
      GetErrFileName = ErrPath                                          ' Gibt Pfad der Protokolldatei zurück
End Property

Public Property Let SetComuterName(ByVal szComputername As String)
      Computer = szComputername                                         ' Setzt den Computernamen
End Property

Public Property Let SetUserName(ByVal szUsername As String)
      User = szUsername                                                 ' Setzt den Benutzernamen
End Property

Public Property Let SetMainForm(ByVal fMain As Object)
    Set MainForm = fMain                                                ' Setzt referenz aufs Mainform
End Property
                                                                        ' *****************************************
                                                                        ' Methoden
Public Function ErrorHandler(szModName As String, _
        szFktName As String, _
        szErrNr As String, _
        Optional szErrMsg As String, _
        Optional szDetails As String)
    Dim szErrText As String                                             ' Text der Fehlermeldung
    Dim szDetailText As String                                          ' Optionale Detailinformationen
    Dim bDetails As Boolean                                             ' Sollen Details angezeigt werden
    Dim bIgnore As Boolean
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    Err.Clear                                                           ' Evtl. Error clearen
    szErrText = "In der Funktion: " & szFktName & vbCrLf _
                & "im Modul: " & szModName & vbCrLf _
                & "ist folgender Fehler aufgetreten." & vbCrLf _
                & "FehlerNr: " & szErrNr & vbCrLf _
                & "Fehlertext: " & szErrMsg & vbCrLf                    ' Meldungstext zusammensetzen
    If szDetails <> "" Then                                             ' Falls Details vorhanden
        szDetailText = "Details: " & vbCrLf & szDetails                 ' Deatiltest preparieren
        bDetails = True                                                 ' Details anzeigen
    End If
    If Not CheckErrMsgIgnore(szModName, szFktName, szErrNr) Then        ' Wenn Meldung Nicht im Ignore Array
        bIgnore = True                                                  ' Bei Fehleren immer Ignorieren anbieten
        Call ShowErrMsg(szErrText, vbCritical, "Fehler", bDetails, szDetailText, , bIgnore) ' Fehlermeldung Anzeigen
        If bIgnore Then                                                 ' Nach der Meldung immer noch Ignorieren (im error form gesetzt)
            Call CheckErrMsgIgnore(szModName, szFktName, szErrNr, True) ' in Ignore Array Eintragen
        End If
    End If
    Call WriteErrProt(szModName, szFktName, szErrNr, szErrMsg, szDetails) ' Fehler ins Protokoll schreiben
    Err.Clear                                                           ' Evtl Error Clearen
End Function

Public Function WriteProt(szText As String, _
        Optional ProtFilePath As String, _
        Optional bWithOutDate As Boolean, _
        Optional bWithOutUser As Boolean, _
        Optional bWithOutComputer As Boolean, _
        Optional bAppName As Boolean)
'Schreibt Meldung in Protokolldatei
'    Dim FN As Integer
    Dim filepath As String                                              ' Dateipfad
    Dim szMsg As String                                                 ' Meldungstext
    If ProtFilePath <> "" Then                                          ' Spez. Protokollpfad übergeben ?
        filepath = ProtFilePath                                         ' dann diesen übernehmen
    Else                                                                ' Nein
        filepath = ProtPath                                             ' Standartpfad nehmen
    End If
    filepath = Trim(filepath)                                           ' Pfad Trimmen
    If filepath = "" Then filepath = "c:\" & objObjectBag.GetAppTitle & "Prot.txt"  ' Falls Pfad Leer in C legen
    szMsg = CStr(Now()) & vbTab                                         ' Datum & uhrzeit im Meldung
    If Computer <> "" And Not bWithOutComputer Then szMsg = szMsg & Computer & vbTab ' PC name im Meldung
    If User <> "" And Not bWithOutUser Then szMsg = szMsg & User & vbTab ' Username im Meldung
    szMsg = szMsg & objObjectBag.GetAppTitle & vbTab                    ' Anwendungstitel im Meldung
    szMsg = szMsg & szText & vbCrLf                                     ' Eigentlicher Meldungstext
    Call WriteTextFile(filepath, szMsg)                                 ' In text datei schreiben
'On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
'    If Len(Dir$(filepath, vbNormal)) > 0 Then                           ' Wenn Datei existiert
'        FN = FreeFile
'        Open filepath For Append As #FN                                 ' Datei zum anhängen öffnen
'        Print #FN, szMsg;                                               ' Meldung Schreiben
'        Close FN                                                        ' Datei Schliessen
'    Else
'        FN = FreeFile
'        Open filepath For Output As #FN                                 ' Neue Datei erstellen
'        Print #FN, szMsg;                                               ' Meldung Schreiben
'        Close FN                                                        ' Datei Schliessen
'    End If
    Err.Clear                                                           ' Evtl. Error clearen
End Function

Public Function ShowErrMsg(szText As String, _
            Buttons As VbMsgBoxStyle, szTitel As String, _
            Optional bDetails As Boolean, Optional szDetails As String, _
            Optional CallerForm As Object, Optional ByRef bShowIgnore As Boolean) As VbMsgBoxResult
' Icons für Frage, Information , Warnung, Critical zusammenstellen
    Dim f As frmMsgBox                                                  ' Referenz aufs FehlerForm
    Dim result                                                          ' Rükgabe wert des Fehlerforms
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bMSGInit Then Exit Function                                      ' Falls schon eine Meldung offen -> fertig
    Set f = New frmMsgBox                                               ' Neues FehlerForm
    Call f.InitPicture(Buttons)                                         ' Symbole initialisieren
    Call f.InitButtons(Buttons)                                         ' Button Initialisieren
    f.lblMSG.Caption = szText                                           ' Fehler text setzen
    f.lblDetails.Caption = szDetails                                    ' Detail Text setzen
    f.chkIgnoreErr.Visible = bShowIgnore                                ' Checkbox meldung Ignorieren anzeigen/ausblenden
    If f.lblMSG.Height > 915 Then                                       ' Höhe einstellen ?
        If bShowIgnore Then                                             ' Abhängig von Chekbox
            f.chkIgnoreErr.Top = f.lblMSG.Top + f.lblMSG.Height         ' CkeckIgnore Verschieben
            f.cmdOK.Top = f.chkIgnoreErr.Top + f.chkIgnoreErr.Height    ' OK Button verschieben
        Else                                                            ' Sonst
            f.cmdOK.Top = f.lblMSG.Top + f.lblMSG.Height + 50           ' OK Button verschieben
        End If
        f.cmdRetry.Top = f.cmdOK.Top                                    ' Nochmal Button verschieben
        f.cmdEsc.Top = f.cmdOK.Top                                      ' ESC Button verschieben
        f.cmdDetails.Top = f.cmdOK.Top                                  ' Deatil Button verschieben
        f.Height = f.cmdOK.Top + f.cmdOK.Height + 50 + 405              ' Form höhe anpassen
    End If
    'f.BorderStyle = 3
    Call objObjectBag.CheckFormStyle(f)                                 ' Form Formatierung behandeln
    f.Refresh                                                           ' Form Aktualisieren
    f.Caption = szTitel                                                 ' Fenster Titel setzen
    f.cmdDetails.Visible = bDetails                                     ' Button details ein oder ausblenden
    bMSGInit = True                                                     ' Ab hier keine Mehfach bearbbeitung
    f.Show vbModal, CallerForm                                          ' Anzeigen (Modal)
    ShowErrMsg = f.result                                               ' Ergebnis zurück geben
    bShowIgnore = f.bIgnorErr                                           ' Sol lder Fehler in zukunft ignoriert werden?
    Unload f                                                            ' Form Schliessen
    bMSGInit = False                                                    ' Jetzt können ach wieder andere Fehler zugelassen wedren
    Err.Clear                                                           ' Evtl. Error Clearen
End Function
                                                                        ' *****************************************
                                                                        ' Hilfs funktionen
Private Function CheckErrMsgIgnore(szModName As String, _
        szFktName As String, _
        szErrNr As String, Optional bWrite As Boolean) As Boolean
' Überprüft das szIgnoreMsg() Array ob meldung enthalten
' bWrite = True wird auch gleich reingeschrieben
    Dim i As Integer                                                    ' Counter für szIgnoreMsg()
    Dim szCheckMsg As String                                            ' Eintrag fürs Array / zum überprüfen
    Dim bIn As Boolean                                                  ' True wenn Meldung im Array Enthalten
    szCheckMsg = szModName & ";" & szFktName & ";" & szErrNr            ' Array Eintrag aus Modul Fkt und ErrNr zusammenstellen
On Error GoTo FirstError                                                ' Bei Fehler ist Array Leer
    For i = 0 To UBound(szIgnoreMsg)                                    ' Gesamtes Array durchlaufen
        If szIgnoreMsg(i) = szCheckMsg Then                             ' Prüfen ob eintrag übereinstimmt
            bIn = True                                                  ' Gefunden
        End If
    Next i                                                              ' Nächstes Array Item
FirstError:
    Err.Clear                                                           ' Evtl Error clearen
On Error Resume Next                                                    ' Hier keine Fehlerbehandlung mehr
    If (Not bIn) And bWrite Then                                        ' wenn nicht gefunden und Schreiben = True
        If CheckIgnoreMsgArray(szIgnoreMsg) Then                        ' Wenn Arry Schon einträge enthält
            ReDim Preserve szIgnoreMsg(UBound(szIgnoreMsg) + 1)         ' Einen anhängen
        Else                                                            ' Sonst
            ReDim szIgnoreMsg(0)                                        ' Ersten Eintrag festlegen
        End If
        szIgnoreMsg(UBound(szIgnoreMsg)) = szCheckMsg                   ' Neuen Eintrag Setzen
        bIn = True                                                      ' Jetzt ist er drin
    End If
    CheckErrMsgIgnore = bIn                                             ' Ergebnis zurück
    Err.Clear                                                           ' Evtl Error clearen
End Function

Public Function CheckIgnoreMsgArray(VarArray() As String) As Boolean
' prüft ob das CheckIgnoreMsgArray schon einen Eintrag hat
' zur Fehlervermeidung
    Dim i As Integer                                                    ' Array größe
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    i = UBound(VarArray)                                                ' Array Größe ermitteln
    If Err.Number <> 0 Then                                             ' Ist ein Fehler aufgetreten
        Err.Clear                                                       ' Err Clearen
        CheckIgnoreMsgArray = False                                     ' Array ist leer
        Exit Function                                                   ' Fertig
    End If
    CheckIgnoreMsgArray = True                                          ' Sonst Array Definiert
    Err.Clear                                                           ' Evtl. error clearen
End Function

Private Function WriteErrProt(szModName As String, _
        szFktName As String, _
        szErrNr As String, _
        szErrMsg As String, _
        Optional szDetails As String)
' Stellt fehler Protokoll meldung zusammen und schreibt diese in Text datei
'    Dim aa As String
'    Dim FN As Integer
    Dim szMsg As String                                                 ' Meldungs text
On Error Resume Next                                                    ' Hier keine Fehler
    szMsg = CStr(Now()) & vbTab                                         ' Datum in Meldungstext
    If Computer <> "" Then szMsg = szMsg & Computer & vbTab             ' evtl. Clientnamen in Meldungstext
    If User <> "" Then szMsg = szMsg & User & vbTab                     ' Evtl. Usernamen in Meldungstext
    szMsg = szMsg & szModName & vbTab                                   ' Modulname in Meldungstext
    szMsg = szMsg & szFktName & vbTab                                   ' Funktionsname in Meldungstext
    szMsg = szMsg & szErrNr & vbTab                                     ' Fehler nummer in Meldungstext
    szMsg = szMsg & szErrMsg & vbTab                                    ' Fehler Beschreibung in Meldungstext
    szMsg = szMsg & szDetails & vbCrLf                                  ' Details in Meldungstext
    Call WriteTextFile(ErrPath, szMsg)                                  ' In TXT datei schreiben
'    If Len(Dir$(ErrPath, vbNormal)) > 0 Then
'        FN = FreeFile
'        Open ErrPath For Append As #FN
'        Print #FN, szMsg;
'        Close FN
'    Else
'        FN = FreeFile
'        Open ErrPath For Output As #FN
'        Print #FN, szMsg;
'        Close FN
'    End If
    Err.Clear                                                           ' Evtl. Error clearen
End Function

Private Function WriteTextFile(szTextFile As String, _
        szText As String)
    Dim FN As Integer
    Dim filepath As String                                              ' Dateipfad
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If Len(Dir$(filepath, vbNormal)) > 0 Then                           ' Wenn Datei existiert
        FN = FreeFile
        Open filepath For Append As #FN                                 ' Datei zum anhängen öffnen
        Print #FN, szText;                                              ' Meldung Schreiben
        Close FN                                                        ' Datei Schliessen
    Else
        FN = FreeFile
        Open filepath For Output As #FN                                 ' Neue Datei erstellen
        Print #FN, szText;                                              ' Meldung Schreiben
        Close FN                                                        ' Datei Schliessen
    End If
    Err.Clear                                                           ' Evtl. Error clearen
End Function
