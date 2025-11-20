Attribute VB_Name = "modMain"
Option Explicit                                                     ' Variaben Deklaration erzwingen
Option Compare Text                                                 ' Sortierreihenfolge festlegen

Private Const MODULNAME = "modMain"                                 ' Modulname für Fehlerbehandlung

Public bDebug As Boolean        ' Zum Entwickeln

Public objObjectBag As Object                                       ' ObjectBag object
Public objError  As Object                                          ' Glob Error object
Public objRegTools As Object                                        ' Registry Tools
Public objTools As Object                                           ' Hilfreiches
Public objOptions As Object                                         ' Optionen einlesen & Speichern
'Public objDBconn As Object                                          ' DB Connection
'Public objSQLTools As Object                                        ' SQL Tools
'Public objOffice As Object                                          ' Office Verbindung
'Public bAutoConnect As Boolean
Public bNotShowSplash As Boolean                                    ' Wenn True Kein Splash
Private bLEET As Boolean
Private bConsole As Boolean
Private bE As Boolean

Public Sub Main()                                                   ' Start Procedur
    Dim szLastCon As String                                         ' LastConnection Value aus Reg
    Dim szConnArray() As String                                     ' (0)=Servername, (1)=DBName, (1)=DBUser, (4)=PWD
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Call ReadCmdParams                                              ' Startparameter auslesen
    Set objObjectBag = CreateObject("ObjectBag.clsObjectBag")       ' ObjectBag holen
    If Not objObjectBag Is Nothing Then                             ' Prüfen ob ObjectBAg erfolgreich initialisiert
        Call InitObjectBag                                          ' Globale einstellungen ermitteln
        Call objError.WriteProt(PROT_APP_START)                     ' Start ins Log schreiben
        Call objError.WriteProt("Version " & objObjectBag.GetAppVersion) ' Version ins Protokoll
        Call objError.WriteProt("Anwendungsverz.: " & objObjectBag.Getappdir()) ' Anwendungs Verz ins Protokoll
'        Call objError.WriteProt("Eigene Dateien: " & objObjectBag.GetPersonalDir())
    Else
    
    End If
    Call OpenMainForm                                               ' Main Form Zeigen
exithandler:
On Error Resume Next                                                ' Hier keine Fehler Mehr
    Call ShowSplash(False)                                          ' Splash ausblenden
Exit Sub
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    If objError Is Nothing Then
        Call MsgBox("Im Modul '" & MODULNAME & "' in der Funktion ist ein Fehler aufgetreten." & vbCrLf & _
            "Fehlernr.: " & errNr & vbCrLf & "Beschreibung: " & errDesc, vbCritical, "Fehler")
        End
    Else
        Call objError.Errorhandler(MODULNAME, "Main", errNr, errDesc)   ' Fehler behandlung aufrufen
    End If
    
    Resume exithandler                                              ' Weiter mit Exithandler
End Sub

Public Function OpenMainForm(Optional dbConn As Object)             ' Zeigt das Hauptformular an
On Error Resume Next                                                ' Fehler behandlung deaktivieren
    objObjectBag.SetMainForm = frmMain                              ' form Ref im ObjBag merken
    Call objObjectBag.CheckFormStyle(objObjectBag.getMainForm)
'    If dbConn Is Nothing Then Exit Function                         ' DB Verbindung (autoconnent) übergeben
'    Set frmMain.ThisDBCon = dbConn                                  ' Datenbank verbindung übergeben
    frmMain.Show                                                    ' Form Anzeigen
    Err.Clear                                                       ' Evtl. ERror Clearen
End Function

Public Function AppExit()                                           ' Beendet diese anwendung
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Call objOptions.SaveOptions                                     ' optionen in Registry Speichern
    Set objObjectBag = Nothing                                      ' Objectbag Schliessen
    objError.WriteProt (PROT_APP_END)                               ' Beenden ins Log
    End                                                             ' Anwendung beenden
exithandler:
On Error Resume Next
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "AppExit", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ShowHelp()                                          ' Zeigt die Hilfe an
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
'    Call objTools.ShellExec(objObjectBag.GetAppDir() & "MegaAdminToolHilfe.chm", "", vbNormalFocus)
    Call objTools.HTMLHelp_ShowTopic(objObjectBag.Getappdir() & "MegaAdminToolHilfe.chm")
exithandler:
On Error Resume Next
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ShowHelp", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ShowSplash(bVisible As Boolean)                     ' Zeigt Splash Form an und blendet es aus
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call objObjectBag.ShowSplash(bVisible)                          ' Spalsh form öffen/Schliessen
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Public Sub ShowReadMe()                                             ' Zeigt ReadMeDatei an
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call objObjectBag.ShowReadMe
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Public Sub ShowAbout()                                              ' Zeigt das About Form an
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call objObjectBag.ShowAbout("", False, objDBconn)
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Public Function ShowOptions()                                       ' Zeigt Options Form an und blendet es aus
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call objOptions.ShowOptions                                     ' Options Form anzeigen
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Public Function ReportBug()
    Call objObjectBag.ReportBug
End Function

Private Sub InitObjectBag()                                         ' Initialisiert den Object bag
Dim szDeteils As String                                             ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                          ' Fehlerbehandlung akticieren
                                                                    ' Anwendungsinformationen setzen
    objObjectBag.SetShowSplash = Not bNotShowSplash                 ' Startparameter Splash anzeigen
    objObjectBag.SetAppDir = App.Path                               ' Anwendungsverzeichnis (ordner in dem die .exe liegt)
    objObjectBag.SetAppTitle = SZ_APPTITLE                          ' Anwendungtitel aus modConst
    objObjectBag.SetCopyright = SZ_COPYRIGHT                        ' Copyright aus modConst
    objObjectBag.SetMajor = App.Major                               ' Major Version
    objObjectBag.SetMinor = App.Minor                               ' Minor Version
    objObjectBag.SetRevision = App.Revision                         ' Revision
    objObjectBag.setXMLFile = INI_XMLFILE                           ' Dateiname XML File aus modConst
    objObjectBag.setreadmepath = SZ_READMEFILE                      ' Name Readme Datei aus modConst
    objObjectBag.SetSupportMail = SZ_SUPPORTMAIL                    ' MailAdresse für Support aus modConst
    objObjectBag.SetWWW = SZ_WWW                                    ' InternetAdresse aus modConst
    objObjectBag.SetLeet = bLEET
    objObjectBag.setconsole = bConsole
    szDeteils = "Anwendungs Informationen eingelesen."              ' Details für Fehlerbehandlung
                                                                    ' Sonstige benötigte Objecte aus ObjectBag holen
    Set objError = objObjectBag.GetErrorObj()                       ' Fehlerbehandlung
    szDeteils = "Fehlerbehandlungs Object Initialisiert."           ' Details für Fehlerbehandlung
    Set objRegTools = objObjectBag.GetRegToolsObj()                 ' Registry Tools
    szDeteils = "Registry Object Initialisiert."                    ' Details für Fehlerbehandlung
    Set objTools = objObjectBag.GetToolsObj()                       ' Allg. Tools
    szDeteils = "Allg. Tools Object Initialisiert."                 ' Details für Fehlerbehandlung
'    If objObjectBag.InitSQLTools() Then                             ' SQL Tools initialisieren
'        Set objSQLTools = objObjectBag.GetSqlToolsobj()
'        szDeteils = "SQL Object Initialisiert."                     ' Details für Fehlerbehandlung
'    Else
'        'Fehlerbehandlung ???
'        szDeteils = "SQL Object nicht Initialisiert."               ' Details für Fehlerbehandlung
'    End If
    If objObjectBag.InitOptions() Then                              ' Options Object initialisieren
        Set objOptions = objObjectBag.GetOptionsObj()               ' ObJekt holen
        szDeteils = "Options Object Initialisiert."                 ' Details für Fehlerbehandlung
        objOptions.SetOptioniniPath = INI_OPTIONSINI                ' Optionsini bekannt geben
        Call objOptions.InitOptions                                 ' Options aus ini auslesen
        szDeteils = "Optionen eingelesen."                          ' Details für Fehlerbehandlung
                                                                    ' Anwendungsspez Änderungen an den Optionen
        objError.SetProtFileName = objOptions.getOptionByName(OPTION_APPLOG)  ' Gleich an error Obj Weitergeben
        objError.SetErrFileName = objOptions.getOptionByName(OPTION_ERRLOG)   ' Gleich an error Obj Weitergeben
        bNotShowSplash = objOptions.getOptionByName(OPTION_SPLASH)
'        If bE Then bNotShowSplash = Not bE
    End If
'    Call CheckFirstStart                                            ' Prüfen ob 1. Start oder neue Version
    szDeteils = "Splash anzeigen."                                  ' Details für Fehlerbehandlung
    Call ShowSplash(Not bNotShowSplash)                             ' Splash zeigen
    Call objObjectBag.ShowMSGForm(True, "Initialisiere Anwendung ...") ' Aktuelle Aktion im Spalsh oder MSG Form anzeigen
'    If objObjectBag.GetWordVersion <> "" Then                       ' Wenn Word Installiert
'        If objObjectBag.InitOfficeTools Then                        ' Office Schittstelle Initialisieren
'            Set objOffice = objObjectBag.GetOfficeObj()
'            szDeteils = "Office Object initialisiert."
'        End If
'    End If
exithandler:
On Error Resume Next
Exit Sub
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "InitObjectBag", errNr, errDesc, szDeteils) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit exithandler
End Sub

Public Function ReadCmdParams()                                     ' Start Parameter auslesen
    Dim szCmdString As String                                       ' String mit Startparametern
    Dim cmdArray() As String                                        ' Array mit startparametern
    Dim i As Integer                                                ' Counter
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    szCmdString = Command$                                          ' Startparameter einlesen
    Do While InStr(szCmdString, Space(2)) > 0                       ' Nicht benötigte Spaces entfernen
        szCmdString = Replace(szCmdString, Space(2), Space(1))
    Loop
    szCmdString = UCase(szCmdString)                                ' Generell Gross- oder Kleinschreibung durch UCase hier also nur Grossbuchstaben
    cmdArray = Split(szCmdString, Space(1))                         ' In Array Spliten
    
    For i = 0 To UBound(cmdArray)                                   ' Ganzes Array abarbeiten
        Select Case UCase(cmdArray(i))
        Case UCase(CMD_NOSplash)                                    ' kein Splash
            bNotShowSplash = True
        Case UCase(CMD_EGG)
            bNotShowSplash = False
'            bE = Not bNotShowSplash
        Case UCase(CMD_DOS), UCase(CMD_CMD), UCase(CMD_CONSOLE)     ' Console
            bConsole = True
        Case UCase(CMD_AUTOCON)                                     ' Autoconnect
'            bAutoConnect = True
        Case UCase(CMB_EXPERT)
'            bExpert = True                                          ' Experten modus
        Case UCase(CMB_LEET)
            bLEET = True
        Case UCase(CMD_HELP), UCase(CMD_HELP2)
'            Call MsgBox(CMD_HELP_TXT, vbInformation, SZ_APPTITLE & " " _
                    & App.Major & "." & App.Minor & "." & App.Revision)
            End                                                      ' Anwendung beenden
        Case Else
        
        End Select
    Next
exithandler:
On Error Resume Next
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ReadCmdParams", errNr, errDesc)
    Resume exithandler
End Function


