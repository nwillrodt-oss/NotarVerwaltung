Attribute VB_Name = "modMain"
Option Explicit                                                     ' Variaben Deklaration erzwingen
Private Const MODULNAME = "modMain"                                 ' Modulname für Fehlerbehandlung

Public bDebug As Boolean                                            ' Zum Entwickeln
Private StartTick As Double                                         ' Ticks zur performance Messung

Public objObjectBag As Object                                       ' ObjectBag object
Public objError  As Object                                          ' Glob Error object
Public objDBconn As Object                                          ' DB Connection
Public objRegTools As Object                                        ' Registry Tools
Public objTools As Object                                           ' Hilfreiches
Public objSQLTools As Object                                        ' SQL Tools
Public objOptions As Object                                         ' Optionen einlesen & Speichern
Public objOffice As Object                                          ' Office Verbindung
    
Public bAutoConnect As Boolean                                      ' Wenn True Automatische DB Verbindung
Public bNotShowSplash As Boolean                                    ' Wenn True Kein Splash
Private bExpert As Boolean                                          ' Experten Modus
Private bConsole As Boolean                                         ' Consolen modus
Public bE As Boolean
Private bLEET As Boolean
Private bMorse As Boolean

Public User As UserInfo                                             ' Akt Benutzer informationen

Public Type UserInfo
    ID As String                                                    ' ID In User Tabelle
    NTUsername As String                                            ' NT Anmeldename
    Username As String                                              ' Benutername in User Tabelle
    Vorname As String                                               ' Vorname
    Nachname As String                                              ' Nachname
    Fullname As String                                              ' Nachname, Vorname
    tel As String                                                   ' Tel in User Tabelle
    fax As String                                                   ' FAx in User Tabelle
    eMail As String                                                 ' Mail in User Tabelle
    System As Boolean                                               ' Wenn True ist systemverwalter
End Type

Public Type ListViewInfo                                            ' ListView Informationen
    szSQL As String                                                 ' zugrundeliegendes SQL Statement
    szTag As String                                                 ' Tag des ListViews (welche Daten werden angezeigt)
    bValueList As Boolean                                           ' Darstellung als Valuelist (1.DS pro Wert ein Item)
    DelFlagField As String                                          ' Feld in dem ein gelöscht flag gesetzt werden kann
    WhereNoDel As String                                            ' Where Part mit gelöschten DS (Flag)
    szWhere As String                                               ' evtl. Where Statement
    lngImage As Integer                                             ' Image Index für Item
    AltImage As Integer                                             ' Alternatives Image
    AltImgField As String                                           ' Feld das für alt Image geprüft wird
    AltImgValue As String                                           ' Value der für alt image geprüft wird
    bListSubNodes As Boolean                                        ' Sollen Subnodes im Liszview mitangezeigt werden
    bEdit As Boolean                                                ' Einträge können bearbeitet werden
    bNew As Boolean                                                 ' Es können neue einträge angelegt werden
    bDelete As Boolean                                              ' Es dürfen Einträge gelöscht werden
    bSelectNode As Boolean                                          ' dbl klick selectet den entsprechenden Node
    bShowKontextMenue As Boolean                                    ' Kontextmenü zulässig
End Type

Public Type TreeViewNodeInfo                                        ' TreeNode Informationen
    szName As String                                                ' Node Name (nicht angezeigt)
    szTag As String                                                 ' Tag des Nodes (welche Daten werden angezeigt)
    szText As String                                                ' Text des Nodes (angezeigt)
    szDesc As String                                                ' Beschreibung des Nodes (tooltip)
    szKey As String                                                 ' Eindeutiger Schlüssel des Nodes
    szSQL As String                                                 ' zugrundeliegendes SQL Statement
    szWhere As String                                               ' evtl. Where Statement
    szTyp As String                                                 ' Knoten Typ (statisch / dynamisch)
    bShowSubnodes As Boolean                                        ' Sollen sofort alle subnodes mit angezeigt werden
    lngImage As Integer                                             ' Image Index für Node
    AltImage As Integer                                             ' Alternatives Image
    AltImgField As String                                           ' Feld das für alt Image geprüft wird
    AltImgValue As String                                           ' Value der für alt image geprüft wird
    bShowKontextMenue As Boolean                                    ' Kontextmenü zulässig
End Type

Public DBFormArray() As Object                                      ' Auflistung aller geöffneten DB Formulare

Public Sub ConnectSuccess(objCon As Object)
   ' Call OpenDBForm(objCon)                                        ' DB Form anzeigen
End Sub

Public Sub Main()
' Start Funktion
    Dim szLastCon As String                                         ' LastConnection Value aus Reg
    Dim szConnArray() As String                                     ' (0)=Servername, (1)=DBName, (1)=DBUser, (4)=PWD
    Dim szAutUsername As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Call ReadCmdParams                                              ' Startparameter auslesen
    bAutoConnect = True                                             ' Für Diese aplication immer true
    Set objObjectBag = CreateObject("ObjectBag.clsObjectBag")       ' ObjectBag holen
    If Not objObjectBag Is Nothing Then                             ' Prüfen ob ObjectBAg erfolgreich initialisiert
        Call InitObjectBag                                          ' Globale einstellungen ermitteln
        Call objError.WriteProt(PROT_APP_START)                     ' Start ins Log schreiben
        Call objError.WriteProt("Version " & objObjectBag.GetAppVersion) ' Version ins Protokoll
        Call objError.WriteProt("Anwendungsverz.: " & objObjectBag.Getappdir()) ' Anwendungsverz ins Protokoll
        Call objError.WriteProt("Eigene Dateien: " & objObjectBag.GetPersonalDir()) ' Eigene Dokumente ins Protokoll
        ' DB Verbindung nur SQL kein Access
        If objObjectBag.InitDBConnection(False, True) Then          ' DB Object holen
            Set objDBconn = objObjectBag.GetDBConObj()
            If bAutoConnect Then                                    ' Automatische anmeldung
                szLastCon = objOptions.GetOptionByName(OPTION_LASTCON) ' Option LastConnection auslesen
                If szLastCon <> "" Then
                    If InStr(szLastCon, ";") = 0 Then               ' Enthält der ConString leine Semikolons ist er verschlüsselt
                        szLastCon = objTools.crypt(szLastCon, False) ' Dann entschlüsseln
                    End If
                    szConnArray = Split(szLastCon, ";")             ' Value aufspliten
                    If objDBconn.GetAdodbConn(CLng(szConnArray(0)), CStr(szConnArray(1)), _
                                CStr(szConnArray(2)), CStr(szConnArray(3)), _
                                CStr(szConnArray(4)), _
                                CBool(szConnArray(5))) Then
                        Call objError.WriteProt(PROT_DB_AUTOCON)    ' DB anmeldung ins Log schreiben
                        
                    Else                                            ' verbindungsproblem
                        If objDBconn.GetAdodbConn(2, "", "", "", "") Then  ' manuel versuchen
                        
                        Else
                            ' Evtl. noch ne meldung
                            Call AppExit                            ' Wenn immer noch nicht dan raus
                        End If
                    End If
                Else                                                ' Keine letzte verbindung gespeichert
                    If objDBconn.GetAdodbConn(2, "", "", "", _
                            "") Then                                ' manuel versuchen
            
                    Else
                        ' Evtl. noch ne meldung
                        Call AppExit                                ' Wenn immer noch nicht dan raus
                    End If
                End If  ' szLastCon <> ""
            Else
                If objDBconn.GetAdodbConn(2, "", "", "", _
                            "") Then                                ' manuel versuchen
                Else
                    ' Evtl. noch ne meldung
                    Call AppExit                                    ' Wenn immer noch nicht dan raus
                End If
            End If ' bAutoConnect
        End If  ' objObjectBag.InitDBConnection
    Else                                                            ' Kein Object Bag
        ' Was nun ?
        'Stop
    End If ' Not objObjectBag Is Nothing
    Call objError.WriteProt("Connect with " & objDBconn.getDBName)    ' DB anmeldung erfolgreich ins Log schreiben
    Call objError.WriteProt("Connect successfull")                  ' DB anmeldung erfolgreich ins Log schreiben
    Call objObjectBag.ShowMSGForm(True, _
            "Initialisiere Datenbank anmeldung ...")                ' MSG Form Meldung
    Call objDBconn.InitUserLogIn(objOptions.GetOptionByName("UserLoginTable"), _
        objOptions.GetOptionByName("UserNameField"), _
        objOptions.GetOptionByName("UserPWDField"), _
        objOptions.GetOptionByName("LoginText"), _
        CLng(objOptions.GetOptionByName("UserMaxCount")) _
        , CBool(objOptions.GetOptionByName("SingleSignOn")))        ' Wenn Db verbindung erfolgreich USER LOGIN Initialisieren
    szAutUsername = objDBconn.ShowUserLogin()                       ' Anmelden
    If szAutUsername = "" Then
        Call objError.WriteProt("Userlogin cancel")                 ' Benutzer abbruch ins Protokoll
        Call AppExit                                                ' Anwendung schliessen
        GoTo exithandler                                            ' Wir sind hier fertig
    Else
        Call objError.WriteProt("Userlogin successfull")            ' User Login in Protokoll
        Call ReadUserContext(szAutUsername)                         ' Benutzer informationen einlesen
    End If
    Call objTools.Wait(3000)                                        ' Kurz warten
    Call OpenMainForm                                               ' Main Form Zeigen
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Call ShowSplash(False)                                          ' Splash ausblenden
    Err.Clear                                                       ' Evtl. error clearen
Exit Sub                                                            ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "Main", errNr, errDesc)   ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Sub

Private Sub InitObjectBag()
' Initialisiert den Object bag
Dim szDeteils As String                                             ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
                                                                    ' Anwendungsinformationen setzen
    objObjectBag.SetShowSplash = Not bNotShowSplash                 ' Startparameter Splash anzeigen
    objObjectBag.SetAppDir = App.Path                               ' Anwendungsverzeichnis (ordner in dem die .exe liegt)
    objObjectBag.SetAppTitle = SZ_APPTITLE                          ' Anwendungtitel aus modConst
    objObjectBag.SetAppRegRoot = SZ_REGROOT                         ' Basis Registry Verz
    objObjectBag.SetCopyright = SZ_COPYRIGHT                        ' Copyright aus modConst
    objObjectBag.SetMajor = App.Major                               ' Major Version
    objObjectBag.SetMinor = App.Minor                               ' Minor Version
    objObjectBag.SetRevision = App.Revision                         ' Revision
    objObjectBag.SetIniFile = INI_FILENAME                          ' Dateiname INI File aus modConst (Alt)
    objObjectBag.setXMLFile = INI_XMLFILE                           ' Dateiname XML File aus modConst
    objObjectBag.setreadmepath = SZ_READMEFILE                      ' Name Readme Datei aus modConst
    objObjectBag.SetSupportMail = SZ_SUPPORTMAIL                    ' MailAdresse für Support aus modConst
    objObjectBag.SetWWW = SZ_WWW                                    ' ImternetAdresse aus modConst
    objObjectBag.SetE = bE
    objObjectBag.setexpert = bExpert
    objObjectBag.setconsole = bConsole
    objObjectBag.SetLeet = bLEET
    objObjectBag.SetMorse = bMorse
    szDeteils = "Anwendungs Informationen eingelesen."
                                                                    ' Sonstige benötigte Objecte aus ObjectBag holen
    Set objError = objObjectBag.GetErrorObj()                       ' Fehlerbehandlung
    szDeteils = "Fehlerbehandlungs Object Initialisiert."
    Set objRegTools = objObjectBag.GetRegToolsObj()                 ' Registry Tools
    szDeteils = "Registry Object Initialisiert."
    Set objTools = objObjectBag.GetToolsObj()                       ' Allg. Tools
    szDeteils = "Allg. Tools Object Initialisiert."
    If objObjectBag.InitSQLTools() Then                             ' SQL Tools initialisieren
        Set objSQLTools = objObjectBag.GetSqlToolsobj()
        szDeteils = "SQL Object Initialisiert."
    Else
        'Fehlerbehandlung ???
        szDeteils = "SQL Object nicht Initialisiert."
    End If
    If objObjectBag.InitOptions() Then                              ' Options Object initialisieren
        Set objOptions = objObjectBag.GetOptionsObj()
        szDeteils = "Options Object Initialisiert."
        objOptions.SetOptioniniPath = INI_OPTIONSINI                ' Optionsini bekannt geben
        Call objOptions.InitOptions                                 ' Options aus ini auslesen
        szDeteils = "Optionen eingelesen."
                                                                    ' Anwendungsspez Änderungen an den Optionen
        objError.SetProtFileName = objOptions.GetOptionByName(OPTION_APPLOG)  ' Gleich an error Obj Weitergeben
        objError.SetErrFileName = objOptions.GetOptionByName(OPTION_ERRLOG)   ' Gleich an error Obj Weitergeben
        bNotShowSplash = objOptions.GetOptionByName(OPTION_SPLASH)
        If bE Then bNotShowSplash = Not bE
    End If
    Call CheckFirstStart                                            ' Prüfen ob 1. Start oder neue Version
    szDeteils = "Splash anzeigen."
    Call ShowSplash(Not bNotShowSplash)                             ' Splash zeigen
    Call objObjectBag.ShowMSGForm(True, "Initialisiere Anwendung ...") ' Aktuelle Aktion im Spalsh oder MSG Form anzeigen
    If objObjectBag.GetWordVersion <> "" Then                       ' Wenn Word Installiert
        If objObjectBag.InitOfficeTools Then                        ' Office Schittstelle Initialisieren
            Set objOffice = objObjectBag.GetOfficeObj()
            szDeteils = "Office Object initialisiert."
        End If
    End If
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
    Call objError.Errorhandler(MODULNAME, "InitObjectBag", errNr, errDesc, szDeteils) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Sub

Private Sub ReadUserContext(szUserLogInName As String)
' Liest Userdaten aus der Tabelle USER001 ein Type User ein
    Dim szTab As String                                             ' Tabellen name
    Dim szNamefield As String                                       ' Feld mit Usernamen
    Dim szPWDField As String                                        ' Feld mit Kennwort
    Dim rsUser  As ADODB.Recordset                                  ' RS mit User daten
    Dim szSQL As String                                             ' SQL Statement
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktvieren
    szTab = objOptions.GetOptionByName("UserLoginTable")
    szNamefield = objOptions.GetOptionByName("UserNameField")
    szSQL = "SELECT * FROM " & szTab & " WHERE " & szNamefield & "='" & szUserLogInName & "'"
    Set rsUser = objDBconn.fillrs(szSQL, False)                     ' User Daten holen
    With User                                                       ' User daten einlesen
        .ID = objTools.checknull(rsUser.Fields("ID001").Value, "")  ' ID bei Datenbank
        .NTUsername = objObjectBag.GetUserName                      ' NT Username
        .Username = szUserLogInName                                 ' Anmeldename
        .tel = objTools.checknull(rsUser.Fields("tel001").Value, "") ' Telefon
        .fax = objTools.checknull(rsUser.Fields("fax001").Value, "") ' Fax
        .eMail = objTools.checknull(rsUser.Fields("email001").Value, "") ' Email
        .Vorname = objTools.checknull(rsUser.Fields("vorname001").Value, "") ' Vorname
        .Nachname = objTools.checknull(rsUser.Fields("Nachname001").Value, "") ' Nachanme
        .Fullname = .Vorname & " " & .Nachname                      ' Fullname
        .System = objTools.checknull(rsUser("SYSTEM001").Value, 0)  ' Ist admin
    End With
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
    Call objError.Errorhandler(MODULNAME, "ReadUserContext", errNr, errDesc)   ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Sub

Public Function AskForExit()
' Frag ob die Anwendung beendet werden soll -> falls JA ruft Appexit auf
    Dim szTitle As String                                           ' Msg Title
    Dim szMSG As String                                             ' Msg Text
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szTitle = "Beenden"                                             ' Meldungstitel festlegen
    szMSG = "Möchten Sie die " & objObjectBag.GetAppTitle & " beenden?" ' Meldung festlegen
    If objError.ShowErrMsg(szMSG, vbOKCancel + vbQuestion, szTitle) <> vbCancel Then
        Call AppExit                                                ' Anwendung Beenden
    End If
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Err.Clear                                                       ' Evtl. error clearen
Exit Function                                                            ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "AskForExit", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function AppExit()
' Beendet diese anwendung
On Error Resume Next                                                ' Fehlerbehandlung deaktiviren
    Call Unload(frmMain)                                            ' Main Form schliessen
    Err.Clear                                                       ' Errorhandling deak. falls schon geschehen
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    objError.WriteProt (PROT_APP_END)                               ' Beenden ins Log
    Call objOptions.SaveOptions                                     ' optionen in Registry Speichern
    Set objObjectBag = Nothing                                      ' Objectbag Schliessen
    End                                                             ' Anwendung beenden
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Err.Clear                                                       ' Evtl. error clearen
Exit Function                                                            ' Function Beenden
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

Public Function OpenMainForm()
' Zeigt das Hauptformular an
On Error Resume Next                                                ' Fehler behandlung deaktivieren
    Call objObjectBag.ShowMSGForm(True, "Lade Hauptfenster ...")    ' MSG Form Meldung
    objObjectBag.SetMainForm = frmMain                              ' Ref im ObjBag merken
    Set frmMain.SetDBConn = objDBconn                               ' Datenbank verbindung übergeben
    frmMain.Show                                                    ' Anzeigen
    Err.Clear                                                       ' Evtl. Error Clearen
End Function

Public Function ReadCmdParams()
' Start Parameter auslesen
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
            If Not bE Then bNotShowSplash = True
        Case UCase(CMD_EGG)
            bNotShowSplash = False
            bE = Not bNotShowSplash
        Case UCase(CMD_DOS), UCase(CMD_CMD), UCase(CMD_CONSOLE)     ' Console
            bConsole = True
        Case UCase(CMD_AUTOCON)                                     ' Autoconnect
            bAutoConnect = True
        Case UCase(CMB_EXPERT)
            bExpert = True                                          ' Experten modus
        Case UCase(CMD_MORSE)
            bMorse = True
        Case UCase(CMB_LEET)
            bLEET = True
        Case UCase(CMD_HELP), UCase(CMD_HELP2)
'            Call MsgBox(CMD_HELP_TXT, vbInformation, SZ_APPTITLE & " " _
                    & App.Major & "." & App.Minor & "." & App.Revision)
            End                                                      ' Anwendung beenden
        Case Else
        
        End Select
    Next                                                            ' nächstes Array Item
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
    Call objError.Errorhandler(MODULNAME, "ReadCmdParams", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function CheckFirstStart() As Boolean
' Überprüft aob der User die Anwendung das erste mal statet
' und führt besondere Funktionen für den ersten start aus
    Dim bFirstStart As Boolean                                      ' True wenn 1. Start
    Dim bNewVersion As Boolean                                      ' True wenn neue Version
    Dim szVersion As String                                         ' Versionsnummer
    Dim szTmp As String
On Error GoTo Errorhandler                                          ' Fehler behandlung aktiviernen
    bFirstStart = True                                              ' Erstmal immer true
    szTmp = objRegTools.ReadRegValue("HKCU", "SOFTWARE\" & objObjectBag.GetAppTitle, "FirstStart")
    If szTmp <> "" Then bFirstStart = CBool(szTmp)
    szVersion = objRegTools.ReadRegValue("HKCU", "SOFTWARE\" & objObjectBag.GetAppTitle, "Version", objObjectBag.GetAppVersion)
    szTmp = ""
    ' VersionsSprung
    If szVersion <> objObjectBag.GetAppVersion Then bNewVersion = True
    If bFirstStart Then                                             ' Falls 1. Start
        bNewVersion = True                                          ' Dann auch neue Version
        Call objRegTools.WriteRegValue("HKCU", "SOFTWARE\" & objObjectBag.GetAppTitle, "FirstStart", False)
        ' Protokollpfad in Eigene dateien Festlegen
        szTmp = objOptions.GetOptionByName(OPTION_APPLOG)           ' OPTION_APPLOG = "ProtokollFile"
        If szTmp = "" Then
            Call objOptions.SetOptionByName(OPTION_APPLOG, objObjectBag.GetPersonalDir())
        Else
            If InStr(szTmp, "\") = 0 Then
                Call objOptions.SetOptionByName(OPTION_APPLOG, objObjectBag.GetPersonalDir() & "\" & szTmp)
            End If
        End If
        szTmp = ""
        ' Errorlog in Eigene Dateien festlegen
        szTmp = objOptions.GetOptionByName(OPTION_ERRLOG)           ' OPTION_ERRLOG = "ErrorProtokollFile"
        If szTmp = "" Then
            Call objOptions.SetOptionByName(OPTION_ERRLOG, objObjectBag.GetPersonalDir())
         Else
            If InStr(szTmp, "\") = 0 Then
                Call objOptions.SetOptionByName(OPTION_ERRLOG, objObjectBag.GetPersonalDir() & "\" & szTmp)
            End If
        End If
        szTmp = ""
        ' Ablage Pfad in Eigene Dateien festlegen wenn leer
        szTmp = objOptions.GetOptionByName(OPTION_ABLAGE)           ' OPTION_ABLAGE = "Ablageverzeichnis"
        If szTmp = "" Then
            Call objOptions.SetOptionByName(OPTION_ABLAGE, objObjectBag.GetPersonalDir())
        End If
        szTmp = ""
        ' Vorlagenverz Pfad in App verz festlegen wenn leer
         szTmp = objOptions.GetOptionByName(OPTION_TEMPLATES)       ' OPTION_TEMPLATES = "Vorlagenverzeichnis"
        If szTmp = "" Then
            Call objOptions.SetOptionByName(OPTION_TEMPLATES, objObjectBag.Getappdir() & "\Vorlagen")
        End If
        szTmp = ""
        objError.SetProtFileName = objOptions.GetOptionByName(OPTION_APPLOG) ' Gleich an error Obj Weitergeben
        objError.SetErrFileName = objOptions.GetOptionByName(OPTION_ERRLOG)  ' Gleich an error Obj Weitergeben
        'Call objTools.ShellExec("regedit", "/s ColumsSettings.reg", 0)
    End If
    
    If bNewVersion Then
        ' Option Last Connetion löschen
        Call objRegTools.WriteRegValue("HKCU", "SOFTWARE\" & objObjectBag.GetAppTitle, "LastConnection", "") ' Last Connection löschen
        ' Version setzen
        Call objRegTools.WriteRegValue("HKCU", "SOFTWARE\" & objObjectBag.GetAppTitle, "Version", objObjectBag.GetAppVersion)
        Call objOptions.InitOptions(True)                           ' Optionen neueinlesen
        ' Anwendungsspez Änderungen an den Optionen
        objError.SetProtFileName = objOptions.GetOptionByName(OPTION_APPLOG) ' Gleich an error Obj Weitergeben
        objError.SetErrFileName = objOptions.GetOptionByName(OPTION_ERRLOG)  ' Gleich an error Obj Weitergeben
        ' Columns in Reg löschen
        'Call objRegToolsRegKeyDelete("HKCU", "Columns")
        Call objTools.ShellExec("regedit", "/s ColumsSettings.reg", 0)
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
    Call objError.Errorhandler(MODULNAME, "CheckFirstStart", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Sub ShowHelp()                                               ' Zeigt die Hilfe an
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call objTools.HTMLHelp_ShowTopic(objObjectBag.Getappdir() & SZ_HELPFILE)
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Public Function OpenNewDB()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call objDBconn.GetAdodbConn(2, "", "", "", "")                  ' neue DB Verbindung ohne Parameter (nur SQL)
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Public Sub ShowReadMe()                                             ' Zeigt ReadMeDatei an
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call objObjectBag.ShowReadMe
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Public Sub ShowAbout(Optional szDesc As String, _
        Optional bWithOfficeInfo As Boolean, _
        Optional DBConn As Object)                                  ' Zeigt das About Form an
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call objObjectBag.ShowAbout(szDesc, bWithOfficeInfo, DBConn)
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Public Function ShowChangePWDForm(szUsername As String, bCancel As Boolean, _
        Optional bShowNextLogin As Boolean, Optional bChangeOnNextLogin As Boolean) As String
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    ShowChangePWDForm = objObjectBag.ShowFormChangePWD(szUsername, bCancel, _
                bShowNextLogin, bChangeOnNextLogin)                 ' An Object Bag durchreichen
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Public Function ShowSplash(bVisible As Boolean)
' Zeigt Splash Form an und blendet es aus
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call objObjectBag.ShowSplash(bVisible)                          ' Spalsh form öffen/Schliessen
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Public Function ShowOptions()
' Zeigt Options Form an und blendet es aus
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call objOptions.ShowOptions(bExpert)                            ' Options Form anzeigen
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Public Function ShowSearch(DBConn As Object, RootKey As String, SearchField As String, _
        Optional Suchtext As String, Optional OptWhereID As String, Optional SuchTitel As String) As String
' Zeigt die Suchfunktion an
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    ShowSearch = objObjectBag.ShowSearch(DBConn, RootKey, SearchField, Suchtext, OptWhereID, SuchTitel)   ' Such form anzeigen
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Public Function ReportBug()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call objObjectBag.ReportBug                                     ' Fehlermelde Aktion
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Public Function DeleteDS(szRootkey As String, szID As String)
' Reich die Aktion DS Löschen ans MainForm durch
    Dim fm As frmMain                                               ' Referenz auf frmMain
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Set fm = objObjectBag.getMainForm                               ' Referenz auf frmMain aus OBag holen
    Call fm.DeleteDS(szRootkey, szID)                               ' Funktion DeleteDS aufrufen
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Public Function NewDS(szRootkey As String, Optional bDialog As Boolean, _
    Optional frmParentForm As Form)
' Reich die Aktion Neuen DS ans MainForm durch
    Dim fm As frmMain                                               ' Referenz auf frmMain
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Set fm = objObjectBag.getMainForm                               ' Referenz auf frmMain aus OBag holen
    If frmParentForm Is Nothing Then
        Call fm.OpenEditForm(szRootkey, "", fm, bDialog)            ' Funktion OpenEditForm ohne ID aufrufen mit MainForm als Parent
    Else
        Call fm.OpenEditForm(szRootkey, "", frmParentForm, bDialog) ' Funktion OpenEditForm ohne ID aufrufen mit Parentform als Parent
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Public Function EditDS(szRootkey As String, szID As String, Optional bDialog As Boolean, _
        Optional frmParentForm As Form)
' Reich die Aktion DS bearbeiten ans MainForm durch
    Dim fm As frmMain                                               ' Referenz auf frmMain
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Set fm = objObjectBag.getMainForm                               ' Referenz auf frmMain aus OBag holen
    If frmParentForm Is Nothing Then
        Call fm.OpenEditForm(szRootkey, szID, fm, bDialog)          ' Funktion OpenEditForm mit ID aufrufen mit MainForm als Parent
    Else
        Call fm.OpenEditForm(szRootkey, szID, frmParentForm, bDialog) ' Funktion OpenEditForm mit ID aufrufen mit ParentForm als Parent
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Public Function ShowWordDoc(DocID As String)
' Anhand der DocID wird in der TAb DOC018 der Dokumenten DS ermittelt
' Dieser enthält den Dunkumenten namen im Filesystem Evtl. mit unter ordner
' relativ zum Ablageverzeichnis
    Dim rsDoc As ADODB.Recordset                                    ' Dokumenten RS
    Dim szSQL As String                                             ' SQL Statement
    Dim szAblgenVerz As String                                      ' Ablageverz.
    Dim szDocVerz As String                                         ' Unterverz. des Ablage Verz
    Dim objDoc As Object                                            ' Word Dokumenten Object
    Dim szMSG As String                                             ' Test einer evtl. Fehlermeldung
    Dim szDetails As String                                         ' Zusatzinfos für Fehlerbehnadlung
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If DocID = "" Then GoTo exithandler                             ' keine ID -> fertig
    szDetails = "DocID: " & DocID & vbCrLf
    szSQL = "SELECT * FROM DOC018 WHERE ID018='" & DocID & "'"      ' SQL Festlegen
    Set rsDoc = objDBconn.fillrs(szSQL, False)                      ' Daten Holen
    If rsDoc Is Nothing Then GoTo exithandler                       ' Keine Daten -> raus
    If rsDoc.RecordCount = 0 Then GoTo exithandler                  ' Keine Daten -> raus
    szAblgenVerz = objOptions.GetOptionByName(OPTION_ABLAGE) & "\"  ' Ablageberz. ermitteln
    szDetails = szDetails & "Ablageverz.: " & szAblgenVerz & vbCrLf
    szDocVerz = rsDoc.Fields("DOCPATH018").Value                    ' evtl. Unterverz mit Doknamen aus RS
    szDetails = szDetails & "Dokverz.: " & szDocVerz & vbCrLf
    If szDocVerz = "" Or szAblgenVerz = "" Then GoTo exithandler
    If objTools.FileExist(szAblgenVerz & szDocVerz) Then            ' Prüfen ob Doc existiert
        szDetails = szDetails & "Doc. existiert: Wahr" & vbCrLf
        Set objDoc = objOffice.OpenNewWordDoc(szAblgenVerz & szDocVerz, False) ' Öffnen
    Else                                                            ' Sonst
        szDetails = szDetails & "Doc. existiert: Falsch" & vbCrLf
        ' Fehlemeldung
        szMSG = "Das Dokument '" & szAblgenVerz & szDocVerz & "' kann nicht gefunden werden." & vbCrLf & _
                "Überprüfen Sie in Ihnen Optionen den Pfad des Ablageverzeichnisses." & vbCrLf & _
                "bzw. wenden Sie sich na Ihren Systembetreuer."
        Call objError.ShowErrMsg(szMSG, vbCritical, "Dokument nicht gefunden.")
    End If
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Err.Clear                                                       ' Evtl. error clearen
Exit Function                                                            ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ShowWordDoc", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ImportWordDocFolder(ThisDBCon As Object, _
        Optional PersID As String, _
        Optional StellenID As String, _
        Optional AuschrID As String)
    Dim szDetails As String                                         ' Details für fehlerbehandlung
    Dim szImportFolder As String                                    ' Ordner der Importiert werden Soll
    Dim szFilelist As String                                        ' Dateiliste
    Dim Filenames() As String                                       ' Array mit dateinamen
    Dim i As Integer                                                ' Array Counter
    Dim x As Integer                                                ' Import Counter
    Dim bSaveAsDocx As Boolean                                      ' DocxFormat zulassig
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
' MW 01.09.11 {
    bSaveAsDocx = objOptions.GetOptionByName(OPTION_DOCX)           ' Lassen wir Docx Format zu
' MW 01.09.11 }
    Call objError.WriteProt(PROT_IMPORT_START)                      ' Protokolieren
    If SelectImport(ThisDBCon, "", "", szImportFolder, _
            PersID, StellenID, AuschrID, True) Then                 ' Wenn Import Folder ausgewählt
        If Not objTools.FolderExist(szImportFolder) Then GoTo exithandler ' Prüfen ob Quell ordner existiert
        Call objError.WriteProt("   Importordner: " & szImportFolder) ' Protokolieren
        szFilelist = objTools.GetFileList(szImportFolder, ";")      ' Dateiliste ermitteln
        If szFilelist = "" Then GoTo exithandler                    ' Keine Dateiliste -> Fertig
        Filenames = Split(szFilelist, ";")                          ' in Array Spalten
        For i = 0 To UBound(Filenames)                              ' Array Duchlaufen
            If Filenames(i) <> "" Then                              ' Leere Dateinamen können wir nicht gebrauchen
                If UCase(Right(Filenames(i), 4)) = ".DOC" _
                        Or (UCase(Right(Filenames(i), 5)) = ".DOCX" And bSaveAsDocx) Then ' nur word doks
                    If SaveNewDoc(ThisDBCon, , szImportFolder & Filenames(i), "", PersID, StellenID, AuschrID, True) <> "" Then
                        x = x + 1                                   ' Import Counter hochzählen
                        Call objError.WriteProt("   Datei: " & Filenames(i))  ' Protokolieren
                    Else                                            ' Sonst
                        Call objError.WriteProt("   ImportFehler: " & Filenames(i))  ' Protokolieren
                    End If

                Else
                    Call objError.WriteProt("   Nicht Importiert: " & Filenames(i))  ' Protokolieren
                End If
            End If
        Next i                                                      ' Nächstes Array item
    End If
    Call objError.WriteProt(PROT_IMPORT_END)                        ' Protokolieren
    If x > 0 Then
        Call objError.ShowErrMsg(CStr(x) & " Dokumente importiert", vbInformation, "Import beendet") ' Meldung an user
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
    Call objError.Errorhandler(MODULNAME, "ImportWordDocFolder", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ImportWordDoc(ThisDBCon As Object, _
        Optional PersID As String, _
        Optional StellenID As String, _
        Optional AuschrID As String)
' Importiert ein einzelnes Word dokument
    Dim szDetails As String                                         ' Details fü fehlerbehandlung
    Dim szImportFolder As String                                    ' Ordner der Importiert werden Soll (noch nicht im einsatz)
    Dim szImportFile As String                                      ' Zu Importierendes File
    Dim szAblagePath As String                                      ' Pfad der Datei ablage
    Dim szFilename As String                                        ' Dateiname des Importfiles (ohne Pfad)
    Dim szDocPath As String                                         ' Pfad Segment der Datei in der Ablage
    Dim szDocTitle As String                                        ' Neuer Dokumentname
    Dim szAZ As String                                              ' Aktenzeichen für Dateinamen und Ablage pfad
    Dim szFileSuffix As String                                      ' Suffix des Importfiles
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Call objError.WriteProt(PROT_IMPORT_START)                      ' Protokolieren
    If SelectImport(ThisDBCon, "", szImportFile, szImportFolder, _
            PersID, StellenID, AuschrID, False) Then                ' Wenn Import File ausgewählt
        If Not objTools.FileExist(szImportFile) Then GoTo exithandler ' Prüfen ob Quell File existiert
        If SaveNewDoc(ThisDBCon, , szImportFile, "", PersID, StellenID, AuschrID) <> "" Then
            Call objError.ShowErrMsg("Dokument wurde in der Datenbank gespeichert.", _
                            vbInformation + vbOKOnly, "Word")       ' Meldung Dokument erstellt und abgespeichert
            Call objError.WriteProt(PROT_IMPORT_END)                ' Protokolieren
        End If
    Else                                                            ' Kein Import File/Folder ausgewählt
        Call objError.WriteProt("Dokumenten Import abgebrochen")    ' Protokolieren
        Call objError.WriteProt(PROT_IMPORT_END)                    ' Protokolieren
        GoTo exithandler                                            ' -> Fertig
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
    Call objError.Errorhandler(MODULNAME, "ImportWordDoc", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function SaveNewDoc(ThisDBCon As Object, _
        Optional objDoc As Object, _
        Optional szSourceFile As String, _
        Optional szVorlage As String, _
        Optional PersID As String, _
        Optional StellenID As String, _
        Optional AuschrID As String, _
        Optional bHoldFilename As Boolean) As String
' Speichert neues Dokument (entweder ObjDoc bei Schreibwerk oder szSourceFile bei Import
' Leegt Ablagepad anhand vorhandener ID fest
    Dim szDetails As String                                         ' Details für fehlerehandlung
    Dim szDocTitle As String                                        ' Neuer Dokumenten Name
    Dim szAblagePath As String                                      ' Pfad der Dokumenten abblage
    Dim szDocPath As String                                         ' Evtl. Pfad inerhalb der Ablage
    Dim szFilename As String                                        ' Dateiname bei Import
    Dim szFileSuffix As String                                      ' FileSuffix bei import
    Dim szAZ As String                                              ' AZ für Ablagepfad und Dateiname
    Dim szBezirk As String                                          ' Bezirk für Ablagepfad und Dateiname
    Dim szPersName As String                                        ' Personen name für Ablagepfad und Dateiname
    Dim bSaveAsDocx As Boolean                                      ' Als Docx (word 2007) Speichern
    Dim szMSG As String                                             ' Meldungstext
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If szSourceFile = "" And objDoc Is Nothing Then GoTo exithandler ' Nix zu tun
    szAblagePath = objOptions.GetOptionByName(OPTION_ABLAGE) & "\"  ' AblagePfad ermitteln
    If szAblagePath = "\" Then                                      ' Haben wir einen Speicherpfad ?
        szMSG = "Es ist kein Ablageverzeichnis hinterlegt. Die Ausführung wird abgebrochen."
        Call objError.ShowErrMsg(szMSG, vbExclamation, "Fehlerhafte Konfiguration", False, "")   ' Meldung ausgeben
        GoTo exithandler                                            ' Fertig
    End If
' MW 01.09.11 {
    bSaveAsDocx = objOptions.GetOptionByName(OPTION_DOCX)           ' Soll im Docx Format gespeichert werden
' MW 01.09.11 }
    If szSourceFile <> "" Then
        szFilename = objTools.GetFileNameFromPath(szSourceFile)     ' Filename aus Pfad holen
' MW 01.09.11 {
'        szFileSuffix = Right(szFilename, 3)                         ' File suffix ermitteln
        szFileSuffix = Right(szFilename, 4)                         ' File suffix ermitteln
        If Left(szFileSuffix, ".") Then szFileSuffix = Right(szFileSuffix, 3)
' MW 01.09.11 }
        szVorlage = szFilename
    Else
        szFileSuffix = "doc"                                        ' Ansonsten Doc
    End If
' MW 01.09.11 {
    If bSaveAsDocx Then szFileSuffix = "docx"                       ' Wenn docx dann filesuffix anpassen
' MW 01.09.11 }
    If bHoldFilename Then                                           ' dateinamen beibehalten (import)
        szDocTitle = szFilename
    Else
        szDocTitle = GetDocTitleFromIDs(ThisDBCon, PersID, StellenID, AuschrID, True) ' Dokumentennamen zusammen setzen
    End If
    If szDocTitle = "" Then                                         ' Wenn weder PersId noch stellen ID notwendig kann leerer DocTitel vorkommen
        szDocTitle = Replace(szVorlage, "." & szFileSuffix, "") & " " & Now()
    End If
    szDocTitle = PrepareDocTitle(szDocTitle, szFileSuffix)          ' Sonderzeichen aus Docnamen löschen
    szDetails = "Docname: " & szDocTitle                            ' Details für fehlerbehandlung
    szAZ = GetAZfromIds(objDBconn, PersID, StellenID, AuschrID)     ' AZ für Pfad ermitteln
    szAZ = Trim(szAZ)                                               ' Aktenzeichen trimmen
    szBezirk = GetBezirkFromIDs(objDBconn, PersID, StellenID, AuschrID) ' Bezirk für Pfad ermiiteln
    szBezirk = Trim(szBezirk)                                       ' Bezirk trimmen
    If PersID <> "" Then                                            ' Keine Personen ID
        szPersName = ThisDBCon.GetValueFromSQL("SELECT TOP 1 NACHNAME010 FROM RA010 " & _
                " WHERE ID010='" & PersID & "'")                    ' Nachname für Pfad ermitteln
    End If
    szPersName = Trim(szPersName)                                   ' Personen name Trimmen
    If szAZ = "" Then
        ' Müssen wir hir was tun ?
    End If
    szDetails = "AZ: " & szAZ & " Bezirk: " & szBezirk & " Pers: " & szPersName ' Details für fehlerbehandlung
'    szAblagePath = objOptions.GetOptionByName(OPTION_ABLAGE) & "\"  ' AblagePfad ermitteln
    If szAZ <> "" Then szDocPath = szAZ & "\"                       ' Ablage pfad um AZ ergänzen
    If szBezirk <> "" Then szDocPath = szDocPath & szBezirk & "\"   ' Ablage pfad um Bezirk ergänzen
    If szPersName <> "" Then szDocPath = szDocPath & szPersName & "\"   ' Ablage pfad um Nachname ergänzen
    szDocPath = PrepareFolderPath(szDocPath)                        ' Evtl. Sonderzeichen aus Docpath entfernen
    szDetails = "Ablage: " & szAblagePath & szDocPath               ' Details für fehlerbehandlung
    If objTools.CheckPath(szAblagePath & szDocPath, True) Then      ' Prüfen Verz. existiert sonst anlegen
        szDetails = "Zielpfad existiert"                            ' Details für fehlerbehandlung
        If szSourceFile <> "" Then                                  ' Docimport vom Quellfile
            If objTools.FileCopy(szSourceFile, szAblagePath & szDocPath & _
                szDocTitle, True) Then                              ' Datei Kopieren
                szDetails = "Document unter " & szAblagePath & szDocPath & szDocTitle & " gespeichert"
                Call objError.WriteProt(szDetails)                  ' Erfolg Protokolieren
            Else
                GoTo DocSaveError                                   ' Datei konte nicht importiert werden
            End If
        ElseIf Not objDoc Is Nothing Then                           ' Word object Speichern
            If objOffice.SaveWordDoc(objDoc, szAblagePath & _
                szDocPath & szDocTitle) Then                        ' Doc im Filesystem speichern
                szDetails = "Document unter " & szAblagePath & szDocPath & szDocTitle & " gespeichert"
                Call objError.WriteProt(szDetails)                  ' Erfolg Protokolieren
            Else
                GoTo DocSaveError                                   ' DocObj konte nicht gespeichert werden
            End If
        Else
            GoTo exithandler                                        ' Fertig
        End If
        If SaveDocRecord(PersID, StellenID, AuschrID, szVorlage, _
                        szDocTitle, szDocPath) Then                 ' INSERT Statement ausführen (DS in DOC018 Speichern)
                    szDetails = "Document DS gespeichert"
            Call objError.WriteProt(szDetails)                      ' Erfolg Protokolieren
        Else
            GoTo DocSaveError                                       ' Datei konte nicht gespeichert werden
        End If
    Else                                                            ' Ablage Verz konnte nicht gefunder oder erstellt werden
        GoTo DocSaveError                                           ' Datei konte nicht gespeichert werden
    End If
    SaveNewDoc = szAblagePath & szDocPath & szDocTitle              ' Erfolg melden & Pfad zurück
    
exithandler:

Exit Function
DocSaveError:
On Error Resume Next
    Call objTools.FileDelete(szAblagePath & szDocPath & szDocTitle, True)
    Call objObjectBag.ShowMSGForm(False, "")                        ' MSG Form ausbelnden
    Call objError.ShowErrMsg("Beim erstellen des Dokuments ist ein Fehler aufgetreten!" & vbCrLf, _
            "Das Dokument wurde nicht gespeichert.", _
            vbInformation + vbOKOnly, "Word")                       ' Meldung Dokument Nicht erstellt oder abgespeichert
    Call objOffice.CloseWordObj                                     ' Word dichtmachen
    GoTo exithandler                                                ' Fertig
        
Exit Function
Errorhandler:
On Error Resume Next
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SaveNewDoc", errNr, errDesc)
    Resume exithandler
End Function

Public Function WriteWord(Optional szTemplate As String, _
                Optional PersID As String, _
                Optional StellenID As String, _
                Optional AuschrID As String, _
                Optional bHide As Boolean, _
                Optional CurForm As Form)
' Mittels dieser Fkt wird eine Vorlage (kann als Parameter angegeben werden) mit daten aus Der Datenbank gefüllt
    Dim objDoc As Object                                            ' Word Dokumenten Object
    Dim szIniFilePath As String                                     ' XML Konfig Datei
    Dim rsAus As ADODB.Recordset
    Dim rsListeStellen As ADODB.Recordset
    Dim rsListeBewerber As ADODB.Recordset
    Dim rsListeBewerberA As ADODB.Recordset
    Dim rsListeBewerberZ As ADODB.Recordset
    Dim rsRA As ADODB.Recordset                                     ' Recordset mit Personen Daten
    Dim rsBewerber As ADODB.Recordset                               ' Recordset mit Bewerber Daten  (inclusive Stelle)
    Dim rsMitbewerber As ADODB.Recordset                            ' Recordset mit Mitbewerber Daten
    Dim rsZugesagt As ADODB.Recordset                               ' Recordset mit Zugesagten Mitbewerber Daten
    Dim rsAbgesagt As ADODB.Recordset                               ' Recordset mit Abgesagten Mitbewerber Daten
    Dim rsStellen As ADODB.Recordset                                ' Recordset mit Stellen daten
    Dim rsUser As ADODB.Recordset                                   ' Recordset mit User Daten
    Dim szVorlagePath As String                                     ' kompletter Vorlagen Pfad
    Dim szDestPath As String                                        ' Pad unter dem das Doc gespeichert wurde
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    Dim szMSG As String                                             ' Meldungstext
    Dim bPersIDRequ As Boolean                                      ' ID010 (Personen ID) ist erfoderlich
    Dim bStellenIDRequ As Boolean                                   ' ID012 (Stellen ID) ist erfoderlich
    Dim bAussrIDRequ As Boolean                                     ' ID020 (Ausschreibung ID) ist erfoderlich
    Dim Step As Integer                                             ' Schrittweite Für Progess
    Dim n As Integer                                                ' Counter
    Dim szReqData As String
    Dim szReqLists As String
    Dim szReqListsArray() As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szVorlagePath = objOptions.GetOptionByName(OPTION_TEMPLATES) & "\"  ' Vorlagen Pfad holen
    If szVorlagePath = "\" Then                                      ' Kein Vorlageverz.
        szMSG = "Es ist kein Vorlagenverzeichnis hinterlegt. Die Ausführung wird abgebrochen."
        Call objError.ShowErrMsg(szMSG, vbExclamation, "Fehlerhafte Konfiguration", False, "")   ' Meldung ausgeben
        GoTo exithandler                                            ' -> fertig
    End If
    szIniFilePath = objObjectBag.Getappdir & objObjectBag.GetXMLFile ' XML inifile festlegen
    szDetails = "Starte SAT"                                        ' Details für Fehlerbehandlung
    Call objError.WriteProt(PROT_DOC_START)                         ' SAT Start Protokolieren
    If szTemplate = "" Then                                         ' Keine Vorlage angegeben
        szTemplate = SelectDocTemplate()                            ' Vorlage auswählen
    End If
    If szTemplate = "" Then GoTo exithandler                        ' Immer noch Keine Vorlage -> fertig
    
    szDetails = "Vorlage: " & szVorlagePath & szTemplate            ' Details für Fehlerbehandlung
    Call objError.WriteProt("Template: " & szVorlagePath & szTemplate) ' Vorlage Protokolieren
    Set objDoc = objOffice.OpenNewWordDoc(szVorlagePath & szTemplate, True)      ' Doc mit vorlage öffnen (nicht anzeigen)
    If objDoc Is Nothing Then GoTo WordError                        ' kein Dok Spezielle fehlerbehandlung und -> Fertig
    szDetails = "Vorlage geöffnet"                                  ' Details für Fehlerbehandlung
    If Not CheckRequieredIDForDoc(objDoc, bPersIDRequ, PersID, bStellenIDRequ, StellenID, bAussrIDRequ, AuschrID, szReqData) And _
        Not CheckRequieredListForDoc(objDoc, szReqLists) Then _
            GoTo exithandler                                        ' Vorlage nach benötigten ID überprüfen
    'Public Function ShowMSGForm(bShow As Boolean, szMSGtext As String, Optional ProgBarMax As Integer, Optional szAction As String)
    Call objObjectBag.ShowMSGForm(True, "Dokumenterstellung", 0, "Vorlage wird geladen") ' MSG Form anzeigen mit ProgBar
    Call objError.WriteProt("Address: " & PersID)                   ' Empfänger Protokolieren
    Call objError.WriteProt("Stelle: " & StellenID)                 ' Empfänger Protokolieren
    If Not CurForm Is Nothing Then
        CurForm.MousePointer = vbHourglass                          ' Sanduhr
    Else
        frmMain.MousePointer = vbHourglass                          ' Sanduhr
    End If
    Call objObjectBag.ShowMSGForm(True, "Dokumenterstellung", 0, "Datenbank zugriff") ' Meldungsform unzeigen
    Set rsUser = GetDocData(szIniFilePath, "User", " ID001 = '" _
            & User.ID & "'")                                        ' Absender Daten holen
    If PersID <> "" And bPersIDRequ Then                            ' Wenn PersID Vorh. und Notwendig
        Set rsRA = GetDocData(szIniFilePath, "Empfaenger", _
                " PersID = '" & PersID & "'")                       ' Personen Daten holen
        If StellenID <> "" Then                                     ' Wenn zusätzlich Stellen ID
            Set rsBewerber = GetDocData(szIniFilePath, "Bewerber", " PersID = '" & PersID _
                    & "' AND STELLENID = '" & StellenID & "'")      ' Dann Bewerber Daten holen
        End If
    End If
    If StellenID <> "" And bStellenIDRequ Then                      ' Wenn Stellen ID Vorh. und Notwendig
        Set rsStellen = GetDocData(szIniFilePath, "Stellen", _
                    " StellenID = '" & StellenID & "'")             ' Stellen Daten holen
        'Mitbewerber;AMitbewerber;ZMitbewerber;
        If InStr(szReqData, ";Mitbewerber;") > 0 Then
            Set rsMitbewerber = GetDocData(szIniFilePath, "Mitbewerber", _
                    " StellenID = '" & StellenID & "'")             ' Mitbewerber Daten holen
        End If
        If InStr(szReqData, "AMitbewerber;") > 0 Then
            Set rsAbgesagt = GetDocData(szIniFilePath, "MitbewerberAbsage", _
                    " StellenID = '" & StellenID & "'")             ' Abgesagte Mitbewerber Daten holen
        End If
        If InStr(szReqData, "ZMitbewerber;") > 0 Then
             Set rsZugesagt = GetDocData(szIniFilePath, "MitbewerberZusage", _
                    " StellenID = '" & StellenID & "'")             ' Zugesagte Mitbewerber Daten holen
        End If
        If InStr(szReqLists, "ListeBewerber_") > 0 Then
            Set rsListeBewerber = GetDocListData(szIniFilePath, "ListeBewerber" _
                    , "'" & StellenID & "'")                        ' Daten Aller Bewerber Holen
        End If
        If InStr(szReqLists, "ListeBewerberAbgelehnt") > 0 Then
            Set rsListeBewerberA = GetDocListData(szIniFilePath, "ListeBewerberAbgelehnt" _
                    , "'" & StellenID & "'")                        ' Daten Aller Abgelenten Bewerber holen
        End If
        If InStr(szReqLists, "ListeBewerberZugesagt") > 0 Then
            Set rsListeBewerberZ = GetDocListData(szIniFilePath, "ListeBewerberZugesagt" _
                    , "'" & StellenID & "'")                        ' Daten Aller Zugesagten Bewerber holen
        End If
    End If

    If AuschrID <> "" And bAussrIDRequ Then                         ' Wenn Ausschreibungs ID Vorh. und Notwendig
        Set rsAus = GetDocData(szIniFilePath, "Ausschreibung", _
                    " AusschreibungID = '" & AuschrID & "'")        ' Daten der Ausschreibung Holen
        If InStr(szReqLists, "ListeAusgeschreibeneStellen") > 0 Then
            Set rsListeStellen = GetDocListData(szIniFilePath, "ListeAusgeschriebeneStellen" _
                    , "'" & AuschrID & "'")                         ' Daten der Stellenausschreibungen holen
        End If
    End If
    szDestPath = SaveNewDoc(objDBconn, objDoc, , szTemplate, PersID, StellenID, AuschrID)   ' Doc Speichern
    If szDestPath <> "" Then
        If Not rsListeStellen Is Nothing Then
            Call objOffice.SetDataInDocNeu(objDoc, rsListeStellen, , "Dokumenten erstellung", _
                    "Stellenliste wird eingefügt")                  ' Liste aller Stellen ins Doc
        End If
        If Not rsAus Is Nothing Then
            Call objOffice.SetDataInDocNeu(objDoc, rsAus, , "Dokumenten erstellung", _
                    "Daten der Ausschreibung werden eingefügt")     ' Ausschreibungs Daten ins Doc
        End If
        If Not rsUser Is Nothing Then '
            Call objOffice.SetDataInDocNeu(objDoc, rsUser, , "Dokumenten erstellung", _
                    "Benutzerdaten werden eingefügt")               ' Absender Daten ins Doc
            szDetails = "Userdaten eingefügt"
        End If
        If Not rsListeBewerber Is Nothing Then
            Call objOffice.SetDataInDocNeu(objDoc, rsListeBewerber, , "Dokumenten erstellung", _
                    "Daten der Mitbewerber werden eingefügt")       ' Mitbewerber Daten ins Doc
        End If
        
        If Not rsListeBewerberA Is Nothing Then
            Call objOffice.SetDataInDocNeu(objDoc, rsListeBewerberA, , "Dokumenten erstellung", _
                    "Daten der abgelenten Mitbewerber werden eingefügt") ' Abgelente Mitbewerber Daten ins Doc
        End If
        
        If Not rsListeBewerberZ Is Nothing Then
            Call objOffice.SetDataInDocNeu(objDoc, rsListeBewerberZ, , "Dokumenten erstellung", _
                    "Daten der abgelenten Mitbewerber werden eingefügt") ' Zugesagte Mitbewerber Daten ins Doc
        End If
        If Not rsStellen Is Nothing Then
            Call objOffice.SetDataInDocNeu(objDoc, rsStellen, , "Dokumenten erstellung", _
                    "Stellendaten werden eingefügt")                ' Stellen Daten ins Doc
            szDetails = "Stellen Daten eingefügt"
        End If
        If Not rsRA Is Nothing Then
            Call objOffice.SetDataInDocNeu(objDoc, rsRA, , "Dokumenten erstellung", _
                    "Empfängerdaten werden eingefügt")              ' Empfänger Daten ins Doc
            szDetails = "Empfänger Daten eingefügt"
        End If
        If Not rsBewerber Is Nothing Then
            Call objOffice.SetDataInDocNeu(objDoc, rsBewerber, , "Dokumenten erstellung", _
                    "Bewerberdaten werden eingefügt")               ' Bewerber Daten ins Doc
            szDetails = "Bewerber Daten eingefügt"
        End If
        If Not rsMitbewerber Is Nothing Then
            Call objOffice.SetDataInDocNeu(objDoc, rsMitbewerber, True, "Dokumenten erstellung", _
                    "Mitbewerberdaten werden eingefügt")            ' Mitbewerber Daten ins Doc
            szDetails = "Mitbewerber Daten eingefügt"
        End If
        If Not rsZugesagt Is Nothing Then
            Call objOffice.SetDataInDocNeu(objDoc, rsZugesagt, True, "Dokumenten erstellung", _
                    "Daten abgesagter Mitbewerber werden eingefügt") ' Absender Daten ins Doc
            szDetails = "Zugesagte Bewerber Daten eingefügt"
        End If
        If Not rsAbgesagt Is Nothing Then '
            Call objOffice.SetDataInDocNeu(objDoc, rsAbgesagt, True, "Dokumenten erstellung", _
                    "Daten zugesagter Mitbewerber werden eingefügt") ' Absender Daten ins Doc
            szDetails = "Abgesagte Bewerber Daten eingefügt"
        End If
                
        If objOffice.SaveWordDoc(objDoc, szDestPath) Then           ' Doc mit änderungen im Filesystem speichern
            Call objObjectBag.ShowMSGForm(False, "")                ' MSG Form ausblenden
            Call objError.ShowErrMsg("Dokument wurde erstellt und gespeichert.", _
                    vbInformation + vbOKOnly, "Word")               ' Meldung Dokument erstellt und abgespeichert
            Call objOffice.ShowWord(objDoc, True)                   ' Jetzt doc anzeigen
            Call objOffice.BringWordOnTop                           ' Word in den Vordergrund
        Else
            GoTo WordError
        End If
                
    End If
    frmMain.MousePointer = vbDefault                                ' Sanduhr abschalten
    Call objError.WriteProt(PROT_DOC_END)                           ' Protokolieren
    
exithandler:
On Error Resume Next                                                ' Hier keine Fehlerbehandlung mehr
    If Not CurForm Is Nothing Then
        CurForm.MousePointer = vbDefault                            ' Sanduhr abschalten
    Else
        frmMain.MousePointer = vbDefault                            ' Sanduhr abschalten
    End If
    
    rsRA.Close
    rsAbgesagt.Close
    rsBewerber.Close
    rsMitbewerber.Close
    rsAus.Close
    rsListeBewerber.Close
    rsStellen.Close
    rsUser.Clone
    rsZugesagt.Close
    rsListeBewerberA.Close
    rsListeBewerberZ.Close
    
Exit Function
WordError:
On Error Resume Next
    Call objTools.FileDelete(szDestPath, True)                      ' Evtl. angelegte Datei Löschen
    Call objObjectBag.ShowMSGForm(False, "")                        ' MSG Form ausbelnden
    Call objError.ShowErrMsg("Beim erstellen des Dokuments ist ein Fehler aufgetreten!" & vbCrLf, _
            "Das Dokument wurde nicht gespeichert.", _
            vbInformation + vbOKOnly, "Word")                       ' Meldung Dokument Nicht erstellt oder abgespeichert
    Call objOffice.CloseWordObj                                     ' Word dichtmachen
    GoTo exithandler                                                ' Fertig
Exit Function
Errorhandler:
On Error Resume Next
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objObjectBag.ShowMSGForm(False, "")                       ' MSG Form ausblenden
    Call objError.Errorhandler(MODULNAME, "WriteWord", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Private Function SaveDocRecord(PersID As String, _
        StellenID As String, _
        AusschrID As String, _
        szTemplate As String, _
        szDocTitle As String, _
        szDocPath As String) As Boolean
    Dim szSQL As String                                             ' SQL Statement
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szSQL = "INSERT INTO DOC018 ( "                                 ' INSERT Statement zusammensetzen
    If PersID <> "" Then szSQL = szSQL & "FK010018,"                ' Wenn PersID dann Feldliste erweitern
    If StellenID <> "" Then szSQL = szSQL & "FK012018,"             ' Wenn StellenID dann Feldliste erweitern
    If AusschrID <> "" Then szSQL = szSQL & "FK020018,"             ' Wenn AuschreibungsID dann Feldliste erweitern
    szSQL = szSQL & "DOCPATH018, TEMPLATE018,DOCNAME018,CFROM018 ) Values ("
    If PersID <> "" Then szSQL = szSQL & "'" & PersID & "',"        ' Wenn PersID dann Wertliste erweitern
    If StellenID <> "" Then szSQL = szSQL & "'" & StellenID & "',"  ' Wenn StellenID dann Wertliste erweitern
    If AusschrID <> "" Then szSQL = szSQL & "'" & AusschrID & "',"  ' Wenn AuschreibungsID dann Wertliste erweitern
    szSQL = szSQL & "'" & szDocPath & szDocTitle & "','" & szTemplate & _
                    "','" & szDocTitle & "','" & objObjectBag.GetUserName() & "' )"
    SaveDocRecord = objDBconn.execSql(szSQL)                         'SQL Ausführen und erg. zurück
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
On Error Resume Next
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SaveDocRecord", errNr, errDesc)
    Resume exithandler
End Function

Private Function GetDocListData(XMLPath As String, _
        XMLSATNodeName As String, _
        szParameter As String) As ADODB.Recordset
    Dim szSQL As String                                             ' SQL Statement
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szSQL = objTools.GetDocSQL(XMLPath, XMLSATNodeName)             ' SQL Statemet Empfäger daten aus XML holen
    If szSQL <> "" Then                                             ' SQL Statement Leer -> fertig
        If szParameter <> "" Then                                   ' Wenn Parameter vorhanden
            szSQL = szSQL & " " & Trim(szParameter)                 ' Ans SQL Statement anhängen
        End If
        Set GetDocListData = objDBconn.fillrs(szSQL, False)         ' Daten Holen
        szDetails = XMLSATNodeName & " Daten da"                    ' Details für fehlerbehandlung
    End If
    
exithandler:

Exit Function
Errorhandler:
On Error Resume Next
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetDocListData", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Private Function GetDocData(XMLPath As String, _
        XMLSATNodeName As String, _
        szWhere As String, _
        Optional szOrder As String) As ADODB.Recordset
'Private Function GetDocData(XMLPath As String, _
        XMLSATNodeName As String, _
        szWhere As String) As ADODB.Recordset
    Dim szSQL As String                                             ' SQL Statement
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    szSQL = objTools.GetDocSQL(XMLPath, XMLSATNodeName)             ' SQL Statemet Empfäger daten aus XML holen
    If szSQL <> "" Then                                             ' SQL Vorhanden
        If szWhere <> "" Then                                       ' Where Part vorhanden
            szSQL = objSQLTools.AddWhereInFullSQL(szSQL, szWhere)   ' Where bed. anhängen
' MW 15.09.2011 {
            If szOrder <> "" Then
                szSQL = szSQL & " ORDER BY " & szOrder
            End If
' MW 15.09.2011 }
        End If
        If InStr(UCase(szSQL), UCase("@Jahr")) > 0 Then
            
        End If
        Set GetDocData = objDBconn.fillrs(szSQL, False)             ' Daten Holen
        szDetails = XMLSATNodeName & " Daten da"                    ' Details für Fehlerbehandlung
    End If
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
On Error Resume Next
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetDocData", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Private Function CheckRequieredIDForDoc(objDoc As Object, _
            bPersIDRequired As Boolean, _
            PersID As String, _
            bStellenIDRequired As Boolean, _
            StellenID As String, _
            bAussrIDRequired As Boolean, _
            AussrID As String, _
            Optional szResult As String) As Boolean
' Überprüft ob bestimmte DocVariaablen in der Vorlage benötigtwerden
' Und stellet sicher das die benötigten IDs für diese Abfragen vorhanden sind
    Dim bAussrIDOK As Boolean                                       ' Vorraussetzung Auschreibung ID  erfüllt
    Dim bPersIDOK As Boolean                                        ' Vorraussetzung Personen ID erfüllt
    Dim bStellenIDOK As Boolean                                     ' Vorraussetzung Stellen ID erfüllt
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    'Call ZeitMessungStart                                          ' Zeitmessung zum Debuggen
                                                                    ' Prüfen welche DocVariaben in der Vorlage vorkommen
    szResult = objOffice.DocHasVarFields(objDoc, "RA;Bewerber;Stellen;Mitbewerber;AMitbewerber;ZMitbewerber;Ausschreibung", True)
    If InStr(szResult, "RA;") > 0 Then bPersIDRequired = True       ' Benötigt Pers ID
    If InStr(szResult, "Bewerber;") > 0 Then
        bPersIDRequired = True                                      ' Benötigt Stellen ID
        bStellenIDRequired = True                                   ' + Pers ID
    End If
    If InStr(szResult, "Ausschreibung") > 0 Then bAussrIDRequired = True ' Benötigt Ausschreibungs ID
    If InStr(szResult, "Stellen;") > 0 Then bStellenIDRequired = True ' Benötigt Stellen ID
    If InStr(szResult, "AMitbewerber") > 0 Then bStellenIDRequired = True ' Benötigt Stellen ID
    If InStr(szResult, "ZMitbewerber") > 0 Then bStellenIDRequired = True ' Benötigt Stellen ID
    If InStr(szResult, "Mitbewerber") > 0 Then bStellenIDRequired = True ' Benötigt Stellen ID
    'Call ZeitMessungEnde("CheckRequieredIDForDoc (Neuer weg)")     ' Zeitmessung zum Debuggen
    szDetails = "PersID: " & PersID & vbCrLf & "StellenID: " & StellenID
    If StellenID = "" And bStellenIDRequired Then                   ' Wenn StellenID leer und benötigt
        StellenID = ShowSearch(objDBconn, "Stellen", "Bezirk")      ' Suche nach stellen auffrufen
        If bPersIDRequired Then                          ' Wenn PersID Leer und benötigt
            If StellenID <> "" Then                                     ' Wenn Stellen ID vorhanden
                PersID = ShowSearch(objDBconn, "PersonenNachStellen", "Nachname", "", StellenID) ' Person in dieser Stelle suchen
                szDetails = "PersID: " & PersID & vbCrLf & "StellenID: " & StellenID
            Else
                PersID = ShowSearch(objDBconn, "Personen", "Nachname")  ' Sonst Person Suchen
                szDetails = "PersID: " & PersID & vbCrLf & "StellenID: " & StellenID
            End If
        End If
        szDetails = "PersID: " & PersID & vbCrLf & "StellenID: " & StellenID
    End If
    If PersID = "" And bPersIDRequired Then                         ' Wenn PersID Leer und benötigt
        If StellenID <> "" Then                                     ' Wenn Stellen ID vorhanden
            PersID = ShowSearch(objDBconn, "PersonenNachStellen", "Nachname", "", StellenID) ' Person in dieser Stelle suchen
            szDetails = "PersID: " & PersID & vbCrLf & "StellenID: " & StellenID
        Else
            PersID = ShowSearch(objDBconn, "Personen", "Nachname")  ' Sonst Person Suchen
            szDetails = "PersID: " & PersID & vbCrLf & "StellenID: " & StellenID
        End If
    End If
    If (bPersIDRequired And PersID <> "") _
        Or Not bPersIDRequired Then bPersIDOK = True
    If (bStellenIDRequired And StellenID <> "") _
        Or Not bStellenIDRequired Then bStellenIDOK = True
    If (bAussrIDRequired And AussrID <> "") _
        Or Not bAussrIDRequired Then bAussrIDOK = True
    szDetails = "PersID: " & PersID & vbCrLf & "StellenID: " & StellenID & "AusschreibungsID: " & AussrID
    If bStellenIDOK And bPersIDOK And bAussrIDOK Then CheckRequieredIDForDoc = True
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
On Error Resume Next
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "CheckRequieredIDForDoc", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Private Function CheckRequieredListForDoc(objDoc As Object, _
            szRequLists As String) As Boolean
' Überprüft ob bestimmte DocVariaablen für Listen in der Vorlage benötigt werden
    Dim bPersIDOK As Boolean                                        ' Vorraussetzung Personen ID erfüllt
    Dim bStellenIDOK As Boolean                                     ' Vorraussetzung Stellen ID erfüllt
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    Dim szResult As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
                                                                        
    'Call ZeitMessungStart                                          ' Zeitmessung zum debuggen
                                                                    ' Prüfen welche DocVariaben in der Vorlage vorkommen
    szResult = objOffice.DocHasVarFields(objDoc, "ListeBewerberZugesagt_RangNamePunkte;ListeBewerberAbgelehnt_RangNamePunkte;" & _
            "ListeAusgeschreibeneStellen;ListeBewerber_RangNamePunkte;ListeBewerber_RANachname;ListeBewerberAbgelehnt_RANachname;" & _
            "ListeBewerberZugesagt_RANachname", True)
    szRequLists = szResult
    'Call ZeitMessungEnde("CheckRequieredListForDoc ")              ' Zeitmessung zum debuggen
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
On Error Resume Next
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "CheckRequieredListForDoc", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Private Function SelectImport(ThisDBCon As Object, _
        szDefFolder As String, _
        szImpFile As String, szImpFolder As String, _
        Optional PersID As String, _
        Optional StellenID As String, _
        Optional AusschrID As String, _
        Optional bFolder As Boolean) As Boolean
    Dim bCancel As Boolean                                          ' Abbruch variable
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If szDefFolder = "" Then
        szDefFolder = objObjectBag.GetPersonalDir                   ' Evtl. Eigene Dateien holen
    End If
    frmDocImport.txtFile = szImpFile
    frmDocImport.txtIDPerson = PersID
    frmDocImport.txtIDStelle = StellenID
    frmDocImport.txtIDAusschreibung = AusschrID
    frmDocImport.szDefFolder = szDefFolder                          ' Standart Suchordner übergeben
    Call frmDocImport.InitForm(objObjectBag.getMainForm, ThisDBCon, , , bFolder) ' Auswahl Form  initialisieren
    frmDocImport.Show vbModal, objObjectBag.getMainForm             ' Auswahl Form für Import file anzeigen
On Error Resume Next                                                ' Hier Fehlerbehandlung sicherheitshalber deaktivieren
    PersID = frmDocImport.txtIDPerson
    StellenID = frmDocImport.txtIDStelle
    AusschrID = frmDocImport.txtIDAusschreibung
    If bFolder Then                                                 ' Wenn Ordner import
        szImpFolder = frmDocImport.txtFile                          ' Ordnername zurück
        If Not objTools.CheckLastChar(szImpFolder, "\") Then        ' evtl. Letzten \ anhängen
            szImpFolder = szImpFolder & "\"
        End If
    Else                                                            ' Sonst
        szImpFile = frmDocImport.txtFile                            ' dateiname zurück
    End If
    Err.Clear
On Error GoTo Errorhandler                                          ' fehlerbehandlung wieder aktivieren
            
    If frmDocImport.bCancel Then bCancel = True                     ' Abbruch?
    
    frmDocImport.bCancel = False
    Unload frmDocImport                                             ' Auswahl form schliessen
        
    If bCancel Then GoTo exithandler                                ' keine Import dann raus
    SelectImport = True                                             ' Kann weitergehen
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
On Error Resume Next
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SelectImport", errNr, errDesc)
    Resume exithandler
End Function

Private Function SelectDocTemplate() As String
' Zeigt einen Auswahl dialog für eine Word vorlage an
    Dim szTemplate As String                                        ' Name der Vorlage
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
        frmVorlagenSelect.Show vbModal, objObjectBag.getMainForm    ' Auswahl Form für Vorlagen anzeigen
On Error Resume Next                                                ' Fehlerbehandlung erstmal deaktivieren
        szTemplate = frmVorlagenSelect.LVVorlagen.SelectedItem      ' Ausgewählte vorlage übernehmen
        Err.Clear
On Error GoTo Errorhandler                                          ' Fehlerbehandlung Wieder aktivieren
        If frmVorlagenSelect.bCancel Then szTemplate = ""           ' Abbruch?
        frmVorlagenSelect.bCancel = False
        Unload frmVorlagenSelect                                    ' Auswahl form schliessen
        If szTemplate = "" Then GoTo exithandler                    ' keine Vorlage dann raus
        SelectDocTemplate = szTemplate                              ' Ausgesuchte vorlage zurück geben
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
On Error Resume Next
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SelectDocTemplate", errNr, errDesc)
    Resume exithandler
End Function

Private Function GetDocTitleFromIDs(ThisDBCon As Object, _
        Optional PersID As String, _
        Optional StellenID As String, _
        Optional AusschrID As String, _
        Optional bWithDate As Boolean) As String
' Ermittelt einen Eindeutigen Dokumenten namen aus AZ, PerdID, StellenID, und AusschreibungsID

    Dim szAZ As String                                              ' AZ (Vorgang oder Notar)
    Dim szDocTitle As String                                        ' Dokumenten Titel
    Dim szBezirk As String                                          ' Bezirk Der Stelle oder des Notars
    Dim szNachname As String                                        ' Nachname des Bewerbers oder des Notars
    Dim szDateTime As String                                        ' Datum und Uhrzeit als String
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

        szAZ = GetAZfromIds(objDBconn, PersID, StellenID, AusschrID) ' Aktenzeichen Ermitteln
        If StellenID <> "" Then
            szBezirk = objTools.checknull(ThisDBCon.GetValueFromSQL( _
                "SELECT TOP 1 BEZIRK012 FROM STELLEN012 WHERE ID012='" _
                & StellenID & "'"), " ")                            ' Bezirk der Stelle
        End If
        If PersID <> "" Then
            szNachname = objTools.checknull(ThisDBCon.GetValueFromSQL( _
                "SELECT TOP 1 NACHNAME010 FROM RA010 WHERE ID010='" _
                & PersID & "'"), " ")                               ' Nachname der Person
        End If
        
        szDocTitle = szAZ                                           ' Neuer Doc name setzt sich zusammen aus AZ
        If szBezirk <> "" Then szDocTitle = szDocTitle & " " & szBezirk ' Bezirk
        If szNachname <> "" Then szDocTitle = szDocTitle & " " & szNachname ' Nachname
        If bWithDate Then
            szDateTime = CStr(Now())
            szDateTime = Replace(szDateTime, ".", "")
            szDateTime = Replace(szDateTime, ":", "")
            szDateTime = Replace(szDateTime, " ", "")
            szDocTitle = szDocTitle & " " & szDateTime              ' Evtl. Datum anhängen
        End If
        GetDocTitleFromIDs = Trim(szDocTitle)
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetDocTitleFromIDs", errNr, errDesc)
    Resume exithandler
End Function

Private Function GetBezirkFromIDs(ThisDBCon As Object, _
        Optional PersID As String, _
        Optional StellenID As String, _
        Optional AusschrID As String) As String
' ermittelt Bezirk aus mehr oder minder vollständigem ID Convult
    Dim szBezirk As String                                          ' Ermittelter Bezirk
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    'If AusschrID <> "" Then                                         ' Wenn Auschreibung angegeben
'        szAZ = ThisDBcon.GetValueFromSQL("SELECT TOP 1 AZ020 FROM AUSSCHREIBUNG020 " & _
                " WHERE ID020='" & AusschrID & "'")
            
    'Else
    If StellenID <> "" Then                                         ' Wenn Stelle angegeben
        szBezirk = ThisDBCon.GetValueFromSQL("SELECT TOP 1 BEZIRK012 FROM STELLEN012 " & _
                " WHERE ID012='" & StellenID & "'")
    ElseIf PersID <> "" Then                                        ' Wenn Person angegeben
        szBezirk = ThisDBCon.GetValueFromSQL("SELECT TOP 1 AG010 FROM RA010 " & _
                " WHERE ID010='" & PersID & "'")                    ' Notar AG Bezirk
        If Trim(szBezirk) = "" Then
            szBezirk = ThisDBCon.GetValueFromSQL("SELECT TOP 1 BEZIRK012 FROM  " & _
                    " Stellen012 " & _
                    " INNER JOIN Bewerb013 On FK012013 = FK010013 " & _
                    " WHERE FK010013='" & PersID & "'")
        End If
    Else                                                            ' Nix
        
    End If
    GetBezirkFromIDs = szBezirk                                     ' Bezirk zurück geben
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
    Call objError.Errorhandler(MODULNAME, "GetBezirkFromIDs", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Private Function GetAZfromIds(ThisDBCon As Object, _
        Optional PersID As String, _
        Optional StellenID As String, _
        Optional AusschrID As String) As String
' ermittelt AZ aus mehr oder minder vollständigem ID Convult
    Dim szAZ As String                                              ' Ermitteltes Aktenzeichen
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If AusschrID <> "" Then                                         ' Wenn Auschreibung angegeben
        szAZ = ThisDBCon.GetValueFromSQL("SELECT TOP 1 AZ020 FROM AUSSCHREIBUNG020 " & _
                " WHERE ID020='" & AusschrID & "'")
    ElseIf StellenID <> "" Then                                     ' Wenn Stelle angegeben
        szAZ = ThisDBCon.GetValueFromSQL("SELECT TOP 1 AZ020 FROM AUSSCHREIBUNG020 " & _
                " INNER JOIN STELLEN012 ON ID020 = FK020012 " & _
                " WHERE ID012='" & StellenID & "'")
    ElseIf PersID <> "" Then                                        ' Wenn Person angegeben
        szAZ = ThisDBCon.GetValueFromSQL("SELECT TOP 1 AZ010 FROM RA010 " & _
                " WHERE ID010='" & PersID & "'")                    ' Notar AZ VI
        If Trim(szAZ) = "" Then
            szAZ = ThisDBCon.GetValueFromSQL("SELECT TOP 1 AZ020 FROM (AUSSCHREIBUNG020 " & _
                    " INNER JOIN Stellen012 ON ID020 = FK020012) " & _
                    " INNER JOIN Bewerb013 On FK012013 = FK010013 " & _
                    " WHERE FK010013='" & PersID & "'")
        End If
    Else                                                            ' Nix
        
    End If
    GetAZfromIds = szAZ                                             ' Aktenzeichen zurück geben
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
    Call objError.Errorhandler(MODULNAME, "GetAZfromIDs", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function PrepareFolderPath(FolderPath As String) As String
' Prepariert einen FolderPath so das Punkte, Doppelpunkte und Leerzeichen
' durch Unterstriche ersetzt werden
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If FolderPath = "" Then GoTo exithandler                        ' Kein FolderPath -> fertig
    FolderPath = Trim(FolderPath)                                   ' FolderPath Trimmen
    FolderPath = Replace(FolderPath, " ", "_")                      ' Leerzeichen aus DocTitle entfernen
    FolderPath = Replace(FolderPath, ".", "_")                      ' . aus DocTitle entfernen
    FolderPath = Replace(FolderPath, ":", "_")                      ' : aus DocTitle entfernen
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    PrepareFolderPath = FolderPath                                  ' Neuen Pfad zurück
    Err.Clear                                                       ' Evtl. error clearen
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "PrepareFolderPath", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function PrepareDocTitle(Filename As String, Optional FileSuffix As String) As String
' Prepariert einen Filenamen so das Punkte, Doppelpunkte und Leerzeichen
' durch Unterstriche ersetzt werden
    Dim bSaveAsDocx As Boolean                                      ' Soll als Docx (word 2007) gespeichert werden?
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If Filename = "" Then GoTo exithandler                          ' Kein Filename -> fertig
    Filename = Trim(Filename)                                       ' Filename Trimmen
    Filename = Replace(Filename, " ", "_")                          ' Leerzeichen aus DocTitle entfernen
    Filename = Replace(Filename, ".", "_")                          ' . aus DocTitle entfernen
    Filename = Replace(Filename, ":", "_")                          ' : aus DocTitle entfernen
    If FileSuffix = "" Then                                         ' Wenn Filesuffix nicht vorhanden
        If Left(Right(Filename, 4), 1) = "." Then                   ' Filesuffix (word 2003 & sonst) schon im Dateinamen
            FileSuffix = Right(Filename, 4)                         ' Suffix ermitteln
            Filename = Left(Filename, Len(Filename) - 4)            ' Dateinamen ohne suffix
        ElseIf Left(Right(Filename, 5), 1) = "." Then               ' Filesuffix (word 3007) schon im Dateinamen
            FileSuffix = Right(Filename, 5)                         ' Suffix ermitteln
            Filename = Left(Filename, Len(Filename) - 5)            ' Dateinamen ohne suffix
        Else                                                        ' Sonst
            FileSuffix = ".___"                                     ' Defaultsuffix
        End If
    Else
        FileSuffix = Trim(FileSuffix)                               ' Filesuffix Trimmen
    End If
' MW 01.09.11 {
    If Left(FileSuffix, 1) <> "." Then
        FileSuffix = "." & FileSuffix
    End If
'    If Len(FileSuffix) > 4 Then                                     ' Wenn Filesuffix Fehlerhaft
'        FileSuffix = Left(FileSuffix, 3)
'        FileSuffix = Replace(FileSuffix, ".", "_")
'        FileSuffix = "." & FileSuffix
'    ElseIf Len(FileSuffix) < 3 Then
'        FileSuffix = ".___"                                         ' Defaultsuffix setzen
'    End If
'    If Len(FileSuffix) = 3 Then
'        FileSuffix = Replace(FileSuffix, ".", "_")
'        FileSuffix = "." & FileSuffix
'    End If
'    If Len(FileSuffix) = 4 Then
'        If Left(FileSuffix, 1) <> "." Then
'            FileSuffix = Left(FileSuffix, 3)
'            FileSuffix = Replace(FileSuffix, ".", "_")
'            FileSuffix = "." & FileSuffix
'        End If
'    End If
' MW 01.09.11 }
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    PrepareDocTitle = Filename & FileSuffix                         ' Documenten namen zurück geben
    Err.Clear                                                       ' Evtl. error clearen
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "PrepareDocTitle", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function CheckTodayDeadlins(dbCon As Object) As Boolean
' Die funktion prüft ob Heute Fristen ablaufen , falls ja wird Ture zurück geliefert
    Dim szSQL  As String                                            ' SQL Statement
    Dim Result As Integer                                           ' Ergebnis (anzahl heute fälliger fristen)
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szSQL = "SELECT Count(*) FROM Frist024 " & _
        "WHERE Frist024.Frist024 = DateAdd(dd, 0, DateDiff(dd, 0, GETDATE()))" ' SQL Statement festlegen
    Result = dbCon.GetValueFromSQL(szSQL)                           ' Abfrage ausführen
    If Result > 0 Then                                              ' Mehr als null fristen?
        CheckTodayDeadlins = True                                   ' True zurück
    Else                                                            ' Sonst
        CheckTodayDeadlins = False                                  ' eben false
    End If
exithandler:
Exit Function                                                       ' Funktion beenden
Errorhandler:
Dim errNr As String                                                 ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "CheckTodayDeadlins", errNr, errDesc)  ' Fehler behandlung aufrufen
    CheckTodayDeadlins = False                                      ' Bei Fehler auch keine Fristen gefunden
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function SetLVDataInWorkSheet(LV As ListView, _
        Optional bWithoutColumnHeader As Boolean, _
        Optional bPrintNow As Boolean, _
        Optional bQuerformat As Boolean) As Boolean
' Schaufelt die Daten eines Listviews in eine neue excelmappe
    Dim objWorkBook As Object                                       ' Arbeitsmappen object
    Dim objWorkSheet As Object                                      ' Arbeitsblatt object
    Dim i As Integer                                                ' Counter LV items
    Dim n As Integer                                                ' Counter LV Subitems
    Dim LVItem As ListItem                                          ' Akt ListView Item
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    Dim szSpaltenBezug As String
    Dim lngZeilenoffset As Integer
    Dim szTitelRange As String                                      ' Wiederholter Titel
    Dim szPrintRange As String                                      ' Druckbereich
    Dim szTVNodeKey                                                 ' NodeKey als Vorlage für SheetNamen
    Dim NameArray() As String                                       ' Aufgespaltener NodeKey
    Dim szSheetName As String
    Dim Step As Integer                                             ' Schrittweite Für Progess
    Dim MaxSteps As Integer                                         ' Max. Schritte für Progressbar
On Error Resume Next
    szTVNodeKey = frmMain.TVMain.SelectedItem.FullPath
    Err.Clear
On Error GoTo Errorhandler

    If LV.ListItems.Count = 0 Then
        MaxSteps = 100
        Step = 10
    Else
        MaxSteps = objObjectBag.GetMaxProgresSteps(LV.ListItems.Count)
        Step = objObjectBag.GetProgresStepsWidth(LV.ListItems.Count, MaxSteps)
    End If
    
    szDetails = "Starte Excelexport"
    Call objError.WriteProt("Starte Excelexport")                   ' ExcelExport Start Protokolieren
    
    frmMain.MousePointer = vbHourglass                              ' Sanduhranzeigen
    
'    Set mainform = objObjectBag.GetMainForm()
'    NewName = DiscribeTreeNode(mainform.TVMain, "", True)
    If szTVNodeKey <> "" Then
        NameArray = Split(szTVNodeKey, TV_KEY_SEP)
        For i = 0 To UBound(NameArray)
            szSheetName = szSheetName & " > " & NameArray(i)
        Next i
    End If
    
    Call objObjectBag.ShowMSGForm(True, "Export nach Excel", MaxSteps, "Lade Excelmappe") ' MSG Form anzeigen mit ProgBar
    
    Set objWorkBook = objOffice.OpenNewExcelWorkbook("", True)      ' Arbeitsmappe unsichtbar öffnen
    
    Set objWorkSheet = objOffice.GetExcelWorkSheet(objWorkBook, 1)  ' Arbeitsblatt öffnen
    
    If Not bWithoutColumnHeader Then                                ' 1. Zeile Spalten überschriften
        Call objObjectBag.ShowMsgNextStep("übertrage Spaltenüberschriften", Step)  ' Progbar
        lngZeilenoffset = 1
        szSpaltenBezug = Chr(65)                                    ' Spalte A festlegen
        DoEvents                                                    ' Andere Aktionen zulassen
            For i = 1 To LV.ColumnHeaders.Count
                If LV.ColumnHeaders(i).Width > 0 Then
                    Call objOffice.WriteExcelCell(objWorkSheet, szSpaltenBezug & lngZeilenoffset, LV.ColumnHeaders(i).Text)    ' Schreiben
                    Call objObjectBag.ShowMsgNextStep("übertrage Spaltenüberschriften", Step)  ' Progbar
                End If
                If i = 1 Then szTitelRange = szSpaltenBezug & lngZeilenoffset  ' Start TitleRange merken
                If i = LV.ColumnHeaders.Count Then szTitelRange = szTitelRange & ":" & szSpaltenBezug & lngZeilenoffset ' Ende TitleRange merken
                If 65 + n < 90 Then
                    szSpaltenBezug = Chr(65 + i)
                Else
                    szSpaltenBezug = "A" & Chr(65 + i - 25)
                End If
            Next i
    Else
        lngZeilenoffset = 0
    End If
    
    Call objObjectBag.ShowMsgNextStep("übertrage Daten", Step)      ' Progbar
    
    For i = 1 To LV.ListItems.Count                                 ' Alle Listitems Durchlaufen
        szSpaltenBezug = Chr(65)                                    ' Spalte A festlegen
        Set LVItem = LV.ListItems(i)
        Call objOffice.WriteExcelCell(objWorkSheet, szSpaltenBezug & i + lngZeilenoffset, LVItem.Text)   ' Schreiben
        If i = 1 Then szPrintRange = szSpaltenBezug & i + lngZeilenoffset  ' Start DruckBereich merken
            
        DoEvents                                                    ' Andere Aktionen zulassen
        For n = 1 To LVItem.ListSubItems.Count
            If 65 + n < 90 Then
                szSpaltenBezug = Chr(65 + n)
            Else
                szSpaltenBezug = "A" & Chr(65 + n - 25)
            End If
            If LV.ColumnHeaders(n + 1).Width > 0 Then
                Call objObjectBag.ShowMsgNextStep("übertrage Daten", Step) ' Progbar
                Call objOffice.WriteExcelCell(objWorkSheet, szSpaltenBezug & i + lngZeilenoffset, LVItem.SubItems(n))
            End If
            If i = LV.ListItems.Count And n = LVItem.ListSubItems.Count Then
                szPrintRange = szPrintRange & ":" & szSpaltenBezug & i + lngZeilenoffset ' Ende DruckBereich merken
            End If
        Next n
    Next
    
    'Public Function SetExcelPrintSetup(objWorkSheet As Object, szPrintArea As String, szPrintTitleRows As String, _
    Optional bTitleBold As Boolean, Optional bPagenubers As Boolean, Optional szHeaderText As String, _
    Optional bPrintNow As Boolean, Optional bPrintPreview As Boolean) As Boolean
    Call objObjectBag.ShowMsgNextStep("lege Druckbereich fest", Step) ' Progbar 25 %
    
    If bPrintNow Then Call objObjectBag.ShowMSGForm(False, "")      ' MSG Form ausblenden
    
    If objOffice.SetExcelPrintSetup(objWorkSheet, szPrintRange, szTitelRange, _
            True, True, szSheetName, , bPrintNow, bQuerformat) Then             ' Druckbereich festlegen
        If Not bPrintNow Then
            
            Call objOffice.showexcel(True)                          ' Anzeigen
            szDetails = "Excelexport Abgeschlossen"
            Call objError.WriteProt("Excelexport Abgeschlossen")    ' Protokolieren
        Else
            objWorkBook.Saved = True
            Call objOffice.CloseExcelObj
            szDetails = "Excelexport Abgeschlossen"
            Call objError.WriteProt("Excelexport Abgeschlossen und Ausgedruckt") ' Protokolieren
        End If
    Else
        Call objOffice.showexcel(True)                              ' Anzeigen
        szDetails = "Excelexport Abgeschlossen"
        Call objError.WriteProt("Excelexport Fehler ")              ' Protokolieren
    End If
    
    
exithandler:
On Error Resume Next
    frmMain.MousePointer = vbDefault                                ' Sanduhr abschalten
    Call objObjectBag.ShowMSGForm(False, "")                        ' MSG Form ausblenden
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objObjectBag.ShowMSGForm(False, "")                        ' MSG Form ausblenden
    Call objError.Errorhandler(MODULNAME, "SetLVDataInWorkSheet", errNr, errDesc)
    SetLVDataInWorkSheet = False
    Resume exithandler
End Function

Public Sub ZeitMessungStart()
On Error Resume Next
    StartTick = SignedToUnsignedLong(GetTickCount())            ' Zeitmessung für Performance
    Err.Clear
End Sub

Public Sub ZeitMessungEnde(szGemessen As String)
    Dim EndTick As Double
    Dim TickDiff As Double
On Error Resume Next
    EndTick = SignedToUnsignedLong(GetTickCount())              ' Zeitmessung für Performance
    TickDiff = (EndTick - StartTick) / 1000
    Debug.Print szGemessen & " dauert " & TickDiff & " sec."
    StartTick = 0
    Err.Clear
End Sub


