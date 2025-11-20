VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObjectBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit                                                     ' Variaben Deklaration erzwingen
Option Compare Text                                                 ' Sortierreihenfolge festlegen
Private Const MODULNAME = "clsObjectBag"                            ' Modulname für Fehlerbehandlung

Private bDebug As Boolean                                           ' Hilfsvar für Debugging

Private objError  As Object                                         ' Glob Error object
Private objDBconn As Object                                         ' DB Connection
Private objReg As Object                                            ' Registry Tools
Private objTools As Object                                          ' Hilfreiches
Private objSQLTools As Object                                       ' Tools Für SQL strings
Private objOptions As Object                                        ' Optionen Speichern und Lesen
Private objOfficeTools As Object                                    ' Word Schnittstelle

Private frmMainWindow As Object                                     ' Referenz aufs Hauptform
Private fSplash As frmSplash                                        ' Referenz auf denSplash Screen

Private szIniFile As String                                         ' Name der IniDatei
Private szXmlFile As String                                         ' Name der XML inidatei
Private szUsername As String                                        ' der Angemeldete Benutzer
Private szMSInfoExe As String                                       ' Pfad zur MSInfo
Private szMStscExe As String                                        ' Pfad zum Remotedesktop
Private szNotepadExe As String                                      ' Pfad zur Notepad.exe
Private szReadmePath As String                                      ' Pfad zur Readme Datei
Private bAdmin As Boolean                                           ' Akt. User ist Administrator
Private DeskWidth As Integer                                        ' Dektop breite
Private DeskHeight As Integer                                       ' Desktop Höhe
Private CompInfo As ComputerInfo                                    ' Akt. Computer informationen
Private OfficeVersionen As OfficeInfo                               ' Akt Office Versionen
Private ApplicationInfo As AppInfo                                  ' Infos dieser Anwendung
Private bE As Boolean
Private bConsole As Boolean                                         ' Consolen Modus
Private bExpert As Boolean                                          ' Experten Modus
Private bLeet As Boolean
Private bNotShowSplash As Boolean                                   ' Kein Splash anzeigen
Private bMorse As Boolean
Private Att As String
'Private AuthLogin As Boolean
Private HTML As Boolean
'Private MyEncodeType As ENCODE_METHODe

Private Sub Class_Initialize()
    Dim szDetails As String                                         ' fehler details
On Error GoTo Errorhandler
'    bDebug = True                                                   ' Vor release auf false setzen
    If InitErrorHandler Then                                        ' Fehlerbehandlung initialisieren
        'Call objError.InitObjectBag(Me)  ' OB ref. weitergeben
    Else
       GoTo InitError                                               ' Hier eine abbruch behandlung
    End If
    If InitRegTools Then                                            ' Registry Tools Initialisieren
        'Call objReg.InitObjectBag(Me)                              ' OB ref. weitergeben
    Else
        GoTo InitError                                              ' Hier eine abbruch behandlung
    End If
    If InitTools() Then
        'Call objTools.InitObjectBag(Me)                            ' OB ref. weitergeben
    Else
        GoTo InitError                                              ' Hier eine abbruch behandlung
    End If
    bDebug = objTools.IsIde()                                       ' Wenn in Ide Ausgeführt dann debug true
    Call GetSysteminfo                                              ' System infos abfragen
exithandler:
Exit Sub
InitError:
    szDetails = "Fehler bei der initialisierung der Anwendung." & vbCrLf _
            & "Eine Neuinstallation könnte das Problem beheben."
    Call OBErrorhandler(MODULNAME, "Class_Initialize", 91, "", szDetails)
Exit Sub
Errorhandler:
On Error Resume Next
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
On Error Resume Next                                                ' Kein Fehler in der Fehlerbehandlung
    Call OBErrorhandler(MODULNAME, "Class_Initialize", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub Class_Terminate()
On Error Resume Next                                                ' Hier keine Fehler mehr
' Alle aufgemachten objecte schliessen
    objDBconn.Close
    objReg.Close
    objTools.Close
    objError.Close
    objSQLTools.Close
    objOptions.Close
    objOfficeTools.Close
    
    Set objOfficeTools = Nothing
    Set objDBconn = Nothing
    Set objReg = Nothing
    Set objTools = Nothing
    Set objSQLTools = Nothing
    Set objError = Nothing
    Set objOptions = Nothing
    
exithandler:
Exit Sub
Errorhandler:
On Error Resume Next
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
On Error Resume Next                                                ' Kein Fehler in der Fehlerbehandlung
    Call OBErrorhandler(MODULNAME, "Class_Terminate", errNr, errDesc)
    Resume exithandler
End Sub
                                                                    ' *****************************************
                                                                    ' Properties
Public Property Get GetMainForm() As Object
    Set GetMainForm = frmMainWindow                                 ' liefert referenz aufs Main (MDI) form
End Property

Public Property Let SetMainForm(ByVal fMain As Object)
       Set frmMainWindow = fMain
       objError.SetMainForm = fMain                                 ' Gleich an Fehlerbehandlung weiterreichen
End Property

Public Property Get GetOfficeObj() As Object
    Set GetOfficeObj = objOfficeTools                               ' Office Schnittstelle
End Property

Public Property Get GetErrorObj() As Object
    Set GetErrorObj = objError                                      ' Error Object
End Property

Public Property Get GetDBConObj() As Object
    Set GetDBConObj = objDBconn                                     ' DB Verbindungsobject
End Property

Public Property Get GetSQLToolsObj() As Object
    Set GetSQLToolsObj = objSQLTools                                ' Object mit SQL Hilfsfunktion
End Property

Public Property Get GetRegToolsObj() As Object
    Set GetRegToolsObj = objReg                                     ' Object mit Reg Funktionen
End Property

Public Property Get GetToolsObj() As Object
    Set GetToolsObj = objTools                                      ' Object mit Allg hilfsfunktionen
End Property

Public Property Get GetOptionsObj() As Object
    Set GetOptionsObj = objOptions                                  ' Options Object
End Property

Public Property Let SetAppDir(ByVal vNewValue As String)
    ApplicationInfo.AppFolder = vNewValue
    Call InitAppDir                                                 ' Anwendungsverz. Initialisieren
End Property

Public Property Get GetAppDir() As String
    GetAppDir = ApplicationInfo.AppFolder
End Property

Public Property Let SetSQLDir(ByVal vNewValue As String)
    ApplicationInfo.AppSQLFolder = vNewValue
End Property

Public Property Get GetSQLDir() As String
    GetSQLDir = ApplicationInfo.AppSQLFolder
End Property

Public Property Get GetPersonalDir() As String
    GetPersonalDir = WinFolder.UserEigeneDateien                    ' Eigene Dateien
    If Right(GetPersonalDir, 1) <> "\" Then
        GetPersonalDir = GetPersonalDir & "\"
    End If
End Property

Public Property Let SetAppTitle(ByVal vNewValue As String)
    ApplicationInfo.AppTitel = vNewValue                            ' AnwendungsTitel Initialisieren
End Property

Public Property Get GetAppTitle() As String
    GetAppTitle = ApplicationInfo.AppTitel
End Property

Public Property Let SetAppRegRoot(ByVal vNewValue As String)
    ApplicationInfo.AppRegRoot = vNewValue                          ' Basis Regschlüssel Initialisieren
End Property

Public Property Get GetAppRegRoot() As String
    GetAppRegRoot = ApplicationInfo.AppRegRoot
End Property

Public Property Let SetCopyright(ByVal vNewValue As String)
    ApplicationInfo.AppCopyright = vNewValue                        ' Copyright String Initialisieren
End Property

Public Property Get GetCopyright() As String
    GetCopyright = ApplicationInfo.AppCopyright
End Property

Public Property Let SetMajor(ByVal vNewValue As String)
    With ApplicationInfo
        .AppVersionMajor = vNewValue                                ' Major Version. Initialisieren
        .AppVersion = .AppVersionMajor & "." & .AppVersionMinor & "." & .AppVersionRev
    End With
End Property

Public Property Get GetMajor() As String
    GetMajor = ApplicationInfo.AppVersionMajor
End Property

Public Property Let SetMinor(ByVal vNewValue As String)
    With ApplicationInfo
        .AppVersionMinor = vNewValue                                ' Minor Version. Initialisieren
        .AppVersion = .AppVersionMajor & "." & .AppVersionMinor & "." & .AppVersionRev
    End With
End Property

Public Property Get GetMinor() As String
    GetMinor = ApplicationInfo.AppVersionMinor
End Property

Public Property Let SetRevision(ByVal vNewValue As String)
    With ApplicationInfo
        .AppVersionRev = vNewValue                                  ' Revision Version. Initialisieren
        .AppVersion = .AppVersionMajor & "." & .AppVersionMinor & "." & .AppVersionRev
    End With
End Property

Public Property Get GetRevision() As String
    GetRevision = ApplicationInfo.AppVersionRev
End Property

Public Property Get GetAppVersion() As String
    GetAppVersion = ApplicationInfo.AppVersion
End Property

Public Property Get GetWordVersion() As String
    GetWordVersion = OfficeVersionen.WordVersion                    ' Word Version
End Property

Public Property Get GetWordPath() As String
    GetWordPath = OfficeVersionen.WordPath                          ' Pfad zur Word exe
End Property

Public Property Get GetExcelVersion() As String
    GetExcelVersion = OfficeVersionen.ExcelVersion                  ' Excel Version
End Property

Public Property Get GetExcelPath() As String
    GetExcelPath = OfficeVersionen.ExcelPath                        ' Pfad zur Excel.exe
End Property

Public Property Get GetAccessVersion() As String
    GetAccessVersion = OfficeVersionen.AccessVersion                ' Access Version
End Property

Public Property Get GetAccessPath() As String
    GetAccessPath = OfficeVersionen.AccessPath                      ' Access zur MSAccess exe
End Property

Public Property Let SetSupportMail(ByVal vNewValue As String)
    ApplicationInfo.AppSuportMail = vNewValue                       ' Support Emailadr. Initialisieren
End Property

Public Property Get GetSupportMail() As String
    GetSupportMail = ApplicationInfo.AppSuportMail
End Property

Public Property Let SetWWW(ByVal vNewValue As String)
    ApplicationInfo.AppWWW = vNewValue                              ' INetadr. Initialisieren
End Property

Public Property Get GetWWW() As String
    GetWWW = ApplicationInfo.AppWWW
End Property

Public Property Let SetReadMePath(ByVal vNewValue As String)
    szReadmePath = vNewValue
    If InStr(szReadmePath, "\") = 0 Then
        szReadmePath = ApplicationInfo.AppFolder & szReadmePath
    End If
End Property

Public Property Get GetReadMePath() As String
    GetReadMePath = szReadmePath
End Property

Public Property Let SetExpert(ByVal vNewValue As Boolean)
    bExpert = vNewValue                                             ' Experten Modus
End Property

Public Property Get GetExpert() As Boolean
    GetExpert = bExpert
End Property

Public Property Let SetConsole(ByVal vNewValue As Boolean)
    bConsole = vNewValue                                            ' Consolen Modus
End Property

Public Property Get GetConsole() As Boolean
    GetConsole = bConsole
End Property

Public Property Get GetLeet() As Boolean
    GetLeet = bLeet
End Property

Public Property Let SetLeet(ByVal vNewValue As Boolean)
    bLeet = vNewValue
End Property

Public Property Get GetMorse() As Boolean
    GetMorse = bMorse
End Property

Public Property Let SetMorse(ByVal vNewValue As Boolean)
    bMorse = vNewValue
End Property

Public Property Let SetE(ByVal vNewValue As Boolean)
    bE = vNewValue
End Property

Public Property Get GetE() As Boolean
    GetE = bE
End Property

Public Property Let SetShowSplash(ByVal vNewValue As Boolean)
    bNotShowSplash = Not vNewValue
End Property

Public Property Get GetShowSplash() As String
    GetShowSplash = Not bNotShowSplash                              ' Splash anzeigen
End Property

Public Property Let SetINIFile(ByVal vNewValue As String)
    szIniFile = vNewValue                                           ' Name INI Datei
End Property

Public Property Get GetINIFile() As String
    GetINIFile = szIniFile
End Property

Public Property Let SetXMLFile(ByVal vNewValue As String)
    szXmlFile = vNewValue                                           ' Name XML Datei
End Property

Public Property Get GetXMlFile() As String
    GetXMlFile = szXmlFile
End Property

Public Property Get GetImageDir() As String
    GetImageDir = ApplicationInfo.AppImageFolder                    ' Image Verz.
End Property

Public Property Get bUserIsAdmin() As Boolean
    bUserIsAdmin = bAdmin                                           ' User ist Admin
End Property

Public Property Get GetComuterName() As String
    GetComuterName = CompInfo.Compname                              ' Computername dieses Hosts
End Property

Public Property Get GetDomain() As String
    GetDomain = CompInfo.Domain                                     ' ComputerDomain dieses Hosts
End Property

Public Property Get GetUserName() As String
    GetUserName = szUsername                                        ' Name des angemeldeten Users
End Property

Public Property Get GetWinDir() As String
    GetWinDir = WinFolder.Windows                                   ' Windows Verz.
End Property

Public Property Get GetSys32Dir() As String
    GetSys32Dir = WinFolder.WinSystem                               ' System32 Verz.
End Property

Public Property Get GetMSInfoExe() As String
    GetMSInfoExe = szMSInfoExe                                      ' Pfad MSinfo.exe
End Property

Public Property Get GetMStscExe() As String
    GetMStscExe = szMStscExe                                        ' Pfad Remotdesktop
End Property

Public Property Get GetNotepadExe() As String
    GetNotepadExe = szNotepadExe                                    ' Pfad Notepad.exe
End Property
                                                                    ' *****************************************
                                                                    ' Methoden
Public Function FolowLink(szLink As String)
'Public Function FolowLink(f As Form, szLink As String)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If szLink <> "" Then
'        If f Is Nothing Then                                        ' Kein Form angegeben
'            Set f = objObjectBag.GetMainform                        ' Hauptfenster holen
'        End If
        Call objTools.ShellRun(szLink, "")                          ' Link öffnen
        'Call objTools.ShellExec(szLink)                            ' Link öffnen
        'Call ShellExecute(f.hWnd, "Open", szLink, _
                    "", App.Path, 1)                                ' Link öffnen
    End If
    Err.Clear                                                       ' Evtl. error Clearen
End Function

Public Sub ReportBug()
' Bereitet eine Email an AppSuportmail vor mit Protokoll datei im anhang
    Dim szProtPath As String                                        ' Pfad der Protokolldatei
    'Dim szMailTo As String
    'Dim Result
On Error GoTo Errorhandler                                          ' fehlerbehandlung aktivieren
    If ApplicationInfo.AppSuportMail = "" Then GoTo exithandler     ' Keine Suportmail angegeben -> Fertig
    szProtPath = objError.GetErrFileName()                          ' Protokoll pfad ermitteln
    Call ShowNewMail(ApplicationInfo.AppSuportMail, "Bugreport: Mega Admin", _
            "", szProtPath)                                         ' Mail erstellen
exithandler:
Exit Sub                                                            ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ReportBug", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Sub

Public Sub ShowNewMail(szMailAdress As String, szSubject As String, Optional szMailText As String, Optional szDateianhang As String)
    Dim szMailTo As String
    Dim Result
On Error GoTo Errorhandler
    'If szMailAdress = "" Then GoTo exithandler
    If szSubject = "" Then szSubject = GetAppTitle
    'If szMailText <> "" Then szMailText = SetSyntax(szMailText)
    'mailto:test@schwider.de?param1=wert1&amp;param2=wert2&amp;...
    szMailTo = "mailto:" & szMailAdress
    If szSubject <> "" Then
        szMailTo = szMailTo & "?Subject=" & szSubject
    End If
    If szMailText <> "" Then
        szMailTo = szMailTo & "&Body=" & szMailText
    End If
    If szDateianhang <> "" Then
        szMailTo = szMailTo & "&Attach=" & Chr(34) & Chr(34) & szDateianhang & Chr(34) & Chr(34)
    Else
            szMailTo = szMailTo & "&Attach="
    End If
    'szMailTo = "mailto:" & szMailAdress & "?Subject=" & szSubject & "&Body=" & szMailText & "&Attach=" & Chr(34) & Chr(34) & szDateianhang & Chr(34) & Chr(34)
    Result = ShellExecute(0&, "Open", szMailTo, "", "", 1)
    'Call objTools.ShellExec(szMailTo)
exithandler:

Exit Sub                                                            ' Function beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ShowNewMail", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Sub

Public Sub ShowMSInfo()                                             ' Startet MSInfo.exe
On Error Resume Next                                                ' Hier keine Fehlerbehandlung
    If szMSInfoExe = "" Then Exit Sub                               ' MS info nicht gefunden
    If objTools.ExistFile(szMSInfoExe, False) Then                  ' Existiert Datei
        'Call objTools.ShellRun(szMSInfoExe, "")                     ' starten
        Call Shell(szMSInfoExe, vbNormalFocus)                      ' starten
    End If
    Err.Clear                                                       ' Evtl. error clearen
End Sub

Public Function ShowFormChangePWD(szUsername As String, bCancel As Boolean, _
        Optional bShowNextLogin As Boolean, Optional bChangeOnNextLogin As Boolean) As String
    Dim fChangePwd As frmChangePWD
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Set fChangePwd = New frmChangePWD                               ' Verweis aufs Form
    Call fChangePwd.InitObjectBag(Me)
    fChangePwd.lblUsername.Caption = szUsername                     ' Username anzeigen
    'If bShowNextLogin Then
        fChangePwd.chkNextLoginChange.Visible = bShowNextLogin
    'End If
    fChangePwd.Show vbModal, GetMainForm                            ' Form anzeigen
    bCancel = fChangePwd.bCancel                                    ' Abbruch duch benutzer
    bChangeOnNextLogin = fChangePwd.chkNextLoginChange.Value
    ShowFormChangePWD = fChangePwd.txtPWDNew2.Text
exithandler:
On Error Resume Next                                                ' Hier keine Fehler Mehr
Exit Function                                                       ' Function beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ShowFormChangePWD", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Sub ShowReadMe()
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If szReadmePath = "" Then GoTo exithandler                      ' keine Datei angegeben -> Fertig
    If Not objTools.FileExist(szReadmePath) Then GoTo exithandler   ' Datei existiert nicht -> Fertig
    Call objTools.ShellRun(szReadmePath, "", vbNormalFocus)         ' starten
    'Call objTools.ShellExec(szReadmePath, "")
'    Call Shell(szReadmePath, vbNormalFocus)
exithandler:
On Error Resume Next                                                ' Hier keine Fehler Mehr
Exit Sub                                                            ' Function beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ShowReadMe", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Sub

Public Sub ShowAbout(Optional szDesc As String, _
        Optional bWithOfficeInfo As Boolean, _
        Optional DBConn As Object)                                  ' zeigt das info formular an
    Dim fAbout As Form                                              ' Das Info Form
    Dim szImagePath  As String                                      ' Pfad zur bilddatei
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szDetails = "ShowAbout gestartet"                               ' Fehlerdetails festlegen
    Set fAbout = New frmAbout                                       ' Neu initialisieren
    Set fAbout.objObjectBag = Me                                    ' Object bag übergeben
    szDetails = "fAbout initialisiert"                              ' Fehlerdetails festlegen
    fAbout.ImgAbout.Stretch = True                                  ' Image Stretschen
On Error Resume Next                                                ' Fehlerbehandlung deak da Image date evt. nicht vorhanden
    If GetConsole Then
        szImagePath = GetImageDir & "piraten2.bmp"
    Else
        szImagePath = objOptions.GetOptionByName(OPTION_ABOUT_IMG)      ' Image Pfad aus Optionen
    End If
    szDetails = "Picture Pfad. geladen Err:" & Err.Number           ' Details für Fehlerbehandlung
    fAbout.ImgAbout.Picture = LoadPicture(szImagePath)              ' Image laden
    szDetails = "Picture geladen. Err:" & Err.Number                ' Details für Fehlerbehandlung
    Err.Clear                                                       ' Error resetten
On Error GoTo Errorhandler                                          ' Fehlerbehandlung wider aktivieren
    With ApplicationInfo                                            ' Anwendungs info durchgehen
        fAbout.Caption = "Info zu " & .AppTitel                     ' Fenstertitel
        fAbout.lblVersion.Caption = "Version " & .AppVersion        ' version
        fAbout.lblTitle.Caption = .AppTitel                         ' Anwendungstitel
        fAbout.lblDescription.Caption = .AppCopyright               ' Copyright
        If Not DBConn Is Nothing Then                               ' DB Verbindung ist vorhanden
            If DBConn.IsSQLServer And DBConn.bConnect Then          ' DB Ist ein SQL Server
                fAbout.lblDescription.Caption = fAbout.lblDescription.Caption _
                    & vbCrLf & vbCrLf & DBConn.GetSQLVersion()      ' Beschreibung um SQL Version erweitern
            End If
        End If
        If szDesc <> "" Then                                        ' Wenn optinale Beschreibung vorhanden
            fAbout.lblDescription.Caption = fAbout.lblDescription.Caption _
                    & vbCrLf & szDesc                               ' Beschreibung erweitern
        End If
        If .AppWWW <> "" Then                                       ' Anwendugs Web site ist angegeben
            fAbout.lblWWW.Caption = .AppWWW                         ' Eintragen
            fAbout.lblWWW.Visible = True                            ' Anzeigen
        Else                                                        ' Sonst
            fAbout.lblWWW.Visible = False                           ' Nicht anzeigen
        End If
        If bWithOfficeInfo Then                                     ' Office Info soll angezeigt werden
            fAbout.lblDescription.Caption = fAbout.lblDescription & vbCrLf & vbCrLf _
                    & " Instalierte Office Progaramme:" & vbCrLf
            If OfficeVersionen.AccessVersionLong <> "" Then
                fAbout.lblDescription.Caption = fAbout.lblDescription.Caption & " Access " _
                        & OfficeVersionen.AccessVersionLong & vbCrLf
            End If
            If OfficeVersionen.WordVersionLong <> "" Then
                fAbout.lblDescription.Caption = fAbout.lblDescription.Caption & " Word " _
                        & OfficeVersionen.WordVersionLong & vbCrLf
            End If
            If OfficeVersionen.ExcelVersionLong <> "" Then
                fAbout.lblDescription.Caption = fAbout.lblDescription.Caption & " Excel " _
                        & OfficeVersionen.ExcelVersionLong & vbCrLf
            End If
            If OfficeVersionen.OutlookVersionLong <> "" Then
                fAbout.lblDescription.Caption = fAbout.lblDescription.Caption & " Outlook " _
                        & OfficeVersionen.OutlookVersionLong & vbCrLf
            End If
        End If
    End With
    Call CheckFormStyle(fAbout)
    fAbout.Show                                                     ' Form anzeigen
exithandler:
On Error Resume Next                                                ' Hier keine Fehler Mehr
Exit Sub                                                            ' Function beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ShowAbout", errNr, errDesc, szDetails) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Sub

Public Function ShowSplash(bVisible As Boolean)
' Zeigt Splash Form an und blendet es aus
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If Not bVisible Then                                            ' Splash ausbelenden
        If Not fSplash Is Nothing Then                              ' Wenn form geladen
            fSplash.Visible = False                                 ' Unsichtbar
            Unload fSplash                                          ' Splash entladen
        End If
        GoTo exithandler                                            ' Fertig
    End If
    Set fSplash = New frmSplash                                     ' Neues Splash form laden
    Set fSplash.objObjectBag = Me
    Call fSplash.ImageSplash.Move(0, 0, fSplash.ScaleWidth, _
            fSplash.ScaleHeight)                                    ' Pos des Images festlegen
    fSplash.ImageSplash.Stretch = True                              ' Image Stretchen
On Error Resume Next                                                ' Fehlerbehandlung deak.
    fSplash.ImageSplash.Picture = LoadPicture(objOptions.GetOptionByName(OPTION_SPLASH_IMG))
    Err.Clear                                                       ' Error resetten
On Error GoTo Errorhandler                                          ' Fehlerbehandlung wieder aktivieren
    If bE Then fSplash.ImageSplash.Picture = LoadPicture("176746a")
    fSplash.Show                                                    ' Splash form anzeigen
    fSplash.ZOrder
    DoEvents                                                        ' Andere aktionen zulassen
    fSplash.TimerSplash.Interval = 3000                             ' Timer interval setzen
    fSplash.TimerSplash.Enabled = True                              ' Timer starten
exithandler:
    
Exit Function                                                       ' Function beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ShowSplash", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ShowSearch(ThisDBConn As Object, Rootkey As String, SearchField As String, _
        Optional szSearchtext As String, _
        Optional OptWhereID As String, _
        Optional szTitel As String) As String
' Ruft ein Suchformular auf
    Dim NewFrmSearch As Form                                        ' Such Form
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Set NewFrmSearch = frmSuche                                     ' Form Referenz Holen
    Call frmSuche.InitSuche(Me, ThisDBConn, Rootkey, SearchField, _
            szSearchtext, OptWhereID, szTitel)                      ' Such form initialisieren
    If Not NewFrmSearch.Visible Then
    NewFrmSearch.Show vbModal, frmMainWindow                        ' Such Form anzeigen
    End If
    Rootkey = NewFrmSearch.szRootkey                                ' Ergebnis Rootkey (was haben wir gesucht)
    ShowSearch = NewFrmSearch.szResultID                            ' Ergebnis ID zurück geben
exithandler:
On Error Resume Next
    Unload NewFrmSearch                                             ' Such Form entladen
Exit Function                                                       ' Function beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ShowSearch", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function OpenFile(Filter As String, AktForm As Object, Optional DefPath As String) As String
' Ruft datei öffnen Dialog auf
    Dim FileName As String                                          ' zu öffnende Datei
    Dim Flags As Long
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Const OFN_FILEMUSTEXIST As Long = &H1000&
    Const OFN_HIDEREADONLY As Long = &H4&
    Const OFN_PATHMUSTEXIST As Long = &H800&
    Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or _
            OFN_PATHMUSTEXIST                                       ' Flags Setzen
    If DefPath = "" Then DefPath = WinFolder.UserEigeneDateien      ' Besser eigene Dateien als Default verz
    If AktForm Is Nothing Then Set AktForm = frmMainWindow          ' Falls AktForm nicht angegeben MainWindowHolen
    FileName = objTools.ShowOpen(Filter, Flags, _
            AktForm.hwnd, DefPath)                                  ' An objTools weiterreichen
    OpenFile = FileName                                             ' Erg. Zurückliefern
    Err.Clear                                                       ' Evtl Error clearen
End Function

Public Function SaveFile(Filter As String, AktForm As Object, Optional DefPath As String) As String
    Dim Flags As Long
    Dim FileName As String                                          ' zu speichernde Datei
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Const OFN_FILEMUSTEXIST As Long = &H1000&
    Const OFN_HIDEREADONLY As Long = &H4&
    Const OFN_PATHMUSTEXIST As Long = &H800&
    Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or _
            OFN_PATHMUSTEXIST                                       ' Flags Setzen
    If DefPath = "" Then DefPath = WinFolder.UserEigeneDateien      ' Besser eigene Dateien als Default verz
    If AktForm Is Nothing Then Set AktForm = frmMainWindow          ' Falls AktForm nicht angegeben MainWindowHolen
    FileName = objTools.ShowSave(Filter, Flags, _
            AktForm.hwnd, DefPath)                                  ' An objTools weiterreichen
    SaveFile = FileName                                             ' Erg. Zurückliefern
    Err.Clear                                                       ' Evtl Error clearen
End Function

Public Function SelectFolder(szText As String, szDefPath As String) As String
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    SelectFolder = objTools.ShowDirSelect(szText, szDefPath)        ' An objTools Weietrreichen
    Err.Clear                                                       ' Evtl Error clearen
End Function

Public Function CheckFormStyle(f As Object, Optional bForceWinThemes As Boolean)
    Dim c As Control                                                ' Akt. Control
    Dim szTranzRate As String                                       ' Tranzparenz asl str in %
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If f Is Nothing Then Exit Function                              ' Kein Form angegeben -> Fertig
    'Call objTools.FormTo0815(f)
' MW 01.09.11 {
    szTranzRate = objOptions.GetOptionByName(OPTION_TRANZ)          ' Wert Tranzparenz aus Options lesen
    If szTranzRate <> "" Then                                       ' Nur wenn wert vorhanden
        If CSng(szTranzRate) < 1 Then                               ' Nur wenn was zu tun ist
            Call objTools.SetWindowTransparency(f, CSng(szTranzRate)) ' Form Transparent setzen
        End If
    End If
' MW 01.09.11 }
    If GetConsole Then Call objTools.FormToConsole(f)
    If GetLeet Then Call objTools.FormTo1337(f)
    If GetMorse Then Call objTools.FormToMorse(f)
    If Not (GetConsole And GetLeet And GetMorse) Then
        Call objTools.WinTheme(f, True)                             ' Win Themes fürs Form Aktivieren
        For Each c In f                                             ' Alle controls Durchlaufen
            Call objTools.WinTheme(c, True)                         ' Win Themes fürs jeweilige Control aktivieren
        Next                                                        ' nächstes Kontrol
    End If
exithandler:
Exit Function                                                       ' Function beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "CheckFormStyle", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function GetMaxProgresSteps(n As Integer) As Integer
' Ermittelt aus Richtwert (n) möglicher Arbeitschritte MaxWert für Progressbar
    Dim szN As String                                               ' Richtwert als String
    Dim szMAxSteps As String                                        ' Max Steps als String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szN = CStr(n)                                                   ' Richtwert in Str wandeln
    If szN <> "" Then
        szMAxSteps = "1"                                            ' Max wert mit 1 beginnen
        While Len(szN) > 0                                          ' Für jede Stelle
           szMAxSteps = szMAxSteps & "0"                            ' 0 anhängen
           szN = Left(szN, Len(szN) - 1)
        Wend
    End If
    GetMaxProgresSteps = CLng(szMAxSteps)                           ' In Integer wandeln
exithandler:
Exit Function                                                       ' Function beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "GetMaxProgresSteps", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function GetProgresStepsWidth(n As Integer, MaxSteps As Integer) As Integer
' Ermittelt aus Richtwert (n) möglicher Arbeitschritte um MaxSteps eine Schritweite für Progressbar
    Dim Step As Integer
On Error GoTo Errorhandler
    Step = MaxSteps / n
    If Step = 0 Then Step = 1
    GetProgresStepsWidth = Step
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
Exit Function                                                       ' Function beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "GetProgresStepsWidth", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ShowMSGForm(bShow As Boolean, _
        szMSGtext As String, _
        Optional ProgBarMax As Integer, _
        Optional szAction As String)
' Initialisiert und Zeigt msg dialog an
    Dim bShowInSlpash As Boolean                                    ' Meldung im Splash anzeigen
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szMSGtext = Trim(szMSGtext)                                     ' Meldungstext Trimmen
    szAction = Trim(szAction)                                       ' Aktionstext sicherheitshalber trimmen
    If fSplash Is Nothing Then
        bShowInSlpash = False
    Else
        If fSplash.Visible Then
            bShowInSlpash = True
        Else
            bShowInSlpash = False
        End If
    End If
    If bShowInSlpash Then                                           ' Wenn Splach dann kein MSG Form
        If szMSGtext <> "" Then                                     ' Evtl. Action Label initialisieren
            fSplash.lblAction.Caption = szMSGtext                   ' Meldungtext in Splash setzen
            fSplash.lblAction.Visible = True                        ' Meldung sichtbar
        Else
            frmMSG.lblAction.Visible = False
        End If
    Else
        If ProgBarMax > 0 Then                                      ' Evtl. Prog Bar initialisieren
            frmMSG.ProgressBarMSG.Max = ProgBarMax                  ' Max Wert für Progbar festlegen
            frmMSG.ProgressBarMSG.Value = 1                         ' Startwert Progbar festlegen
            frmMSG.ProgressBarMSG.Visible = True                    ' Progbar einblenden
        Else
            frmMSG.ProgressBarMSG.Visible = False                   ' Progbar ausblenden
        End If
    
        If szAction <> "" Then                                      ' Evtl. Action Label initialisieren
            frmMSG.lblAction.Caption = szAction                     ' Aktionstext in MSG Form setzen
            frmMSG.lblAction.Visible = True
        Else
            frmMSG.lblAction.Visible = False                        ' Aktionstext ausblenden
        End If
    
        If bShow Then                                               ' MSG Form anzeigen ?
            frmMSG.lblMSG.Caption = szMSGtext                       ' Meldungstext in MSG form stzen
            frmMSG.Show                                             ' Anzeigen
            frmMSG.Refresh                                          ' MSG Form aktualisieren
        Else                                                        ' MSG Form Ausblenden?
            frmMSG.Hide                                             ' Ausblenden
        End If
    End If
    
exithandler:
    
Exit Function                                                       ' Function beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ShowMSGForm", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ShowMsgNextStep(szAction As String, Optional lngStep As Integer)
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    frmMSG.lblAction.Caption = szAction
    If frmMSG.ProgressBarMSG.Visible Then
        If lngStep = 0 Then
            If frmMSG.ProgressBarMSG.Value + 1 > frmMSG.ProgressBarMSG.Max Then
                frmMSG.ProgressBarMSG.Value = frmMSG.ProgressBarMSG.Max
            Else
                frmMSG.ProgressBarMSG.Value = frmMSG.ProgressBarMSG.Value + 1
            End If
        Else
            If frmMSG.ProgressBarMSG.Value + lngStep > frmMSG.ProgressBarMSG.Max Then
                frmMSG.ProgressBarMSG.Value = frmMSG.ProgressBarMSG.Max
            Else
                frmMSG.ProgressBarMSG.Value = frmMSG.ProgressBarMSG.Value + lngStep
            End If
        End If
    End If
    
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
Exit Function                                                       ' Function beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ShowMsgNextStep", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function InitDBConnection(Optional bAccPossible As Boolean, _
        Optional bSQLPossible As Boolean) As Boolean
' DBConnKlasse Initialisieren (Latebinding)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Set objDBconn = CreateObject("DBConnect.clsDBConn")
    If Not objDBconn Is Nothing Then                                ' Prüfen ob erfolg
        InitDBConnection = True
        Call objDBconn.InitObjectBag(Me)                            ' OB ref. weitergeben
        If (bAccPossible = False And bSQLPossible = False) Then
            bAccPossible = True
            bSQLPossible = True
        End If
        objDBconn.bAccessPossible = bAccPossible
        objDBconn.bMSSQLSrvPossible = bSQLPossible
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Public Function InitSQLTools() As Boolean
' SQLToolsKlasse  Initialisieren (Latebinding)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Set objSQLTools = CreateObject("SQLTools.clsSQLTools")
    If Not objSQLTools Is Nothing Then                              ' Prüfen ob erfolg
        InitSQLTools = True
        Call objSQLTools.InitObjectBag(Me)                          ' OB ref. weitergeben
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Public Function InitOfficeTools() As Boolean
' SQLToolsKlasse  Initialisieren (Latebinding)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Set objOfficeTools = CreateObject("OfficeTools.clsOfficeTools")
    If Not objOfficeTools Is Nothing Then                           ' Prüfen ob erfolg
        InitOfficeTools = True
        Call objOfficeTools.InitObjectBag(Me)                       ' OB ref. weitergeben
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Public Function InitOptions() As Boolean
' SQLToolsKlasse  Initialisieren (Latebinding)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Set objOptions = CreateObject("Options.clsOptions")
    If Not objOptions Is Nothing Then                               ' Prüfen ob erfolg
        InitOptions = True
        Call objOptions.InitObjectBag(Me)                           ' OB ref. weitergeben
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Function
                                                                    ' *****************************************
                                                                    ' Hilfsfunktionen
Private Function GetSysteminfo()
' Informationen zur Systemumgebung zusammtragen.
Dim objWMIService As Object
Dim colOSItems As Object
Dim objOSItem As Object
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    CompInfo.Compname = objTools.GetCurCompName                     ' Computername Auslesen
    objError.SetComuterName = CompInfo.Compname                     ' Für Prot + errorlog an objerror
    szUsername = objTools.getLokalUserName                          ' Username Auslesen
    objError.SetUserName = szUsername                               ' Für Prot + errorlog an objerror
    Set objWMIService = GetObject("winmgmts:")                      ' Anmelde domäne ermitteln Per WMI
    Set colOSItems = objWMIService.InstancesOf("Win32_ComputerSystem")
    If colOSItems.Count > 0 Then
        For Each objOSItem In colOSItems
            DoEvents                                                ' Anderere Events zulassen
            CompInfo.Domain = objOSItem.Domain
        Next
    End If
    If IsUserAdmin(szUsername, CompInfo.Domain) Then bAdmin = True  ' User (lokaler) Admin?
    If CompInfo.Domain = CompInfo.Compname Then CompInfo.Domain = "" ' Domain?
    Call GetSytemFolder                                             ' Windows verzeichnisse ermitteln
    DeskWidth = Screen.Width / Screen.TwipsPerPixelX                ' Bildschirm auflösung (breite)
    DeskHeight = Screen.Height / Screen.TwipsPerPixelY              ' Bildschirm auflösung (Höhe)
    szMSInfoExe = ReadMSInfoPath()                                  ' Pfad MSINFO.exe aus reg ermitteln
    szMStscExe = WinFolder.WinSystem & "\mstsc.exe"                 ' mstsc.exe (Remotdesktop)
    szNotepadExe = WinFolder.WinSystem & "\notepad.exe"             ' notepad.exe
    Call ReadOfficeInfo                                             ' Informationen über inst Office versionen
' ...
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "GetSysteminfo", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Private Function ReadMSInfoPath()
    Dim path As String                                              ' Pfad zur MSinfp.exe
On Error Resume Next                                                ' Hier keine Fehlerbehandlung
    path = objReg.ReadRegValue("HKLM", "SOFTWARE\Microsoft\Shared Tools\MSINFO", _
            "PATH")                                                 ' 1. Peg Schlüssel prüfen
    If path = "" Then                                               ' Wenn nicht gefunden
        path = objReg.ReadRegValue("HKLM", "SOFTWARE\Microsoft\Shared Tools Location", _
                "MSINFO")                                           ' 2. Schlüssel Prüfen
        If path <> "" Then path = path & "\MSINFO32.EXE"            ' Wenn was gefunden exe an Pfad anhängen
    End If
    If Not (Dir(path) <> "") Then                                   ' Prüfen ob Pfad ex.
        path = ""                                                   ' Sonst Leer
    End If
    ReadMSInfoPath = path                                           ' Ergebnis zurück
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Private Sub ReadOfficeInfo()
' Ermittelt Pfad und Version von Word, Excel, Access & Outlook
' aus der Registry
    Dim KeyName As String                                           ' zu überprüfender Reg schlüssel
    Dim KeyValue As String                                          ' Reg Wert
    Dim path As String                                              ' Exe Path
    Dim Version As String                                           ' Aktuelle version
    Dim Produkt As String                                           ' Überprüftes Produkt (Word, Access,...)
' MW 17.04.2013 Abbruch bed. ist Maxversion -1 damit zu früh für Office 2010 {
'    Const MaxVersion = "14"                                         ' 12 ist office 2007 (in die zukunft gedacht
    Const MaxVersion = "20"                                         ' 12 ist office 2007 & 14 = office 2010 (in die zukunft gedacht
' MW 17.04.2013 }
    Const MinVersion = "8"                                          ' Office 97
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Produkt = "Access"                                              ' AccessVersion
    Version = MinVersion                                            ' von Version 8.0 hochzählen
    While Version <> ""
        'KeyName = "SOFTWARE\Microsoft\Office\" & Version & ".0\" & Produkt & "\InstallRoot"
        'KeyName = "SOFTWARE\Classes\" & Produkt & ".Application." & Version & "\shell\open\command\"
        KeyName = "SOFTWARE\Microsoft\Office\" & Version & ".0\Common\FileNew\NFT\General\Blank Database"
        KeyValue = "Command"
        path = Trim(objReg.ReadRegValue("HKLM", KeyName, KeyValue)) ' Zeichenfolge auslesen
        If path = "" Then                                           ' Wenn Kein Pfad
            Version = Version + 1                                   ' Mit neuer Version prüfen
            If Version = MaxVersion Then Version = ""               ' Version 14 (nachfolger vom Nachfolger 2007)
        Else                                                        ' Sonst
            If InStr(path, " ") > 0 Then                            ' evtl Leerzeichen im pfad
                path = Trim(Left(path, InStr(path, "/") - 2))       ' rausschmeissen
'                path = Right(path, Len(path) - 1)
            End If
            OfficeVersionen.AccessPath = path                       ' Exe Pfad merken
            OfficeVersionen.AccessVersion = Version                 ' Version Merken
            Version = ""                                            ' Abbruch bed. setzen
        End If
    Wend
    Produkt = "Word"                                                ' Word Versionen
    Version = MinVersion                                            ' von Version 8.0 hochzählen
    While Version <> ""                                             ' Bis Abbruch
        KeyName = "SOFTWARE\Microsoft\Office\" & Version & ".0\" & Produkt & "\InstallRoot"
        'KeyName = "SOFTWARE\Classes\" & Produkt & ".Application." & Version & "\shell\open\command\"
        path = Trim(objReg.ReadRegValue("HKLM", KeyName, "path"))   ' Zeichenfolge auslesen
        If path = "" Then                                           ' Wenn Kein Pfad
            Version = Version + 1                                   ' Mit neuer Version prüfen
            If Version = MaxVersion Then Version = ""               ' Version 14 (nachfolger vom Nachfolger 2007)
        Else                                                        ' Sonst
            OfficeVersionen.WordPath = path & "WINWORD.EXE"         ' ExeP fad merken
            OfficeVersionen.WordVersion = Version                   ' Version merken
            Version = ""                                            ' Abbruch bed. setzen
        End If
    Wend
    Produkt = "Excel"                                               ' Excel Versionen
    Version = MinVersion                                            ' von Version 8.0 hochzählen
    While Version <> ""                                             ' Bis Abbruch
        KeyName = "SOFTWARE\Microsoft\Office\" & Version & ".0\" & Produkt & "\InstallRoot"
        'KeyName = "SOFTWARE\Classes\" & Produkt & ".Application." & Version & "\shell\open\command\"
        path = Trim(objReg.ReadRegValue("HKLM", KeyName, "path"))   ' Zeichenfolge auslesen
        If path = "" Then                                           ' Wenn Kein Pfad
            Version = Version + 1                                   ' Mit neuer Version prüfen
            If Version = MaxVersion Then Version = ""               ' Version 14 (nachfolger vom Nachfolger 2007)
        Else                                                        ' Sonst
            OfficeVersionen.ExcelPath = path & "EXCEL.EXE"          ' exe pfad merken
            OfficeVersionen.ExcelVersion = Version                  ' Version merken
            Version = ""                                            ' Abbruch bed. setzen
        End If
    Wend
    Produkt = "Outlook"                                             ' Outlook versionen
    Version = MinVersion                                            ' von Version 8.0 hochzählen
    While Version <> ""                                             ' Bis Abbruch
        KeyName = "SOFTWARE\Microsoft\Office\" & Version & ".0\" & Produkt & "\InstallRoot"
        'KeyName = "SOFTWARE\Classes\" & Produkt & ".Application." & Version & "\shell\open\command\"
        path = Trim(objReg.ReadRegValue("HKLM", KeyName, "path"))   ' Zeichenfolge auslesen
        If path = "" Then                                           ' Wenn Kein Pfad
            Version = Version + 1                                   ' Mit neuer Version prüfen
            If Version = MaxVersion Then Version = ""               ' Version 14 (nachfolger vom Nachfolger 2007)
        Else                                                        ' Sonst
            OfficeVersionen.OutlookPath = path & "OUTLOOK.EXE"      ' exe Pfad merken
            While Version <> ""                                             ' Bis Abbruch
                If InStr(UCase(path), "OFFICE" & Version & "\") > 0 Then
                    OfficeVersionen.OutlookVersion = Version            ' Version merken
                    Version = ""
                Else
                     Version = Version + 1
                End If
            Wend
            Version = ""                                            ' Abbruch bed. setzen
        End If
    Wend
exithandler:
    Call GetOfficeVersionLong                                       ' Für gefundene Versionen Langbez. ermitteln
Exit Sub                                                            ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ReadOfficeInfo", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Sub

Private Function GetOfficeVersionLong() As String
On Error Resume Next                                                ' Keine Fehlerbehandlung
    With OfficeVersionen
        .AccessVersionLong = TranslateOfficeVers(.AccessVersion)    ' Access
        .ExcelVersionLong = TranslateOfficeVers(.ExcelVersion)      ' Excel
        .WordVersionLong = TranslateOfficeVers(.WordVersion)        ' Word
        .OutlookVersionLong = TranslateOfficeVers(.OutlookVersion)  ' Outlook
    End With
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Private Function TranslateOfficeVers(szVersion As String) As String
On Error Resume Next                                                ' Keine Fehlerbehandlung
    If szVersion = "" Then Exit Function                            ' Kein Versionsstring -> fertig
    If Not IsNumeric(szVersion) Then                                ' Version nicht numerisch
        TranslateOfficeVers = szVersion                             ' Dann direkt angeben
        Exit Function                                               ' Fertig
    End If
    If CLng(szVersion) < 7 Then TranslateOfficeVers = szVersion     ' Version kleiner 7 -> will ich nicht kennen
    If CLng(szVersion) = 7 Then TranslateOfficeVers = "95"          ' Wer sowas noch im Produktivbetrieb hat, bekommt von mir einen Trostbonscher
    If CLng(szVersion) = 8 Then TranslateOfficeVers = "97"          ' Access 97 Schimmelt hier überall noch rum
    If CLng(szVersion) = 9 Then TranslateOfficeVers = "2000"        ' Solls noch irgendwo geben
    If CLng(szVersion) = 10 Then TranslateOfficeVers = "XP"         ' Haben einge
    If CLng(szVersion) = 11 Then TranslateOfficeVers = "2003"       ' Unsere Version
    If CLng(szVersion) = 12 Then TranslateOfficeVers = "2007"       ' Aktuelle version
    If CLng(szVersion) = 14 Then TranslateOfficeVers = "2010"       ' Ist Geplant (Marketingname steht noch nicht fest)
    If CLng(szVersion) > 14 Then TranslateOfficeVers = szVersion    ' Verion -> 14 Gibts noch nicht
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Private Function GetGlobalVars()
' Protokoll path
' Megs DB Connection
' Error
On Error GoTo Errorhandler

exithandler:

Exit Function                                                       ' Function Beenden
Errorhandler:
On Error Resume Next
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
On Error Resume Next                                                ' Kein Fehler in der Fehlerbehandlung
    Call OBErrorhandler(MODULNAME, "GetGlobalVars", errNr, errDesc)
    Resume exithandler
End Function

Private Function InitAppDir()
' initialisiert Ordner Struktur im anwendungsverzeichniss
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    With ApplicationInfo
        If Not Right(.AppFolder, 1) = "\" Then _
                .AppFolder = .AppFolder & "\"                       ' Evtl. \ an Appfolder anhängen
        .AppImageFolder = .AppFolder & "images\"                    ' BildVerz festlegen
        Call objTools.FileExist(.AppImageFolder, True)              ' Sicherstellen das Bildverz. existiert
        .AppTemplateFolder = .AppFolder & "Vorlagen\"               ' Vorlagenverz. festlegen
        Call objTools.FileExist(.AppTemplateFolder, True)           ' Sicherstellen das Vorlagenverz. existiert
        .AppSQLFolder = .AppFolder & "SQL\"                         ' SQL verz. festlegen
        Call objTools.FileExist(.AppSQLFolder, True)                ' Sicherstellen das SQL verz. existiert
    End With
    
exithandler:

Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "InitAppDir", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Private Function InitErrorHandler() As Boolean
' ErrorKlasse Initialisieren (Latebinding)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Set objError = CreateObject("Error.clsErrorHandler")            ' Object erstellen
    If Not objError Is Nothing Then                                 ' Prüfen ob erfolg
        InitErrorHandler = True                                     ' Erfolg zurück liefern
        Call objError.InitObjectBag(Me)                             ' OB ref. weitergeben
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Function
                                                                    
Private Function InitRegTools() As Boolean
' RegToolsKlasse Initialisieren (Latebinding)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Set objReg = CreateObject("RegTools.clsRegTools")
    If Not objReg Is Nothing Then                                   ' Prüfen ob erfolg
        InitRegTools = True
        Call objReg.InitObjectBag(Me)                               ' OB ref. weitergeben
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Private Function InitTools() As Boolean
' RegToolsKlasse Initialisieren (Latebinding)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Set objTools = CreateObject("Tools.clsTools")
    If Not objTools Is Nothing Then                                 ' Prüfen ob erfolg
        InitTools = True
        Call objTools.InitObjectBag(Me)                             ' OB ref. weitergeben
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Public Function GetCMDIco() As Object
    Set GetCMDIco = frmAbout.ILAbout.ListImages(3).Picture
End Function

Public Function OBErrorhandler(szModulname As String, szFktName As String, szErrNr As String, szErrDesc As String, Optional szErrDetails)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If objError Is Nothing Then
        MsgBox "Initialierung der Fehlerbehandlung fehlgeschlagen." & vbCrLf & szErrDetails, vbCritical, "Schwerer Fehler"
    Else
       Call objError.Errorhandler(szModulname, szFktName, szErrNr, szErrDesc)
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Function
