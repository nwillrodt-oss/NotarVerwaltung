VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit                                                     ' Variaben Deklaration erzwingen
Option Compare Text                                                 ' Sortierreihenfolge festlegen

Private Const MODULNAME = "clsDBConn"                               ' Modulname für Fehlerbehandlung

Private objObjectBag As Object                                      ' ObjectBag Class
Private objError  As Object                                         ' Glob Error Class
Private objTools As Object                                          ' Tools Class
Private objOptions As Object                                        ' Optionen Class
Private objRegTools As Object                                       ' Regtool Class
Private objSQLTools As Object                                       ' SQL Hilfsmittel
Private fMain As Object                                             ' Ref. aufs Hauptform
Public dbConAdo As adodb.Connection                                 ' Aktuelle Connection
Private ConParam As ConnectionParams                                ' Aktuelle Verbindungsparameter
Private szRegTrenn                                                  ' Trennzeichen für Connectstring in registry
Public bConnect As Boolean                                          ' Verbunden
Public bCancel As Boolean                                           ' Abbruch
Public bTransaction As Boolean                                      ' True wenn transaktion läuft
Public bAccessPossible As Boolean                                   ' Access DB zulassig
Public bMSSQLSrvPossible As Boolean                                 ' SQL Server zulässig
Private bSingleSignOnFailed As Boolean                              ' SingelSignOn Gescheitert

Private lngMaxRegCon As Integer                                     ' Max Anzahl von gespeicherten Verbindungen in Registry
Private Const REG_KEY_CON = "Connections\"                          ' Unterschlüssel in dem DB Verbindungen gespeichert werden"
Private Const TYP_ACC = 1                                           ' Verbindungstyp Access
Private Const TYP_SQL = 2                                           ' Verbindungstyp MS SQL Server

Public Type ConnectionParams                                        ' Verbindungsparameter
    bSQL As Boolean                                                 ' True wenn SQL Server ; False wenn Access
    Provider As String                                              ' DB Provider
    SQLServer As String                                             ' SQL Server name
    DBText As String                                                ' Beschreibung der Datenbank für Anzeige
    DBName As String                                                ' DB Name (bei Access mit Pfad)
    DBUser As String                                                ' Benutzername der datenbank
    DBPWD As String                                                 ' DB Kennwort
    CryptPWD As String                                              ' DB Kennwort Verschlüsselt
    bNt As Boolean                                                  ' NT Autentifizierung (nur SQL)
    ConnectString As String                                         ' Completer Connect String
    bUserlogIn As Boolean                                           ' Gibt es ein seperates Benutzer login
    bSingleSignOn As Boolean                                        ' Anmeldung von win durchreichen
    LoginText As String                                             ' Text Für Benutzerlogin
    UserTable As String                                             ' Tabelle mit Benutzerdaten
    AddWhere As String                                              ' Opt. Where für Benutzertabelle
    UsernameField As String                                         ' Feld mit Benutzernamen
    UsernameField2 As String                                        ' Feld mit alternativen Benutzernamen (z.B. angezeigter name und NtAnmeldename)
    PWDField As String                                              ' Feld Mit Kennwort
    UserLoginMaxCount As Integer                                    ' Max Anmelde versuche
End Type

Private Sub Class_Initialize()
    szRegTrenn = ";"                                                ' Trennzeichen für Registry Festlegen
    lngMaxRegCon = 20                                               ' Max gespeicherte Verbindungen
End Sub

Private Sub Class_Terminate()
On Error Resume Next                                                ' Hier keine Fehlerbehandlung
    Set objOptions = Nothing                                        ' Options schliesssen
    Set objRegTools = Nothing                                       ' Reg tools schliessen
    Set objTools = Nothing                                          ' Tools Object schliessen
    Set objSQLTools = Nothing                                       ' SQL Tools schliessen
    Set objObjectBag = Nothing                                      ' Object Bag schliesen
    Set objError = Nothing                                          ' Error Class Schliessen
    Err.Clear                                                       ' Evtl. error clearen
End Sub

'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=mega;Initial Catalog=dbolgtest;Data Source=OLG-mobile003
                                                                    ' *****************************************
                                                                    ' Properties
Public Property Get TransStatus()
    TransStatus = bTransaction                                      ' Läuft eine Transaktion
End Property

Public Property Get GetServername()
    If bConnect Then                                                ' Wenn Verbunden
        GetServername = ConParam.SQLServer                          ' Akt Servername
    End If
End Property

Public Property Get GetDBName()
    If bConnect Then                                                ' Wenn Verbunden
        GetDBName = ConParam.DBName                                 ' Akt Datenbankname
    End If
End Property

Public Property Get GetDBUsername()
    If bConnect Then                                                ' Wenn Verbunden
        GetDBUsername = ConParam.DBUser                             ' Akt DB Benutzer
    End If
End Property

Public Property Get GetConnectString()
    If bConnect Then                                                ' Wenn Verbunden
        GetConnectString = ConParam.ConnectString                   ' Connect string
    End If
End Property

Public Property Get GetDBtext()
    If bConnect Then                                                ' Wenn Verbunden
        GetDBtext = ConParam.DBText                                 ' ?
    End If
End Property

Public Property Get IsSQLServer() As Boolean
    IsSQLServer = ConParam.bSQL                                     ' SQL Server Verbindung
End Property

Public Property Get GetMaxRegCon() As Integer
    GetMaxRegCon = lngMaxRegCon                                     ' MAx Gespeicherte Verbindngen in der Registry
End Property
                                                                    ' *****************************************
                                                                    ' Methoden
Public Function InitObjectBag(objOb As Object)
'Initialisiert den ObjectBag
    Set objObjectBag = objOb
    If Not objObjectBag Is Nothing Then
        Set objError = objObjectBag.GetErrorObj                     ' Error Class aus Objectbag holen
        Set objOptions = objObjectBag.GetOptionsObj                 ' Options Class aus Objectbag holen
        Set objTools = objObjectBag.getToolsObj                     ' Tools Class aus Objectbag holen
        Set objSQLTools = objObjectBag.getSQLToolsObj
        Set objRegTools = objObjectBag.GetRegToolsObj               ' Registry Class holen
        Set fMain = objObjectBag.GetMainForm                        ' Referenz aufs Main form
    End If
End Function

Public Function GetDBErrors()
    
    
End Function

Public Function GetADODBConn(ingType As Integer, _
        szServername As String, _
        szDBName As String, _
        Optional szDBUser As String, _
        Optional szPWD As String, _
        Optional bNtAut As Boolean, _
        Optional bCancel As Boolean, _
        Optional bNoMsg As Boolean) As Boolean
' Stellt eine ADO datenbank verbindung her
    Dim szConnect As String                                         ' Connect String
    Dim szDetails As String                                         ' Genauere Fehlerbeschreibung
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    With ConParam
        .bNt = bNtAut                                               ' NT Autentifikation
        .SQLServer = szServername                                   ' Servername (nur SQL)
        .DBName = szDBName                                          ' Datenank Name
        .DBUser = szDBUser                                          ' Benutzername
        .DBPWD = szPWD                                              ' Benutzer Passwort
        Select Case ingType                                         ' Art der Datenbank
        Case 1, TYP_ACC                                             ' Access
'            .Provider = "Microsoft.Jet.OLEDB.3.51"
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .bSQL = False                                           ' Kein SQL Server
            .bNt = False                                            ' Keine NT Autentifikation
        Case 2, TYP_SQL                                             ' SQL Server
            .Provider = "sqloledb.1"
            .bSQL = True                                            ' ist SQL Server
        Case Else
            .Provider = "sqloledb.1"
            .bSQL = True                                            ' ist SQL Server
            .bNt = False                                            ' Keine NT Autentifikation
        End Select
        If .DBName = "" Then                                        ' Kein DB Name angegeben
            Call OpenConnectParameterForm(bCancel)                  ' Verbindungs parameter abfagen
            GoTo exithandler                                        ' und raus
        End If
        If .DBUser = "" And Not .bNt And .bSQL Then                 ' Bein DB User bei SQL (ohne NTAut.)
            Call OpenConnectParameterForm(bCancel)                  ' Verbindungs parameter abfagen
            GoTo exithandler                                        ' und raus
        End If
    End With
    If Not bNoMsg Then                                              ' Wenn Meldung gewünscht
        If ingType = TYP_SQL Then
            Call objObjectBag.ShowMSGForm(True, "Datenbankverbindung zu " & ConParam.DBName & " auf " _
                    & ConParam.SQLServer & " wird hergestellt ...") ' Meldung für SQL anzeigen
        Else
            Call objObjectBag.ShowMSGForm(True, "Datenbankverbindung zu " & ConParam.DBName _
                    & " wird hergestellt ...")                      ' Meldung für Acc anzeigen
        End If
    End If
    'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Referend\DataBase.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False
'    Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Referend\DataBase.mdb;Mode=ReadWrite|Share Deny Write
    With ConParam
        Select Case ingType
        Case 1                                                      ' Access ohn eKennwort und User
'            .ConnectString = "Provider=" & .Provider & ";Data Source=" & .DBName 'szProvider = "Microsoft.Jet.OLEDB.3.51"
'            "Driver={Microsoft Access Driver (*.mdb)};Dbq=nwind.mdb;" & _
                     "DefaultDir=C:\program files\devstudio\vb;Uid=Admin;Pwd=;"
            .ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
                    & .DBName & ";Mode=ReadWrite"
            .DBText = "Access: " & .DBName
        Case 2                                                      ' MS SQL Server mit Sql autentifizierung
            .ConnectString = "Provider=" & .Provider & ";Data Source=" & .SQLServer & ";" _
                & "Initial Catalog=" & .DBName & ";"                ' szProvider = "sqloledb"
            If .bNt Then                                            ' NT autentifikation
                .ConnectString = "Data Source=" & .SQLServer & ";" _
                        & "Initial Catalog=" & .DBName & ";"
                .ConnectString = .ConnectString & "Integrated Security=SSPI;"
            Else                                                    ' SQL Autentifikation
                If .DBUser <> "" Then
                    .ConnectString = .ConnectString & "User ID=" & .DBUser & ";"
                End If
                If szPWD <> "" Then
                    .ConnectString = .ConnectString & "Password=" & .DBPWD & ";"
                End If
            End If
           .DBText = "MS SQL: " & .DBName & " auf " & .SQLServer
        'Case 3                                                     ' MS SQL Server mit NT autentifizierung
        Case Else
    
        End Select
    End With
    Set dbConAdo = New adodb.Connection                             ' Neue Verbindung öffnen
    With dbConAdo                                                   ' Verbindungs Parameter setzen
        .Mode = adModeReadWrite
        .CursorLocation = adUseClient
        .Provider = ConParam.Provider
        .ConnectionString = ConParam.ConnectString
'        .Mode = adModeShareDenyNone
        .Open                                                       ' Verbindung öffnen
    End With
    Call SaveConnection                                             ' Verbindung Speichern
    If Not bNoMsg Then                                              ' Wenn Meldung gewünscht
        If ingType = TYP_SQL Then
            Call objObjectBag.ShowMSGForm(True, "Datenbankverbindung zu " & ConParam.DBName & " auf " _
                    & ConParam.SQLServer & " ist hergestellt!")     ' Meldung für SQL anzeigen
        Else
            Call objObjectBag.ShowMSGForm(True, "Datenbankverbindung zu " & szDBName _
                    & " ist hergestellt!")                          ' Meldung für Acc anzeigen
        End If
    End If
    bConnect = True                                                 ' Connectflag auf erfolg
exithandler:
On Error Resume Next                                                ' Hier keinen Fehler mehr
    GetADODBConn = bConnect                                         ' Connect Flag zurück geben
    Call objObjectBag.ShowMSGForm(False, "...")                     ' Meldungs Form gegebenenfalls Ausblenden
    szDetails = ""                                                  ' Fehler details löschen
    Err.Clear                                                       ' Evtl Error Clearen
Exit Function                                                       ' Funkion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objObjectBag.ShowMSGForm(False, "")                        ' Message form evtl .asblenden
    Select Case errNr                                               ' Nur bestimmte Fehler ausgeben
    Case "-2147217843"
        If ingType = 1 Then
            Call objError.Errorhandler(MODULNAME, "GetADODBConn", errNr, errDesc)
        ElseIf ingType = 2 Then                                     ' Falsches Kennwort oder falscher benutzername
'            szDetails = "Der SQL Server '" & szServername & "' ist nicht erreichbar." _
'            & "Überprüfen Sie die schreibweise des Servernamens und " _
'            & "stellen Sie sicher das die Namensauflösung einwandfrei funktioniert."
            Call objError.Errorhandler(MODULNAME, "GetADODBConn", errNr, errDesc, szDetails)
        Else
        
        End If
    Case "-2147467259"
        If ingType = 1 Then                                         ' Meldung das DB Pfad nicht gefunden werden konnte bitte in oPtionen überprüfen
                                                                    ' Weiter mit usedb false
            Call objError.Errorhandler(MODULNAME, "GetADODBConn", errNr, errDesc)
        ElseIf ingType = 2 Then
                                                                    ' 1. SQL Server existiert nicht oder ist nicht erreichbar
                                                                    ' 2. Falscher Datenbank name
                                                                    ' 3. Benutzername Falsch
'            szDetails = "Der SQL Server '" & szServername & "' ist nicht erreichbar." _
'            & "Überprüfen Sie die schreibweise des Servernamens und " _
'            & "stellen Sie sicher das die Namensauflösung einwandfrei funktioniert."
            Call objError.Errorhandler(MODULNAME, "GetADODBConn", errNr, errDesc, szDetails)
        Else
        
        End If
    Case Else
        Call objError.Errorhandler(MODULNAME, "GetADODBConn", errNr, errDesc)
    End Select
    Resume exithandler
End Function
' Ein paar bsp. connect strings
'For Standard security:
'strConnect = _T("Provider=sqloledb;Data Source=MyServerName;" _
        "Initial Catalog=MyDatabaseName;" _
        "User Id=MyUsername;Password=MyPassword;");
'For Trusted Connection security (Microsoft Windows NT integrated security):
'strConnect = _T("Provider=sqloledb;Data Source=MyServerName;" _
        "Initial Catalog=MyDatabaseName;" _
        "Integrated Security=SSPI;");
'If you want to connect to a "Named Instance" (SQL Server 2000), you must to specify Data Source=Servere Name\Instance Name like in the following example:
'strConnect = _T("Provider=sqloledb;Data Source=MyServerName\MyInstanceName;" _
    "Initial Catalog=MyDatabaseName;User Id=MyUsername;Password=MyPassword;");
'If you want to connect with a SQL Server running on the same computer, you must specify the keyword (local) in the Data Source like in the following example:
'strConnect = _T("Provider=sqloledb;Data Source=(local);" _
        "Initial Catalog=myDatabaseName;" _
        "User ID=myUsername;Password=myPassword;");
' To connect to SQL Server running on a remote computer (via an IP address):
'strConnect = _T("Provider=sqloledb;Network Library=DBMSSOCN;" _
        "Data Source=130.120.110.001,1433;" _
        "Initial Catalog=MyDatabaseName;User ID=MyUsername;" _
        "Password=MyPassword;");
        
Public Function ShowUserLogin() As String
' Zeigt Benutzer Loginform
    Dim UserAnmeldename As String                                   ' für UserLogin
    Dim UserPWD As String                                           ' Benutzer PWD
    Dim rsUser As adodb.Recordset                                   ' RS mit Ergebissen
    Dim szSQL As String                                             ' SQL Statement
    Dim bCancel As Boolean                                          ' Abbruch durch benutzer
    Dim bCount As Boolean                                           ' Abbruch wegen überzähliger anmelde versuche
    Dim bOK As Boolean                                              ' Anmeldung Erfolgreich
    Dim Count As Integer                                            ' Counter für Anmelde Versuche
    Dim szFieldlist As String                                       ' Feldliste für SQL Statement
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    With ConParam
        If .bUserlogIn Then                                         ' Benutzer login ist gefordert
            If Not .bSingleSignOn Then bSingleSignOnFailed = True
            If .UsernameField <> "" Then szFieldlist = .UsernameField & ", "
            If .UsernameField2 <> "" Then szFieldlist = szFieldlist & .UsernameField2 & ", "
            If .PWDField <> "" Then szFieldlist = szFieldlist & .PWDField
            While bCount Or bCancel Or Not bOK                      ' Solange kein Abbruch, kein Erfolg oder noch anmelde versuche
                szSQL = "SELECT " & szFieldlist & _
                        " FROM " & .UserTable & " WHERE "           ' Anfang SQL Statment zusammen setzen
                If .AddWhere <> "" Then                             ' Ist Optionale Where vorhanden
                    szSQL = szSQL & .AddWhere & " AND "             ' Anhängen + AND (da koommt nochr mehr)
                End If
                If Not .bSingleSignOn Or Count > 0 Then             ' Kein Single SignOn
                    frmUserLogin.lblLogintext = .LoginText          ' Anmelde fenster vorbereiten
                    Call objObjectBag.CheckFormStyle(frmUserLogin)
                    frmUserLogin.Show 1, fMain                      ' und anzeigen
                    UserAnmeldename = frmUserLogin.txtMegaUser.Text ' Benutzername auslesen
                    UserPWD = frmUserLogin.txtMwgaPWD.Text          ' PWD auslesen
                    bCancel = frmUserLogin.bCancel                  ' Cancel auslesen
                    szSQL = szSQL & "(" & .UsernameField & "='" & UserAnmeldename & "'"
                    If .UsernameField2 <> "" Then szSQL = szSQL & " OR " & _
                            .UsernameField2 & "='" & UserAnmeldename & "'" ' evtl. 2, Namensfeld abfragen
                    szSQL = szSQL & ") AND " & .PWDField & "='" & UserPWD & "'" ' SQL Statement mit pwd & username vervollständigen
                Else                                                ' Sonst
                    UserAnmeldename = objObjectBag.getUsername()    ' NT anmelde namen verwenden
                    szSQL = szSQL & "(" & .UsernameField & "='" & UserAnmeldename & "'"
                    If .UsernameField2 <> "" Then szSQL = szSQL & " OR " & _
                            .UsernameField2 & "='" & UserAnmeldename & "'" ' evtl. 2, Namensfeld abfragen
                    szSQL = szSQL & ")"                             ' Letzte Schliessende Klammer
                End If ' .bSingleSignOn Or Count > 0
                
                
                
                Count = Count + 1                                   ' Anmelde versuch hochzählen
                If Count > .UserLoginMaxCount Then GoTo exithandler ' Zu Viele anmelde versuche
                If bCancel Then GoTo exithandler                    ' Abbruch durch den Anwender
                Set rsUser = FillRS(szSQL, False)                   ' RS aus SQL Statement
                If rsUser Is Nothing Then                           ' Nix da
                    bOK = False                                     ' Anmeldung gescheitert
                    Call objError.ShowErrMsg("Benutzername oder Kennwort falsch.", vbCritical + vbOKOnly, "Anmeldung") ' Anwender informieren
                    frmUserLogin.txtMwgaPWD.Text = ""               ' PWD im form löschen
                Else                                                ' Sonst
                    If rsUser.RecordCount = 0 Then                  ' keine DS gefunden ?
                        bOK = False                                 ' Anmeldung gescheitert
                        If bSingleSignOnFailed Then
                            Call objError.ShowErrMsg("Benutzername oder Kennwort falsch.", _
                                    vbCritical + vbOKOnly, "Anmeldung") ' Benutzer infoemieren
                        End If
                        bSingleSignOnFailed = True
                        frmUserLogin.txtMwgaPWD.Text = ""           ' PWD im form löschen
                    Else                                            ' Sonst (DS gefunden)
                        bOK = True                                  ' Anmeldung erfolgreich
                    End If ' rsUser.RecordCount = 0
                End If ' rsUser Is Nothing
            Wend
        End If ' .bUserlogIn
    End With
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    If Not bOK Then UserAnmeldename = ""                            ' Anmelde name Leer
    ShowUserLogin = UserAnmeldename                                 ' Anmelde name zurück geben
    Err.Clear                                                       ' Evtl. error clearen
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ShowUserLogin", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function InitUserLogIn(UserLoginTable As String, _
        UsernameField As String, PWDField As String, LoginText As String, _
        UserLoginMaxCount As Integer, Optional bSingleSignOn As Boolean, _
        Optional AddWhere As String, Optional UsernameField2 As String)
' Initialisiert die zusätzlicher Benutzeranmeldung. (eigene Tabelle mit Userkennungen)
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If UserLoginTable <> "" And UsernameField <> "" And PWDField <> "" Then
        With ConParam
            .bUserlogIn = True
            .UserLoginMaxCount = UserLoginMaxCount
            .UserTable = UserLoginTable
            .LoginText = LoginText
            .UsernameField = UsernameField
            .UsernameField2 = UsernameField2
            .PWDField = PWDField
            .AddWhere = AddWhere
            .bSingleSignOn = bSingleSignOn
        End With
    End If
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "InitUserLogIn", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function TestDBConn(ingType As Integer, _
        szServername As String, _
        szDBName As String, _
        Optional szDBUser As String, _
        Optional szPWD As String, _
        Optional bNtAut As Boolean) As String
    Dim szProvider As String
    Dim szConnectString As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If ingType = 1 Then                                             ' Access DB
        szProvider = "Microsoft.Jet.OLEDB.4.0"
'            .ConnectString = "Provider=" & .Provider & ";Data Source=" & .DBName 'szProvider = "Microsoft.Jet.OLEDB.3.51"
'            "Driver={Microsoft Access Driver (*.mdb)};Dbq=nwind.mdb;" & _
                     "DefaultDir=C:\program files\devstudio\vb;Uid=Admin;Pwd=;"
        szConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
                    & szDBName & ";Mode=ReadWrite"
            
    Else                                                            ' Sonst (SQL DB)
        szProvider = "sqloledb.1"
        szConnectString = "Provider=" & szProvider & ";Data Source=" & szServername & ";" _
                & "Initial Catalog=" & szDBName & ";"                ' szProvider = "sqloledb"
        If bNtAut Then                                            ' NT autentifikation
            szConnectString = "Data Source=" & szServername & ";" _
                        & "Initial Catalog=" & szDBName & ";"
            szConnectString = szConnectString & "Integrated Security=SSPI;"
        Else                                                    ' SQL Autentifikation
            If szDBUser <> "" Then
                szConnectString = szConnectString & "User ID=" & szDBUser & ";"
            End If
            If szPWD <> "" Then
                szConnectString = szConnectString & "Password=" & szPWD & ";"
            End If
        End If
    End If
On Error Resume Next
    Set dbConAdo = New adodb.Connection                             ' Neue Verbindung öffnen
    With dbConAdo                                                   ' Verbindungs Parameter setzen
        .Mode = adModeReadWrite
        .CursorLocation = adUseClient
        .Provider = szProvider
        .ConnectionString = szConnectString
        .Open                                                       ' Verbindung öffnen
    End With
On Error GoTo Errorhandler
    If Err.Number <> 0 Then GoTo Errorhandler
TestErfolg:
    TestDBConn = GetSQLVersion()


exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr

Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "TestDBConn", errNr, errDesc)  ' Fehler behandlung aufrufen
    TestDBConn = "-1"
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function UserChangePWD(szCurUsername As String, Optional NewPWD As String)
' Änder das Benutzer kennwort
    Dim bCancel As Boolean                                          ' Abbruch duch anwender
    Dim szSQL As String                                             ' SQL Statement
    Dim fUSerPWD As frmUserChangePWD                                ' Form für PWD änderung
    Dim szOldPWD As String                                          ' Altes Kennwort
    Dim szNewPWD As String                                          ' Neues Kennwort
    Dim rsUser As adodb.Recordset                                   ' RS mit User daten
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If szCurUsername = "" Then GoTo exithandler                     ' Kein Benutzername angegeben -> fertig
    With ConParam
        If .bUserlogIn = False Then GoTo exithandler                ' Nur sinnvoll wenn Userlogin vorgesehen
        If NewPWD = "" Then                                         ' Wenn kein PWD angegeben dialog öffnen
            Set fUSerPWD = New frmUserChangePWD
            fUSerPWD.lblUsername.Caption = szCurUsername
            fUSerPWD.cmdOK.Enabled = False
            fUSerPWD.Show vbModal, fMain                            ' Form zur Kennwort eingabe anzeigen
            bCancel = fUSerPWD.bCancel
            szOldPWD = fUSerPWD.txtPWDOld
            szNewPWD = fUSerPWD.txtPWDNew2
        End If
        szSQL = "SELECT " & .UsernameField & ", " & .UsernameField2 & ", " & .PWDField & " FROM " & .UserTable & _
                 " WHERE (" & .UsernameField & "='" & szCurUsername & "' OR " & .UsernameField2 & "='" & szCurUsername & "')" & _
                 " AND " & .PWDField & "='" & szOldPWD & "'"
        Set rsUser = FillRS(szSQL, False)                           ' RS mit Userdaten holen
        If Not rsUser Is Nothing Then                               ' RS vorhanden
            If rsUser.RecordCount = 1 Then                          ' 1 DS gefunden
            
                szSQL = "UPDATE " & .UserTable & " SET " & .PWDField & "='" & szNewPWD & _
                        "' WHERE WHERE (" & .UsernameField & "='" & szCurUsername & "' OR " & _
                        .UsernameField2 & "='" & szCurUsername & "')" ' Update statement zusammenstellen
                If ExecSQL(szSQL) Then                              ' Wenn Statement erfolgreich ausgeführt
                    Call objError.ShowErrMsg("Ihr Kennwort wurde erfolgreich geändert.", _
                            vbInformation, "Kennwort änderung")     ' Benutzer informieren
                    GoTo exithandler                                ' Fertig
                End If
            End If
        End If
    End With
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "UserChangePWD", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function BeginTrans() As Boolean
' Startet eine transaktion
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If Not bConnect Then GoTo exithandler                           ' Kein Connect -> Raus
    dbConAdo.BeginTrans                                             ' Transaktion starten
    bTransaction = True                                             ' Falg setzen
    BeginTrans = True                                               ' Erfolg melden
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "BeginTrans", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function CommitTrans() As Boolean
' Schliest die laufende Transaktion ab
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If Not bConnect Then GoTo exithandler                           ' Kein Connect -> Raus
    If bTransaction Then                                            ' Wenn Transaktion läuft
        dbConAdo.CommitTrans                                        ' Transaktion Commiten
        bTransaction = False                                        ' Flag Setzen
        CommitTrans = True                                          ' erfolg melden
    End If
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "CommitTrans", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function RollbackTrans()
' Setzt die laufende Transaktion zurück
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If Not bConnect Then GoTo exithandler                           ' Kein Connect -> Raus
    If bTransaction Then                                            ' Wenn Transaktion läuft
        dbConAdo.RollbackTrans                                      ' Transaktion zurück stzen
        bTransaction = False                                        ' Flag Setzen
        RollbackTrans = True                                        ' erfolg melden
    End If
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "RollbackTrans", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function CloseConnect()
' Schliesst die Datenbank verbindung
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    dbConAdo.Close                                                  ' Verbindung schliessen
    bConnect = False                                                ' Connect flag setzen
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "CloseConnect", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function
        
Public Function Count(Tablename As String, Optional szWhere As String, Optional bSilent As Boolean) As Long
' Liefert die anzahl der Datensätz dieser Tabelle zurück
    Dim rs As New adodb.Recordset                                   ' Ergebiss RS
    Dim szSQL As String                                             ' SQL Statement
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If Not bConnect Then GoTo exithandler                           ' Kein Connect -> Raus
    If Tablename = "" Then GoTo exithandler                         ' Kein TAbellenname -> Raus
    szSQL = "SELECT Count(*) FROM " & Tablename                     ' Count Statement zusammen setzen
    If szWhere <> "" Then szSQL = szSQL & " " & szWhere             ' Evtl- Where bed anhängen
    Set rs = FillRS(szSQL, False, bSilent)                                  ' RS Füllen
    If Not rs Is Nothing Then                                       ' Wenn RS vorhanden
        Count = rs.Fields(0).Value                                  ' Anzahl auslesen
    End If
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    If bSilent Then                                                 ' Keine Fehlermeldung
        Err.Clear                                                   ' Fehler Clearen
        Resume exithandler                                          ' Weiter mit Exithandler
    End If
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "Count", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function GetSQLVersion() As String
' Ermittelt die version des SQL servers
    Dim szSQL As String                                             ' SQL Statement
    Dim rsVersion As New adodb.Recordset                            ' Ergebnis RS
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If Not bConnect Then GoTo exithandler                           ' Kein Connect -> Raus
    szSQL = "Declare @szVersion varchar(255) "                      ' SQL zum Version ermitteln
    szSQL = szSQL & "Set @szVersion =(SELECT @@VERSION) "
    szSQL = szSQL & "SELECT 'SQLVersion' = Case WHEN SUBSTRING(@szVersion,23,4)='2000' THEN " & _
            "'MS SQL Server ' + SUBSTRING(@szVersion,23,4) + ' ' + Cast(SERVERPROPERTY ('productlevel') as varchar(10))" & _
            "+ ' - ' + CAST(SERVERPROPERTY('productversion')as varchar(10)) + ' '" & _
            "+ Cast(SERVERPROPERTY ('edition') as varchar(255)) "
    szSQL = szSQL & "Else Case WHEN SUBSTRING(@szVersion,22,4)='2005' Or SUBSTRING(@szVersion,22,4)='2008' THEN " & _
            "'MS SQL Server ' + SUBSTRING(@szVersion,22,4) + ' ' + Cast(SERVERPROPERTY ('productlevel') as varchar(10))" & _
            "+ ' - ' + CAST(SERVERPROPERTY('productversion')as varchar(10)) + ' '" & _
            "+ Cast(SERVERPROPERTY ('edition') as varchar(255)) "
    szSQL = szSQL & "Else @szVersion End End"
'    szSQL = szSQL & "Set @szVersion =(SELECT @@VERSION) "
'    szSQL = szSQL & "IF SUBSTRING(@szVersion,23,4)='2000' Or SUBSTRING(@szVersion,22,4)='2005' " & _
'            "Or SUBSTRING(@szVersion,22,4)='2008' "
'    szSQL = szSQL & "SELECT  'MS SQL Server ' + SUBSTRING(@szVersion,22,4) + ' ' + Cast(SERVERPROPERTY ('productlevel') as varchar(10)) " _
'            & "+ ' - ' + CAST(SERVERPROPERTY('productversion')as varchar(10)) + ' ' + Cast(SERVERPROPERTY ('edition') as varchar(255)) , SERVERPROPERTY ('ServerName') "
'    szSQL = szSQL & "Else "
'    szSQL = szSQL & "Select @szVersion "
    Set rsVersion = FillRS(szSQL, False)                            ' RS Füllen
    If Not rsVersion Is Nothing Then                                ' Wenn RS vorhanden
        GetSQLVersion = rsVersion.Fields(0).Value                   ' Wert zürück geben
    End If
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "GetSQLVersion", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function GetValueFromSQL(szSQL As String) As Variant
' Liefert den 1. Wert der 1. DS des SQL statements zurück
    Dim rs As adodb.Recordset                                       ' Ergebnis RS
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If Not bConnect Then GoTo exithandler                           ' Wenn kein Connect dann Raus
    If szSQL = "" Then GoTo exithandler                             ' Wenn Kein SQL dann Raus
    Set rs = FillRS(szSQL, False)                                   ' RS Füllen
    If rs.RecordCount > 0 Then                                      ' Wenn Mind 1. DS
        On Error Resume Next                                        ' Hier keine Fehlerbehandlung
            GetValueFromSQL = rs.Fields(0).Value                    ' Wert holen
            If Err.Number <> 0 Then                                 ' Auf Fehler prüfen
                Err.Clear                                           ' Evtl. Error clearen
            End If
        On Error GoTo exithandler                                   ' Fehler behandlung wieder Aktivieren
    Else
        GetValueFromSQL = ""
    End If
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "GetValueFromSQL", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function FillRS(szSQL As String, Optional bUpdate As Boolean, Optional bSilent As Boolean) As adodb.Recordset
' Lifert ein RS mit daten aus SQL statement
    Dim rs As New adodb.Recordset                                   ' Ergebnis recordset
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If Not bConnect Then GoTo exithandler                           ' Wenn kein Connect dann Raus
    If szSQL = "" Then GoTo exithandler                             ' Wenn Kein SQL dann Raus
    szDetails = szSQL                                               ' SQL Statement als Details für Fehlerbehandlung
    With rs
        If bUpdate Then                                             ' Daten Update Möglich
            .LockType = adLockOptimistic                            ' Loctype auf Optimistic
        End If
        .Open szSQL, dbConAdo                                       ' RS Öffnen
    End With
    Set FillRS = rs                                                 ' RS Zurück geben
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    If bSilent Then                                                 ' Keine Fehlermeldung
        Err.Clear                                                   ' Fehler Clearen
        Resume exithandler                                          ' Weiter mit Exithandler
    End If
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "FillRS", errNr, errDesc, szDetails) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ExecSQL(szSQL As String, Optional bSilent As Boolean) As Boolean
' Führt das SQL statement aus.
' Liefert True wenn kein Fehler aufgetreten ist
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If Not bConnect Then GoTo exithandler                           ' Wenn kein Connect dann Raus
    If szSQL = "" Then GoTo exithandler                             ' Wenn Kein SQL dann Raus
    szDetails = szSQL                                               ' SQL Statement als Details für Fehlerbehandlung
'    dbConAdo.Execute szSQL, , adAsyncFetch                           ' Ausführen
    dbConAdo.Execute szSQL, , adCmdText                              ' Ausführen
    ExecSQL = True                                                  ' Erfolg zurück
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    If bSilent Then                                                 ' Keine Fehlermeldung
        Err.Clear                                                   ' Fehler Clearen
        Resume exithandler                                          ' Weiter mit Exithandler
    End If
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ExecSQL", errNr, errDesc, szDetails) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ColumnType(ColumnName As String, ByVal Tablename As String, _
        Optional bForSQL As Boolean) As String
' Liefert Datentyp und Länge des Feldes
    Dim cat As New ADOX.Catalog                                     ' Verbindung
    Dim sName As String                                             ' Typ als Zahl
    Dim szTypName As String                                         ' Datentyp name
    Dim szTypLength As String                                       ' Datentyp länge
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    Set cat.ActiveConnection = dbConAdo                             ' Set the catalogs
    sName = cat.Tables(Tablename).Columns(ColumnName).Type          ' Type ermitteln
    Select Case sName                                               ' Type übersetzen
    Case 2
        szTypName = "smallint"
        szTypLength = ""
    Case 3
        szTypName = "int"
        szTypLength = ""
    Case 5
        szTypName = "float"
        szTypLength = ""
    Case 6
        szTypName = "money"
        szTypLength = ""
    Case 7                                                          ' Access date time
        szTypName = "datetime"                                     ' ??
        szTypLength = ""
    Case 11                                                         ' Access Bool
        szTypName = "smallint"                                     ' ??
        szTypLength = ""
    Case 17
        szTypName = "tinyint"
        szTypLength = ""
    Case 131
        szTypName = "decimal"
        szTypLength = ""
    Case 129
        szTypName = "char"
        szTypLength = cat.Tables(Tablename).Columns(ColumnName).DefinedSize
    Case 135
        szTypName = "datetime"
    Case 200
        szTypName = "varchar"
        szTypLength = cat.Tables(Tablename).Columns(ColumnName).DefinedSize
    Case 201
        szTypName = "text"
    Case 202
        szTypName = "nvarchar"
        szTypLength = cat.Tables(Tablename).Columns(ColumnName).DefinedSize
    Case Else
       Stop
    End Select
    If bForSQL Then                                                 ' Ausgabe für SQL Script
        ColumnType = "[" & szTypName & "]"
    Else                                                            ' Sonst
        ColumnType = szTypName
    End If
    If szTypLength <> "" Then ColumnType = ColumnType & " (" & szTypLength & ")"
    
' if table does not exist
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ColumnType", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ConvertColumnType(szType As String, _
        Optional szSize As String, _
        Optional szFormat As String) As String
' Übersetzt Field.Type  vereinfacht ins Deutsche
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    Select Case szType                                              ' Type übersetzen
    Case 2
        ConvertColumnType = "Zahlenwert"
    Case 3
        ConvertColumnType = "Zahlenwert"
    Case 5
        ConvertColumnType = "KommaZahl"
    Case 6
        ConvertColumnType = "Währung"
    Case 7                                                          ' Access date time
        ConvertColumnType = "Datum/Uhrzeit"
    Case 11                                                         ' Access Bool
        ConvertColumnType = "Ja/Nein"
    Case 17
        ConvertColumnType = "Zahlenwert"
    Case 131
        ConvertColumnType = "Kommazahl"
    Case 129
        ConvertColumnType = "Text"
        If szSize <> "" Then
            ConvertColumnType = ConvertColumnType & " [" & szSize & " Zeichen]"
        End If
    Case 135
        ConvertColumnType = "Datum"
    Case 200
        ConvertColumnType = "Text"
        If szSize <> "" Then
            ConvertColumnType = ConvertColumnType & " [" & szSize & " Zeichen]"
        End If
    Case 201
        ConvertColumnType = "Text"
        If szSize <> "" Then
            ConvertColumnType = ConvertColumnType & " [" & szSize & " Zeichen]"
        End If
    Case 202
        ConvertColumnType = "Text"
        If szSize <> "" Then
            ConvertColumnType = ConvertColumnType & " [" & szSize & " Zeichen]"
        End If
    Case 203                                                        ' Access Memo
        ConvertColumnType = "Text"
        If szSize <> "" Then
            ConvertColumnType = ConvertColumnType & " [" & szSize & " Zeichen]"
        End If
    Case Else
       Stop
    End Select
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
Exit Function                                                       ' Funkion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ConvertColumnType", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function GetFieldListFromRS(rs As Recordset, Optional szDelemiter As String) As String
    Dim i As Integer                                                ' Counter (fields)
    Dim szFieldlist As String                                       ' Feldliste getrente liste
    Dim szFieldName As String                                       ' Name des Aktuellen Feldes
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If rs Is Nothing Then GoTo exithandler                          ' Kein Recordset -> Fertig
    If szDelemiter = "" Then szDelemiter = ";"                      ' Standart trennzeichen angeben
    For i = 0 To rs.Fields.Count - 1                                ' Alle Felder Durchlaufen
        szFieldName = rs.Fields(i).Name                             ' Feldnamen auslesen
        If InStr(szFieldName, " ") > 0 Or objSQLTools.IsReservedWord(szFieldName) _
                Then szFieldName = "[" & szFieldName & "]"          ' Wenn Feldnamen ein Leerzeichen enthält -> in []
        szFieldlist = szFieldlist & szFieldName & szDelemiter       ' Fedliste zusammensetzen
    Next i                                                          ' Nächstes Feld
    szFieldlist = objTools.Cutlastchar(szFieldlist, szDelemiter)    ' letztes Trennzeichen abschneiden
    GetFieldListFromRS = szFieldlist                                ' Ergebnis zurück
exithandler:
On Error Resume Next                                                ' Hier keine Fehler Mehr
Exit Function                                                       ' Funkion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "GetFieldListFromRS", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function GetValueListFromRS(rs As adodb.Recordset, RecordIndex As Long, _
        Optional szFieldlist As String, Optional szDelemiter As String) As String
    Dim i As Integer                                                ' Counter (fields)
    Dim n As Integer                                                ' Counter (array)
    Dim szValueList As String                                       ' Getrennte Werteliste
    Dim szFieldArray() As String                                    ' FeldArray
    Dim szYear As String
    Dim szMonth As String
    Dim szDay As String
    Dim szHour As String
    Dim szMin As String
    Dim szSec As String
    Dim sztmpVal As String
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If rs Is Nothing Then GoTo exithandler                          ' Kein Recordset -> Fertig
    If szDelemiter = "" Then szDelemiter = ";"                      ' Standart trennzeichen angeben
    rs.MoveFirst                                                    ' Erstmal zu 1. DS
    rs.Move (RecordIndex)                                           ' Dann entsprechend Weiterspringen
    If szFieldlist <> "" Then                                       ' Feldliste angegeben
        szFieldArray = Split(szFieldlist, szDelemiter)              ' Feldliste in arry splaten
        If Not objTools.CheckArray(szFieldArray) Then               ' Wenn array leer
            szFieldlist = ""                                        ' Feldliste löschen
        End If
    End If
    n = 0                                                           ' Array Counter initialisieren
    For i = 0 To rs.Fields.Count - 1                                ' Alle Felder Durchlaufen
        If szFieldlist <> "" Then
            If UCase(rs.Fields(i).Name) = UCase(szFieldArray(n)) Then
                n = n + 1
                GoTo Check
            Else
                GoTo NextField
            End If
        End If
Check:
        Select Case rs.Fields(i).Type                               ' FeldTyp auswerten
        Case 2, 3, 17                                               ' Zahlen
        'ConvertColumnType = "Zahlenwert"
            sztmpVal = objTools.CheckNull(rs.Fields(i).Value)       ' Wert auslesen
            If sztmpVal = "" Then                                   ' Wenn Leer
                szValueList = szValueList & "NULL" & szDelemiter    ' NULL einsetzen
            Else                                                    ' Sonst
                szValueList = szValueList & rs.Fields(i).Value & szDelemiter ' Einfach anhängen
            End If
            sztmpVal = ""                                           ' Tmp Value initialisieren
        Case 5, 131                                                 ' Float
            sztmpVal = objTools.CheckNull(rs.Fields(i).Value)       ' Wert auslesen
            If sztmpVal = "" Then                                   ' Wenn Leer
                szValueList = szValueList & "NULL" & szDelemiter    ' NULL einsetzen
            Else                                                    ' Sonst
                If InStr(sztmpVal, ",") Then sztmpVal = Replace(sztmpVal, ",", ".") ' , durch . erstezen
                szValueList = szValueList & sztmpVal & szDelemiter  ' An Valuelist anhängen
            End If
            sztmpVal = ""                                           ' Tmp Value initialisieren
        Case 6                                                      ' Money (wie float behandeln)
            sztmpVal = objTools.CheckNull(rs.Fields(i).Value)       ' Wert auslesen
            If sztmpVal = "" Then                                   ' Wenn Leer
                szValueList = szValueList & "NULL" & szDelemiter    ' NULL einsetzen
            Else                                                    ' Sonst
                If InStr(sztmpVal, ",") Then sztmpVal = Replace(sztmpVal, ",", ".") ' , durch . erstezen
                szValueList = szValueList & sztmpVal & szDelemiter  ' An Valuelist anhängen
            End If
            sztmpVal = ""                                           ' Tmp Value initialisieren
        Case 7                                                      ' Access date time
        
        Case 11                                                     ' Access Bool
            szValueList = szValueList & rs.Fields(i).Value & szDelemiter ' eifach anhängen
        Case 72                                                     ' GUID (wie text behandeln)
            sztmpVal = objTools.CheckNull(rs.Fields(i).Value)       ' Wert auslesen
            If sztmpVal = "" Then                                   ' Wenn Leer
                szValueList = szValueList & "NULL" & szDelemiter    ' NULL einsetzen
            Else                                                    ' Sonst
                szValueList = szValueList & "'" & Trim(rs.Fields(i).Value) & "'" & szDelemiter ' in tütelchen packen
            End If
            sztmpVal = ""                                           ' Tmp Value initialisieren
        Case 129, 200, 201, 202                                     ' text
            sztmpVal = objTools.CheckNull(rs.Fields(i).Value)       ' Wert auslesen
            If sztmpVal = "" Then                                   ' Wenn Leer
                szValueList = szValueList & "NULL" & szDelemiter    ' NULL einsetzen
            Else                                                    ' Sonst
                szValueList = szValueList & "'" & Trim(rs.Fields(i).Value) & "'" & szDelemiter ' in tütelchen packen
            End If
            sztmpVal = ""                                           ' Tmp Value initialisieren
        Case 135                                                    ' Datum
            sztmpVal = objTools.CheckNull(rs.Fields(i).Value)
            If sztmpVal <> "" Then
                szValueList = szValueList & "'" & sztmpVal & "'" & szDelemiter
'                szYear = Year(sztmpVal)
'                szMonth = Month(sztmpVal)
'                szMonth = objTools.SetPreZero(szMonth, 2)
'                szDay = Day(sztmpVal)
'                szDay = objTools.SetPreZero(szDay, 2)
'                szHour = Hour(sztmpVal)
'                szValueList = szValueList & "'" & szYear & szMonth & szDay & "'" & szDelemiter
            Else                                                    ' Sonst
                szValueList = szValueList & "NULL" & szDelemiter    ' NULL einsetzen
            End If
        Case Else
        End Select
NextField:
    Next i                                                          ' Nächstes Feld
    szValueList = objTools.Cutlastchar(szValueList, szDelemiter)    ' letztes Trennzeichen abschneiden
    GetValueListFromRS = szValueList                                ' Ergebnis zurück
exithandler:
On Error Resume Next                                                ' Hier keine Fehler Mehr
Exit Function                                                       ' Funkion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "GetValueListFromRS", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function GetColumnList(ByVal Tablename As String, Optional szDelimiter As String) As String
' Liefert eine Liste aller Spalten dieser Tabelle
    Dim lngColCount As String                                       ' Anz Spalten
    Dim szColList As String                                         ' Spalten komagetrent
    Dim cat As New ADOX.Catalog                                     ' ADO Catalog
    Dim i As Integer                                                ' counter
    Dim szErrDetails As String                                      ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If Tablename = "" Then GoTo exithandler                         ' Keine Tabelle -> Fertig
    If Not bConnect Then GoTo exithandler                           ' Wenn kein Connect dann Raus
    If szDelimiter = "" Then szDelimiter = ","                      ' Default Trennzeichen setzen
    Set cat.ActiveConnection = dbConAdo                             ' Set the catalogs
    lngColCount = cat.Tables(Tablename).Columns.Count               ' Felder durchzählen
    szErrDetails = "Tabelle: " & Tablename & vbCrLf & "Anz. Splaten: " & lngColCount ' Exaktere Fehlermeldung
    For i = 0 To lngColCount - 1                                    ' Alle Spalten durchlaufen
        szErrDetails = "Tabelle: " & Tablename & vbCrLf & "Anz. Splaten: " & lngColCount & vbCrLf _
            & "i: " & i                                             ' Exaktere Fehlermeldung
        szColList = szColList & cat.Tables(Tablename).Columns(i).Name & szDelimiter
        szErrDetails = "Tabelle: " & Tablename & vbCrLf & "Anz. Splaten: " & lngColCount & vbCrLf _
                & "i: " & i & vbCrLf & "Spalte: " & cat.Tables(Tablename).Columns(i).Name ' Exaktere Fehlermeldung
    Next i                                                          ' Nächstes Feld
    szColList = objTools.Cutlastchar(szColList, szDelimiter)
'    If Right(szColList, 1) = "," Then szColList = Left(szColList, Len(szColList) - 1)
    GetColumnList = szColList                                       ' Ergebniss zurück
exithandler:
Exit Function                                                       ' Funkion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "GetColumnList", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function TableExist(ByVal Tablename As String) As Boolean
' Liefert True wenn Tabelle existiert
   Dim sName As String                                              ' Tabellen name zur überprüfung
   Dim cat As New ADOX.Catalog                                      ' ADO Catalog
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If Not bConnect Then GoTo exithandler                           ' Wenn kein Connect dann Raus
    If Tablename = "" Then GoTo exithandler                         ' Keine Tabelle -> Fertig
    Set cat.ActiveConnection = dbConAdo                             ' Set the catalogs
    On Error Resume Next
    Err.Clear                                                       ' Clear Errors
    sName = cat.Tables(Tablename).Name                              ' produce an error,
' if table does not exist
    TableExist = (Err = 0)                                          ' check on Error = 0
exithandler:
On Error Resume Next                                                ' Hier keine Fehler Mehr
Exit Function                                                       ' Funkion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "TableExist", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function TableHasPK(szTabname As String) As Boolean
' Gibt true zurück wenn Taelle über einen Primary Key Verfügt
    Dim cat As New ADOX.Catalog                                     ' ADO Catalog
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If szTabname = "" Then GoTo exithandler                         ' Keine Tabelle -> Fertig
    If Not bConnect Then GoTo exithandler                           ' Wenn kein Connect dann Raus
    Set cat.ActiveConnection = dbConAdo                             ' Set the catalogs
    If TableExist(szTabname) Then                                   ' Wenn TAbelle exisistiert
        If cat.Tables(szTabname).Keys.Count > 0 Then                 'Schlüssel vorhanden
            TableHasPK = True                                       ' Tabelle Halt Primary Key
        End If
    End If
exithandler:
On Error Resume Next                                                ' Hier keine Fehler Mehr
Exit Function                                                       ' Funkion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    TableHasPK = False                                              ' Misserfolg zurück
    Call objError.Errorhandler(MODULNAME, "TableHasPK", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function GetPKFieldList(szTabname As String, Optional szDelimiter As String) As String
    Dim cat As New ADOX.Catalog                                     ' ADO Catalog
    Dim i As Integer                                                ' Counter für indizes
    Dim x As Integer                                                ' Counter für PK Fields
    Dim szPKFieldlist As String                                     ' Feldliste des Primäschlüssels
    Dim szPkIndexList As String                                     ' Liste der PK Indexe
    Dim szPkIndexArray() As String                                  ' Array mit PK
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If Not TableHasPK(szTabname) Then GoTo exithandler              ' Keine PK Fields ->Fertig
    If szDelimiter = "" Then szDelimiter = ","                      ' Trennzeichen setzten falls nicht angegeben
    szPkIndexList = GetIndexList(szTabname, True)                   ' Liste der PKs serstellen
    If szPkIndexList = "" Then GoTo exithandler                     ' Liste leer -> fertig
    szPkIndexArray = Split(szPkIndexList, ",")                      ' Liste in Array
    If Not objTools.CheckArray(szPkIndexArray) Then GoTo exithandler
    Set cat.ActiveConnection = dbConAdo                             ' Set the catalogs
    For i = 0 To UBound(szPkIndexArray)                             ' Alle PKs duchgehen
        For x = 0 To cat.Tables(szTabname).Indexes(szPkIndexArray(i)).Columns.Count - 1
            szPKFieldlist = szPKFieldlist & cat.Tables(szTabname).Indexes(szPkIndexArray(i)).Columns(x).Name & szDelimiter
        Next x                                                      ' Nächstes Feld
    Next i                                                          ' Nächster Primary Key
    szPKFieldlist = objTools.Cutlastchar(szPKFieldlist, szDelimiter) ' Abschneiden
    GetPKFieldList = szPKFieldlist                                  ' Liste zurück
exithandler:
On Error Resume Next                                                ' Hier keine Fehler Mehr
Exit Function                                                       ' Funkion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "TableHasPK", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function GetTableList(Optional bAllTabs As Boolean, Optional szDelimiter As String, Optional bViews As Boolean) As String
' Liefert eine Liste aller Tabellen
    Dim lngTabCount As String                                       ' Anz Tabellen
    Dim szTabList As String                                         ' Tabellen Liste (; getrennt)
    Dim cat As New ADOX.Catalog                                     ' ADO Catalog
    Dim i As Integer                                                ' counter
    Dim x As Integer                                                ' Array Counter
    Dim szErrDetails As String                                      ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If szDelimiter = "" Then szDelimiter = ";"                      ' Defaulttrennzeichen setzen
    Set cat.ActiveConnection = dbConAdo                             ' Set the catalogs
    lngTabCount = cat.Tables.Count                                  ' Tabellen Zählen
    szErrDetails = "Anz. Tabellen: " & lngTabCount                  ' Exaktere Fehlermeldung
    For i = 0 To lngTabCount - 1                                    ' Alle TAb durchlaufen
        szErrDetails = "Anz. Tabellen: " & lngTabCount & vbCrLf _
            & "i: " & i                                             ' Exaktere Fehlermeldung
        If Not bAllTabs And cat.Tables(i).Type = "TABLE" Then
            szTabList = szTabList & cat.Tables(i).Name & ";"        ' Liste zusammensetzen
            szErrDetails = "Anz. Tabellen: " & lngTabCount & vbCrLf _
                & "i: " & i & vbCrLf & "Tabelle: " & cat.Tables(i).Name ' Exaktere Fehlermeldung
        ElseIf bViews And cat.Tables(i).Type = "VIEW" Then
            szTabList = szTabList & cat.Tables(i).Name & ";"        ' Liste zusammensetzen
            szErrDetails = "Anz. Tabellen: " & lngTabCount & vbCrLf _
                & "i: " & i & vbCrLf & "Sicht: " & cat.Tables(i).Name ' Exaktere Fehlermeldung
        End If
'        If Not bAllTabs And Not (cat.Tables(i).Type = "SYSTEM TABLE" Or _
'                    cat.Tables(i).Type = "SYSTEM VIEW") Then
'            szTabList = szTabList & cat.Tables(i).Name & ";"        ' Liste zusammensetzen
'                                                                    ' Exaktere Fehlermeldung
'            szErrDetails = "Anz. Tabellen: " & lngTabCount & vbCrLf _
'                & "i: " & i & vbCrLf & "Tabelle: " & cat.Tables(i).Name
'        End If
    Next i                                                          ' Nächste Tabelle
    

    szTabList = objTools.Cutlastchar(szTabList, ";")                ' Letztes ; abschneiden
    GetTableList = szTabList                                        ' Ergebniss zurück
exithandler:
Exit Function                                                       ' Funkion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "GetTableList", errNr, szErrDetails) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ColumnExist(ColumnName As String, ByVal Tablename As String) As Boolean
' Liefert True wenn Spalte existiert
    Dim sName As String
    Dim cat As New ADOX.Catalog                                     ' ADO Catalog
    Dim szErrDetails As String                                      ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    szErrDetails = "Table: " & Tablename & " Column: " & ColumnName ' Details für Fehlerbehandlung festlegen
    Set cat.ActiveConnection = dbConAdo                             ' Set the catalogs
    On Error Resume Next                                            ' Fehler behandlung deaktivieren
    Err.Clear                                                       ' Clear Errors
    sName = cat.Tables(Tablename).Columns(ColumnName).Name          ' produce an error,
    ColumnExist = (Err = 0)                                         ' check on Error = 0
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ColumnExist", errNr, szErrDetails) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function IndexExist(IndexName As String, ByVal Tablename As String) As Boolean
' Liefert True wenn die Tabelle existiert
    Dim sName As String
    Dim cat As New ADOX.Catalog                                     ' ADO Catalog
    Dim szErrDetails As String                                      ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    szErrDetails = "Table: " & Tablename & " Index: " & IndexName   ' Details für Fehlerbehandlung festlegen
    Set cat.ActiveConnection = dbConAdo                             ' Set the catalogs
    On Error Resume Next
    Err.Clear                                                       ' Clear Errors
    sName = cat.Tables(Tablename).Indexes(IndexName).Name           ' produce an error,
    IndexExist = (Err = 0)                                          ' check on Error = 0
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "IndexExist", errNr, szErrDetails) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function GetIndexList(Tablename As String, Optional bPKonly As Boolean) As String
' Liefert eine Liste aller indizes dieser Tabelle
    Dim lngIndCount As String                                       ' Anz Indexe
    Dim szIndList As String                                         ' Indexe komagetrent
    Dim cat As New ADOX.Catalog                                     ' ADO Catalog
    Dim i As Integer                                                ' counter
    Dim szErrDetails As String                                      ' Fehler Details
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    Set cat.ActiveConnection = dbConAdo                             ' Set the catalogs
    lngIndCount = cat.Tables(Tablename).Indexes.Count               ' Anz. Indizes ermitteln
    szErrDetails = "Tabelle: " & Tablename & vbCrLf & "Anz. Indizes: " & lngIndCount
    For i = 0 To lngIndCount - 1                                    ' Alle Indexe durchlaufen
        szErrDetails = "Tabelle: " & Tablename & vbCrLf & "Anz. Splaten: " & lngIndCount & vbCrLf _
            & "i: " & i                                             ' Exaktere Fehlermeldung
        If bPKonly Then                                             ' Wenn nur PKs
            If cat.Tables(Tablename).Indexes(i).PrimaryKey Then
                 szIndList = szIndList & cat.Tables(Tablename).Indexes(i).Name & ","
            End If
        Else                                                        ' Sonst
            szIndList = szIndList & cat.Tables(Tablename).Indexes(i).Name & ","
            szErrDetails = "Tabelle: " & Tablename & vbCrLf & "Anz. Indizes: " & lngIndCount & vbCrLf _
                & "i: " & i & vbCrLf & "Indize: " & cat.Tables(Tablename).Indexes(i).Name
        End If
    Next i
    If Right(szIndList, 1) = "," Then szIndList = Left(szIndList, Len(szIndList) - 1)
    GetIndexList = szIndList                                        ' Ergebnis zurück
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "GetIndexList", errNr, szErrDetails) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function GetIndexFieldList(Tablename As String, IndexName As String) As String
' Liefert eine Liste aller Felder des indizes
    Dim cat As New ADOX.Catalog                                     ' ADO Catalog
    Dim lngIndFieldCount  As Integer                                ' Anzahl der index Felder
    Dim szIndFieldList As String                                    ' Feldliste
    Dim i As Integer                                                ' counter
    Dim szErrDetails As String                                      ' Fehler details
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    szErrDetails = "Tabelle: " & Tablename & vbCrLf & " Index: " & IndexName
    Set cat.ActiveConnection = dbConAdo                             ' Set the catalogs
    lngIndFieldCount = cat.Tables(Tablename).Indexes(IndexName).Columns.Count
    For i = 0 To lngIndFieldCount - 1                               ' Alle Indexfelder durchlaufen
        szIndFieldList = szIndFieldList & cat.Tables(Tablename).Indexes(IndexName).Columns(i).Name & ","
        szErrDetails = "Tabelle: " & Tablename & vbCrLf & " Index: " _
                        & IndexName & " Feld: " & szIndFieldList
    Next i                                                          ' Nächstes Feld
    If Right(szIndFieldList, 1) = "," Then szIndFieldList = Left(szIndFieldList, Len(szIndFieldList) - 1)
    GetIndexFieldList = szIndFieldList                              ' Ergebnis zurück
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "GetIndexFieldList", errNr, szErrDetails) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function IndexIsUnique(Tablename As String, IndexName As String) As Boolean
' Liefert True wenn der angegebene Index ein Unique index ist
    Dim cat As New ADOX.Catalog                                     ' ADO Catalog
    Dim bUnique  As Integer                                         ' Index ist Unique
    Dim szErrDetails As String                                      ' Fehler Details
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    szErrDetails = "Tabelle: " & Tablename & vbCrLf & " Index: " & IndexName
    Set cat.ActiveConnection = dbConAdo                             ' Set the catalogs
    bUnique = cat.Tables(Tablename).Indexes(IndexName).Unique       ' Auf Unique prüfen
    IndexIsUnique = bUnique                                         ' Ergebniss zurück liefern
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "IndexIsUnique", errNr, szErrDetails) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function SaveConnection() As Boolean
' Letze Connection Speichern ( Wichtig!! PWD Verschlüsseln!)
    Dim szConName As String                                         ' Verbindungswert                                              ' Counter
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If Not SaveLastConnection(szConName) Then GoTo exithandler      ' Akt. Verbindung als Last Connection speichern
    If Not SaveConnectionList(szConName) Then GoTo exithandler
    SaveConnection = True                                           ' Erfolg zurück
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "SaveConnection", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ReadConnectionFromList(Index As Integer, _
        Optional bNoSQL As Boolean, _
        Optional bNoAccess As Boolean) As String
    Dim szRegKey As String                                          ' Registry Schlüssel
    Dim szTmpValue As String                                        ' Evtl gefundener wert
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    szRegKey = "SOFTWARE\" & objObjectBag.GetAppRegRoot & "\" & REG_KEY_CON ' RegKey festlegen
    szTmpValue = objRegTools.ReadRegValue("HKCU", szRegKey, CStr(Index)) ' und auslesen
    If szTmpValue <> "" Then                                        ' Wert gefunden
        If InStr(szTmpValue, szRegTrenn) = 0 Then                   ' Wenn kein ; darin dann ists wohl verschlüsselt
            szTmpValue = objTools.Crypt(szTmpValue, False)          ' Entschlüsseln
        End If
        If Left(szTmpValue, 2) = TYP_ACC & szRegTrenn _
                And bNoAccess Then GoTo exithandler                 ' Verb. Ist Access aber Access nicht gewünscht -> Ferig
        If Left(szTmpValue, 2) = TYP_SQL & szRegTrenn _
                And bNoSQL Then GoTo exithandler                    ' Verb. Ist SQL aber SQL nicht gewünscht -> Ferig
        ReadConnectionFromList = szTmpValue                         ' Wert zurück geben
    End If
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ReadConnectionFromList", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function FindReadableConInReg(szReadebleConnect As String) As String
    Dim szTmpValue As String                                        ' tmp Reg Value
    Dim i As Integer                                                ' Counter
    Dim szRegKey As String                                          ' Regschlüssel
    Dim szDBName As String
    Dim szServername As String
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If szReadebleConnect = "" Then GoTo exithandler                 ' Kein ReadebleConnect -> Fertig
    If InStr(szReadebleConnect, " auf ") > 0 Then
        szDBName = Left(szReadebleConnect, InStr(szReadebleConnect, " auf ") - 1)
        szServername = Replace(szReadebleConnect, szDBName & " auf ", "")
    End If
    szRegKey = "SOFTWARE\" & objObjectBag.GetAppRegRoot & "\" & REG_KEY_CON
    szTmpValue = "@"                                                ' Abbruch bed umgehen
    i = 0                                                           ' Counter initialiseren
    While szTmpValue = "@"                                          ' Abbruch bed. prüfen
        szTmpValue = objRegTools.ReadRegValue("HKCU", szRegKey, CStr(i)) 'Reg wert auslesen
        If szTmpValue <> "" Then                                    ' Wert gefunden
            If InStr(UCase(szTmpValue), UCase(szRegTrenn & szServername & szRegTrenn & szDBName & szRegTrenn)) > 0 Then
                FindReadableConInReg = szTmpValue                   ' Gefunden
                szTmpValue = ""
            Else
                szTmpValue = "@"
            End If
                'If UCase(szTmpValue) <> UCase(szConName) Then       ' Connection schon gespeichert
                '    ReDim Preserve szConArray(i + 1)                ' Array erweitern
                '    szConArray(i + 1) = szTmpValue                  ' Wert eintragen
                'End If
        End If
        i = i + 1                                                   ' Counter hochzählen
    Wend                                                            ' Weiter
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "FindReadableConInReg", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function GetReadableConnect(szRegConParams As String) As String
' Übersetzt die internen Verbindungsfolgen in etwas lesbares
    Dim szParamArry() As String                                     ' Array mit verbindungsparametern
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If szRegConParams = "" Then GoTo exithandler                    ' Parameter Zeichefolge Leer -> Fertig
    If InStr(szRegConParams, szRegTrenn) = 0 Then                   ' Wenn Kein ; enthalten -> Verschlüsselt
        szRegConParams = objTools.Crypt(szRegConParams, False)      ' Entschlüsseln
    End If
    szParamArry = Split(szRegConParams, szRegTrenn)                 ' Parameter aufspalten in Array
    If objTools.CheckArray(szParamArry) Then                        ' Array Überprüfen
        If szParamArry(0) = CStr(TYP_ACC) Then                      ' Access Verbindung
        
        ElseIf szParamArry(0) = CStr(TYP_SQL) Then                  ' MS SQL Server Verbindung
            GetReadableConnect = szParamArry(2) & " auf " & szParamArry(1)
        End If
    End If
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "GetReadableConnect", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function
                                                                    ' *****************************************
                                                                    ' Hilfsfunktionen
Private Sub OpenConnectParameterForm(Optional Cancel As Boolean)
    Dim f As New frmDBConnProperties
    Dim szConn As String
    Dim szConnArray() As String
On Error GoTo Errorhandler
     If ConParam.SQLServer = "" Or ConParam.DBName = "" Then
        szConn = objOptions.GetOptionByName("LastConnection")
        If Trim(szConn) <> "" Then
            If InStr(szConn, szRegTrenn) = 0 Then                   ' Kein ; in ConnectSTring dann ist dieser Verschlüsselt
                szConn = objTools.Crypt(szConn, False)              ' Entschlüsseln
            End If
            szConnArray = Split(szConn, szRegTrenn)                 ' Connect String in Array aufspalten
            If UBound(szConnArray) >= 5 Then
                If szConnArray(0) = 1 Then
                    ConParam.bSQL = False
                Else
                    ConParam.bSQL = True
                End If
                ConParam.SQLServer = szConnArray(1)
                ConParam.DBName = szConnArray(2)
                ConParam.DBUser = szConnArray(3)
                ConParam.DBPWD = szConnArray(4)
'                ConParam.DBPWD = objTools.Crypt(szConnArray(4), False)
                ConParam.bNt = CBool(szConnArray(5))
            End If
        End If
    End If
    
    If Not ConParam.bSQL Then  ' Access
        f.OptAcc.Value = True
        f.optMSSQL.Value = False
    Else    ' SQL Server
        f.OptAcc.Value = False
        f.optMSSQL.Value = True
    End If
    
    If Not (bAccessPossible And bMSSQLSrvPossible) Then
        f.OptAcc.Enabled = False
        f.optMSSQL.Enabled = False
        f.OptAcc.Visible = False
        f.optMSSQL.Visible = False
    Else
        f.OptAcc.Visible = True
        f.optMSSQL.Visible = True
        f.OptAcc.Enabled = True
        f.optMSSQL.Enabled = True
    End If
    With f
        Call .InitObjectBag(objObjectBag)
        .txtServer.Text = ConParam.SQLServer
        .txtDBName.Text = ConParam.DBName
        .txtDBUser.Text = ConParam.DBUser
        .txtPWD.Text = ConParam.DBPWD
        If ConParam.bNt Then
            .chkNT.Value = 1
        Else
            .chkNT.Value = 0
        End If
        Call objObjectBag.CheckFormStyle(f)
        .Show 1
        Cancel = .bCancel
    End With
    'Stop
               
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
Exit Sub                                                            ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "OpenConnectParameterForm", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Sub

Public Function SaveLastConnection(ByRef szConName As String) As Boolean
' Speichert die Aktuell DB Verbindung als Letzte Verbindung
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    With ConParam
        .CryptPWD = objTools.Crypt(.DBPWD, True)                    ' DB PWD Verschlüsseln
        If .bSQL Then                                               ' Ist SQL Server
            szConName = "2" & szRegTrenn & .SQLServer & szRegTrenn & .DBName & szRegTrenn & .DBUser _
                    & szRegTrenn & .CryptPWD & szRegTrenn & .bNt    ' Verbindungsparameter als Zeichenfolge speichern
        Else                                                        ' Sonst (Access)
            szConName = "1" & szRegTrenn & .SQLServer & szRegTrenn & .DBName & szRegTrenn & .DBUser _
                    & szRegTrenn & .CryptPWD & szRegTrenn & .bNt    ' Verbindungsparameter als Zeichenfolge speichern
        End If
        ' Jetzt das ganze nochmal für die Option Last Connection mit erstmal unverschlüsseltem PWD
        If .bSQL Then                                               ' Ist SQL Server
            szConName = "2" & szRegTrenn & .SQLServer & szRegTrenn & .DBName & szRegTrenn & .DBUser _
                    & szRegTrenn & .DBPWD & szRegTrenn & .bNt       ' Verbindungsparameter als Zeichenfolge speichern
        Else                                                        ' Sonst (Access)
            szConName = "1" & szRegTrenn & .SQLServer & szRegTrenn & .DBName & szRegTrenn & .DBUser _
                    & szRegTrenn & .DBPWD & szRegTrenn & .bNt       ' Verbindungsparameter als Zeichenfolge speichern
        End If
        szConName = objTools.Crypt(szConName, True)                 ' Und das genze geschleuder verschlüsseln
        Call objOptions.SetOptionByName("LastConnection", szConName) ' In Optionen Speichern
     End With
     SaveLastConnection = True                                      ' Erfolg zurück
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "SaveLastConnection", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function SaveConnectionList(szConName As String) As Boolean
    Dim szRegKey As String                                          ' Reg schlüssel
    Dim i As Integer                                                ' Counter
    Dim szTmpValue As String                                        ' AKt Reg Value
    Dim szConArray() As String                                      ' Array mit Verbindungen
    Dim bFound As Boolean                                           ' True wenn szConName schon in Reg liste
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    szTmpValue = "@"                                                ' Abbruch bedingung umgehen
    i = 0                                                           ' Array Zähler initialisieren
    While szTmpValue <> ""                                          ' Abbruch bedingung
        szRegKey = "SOFTWARE\" & objObjectBag.GetAppRegRoot & "\" & REG_KEY_CON ' Reg SChlüssel
        szTmpValue = objRegTools.ReadRegValue("HKCU", szRegKey, CStr(i)) ' Reg wert auslesen
        If szTmpValue <> "" Then                                    ' Wert gefunden
            If UCase(szTmpValue) <> UCase(szConName) Then           ' Connection schon gespeichert
                ReDim Preserve szConArray(i)                        ' Array erweitern
                szConArray(i) = szTmpValue                          ' Wert eintragen
            Else                                                    ' Sonst
                bFound = True                                       ' Gefunden
            End If
        End If
        i = i + 1                                                   ' Counter hochzählen
    Wend                                                            ' Weiter bis Abbrch
    Call objRegTools.DelRegKey("HKCU", szRegKey)                    ' Reg schlüssel löschen
    Call objRegTools.writeRegValue("HKCU", szRegKey, "0", szConName) ' Akt Verbindung auf 1 in reg Speichern
    If objTools.CheckArray(szConArray) Then                         ' Prüfen ob Array werte enthält
        For i = 0 To UBound(szConArray)                             ' Gesamtest Array Duchlaufen
            If i > lngMaxRegCon Then Exit For                       ' Nur 20 verbindungen spichern
            If szConArray(i) <> "" Then                             ' Wenn Array Wert enthält
                Call objRegTools.writeRegValue("HKCU", szRegKey, CStr(i + 1), szConArray(i)) ' Akt Verbindung mit lfdnr speichern
            End If
        Next i                                                      ' Nächstes Array Item
    End If
    SaveConnectionList = True                                       ' Erfolg zurück
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "SaveConnectionList", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function GetConParamsFromRegString(szConnRegStr As String, _
        lngConTyp As Integer, _
        szServer As String, _
        szDBName As String, _
        szUser As String, _
        szPWD As String, _
        bNtAut As Boolean) As Boolean
    Dim szConnArray() As String                                     ' Array mit verbindungsparametern
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If szConnRegStr = "" Then GoTo exithandler
    If InStr(szConnRegStr, szRegTrenn) = 0 Then                     ' Kein ; enthalten -> Verschlüsselt
        szConnRegStr = objTools.Crypt(szConnRegStr, False)          ' Entschlüsseln
    End If
    szConnArray = Split(szConnRegStr, szRegTrenn)                   ' Connect String in Array Aufspalten
    If Not objTools.CheckArray(szConnArray) Then GoTo exithandler
    If UBound(szConnArray) > 0 Then lngConTyp = CLng(szConnArray(0))
    If UBound(szConnArray) > 1 Then szServer = szConnArray(1)
    If UBound(szConnArray) > 2 Then szDBName = szConnArray(2)
    If UBound(szConnArray) > 3 Then szUser = szConnArray(3)
    If UBound(szConnArray) > 4 Then szPWD = szConnArray(4)
    If UBound(szConnArray) > 5 Then bNtAut = CBool(szConnArray(5))
    GetConParamsFromRegString = True                                ' Erfolg zurück
exithandler:
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "GetConParamsFromRegString", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function
