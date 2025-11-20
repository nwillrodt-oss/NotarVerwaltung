VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMDIMain 
   Appearance      =   0  '2D
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6000
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10665
   Icon            =   "frmMDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin MSComctlLib.ImageList ILMainMenu 
      Left            =   1440
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBarMain 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5745
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Picture         =   "frmMDIMain.frx":0CCA
            Key             =   "PanelDate"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Picture         =   "frmMDIMain.frx":1264
            Key             =   "PanelTime"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMainDatei 
      Caption         =   "&Datei"
      NegotiatePosition=   1  'Links
      Begin VB.Menu mnuMainDateiDBselect 
         Caption         =   "&Datenbank auswählen"
      End
      Begin VB.Menu mnuMainDateiExit 
         Caption         =   "&Beenden"
      End
   End
   Begin VB.Menu mnuMainEdit 
      Caption         =   "&Bearbeiten"
   End
   Begin VB.Menu mnuMainView 
      Caption         =   "&Ansicht"
      Begin VB.Menu mnuMainViewList 
         Caption         =   "&Liste"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuMainViewDetails 
         Caption         =   "&Details"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuMainExtras 
      Caption         =   "&Extras"
      Begin VB.Menu mnuMainExtrasOptions 
         Caption         =   "&Optionen"
      End
   End
   Begin VB.Menu mnuMainWindow 
      Caption         =   "&Fenster"
      WindowList      =   -1  'True
      Begin VB.Menu mnuMainWindowHSplit 
         Caption         =   "&Untereinander"
      End
      Begin VB.Menu mnuMainWindowVSplit 
         Caption         =   "&Nebeneinander"
      End
      Begin VB.Menu mnuMainWindowCascade 
         Caption         =   "Ü&berlappend"
      End
   End
   Begin VB.Menu mnuMainInfo 
      Caption         =   "&?"
      Begin VB.Menu mnuMainInfoHelp 
         Caption         =   "&Hilfe"
      End
      Begin VB.Menu mnuMainInfoAbout 
         Caption         =   "&Info"
      End
      Begin VB.Menu mnuMainInfoBug 
         Caption         =   "&Fehler Melden"
      End
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULNAME = "frmMDIMain"

Private Sub MDIForm_Load()

On Error GoTo Errorhandler

Dim objError As Object
    
    Set objError = objObjectBag.GetErrorObj()                       ' Error Objekt holen
    
    Call InitMDIForm                                                ' Main (MDI) Form initialisieren
    
exithandler:

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "MDIForm_Load", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
    
    Call AppExit                                                    ' Application beenden
End Sub

Private Function InitMDIForm()
    
    Dim szSize As String                                            ' SizeValue aus Optionen
    Dim lngState As Integer                                         ' State Value aus Optionen
    Dim szSizeArray() As String                                     ' (0) = Width , (1) = Height
    Dim szLastCon As String                                         ' LastConnection Value aus Reg
    Dim szConnArray() As String                                     ' (0)=Servername, (1)=DBName, (1)=DBUser, (4)=PWD
    Dim bCancel As Boolean
    Dim bOK As Boolean
    
On Error GoTo Errorhandler

    Me.Caption = SZ_APPTITLE                                        ' Form Caption Setzten
    
    lngState = CLng(objTools.checkNull(objOptions.GetOptionByName(OPTION_MAINSTATE), 0)) ' Option WindowState auslesen
    If lngState >= 0 And lngState <= 2 Then                         ' 0=normal 1=min 2=max
        Me.WindowState = lngState                                   ' State setzen
    End If
    
    szSize = objOptions.GetOptionByName(OPTION_MAINSIZE)            ' Option WindowSize auslesen
    If szSize <> "" Then
        szSizeArray = Split(szSize, "/")                            ' Value aufspliten
        Me.Width = szSizeArray(0)                                   ' (0) = Width
        Me.Height = szSizeArray(1)                                  ' (1) = Height
    End If
    
    Call InitStatusBarMain                                          ' Statusbar initialisieren
    
'    Call objObjectBag.InitDBConnection(False, True)                 ' Acc nein ; SQl ja
'    Set objDBconn = objObjectBag.GetDBConObj()
            
    If bAutoConnect Or _
            objOptions.GetOptionByName(OPTION_AUTOCON) Then         ' Automatische anmeldung
        Call OpenNewDB                                              ' DB Form öffnen
    End If
' connection im DB form reslisieren !!!!!!

'    If bAutoConnect Or _
'            objOptions.GetOptionByName(OPTION_AUTOCON) Then         ' Automatische anmeldung
'        szLastCon = objOptions.GetOptionByName(OPTION_LASTCON)      ' Option LastConnection auslesen
'        If szLastCon <> "" Then
'            szConnArray = Split(szLastCon, ";")                     ' Value aufspliten
'            If UBound(szConnArray) <= 5 Then GoTo manConnection
'            If objDBconn.GetADODBConn(CLng(szConnArray(0)), _
'                    szConnArray(1), szConnArray(2), szConnArray(3), _
'                    szConnArray(4), CBool(szConnArray(5)), , bCancel) Then
'                objError.WriteProt (PROT_DB_AUTOCON)
'                bOK = True
'                Call OpenDBForm(objDBconn)                          ' DB Form anzeigen
'            Else                                                    ' keine Verbindung
'                GoTo manConnection
'            End If
'        Else                                                        ' Keine letzte verbindung gespeichert
'           GoTo manConnection
'        End If
'    Else                                                            ' Kein AutoConnect
'        GoTo manConnection
'    End If
    
'Test Mega Komponenten {
            'Dim mServer As cMegaDbServer
            'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
'Dim MTDB As TDB
'Dim MTConnectParams As TConnectParams
'            With MTConnectParams
'                .mvarAnmeldename = "mega"
'                .mvarDbName = "dbolg"
'                .mvarNTSicherheit = 0
'                .mvarPasswortEntschluesselt = "mega"
'                .mvarServername = "OLG-SRV6"
'            End With
'            Call mServer.Anmeldung(MTDB, MTConnectParams) 'Klappt nicht

'            Set mServer = New cMegaDbServer
'            Call mServer.DBVerbinden("MegaTest")

            
'Test Mega Komponenten }


exithandler:
On Error Resume Next
    Call ShowSplash(False)
    'Set objDBconn = Nothing
Exit Function

manConnection:
    ' manuel versuchen
    'GetADODBConn(ingType As Integer, _
        szServername As String, _
        szDBName As String, _
        Optional szDBUser As String, _
        Optional szPWD As String, _
        Optional bNtAut As Boolean, Optional bCancel As Boolean)
    While Not bOK
        'bOK = objDBconn.GetADODBConn(2, "", "", "", "", , bCancel)
        'If bOK Then Call OpenDBForm(objDBconn)                      ' DB Form anzeigen
        'If bCancel Then bOK = True
    Wend
    If bCancel Then Call AppExit
Exit Function

Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitMDIForm", errNr, errDesc)
    Resume exithandler
End Function

Private Function InitStatusBarMain()
' Initialisiert den Statusbar des Mainforms

On Error GoTo Errorhandler
    
    StatusBarMain.Panels(1).Alignment = sbrCenter                   ' Datum
    StatusBarMain.Panels(1).Text = Left(CStr(Now()), 10)
    
    StatusBarMain.Panels(2).Alignment = sbrLeft                     ' Angemelderter User
    StatusBarMain.Panels(2).Text = objObjectBag.GetUserName
    
    StatusBarMain.Panels(3).Alignment = sbrLeft                     ' Admin
    If objObjectBag.bUserIsAdmin Then
        StatusBarMain.Panels(3).Text = "(Admin)"
    Else
        StatusBarMain.Panels(3).Text = "(Benutzer)"
    End If
    
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitMDIForm", errNr, errDesc)
    Resume exithandler
End Function

Public Function OpenNewDB()                                         ' Public kann aus DB form aufgerufen werden

    Dim newDBCon As Object                                          ' Neues DB Connection Object
    Dim bOK As Boolean
    Dim szLastCon As String                                         ' verbindungsparameter als String
    Dim szConnArray() As String                                     ' verbindungsparameter als Array
    Dim bCancel As Boolean                                          ' Benutzerabbruch
    
On Error GoTo Errorhandler
    
    Me.MousePointer = vbHourglass                                   ' Sanduhr
    Call objObjectBag.InitDBConnection(False, True)                 ' Acc nein ; SQl ja
    Set newDBCon = objObjectBag.GetDBConObj()                       ' connection Object Holen
    
    If objTools.ArrayCount(DBFormArray) > -1 Then                   ' Wenn schon DB Form offen
        GoTo manConnection                                          ' Dann manuell Verbinden
    End If
    
    If bAutoConnect Or _
            objOptions.GetOptionByName(OPTION_AUTOCON) Then         ' Automatische anmeldung ?
        szLastCon = objOptions.GetOptionByName(OPTION_LASTCON)      ' Option LastConnection auslesen
        objError.WriteProt (PROT_DB_AUTOCON)                        ' Autoconnect Protokolieren
        If szLastCon <> "" Then
            szConnArray = Split(szLastCon, ";")                     ' Value aufspliten
            'If UBound(szConnArray) <= 5 Then GoTo manConnection
            bOK = newDBCon.GetADODBConn(CLng(szConnArray(0)), _
                    szConnArray(1), szConnArray(2), szConnArray(3), _
                    szConnArray(4), CBool(szConnArray(5)), bCancel) ' Verbinden
        End If
    Else
        GoTo manConnection                                          ' Manuell Verbinden
        'bOK = newDBCon.GetADODBConn(2, "", "", "", "")               ' Wenn Anmelden erfolgreich
    End If


        If bOK Then
        objError.WriteProt (PROT_DB_CON & newDBCon.GetServername & ";" _
                & newDBCon.GetDBname & ";" & newDBCon.GetDBUsername) ' Erfolgreiche anmeldung Protokolieren
        Call OpenDBForm(newDBCon)                                   ' DB form öffnen
    Else
        objError.WriteProt (PROT_DB_CONFAIL & newDBCon.GetServername & ";" _
                & newDBCon.GetDBname & ";" & newDBCon.GetDBUsername) ' Fehlerhafte anmeldung Protokolieren
    End If
    
exithandler:
    Me.MousePointer = vbDefault                                     ' Mauszeiger wieder normal
    
Exit Function
manConnection:                                                      ' manuel versuchen
    While Not bOK
        bOK = newDBCon.GetADODBConn(2, "", "", "", "", , bCancel)   ' Verbinden
        If bOK Then
            objError.WriteProt (PROT_DB_CON & newDBCon.GetServername & ";" _
                & newDBCon.GetDBname & ";" & newDBCon.GetDBUsername) ' Erfolgreiche anmeldung Protokolieren
            Call OpenDBForm(newDBCon)                               ' DB Form anzeigen
        Else
            objError.WriteProt (PROT_DB_CONFAIL & newDBCon.GetServername & ";" _
                & newDBCon.GetDBname & ";" & newDBCon.GetDBUsername) ' Fehlerhafte anmeldung Protokolieren
        End If
        If bCancel Then bOK = True
    Wend
    If bCancel Then Call AppExit                                    ' Bei Benutzerabbruch -> Anwendung beenden
    GoTo exithandler
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "OpenNewDB", errNr, errDesc)
    Resume exithandler
End Function

Private Sub mnuMainDateiDBselect_Click()
    Call OpenNewDB                                                  ' Neue DB öffnen
End Sub

Private Sub mnuMainDateiExit_Click()
    Call AppExit                                                    ' Anwendung beenden
End Sub

Private Sub mnuMainExtrasOptions_Click()
    Call ShowOptions                                                ' Options Dialog anzeigen
End Sub

Private Sub mnuMainInfoAbout_Click()
    'Call ShowAbout                                                  ' About Form Anzeigen
End Sub

Private Sub mnuMainInfoBug_Click()
'    Call ReportBug
End Sub

Private Sub mnuMainInfoHelp_Click()
    Call ShowHelp                                                   ' Hilfe anzeigen
End Sub

Public Sub mnuMainWindowCascade_Click()
    Call MDIChildsCascade(Me)                                       ' Alle db Fenster Überlappend Versetzt anzeigen
End Sub

Public Sub mnuMainWindowHSplit_Click()
    Call MDIChildsHSplit(Me)                                        ' Alle db Fenster untereinader anzeigen
End Sub

Public Sub mnuMainWindowVSplit_Click()
    Call MDIChildsVSplit(Me)                                        ' Alle db Fenster nebeneinader anzeigen
End Sub
