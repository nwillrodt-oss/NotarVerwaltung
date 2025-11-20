VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Info zu meiner Anwendung"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   5380.766
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   3120
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&Systeminfo..."
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   3075
      Width           =   1485
   End
   Begin MSComctlLib.ImageList ILAbout 
      Left            =   2520
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":038A
            Key             =   ""
            Object.Tag             =   "Word"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":13FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":20D8
            Key             =   ""
            Object.Tag             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":2472
            Key             =   ""
            Object.Tag             =   "Info"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":280C
            Key             =   ""
            Object.Tag             =   "Search"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":2DA6
            Key             =   ""
            Object.Tag             =   "Refresh"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":3140
            Key             =   ""
            Object.Tag             =   "Back"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":34DA
            Key             =   ""
            Object.Tag             =   "Forward"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":3874
            Key             =   ""
            Object.Tag             =   "Add"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":3C0E
            Key             =   ""
            Object.Tag             =   "Print"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblWWW 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmAbout.frx":41A8
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   6
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Image ImgAbout 
      Height          =   2295
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Innen ausgefüllt
      Index           =   1
      X1              =   112.686
      X2              =   5337.57
      Y1              =   2070.653
      Y2              =   2070.653
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Beschreibung"
      ForeColor       =   &H00000000&
      Height          =   1890
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name der Anwendung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   2655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   2070.653
      Y2              =   2070.653
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   225
      Left            =   1800
      TabIndex        =   5
      Top             =   720
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3840
      TabIndex        =   3
      Top             =   2625
      Width           =   1830
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MODULNAME = "frmAbout"                                ' Modulname für Fehlerbehandlung

Public objObjectBag As Object                                       ' ObjectBag object

''' Registrierungsschlüssel-Sicherheitsoptionen...
'Const READ_CONTROL = &H20000
'Const KEY_QUERY_VALUE = &H1
'Const KEY_SET_VALUE = &H2
'Const KEY_CREATE_SUB_KEY = &H4
'Const KEY_ENUMERATE_SUB_KEYS = &H8
'Const KEY_NOTIFY = &H10
'Const KEY_CREATE_LINK = &H20
'Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
'                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
'                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
'
'' Registrierungsschlüssel-Stammtypen...
'Const HKEY_LOCAL_MACHINE = &H80000002
'Const ERROR_SUCCESS = 0
'Const REG_SZ = 1                         ' Null-terminierte Unicode-Zeichenfolge
'Const REG_DWORD = 4                      ' 32-Bit-Zahl
'
'Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
'Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
'Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub Form_Load()
'    'Me.Caption = "Info zu " & App.Title
'    'Me.Caption = "Info zu " & szAppTitel
'    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
'    'lblTitle.Caption = App.Title
'    'lblTitle.Caption = szAppTitel
'    'lblDescription = szCopyright
End Sub

Private Sub cmdSysInfo_Click()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call objObjectBag.ShowMSInfo                                    ' MS Info öffnen
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub cmdOK_Click()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Unload Me                                                       ' Form entladen
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub lblWWW_Click()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call objObjectBag.FolowLink(lblWWW.Caption)                 ' Link Folgen
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

'Public Sub StartSysInfo()
'
'    Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
'Const gREGVALSYSINFOLOC = "MSINFO"
'Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
'Const gREGVALSYSINFO = "PATH"
'
'    On Error GoTo SysInfoErr
'
'    Dim rc As Long
'    Dim SysInfoPath As String
'
'    ' Versuchen, den Systeminfo-Programmpfad/-namen aus der Registrierung abzurufen...
'    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
'    ' Versuchen, nur den Systeminfo-Programmpfad aus der Registrierung abzurufen...
'    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
'        ' Überprüfen, ob bekannte 32-Dateiversion vorhanden ist
'        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
'            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
'
'        ' Fehler - Datei wurde nicht gefunden...
'        Else
'            GoTo SysInfoErr
'        End If
'    ' Fehler - Registrierungseintrag wurde nicht gefunden...
'    Else
'        GoTo SysInfoErr
'    End If
'
'    Call Shell(SysInfoPath, vbNormalFocus)
'
'    Exit Sub
'SysInfoErr:
'    MsgBox "Systeminformationen sind momentan nicht verfügbar", vbOKOnly
'End Sub
'
'Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
'    Dim i As Long                                           ' Schleifenzähler
'    Dim rc As Long                                          ' Rückgabe-Code
'    Dim hKey As Long                                        ' Zugriffsnummer für einen offenen Registrierungsschlüssel
'    Dim hDepth As Long                                      '
'    Dim KeyValType As Long                                  ' Datentyp eines Registrierungsschlüssels
'    Dim tmpVal As String                                    ' Temporärer Speicher eines Registrierungsschlüsselwertes
'    Dim KeyValSize As Long                                  ' Größe der Registrierungsschlüsselvariablen
'    '------------------------------------------------------------
'    ' Registrierungsschlüssel unter KeyRoot {HKEY_LOCAL_MACHINE...} öffnen
'    '------------------------------------------------------------
'    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Registrierungsschlüssel öffnen
'
'    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Fehler behandeln...
'
'    tmpVal = String$(1024, 0)                             ' Platz für Variable reservieren
'    KeyValSize = 1024                                       ' Größe der Variable markieren
'
'    '------------------------------------------------------------
'    ' Registrierungsschlüsselwert abrufen...
'    '------------------------------------------------------------
'    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
'                         KeyValType, tmpVal, KeyValSize)    ' Schlüsselwert abrufen/erstellen
'
'    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Fehler behandeln
'
'    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 fügt null-terminierte Zeichenfolge hinzu...
'        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null gefunden, aus Zeichenfolge extrahieren
'    Else                                                    ' Keine null-terminierte Zeichenfolge für WinNT...
'        tmpVal = Left(tmpVal, KeyValSize)                   ' Null nicht gefunden, nur Zeichenfolge extrahieren
'    End If
'    '------------------------------------------------------------
'    ' Schlüsselwerttyp für Konvertierung bestimmen...
'    '------------------------------------------------------------
'    Select Case KeyValType                                  ' Datentypen durchsuchen...
'    Case REG_SZ                                             ' Zeichenfolge für Registrierungsschlüsseldatentyp
'        KeyVal = tmpVal                                     ' Zeichenfolgenwert kopieren
'    Case REG_DWORD                                          ' Registrierungsschlüsseldatentyp DWORD
'        For i = Len(tmpVal) To 1 Step -1                    ' Jedes Bit konvertieren
'            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Wert Zeichen für Zeichen erstellen
'        Next
'        KeyVal = Format$("&h" + KeyVal)                     ' DWORD in Zeichenfolge konvertieren
'    End Select
'
'    GetKeyValue = True                                      ' Erfolgreiche Ausführung zurückgeben
'    rc = RegCloseKey(hKey)                                  ' Registrierungsschlüssel schließen
'    Exit Function                                           ' Beenden
'
'GetKeyError:      ' Bereinigen, nachdem ein Fehler aufgetreten ist...
'    KeyVal = ""                                             ' Rückgabewert auf leere Zeichenfolge setzen
'    GetKeyValue = False                                     ' Fehlgeschlagene Ausführung zurückgeben
'    rc = RegCloseKey(hKey)                                  ' Registrierungsschlüssel schließen
'End Function

