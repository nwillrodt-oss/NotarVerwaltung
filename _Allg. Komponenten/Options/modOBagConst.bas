Attribute VB_Name = "modOBagConst"
Option Explicit                                                     ' Variaben Deklaration erzwingen
Option Compare Text                                                 ' Sortierreihenfolge festlegen
Private Const MODULNAME = "modOBagConst"                            ' Modulname für Fehlerbehandlung

Public Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal lpParameters _
        As String, ByVal lpDirectory As String, ByVal nShowCmd _
        As Long) As Long
        
' ************API für Top Most
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Const HWND_TOPMOST As Long = -1&
Public Const HWND_NOTOPMOST As Long = -2&
        
Public Const OPTION_SPLASH_IMG = "SplashImage"
Public Const OPTION_ABOUT_IMG = "AboutImage"
Public Const OPTION_TRANZ = "Tranzparenz"

'************API für Admin Check
'Public Declare Function OpenSCManager Lib "advapi32.dll" _
'    Alias "OpenSCManagerA" (ByVal lpMachineName As String, _
'    ByVal lpDatabaseName As String, ByValdwDesiredAccess As Long) As Long
'
'Public Const GENERIC_EXECUTE = &H20000000
'Public Const GENERIC_READ = &H80000000
'Public Const GENERIC_WRITE = &H40000000

' Benötigte API-Deklarationen für Userinfo
Public Const SUCCESS As Long = 0&

Public Type USER_INFO_3
  usri3_name As Long
  usri3_password As Long
  usri3_password_age As Long
  usri3_priv As Long
  usri3_home_dir As Long
  usri3_comment As Long
  usri3_flags As Long
  usri3_script_path As Long
  usri3_auth_flags As Long
  usri3_full_name As Long
  usri3_usr_comment As Long
  usri3_parms As Long
  usri3_workstations As Long
  usri3_last_logon As Long
  usri3_last_logoff As Long
  usri3_acct_expires As Long
  usri3_max_storage As Long
  usri3_units_per_week As Long
  usri3_logon_hours As Long
  usri3_bad_pw_count As Long
  usri3_num_logons As Long
  usri3_logon_server As Long
  usri3_country_code As Long
  usri3_code_page As Long
  usri3_user_id As Long
  usri3_primary_group_id As Long
  usri3_profile As Long
  usri3_home_dir_drive As Long
  usri3_password_expired As Long
End Type

Public Declare Function NetUserGetInfo Lib "Netapi32.dll" ( _
  ServerName As Any, _
  UserName As Any, _
  ByVal Level As Long, _
  Buffer As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" _
  Alias "RtlMoveMemory" ( _
  Dest As Any, _
  ByVal Source As Long, _
  ByVal cbCopy As Long)

Public Declare Function NetApiBufferFree Lib "Netapi32.dll" ( _
  ByVal lpBuffer As Long) As Long

' in diesen Variablen werden die Benutzer-Infos gespeichert
Dim LastLogon, LastLogoff, NumLogons, Expire, PWAge, Priv, FullName As String

' **** API + Konstanten für Windows Folder {
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" _
        Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal _
        pszPath As String) As Long
        
Public Declare Function SHGetSpecialFolderLocation Lib _
        "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder _
        As Long, pidl As ITEMIDLIST) As Long
        
Public Declare Function GetTempPath Lib "kernel32" Alias _
        "GetTempPathA" (ByVal nBufferLength As Long, ByVal _
        lpBuffer As String) As Long
        
Public Type ITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As ITEMID
End Type

Private Const CSIDL_FLAG_CREATE As Long = &H8000
Private Const CSIDL_ADMINTOOLS As Long = &H30
Private Const CSIDL_ALTSTARTUP As Long = &H1D
Private Const CSIDL_APPDATA As Long = &H1A          ' Anwendungsdaten
Private Const CSIDL_BITBUCKET As Long = &HA
Private Const CSIDL_CDBURN_AREA As Long = &H3B      ' CD Burning
Private Const CSIDL_COMMON_ADMINTOOLS As Long = &H2F    ' Für alle Benutzer (All Users)
Private Const CSIDL_COMMON_ALTSTARTUP As Long = &H1D
Private Const CSIDL_COMMON_APPDATA As Long = &H23
Private Const CSIDL_COMMON_DESKTOPDIRECTORY As Long = &H19  ' Desktop
Private Const CSIDL_COMMON_DOCUMENTS As Long = &H2E
Private Const CSIDL_COMMON_FAVORITES As Long = &H1F     ' Favoriten
Private Const CSIDL_COMMON_MUSIC As Long = &H35     ' Gemeinsame Musik
Private Const CSIDL_COMMON_PICTURES As Long = &H36  ' Gemeinsame Bilder
Private Const CSIDL_COMMON_PROGRAMS As Long = &H17  ' Programme-Ordner im Startmenü
Private Const CSIDL_COMMON_STARTMENU As Long = &H16 ' Starmenü
Private Const CSIDL_COMMON_STARTUP As Long = &H18   ' Autostart
Private Const CSIDL_COMMON_TEMPLATES As Long = &H2D ' Vorlagen
Private Const CSIDL_COMMON_VIDEO As Long = &H37     ' Gemeinsame Videos
Private Const CSIDL_CONTROLS As Long = &H3
Private Const CSIDL_COOKIES As Long = &H21          ' Cookies
Private Const CSIDL_DESKTOP As Long = &H0           ' Desktop
'Private Const CSIDL_DESKTOPDIRECTORY As Long = &H10
Private Const CSIDL_DRIVES As Long = &H11           ' Treiber
Private Const CSIDL_FAVORITES As Long = &H6         ' Favoriten
Private Const CSIDL_FONTS As Long = &H14            ' Schriftarten
Private Const CSIDL_HISTORY As Long = &H22          ' Verlauf
Private Const CSIDL_INTERNET As Long = &H1
Private Const CSIDL_INTERNET_CACHE As Long = &H20   ' Temporäre Internetdateien
Private Const CSIDL_LOCAL_APPDATA As Long = &H1C    ' Anwendungsdaten
Private Const CSIDL_MYDOCUMENTS As Long = &HC
Private Const CSIDL_MYMUSIC As Long = &HD       ' Eigene Musik
Private Const CSIDL_MYPICTURES As Long = &H27    ' Eigene Bilder
Private Const CSIDL_MYVIDEO As Long = &HE        ' Eigene Videos
Private Const CSIDL_NETHOOD As Long = &H13       ' Netzwerkumgebung
Private Const CSIDL_NETWORK As Long = &H12
'Private Const CSIDL_PERSONAL As Long = &H5       ' Eigene Dateien
Private Const CSIDL_PRINTERS As Long = &H4
Private Const CSIDL_PRINTHOOD As Long = &H1B     ' Druckerumgebung
Private Const CSIDL_PROFILE As Long = &H28       ' Profil
Private Const CSIDL_PROGRAM_FILES As Long = &H26 ' Programme
Private Const CSIDL_PROGRAM_FILES_COMMON As Long = &H2B      ' Gemeinsamme Dateien
Private Const CSIDL_PROGRAMS As Long = &H2       ' Programme (im Startmenü)
Private Const CSIDL_RECENT As Long = &H8         ' Zuletzt verwendete Dokumente
Private Const CSIDL_SENDTO As Long = &H9         ' Senden An
Private Const CSIDL_STARTMENU As Long = &HB      ' Startmenü
Private Const CSIDL_STARTUP As Long = &H7        ' Autostart
'Private Const CSIDL_SYSTEM As Long = &H25        ' System (bzw. System32)
Private Const CSIDL_TEMPLATES As Long = &H15     ' Vorlagen
'Private Const CSIDL_WINDOWS As Long = &H24       ' Windows
'Private Const NOERROR As Long = 0&

Public Type Folders
    Windows As String
    WinSystem As String
    UserEigeneDateien As String
    UserDesktop As String
End Type
' **** API + Konstanten für Windows Folder }

Public WinFolder As Folders

Public Type AppInfo
    AppTitel As String
    AppCopyright As String
    AppRegRoot As String
    AppWWW As String
    AppSuportMail As String
    AppFolder As String
    AppImageFolder As String
    AppTemplateFolder As String
    AppSQLFolder As String
    AppVersion As String
    AppVersionMajor As String
    AppVersionMinor As String
    AppVersionRev As String
End Type

Public Type ComputerInfo
    Compname As String
    Domain As String
End Type

Public Type OfficeInfo
    AccessVersion As String
    AccessPath As String
    AccessVersionLong As String
    WordVersion As String
    WordPath As String
    WordVersionLong As String
    ExcelVersion As String
    ExcelPath As String
    ExcelVersionLong As String
    OutlookVersion As String
    OutlookPath As String
    OutlookVersionLong As String
End Type

Public Sub GetSytemFolder()
    'Dim Result As Long
    'Dim Buff As String

    Const CSIDL_WINDOWS As Long = &H24       ' Windows
    Const CSIDL_SYSTEM As Long = &H25        ' System (bzw. System32)
    Const CSIDL_PERSONAL As Long = &H5       ' Eigene Dateien
    Const CSIDL_DESKTOPDIRECTORY As Long = &H10
    
    With WinFolder
        .UserDesktop = GetPath(CSIDL_DESKTOPDIRECTORY)
        .UserEigeneDateien = GetPath(CSIDL_PERSONAL)
        .Windows = GetPath(CSIDL_WINDOWS)
        .WinSystem = GetPath(CSIDL_SYSTEM)
    End With
    'Label2.Caption = GetPath(CSIDL_STARTMENU)
    'Label3.Caption = GetPath(CSIDL_PROGRAM_FILES)
'    Label5.Caption = GetPath(CSIDL_FAVORITES)
'    Label6.Caption = GetPath(CSIDL_COMMON_STARTUP)
'    Label7.Caption = GetPath(CSIDL_RECENT)
'    Label8.Caption = GetPath(CSIDL_SENDTO)
'    Label9.Caption = GetPath(CSIDL_TEMPLATES)
'    Label10.Caption = GetPath(CSIDL_NETWORK)
'    Label11.Caption = GetPath(CSIDL_FONTS)
'    Label12.Caption = GetPath(CSIDL_INTERNET_CACHE)
'
'    Buff = Space$(512)
'    Result = GetTempPath(Len(Buff), Buff)
'    Label13.Caption = Trim$(Buff)
End Sub

Public Function IsUserAdmin(szUser As String, Optional szAnmeldeDomäne As String) As Boolean
  
    Dim oUser   As Object
    Dim oGroup As Object
    Dim oMember As Object
    Dim sUserName As String
    Dim Admin As Boolean

On Error GoTo exithandler
  
    ' Methode 1. User ist in der Lokalen Administratoren gruppe
    ' Name des aktuellen Benutzers
    sUserName = Environ$(szUser)
  
    ' Gruppe der Administratoren für den lokalen Rechner
    Set oGroup = GetObject("WinNT://./Administratoren,group")
  
    ' alle User in der Administratoren-Gruppe durchlaufen
    For Each oMember In oGroup.Members
        'If bDebug Then Debug.Print oMember.Name
        If oMember.Name = szUser Then
            ' aktueller Benutzer ist vorhanden!
            Admin = True
            Exit For
        End If
    Next
  
    If Admin Then GoTo exithandler
    If szAnmeldeDomäne <> "" Then
        ' Methode 2: User ist i der Gruppe der Domänen-Administratoren
        Set oUser = GetObject("WinNT://" & szAnmeldeDomäne & "/" & szUser & ",user")
        For Each oGroup In oUser.Groups
            If oGroup.Name = "Administratoren" Or oGroup.Name = "Domänen-Admins" Then
                Admin = True
            End If
        Next
    End If
    
    If Admin Then GoTo exithandler
    ' Methode 3: (Hadcore) versuchen in system32 zu schreiben
    

    
exithandler:
On Error Resume Next
    If Admin Then IsUserAdmin = True
  ' Objekte zerstören
  Set oMember = Nothing
  Set oGroup = Nothing
  Set oUser = Nothing

Exit Function
Errorhandler:
On Error Resume Next
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    'Call OBErrorhandler(MODULNAME, "IsUserAdmin", errNr, errDesc)
    Resume exithandler
End Function

Public Function GetPath(Num As Long) As String
    Dim Result As Long
    Dim Buff As String
    Dim idl As ITEMIDLIST
    Const NOERROR As Long = 0&
   
    'Result = SHGetSpecialFolderLocation(Me.hwnd, Num, idl)
    Result = SHGetSpecialFolderLocation(0, Num, idl)
    If Result = NOERROR Then
        Buff = Space$(512)
        Result = SHGetPathFromIDList(ByVal idl.mkid.cb, ByVal Buff)
        If Result Then
            GetPath = Left(Buff, InStr(Buff, Chr(0)) - 1)
        End If
    End If
End Function

Public Sub AddLVColumn(ctlListView As ListView, ColumnName As String)

    Dim i As Integer
    
On Error GoTo Errorhandler
    
    'ctlListView.columnHeaders.Clear
    If ctlListView.ColumnHeaders.Count > 0 Then
        For i = 1 To ctlListView.ColumnHeaders.Count
            If ColumnName = ctlListView.ColumnHeaders(i).Text Then GoTo exithandler
        Next i
    End If
    
    ctlListView.ColumnHeaders.Add ctlListView.ColumnHeaders.Count + 1, ColumnName, ColumnName
    
    ctlListView.Sorted = False

exithandler:

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    'Call OBErrorhandler(MODULNAME, "AddLVColumn", errNr, errDesc)
    Resume exithandler
End Sub

Public Sub AddListViewSubItem(ByVal itemMain As ListItem, szSubItemText As String)
    
    Dim Index As Integer

On Error GoTo Errorhandler
    
    Index = itemMain.ListSubItems.Count + 1
    itemMain.SubItems(Index) = Left(szSubItemText, 50)
    
exithandler:
On Error Resume Next
    'Set itemX = Nothing
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    'Call OBErrorhandler(MODULNAME, "AddListViewSubItem", errNr, errDesc)
    Resume exithandler
End Sub

Public Function AddListViewItem(ctlListView As ListView, _
        szItemText As String, _
        szItemID As String, _
        Optional szSubItemText As String, _
        Optional szValueName As String, _
        Optional intImage As Integer) As ListItem
        
     Dim itemX As ListItem
     
On Error GoTo Errorhandler
   
    Set itemX = ctlListView.ListItems.Add(, , szItemText, intImage)
    itemX.SmallIcon = intImage          ' Image setzen
    itemX.Tag = szItemID                ' ItemID Setzen
    If szSubItemText <> "" Then
        Call AddListViewSubItem(itemX, szSubItemText)   ' Sub Item Anlegen
    End If
    Set AddListViewItem = itemX
    
exithandler:
On Error Resume Next
    Set itemX = Nothing
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    'Call OBErrorhandler(MODULNAME, "AddListViewItem", errNr, errDesc)
    Resume exithandler
End Function


