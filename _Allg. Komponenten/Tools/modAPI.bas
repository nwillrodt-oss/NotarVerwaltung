Attribute VB_Name = "modAPI"
Option Explicit                                                     ' Variaben Deklaration erzwingen
Option Compare Text                                                 ' Sortierreihenfolge festlegen
Const MODULNAME = "modAPI"                                          ' Modulname für Fehlerbehandlung

Public Declare Function Sleep Lib "kernel32" (ByVal dwMilliSeconds As Long)

' API Für WIn Designs {
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function ActivateWindowTheme Lib "uxtheme" _
        Alias "SetWindowTheme" (ByVal hwd As Long, _
        Optional ByVal pszSubAppName As Long = 0, _
        Optional ByVal pszSubIdList As Long = 0) As Long
Public Declare Function DeactivateWindowTheme Lib "uxtheme" _
        Alias "SetWindowTheme" (ByVal hwd As Long, _
        Optional ByVal pszSubAppName As String = " ", _
        Optional ByVal pszSubIdList As String = " ") As Long
Public Declare Function IsThemeActive Lib "UxTheme.dll" () As Boolean
' }

' Api für Check exe  oder IDE {
Public Declare Function GetWindow Lib "user32" _
        (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
        (ByVal hWnd As Long, ByVal ipCalssName As String, _
        ByVal nMaxCount As Long) As Long
Public Const GW_OWNER = 4
' }

' Api für Pfad kürzen {
Public Declare Function PathCompactPath Lib "shlwapi" _
        Alias "PathCompactPathA" (ByVal hdc As Long, _
        ByVal lpszPath As String, ByVal dX As Long) _
        As Long
' }

' HTML-Help {
Public Declare Function HtmlHelp Lib "hhctrl.ocx" _
  Alias "HtmlHelpA" (ByVal hwndCaller As Long, _
  ByVal pszFile As String, ByVal uCommand As Long, _
  ByVal dwData As Long) As Long

Public Declare Function HtmlHelpTopic Lib "hhctrl.ocx" _
  Alias "HtmlHelpA" (ByVal hWnd As Long, _
  ByVal lpHelpFile As String, ByVal wCommand As Long, _
  ByVal dwData As String) As Long

Public Const HH_DISPLAY_TOPIC = &H0
' }

' ******************* Api Für Zeitmessungen
Public Declare Function GetTickCount Lib "kernel32" () As Long

' Benötigte API-Deklaration für PING
Public Declare Function IsDestinationReachable Lib _
  "Sensapi.dll" Alias "IsDestinationReachableA" _
  (ByVal lpszDestination As String, _
  lpQOCInfo As QOCINFO) As Long

Public Type QOCINFO
  dwSize As Long
  dwFlags As Long
  dwInSpeed As Long
  dwOutSpeed As Long
End Type

Public Declare Function FindWindow Lib "user32" Alias _
        "FindWindowA" (ByVal lpClassName As String, ByVal _
        lpWindowName As String) As Long

Public Declare Function GetParent Lib "user32" (ByVal hWnd As _
        Long) As Long
        
Public Declare Function GetWindowThreadProcessId Lib "user32" _
        (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Public Const GW_HWNDNEXT As Long = 2&

' API für Lokalen Computernamen
Public Declare Function GetComputerName Lib "kernel32" _
        Alias "GetComputerNameA" (ByVal lpBuffer As String, _
        nSize As Long) As Long

' API für Lokalen Benutzernamen
Public Declare Function GetUserName Lib "advapi32.dll" Alias _
     "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) _
      As Long

' API für ini
Public Declare Function WritePrivateProfileString Lib _
        "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal _
        lpKeyName As Any, ByVal lpString As Any, ByVal _
        lpFileName As String) As Long
        
Public Declare Function GetPrivateProfileString Lib _
        "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal _
        lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize _
        As Long, ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileSection Lib _
        "kernel32" Alias "WritePrivateProfileSectionA" _
        (ByVal lpAppName As String, ByVal lpString As _
        String, ByVal lpFileName As String) As Long
        
Public Declare Function GetPrivateProfileSection Lib _
        "kernel32" Alias "GetPrivateProfileSectionA" _
        (ByVal lpAppName As String, ByVal lpReturnedString _
        As String, ByVal nSize As Long, ByVal lpFileName _
        As String) As Long

'API für Open/SaveDialog
Public Declare Function GetOpenFileName Lib "comdlg32.dll" _
        Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) _
        As Long
        
Public Declare Function GetSaveFileName Lib "comdlg32.dll" _
        Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) _
        As Long

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Const OFN_ALLOWMULTISELECT As Long = &H200&
Public Const OFN_CREATEPROMPT As Long = &H2000&
Public Const OFN_ENABLEHOOK As Long = &H20&
Public Const OFN_ENABLETEMPLATE As Long = &H40&
Public Const OFN_ENABLETEMPLATEHANDLE As Long = &H80&
Public Const OFN_EXPLORER As Long = &H80000
Public Const OFN_EXTENSIONDIFFERENT As Long = &H400&
Public Const OFN_FILEMUSTEXIST As Long = &H1000&
Public Const OFN_HIDEREADONLY As Long = &H4&
Public Const OFN_LONGNAMES As Long = &H200000
Public Const OFN_NOCHANGEDIR As Long = &H8&
Public Const OFN_NODEREFERENCELINKS As Long = &H100000
Public Const OFN_NOLONGNAMES As Long = &H40000
Public Const OFN_NONETWORKBUTTON As Long = &H20000
Public Const OFN_NOREADONLYRETURN As Long = &H8000&
Public Const OFN_NOTESTFILECREATE As Long = &H10000
Public Const OFN_NOVALIDATE As Long = &H100&
Public Const OFN_OVERWRITEPROMPT As Long = &H2&
Public Const OFN_PATHMUSTEXIST As Long = &H800&
Public Const OFN_READONLY As Long = &H1&
Public Const OFN_SHAREAWARE As Long = &H4000&
Public Const OFN_SHAREFALLTHROUGH As Long = 2&
Public Const OFN_SHARENOWARN As Long = 1&
Public Const OFN_SHAREWARN As Long = 0&
Public Const OFN_SHOWHELP As Long = &H10&

' ******* API Für BrowsforFolder {
Public Type BROWSEINFO
  hwndOwner       As Long
  pIDLRoot        As Long
  pszDisplayName  As Long
  lpszTitle       As String
  ulFlags         As Long
  lpfnCallback    As Long
  lParam          As Long
  iImage          As Long
End Type
    
Public Declare Function SHBrowseForFolder Lib "Shell32" _
       (lpbi As BROWSEINFO) As Long

Public Declare Function SHGetPathFromIDList Lib "Shell32" _
       (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)

Public Const MAX_PATH = 260
Public Const BIF_RETURNONLYFSDIRS = &H1& 'Nur Verzeichnisse zurückgeben
' }

'**** API für Menü Color ****

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal _
      crColor As Long) As Long

Private Declare Function GetMenu Lib "user32" (ByVal hWnd _
      As Long) As Long

Private Declare Function GetMenuItemCountA Lib "user32" Alias _
      "GetMenuItemCount" (ByVal hMenu As Long) As Long

Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu _
      As Long, ByVal nPos As Long) As Long

Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd _
      As Long, ByVal bRevert As Long) As Long

Private Declare Function SetMenuInfo Lib "user32" (ByVal hMenu _
      As Long, lpcmi As MENUINFO) As Long

Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd _
      As Long) As Long

Private Declare Sub OleTranslateColor Lib "olepro32.dll" (ByVal _
      clr As Long, ByVal hPal As Long, pcolorref As Long)

Private Type MENUINFO
    cbSize          As Long
    fMask           As Long
    dwStyle         As Long
    cyMax           As Long
    hbrBack         As Long
    dwContextHelpID As Long
    dwMenuData      As Long
End Type

Public Enum MenuNFO
    mMenuBarColor = 1
    mMenuColor = 2
    mSysMenuColor = 3
End Enum

Private Const MIM_BACKGROUND As Long = &H2&
Private Const MIM_APPLYTOSUBMENUS As Long = &H80000000
' }

' API für Screenshot {
Public Declare Function GetDesktopWindow Lib "user32" () _
        As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As _
        Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal _
        hWnd As Long, lpRect As RECT) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc _
        As Long, ByVal x As Long, ByVal y As Long, ByVal _
        nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC _
        As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
        ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
        ByVal dwRop As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd _
        As Long, ByVal hdc As Long) As Long

Public Type RECT
  Left As Long
  Top As Long
  Width As Long
  Height As Long
End Type
' }
                                                                    ' API Fur Transparente Fenster {
Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, _
    ByVal crKey As Long, _
    ByVal bAlpha As Byte, _
    ByVal dwFlags As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
'Public Const LWA_COLORKEY = &H1                                     ' Macht nur eine Farbe transparent
Public Const LWA_ALPHA = &H2                                        ' Macht das ganze Fenster transparent
                                                                    ' API Fur Transparente Fenster }
                                                                    
Public Function Set_MenuColor(SetWhat As MenuNFO, _
      ByVal hWnd As Long, ByVal Color As Long, _
      Optional MenuIndex As Integer, Optional _
      IncludeSubmenus As Boolean = True) As Boolean
      
  Dim MI As MENUINFO
  Dim clrref As Long, hSysMenu As Long, mHwnd As Long

  On Local Error GoTo errQuit

  clrref = Convert_OLEtoRBG(Color)

  MI.cbSize = Len(MI)
  MI.hbrBack = CreateSolidBrush(clrref)

  Select Case SetWhat
    Case mMenuBarColor
      MI.fMask = MIM_BACKGROUND
      SetMenuInfo GetMenu(hWnd), MI

    Case mMenuColor
      If MenuIndex = 0 Then
        Set_MenuColor = Set_MenuColor(mMenuBarColor, hWnd, Color)
        Exit Function
      End If

      If MenuIndex < 1 Or Get_MenuItemCount(hWnd) < MenuIndex Then
        Exit Function
      End If

      MI.fMask = IIf(IncludeSubmenus, MIM_BACKGROUND Or _
                    MIM_APPLYTOSUBMENUS, MIM_BACKGROUND)

      mHwnd = GetMenu(hWnd)
      mHwnd = GetSubMenu(mHwnd, MenuIndex - 1)

      SetMenuInfo mHwnd, MI
      hWnd = mHwnd

    Case mSysMenuColor
      hSysMenu = GetSystemMenu(hWnd, False)

      MI.fMask = MIM_BACKGROUND Or MIM_APPLYTOSUBMENUS

      SetMenuInfo hSysMenu, MI
      hWnd = hSysMenu

    Case Else
  End Select

  DrawMenuBar hWnd

  Set_MenuColor = True

errQuit:
End Function

Private Function Convert_OLEtoRBG(ByVal OLEcolor As Long) As Long
  OleTranslateColor OLEcolor, 0, Convert_OLEtoRBG
End Function

Private Function Get_MenuItemCount(ByVal hWnd As Long) As Long
  Get_MenuItemCount = GetMenuItemCountA(Get_MenuHwnd(hWnd))
End Function

Private Function Get_MenuHwnd(ByVal hWnd As Long) As Long
  Get_MenuHwnd = GetMenu(hWnd)
End Function

