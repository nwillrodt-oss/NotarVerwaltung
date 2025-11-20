Attribute VB_Name = "modRegAPI"
Option Explicit
Private Const MODULNAME = "modRegAPI"                               ' Modulname für Fehlerbehandlung

Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias _
        "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex _
        As Long, ByVal lpName As String, lpcbName As Long, _
        ByVal lpReserved As Long, ByVal lpClass As String, _
        lpcbClass As Long, lpftLastWriteTime As FILETIME) _
        As Long
        
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias _
        "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, _
        ByVal ipValueName As String, ipcbValueName As Long, _
        ByVal ipReserved As Long, ipType As Long, ipData As Byte, _
        ipcpData As Long) As Long

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal _
        lpSubKey As String, ByVal ulOptions As Long, ByVal _
        samDesired As Long, phkResult As Long) As Long
        
Public Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long
        
Public Declare Function RegQueryValueEx Lib "advapi32.dll" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal _
        lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As Any, lpcbData As Any) As Long
        
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" _
        Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal _
        lpSubKey As String, ByVal Reserved As Long, ByVal _
        lpClass As String, ByVal dwOptions As Long, ByVal _
        samDesired As Long, ByVal lpSecurityAttributes As Any, _
        phkResult As Long, lpdwDisposition As Long) As Long
        
Public Declare Function RegFlushKey Lib "advapi32.dll" (ByVal _
        hKey As Long) As Long
        
Public Declare Function RegSetValueEx Lib "advapi32.dll" _
        Alias "RegSetValueExA" (ByVal hKey As Long, ByVal _
        lpValueName As String, ByVal Reserved As Long, ByVal _
        dwType As Long, lpData As Long, ByVal cbData As Long) _
        As Long
        
Public Declare Function RegSetValueEx_Str Lib "advapi32.dll" _
        Alias "RegSetValueExA" (ByVal hKey As Long, ByVal _
        lpValueName As String, ByVal Reserved As Long, ByVal _
        dwType As Long, ByVal lpData As String, ByVal cbData As _
        Long) As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias _
        "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As _
        String) As Long
        
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
        "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName _
        As String) As Long

'Nur für Windows 2k/XP
Public Declare Function SHDeleteKey Lib "shlwapi.dll" Alias _
"SHDeleteKeyA" _
        (ByVal hKey As Long, ByVal pszSubKey As String) As Long

Public Const HKEY_CLASSES_ROOT As Long = &H80000000
Public Const HKEY_CURRENT_USER As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_USERS As Long = &H80000003
Public Const HKEY_PERFORMANCE_DATA As Long = &H80000004
Public Const HKEY_CURRENT_CONFIG As Long = &H80000005
Public Const HKEY_DYN_DATA As Long = &H80000006

Public Const KEY_QUERY_VALUE As Long = &H1
Public Const KEY_SET_VALUE As Long = &H2
Public Const KEY_CREATE_SUB_KEY As Long = &H4
Public Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Public Const KEY_NOTIFY As Long = &H10
Public Const KEY_CREATE_LINK As Long = &H20

Public Const KEY_READ As Long = KEY_QUERY_VALUE Or _
                 KEY_ENUMERATE_SUB_KEYS _
                 Or KEY_NOTIFY
                 
Public Const KEY_ALL_ACCESS As Long = KEY_QUERY_VALUE Or _
                       KEY_SET_VALUE Or _
                       KEY_CREATE_SUB_KEY Or _
                       KEY_ENUMERATE_SUB_KEYS Or _
                       KEY_NOTIFY Or _
                       KEY_CREATE_LINK
                       
Public Const ERROR_SUCCESS As Long = 0&                             ' Kein Fehler  (error = 0)
Public Const REG_NONE As Long = 0&
Public Const REG_SZ As Long = 1&
Public Const REG_EXPAND_SZ As Long = 2&
Public Const REG_BINARY As Long = 3&
Public Const REG_DWORD As Long = 4&
Public Const REG_DWORD_LITTLE_ENDIAN As Long = 4&
Public Const REG_DWORD_BIG_ENDIAN As Long = 5&
Public Const REG_LINK As Long = 6&
Public Const REG_MULTI_SZ As Long = 7&

Public Const REG_OPTION_NON_VOLATILE As Long = &H0&

Public Const HKLM = "HKEY_LOCAL_MACHINE"
Public Const HKCU = "HKEY_CURRENT_USER"
Public Const HKCR = "HKEY_CLASSES_ROOT"

Public Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

