VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit                                                     ' Variaben Deklaration erzwingen
Option Compare Text                                                 ' Sortierreihenfolge festlegen
Private Const MODULNAME = "clsRegTools"                             ' Modulname für Fehlerbehandlung

Private objObjectBag As Object                                      ' ObjectBag Class
Private objError As Object                                          ' Error Class

Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
On Error Resume Next                                                ' Hier keine Fehler mehr
    Set objObjectBag = Nothing                                      ' Objectbag Class Schliessen
    Set objError = Nothing                                          ' Error Class Schliesen
    Err.Clear                                                       ' Evtl. erroe clearen
End Sub

Public Function InitObjectBag(objOb As Object)
    Set objObjectBag = objOb
    If Not objObjectBag Is Nothing Then
        Set objError = objObjectBag.GetErrorObj
    End If
End Function

'Public Function RegDir(szRootKey As String, szRegKey As String) As String()
'' Liefert ein Array mit Regwerten zurück
'    Dim hKey&, Result&
'    Dim Cnt As Integer                                              ' Counter
'    Dim fTime As FILETIME
'    Dim Key$, Value$, ValueLen&, RegTyp&
'    Dim KeyName As String
'    Dim KeyLen As String
'    Dim lngRootKey As Long                                          ' RootKey als Long
'    Dim SubKeyArray() As String                                     ' Ergebnis array
'    Dim NewIndex As Integer                                         ' Neuer Array index
'    Dim tmp(0 To 256) As Byte
'On Error GoTo errorhandler                                          ' Fehlerbehandlung aktivieren
'    ReDim SubKeyArray(0)                                            ' Array Dimensonieren
'    lngRootKey = CheckRootKey(szRootKey)                            ' RootKey übersetzen
'    Result = RegOpenKeyEx(lngRootKey, szRegKey, 0, KEY_QUERY_VALUE, hKey) ' Regschlüssel öffnen
'    If Result = ERROR_SUCCESS Then                                  ' Wenn RegKy Existiert (bzw. gelesen werden kann)
''        Do
''            KeyLen = 256
''            Key = Space(KeyLen)
''            Result = RegEnumKeyEx(hKey, Cnt, Key, KeyLen, 0&, _
''                              vbNullChar, 0&, fTime)
''            If Result = ERROR_SUCCESS Then
''                NewIndex = UBound(SubKeyArray) + 1
''                ReDim Preserve SubKeyArray(NewIndex)                ' Array Redimensionieren
''                SubKeyArray(NewIndex) = UCase(Left$(Key, KeyLen))
''            End If
''            Cnt = Cnt + 1                                           ' hochzählen
''        Loop Until Result <> ERROR_SUCCESS                          ' Weiter bis Fehler
'        Cnt = 0                                                     ' Counter zurücksetzen
'        Do
'            KeyName = Space(255)
'            KeyLen = Len(KeyName)
'            Result = RegEnumValue(hKey, Cnt, KeyName, ByVal KeyLen, 0&, _
'                              0&, tmp(0), 256)
'            If Result = ERROR_SUCCESS Then
'                NewIndex = UBound(SubKeyArray) + 1
'                ReDim Preserve SubKeyArray(NewIndex)                ' Array Redimensionieren
'                SubKeyArray(NewIndex) = UCase(Left$(Value, ValueLen))
'            End If
'            Cnt = Cnt + 1                                           ' hochzählen
'        Loop Until Result <> ERROR_SUCCESS                          ' Weiter bis Fehler
'    End If
'    RegDir = SubKeyArray                                            ' Ergebnis zurück
'Exit Function
'errorhandler:
'    Stop
'    Debug.Print CStr(Err.Number) & " " & Err.Description
'
'End Function

Public Function DelRegKey(szRootKey As String, _
        szRegKey As String)
' Löscht regkey
    Dim lngRootKey As Long                                          ' RootKey als Long
On Error Resume Next                                                ' Hier keine Fehlerbehandlung
    lngRootKey = CheckRootKey(szRootKey)                            ' szRootKey übersetzen
    Call RegKeyDelete(lngRootKey, szRegKey)                         ' Regkey löschen
    Err.Clear                                                       ' Evtl. Error resetten
End Function

Public Function ReadRegValue(szRootKey As String, _
        szRegKey As String, szField As String, Optional Default As String) As String
    Dim lngRootKey As Long                                          ' RootKey als Long
    Dim Val As Variant
    Dim Result As Long
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    lngRootKey = CheckRootKey(szRootKey)                            ' szRootKey übersetzen
    Result = RegValueGet(lngRootKey, szRegKey, szField, Val)
    If Result <> ERROR_SUCCESS Then
        If Default <> "" Then
            Result = RegKeyCreate(lngRootKey, szRegKey)
            Result = RegValueSet(lngRootKey, szRegKey, szField, Default)
            Val = Default
        End If
    End If
    ReadRegValue = Val
exithandler:
Exit Function                                                       ' Function beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ReadRegValue", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function WriteRegValue(szRootKey As String, _
        szRegKey As String, szField As String, szValue As String) As String
' Schreibt regwert in registry
    Dim lngRootKey As Long                                          ' RootKey als Long
    Dim Result As Long
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    lngRootKey = CheckRootKey(szRootKey)
    Result = RegValueSet(lngRootKey, szRegKey, szField, szValue)
    If Result <> ERROR_SUCCESS Then
        Result = RegKeyCreate(lngRootKey, szRegKey)
        Result = RegValueSet(lngRootKey, szRegKey, szField, szValue)
    End If
    WriteRegValue = szValue                                         ' Wert zurück geben
exithandler:
Exit Function                                                       ' Function beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "WriteRegValue", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Private Function CheckRootKey(szRootKey As String) As Long
On Error Resume Next                                                ' Hier keine Fehlerbehandlung
    Select Case szRootKey
    Case HKCU, "HKCU"                                               ' Current User
        CheckRootKey = HKEY_CURRENT_USER
    Case HKLM, "HKLM"                                               ' Local Machine
        CheckRootKey = HKEY_LOCAL_MACHINE
    Case HKCR, "HKCR"                                               ' Classes Root
        CheckRootKey = HKEY_CLASSES_ROOT
    Case Else
    End Select
    Err.Clear                                                       ' Evtl. Error resetten
End Function

Private Function RegKeyExist(Root&, Key$) As Long
'Prüfen ob ein Schlüssel existiert

    Dim Result As Long
    Dim hKey&

    Result = RegOpenKeyEx(Root, Key, 0, KEY_READ, hKey)
    If Result = ERROR_SUCCESS Then Call RegCloseKey(hKey)
    RegKeyExist = Result
    
End Function

Private Function RegKeyCreate(Root&, Newkey$) As Long
'Neuen Schlüssel erstellen

    Dim Result&, hKey&, Back&
    
    Result = RegCreateKeyEx(Root, Newkey, 0, vbNullString, _
                            REG_OPTION_NON_VOLATILE, _
                            KEY_ALL_ACCESS, 0&, hKey, Back)
    If Result = ERROR_SUCCESS Then
      Result = RegFlushKey(hKey)
      If Result = ERROR_SUCCESS Then Call RegCloseKey(hKey)
        RegKeyCreate = Back
    End If
    
Exit Function
Errorhandler:
    Stop
End Function

Private Function RegKeyDelete(Root&, Key$) As Long
'Schlüssel löschen
  
      RegKeyDelete = RegDeleteKey(Root, Key)
      
End Function

Private Function RegFieldDelete(Root&, Key$, Field$) As Long
'Feld löschen
    Dim Result&, hKey&
    
    Result = RegOpenKeyEx(Root, Key, 0, KEY_ALL_ACCESS, hKey)
    If Result = ERROR_SUCCESS Then
      Result = RegDeleteValue(hKey, Field)
      Result = RegCloseKey(hKey)
    End If
    RegFieldDelete = Result
End Function

Private Function RegValueSet(Root&, Key$, Field$, Value As Variant) As Long
'Wert in ein Feld der Registry schreiben

    Dim Result&, hKey&, s$, l&
    
    Result = RegOpenKeyEx(Root, Key, 0, KEY_ALL_ACCESS, hKey)
    If Result = ERROR_SUCCESS Then
      Select Case VarType(Value)
        Case vbInteger, vbLong
          l = CLng(Value)
          Result = RegSetValueEx(hKey, Field, 0, REG_DWORD, l, 4)
        Case vbString
          s = CStr(Value)
          Result = RegSetValueEx_Str(hKey, Field, 0, REG_SZ, s, _
                                        Len(s) + 1)
      End Select
      Result = RegCloseKey(hKey)
    End If
    
    RegValueSet = Result
End Function

Private Function RegValueGet(Root&, Key$, Field$, Value As Variant) As Long
'Wert aus einem Feld der Registry auslesen

    Dim Result&, hKey&, dwType&, Lng&, Buffer$, l&
    
    Result = RegOpenKeyEx(Root, Key, 0, KEY_READ, hKey)
    If Result = ERROR_SUCCESS Then
      Result = RegQueryValueEx(hKey, Field, 0&, dwType, ByVal 0&, l)
      If Result = ERROR_SUCCESS Then
        Select Case dwType
          Case REG_SZ
            Buffer = Space$(l + 1)
            'Buffer = Space$(l)
            Result = RegQueryValueEx(hKey, Field, 0&, _
                                     dwType, ByVal Buffer, l)
            If Result = ERROR_SUCCESS Then Value = Left(Buffer, Len(Buffer) - 2)
          Case REG_DWORD
            Result = RegQueryValueEx(hKey, Field, 0&, dwType, Lng, l)
            If Result = ERROR_SUCCESS Then Value = Lng
        End Select
      End If
    End If
    
    If Result = ERROR_SUCCESS Then Result = RegCloseKey(hKey)
    RegValueGet = Result
End Function



