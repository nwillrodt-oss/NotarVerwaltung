Imports Microsoft.Win32                                                         ' Win32 Importieren (Spart schreibarbeit)

Module modRegTools
    Private Const MODULNAME = "modRegTools"                                     ' Modulname für Fehlerbehandlung

    Public Function ReadRegValue(ByVal szRegKey As String, ByVal szValueName As String, _
                               Optional ByVal szDefValue As String = "", _
                               Optional ByRef oBag As clsObjectBag = Nothing) As String
        Dim RegKey As RegistryKey                                               ' Aktueller RegKey
        Dim szValue As String                                                   ' Ausgelesener wert
        Try                                                                     ' Fehlerbehandlung aktivieren
            If szRegKey = "" Then Return "" ' Kein Regschlüssel -> fertig
            If szValueName = "" Then Return "" ' Kein Valuename -> Fertig
            RegKey = Registry.LocalMachine.OpenSubKey(szRegKey)                 ' Erstmal HKLM öffnen
            If IsNothing(RegKey) Then
                RegKey = Registry.CurrentUser.OpenSubKey(szRegKey, True)        ' Dann HKCU öffnen , gegebenenfalls anlegen)
            End If
            If IsNothing(RegKey) Then
                Return ""
            End If
            szValue = RegKey.GetValue(szValueName, szDefValue)                  ' Wert auslesen
            RegKey.Close()                                                      ' Regkey Schliessen
            Return szValue
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "ReadRegValue", ex)           ' Fehler behandlung aufrufen
            End If
            Return ""                                                           ' Misserfolg zurück
        End Try
    End Function

    Public Function WriteRegValue(ByVal szRegKey As String, ByVal szValueName As String, _
                               ByVal szValue As Object, _
                               Optional ByRef oBag As clsObjectBag = Nothing) As String
        Dim RegKey As RegistryKey                                               ' Aktueller RegKey
        Try                                                                     ' Fehlerbehandlung aktivieren
            If szRegKey = "" Then Return "" ' Kein Regschlüssel -> fertig
            If szValueName = "" Then Return "" ' Kein Valuename -> Fertig
            'RegKey = Registry.LocalMachine.OpenSubKey(szRegKey)                 ' Erstmal HKLM öffnen
            'If IsNothing(RegKey) Then
            RegKey = Registry.CurrentUser.OpenSubKey(szRegKey, True)            ' HKCU öffnen , gegebenenfalls anlegen)
            'End If
            If IsNothing(RegKey) Then
                Return ""
            End If
            RegKey.SetValue(szValueName, szValue)                               ' Wert auslesen
            RegKey.Close()                                                      ' Regkey Schliessen
            Return szValue
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "WriteRegValue", ex)          ' Fehler behandlung aufrufen
            End If
            Return ""                                                           ' Misserfolg zurück
        End Try
    End Function


End Module
