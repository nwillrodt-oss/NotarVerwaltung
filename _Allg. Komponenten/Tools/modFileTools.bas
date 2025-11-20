Attribute VB_Name = "modFileTools"
Option Explicit                                                     ' Variaben Deklaration erzwingen
Option Compare Text                                                 ' Sortierreihenfolge festlegen
Const MODULNAME = "modFileTools"                                    ' Modulname für Fehlerbehandlung

Public objError As Object                                           ' Error Object

Public Function DeleteFolder(szFolderPath As String, bForce As Boolean) As Boolean
    Dim oFSO As New FileSystemObject                                ' File System Object
    Dim oFolder As Folder                                           ' Folder object
    Dim oFile As File                                               ' File object
    Dim bNotEmpty As Boolean                                        ' Leer Variable
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Set oFolder = oFSO.GetFolder(szFolderPath)                      ' Folder object holen
    If oFolder Is Nothing Then GoTo exithandler                     ' Kein folder -> fertig
    For Each oFile In oFolder.Files
        bNotEmpty = True                                            ' File gefunden Ordner nicht leer
    Next                                                            ' nächstes file
    If Not bNotEmpty Then                                           ' Wenn Leer
         Call oFSO.DeleteFolder(szFolderPath, bForce)               ' Ordner Löschen
         DeleteFolder = True                                        ' Erfolg zurück
    End If
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Set oFSO = Nothing                                              ' FileSystemObject killen
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "DeleteFolder", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function DeleteFile(szFilePath As String, Optional bForce As Boolean) As Boolean
    Dim oFSO As New FileSystemObject                                ' File System Object
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If szFilePath = "" Then GoTo exithandler                        ' Kein Pfad Fertig
    If ExistFile(szFilePath) Then                                   ' Wenn File Existiert
        Call oFSO.DeleteFile(szFilePath, bForce)                    ' Löschen
        DeleteFile = True
    Else
        DeleteFile = True
    End If
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Set oFSO = Nothing                                              ' FileSystemObject killen
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "DeleteFile", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ExistFile(szFilePath As String, Optional bForce As Boolean) As Boolean
    Dim oFSO As New FileSystemObject                                ' File System Object
    Dim szDirPath As String                                         ' VerzeichnisPfad (ohne Filename)
    Dim szFileName As String                                        ' Dateiname ohne verzeichnisPfad
    Dim szFileExtenntion As String
    Dim oFolder As Folder                                           ' Folder object
    Dim oFile As File                                               ' File object
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If szFilePath = "" Then GoTo exithandler                        ' Kein Pfad Fertig
    If InStr(szFilePath, "\*.") Then
        szFileExtenntion = Right(szFilePath, 3)
        szDirPath = Replace(szFilePath, "*." & szFileExtenntion, "")
        Set oFolder = oFSO.GetFolder(szDirPath)
        For Each oFile In oFolder.Files
                ' nur .txt Dateien!
            If LCase(oFSO.GetExtensionName(oFile)) = LCase(szFileExtenntion) Then
                ExistFile = True
                GoTo exithandler
            End If
        Next
    End If
    If Not oFSO.FileExists(szFilePath) Then                         ' Kein File vorhanden
        ExistFile = False                                           ' Falsch zurück
        If bForce Then                                              ' Wenn bForce
            Call oFSO.CreateTextFile(szFilePath, True)              ' Anlegen
            ExistFile = True
        End If
    Else
        ExistFile = True                                            ' Alles OK
    End If
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Set oFSO = Nothing                                              ' FileSystemObject killen
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ExistFile", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ExistFolder(szFolderPath As String, Optional bForce As Boolean) As Boolean
    Dim oFSO As New FileSystemObject                                ' File System Obj
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If szFolderPath = "" Then GoTo exithandler                      ' kein Ordner angegeben - > Fertig
    If oFSO.FolderExists(szFolderPath) Then                         ' Wenn Ornder ex.
        ExistFolder = True                                          ' Ordner ex. zurück
    Else                                                            ' Sonst
        If bForce Then                                              ' Wenn Ordner anlegen erzwingen
            If CreateFolder(szFolderPath) Then                      ' Ordner anlegen
                ExistFolder = True                                  ' Erfolg zurück
            End If
        Else                                                        ' Sonst
            ExistFolder = False                                     ' Ex. nicht zurück
        End If
    End If
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Set oFSO = Nothing                                              ' FileSystemObject killen
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ExistFolder", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function CreateFolder(ByVal szFolderPath As String) As Boolean
    Dim tmpPathArray() As String                                    ' Pfad in Array
    Dim i As Integer                                                ' Array Counter
    Dim oFSO As New FileSystemObject                                ' File System Obj
    Dim tmpPath As String                                           ' Teilpfad zum stückweisen erzeugen
    Dim bUncPath As Boolean                                         ' True wenn UNC Pfad
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szFolderPath = Trim(szFolderPath)                               ' Pfad trimmen
    If szFolderPath = "" Then GoTo exithandler                      ' Kein Pfad -> fertig
    bUncPath = CBool(Left(szFolderPath, 2) = "\\")                  ' Feststellen ob UNCPfad
    szFolderPath = Replace(szFolderPath, "\\", "")                  ' führende \\ abschneiden
    tmpPathArray = Split(szFolderPath, "\")                         ' Pfad in Array aufspalten
    If bUncPath Then tmpPath = "\\"                                 ' Wenn UNC fügrende \\ an tmpPath
    For i = 0 To UBound(tmpPathArray)                               ' Array durchlaufen
        If tmpPathArray(i) <> "" Then                               ' Wenn Array Item <>"""
            tmpPath = tmpPath & tmpPathArray(i) & "\"               ' an tmpPath anhängen
            If Not ExistFolder(tmpPath, False) Then
                If i = 0 Then
                
                Else
                    Call oFSO.CreateFolder(tmpPath)
                End If
            End If
'            If i = 0 And Not bUncPath Then
'                                                                    ' LW -> nix tun
'            Else
'                Call oFSO.CreateFolder(tmpPath)                     ' Ordner erstellen
'            End If
            
        End If
    Next i                                                          ' Nächstes Array Item
    CreateFolder = True                                             ' Erfolg zurück

exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Set oFSO = Nothing                                              ' FileSystemObject killen
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "CreateFolder", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function CopyFile(szSourcePath As String, szDestPath As String, bQverwrite As Boolean) As Boolean
    Dim oFSO As New FileSystemObject                                ' File System Obj
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If szSourcePath = "" Then GoTo exithandler                      ' Kein SourceFile -> Fertig
    If szDestPath = "" Then GoTo exithandler                        ' Kein Zielpfad -> fertig
    If Not ExistFile(szSourcePath) Then GoTo exithandler            ' Kein SourceFile -> Fertig
    Call oFSO.CopyFile(szSourcePath, szDestPath, bQverwrite)        ' Copieren und Überschreiben
    CopyFile = True                                                 ' Erfolg zurück
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Set oFSO = Nothing                                              ' FileSystemObject killen
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "CopyFile", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function
