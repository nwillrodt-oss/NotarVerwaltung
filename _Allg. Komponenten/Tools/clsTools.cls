VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit                                                     ' Variaben Deklaration erzwingen
Option Compare Text                                                 ' Sortierreihenfolge festlegen
Const MODULNAME = "clsTools"                                        ' Modulname für Fehlerbehandlung

Private objObjectBag As Object                                      ' ObjectBag class
Public objError As Object                                           ' Error class
Private objWMIService As Object                                     ' winmgmts:\\"
Public objLocator As Object
Private fs As frmScreenshot                                         ' Screenshot form

Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Set objError = Nothing                                          ' Errorobject Weg
    Set objObjectBag = Nothing                                      ' Object Bag weg
    Set objWMIService = Nothing                                     ' WMI Object Weg
    Err.Clear                                                       ' Evtl error clearen
End Sub

Public Function InitObjectBag(objOb As Object)
    Set objObjectBag = objOb                                        ' ObjectBag übernehmen
    Call InitCommonControls                                         ' Windows Themes zulassen
    If Not objObjectBag Is Nothing Then                             ' Wenn object bag da
        Set objError = objObjectBag.GetErrorObj                     ' Dann error object übernehmen
    End If
End Function
                                                                    '************************************************
                                                                    ' ScreenShot
Public Function ScreenShotTake()
' Läd eine Aktuellen Schreenshot in dei Picutebox auf femScreenshot
    Dim Result As Long
    Dim DesktopHwnd As Long                                         ' Desktop Handel
    Dim DesktopHdc&                                                 ' Desktop Geräte kontext
    Dim Desktop As RECT                                             ' Desktop abmessungen
    Const SRCCOPY = &HCC0020
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If Not fs Is Nothing Then                                       ' Scrennshot form noch geladen
        Unload fs                                                   ' Schliessen
        Set fs = Nothing                                            ' Referens aufs Form freigeben
    End If
    Set fs = New frmScreenshot                                      ' Screenshot form öffnen
    fs.Hide                                                         ' Form Verstecken
    fs.picScreenShot.AutoRedraw = True                              '
                                                                    '### Desktopgröße in Pixeln ermitteln
    DesktopHwnd = GetDesktopWindow()                                ' Desktop Handel ermitteln
    DesktopHdc = GetDC(DesktopHwnd)
    Result = GetWindowRect(DesktopHwnd, Desktop)                    ' Desktop abmessungen ermitteln
                                                                    '### Zielbild und Scrollbalken der Desktopgröße anpassen
    fs.picScreenShot.Width = Desktop.Width * 15                     ' Breite Picturebox anpassen
    fs.picScreenShot.Height = Desktop.Height * 15                   ' Höhe Picturebox anpassen
                                                                    '### Der eigentliche Screenshot
    Result = StretchBlt(fs.picScreenShot.hdc, Desktop.Left, Desktop.Top, Desktop.Width, Desktop.Height, _
            DesktopHdc, 0, 0, Desktop.Width, Desktop.Height, _
            SRCCOPY)                                                ' Screnshot im Picutebox laden
                                                                    '### Gerätekontext löschen
    Result = ReleaseDC(DesktopHwnd, DesktopHdc)
    fs.picScreenShot.Refresh
    fs.picScreenShot.AutoRedraw = False
exithandler:
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ScreenShotTake", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ScreenshotSave(szBMPFileName As String) As Boolean
' Speichert den akuellen Screenshot als bmp datei
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If fs Is Nothing Then GoTo exithandler                          ' Scrennshot form nicht geladen -> Fertig
    If szBMPFileName = "" Then                                      ' Kein Dateiname angegeben
        szBMPFileName = Year(Now()) & Month(Now()) & Day(Now()) & Hour(Now()) & Minute(Now()) & Second(Now()) & ".bmp"
        szBMPFileName = objObjectBag.GetPersonalDir & szBMPFileName
    End If
    SavePicture fs.picScreenShot, szBMPFileName                     ' Screenshot speichern
    ScreenshotSave = True
exithandler:
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ScreenshotSave", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function
                                                                    '************************************************
                                                                    ' Datei und Path Select
Public Function ShowOpen(Filter As String, Flags As Long, hWnd As Long, Optional szDir As String) As String
' Zeigt Dialog zum Datei öffnen an
    Dim Buffer As String
    Dim Result As Long
    Dim ComDlgOpenFileName As OPENFILENAME
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Buffer = String$(128, 0)
    With ComDlgOpenFileName                                         ' Mit Öffnen Dialog
        .lStructSize = Len(ComDlgOpenFileName)
        .hwndOwner = hWnd                                           ' Owner Form übergeben
        .Flags = Flags                                              ' Flags setzen
        .nFilterIndex = 1&
        .nMaxFile = Len(Buffer)
        .lpstrFile = Buffer
        .lpstrFilter = Filter                                       ' Datei Filter setzen
        .lpstrInitialDir = szDir                                    ' Start Verz. angeben
    End With
    Result = GetOpenFileName(ComDlgOpenFileName)                    ' Ergebnis ?
    If Result <> 0 Then                                             ' Dateinemen zuück geben
        ShowOpen = Left$(ComDlgOpenFileName.lpstrFile, _
                   InStr(ComDlgOpenFileName.lpstrFile, _
                   Chr$(0)) - 1)
    End If
exithandler:
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ShowOpen", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ShowSave(Filter As String, Flags As Long, _
                         hWnd As Long, Filename As String) As String
' Zeigt Dialog zum Datei Speichern an
    Dim Buffer As String
    Dim Result As Long
    Dim ComDlgOpenFileName As OPENFILENAME
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Buffer = Filename & String$(128 - Len(Filename), 0)
    With ComDlgOpenFileName                                         ' Mit Speichenrn Dialog
        .lStructSize = Len(ComDlgOpenFileName)
        .hwndOwner = hWnd
        .Flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST
        .nFilterIndex = 1&
        .nMaxFile = Len(Buffer)
        .lpstrFile = Buffer
        .lpstrFilter = Filter
    End With
    Result = GetSaveFileName(ComDlgOpenFileName)
    If Result <> 0 Then
        ShowSave = Left$(ComDlgOpenFileName.lpstrFile, _
                   InStr(ComDlgOpenFileName.lpstrFile, _
                   Chr$(0)) - 1)
    End If
exithandler:
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ShowSave", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ShowDirSelect(szText As String, szDefPath As String, Optional Parent As Variant) As String
    Dim tBI         As BROWSEINFO
    Dim lhWndParent As Long
    Dim lngPIDL     As Long
    Dim strPath     As String
    If IsMissing(szText) Then szText = "Wählen Sie einen Ordner aus"
    If IsMissing(Parent) = False Then lhWndParent = Parent.hWnd
'GoTo PlanB
    With tBI
        .hwndOwner = lhWndParent
        .lpszTitle = szText
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    lngPIDL = SHBrowseForFolder(tBI)
    If (lngPIDL <> 0) Then                                          ' Pfad aus Item ID List ermitteln:
        strPath = Space$(MAX_PATH)
        SHGetPathFromIDList lngPIDL, strPath

        strPath = Left$(strPath, InStr(strPath, Chr$(0)) - 1)       ' Pfad preparieren
        CoTaskMemFree lngPIDL                                       ' PIDL freigeben:
        ShowDirSelect = strPath                                     ' Pfad zurück
    Else
        ShowDirSelect = szDefPath                                   ' Default Pfad zurück
    End If
Exit Function
' **** Plan B
'PlanB:
'
'    Dim objShell As Object
'    Dim objFolder As Object
'    Dim objItem As Object
'    Set objShell = CreateObject("Shell.Application")
'    With objShell
'        Set objFolder = .BrowseForFolder(0&, szText, 0, szDefPath)
'    End With
'    If Not objFolder Is Nothing Then
'        Set objItem = objFolder.Self
'        ShowDirSelect = objItem.Path
'    End If
''
End Function
                                                                    '************************************************
                                                                    ' textFile Schreiben / Lesen
Public Function TextFileWrite(FullFilePath As String, szText As String, _
        Optional bAppend As Boolean) As Boolean
' Schreibt szTest in tesxtdatei
    Dim FN As Integer
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If FileExist(FullFilePath, True) Then                           ' Ext Datei sonst anlegen
On Error Resume Next                                                ' Fehlerbehandlung Deaktivieren
        If Len(Dir$(FullFilePath, vbNormal)) > 0 Or bAppend Then    ' Text anhängen
            FN = FreeFile
            Open FullFilePath For Append As #FN                     ' Öffnen zum anhägen
            Print #FN, szText;                                      ' Text in Datei Schreiben
            Close FN                                                ' Datei Schliessen
        Else                                                        ' Sonst neu
            FN = FreeFile
            Open FullFilePath For Output As #FN                     ' Öffnen zum Schreiben
            Print #FN, szText;                                      ' Text in Datei Schreiben
            Close FN                                                ' Datei Schliessen
        End If
    End If
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "TextFileWrite", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit exithandler
End Function

Public Function TextFileRead(FullFilePath As String) As String
    Dim FN As Integer
    Dim szText As String                                            ' gelesener text
    Dim FileLength As Integer
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If FileExist(FullFilePath, False) Then                          ' Prüfen ob Datei ex (nicht anlegen)
'On Error Resume Next
        If Len(Dir$(FullFilePath, vbNormal)) > 0 Then
            FN = FreeFile
            Open FullFilePath For Input As #FN                      ' Datei zu lesen öffnen
            FileLength = LOF(FN)
            szText = Input(FileLength, #FN)                         ' Text einlesen
'            Input #FN, szText
            Close FN                                                ' Datei schliessen
        End If
    End If
    TextFileRead = szText                                           ' Gelesenen Text zurück
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "TextFileRead", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit exithandler
End Function
                                                                    '************************************************
                                                                    ' File & Folder
Public Function GetShortPath(F As Object, ByVal szPath As String, _
        Optional lngMaxChar As Long) As String
' Kürzt eine Pfad angabe damit diese in eine Textbox past
    Dim Result As Long                                              ' API Result
On Error Resume Next                                                ' Fehler behandlung deaktivieren
    Dim x As Long                                                   ' Trenn position
    If szPath = "" Then Exit Function                               ' Kein Pfad -> fertig
    If lngMaxChar = 0 Then lngMaxChar = 100                         ' Def Max. Länge eintragen
    Result = PathCompactPath(F.hdc, szPath, lngMaxChar)             ' Pfad kürzen
    x = InStr(1, szPath, Chr$(0))                                   ' Trenn Pos ermitteln
    If x <> 0 Then                                                  ' Trennzeichen vorhanden
        GetShortPath = Left$(szPath, x - 1)                         ' Alles links von Trennpos zurück
    Else                                                            ' Sonst
        GetShortPath = szPath                                       ' Muste nix gekürzt werden
    End If
    Err.Clear                                                       ' Evtl. Error Clearen
End Function

Public Function FileDelete(szPath As String, _
        Optional bForce As Boolean, _
        Optional bWithEmptyFolder As Boolean) As Boolean
' Löscht eine Datei
    Dim szFolderPath As String                                      ' Pfad ohne Dateiname
On Error Resume Next                                                ' Fehler behandlung deaktivieren
    FileDelete = DeleteFile(szPath, bForce)                         ' Datei löschen
    If FileDelete And bWithEmptyFolder Then
        szFolderPath = CutFilenameFromPath(szPath, "")              ' FolderPfad extrahieren
        FileDelete = DeleteFolder(szFolderPath, False)              ' Ordner löschen wenn leer
    End If
    Err.Clear                                                       ' Evtl. Error Clearen
End Function

Public Function FileCreate(szPath As String) As Boolean
' Legt Datei an
On Error Resume Next                                                ' Fehler behandlung deaktivieren
    FileCreate = FileExist(szPath, True)                            ' über Fileexist mit bForce True ausführen
    Err.Clear                                                       ' Evtl. Error Clearen
End Function

Public Function FileExist(szPath As String, Optional bForce As Boolean) As Boolean
' Prüft ob Datei existiert.  Bei bForce dird diese samt pfad angelegt (soweit rechte vorhanden)
    Dim szFileName As String                                        ' Dateiname Ohne pfad
    Dim szFolderPath As String                                      ' Dateipfad ohne File
On Error Resume Next                                                ' Fehler behandlung deaktivieren
    szFolderPath = CutFilenameFromPath(szPath, szFileName)          ' Pfad in Dateimanen und Ordnerpfad aufspalten
    If ExistFolder(szFolderPath, bForce) Then                       ' Erst Pfad prüfen
        FileExist = ExistFile(szPath, bForce)                       ' Datei Prüfen + Bei bForce = true anlegen
    Else                                                            ' Kein Pfad
        FileExist = False                                           ' -> kein File
    End If
    Err.Clear                                                       ' Evtl. Error Clearen
End Function

Public Function FileCopy(szSourcePath As String, szDestPath As String, _
        Optional bQverwrite As Boolean) As Boolean
    Dim Filename As String                                          ' Dateiname ohne Verzeichnise
    Dim DestPathWithoutFilename As String                           ' Ordner Pfad der Datei
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Filename = GetFileNameFromPath(szDestPath)                      ' Dateinamen extrahieren
    If Filename = "" Then GoTo exithandler                          ' Keine Datei zum kopieren -> Fertig
    DestPathWithoutFilename = CutFilenameFromPath(szDestPath, Filename) ' Pfad ohne Dateinemen ermitteln
    If ExistFolder(DestPathWithoutFilename, True) Then              ' Evtl. Verz anlegen
        FileCopy = CopyFile(szSourcePath, szDestPath, bQverwrite)   ' Datei kopieren
    End If
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "FileCopy", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function FolderCreate(szPath As String) As Boolean
' Legt Ordner an
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    FolderCreate = FolderExist(szPath, True)                        ' Über FolderExist mit bForce = True ausführen
    Err.Clear                                                       ' Evtl. Error Clearen
End Function

Public Function CheckPath(szPath As String, bCreate As Boolean) As Boolean
    ' Existiert der Ordner?
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If ExistFolder(szPath, bCreate) Then
        CheckPath = True
    Else
        CheckPath = False
    End If
    Err.Clear                                                       ' Error clearen
End Function

Public Function FolderExist(szPath As String, Optional bForce As Boolean) As Boolean
' Prüft ob Ordner existiert. Bei bForce wird diese Ordner struktur angelegt
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    FolderExist = ExistFolder(szPath, bForce)                       ' Ordner Prüfen + Bei bForce = true anlegen
    Err.Clear                                                       ' Evtl. Error Clearen
End Function

Public Function GetFileList(szFolder As String, Optional szDelimiter As String) As String
' Liefert Trennzeichen (delimiter) getrennte Liste aller dateien
    Dim oFSO As New FileSystemObject                                ' File System Object
    Dim oFolder As Folder                                           ' Ordner Object
    Dim oFile As File                                               ' File object
    Dim szFileList As String                                        ' Dateiliste
On Error GoTo Errorhandler                                          ' fehlerbehandlung aktivieren
    If szDelimiter = "" Then szDelimiter = ";"                      ' Simikolon wenn kein Trennzeichen angegeben
    If Not ExistFolder(szFolder, False) Then GoTo exithandler       ' Verz ex. nicht -> Fertig
    Set oFolder = oFSO.GetFolder(szFolder)                          ' ordner auswählen
    For Each oFile In oFolder.Files                                 ' Alle Files Duchlaufen
        szFileList = szFileList & oFile.Name & szDelimiter          ' An liste hängen
    Next                                                            ' nächste Datei
    GetFileList = CutLastChar(szFileList, szDelimiter)              ' letztes Trennzeichen abschneiden und zurück
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Set oFSO = Nothing                                              ' oFSO killen
    Set oFolder = Nothing                                           ' oFolder Killen
    Set oFile = Nothing                                             ' oFile killen
    Err.Clear                                                       ' Evtl error clearen
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "GetFileList", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function CutFilenameFromPath(szFilePath As String, szFileName As String) As String
' Trennt Dateinamen und Datei pfad
    Dim szFilePathArray() As String                                 ' Array mit Pafd elementen
    Dim szFolderPath As String                                      ' nur Verz. Pfad
    Dim i As Integer                                                ' Counter
    Const szPathSep = "\"
On Error GoTo Errorhandler                                          ' fehlerbehandlung aktivieren
    If szFilePath = "" Then GoTo exithandler                        ' Kein Pfad -> Fertig
    szFilePathArray = Split(szFilePath, szPathSep)                  ' FilePath in Array
    If InStr(szFilePathArray(UBound(szFilePathArray)), ".") > 0 Then ' . im Letzen element
        szFileName = szFilePathArray(UBound(szFilePathArray))       ' Letzes Element ist Dateiname
        For i = 0 To UBound(szFilePathArray) - 1                    ' Rest des Arrays duchlaufen
            szFolderPath = szFolderPath & szFilePathArray(i) & szPathSep ' DateiPfad wieder zusammensetzen
        Next                                                        ' Nächtes Array Item
        CutFilenameFromPath = szFolderPath                          ' DateiPfad zurück geben
        GoTo exithandler                                            ' Fertig
    Else                                                            ' Keine Dateiendung gefunden
        If szFileName = "" Then GoTo exithandler                    ' Filename nicht angegeben -> Fertig
        If InStr(UCase(szFilePathArray(UBound(szFilePathArray))), UCase(szFileName)) > 0 Then
            For i = 0 To UBound(szFilePathArray) - 1
                szFolderPath = szFolderPath & szFilePathArray(i) & szPathSep
            Next
            CutFilenameFromPath = szFolderPath                      ' DateiPfad zurück geben
            GoTo exithandler                                        ' Fertig
        End If
    End If
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "CutFilenameFromPath", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function GetFileNameFromPath(szPath As String) As String
    Dim PathArray() As String                                       ' Array mit Pfad elementen
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If szPath = "" Then GoTo exithandler                            ' Kein Pfad -> Fertig
    PathArray = Split(szPath, "\")                                  ' Pfad in array aufspalten
    If Not CheckArray(PathArray) Then GoTo exithandler              ' Wenn Array Leer -> Fertig
    GetFileNameFromPath = PathArray(UBound(PathArray))              ' Letzes Element ist Dateiname
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "GetFileNameFromPath", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function
                                                                    '************************************************
                                                                    ' XLM
                                                                    '************************************************
Public Function LoadXMLDoc(ByVal XMLDocPath As String) As DOMDocument
' Öffnet ein XML Dokument
    Dim XMLDoc As DOMDocument                                       ' XML Dokument
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If XMLDocPath = "" Then GoTo exithandler                        ' Kein Doc -> Raus
    Set XMLDoc = New DOMDocument
    XMLDoc.async = False
    XMLDoc.Load XMLDocPath                                          ' XML Doc Laden
    Set LoadXMLDoc = XMLDoc                                         ' Zurückgeben
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "LoadXMLDoc", errNr, errDesc)
End Function

Public Function GetDocSQL(XMLDocPath As String, Optional szSQLName As String) As String
' Liefert aus dem unterknoten = szSQLName von <Root><SAT>
' sen Wert des Atributs SQL
' Falls szSQLName ="" wird er erste unterknoten genommen
    Dim XMLDoc As DOMDocument                                       ' XML Dokument
    Dim cXMLNode As IXMLDOMNode                                     ' Akt. XML Node
    Dim XMLSQLNode As IXMLDOMNode                                   ' Selektierte ChildNode von cXMLNode
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If XMLDocPath = "" Then GoTo exithandler                        ' Kein Doc -> Raus
    Set XMLDoc = LoadXMLDoc(XMLDocPath)                             ' XMLDokument Laden
    Set cXMLNode = GetXMLNode(XMLDoc, "Root" & XML_PATH_SEP & "SAT") ' Hauptknoten <Root><SAT> auswählen
    Set XMLSQLNode = GetXmlChildNode(cXMLNode, szSQLName)           ' Childnode mit namen szSQLName auswählen
    GetDocSQL = GetXMLAttributByName(XMLSQLNode, "SQL")             ' Wert Attribut "SQL" holen
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetDocSQL", errNr, errDesc)
End Function

Public Function GetOptKategorieListFromXML(ByVal XMLDocPath As String) As String
' Liefert eine String Liste (; getrennt) <alle Kategorie>  von <Options>
    Dim XMLDoc As DOMDocument                                       ' XML Dokument
    Dim XMLTreeNode As IXMLDOMNode                                  ' Jeweiliger <Kategorie>
    Dim szResult As String                                          ' Ergebnis String mit ; getrent
    Dim cXMLNode As IXMLDOMNode                                     ' Hauptknoten <Options>
    Dim XMLNodeAtribute As IXMLDOMAttribute                         ' XML Node Atributliste
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If XMLDocPath = "" Then GoTo exithandler                        ' Kein Doc -> Raus
    Set XMLDoc = LoadXMLDoc(XMLDocPath)                             ' XMLDokument Laden
    Set cXMLNode = GetXMLNode(XMLDoc, "Options")                    ' Hauptknoten <Options> auswählen
    For Each XMLTreeNode In cXMLNode.childNodes                     ' Für alle Childnodes die Attributwerte von "Tag" auslesen
        szResult = szResult & GetXMLAttributByName(XMLTreeNode, "Name") & ";"
    Next                                                            ' Nächster Childnode
exithandler:
On Error Resume Next
    GetOptKategorieListFromXML = CutLastChar(szResult, ";")         ' Letztes ; abschneiden
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetOptKategorieListFromXML", errNr, errDesc)
End Function

Public Function GetRootNodeListFromXML(ByVal XMLDocPath As String) As String
' Liefert eine String Liste (; getrent) alle direkten unterkonten von <Root><Tree>
' Dies sind die Root Konzen iner TreeView ansicht
    Dim XMLDoc As DOMDocument                                       ' XML Dokument
    Dim XMLTreeNode As IXMLDOMNode                                  ' Jeweiliger <TreeNode>
    Dim szResult As String                                          ' Ergebnis String mit ; getrent
    Dim cXMLNode As IXMLDOMNode                                     ' HAuptknoten <Tree>
    Dim XMLNodeAtribute As IXMLDOMAttribute                         ' XML Node Atributliste
    Dim szChildAtrValue As String                                   ' Gesuchter Atribut wert
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If XMLDocPath = "" Then GoTo exithandler                        ' Kein Doc -> Raus
    Set XMLDoc = LoadXMLDoc(XMLDocPath)                             ' XMLDokument Laden
    Set cXMLNode = GetXMLNode(XMLDoc, "Root" & XML_PATH_SEP & "Tree") ' Hauptknoten <Root><Tree> auswählen
    For Each XMLTreeNode In cXMLNode.childNodes                     ' Für alle Childnodes die Attributwerte von "Tag" auslesen
        szChildAtrValue = GetXMLAttributByName(XMLTreeNode, "Tag")  ' Atribut wert des Childnode auslesen
        If szChildAtrValue <> "" Then                               ' Atribut Wert nicht leer
            szResult = szResult & GetXMLAttributByName(XMLTreeNode, "Tag") & ";" ' Ergebnisliste zusammen setzen
        End If
        szChildAtrValue = ""                                        ' Atributwert initialisieren
    Next                                                            ' Nächsten Child node
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    GetRootNodeListFromXML = CutLastChar(szResult, ";")             ' Letztes ; abschneiden
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "GetRootNodeListFromXML", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit exithandler
End Function

Public Function GetSubNodeListFromXML(ByVal XMLDocPath As String, ParentNodeTag As String) As String
    Dim XMLSubNodes As IXMLDOMNodeList                              ' Liste aller <Subnodes> von cXMLNode
    Dim szResult As String                                          ' Ergebnis String
    Dim cXMLNode As IXMLDOMNode                                     ' Akt. Node
    Dim XMLNodeAtribute As IXMLDOMAttribute                         ' XMLNode Attribute
    Dim i As Integer                                                ' Counter
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If XMLDocPath = "" Then GoTo exithandler                        ' Kein Doc -> Raus
    Set cXMLNode = GetTreeNodeFromXML(XMLDocPath, ParentNodeTag)
    Set XMLSubNodes = cXMLNode.selectNodes("SubNode")
    For i = 0 To XMLSubNodes.length - 1
        szResult = szResult & GetXMLAttributByName(XMLSubNodes(i), "Tag") & ";"
    Next i
exithandler:
On Error Resume Next
    GetSubNodeListFromXML = CutLastChar(szResult, ";")              ' Letztes ; abschneiden
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetSubNodeListFromXML", errNr, errDesc)
End Function

Public Function GetNoKontextListFromXML(ByVal XMLDocPath As String) As String
    Dim XMLDoc As DOMDocument                                       ' XML Dokument
    Dim cXMLNode As IXMLDOMNode                                     ' Akt. Node
    Dim XMLSubNodes As IXMLDOMNode                                  ' Liste aller <Subnodes> von cXMLNode
    Dim szResult As String                                          ' Ergebnis String
    Dim XMLNodeAtribute As IXMLDOMAttribute                         ' XMLNode Attribute
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If XMLDocPath = "" Then GoTo exithandler                        ' Kein Doc -> Raus
    Set XMLDoc = LoadXMLDoc(XMLDocPath)                             ' XMLDokument Laden
    Set cXMLNode = GetXMLNode(XMLDoc, "Root" & XML_PATH_SEP & "Tree") ' Hauptknoten <Root><Tree> auswählen
    For Each XMLSubNodes In cXMLNode.childNodes                     ' Für alle Childnodes die Attributwerte von "bShowKomtextMenue" auslesen
        If Not CBool(GetXMLAttributByName(XMLSubNodes, "bShowKomtextMenue")) Then
            If XMLSubNodes.baseName = "List" Then
                szResult = szResult & GetXMLAttributByName(XMLSubNodes, "Tag")
            Else
                szResult = szResult & GetXMLAttributByName(XMLSubNodes, "Name")
            End If
        End If
    Next
exithandler:
On Error Resume Next
    GetNoKontextListFromXML = CutLastChar(szResult, ";")            ' Letztes ; abschneiden
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetNoKontextListFromXML", errNr, errDesc)
End Function

Public Function GetOptionsNodeList(ByVal XMLDocPath As String, Optional KategorieName As String) As String
    Dim XMLDoc As DOMDocument                                       ' XML Dokument
    Dim cXMLNode As IXMLDOMNode                                     ' Akt. Root Node hier <Options>
    Dim cKatNode As IXMLDOMNode                                     ' Akt. Kategorie Node
    Dim cOptNode As IXMLDOMNode                                     ' Akt Options Node
    Dim szResult As String                                          ' Ergebnis String
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If XMLDocPath = "" Then GoTo exithandler                        ' Kein Doc -> Raus
    Set XMLDoc = LoadXMLDoc(XMLDocPath)                             ' XMLDokument Laden
    Set cXMLNode = GetXMLNode(XMLDoc, "Options")                    ' Hauptknoten <Options> auswählen
    If KategorieName <> "" Then                                     ' Wenn Kategorie name angegeben
        Set cKatNode = GetXmlChildNode(cXMLNode, "", "Name", KategorieName) ' Unterknoten mit att Name= KategorieName auswählen
        If Not cKatNode Is Nothing Then                             ' Wenn kategorie Node gefunden
            For Each cOptNode In cKatNode.childNodes                ' Alle unterkonten durchlaufen
                szResult = szResult & GetXMLAttributByName(cOptNode, "Name") & ";"    ' Attribut Name auslesen
            Next
        End If
    Else
        For Each cKatNode In cXMLNode.childNodes                    ' Alle Kategorien durchlaufen
            For Each cOptNode In cKatNode.childNodes                ' Alle unterkonten (Optionen) durchlaufen
                szResult = szResult & GetXMLAttributByName(cOptNode, "Name") & ";"   ' Attribut Name auslesen
            Next
        Next
    End If
    
exithandler:
On Error Resume Next
    GetOptionsNodeList = CutLastChar(szResult, ";")                 ' Letztes ; abschneiden
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetOptionsNodeList", errNr, errDesc)
End Function

Public Function GetSearchListFromXML(ByVal XMLDocPath As String) As String
    Dim XMLDoc As DOMDocument                                       ' XML Dokument
    Dim cXMLNode As IXMLDOMNode                                     ' Akt. Node
    Dim XMLSubNodes As IXMLDOMNode                                  ' Liste aller Subnodes von cXMLNode
    Dim XMLSubSubNodes As IXMLDOMNode                               ' Liste aller Subnodes von cXMLNode.Subnodes
    Dim szResult As String                                          ' ; getrente Stringliste
    Dim szTmp As String                                             ' Name der Akt. Suche
    Dim bInUserList As Boolean                                      ' Akt. Suche sol für Userauswahl bereitstehen
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If XMLDocPath = "" Then GoTo exithandler                        ' Kein Doc -> Raus
    Set XMLDoc = LoadXMLDoc(XMLDocPath)                             ' XMLDokument Laden
    Set cXMLNode = GetXMLNode(XMLDoc, "Root" & XML_PATH_SEP & "Search") ' Hauptknoten <Root><Search> auswählen
    For Each XMLSubNodes In cXMLNode.childNodes                     ' Für alle Childnodes von cXMLNode
        For Each XMLSubSubNodes In XMLSubNodes.childNodes           ' Für alle Childnodes von XMLSubNodes
            szTmp = GetXMLAttributByName(XMLSubSubNodes, "Name")    ' Atrribut "Name" auslesen
            If szTmp <> "" Then                                     ' Wert vorhanden
                bInUserList = CheckXMLValueForBool( _
                        GetXMLAttributByName(XMLSubSubNodes, _
                        "InUserList"), False)                       ' Prüfen ob in Liste angezeigt
                If bInUserList Then                                 ' Soll angezeigt werden
                    szResult = szResult & szTmp & ","               ' in diel Liste aufnehmen
                  
                End If
            End If
            szTmp = ""                                              ' Wert initialisieren
            bInUserList = False                                     ' initialisieren
        Next
    Next
exithandler:
On Error Resume Next
    GetSearchListFromXML = CutLastChar(szResult, ";")               ' Letztes ; abschneiden
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetSearchListFromXML", errNr, errDesc)
End Function

Public Function GetSearchSQLFromXML(ByVal XMLDocPath As String, szSearchRootKey As String) As String
' ermittelt das Arrtibut "SQL" des XML Nodes
' <Root><Search><szSearchRootKey>
    Dim XMLDoc As DOMDocument                                       ' XML Dokument
    Dim cXMLNode As IXMLDOMNode                                     ' Akt. Node
    Dim cXMLSeachNode As IXMLDOMNode                                ' SeachNode Node
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If XMLDocPath = "" Then GoTo exithandler                        ' Kein Doc -> Raus
    Set XMLDoc = LoadXMLDoc(XMLDocPath)                             ' XMLDokument Laden
    Set cXMLNode = GetXMLNode(XMLDoc, "Root" & XML_PATH_SEP & "Search") ' Hauptknoten <Root><Search> auswählen
    Set cXMLSeachNode = GetXmlChildNode(cXMLNode, szSearchRootKey)  ' Unterknoten = szSearchRootKey auswählen
    GetSearchSQLFromXML = GetXMLAttributByName(cXMLSeachNode, "SQL") ' Atribut SQL auslesen
exithandler:
On Error Resume Next
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetSearchSQLFromXML", errNr, errDesc)
End Function

Public Function GetSearchInfoFromXML(ByVal XMLDocPath As String, _
            szSearchRootKey As String, _
            szSeachfield As String, Optional szSQL As String, _
            Optional szOptWhere As String, _
            Optional szSuchname As String, _
            Optional Fieldname As String, _
            Optional bInUserList As Boolean) As String
' ermittelt das Arrtibut "Field" des XML Nodes
' <Root><Search><szSearchRootKey><szSeachfield>
    Dim XMLDoc As DOMDocument                                       ' XML Dokument
    Dim cXMLNode As IXMLDOMNode                                     ' Akt. Node
    Dim cXMLSeachNode As IXMLDOMNode                                ' SeachNode Node
    Dim cXMLSeachSubNode As IXMLDOMNode                             ' Childnaode von Searchnode
    Dim XMLNodeAtribute As IXMLDOMAttribute                     ' XML Node Atribut
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If XMLDocPath = "" Then GoTo exithandler                        ' Kein Doc -> Raus
    Set XMLDoc = LoadXMLDoc(XMLDocPath)                             ' XMLDokument Laden
    Set cXMLNode = GetXMLNode(XMLDoc, "Root" & XML_PATH_SEP & "Search") ' Hauptknoten <Root><Search> auswählen
    Set cXMLSeachNode = GetXmlChildNode(cXMLNode, szSearchRootKey)  ' Unterknoten = szSearchRootKey auswählen
    For Each XMLNodeAtribute In cXMLSeachNode.attributes            ' Alle Attribute duchlaufen
        Select Case XMLNodeAtribute.Name
        Case "SQL"                                                  ' Attibut SQL auslesen
            szSQL = XMLNodeAtribute.Value
        Case "OptWhere"                                             ' Attibut OptWhere auslesen
            szOptWhere = XMLNodeAtribute.Value
        Case Else
            
        End Select
    Next
    Set cXMLSeachSubNode = GetXmlChildNode(cXMLSeachNode, "", "Name", _
            szSearchRootKey & " " & szSeachfield)                   ' Unterknoten = szSearchRootKey auswählen
    For Each XMLNodeAtribute In cXMLSeachSubNode.attributes         ' Alle Attribute duchlaufen
        Select Case XMLNodeAtribute.Name
        Case "Name"
            szSuchname = XMLNodeAtribute.Value
        Case "Field"
            Fieldname = XMLNodeAtribute.Value
        Case "InUserList"
            bInUserList = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
        Case Else
         
        End Select
    Next
'    GetRealSearchFieldFromXML = GetXMLAttributByName(cXMLSeachSubNode, "Field") ' Atribut SQL auslesen
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetRealSearchFieldFromXML", errNr, errDesc)
End Function

Public Function GetTreeNodeFromXML(ByVal XMLDocPath As String, NodeTag As String) As IXMLDOMNode
' ermittelt aus <Root><Tree><Treenode> Childnode dessen Atrib. "Tag" dem Pfad NodeTag folgt
    Dim XMLDoc As DOMDocument                                       ' XML Dokument
    Dim XMLTreeNode As IXMLDOMNode                                  ' Jeweiliger <TreeNode>
    Dim szResult As String                                          ' Ergebnis String mit ; getrent
    Dim cXMLNode As IXMLDOMNode                                     ' HAuptknoten <Tree>
    Dim XMLNodeAtribute As IXMLDOMAttribute                         ' XML Node Atributliste
    Dim XMLRoot As IXMLDOMNodeList                                  ' Liste der XML Root Nodes
    Dim cXMLRootNode As IXMLDOMNode                                 ' Ergebis XML Node
    Dim szTagRest As String                                         ' Rest Tag zum suchen in Unter knoten
    Dim TagArray()  As String                                       ' Nodes Tag in Array augespalten
    Dim i As Integer                                                ' Counter
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    TagArray = Split(NodeTag, "\")                                  ' Nodetag aufspalten
    If XMLDocPath = "" Then GoTo exithandler                        ' Kein Doc -> Raus
    Set XMLDoc = LoadXMLDoc(XMLDocPath)                             ' XMLDokument Laden
    Set cXMLNode = GetXMLNode(XMLDoc, "Root" & XML_PATH_SEP & "Tree") ' Hauptknoten <Root><Tree> auswählen
    Set cXMLRootNode = GetXmlChildNode(cXMLNode, "TreeNode", "Tag", TagArray(0))    ' Childnode mit Attribut "Tag"=TagArray(0) auswählen
    If UBound(TagArray) > 0 Then                                    ' Wenn Tag abgearbeitet Fertig
        For i = 1 To UBound(TagArray)                               ' Sonst
            szTagRest = szTagRest & TagArray(i) & "\"               ' restlichen Tag wieder zusammensetzen
        Next i
        szTagRest = CutLastChar(szTagRest, "\")
        Set cXMLNode = FindXMLSubNode(cXMLRootNode, szTagRest)      ' in SubNode suchen
    Else
        Set cXMLNode = cXMLRootNode
    End If
    Set GetTreeNodeFromXML = cXMLNode                               ' Gefundenen Knoten zurück geben
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetTreeNodeFromXML", errNr, errDesc)
End Function

Public Function GetEditNodeFromXML(ByVal XMLDocPath As String, NodeName As String) As IXMLDOMNode
' Ermittelt Childnode von <Root><Edit> mit Atribut "Name" = Nodename
    Dim XMLDoc As DOMDocument                                       ' XML Dokument
    Dim XMLRoot As IXMLDOMNodeList                                  ' Liste der XML Root Nodes
    Dim cXMLNode As IXMLDOMNode                                     ' Aktueller XML Node
    Dim cXMLRootNode As IXMLDOMNode                                 ' Ergebis XML Node
    Dim XMLNodeAtribute As IXMLDOMAttribute                         ' XML Node Atribut
    Dim szTagRest As String                                         ' Rest Tag zum suchen in Unter knoten
    Dim TagArray()  As String                                       ' Nodes Tag in Array augespalten
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If XMLDocPath = "" Then GoTo exithandler                        ' Kein Doc -> Raus
    Set XMLDoc = LoadXMLDoc(XMLDocPath)                             ' XMLDokument Laden
    Set cXMLNode = GetXMLNode(XMLDoc, "Root" & XML_PATH_SEP & "Edit") ' Parentnode Ermitteln hier <Root><Edit>
    Set cXMLRootNode = GetXmlChildNode(cXMLNode, "", "Name", NodeName)
    Set GetEditNodeFromXML = cXMLRootNode                           ' Gefundenen Knoten zurück geben
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetEditNodeFromXML", errNr, errDesc)
End Function

Public Function GetOptionNodeFromXML(ByVal XMLDocPath As String, OptionsName As String) As IXMLDOMNode
' Ermittelt <Option> von <Options> mit Atribut "Name" = Nodename
    Dim XMLDoc As DOMDocument                                       ' XML Dokument
    Dim cXMLNode As IXMLDOMNode                                     ' Aktueller XML Node hier <Options>
    Dim KatNode As IXMLDOMNode                                      ' Aktueller Kategorie node
    Dim cXMLOptNode As IXMLDOMNode                                  ' Options Node
    Dim cResultNode As IXMLDOMNode                                  ' Ergebis Node
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    If XMLDocPath = "" Then GoTo exithandler                        ' Kein Doc -> Raus
    Set XMLDoc = LoadXMLDoc(XMLDocPath)                             ' XMLDokument Laden
    Set cXMLNode = GetXMLNode(XMLDoc, "Options")                    ' Parentnode Ermitteln hier <Options>
    For Each KatNode In cXMLNode.childNodes
        'For Each cXMLOptNode In KatNode.childNodes
        Set cResultNode = GetXmlChildNode(KatNode, "", "Name", OptionsName)
        If Not cResultNode Is Nothing Then Exit For
    Next
    Set GetOptionNodeFromXML = cResultNode                           ' Gefundenen Knoten zurück geben
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetOptionNodeFromXML", errNr, errDesc)
End Function

Public Function FindXMLSubNode(XMLParentNode As IXMLDOMNode, szTag As String) As IXMLDOMNode
' Ermittelt Childnode von XMLParentNode
    Dim cXMLRootNode As IXMLDOMNode                                 ' RootNode
    Dim cXMLNode As IXMLDOMNode                                     ' Aktueller XML Node
    Dim XMLNodeAtribute As IXMLDOMAttribute                         ' XML Node Atribut
    Dim szTagRest As String                                         ' Relevater rest vom Tag
    Dim TagArray()  As String                                       ' in Array aufgespaltener Tag
    Dim i As Integer                                                ' Counter
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    TagArray = Split(szTag, "\")                                    ' Tag in Array aufspalten
    ' Ermittle XMLParentNode.Childnode mit <SubNode> und Atrb. "Tag" = TagArray(0)
    Set cXMLRootNode = GetXmlChildNode(XMLParentNode, "SubNode", "Tag", TagArray(0))
    If UBound(TagArray) = 0 Then                                    ' Wenn Array Abgearbeitet
        Set cXMLNode = cXMLRootNode                                 ' Fertig
    Else                                                            ' Sonst
        For i = 1 To UBound(TagArray)                               ' Rest von Array ermitteln
            szTagRest = szTagRest & TagArray(i) & "\"
        Next i
        szTagRest = CutLastChar(szTagRest, "\")                     ' Letzten \ abschneiden
        Set cXMLNode = FindXMLSubNode(cXMLRootNode, szTagRest)      ' Rekursiver aufruf
    End If
    Set FindXMLSubNode = cXMLNode
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "FindXMLSubNode", errNr, errDesc)
End Function

Public Function GetEditRelInfoFromXML(ByVal XMLDocPath As String, NodeName As String, _
            RelName As String, szSQL As String, Optional szWhere As String, Optional lngImage As Integer, _
            Optional szDelFlag As String, Optional lngAltImage As Integer, Optional szAktImgField As String, _
            Optional szAltImgValue As String)
' Ermittelt Childnode von <Root><Edit> mit Atribut "Name" = Nodename
' dafon Childnode mit Attribut "Name"=RelName
    Dim newXMLNode As IXMLDOMNode                                   ' Edit Node
    Dim cXMLNode As IXMLDOMNode                                     ' Realtion Subnode
    Dim XMLNodeAtribute As IXMLDOMAttribute                         ' XML Node Atribut
    Dim szTmp As String
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    Set newXMLNode = GetEditNodeFromXML(XMLDocPath, NodeName)       ' Edit Node Ermitteln
    If newXMLNode Is Nothing Then GoTo exithandler                  ' Kein edit Node -> Fertig
    Set cXMLNode = GetXmlChildNode(newXMLNode, "", "Name", RelName) ' Subnode mit NAme = Relnode suchen
    If cXMLNode Is Nothing Then GoTo exithandler                    ' Keine Subnode gefunden -> Raus
    For Each XMLNodeAtribute In cXMLNode.attributes                 ' Alle Attribute duchlaufen
        Select Case XMLNodeAtribute.Name
        Case "SQL"                                                  ' Attibut SQL auslesen
            szSQL = XMLNodeAtribute.Value
        Case "WHERE"                                                ' Attibut WHERE auslesen
            szWhere = XMLNodeAtribute.Value
        Case "DelFlag"                                              ' Attibut DelFlag auslesen
            szDelFlag = XMLNodeAtribute.Value
        Case "AltImage"                                             ' Attibut AltImage auslesen
            lngAltImage = CheckXMLValueForNumeric(XMLNodeAtribute.Value, 1) ' Attributvalue auf Numerisch prüfen
        Case "AltImageField"                                        ' Attibut AltImageField auslesen
            szAktImgField = XMLNodeAtribute.Value
        Case "AltImageValue"                                        ' Attibut AltImageValue auslesen
            szAltImgValue = XMLNodeAtribute.Value
        Case "Image"                                                ' Attibut Image auslesen
            lngImage = CheckXMLValueForNumeric(XMLNodeAtribute.Value, 1) ' Attributvalue auf Numerisch prüfen
        Case Else
            
        End Select
    Next                                                            ' Nächstes Attibut
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetEditRelInfoFromXML", errNr, errDesc)
End Function

Public Function GetEditRelationList(ByVal XMLDocPath As String, NodeName As String) As String
    Dim newXMLNode As IXMLDOMNode                                   ' Edit Node
    Dim szRelList As String                                         ' ; getrente Liste aller realtionen
    Dim RelNode As IXMLDOMNode                                      ' Realtion Subnode
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Set newXMLNode = GetEditNodeFromXML(XMLDocPath, NodeName)       ' Edit Node Ermitteln
    If newXMLNode Is Nothing Then GoTo exithandler                  ' Kein edit Node -> Fertig
    For Each RelNode In newXMLNode.childNodes                       ' Alle Childnodes Duchlaufen
        If RelNode.baseName = "Relation" Then                       ' nur Relation Nodes
             'GetXMLAttributByName(RelNode,"Name")
            szRelList = szRelList & _
                    GetXMLAttributByName(RelNode, "Name") & ";"     ' Liste zusammen stellen
        End If
    Next
exithandler:
    GetEditRelationList = CutLastChar(szRelList, ";")               ' Letztes ; abscheiden und zurück
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetEditRelationList", errNr, errDesc)

End Function

Public Function GetEditInfoFromXML(ByVal XMLDocPath As String, NodeName As String, _
            szSQL As String, Optional szWhere As String, Optional lngImage As Integer, _
            Optional szDelFlag As String, Optional szIDorder As String, _
            Optional ExpertSQL As String)
' Liest attribute aus Edit knoten
    Dim newXMLNode As IXMLDOMNode                                   ' Edit Node
    Dim XMLNodeAtribute As IXMLDOMAttribute                         ' XML Node Atribut
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Set newXMLNode = GetEditNodeFromXML(XMLDocPath, NodeName)       ' Edit Node Ermitteln
    If newXMLNode Is Nothing Then GoTo exithandler                  ' Kein edit Node -> Fertig
    For Each XMLNodeAtribute In newXMLNode.attributes               ' Alle Attribute duchlaufen
        Select Case XMLNodeAtribute.Name
        'Case "NAme"
        '    newTag = XMLNodeAtribute.Value
        Case "IDOrder"
            szIDorder = XMLNodeAtribute.Value
        Case "DelFlag"
            szDelFlag = XMLNodeAtribute.Value
        'Case "bValueList"
        Case "ESQL"
            ExpertSQL = XMLNodeAtribute.Value
        Case "SQL"                                                  ' Attibut SQL auslesen
            szSQL = XMLNodeAtribute.Value
        Case "WHERE"                                                ' Attibut WHERE auslesen
            szWhere = XMLNodeAtribute.Value
        Case "Image"                                                ' Attibut Image auslesen
            lngImage = CheckXMLValueForNumeric(XMLNodeAtribute.Value, 1) ' Attributvalue auf Numerisch prüfen
        Case Else
            
        End Select
    Next                                                            ' Nächstes Attibut
        
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetEditInfoFromXML", errNr, errDesc)
End Function

Public Function GetDeleteInfoFromXML(ByVal XMLDocPath As String, NodeName As String, _
        bPosible As Boolean, DelSQL As String, Optional AskSql As String)
' Liest attribute aus Delete knoten (unterknoten von Edit)
    Dim newXMLNode As IXMLDOMNode                                   ' Edit Node
    Dim cXMLNode As IXMLDOMNode                                     ' Delete Subnode
    Dim XMLNodeAtribute As IXMLDOMAttribute                         ' XML Node Atribut
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Set newXMLNode = GetEditNodeFromXML(XMLDocPath, NodeName)       ' Edit Node Ermitteln
    If newXMLNode Is Nothing Then GoTo exithandler                  ' Kein edit Node -> Fertig
    Set cXMLNode = GetXmlChildNode(newXMLNode, "Delete")            ' Subnode Delete suchen
    If cXMLNode Is Nothing Then GoTo exithandler                    ' Keine Subnode gefunden -> Raus
    For Each XMLNodeAtribute In cXMLNode.attributes                 ' Alle Attribute duchlaufen
        Select Case XMLNodeAtribute.Name
        Case "Posible"                                              ' Attibut Posible auslesen
            bPosible = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
        Case "ASKSQL"                                               ' Attibut ASKSQL auslesen
            AskSql = XMLNodeAtribute.Value
        Case "DelSQL"                                               ' Attribut DelSQl auslesen
            DelSQL = XMLNodeAtribute.Value
        Case Else
            
        End Select
    Next                                                            ' Nächstes Attibut
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetDeleteInfoFromXML", errNr, errDesc)
End Function

Public Function GetOptionInfoFromXML(ByVal XMLDocPath As String, OptionsName As String, _
            Caption As String, Value As Variant, Optional bCrypt As Boolean, _
            Optional bEdit As Boolean, Optional bDisabled As Boolean, _
            Optional bPath As Boolean, Optional bFile As Boolean, _
            Optional bBool As Boolean, Optional bExpert As Boolean, _
            Optional szList As String)
' Liest attribute aus Options knoten
    Dim newXMLNode As IXMLDOMNode                                   ' Options Node
    Dim XMLNodeAtribute As IXMLDOMAttribute                         ' XML Node Atribut
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Set newXMLNode = GetOptionNodeFromXML(XMLDocPath, OptionsName)  ' Option Node ermitteln
    If newXMLNode Is Nothing Then GoTo exithandler
    If newXMLNode Is Nothing Then GoTo exithandler                  ' Kein edit Node -> Fertig
    For Each XMLNodeAtribute In newXMLNode.attributes               ' Alle Attribute duchlaufen
        Select Case XMLNodeAtribute.Name
        Case "Caption"
            Caption = XMLNodeAtribute.Value
        Case "Value"
            Value = XMLNodeAtribute.Value
        Case "bCrypt"
            bCrypt = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
        Case "bEdit"
             bEdit = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
        Case "bDisabled"
            bDisabled = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
        Case "bPath"
            bPath = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
        Case "bFile"
            bFile = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
        Case "bBool"
            bBool = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
        Case "bExpert"
            bExpert = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
        Case "List"
            szList = XMLNodeAtribute.Value
        Case Else
            
        End Select
    Next                                                            ' Nächstes Attibut
        
exithandler:
On Error Resume Next
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetOptionInfoFromXML", errNr, errDesc)
End Function

Public Function GetEditDefaultValue(ByVal XMLDocPath As String, NodeName As String, _
           Fieldname As String, Optional SQL As String)
' Ermittelt Defaultwert aus Edit Node für feldnamen
    Dim newXMLNode As IXMLDOMNode                                   ' Edit Node
    Dim cXMLNode As IXMLDOMNode
    Dim DefValXMLNode As IXMLDOMNode                                ' Default Value Node
    Dim XMLNodeAtribute As IXMLDOMAttribute                         ' XML Node Atribut
    Dim szTmp As String
    Dim bFound As Boolean                                           ' Node gefunden
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Set newXMLNode = GetEditNodeFromXML(XMLDocPath, NodeName)       ' Edit Node Ermitteln
    If newXMLNode Is Nothing Then GoTo exithandler                  ' Kein edit Node -> Fertig
    For Each cXMLNode In newXMLNode.childNodes                      ' Alle child Nodes duchlaufen
        If cXMLNode.baseName = "Defaultwert" Then                   ' Defaultwert node suchen
            For Each XMLNodeAtribute In cXMLNode.attributes         ' Alle Attribute duchlaufen
                Select Case XMLNodeAtribute.Name
                Case "Fieldname"                                    ' Wenn Attibut feldname
                    If UCase(XMLNodeAtribute.Value) = UCase(Fieldname) Then ' Feldname Gleich
                        Set DefValXMLNode = cXMLNode                ' Def Node merken
                        bFound = True                               ' Gefunden
                        Exit For                                    ' Raus
                    End If
                End Select
            Next
            If bFound Then Exit For                                 ' Gefunden -> Raus
        End If
    Next
    
    If bFound Then                                                  ' Wenn gefunden
        For Each XMLNodeAtribute In DefValXMLNode.attributes        ' Alle attribute duchlaufen
            Select Case XMLNodeAtribute.Name
            Case "SQL"                                              ' Attibut SQL auslesen
                SQL = XMLNodeAtribute.Value
            Case "Value"                                            ' Attribut Value auslesen
                GetEditDefaultValue = XMLNodeAtribute.Value
            End Select
        Next
    End If
    
exithandler:
On Error Resume Next
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetEditDefaultValue", errNr, errDesc)
End Function

Public Function GetEditValueType(ByVal XMLDocPath As String, NodeName As String, _
           Fieldname As String, IsBool As Boolean, _
           IsDate As Boolean, szList As String, _
           szSQLList As String) As Boolean
    Dim newXMLNode As IXMLDOMNode                                   ' Edit Node
    Dim cXMLNode As IXMLDOMNode                                     ' Child nodes
    Dim ValTypeXMLNode As IXMLDOMNode                               ' gesuchter ValueTyp Node
    Dim XMLNodeAtribute As IXMLDOMAttribute                         ' Node Atrribute
    Dim szTmp As String
    Dim bFound As Boolean                                           ' Gefunden Flag (es gibt einen Eitrag für dies Feld)
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Set newXMLNode = GetEditNodeFromXML(XMLDocPath, NodeName)       ' Edit Knoten ermitteln
    If newXMLNode Is Nothing Then GoTo exithandler                  ' Kein Edit Knoten -> Fertig
    For Each cXMLNode In newXMLNode.childNodes                      ' Alle child Nodes duchlaufen
        If cXMLNode.baseName = "Valuetype" Then                     ' Childnode "ValueType" Suchen
            For Each XMLNodeAtribute In cXMLNode.attributes         ' Attribute duchlaufen
                Select Case XMLNodeAtribute.Name
                Case "Fieldname"                                    ' Für Atrribut "Fieldname"
                    If UCase(XMLNodeAtribute.Value) = UCase(Fieldname) Then ' Ist unser gesuchtes Feld
                        Set ValTypeXMLNode = cXMLNode               ' Node Merken
                        bFound = True                               ' Gefunden Flag setzen
                        Exit For                                    ' Fertig im FOR
                    End If
                End Select
            Next                                                    ' Nächstes Attribut
            If bFound Then Exit For                                 ' Wenn gefunden fertig mit FOR
        End If
    Next                                                            ' Nächster Knoten
    If bFound Then                                                  ' Wir haben was Gefunden
        For Each XMLNodeAtribute In ValTypeXMLNode.attributes       ' Attribbute auswerten
            Select Case XMLNodeAtribute.Name
            Case "bBool"
                IsBool = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
            Case "bDate"
                IsDate = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
            Case "List"
                szList = XMLNodeAtribute.Value                      ' Attributvalue ist text
            Case "SQLList"
                szSQLList = XMLNodeAtribute.Value                   ' Attributvalue ist text
            End Select
        Next
        GetEditValueType = True                                     ' Erfolg zurück
    End If
    
exithandler:
On Error Resume Next
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetEditValueType", errNr, errDesc)
End Function

Public Function GetEditLockedField(ByVal XMLDocPath As String, NodeName As String, _
           Fieldname As String, Locked As Boolean, _
           Optional Enabled As Boolean, _
           Optional Visible As Boolean, _
           Optional Multiline As Integer, _
           Optional LockedWhenNew As Boolean) As Boolean
    Dim newXMLNode As IXMLDOMNode                                   ' Edit Node
    Dim cXMLNode As IXMLDOMNode                                     ' Child nodes
    Dim LockXMLNode As IXMLDOMNode                                  ' Unser gesucchter LockedField Node
    Dim XMLNodeAtribute As IXMLDOMAttribute                         ' Node Atrribute
    Dim szTmp As String
    Dim bFound As Boolean                                           ' Gefunden Flag (es gibt einen Eitrag für dies Feld)
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Set newXMLNode = GetEditNodeFromXML(XMLDocPath, NodeName)       ' Edit Knoten ermitteln
    If newXMLNode Is Nothing Then GoTo exithandler                  ' Kein Edit Knoten -> Fertig
    For Each cXMLNode In newXMLNode.childNodes                      ' Alle child Nodes duchlaufen
        If cXMLNode.baseName = "LockedField" Then                   ' Childnode "LockedField" Suchen
            For Each XMLNodeAtribute In cXMLNode.attributes         ' Attribute duchlaufen
                Select Case XMLNodeAtribute.Name
                Case "Fieldname"                                    ' Für Atrribut "Fieldname"
                    If UCase(XMLNodeAtribute.Value) = UCase(Fieldname) Then ' Ist unser gesuchtes Feld
                        Set LockXMLNode = cXMLNode                  ' Node Merken
                        bFound = True                               ' Gefunden Flag setzen
                        Exit For                                    ' Fertig im FOR
                    End If
                End Select
            Next                                                    ' Nächstes Attribut
            If bFound Then Exit For                                 ' Wenn gefunden fertig mit FOR
        End If
    Next                                                            ' Nächster Knoten
    If bFound Then                                                  ' Wir haben was Gefunden
        For Each XMLNodeAtribute In LockXMLNode.attributes          ' Attribbute auswerten
            Select Case XMLNodeAtribute.Name
            Case "Locked"
                Locked = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
            Case "Enabled"
                Enabled = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
            Case "Visible"
                Visible = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
            Case "Multiline"
                Multiline = CheckXMLValueForNumeric(XMLNodeAtribute.Value, 1) ' Attributvalue auf Numerisch prüfen
            Case "LockedNew"
                LockedWhenNew = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
            End Select
        Next
        GetEditLockedField = True                                   ' Erfolg zurück
    End If
    
exithandler:
On Error Resume Next
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetEditLockedField", errNr, errDesc)
End Function

Public Function GetLVInfoFromXML(ByVal XMLDocPath As String, NodeTag As String, _
        szSQL As String, newTag As String, szWhere As String, Optional newImage As Integer, _
        Optional bValueList As Boolean, Optional bSubNodes As Boolean, _
        Optional bEdit As Boolean, Optional bNew As Boolean, Optional bSelectNode As Boolean, _
        Optional lngAltImage As Integer, Optional szAktImgField As String, _
        Optional szAltImgValue As String, Optional szDelFlag As String, _
        Optional bShowkontextMenue As Boolean, Optional bDelete As Boolean)
    Dim newXMLNode As IXMLDOMNode
    Dim cXMLNode As IXMLDOMNode
    Dim XMLNodeAtribute As IXMLDOMAttribute
    Dim szTmp As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Set newXMLNode = GetTreeNodeFromXML(XMLDocPath, NodeTag)
    If newXMLNode Is Nothing Then GoTo exithandler
    For Each cXMLNode In newXMLNode.childNodes                      ' Alle root Nodes >TreeNode> duchlaufen
        If cXMLNode.baseName = "List" Then
            For Each XMLNodeAtribute In cXMLNode.attributes
                Select Case XMLNodeAtribute.Name
                Case "Tag"
                    newTag = XMLNodeAtribute.Value
                Case "Image"
                    newImage = CheckXMLValueForNumeric(XMLNodeAtribute.Value, 1) ' Attributvalue auf Numerisch prüfen
                Case "AltImage"                                     ' Attibut AltImage auslesen
                    lngAltImage = CheckXMLValueForNumeric(XMLNodeAtribute.Value, 1) ' Attributvalue auf Numerisch prüfen
                Case "AltImageField"                                ' Attibut AltImageField auslesen
                    szAktImgField = XMLNodeAtribute.Value
                Case "AltImageValue"                                ' Attibut AltImageValue auslesen
                    szAltImgValue = XMLNodeAtribute.Value
                Case "bValueList"
                    bValueList = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
                Case "DelFlag"
                    szDelFlag = XMLNodeAtribute.Value
                Case "SQL"
                    szSQL = XMLNodeAtribute.Value
                Case "WHERE"
                    szWhere = XMLNodeAtribute.Value
                Case "bListSubNodes"
                    bSubNodes = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
                Case "Edit"
                    bEdit = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
                Case "New"
                    bNew = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
                Case "SelectNode"
                    bSelectNode = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
                Case "bShowKontextMenue"
                    bShowkontextMenue = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
                Case "Delete"
                     bDelete = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
                Case Else
            
                End Select
                szTmp = ""
            Next
        End If
    Next
exithandler:
On Error Resume Next
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetLVInfoFromXML", errNr, errDesc)
End Function

Public Function GetTVNodeInfofromXML(ByVal XMLDocPath As String, NodeTag As String, _
        newTag As String, NewText As String, NewKey As String, bShowSubNodes As Boolean, _
        Optional szSQL As String, Optional szWhere As String, Optional newImage As Integer, _
        Optional bShowkontextMenue As Boolean, Optional szDescription As String)
    Dim newXMLNode As IXMLDOMNode
    Dim xmlChildNode As IXMLDOMNodeList
    Dim XMLNodeAtribute As IXMLDOMAttribute
    Dim szTmp As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Set newXMLNode = GetTreeNodeFromXML(XMLDocPath, NodeTag)
    If newXMLNode Is Nothing Then GoTo exithandler
    For Each XMLNodeAtribute In newXMLNode.attributes
        Select Case XMLNodeAtribute.Name
        Case "Tag"
            newTag = XMLNodeAtribute.Value
        Case "Key"
            NewKey = XMLNodeAtribute.Value
        Case "Text"
            NewText = XMLNodeAtribute.Value
        Case "Description"
            szDescription = XMLNodeAtribute.Value
        Case "Image"
            'newImage = XMLNodeAtribute.Value
            newImage = CheckXMLValueForNumeric(XMLNodeAtribute.Value, 1) ' Attributvalue auf Numerisch prüfen
        Case "bShowSubnodes"
            bShowSubNodes = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
        Case "bShowKontextMenue"
            bShowkontextMenue = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
        Case "SQL"
            szSQL = XMLNodeAtribute.Value
        Case "WHERE"
            szWhere = XMLNodeAtribute.Value
        Case Else
        
        End Select
    Next
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetTVNodeInfofromXML", errNr, errDesc)
End Function
                                                                    ' *****************************************
                                                                    ' INI Funktionen
Public Function GetINIValue(ByVal Path As String, ByVal Sect As String, ByVal Key As String) _
                             As String
  'Dim Result&, Buffer$
  Dim Buffer As String
  Dim Result As Long
  
    'Wert lesen
    Buffer = Space$(2000)
    Result = GetPrivateProfileString(Sect, Key, vbNullString, _
                                     Buffer, Len(Buffer), Path)
    GetINIValue = Left$(Buffer, Result)
End Function

Public Sub SetINIValue(ByVal Path As String, ByVal Sect As String, ByVal Key As String, _
                        ByVal Value As String)
  Dim Result As Long
    'Wert schreiben
    Result = WritePrivateProfileString(Sect, Key, Value, Path)
End Sub

Public Function SetINIArray(ByVal Path As String, ByVal Sect As String, xArray() _
                             As String)
  
  Dim x As Integer                                                  ' Counter
  Dim Buffer As String
  Dim Result As Long
  
    For x = LBound(xArray) To UBound(xArray)                        ' Feld in einen String mit Trennzeichen Chr$(0) umwandeln
      Buffer = Buffer & xArray(x) & Chr$(0)
    Next x

    Buffer = Left$(Buffer, Len(Buffer) - 1)                         ' String schreiben
    Result = WritePrivateProfileSection(Sect, Buffer, Path)
End Function

Public Sub DeleteINIKey(ByVal Path As String, ByVal Sect As String, ByVal Key As String)
  Call WritePrivateProfileString(Sect, Key, 0&, Path)
End Sub
 
Public Sub DeleteINISection(ByVal Path As String, ByVal Sect As String)
  Call WritePrivateProfileString(Sect, 0&, 0&, Path)
End Sub
                                                                    ' *****************************************
                                                                    ' Validierungen
Public Function CheckArray(VarArray As Variant) As Boolean
' Liefert True zurück wenn Array mind. ein Item enthält
    Dim i As Integer                                                ' Array Counter
On Error Resume Next                                                ' Hier keine Fehlerbehandlung
    i = UBound(VarArray)                                            ' Array Count ermitteln
    If Err.Number <> 0 Then                                         ' Fehler abfragen
        Err.Clear                                                   ' Fehler resetten
        CheckArray = False                                          ' Falls zurück
        Exit Function                                               ' fertig
    End If
    CheckArray = True                                               ' Sonst True zurück
    Err.Clear                                                       ' Evtl. error clearen
End Function

Public Function ArrayCount(VarArray As Variant) As Integer
' Liefert Ubound(VarArray) zurück
' eleminiert Fehler wenn kein item vorhanden
    Dim i As Integer                                                ' Array Counter
On Error Resume Next                                                ' Hier keine Fehlerbehandlung
    i = UBound(VarArray)                                            ' Array Count ermitteln
    If Err.Number > 0 Then                                          ' Fehler abfragen
        Err.Clear                                                   ' Fehler resetten
        ArrayCount = -1                                             ' Array enthält kein item
    Else
        ArrayCount = i                                              ' Array Count zurück
    End If
    Err.Clear                                                       ' Evtl. error clearen
End Function

Public Function AddInArray(VarArray As Variant, bPreserve As Boolean) As Boolean

On Error Resume Next                                                ' Hier keine Fehlerbehandlung
    If bPreserve Then                                               ' Werte erhalten
        ReDim Preserve VarArray(UBound(VarArray) + 1)               ' Array Hochzählen und werte erhalten
    Else                                                            ' Sonst
        ReDim VarArray(UBound(VarArray) + 1)                        ' Nur holzählen
    End If
    
    If Err.Number <> 0 Then                                         ' Wenn Fehler
        ReDim VarArray(0)                                           ' 1. array Item
        Err.Clear                                                   ' Fehler Resetten
    End If
End Function

Public Function CheckNull(ReturnValue As Variant, Optional DefaultReturn As Variant) As Variant
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If IsMissing(DefaultReturn) Then DefaultReturn = ""
    If IsEmpty(DefaultReturn) Then DefaultReturn = ""
    If IsNull(ReturnValue) Then
        CheckNull = DefaultReturn
    Else
        CheckNull = ReturnValue
    End If

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "CheckNull", errNr, errDesc)
End Function


Public Function IsIP(szIP As String) As Boolean
    Dim testArray() As String
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    IsIP = False
    If Trim(szIP) = "" Then GoTo exithandler
    If InStr(szIP, ".") = 0 Then GoTo exithandler
    testArray = Split(szIP, ".")
    If UBound(testArray) < 3 Or UBound(testArray) > 3 Then GoTo exithandler
   
   If IsNumeric(testArray(0)) And IsNumeric(testArray(1)) And IsNumeric(testArray(2)) _
            And IsNumeric(testArray(3)) Then IsIP = True
       
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ExecVerwalten", errNr, errDesc)
    Resume exithandler
End Function
                                                                    ' ***********************************************
                                                                    ' WMI
Public Function GetIPfromNameWMI(szCompName As String) As String
    Dim objPing As Object
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Set objPing = GetObject("winmgmts:Win32_PingStatus.address='" & szCompName & "'")
    DoEvents
    If IsNull(objPing.StatusCode) Or objPing.StatusCode <> 0 Then
    Else
        GetIPfromNameWMI = objPing.ProtocolAddress
    End If
exithandler:
On Error Resume Next
    Set objPing = Nothing
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetIPfromNameWMI", errNr, errDesc)
    Resume exithandler
End Function

Public Function SetWMIServiceObj(Optional szCompName As String, _
                    Optional szDomain As String, _
                    Optional szUser As String, _
                    Optional szPWD As String) As Boolean
    Dim szConStr As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If Trim(szCompName) = "" Then szCompName = "."
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    If Trim(szDomain) = "" Or Trim(szUser) = "" Then
        szConStr = ""
        szPWD = ""
            Set objWMIService = objLocator.ConnectServer(szCompName, "root\cimv2")
    Else
        szConStr = "szDomain + " \ " + szUser"
        Set objWMIService = objLocator.ConnectServer(szCompName, "root\cimv2", szDomain + "\" + szUser, szPWD)
    End If
    objWMIService.Security_.ImpersonationLevel = 3
    SetWMIServiceObj = True
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    If errNr = -2147023174 Then
        SetWMIServiceObj = False
        Resume exithandler
    Else
        Call objError.Errorhandler(MODULNAME, "SetWMIServiceObj", errNr, errDesc)
        Resume exithandler
    End If
End Function

Public Function GetUserOnCompWMI(szCompName As String) As String
    Dim objServ As Object
    Dim objCompSys As Object
    Dim objCompSysSet As Object
    Dim szUsername As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If Not PingWMI(szCompName) Then GoTo exithandler
    If Not SetWMIServiceObj(szCompName) Then GoTo exithandler
    Set objCompSysSet = objWMIService.InstancesOf("Win32_ComputerSystem")
    On Error Resume Next
    Set objCompSys = objWMIService.Get("Win32_ComputerSystem='" & szCompName & "'")
    If Err.Number <> 0 Then
        Err.Clear
        GoTo exithandler
    End If
    On Error GoTo Errorhandler
    'If objCompSysSet.Count > 0 Then
'        For Each objCompSys In objCompSysSet
            DoEvents
            szUsername = CheckNull(objCompSys.USERName, "")
 '       Next
    'End If
    GetUserOnCompWMI = szUsername
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetUserOnCompWMI", errNr, errDesc)
    Resume exithandler
End Function

Public Function DoShutdownWMI(szCompName As String, _
            Optional bReboot As Boolean, _
            Optional bLogOut As Boolean)
    Dim objOSSet As Object
    'Dim szCompname As String
    Dim objOS As Object
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If szCompName = "" Then GoTo exithandler
    If Not PingWMI(szCompName) Then GoTo exithandler
'    szCompname = varcomp.GetValueByName("Caption").szValue
    ' --- Bindung an Menge der Betriebssysteme
    Set objOSSet = GetObject("winmgmts:{impersonationLevel=impersonate,(RemoteShutdown)}//" _
            & szCompName & "").ExecQuery("select * from Win32_OperatingSystem where Primary=true")
    ' --- Neustart "aller" BS (ist aber immer nur eins!)
    For Each objOS In objOSSet
        If bLogOut Then     ' Abmelden
            objOS.win32shutdown 0, 0
        End If
        If bReboot Then     ' Neustarten
            objOS.reboot
        Else                ' Herunterfahren
            objOS.Shutdown
        End If
    Next
exithandler:
On Error Resume Next
    Set objOSSet = Nothing
    Set objOS = Nothing
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "DoShutdownWMI", errNr, errDesc)
    Resume exithandler
End Function

Public Function PingWMI(szCompName As String) As Boolean
    Dim objPing As Object
    Dim ticks As Integer
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Set objPing = GetObject("winmgmts:Win32_PingStatus.address='" & szCompName & "'")
    DoEvents
    'say objPing.StatusCode
    If IsNull(objPing.StatusCode) Or objPing.StatusCode <> 0 Then
        PingWMI = False
    Else
        PingWMI = True
    End If
      ' objPing.ProtocolAddress = IP
      '?objPing.ResponseTime
exithandler:
On Error Resume Next
    Set objPing = Nothing
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "PingWMI", errNr, errDesc)
    Resume exithandler
End Function

Public Function Ping(ByVal sHost As String) As Single
' Server anpingen und Reaktionszeit zurückgeben
  Dim QI As QOCINFO
  Dim vTime As Single
  Dim ticks1 As Double
  Dim ticks2 As Double

On Error Resume Next

  QI.dwSize = Len(QI)
  vTime = Timer
  ticks1 = GetTickCount
  If IsDestinationReachable(sHost, QI) = 1 Then
    DoEvents
    Ping = Round(Timer - vTime, 5)
    ticks2 = GetTickCount
    If ticks2 = ticks1 Then ticks2 = ticks1 + 10
    Ping = (ticks2 - ticks1) / 1000
  Else
    Ping = -1
  End If
  'ticks = (GetTickCount - ticks) / 1000
End Function
                                                                    ' ***********************************************
                                                                    ' Prog. Starten ausführen
Public Function ExecRemoteDesk(Computername As String, _
        Optional RemoteDesktopPath As String, _
        Optional VNCPath As String)
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If RemoteDesktopPath = "" And VNCPath = "" Then GoTo exithandler
    If RemoteDesktopPath <> "" Then
        Call ShellExec(RemoteDesktopPath, "/v " & Computername, vbNormalFocus)
    ElseIf VNCPath <> "" Then
        Call ShellExec(VNCPath, Computername, vbNormalFocus)
    End If
exithandler:
On Error Resume Next
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ExecRemoteDesk", errNr, errDesc)
    Resume exithandler
End Function

Public Function ExecSysInfo(Optional szComputername As String)
    '\\olg_sl_ws337\c$\Programme\Gemeinsame Dateien\Microsoft Shared\MSInfo
      Dim bOnline As Boolean
      Dim szcmd As String
      Dim szArgs As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    bOnline = PingWMI(szComputername)
    If Not bOnline Then GoTo exithandler
    'szCMD = "\\" & szComputername & "\C:\Programme\Gemeinsame Dateien\Microsoft Shared\MSInfo\MSInfo32.exe"
    'szCMD = "C:\Programme\Gemeinsame Dateien\Microsoft Shared\MSInfo\MSInfo32.exe"
    szcmd = objObjectBag.GetMSInfoExe
    If szComputername <> "" Then szArgs = "/computer " & szComputername
    ExecSysInfo = ShellExec(szcmd, szArgs, vbNormalFocus)
exithandler:
On Error Resume Next
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ExecSysInfo", errNr, errDesc)
    Resume exithandler
End Function

Public Function ExecVerwalten(szCompName As String)
'C:\WINDOWS\system32\mmc.exe C:\WINDOWS\system32\compmgmt.msc /computer=OLG_SL_WS337
    'Dim VarNet As clsValuelist
    Dim szcmd As String
    Dim szArgs As String
    Dim bOnline As Boolean
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    bOnline = PingWMI(szCompName)
    If Not bOnline Then GoTo exithandler
    'szCMD = "C:\WINDOWS\system32\mmc.exe"
    szcmd = objObjectBag.GetSys32Dir & "\mmc.exe"
    'szArgs = "C:\WINDOWS\system32\compmgmt.msc " _
        & "/computer=" & szCompname
    szArgs = objObjectBag.GetSys32Dir & "\compmgmt.msc " _
        & "/computer=" & szCompName
    ExecVerwalten = ShellExec(szcmd, szArgs, vbNormalFocus)
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ExecVerwalten", errNr, errDesc)
    Resume exithandler
End Function

Public Function ExecExplorer(szPath As String, Optional szSelectFile As String) As Long
' MW 01.10.09 öffnet den WinExplorer mit dem angegebenen verz.
    Dim ShellResult
    Dim szWinfolder As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
  'ExecExplorer = ShellExec("explorer.exe", szPath, vbNormalFocus)
    szPath = Trim(szPath)                                           ' Pfad angabe trimmen
    If szSelectFile <> "" Then szSelectFile = Trim(szSelectFile)    ' Datei trimmen
    szWinfolder = objObjectBag.GetWinDir                            ' Windows verz ermitteln
    If Right(szWinfolder, 1) <> "\" Then szWinfolder = szWinfolder & "\"
    If szSelectFile = "" Then                                       ' keine bestimmte Datei ?
        ShellResult = ShellExec(szWinfolder & "explorer.exe", szPath, vbNormalFocus)
        'ShellResult = Shell(szWinfolder & "explorer.exe " & szPath, vbNormalFocus)
    Else                                                            ' Sonst
        ShellResult = ShellExec(szWinfolder & "explorer.exe", "/select," & szSelectFile, vbNormalFocus)
       ' ShellResult = Shell(szWinfolder & "explorer.exe /select," & szSelectFile, vbNormalFocus)
    End If
exithandler:
On Error Resume Next
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ExecExplorer", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit exithandler
End Function

Public Function ExecEditor(szFilePath As String, Optional szOptEditorPath As String) As Long
' MW 01.10.09 öffnet den WinExplorer mit dem angegebenen verz.
    Dim ShellResult
    Dim szWinfolder As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szFilePath = Trim(szFilePath)                                   ' Pfad angabe trimmen
    szWinfolder = objObjectBag.GetWinDir                            ' Windows verz ermitteln
    If Right(szWinfolder, 1) <> "\" Then szWinfolder = szWinfolder & "\"
    If szOptEditorPath = "" Then
         ShellResult = ShellExec(szWinfolder & "Notepad.exe", szFilePath, vbNormalFocus)
    Else
    
    End If
exithandler:
On Error Resume Next
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ExecExplorer", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit exithandler
End Function

Public Sub HTMLHelp_ShowTopic(szHelpFile As String, _
            Optional ByVal szTopicFile As String)
    Dim sHelpFile As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    ' Pfad zur HTML-Hilfe
    If szTopicFile = "" Then
        ' Startseite anzeigen
        HtmlHelp 0, szHelpFile, HH_DISPLAY_TOPIC, ByVal 0&
    Else
        ' belibiege Seite anzeigen
        HtmlHelpTopic 0, szHelpFile, HH_DISPLAY_TOPIC, _
                szTopicFile
  End If
exithandler:
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "HTMLHelp_ShowTopic", errNr, errDesc)
    Resume exithandler
End Sub

Public Function RunShell(szcmd As String, szArgs As String, Optional bHide As Boolean)
    Dim lngWindowsStyle As Integer
    lngWindowsStyle = 1
    If bHide Then lngWindowsStyle = 0
    Call Shell(szcmd & " " & szArgs, vbNormalFocus)
End Function

Public Function ShellRun(szcmd As String, szArgs As String, _
        Optional bHide As Boolean) As Integer
    Dim wshshell As Object
    Dim ret As Integer
    Dim lngWindowsStyle As Integer                                  ' Fenster Style (Minimiert oder Normal)
    Dim szExecString As String                                      ' Befehlszeilen Szting (evtl. mit Argumenten)
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If InStr(szcmd, " ") > 0 Then szcmd = Chr(34) & szcmd & Chr(34)
    lngWindowsStyle = 1                                             ' Fenster anzeigen
    If bHide Then lngWindowsStyle = 0                               ' Wenn bHide Fenster verstecken
    If szArgs <> "" Then                                            ' Sind argumente angegeben
        szExecString = szcmd & " " & szArgs                         ' cmd & Argumente in einen String
    Else                                                            ' Sonst
        szExecString = szcmd                                        ' Lassen wie es ist
    End If
    Set wshshell = CreateObject("WScript.Shell")                    ' Script Shell holen
    wshshell.Run szExecString, lngWindowsStyle                      ' Ausführen
    ShellRun = ret
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Set wshshell = Nothing                                          ' WScript.Shell killen
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ShellRun", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function
    
Public Function ShellExec(szExecPath As String, _
            Optional szArgs As String, _
            Optional Mode As VbAppWinStyle) As Long
    Dim ProcHWND As Long                                            ' Fenster Handel des ausgeführten Programms
    Dim ProcHWN As Long                                             ' Temp Fenster Handel
    Dim RetVal
    Dim szExecString As String                                      ' Befehlszeilen Szting (evtl. mit Argumenten)
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If Trim(szArgs) <> "" Then                                      ' Sind Argumente angegeben
        szExecString = szExecPath & " " & Trim(szArgs)              ' Dann anhängen
    Else                                                            ' Sonst
        szExecString = szExecPath                                   ' Ohne Argumente
    End If
    ProcHWND = Shell(szExecString, Mode)                            ' Ausführen und Fenster Handel ermitteln
    RetVal = FindWindow(vbNullString, vbNullString)
    Do While RetVal <> 0
        If GetParent(RetVal) = 0 Then
            Call GetWindowThreadProcessId(RetVal, ProcHWN)
            If ProcHWN = ProcHWND Then
                ShellExec = RetVal
                Exit Do
            End If
        End If
On Error Resume Next                                                ' fehlerbehandlung deaktivieren
        RetVal = GetWindow(RetVal, GW_HWNDNEXT)
On Error GoTo Errorhandler                                          ' Fehlerbehandlung wieder aktivieren
    Loop
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ShellExec", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function
                                                                    ' **********************************************
                                                                    ' System Enviroment
Public Function GetCurCompName() As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Dim Buffer As String, CName As String
    Dim Result As Long, l As Long
    l = 15 + 1
    Buffer = Space$(l)
    Result = GetComputerName(Buffer, l)
    If Result = 1 Then
        CName = Left$(Buffer, InStr(1, Buffer, Chr$(0)) - 1)
        Buffer = Trim(Buffer)
        If Asc(Right(Buffer, 1)) = 0 Then Buffer = Left(Buffer, Len(Buffer) - 1)
        GetCurCompName = Trim(Buffer)
    End If
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetCurCompName", errNr, errDesc)
    Resume exithandler
End Function

Public Function getLokalUserName() As String
  Dim strName   As String
  Dim nSize     As Long
  Dim lngResult As Long

  nSize = 100
  strName = Space$(100)

  lngResult = GetUserName(strName, nSize)
  If lngResult <> 0 Then
    getLokalUserName = Left$(strName, nSize - 1)
  End If
End Function

Public Function FormatWMIDate(szWMIDate As String) As Date
    
    Dim szYear, szMonth, szDay, szHour, szMin As String
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    szYear = Left(szWMIDate, 4)
    szMonth = Mid(szWMIDate, 5, 2)
    szDay = Mid(szWMIDate, 7, 2)
    szHour = Mid(szWMIDate, 9, 2)
    szMin = Mid(szWMIDate, 11, 2)
    FormatWMIDate = CDate(szDay & "." & szMonth & "." & szYear & " " & szHour & ":" & szMin)
    
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "FormatWMIDate", errNr, errDesc)
    Resume exithandler
End Function

Public Function SetMenueColor(F As Object)
       'Farbe der Menüleiste setzen (hier: rot)
         Set_MenuColor mMenuBarColor, F.hWnd, vbBlack
   'Farbe der Menüelemente des 1. Menü setzen (hier: blau)
  Set_MenuColor mMenuColor, F.hWnd, vbBlack, 1, False
   'Farbe des Systemmenüs setzen (hier: gelb)
  Set_MenuColor mSysMenuColor, F.hWnd, vbBlack

End Function

Private Function SplitDateTimeBias(leasedate As String, _
                                   leasedatepart As String, _
                                   leasetimepart As String) As Long
  'takes a datetime returned
  'and splits out the date and time
  'components, returns them in the
  'leasedatepart and leasetimepart
  'passed variables, and returns the
  'bias to be applied to the resultant date.
    Dim pos     As Long
    Dim bias    As Long
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    pos = InStr(leasedate, ".")
    If pos > 0 Then
    bias = StripTimeZoneBias(leasedate)
        leasedatepart = Left$(leasedate, 8)
        leasetimepart = Mid$(leasedate, 9, pos - Len(leasedatepart) - 1)
        leasedatepart = InsertInString(leasedatepart, "-", 5, "")
        leasedatepart = InsertInString(leasedatepart, "-", 8, "")
        leasetimepart = InsertInString(leasetimepart, ":", 3, "")
        leasetimepart = InsertInString(leasetimepart, ":", 6, "")
        SplitDateTimeBias = bias
   Else

   End If
exithandler:
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "SplitDateTimeBias", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Private Function StripTimeZoneBias(leasedate As String) As Long
    Dim pos As Long
    Dim tmp As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    pos = InStr(leasedate, "-")
    If pos = 0 Then
        pos = InStr(leasedate, "+")
        If pos = 0 Then
            StripTimeZoneBias = 0
        Else
        End If
    Else
        tmp = Mid$(leasedate, pos, Len(leasedate))
        leasedate = Mid$(leasedate, 1, pos - 1)
        StripTimeZoneBias = CLng(tmp)
    End If
exithandler:
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "StripTimeZoneBias", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Function ConvertDate(dtb) As String
   Dim d As String
   Dim t As String
   Dim bias As Long
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
   If Not (dtb = "Not Available") Then
      bias = SplitDateTimeBias(CStr(dtb), d, t)
      ConvertDate = Format$(d, "dddd mmm d, yyyy") & " at " & _
                    Format$(t, "hh:mm:ss")
   Else
      ConvertDate = "Not Available"
   End If
               
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ConvertDate", errNr, errDesc)
    Resume exithandler
End Function
                                                                    '************************************************
                                                                    ' String Funktionen
Public Function TextTo(szText As String, Optional lngType As Integer) As String
On Error Resume Next                                                ' Hier keine Felhler behandlung
Select Case lngType
    Case 1                                                          ' leet
        TextTo = TextTo1337(szText)
    Case 2                                                          ' Hex
        TextTo = TextToHex(szText)
    Case 3                                                          ' Binär
        TextTo = TextToBin(szText)
    Case 4                                                          ' Morsezeichen
        TextTo = TextToMorse(szText)
    Case 5                                                          ' HTTP Sysntax
        TextTo = SetHTTPSyntax(szText)
    Case 6                                                          ' XML Syntax
        TextTo = SetXMLSyntax(szText)
    Case Else
        TextTo = TextTo1337(szText)
    End Select
    Err.Clear                                                       ' Evtl. error clearen
End Function
                                                                
Private Function InsertInString(ByVal sOriginal As String, _
                                sReplace As String, _
                                nField As Long, _
                                sDelimeter As String) As String
'Replaces or inserts a string into a (any char) delimeted string
    Dim lnCount As Long
    Dim lnStart As Long
    Dim lnLast As Long
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Do While InStr(lnStart + 1, sOriginal, sDelimeter) > 0
        lnStart = InStr(lnStart + 1, sOriginal, sDelimeter)
        lnCount = lnCount + 1
        If lnCount >= nField Then
            Exit Do
        End If
        lnLast = lnStart
    Loop
    Select Case lnCount
    Case 1
        InsertInString = sReplace & Mid$(sOriginal, lnStart)
    Case Is >= nField
        InsertInString = Mid$(sOriginal, 1, lnLast) & _
                sReplace & Mid$(sOriginal, lnStart)
    Case Else
        InsertInString = sOriginal & _
            String$((nField - 1) - lnCount, sDelimeter) & sReplace
    End Select
exithandler:
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "InsertInString", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function CheckStringOperation(szOperation As String, Value1 As Variant, Value2 As Variant) As Boolean
'Vergleicht Value1 mit Value2 und verwendet szOperator als Vergleich Operation
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szOperation = Trim(szOperation)                                 ' Sicherheitshalber trimmen
    If szOperation = "" Then szOperation = "="                      ' Keine Operation angegeben -> Fertig
    Select Case szOperation
    Case "="                                                        ' Ist Gleich
        If Value1 = Value2 Then CheckStringOperation = True
    Case ">"                                                        ' Größer als
        If Value1 > Value2 Then CheckStringOperation = True
    Case "<"                                                        ' Kleiner als
        If Value1 < Value2 Then CheckStringOperation = True
    Case ">="                                                       ' Größer gleich
        If Value1 >= Value2 Then CheckStringOperation = True
    Case "<="                                                       ' Kleiner Gleich
        If Value1 <= Value2 Then CheckStringOperation = True
    Case Else                                                       ' Nicht unterstützte Operatoren

    End Select
exithandler:
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "CheckStringOperation", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function CheckLastChar(szStr As String, szChar As String) As Boolean
' Überprüft Letztes Vorkommendes Zeichen on szStr = szCahr ist
    Dim intCharLen As Integer                                       ' Prüf Zeichen länge
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    intCharLen = Len(szChar)                                        ' Länge Prüf zeichen bestimmen
    If Right(szStr, intCharLen) = szChar Then CheckLastChar = True  ' Wenn Prüf zeichen vorkommt -> True zurück
exithandler:
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "CheckLastChar", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function CutLastChar(szStr As String, szChar As String) As String
' Schneidet das/die Zeichen szChar vom szStr ab sofer diese Hinten stehen
    Dim intCharLen As Integer                                       ' Prüf Zeichen länge
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    intCharLen = Len(szChar)                                        ' Länge Prüf zeichen bestimmen
    If CheckLastChar(szStr, szChar) Then                            ' Prüfen ob Prüf zeichen vorkommt
        szStr = Left(szStr, Len(szStr) - intCharLen)                ' Abschneiden
    End If
    CutLastChar = szStr                                             ' Ergebnis zurück
exithandler:
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "CutLastChar", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function SetPreZero(NumericText As String, lngLen As Integer) As String
' Stellt eine 0 vor den (numerischen) Text
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If lngLen = 0 Then lngLen = 2                                   ' Min 2 sonst aufruf der Fkt nicht sinnvoll
    NumericText = Trim(NumericText)                                 ' Sicherheitshalber trimmen
    If NumericText = "" Then GoTo exithandler                       ' Kein Text -> fertig
    If Len(NumericText) > lngLen Then                               ' Text schon lang genug
        SetPreZero = NumericText                                    ' Text zurück geben
        GoTo exithandler                                            ' -> Fertig
    End If
    While Len(NumericText) < lngLen                                 ' Solange Text zu kurz
        NumericText = "0" & NumericText                             ' Null Davor schreiben
    Wend
    SetPreZero = NumericText                                        ' Text zurück geben
exithandler:
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "SetPreZero", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ArabicToRomanInt(ArabicInt As Integer) As String
' Übersetzt Arabische Zahlen in Römische I,II,III usw
' Zeichen     I   V   X   L   C   D   M
' Wert        1   5   10  50  100 500 1000
' Umrechnung erfolg durch Teilem mit Rest vo der Hösten zur niedrigsten Zahl
' Um hier keine Endlosen Zeichenfolgen zu Prod. wird hier nur int verwendet
    Dim lngRest As String                                           ' Noch umzuwandlender Rest
    Dim lngFactor As Integer                                        ' Anz der Römischen Zeichen
    Dim szResult As String                                          ' Ergebis String
    Dim szRom() As String                                           ' Array mit Römischen Ziffern
    Dim szArab() As String                                          ' Array Mit Arabischen Ziffern
    Dim i As Integer                                                ' Counter
    Dim n As Integer                                                ' Noch ein Counter
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If ArabicInt > 100000 Then GoTo exithandler                     ' > 100.000 -> Fertig
    lngRest = ArabicInt
    szRom = Split("M;D;C;L;X;IX;VIII;VII;VI;V;IV;III;II;I", ";")    ' Römische Ziffern in Array
    szArab = Split("1000;500;100;50;10;9;8;7;6;5;4;3;2;1", ";")     ' Arabische Ziffern in Array
    For i = 0 To UBound(szRom)                                      ' Beide Arrays durchlaufen
        If lngRest >= CInt(szArab(i)) Then                          ' Wird erg. > 1 ?
            If lngRest > 15 Then                                    ' Wenn > 15
                lngFactor = Int(lngRest / CInt(szArab(i)))
                lngRest = lngRest - (lngFactor * CInt(szArab(i)))
                For n = 1 To lngFactor
                    szResult = szResult & szRom(i)
                Next n
            Else                                                    ' Wenn <= 15
                If CInt(szArab(i)) <= lngRest Then
                    lngRest = lngRest - CInt(szArab(i))
                    szResult = szResult & szRom(i)
                End If
            End If
        End If
    Next i
    ArabicToRomanInt = szResult                                     ' Ergebnis zurück
exithandler:
    Exit Function                                                   ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ArabicToRomanInt", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function
                                                                    '************************************************
                                                                    ' String Funktionen
Public Function SetWindowTransparency(F As Object, sinPercent As Single)
' Setzt Fenster transparenz
' funktioniert nur unter Windows 2000 oder XP!!!
' RateOfT: 254 = normal 0 = ganz transparent (also unsichtbar)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call SetFormTransparency(F, sinPercent)
    Err.Clear                                                       ' Evtl. error clearen
End Function

Public Function FormTo1337(F As Object) As Boolean
    Dim Ctl As Control                                              ' Aktuelles Control
    Dim frm As Form                                                 ' Form referenz
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If F Is Nothing Then GoTo exithandler                           ' Kein form -> fertig
    Set frm = F
    frm.Caption = TextTo1337(frm.Caption)
    For Each Ctl In frm.Controls                                    ' Alle Controls duchgehen
        If CtlIsTextBox(Ctl) Then                                   ' Ctl ist Textbox
           Ctl.Text = TextTo1337(Ctl.Text)
        End If
        If CtlIsLabel(Ctl) Then                                     ' Ctl ist Label
           Ctl.Caption = TextTo1337(Ctl.Caption)
        End If
        If CtlIsButton(Ctl) Then                                    ' Ctl ist Commandbar
            Ctl.Caption = TextTo1337(Ctl.Caption)
        End If
        If CtlIsCheck(Ctl) Then                                     ' Ctl is Checkbox
            Ctl.Caption = TextTo1337(Ctl.Caption)
        End If
    Next                                                            ' Nächstes Control
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "FormToConsole", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function FormToConsole(F As Object) As Boolean
    Dim Ctl As Control                                              ' Aktuelles Control
    Dim frm As Form                                                 ' Form referenz
    Dim il As Image
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If F Is Nothing Then GoTo exithandler                           ' Kein form -> fertig
    Set frm = F
    frm.Icon = objObjectBag.GetCMDIco
    frm.BackColor = &H0&
    For Each Ctl In frm.Controls                                    ' Alle Controls duchgehen
        If CtlIsTextBox(Ctl) Then                                   ' Ctl ist Textbox
            Ctl.BackColor = &H0&
            Ctl.ForeColor = &HC0C0C0
            Ctl.Font = "Lucida Console"
        End If
        If CtlIsLabel(Ctl) Then                                     ' Ctl ist Label
            Ctl.BackColor = &H0&
            Ctl.ForeColor = &HC0C0C0
            Ctl.Font = "Lucida Console"
        End If
        If CtlIsButton(Ctl) Then                                    ' Ctl ist Commandbar
            Ctl.Font = "Lucida Console"
        End If
        If CtlIsCheck(Ctl) Then                                     ' Ctl ist Checkbox
            Ctl.BackColor = &H0&
            Ctl.ForeColor = &HC0C0C0
            Ctl.Font = "Lucida Console"
        End If
        If CtlIsFrame(Ctl) Then                                     ' Ctl ist Frame
            Ctl.BackColor = &H0&
            Ctl.ForeColor = &HC0C0C0
            Ctl.Font = "Lucida Console"
        End If
        If CtlIsStatusbar(Ctl) Then                                 ' Ctl ist Statusbar
            ' ...
        End If
        If CtlIsOptionButton(Ctl) Then                              ' Ctl ist Optionsbutton
            Ctl.BackColor = &H0&
            Ctl.ForeColor = &HC0C0C0
            Ctl.Font = "Lucida Console"
        End If
        If CtlIsCombo(Ctl) Then                                     ' Ctl ist Kombobox
            Ctl.BackColor = &H0&
            Ctl.ForeColor = &HC0C0C0
            Ctl.Font = "Lucida Console"
        End If
    Next                                                            ' Nächstes Control
exithandler:
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "FormToConsole", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function FormToMorse(F As Object) As Boolean
    Dim Ctl As Control                                              ' Aktuelles Control
    Dim frm As Form                                                 ' Form referenz
    Dim il As Image
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If F Is Nothing Then GoTo exithandler                           ' Kein form -> fertig
    Set frm = F
    'frm.Icon = objObjectBag.GetCMDIco
    frm.Caption = TextTo(frm.Caption, 4)
    'frm.BackColor = &HFFFFFF
    For Each Ctl In frm.Controls                                    ' Alle Controls duchgehen
        If CtlIsTextBox(Ctl) Then                                   ' Ctl ist Textbox
            'Ctl.BackColor = &H0&
            'Ctl.ForeColor = &HC0C0C0
            'Ctl.Font = "Lucida Console"
            'Ctl.BorderStyle = 0
            'Ctl.Appearance = 0
            Ctl.Text = TextTo(Ctl.Text, 4)
        End If
        If CtlIsLabel(Ctl) Then                                     ' Ctl ist Label
 '           Ctl.BackColor = &H0&
 '           Ctl.ForeColor = &HC0C0C0
            'Ctl.Font = "Lucida Console"
            'Ctl.BackStyle = 0
            Ctl.Caption = TextTo(Ctl.Caption, 4)
        End If
        If CtlIsButton(Ctl) Then                                    ' Ctl ist Commandbar
'            Ctl.BackColor = &HC0C0C0
            'Ctl.Font = "Lucida Console"
            Ctl.Caption = TextTo(Ctl.Caption, 4)
        End If
        If CtlIsCheck(Ctl) Then                                     ' Ctl ist Checkbox
            'Ctl.BackColor = &H0&
            'Ctl.ForeColor = &HC0C0C0
            'Ctl.Font = "Lucida Console"
            Ctl.Caption = TextTo(Ctl.Caption, 4)
        End If
        If CtlIsFrame(Ctl) Then                                     ' Ctl ist Frame
            'Ctl.BackColor = &H0&
            'Ctl.ForeColor = &HC0C0C0
            'Ctl.Font = "Lucida Console"
            Ctl.Caption = TextTo(Ctl.Caption, 4)
        End If
        If CtlIsStatusbar(Ctl) Then                                 ' Ctl ist Statusbar
            Stop
            
        End If
        If CtlIsCombo(Ctl) Then                                     ' Ctl ist Kombobox
            'Ctl.BackColor = &H0&
            'Ctl.ForeColor = &HC0C0C0
            'Ctl.Font = "Lucida Console"
            Ctl.Text = TextTo(Ctl.Text, 4)
        End If
    Next                                                            ' Nächstes Control
exithandler:
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "FormToMorse", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function FormTo0815(F As Object) As Boolean
    Dim Ctl As Control                                              ' Aktuelles Control
    Dim frm As Form                                                 ' Form referenz
    Dim il As Image
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If F Is Nothing Then GoTo exithandler                           ' Kein form -> fertig
    Set frm = F
    frm.Icon = objObjectBag.GetCMDIco
    frm.BackColor = &HFFFFFF
    For Each Ctl In frm.Controls                                    ' Alle Controls duchgehen
        If CtlIsTextBox(Ctl) Then                                   ' Ctl ist Textbox
            'Ctl.BackColor = &H0&
            'Ctl.ForeColor = &HC0C0C0
            'Ctl.Font = "Lucida Console"
            Ctl.BorderStyle = 0
            Ctl.Appearance = 0
        End If
        If CtlIsLabel(Ctl) Then                                     ' Ctl ist Label
 '           Ctl.BackColor = &H0&
 '           Ctl.ForeColor = &HC0C0C0
            'Ctl.Font = "Lucida Console"
            Ctl.BackStyle = 0
        End If
        If CtlIsButton(Ctl) Then                                    ' Ctl ist Commandbar
'            Ctl.BackColor = &HC0C0C0
            'Ctl.Font = "Lucida Console"
        End If
        If CtlIsCheck(Ctl) Then                                     ' Ctl ist Checkbox
            Ctl.BackColor = &H0&
            Ctl.ForeColor = &HC0C0C0
            'Ctl.Font = "Lucida Console"
        End If
        If CtlIsFrame(Ctl) Then                                     ' Ctl ist Frame
            Ctl.BackColor = &H0&
            Ctl.ForeColor = &HC0C0C0
            'Ctl.Font = "Lucida Console"
        End If
        If CtlIsStatusbar(Ctl) Then                                 ' Ctl ist Statusbar
            Stop
            
        End If
        If CtlIsCombo(Ctl) Then                                     ' Ctl ist Kombobox
            Ctl.BackColor = &H0&
            Ctl.ForeColor = &HC0C0C0
            'Ctl.Font = "Lucida Console"
        End If
    Next                                                            ' Nächstes Control
exithandler:
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "FormTo0815", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function ExecNetSend(szCompName As String) As Boolean
Dim szText As String
Dim szcmd As String
Dim szArgs As String
Dim Result
    If Not PingWMI(szCompName) Then Exit Function
    szText = InputBox("Nachricht an " & szCompName)
    szcmd = "NET SEND"
    szArgs = szCompName & " " & szText
    Result = ShellExec(szcmd, szArgs, vbNormalFocus)
End Function

Public Sub wait(lngMilisec As Long)
    Dim TimerA
    Dim TimerB
    TimerA = GetTickCount
    TimerB = GetTickCount
    While TimerB < TimerA + lngMilisec
        DoEvents
        TimerB = GetTickCount
    Wend
End Sub

Public Function Crypt(Inp As String, Mode As Boolean, Optional ThisCryptKey As String) As String
' Ver/entschlüsselt Strings
    Dim z As String
    Dim i As Integer, Position As Integer
    Dim cptZahl As Long, orgZahl As Long
    Dim keyZahl As Long, cptString As String
    Dim szKey As String                                             ' Schlüssel String
On Error GoTo Errorhandler                                          ' Fehlerbahandlung aktivieren
    If ThisCryptKey <> "" Then                                      ' Feststellen welcher schlüssel verwendet wird
        szKey = ThisCryptKey                                        ' Angegebener schlüssel
    Else                                                            ' Sonst
        szKey = CryptKey                                            ' Default schlüssel
    End If
    For i = 1 To Len(Inp)
        Position = Position + 1
        If Position > Len(szKey) Then Position = 1
        keyZahl = Asc(Mid(szKey, Position, 1))
        If Mode Then                                                ' Verschlüsseln
            orgZahl = Asc(Mid(Inp, i, 1))
            cptZahl = orgZahl Xor keyZahl
            cptString = Hex(cptZahl)
            If Len(cptString) < 2 Then cptString = "0" & cptString
            z = z & cptString
        Else                                                        ' Entschlüsseln
            If i > Len(Inp) \ 2 Then Exit For
            cptZahl = CByte("&H" & Mid$(Inp, i * 2 - 1, 2))
            orgZahl = cptZahl Xor keyZahl
            z = z & Chr$(orgZahl)
        End If
    Next i
    Crypt = z                                                       ' ergebnis zurück
exithandler:
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "Crypt", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function GetVersion(Optional bVersNrOnly As Boolean) As String
    Dim szVersion As String
    If bVersNrOnly Then szVersion = "Version "
    GetVersion = szVersion & App.Major & "." & App.Minor
End Function

Public Function WinTheme(F As Object, Optional bActivate As Boolean)
' Aktiviert/Deaktiviert für f (Frm oder Ctl) die Windows Themes
On Error Resume Next                                                ' Hier keine Fehlerbehandlung
    If F Is Nothing Then Exit Function                              ' Kein f dann -> Fertig
    Call InitCommonControls
    If bActivate Then                                               ' Soll aktiviert werden
        ActivateWindowTheme F.hWnd                                  ' Dannn naktivieren
    Else                                                            ' Sonst
        DeactivateWindowTheme F.hwd                                 ' Deaktivieren
    End If
    Err.Clear                                                       ' Evtl Error clearen
End Function

Public Function IsIDE(Optional F As Object) As Boolean
    Dim Buffer As String
    Dim ParentWnd As Long
On Error GoTo Errorhandler                                          ' Fehlerbahandlung aktivieren
    If Not F Is Nothing Then                                        ' Saubere Methode
        Buffer = Space(128)
        ParentWnd = GetWindow(F.hWnd, GW_OWNER)
        Call GetClassName(ParentWnd, Buffer, Len(Buffer))
        MsgBox Buffer
        If Left(Buffer, 11) = "ThunderMain" Then
            
        Else
            
        End If
    Else                                                            ' Sonst Dirty Methode
On Error GoTo DirtyHandler                                          ' Hier alternative Fehlerbehandlung
        Debug.Print 1 / 0
        GoTo exithandler                                            ' Weiter mit Exithandler
    End If
GoTo exithandler                                                    ' Weiter mit Exithandler
DirtyHandler:
    IsIDE = True
exithandler:
Exit Function                                                       ' Funktion beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "IsIDE", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function
