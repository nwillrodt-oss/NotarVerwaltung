Public Class clsApp

    Private Const MODULNAME = "clsApp"                                          ' Modulname für Fehlerbehandlung
    Private ObjBag As clsObjectBag                                              ' Sammelklasse
    Private bInitOK As Boolean                                                  ' Gibt an das die Klasse erfolgreich initialisiert wurde

    Private szVersion As String                                                 ' Frontend Version
    Private szAppTitle As String                                                ' Anwendungstitel (Vorsicht: Wird als Regkey benutzt)
    Private szRegRoot As String                                                 ' Anwendungspfad in Registry
    Private szAppDesc As String                                                 ' Anwendungs bechreibung
    Private szCopyright As String                                               ' Copyright
    Private szWWW As String                                                     ' Web Adresse
    Private szAppPath As String                                                 ' Anwendungsverz.
    Private szSoundPath As String                                               ' Sound Verz.
    Private szImagePath As String                                               ' Abbildungsverz.
    'Private szTemplatePath As String                                            ' Vorlagen Verz.
    Private szComputername As String                                            ' Name des Aktuellen PCs
    Private szOS As String
    Private szWinDir As String                                                  ' Windows ordner
    Private szSystemDir As String                                               ' System ordner (system32)
    'Private szTempDir As String                                                 ' Temp ordner

    'Private szWordPath As String                                                ' Pfad der installierten WinWord.exe
    'Private szWordVersion As String                                             ' Word Version
    'Private szExcelPath As String                                               ' Pfad der installierten Excel.exe
    'Private szExcelVersion As String                                            ' Excel Version

#Region "Constructor"

    Public Sub New(ByVal oBag As clsObjectBag, ByVal AppVersion As String, ByVal AppPath As String)
        Try                                                                     ' Fehlerbehandlung aktivieren
            ObjBag = oBag                                                       ' ObjectBag Übergeben
            szAppTitle = SZ_APPTITLE
            szRegRoot = SZ_REGROOT
            szCopyright = SZ_COPYRIGHT
            szAppDesc = SZ_APPDESC
            szWWW = SZ_WWW
            szAppPath = AppPath
            szVersion = AppVersion
            If Not ReadEnviroment() Then                                        ' Evironment auslesen
                bInitOK = False                                                 ' Misserfolg Zurück
                Exit Sub
            End If
            If Not InitAppPath() Then                                           ' Anwendungsverz. initialisieren
                bInitOK = False                                                 ' Misserfolg Zurück
                Exit Sub
            End If
            bInitOK = True                                                      ' Erfolg Zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            bInitOK = False                                                     ' Misserfolg zurück
            Call ObjBag.ErrorHandler(MODULNAME, "New", ex)                      ' Fehlermeldung ausgeben
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#End Region

#Region "Pproperties"

    Public ReadOnly Property InitOK As Boolean                                  ' Gibt zurück ob die initialieierung Fehlerfrei war
        Get
            Return bInitOK
        End Get
    End Property

    Public ReadOnly Property AppDir As String                                   ' Anwendungsverz. 
        Get
            Return szAppPath
        End Get
    End Property

    Public ReadOnly Property ImageDir As String                                 ' Immage Ordner im Anwendungsverz
        Get
            If Right(szImagePath, 1) <> "\" Then szImagePath = szImagePath & "\"
            Return szImagePath
        End Get
    End Property

    Public ReadOnly Property SoundDir As String                                 ' Sound Ordner im Anwendungsverz
        Get
            If Right(szSoundPath, 1) <> "\" Then szSoundPath = szSoundPath & "\"
            Return szSoundPath
        End Get
    End Property

    Public ReadOnly Property AppVersion As String                               ' Anwendungsversion
        Get
            Return szVersion
        End Get
    End Property

    Public ReadOnly Property Copyright As String                                ' Copyrightstring
        Get
            Return szCopyright
        End Get
    End Property

    Public ReadOnly Property Comutername As String                              ' Computername
        Get
            Return szComputername
        End Get
    End Property

    Public ReadOnly Property Operatingsystem As String                          ' Betriebssystem
        Get
            Return szOS
        End Get
    End Property

    Public ReadOnly Property Windir As String                                   ' Windows ordner
        Get
            If Right(szWinDir, 1) <> "\" Then szWinDir = szWinDir & "\"
            Return szWinDir
        End Get
    End Property

    Public ReadOnly Property Sys32Dir As String                                 ' System32 ordner
        Get
            If Right(szSystemDir, 1) <> "\" Then szSystemDir = szSystemDir & "\"
            Return szSystemDir
        End Get
    End Property

    Public ReadOnly Property AppTitle As String                                 ' Anwendungstitel
        Get
            Return szAppTitle
        End Get
    End Property

    Public ReadOnly Property AppDesc As String                                  ' Anwendungsbeschreibung
        Get
            Return szAppDesc
        End Get
    End Property

    Public ReadOnly Property RegRoot As String                                  ' Anwendungspfad in Reg
        Get
            Return szRegRoot
        End Get
    End Property

#End Region

    Private Function InitAppPath() As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            If szAppPath <> "" Then                                             ' Kein Anwendungsverz -> Nichts zu tun
                If Not CheckLastChar(szAppPath, "\", ObjBag) Then               ' Wenn letztes \ nicht vorhanden
                    szAppPath = szAppPath & "\"                                 ' Anhängen
                End If

                szSoundPath = szAppPath & "Sounds\"                             ' Sound verz. definieren
                szImagePath = szAppPath & "Images\"                             ' Imageverz. definieren

                Return True                                                     ' Erfolg Zurück
            Else                                                                ' Sonst
                Return False                                                    ' Misserfolg zurück
            End If

        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "InitAppPath", ex)              ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function ReadEnviroment() As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            szSystemDir = Environment.GetFolderPath(Environment.SpecialFolder.System) 'System32
            szWinDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows) ' Windows
            szComputername = Environment.MachineName                            ' Akt. Hostnamen
            szOS = Environment.OSVersion.ToString()                             ' Betriebssystem

            Return True                                                         ' Erfolg Zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "ReadEnviroment", ex)           ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

End Class

