Public Class clsUser

    Private Const MODULNAME = "clsUser"                                         ' Modulname für Fehlerbehandlung
    Private ObjBag As clsObjectBag                                              ' Sammelklasse
    Private bInitOK As Boolean                                                  ' Gibt an das die Klasse erfolgreich initialisiert wurde

    Private szNTUsername As String                                              ' NT Anmelde name
    Private szVorname As String                                                 ' Vorname des Benutzers
    Private szNachname As String                                                ' Nachname des Benutzers
    Private szPersonalDir As String                                             ' Eigene Dokumente des Users
    Private szEmail As String                                                   ' Email Adresse

    Public Sub New(ByVal oBag As clsObjectBag)
        Try                                                                     ' Fehlerbehandlung aktivieren
            ObjBag = oBag                                                       ' ObjectBag Übergeben
            If Not ReadEnviroment() Then                                        ' Evironment auslesen
                bInitOK = False                                                 ' Misserfolg zurück
                Exit Sub
            End If

            bInitOK = True                                                      ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            bInitOK = False                                                     ' Misserfolg zurück
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#Region "Properties"

    Public ReadOnly Property InitOK As Boolean                                  ' Gibt zurück ob die initialieierung Fehlerfrei war
        Get
            Return bInitOK
        End Get
    End Property

    Public ReadOnly Property UserName As String                                 ' Benutzername
        Get
            Return szNTUsername
        End Get
    End Property

    Public ReadOnly Property PersonalDir As String                              ' Eigene Dokumente
        Get
            If Right(szPersonalDir, 1) <> "\" Then szPersonalDir = szPersonalDir & "\"
            Return szPersonalDir
        End Get
    End Property

#End Region

    Private Function ReadEnviroment() As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            szNTUsername = Environment.UserName                                 ' Akt User Auslesen
            szPersonalDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) ' Eigene Dateien

            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Return False                                                        ' Misserfolg zurück
        End Try

    End Function

End Class
