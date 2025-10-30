Imports System.IO                                                               ' IO Klasse importieren (Spart schreibarbeit)

Public Class clsError

    Private Const MODULNAME = "clsError"                                        ' Modulname für Fehlerbehandlung

    Private ProtFilePath As String                                              ' Protokoll Pfad
    Private ErrProtFilePath As String                                           ' Fehler Protokoll Pfad

#Region "Constructor"

    Public Sub New()

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#End Region

#Region "Properties"

    Public Property ErrProtFile() As String
        Get
            Return ErrProtFilePath
        End Get
        Set(ByVal value As String)
            ErrProtFilePath = value
        End Set
    End Property

    Public Property ProtFile() As String
        Get
            Return ProtFilePath
        End Get
        Set(ByVal value As String)
            ProtFilePath = value
        End Set
    End Property

#End Region

    Public Function ErrorHandler(ByVal szModName As String, _
                                 ByVal szFktName As String, _
                                 ByVal ex As Exception, _
                                 Optional ByVal szDetails As String = "") As Boolean
        Dim szErrText As String                                                 ' Meldungstext
        On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
        szErrText = "In der Funktion: " & szFktName & vbCrLf _
                & "im Modul: " & szModName & vbCrLf _
                & "ist folgender Fehler aufgetreten." & vbCrLf _
                & "Fehlertext: " & ex.Message & vbCrLf                          ' Meldungstext zusammensetzen
        If szDetails <> "" Then                                                 ' Wenn Details vorhanden
            szErrText = szErrText & vbCrLf & "Details: " & szDetails            ' Detail info anhängen
        End If
        Call ShowErrMsg(szErrText, MsgBoxStyle.Critical, "Fehler")              ' Melsung anzeigen (bis auf weiteres
        ' Protokoll schreiben
        ' Fehler Tabelle füllen ?
        Err.Clear()                                                             ' evtl auftretenden Error Clearen
        Return True
    End Function

    Public Function ShowErrMsg(ByVal szText As String, _
                               ByVal Buttons As MsgBoxStyle, _
                               ByVal szTitel As String) As MsgBoxResult
        On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
        Return MsgBox(szText, Buttons, szTitel)                                 ' erstmal nur MSG Box ausgeben
        Err.Clear()                                                             ' evtl auftretenden Error Clearen
    End Function

    Public Function WriteProtokoll(ByVal szText As String, _
                                   Optional ByVal FilePath As String = "", _
                                   Optional ByVal bWithOutDate As Boolean = False, _
                                   Optional ByVal bWithOutUser As Boolean = False, _
                                   Optional ByVal bWithOutComputer As Boolean = False, _
                                   Optional ByVal bAppName As Boolean = True) As Boolean
        Dim szMSG As String                                                     ' Zusammengesetzter meldungtest
        Try                                                                     ' Fehlerbehandlung aktivieren
            If szText = "" Then Return False ' kein Protokoll text -> Fertig
            If FilePath = "" Then                                               ' Keine Protokoll datei 
                FilePath = ErrProtFile                                          ' -> Fehler protokoll
            End If
            szMSG = CStr(Now()) & vbTab                                         ' Datum & uhrzeit im Meldung
            'If oClsApp.Com <> "" And Not bWithOutComputer Then szMSG = szMSG & Computer & vbTab ' PC name im Meldung
            'If User <> "" And Not bWithOutUser Then szMsg = szMsg & User & vbTab ' Username im Meldung
            'szMsg = szMsg & objObjectBag.GetAppTitle & vbTab                    ' Anwendungstitel im Meldung
            szMsg = szMsg & szText & vbCrLf                                     ' Eigentlicher Meldungstext
            Return WriteTextFile(szMSG, FilePath)
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ErrorHandler(MODULNAME, "WriteProtokoll", ex)                  ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function WriteErrorProtokoll(ByVal szText As String) As Boolean
        Dim szMSG As String                                                     ' Zusammengesetzter meldungtest
        Try                                                                     ' Fehlerbehandlung aktivieren
            If szText = "" Then  ' kein Protokoll text 
                szText = "Kein Fehlertext angegeben."
            End If
            If ErrProtFilePath = "" Then Return False ' Keine Protokoll datei 
            szMSG = CStr(Now()) & vbTab                                         ' Datum & uhrzeit im Meldung
            Return WriteTextFile(szMSG, ErrProtFilePath)                        ' Fehler meldung schreiben
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ErrorHandler(MODULNAME, "WriteProtokoll", ex)                  ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function WriteTextFile(ByVal szText As String, _
                                   Optional ByVal FilePath As String = "") As Boolean
        Dim EFSR As IO.StreamWriter
        Try                                                                     ' Fehlerbehandlung aktivieren
            If szText = "" Then Return False ' kein Protokoll text -> Fertig
            If File.Exists(FilePath) Then                                       ' Wenn Datei existiert
                EFSR = New IO.StreamWriter(FilePath)                            ' Zum Schreiben öffnen
                EFSR.WriteLine(szText)                                          ' Schreiben
                EFSR.Close()                                                    ' Datei schliessen
                Return True
            Else
                Return False
            End If
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ErrorHandler(MODULNAME, "WriteTextFile", ex)                   ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

End Class
