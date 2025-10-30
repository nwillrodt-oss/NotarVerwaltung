Module modStringTools
    Private Const MODULNAME = "modStringTools"                                  ' Modulname für Fehlerbehandlung


    Public Function CheckLastChar(ByVal szStr As String, ByVal szChar As String, Optional ByRef oBag As clsObjectBag = Nothing) As Boolean
        ' Überprüft Letztes Vorkommendes Zeichen on szStr = szCahr ist
        Dim intCharLen As Integer                                               ' Prüf Zeichen länge
        Try                                                                     ' Fehlerbehandlung aktivieren
            intCharLen = Len(szChar)                                            ' Länge Prüf zeichen bestimmen
            If Right(szStr, intCharLen) = szChar Then                           ' Wenn Prüf zeichen vorkommt
                Return True                                                     ' True zurück
            Else                                                                ' Sonst
                Return False                                                    ' False zurück
            End If
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "CheckLastChar", ex)          ' Fehler behandlung aufrufen
            End If
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function


End Module
