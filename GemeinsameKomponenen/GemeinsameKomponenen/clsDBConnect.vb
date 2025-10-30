Imports System.Data.Common                                                      ' Data.Common Klasse Importieren (Spart schreibarbeit)
Imports System.Data.SqlClient                                                   ' Data.SQLClient Klasse Importieren (Spart schreibarbeit)

Public Class clsDBConnect
    Private Const MODULNAME = "clsDBConnect"                                    ' Modulname für Fehlerbehandlung
    Private ObjBag As clsObjectBag                                              ' Sammelklasse
    Private bInitOK As Boolean                                                  ' Gibt an das die Klasse erfolgreich initialisiert wurde

    Private objSQLCon As SqlConnection                                          ' Aktuelle SQL DB Verbindung
    Private objAccCon As OleDb.OleDbConnection                                  ' Aktuelle Acces verb (noch nicht unterstützt

    Private bAccessPossible As Boolean                                          ' Verbindung zu Access DB zulassig
    Private bMSSQLSrvPossible As Boolean                                        ' Verbindung zu SQL Server zulässig
    Private ConParams As ConnetionParameter                                     ' Aktuelle Verbindungsparameter

    Private Structure ConnetionParameter
        Public bSQL As Boolean                                                  ' True wenn SQL Server ; False wenn Access
        Public Provider As String                                               ' DB Provider
        Public SQLServer As String                                              ' SQL Server name
        Public DBText As String                                                 ' Beschreibung der Datenbank für Anzeige
        Public DBName As String                                                 ' DB Name (bei Access mit Pfad)
        Public DBUser As String                                                 ' Benutzername der datenbank
        Public DBPWD As String                                                  ' DB Kennwort
        Public CryptPWD As String                                               ' DB Kennwort Verschlüsselt
        Public bNt As Boolean                                                   ' NT Autentifizierung (nur SQL)
        Public ConnectString As String                                          ' Completer Connect String
        Public bUserlogIn As Boolean                                            ' Gibt es ein seperates Benutzer login
        Public bSingleSignOn As Boolean                                         ' Anmeldung von win durchreichen
        Public LoginText As String                                              ' Text Für Benutzerlogin
        Public UserTable As String                                              ' Tabelle mit Benutzerdaten
        Public AddWhere As String                                               ' Opt. Where für Benutzertabelle
        Public UsernameField As String                                          ' Feld mit Benutzernamen
        Public UsernameField2 As String                                         ' Feld mit alternativen Benutzernamen (z.B. angezeigter name und NtAnmeldename)
        Public PWDField As String                                               ' Feld Mit Kennwort
        Public UserLoginMaxCount As Integer                                     ' Max Anmelde versuche
    End Structure

#Region "Constructor"

    Public Sub New(ByVal oBag As clsObjectBag, _
                    ByVal bAccess As Boolean, _
                    ByVal bSQL As Boolean)
        Try                                                                     ' Fehlerbehandlung aktivieren
            ObjBag = oBag                                                       ' ObjectBag Übergeben
            bAccessPossible = bAccess                                           ' Übernehmen ob Access Zulässig
            bMSSQLSrvPossible = bSQL                                            ' Übernehmen ob SQL zulässig
            ConParams = New ConnetionParameter                                  ' Verbindungsparameter initialisieren
            bInitOK = True                                                      ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            bInitOK = False                                                     ' Misserfolg zurück
        End Try
    End Sub

    Public Sub New(ByVal oBag As clsObjectBag)
        Try                                                                     ' Fehlerbehandlung aktivieren
            ObjBag = oBag                                                       ' ObjectBag Übergeben
            bAccessPossible = True                                              ' Access Zulässig
            bMSSQLSrvPossible = True                                            ' SQL Zulässig
            ConParams = New ConnetionParameter                                  ' Verbindungsparameter initialisieren
            bInitOK = True                                                      ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            bInitOK = False                                                     ' Misserfolg zurück
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        If ConParams.bSQL Then                                                  ' SQL Server verbindung
            If Not IsNothing(objSQLCon) Then                                    ' Wenn Verbindung noch existiert
                If objSQLCon.State <> ConnectionState.Closed Then               ' Wenn Verbindung nicht closed
                    Call objSQLCon.Close()                                      ' Schliessen
                End If
            End If
        Else                                                                    ' Sonst Access
            If Not IsNothing(objAccCon) Then                                    ' Wenn Verbindung noch existiert
                If objAccCon.State <> ConnectionState.Closed Then               ' Wenn Verbindung nicht closed
                    Call objAccCon.Close()                                      ' Schliessen
                End If
            End If
        End If
        
        MyBase.Finalize()
    End Sub

#End Region

#Region "Properties"

    Public ReadOnly Property SQLServer As String                                ' Gibt den SQL Server namen (+ instanz) zurück
        Get
            Return ConParams.SQLServer
        End Get
    End Property

    Public ReadOnly Property DBName As String                                   ' Gibt den Datenbanknamen zurück
        Get
            Return ConParams.DBName
        End Get
    End Property

    Public ReadOnly Property DBUser As String                                   ' Gibt den Datenbankbenutzer zurück
        Get
            Return ConParams.DBUser
        End Get
    End Property

    Public ReadOnly Property DBCon As DbConnection                              ' Gibt zurück ob die initialieierung Fehlerfrei war
        Get
            If ConParams.bSQL Then                                              ' Ist SQL Server verbindung
                Return objSQLCon                                                ' Diese zurück
            Else                                                                ' Sonst Access
                Return objAccCon
            End If
        End Get
    End Property

    Public ReadOnly Property InitOK As Boolean                                  ' Gibt zurück ob die initialieierung Fehlerfrei war
        Get
            Return bInitOK
        End Get
    End Property

    Public ReadOnly Property SQLPossible As Boolean
        Get
            SQLPossible = bMSSQLSrvPossible
        End Get
    End Property

    Public ReadOnly Property AccessPossible As Boolean
        Get
            AccessPossible = bAccessPossible
        End Get
    End Property

    Public Property UserLogIn As Boolean
        Get
            UserLogIn = ConParams.bUserlogIn
        End Get
        Set(ByVal value As Boolean)
            ConParams.bUserlogIn = value
        End Set
    End Property

    Public Property SingleSignOn As Boolean
        Get
            SingleSignOn = ConParams.bSingleSignOn
        End Get
        Set(ByVal value As Boolean)
            ConParams.bSingleSignOn = value
        End Set
    End Property

#End Region

    Public Function BulidDBConnection(ByVal lngType As Integer, _
         ByVal szServername As String, _
         ByVal szDBName As String, _
         Optional ByVal szDBUser As String = "", _
         Optional ByVal szPWD As String = "", _
         Optional ByVal bNtAut As Boolean = False) As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            With ConParams
                .bSQL = True                                                    ' Kein Access
                .SQLServer = szServername                                       ' Servernamen Speichern
                .DBName = szDBName                                              ' Datenbank namen Speichern
                .DBUser = szDBUser                                              ' Datenbank User Speichern
                .DBPWD = szPWD                                                  ' PWD Speichern
                .bNt = bNtAut                                                   ' SQL Autentifizierung speichern
            End With
            Select Case lngType                                                 ' Verbindungstyp auswerten
                Case 1                                                          ' 1 = SQL Server
                    If Not bMSSQLSrvPossible Then Return False
                    If szServername = "" Or szDBName = "" Then
                        If Not AskForConParams() Then Return False
                    End If
                    Return BuildSQLConnection(szServername, szDBName, szDBUser, szPWD, bNtAut)

                Case 2                                                          ' 1 = Access
                    If Not bAccessPossible Then Return False
                    If szDBName = "" Then
                        If Not AskForConParams() Then Return False
                    End If
                    Return BuildAccessConnection(szDBName, szDBUser, szPWD)

            End Select
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "BulidDBConnection", ex)        ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Public Function BuildAccessConnection(ByVal szDBName As String, _
        Optional ByVal szDBUser As String = "", _
        Optional ByVal szPWD As String = "", _
        Optional ByVal bWithTest As Boolean = False) As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            With ConParams
                .bSQL = False                                                   ' Kein SQL

                .DBText = "Access: " & .DBName
            End With
            If bWithTest Then                                                   ' Wenn gewünscht verbindung testen
                Return TestConnetion()
            End If
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "BuildAccessConnection", ex)    ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Public Function BuildSQLConnection(ByVal szServername As String, _
        ByVal szDBName As String, _
        Optional ByVal szDBUser As String = "", _
        Optional ByVal szPWD As String = "", _
        Optional ByVal bNtAut As Boolean = False, _
        Optional ByVal bWithTest As Boolean = False) As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            Dim builder As New SqlConnectionStringBuilder
            With ConParams
                .bSQL = True                                                    ' Kein Access
                .SQLServer = szServername                                       ' Servernamen Speichern
                .DBName = szDBName                                              ' Datenbank namen Speichern
                .DBUser = szDBUser                                              ' Datenbank User Speichern
                .DBPWD = szPWD                                                  ' PWD Speichern
                .bNt = bNtAut                                                   ' SQL Autentifizierung speichern

                builder("Initial Catalog") = .DBName                            ' Datenbanknamen an ConStringBuilder übergeben
                builder("Data Source") = .SQLServer                             ' Servernamen an ConStringBuilder übergeben
                builder("User ID") = .DBUser                                    ' Usernamen an ConStringBuilder übergeben
                builder("Integrated Security") = .bNt
                If Not .bNt Then                                                ' Wenn keine NT security
                    builder.Password = .DBPWD                                   ' Dann brauchen wir das PWD
                End If
                .DBText = "MS SQL: " & .DBName & " auf " & .SQLServer
                .ConnectString = builder.ConnectionString                       ' Connetionstring zusammensetzen lassen
            End With
            objSQLCon = New SqlConnection(ConParams.ConnectString)              ' Neue Verbindung mit connectionString
            If bWithTest Then                                                   ' Wenn gewünscht verbindung testen
                Return TestConnetion()
            End If

            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "BuildSQLConnection", ex)       ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Public Function ExecSQL(ByVal szSQL As String, _
                            Optional ByVal bSilent As Boolean = False) As Boolean
        ' Führt das SQL statement aus.
        ' Liefert True wenn kein Fehler aufgetreten ist
        Dim szDetails As String                                                 ' Details für Fehlerbehandlung
        Dim ConDB As DbConnection
        szDetails = "SQL: " & szSQL                                             ' SQL Statement als Details für Fehlerbehandlung
        Try                                                                     ' Fehlerbehandlung aktivieren
            If szSQL = "" Then                                                  ' Wenn Kein SQL 
                Return False                                                    ' dann Raus
            End If
            ConDB = DBCon
            Call ConDB.Open()
            'Call objDBCon.Open()                                                ' DB Verbindung öffnen
            'Dim cmd As New OleDb.OleDbCommand(szSQL, objDBCon)                  ' SQL Commando mit DBCon erstellen
            Dim cmd As New SqlCommand(szSQL, ConDB)                  ' SQL Commando mit DBCon erstellen
            Call cmd.ExecuteNonQuery()                                          ' Ausführen
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not bSilent Then                                                 ' Wenn nicht silent gewünscht
                Call ObjBag.ErrorHandler(MODULNAME, "ExecSQL", ex, szDetails)   ' Fehlermeldung ausgeben
            End If
            Return False                                                        ' Misserfolg zurück
        Finally
            Call CloseConection()                                               ' Verbindung schliessen       
        End Try
    End Function

    Public Function FillDS(ByVal Da As DbDataAdapter, _
                                 Optional ByVal szDSName As String = "",
                                 Optional ByVal bSilent As Boolean = False) As DataSet
        ' Liefert auf basis des SQL Statements ein Dataset zurück
        Return FillDataSet(Da, szDSName, bSilent)
    End Function

    Public Function FillDS(ByVal szSQL As String, _
                             Optional ByVal szDSName As String = "",
                             Optional ByVal bSilent As Boolean = False) As DataSet
        ' Liefert auf basis des SQL Statements ein Dataset zurück
        Return FillDataSet(szSQL, szDSName, bSilent)
    End Function

    Public Function GetDA(ByVal szSQL As String, _
                          Optional ByVal bSilent As Boolean = False) As DbDataAdapter
        'Public Function GetDA(ByVal szSQL As String, _
        '                          Optional ByVal bSilent As Boolean = False) As OleDb.OleDbDataAdapter
        Return GetDataAdapter(szSQL, bSilent)
    End Function

    Public Function UpdateDS(ByVal Da As DbDataAdapter, _
                              ByVal DS As DataSet, _
                              Optional ByVal szTablename As String = "", _
                              Optional ByVal bSilent As Boolean = False) As DataSet
        Return UpdateDataSet(Da, DS, szTablename, bSilent)
    End Function

    Public Function GetValueFromSQL(ByVal szSQL As String, _
                                    Optional ByVal bSilent As Boolean = False) As Object
        ' Liefert den 1. Wert der 1. DS des SQL statements zurück
        Dim DS As DataSet                                                       ' Ergebnis Dataset
        Dim szDetails As String = ""                                            ' Details für Fehlerbehandlung
        Try                                                                     ' Fehlerbehandlung aktivieren
            szDetails = "SQL: " & szSQL                                         ' SQL Statement als Details für Fehlerbehandlung
            If szSQL = "" Then Return Nothing ' Kein SQL angegeben -> Fertig
            DS = FillDataSet(szSQL)                                             ' DS Füllen
            If DS.Tables(0).Rows.Count > 0 Then                                 ' Wenn Mind 1. DS
                Return DS.Tables(0).Rows(0).Item(0)                             ' Wert holen
            Else
                Return Nothing                                                  ' Keine Daten gefunden -> Nichts zurück
            End If
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not bSilent Then                                                 ' Wenn nicht silent gewünscht
                Call ObjBag.ErrorHandler(MODULNAME, "GetValueFromSQL", ex, szDetails) ' Fehlermeldung ausgeben
            End If
            Return Nothing                                                      ' Misserfolg zurück
        Finally
            Call CloseConection()                                               ' Verbindung schliessen
        End Try
    End Function

    Private Function FillDataSet(ByVal szSQL As String, _
                                 Optional ByVal szTablename As String = "",
                                 Optional ByVal bSilent As Boolean = False) As DataSet
        ' Liefert auf basis des SQL Statements ein Dataset zurück
        Dim szDetails As String = ""                                             ' Details für Fehlerbehandlung
        Dim ds As New DataSet                                                   ' Unser rDataset
        Try                                                                     ' Fehlerbehandlung aktivieren
            szDetails = "SQL: " & szSQL                                         ' SQL Statement als Details für Fehlerbehandlung
            If szSQL = "" Then Return Nothing ' Kein SQL Statement -> Fertig
            'Dim da As OleDb.OleDbDataAdapter = GetDataAdapter(szSQL, bSilent)   ' DataAdapter mit SQL statement & DBConnection erstellen
            Dim da As DbDataAdapter = GetDataAdapter(szSQL, bSilent)   ' DataAdapter mit SQL statement & DBConnection erstellen
            If OpenConnection() Then
                If szTablename <> "" Then                                           ' Wenn Tabellenname angegeen
                    Call da.Fill(ds, szTablename)                                   ' DataSet über DataAdapter füllen mit Tabelennamen
                Else                                                                ' Sonst
                    Call da.Fill(ds)                                                ' Ohne Tabelennamen
                End If
                Call CloseConection()
            End If
            Return ds                                                           ' DataSet zurückgeben
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not bSilent Then                                                 ' Wenn nicht silent gewünscht
                Call ObjBag.ErrorHandler(MODULNAME, "FillDataSet", ex, szDetails) ' Fehlermeldung ausgeben
            End If
            Return Nothing                                                      ' Misserfolg zurück
        Finally
            Call CloseConection()                                                  ' Verbindung schliessen
        End Try
    End Function

    Private Function FillDataSet(ByVal da As DbDataAdapter, _
                                 Optional ByVal szTablename As String = "",
                                 Optional ByVal bSilent As Boolean = False) As DataSet
        ' Liefert auf basis des SQL Statements ein Dataset zurück
        Dim szDetails As String = ""                                             ' Details für Fehlerbehandlung
        Dim ds As New DataSet                                                   ' Unser rDataset
        Try                                                                     ' Fehlerbehandlung aktivieren
            If IsNothing(da) Then Return Nothing ' Kein Dataadapter -> Fertig
            If szTablename <> "" Then                                           ' Wenn Tabellenname angegeen
                Call da.Fill(ds, szTablename)                                   ' DataSet über DataAdapter füllen mit Tabelennamen
            Else                                                                ' Sonst
                Call da.Fill(ds)                                                ' Ohne Tabelennamen
            End If
            Return ds                                                           ' DataSet zurückgeben
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not bSilent Then                                                 ' Wenn nicht silent gewünscht
                Call ObjBag.ErrorHandler(MODULNAME, "FillDataSet", ex, szDetails) ' Fehlermeldung ausgeben
            End If
            Return Nothing                                                      ' Misserfolg zurück
        Finally
            Call CloseConection()                                               ' Verbindung schliessen
        End Try
    End Function

    Private Function GetDataAdapter(ByVal szSQL As String, _
                                     Optional ByVal bSilent As Boolean = False) As DbDataAdapter
        Dim szDetails As String = ""                                            ' Details für Fehlerbehandlung
        Dim DA As DbDataAdapter
        Try                                                                     ' Fehlerbehandlung aktivieren
            szDetails = "SQL: " & szSQL                                         ' SQL Statement als Details für Fehlerbehandlung
            If szSQL = "" Then Return Nothing
            If ConParams.bSQL Then
                DA = New SqlDataAdapter(szSQL, objSQLCon)

                Dim CB As New SqlCommandBuilder(DA)
                DA.UpdateCommand = CB.GetUpdateCommand
                Return DA
            Else
                DA = New OleDb.OleDbDataAdapter(szSQL, objAccCon)                  ' DataAdapter mit SQL statement & DBConnection erstellen
                'DA.
                Return DA
            End If

        Catch ex As Exception                                                   ' Fehler behandeln
            If Not bSilent Then                                                 ' Wenn nicht silent gewünscht
                Call ObjBag.ErrorHandler(MODULNAME, "GetDataAdapter", ex, szDetails) ' Fehlermeldung ausgeben
            End If
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

    Private Function UpdateDataSet(ByVal Da As DbDataAdapter, _
                                   ByVal Ds As DataSet, _
                                   Optional ByVal szTablename As String = "", _
                                   Optional ByVal bSilent As Boolean = False) As DataSet
        ' Liefert auf basis des SQL Statements ein Dataset zurück
        Dim szDetails As String = ""                                            ' Details für Fehlerbehandlung
        Try                                                                     ' Fehlerbehandlung aktivieren
            If IsNothing(Da) Then Return Nothing ' Kein Dataadapter -> Fertig
            If IsNothing(Ds) Then Return Nothing ' Kein DataSet -> Fertig
            If szTablename <> "" Then                                           ' Wenn Tabellenname angegeen
                Call Da.Update(Ds, szTablename)                                 ' DataSet über DataAdapter füllen mit Tabelennamen
            Else                                                                ' Sonst
                Call Da.Update(Ds)                                              ' Ohne Tabelennamen
            End If
            Return Ds                                                           ' DataSet zurückgeben
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not bSilent Then                                                 ' Wenn nicht silent gewünscht
                Call ObjBag.ErrorHandler(MODULNAME, "UpdateDataSet", ex, szDetails) ' Fehlermeldung ausgeben
            End If
            Return Nothing                                                      ' Misserfolg zurück
        Finally
            Call CloseConection()                                               ' Verbindung schliessen
        End Try
    End Function

    Public Function OpenConnection() As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            If DBCon.State <> ConnectionState.Open Then                         ' Wenn Verbindung nicht offen
                Call DBCon.Open()                                               ' DB Verbindung öffnen
            End If
            'If objDBCon.State <> ConnectionState.Open Then                      ' Wenn Verbindung nicht closes
            '    Call objDBCon.Open()                                            ' DB Verbindung öffnen
            'End If
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "OpenConnection", ex)           ' Fehlermeldung ausgeben
            Return False
        End Try
    End Function

    Public Function CloseConection() As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            If DBCon.State <> ConnectionState.Closed Then                       ' Wenn Verbindung nicht geschlossen
                Call DBCon.Close()                                              ' DB Verbindung Schliessen
            End If
            'If objDBCon.State <> ConnectionState.Closed Then                    ' Wenn Verbindung nicht closes
            '    Call objDBCon.Close()                                           ' DB Verbindung Schliessen
            'End If
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "CloseConection", ex)           ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function TestConnetion() As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            Call DBCon.Open()                                                   ' DB Verbindung öffnen
            'objDBCon.Open()                                                     ' DB Verbindung öffnen
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "TestConnetion", ex)            ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        Finally
            'If objDBCon.State <> ConnectionState.Closed Then                    ' Wenn Verbindung nicht closes
            '    Call objDBCon.Close()                                           ' Schliessen
            'End If
        End Try
    End Function

    Private Function AskForConParams() As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            Dim fCon As New frmConParams(ObjBag)                                ' Form für Coonect Parameter vorbereiten
            fCon.ShowDialog()                                                   ' Form als Dialog anzeigen
            With fCon                                                           ' Form auslesen
                ConParams.bSQL = .radSQL.Checked                                ' SQL 
                ConParams.SQLServer = .txtServer.Text                           ' Servernamen auslesen
                ConParams.DBName = .txtDBName.Text                              ' DB Namen auslesen
                ConParams.DBUser = .txtDBUSer.Text                              ' Usernamen auslesen
                ConParams.DBPWD = .txtPWD.Text                                  ' PWD auslesen
                ConParams.bNt = .chkNT.Checked                                  ' Nt Authent.. 
            End With
            Return Not fCon.bCancel                                             ' evtl. Cancel zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "AskForConParams", ex)          ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Public Function ConvertColumnType(ByVal szType As String, _
                                      Optional ByVal szSize As String = "", _
                                      Optional ByVal szFormat As String = "") As String
        ' Übersetzt Field.Type  vereinfacht ins Deutsche
        Dim szResult As String = ""
        Try                                                                     ' Fehlerbehandlung aktivieren
            Select Case szType                                                  ' Type übersetzen
                Case 2
                    szResult = "Zahlenwert"
                Case 3
                    szResult = "Zahlenwert"
                Case 5
                    szResult = "KommaZahl"
                Case 6
                    szResult = "Währung"
                Case 7                                                          ' Access date time
                    szResult = "Datum/Uhrzeit"
                Case 11                                                         ' Access Bool
                    szResult = "Ja/Nein"
                Case 17
                    szResult = "Zahlenwert"
                Case 131
                    szResult = "Kommazahl"
                Case 129
                    szResult = "Text"
                    If szSize <> "" Then
                        szResult = szResult & " [" & szSize & " Zeichen]"
                    End If
                Case 135
                    szResult = "Datum"
                Case 200
                    szResult = "Text"
                    If szSize <> "" Then
                        szResult = szResult & " [" & szSize & " Zeichen]"
                    End If
                Case 201
                    szResult = "Text"
                    If szSize <> "" Then
                        szResult = szResult & " [" & szSize & " Zeichen]"
                    End If
                Case 202
                    szResult = "Text"
                    If szSize <> "" Then
                        szResult = szResult & " [" & szSize & " Zeichen]"
                    End If
                Case 203                                                        ' Access Memo
                    szResult = "Text"
                    If szSize <> "" Then
                        szResult = szResult & " [" & szSize & " Zeichen]"
                    End If
                Case Else
                    Stop
            End Select
            Return szResult
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "ConvertColumnType", ex)        ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

End Class
