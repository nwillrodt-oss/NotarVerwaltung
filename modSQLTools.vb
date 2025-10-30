Module modSQLTools
    Private Const MODULNAME = "modSQLTools"                                     ' Modulname für Fehlerbehandlung

    Private Const SQL_WHERE = "WHERE"
    Private Const SQL_SELECT = "SELECT"
    Private Const SQL_FROM = "FROM"
    Private Const SQL_ORDER = "ORDER BY"
    Private Const SQL_UNION = "UNION"
    Private Const SQL_IJOIN = "INNER JOIN"
    Private Const SQL_LJOIN = "LEFT JOIN"
    Private Const SQL_RJOIN = "RIGHT JOIN"
    Private Const SQL_AS = "AS"
    Private Const SQL_ON = "ON"
    Private Const SQL_INSERT = "INSERT INTO"
    Private Const SQL_DELETE = "DELETE"
    Private Const SQL_VALUES = "VALUES"
    Private Const SQL_UPDATE = "UPDATE"
    Private Const SQL_SET = "SET"
    Private Const SQL_GROUP = "GROUP BY"
    Private Const SQL_HAVING = "HAVING"
    Private Const SQL_ALTER = "ALTER"
    Private Const SQL_TAB = "TABLE"

    Public Function AddWhere(ByVal szWhere As String, _
                             ByVal szNewCondition As String, _
                             Optional ByVal bOr As Boolean = False, _
                             Optional ByRef oBag As clsObjectBag = Nothing) As String
        Dim szAnd As String                                                     ' Verknüpfugs String
        Try                                                                     ' Fehlerbehandlung aktivieren
            szWhere = Trim(szWhere)                                             ' Alte Condition Trimmen
            szNewCondition = Trim(szNewCondition)                               ' Neue Condition Trimmen

            If Left(szWhere.ToUpper, 6) = SQL_WHERE & " " Then
                szWhere = Right(szWhere, szWhere.Length - 6)                     ' Falls Where davor -> abschneiden
            End If
            If Left(szNewCondition.ToUpper, 6) = SQL_WHERE & " " Then
                szNewCondition = Right(szNewCondition, Len(szNewCondition) - 6) ' Falls Where davor -> abschneiden
            End If
            If szWhere <> "" Then szWhere = "(" & szWhere & ")" ' klammern
            szNewCondition = "(" & szNewCondition & ")"
            szAnd = " AND "                                                     ' Verknüpfung bstimmen AND / OR
            If bOr Then szAnd = " OR "
            If szWhere <> "" Then
                szWhere = "(" & szWhere & szAnd & szNewCondition & ")"          ' Zusammensetzen
            Else
                szWhere = szNewCondition
            End If
            If Left(szWhere, 6) <> SQL_WHERE & " " Then szWhere = SQL_WHERE & " " & szWhere ' Schlüsselwort voran
            AddWhere = szWhere
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "AddWhere", ex)               ' Fehler behandlung aufrufen
            End If
            Return ""                                                           ' Misserfolg zurück
        End Try
    End Function

    Public Function AddWhereInFullSQL(ByVal szSQL As String, _
                                      ByVal szNewCondition As String, _
                                      Optional ByVal bOr As Boolean = False, _
                                      Optional ByRef oBag As clsObjectBag = Nothing) As String
        Dim lngWhereStartPos As Integer                                         ' Startpos. den Wherestatements im SQL Statement
        Dim lngWhereEndPos As Integer                                           ' Endpos. den Wherestatements im SQL Statement
        Dim lngOrderStartPos As Integer                                         ' Startpos. den Orderstatements im SQL Statement
        Dim lngGroupStartPos As Integer                                         ' Startpos. den Groupstatements im SQL Statement
        Dim lngHavingStartPos As Integer                                        ' Startpos. den Havingstatements im SQL Statement
        Dim lngUnionStartPos As Integer
        Dim lngSQLLen As Integer                                                ' länge des SQL Statements
        Dim szWherePart As String                                               ' Wherestatement
        Dim szSelectPart As String                                              ' Alles vor dem Wherestatement
        Dim szAfterWherePart As String                                          ' Alles nach dem Wherestatement
        Dim szSQL2 As String
        Try                                                                     ' Fehlerbehandlung aktivieren
            szSQL = Trim(szSQL)                                                 ' Evtl. leerzeichen davor/dahinter abscheiden
            If InStr(szSQL.ToUpper, SQL_WHERE) > 0 And _
                InStr(szNewCondition.ToUpper, SQL_WHERE) > 0 Then               ' Wenn WHERE im SQL und New Condition
                szNewCondition = Replace(szNewCondition, SQL_WHERE, "")         ' Einen Rauschneiden
            End If
            ' Select From Where Group Order
            lngSQLLen = szSQL.Length                                            ' Länge des SQL Statements ermitteln
            lngUnionStartPos = InStr(szSQL.ToUpper, SQL_UNION)                  ' Feststellen ob Union enthalten
            lngWhereStartPos = InStr(szSQL.ToUpper, SQL_WHERE)                  ' Feststellen ob schon Where enthalten
            lngOrderStartPos = InStr(szSQL.ToUpper, SQL_ORDER)                  ' Feststellen ob Order By enthalten
            lngGroupStartPos = InStr(szSQL.ToUpper, SQL_GROUP)                  ' Feststellen ob Group By enthalten
            lngHavingStartPos = InStr(szSQL.ToUpper, SQL_HAVING)
            If lngUnionStartPos > 1 Then
                szSQL2 = Left(szSQL, lngUnionStartPos - 1)
                szSQL2 = AddWhereInFullSQL(szSQL2, szNewCondition, bOr, oBag)
                szSQL = Right(szSQL, Len(szSQL) - lngUnionStartPos + 1)
                szSQL = AddWhereInFullSQL(szSQL, szNewCondition, bOr, oBag)
                szSQL = szSQL2 & " " & szSQL
            ElseIf lngGroupStartPos + lngWhereStartPos + lngOrderStartPos = 0 Then  ' Einfaches SQL Statement ohne Where, Order usw.
                szSQL = szSQL & " " & AddWhere("", szNewCondition, bOr, oBag)
            Else
                If lngHavingStartPos > 0 Then lngWhereEndPos = lngHavingStartPos - 1
                If lngOrderStartPos > 0 Then lngWhereEndPos = lngOrderStartPos - 1
                If lngGroupStartPos > 0 Then lngWhereEndPos = lngGroupStartPos - 1
                If lngWhereEndPos = 0 Then lngWhereEndPos = Len(szSQL)
                If lngWhereStartPos = 0 Then lngWhereStartPos = lngWhereEndPos
                szSelectPart = Left(szSQL, lngWhereStartPos - 1)
                szAfterWherePart = Right(szSQL, lngSQLLen - lngWhereEndPos)
                If lngWhereStartPos < lngWhereEndPos Then
                    szWherePart = Mid(szSQL, lngWhereStartPos, lngWhereEndPos - lngWhereStartPos + 1)
                End If
                szSQL = szSelectPart & " " & AddWhere(szWherePart, szNewCondition, bOr, oBag) & " " & szAfterWherePart
            End If
            Return szSQL
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "AddWhereInFullSQL", ex)      ' Fehler behandlung aufrufen
            End If
            Return ""                                                           ' Misserfolg zurück
        End Try
    End Function

End Module
