VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSQLTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const MODULNAME = "modSQLTools"                             ' Modulname für Fehlerbehandlung

Private objObjectBag As Object                                      ' ObjectBag object
Private objError As Object                                          ' Error Object

Private Sub Class_Terminate()
On Error Resume Next                                                ' Hier keine Fehlerbehandlung mehr
    Set objError = Nothing
    Set objObjectBag = Nothing
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Public Function InitObjectBag(objOb As Object)

    Set objObjectBag = objOb
    If Not objObjectBag Is Nothing Then
        Set objError = objObjectBag.GetErrorObj
    End If
    
End Function

                                                                    ' *****************************************
                                                                    ' Methoden
Public Function GetFieldFromRSName(ByVal szSQL As String, szRSFieldName As String) As String
    Dim FieldArray() As String                                      ' Feldarray
    Dim tmpField As String
    Dim AliasStartPos As String
    Dim LeerPos As Integer
    Dim CasePos As Integer
    Dim AsPos As Integer
    Dim i As Integer                                                ' counter
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szSQL = Left(szSQL, InStr(UCase(szSQL), "FROM") - 1)            ' Startpos ermitteln
    FieldArray = Split(szSQL, ",")
    For i = 0 To UBound(FieldArray)
        FieldArray(i) = Trim(FieldArray(i))
        If InStr(FieldArray(i), szRSFieldName) > 0 Then
            FieldArray(i) = Trim(Replace(UCase(FieldArray(i)), SQL_SELECT, ""))
            AliasStartPos = InStr(Trim(UCase(FieldArray(i))), " " & Trim(UCase(szRSFieldName)))
            If AliasStartPos > 0 Then
                'tmpField = Trim(Left(FieldArray(i), AliasStartPos))
                AsPos = InStr(UCase(FieldArray(i)), "AS ")
                LeerPos = InStr(UCase(FieldArray(i)), " ")
                CasePos = InStr(UCase(FieldArray(i)), "CASE ")
                If CasePos > 0 Then
                    FieldArray(i) = Trim(Right(FieldArray(i), Len(FieldArray(i)) - 4))
                    LeerPos = InStr(UCase(FieldArray(i)), " ")
                    GetFieldFromRSName = UCase(Trim(Left(FieldArray(i), LeerPos)))
                    Exit For
                ElseIf AsPos > 0 Then
                    If AliasStartPos < AsPos And AliasStartPos > 1 Then
                    
                    Else
                        GetFieldFromRSName = UCase(Trim(Left(FieldArray(i), AsPos - 1)))
                        Exit For
                    End If
                    
                ElseIf LeerPos > 0 And AsPos = 0 Then
                    GetFieldFromRSName = UCase(Trim(Left(FieldArray(i), LeerPos)))
                    Exit For
                    
                ElseIf AsPos + LeerPos + CasePos = 0 Then
                   GetFieldFromRSName = UCase(Trim(FieldArray(i)))
                   Exit For
                Else
                   
                End If
            Else
                If Trim(UCase(FieldArray(i))) = Trim(UCase(szRSFieldName)) Then
                    GetFieldFromRSName = UCase(Trim(FieldArray(i)))
                     Exit For
                End If
                If InStr(Trim(UCase(FieldArray(i))), "'" & Trim(UCase(szRSFieldName)) & "'") > 0 Then
                    If InStr(Trim(UCase(FieldArray(i))), " AS") > 0 Then
                        
                        GetFieldFromRSName = Left(UCase(Trim(FieldArray(i))), InStr(Trim(UCase(FieldArray(i))), " AS"))
                        
                        Exit For
                    End If
                End If
            End If
        End If
    Next i
exithandler:
On Error Resume Next
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "GetFieldFromRSName", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function GetOneFieldSQL(szSQL As String) As String
' Kürzt die feldliste des Select statements auf ein Feld
    Dim FieldArray() As String
    Dim szFromPart As String
    Dim szResult As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If InStr(szSQL, SQL_FROM) > 0 Then
        szFromPart = Right(szSQL, Len(szSQL) - InStr(szSQL, SQL_FROM) + 1)
        
        FieldArray = Split(szSQL, ",")
        szResult = FieldArray(0) & " " & szFromPart
    End If
    GetOneFieldSQL = szResult
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetOneFieldSQL", errNr, errDesc)
    Resume exithandler
End Function

Public Function GetTableFromSQL(szSQL As String) As String
    
    Dim szFrom As String
    Dim LeerPos As Integer
    Dim KommaPos As Integer
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    szFrom = GetFromFromSQL(szSQL)
    If Left(szFrom, 1) = "(" Then
        szFrom = Trim(Mid(szFrom, 2, Len(szFrom)))
    End If
    If szFrom = "" Then GoTo exithandler
    
    LeerPos = InStr(szFrom, " ")
    KommaPos = InStr(szFrom, ",")
    If LeerPos > 0 Or KommaPos > 0 Then    ' Join oder mehrere Tabellen
    ' Nur erste Tabelle ausgeben (Notlösung)
    
        If LeerPos > KommaPos Then
            GetTableFromSQL = Trim(Left(szFrom, LeerPos))
        Else
            GetTableFromSQL = Trim(Left(szFrom, KommaPos))
        End If
    Else
        GetTableFromSQL = szFrom
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
    Call objError.Errorhandler(MODULNAME, "GetTableFromSQL", errNr, errDesc)
    Resume exithandler
End Function

Public Function GetFromFromSQL(szSQL As String, Optional bWithFrom As Boolean)
' Liefert den FROM Part des SQL Statements

    Dim FromStartPos As Integer
    Dim FromEndPos As Integer
    Dim szFrom As String
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    If InStr(UCase(szSQL), "FROM") = 0 Then GoTo exithandler        ' Kein FROM
    
    FromStartPos = InStr(UCase(szSQL), SQL_FROM)                    ' Startpos ermitteln
    
    If InStr(FromStartPos, szSQL, SQL_WHERE) > 0 Then               ' Where vorhanden
        FromEndPos = InStr(UCase(szSQL), SQL_WHERE)
    ElseIf InStr(UCase(szSQL), SQL_ORDER) > 0 Then                  ' Order by vorhanden
        FromEndPos = InStr(UCase(szSQL), SQL_ORDER)
    Else
        FromEndPos = Len(szSQL)
    End If
    
    If Not bWithFrom Then FromStartPos = FromStartPos + 4
    szFrom = Mid(szSQL, FromStartPos, FromEndPos - FromStartPos + 1)
    GetFromFromSQL = Trim(szFrom)
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetFromFromSQL", errNr, errDesc)
    Resume exithandler
End Function

Public Function GetIDField(szTabname As String) As String

On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    If szTabname = "" Then GoTo exithandler
    GetIDField = "ID" & Right(szTabname, 3)
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetIDField", errNr, errDesc)
    Resume exithandler
End Function

Public Function ReplaceWidcarts(dbConn As Object, szSearch As String) As String

    If InStr(szSearch, "*") > 0 Then szSearch = Replace(szSearch, "*", "%")
 
    'If dbConn.IsSQLServer Then
    '    If InStr(szSearch, "*") > 0 Then szSearch = Replace(szSearch, "*", "%")
    'Else
    '    If InStr(szSearch, "%") > 0 Then szSearch = Replace(szSearch, "%", "*")
    'End If
    ReplaceWidcarts = szSearch
    
End Function

Public Function IsReservedWord(szString) As Boolean
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If szString = "" Then GoTo exithandler                          ' Kein String -> Fertig
    szString = Trim(szString)
    Select Case UCase(szString)
    Case SQL_SELECT
        IsReservedWord = True
    Case SQL_FROM
        IsReservedWord = True
    Case "ORDER", "BY"
        IsReservedWord = True
    Case SQL_WHERE
        IsReservedWord = True
    Case SQL_UNION
        IsReservedWord = True
    Case "INNER", "JOIN", "LEFT", "RIGHT"
        IsReservedWord = True
    Case SQL_AS
        IsReservedWord = True
    Case SQL_ON
        IsReservedWord = True
    Case "INSERT", "INTO"
        IsReservedWord = True
    Case SQL_DELETE
        IsReservedWord = True
    Case SQL_VALUES
        IsReservedWord = True
    Case SQL_UPDATE
        IsReservedWord = True
    Case SQL_SET
        IsReservedWord = True
    Case "GROUP", "BY"
        IsReservedWord = True
    Case SQL_HAVING
        IsReservedWord = True
    Case SQL_ALTER
        IsReservedWord = True
    Case SQL_TAB
        IsReservedWord = True
    Case Else
    
    End Select

exithandler:
On Error Resume Next
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "IsReservedWord", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function AddWhereInFullSQL(szSQL As String, _
        ByVal szNewCondition As String, _
        Optional bOr As Boolean) As String
    Dim lngWhereStartPos As Integer                                 ' Startpos. den Wherestatements im SQL Statement
    Dim lngWhereEndPos As Integer                                   ' Endpos. den Wherestatements im SQL Statement
    Dim lngOrderStartPos As Integer                                 ' Startpos. den Orderstatements im SQL Statement
    Dim lngGroupStartPos As Integer                                 ' Startpos. den Groupstatements im SQL Statement
    Dim lngHavingStartPos As Integer                                ' Startpos. den Havingstatements im SQL Statement
    Dim lngUnionStartPos As Integer
    Dim lngSQLLen As Integer                                        ' länge des SQL Statements
    Dim szWherePart As String                                       ' Wherestatement
    Dim szSelectPart As String                                      ' Alles vor dem Wherestatement
    Dim szAfterWherePart As String                                  ' Alles nach dem Wherestatement
    Dim szSQL2 As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If InStr(UCase(szSQL), "WHERE") > 0 And _
        InStr(UCase(szNewCondition), "WHERE") > 0 Then              ' Wenn WHERE im SQL und New Condition
        szNewCondition = Replace(szNewCondition, "WHERE", "")       ' Einen Rauschneiden
    End If
     ' Select From Where Group Order
    szSQL = Trim(szSQL)
    lngSQLLen = Len(szSQL)
    lngUnionStartPos = InStr(UCase(szSQL), SQL_UNION)               ' Feststellen ob Union enthalten
    lngWhereStartPos = InStr(UCase(szSQL), SQL_WHERE)               ' Feststellen ob schon Where enthalten
    lngOrderStartPos = InStr(UCase(szSQL), SQL_ORDER)               ' Feststellen ob Order By enthalten
    lngGroupStartPos = InStr(UCase(szSQL), SQL_GROUP)               ' Feststellen ob Group By enthalten
    lngHavingStartPos = InStr(UCase(szSQL), SQL_HAVING)
    If lngUnionStartPos > 1 Then
        szSQL2 = Left(szSQL, lngUnionStartPos - 1)
        szSQL2 = AddWhereInFullSQL(szSQL2, szNewCondition, bOr)
        szSQL = Right(szSQL, Len(szSQL) - lngUnionStartPos + 1)
        szSQL = AddWhereInFullSQL(szSQL, szNewCondition, bOr)
        szSQL = szSQL2 & " " & szSQL
    
    ElseIf lngGroupStartPos + lngWhereStartPos + lngOrderStartPos = 0 Then  ' Einfaches SQL Statement ohne Where, Order usw.
        szSQL = szSQL & " " & AddWhere("", szNewCondition, bOr)
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
        szSQL = szSelectPart & " " & AddWhere(szWherePart, szNewCondition, bOr) & " " & szAfterWherePart

    
    End If
    AddWhereInFullSQL = szSQL
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "AddWhereInFullSQL", errNr, errDesc)
    Resume exithandler
End Function

Public Function AddWhere(ByVal szWhere As String, _
        szNewCondition As String, _
        Optional bOr As Boolean) As String
    Dim szAnd As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szWhere = Trim(szWhere)                                         ' Alte Condition Trimmen
    szNewCondition = Trim(szNewCondition)                           ' Neue Condition Trimmen
    
    If Left(UCase(szWhere), 6) = SQL_WHERE & " " Then
        szWhere = Right(szWhere, Len(szWhere) - 6)                  ' Falls Where davor -> abschneiden
    End If
    If Left(UCase(szNewCondition), 6) = SQL_WHERE & " " Then
        szNewCondition = Right(szNewCondition, Len(szNewCondition) - 6) ' Falls Where davor -> abschneiden
    End If
    If szWhere <> "" Then szWhere = "(" & szWhere & ")"             ' klammern
    szNewCondition = "(" & szNewCondition & ")"
    szAnd = " AND "                                                 ' Verknüpfung bstimmen AND / OR
    If bOr Then szAnd = " OR "
    If szWhere <> "" Then
        szWhere = "(" & szWhere & szAnd & szNewCondition & ")"      ' Zusammensetzen
    Else
        szWhere = szNewCondition
    End If
    If Left(szWhere, 6) <> SQL_WHERE & " " Then szWhere = SQL_WHERE & " " & szWhere ' Schlüsselwort voran
    AddWhere = szWhere
exithandler:
On Error Resume Next
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "AddWhere", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function BuildInsertSQL(szInsertTab As String, _
        szFieldlist As String, _
        szValuelist As String, _
        Optional bCRInInsert As Boolean) As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If szInsertTab = "" Then GoTo exithandler                       ' Keine Zieltabelle angegeben -> Fertig
    If szFieldlist = "" Then GoTo exithandler                       ' Keine Feldliste angegeben -> Fertig
    If szValuelist = "" Then GoTo exithandler                       ' Keine Werteliste angegeben -> Fertig
    BuildInsertSQL = SQL_INSERT & " " & szInsertTab & " ( " & szFieldlist & " ) " ' Insert TAb mit Feldliste
    If bCRInInsert Then BuildInsertSQL = BuildInsertSQL & vbCrLf    ' Evtl. Zeilenumbruch
    BuildInsertSQL = BuildInsertSQL & SQL_VALUES & " ( " & szValuelist & ") " ' Werteliste
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "BuildInsertSQL", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function BuildDeleteAllSQL(szDelTable As String) As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If szDelTable = "" Then GoTo exithandler                        ' Keine Tabele angegeben -> fertig
    BuildDeleteAllSQL = SQL_DELETE & " " & SQL_FROM & " " & szDelTable    ' Delete aller DS
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "BuildDeleteAllSQL", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function BuildDeleteSQL(ByVal szDeleteTable As String, _
        Optional szWhere As String) As String
    Dim szSQL As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If szDeleteTable = "" Then GoTo exithandler                     ' Keine Tabelle angegeben -> Fertig
    If Len(szDeleteTable) <= 4 Then GoTo exithandler                '
    If UCase(Left(Trim(szDeleteTable), 4)) <> SQL_FROM Then szDeleteTable = " " & SQL_FROM & " " & szDeleteTable
    BuildDeleteSQL = SQL_DELETE & " " & szDeleteTable & " " & szWhere   ' Statement zusamensetzen und zurück
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "BuildDeleteSQL", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function BuildUpdateSQL(szUpdateTab As String, _
        szValuelist As String, _
        szWhere As String) As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If szUpdateTab = "" Then GoTo exithandler
    If szValuelist = "" Then GoTo exithandler
    BuildUpdateSQL = SQL_UPDATE & " " & szUpdateTab & " " _
            & SQL_SET & " " & szValuelist & " " & szWhere
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "BuildUpdateSQL", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function BuildSelectSQL(ByVal szFromTable As String, _
        szFieldlist As String, _
        szWhere As String, Optional szOrder As String) As String
    Dim szSQL As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If szFromTable = "" Then GoTo exithandler                       ' Keine Tabelle angegeben -> Fertig
    If Len(szFromTable) <= 4 Then GoTo exithandler                  '
    If UCase(Left(Trim(szFromTable), 4)) <> SQL_FROM Then szFromTable = " " & SQL_FROM & " " & szFromTable
    If szFieldlist = "" Then szFieldlist = " * "                    ' Keine Fieldlistangegeben dan mit *
    BuildSelectSQL = SQL_SELECT & " " & szFieldlist & " " & szFromTable _
            & " " & szWhere & " " & szOrder                         ' Statement zusamensetzen und zurück
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "BuildSelectSQL", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function




