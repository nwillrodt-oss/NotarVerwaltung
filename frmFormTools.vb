Module frmFormTools
    Private Const MODULNAME = "frmFormTools"                                    ' Modulname für Fehlerbehandlung

    Public Const TV_KEY_SEP = "\"

    'Public Function AddTVNode(ByVal TV As TreeView, _
    '                           ByVal NodeInfo As TVNodeInfo, _
    '                           Optional ByRef oBag As clsObjectBag = Nothing) As Boolean
    '    Dim newNode As TreeNode
    '    Try                                                                     ' Fehlerbehandlung aktivieren
    '        With NodeInfo
    '            If .ParentKey = "" Then                                         ' Rootnode anlegen
    '                newNode = TV.Nodes.Add(.Key, .Text, .ImageIndex)
    '                newNode.Tag = .Tag
    '            Else                                                            ' Sonst in den Baum einhängen
    '                'newNode = TV.Nodes.Add(.ParentKey, tvwChild, ParentKey & "\" & .Key, .Text)
    '                'newNode.Tag = newNode.Parent.Tag & TV_KEY_SEP & Tag         ' Tag setzen
    '            End If
    '        End With
    '        Return True                                                         ' Erfolg zurück
    '    Catch ex As Exception                                                   ' Fehler behandeln
    '        If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
    '            Call oBag.ErrorHandler(MODULNAME, "AddTVNode", ex)              ' Fehlermeldung
    '        End If
    '        Return False                                                        ' Misserfolg zurück
    '    End Try
    'End Function

    Public Function GetSelectTreeNode(ByVal TV As TreeView) As TreeNode
        On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
        GetSelectTreeNode = TV.SelectedNode                                     ' Ermittelt aktuell ausgewählten Tree node
        Err.Clear()                                                             ' Evtl Error Clearen
    End Function

    Public Function RefreshListView(ByVal LV As ListView, _
                                    ByVal TV As TreeView, _
                                    Optional ByVal cNode As TreeNode = Nothing, _
                                    Optional ByVal ID As String = "", _
                                    Optional ByRef oBag As clsObjectBag = Nothing) As Boolean
        ' Aktualisiert nur das Listview
        Dim szLvItemKey As String                                               ' List view Item Key
        Dim LVInfo As LVInfo                                                    ' Infos zum LV Handling aus XML
        Try                                                                     ' Fehlerbehandlung aktivieren
            If cNode Is Nothing Then                                            ' Kein Node angegeben
                cNode = GetSelectTreeNode(TV)                                   ' Dann Akt Node holen
            End If
            If cNode Is Nothing Then Return Nothing ' Immer noch kein Konten dann fertig
            If ID = "" Then ID = GetIDFromNode(cNode) ' Akt ID aus Node.Namen ermitteln
            With LVInfo
                'Call objTools.GetLVInfoFromXML(App.Path & "\" & INI_XMLFILE, cNode.Tag, .szSQL, .szTag, .szWhere, _
                '        .lngImage, .bValueList, .bListSubNodes)                 ' LV infos aus mxl datei holen
                'Call ListLVByTag(LV, ThisDBCon, cNode.Tag, ID, .bValueList, _
                '        cNode.Image)                                            ' Listitems anzeigen
                'If .bListSubNodes Then Call ListLVFromSubNodes(LV, _
                '        TV, cNode) ' Subnodes im LV anzeigen
                'If Not .bValueList Then
                '    Call CountLVItems(LV)                                       ' Anzahl der listitem in statusbar anzeigen
                'Else
                '    Call CountLVItems(LV, 1)                                    ' Wenn .bValueList anzahl ist immer 1
                'End If
            End With

        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "RefreshListView", ex)        ' Fehlermeldung
            End If
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

    Public Function FillLVByDS(ByVal LV As ListView, ByVal szTag As String, ByVal DSList As DataSet, _
        Optional ByVal bShowValueList As Boolean = False, _
        Optional ByVal lngImagindex As Integer = 0, _
        Optional ByVal szIndexField As String = "", _
        Optional ByVal bOptColumnWidth As Boolean = False, _
        Optional ByVal bNoColWidth As Boolean = False, _
        Optional ByVal szKey As String = "", _
        Optional ByVal lngAltImageIndex As Integer = 0, _
        Optional ByVal AltImageConditionField As String = "", _
        Optional ByVal AltImageConditionValue As Object = Nothing, _
        Optional ByVal AltImageConditionOperation As String = "", _
                                    Optional ByRef oBag As clsObjectBag = Nothing) As Boolean
        ' Füllt ListView indem anhand des RS
        ' bShowValueList = True gibt an ob ein DS untereinader angezeit wird Pro Feld ein Item)
        ' bShowValueList = False Listet alle DS auf
        ' bOptColumnWidth = true die Splatenbreite wird optimal eingestellt
        '   Sonst wird versicht die Splatenbreite aus der Reg zu laden
        Dim LVItem As ListViewItem                                              ' ListView item
        Dim i As Integer                                                        ' Counter
        Dim ci As Integer                                                       ' Akt ColumnIndex
        Dim szTmpName As String                                                 ' Item text
        Dim szTmpID As String                                                   ' Evtl DS ID
        Dim szTagArray() As String                                              ' LV.Tag in array aufgespalten
        Dim NewTag As String                                                    ' Evtl. Neu generierte Tag
        Dim szDetails As String                                                 ' Details für Fehlerbehandlung
        Dim lngUsedImgIndex As Integer                                          ' ImgIndex der tatsächlich verwendet wird
        Dim dTable As DataTable
        Dim dRow As DataRow
        Try                                                                     ' Fehlerbehandlung aktivieren
            If DSList Is Nothing Then Return False ' Kein Dataset -> Fertig
            If DSList.Tables.Count = 0 Then Return False ' Keine Daten -> Fertig
            dTable = DSList.Tables(0)                                           ' DataTable auslesen
            Call LV.Clear()                                                     ' ListView Aufräumen
            'If rs.RecordCount = 0 Then GoTo exithandler ' ?
            'If IsEmpty(lngImagindex) Then lngImagindex = 0 ' Defaultwert Imagindex
            'If IsEmpty(lngAltImageIndex) Then lngAltImageIndex = 0 ' Defaultwert Alternatives Image Indes
            If bShowValueList Then                                              ' Alle felder eines DS untereinader auflisten
                Call AddLVColumn(LV, "Eigenschaft", 2000, oBag)                 ' Feste Spalte Eigenschaft (enthält Feldnamen)
                Call AddLVColumn(LV, "Wert", 4000, oBag)                        ' Feste Spalte Wert (enthält Feldwert)

                If dTable.Rows.Count = 1 Then                                   ' Sind Zeilen vorhanden
                    'rsList.MoveFirst()                                            ' Nur ersten DS anzeigen
                    'szTmpID = Trim(objTools.checknull(rsList.Fields(0).Value, ""))
                    szTmpID = dTable.Rows(0).Item(0).ToString
                    szTagArray = Split(LV.Tag, TV_KEY_SEP)                      ' LV.Tag in array aufspalten
                    ' Das Letzte * im LV.Tag durch ID ersetzen
                    For i = 0 To UBound(szTagArray) - 1                         ' Bis auf letzten eintrag wieder zusammen setzen
                        NewTag = NewTag & szTagArray(i) & TV_KEY_SEP
                    Next
                    NewTag = NewTag & szTmpID                                   ' DS ID an NeuenTag anhängen
                    'For i = 0 To rsList.Fields.Count - 1                        ' Für jedes Feld eine Zeile
                    '    If Left(rsList.Fields(i).Name, 2) <> "ID" Then           ' Wenn Feld mit ID anfängt ausblenden
                    '        szTmpName = Trim(objTools.checknull(rsList.Fields(i).Value, ""))
                    '        LVItem = AddListViewItem(LV, rsList.Fields(i).Name, _
                    '                szTmpName, NewTag, lngImagindex)            ' ListViewItem anlegen
                    '        'LVItem.Key = szKey
                    '        LVItem.Icon = lngImagindex                          ' Icon Setzen
                    '        szTmpName = ""
                    '    End If
                    'Next i
                    For i = 0 To dTable.Columns.Count - 1
                        If Left(dTable.Columns(i).ColumnName, 2) <> "ID" Then
                            LVItem = AddListViewItem(LV, dTable.Rows(0).Item(i).ToString, , , , oBag)
                        End If
                    Next

                End If
            Else                                                                ' Alle DS untereinader auflisten
                'LV.MousePointer = vbHourglass                                   ' Sanduhr anzeigen
                ' ColumsHeader anlegen
                'For i = 1 To rsList.Fields.Count - 1                            ' Erste Spalte ist ID -> auslassen
                '    ci = AddLVColumn(LV, rsList.Fields(i).Name)                 ' Colum hinzufügen
                '    Select Case rsList.Fields(i).Type                           ' Daten Typ Prüfen
                '        Case adSmallInt, adInteger, adBigInt, adSingle          ' Zahlen
                '            LV.ColumnHeaders(ci).Tag = adInteger                ' DatenFormat in ColumnKey zum sortieren
                '        Case adDate, adDBDate, adDBTimeStamp, adDBTime          ' Datum
                '            LV.ColumnHeaders(ci).Tag = adDate                   ' DatenFormat in ColumnKey zum sortieren
                '        Case Else
                '    End Select
                'Next i                                                          ' Nächstes Feld
                For Each dCol As DataColumn In dTable.Columns                   ' Alle Spalten durchlaufen
                    ci = AddLVColumn(LV, dCol.ColumnName, 0, oBag)              ' Colum hinzufügen
                Next
                ' Jedes Listitem bekommet als Tag den Tag des LVs + Datensat ID eingetragen
                For Each dRow In dTable.Rows                                    ' Alle Datensätze durchlaufen
                    If szIndexField <> "" Then                                  ' explicit angegebenes ID feld
                        szTmpID = Trim(dRow.Item(szIndexField).ToString)        ' ID Auslesen
                    End If
                    If szTmpID = "" Then                                        ' Wenn keine ID gefunden
                        szTmpID = dRow.Item(0).ToString                         ' 1. Feldwert ID
                    End If
                    szTmpName = Trim(dRow.Item(0).ToString)                     ' 2. Feldwert ItemText
                    If szTmpName <> "" And szTmpID <> "" Then                   ' Nur wenn ID und Itemtext vorhanden
                        szDetails = "ItemKey: " & LV.Tag & TV_KEY_SEP & szTmpID & vbCrLf & "ItemTag: " & LV.Tag & TV_KEY_SEP & "*" & vbCrLf
                        ' DAs mit dem Alternativen Image überleg ich mir noch mal
                        LVItem = AddListViewItem(LV, szTmpName, , LV.Tag & TV_KEY_SEP & _
                                szTmpID, lngUsedImgIndex, oBag)                 ' ListView Item anlegen
                        szDetails = szDetails & "Item angelegt."                ' Detail infos für Fehlermeldung
                        LVItem.Tag = LV.Tag & TV_KEY_SEP & "*"                  ' Item Tag setzen
                        If szKey <> "" Then                                     ' Item Key Setzen
                            LVItem.Name = szKey & TV_KEY_SEP & szTmpID
                        Else
                            LVItem.Name = LV.Tag & TV_KEY_SEP & szTmpID
                        End If
                        For i = 2 To LV.Columns.Count                           ' Für jedes Feld ein SubItem
                            Call AddListViewSubItem(LVItem, dRow.Item(LV.Columns(i).Text).Value)  ' ListViewSubItem anlegen
                        Next i                                              ' nächstes Feld
                        szTmpName = ""
                        szTmpID = ""
                    End If
                Next
                'Do While Not rsList.EOF                                         ' Für jeden DS einen eintrag
                '    If szIndexField <> "" Then                                  ' explicit angegebenes ID feld
                '        szTmpID = Trim(objTools.checknull(rsList.Fields(szIndexField).Value, ""))
                '    End If
                '    If szTmpID = "" Then szTmpID = Trim(objTools.checknull( _
                '            rsList.Fields(0).Value, "")) ' Sonst 1. Feldwert ID
                '    szTmpName = Trim(objTools.checknull( _
                '            rsList.Fields(1).Value, ""))                        ' 2. Feldwert ItemText
                '    If szTmpID <> "" And szTmpName = "" Then szTmpName = " (Leer) "
                '    If szTmpName <> "" And szTmpID <> "" Then
                '        szDetails = "ItemKey: " & LV.Tag & TV_KEY_SEP & szTmpID & vbCrLf & "ItemTag: " & LV.Tag & TV_KEY_SEP & "*" & vbCrLf
                '        'Set LVItem = AddListViewItem(LV, szTmpName, , LV.Tag & TV_KEY_SEP & szTmpID, lngImageIndex)
                '        'LVItem.Key = LV.Tag & TV_KEY_SEP & "*"
                '        lngUsedImgIndex = lngImagindex                      ' ImageIndex Setzen
                '        If AltImageConditionField <> "" Then   ' Auf alternatives Image prüfen
                '            If objTools.CheckStringOperation(AltImageConditionOperation, _
                '                    objTools.checknull(rsList.Fields(AltImageConditionField).Value, ""), _
                '                    AltImageConditionValue) Then            '
                '                lngUsedImgIndex = lngAltImageIndex          ' Alternatives Image setzen
                '            End If
                '        End If

                '        LVItem = AddListViewItem(LV, szTmpName, , LV.Tag & TV_KEY_SEP & _
                '                szTmpID, lngUsedImgIndex)                   ' ListView Item anlegen
                '        szDetails = szDetails & "Item angelegt."            ' Detail infos für Fehlermeldung
                '        LVItem.Tag = LV.Tag & TV_KEY_SEP & "*"              ' Item Tag setzen
                '        If szKey <> "" Then                                 ' Item Key Setzen
                '            LVItem.Key = szKey & TV_KEY_SEP & szTmpID
                '        Else
                '            LVItem.Key = LV.Tag & TV_KEY_SEP & szTmpID
                '        End If
                '        For i = 2 To LV.ColumnHeaders.Count                 ' Für jedes Feld ein SubItem
                '            Call AddListViewSubItem(LVItem, Trim(objTools.checknull(rsList.Fields( _
                '                    LV.ColumnHeaders(i).Text).Value, "")))  ' ListViewSubItem anlegen
                '        Next i                                              ' nächstes Feld
                '        szTmpName = ""
                '        szTmpID = ""
                '    End If
                '    rsList.MoveNext()                                         ' Nächster DS
                'Loop
            End If

            'If Not bNoColWidth Then
            '    If LV.ListItems.Count > 0 And bOptColumnWidth Then
            '        Call OptimalHeaderWidth(LV)                             ' Optiomale Spalten breite einstellen
            '    Else
            '        Call LoadColumnWidth(LV, "", True)                      ' Spalten breite anhand Tag aus Registry
            '    End If
            'End If

            Return True
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "FillLVByDS", ex)             ' Fehlermeldung
            End If
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Public Function FillLVBySQL(ByVal LV As ListView, _
                                ByVal szSQL As String, _
                                ByVal dbCon As clsDBConnect, _
                                Optional ByVal bShowValueList As Boolean = False, _
                                Optional ByVal lngImageIndex As Integer = 0, _
                                Optional ByVal szIndexField As String = "", _
                                Optional ByVal bOptColumnWidth As Boolean = False, _
                                Optional ByVal bNoColWidth As Boolean = False, _
                                Optional ByVal szKey As String = "", _
                                Optional ByVal lngAltImageIndex As Integer = -1, _
                                Optional ByVal AltImageConditionField As String = "", _
                                Optional ByVal AltImageConditionValue As String = "", _
                                Optional ByRef oBag As clsObjectBag = Nothing) As DataSet
        ' Füllt ein Listview mit Items aus SQl Statement
        ' bShowValueList = True gibt an ob ein DS untereinader angezeit wird Pro Feld ein Item)
        ' bShowValueList = False Listet alle DS auf
        ' bOptColumnWidth = true die Splatenbreite wird optimal eingestellt
        '   Sonst wird versicht die Splatenbreite aus der Reg zu laden
        Dim DSList As New DataSet                                               ' RS mit Daten
        Dim szDetails As String                                                 ' Detailinfos für Fehlerbehandlung
        Try                                                                     ' Fehlerbehandlung aktivieren
            If szSQL = "" Then Return Nothing ' Kein SQL -> raus
            LV.Clear()                                                          ' ListItems und Columns löschen
            'If lngImageIndex = -1 Then lngImageIndex = 0 ' Default wert für image index
            szDetails = "SQL: " & szSQL
            DSList = dbCon.FillDS(szSQL, True)                                  ' Daten holen
            'If rsList Is Nothing Then GoTo exithandler                          ' Keine Daten (fehler) -> Fertig
            Call FillLVByDS(LV, "", DSList, bShowValueList, lngImageIndex, szIndexField, _
                    bOptColumnWidth, bNoColWidth, szKey, lngAltImageIndex, _
                    AltImageConditionField, AltImageConditionValue)             ' ListView aus Recordet füllen
            Return DSList                                                       ' DS mit daten als Rückgabe wert
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "FillLVBySQL", ex)            ' Fehlermeldung
            End If
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

    Public Function FillLV(ByVal LV As ListView, _
                           ByVal dbCon As Object, _
                           ByVal szTag As String, _
                           Optional ByVal szWhereKey As String = "", _
                           Optional ByVal bShowValueList As Boolean = False, _
                           Optional ByVal lngImageIndex As Integer = 0, _
                           Optional ByVal bNotColWidth As Boolean = False, _
                           Optional ByVal szKey As String = "", _
                           Optional ByRef oBag As clsObjectBag = Nothing) As Boolean
        ' Füllt ListView indem anhand des Tag ein SQLStatement aus XML gelesen Wird
        'Dim LVInfo As ListViewInfo                                      ' ListView Infos
        Dim bShowDel As Boolean                                                 ' Als gelöscht gesetzte DS anzeigen
        Try                                                                     ' Fehlerbehandlung altivieren
            LV.Clear()                                                          ' ListView Aufräumen
            'bShowDel = objOptions.GetOptionByName(OPTION_SHOWDELREL)            ' Option Gelöschte DS anzeigen auslesen
            'LV.Tag = szTag '& TV_KEY_SEP & "*"                                  ' Tag des Listviews setzen
            With LVInfo
                If .szSQL = "" Then Return False ' Kein SQL Statement -> Fertig
                'Call objTools.GetLVInfoFromXML(App.Path & "\" & INI_XMLFILE, szTag, .szSQL, .szTag, _
                '        .szWhere, .lngImage, .bValueList, .bListSubNodes, , , , .AltImage, .AltImgField, _
                '            .AltImgValue, .DelFlagField)                        ' Listview infos aus XMl datei holen
                If szWhereKey <> "" And .szSQL <> "" Then                       ' Evtl. Where bedingung anhängen
                    If .szWhere <> "" Then .szWhere = .szWhere & "'" & szWhereKey & "'"
                    'If .DelFlagField <> "" Then                                 ' Gibt es ein gelöscht flag ?
                    '    .WhereNoDel = AddWhere(.szWhere, .DelFlagField & "=0", False, oBag) ' Delflag mit in Where einbauen
                    'Else
                    '    .WhereNoDel = .szWhere                                  ' Sont Where statements gleich
                    'End If
                    'If bShowDel Then                                            ' Sollen als gelöscht gekennzeichnete DS angezeigt werden?
                    '    If .szWhere <> "" Then .szSQL = AddWhereInFullSQL(.szSQL, .szWhere, False, oBag) ' Where Ohne Delflag Filter
                    'Else
                    '    If .WhereNoDel <> "" Then .szSQL = AddWhereInFullSQL(.szSQL, .WhereNoDel, False, oBag) ' Where mit delFlag Filter
                    'End If
                End If
                lngImageIndex = .lngImage                                       ' Imageindex Setzen
                'If .lngImage <> "" And IsNumeric(.lngImage) Then lngImageindex = CLng(.lngImage) ' Evtl. Image index  aus Ini lesen

                Return FillLVBySQL(LV, .szSQL, dbCon, .bValueList, lngImageIndex, _
                        , , bNotColWidth, szKey, .AltImage, .AltImgField, _
                        .AltImgValue)                                           ' Listview aus SQL statement füllen
            End With

        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "RefreshListView", ex)        ' Fehlermeldung
            End If
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

    Public Function GetIDFromNode(ByVal cNode As TreeNode, _
                                  Optional ByRef oBag As clsObjectBag = Nothing) As String
        ' Ermitteli DS ID aus cNode.Namen
        Dim szTmp As String                                                     ' Hilfsvariable
        Dim ID As String                                                        ' Evtl ID des Detaildatensatzes
        Dim i As Integer                                                        ' Counter
        Try                                                                     ' Fehlerbehandlung altivieren
           
            If InStr(cNode.Name, "ID:") > 0 Then                                ' detaildatensatz
                ID = Replace(cNode.Name, "ID:", "")                             ' ID aus Namen ermitten
                GetIDFromNode = ID                                              ' ID Zurück
            Else                                                                ' Wenn kein Detail Datensatz
                Return ""
                'If InStr(cNode.FullPath, "ID:") Then                               ' Statischer unterknoten eines Detaildatensatzes
                '    szTmp = ""                                              ' Tmp Leeren
                '    i = UBound(szTagArray) + 1                              ' Max Arra Index festlegen
                '    While szTmp <> "*"                                      ' Tag bis * rückwärts durchlaufen
                '        i = i - 1                                           ' arrayindex herunterzählen
                '        szTmp = szTagArray(i)                               ' array wert merken
                '        ID = szKeyArray(i)                                  ' ID am gleicher stelle aus Tag
                '    End While
                'End If
                'If ID = "" And InStr(cNode.Key, TV_KEY_SEP) > 0 Then ID = GetLastKey(cNode.Key, TV_KEY_SEP) ' sonst mit gewalt
            End If
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "GetIDFromNode", ex)          ' Fehlermeldung
            End If
            Return ""                                                           ' Misserfolg zurück
        End Try
    End Function

    Public Function AddLVColumn(ByVal LV As ListView, ByVal szHeaderName As String, _
                            Optional ByVal ColWidith As Integer = 0, _
                            Optional ByRef oBag As clsObjectBag = Nothing) As Integer
        ' Fügt eine Spalte dem Listview hinzu
        Dim i As Integer                                                        ' counter
        Try                                                                     ' Fehlerbehandlung aktivieren
            'ColWidith = CLng(ColWidith)                                     ' Fall ColWidth nicht angegeben
            If szHeaderName <> "" Then                                          ' Header Name vorhanden
                For i = 1 To LV.Columns.Count                                   ' Feststellen ob es den Header schon gibt
                    If LV.Columns(i).Text = szHeaderName Then
                        Return i                                                ' Column index als Rückgabe wert
                    End If
                Next i                                                          ' Nächste Spalte
                LV.HeaderStyle = ColumnHeaderStyle.Clickable                    ' ColumHeaders Anzeigen einstellen
                If ColWidith <> 0 Then                                          ' Spaltenbreite angegeben
                    LV.Columns.Add(szHeaderName, ColWidith)                     ' ColumHeader mit breite setzen
                Else                                                            ' Sonst
                    LV.Columns.Add(szHeaderName)                                ' ColumnHeader ohne breite setzen
                End If
                Return LV.Columns.Count                                         ' Column index als Rückgabe wert
            Else
                LV.Columns.Add("", 0)                                           ' Unsichtbare spalte sezen
                Return LV.Columns.Count                                         ' Column index als Rückgabe wert
            End If
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "AddLVColumn", ex)            ' Fehlermeldung
            End If
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

    Public Function AddListViewItem(ByVal ctlListView As ListView, _
                                ByVal szItemText As String, _
                                Optional ByVal szSubItemText As String = "", _
                                Optional ByVal szValueName As String = "", _
                                Optional ByVal intImage As Integer = 0, _
                                Optional ByRef oBag As clsObjectBag = Nothing) As ListViewItem
        ' Legt ein ListView Item an und evt. ein Sub Item
        ' szValuename wird der Item Tag
        Dim itemX As ListViewItem                                               ' angelegtes Listview item
        Try                                                                     ' Fehlerbehandlung aktivieren
            itemX = ctlListView.Items.Add(szItemText, intImage)                 ' List Item anlegen
            'itemX.SmallIcon = intImage                                          ' Item Image setzten
            itemX.Tag = szValueName                                             ' Tag setzen
            If szSubItemText <> "" Then Call AddListViewSubItem(itemX, szSubItemText) ' Evtl. sub item setzen
            Return itemX                                     ' Item als RÜckgabe wert
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "AddListViewItem", ex)        ' Fehlermeldung
            End If
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

    Public Sub AddListViewSubItem(ByRef itemMain As ListViewItem, _
                              ByVal szSubItemText As String, _
                              Optional ByRef oBag As clsObjectBag = Nothing)
        ' Legt ein Sub Item an (nur in ListView mit ansicht Report sichtbar)
        'Dim Index As Integer                                                    ' Neuer SubItem index
        Try                                                                     ' Fehlerbehandlung aktivieren
            'Index = itemMain.SubItems.Count + 1                                 ' Neuen Index Ermitteln
            'itemMain.SubItems(Index) = Left(szSubItemText, 50)                  ' SubItem anlegen
            itemMain.SubItems.Add(szSubItemText)

        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "AddListViewItem", ex)        ' Fehlermeldung
            End If
        End Try
    End Sub

End Module
