Imports Notarverwaltung.ClsTreeNodeEx

Module frmFormTools
    Private Const MODULNAME = "frmFormTools"                                    ' Modulname für Fehlerbehandlung

    Public Const TV_KEY_SEP = "\"

    Public Function SetFormIcon(ByVal frm As Form, _
                                ByVal IL As ImageList, _
                                ByVal index As Integer, _
                                Optional ByRef oBag As clsObjectBag = Nothing)
        Try                                                                     ' Fehlerbehandlung aktivieren
            If Not IL Is Nothing Then
                Dim hbitmap As Bitmap = IL.Images(index)
                Dim hIcon As IntPtr = hbitmap.GetHicon
                frm.Icon = Icon.FromHandle(hIcon)
            End If
            'frmEdit.Icon = frmParent.ILTree.ListImages(CLng(ThisClass.GetImage)).Picture
            'If Err() > 0 Then
            '    Err.Clear()
            '    frmEdit.Icon = frmEdit.ILTree.ListImages(CLng(ThisClass.GetImage)).Picture  ' Form Icon aus XML setzen
            'End If
            'If Err() > 0 Then                                                 ' Wenn Fehler
            '    Err.Clear()
            '    frmEdit.Icon = frmEdit.ILTree.ListImages(1).Picture         ' Form Icon 1 setzen
            'End If
            'If Err() > 0 Then                                                 ' Wenn Fehler
            '    frmEdit.Icon = LoadPicture("images\" & ThisClass.Name & ".ico") ' Form Icon aus Filesystem setzen
            '    Err.Clear()
            'End If
            Return True
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "SetFormIcon", ex)            ' Fehlermeldung
            End If
            Return False
        End Try
    End Function

#Region "TreeView Funktionen"

    Public Function GetSelectTreeNode(ByVal TV As TreeView) As TreeNode
        On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
        GetSelectTreeNode = TV.SelectedNode                                     ' Ermittelt aktuell ausgewählten Tree node
        Err.Clear()                                                             ' Evtl Error Clearen
    End Function

    Public Function SelectTreeNode(ByVal TV As TreeView, _
                                   ByVal SelectedNode As clsTreeNodeEx, _
                                   Optional ByVal bExpand As Boolean = False)
        On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
        TV.SelectedNode = SelectedNode                                          ' SelectedNode auf ausgewählt setzen
        If bExpand Then TV.SelectedNode.Expand() ' Expanden
        Err.Clear()                                                             ' Evtl Error Clearen
    End Function

    Public Function GetTreeNodeByKey(ByVal TV As TreeView, _
                                     ByVal szKey As String, _
                                     Optional ByVal oBag As clsObjectBag = Nothing) As clsTreeNodeEx
        ' ermittelt Node aus angegebenen Node Key
        Dim cNode As clsTreeNodeEx
        Dim fNode As clsTreeNodeEx
        Try                                                                     ' Fehlerbehandlung aktivieren
            'FullPatharray = Split(szKey, TV.PathSeparator)
            For Each cNode In TV.Nodes
                If cNode.Name = szKey Then
                    Return cNode
                Else
                    fNode = GetChildnodeByKey(cNode, szKey, oBag)
                    If Not IsNothing(fNode) Then Return fNode
                End If
            Next
            Return Nothing                                                      ' Wenn wir hier ankommen haben wir nichts gefunden
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "GetTreeNodeByKey", ex)       ' Fehlermeldung
            End If
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

    Private Function GetChildnodeByKey(ByVal pNode As clsTreeNodeEx, _
                                     ByVal szKey As String, _
                                     Optional ByVal oBag As clsObjectBag = Nothing) As clsTreeNodeEx
        Dim fNode As clsTreeNodeEx
        Try                                                                     ' Fehlerbehandlung aktivieren
            For Each cNode In pNode.Nodes
                If cNode.Name = szKey Then
                    Return cNode
                Else
                    fNode = GetChildnodeByKey(cNode, szKey, oBag)
                    If Not IsNothing(fNode) Then Return fNode
                End If
            Next
            Return Nothing                                                      ' Wenn wir hier ankommen haben wir nichts gefunden
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "GetTreeNodeByKey", ex)       ' Fehlermeldung
            End If
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

#End Region

#Region "ListView Funktionen"

    Public Function RefreshListView(ByVal LV As ListView, _
                                    ByVal TV As TreeView, _
                                    ByVal cNode As clsTreeNodeEx, _
                                    Optional ByVal ID As String = "", _
                                    Optional ByRef oBag As clsObjectBag = Nothing) As Boolean
        ' Aktualisiert nur das Listview
        Dim LVInfo As New ListViewInfoEx                                        ' Infos zum LV Handling aus XML
        Try                                                                     ' Fehlerbehandlung aktivieren
            If cNode Is Nothing Then                                            ' Kein Node angegeben
                cNode = GetSelectTreeNode(TV)                                   ' Dann Akt Node holen
            End If
            If cNode Is Nothing Then Return Nothing ' Immer noch kein Konten dann fertig
            If ID = "" Then ID = cNode.ID ' Akt ID aus Node.Namen ermitteln
            'LVInfo = cNode.ExListInfo                                           ' ListView informationen aus TreeNodeEx übernehmen
            'With LVInfo
            'Call objTools.GetLVInfoFromXML(App.Path & "\" & INI_XMLFILE, cNode.Tag, .szSQL, .szTag, .szWhere, _
            '        .lngImage, .bValueList, .bListSubNodes)                 ' LV infos aus mxl datei holen
            '    Call FillLVByTag(LV, oBag.ObjDBConnect, cNode.Tag, ID, .bValueList, cNode.SelectedImageIndex, , , oBag) ' Listitems anzeigen
            Call FillLVByTag(LV, oBag.ObjDBConnect, cNode, ID, oBag) ' Listitems anzeigen
            If cNode.ExListInfo.bListSubNodes Then                              ' Sollen Subnodes im LV mitangezeigt werden
                Call ListLVFromSubNodes(LV, TV, cNode)                          ' Subnodes im LV anzeigen
            End If
            'If Not .bValueList Then
            '    Call CountLVItems(LV)                                       ' Anzahl der listitem in statusbar anzeigen
            'Else
            '    Call CountLVItems(LV, 1)                                    ' Wenn .bValueList anzahl ist immer 1
            'End If
            ' End With
            Return True
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "RefreshListView", ex)        ' Fehlermeldung
            End If
            Return False                                                      ' Misserfolg zurück
        End Try
    End Function

    Public Function FillLVByTag(ByRef LV As ListView, _
                         ByVal dbCon As Object, _
                         ByVal cNode As clsTreeNodeEx, _
                         Optional ByVal szWhereKey As String = "", _
                         Optional ByRef oBag As clsObjectBag = Nothing) As DataSet
        ' Füllt ListView indem anhand des Tag ein SQLStatement aus XML gelesen Wird
        'Dim bShowDel As Boolean                                                 ' Als gelöscht gesetzte DS anzeigen
        Dim LVInfo As New ListViewInfoEx                                        ' Infos zum LV Handling aus XML
        Try                                                                     ' Fehlerbehandlung altivieren
            If cNode Is Nothing Then Return Nothing ' Immer noch kein Konten dann fertig
            LVInfo = cNode.ExListInfo                                           ' ListView informationen aus TreeNodeEx übernehmen
            'bShowDel = objOptions.GetOptionByName(OPTION_SHOWDELREL)            ' Option Gelöschte DS anzeigen auslesen
            LV.Tag = cNode.Name '& TV_KEY_SEP & "*"                                  ' Tag des Listviews setzen
            With LVInfo
                If .SQL = "" Then Return Nothing ' Kein SQL Statement -> Fertig
                If szWhereKey <> "" And .SQL <> "" Then                         ' Evtl. Where bedingung anhängen

                    If .WHERE <> "" Then .WHERE = .WHERE & "'" & szWhereKey & "'"
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
                    .SQL = AddWhereInFullSQL(.SQL, .WHERE, , oBag)
                End If
                Return FillLVBySQL(LV, .SQL, dbCon, LVInfo, , , oBag)      ' Listview aus SQL statement füllen
            End With

        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "FillLVByTag", ex)            ' Fehlermeldung
            End If
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

    Public Function FillLVBySQL(ByRef LV As ListView, _
                                ByVal szSQL As String, _
                                ByVal dbCon As clsDBConnect, _
                                ByVal LVInfo As ListViewInfoEx, _
                                Optional ByVal szIndexField As String = "", _
                                Optional ByVal szKey As String = "", _
                                Optional ByRef oBag As clsObjectBag = Nothing) As DataSet
        ' Füllt ein Listview mit Items aus SQL Statement
        ' bShowValueList = True gibt an ob ein DS untereinader angezeit wird Pro Feld ein Item)
        ' bShowValueList = False Listet alle DS auf
        ' bOptColumnWidth = true die Splatenbreite wird optimal eingestellt
        '   Sonst wird versiucht die Splatenbreite aus der Reg zu laden
        Dim DSList As New DataSet                                               ' RS mit Daten
        Dim szDetails As String                                                 ' Detailinfos für Fehlerbehandlung
        Try                                                                     ' Fehlerbehandlung aktivieren
            If szSQL = "" Then Return Nothing ' Kein SQL -> raus
            szDetails = "SQL: " & szSQL
            DSList = dbCon.FillDS(szSQL, LV.Tag)                                  ' Daten holen
            If IsNothing(DSList) Then Return Nothing ' Keine Daten (fehler) -> Fertig
            Call FillLVByDS(LV, DSList, LVInfo, szIndexField, szKey, oBag) ' ListView aus Dataset füllen
            Return DSList                                                       ' DS mit daten als Rückgabe wert
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "FillLVBySQL", ex)            ' Fehlermeldung
            End If
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

    Public Function FillLVByDS(ByVal LV As ListView, _
                            ByVal DSList As DataSet, _
                            ByVal LVInfo As ListViewInfoEx, _
                            Optional ByVal szIndexField As String = "", _
                            Optional ByVal szKey As String = "", _
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
        Dim NewTag As String                                                    ' Evtl. Neu generierte Tag
        Dim szDetails As String                                                 ' Details für Fehlerbehandlung
        'Dim lngUsedImgIndex As Integer                                          ' ImgIndex der tatsächlich verwendet wird
        Dim dTable As DataTable
        Dim dRow As DataRow
        Try                                                                     ' Fehlerbehandlung aktivieren
            If DSList Is Nothing Then Return False ' Kein Dataset -> Fertig
            If DSList.Tables.Count = 0 Then Return False ' Keine Daten -> Fertig
            dTable = DSList.Tables(0)                                           ' DataTable auslesen
            Call LV.Clear()                                                     ' ListView Aufräumen
            If LVInfo.bValueList Then                                           ' Alle felder eines DS untereinader auflisten
                Call AddLVColumn(LV, "Eigenschaft", 200, oBag)                  ' Feste Spalte Eigenschaft (enthält Feldnamen)
                Call AddLVColumn(LV, "Wert", 400, oBag)                         ' Feste Spalte Wert (enthält Feldwert)
                NewTag = ""
                If dTable.Rows.Count = 1 Then                                   ' Sind Zeilen vorhanden                   
                    NewTag = LV.Tag
                    For i = 0 To dTable.Columns.Count - 1                       ' Alle Tabellen spalten durchlaufen
                        If Left(dTable.Columns(i).ColumnName, 2) <> "ID" Then   ' Das ID feld zeigen wir nicht mit an
                            LVItem = AddListViewItem(LV, dTable.Columns(i).ColumnName, _
                                                     dTable.Rows(0).Item(i).ToString, NewTag, LVInfo.ImageIndex, oBag)
                        End If
                    Next                                                        ' Nächste Spalte
                End If
            Else                                                                ' Alle DS untereinader auflisten
                For i = 1 To dTable.Columns.Count - 1                           ' Erste Spalte ist ID -> auslassen
                    ci = AddLVColumn(LV, dTable.Columns(i).ColumnName, , oBag)  ' Colum hinzufügen
                Next i
                ' Jedes Listitem bekommet als Tag den Tag des LVs + Datensat ID eingetragen
                szTmpID = ""
                For Each dRow In dTable.Rows                                    ' Alle Datensätze durchlaufen
                    If szIndexField <> "" Then                                  ' explicit angegebenes ID feld
                        szTmpID = Trim(dRow.Item(szIndexField).ToString)        ' ID Auslesen
                    End If
                    If szTmpID = "" Then                                        ' Wenn keine ID gefunden
                        szTmpID = dRow.Item(0).ToString                         ' 1. Feldwert ID
                    End If
                    szTmpName = Trim(dRow.Item(1).ToString)                     ' 2. Feldwert ItemText
                    If szTmpName <> "" And szTmpID <> "" Then                   ' Nur wenn ID und Itemtext vorhanden
                        szDetails = "ItemKey: " & LV.Tag & TV_KEY_SEP & szTmpID & vbCrLf & _
                                    "ItemTag: " & LV.Tag & TV_KEY_SEP & "*" & vbCrLf
                        ' DAs mit dem Alternativen Image überleg ich mir noch mal
                        LVItem = AddListViewItem(LV, szTmpName, , LV.Tag & TV_KEY_SEP & _
                                szTmpID, LVInfo.ImageIndex, oBag)                 ' ListView Item anlegen
                        szDetails = szDetails & "Item angelegt."                ' Detail infos für Fehlermeldung
                        LVItem.Tag = LV.Tag & TV_KEY_SEP & szTmpID              ' Item Tag setzen
                        If szKey <> "" Then                                     ' Item Key Setzen
                            LVItem.Name = szKey & TV_KEY_SEP & szTmpID
                        Else
                            LVItem.Name = LV.Tag & TV_KEY_SEP & szTmpID
                        End If
                        For i = 2 To LV.Columns.Count                           ' Für jedes Feld ein SubItem
                            Call AddListViewSubItem(LVItem, dRow.Item(LV.Columns(i - 1).Text).ToString, oBag)  ' ListViewSubItem anlegen
                        Next i                                              ' nächstes Feld
                        szTmpName = ""
                        szTmpID = ""
                    End If
                Next

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

    Public Function ListLVFromSubNodes(ByVal LV As ListView, _
                                      ByVal TV As TreeView, _
                                      ByVal ThisNode As clsTreeNodeEx, _
                                      Optional ByVal oBag As clsObjectBag = Nothing)
        ' fügt die 1. Ebene Unterknoten dem Listviw als Items hinzu
        Dim subNode As clsTreeNodeEx                                            ' sub node der im LV angezeigt werden soll
        Dim LVItem As ListViewItem                                              ' neu angelegte LV Item
        'Dim KeyArray() As String                                                ' Node Key in Array
        'Dim szDescription As String                                             ' Evtl Beschreibung des Subnodes
        'Dim Nodeinfo As TVNodeInfoEx                                             ' Infos des subnodes aus XML
        Try                                                                     ' Fehlerbehandlung aktivieren
            'KeyArray = Split(ThisNode.Key, TV_KEY_SEP)                          ' Node Key in Array aufspalten
            If LV.Columns.Count = 0 Then                                        ' Wenn keine Column header
                Call AddLVColumn(LV, ThisNode.Text, 200)                       ' einen anlegen
                Call AddLVColumn(LV, "Beschreibung", LV.Width - 200)           ' einen anlegen
            End If
            'subNode.
            If ThisNode.Nodes.Count > 0 Then                                    ' Wenn Node Subnodes Besitzt
                subNode = ThisNode.Nodes(0)                                     ' 1. Subnode holen
                While Not subNode Is Nothing                                    ' Solage Subnodes gefunden werden
                    If Not subNode Is Nothing Then                              ' Wenn SubNode nicht Nothing
                        'KeyArray = Split(subNode.Key, TV_KEY_SEP)               ' SubnodeKey aufspalten
                        With subNode.ExTVNodeInfo
                            LVItem = AddListViewItem(LV, subNode.Text, .Desc, , .ImageIndex) ' List Item anlegen
                            'LVItem.Tag = subNode.Tag                            ' Tag setzen
                            'LVItem.Key = subNode.Key                            ' Key setzen
                        End With
                        subNode = subNode.NextNode                              ' nächster subnode
                    End If
                End While
            End If
            Return True
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "ListLVFromSubNodes", ex)     ' Fehlermeldung
            End If
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Public Function AddLVColumn(ByRef LV As ListView, ByVal szHeaderName As String, _
                            Optional ByVal ColWidith As Integer = 0, _
                            Optional ByRef oBag As clsObjectBag = Nothing) As Integer
        ' Fügt eine Spalte dem Listview hinzu
        Dim i As Integer                                                        ' counter
        Try                                                                     ' Fehlerbehandlung aktivieren
            'ColWidith = CLng(ColWidith)                                     ' Fall ColWidth nicht angegeben
            If szHeaderName <> "" Then                                          ' Header Name vorhanden
                For i = 0 To LV.Columns.Count - 1                                   ' Feststellen ob es den Header schon gibt
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
            If szSubItemText <> "" Then Call AddListViewSubItem(itemX, szSubItemText, oBag) ' Evtl. sub item setzen
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
                Call oBag.ErrorHandler(MODULNAME, "AddListViewSubItem", ex)     ' Fehlermeldung
            End If
        End Try
    End Sub

#End Region

    Public Sub SetWindowSizeFromString(ByVal F As Form, _
                                       ByVal szSize As String, _
                                       Optional ByVal szDelimiter As String = "/", _
                                       Optional ByRef oBag As clsObjectBag = Nothing)
        Dim szSizeArray() As String                                             ' Array mit größen
        Try                                                                     ' Fehlerbehandlung aktivieren
            If F Is Nothing Then Exit Sub ' Kein Form dann fertig
            If szSize <> "" Then                                                ' Ist szSize leer ?
                szSizeArray = Split(szSize, szDelimiter)                        ' Value aufspliten mit szDelimiter
                If szSizeArray.Length < 1 Then Exit Sub ' Array zuklein -> fertig
                If szSizeArray(0) <> "" And szSizeArray(1) <> "" Then           ' Werte vorhanden ?
                    If IsNumeric(szSizeArray(0)) _
                            And IsNumeric(szSizeArray(1)) Then                  ' Werte sind zahlen
                        F.Width = CLng(szSizeArray(0))                          ' (0) = Width
                        F.Height = CLng(szSizeArray(1))                         ' (1) = Height
                    End If
                End If
            End If
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "SetWindowSizeFromString", ex) ' Fehlermeldung
            End If
        End Try
    End Sub

    Public Sub SetWindowStateFromString(ByVal F As Form, _
                                        ByVal szWinState As String, _
                                        Optional ByRef oBag As clsObjectBag = Nothing)
        ' Setzt f.WindowState aus szWinState prüft und wandelt sting um
        Dim lngState As Integer                                                 ' Integer von szWinState
        Try                                                                     ' Fehlerbehandlung aktivieren
            If F Is Nothing Then Exit Sub ' Kein Form dann fertig
            If szWinState <> "" Then                                            ' Ist szWinState leer ?
                If IsNumeric(szWinState) Then                                   ' Ist szWinState Zahl ?
                    lngState = CLng(szWinState)                                 ' In Intwandeln
                    If lngState >= 0 And lngState <= 2 Then                     ' 0=normal 1=min 2=max
                        F.WindowState = lngState                                ' Window State setzen
                    Else
                        F.WindowState = vbNormal                                ' <0 bzw. >2 dann Normal
                    End If
                Else
                    F.WindowState = vbNormal                                    ' keine Zahl dann Normal
                End If
            Else
                F.WindowState = vbNormal                                        ' Nix angegeben dann Normal
            End If
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "SetWindowStateFromString", ex) ' Fehlermeldung
            End If
        End Try
    End Sub

#Region "Combobox Funktionen"

    Public Sub FillCMBlist(ByVal cmbCTL As ComboBox, _
                                 ByVal szValueList As String, _
                                 Optional ByRef oBag As clsObjectBag = Nothing)
        Dim szListArray() As String                                             ' Array mit werten
        Dim i As Integer                                                        ' Counter (Array)
        Try                                                                     ' Fehlerbehandlung aktivieren
            If cmbCTL Is Nothing Then Return ' Keine ComboBox -> Fertig
            cmbCTL.Items.Clear()                                                ' Alte Werte raus
            If szValueList = "" Then Return ' Keine Werteliste -> Fertig
            szListArray = Split(szValueList, ";")                               ' Liste in Array aufsplaten
            If Not IsNothing(szListArray) Then                                  ' Array nicht leer
                For i = 0 To szListArray.Length - 1                             ' Alle Array Items durchlaufen
                    If szListArray(i) <> "" Then cmbCTL.Items.Add(szListArray(i)) ' Wert hinzufügen
                Next i                                                          ' Nächstes Array item
            End If
        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "FillCMBlist", ex)            ' Fehlermeldung
            End If
        End Try
    End Sub

    Public Sub FillCMBlistFromSQL(ByVal cmbCTL As ComboBox, _
                                 ByVal szSQL As String, _
                                 Optional ByRef oBag As clsObjectBag = Nothing)
        Dim szValueList As String = ""
        Dim i As Integer                                                        ' Counter (Array)
        Dim DS As DataSet
        Try                                                                     ' Fehlerbehandlung aktivieren
            If cmbCTL Is Nothing Then Return ' Keine ComboBox -> Fertig
            If szSQL = "" Then Return ' Keine Werteliste -> Fertig
            DS = oBag.ObjDBConnect.FillDS(szSQL, "", True)                      ' dataset füllen
            If IsNothing(DS) Then Return ' Kein DataSet -> Fertig
            If DS.Tables.Count = 0 Then Return
            If DS.Tables(0).Rows.Count = 0 Then Return
            For i = 0 To DS.Tables(0).Columns.Count - 1                         ' Alle Spalten durchlaufen
                szValueList = szValueList & DS.Tables(0).Rows(0).Item(i).ToString & ";"
            Next                                                                ' Nächste Spalte
            If Right(szValueList, 1) = ";" Then szValueList = Left(szValueList, Len(szValueList) - 1)
            Call FillCMBlist(cmbCTL, szValueList, oBag)

        Catch ex As Exception                                                   ' Fehler behandeln
            If Not IsNothing(oBag) Then                                         ' Nur wenn obag angegeben
                Call oBag.ErrorHandler(MODULNAME, "FillCMBlist", ex)            ' Fehlermeldung
            End If
        End Try
    End Sub

#End Region
End Module
