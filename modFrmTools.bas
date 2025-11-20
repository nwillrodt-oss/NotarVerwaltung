Attribute VB_Name = "modFrmTools"
Option Explicit                                                     ' Variaben Deklaration erzwingen
Private Const MODULNAME = "modFrmTools"                             ' Modulname für Fehlerbehandlung

                                                                    ' *****************************************
                                                                    ' API gefrotzel
                                                                    ' API für Optimale Coulum Header {
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
                                                                    ' API für Optimale Coulum Header }
                                                                    ' API für Top Most {
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Const HWND_TOPMOST As Long = -1&
Public Const HWND_NOTOPMOST As Long = -2&
                                                                    ' API für Top Most }
                                                                    ' API für TreeView Backcolor {
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, _
    ByVal lHPalette As Long, _
    ByRef lColorRef As Long) As Long
Private Const TVM_SETBKCOLOR = 4381&
                                                                    ' API für TreeView Backcolor }
'                                                                    ' API Fur Transparente Fenster {
'Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, _
'    ByVal nIndex As Long) As Long
'Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, _
'    ByVal nIndex As Long, _
'    ByVal dwNewLong As Long) As Long
'Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, _
'    ByVal crKey As Long, _
'    ByVal bAlpha As Byte, _
'    ByVal dwFlags As Long) As Long
'Public Const GWL_EXSTYLE = (-20)
'Public Const WS_EX_LAYERED = &H80000
''Public Const LWA_COLORKEY = &H1                                     ' Macht nur eine Farbe transparent
'Public Const LWA_ALPHA = &H2                                        ' Macht das ganze Fenster transparent
'                                                                    ' API Fur Transparente Fenster }
                                                                    
                                                                    ' *****************************************
                                                                    ' Controls Allgemein
Public Sub RemoveTabByCaption(ctlTabStrip As TabStrip, szTabCaption As String)
' Enthern Registerkarte aus Tabstrip mit angegebener Caption
Dim i As Integer                                                    ' Counter
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If ctlTabStrip Is Nothing Then GoTo exithandler                 ' Kein TabStrip Control -> Fertig
    If szTabCaption = "" Then GoTo exithandler                      ' Keine Tab Captin -> fertig
    For i = 1 To ctlTabStrip.Tabs.Count                             ' Alle Tabs (registerkarten) durchlaufen
        If UCase(szTabCaption) = UCase(ctlTabStrip.Tabs(i)) Then    ' Caption gefunden ?
            If ctlTabStrip.SelectedItem.Index = i Then              ' ist das der Aktive Tab ?
                If ctlTabStrip.SelectedItem.Index = 1 Then          ' Ist das der 1 TAb
                    ctlTabStrip.Tabs(2).Selected = True             ' 2. Auswählen
                Else
                    ctlTabStrip.Tabs(1).Selected = True             ' 1. Auswählen
                End If
            End If
            ctlTabStrip.Tabs.Remove (i)                             ' Tab entfernen
            Exit For                                                ' Schleife abbrechen
        End If
    Next i                                                          ' Nächsten Tab
exithandler:

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "RemoveTabByCaption", errNr, errDesc)
    Resume exithandler
End Sub

Public Function GetControlByName(ThisForm As Form, szContolname As String) As Control
' Liefert Referenz auf das Contol mit Namen = szContolname

    Dim CTL As Control                                              ' Akt. Control
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    
On Error GoTo Errorhandler
    
    For Each CTL In ThisForm                                        ' Alle Controls des Forms Duchlaufen
        szDetails = "Controlname: " & CTL.Name
        If UCase(CTL.Name) = UCase(szContolname) Then               ' Namen Prüfen
            Set GetControlByName = CTL                              ' CTL Referenz zurück
            GoTo exithandler                                        ' Fertig
        End If
    Next                                                            ' Nächstes  Control
    
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetControlByName", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Public Function GetControlByDatafield(ThisForm As Form, szDatafield As String) As Control
' Liefert Referenz auf das Contol mit Datafield = szDatafield

    Dim CTL As Control                                              ' Akt. Control
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    Dim szCtlPref As String                                         ' Prefix des Controlnamens
    
On Error GoTo Errorhandler
    
    For Each CTL In ThisForm                                        ' Alle Controls des Forms Duchlaufen
        szDetails = "Controlname: " & CTL.Name
        szCtlPref = UCase(Left(CTL.Name, 3))                        ' Prfix ermitteln
        If szCtlPref = "TXT" Or szCtlPref = "CMB" Then
            If UCase(CTL.DataField) = UCase(szDatafield) Then       ' Namen Prüfen
                Set GetControlByDatafield = CTL                     ' CTL Referenz zurück
                GoTo exithandler                                    ' Fertig
            End If
        End If
    Next                                                            ' Nächstes  Control
    
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetControlByDatafield", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Public Sub HoverLabel(lblCTL As Label, bHover As Boolean)
' Setzt Label.Caption unterstrichen bzw. nicht unterstrichen
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If lblCTL Is Nothing Then Exit Sub                              ' Kein Label -> Fertig
    If bHover Then                                                  ' Unterstrichen
        lblCTL.Font.Underline = True
        Call MousePointerLink(, lblCTL)                             ' Mouspointer sezen
        'lblCtl.ForeColor = 16576
    Else                                                            ' Nicht Unterstrichen
        lblCTL.Font.Underline = False
        'lblCtl.ForeColor = -2147483646
    End If
    Err.Clear                                                       ' Evtl. Error Cleraen
End Sub

Public Function FillCmbListWithSQL(cmbctl As ComboBox, szSQL As String, dbCon As Object) As ADODB.Recordset
' Liste einer Combobox mit werten aus SQL statement füllen
' Es wird nur die Erste Spalte berücksichtigt

    Dim rsList As ADODB.Recordset                                   ' RS mit Listen werten
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
        
    If szSQL = "" Then GoTo exithandler                             ' Kein SQL - > raus
    cmbctl.Clear                                                    ' Evtl Liste löschen
    szDetails = "Combobox: " & cmbctl.Name & vbCrLf & "SQL: " & szSQL
    Set rsList = dbCon.fillrs(szSQL, False)
    If Not rsList Is Nothing Then                                   ' Kein RS (Fehler) -> raus
        If rsList.RecordCount > 0 Then                              ' Wenn DS vorhanden
            rsList.MoveFirst                                        ' zum ersten springen
            While Not rsList.EOF                                    ' Alle DS durchlaufen
'                If rsList.Fields.Count > 1 Then
'                    cmbctl.AddItem Trim(objTools.checknull(rsList.Fields(0).Value, "")), objTools.checknull(rsList.Fields(1).Value, 0)
'                Else
                    cmbctl.AddItem Trim(objTools.checknull(rsList.Fields(0).Value, ""))
'                End If
                rsList.MoveNext                                     ' Nächster DS
            Wend
        End If
        Set FillCmbListWithSQL = rsList                             ' RS als Rückgabe Wert
    End If
    
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "FillCmbListWithSQL", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Public Sub FillFrameWithLV(CurFrame As Frame, CurLV As ListView)
    Call CurLV.Move(120, 240, CurFrame.Width - 240, CurFrame.Height - 360)  ' ListView auf ges Frame ausdehnen
End Sub
                                                                    ' *****************************************
                                                                    ' TreeView Funktionen
Public Sub TreeViewBackColor(tvw As TreeView, _
        Optional ByVal BackColor As Variant = vbWindowBackground, _
        Optional ByVal NodesBackColor As Boolean = True)
        
    Dim nNode As node
    Dim nIL As ImageList
    Dim nBackColor As Long
    Dim nStyle As Long

On Error Resume Next

    With tvw
        Set nIL = tvw.ImageList
        If Not (nIL Is Nothing) Then
            nIL.BackColor = BackColor
        End If
        nStyle = .Style
        OleTranslateColor BackColor, 0&, nBackColor
        SendMessage .hWnd, TVM_SETBKCOLOR, 0, ByVal nBackColor
        .Style = 0
        .Style = nStyle
        If NodesBackColor Then
            For Each nNode In .Nodes
               nNode.BackColor = BackColor
            Next nNode
        End If
    End With
End Sub

Public Sub AddTreeNode_New(TV As TreeView, ParentKey As String, _
        Key As String, Tag As String, Text As String, _
        dbCon As Object, Optional ImageIndex As Integer, Optional bWithoutSubnodes As Boolean)

    Dim newNode As node                                             ' Neuer Konten
    Dim szSubNodeSQL As String                                      ' SQl für Sub nodes
    Dim lngSubNodeImage As Integer                                  ' Image index für subnode
    Dim szSubNodeImage As String                                    ' Image index für subnode als String aus ini
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    If ParentKey = "" Then                                          ' Rootnode anlegen
        Set newNode = TV.Nodes.Add(, , Key, Text)
        newNode.Tag = Tag
    Else                                                            ' Sonst in den Baum einhängen
        Set newNode = TV.Nodes.Add(ParentKey, tvwChild, ParentKey & "\" & Key, Text)
        newNode.Tag = newNode.Parent.Tag & TV_KEY_SEP & Tag         ' Tag setzen
    End If
    
    
    newNode.Image = ImageIndex                                      ' Image setzen
    
    If bWithoutSubnodes Then GoTo exithandler                       ' keine Subnodes mitanlegen -> fertig
    
    Call AddSubTreeNodes(TV, newNode, dbCon, ImageIndex, bWithoutSubnodes)  'Auf subnodes prüfen (recursiv)
    
exithandler:
On Error Resume Next
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "AddTreeNode_New", errNr, errDesc)
    Resume exithandler
End Sub

Public Function DelSubTreeNodes(TV As TreeView, ParentNode As node)
' Löscht alle Unterknoten des Parentnode

    Dim delNode As node                                             ' zu löschender Node
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    If ParentNode Is Nothing Then GoTo exithandler
    
    While ParentNode.Children > 0                                   ' Alle Child nodes durchlaufen
        Set delNode = ParentNode.Child                              ' zu löschenden Node festlegen
        If Not delNode Is Nothing Then TV.Nodes.Remove (delNode.Index)
    Wend
    
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "DelSubTreeNodes", errNr, errDesc)
End Function

'Public Sub AddSubTreeNodes(TV As TreeView, ParentNode As node, _
'            DBCon As Object, Optional ImageIndex As Integer, Optional bWithoutSubnodes As Boolean)
'' Legt Tree Nodes mit informationen aus XML file an
'
'    Dim szDetails As String                                         ' Detailinfos für Fehlerbehandlung
'    Dim szSubNodeSQL As String                                      ' SQl für Sub nodes
'    Dim lngSubNodeImage As Integer                                  ' Image index für subnode
'    Dim szSubNodeImage As String                                    ' Image index für subnode als String aus ini
'    Dim szWhere As String                                           ' Evtl. Where statement
'    Dim ID As String                                                ' Letzter part des NodeKey as ID
'    Dim szSubNodeList As String                                     ' Liste von subnodes
'    Dim SubNodeArray() As String                                    ' Array mit subnodes
'    Dim TVNode As TreeViewNodeInfo                                  ' tree nodes infos
'    Dim i As Integer                                                ' Counter
'
'On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
'
'    szSubNodeList = objTools.GetSubNodeListFromXML(App.Path & "\" _
'            & INI_XMLFILE, ParentNode.Tag)                          ' Subnodes List zu akt Node ermitteln
'    SubNodeArray = Split(szSubNodeList, ";")                        ' Subnodelist in array Aufspalten
'    For i = 0 To UBound(SubNodeArray)                               ' SubnodeListArray durchlaufen
'        With TVNode                                                 ' SubTreeNode Infos einlesen
'            Call objTools.GetTVNodeInfofromXML(App.Path & "\" & INI_XMLFILE, _
'                        ParentNode.Tag & TV_KEY_SEP & SubNodeArray(i), _
'                        .szTag, .szText, .szKey, .bShowSubnodes, .szSQL, .szWhere, .lngImage)
'            If .lngImage <> 0 Then ImageIndex = .lngImage           ' Image wert konvertieren
'            If .szSQL <> "" Then                                    ' SQL Statement vorhanden
'                If .szWhere <> "" Then                              ' Where Statement vorhanden
'                    ID = GetLastKey(ParentNode.Key, "\")            ' ID ermitteln
'                    .szSQL = objSQLTools.AddWhereInFullSQL(.szSQL, .szWhere & "'" & _
'                            ID & "'")                               ' ID & Where Statement integrieren
'                End If
'                Call AddTreeNodeFromSQL(TV, ParentNode, .szSQL, DBCon, ImageIndex, _
'                        Not .bShowSubnodes)                         ' Tree node aus SQL hinzufühgen
'            ElseIf .szTag <> "" And .szText <> "" And .szKey <> "" Then
'                Call AddTreeNode_New(TV, ParentNode.Key, .szKey, .szTag, .szText, _
'                        DBCon, ImageIndex, Not .bShowSubnodes)      ' statischen Tree node hinzufügen
'
'            End If
'        End With
'     Next i
'
'exithandler:
'
'Exit Sub
'Errorhandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.Errorhandler(MODULNAME, "AddSubTreeNodes", errNr, errDesc)
'    Resume exithandler
'End Sub

Public Sub AddSubTreeNodes(TV As TreeView, ParentNode As node, _
            dbCon As Object, Optional ImageIndex As Integer, Optional bWithoutSubnodes As Boolean)

    Dim ID As String                                                ' Letzter part des NodeKey as ID
    Dim i As Integer                                                ' Counter
    Dim szSubNodeList As String                                     ' Liste der Subnodes
    Dim SubNodeArray() As String                                    ' Array der Subnodes
    Dim TVNode As TreeViewNodeInfo                                  ' Node informationen
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    TV.MousePointer = vbHourglass                                   ' Sanduhr anzeigen

    szSubNodeList = objTools.GetSubNodeListFromXML(App.Path & "\" _
            & INI_XMLFILE, ParentNode.Tag)                          ' Liste der Subnodes aus XMl laden
    SubNodeArray = Split(szSubNodeList, ";")                        ' Sub nodelist in Array aufspalten
    For i = 0 To UBound(SubNodeArray)                               ' Für alle subnodes
        DoEvents                                                    ' andere aktionen zulassen
        With TVNode
            Call objTools.GetTVNodeInfofromXML(App.Path & "\" & INI_XMLFILE, _
                        ParentNode.Tag & TV_KEY_SEP & SubNodeArray(i), _
                        .szTag, .szText, .szKey, .bShowSubnodes, _
                        .szSQL, .szWhere, .lngImage)                ' Tree node informationen aus XML laden
            'If .lngImage <> "" Then ImageIndex = CLng(.lngImage)
            If .szSQL <> "" Then                                    ' Wenn ein SQL statementvorliegt
                If .szWhere <> "" Then                              ' Wenn Where Part vorliegt
                    ID = GetLastKey(ParentNode.Key, "\")            ' ID ermitteln
                    .szSQL = objSQLTools.AddWhereInFullSQL(.szSQL, _
                                .szWhere & "'" & ID & "'")          ' Where Part in SQL einbauen
                End If
                Call AddTreeNodeFromSQL(TV, ParentNode, .szSQL, dbCon, _
                        .lngImage, Not .bShowSubnodes)             ' tree node aus SQL hinzufügen
            ElseIf .szTag <> "" And .szText <> "" And .szKey <> "" Then
                Call AddTreeNode_New(TV, ParentNode.Key, .szKey, .szTag, .szText, _
                        dbCon, .lngImage, Not .bShowSubnodes)      ' Tree node normal hinzufügen
            End If
        End With
     Next i                                                         ' Nächster Subnode
     
exithandler:
On Error Resume Next
    TV.MousePointer = vbDefault                                     ' Mauszeiger wieder normal
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "AddSubTreeNodes", errNr, errDesc)
    Resume exithandler
End Sub

Public Sub AddTreeNodeFromSQL(TV As TreeView, ParentNode As node, szNodeSQL As String, _
        dbCon As Object, Optional ImageIndex As Integer, Optional bWithoutSubnodes As Boolean)
' Left Tree nodes aus SQL statement an

    Dim rsNode As ADODB.Recordset                                   ' RS mit Node Daten
    Dim szKey As String                                             ' Neuer Node Key
    Dim szTag As String                                             ' Neuer Node Tag
    Dim szText As String                                            ' Neuer Node Text
    Dim lngSubNodeImage As Integer                                  ' Image index für subnode
    Dim szSubNodeImage As String                                    ' Image index für subnode als String aus ini
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    If szNodeSQL = "" Then GoTo exithandler                         ' Kein SQL -> Raus
    Set rsNode = objDBconn.fillrs(szNodeSQL, False)                 ' RS füllen
    If Not rsNode Is Nothing Then
        If Not rsNode.EOF Then rsNode.MoveFirst                     ' zum ersten DS
        While Not rsNode.EOF                                        ' RS durchlaufen
            szText = rsNode.Fields("nodetext").Value                ' NodeText aus RS
            szKey = rsNode.Fields("nodekey").Value                  ' NodeKey as RS
            szTag = rsNode.Fields("nodetag").Value                  ' NodeTag aus RS
            If szText <> "" Or szKey <> "" Or szTag <> "" Then
                lngSubNodeImage = ImageIndex
                If szSubNodeImage <> "" And IsNumeric(szSubNodeImage) Then _
                        lngSubNodeImage = CLng(szSubNodeImage)      ' SubnodeImage setzen
                Call AddTreeNode_New(TV, ParentNode.Key, szKey, szTag, szText, _
                        dbCon, lngSubNodeImage, bWithoutSubnodes)   ' Sub Node hinzufügen
            End If
            rsNode.MoveNext                                         ' Nächster DS
        Wend
    End If
   
exithandler:
On Error Resume Next
     Set rsNode = Nothing
     
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "AddTreeNodeFromSQL", errNr, errDesc)
    Resume exithandler
End Sub

Public Sub SelectTreeNode(TV As TreeView, SelectedNode As node, Optional bExpand As Boolean)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    TV.SelectedItem = SelectedNode                                  ' SelectedNode auf ausgewählt setzen
    If bExpand Then TV.SelectedItem.Expanded = True                 ' Expanden
    Err.Clear                                                       ' Evtl Error Clearen
End Sub

Public Function GetSelectTreeNode(TV As TreeView) As node
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Set GetSelectTreeNode = TV.SelectedItem                         ' Ermittelt aktuell ausgewählten Tree node
    Err.Clear                                                       ' Evtl Error Clearen
End Function

Public Function GetIDFromNode(Optional TV As TreeView, Optional cNode As node)
' Ermitteli DS ID aus cNode.Tag & Key
    Dim fMain As Form                                               ' Hauptformular (mit TV & LV)
    Dim szKeyArray() As String                                      ' Node Key in array aufgespalten
    Dim szTagArray() As String                                      ' Node Tag in array aufgespalten
    Dim szTmp As String                                             ' Hilfsvariable
    Dim ID As String                                                ' Evtl ID des Detaildatensatzes
    Dim i As Integer                                                ' Counter
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung altivieren
    
    If TV Is Nothing Then                                           ' Kein TV angegeben
        Set fMain = objObjectBag.getMainForm                        ' Hautform aus OBag holen
        Set TV = fMain.GetTV                                        ' TV Prop. Abfragen
    End If
    
    If cNode Is Nothing Then                                        ' kein Node angegeben
        Set cNode = GetSelectTreeNode(TV)                           ' Akt. Node ermitteln
    End If
    If cNode Is Nothing Then GoTo exithandler                       ' immer noch kein node dann raus
    
    szTagArray = Split(cNode.Tag, "\")                              ' Node Tag aufspalten
    szKeyArray = Split(cNode.Key, "\")                              ' Node Key aufspalten
    
    If szTagArray(UBound(szTagArray)) = "*" Then                    ' detaildatensatz
        ID = GetLastKey(cNode.Key, TV_KEY_SEP)                      ' ID aus Key ermitten
    Else                                                            ' Wenn kein Detail Datensatz
        If InStr(cNode.Tag, "*") Then                               ' Statischer unterknoten eines Detaildatensatzes
            szTmp = ""                                              ' Tmp Leeren
            i = UBound(szTagArray) + 1                              ' Max Arra Index festlegen
            While szTmp <> "*"                                      ' Tag bis * rückwärts durchlaufen
                i = i - 1                                           ' arrayindex herunterzählen
                szTmp = szTagArray(i)                               ' array wert merken
                ID = szKeyArray(i)                                  ' ID am gleicher stelle aus Tag
            Wend
        End If
        If ID = "" And InStr(cNode.Key, TV_KEY_SEP) > 0 Then ID = GetLastKey(cNode.Key, TV_KEY_SEP)     ' sonst mit gewalt
    End If
    
exithandler:
    GetIDFromNode = ID                                              ' ID Zurück geben
     
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetIDFromNode", errNr, errDesc)
    Resume exithandler
End Function

'Public Function GetRootNodeParent(node As MSComctlLib.node, Optional szRootkey As String) As node
'
'    Dim ParentNode As node
'    Dim tmpNode As node
'
'On Error GoTo Errorhandler
'
'    Set ParentNode = node
'    While Not (ParentNode.Parent Is Nothing)
'        If ParentNode.Parent.Key = szRootkey Then
'            Set tmpNode = ParentNode
'        End If
'        Set ParentNode = ParentNode.Parent
'    Wend
'    If szRootkey = "" Then
'        Set GetRootNodeParent = ParentNode
'    Else
'        If node.Key = szRootkey Then
'            Set GetRootNodeParent = node
'        Else
'            Set GetRootNodeParent = tmpNode
'        End If
'    End If
'
'exithandler:
'On Error Resume Next
'
'Exit Function
'Errorhandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.Errorhandler(MODULNAME, "GetRootNodeParent", errNr, errDesc)
'End Function

Public Function GetNodeByKey(TV As TreeView, szKey As String) As node
' ermittelt Node aus angegebenen Node Key

    Dim x As Integer                                                ' Counter
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    For x = 1 To TV.Nodes.Count                                     ' Alle Nodes duchlaufen
        If TV.Nodes(x).Key = szKey Then
            Set GetNodeByKey = TV.Nodes(x)                          ' Node gefunden
            Exit For                                                ' fertig
        End If
    Next x
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetNodeByKey", errNr, errDesc)
End Function

Public Function GetNodeByFullPath(TV As TreeView, szFullPath As String) As node
' ermittelt Node aus angegebenen Node.FullPath
    
    Dim x As Integer                                                ' Counter
    Dim cNode As node                                               ' gefundener Node
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    For x = 1 To TV.Nodes.Count                                     ' Alle Nodes duchlaufen
        If TV.Nodes(x).FullPath = szFullPath Then                   ' Fullpath vergleichen
            Set GetNodeByFullPath = TV.Nodes(x)                     ' Node gefunden
            Exit For                                                ' fertig
        End If
    Next x
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetNodeByFullPath", errNr, errDesc)
End Function

                                                                    ' *****************************************
                                                                    ' ListView Funktionen
Public Sub CangeView(ctlListView As ListView, lngView As Integer)
' Ändert LV ansicht
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    ctlListView.View = lngView                                      ' List View ansicht ändern
    ctlListView.Refresh                                             ' List View Aktualisieren
    Err.Clear                                                       ' Evtl Error Clearen
End Sub

Public Function GetSelectListItem(LV As ListView) As ListItem
' ermittel akt. ListView Item
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Set GetSelectListItem = LV.SelectedItem                         ' Ermittelt aktuell ausgewähltes List Item
    Err.Clear                                                       ' Evtl Error Clearen
End Function

Public Sub InitDefaultListViewResult(LV As ListView)
' Default einstellungen für ListView
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    LV.View = lvwReport                                             ' Ansicht Report (Details)
    LV.FullRowSelect = True                                         ' Ganze Zeile Selecten
    LV.GridLines = False                                            ' Keine Gitter Linien
    LV.LabelEdit = lvwManual                                        ' Label bearbeiten nicht automatisch
    LV.Refresh                                                      ' einmal aktualisieren
    Err.Clear                                                       ' Evtl Error Clearen
End Sub

Public Sub ClearLV(LV As ListView)
' Listitems und columheaders löschen (für neue ansicht)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    LV.Sorted = False                                               ' Sortierung abschalten
    LV.ListItems.Clear                                              ' Evtl. vorhandene Items löschen
    LV.ColumnHeaders.Clear                                          ' Evtl. vorhandene Spaltenköpfe löschen
    Err.Clear                                                       ' Evtl Error Clearen
End Sub

Public Sub SetColumnWidth(ctlListView As ListView, ColIndex As Integer, Width As Integer)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    ctlListView.ColumnHeaders(ColIndex).Width = Width               ' Setzt Spalten breite
    Err.Clear                                                       ' Evtl Error Clearen
End Sub

Public Sub HideColumn(ctlListView As ListView, ColIndex As Integer)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
     ctlListView.ColumnHeaders(ColIndex).Width = 0                  ' Splate ausblenden
     Err.Clear                                                      ' Evtl Error Clearen
End Sub

Public Sub SelectLVItem(LV As ListView, ItemKey As String)
' Selectiert ListView Item mit Key = ItemKey
    
    Dim i As Integer                                                ' Counter
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
       
    For i = 1 To LV.ListItems.Count                                 ' Für jedes List view Item
        If LV.ListItems(i).Key = ItemKey Then                       ' Key = ItemKey
            LV.ListItems(i).Selected = True                         ' Wenn JA selecten
            LV.ListItems(i).EnsureVisible                           ' Stellt sicher das Item auch sichtbar ist (Scroll)
            LV.Refresh                                              ' LV aktialiseren
            Exit For                                                ' Fertig
        End If
    Next i
    
exithandler:
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SelectLVItem", errNr, errDesc)
End Sub

Public Sub AddListViewSubItem(ByVal itemMain As ListItem, szSubItemText As String)
' Legt ein Sub Item an (nur in ListView mit ansicht Report sichtbar)
    Dim Index As Integer                                            ' Neuer SubItem index

On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    Index = itemMain.ListSubItems.Count + 1                         ' Neuen Index Ermitteln
    itemMain.SubItems(Index) = Left(szSubItemText, 50)              ' SubItem anlegen
    
exithandler:
On Error Resume Next
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "AddListViewSubItem", errNr, errDesc)
    Resume exithandler
End Sub

Public Function AddListViewItem(ctlListView As ListView, _
        szItemText As String, _
        Optional szSubItemText As String, _
        Optional szValueName As String, _
        Optional intImage As Integer) As ListItem
' Legt ein ListView Item an und evt. ein Sub Item
' szValuename wird der Item Tag

    Dim itemX As ListItem                                            'angelegtes Listview item
     
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
   
    Set itemX = ctlListView.ListItems.Add(, , szItemText, intImage) ' List Item anlegen
    itemX.SmallIcon = intImage                                      ' Item Image setzten
    itemX.Tag = szValueName                                         ' Tag setzen
    If szSubItemText <> "" Then Call AddListViewSubItem(itemX, szSubItemText) ' Evtl. sub item setzen
    Set AddListViewItem = itemX                                     ' Item als RÜckgabe wert
    
exithandler:
On Error Resume Next
    Set itemX = Nothing
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "AddListViewItem", errNr, errDesc)
    Resume exithandler
End Function

Public Function FillLVBySQL(LV As ListView, _
        szSQL As String, _
        dbCon As Object, _
        Optional bShowValueList As Boolean, _
        Optional lngImageIndex As Integer, _
        Optional szIndexField As String, _
        Optional bOptColumnWidth As Boolean, _
        Optional bNoColWidth As Boolean, _
        Optional szKey As String, _
        Optional lngAltImageIndex As Integer, _
        Optional AltImageConditionField As String, _
        Optional AltImageConditionValue As String) As ADODB.Recordset
' Füllt ein Listview mit Items aus SQl Statement
' bShowValueList = True gibt an ob ein DS untereinader angezeit wird Pro Feld ein Item)
' bShowValueList = False Listet alle DS auf
' bOptColumnWidth = true die Splatenbreite wird optimal eingestellt
'   Sonst wird versicht die Splatenbreite aus der Reg zu laden

    Dim rsList As New ADODB.Recordset                               ' RS mit Daten
    Dim szDetails As String                                         ' Detailinfos für Fehlerbehandlung
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    Call ClearLV(LV)                                                ' ListItems und Columns löschen
    If IsEmpty(lngImageIndex) Then lngImageIndex = 0                ' Default wert für image index
    If szSQL = "" Then GoTo exithandler                             ' Kein SQL -> raus
    szDetails = "SQL: " & szSQL
    Set rsList = dbCon.fillrs(szSQL, True)                          ' Daten holen
    'If rsList Is Nothing Then GoTo exithandler                     ' Keine Daten (fehler) -> Fertig
    Call FillLVByRS(LV, "", rsList, bShowValueList, lngImageIndex, szIndexField, _
            bOptColumnWidth, bNoColWidth, szKey, lngAltImageIndex, _
            AltImageConditionField, AltImageConditionValue)         ' ListView aus Recordet füllen
    Set FillLVBySQL = rsList                                        ' RS mit daten als Rückgabe wert
    
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "FillLVBySQL", errNr, errDesc, szDetails)
    Resume exithandler
End Function

'Public Function FillLVBySQL(LV As ListView, _
'        szSQL As String, _
'        DBCon As Object, _
'        Optional bShowValueList As Boolean, _
'        Optional lngImageIndex As Integer, _
'        Optional szIndexField As String, _
'        Optional bOptColumnWidth As Boolean, _
'        Optional bNoColWidth As Boolean, _
'        Optional bNew As Boolean) As ADODB.Recordset
'' Füllt ein Listview mit Items aus SQl Statement
'' bShowValueList = True gibt an ob ein DS untereinader angezeit wird Pro Feld ein Item)
'' bShowValueList = False Listet alle DS auf
'' bOptColumnWidth = true die Splatenbreite wird optimal eingestellt
''   Sonst wird versicht die Splatenbreite aus der Reg zu laden
'' bNew = True dann werden nur ColumHeader geladen keine Daten
'
'    Dim rsList As New ADODB.Recordset                               ' RS mit Daten
'    Dim LVItem As ListItem                                          ' ListView item
'    Dim i As Integer                                                ' Counter
'    Dim ci As Integer                                               ' Akt ColumnIndex
'    Dim szTmpName As String                                         ' Item text
'    Dim szTmpID As String                                           ' Evtl DS ID
'    Dim szTagArray() As String                                      ' LV.Tagin array aufgespalten
'    Dim NewTag As String                                            ' Evtl. Neu generierte Tag
'    Dim szDetails As String
'
'On Error GoTo Errorhandler
'
'    Call ClearLV(LV)                                                ' ListItems und Columns löschen
'
'    If IsEmpty(lngImageIndex) Then lngImageIndex = 0
'
'    If szSQL = "" Then GoTo exithandler                             ' Kein SQL -> raus
'    Set rsList = DBCon.fillrs(szSQL, True)                          ' Daten holen
'
'    'If rsList Is Nothing Then GoTo exithandler                     ' Keine Daten (fehler) -> Fertig (auskomm das sonst keien spalten köpfe)
'
'    If bShowValueList Then                                          ' Alle felder eines DS untereinader auflisten
'                                                                    ' ColumsHeader anlegen
'        Call AddLVColumn(LV, "Eigenschaft", 2000)                   ' Feste Spalte Eigenschaft (enthält Feldnamen)
'        Call AddLVColumn(LV, "Wert", 4000)                          ' Feste Spalte Wert (enthält Feldwert)
'        If rsList Is Nothing Then GoTo exithandler
'        If rsList.RecordCount = 0 Then GoTo exithandler             ' Keine Daten -> Fertig
'        If bNew Then GoTo exithandler                               ' Neuer DS keine Daten anzeigen
'        rsList.MoveFirst                                            ' Nur ersten DS anzeigen
'        szTmpID = Trim(objTools.checknull(rsList.Fields(0).Value, ""))
'        ' Das Letzte * im LV.Tag durch ID ersetzen
'        szTagArray = Split(LV.Tag, TV_KEY_SEP)                      ' LV.Tag in array aufspalten
'        For i = 0 To UBound(szTagArray) - 1                         ' Bis auf letzten eintrag wieder zusammen setzen
'            NewTag = NewTag & szTagArray(i) & TV_KEY_SEP
'        Next
'        NewTag = NewTag & szTmpID                                   ' DS ID an NeuenTag anhängen
'        For i = 0 To rsList.Fields.Count - 1                        ' Für jedes Feld eine Zeile
'            If Left(rsList.Fields(i).Name, 2) <> "ID" Then          ' Wenn Feld mit ID anfängt ausblenden
'                szTmpName = Trim(objTools.checknull(rsList.Fields(i).Value, ""))
'                Set LVItem = AddListViewItem(LV, rsList.Fields(i).Name, szTmpName, NewTag, lngImageIndex)
'                LVItem.Icon = lngImageIndex                         ' Icon Setzen
'                szTmpName = ""
'            End If
'        Next i
'    Else                                                            ' Alle DS untereinader auflisten
'                                                                    ' ColumsHeader anlegen
'        For i = 1 To rsList.Fields.Count - 1                        ' Erste Spalte ist ID -> auslassen
'            ci = AddLVColumn(LV, rsList.Fields(i).Name)             ' Colum hinzufügen
'            Select Case rsList.Fields(i).Type                       ' Daten Typ Prüfen
'            Case adSmallInt, adInteger, adBigInt, adSingle          ' Zahlen
'                LV.ColumnHeaders(ci).Tag = adInteger                ' DatenFormat in ColumnKey zum sortieren
'            Case adDate, adDBDate, adDBTimeStamp, adDBTime          ' Datum
'                LV.ColumnHeaders(ci).Tag = adDate                   ' DatenFormat in ColumnKey zum sortieren
'            Case Else
'            End Select
'        Next i
'
'        If bNew Then GoTo exithandler                               ' Neuer DS keine Daten anzeigen
'
'      ' Alle DS anzeigen
'      ' Jedes Listitem bekommet als Tag den Tag des LVs + Datensat ID eingetragen
'        Do While Not rsList.EOF                                     ' Für jeden DS einen eintrag
'            If szIndexField <> "" Then                              ' explicit angegebenes ID feld
'                szTmpID = Trim(objTools.checknull(rsList.Fields(szIndexField).Value, ""))
'            End If
'            If szTmpID = "" Then szTmpID = Trim(objTools.checknull(rsList.Fields(0).Value, "")) ' Sonst 1. Feldwert ID
'            szTmpName = Trim(objTools.checknull(rsList.Fields(1).Value, ""))    ' 2. Feldwert ItemText
'            If szTmpName <> "" And szTmpID <> "" Then               ' Name und ID <>""
'                szDetails = "ItemKey: " & LV.Tag & TV_KEY_SEP & szTmpID & vbCrLf & "ItemTag: " & LV.Tag & TV_KEY_SEP & "*" & vbCrLf
'
'                'Set LVItem = AddListViewItem(LV, szTmpName, , LV.Tag & TV_KEY_SEP & szTmpID, lngImageIndex)
'                'LVItem.Key = LV.Tag & TV_KEY_SEP & "*"
'                Set LVItem = AddListViewItem(LV, szTmpName, , LV.Tag & TV_KEY_SEP & szTmpID, lngImageIndex)
'                szDetails = szDetails & "Item angelegt."
'                LVItem.Tag = LV.Tag & TV_KEY_SEP & "*"              ' Item Tag festlegen (* statt ID)
'                LVItem.Key = LV.Tag & TV_KEY_SEP & szTmpID          ' Item Key Festlegen mit ID
'                For i = 2 To LV.ColumnHeaders.Count                 ' Für jedes Feld ein SubItem
'                    Call AddListViewSubItem(LVItem, Trim(objTools.checknull(rsList.Fields(LV.ColumnHeaders(i).Text).Value, "")))
'                Next i
'                szTmpName = ""
'                szTmpID = ""
'            End If
'            rsList.MoveNext                                         ' Nächster DS
'        Loop
'    End If
'
'    If Not bNoColWidth Then                                         ' Spalten Breite einstellen
'        If LV.ListItems.Count > 0 And bOptColumnWidth Then          ' Opt. Spalten breite
'            Call OptimalHeaderWidth(LV)                             ' Optiomale Spalten breite einstellen
'        Else
'            Call LoadColumnWidth(LV)                                ' Spalten breite anhand Tag aus Registry
'        End If
'    End If
'
'    Set FillLVBySQL = rsList                                        ' RS mit daten als Rückgabe wert
'
'exithandler:
'
'Exit Function
'Errorhandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.Errorhandler(MODULNAME, "FillLVBySQL", errNr, errDesc, szDetails)
'    Resume exithandler
'End Function

Public Function FillLVByRS(LV As ListView, szTag As String, rsList As ADODB.Recordset, _
        Optional bShowValueList As Boolean, _
        Optional lngImagindex As Integer, _
        Optional szIndexField As String, _
        Optional bOptColumnWidth As Boolean, _
        Optional bNoColWidth As Boolean, _
        Optional szKey As String, _
        Optional lngAltImageIndex As Integer, _
        Optional AltImageConditionField As String, _
        Optional AltImageConditionValue As Variant, _
        Optional AltImageConditionOperation As String)
' Füllt ListView indem anhand des RS
' bShowValueList = True gibt an ob ein DS untereinader angezeit wird Pro Feld ein Item)
' bShowValueList = False Listet alle DS auf
' bOptColumnWidth = true die Splatenbreite wird optimal eingestellt
'   Sonst wird versicht die Splatenbreite aus der Reg zu laden
    Dim LVItem As ListItem                                          ' ListView item
    Dim i As Integer                                                ' Counter
    Dim ci As Integer                                               ' Akt ColumnIndex
    Dim szTmpName As String                                         ' Item text
    Dim szTmpID As String                                           ' Evtl DS ID
    Dim szTagArray() As String                                      ' LV.Tag in array aufgespalten
    Dim NewTag As String                                            ' Evtl. Neu generierte Tag
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    Dim lngUsedImgIndex As Integer                                  ' ImgIndex der tatsächlich verwendet wird
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Call ClearLV(LV)                                                ' ListView Aufräumen
    If rsList Is Nothing Then GoTo exithandler                      ' Kein RS -> Fertig
    'If rs.RecordCount = 0 Then GoTo exithandler ' ?
    If IsEmpty(lngImagindex) Then lngImagindex = 0                  ' Defaultwert Imagindex
    If IsEmpty(lngAltImageIndex) Then lngAltImageIndex = 0          ' Defaultwert Alternatives Image Indes
    If bShowValueList Then                                          ' Alle felder eines DS untereinader auflisten
        Call AddLVColumn(LV, "Eigenschaft", 2000)                   ' Feste Spalte Eigenschaft (enthält Feldnamen)
        Call AddLVColumn(LV, "Wert", 4000)                          ' Feste Spalte Wert (enthält Feldwert)
        If rsList Is Nothing Then GoTo exithandler
        If rsList.RecordCount = 0 Then GoTo exithandler             ' Keine Daten -> Fertig
        rsList.MoveFirst                                            ' Nur ersten DS anzeigen
        szTmpID = Trim(objTools.checknull(rsList.Fields(0).Value, ""))
                                                                    ' Das Letzte * im LV.Tag durch ID ersetzen
        szTagArray = Split(LV.Tag, TV_KEY_SEP)                      ' LV.Tag in array aufspalten
        For i = 0 To UBound(szTagArray) - 1                         ' Bis auf letzten eintrag wieder zusammen setzen
            NewTag = NewTag & szTagArray(i) & TV_KEY_SEP
        Next
        NewTag = NewTag & szTmpID                                   ' DS ID an NeuenTag anhängen
        For i = 0 To rsList.Fields.Count - 1                        ' Für jedes Feld eine Zeile
            If Left(rsList.Fields(i).Name, 2) <> "ID" Then           ' Wenn Feld mit ID anfängt ausblenden
                szTmpName = Trim(objTools.checknull(rsList.Fields(i).Value, ""))
                Set LVItem = AddListViewItem(LV, rsList.Fields(i).Name, _
                        szTmpName, NewTag, lngImagindex)            ' ListViewItem anlegen
                'LVItem.Key = szKey
                LVItem.Icon = lngImagindex                          ' Icon Setzen
                szTmpName = ""
            End If
        Next i                                                      ' Nächstes Feld
    Else                                                            ' Alle DS untereinader auflisten
        LV.MousePointer = vbHourglass                               ' Sanduhr anzeigen
                                                                    ' ColumsHeader anlegen
        For i = 1 To rsList.Fields.Count - 1                        ' Erste Spalte ist ID -> auslassen
            ci = AddLVColumn(LV, rsList.Fields(i).Name)             ' Colum hinzufügen
            Select Case rsList.Fields(i).Type                       ' Daten Typ Prüfen
            Case adSmallInt, adInteger, adBigInt, adSingle          ' Zahlen
                LV.ColumnHeaders(ci).Tag = adInteger                ' DatenFormat in ColumnKey zum sortieren
            Case adDate, adDBDate, adDBTimeStamp, adDBTime          ' Datum
                LV.ColumnHeaders(ci).Tag = adDate                   ' DatenFormat in ColumnKey zum sortieren
            Case Else
            End Select
        Next i                                                      ' Nächstes Feld
                                                                    ' Jedes Listitem bekommet als Tag den Tag des LVs + Datensat ID eingetragen
        Do While Not rsList.EOF                                     ' Für jeden DS einen eintrag
            If szIndexField <> "" Then                              ' explicit angegebenes ID feld
                szTmpID = Trim(objTools.checknull(rsList.Fields(szIndexField).Value, ""))
            End If
            If szTmpID = "" Then szTmpID = Trim(objTools.checknull( _
                    rsList.Fields(0).Value, ""))                    ' Sonst 1. Feldwert ID
            szTmpName = Trim(objTools.checknull( _
                    rsList.Fields(1).Value, ""))                    ' 2. Feldwert ItemText
            If szTmpID <> "" And szTmpName = "" Then szTmpName = " (Leer) "
            If szTmpName <> "" And szTmpID <> "" Then
                szDetails = "ItemKey: " & LV.Tag & TV_KEY_SEP & szTmpID & vbCrLf & "ItemTag: " & LV.Tag & TV_KEY_SEP & "*" & vbCrLf
                'Set LVItem = AddListViewItem(LV, szTmpName, , LV.Tag & TV_KEY_SEP & szTmpID, lngImageIndex)
                'LVItem.Key = LV.Tag & TV_KEY_SEP & "*"
                lngUsedImgIndex = lngImagindex                      ' ImageIndex Setzen
                If AltImageConditionField <> "" Then   ' Auf alternatives Image prüfen
                    If objTools.CheckStringOperation(AltImageConditionOperation, _
                            objTools.checknull(rsList.Fields(AltImageConditionField).Value, ""), _
                            AltImageConditionValue) Then            '
                        lngUsedImgIndex = lngAltImageIndex          ' Alternatives Image setzen
                    End If
                End If
'                If AltImageConditionField <> "" Then                ' Auf alternatives Image prüfen
'                    If objTools.checknull(rsList.Fields(AltImageConditionField).Value, "") _
'                            = AltImageConditionValue Then
'                        lngUsedImgIndex = lngAltImageIndex          ' Alternatives Image setzen
'                    End If
'                End If
                Set LVItem = AddListViewItem(LV, szTmpName, , LV.Tag & TV_KEY_SEP & _
                        szTmpID, lngUsedImgIndex)                   ' ListView Item anlegen
                szDetails = szDetails & "Item angelegt."            ' Detail infos für Fehlermeldung
                LVItem.Tag = LV.Tag & TV_KEY_SEP & "*"              ' Item Tag setzen
                If szKey <> "" Then                                 ' Item Key Setzen
                    LVItem.Key = szKey & TV_KEY_SEP & szTmpID
                Else
                    LVItem.Key = LV.Tag & TV_KEY_SEP & szTmpID
                End If
                For i = 2 To LV.ColumnHeaders.Count                 ' Für jedes Feld ein SubItem
                    Call AddListViewSubItem(LVItem, Trim(objTools.checknull(rsList.Fields( _
                            LV.ColumnHeaders(i).Text).Value, "")))  ' ListViewSubItem anlegen
                Next i                                              ' nächstes Feld
                szTmpName = ""
                szTmpID = ""
            End If
            rsList.MoveNext                                         ' Nächster DS
        Loop
    End If
    
    If Not bNoColWidth Then
        If LV.ListItems.Count > 0 And bOptColumnWidth Then
            Call OptimalHeaderWidth(LV)                             ' Optiomale Spalten breite einstellen
        Else
            Call LoadColumnWidth(LV, "", True)                      ' Spalten breite anhand Tag aus Registry
        End If
    End If
    
    Set FillLVByRS = rsList                                         ' RS mit daten als Rückgabe wert
    
exithandler:
On Error Resume Next
    LV.MousePointer = vbDefault                                     ' Mauszeiger wieder normal
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "FillLVByRS", errNr, errDesc, szDetails)
    If errNr = 35602 Then Resume Next
    Resume exithandler
End Function

'Public Function FillLVByRS(LV As ListView, szTag As String, rsList As ADODB.Recordset, _
'        Optional bShowValueList As Boolean, _
'        Optional lngImagindex As Integer, _
'        Optional szIndexField As String, _
'        Optional bOptColumnWidth As Boolean, _
'        Optional bNoColWidth As Boolean, _
'        Optional szKey As String, _
'        Optional lngAltImageIndex As Integer, _
'        Optional AltImageConditionField As String, _
'        Optional AltImageConditionValue As Variant, _
'        Optional AltImageConditionOperation As String)
'
'' Füllt ListView indem anhand des RS
'' bShowValueList = True gibt an ob ein DS untereinader angezeit wird Pro Feld ein Item)
'' bShowValueList = False Listet alle DS auf
'' bOptColumnWidth = true die Splatenbreite wird optimal eingestellt
''   Sonst wird versicht die Splatenbreite aus der Reg zu laden
'
'    Dim LVItem As ListItem                                          ' ListView item
'    Dim i As Integer                                                ' Counter
'    Dim ci As Integer                                               ' Akt ColumnIndex
'    Dim szTmpName As String                                         ' Item text
'    Dim szTmpID As String                                           ' Evtl DS ID
'    Dim szTagArray() As String                                      ' LV.Tag in array aufgespalten
'    Dim NewTag As String                                            ' Evtl. Neu generierte Tag
'    Dim szDetails As String                                         ' Details für Fehlerbehandlung
'    Dim lngUsedImgIndex As Integer                                  ' ImgIndex der tatsächlich verwendet wird
'
'On Error GoTo Errorhandler
'
'    Call ClearLV(LV)                                                ' ListView Aufräumen
'
'
'    If rsList Is Nothing Then GoTo exithandler                      ' Kein RS -> Fertig
'    'If rs.RecordCount = 0 Then GoTo exithandler ' ?
'
'    If IsEmpty(lngImagindex) Then lngImagindex = 0
'    If IsEmpty(lngAltImageIndex) Then lngAltImageIndex = 0
'    If AltImageConditionOperation = "" Then AltImageConditionOperation = "="
'
'    If bShowValueList Then                                          ' Alle felder eines DS untereinader auflisten
'        Call AddLVColumn(LV, "Eigenschaft", 2000)                   ' Feste Spalte Eigenschaft (enthält Feldnamen)
'        Call AddLVColumn(LV, "Wert", 4000)                          ' Feste Spalte Wert (enthält Feldwert)
'        If rsList Is Nothing Then GoTo exithandler
'        If rsList.RecordCount = 0 Then GoTo exithandler             ' Keine Daten -> Fertig
'        rsList.MoveFirst                                            ' Nur ersten DS anzeigen
'        szTmpID = Trim(objTools.checknull(rsList.Fields(0).Value, ""))
'                                                                    ' Das Letzte * im LV.Tag durch ID ersetzen
'        szTagArray = Split(LV.Tag, TV_KEY_SEP)                      ' LV.Tag in array aufspalten
'        For i = 0 To UBound(szTagArray) - 1                         ' Bis auf letzten eintrag wieder zusammen setzen
'            NewTag = NewTag & szTagArray(i) & TV_KEY_SEP
'        Next
'        NewTag = NewTag & szTmpID                                   ' DS ID an NeuenTag anhängen
'        For i = 0 To rsList.Fields.Count - 1                        ' Für jedes Feld eine Zeile
'                                                                    ' Wenn Feld mit ID anfängt ausblenden
'            If Left(rsList.Fields(i).Name, 2) <> "ID" Then
'                szTmpName = Trim(objTools.checknull(rsList.Fields(i).Value, ""))
'                Set LVItem = AddListViewItem(LV, rsList.Fields(i).Name, szTmpName, NewTag, lngImagindex)
'                'LVItem.Key = szKey
'                LVItem.Icon = lngImagindex                          ' Icon Setzen
'                szTmpName = ""
'            End If
'        Next i
'    Else                                                            ' Alle DS untereinader auflisten
'        LV.MousePointer = vbHourglass                               ' Sanduhr anzeigen
'                                                                    ' ColumsHeader anlegen
'        For i = 1 To rsList.Fields.Count - 1                        ' Erste Spalte ist ID -> auslassen
'            ci = AddLVColumn(LV, rsList.Fields(i).Name)             ' Colum hinzufügen
'            Select Case rsList.Fields(i).Type                       ' Daten Typ Prüfen
'            Case adSmallInt, adInteger, adBigInt, adSingle          ' Zahlen
'                LV.ColumnHeaders(ci).Tag = adInteger                ' DatenFormat in ColumnKey zum sortieren
'            Case adDate, adDBDate, adDBTimeStamp, adDBTime          ' Datum
'                LV.ColumnHeaders(ci).Tag = adDate                   ' DatenFormat in ColumnKey zum sortieren
'            Case Else
'            End Select
'        Next i
'                                                                    ' Jedes Listitem bekommet als Tag den Tag des LVs + Datensat ID eingetragen
'        Do While Not rsList.EOF                                     ' Für jeden DS einen eintrag
'            If szIndexField <> "" Then                              ' explicit angegebenes ID feld
'                szTmpID = Trim(objTools.checknull(rsList.Fields(szIndexField).Value, ""))
'            End If
'            If szTmpID = "" Then szTmpID = Trim(objTools.checknull(rsList.Fields(0).Value, "")) ' Sonst 1. Feldwert ID
'            szTmpName = Trim(objTools.checknull(rsList.Fields(1).Value, "")) ' 2. Feldwert ItemText
'            If szTmpID <> "" And szTmpName = "" Then szTmpName = " (Leer) "
'            If szTmpName <> "" And szTmpID <> "" Then
'                szDetails = "ItemKey: " & LV.Tag & TV_KEY_SEP & szTmpID & vbCrLf & "ItemTag: " & LV.Tag & TV_KEY_SEP & "*" & vbCrLf
'
'                'Set LVItem = AddListViewItem(LV, szTmpName, , LV.Tag & TV_KEY_SEP & szTmpID, lngImageIndex)
'                'LVItem.Key = LV.Tag & TV_KEY_SEP & "*"
'                lngUsedImgIndex = lngImagindex
'                If AltImageConditionField <> "" Then   ' Auf alternatives Image prüfen
'                    Select Case AltImageConditionOperation
'                    Case "="
'                        If objTools.checknull(rsList.Fields(AltImageConditionField).Value, "") _
'                                = AltImageConditionValue Then
'                            lngUsedImgIndex = lngAltImageIndex      ' Alternatives Image setzen
'                        End If
'                    Case ">"
'                        If objTools.checknull(rsList.Fields(AltImageConditionField).Value, "") _
'                                > AltImageConditionValue Then
'                            lngUsedImgIndex = lngAltImageIndex      ' Alternatives Image setzen
'                        End If
'                    Case "<"
'                        If objTools.checknull(rsList.Fields(AltImageConditionField).Value, "") _
'                                < AltImageConditionValue Then
'                            lngUsedImgIndex = lngAltImageIndex      ' Alternatives Image setzen
'                        End If
'                    Case ">="
'                        If objTools.checknull(rsList.Fields(AltImageConditionField).Value, "") _
'                                >= AltImageConditionValue Then
'                            lngUsedImgIndex = lngAltImageIndex      ' Alternatives Image setzen
'                        End If
'                    Case "<="
'                        If objTools.checknull(rsList.Fields(AltImageConditionField).Value, "") _
'                                <= AltImageConditionValue Then
'                            lngUsedImgIndex = lngAltImageIndex      ' Alternatives Image setzen
'                        End If
'                    Case Else
'
'                    End Select
'                End If
'                Set LVItem = AddListViewItem(LV, szTmpName, , LV.Tag & TV_KEY_SEP & szTmpID, lngUsedImgIndex)
'                szDetails = szDetails & "Item angelegt."            ' Detail infos für Fehlermeldung
'                LVItem.Tag = LV.Tag & TV_KEY_SEP & "*"              ' Item Tag setzen
'                If szKey <> "" Then                                 ' Item Key Setzen
'                    LVItem.Key = szKey & TV_KEY_SEP & szTmpID
'                Else
'                    LVItem.Key = LV.Tag & TV_KEY_SEP & szTmpID
'                End If
'                For i = 2 To LV.ColumnHeaders.Count                 ' Für jedes Feld ein SubItem
'                    Call AddListViewSubItem(LVItem, Trim(objTools.checknull(rsList.Fields(LV.ColumnHeaders(i).Text).Value, "")))
'                Next i
'                szTmpName = ""
'                szTmpID = ""
'            End If
'            rsList.MoveNext                                         ' Nächster DS
'        Loop
'    End If
'
'    If Not bNoColWidth Then
'        If LV.ListItems.Count > 0 And bOptColumnWidth Then
'            Call OptimalHeaderWidth(LV)                             ' Optiomale Spalten breite einstellen
'        Else
'            Call LoadColumnWidth(LV, "", True)                      ' Spalten breite anhand Tag aus Registry
'        End If
'    End If
'
'    Set FillLVByRS = rsList                                         ' RS mit daten als Rückgabe wert
'
'exithandler:
'On Error Resume Next
'    LV.MousePointer = vbDefault                                     ' Mauszeiger wieder normal
'Exit Function
'Errorhandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.Errorhandler(MODULNAME, "FillLVByRS", errNr, errDesc, szDetails)
'    If errNr = 35602 Then Resume Next
'    Resume exithandler
'End Function

Public Function ListLVByTag(LV As ListView, dbCon As Object, _
        szTag As String, _
        Optional szWhereKey As String, _
        Optional bShowValueList As Boolean, _
        Optional lngImageIndex As Integer, _
        Optional bNotColWidth As Boolean, _
        Optional szKey As String) As ADODB.Recordset
' Füllt ListView indem anhand des Tag ein SQLStatement aus XML gelesen Wird

    Dim LVInfo As ListViewInfo                                      ' ListView Infos
    Dim bShowDel As Boolean                                         ' Als gelöscht gesetzte DS anzeigen
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
      
    Call ClearLV(LV)                                                ' ListView Aufräumen
    
    bShowDel = objOptions.GetOptionByName(OPTION_SHOWDELREL)        ' Option Gelöschte DS anzeigen auslesen
    
    LV.Tag = szTag '& TV_KEY_SEP & "*"                              ' Tag des Listviews setzen

    With LVInfo
        Call objTools.GetLVInfoFromXML(App.Path & "\" & INI_XMLFILE, szTag, .szSQL, .szTag, _
                .szWhere, .lngImage, .bValueList, .bListSubNodes, , , , .AltImage, .AltImgField, _
                    .AltImgValue, .DelFlagField)                    ' Listview infos aus XMl datei holen
        'GetLVInfoFromXML(ByVal XMLDocPath As String, NodeTag As String, _
    szSQL As String, newTag As String, szWhere As String, Optional newImage As Integer, _
    Optional bValueList As Boolean, Optional bSubNodes As Boolean, _
    Optional bEdit As Boolean, Optional bNew As Boolean, Optional bSelectNode As Boolean, _
    Optional lngAltImage As Integer, Optional szAktImgField As String, _
            Optional szAltImgValue As String)
        If szWhereKey <> "" And .szSQL <> "" Then                   ' Evtl. Where bedingung anhängen
            If .szWhere <> "" Then .szWhere = .szWhere & "'" & szWhereKey & "'"
            If .DelFlagField <> "" Then                             ' Gibt es ein gelöscht flag ?
                .WhereNoDel = objSQLTools.AddWhere(.szWhere, .DelFlagField & "=0") ' Delflag mit in Where einbauen
            Else
                .WhereNoDel = .szWhere                              ' Sont Where statements gleich
            End If
            If bShowDel Then                                        ' Sollen als gelöscht gekennzeichnete DS angezeigt werden?
                If .szWhere <> "" Then .szSQL = objSQLTools.AddWhereInFullSQL( _
                        .szSQL, .szWhere)                           ' Where Ohne Delflag Filter
            Else
                If .WhereNoDel <> "" Then .szSQL = objSQLTools.AddWhereInFullSQL( _
                        .szSQL, .WhereNoDel)                        ' Where mit delFlag Filter
            End If
        End If
    
        If .szSQL = "" Then GoTo exithandler                        ' Kein SQL Statement -> Fertig
        lngImageIndex = .lngImage                                   ' Imageindex Setzen
        'If .lngImage <> "" And IsNumeric(.lngImage) Then lngImageindex = CLng(.lngImage) ' Evtl. Image index  aus Ini lesen
    
        Set ListLVByTag = FillLVBySQL(LV, .szSQL, dbCon, .bValueList, lngImageIndex, _
                , , bNotColWidth, szKey, .AltImage, .AltImgField, _
                .AltImgValue)                                       ' Listview aus SQL statement füllen
     End With
     
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ListLVByTag", errNr, errDesc)
End Function

Public Function ListLVFromSubNodes(LV As ListView, TV As TreeView, ThisNode As node)
' fügt die 1. Ebene Unterknoten dem Listviw als Items hinzu

    Dim subNode As node                                             ' sub node der im LV angezeigt werden soll
    Dim LVItem As ListItem                                          ' neu angelegte LV Item
    Dim KeyArray() As String                                        ' Node Key in Array
    Dim szDescription As String                                     ' Evtl Beschreibung des Subnodes
    Dim Nodeinfo As TreeViewNodeInfo                                ' Infos des subnodes aus XML
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    KeyArray = Split(ThisNode.Key, TV_KEY_SEP)                      ' Node Key in Array aufspalten

    If LV.ColumnHeaders.Count = 0 Then                              ' Wenn keine Column header
        Call AddLVColumn(LV, ThisNode.Text, 2000)                   ' einen anlegen
        Call AddLVColumn(LV, "Beschreibung", LV.Width - 2100)       ' einen anlegen
    End If

    If ThisNode.Children > 0 Then                                   ' Wenn Node Subnodes Besitzt
        Set subNode = ThisNode.Child                                ' Subnode holen

        While Not subNode Is Nothing                                ' Solage Subnodes gefunden werden
            If Not subNode Is Nothing Then                          ' Wenn SubNode nicht Nothing
                KeyArray = Split(subNode.Key, TV_KEY_SEP)           ' SubnodeKey aufspalten
                With Nodeinfo
                    Call objTools.GetTVNodeInfofromXML(App.Path & "\" & INI_XMLFILE, _
                            subNode.Tag, .szTag, .szText, .szKey, .bShowSubnodes, .szSQL, .szWhere, _
                            .lngImage, .bShowKontextMenue, .szDesc) ' Tree node informationen aus XML laden
                    Set LVItem = AddListViewItem(LV, subNode.Text, .szDesc, _
                            subNode.Key, subNode.Image)             ' List Item anlegen
                    LVItem.Tag = subNode.Tag                        ' Tag setzen
                    LVItem.Key = subNode.Key                        ' Key setzen
                End With
                Set subNode = subNode.Next                          ' nächster subnode
            End If
        Wend
    End If

exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ListLVFromSubNodes", errNr, errDesc)
    Resume exithandler
End Function
                                                                    ' *****************************************
                                                                    ' LV Colum Header Funktionen
Public Function OptimalHeaderWidth(LV As ListView)
' Setzt optimale Spalten breite
' Columntext wird nicht berücksichtigt
' Daher wenn leere Spalte Breite = 0

    Dim i As Integer                                                ' counter
    Const LVSCW_AUTOSIZE = (-1)                                     ' automatisch optimale Spaltenbreite
    Const LVSCW_AUTOSIZE_USEHEADER = (-2)                           ' sorgt dafür, dass alle Spaltenüberschriften lesbar sind

On Error GoTo Errorhandler
    
    For i = 0 To LV.ColumnHeaders.Count - 1                         ' Alle Spaltenköpfe durchlaufen
        SendMessage LV.hWnd, &H101E, i, LVSCW_AUTOSIZE
    Next
    
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "OptimalHeaderWidth", errNr, errDesc)
    Resume exithandler
End Function

Public Function AddLVColumn(LV As ListView, szHeaderName As String, Optional ColWidith As Integer) As Integer
' Fügt eine Spalte dem Listview hinzu

    Dim i As Integer                                                ' counter
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    ColWidith = CLng(ColWidith)                                     ' Fall ColWidth nicht angegeben
    
    If szHeaderName <> "" Then                                      ' Header Name vorhanden
        For i = 1 To LV.ColumnHeaders.Count                         ' Feststellen ob es den Header schon gibt
            If LV.ColumnHeaders(i).Text = szHeaderName Then
                AddLVColumn = i                                     ' Column index als Rückgabe wert
                GoTo exithandler                                    ' Raus
            End If
        Next i                                                      ' Nächste Spalte
        
        LV.HideColumnHeaders = False                                ' ColumHeaders Anzeigen einstellen
        If ColWidith <> 0 Then
            LV.ColumnHeaders.Add , , szHeaderName, ColWidith        ' ColumHeader mit breite setzen
        Else
            LV.ColumnHeaders.Add , , szHeaderName                   ' ColumnHeader ohne breite setzen
        End If
        AddLVColumn = LV.ColumnHeaders.Count                        ' Column index als Rückgabe wert
    Else
        LV.ColumnHeaders.Add , , , 0                                ' Unsichtbare spalte sezen
        AddLVColumn = LV.ColumnHeaders.Count                        ' Column index als Rückgabe wert
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
    Call objError.Errorhandler(MODULNAME, "AddLVColumn", errNr, errDesc)
    Resume exithandler
End Function

Public Function SetColumnOrder(LV As ListView, ColumnHeader As MSComctlLib.ColumnHeader)
' Sortiert die angegebene Spalte ja nach eltl forhandener Sortierung entgegengesetzt
' Auch Datm und Zahlen werden korrekt sortiert, wenn im Tag des Headers
' die konstaten adInteger oder ad Date eingetragen sind
' SortOrder=0 Aufsteigend
' SortOrder=1 Absteigend
    
    Dim i As Integer                                                ' Counter (listItems)
    Dim x As Integer                                                ' noch ein Counter (Column header)
    Dim NewSub As Long                                              ' Index der Dummy Spalte bei Datum oder Zahl
    Dim sFormat As String                                           ' Sortierbares Datumsformat
    Dim Li As ListItem                                              ' LI Current Listview Item
    Dim bDummyColum As Boolean                                      ' Hilfs Spalte
    Dim OldLVItem As ListItem                                       ' Aktives ListItem Vor sortierung

On Error GoTo Errorhandler                                          ' Fehlerbehandlung wieder aktivieren
    
    Set OldLVItem = GetSelectListItem(LV)                           ' Akt. ausgewähltes ListItem merken
    LV.MousePointer = vbHourglass                                   ' Sanduhr anzeigen (kann bei datum etwas dauern)
    sFormat = "yyyy.mm.dd hh:mm:ss"                                 ' Sortierbares Datumsformat setzen
    i = 0

    With LV
        .Visible = False                                            ' ListView ruhig halten, Sichtbarkeit bleibt trotzdem erhalten
        For x = 1 To .ColumnHeaders.Count                           ' Erst für alle Columnheader
            .ColumnHeaders(x).Icon = 0                              ' alle icons ausblenden
        Next x                                                      ' nächster ColumnHeader
        .SortKey = ColumnHeader.Index - 1                           ' zu sortierende Spalte bestimmen
        .ColumnHeaders.Add , , "Dummy", 0                           ' Dummy-Spalte einfügen mit Breite 0
        bDummyColum = True                                          ' DummeSpalte angefügt
        NewSub = .ColumnHeaders.Count - 1                           ' Nummer der Dummy-Spalte
        If ColumnHeader.Tag = adDate Then                           ' abfragen auf Spalte mit Datum
            For i = .ListItems.Count To 1 Step -1                   ' Sortiere nach Datum
                Set Li = .ListItems(i)                              ' Akt ListItem holen
                If Li.SubItems(ColumnHeader.Index - 1) = "" Then    ' Dummy-Spalte mit sortierfähigem Datum belegen
                    Li.SubItems(NewSub) = Format(CDate(vbNull), sFormat)
                Else
                    Li.SubItems(NewSub) = Format(CDate(Li.SubItems(ColumnHeader.Index - 1)), sFormat)
                End If
            Next i                                                  ' Nächstes Listitem
            .SortKey = NewSub                                       ' zu sortierende Spalte umbiegen
        ElseIf ColumnHeader.Tag = adInteger Then                    ' abfragen auf Spalte mit Zahlen
            For i = .ListItems.Count To 1 Step -1                   ' Sortiere nach Zahlen
                Set Li = .ListItems(i)                              ' Akt ListItem holen
                Li.SubItems(NewSub) = Right(Space(20) & _
                        Li.SubItems(ColumnHeader.Index - 1), 20)    ' Dummy-Spalte mit sortierfähiger Zahl belegen
            Next i                                                  ' Nächstes Listitem
            .SortKey = NewSub                                       ' zu sortierende Spalte umbiegen
        End If
        
        If .SortOrder = 0 Then                                      ' SortOrder bestimmen Asc oder Desc
            .SortOrder = 1                                          ' Sort Order umdrehen
            .ColumnHeaders(ColumnHeader.Index).Icon = IMG_SORTUP    ' Entsprechendes Icon im Columheader setzen
        Else
            .SortOrder = 0                                          ' Sort Order umdrehen
            .ColumnHeaders(ColumnHeader.Index).Icon = IMG_SORTDOWN  ' Entsprechendes Icon im Columheader setzen
        End If
        
        .Sorted = True                                              ' Sort anstossen
        If bDummyColum Then .ColumnHeaders.Remove .ColumnHeaders.Count  ' Dummy-Spalte entfernen
        If .ListItems.Count > 0 Then                                ' Liste nicht Leer
            .ListItems(1).Selected = True                           ' Zeiger auf 1. Zeile und scrollen
            .ListItems(1).EnsureVisible
        End If
        .Visible = True                                             ' sichtbar machen
    End With
    
exithandler:
On Error Resume Next
    Call SelectLVItem(LV, OldLVItem.Key)                            ' Evtl. Vor sort. selected Item wieder selecten
    LV.MousePointer = vbDefault                                     ' Mauszeiger wieder normal
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SetColumnOrder", errNr, errDesc)
End Function

Public Sub SaveColumnWidth(LV As ListView, Optional TagPreFix As String, Optional bForce As Boolean)
' Speichert Spaltenbreiten eines ListViews in der Registry unter HKCurrentUser
' Analog zu LoadColumnWidth

    Dim i As Integer                                                ' Counter
    Dim szRegKey As String                                          ' Reg Schlüssel
    Dim szRegValue As String                                        ' Reg Wert
    
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren
    
    If LV.Tag = "" Then GoTo exithandler                            ' Kein LV Tag -> raus
    
    szRegKey = "SOFTWARE\" & objObjectBag.GetAppTitle & "\Columns"  ' Reg Schlüssel mit anwendungsnahmen festlegen
    
    If LV.ColumnHeaders.Count = 0 Then GoTo exithandler             ' keine Spalten -> Fertig
    If LV.ColumnHeaders.Count > 2 Or bForce Then                    ' 2 Spaltige LV nicht berücksichtigen, se sei den der Parameter bForce = true
        For i = 1 To LV.ColumnHeaders.Count                         ' Alle Spalten durchlaufen
            szRegValue = szRegValue & LV.ColumnHeaders(i).Width & ";" ' einzelne Splatenbreiten mit ; Trennen
        Next i                                                      ' Nächste Spalte
        Call objRegTools.WriteRegValue("HKCU", szRegKey, TagPreFix & LV.Tag, _
                Left(szRegValue, Len(szRegValue) - 1))              ' in registry Schreiben, evtl. otionalen Prefix anhängen
    End If
    
exithandler:

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SaveColumnWidth", errNr, errDesc)
    Resume exithandler
End Sub

Public Sub LoadColumnWidth(LV As ListView, Optional TagPreFix As String, Optional bForce As Boolean)
' Lädt Spaltenbreiten eines ListViews aus der Registry unter HKCurrentUser
' Analog zu SaveColumnWidth

    Dim i As Integer                                                ' Counter
    Dim szRegKey As String                                          ' Reg Schlüssel
    Dim szRegValue As String                                        ' Reg Wert
    Dim ColArray() As String                                        ' Array mit Spalten breiten
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    If LV.Tag = "" Then GoTo exithandler                            ' Kein LV Tag -> raus
    LV.Visible = False                                              ' ListView ruhig halten, Sichtbarkeit bleibt trotzdem erhalten
    szRegKey = "SOFTWARE\" & objObjectBag.GetAppTitle & "\Columns"  ' Reg Schlüssel mit anwendungsnahmen festlegen
    szRegValue = objRegTools.ReadRegValue("HKCU", szRegKey, TagPreFix & LV.Tag) ' Reg Wert lesen, evtl. otionalen Prefix anhängen
    If szRegValue <> "" Then                                        ' Wert gefunden
        If LV.ColumnHeaders.Count > 2 Or bForce Then                ' 2 Spaltige LV nicht berücksichtigen, se sei den derPArameter bForce = true
            ColArray = Split(szRegValue, ";")                       ' Wert in Array aufspalten
            For i = 0 To UBound(ColArray)                           ' Array durchlaufen
                Call SetColumnWidth(LV, i + 1, CLng(ColArray(i)))   ' Spalten breite setzen
            Next i                                                  ' Nächstes Array Item
        End If
    End If
    
exithandler:
On Error Resume Next
    LV.Visible = True                                               ' LV wieder einblenden
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "LoadColumnWidth", errNr, errDesc)
    Resume exithandler
End Sub
                                                                    ' *****************************************
                                                                    ' Form Funktionen
'Public Sub SetWindowTransparency(F As Form, sinPercent As Single)
'' Setzt Fenster transparenz
'' funktioniert nur unter Windows 2000 oder XP!!!
'' RateOfT: 254 = normal 0 = ganz transparent (also unsichtbar)
'    Dim WinInfo As Long                                             ' Windows Handel
'    Dim RateOfT As Byte                                             ' Transparenz Rate 255 Keine, 0 Totale Tranzparenz
'    If F Is Nothing Then GoTo exithandler                           ' Kein Form -> Fertig
'    WinInfo = GetWindowLong(F.hWnd, GWL_EXSTYLE)
'    If sinPercent < 0.1 Then sinPercent = 0.1                       ' Minimal tranzparenz festlegen
'    RateOfT = sinPercent * 254
'    If RateOfT < 255 Then                                           ' Tranzparenz setzen
'        WinInfo = WinInfo Or WS_EX_LAYERED
'        SetWindowLong F.hWnd, GWL_EXSTYLE, WinInfo
'        SetLayeredWindowAttributes F.hWnd, 0, RateOfT, LWA_ALPHA
'    Else                                                            ' Wenn als Rate 255 angegeben wird,
'        WinInfo = WinInfo Xor WS_EX_LAYERED                         ' so wird der Ausgangszustand wiederhergestellt
'        SetWindowLong F.hWnd, GWL_EXSTYLE, WinInfo
'    End If
'exithandler:
'
'End Sub

Public Function MousePointerHourglas(Optional F As Form, Optional CTL As Control)
' Setzt MousPointer auf Hourglass fürs angegebene Form sonst global
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If Not CTL Is Nothing Then                                      ' CTL vorhanden
        CTL.MousePointer = vbHourglass                              ' Mouspointer setzen
        GoTo exithandler                                            ' Fertig
    End If
     If Not F Is Nothing Then                                       ' Formular angegeben
        F.MousePointer = vbHourglass                                ' Mouspointer setzen
        GoTo exithandler                                            ' Fertig
    End If
    If F Is Nothing Then                                            ' Kein Form angegeben
        Set F = objObjectBag.getMainForm                            ' Main Form auf Obag holen
        F.MousePointer = vbHourglass                                ' Mouspointer setzen
    End If
    
exithandler:
    Err.Clear                                                       ' Evtl Error Clearen
End Function

Public Function MousePointerDefault(Optional F As Form, Optional CTL As Control)
' Setzt MousPointer auf Default fürs angegebene Form sonst global
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If Not CTL Is Nothing Then                                      ' CTL vorhanden
        CTL.MousePointer = vbDefault                                ' Mouspointer setzen
        GoTo exithandler                                            ' Fertig
    End If
    If Not F Is Nothing Then                                        ' Formular angegeben
        F.MousePointer = vbDefault                                  ' Mouspointer setzen
        GoTo exithandler                                            ' Fertig
    End If
    If F Is Nothing Then                                            ' Kein Form angegeben
        Set F = objObjectBag.getMainForm                            ' Main Form auf Obag holen
        F.MousePointer = vbDefault                                  ' Mouspointer setzen
    End If

exithandler:
    Err.Clear                                                       ' Evtl Error Clearen
End Function

Public Function MousePointerLink(Optional F As Form, Optional CTL As Control)
' Setzt MousPointer auf Hand fürs angegebene Form / CTL sonst global
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If Not CTL Is Nothing Then                                      ' CTL vorhanden
        CTL.MousePointer = 99                                       ' Mouspointer auf benutzerdefiniert setzen
        CTL.MouseIcon = LoadPicture(objObjectBag.GetImageDir _
                & "Hand.cur")                                       ' Cursor laden & setzen
        GoTo exithandler                                            ' Fertig
    End If
    If Not F Is Nothing Then                                        ' Formular angegeben
        CTL.MousePointer = 99                                       ' Mouspointer auf benutzerdefiniert setzen
        F.MouseIcon = LoadPicture(objObjectBag.GetImageDir _
                & "Hand.cur")                                       ' Cursor laden & setzen
        GoTo exithandler                                            ' Fertig
    End If
    If F Is Nothing And CTL Is Nothing Then                         ' Kein Form & kein Control angegeben
        Set F = objObjectBag.getMainForm                            ' Main Form auf Obag holen
        CTL.MousePointer = 99                                       ' Mouspointer auf benutzerdefiniert setzen
        F.MouseIcon = LoadPicture(objObjectBag.GetImageDir _
            & "Hand.cur")                                           ' Cursor laden & setzen
    End If
    
exithandler:
    Err.Clear                                                       ' Evtl Error Clearen
End Function

Public Sub SetWindowStateFromString(F As Form, szWinState As String)
' Setzt f.WindowState aus szWinState prüft und wandelt sting um
    Dim lngState As Integer                                         ' Integer von szWinState
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If F Is Nothing Then Exit Sub                                   ' Kein Form dann fertig
    If szWinState <> "" Then                                        ' Ist szWinState leer ?
        If IsNumeric(szWinState) Then                               ' Ist szWinState Zahl ?
            lngState = CLng(szWinState)                             ' In Intwandeln
            If lngState >= 0 And lngState <= 2 Then                 ' 0=normal 1=min 2=max
                F.WindowState = lngState                            ' Window State setzen
            Else
                F.WindowState = vbNormal                            ' <0 bzw. >2 dann Normal
            End If
        Else
            F.WindowState = vbNormal                                ' keine Zahl dann Normal
        End If
    Else
        F.WindowState = vbNormal                                    ' Nix angegeben dann Normal
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Public Sub SetWindowSizeFromString(F As Form, szSize As String, Optional szDelimiter As String)
    Dim szSizeArray() As String                                     ' Array mit größen
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If F Is Nothing Then Exit Sub                                   ' Kein Form dann fertig
    If szSize <> "" Then                                            ' Ist szSize leer ?
        If szDelimiter = "" Then                                    ' Ist Trennzeichen angegeben
            szSizeArray = Split(szSize, "/")                        ' Value aufspliten mit Standard trennzeichen
        Else
            szSizeArray = Split(szSize, szDelimiter)                ' Value aufspliten mit szDelimiter
        End If
        If UBound(szSizeArray) < 1 Then Exit Sub                    ' Array zuklein -> fertig
        If szSizeArray(0) <> "" And szSizeArray(1) <> "" Then       ' Werte vorhanden ?
            If IsNumeric(szSizeArray(0)) _
                    And IsNumeric(szSizeArray(1)) Then              ' Werte sind zahlen
                F.Width = CLng(szSizeArray(0))                      ' (0) = Width
                F.Height = CLng(szSizeArray(1))                     ' (1) = Height
            End If
        End If
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Sub
                                                                    ' *****************************************
                                                                    ' MDI Form Funktionen
Public Sub MDIChildsCascade(MDIParent As Form)
'Alle db Fenster Überlappend Versetzt anzeigen
    
    Dim DBWindowCount As Integer                                    ' Fenster anzahl
    Dim i As Integer                                                ' Counter
    Dim MainWidth As Integer                                        ' Innere Breite des MDI Parent
    Dim MainHeight As Integer                                       ' Innere Höhe des MDI Parent
    Dim DBDiff As Integer                                           ' Versatz der Fenster
    Dim CurDBFrm As Form
    
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    
    DBWindowCount = UBound(DBFormArray)                             ' Anzahl der DB Fenster ermitteln
    Err.Clear                                                       ' Evtl. Error clearen
    
On Error GoTo Errorhandler                                           'Fehlerbehandlung wieder Aktivieren

    If DBWindowCount < 0 Then GoTo exithandler                      ' Keine DB fenster dann raus
    MainWidth = MDIParent.ScaleWidth                                ' Innere Breite des MDI Parent ermitteln
    MainHeight = MDIParent.ScaleHeight                              ' Innere Höhe des MDI Parent ermitteln
    DBDiff = 400
    If DBWindowCount = 0 Then GoTo exithandler                      ' Nur ein fenster -> nix zu tun
    For i = 0 To DBWindowCount
        Set CurDBFrm = DBFormArray(i)
        CurDBFrm.WindowState = vbNormal                             ' Fall max or min erstmal auf Normal
        CurDBFrm.Top = i * DBDiff
        CurDBFrm.Left = i * DBDiff
        CurDBFrm.Width = MainWidth - (i * DBDiff)
        CurDBFrm.Height = MainHeight - (i * DBDiff)
    Next i
    
exithandler:

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "MDICildsCascade", errNr, errDesc)
    Resume exithandler
End Sub

Public Sub MDIChildsHSplit(MDIParent As Form)
'Alle db Fenster untereinader anzeigen

    Dim DBWindowCount As Integer                                    ' Fenster anzahl
    Dim i As Integer                                                ' Counter
    Dim MainWidth As Integer                                        ' Innere Breite des MDI Parent
    Dim MainHeight As Integer                                       ' Innere Höhe des MDI Parent
    Dim DBHeight As Integer                                         ' Höhe der einzelnen DB Windows
    Dim CurDBFrm As Form
    
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    
    DBWindowCount = UBound(DBFormArray)                             ' Anzahl der DB Fenster ermitteln
    Err.Clear                                                       ' Evtl. Error clearen
    
On Error GoTo Errorhandler                                           'Fehlerbehandlung wieder Aktivieren

    If DBWindowCount < 0 Then GoTo exithandler                      ' Keine DB fenster dann raus
    MainWidth = MDIParent.ScaleWidth                                ' Innere Breite des MDI Parent ermitteln
    MainHeight = MDIParent.ScaleHeight                              ' Innere Höhe des MDI Parent ermitteln
    DBHeight = MainHeight / (DBWindowCount + 1)
    If DBHeight < 1000 Then DBHeight = 1000
    If DBWindowCount = 0 Then GoTo exithandler                      ' Nur ein fenster -> nix zu tun
    For i = 0 To DBWindowCount
        Set CurDBFrm = DBFormArray(i)
        CurDBFrm.WindowState = vbNormal                             ' Fall max or min erstmal auf Normal
        CurDBFrm.Top = i * DBHeight
        CurDBFrm.Left = 0
        CurDBFrm.Width = MainWidth
        CurDBFrm.Height = DBHeight
    Next i
    
exithandler:

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "MDIChildsHSplit", errNr, errDesc)
    Resume exithandler
End Sub

Public Sub MDIChildsVSplit(MDIParent As Form)
'Alle db Fenster Nebeneinader anzeigen

    Dim DBWindowCount As Integer                                    ' Fenster anzahl
    Dim i As Integer                                                ' Counter
    Dim MainWidth As Integer                                        ' Innere Breite des MDI Parent
    Dim MainHeight As Integer                                       ' Innere Höhe des MDI Parent
    Dim DBWidth As Integer                                          ' Breite der einzelnen DB Windows
    Dim CurDBFrm As Form
    
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    
    DBWindowCount = UBound(DBFormArray)                             ' Anzahl der DB Fenster ermitteln
    Err.Clear                                                       ' Evtl. Error clearen
    
On Error GoTo Errorhandler                                           'Fehlerbehandlung wieder Aktivieren

    If DBWindowCount < 0 Then GoTo exithandler                      ' Keine DB fenster dann raus
    
    MainWidth = MDIParent.ScaleWidth                                ' Innere Breite des MDI Parent ermitteln
    MainHeight = MDIParent.ScaleHeight                              ' Innere Höhe des MDI Parent ermitteln
    DBWidth = MainWidth / (DBWindowCount + 1)
    If DBWidth < 3000 Then DBWidth = 3000
    If DBWindowCount = 0 Then GoTo exithandler                      ' Nur ein fenster -> nix zu tun
    For i = 0 To DBWindowCount
        Set CurDBFrm = DBFormArray(i)
        CurDBFrm.WindowState = vbNormal                             ' Fall max or min erstmal auf Normal
        CurDBFrm.Top = 0
        CurDBFrm.Left = i * DBWidth
        CurDBFrm.Width = DBWidth
        CurDBFrm.Height = MainHeight
    Next i
    
exithandler:

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "MDIChildsVSplit", errNr, errDesc)
    Resume exithandler
End Sub

Public Function GetLastKey(szFullKey As String, szDelimiter As String) As String
    
    Dim szKeyArray() As String

On Error GoTo Errorhandler

    If szDelimiter = "" Then szDelimiter = "\"
    szKeyArray = Split(szFullKey, szDelimiter)
    GetLastKey = szKeyArray(UBound(szKeyArray))
    
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetLastKey", errNr, errDesc)
    Resume exithandler
End Function

'Public Function ListSubNodesByKey(TV As TreeView, szKey As String, _
'        Optional lngImageIndex As Integer, _
'        Optional bWithSubnodes As Boolean, Optional bExpand As Boolean)
'
'    Dim szSQL As String                                             ' SQL Statement
'    Dim rsList As New ADODB.Recordset                               ' Recordset für NodeList
'    Dim szTmpName As String                                         ' Node Bezeichnung
'    Dim szTmpID As String                                           ' Node ID (eindeutig)
'    Dim i As Integer                                                ' Counter
'    Dim cNode As node                                               '
'    Dim KeyArray() As String
'
'On Error GoTo Errorhandler
'
'    lngImageIndex = CLng(lngImageIndex)
'     'TV.Enabled = False
'
'    KeyArray = Split(szKey, TV_KEY_SEP)
'    ' SQL Statement holen
'    szSQL = objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_SQL, KeyArray(UBound(KeyArray)))
'    If szSQL = "" Then GoTo exithandler                             ' Kein SQL -> fertig
'    Set rsList = objDBconn.fillrs(szSQL)                            ' Recordet Füllen
'    Do While Not rsList.EOF                                         ' Alle DS durchgehen
'        szTmpID = Trim(objTools.checknull(rsList.Fields(0).Value, ""))  ' 1. Feldwert ID
'        szTmpName = Trim(objTools.checknull(rsList.Fields(1).Value, "")) ' 2. Feldwert Bezeichnung
'        If szTmpName <> "" And szTmpID <> "" Then
'
'            Call AddTreeNode(TV, szKey, szTmpID, szTmpName, lngImageIndex, False, , True)
'            'Call AddTreeNode(TV, szKey, szTmpID, szTmpName, lngImageIndex, bExpand, , True)
'            'Call ListSubNodesByRelation(TV, szKey & "\" & szTmpID, szKey, szTmpID, , bExpand)
'           Call ListSubNodesByRelation(TV, szKey & "\" & szTmpID, szKey, szTmpID, , False)
'            szTmpName = ""
'        End If
'        DoEvents
'        rsList.MoveNext
'    Loop
''     TV.Enabled = True
'    ' Sicherstellen das Parentnode nicht ausgefaltet
'''    Set cnode = GetNodeByKey(TV, szKey)     ' Parentnode ermitteln
'''    cnode.Expanded = False                  ' Einfalten
'
'exithandler:
'On Error Resume Next
'
'Exit Function
'Errorhandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.Errorhandler(MODULNAME, "ListSubNodesByKey", errNr, errDesc)
'    Resume exithandler
'End Function

'Public Function ListSubNodesByKey(TV As TreeView, szKey As String, _
'        Optional lngImageIndex As Integer, _
'        Optional bWithSubnodes As Boolean, Optional bExpand As Boolean)
'
'    Dim szSQL As String                 ' SQL Statement
'    Dim rsList As New ADODB.Recordset   ' Recordset für NodeList
'    Dim szTmpName As String             ' Node Bezeichnung
'    Dim szTmpID As String               ' Node ID (eindeutig)
'    Dim i As Integer                    ' Counter
'    Dim cnode As node                   '
'    Dim KeyArray() As String
'    Dim szTmp As String
'    Dim bSubNodes As Boolean
'
'On Error GoTo Errorhandler
'
'    lngImageIndex = CLng(lngImageIndex)
'
'    KeyArray = Split(szKey, "\")
'    ' SQL Statement holen
'    szSQL = objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_SQL, KeyArray(UBound(KeyArray)))
'    If szSQL = "" Then GoTo exithandler     ' Kein SQL -> fertig
'    Set rsList = objDBconn.fillrs(szSQL)    ' Recordet Füllen
'
'    Do While Not rsList.EOF                 ' Alle DS durchgehen
'        szTmpID = Trim(objTools.checknull(rsList.Fields(0).Value, ""))  ' 1. Feldwert ID
'        szTmpName = Trim(objTools.checknull(rsList.Fields(1).Value, "")) ' 2. Feldwert Bezeichnung
'        If szTmpName <> "" And szTmpID <> "" Then
'            Call AddTreeNode(TV, szKey, szTmpID, szTmpName, lngImageIndex, bExpand, , True)
'            Call ListSubNodesByRelation(TV, szKey, szTmpID, , bExpand)
'            szTmpName = ""
'        End If
'        DoEvents
'        rsList.MoveNext
'    Loop
'
'    ' Sicherstellen das Parentnode nicht ausgefaltet
'    Set cnode = GetNodeByKey(TV, szKey)     ' Parentnode ermitteln
'    cnode.Expanded = False                  ' Einfalten
'
'exithandler:
'On Error Resume Next
'
'Exit Function
'Errorhandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.Errorhandler(MODULNAME, "ListSubNodesByKey", errNr, errDesc)
'    Resume exithandler
'End Function

'Public Function ListSubNodesByFullKey(TV As TreeView, szFullKey As String, _
'        Optional lngImageIndex As Integer, _
'        Optional bWithSubnodes As Boolean)
'
'    Dim szSQL As String                 ' SQL Statement
'    Dim rsList As New ADODB.Recordset   ' Recordset für NodeList
'    Dim szTmpName As String             ' Node Bezeichnung
'    Dim szTmpID As String               ' Node Id (eindeutig)
'    Dim i As Integer                    ' Counter
'    Dim cNode As node                   '
'    Dim szKeyArray() As String
'
'On Error GoTo Errorhandler
'
'    ' Full Key analysieren
'    szKeyArray = Split(szFullKey, "\")
'
'    ' SQL Statement holen
'    szSQL = objTools.GetINIValue(App.Path & "\" & INI_FILENAME, "TreeView", szFullKey)
'    If szSQL = "" Then GoTo exithandler     ' Kein SQL -> fertig
'    Set rsList = objDBconn.fillrs(szSQL)    ' Recordet Füllen
'    Do While Not rsList.EOF                 ' Alle DS durchgehen
'        szTmpID = Trim(objTools.checknull(rsList.Fields(0).Value, ""))  ' 1. Feldwert ID
'        szTmpName = Trim(objTools.checknull(rsList.Fields(1).Value, "")) ' 2. Feldwert Bezeichnung
'
'        If szTmpName <> "" And szTmpID <> "" Then
'            Call AddTreeNode(TV, szFullKey, szTmpID, szTmpName, lngImageIndex)
'            Call ListSubNodesByRelation(TV, szFullKey, szTmpName)
'            szTmpName = ""
'        End If
'        DoEvents
'        rsList.MoveNext
'    Loop
'
'    ' Sicherstellen das Parentnode nicht ausgefaltet
'    Set cNode = GetNodeByKey(TV, szFullKey)     ' Parentnode ermitteln
'    cNode.Expanded = False                  ' Einfalten
'
'exithandler:
'On Error Resume Next
'
'Exit Function
'Errorhandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.Errorhandler(MODULNAME, "ListSubNodesByFullKey", errNr, errDesc)
'    Resume exithandler
'End Function



'Public Function ListLVByKey(LV As ListView, _
'        szKey As String, _
'        Optional szWhereKey As String, _
'        Optional bShowValueList As Boolean, _
'        Optional lngImageIndex As Integer)
'
'    Dim szSQL As String                 ' SQL Statement
'    Dim szTmpName As String             ' Item Text
'    Dim szTmpID As String               ' Item ID (eindeutig)
'
'
'On Error GoTo Errorhandler
'
'    Call ClearLV(LV)
'
'    LV.Tag = szKey
'    ' SQL Statement aus ini holen
'    szSQL = objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_SQL, szKey)
'    If szWhereKey <> "" Then
'        LV.Tag = LV.Tag & "\" & szWhereKey
'        ' Evtl. Where bedingung anhängen
'        szSQL = objSQLTools.AddWhereInFullSQL(szSQL, objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_SQL, "WHERE" & szKey) & "'" & szWhereKey & "'")
'    End If
'
'    If szSQL = "" Then GoTo exithandler         ' Kein SQL Statement -> Fertig
'
'    Call FillLVBySQL(LV, szSQL, objDBconn, bShowValueList, lngImageIndex)
'
'exithandler:
'On Error Resume Next
'
'Exit Function
'Errorhandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.Errorhandler(MODULNAME, "ListLVByKey", errNr, errDesc)
'End Function



'Public Function ListLVByTag(LV As ListView, DBCon As Object, _
'        szTag As String, _
'        Optional szWhereKey As String, _
'        Optional bShowValueList As Boolean, _
'        Optional lngImageIndex As Integer, _
'        Optional bNotColWidth As Boolean, _
'        Optional bNew As Boolean) As ADODB.Recordset
'' Füllt ListView indem anhand des Tag ein SQLStatement aus der INI gelesen Wird
'
'    Dim szTmpName As String                 ' Item Text
'    Dim szTmpID As String                   ' Item ID (eindeutig)
'    Dim szItemImage As String               ' Image index für subnode als String aus ini
'
'On Error GoTo Errorhandler
'
'    Dim LVInfo As ListViewInfo
'
'    Call ClearLV(LV)                        ' ListView Aufräumen
'
'    LV.Tag = szTag '& TV_KEY_SEP & "*"       ' Tag des Listviews setzen
'    ' SQL Statement aus ini holen
'    'szSQL = objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_SQLLIST, szTag)
'    ' aus XML
'    With LVInfo
'        Call objTools.GetLVInfoFromXML(App.Path & "\" & INI_XMLFILE, szTag, .szSQL, .szTag, .szWhere, .lngImage, .bValueList, .bListSubNodes)
'        If szWhereKey <> "" And .szSQL <> "" Then
'            ' Evtl. Where bedingung anhängen
'            'szWhere = objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_SQLLIST, "WHERE" & szTag)
'            If .szWhere <> "" Then .szSQL = objSQLTools.AddWhereInFullSQL(.szSQL, .szWhere & "'" & szWhereKey & "'")
'        End If
'
'        If .szSQL = "" Then GoTo exithandler     ' Kein SQL Statement -> Fertig
'
'        ' Evtl. Image index  aus Ini lesen
'        'szItemImage = objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_IMAGE, szTag & TV_KEY_SEP & "*")
'        lngImageIndex = .lngImage
'        'If .lngImage <> "" And IsNumeric(.lngImage) Then lngImageIndex = CLng(.lngImage)
'
'        Set ListLVByTag = FillLVBySQL(LV, .szSQL, DBCon, .bValueList, lngImageIndex, , , bNotColWidth, bNew)
''        If .bListSubNodes Then call fill
'
'        'End If
'     End With
'
'exithandler:
'On Error Resume Next
'
'Exit Function
'Errorhandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.Errorhandler(MODULNAME, "ListLVByTag", errNr, errDesc)
'End Function


'Public Function ListLVRelation(LV As ListView, FirstRelKey As String, _
'        SecRelKey As String, _
'        szWhere As String, _
'        Optional ImgIndex As Integer)
'
'    Dim szSQL As String         ' SQL Statement
'    Dim szIndexField As String
'
'On Error GoTo Errorhandler
'
'    ImgIndex = CLng(ImgIndex)
'    Call ClearLV(LV)
'
'    'lv.Tag = FirstRelKey & "\" & szWhere & "\" & SecRelKey
'    LV.Tag = FirstRelKey & "\" & SecRelKey      ' LV Tag festlegen
'
'    ' SQL Statement aus Ini datei
'    szSQL = objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_RELATIONS, FirstRelKey & SecRelKey)
'    szIndexField = objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_RELATIONS, FirstRelKey & SecRelKey & "INDEXFIELD")
'    If szSQL = "" Then GoTo exithandler         ' Kein SQL Statement -> Fertig
'    ' Where Statement aus INI Datei
'    szWhere = objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_SQL, "WHERE" & FirstRelKey) & "'" & szWhere & "'"
'    ' Where Statement zusammensetzen
'    If szWhere <> "" Then szSQL = objSQLTools.AddWhereInFullSQL(szSQL, szWhere)
'
'    Call FillLVBySQL(LV, szSQL, objDBconn, False, ImgIndex, szIndexField)       ' ListView Füllen
'
'exithandler:
'On Error Resume Next
'
'Exit Function
'Errorhandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.Errorhandler(MODULNAME, "ListLVRelation", errNr, errDesc)
'End Function
'Public Function ListSubNodesByRelation(TV As TreeView, ParentNode As node, FirstRelKey As String, szWhere As String, _
'        Optional lngImageIndex As Integer, Optional bExpand As Boolean)
'
'    Dim SecRelKey As String         ' 2. Teil der Relation
'    Dim szTmp As String
'    Dim i As Integer                'Counter
'    Dim bSubNodes As Boolean        ' Relation hat unterknoten
'    Dim szRelationName As String    ' Name der Relation
'    Dim szRelType As String
'
' On Error GoTo Errorhandler
'
'    lngImageIndex = CLng(lngImageIndex)
'
'    SecRelKey = "@"
'    While SecRelKey <> ""
'        ' Alle Ini Einträge unter "Relations" mit Firstrelkey + nummer
'        ' durchgehen und 2. teil der Relation ermitteln
'        SecRelKey = objTools.GetINIValue(App.Path & "\" _
'                & INI_FILENAME, INI_RELATIONS, FirstRelKey & CStr(i))
'        If SecRelKey <> "" Then                         ' Wenn 2. Teil gefunden
'            bSubNodes = True
'            szRelationName = FirstRelKey & SecRelKey    ' Relationsname zusammensetzen
'            szRelType = objTools.GetINIValue(App.Path & "\" _
'                    & INI_FILENAME, INI_RELATIONS, szRelationName & "TYP")
'            If szRelType = "n" Then
'                ' Ermitteln ob weitere unter knoten in dieser relation erwünscht
'                szTmp = objTools.GetINIValue(App.Path & "\" _
'                        & INI_FILENAME, INI_RELATIONS, szRelationName & "SUBNODES")
'                If szTmp <> "" Then bSubNodes = CBool(szTmp)
'                If bSubNodes Then
'                    If lngImageIndex = 0 Then
'                        ' Image index aus INI lesen
'                        szTmp = objTools.GetINIValue(App.Path & "\" & _
'                                    INI_FILENAME, INI_IMAGE, szRelationName)
'                        If szTmp <> "" And IsNumeric(szTmp) Then lngImageIndex = CLng(szTmp)
'                    End If
'                    Call AddTreeNode(TV, FirstRelKey & "\" & szWhere, SecRelKey, SecRelKey, lngImageIndex, bExpand)
'                End If ' bSubNodes
'                szRelType = ""
'            End If ' szRelType = "n"
'        End If  ' SecRelKey <> ""
'        lngImageIndex = 0
'        i = i + 1               'Counter hochzählen für nächste Relation
'    Wend
'
'exithandler:
'On Error Resume Next
'
'Exit Function
'Errorhandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.Errorhandler(MODULNAME, "ListSubNodesByRelation", errNr, errDesc)
'    Resume exithandler
'End Function

'Public Sub AddTreeNode(TV As TreeView, szKey As String, _
'        szNodeID As String, _
'        szNodeValue As String, _
'        Optional ingImgIndex As Integer, _
'        Optional bExpand As Boolean, _
'        Optional bRootNode As Boolean, _
'        Optional bWithoutNodeCheck As Boolean)
'' Legt einen Tree node an
'
'    Dim cNode As node
'    Dim szNodetext, sznodeKey As String                             ' Koten text
'    Dim newNode As node                                             ' Neuer Konten
'    Dim SubNodeName As String                                       ' Unterknoten name
'    Dim i As Integer                                                ' Counter
'    Dim szIndex As String                                           ' Counter als 2 stelliger String
'    Dim szTmpImageIndex As String                                   ' Evtl. Image index
'
'On Error GoTo Errorhandler
'
'     If Not bWithoutNodeCheck Then                                  ' soll evtl. vohanden sein des Nodes geprüft werden
'        Set cNode = GetNodeByKey(TV, szKey & "\" & szNodeID)        ' Node anhand von Key ermitteln
'        If Not cNode Is Nothing Then                                ' Node schon vorhanden
'            Call SelectTreeNode(TV, cNode)                          ' Node auswählen
'            GoTo exithandler                                        ' -> Raus
'        End If
'    End If
'
'    i = 1
'    szNodetext = szNodeValue                                        ' Node Bez setzen
'
'    If bRootNode Then                                               ' Ist Wurzel (Root) Node
'        sznodeKey = szNodeID                                        ' Neuen node Key generieren
'        Set newNode = TV.Nodes.Add(, , sznodeKey, szNodetext)       ' Node anlegen
'    Else
'        sznodeKey = szKey & "\" & szNodeID                          ' Neuen node Key generieren (mit evtl. DS ID)
'        Set newNode = TV.Nodes.Add(szKey, tvwChild, sznodeKey, szNodetext) ' Node anlegen
'    End If
'
'    newNode.Image = ingImgIndex                                     ' Image setzen
'    'newNode.Expanded = bExpand                                     ' Knoten ausfalten
'    SubNodeName = "@@"                                              ' Subnodename setzen als Abbruch bed.
'    While SubNodeName <> "" And bRootNode
'        szIndex = CStr(i)
'        If Len(szIndex) = 1 Then szIndex = "0" & szIndex
'        SubNodeName = objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_SUBNODES, szNodeID & szIndex)
'        If SubNodeName <> "" Then                                   ' Wenn Subnode gefunden
'            If Left(SubNodeName, 6) = "SELECT" Then
'                SubNodeName = objDBconn.GetValueFromSQL(SubNodeName)
'                If SubNodeName <> "" Then
'                    Call AddTreeNode(TV, sznodeKey, SubNodeName, SubNodeName, ingImgIndex, bExpand, False)
'                End If
'            Else
'                szTmpImageIndex = objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_IMAGE, SubNodeName)
'                If szTmpImageIndex = "" Then szTmpImageIndex = ingImgIndex
'                Call AddTreeNode(TV, sznodeKey, SubNodeName, SubNodeName, CLng(szTmpImageIndex), bExpand, False)
'            End If
'        End If
'        i = i + 1
'    Wend
'
'    If Not newNode.Parent Is Nothing Then                           ' Vater Node vorhanden
'        newNode.Parent.Expanded = bExpand                           ' Vater Knoten ausfalten
'    End If
'
'exithandler:
'On Error Resume Next
'
'Exit Sub
'Errorhandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.Errorhandler(MODULNAME, "AddTreeNode", errNr, errDesc)
'    Resume exithandler
'End Sub
'Public Function ListSubNodesByRelation(TV As TreeView, ParentNodeKey As String, FirstRelKey As String, szWhere As String, _
'        Optional lngImageIndex As Integer, Optional bExpand As Boolean)
'
'    Dim SecRelKey As String                                         ' 2. Teil der Relation
'    Dim szTmp As String
'    Dim i As Integer                                                'Counter
'    Dim bSubNodes As Boolean                                        ' Relation hat unterknoten
'    Dim szRelationName As String                                    ' Name der Relation
'    Dim szRelType As String
'    Dim bWithRelNode As Boolean
'
' On Error GoTo Errorhandler
'
'    lngImageIndex = CLng(lngImageIndex)
'
'    SecRelKey = "@"
'    While SecRelKey <> ""
'        ' Alle Ini Einträge unter "Relations" mit Firstrelkey + nummer
'        ' durchgehen und 2. teil der Relation ermitteln
'        SecRelKey = objTools.GetINIValue(App.Path & "\" _
'                & INI_FILENAME, INI_RELATIONS, FirstRelKey & CStr(i))
'        If SecRelKey <> "" Then                                     ' Wenn 2. Teil gefunden
'            bSubNodes = True
'            szRelationName = FirstRelKey & SecRelKey                ' Relationsname zusammensetzen
'            szRelType = objTools.GetINIValue(App.Path & "\" _
'                    & INI_FILENAME, INI_RELATIONS, szRelationName & "TYP")
'            If szRelType = "n" Then
'                ' Ermitteln ob weitere unter knoten in dieser relation erwünscht
'                szTmp = objTools.GetINIValue(App.Path & "\" _
'                        & INI_FILENAME, INI_RELATIONS, szRelationName & "SUBNODES")
'                If szTmp <> "" Then bSubNodes = CBool(szTmp)
'                szTmp = ""
'                If bSubNodes Then
'                    If lngImageIndex = 0 Then
'                        ' Image index aus INI lesen
'                        szTmp = objTools.GetINIValue(App.Path & "\" & _
'                                    INI_FILENAME, INI_IMAGE, szRelationName)
'                        If szTmp <> "" And IsNumeric(szTmp) Then lngImageIndex = CLng(szTmp)
'                        szTmp = ""
'                    End If
''                    szTmp = objTools.GetINIValue(App.Path & "\" & _
'                                    INI_FILENAME, INI_RELATIONS, szRelationName & "WITHRELNODE")
''                    If szTmp <> "" Then bWithRelNode = CBool(szTmp)
''                    szTmp = ""
''                    If bWithRelNode Then
'                        Call AddTreeNode(TV, ParentNodeKey, SecRelKey, SecRelKey, lngImageIndex, bExpand)
''                    Else
''                        Call ListSubNodesByKey(TV, SecRelKey, lngImageIndex, True)
''                    End If
'                End If ' bSubNodes
'                szRelType = ""
'            End If ' szRelType = "n"
'        End If  ' SecRelKey <> ""
'        lngImageIndex = 0
'        i = i + 1                                                   'Counter hochzählen für nächste Relation
'    Wend
'
'exithandler:
'On Error Resume Next
'
'Exit Function
'Errorhandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.Errorhandler(MODULNAME, "ListSubNodesByRelation", errNr, errDesc)
'    Resume exithandler
'End Function

