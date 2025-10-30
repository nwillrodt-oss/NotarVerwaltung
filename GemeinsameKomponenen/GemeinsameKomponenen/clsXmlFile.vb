Imports System.Xml                                                              ' XML Klasse Importieren (Spart schreibarbeit)

Public Class clsXmlFile
    Private Const MODULNAME = "clsXmlFile"                                      ' Modulname für Fehlerbehandlung
    Private bInitOK As Boolean                                                  ' Gibt an das die Klasse erfolgreich initialisiert wurde

    Private ObjBag As clsObjectBag                                              ' Sammelobject
    Private ConfigXML As Xml.XmlDocument                                        ' Akt Config XML File
    Private Const XML_PATH_SEP = "\"
    Private _getChildNodeNameList As String

#Region "Constructor"

    Public Sub New(ByVal oBag As clsObjectBag, ByVal xmlFilePath As String)
        Try                                                                     ' Fehlerbehandlung aktivieren
            ObjBag = oBag                                                       ' ObjectBag Übergeben
            bInitOK = LoadXML(xmlFilePath)
        Catch ex As Exception                                                   ' Fehler behandeln

        End Try
    End Sub

    Public Sub New(ByVal oBag As clsObjectBag, ByVal xmlDoc As XmlDocument)
        Try                                                                     ' Fehlerbehandlung aktivieren
            ObjBag = oBag                                                       ' ObjectBag Übergeben

        Catch ex As Exception                                                   ' Fehler behandeln

        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#End Region

#Region "Properties"

    Public ReadOnly Property InitOK As Boolean                                  ' Gibt zurück ob die initialieierung Fehlerfrei war
        Get
            Return bInitOK
        End Get
    End Property

    Public ReadOnly Property XMLDoc() As System.Xml.XmlDocument                 ' Gibt das Config XML Zurück
        Get
            XMLDoc = ConfigXML                                                  ' Und zurück geben
        End Get
    End Property

    Public ReadOnly Property RootElement As XmlElement
        Get
            RootElement = ConfigXML.DocumentElement                             ' Wurzelknoten ermittelm
        End Get
    End Property

#End Region

    Private Function LoadXML(ByVal xmlFilePath As String) As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            If xmlFilePath = "" Then Return False ' Kein Pfad -> fertig
            ConfigXML = New Xml.XmlDocument                                     ' Neues Object Createn
            ConfigXML.Load(xmlFilePath)                                         ' XML Datei als XMLDocument laden
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "LoadXML", ex)                  ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Public Function GetChildNode(ByVal szParentNodePath As String) As XmlElement
        Dim TreeNodeArray() As String
        Dim ConfigRootNode As XmlElement                                        ' Options Wurzelknoten
        Dim TreeNode As XmlElement                                              ' Tree knoten
        Dim szResult As String = ""
        Try                                                                     ' Fehlerbehandlung aktivieren
            If szParentNodePath = "" Then Return Nothing
            TreeNodeArray = Split(szParentNodePath, XML_PATH_SEP)               ' NodePath in array aufspalten
            ConfigRootNode = ConfigXML.DocumentElement                          ' Wurzelknoten ermittelm
            If ConfigRootNode.Name = TreeNodeArray(0) Then
                TreeNode = GetChildNode(ConfigRootNode, TreeNodeArray(1))       ' 1. Unterknoten von Wurzel knoten holen
                For i = 2 To TreeNodeArray.Length - 1                           ' Dann Path Array durchlaufen
                    TreeNode = GetChildNode(TreeNode, TreeNodeArray(i))         ' und unterknoten holen
                Next
                Return TreeNode                                                 ' Ergebnis zurück
            End If
            Return Nothing
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "GetChildNode", ex)             ' Fehlermeldung ausgeben
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

    Public Function GetChildNode(ByVal cXMLParentNode As XmlElement, _
                                ByVal szChildNodeName As String, _
                                Optional ByVal szAttributName As String = "", _
                                Optional ByVal szAttribuValue As String = "") As XmlElement
        ' liefert Child node mit szAttributName = szAttribuValue zurück
        Dim cXMLNode As XmlElement                                              ' Aktueller XML Node
        Dim XMLNodeAtribute As XmlAttribute                                     ' XML Node Atribut
        Try                                                                     ' Fehler behandlung aktivieren
            For Each cXMLNode In cXMLParentNode.ChildNodes                      ' Alle Child Nodes duchlaufen
                If szChildNodeName = "" Or cXMLNode.Name = szChildNodeName Then ' Nur TreeNode unterknoten berücksichtigen
                    If szAttributName <> "" Then                                '  evtl. Attribut prüfen
                        For Each XMLNodeAtribute In cXMLNode.Attributes
                            If XMLNodeAtribute.Name = szAttributName Then       ' Attribut Tag suchen
                                If szAttribuValue = "" Or XMLNodeAtribute.Value = szAttribuValue Then
                                    Return cXMLNode                             ' Knoten gefunden
                                End If
                            End If  ' XMLNodeAtribute.Name = szAttributName
                        Next                                                    ' Nächstes Attribut
                    Else                                                        ' Sonst
                        Return cXMLNode                                         ' Knoten gefunden
                    End If ' szAttributName <> ""
                End If ' cXMLNode.baseName = szChildNodeName
            Next
            Return Nothing
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "GetChildNode", ex)             ' Fehlermeldung ausgeben
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

    Public Function GetChildNodeNameList(ByVal szParentNodePath As String, _
                                        Optional ByVal szDelemiter As String = ";") As String
        Dim TreeNodeArray() As String
        Dim ConfigRootNode As XmlElement                                        ' Options Wurzelknoten
        Dim TreeNode As XmlElement                                              ' Tree knoten
        Dim szResult As String = ""
        Try                                                                     ' Fehlerbehandlung aktivieren
            If szParentNodePath = "" Then Return ""
            TreeNodeArray = Split(szParentNodePath, XML_PATH_SEP)               ' NodePath in array aufspalten
            ConfigRootNode = ConfigXML.DocumentElement                          ' Wurzelknoten ermittelm
            If ConfigRootNode.Name = TreeNodeArray(0) Then
                TreeNode = GetChildNode(ConfigRootNode, TreeNodeArray(1))       ' 1. Unterknoten von Wurzel knoten holen
                For i = 2 To TreeNodeArray.Length - 1                           ' Dann Path Array durchlaufen
                    TreeNode = GetChildNode(TreeNode, TreeNodeArray(i))         ' und unterknoten holen
                Next
                If IsNothing(TreeNode) Then Return ""
                For i = 0 To TreeNode.ChildNodes.Count
                    szResult = szResult & ReadAttribute(TreeNode.ChildNodes(i), "Name") & szDelemiter
                Next
                Return szResult                                                 ' Ergebnis zurück
            End If
            Return ""
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "GetChildNodeNameList", ex)     ' Fehlermeldung ausgeben
            Return ""                                                           ' Misserfolg zurück
        End Try
    End Function

    Public Function GetChildNodeNameArray(ByVal szParentNodePath As String) As String()
        Try                                                                     ' Fehlerbehandlung aktivieren
            Return Split(GetChildNodeNameList(szParentNodePath, ";"), ";")
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "GetChildNodeNameList", ex)     ' Fehlermeldung ausgeben
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

    Private Function ReadAttribute(ByVal xmlE As XmlElement, _
                                   ByVal szAttributname As String) As String
        Try                                                                     ' Fehlerbehandlung aktivieren
            If IsNothing(xmlE) Then Return "" ' Kein xml element -> Fertig
            If szAttributname = "" Then Return "" ' Kein Attributname angegeben -> Fertig
            If xmlE.HasAttributes Then                                          ' Wenn Attribute vorhanden
                For Each Attr As XmlAttribute In xmlE.Attributes                ' Alle Attribute durchlaufen
                    If Attr.Name.ToUpper = szAttributname.ToUpper Then          ' Gesuchter Attribut name dabei
                        Return Attr.Value                                       ' WErt zurück
                    End If
                Next                                                            ' Nächstes Attribut
            End If
            Return ""                                                           ' Nichts gefunden
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "ReadAttribute", ex)            ' Fehlermeldung ausgeben
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

    Public Function GetXmLNode(ByVal XmlStartElement As XmlElement, ByVal szElementName As String, _
                              Optional ByVal szAttrName As String = "", _
                              Optional ByVal szAttrValue As String = "") As XmlElement
        Dim xmlFound As XmlElement
        Try                                                                     ' Fehler behandlung aktivieren
            If XmlStartElement.Name = szElementName Then                        ' Gesuchter elementname
                If szAttrName <> "" Then                                        ' Attribut name ist angegeben
                    ' Erst Attribute dieses Elements Durchsuchen
                    If XmlStartElement.Attributes.Count > 0 Then                ' Wenn Attributte vorhanden
                        If szAttrValue = "" Then                                ' Attribut value ist nicht  angegeben
                            If Not IsNothing(XmlStartElement.Attributes(szAttrName)) Then ' wenn diese Attribut aber vorhanden
                                Return XmlStartElement                          ' Erfolg zurück
                            End If
                        Else                                                    ' Sonst (Attribut value ist angegeben)
                            If XmlStartElement.GetAttribute(szAttrName).ToUpper = szAttrValue.ToUpper Then ' Vergleichen
                                Return XmlStartElement                          ' Erfolg zurück
                            End If
                        End If
                    End If
                Else                                                            ' Sonst (Attribut name ist nicht angegeben)
                    Return XmlStartElement                                      ' Erfolg zurück
                End If
            End If
            ' Dann Recursiv Kindknoten durchsuchen
            If XmlStartElement.HasChildNodes Then                               ' Wenn Kindknoten vorhanden
                For Each cNode In XmlStartElement.ChildNodes                    ' Alle Kindknoten  durchlaufen
                    'If cNode.name = szElementName Then                          ' Wenn Kindknoten gesuchter elementname
                    xmlFound = GetXmLNode(cNode, szElementName, szAttrName, szAttrValue) ' Recursiv aufrufen
                    If Not xmlFound Is Nothing Then Return xmlFound 'Wenn Erfolg bei ReKursion -> erfolg zurück geben
                    'End If
                Next                                                            ' Nächster Kindknoten
            End If
            Return Nothing                                                      ' wenn wir hier landen -> Misserfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "GetXMLNode", ex)               ' Fehlermeldung ausgeben
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

    Public Function GetXmLNode(ByVal StartElementPath As String, ByVal szElementName As String, _
                              Optional ByVal szAttrName As String = "", _
                              Optional ByVal szAttrValue As String = "") As XmlElement
        Dim TreeNodeArray() As String                                           ' Pfad Array
        Dim xmlStartelement As XmlElement                                       ' Start xml element
        Try                                                                     ' Fehler behandlung aktivieren
            If StartElementPath = "" Then Return Nothing
            TreeNodeArray = Split(StartElementPath, XML_PATH_SEP)               ' NodePath in array aufspalten
            xmlStartelement = ConfigXML.DocumentElement                         ' Wurzelknoten ermittelm
            If IsNothing(xmlStartelement) Then Return Nothing ' Wenn kein Wurzelknoten vorhanden -> Fertig
            If xmlStartelement.Name.ToUpper = TreeNodeArray(0).ToUpper Then     ' Wurzelknoten sollte im StartElementPath an 1. stelle stehen
                Return GetXmLNode(xmlStartelement, szElementName, szAttrName, szAttrValue)  ' Recursiv weitermachen
            End If
            Return Nothing                                                      ' wenn wir hier landen -> Misserfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "GetXMLNode", ex)               ' Fehlermeldung ausgeben
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

    'Public Function GetRootNodeListFromXML() As String
    '    Dim ConfigRootNode As XmlElement                                        ' Options Wurzelknoten
    '    Dim TreeNode As XmlElement                                              ' Tree knoten
    '    Dim i As Integer                                                        ' Counter für die Kategorieknoten
    '    Dim result As String = ""                                               ' Ergebnis String
    '    Try                                                                     ' Fehlerbehandlung aktivieren
    '        If LoadConfigXML() Then
    '            ConfigRootNode = ConfigXML.DocumentElement                      ' Wurzelknoten ermittelm
    '            If ConfigRootNode.HasChildNodes Then
    '                TreeNode = ConfigRootNode.ChildNodes(0)                     ' Tree Node auswählen
    '                For i = 0 To TreeNode.ChildNodes.Count - 1                  ' Alle unterkonoten durchlaufen
    '                    result = result & TreeNode.ChildNodes(0).Attributes("Tag").Value.ToString & ";"
    '                Next i
    '            End If
    '            Return result                                                   ' Erfolg zurück
    '        End If
    '        Return ""                                                           ' Misserfolg zurück
    '    Catch ex As Exception                                                   ' Fehler behandeln
    '        Call ErrorHandler(MODULNAME, "GetRootNodeListFromXML", ex)          ' Fehlermeldung ausgeben
    '        Return ""                                                           ' Misserfolg zurück
    '    End Try
    'End Function

    'Public Function GetRootNodeArrayFromXML() As String()
    '    Dim NodeList As String
    '    Dim result() As String                                                  ' Ergebnis Array
    '    Try                                                                     ' Fehlerbehandlung aktivieren
    '        NodeList = GetRootNodeListFromXML()                                 ' RootNodes Als StringListe holen
    '        If NodeList <> "" Then                                              ' Wenn Die lise nicht leer ist
    '            result = Split(NodeList, ";")                                   ' In Array Aufspalten
    '            Return result                                                   ' Erfolg zurück
    '        End If
    '        Return Nothing                                                      ' Misserfolg zurück
    '    Catch ex As Exception                                                   ' Fehler behandeln
    '        Call ErrorHandler(MODULNAME, "GetRootNodeListFromXML", ex)          ' Fehlermeldung ausgeben
    '        Return Nothing                                                      ' Misserfolg zurück
    '    End Try
    'End Function

    'Public Function GetLVInfoFromXML(ByVal Nodefullpath As String) As LVListInfo
    '    Dim newXMLNode As IXMLDOMNode
    '    Dim cXMLNode As IXMLDOMNode
    '    Dim XMLNodeAtribute As IXMLDOMAttribute
    '    Dim szTmp As String
    '    Try                                                                 ' Fehlerbehandlung aktivieren
    '        newXMLNode = GetTreeNodeFromXML(Nodefullpath)
    '        If newXMLNode Is Nothing Then GoTo exithandler
    '        For Each cXMLNode In newXMLNode.childNodes                      ' Alle root Nodes >TreeNode> duchlaufen
    '            If cXMLNode.baseName = "List" Then
    '                For Each XMLNodeAtribute In cXMLNode.Attributes
    '                    Select Case XMLNodeAtribute.Name
    '                        Case "Tag"
    '                            newTag = XMLNodeAtribute.Value
    '                        Case "Image"
    '                            newImage = CheckXMLValueForNumeric(XMLNodeAtribute.Value, 1) ' Attributvalue auf Numerisch prüfen
    '                        Case "AltImage"                                     ' Attibut AltImage auslesen
    '                            lngAltImage = CheckXMLValueForNumeric(XMLNodeAtribute.Value, 1) ' Attributvalue auf Numerisch prüfen
    '                        Case "AltImageField"                                ' Attibut AltImageField auslesen
    '                            szAktImgField = XMLNodeAtribute.Value
    '                        Case "AltImageValue"                                ' Attibut AltImageValue auslesen
    '                            szAltImgValue = XMLNodeAtribute.Value
    '                        Case "bValueList"
    '                            bValueList = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
    '                        Case "DelFlag"
    '                            szDelFlag = XMLNodeAtribute.Value
    '                        Case "SQL"
    '                            szSQL = XMLNodeAtribute.Value
    '                        Case "WHERE"
    '                            szWhere = XMLNodeAtribute.Value
    '                        Case "bListSubNodes"
    '                            bSubNodes = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
    '                        Case "Edit"
    '                            bEdit = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
    '                        Case "New"
    '                            bNew = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
    '                        Case "SelectNode"
    '                            bSelectNode = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
    '                        Case "bShowKontextMenue"
    '                            bShowkontextMenue = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
    '                        Case "Delete"
    '                            bDelete = CheckXMLValueForBool(XMLNodeAtribute.Value, False) ' Attributvalue auf Bool prüfen
    '                        Case Else

    '                    End Select
    '                    szTmp = ""
    '                Next
    '            End If
    '        Next
    '    Catch ex As Exception                                                   ' Fehler behandeln
    '        Call ErrorHandler(MODULNAME, "GetLVInfoFromXML", ex)          ' Fehlermeldung ausgeben
    '        Return Nothing                                                      ' Misserfolg zurück
    '    End Try
    'End Function

End Class
