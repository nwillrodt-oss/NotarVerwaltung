Imports System.Xml                                                              ' XML Klasse Importieren (Spart schreibarbeit)

Public Class clsTreeNodeEx
    Inherits TreeNode                                                           ' Org. TreeNode beerben

    Private Const MODULNAME = "clsTreeNodeEx"                                   ' Modulname für Fehlerbehandlung

    Private objBag As clsObjectBag                                              ' Sammelobject
    'Private ConfigXML As Xml.XmlDocument                                        ' Config XML File
    Private objConfigXML As clsXmlFile
    Public ExTVNodeInfo As TVNodeInfoEx
    Public ExListInfo As ListViewInfoEx

    Private Const XML_PATH_SEP = "\"

    Public Structure TVNodeInfoEx                                               ' TV Konoten informationen
        Public ID As String
        Public Desc As String                                                   ' Angezeigte Beschreibung
        Public ImageIndex As Integer
        Public SQL As String                                                    ' zugrundeliegendes SQL Statement
        Public WHERE As String                                                  ' evtl. Where Statement
        'Public XML As String                                                    ' Entsprechender XML Node als String
        Public XmlDoc As XmlDocument                                            ' Entsprechender XML Node als XML DOc
        Public Typ As String                                                    ' Statisch oder dynamisch
        Public bShowSubnodes As Boolean                                         ' Unterknoten sofort anzeigen
        Public bShowKontextMenue As Boolean
        Public ChildnodeList As String
    End Structure

    Public Structure ListViewInfoEx                                             ' ListView Informationen
        Public SQL As String                                                    ' zugrundeliegendes SQL Statement
        Public XmlDoc As XmlDocument                                            ' Entsprechender XML Node als XML DOc
        Public EditFormName As String
        'Public szTag As String                                                  ' Tag des ListViews (welche Daten werden angezeigt)
        Public bValueList As Boolean                                            ' Darstellung als Valuelist (1.DS pro Wert ein Item)
        Public DelFlagField As String                                           ' Feld in dem ein gelöscht flag gesetzt werden kann
        Public WhereNoDel As String                                             ' Where Part mit gelöschten DS (Flag)
        Public WHERE As String                                                  ' evtl. Where Statement
        Public ImageIndex As Integer                                            ' Image Index für Item
        Public AltImage As Integer                                              ' Alternatives Image
        Public AltImgField As String                                            ' Feld das für alt Image geprüft wird
        Public AltImgValue As String                                            ' Value der für alt image geprüft wird
        Public bListSubNodes As Boolean                                         ' Sollen Subnodes im Liszview mitangezeigt werden
        Public bEdit As Boolean                                                 ' Einträge können bearbeitet werden
        Public bNew As Boolean                                                  ' Es können neue einträge angelegt werden
        Public bDelete As Boolean                                               ' Es dürfen Einträge gelöscht werden
        Public bSelectNode As Boolean                                           ' dbl klick selectet den entsprechenden Node
        Public bShowKontextMenue As Boolean                                     ' Kontextmenü zulässig
    End Structure

#Region "Constructor"

    Public Sub New(ByVal oBag As clsObjectBag, ByVal xmlNodePath As String, _
                   Optional ByVal Parent As TreeNode = Nothing)
        Try                                                                     ' Fehlerbehandlung aktivieren
            ObjBag = oBag                                                       ' ObjectBag Übergeben
            'ConfigXML = ObjBag.ConfigXMLDoc                                     ' ConfigDatei laden
            objConfigXMl = objBag.ConfigXML
            Call InitNodeInfo(xmlNodePath, Parent)                              ' Extendet infos zum Node laden
        Catch ex As Exception                                                   ' Fehler behandeln

        End Try
    End Sub

    Public Sub New(ByVal oBag As clsObjectBag, ByVal xmlNodePath As String, _
                   ByVal ID As String, _
                   ByVal Nodetext As String,
                   Optional ByVal Parent As TreeNode = Nothing)
        Try                                                                     ' Fehlerbehandlung aktivieren
            ObjBag = oBag                                                       ' ObjectBag Übergeben
            'ConfigXML = ObjBag.ConfigXMLDoc                                     ' ConfigDatei laden
            objConfigXMl = objBag.ConfigXML
            Call InitNodeInfo(xmlNodePath, Parent, ID, Nodetext)                        ' Extendet infos zum Node laden
        Catch ex As Exception                                                   ' Fehler behandeln

        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#End Region

#Region "Properties"

    Public ReadOnly Property ID As String
        Get
            ID = Replace(ExTVNodeInfo.ID, "ID:", "")
        End Get
    End Property

#End Region

    Private Function InitNodeInfo(ByVal xmlNodename As String, _
                                  Optional ByVal ParentNode As clsTreeNodeEx = Nothing, _
                                  Optional ByVal ID As String = "",
                                  Optional ByVal NodeText As String = "") As Boolean
        Dim xmlE As XmlElement                                                  ' Aktueller TreeNode
        Dim NewID As String
        Try                                                                     ' Fehlerbehandlung aktivieren
            ExTVNodeInfo = New TVNodeInfoEx                                     ' Zusatzinfo initialisieren
            If GetTreeNodeXML(xmlNodename) Then                                 ' entsprechendem XML part finden
                With ExTVNodeInfo
                    xmlE = .XmlDoc.DocumentElement                              ' XML Doc Des Nodes Einlesen
                    .Typ = xmlE.GetAttribute("Typ")                             ' Typ Setzen
                    .bShowSubnodes = xmlE.GetAttribute("bShowSubnodes")         ' SubNodes anzeigen setzen
                    .bShowKontextMenue = xmlE.GetAttribute("bShowKontextMenue") ' Kontextmenü anzeigen setzen
                    .Desc = xmlE.GetAttribute("Description")                    ' Beschreibung setzen
                    .SQL = xmlE.GetAttribute("SQL")                             ' SQL Statement setzen
                    .WHERE = xmlE.GetAttribute("WHERE")                         ' Optionales Where setzen
                    .ImageIndex = xmlE.GetAttribute("Imageindex")
                    If NodeText <> "" Then
                        Me.Text = NodeText                                      ' Text aus Parameter
                    Else
                        Me.Text = xmlE.GetAttribute("Text")                     ' Text aus XML Attribut Setzen
                    End If
                    If ID = "" Then
                        NewID = Me.Text
                    Else
                        NewID = "ID:" & ID
                        .ID = NewID
                    End If
                    Me.ImageIndex = .ImageIndex
                    Me.SelectedImageIndex = .ImageIndex

                    Call InitListInfo(xmlE, xmlNodename)
                    If Not ParentNode Is Nothing Then
                        Me.Name = ParentNode.Name & "\" & NewID
                    Else
                        Me.Name = NewID
                    End If

                    If .bShowSubnodes Then                                      ' Wenn Subnodes anzeigen
                        Call AddSubNodes()                                      ' SUbNodes Anhängen
                    End If

                End With
            End If
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "InitNodeInfo", ex)             ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function InitListInfo(ByVal xmlE As XmlElement, ByVal xmlNodename As String) As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            ExListInfo = New ListViewInfoEx                                     ' Zusatzinfo initialisieren
            xmlE = objConfigXML.GetXmLNode(xmlE, "List", "Name", xmlNodename)   ' entsprechendem XML part finden
            If Not IsNothing(xmlE) Then                                         ' Wenn xml Element vorhanden
                With ExListInfo
                    .bValueList = xmlE.GetAttribute("bValueList")               ' WerteListe ?
                    .SQL = xmlE.GetAttribute("SQL")                             ' SQL Statement setzen
                    .WHERE = xmlE.GetAttribute("WHERE")                         ' Optionales Where setzen
                    .ImageIndex = xmlE.GetAttribute("Imageindex")               ' Imageindex für ListView Item
                    .bListSubNodes = xmlE.GetAttribute("bListSubNodes")         ' Mögliche SubNodes (tree) in LV mitanzeigen
                    .EditFormName = xmlE.GetAttribute("EditName")
                    .bEdit = xmlE.GetAttribute("Edit")
                End With
                Return True                                                     ' Erfolg zurück
            Else
                Return False                                                    ' Misserfolg zurück
            End If

        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "InitListInfo", ex)             ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Public Function AddSubNodes() As Boolean
        Dim subnode As Xml.XmlElement
        Try                                                                     ' Fehlerbehandlung aktivieren
            Dim xmlE As Xml.XmlElement = ExTVNodeInfo.XmlDoc.DocumentElement
            For Each subnode In xmlE.ChildNodes
                If subnode.Name = "Treenode" Then
                    If subnode.Attributes("Typ").Value = "Static" Then          ' Statischer knoten (1ner)
                        Dim stvn As New clsTreeNodeEx(ObjBag, subnode.Attributes("Name").Value, Me)
                        If Not CheckChildNodeExist(stvn) Then                   ' Prüfen ob wir den Knoten schon hatten
                            Me.Nodes.Add(stvn)                                  ' und anhängen
                            Me.Expand()
                        End If
                    Else                                                        ' Dynamische unterknoten                                            
                        Call AddTreeNodeFromSQL(subnode.Attributes("Name").Value)
                    End If

                End If
            Next
            Return True
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "InitNodeInfo", ex)             ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function AddTreeNodeFromSQL(ByVal xmlNodename As String) As Boolean
        ' Left Tree nodes aus SQL statement an
        Dim DSNode As DataSet                                                   ' DS mit Node Daten
        Dim dTable As DataTable                                                 ' Datatable mit Node Daten
        Dim szSQL As String                                                     ' SQL Statement
        Dim szKey As String                                                     ' Neuer Node Key (ID)
        Dim szTag As String                                             ' Neuer Node Tag
        Dim szText As String                                                    ' Neuer Node Text
        Try                                                                     ' Fehlerbehandlung aktivieren
            Dim xmlE As Xml.XmlElement = ExTVNodeInfo.XmlDoc.DocumentElement
            xmlE = objConfigXML.GetXmLNode(xmlE, "Treenode", "Name", xmlNodename) ' XMlElement für unterknoten holen
            If IsNothing(xmlE) Then Return False ' Kein xmlElement -> Raus
            szSQL = xmlE.GetAttribute("SQL")                                    ' SQL Statement auslesen
            If szSQL = "" Then Return False ' Kein SQL -> Fertig
            DSNode = ObjBag.ObjDBConnect.FillDS(szSQL)                          ' DataSet füllen
            If IsNothing(DSNode) Then Return False ' Keine Daten -> Fertig
            If DSNode.Tables.Count = 0 Then Return False ' Keine Daten -> Fertig
            dTable = DSNode.Tables(0)                                           ' Datentabele holen
            For Each dRow As DataRow In dTable.Rows                             ' Alle Datensätze durchlaufen
                szText = dRow.Item("nodetext").ToString                         ' NodeText aus Datasaet
                szKey = dRow.Item("nodekey").ToString                           ' NodeKey aus DataSet (ist die ID)
                szTag = dRow.Item("nodetag").ToString                           ' NodeTag aus DataSet
                If szText <> "" And szKey <> "" And szTag <> "" Then            ' Haben wir alles beisammen
                    Dim stvn As New clsTreeNodeEx(ObjBag, xmlNodename, szKey, szText, Me) ' Neuen Treenodee Konstruieren
                    If Not CheckChildNodeExist(stvn) Then                       ' Prüfen ob wir den Knoten schon hatten
                        Me.Nodes.Add(stvn)                                      ' und anhängen
                        Me.Expand()
                    End If
                End If
            Next                                                                ' Nächster DS
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "AddTreeNodeFromSQL", ex)       ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function CheckChildNodeExist(ByVal NewNode As clsTreeNodeEx) As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            For Each cNode In Me.Nodes                                          ' Alle Vorhandenen Kindknoten durchlaufen
                If cNode.name.toupper = NewNode.Name.ToUpper Then
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "CheckChildNodeExist", ex)      ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function GetTreeNodeXML(ByVal xmlNodeName As String) As Boolean
        Dim xmlE As XmlElement
        Try                                                                     ' Fehlerbehandlung aktivieren
            'Dim ConfigRootNode As XmlElement = ConfigXML.DocumentElement        ' Wurzelknoten ermittelm
            'If ConfigRootNode.HasChildNodes Then                                ' Wenn Kindknoten vorhanden
            Dim TreeNode As XmlElement = objConfigXML.RootElement               ' Tree Node auswählen
            If TreeNode.HasChildNodes Then                                      ' Wenn Kindknoten vorhanden
                xmlE = objConfigXML.GetXmLNode(TreeNode, "Treenode", "Name", xmlNodeName)
                If IsNothing(xmlE) Then Return False ' Kein knoten gefunden -> Fertig
                'ExTVNodeInfo.XML = xmlE.OuterXml                                ' OuterXML Merken
                ExTVNodeInfo.XmlDoc = New Xml.XmlDocument
                ExTVNodeInfo.XmlDoc.LoadXml(xmlE.OuterXml)
                Return True                                                     ' Erfolg zurück
            End If
            'End If
            Return False                                                        ' Misserfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call objBag.ErrorHandler(MODULNAME, "GetTreeNodeXML", ex)           ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

End Class