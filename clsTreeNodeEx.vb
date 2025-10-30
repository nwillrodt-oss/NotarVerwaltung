Imports System.Xml                                                              ' XML Klasse Importieren (Spart schreibarbeit)

Public Class ClsTreeNodeEx
    Inherits TreeNode                                                           ' Org. TreeNode beerben

    Private Const MODULNAME = "ClsTreeNodeEx"                                   ' Modulname für Fehlerbehandlung

    Private ObjBag As clsObjectBag                                              ' Sammelobject
    Private ConfigXML As Xml.XmlDocument                                        ' Config XML File
    Public ExTVNodeInfo As TVNodeInfoEx
    Public ExListInfo As ListViewInfoEx

    Private Const XML_PATH_SEP = "\"

    Public Structure TVNodeInfoEx                                               ' TV Konoten informationen
        Public ID As String
        Public Desc As String                                                   ' Angezeigte Beschreibung
        Public ImageIndex As Integer
        Public SQL As String                                                    ' zugrundeliegendes SQL Statement
        Public WHERE As String                                                  ' evtl. Where Statement
        Public XML As String                                                    ' Entsprechender XML Node als String
        Public XmlDoc As XmlDocument                                            ' Entsprechender XML Node als XML DOc
        Public Typ As String                                                    ' Statisch oder dynamisch
        Public bShowSubnodes As Boolean                                         ' Unterknoten sofort anzeigen
        Public bShowKontextMenue As Boolean
        Public ChildnodeList As String
    End Structure

    Public Structure ListViewInfoEx                                             ' ListView Informationen
        Public SQL As String                                                    ' zugrundeliegendes SQL Statement
        Public XmlDoc As XmlDocument                                            ' Entsprechender XML Node als XML DOc
        Public szTag As String                                                  ' Tag des ListViews (welche Daten werden angezeigt)
        Public bValueList As Boolean                                            ' Darstellung als Valuelist (1.DS pro Wert ein Item)
        Public DelFlagField As String                                           ' Feld in dem ein gelöscht flag gesetzt werden kann
        Public WhereNoDel As String                                             ' Where Part mit gelöschten DS (Flag)
        Public WHERE As String                                                  ' evtl. Where Statement
        Public lngImage As Integer                                              ' Image Index für Item
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

    Public Sub New(ByVal oBag As clsObjectBag, ByVal xmlNodePath As String, Optional ByVal Parent As TreeNode = Nothing)
        Try                                                                     ' Fehlerbehandlung aktivieren
            ObjBag = oBag                                                       ' ObjectBag Übergeben
            ConfigXML = ObjBag.ConfigXMLDoc                                     ' ConfigDatei laden
            Call InitNodeInfo(xmlNodePath)                                      ' Extendet infos zum Node laden
        Catch ex As Exception                                                   ' Fehler behandeln

        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Function InitNodeInfo(ByVal xmlNodename As String) As Boolean
        Dim xmlE As XmlElement                                                  ' Aktueller TreeNode
        Try                                                                     ' Fehlerbehandlung aktivieren
            ExTVNodeInfo = New TVNodeInfoEx                                     ' Zusatzinfo initialisieren
            If GetNodeXML(xmlNodename) Then                                     ' entsprechendem XML part finden
                With ExTVNodeInfo
                    xmlE = .XmlDoc.DocumentElement                              ' XML DOck Des Nods Einlesen
                    Me.Text = xmlE.Attributes("Text").Value                     ' Text Setzen
                    .Typ = xmlE.Attributes("Typ").Value                         ' Typ Setzen
                    .bShowSubnodes = xmlE.Attributes("bShowSubnodes").Value     ' SubNodes anzeigen setzen
                    .bShowKontextMenue = xmlE.Attributes("bShowKontextMenue").Value ' Kontextmenü anzeigen setzen
                    .Desc = xmlE.Attributes("Description").Value                ' Beschreibung setzen
                    .SQL = xmlE.Attributes("SQL").Value                         ' SQL Statement setzen
                    .WHERE = xmlE.Attributes("WHERE").Value                     ' Optionales Where setzen
                    Call InitListInfo(xmlE, xmlNodename)
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
            xmlE = GetXmLNode(xmlE, "List", xmlNodename)                        ' entsprechendem XML part finden
            If Not xmlE Is Nothing Then
                With ExListInfo
                    .SQL = xmlE.Attributes("SQL").Value                         ' SQL Statement setzen
                    .WHERE = xmlE.Attributes("WHERE").Value                     ' Optionales Where setzen
                End With
                Return True
            Else
                Return False
            End If
           
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "InitNodeInfo", ex)             ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function AddSubNodes() As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            Dim x As Xml.XmlElement = ExTVNodeInfo.XmlDoc.DocumentElement
            For Each subnode In x.ChildNodes
                If subnode.Name = "Treenode" Then
                    Dim stvn As New ClsTreeNodeEx(ObjBag, subnode.Attributes("Name").Value, Me)
                    Me.Nodes.Add(stvn)
                End If
            Next
            Return True
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "InitNodeInfo", ex)             ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function GetNodeXML(ByVal xmlNodeName As String) As Boolean
        Dim NodeNode As XmlElement
        Try                                                                     ' Fehlerbehandlung aktivieren
            Dim ConfigRootNode As XmlElement = ConfigXML.DocumentElement        ' Wurzelknoten ermittelm
            If ConfigRootNode.HasChildNodes Then                                ' Wenn Kindknoten vorhanden
                Dim TreeNode As XmlElement = ConfigRootNode.ChildNodes(0)       ' Tree Node auswählen
                If TreeNode.HasChildNodes Then                                  ' Wenn Kindknoten vorhanden
                    NodeNode = GetXmLNode(TreeNode, "Treenode", "Name", xmlNodeName)
                    ExTVNodeInfo.XML = NodeNode.OuterXml                        ' OuterXML Merken
                    ExTVNodeInfo.XmlDoc = New Xml.XmlDocument
                    ExTVNodeInfo.XmlDoc.LoadXml(NodeNode.OuterXml)
                    Return True                                                 ' Erfolg zurück
                End If
            End If
            Return False                                                        ' Misserfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "GetNodeXML", ex)               ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Public Function GetXmLNode(ByVal xmlE As XmlElement, ByVal szElementName As String, _
                               Optional ByVal szAttrName As String = "", _
                               Optional ByVal szAttrValue As String = "") As XmlElement
        Dim xmlFound As XmlElement
        Try                                                                     ' Fehler behandlung aktivieren
            If xmlE.Name = szElementName Then
                If szAttrName <> "" Then
                    ' Erst Attribute dieses Elements Durchsuchen
                    If xmlE.Attributes.Count > 0 Then                                   ' Wenn Attributte vorhanden
                        If szAttrValue = "" Then
                            If Not IsNothing(xmlE.Attributes(szAttrName)) Then
                                Return xmlE
                            End If
                        Else
                            If xmlE.Attributes(szAttrName).Value = szAttrValue Then
                                Return xmlE
                            End If
                        End If
                    End If
                Else
                    Return xmlE
                End If
            End If

            ' Dann Recursiv Kindknoten durchsuchen
            If xmlE.HasChildNodes Then                                          ' Wurzelknoten ermittelm
                For Each cNode In xmlE.ChildNodes                               ' Alle Kindknoten  durchlaufen
                    If cNode.name = szElementName Then
                        xmlFound = GetXmLNode(cNode, szElementName, szAttrName, szAttrValue)
                        If Not xmlFound Is Nothing Then Return xmlFound
                    End If
                Next
            End If
            Return Nothing
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "GetXMLNode", ex)               ' Fehlermeldung ausgeben
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

End Class
