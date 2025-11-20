Attribute VB_Name = "modXMLTools"
Option Explicit                                                     ' Variaben Deklaration erzwingen
Const MODULNAME = "modXMLTools"                                     ' Modulname für Fehlerbehandlung

Public Function GetXMLNode(XMLDoc As DOMDocument, szXMLNodePath As String) As IXMLDOMNode
    Dim XMLRoot As IXMLDOMNodeList                                  ' Liste der XML Root Nodes
    Dim cXMLNode As IXMLDOMNode                                     ' Aktueller XML Node
    Dim cXMLRootNode As IXMLDOMNode                                 ' Ergebis XML Node
    Dim XMLNodeAtribute As IXMLDOMAttribute                         ' XML Node Atribut
    Dim szTagRest As String                                         ' Rest Tag zum suchen in Unter knoten
    Dim PathArray()  As String                                      ' Nodes Tag in Array augespalten
    Dim i As Integer
On Error GoTo Errorhandler
    If szXMLNodePath = "" Then GoTo exithandler
    PathArray = Split(szXMLNodePath, "\")                           ' Path aufspalten
'    If XMLDocPath = "" Then GoTo Exithandler                       ' Kein Doc -> Raus
'    Set XMLDoc = LoadXMLDoc(XMLDocPath)                            ' XMLDokument Laden
    Set XMLRoot = XMLDoc.selectNodes(PathArray(0))                  ' Node 1.ebene auswählen
    If XMLRoot Is Nothing Then GoTo exithandler
    If XMLRoot.length < 1 Then GoTo exithandler
    
    If UBound(PathArray) = 0 Then
        Set cXMLNode = XMLRoot(0)
    Else
        For i = 1 To UBound(PathArray)
            Set cXMLNode = GetXmlChildNode(XMLRoot(0), PathArray(i))
        Next i
    End If
    'Set XMLRoot = XMLRoot(0).selectNodes("Tree")           ' Hauptknoten Tree auswählen
    If cXMLNode Is Nothing Then GoTo exithandler
'    If cXMLNode.length < 1 Then GoTo Exithandler
    Set GetXMLNode = cXMLNode                   ' Gefundenen Knoten zurück geben
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    'Call objError.ErrorHandler(MODULNAME, "GetXMLNode", errNr, errDesc)
End Function

Public Function GetXmlChildNode(cXMLParentNode As IXMLDOMNode, szChildNodeName As String, _
    Optional szAttributName As String, Optional szAttribuValue As String) As IXMLDOMNode
' liefert Child node mit szAttributName = szAttribuValue zurück

    Dim cXMLChildNode As IXMLDOMNode                                ' Gesuchter Child Node
    Dim cXMLNode As IXMLDOMNode                                     ' Aktueller XML Node
    Dim XMLNodeAtribute As IXMLDOMAttribute                         ' XML Node Atribut
    Dim szTagRest As String                                         ' Rest Tag zum suchen in Unter knoten
    Dim PathArray()  As String                                      ' Nodes Tag in Array augespalten
    Dim bFound As Boolean                                           ' entsprechenden Node gefunden
    'Dim i As Integer
    
On Error GoTo Errorhandler

    For Each cXMLNode In cXMLParentNode.childNodes                  ' Alle Child Nodes duchlaufen
        If szChildNodeName = "" Or cXMLNode.baseName = szChildNodeName Then           ' Nur TreeNode unterknoten berücksichtigen
            If szAttributName <> "" Then                            '  evtl. Attribut prüfen
                For Each XMLNodeAtribute In cXMLNode.attributes
                    If XMLNodeAtribute.Name = szAttributName Then   ' Attribut Tag suchen
                        If szAttribuValue = "" Or XMLNodeAtribute.Value = szAttribuValue Then
                            Set cXMLChildNode = cXMLNode            ' Knoten gefunden
                            bFound = True
                            Exit For
                        End If
                    End If  ' XMLNodeAtribute.Name = szAttributName
                Next
            Else
                Set cXMLChildNode = cXMLNode                        ' Knoten gefunden
                bFound = True
                Exit For
            End If ' szAttributName <> ""
        
        End If ' cXMLNode.baseName = szChildNodeName
        If bFound Then Exit For                                     ' for verlassen
    Next

   Set GetXmlChildNode = cXMLChildNode
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    'Call objError.ErrorHandler(MODULNAME, "GetXmlChildNode", errNr, errDesc)
End Function

Public Function GetXMLAttributByName(cXMLNode As IXMLDOMNode, szAttributName As String) As String

    Dim XMLNodeAtribute As IXMLDOMAttribute                             ' XML Node Atribut

On Error GoTo Errorhandler

    For Each XMLNodeAtribute In cXMLNode.attributes                     ' Alle attribute duchgehen
        If XMLNodeAtribute.Name = szAttributName Then                   ' Attribut namen vergleichen
            GetXMLAttributByName = XMLNodeAtribute.Value
            Exit For
        End If  ' XMLNodeAtribute.Name = szAttributName
    Next
                
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    'Call objError.ErrorHandler(MODULNAME, "GetXmlChildNode", errNr, errDesc)
End Function

Public Function CheckXMLValueForBool(szValue As String, Optional DefaultValue As Boolean) As Boolean
    
    szValue = Trim(szValue)
    If szValue = "" Then
        CheckXMLValueForBool = DefaultValue
    Else
        CheckXMLValueForBool = CBool(szValue)
    End If
    
End Function

Public Function CheckXMLValueForNumeric(szValue As String, Optional DefaultValue As Variant) As Variant

    If IsNumeric(szValue) Then
        CheckXMLValueForNumeric = szValue
    Else
        CheckXMLValueForNumeric = DefaultValue
    End If
End Function
