Imports System.Data                                                             ' System.Data Klasse Importieren (Spart schreibarbeit)
Imports System.Xml                                                              ' XML Klasse Importieren (Spart schreibarbeit)

Public Class clsValueList

    Private Const MODULNAME = "clsValueList"                                    ' Modulname für Fehlerbehandlung
    Private ObjBag As clsObjectBag                                              ' Sammelklasse
    Private bInitOK As Boolean                                                  ' Gibt an das die Klasse erfolgreich initialisiert wurde
    Private ID_DELIMITER As String                                              ' Trennzeichen für zusammengesetzte (Clustered) IDs
    Private bInit As Boolean                                                    ' Gibt an das das Obj gerade initialisiert wird

    Private objConfigXML As clsXmlFile
    Private ObjCon As clsDBConnect                                              ' Aktuelle DB VerbindungsKlasse
    Private Info As ValueListInfo                                               ' Informationen zu dieser Werteliste

    Private Structure ValueListInfo
        Dim bNew As Boolean                                                     ' Gibt an ob DS neu
        Dim bDirty As Boolean                                                   ' Gibt An Ob DS Dirty (ungespeichert)
        Dim bChangeProt As Boolean                                              ' Gibt an Ob ÄnderungsProtokoll gefürt werden soll
        Dim szID As String                                                      ' ID des Ds evtl. zusammengesetzt
        Dim XmlDoc As XmlDocument                                               ' Entsprechender XML Node als XML DOc
        Dim Name As String                                                      ' Objektname (Rootkey)
        Dim SQL As String                                                       ' SQL Statenent
        Dim ExpertSQL As String                                                 ' Evtl. erweitertes SQL Statement
        Dim Table As String                 '                                   ' Datatablename im DS
        Dim Where As String                                                     ' Where Part
        Dim WhereNoDel As String                                                ' Where pArt mit gelöschten DS (Flag)
        Dim objDataAdapter As Common.DbDataAdapter
        Dim objDataSet As DataSet
        Dim objDataView As DataView
        Dim FullSQL As String                                                   ' SQL Statement incl. Where part
        Dim ImageIndex As Integer
        Dim EditCaption As String                                               ' Form Caption des Edit Forms
        Dim DelFlagField As String                                              ' Feld in dem ein gelöscht flag gesetzt werden kann
        Dim bShowDelFlag As Boolean                                             ' Sollen als gelöscht gesetzte Ds angezeigt werden
        Dim bDelPosible As Boolean                                              ' Löschen ist möglich
        Dim AskSQL As String                                                    ' SQL Statement mit Frage
        Dim DelSQL As String                                                    ' SQL Statement mit Löschanweisung
    End Structure

    Public Structure FieldInfo
        Dim Name As String
        Dim Fieldname As String
        Dim bLocked As Boolean
        Dim bEnabled As Boolean
        Dim bVisible As Boolean
        Dim bLockedNew As Boolean
        Dim Valuelist As String
        Dim ValueListSQL As String
        Dim Defaultvalue As String
        Dim DefaultSQL As String
    End Structure

#Region "Constructor"

    Public Sub New(ByVal oBag As clsObjectBag)
        Try                                                                     ' Fehlerbehandlung aktivieren
            ObjBag = oBag                                                       ' ObjectBag Übergeben    
            ObjCon = ObjBag.ObjDBConnect                                        ' Datenbankverbindung holen
            objConfigXML = ObjBag.ConfigXML                                     ' Konfig File holen
            'szIniFilePath = objObjectBag.Getappdir & objObjectBag.GetXMLFile
            bInitOK = True                                                      ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            bInitOK = False                                                     ' Misserfolg zurück
        End Try
    End Sub

    Public Sub New(ByVal oBag As clsObjectBag, _
                   ByVal EditformName As String, _
                   Optional ByVal ID As String = "")
        Try                                                                     ' Fehlerbehandlung aktivieren
            ObjBag = oBag                                                       ' ObjectBag Übergeben    
            ObjCon = ObjBag.ObjDBConnect                                        ' Datenbankverbindung holen
            objConfigXML = ObjBag.ConfigXML                                     ' Konfig File holen

            'szIniFilePath = objObjectBag.Getappdir & objObjectBag.GetXMLFile
            bInitOK = InitValue(ID, EditformName)                            ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            bInitOK = False                                                     ' Misserfolg zurück
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

    Public ReadOnly Property Caption() As String
        Get
            Caption = Info.EditCaption
        End Get
    End Property

    Public ReadOnly Property ID() As String
        Get
            ID = Info.szID
        End Get
    End Property

    Public Property Dirty() As Boolean
        Get
            Dirty = Info.bDirty
        End Get
        Set(ByVal value As Boolean)
            Info.bDirty = value
        End Set
    End Property

    Public ReadOnly Property IsNew() As Boolean
        Get
            IsNew = Info.bNew
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            Name = Info.Name
        End Get
    End Property

    Public ReadOnly Property SQL() As String
        Get
            SQL = Info.SQL
        End Get
    End Property

    Public ReadOnly Property Where() As String
        Get
            Where = Info.Where
        End Get
    End Property

    Public ReadOnly Property FullSQL() As String
        Get
            FullSQL = Info.FullSQL
        End Get
    End Property

    Public ReadOnly Property Imageindex() As Integer
        Get
            Imageindex = Info.ImageIndex
        End Get
    End Property

    Public ReadOnly Property DataSet() As DataSet
        Get
            DataSet = Info.objDataSet
        End Get
    End Property

    Public ReadOnly Property Dataview() As DataView
        Get
            Dataview = Info.objDataView
        End Get
    End Property

    Public ReadOnly Property DefaultValue(ByVal Fieldname As String) As String
        Get
            Dim fInfo As FieldInfo = GetFieldInfo(Fieldname)
            If Not IsNothing(fInfo) Then
                DefaultValue = fInfo.Defaultvalue
            Else
                DefaultValue = ""
            End If
        End Get
    End Property

    Public ReadOnly Property DefaultSQL(ByVal Fieldname As String) As String
        Get
            Dim fInfo As FieldInfo = GetFieldInfo(Fieldname)
            If Not IsNothing(fInfo) Then
                DefaultSQL = fInfo.DefaultSQL
            Else
                DefaultSQL = ""
            End If
        End Get
    End Property

    Public ReadOnly Property Locked(ByVal Fieldname As String) As Boolean
        Get
            Dim fInfo As FieldInfo = GetFieldInfo(Fieldname)
            If Not IsNothing(fInfo) Then
                Locked = fInfo.bLocked
            Else
                Locked = False
            End If
        End Get
    End Property

    Public ReadOnly Property Visible(ByVal Fieldname As String) As Boolean
        Get
            Dim fInfo As FieldInfo = GetFieldInfo(Fieldname)
            If Not IsNothing(fInfo) Then
                Visible = fInfo.bVisible
            Else
                Visible = True
            End If
        End Get
    End Property

    Public ReadOnly Property Enabled(ByVal Fieldname As String) As Boolean
        Get
            Dim fInfo As FieldInfo = GetFieldInfo(Fieldname)
            If Not IsNothing(fInfo) Then
                Enabled = fInfo.bEnabled
            Else
                Enabled = True
            End If
        End Get
    End Property

    Public ReadOnly Property GetFieldInfo(ByVal fieldname As String) As FieldInfo
        Get
            GetFieldInfo = ReadFieldInfo(fieldname)
        End Get
    End Property
    'Public Property Get IsDeletable() As Boolean
    '    IsDeletable = Info.bDelPosible
    'End Property

    'Public Property Get ShowDelFlag() As Boolean
    '    ShowDelFlag = Info.bShowDelFlag
    'End Property

    'Public Property Let SetShowDelFlag(ShowDel As Boolean)
    '    Info.bShowDelFlag = ShowDel
    'End Property

    'Public Property Get GetDelFlagField() As String
    '    GetDelFlagField = Info.DelFlagField
    'End Property

    'Public Property Get GetRS() As ADODB.Recordset
    '    Set GetRS = Info.rs
    'End Property

    ' Public Property Get GetTable() As String
    '    GetTable = Info.Table
    '    End Property
#End Region

    Private Function InitEditInfo(ByVal EditFormName As String, _
                                 Optional ByVal ID As String = "") As Boolean
        Dim xmlE As XmlElement                                                  ' Aktueller TreeNode
        Try                                                                     ' Fehlerbehandlung aktivieren
            Info = New ValueListInfo                                            ' Info Struktur initialisieren
            If GetEditNodeXML(EditFormName) Then                                ' entsprechenden XML part finden
                With Info
                    .Name = EditFormName                                               ' (Rootkey)
                    '.bChangeProt = objOptions.GetOptionByName(OPTION_CHANGEPROT)    ' Option auslesen ob ÄnderungsProtokoll
                    xmlE = .XmlDoc.DocumentElement                              ' XML Doc Des Nodes Einlesen
                    '.DelFlagField = xmlE.GetAttribute("DelFlag")
                    .SQL = xmlE.GetAttribute("SQL")                             ' SQL Statement setzen
                    .Where = xmlE.GetAttribute("WHERE")                         ' Optionales Where setzen
                    .ImageIndex = CLng(xmlE.GetAttribute("Imageindex"))
                    .ExpertSQL = xmlE.GetAttribute("ESQL")
                    If xmlE.HasChildNodes Then
                        xmlE = objConfigXML.GetXmLNode(xmlE, "Delete")
                        If Not IsNothing(xmlE) Then
                            .bDelPosible = xmlE.GetAttribute("Posible")
                            .AskSQL = xmlE.GetAttribute("ASKSQL")
                            .DelSQL = xmlE.GetAttribute("DelSQL")
                        End If
                    End If
                End With
            End If
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "InitEditInfo", ex)             ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function InitValue(ByVal ID As String, _
                          ByVal szRootkey As String) As Boolean
        'Public Function Init(DBConn As Object, ID As String, szRootkey As String, _
        ' Optional AnAdoDC As Object, Optional bShowDel As Boolean) As Boolean
        'Dim szTmpWhere As String
        Try                                                                     ' Fehlerbehandlung aktivieren
            bInit = True                                                        ' Wir initialisieren        
            With Info
                If InitEditInfo(szRootkey, ID) Then
                    If ID = "" Then                                             ' ID Leer dann
                        .bNew = True                                            ' neuer DS
                        .FullSQL = .SQL                                         ' Kein where part im SQL
                    Else
                        If InStr(ID, ID_DELIMITER) > 0 Then                     ' Zusammengesetzte ID?
                            .szID = ID
                        End If
                        .Where = .Where & "'" & .szID & "'"                     ' ID an WherePart anhängen
                        If .DelFlagField <> "" Then
                            .WhereNoDel = AddWhere(.Where, .DelFlagField & "= 0")
                        End If
                        If Not .bShowDelFlag Then
                            .FullSQL = AddWhereInFullSQL(.SQL, .Where)          ' Kompletes SQL Statement erstellen
                        Else
                            .FullSQL = AddWhereInFullSQL(.SQL, .WhereNoDel)     ' Kompletes SQL Statement erstellen
                        End If
                    End If
                End If
                ' If Not GetData(.bNew) Then Return False ' Daten Holen
                .Table = .Name                                                  ' Name der Datatable = .Name
                .objDataAdapter = ObjCon.GetDA(.FullSQL)                        ' DataAdapter generieren
                .objDataSet = ObjCon.FillDS(.objDataAdapter, .Name)             ' mit SQL Statement Dataset füllen
                .objDataView = New DataView(.objDataSet.Tables(.Name))          ' DataView holen
                If .bNew Then                                                   ' Neuer DS ?
                    '.rs.AddNew                                                  ' Neuen DS an RS anhängen
                    .EditCaption = .Name & ": Neuer Datensatz"                  ' Caption Für Edit form
                Else                                                            ' Sonst
                    .EditCaption = .Name & ": - Bearbeiten"                     ' Caption Für Edit form
                End If
            End With
            bInit = False                                                       ' Initialisierung abgeschlossen
            Return True                                                         ' Erfolg zurück liefern
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "InitValue", ex)                ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function GetEditNodeXML(ByVal EditFormName As String) As Boolean
        Dim xmlE As XmlElement                                                  ' Gesuchtes XML Element
        Try                                                                     ' Fehlerbehandlung aktivieren
            Dim TreeNode As XmlElement = objConfigXML.RootElement               ' Tree Node auswählen
            If TreeNode.HasChildNodes Then                                      ' Wenn Kindknoten vorhanden
                xmlE = objConfigXML.GetXmLNode(TreeNode, "EditForm", "Name", EditFormName)
                If IsNothing(xmlE) Then Return False ' Kein knoten gefunden -> Fertig
                'Info.XML = xmlE.OuterXml                                       ' OuterXML Merken
                Info.XmlDoc = New Xml.XmlDocument
                Info.XmlDoc.LoadXml(xmlE.OuterXml)
                Return True                                                     ' Erfolg zurück
            End If
            Return False                                                        ' Misserfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "GetEditNodeXML", ex)           ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function ReadFieldInfo(ByVal szFieldname As String) As FieldInfo
        Dim xmlE As XmlElement                                                  ' Gesuchtes XML Element
        Try                                                                     ' Fehlerbehandlung aktivieren
            If szFieldname = "" Then Return Nothing
            xmlE = objConfigXML.GetXmLNode(Info.XmlDoc.DocumentElement, "Field", "Name", szFieldname)
            Dim fInfo As FieldInfo = New FieldInfo                          ' FieldInfo initialisieren
            If Not IsNothing(xmlE) Then                                         ' Entsprechenden Konoten gefunden
                With fInfo
                    .Name = szFieldname                                          ' Name ist klar
                    .Fieldname = xmlE.GetAttribute("Feldname")
                    .bLocked = xmlE.GetAttribute("Locked")
                    .bEnabled = xmlE.GetAttribute("Enabled")
                    .bVisible = xmlE.GetAttribute("Visible")
                    .bLockedNew = xmlE.GetAttribute("LockedNew")
                    .Valuelist = xmlE.GetAttribute("Valuelist")
                    .ValueListSQL = xmlE.GetAttribute("ValueListSQL")
                    .Defaultvalue = xmlE.GetAttribute("Defaultvalue")
                    .DefaultSQL = xmlE.GetAttribute("DefaultSQL")
                End With
                Return fInfo
            Else
                With fInfo
                    .Name = szFieldname                                          ' Name ist klar
                    .Fieldname = szFieldname
                    .bLocked = False
                    .bEnabled = True
                    .bVisible = True
                    .bLockedNew = False
                    .Valuelist = ""
                    .ValueListSQL = ""
                    .Defaultvalue = ""
                    .DefaultSQL = ""
                End With
                Return fInfo
            End If
            Return Nothing
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "GetFieldInfo", ex)             ' Fehlermeldung ausgeben
            Return Nothing                                                        ' Misserfolg zurück
        End Try
    End Function

    'Private Function GetData(Optional ByVal bNew As Boolean = False) As Boolean
    '    'Private Function GetData(ByVal AnAdoDC As Object, ByVal bNew As Boolean) As Boolean

    '    'AnAdoDC.ConnectionString = ThisDBCon.GetConnectString                   ' Connect String setzen
    '    'Err.Clear()
    '    Try                                                                     ' Fehlerbehandlung aktivieren
    '        With Info
    '            .objDataSet = ObjCon.FillDS(.FullSQL, .Name)                    ' mit SQL Statement Dataset füllen
    '            .objDataView = New DataView(.objDataSet.Tables(.Name))          ' DataView holen
    '            'AnAdoDC.CommandType = adCmdText                                 ' Commandtype Setzen (SQL Statement)
    '            'AnAdoDC.RecordSource = Info.FullSQL                             ' SQL Statement übergeben
    '            'AnAdoDC.Refresh()
    '            If bNew Then
    '                'AnAdoDC.Recordset.AddNew()                                  ' Neue DS anhängen
    '                .objDataSet.Tables(.Name).NewRow()                          ' Neue DS anhängen
    '            Else

    '            End If
    '            'AnAdoDC.Refresh                                                 ' ADOC Aktualisieren
    '            'ThisADODC = AnAdoDC                                             ' Merken
    '            'Info.rs = ThisADODC.Recordset                                   ' RS Holen
    '        End With
    '        Return True                                                         ' Erfolg melden
    '    Catch ex As Exception                                                   ' Fehler behandeln
    '        Call ObjBag.ErrorHandler(MODULNAME, "GetData", ex)                  ' Fehlermeldung ausgeben
    '        Return False                                                        ' Misserfolg zurück
    '    End Try
    'End Function

    Public Function Refresh(Optional ByVal AnAdoDC As Object = Nothing) As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            If bInit Then Return False ' Nicht bei initialisierung
            'If Not GetData(Info.bNew) Then Return False ' Daten Holen
            '    Call GetRelInfos(False)                                         ' Relation Daten Aktualisieren
            Return True
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "Refresh", ex)                  ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Public Function Save() As Boolean
        'Dim szChangeLog As String                                               ' Text für Protokoll
        Try                                                                     ' Fehlerbehandlung aktivieren
            With Info
                If .bDirty Then                                                 ' Nur wenn Dirty (ungespeichert)
                    If .bNew Then                                               ' Bei neuen Datensatz

                    Else

                    End If
                    If .objDataSet.HasChanges Then
                        .objDataSet = ObjCon.UpdateDS(.objDataAdapter, .objDataSet, .Table)
                        '.objDataAdapter.Update(.objDataSet, .Table)             ' Mit DataAdapter Dataset Updaten
                    End If
                    'AnAdoDC.Recordset.Update()                                  ' RS Speichern
                    'ThisADODC = AnAdoDC
                    Call WriteChangeProt(False)                                 ' evtl. änderungen Protokolieren
                    .EditCaption = .Name & ": - Bearbeiten"                     ' Caption Für Edit form
                    .bDirty = False                                             ' DS nicht mehr Dirty
                    .bNew = False                                               ' DS Nicht mehr neu
                End If
            End With
            Return True                                                         ' Erfolg Melden
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "Save", ex)                     ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Public Function Delete() As Boolean
        ' Diesen Datensatz Löschen bzw auf gelöscht setzen (Del flag)
        Dim szMSG As String                                                     ' Message text
        Dim FullAskSQL As String = ""                                           ' SQL Statement mit Frage
        Dim FullDelSQL As String = ""                                           ' SQL Statement mit Löschanweisung
        'Dim bDelSuccess As Boolean                                              ' Löschen erfolgreich
        Try                                                                     ' Fehlerbehandlung aktivieren
            With Info
                If Not .bDelPosible Then Return False ' kein löschen vorgesehen -> fertig
                If .DelFlagField <> "" Then                                     ' DS auf gelöscht sezen möglich ?

                Else
                    If .AskSQL <> "" Then                                       ' SQL für Frage forhanden
                        If .Where <> "" Then                                    ' Wherepart vorhanden
                            FullAskSQL = AddWhereInFullSQL(.AskSQL, .Where)     ' Kompletes SQL Statement erstellen
                        Else
                            FullAskSQL = .AskSQL                                ' Sonst Ohne Where
                        End If
                        szMSG = ObjCon.GetValueFromSQL(FullAskSQL)           ' Frage aus AskSQL ermitteln
                    Else
                        szMSG = "Möchten Sie diesen Datensatz wirklich löschen?" ' Falls kein AskSQL angegeben
                    End If
                    If szMSG = "" Then Return False ' Immernoch kein Fragetext -> Fertig
                    If ObjBag.ShowErrMsg(szMSG, vbQuestion + vbOKCancel, "Löschen") = vbOK Then ' Sicherheitshalber nachfragen
                        If .DelSQL <> "" Then                                   ' Löschanweisung vorhanden
                            If .Where <> "" Then                                ' Wherepart vorhanden
                                FullDelSQL = AddWhereInFullSQL(.DelSQL, .Where)  ' Kompletes SQL Statement erstellen
                            End If
                        End If
                        If FullDelSQL <> "" Then                                ' Löschstatement komplett ?
                            Call WriteChangeProt(True)                          ' Löschung Protokolieren
                            Return ObjCon.ExecSQL(FullDelSQL)                ' Löschstatement ausführen
                        End If
                    End If
                End If
            End With
            Return False
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "Delete", ex)                   ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Public Function ReDelete() As Boolean
        'Diesen Benutzer auf nichtgelöscht (flag) setzen
        'Dim szMSG As String                                                     ' Message text
        'Dim bReDelAllUsers As Boolean                                           ' True wenn auch alle user Widerhergestellt werden sollen
        Try                                                                     ' Fehlerbehandlung aktivieren

            '    szMSG = "Möchten Sie den Benutzer " & Info.Username & " in allen Abteilungen in denen dieser gelöscht wurde wiederherstellen?"
            '    If objError.ShowErrMsg(szMSG, vbQuestion + vbOKCancel, "Wiederherstellen") = vbOK Then  ' Sicherheitshalber nachfragen
            '        bReDelAllUsers = True
            '    End If
            '
            '    ReDelete = ReDelUser(ThisDBCon, Info.PersID, Info.nr, Info.kz, _
            '            bReDelAllUsers, False, Info.bChangeProt)                ' löschen rückgängig
            '    If ReDelete And bReDelAllUsers Then                             ' Wenn erfolg
            '        Call GetRelInfos(True)                                      ' Relation Daten Aktualisieren
            '    End If

            ' User in allen seinen Abteilungen auf nicht gelöscht setzen (a_besetzung)

            ' Entsprechende einträge in a_sicherheit anlegen

            ' Entsprechende einträge in a_bereiche für diese User anlegen
            Return True
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "ReDelete", ex)                 ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function WriteChangeProt(ByVal bDelete As Boolean) As Boolean
        Dim szChangeLog As String                                               ' Text für Protokoll
        'Dim i As Integer                                                        ' Counter für RS felder
        Try                                                                     ' Fehlerbehandlung aktivieren
            With Info
                If .bChangeProt Then                                            ' Wenn Protokolierung
                    If bDelete Then                                             ' DS wird gelöscht
                        szChangeLog = .Name & " gelöscht: "
                    ElseIf .bNew Then                                           ' Bei neuen Datensatz
                        szChangeLog = .Name & " neu angelegt: "
                    Else                                                        ' Bei Bestehendem DS
                        szChangeLog = .Name & " geändert: "
                    End If
                    'Call objError.WriteProt(szChangeLog)                        ' ins Protokoll schreiben
                    'For i = 0 To ThisADODC.Recordset.Fields.Count - 1           ' Alle felder durchlaufen
                    '    szChangeLog = "  " & ThisADODC.Recordset.Fields(i).Name & _
                    '            "=" & ThisADODC.Recordset.Fields(i).Value       ' Protokoll text festlegen
                    '    Call objError.WriteProt(szChangeLog)                    ' ins Protokoll schreiben
                    'Next i                                                      ' Nächstes Feld
                End If
            End With
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "WriteChangeProt", ex)          ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function LoadDefaultValue(ByVal Fieldname As String) As String
        ' Holtdefault werte aus XML
        Dim szDefValue As String = ""                                           ' Default Value
        Try                                                                     ' Fehlerbehandlung aktivieren
            If Fieldname = "" Then Return "" ' kein Fielname -> Fertig
            Dim fInfo As FieldInfo = GetFieldInfo(Fieldname)                    ' Feld infos enlesen (aus xml)
            If Not IsNothing(fInfo) Then                                        ' Feld infos gefunden
                With fInfo
                    If .DefaultSQL <> "" Then                                   ' SQL Statement für defaultwert vorhanden
                        szDefValue = ObjCon.GetValueFromSQL(.DefaultSQL)        ' Ausführen und 1. wert zurück
                    End If
                    If szDefValue = "" Then                                     ' Wenn kein defaut wrt gefunden
                        szDefValue = .Defaultvalue                              ' String angeben
                    End If
                End With
                Return szDefValue                                               ' Gefundenen Default wert zurück
            End If
            Return ""                                                           ' Wenn wir hier ankommen gibt es kein Defaut wert
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "LoadDefaultValue", ex)         ' Fehlermeldung ausgeben
            Return ""                                                           ' Misserfolg zurück
        End Try
    End Function

End Class
