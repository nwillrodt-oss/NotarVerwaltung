VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MODULNAME = "clsPerson"

Private ID_DELIMITER As String                                      ' Trennzeichen für zusammengesetzte (Clustered) IDs
Private bInit As Boolean                                            ' Gibt an das das Ob gerade initialisiert wird
Private szIniFilePath As String                                     ' Pfad der Ini datei
Private ThisDBCon As Object                                         ' Aktuelle DB Verbindung
Private ThisADODC As Object                                         ' AdoDC

Private Info As PersonInfo                                            ' Informationen über diesen Benutzer
'Private RelUnitInfo As RelationInfo                                 ' Relationen Informationen zu Abteilungen (a_besetzung)
'Private RelSecureInfo As RelationInfo                               ' Relationen Informationen zu Sicherheit (a_sicherheit)

Private Type PersonInfo
    bNew As Boolean                                                 ' Gibt an ob DS neu
    bDirty As Boolean                                               ' Gibt An Ob DS Dirty (ungespeichert)
    bChangeProt As Boolean                                          ' Gibt an Ob ÄnderungsProtokoll gefürt werden soll
    szID As String                                                  ' ID des Ds evtl. zusammengesetzt
    Name As String                                                  ' Objektname (Rootkey)
    SQL As String                                                   ' SQL Statenent
    Table As String                 '                               ' Bearbeitete Tabelle ( Hier RA010)
    Where As String                                                 ' Where Part
    WhereNoDel As String                                            ' Where Part mit gelöschten DS (Flag)
    RS As ADODB.Recordset                                           ' Recordset
    FullSQL As String                                               ' SQL Statement incl. Where part
    Image As Integer
    EditCaption As String                                           ' Form Caption des Edit Forms
    DelFlagField As String                                          ' Feld in dem ein gelöscht flag gesetzt werden kann
    bShowDelFlag As Boolean                                         ' Sollen als gelöscht gesetzte Ds angezeigt werden
    PersID As String                                                ' RA010.ID010
    Nachname As String                                              ' RA010.Nachname010
    Vorname As String                                               ' RA010.Vorname010
    Status As String                                                ' RA010.Status010
    Fullname As String                                              ' Nachname , Vorname
End Type

Private Type RelationInfo
    Name As String
    SQL As String
    Table As String                 '                               ' Bearbeitete Tabelle
    Where As String
    WhereNoDel As String                                            ' Where Part mit gelöschten DS (Flag)
    FullSQL As String
    RelRS As ADODB.Recordset
    Image As Integer
    AltImage As Integer                                             ' Alternatives Image
    AltImgField As String                                           ' Feld das für alt Image geprüft wird
    AltImgValue As String                                           ' Value der für alt image geprüft wird
    DelFlagField As String                                          ' Feld in dem ein gelöscht flag gesetzt werden kann
    bShowDelFlag As Boolean                                         ' Sollen als gelöscht gesetzte Ds angezeigt werden
End Type

Private Sub Class_Initialize()
    szIniFilePath = objObjectBag.Getappdir & objObjectBag.GetXMLFile
    ID_DELIMITER = Chr(1)
End Sub

Private Sub Class_Terminate()
'

End Sub
                                                                    ' *****************************************
                                                                    ' Methoden
Public Function Init(DBConn As Object, AnAdoDC As Object, ID As String, _
    Optional bShowDel As Boolean) As Boolean
'
On Error GoTo Errorhandler

    bInit = True                                                    ' Wir initialisieren
    Set ThisDBCon = DBConn                                          ' Aktuelle DB Verbindung übernehmen
    With Info
        .Name = "Personenkartei"                                          ' (Rootkey)
        .Table = "RA010"
        '.bChangeProt = objOptions.GetOptionByName(OPTION_CHANGEPROT) ' Option auslesen ob ÄnderungsProtokoll
        '.bShowDelFlag = objOptions.GetOptionByName(OPTION_SHOWDELREL) ' Option Gelöschte beziehungen anzeigen auslesen
        'If bShowDel And Not .bShowDelFlag Then .bShowDelFlag = bShowDel
        
        Call objTools.GetEditInfoFromXML(szIniFilePath, .Name, .SQL, .Where, .Image, .DelFlagField)
    
        If ID = "" Then                                             ' ID Leer dann
            .bNew = True                                            ' neuer DS
            .FullSQL = .SQL                                         ' Kein where part im SQL
        Else
            .szID = ID
            .Where = .Where & "'" & .szID & "'"                     ' ID an WherePart anhängen
            .FullSQL = objSQLTools.AddWhereInFullSQL(.SQL, .Where)  ' Kompletes SQL Statement erstellen
        End If
    
        If Not GetData(AnAdoDC) Then GoTo exithandler               ' Daten Holen
    
        If .bNew Then                                               ' Neuer DS ?
            .RS.AddNew                                              ' Neuen DS an RS anhängen
            .EditCaption = .Name & ": Neuer Datensatz"              ' Caption Für Edit form
        Else
            .Nachname = .RS.Fields("Nachname010").Value             ' Nachname holen
            .Vorname = .RS.Fields("Vorname010").Value               ' Vorname holen
            .Status = .RS.Fields("Status010").Value                 ' Status
            .Fullname = .Nachname                                   ' Vollen namen zusammen setzen
            If .Vorname <> "" Then .Fullname = .Fullname & ", " & .Vorname
            .EditCaption = GetEditCaption()
            
        End If
    End With
    
    Call GetRelInfos(False)                                         ' Relationdaten holen
    
    Init = True                                                     ' Erfolg Melden
    
exithandler:
    bInit = False                                                   ' Initialisierung abgeschlossen
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "Init", errNr, errDesc)
    Resume exithandler
End Function

Public Function Save(AnAdoDC As Object) As Boolean

    Dim bDelete As Boolean                                          ' Lösch flag gesetzt ?
    Dim bOK As Boolean
    Dim szMSG As String                                             ' Message text
On Error GoTo Errorhandler

    szMSG = "Der Datensatz kann nicht gespeichert werden da eine Schlüsselverletzung vorliegt."
    
    With Info
        If .bDirty Then                                             ' Nur wenn Dirty (ungespeichert)
            If .bNew Then                                           ' Bei neuen Datensatz
            
            Else
                'bDelete = AnAdoDC.Recordset.Fields("loesch_kz").Value
'                If bDelete And Not .bDelete Then                    ' Wenn gelöscht
'                '    AnAdoDC.Recordset.Fields("mega_kz").Value = 0   ' dann auch kein Megazugriff
'                End If
            End If
On Error Resume Next
            AnAdoDC.Recordset.Update                                ' RS Speichern
            If AnAdoDC.Recordset.Status = adRecIntegrityViolation Then
                'Call objError.ShowErrMsg(szMSG, vbCritical, "Speichern")
                    'Stop
                    GoTo exithandler
            End If
    
            Err.Clear
On Error GoTo Errorhandler

            Set ThisADODC = AnAdoDC
            .EditCaption = GetEditCaption()                         ' Caption Für Edit form
            
            Call GetRelInfos(True)                                  ' RelationDaten aktualisieren
'            Call GetSecureRelInfo(True)                             ' Daten für Relation Sicherheit aktualisieren (a_sicherheit)
'            Call GetUnitRelInfo(True)                               ' Daten für Relation Abteilung aktualisieren (a_besetzung)
            
'            If bDelete And Not .bDelete Then
'                bOK = Delete                                        ' DS auf gelöscht setzen
'            ElseIf .bDelete And Not bDelete Then
'                bOK = ReDelete                                      ' Wiederherstellung
'            Else
'                bOK = True
'            End If
            
        End If
        If bOK Then
            .bDirty = False                                         ' DS nicht mehr Dirty
            .bNew = False                                           ' DS Nicht mehr neu
'            .bDelete = CBool(objTools.checkNull(.RS.Fields("loesch_kz").Value, False))
'            .bMegaAccess = CBool(objTools.checkNull(.RS.Fields("mega_kz").Value, True))
            Save = True                                             ' Erfolg Melden
        End If
    End With

   
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "Save", errNr, errDesc)
    Resume exithandler
End Function

Public Function Delete() As Boolean
'Diesen Benutzer auf gelöscht (flag) setzen (benutzer bleibt erhalten)

    Dim szMSG As String                                             ' Message text

On Error GoTo Errorhandler

'    szMSG = "Möchten Sie die Person " & Info.Fullname & " wirklich löschen?"
'    If objError.ShowErrMsg(szMSG, vbQuestion + vbOKCancel, "Löschen") = vbOK Then ' Sicherheitshalber nachfragen
'        Delete = DelUser(ThisDBCon, Info.PersID, Info.nr, Info.kz, False, Info.bChangeProt)  ' löschen
'        If Delete Then
'            Call GetRelInfos(True)                                  ' RelationDaten aktualisieren
'            'Call GetUnitRelInfo(True)                                ' Rel Daten refreshen
'            'Call GetSecureRelInfo(True)                              ' Rel Daten refreshen
'        End If
'    End If
    ' User aus allen Abteilungen entfernen (a_besetzung)
        
    ' alle einträge in a_sicherheit für diesen Benutzer löschen
    
    ' Alle einträge in a_bereiche für diese User löschen
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "Delete", errNr, errDesc)
    Resume exithandler
End Function

'Public Function ReDelete() As Boolean
''Diesen Benutzer auf nichtgelöscht (flag) setzen
'
'    Dim szMSG As String                                             ' Message text
'    Dim bReDelAllUsers As Boolean                                   ' True wenn auch alle user Widerhergestellt werden sollen
'
'On Error GoTo Errorhandler
'
'    szMSG = "Möchten Sie den Benutzer " & Info.Username & " in allen Abteilungen in denen dieser gelöscht wurde wiederherstellen?"
'    If objError.ShowErrMsg(szMSG, vbQuestion + vbOKCancel, "Wiederherstellen") = vbOK Then  ' Sicherheitshalber nachfragen
'        bReDelAllUsers = True
'    End If
'
'    ReDelete = ReDelUser(ThisDBCon, Info.PersID, Info.nr, Info.kz, _
'            bReDelAllUsers, False, Info.bChangeProt)                ' löschen rückgängig
'    If ReDelete And bReDelAllUsers Then                             ' Wenn erfolg
'        Call GetUnitRelInfo(True)                                   ' Rel Daten refreshen
'        Call GetSecureRelInfo(True)                                 ' Rel Daten refreshen
'    End If
'
'    ' User in allen seinen Abteilungen auf nicht gelöscht setzen (a_besetzung)
'
'    ' Entsprechende einträge in a_sicherheit anlegen
'
'    ' Entsprechende einträge in a_bereiche für diese User anlegen
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
'    Call objError.Errorhandler(MODULNAME, "ReDelete", errNr, errDesc)
'    Resume exithandler
'End Function

Public Function SecureAdd(Optional ClusterdRelID As String) As Boolean
' Eintrag In a_sicherheit für diesen User hinzufügen

    Dim szBereich As String                                         ' Welcher Gerichtsbereich
    Dim szAbteilung As String                                       ' Welche Abteilung
    Dim szRegister As String                                        ' Welches register

On Error GoTo Errorhandler

'    If ClusterdRelID = "" Then                                      ' Wenn keine ID dann suchen
'        ClusterdRelID = ShowSearch(ThisDBCon, "UserSicherheit", "Register")
'    End If
'    If ClusterdRelID = "" Then GoTo exithandler                     ' Keine ID fertig
'    szBereich = GetIDFromClusterID(ClusterdRelID, "bereich")        ' Bereich herausdröseln
'    szAbteilung = GetIDFromClusterID(ClusterdRelID, "abteilung")    ' Abteilung herausdröseln
'    szRegister = GetIDFromClusterID(ClusterdRelID, "register")      ' Register herausdröseln
'
'    If szAbteilung = "" Then GoTo exithandler                       ' Keine Abteilung -> fertig
'    If szRegister = "" Then GoTo exithandler                        ' Kein Register -> fertig
'
'    SecureAdd = InsertInSecure(ThisDBCon, Info.PersID, szBereich, szAbteilung, szRegister, False, Info.bChangeProt)
'
'    Call GetSecureRelInfo(True)                                     ' Rel RS refreshen
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SecureAdd", errNr, errDesc)
    Resume exithandler
End Function

Public Function SecureDel(ClusterdRelID As String) As Boolean
' Eintrag In a_sicherheit für diesen User entfernen

    Dim szBereich As String                                         ' Welcher Gerichtsbereich
    Dim szAbteilung As String                                       ' Welche Abteilung
    Dim szRegister As String                                        ' Welches register
    Dim szMSG As String                                             ' Message Text
    
On Error GoTo Errorhandler

'    If ClusterdRelID = "" Then GoTo exithandler                     ' Keine ID fertig
'    szBereich = GetIDFromClusterID(ClusterdRelID, "bereich")        ' Bereich herausdröseln
'    szAbteilung = GetIDFromClusterID(ClusterdRelID, "abteilung")    ' Abteilung herausdröseln
'    szRegister = GetIDFromClusterID(ClusterdRelID, "register")      ' Register herausdröseln
'
'    If szAbteilung = "" Then GoTo exithandler                       ' Keine Abteilung -> fertig
'    If szRegister = "" Then GoTo exithandler                        ' Kein Register -> fertig
'
'    szMSG = "Möchten Sie dem Benutzer " & Info.Username & " den Zugriff auf das Register " _
'                    & szRegister & " der Abteilung " & szAbteilung & " wirklich entziehen?"
'
'    If MsgBox(szMSG, vbQuestion + vbOKCancel, "Löschen") = vbOK Then ' Sicherheitshalber nachfragen
'        SecureDel = DelFromSecure(ThisDBCon, Info.PersID, szBereich, szAbteilung, szRegister, False, Info.bChangeProt)
'        Call GetSecureRelInfo(True)                                  ' Rel RS refreshen
'    End If
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SecureDel", errNr, errDesc)
    Resume exithandler
End Function

Public Function UnitAdd(Optional ClusterdRelID As String) As Boolean
' Diesen User in a_besetzung einer Abteilung hinzufügen

    Dim szBereich As String                                         ' Welcher Gerichtsbereich
    Dim szAbteilung As String                                       ' Welche Abteilung
    'Dim szProt As String                                            ' Protokoll text
    
On Error GoTo Errorhandler

'    If ClusterdRelID = "" Then                                      ' Wenn keine ID dann suchen
'        ClusterdRelID = ShowSearch(ThisDBCon, "Abteilung", "Kurzbezeichnung")
'    End If
'    If ClusterdRelID = "" Then GoTo exithandler                     ' Keine ID fertig
'    szBereich = GetIDFromClusterID(ClusterdRelID, "bereich")        ' Bereich herausdröseln
'    szAbteilung = GetIDFromClusterID(ClusterdRelID, "abteilung")    ' Abteilung herausdröseln
'
'    If szAbteilung = "" Then GoTo exithandler                       ' Keine Abteilung -> fertig
'    If szBereich = "" Then GoTo exithandler                         ' Keine Bereich -> fertig
'
'    'szProt = "Benutzer: " & Info.Username & " der Abteilung " & szAbteilung & " hinzugefügt."
'
'    UnitAdd = InsertInUnit(ThisDBCon, Info.PersID, szAbteilung, Info.nr, Info.kz, szBereich, False, Info.bChangeProt)
'
'    Call GetUnitRelInfo(True)                                       ' Rel RS refreshen
'    Call GetSecureRelInfo(True)                                     ' Rel RS refreshen
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "UnitAdd", errNr, errDesc)
    Resume exithandler
End Function

Public Function UnitDel(ClusterdRelID As String) As Boolean
' Diesen User in a_besetzung aus Abteilung entfernen

    Dim szBereich As String                                         ' Welcher Gerichtsbereich
    Dim szAbteilung As String                                       ' Welche Abteilung
    Dim szMSG As String                                             ' Message Text
    'Dim szProt As String                                            ' Protokoll text
    
On Error GoTo Errorhandler

'    If ClusterdRelID = "" Then GoTo exithandler                     ' Keine ID fertig
'    szBereich = GetIDFromClusterID(ClusterdRelID, "bereich")        ' Bereich herausdröseln
'    szAbteilung = GetIDFromClusterID(ClusterdRelID, "abteilung")    ' Abteilung herausdröseln
'
'    If szAbteilung = "" Then GoTo exithandler                       ' Keine Abteilung -> fertig
'    If szBereich = "" Then GoTo exithandler                         ' Keine Bereich -> fertig
'
'    szMSG = "Möchten Sie den Benutzer " & Info.Username & " aus der Abteilung " _
'                    & szAbteilung & " wirklich entfernen?"
'    'szProt = "Benutzer: " & Info.Username & " aus der Abteilung " & szAbteilung & " entfernt."
'
'    If MsgBox(szMSG, vbQuestion + vbOKCancel, "Löschen") = vbOK Then ' Sicherheitshalber nachfragen
'        UnitDel = DelFromUnit(ThisDBCon, Info.PersID, szAbteilung, Info.nr, Info.kz, szBereich, False, Info.bChangeProt)
'        'If Info.bChangeProt Then Call objError.WriteProt(szProt)    ' Änderungen protokolieren
'        Call GetUnitRelInfo(True)                                   ' Rel RS refreshen
'        Call GetSecureRelInfo(True)                                 ' Rel RS refreshen
'    End If
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "UnitDel", errNr, errDesc)
    Resume exithandler
End Function
                                                                    ' *****************************************
                                                                    ' Hilfsfunktionen
Private Function GetData(AnAdoDC As Object) As Boolean
' Initailisiert das Übergebene ADODC
' und läd die Daten

On Error Resume Next                                                ' Fehlerbehandlung Deakt. Da Sonst Putzige Fehlermeldung
        AnAdoDC.ConnectionString = ThisDBCon.GetConnectString       ' Connect String setzen
        Err.Clear
On Error GoTo Errorhandler                                          ' Fehlerbehandlung Wieder Akivieren
        
        AnAdoDC.CommandType = adCmdText                             ' Commandtype Setzen (SQL Statement)
        AnAdoDC.RecordSource = Info.FullSQL                         ' SQL Statement übergeben
        AnAdoDC.Refresh                                             ' ADOC Aktualisieren
        
        Set ThisADODC = AnAdoDC                                     ' Merken
        Set Info.RS = ThisADODC.Recordset                           ' RS Holen
        
        GetData = True                                              ' Erfolg melden
        
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetData", errNr, errDesc)
    Resume exithandler
End Function

Private Sub GetUnitRelInfo(bRefreshOnly As Boolean)

On Error GoTo Errorhandler

'    With RelUnitInfo                                                ' Daten für Relation Abteilung holen (a_besetzung)
'        If Not bRefreshOnly Then
'            .Name = "Abteilungen"
'            .Table = "a_besetzung"
'            Call objTools.GetEditRelInfoFromXML(szIniFilePath, Info.Name, .Name, .SQL, .Where, .Image, .DelFlagField, _
'                    .AltImage, .AltImgField, .AltImgValue)          ' Realtions Informationen aus XML holen
'            .Where = .Where & "'" & Info.szID & "'"                 ' ID an WherePart anhängen
'            .FullSQL = objSQLTools.AddWhereInFullSQL(.SQL, .Where)  ' Kompletes SQL Statement erstellen
'        End If
'        Set .RelRS = ThisDBCon.fillrs(.FullSQL, True)               ' RS füllen
'    End With
    
exithandler:
                                                
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetUnitRelInfo", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub GetSecureRelInfo(bRefreshOnly As Boolean)

On Error GoTo Errorhandler

'    With RelSecureInfo                                              ' Daten für Relation Sicherheit holen (a_sicherheit)
'        If Not bRefreshOnly Then
'            .Name = "Sicherheit"
'            .Table = "a_sicherheit"
'            Call objTools.GetEditRelInfoFromXML(szIniFilePath, Info.Name, .Name, .SQL, .Where, .Image, .DelFlagField, _
'                    .AltImage, .AltImgField, .AltImgValue)          ' Realtions Informationen aus XML holen
'            .Where = .Where & "'" & Info.szID & "'"                 ' ID an WherePart anhängen
'            .FullSQL = objSQLTools.AddWhereInFullSQL(.SQL, .Where)  ' Kompletes SQL Statement erstellen
'        End If
'        Set .RelRS = ThisDBCon.fillrs(.FullSQL, True)               ' RS füllen
'    End With
    
exithandler:
                                                
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetUnitRelInfo", errNr, errDesc)
    Resume exithandler
End Sub

Private Function GetEditCaption() As String
    Info.EditCaption = Info.Fullname & " - Bearbeiten"              ' Caption Für Edit form
    GetEditCaption = Info.EditCaption
End Function

Private Function GetRelInfos(bRefreshOnly As Boolean)
    Call GetUnitRelInfo(bRefreshOnly)
    Call GetSecureRelInfo(bRefreshOnly)
End Function

Private Function LoadDefaultValue(Fieldname As String) As String
' Holtdefault werte aus XML

    Dim szDefValue As String                                        ' Default Value
    Dim szDefValSQL As String                                       ' SQL Statement das Value holt
    
On Error GoTo Errorhandler

    If Fieldname = "" Then GoTo exithandler                         ' kein Fielname -> Fertig
    
    With Info
        szDefValue = objTools.GetEditDefaultValue(szIniFilePath, .Name, Fieldname, szDefValSQL)
        If szDefValSQL <> "" Then szDefValue = ThisDBCon.GetValueFromSQL(szDefValSQL)
        If szDefValue <> "" Then
            LoadDefaultValue = szDefValue
            GoTo exithandler
        End If
    End With

exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "LoadDefaultValue", errNr, errDesc)
    Resume exithandler
End Function

Private Function GetLockedField(Fieldname As String, bLocked As Boolean, bEnabeld As Boolean, bVisible As Boolean) As Boolean
' Stellt fest ob feld Gelocked enabled und/oder Visible ist aus XML
' gibt True zurück wenn feld gefunden
    
On Error GoTo Errorhandler

    If Fieldname = "" Then GoTo exithandler                         ' kein Fielname -> Fertig
    
    With Info
         GetLockedField = objTools.GetEditLockedField(szIniFilePath, .Name, Fieldname, bLocked, bEnabeld, bVisible)
    End With

exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "LoadDefaultValue", errNr, errDesc)
    Resume exithandler
End Function
                                                                    ' *****************************************
                                                                    ' Properties (Eigenschaften)
Public Property Get GetIniFile() As String
    GetIniFile = szIniFilePath
End Property

Public Property Get GetCaption() As String
    GetCaption = Info.EditCaption
End Property

Public Property Get GetID() As String
    GetID = Info.szID
End Property

Public Property Get GetTable() As String
    GetTable = Info.Table
End Property

Public Property Get GetImage() As Integer
    GetImage = Info.Image
End Property

Public Property Get GetDelFlagField() As String
    GetDelFlagField = Info.DelFlagField
End Property

Public Property Get GetRS() As ADODB.Recordset
    Set GetRS = Info.RS
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = Info.bDirty
End Property

Public Property Get IsDelete() As Boolean
'    IsDelete = Info.bDelete
End Property

Public Property Let SetDirty(Dirty As Boolean)
    Info.bDirty = Dirty
End Property

Public Property Get ShowDelFlag() As Boolean
    ShowDelFlag = Info.bShowDelFlag
End Property

Public Property Let SetShowDelFlag(ShowDel As Boolean)
    Info.bShowDelFlag = ShowDel
End Property

Public Property Get IsNew() As Boolean
    IsNew = Info.bNew
End Property

Public Property Get GetName() As String
    GetName = Info.Name
End Property

Public Property Get GetSQL() As String
    GetSQL = Info.SQL
End Property

Public Property Get GetWhere() As String
    GetWhere = Info.Where
End Property

Public Property Get GetFullSQL() As String
    GetFullSQL = Info.FullSQL
End Property

Public Property Get GetDefaultValue(Fieldname As String) As String
    GetDefaultValue = LoadDefaultValue(Fieldname)
End Property

Public Property Get IsLocked(Fieldname As String) As Boolean
    Dim bLocked As Boolean
    
    If GetLockedField(Fieldname, bLocked, False, False) Then
        IsLocked = bLocked
    End If
    
End Property

Public Property Get IsVisible(Fieldname As String) As Boolean
     Dim bVisible As Boolean
    
    If GetLockedField(Fieldname, False, False, bVisible) Then
        IsVisible = bVisible
    Else
        IsVisible = True
    End If
End Property

Public Property Get IsEnabled(Fieldname As String) As Boolean
    Dim bEnabled As Boolean
    
    If GetLockedField(Fieldname, False, bEnabled, False) Then
        IsEnabled = bEnabled
    Else
        IsEnabled = True
    End If
End Property

Public Property Get GetUnitRS() As ADODB.Recordset
'    Set GetUnitRS = RelUnitInfo.RelRS
End Property

Public Property Get GetSecureRS() As ADODB.Recordset
'    Set GetSecureRS = RelSecureInfo.RelRS
End Property

Public Property Get GetUnitImg() As Integer
'     GetUnitImg = RelUnitInfo.Image
End Property

Public Property Get GetUnitAltImg() As Integer
'     GetUnitAltImg = RelUnitInfo.AltImage
End Property

Public Property Get GetUnitAltImgField() As String
'     GetUnitAltImgField = RelUnitInfo.AltImgField
End Property

Public Property Get GetUnitAltImgValue() As String
'     GetUnitAltImgValue = RelUnitInfo.AltImgValue
End Property

Public Property Get GetSecureImg() As Integer
'     GetSecureImg = RelSecureInfo.Image
End Property

Public Property Get GetSecureAltImg() As Integer
'     GetSecureAltImg = RelSecureInfo.AltImage
End Property

Public Property Get GetSecureAltImgField() As String
'     GetSecureAltImgField = RelSecureInfo.AltImgField
End Property

Public Property Get GetSecureAltImgValue() As String
'     GetSecureAltImgValue = RelSecureInfo.AltImgValue
End Property


