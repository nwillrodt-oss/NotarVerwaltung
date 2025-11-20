Attribute VB_Name = "modFrmEdit"
Option Explicit                                                     ' Variaben Deklaration erzwingen
Private Const MODULNAME = "modFrmEdit"                              ' Modulname für Fehlerbehandlung

Private bHandleTabClickMSG As Boolean

Public Sub InitEditButtonMenue(frmEdit As Form, _
        Optional bShowSave As Boolean, _
        Optional bShowDelete As Boolean, _
        Optional bShowWord As Boolean)
    
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    frmEdit.cmdSave.Visible = bShowSave                             ' Speichern Schaltfläche un/sichtbar
    frmEdit.cmdDelete.Visible = bShowDelete                         ' Löschen Schaltfläche un/sichtbar
    frmEdit.cmdWord.Visible = bShowWord                             ' SAT Schaltfläche un/sichtbar
    Err.Clear                                                       ' Evtl. Error Clearen
End Sub

Public Sub EditFormLoad(frmEdit As Form, szRootkey As String)
    Dim CTL As Control                                              ' Control
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szDetails = "Formname: " & frmEdit.Name & vbCrLf & "ID: " & frmEdit.ID ' Details für Fehlerbehandlung
    'Call SetEditFormIcon(frmEdit, szRootKey)                       ' Form Icon setzen
    frmEdit.cmdUpdate.Enabled = False                               ' Button Übernehmen erstmmal disablen
    frmEdit.cmdSave.Enabled = False                                 ' Speichern Button erstmal disablen
    frmEdit.cmdDelete.Enabled = False
    frmEdit.Adodc1.Visible = False                                  ' Daten Verbindungs Control ausblenden
    For Each CTL In frmEdit.Controls
        Select Case UCase(Left(CTL.Name, 2))
        Case "LV"                                                   ' Bei Listviews Imagelisten setzen
            CTL.Icons = frmMain.ILTree
            CTL.SmallIcons = frmMain.ILTree
        Case "DT"                                                   ' bei DateTimePickern Now als standart wert wen neuer DS
            If frmEdit.IsNew Then
                CTL.Value = Now()
            End If
        Case Else
        
        End Select
    Next
    
exithandler:
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "EditFormLoad", errNr, errDesc, szDetails)
    Resume exithandler
End Sub

Public Sub EditFormUnload(frmEdit As Form)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call frmMain.CloseEditForm(frmEdit)                             ' Edit form schliessen
    Err.Clear                                                       ' Evtl. Error Clearen
End Sub

Public Function InitLVFrame(frmEdit As Form, _
        szRootkey As String, _
        szDetailKey As String, _
        cFrame As Frame, _
        LV As ListView, _
        LvTag As String) As ADODB.Recordset
    ' Positioniert Frame im Tabstrip & LV im Frame
    ' Füllt LV mit daten
    ' Lädt Spaltenbreite
    
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    Dim dbCon As Object                                             ' Datenbank verbindung
    
On Error GoTo Errorhandler


    szDetails = "ID: " & frmEdit.ID & vbCrLf & "bNew: " & frmEdit.bNew & vbCrLf
    cFrame.Visible = False                                          ' Frame erstmal Unsichtbar
    Call PosFrameAndListView(frmEdit, cFrame, True, LV)             ' Frame positionieren
    LV.Tag = LvTag                                                  ' Tag für ListView Setzen
    Set dbCon = frmEdit.GetDBConn                                   ' DB Verbindung aus Form holen
    Set InitLVFrame = ListLVByTag(LV, dbCon, LvTag, szDetailKey, _
            False, , True, frmEdit.bNew)                            ' Daten holen und Listitems setzen
    Call LoadColumnWidth(LV, szRootkey & "LV", True)                ' Spaltenbreite laden
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitLVFrame", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Public Function InitLVEditFrame(frmEdit As Form, RS As ADODB.Recordset, _
        cFrame As Frame, LV As ListView, LvTag As String, _
        szRootkey As String, lngImageIndex As Integer, _
        Optional AltImageIndex As Integer, _
        Optional AltImgField As String, _
        Optional AltImgValue As String) As ADODB.Recordset

    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    szDetails = "ID: " & frmEdit.ID & vbCrLf & "bNew: " & frmEdit.IsNew & vbCrLf
    cFrame.Visible = False                                          ' erstmal ausblenden
    Call PosFrameAndListView(frmEdit, cFrame, True, LV)             ' LV Im Frame Positionieren
    LV.Tag = LvTag                                                  ' LV Tag Setzen
    Call FillLVByRS(LV, "", RS, False, lngImageIndex, , False, True, , _
            AltImageIndex, AltImgField, AltImgValue)                ' Daten is LV einlesen
    
    Call LoadColumnWidth(LV, szRootkey & "LV", True)                ' Spaltenbreiten aus Reg laden
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitLVEditFrame", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Public Function RefreshFrame(frmEdit As Form, _
        EditFrame As Frame, _
        LV As ListView, _
        szRootkey As String, _
        szRelationName As String, _
        Optional bVisible As Boolean) As ADODB.Recordset
    Dim szSQL As String                                             ' SQL Statement
    Dim szWhere As String                                           ' Where PArt
    Dim lngImage As Integer                                         ' Image index
    Dim RS As ADODB.Recordset                                       ' Daten
    Dim dbCon As Object                                             ' Datenbank verbindung
    Dim szIniFilePath As String                                     ' Pfad zum XML file
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szIniFilePath = frmEdit.GetXMLPath()                            ' XML Pfad holen
    Call objTools.GetEditRelInfoFromXML(szIniFilePath, szRootkey, szRelationName, szSQL, _
                    szWhere, lngImage)                              ' Relations Informationen aus XML holen
    szWhere = szWhere & "'" & frmEdit.ID & "'"                      ' ID an WherePart anhängen
    szSQL = objSQLTools.AddWhereInFullSQL(szSQL, szWhere)           ' Kompletes SQL Statement erstellen
    Set dbCon = frmEdit.GetDBConn                                   ' DB Verbindung aus Form holen
    Set RS = dbCon.fillrs(szSQL, True)                              ' RS füllen
    Call InitLVEditFrame(frmEdit, RS, EditFrame, LV, LV.Tag, _
            szRootkey, lngImage)                                    ' Frame Positionieren und mit Daten Füllen
    EditFrame.Visible = bVisible                                    ' Frame Sichtbar ?
    Set RefreshFrame = RS                                           ' Daten zur weiteren Verwendung zurück
    
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "RefreshFrame", errNr, errDesc)
    Resume exithandler
End Function

Public Sub InitFrameInfo(frmEdit As Form, Optional bVisible As Boolean)
' Info Frame initialisieren
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    Dim mail As String                                              ' Email adresse
    Dim ThisDBCon As Object                                         ' DB Verbindung
On Error GoTo Errorhandler                                          ' Fehlerbehandlung akivieren
    szDetails = "Formname: " & frmEdit.Name & vbCrLf & "ID: " & frmEdit.ID
    Call frmEdit.FrameInfo.Move(frmEdit.GetFrameLeft, frmEdit.GetFrameTop, _
            frmEdit.GetFrameWidth, frmEdit.GetFrameHeigth)          ' Frame Positionieren
    frmEdit.FrameInfo.Visible = bVisible                            ' Sichtbar ?
    Set ThisDBCon = frmEdit.GetDBConn                               ' DB Verbindung holen
    If frmEdit.txtModifyFrom.Text <> "" Then                        ' Modify Feld ist nicht leer
        mail = objTools.checknull(ThisDBCon.GetValueFromSQL( _
                "SELECT EMAIL001 FROM USER001 WHERE USERNAME001 = '" _
                & frmEdit.txtModifyFrom.Text & "'"), "")            ' Mitarbeiter email holen und in Tag schreiben
        If mail <> "" Then frmEdit.txtModifyFrom.Tag = mail
        If frmEdit.txtModifyFrom.Tag <> "" Then                     ' Wenn tag gefüllt
            frmEdit.txtModifyFrom.Font.Underline = True             ' Unterstreichen
            frmEdit.txtModifyFrom.ForeColor = -2147483635           ' Farbe setzen
        End If
    End If
    If frmEdit.txtCreateFrom.Text <> "" Then                        ' Create Feld ist nicht leer
        frmEdit.txtCreateFrom.Tag = ThisDBCon.GetValueFromSQL( _
                "SELECT EMAIL001 FROM USER001 WHERE USERNAME001 = '" _
                & frmEdit.txtCreateFrom.Text & "'")                 ' Mitarbeiter email holen und in Tag schreiben
        If frmEdit.txtCreateFrom.Tag <> "" Then                     ' Wenn tag gefüllt
            frmEdit.txtCreateFrom.Font.Underline = True             ' Unterstreichen
            frmEdit.txtCreateFrom.ForeColor = -2147483635           ' Farbe setzen
        End If
    End If
exithandler:
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitFrameInfo", errNr, errDesc, szDetails)
    Resume exithandler
End Sub

Public Sub AskUserAboutThisDS(CTL As Control, Optional szSubject As String)
    Dim UserEmail As String                                         ' eMail adresse des Users
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    If CTL.Tag <> "" Then                                           ' TAg nicht leer
        UserEmail = CTL.Tag                                         ' Mailadresse aus tag lesen
        UserEmail = Trim(UserEmail)                                 ' Sicherheitshalber trimmen
        Call objObjectBag.ShowNewMail(UserEmail, szSubject)
    End If
    'Call objObjectBag.ShowNewMail(szMailAdress As String, szSubject As String, Optional szMailText As String, Optional szDateianhang As String
    
exithandler:
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "AskUserAboutThisDS", errNr, errDesc)
    Resume exithandler
End Sub

Public Sub RefreshRelField(frmEdit As Form, _
        txtRelCtl As Control, _
        txtIDCtl As Control, _
        szSQL As String, _
        szWhere As String, _
        Optional bLocked As Boolean)
' Gedacht für Como felder die eine ID in Text feld schreiben
' Der inhalt des Testfeldes wird mit dim inhalt des Combofeldes abgeglichen
    Dim dbCon As Object                                             ' Datenbank Verbindung
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    If szSQL = "" Then Exit Sub                                     ' Kein SQL -> Fertig
    If txtIDCtl.Text <> "" Then                                     ' Wenn ID Feld gefüllt
        Set dbCon = frmEdit.GetDBConn                               ' DB Verbindung Holen
        szSQL = objSQLTools.AddWhereInFullSQL(szSQL, szWhere _
                & "'" & txtIDCtl.Text & "'")                        ' Where in SQL Statement einsetzen
        txtRelCtl.Text = dbCon.GetValueFromSQL(szSQL)               ' Abfrage ausführen & Erg. in txtRelCtl setzen
        'txtIDCtl.Locked = bLocked                                  ' Feld Locken
    End If
    
exithandler:

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "RefreshRelField", errNr, errDesc)
    Resume exithandler
End Sub

Public Function InitAdoDC(frmEdit As Form, dbCon As Object, szSQL As String, szWhere As String)
    ' Initialisiert ADODC mit SQL Statement & evtl. Where Part
On Error GoTo Errorhandler                                          ' Fehlerbehandlung akticieren
    If szSQL = "" Then GoTo exithandler                             ' Kein SQL -> Fertig
    If Not frmEdit.IsNew And szWhere <> "" Then _
            szSQL = objSQLTools.AddWhereInFullSQL(szSQL, szWhere)   ' Where PArt in SQL integierrn
On Error Resume Next                                                ' Fehlerbehandlung deakt.
    frmEdit.Adodc1.fCancelDisplay = True                            ' Keine Fehler anzeigen
    frmEdit.Adodc1.ConnectionString = dbCon.GetConnectString        ' Connectstring setzen
    Err.Clear                                                       ' Evtl. Fehler Clearen
On Error GoTo Errorhandler                                          ' Fehlerbehandlung wieder aktivieren
    frmEdit.Adodc1.CommandType = adCmdText                          ' Querytyp = SQL
    frmEdit.Adodc1.RecordSource = szSQL                             ' SQL Statement setzen
    frmEdit.Adodc1.Refresh                                          ' Daten holen
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitAdoDC", errNr, errDesc)
    Resume exithandler
End Function

Public Function GetRelLVSelectedID(LV As ListView) As String
    Dim ItemKey As String                                           ' LVItem Key
    Dim szItemKeyArray() As String
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    ItemKey = LV.SelectedItem.Key                                   ' Selected Item.Key ermitteln
    If ItemKey = "" Then GoTo exithandler                           ' Wenn Item Key Vorhanden
    szItemKeyArray = Split(ItemKey, TV_KEY_SEP)                     ' ItemKey In Array aufspalten
    GetRelLVSelectedID = szItemKeyArray(UBound(szItemKeyArray))     ' Letzen Array Inalt ist ID
exithandler:
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Public Sub CheckUpdate(frmEdit As Form)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    frmEdit.cmdUpdate.Enabled = frmEdit.IsDirty                     ' cmdUpdate (Übernehmen) auf Enabled = Dirty setzen
    frmEdit.cmdSave.Enabled = frmEdit.IsDirty                       ' cmdSave auf Enabled = Dirty setzen
    frmEdit.cmdDelete.Enabled = Not frmEdit.IsNew                   ' cmdDelete auf Enabled = Not New setzen
    frmEdit.cmdWord.Enabled = Not frmEdit.IsNew                     ' cmdDelete auf Enabled = Not New setzen
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Public Sub FormatDTPicker(frmEdit As Form, DTCtl As DTPicker, Datum As Date)
Const sFormat = "dd.mm.yyyy"
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    DTCtl.Value = Format(Datum, sFormat)                            ' Standart DAtumsformat setzen
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Public Sub FirstCharUp(frmEdit As Form, CTL As Control)
    ' Setzt 1. Zeichen Groß
    Dim OldText As String                                           ' Alter ctl Text
    Dim FirstChar As String                                         ' 1. Zeichen
    Dim Rest As String                                              ' Alle anderen Zeichen
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    OldText = CTL.Text                                              ' Text holen
    If OldText <> "" Then
        FirstChar = Left(OldText, 1)                                ' 1. Zeichen holen
        Rest = Right(OldText, Len(OldText) - 1)                     ' Rest Holen
        CTL.Text = UCase(FirstChar) & Rest                          ' Neuen Text setzen
        CTL.SelStart = Len(CTL.Text)                                ' Cursor setzen
    End If
    Err.Clear                                                       ' Evtl. Error Clearen
End Sub

Public Sub StandartTextChange(frmEdit As Form, ctlTXT As TextBox)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call ShowTextLenInToolTip(frmEdit, ctlTXT)                      ' Zeichen länge in Tooltip zeigen
    Call CheckTextLen(frmEdit, ctlTXT)                              ' Prüfen ob eingabe zulang
    frmEdit.SetDirty = True                                         ' Dirty Setzen
    Call CheckUpdate(frmEdit)                                       ' Übernehmen enablen
    Err.Clear                                                       ' Evtl. Error Clearen
End Sub

Public Sub HiglightMustFields(frmEdit As Form, Optional bNoHighligt As Boolean)
' Hinterlegt Pflichtfelder (indexfelder) mit andere Farbe
    Dim CTL As Control                                              ' Control
    Dim bMustCtl As Boolean                                         ' Ist Pflichtfeld

On Error GoTo Errorhandler

    For Each CTL In frmEdit.Controls                                ' Alle Controls Durchlaufen
        If UCase(Left(CTL.Name, 3)) = "TXT" Then                    ' TextFeld ?
            If CTL.DataField <> "" Then                             ' Datafield angegeben ?
                If Not CBool(frmEdit.Adodc1.Recordset.Fields(CTL.DataField).Attributes _
                            And adFldIsNullable) Then               ' Ist Pflichtfeld ?
                    If CTL.Name <> "txtID" And CTL.Name <> "txtCreate" Then ' Ausnahmen in Infoframe
                        If bNoHighligt Then                         ' Soll An oder Abgeschaltet werden ?
                            Call HiglightMustField(frmEdit, CTL, True) ' Hervorhebung entfernen
                        Else
                            Call HiglightMustField(frmEdit, CTL)    ' Hervorhebung Aktivieren
                            'ctl.BackColor = &HE3DCFA
                        End If
                    End If
                End If
            End If
        End If
    
        If UCase(Left(CTL.Name, 3)) = "CMB" Then                    ' Combo Feld ?
            If CTL.DataField <> "" Then                             ' Datafield angegeben ?
                If Not CBool(frmEdit.Adodc1.Recordset.Fields(CTL.DataField).Attributes _
                            And adFldIsNullable) Then               ' Ist Pflichtfeld ?
                    If bNoHighligt Then                             ' Soll An oder Abgeschaltet werden ?
                        Call HiglightMustField(frmEdit, CTL, True)  ' Hervorhebung Aktivieren
                    Else
                        Call HiglightMustField(frmEdit, CTL)        ' Hervorhebung Aktivieren
                        'ctl.BackColor = &HE3DCFA
                    End If
                End If
            End If
        End If
    Next                                                            ' Nächstes Control

Exit Sub
Errorhandler:
    Stop
End Sub

Public Sub HiglightCurentField(frmEdit As Form, CTL As Control, Optional bDeHighlight As Boolean)
On Error Resume Next                                                ' Feherbehandlung deaktivieren
    If bDeHighlight Then
        CTL.BackColor = &H80000005                                  ' Standartfarbe setzten
        Call frmEdit.HiglightThisMustFields(Not (frmEdit.IsNew Or frmEdit.IsDirty))
    Else
        CTL.BackColor = 12648447                                    ' Farbe setzten (Hellgelb)
    End If
    Err.Clear                                                       ' Evtl. Error Clearen
End Sub

Public Sub HiglightMustField(frmEdit As Form, CTL As Control, Optional bDeHighlight As Boolean)
On Error Resume Next                                                ' Feherbehandlung deaktivieren
    If bDeHighlight Then
        CTL.BackColor = &H80000005                                  ' Standartfarbe setzten
    Else
        CTL.BackColor = &HE3DCFA                                    ' Farbe setzten (Hellrot / Rosa)
    End If
    Err.Clear                                                       ' Evtl. Error Clearen
End Sub

Public Sub NoHiglight(frmEdit As Form, CTL As Control)
On Error Resume Next                                                ' Feherbehandlung deaktivieren
    CTL.BackColor = &H80000005                                      ' Standartfarbe setzten
    Err.Clear                                                       ' Evtl. Error Clearen
End Sub

Public Sub CheckTextLen(frmEdit As Form, txtCTL As TextBox)
    ' Prüft ob Maximale Textlänge überschritten
    Dim MaxLen As Integer                                           ' Max Länge

On Error Resume Next                                                ' Fehlerbehandlung deaktivieren

    MaxLen = frmEdit.Adodc1.Recordset.Fields( _
            txtCTL.DataField).DefinedSize                           ' Max Länge ermitteln
    If frmEdit.Adodc1.Recordset.Fields(txtCTL.DataField _
            ).Type = 72 Then MaxLen = 38                            ' Guid hat andere länge
    If Len(txtCTL.Text) > MaxLen And MaxLen <> 0 Then               ' Wenn Überschritten
        txtCTL.Text = Left(txtCTL.Text, MaxLen)                     ' Abschneiden
        txtCTL.SelStart = Len(txtCTL.Text)                          ' Cursor setzen
    End If
    Err.Clear                                                       ' Evtl. Error Clearen
End Sub

Public Function ValidateTxtFieldIsYear(txtCTL As Control) As Boolean
    ' Prüft ob Feldinhalt (txt o. Combo) eine Jahreszahl ist und baut meldung zusammen
    ' Gibt False zurück wenn alles OK (ist Jahr)
    ' Bzw True wenn kein Jahr
On Error GoTo Errorhandler

    If txtCTL.Text = "" Then GoTo exithandler                       ' Wenn kein Text -> fertig
    If IsNumeric(txtCTL.Text) Then                                  ' Wenn nur zahlen
        
        If Len(txtCTL.Text) = 4 Then                                ' 4 stellig
            
        ElseIf Len(txtCTL.Text) = 2 Then                            ' 2 Stellig
            If CLng(txtCTL.Text) < 30 Then                          ' Wenn <30
                txtCTL.Text = "19" & txtCTL.Text                    ' Dann 21. Jahrhundert
            Else                                                    ' Sonst
                txtCTL.Text = "20" & txtCTL.Text                    ' 20. Jahrhundert
            End If
        Else
            ValidateTxtFieldIsYear = True
        End If
    Else
        ValidateTxtFieldIsYear = True
    End If
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ValidatetxtFieldIsYear", errNr, errDesc)
    Resume exithandler
End Function

Public Function ValidateTxtFieldOnEmpty(txtCTL As Control, szFieldlable As String, _
        ByRef szMSG As String, ByRef FocusCTL As Control) As Boolean
    ' Prüft ob Feldinhalt (txt o. Combo) leer ist und baut meldung zusammen
    ' Gibt FAllse zurück wenn alles OK (nicht Leer)
    ' Bzw True wenn das Feld Leer ist
On Error GoTo Errorhandler

    If txtCTL.Text = "" Then                                        ' Wenn kein Text vorhanden
        szMSG = "Das Feld '" & szFieldlable _
                & "' darf nicht leer bleiben!"                      ' Meldungstext setzen
        Set FocusCTL = txtCTL                                       ' Focus Control setzen
        ValidateTxtFieldOnEmpty = True                              ' Validierung Gescheitert zurück
    End If
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ValidateTxtFieldOnEmpty", errNr, errDesc)
    Resume exithandler
End Function

Public Sub ShowTextLenInToolTip(frmEdit As Form, txtCTL As TextBox)
    Dim MaxLen As Integer                                           ' Max Feld länge
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    MaxLen = frmEdit.Adodc1.Recordset.Fields( _
            txtCTL.DataField).DefinedSize                           ' Max Feldlänge aus RS ermitteln
    txtCTL.ToolTipText = MaxLen & " Zeichen"                        ' ToolTip setzen
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Public Sub GetLockedControls(frmEdit As Form)
'Public Sub GetLockedControls(frmEdit As Form, ThisEditObj As Object)
    Dim CTL As Control                                              ' Akt Control
    Dim bEnabled As Boolean                                         ' Ctl enablen
    Dim bLocked As Boolean                                          ' Ctl Locken
    Dim CtlNamePre As String                                        ' Ctl Namens Prefix (txt oder cmb)
    Dim bLockWhenNew As Boolean                                     ' Soll das FEld auch für Neue DS gelocked sein
    Dim lngMultiline As Integer                                     ' Mehrere Zeilen
    Dim szDetails As String                                         ' Details für fehlerbehandlung
    Dim lngThisCtlNewHeight As Integer                              ' Neue Control Höhe
    Dim lngCtlTopDiff As Integer
    Dim lngStartTop As Integer
    Dim szOldCtlText As String
    
On Error GoTo Errorhandler

    For Each CTL In frmEdit.Controls                                ' Alle Controls duchlaufen
        szOldCtlText = ""
        bEnabled = True                                             ' Standartwert Enabled = True
        bLocked = False                                             ' Standartwert Locked = False
        bLockWhenNew = False                                        ' Standartwert
        CtlNamePre = UCase(Left(CTL.Name, 3))                       ' Namen Prefix des controls holen
        szDetails = "Controlname: " & CTL.Name                      ' Control name für Fehlerbehandlung
        If CtlNamePre = "TXT" Or CtlNamePre = "CMB" Then            ' Nur für txt oder cmb
            If objTools.GetEditLockedField(frmEdit.GetXMLPath, frmEdit.GetRootkey, _
                    CTL.Name, bLocked, bEnabled, , _
                    lngMultiline, bLockWhenNew) Then                ' Feld Infos aus XML Auslesen
On Error Resume Next                                                ' Um aufsehen zu vermeiden
                If frmEdit.IsNew And Not bLockWhenNew Then          ' DS Neu &  bNoLockWhenNew = True
                    CTL.Locked = False                              ' Nicht locken
                    CTL.Enabled = True                              ' Enablen
                Else
                    CTL.Locked = bLocked                            ' Control auf bLeocked setzen
                    CTL.Enabled = bEnabled                          ' Control auf bEnabeld setzen
                End If
                If lngMultiline > 1 Then                            ' Wenn Mehere Zeilen
                    szOldCtlText = CTL.Text
                    CTL.MultiLine = True
                    If CTL.Height <= 315 Then
                        lngThisCtlNewHeight = CTL.Height * lngMultiline
                        lngCtlTopDiff = lngCtlTopDiff + (lngThisCtlNewHeight - CTL.Height)
                        lngStartTop = CTL.Top
                        CTL.Height = lngThisCtlNewHeight
                    End If
                    CTL.ScrollBars = 0
                    CTL.Text = szOldCtlText
                    CTL.Refresh
                End If
                Err.Clear                                           ' Evtl. Fehler zurücksetzen
On Error GoTo Errorhandler                                          ' Fehlerbehandlung wieder aktivieren
            End If
        End If
    Next                                                            ' Nächstes Control

    If lngCtlTopDiff <> 0 And lngStartTop <> 0 Then                 ' Control grösse geändert
        For Each CTL In frmEdit.Controls                            ' Nochmal alle Controls duchlaufen
            CtlNamePre = UCase(Left(CTL.Name, 3))                   ' Namen Prefix des controls holen
            If CtlNamePre = "TXT" Or CtlNamePre = "CMB" _
                    Or CtlNamePre = "LBL" Then                      ' Nur für lbl, txt oder cmb
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
                If CTL.Top > lngStartTop Then                       ' Akt. CTL Top grösser Start Top
                    CTL.Top = CTL.Top + lngCtlTopDiff               ' CTL Top um Top Diff. erweitern
                End If
                Err.Clear                                           ' Evtl. Error clearen
On Error GoTo Errorhandler                                          ' Fehlerbehandlung wieder aktivieren
            End If
        Next                                                        ' Nächstes Control
    End If
exithandler:
On Error Resume Next
    frmEdit.Refresh                                                 ' Form aktualisieren

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetLockedControls", errNr, errDesc, szDetails)
    Resume exithandler
End Sub

Public Function HandleLVkmnuNew(frmEdit As Form, szCaption As String) As Boolean

     Dim szMSG As String                                            ' Meldungs Text

On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    If Not frmEdit.IsNew Then GoTo exithandler                      ' Wenn nicht bNew dan raus

    If szCaption = "" Then szCaption = "diese Aktion ausführen"     ' Teil der Meldung aus Parameter Konfigurieren
    szMSG = "Sie müssen den Datensatz speichern bevor Sie " _
            & szCaption & " können"                                 ' Rest Meldung Konfigurieren
    
    Call objError.ShowErrMsg(szMSG, vbInformation, "Hinweis")       ' Meldung ausgeben
    HandleLVkmnuNew = True                                          ' OK zurück liefern
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "HandleLVkmnuNew", errNr, errDesc)
    Resume exithandler
End Function

Public Function GetTabStrimClientPos(TabCtl As TabStrip, lngTop As Single, _
        lngLeft As Single, lngHeight As Single, lngWidth As Single)
' Ermitelt Frame positionen aus TabStrip Ctl
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    lngTop = TabCtl.ClientTop + 100
    lngLeft = 120
    lngWidth = TabCtl.ClientWidth - 120
    lngHeight = TabCtl.ClientHeight - 200
    Err.Clear                                                       ' Evtl. Error clearen
End Function

Public Function HandleTabClickNew(frmEdit As Form, CtlTab As TabStrip) As Boolean
' Behandelt TAb Click bei neuen Dirty DS (nur 1. TAb zulässig)
    Dim szMSG As String                                             ' Meldungs Text
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    HandleTabClickNew = True
    If Not frmEdit.IsNew Then GoTo exithandler                      ' Wenn nicht bNew dan raus
     HandleTabClickNew = False
    szMSG = "Sie müssen den Datensatz speichern bevor Sie die " & CtlTab.SelectedItem.Caption _
            & " bearbeiten können"                                  ' Meldung Konfigurieren
    If CtlTab.SelectedItem.Index > 1 Then                           ' 1. Tab ist OK
        Call objError.ShowErrMsg(szMSG, vbInformation, "Hinweis", _
                False, "", frmEdit)                                 ' Meldung ausgeben
        HandleTabClickNew = False
'        CtlTab.Tabs(1).Selected = True                              ' 1. Tab Selecten
    End If
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "HandleTabClickNew", errNr, errDesc)
    Resume exithandler
End Function

Public Sub Hover(frmEdit As Form, CTL As Control, Optional bOn As Boolean)
    If bOn Then
        CTL.FontUnderline = True                                    ' Text unterstreichen
        CTL.ForeColor = -2147483635                                 ' Textfarbe dunkelblau
    Else
        CTL.FontUnderline = False                                   ' Text nicht unterstreichen
        CTL.ForeColor = -2147483640                                 ' Standart testfarbe setzten
    End If
End Sub

Public Sub HandleEditLVDoubleClick(frmEdit As Form, LV As ListView, _
        Optional parentform As Form, _
        Optional bDialog As Boolean)
    Dim szRootkey As String                                         ' Was für ein DS soll angezeigt werden
    Dim szDetailKey As String                                       ' Datensatz ID
    Dim szItemKeyArray() As String                                  ' Array mit Item Keys
    Dim szTmp As String                                             ' ItemKey als string
On Error Resume Next                                                ' Fehlerbehandlung erstmal deakt.
    szTmp = LV.SelectedItem.Key                                     ' Selectet Item ermitteln
    Err.Clear                                                       ' Evtl. err clearen
On Error GoTo Errorhandler                                          ' Fehlerbehandlung wieder aktivieren
    If szTmp = "" Then GoTo exithandler                             ' Wenn Item Key ="" -> Fertig
    szItemKeyArray = Split(szTmp, TV_KEY_SEP)                       ' Item Key aufspalten
    szDetailKey = szItemKeyArray(UBound(szItemKeyArray))            ' DS Id ermitteln (letztes SubItem ausgeblendet)
    szRootkey = szItemKeyArray(UBound(szItemKeyArray) - 1)          ' RootKey steht davor
    If szRootkey <> "" And szDetailKey <> "" Then                   ' Wenn Rootkey und ds ID vorhanden
        Call EditDS(szRootkey, szDetailKey, bDialog)                ' DS zum bearbeiten öffnen
    End If
    
exithandler:
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "HandleEditLVDoubleClick", errNr, errDesc)
    Resume exithandler
End Sub

Public Function HandleKeyDownEdit(frmEdit As Form, KeyCode As Integer, Shift As Integer)
' Behandelt Key DoWn für Shortcuts
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If KeyCode = 27 And Shift = 0 Then                              ' ESC
        Unload frmEdit                                              ' Form Schliessen ohne speichern
    End If
    If KeyCode = 83 And Shift = 2 Then                              ' STGR + S
        Call frmEdit.cmdUpdate_Click                                ' Form Speichern
    End If
    Err.Clear                                                       ' Evtl. Error Clearen
End Function

Public Function EditValidateForm(frmEdit As Form) As Boolean

    Dim szDetails As String                                         ' Details fürfehlerbehandlung
    Dim adoField As ADODB.Field                                     ' Akt Feld
     
On Error GoTo Errorhandler                                          ' Fehler behandlung aktivieren

    szDetails = "Formname: " & frmEdit.Name & vbCrLf & "ID: " & frmEdit.ID
    
    If frmEdit.IsNew Then                                           ' Neuer DS -> Insert
        For Each adoField In frmEdit.Adodc1.Recordset.Fields        ' Für jedes Feld im RS
            If Not CBool(adoField.Attributes And _
                    adFldIsNullable) Then                           ' Wenn Feld kein NULL zulässt
                If adoField.Value = "" Then                         ' Und Feldvalue =""
                       ' Debug.Print "Feld " & adoField.Name & " darf nicht leer sein!"
                End If
            End If
        Next adoField                                               ' Nächstes Feld
    Else                                                            ' Update
    
    End If
    
    EditValidateForm = True                                         ' Validierung erfolgreich
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "EditValidateForm", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Public Sub GetDefaultValues(frmEdit As Form, RootKey As String, szIniFilePath As String)
' Holt default werte aus XML
    Dim szDefValue As String                                        ' Default Value
    Dim szDefValSQL As String                                       ' SQL Statement das Value holt
    Dim i As Integer                                                ' Counter
    Dim szFieldName As String                                       ' aktueller Feldname
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    For i = 0 To frmEdit.Adodc1.Recordset.Fields.Count - 1          ' Für jedes Feld
        szFieldName = frmEdit.Adodc1.Recordset.Fields(i).Name       ' feldname ermitteln
        If szFieldName = "" Then GoTo exithandler                   ' kein Fielname -> Fertig
    
        szDefValue = objTools.GetEditDefaultValue(szIniFilePath, RootKey, szFieldName, szDefValSQL)
        If szDefValSQL <> "" Then szDefValue = frmEdit.GetDBConn.GetValueFromSQL(szDefValSQL)
        If szDefValue <> "" Then
            If szDefValue <> "" Then frmEdit.Adodc1.Recordset.Fields(i).Value = szDefValue
            szDefValue = ""
            szDefValSQL = ""
        End If
    Next i

exithandler:
On Error Resume Next
    frmEdit.Refresh                                                 ' Form aktualisieren

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetDefaultValues", errNr, errDesc)
    Resume exithandler
End Sub

Public Function UpdateEditForm(frmEdit As Form, Optional szRootkey As String) As Boolean
' Speichert DS aus Edit form
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    Dim bNewBeforSave As Boolean                                    ' DS Ist vorm speichern Neu
    Dim fMain As Form                                               ' Hauptform
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szDetails = "Formname: " & frmEdit.Name & vbCrLf & "ID: " & frmEdit.ID ' Details für Fehlerbehandlung
    If Not frmEdit.IsDirty Then GoTo exithandler                    ' Keine Änderungen -> Raus
    If Not EditValidateForm(frmEdit) Then GoTo exithandler          ' Prüfe ob Änderungen Zulässig
    bNewBeforSave = frmEdit.IsNew                                   ' Ist Neuer DS
    If bNewBeforSave Then                                           ' Bei neuen Datensatz
        frmEdit.txtCreateFrom.Text = objObjectBag.GetUserName       ' Benutzer eintragen
        ' Erstellt datum wird über standartwert in der tabelle geregelt
    End If
    frmEdit.txtModify.Text = Now()                                  ' Änderungsdatum eintragen
    frmEdit.txtModifyFrom.Text = objObjectBag.GetUserName           ' Benutzer eintragen
    If bNewBeforSave Then                                           ' Neuer DS -> Insert
        frmEdit.Adodc1.Recordset.Update
'        frmEdit.Adodc1.Recordset.Insert
    Else                                                            ' Update
        frmEdit.Adodc1.Recordset.Update
    End If
    frmEdit.SetDirty = False                                        ' Damit ist das Form nicht mehr Dirty
    Call CheckUpdate(frmEdit)                                       ' Übernehmen button disablen
    Set fMain = objObjectBag.getMainForm                            ' MainForm Holen
    If Not fMain Is Nothing Then                                    ' MainForm vorhanden
        If bNewBeforSave Then                                       ' neuer DS
            Call fMain.RefreshTreeView                              ' TV & LV Aktualisieren
        Else                                                        ' nur geändert
            Call fMain.RefreshListView                              ' Nur LV Aktualisieren
        End If
    End If

    UpdateEditForm = True                                           ' Erfolg zurück
    
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "UpdateEditForm", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Public Function PosFrameAndListView(frmEdit As Form, CurFrame As Frame, _
            bWithBorder As Boolean, Optional LV As ListView)
'Positioniert Frame / ListView kombination
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szDetails = "Formname: " & frmEdit.Name & vbCrLf & " ID: " & frmEdit.ID & vbCrLf _
            & " Frame: " & CurFrame.Name
    If bWithBorder Then                                             ' Soll Frame ramen angezeigt werden ?
        CurFrame.BorderStyle = vbFixedSingle                        ' Frame Rahmen anzeigen
    Else
        CurFrame.BorderStyle = vbBSNone                             ' Kein Rahmen anzeigen
    End If
    Call CurFrame.Move(frmEdit.GetFrameLeft, frmEdit.GetFrameTop, _
                    frmEdit.GetFrameWidth, frmEdit.GetFrameHeigth)  ' Frame positionieren
    If Not LV Is Nothing Then                                       ' LV Angegeben
        Call FillFrameWithLV(CurFrame, LV)                          ' LV an Frame angleichen
        Call LV.Move(120, 240, frmEdit.GetFrameWidth - 240, frmEdit.GetFrameHeigth - 360)
        Call InitDefaultListViewResult(LV)                          ' LV Initialisieren
    End If
    
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "PosFrameAndListView", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Public Function SetEditFormCaption(frmEdit As Form, szRootkey As String, _
                    Optional szAddCaption As String)
    Dim szDetails As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szDetails = "Formname: " & frmEdit.Name & vbCrLf & "ID: " & frmEdit.ID
    If frmEdit.IsNew Then                                           ' Wenn DS Neu
        frmEdit.Caption = szRootkey & ": Neuer Datensatz"           ' Caption Setzen
    Else                                                            ' Sonst
        frmEdit.Caption = szRootkey & ": " & szAddCaption & " -  Bearbeiten"  ' Caption Setzen
    End If

exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SetEditFormCaption", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Public Function DelRelationinLV(frmEdit As Form, _
        RootKey As String, _
        DBConn As Object, _
        LV As ListView, _
        RS As ADODB.Recordset, _
        RelIDField As String, _
        RelationTable As String)
    Dim RelID As String         ' Relation ID
    Dim szDetails As String     ' Zusatz für Fehlermeldung
    Dim i As Integer            ' Counter
    Dim szSQL As String         ' SQL Statement
On Error GoTo Errorhandler
    szDetails = "Formname: " & frmEdit.Name & vbCrLf & "ID: " & frmEdit.ID
    RelID = GetRelLVSelectedID(LV)
    If RelID = "" Then GoTo exithandler
    If LV.SelectedItem = "" Then GoTo exithandler  ' Nur wenn ein DS ausgewählt
    RS.MoveFirst
    RS.Find (LV.ColumnHeaders(1).Text & " = '" & LV.SelectedItem & "'")

    If Not RS.EOF Or Not RS.BOF Then
        RelID = RS.Fields(RelIDField).Value

        If RelID = "" Then GoTo exithandler

        szSQL = "DELETE  FROM " & RelationTable & " WHERE " & RelIDField & "='" & RelID & "'"
        Call DBConn.execSql(szSQL)

    End If

exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "DelRelationinLV", errNr, errDesc, szDetails)
    Resume exithandler
End Function
'Public Function FillRelLV(frmEdit As Form, _
'        dbCon As Object, _
'        IniPath As String, _
'        LV As ListView, _
'        szRootkey As String, _
'        szRelKey As String) As ADODB.Recordset
'
'    Dim szSQLMain As String
'    Dim szWhere As String
'    Dim szImgIndex As String
'
'On Error GoTo Errorhandler
'
'    If IniPath = "" Then GoTo exithandler
'    If szRelKey = "" Then GoTo exithandler
'
'    szSQLMain = objTools.GetINIValue(IniPath, INI_RELATIONS, szRootkey & szRelKey)
'    If szSQLMain = "" Then GoTo exithandler
'
'    szWhere = objTools.GetINIValue(IniPath, INI_RELATIONS, "WHERE" & szRootkey & szRelKey)
'    szImgIndex = objTools.GetINIValue(IniPath, INI_IMAGE, szRootkey & szRelKey)
'    'If szWhere <> "" And Not frmEdit.bNew Then
'    If szWhere <> "" Then
'        If frmEdit.ID = "" Then
'            szWhere = szWhere & "'00000000-0000-0000-0000-000000000000'"
'        Else
'            szWhere = szWhere & "'" & frmEdit.ID & "'"  '"CAST('" & frmEdit.ID & "' as uniqueidentifier)"
'            'szWhere = szWhere & "CAST('" & frmEdit.ID & "' as uniqueidentifier)"
'        End If
'
'
'        szSQLMain = objSQLTools.AddWhereInFullSQL(szSQLMain, szWhere)
'    End If
'    If szImgIndex = "" Then szImgIndex = "0"
'    Set FillRelLV = FillLVBySQL(LV, szSQLMain, dbCon, , CLng(szImgIndex))
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
'    Call objError.Errorhandler(MODULNAME, "FillRelLV", errNr, errDesc)
'    Resume exithandler
'End Function

'Public Function SetRelationinLV(frmEdit As Form, _
'        RootKey As String, _
'        SearchField As String, _
'        DBConn As Object, _
'        LV As ListView, _
'        RS As ADODB.Recordset, _
'        EntityIDField As String, _
'        RelationIDField As String)
'
'    Dim RelID As String         ' Relation ID
'    'Dim szSQLInsert As String
'    Dim szDetails As String     ' Zusatz für Fehlermeldung
'    Dim i As Integer            ' Counter
'    Dim KeyArray() As String
'    Dim bIsIn As Boolean
'
'On Error GoTo Errorhandler
'
'    szDetails = "Formname: " & frmEdit.Name & vbCrLf & "ID: " & frmEdit.ID
'
'    RelID = ShowSearch(DBConn, RootKey, SearchField)
'    If RelID = "" Then GoTo exithandler
'        If LV.ListItems.Count > 0 Then
'            For i = 1 To LV.ListItems.Count
'                KeyArray = Split(LV.ListItems(i).Key, TV_KEY_SEP)
'                If KeyArray(UBound(KeyArray)) = frmEdit.ID Then
'                'If LV.ListItems(i).Text = frmEdit.ID Then
'                    'schon drin
'                    LV.ListItems(i).Selected = True
'                    bIsIn = True
'                    Exit For
'                Else
'                    ' insert
'                   bIsIn = False
'                End If
'            Next i
'            If Not bIsIn Then
'                RS.AddNew
'                RS.Fields(RelationIDField).Value = RelID
'                RS.Fields(EntityIDField).Value = frmEdit.ID
'                RS.Update
'            End If
'        Else
'            RS.AddNew
'            RS.Fields(RelationIDField).Value = RelID
'            RS.Fields(EntityIDField).Value = frmEdit.ID
'            RS.Update
'        End If
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
'    Call objError.Errorhandler(MODULNAME, "SetRelationinLV", errNr, errDesc, szDetails)
'    Resume exithandler
'End Function




 


