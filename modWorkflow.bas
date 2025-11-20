Attribute VB_Name = "modWorkflow"
Option Explicit                                                     ' Variaben Deklaration erzwingen
Private Const MODULNAME = "modWorkflow"                             ' Modulname für Fehlerbehandlung

                                                                    ' *****************************************
                                                                    ' Workflow Gehampel (Neu)
Public Function WorkflowNextStep(frmEdit As Form, szWorkFlowField As String, _
        Optional bAskFor As Boolean) As Boolean

    Dim szAktStep As String
    Dim szNextStep As String
    
On Error GoTo Errorhandler

    szAktStep = frmEdit.GetCurrentStep
    szNextStep = frmEdit.GetNextStep
    
    If ValidateWorkflowCurrentStep(frmEdit) Then                     ' Wenn Evtl. Bedingungen erfolgreich
        If DoWorkflowAction(frmEdit, bAskFor) Then                           ' Evtl. Aktion ausführen
                                                                    ' Dann nächster Schritt
            frmEdit.Adodc1.Recordset.Fields(szWorkFlowField).Value = szNextStep  ' Wert ins RS
            frmEdit.Adodc1.Recordset.Update                         ' Speichern
            Call frmEdit.RefreshFrameWorkflow(frmEdit.FrameWorkflow.Visible)     ' Ansicht Aktualisieren
            WorkflowNextStep = True
        Else
            ' Fehler bei aktion
            
        End If
    Else                                                            ' Sonst Meldung warum nicht
        Call ShowWorkflowValidateMSG(frmEdit)                       ' Meldung
        WorkflowNextStep = False
    End If
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "WorkflowNextStep", errNr, errDesc)
    Resume exithandler
End Function

Public Function ValidateWorkflowCurrentStep(frmEdit As Form) As Boolean

    Dim DBConn As Object                                            ' Aktuelle DB Verbindung
    Dim szAktStep As String                                         ' Akt Schritt
    Dim szSQL As String                                             ' SQL Statement
    Dim szCondition As String                                       ' Prüfbedingung
    Dim szTabName As String                                         ' zu prüfende Tabelle
    Dim ConditionValue As String                                    ' Prüfwert
    Dim szRootkey As String                                         ' Form inhalt
    Dim szIDField As String                                         ' Name des ID Feldes
    
On Error GoTo Errorhandler
    
    Set DBConn = frmEdit.GetDBConn                                  ' DB Verbindung Holen
    szRootkey = frmEdit.GetRootkey                                  ' Rootkey ermitteln
    szIDField = frmEdit.IDField                                     ' ID Feld Ermitteln
    szAktStep = frmEdit.GetCurrentStep                               ' Akt. Schritt ermitteln
    
    szSQL = "SELECT CONDITION006 FROM Workflow006 WHERE STEP006 = " & szAktStep
    szCondition = objTools.checknull(DBConn.GetValueFromSQL(szSQL), "")
    
    szSQL = "SELECT CONDITONVALUE006 FROM Workflow006 WHERE STEP006 = " & szAktStep
    ConditionValue = objTools.checknull(DBConn.GetValueFromSQL(szSQL), "")
    
    szSQL = "SELECT CONDITIONFROM006 FROM Workflow006 WHERE STEP006 = " & szAktStep
    szTabName = objTools.checknull(DBConn.GetValueFromSQL(szSQL), "")
    
    If szCondition <> "" And szTabName <> "" And CStr(ConditionValue) <> "" Then
        szSQL = "IF (SELECT " & ConditionValue & " FROM " & szTabName & _
                " WHERE " & szIDField & " = '" & frmEdit.ID & "') " & _
                szCondition & " SELECT 'TRUE' ELSE SELECT 'FALSE' "
        ValidateWorkflowCurrentStep = CBool(objTools.checknull(DBConn.GetValueFromSQL(szSQL), False))
    Else
        ValidateWorkflowCurrentStep = True                          ' Keine Bed. Alles OK
    End If
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ValidateWorkflowCurrentStep", errNr, errDesc)
    Resume exithandler
End Function

Public Function DoWorkflowAction(frmEdit As Form, Optional bAskFor As Boolean) As Boolean

    Dim szSQL As String                                             ' SQL Statement
    Dim DBConn As Object                                            ' Aktuelle DB Verbindung
    'Dim StepArray() As String                                       ' Array mit Haupt & Teilschritt
    Dim szAction As String                                          ' Aktion als String
    Dim szActionDetails As String
    Dim PersID As String
    Dim StellenID As String
    Dim AusschrID As String
    Dim szRootkey As String
    
    On Error GoTo Errorhandler

    Set DBConn = frmEdit.GetDBConn                                  ' DB Verbindung Holen
    
    If bAskFor Then
        Dim szMSG As String
        szSQL = "SELECT DESC006 FROM WORKFLOW006 WHERE STEP006 = " & frmEdit.GetCurrentStep
        szMSG = " Möchten Sie nun diesen Vorgang mit '" & objTools.checknull( _
            DBConn.GetValueFromSQL(szSQL), "") & "' fortsezen?"
        If objError.ShowErrMsg(szMSG, vbQuestion + vbYesNo, "Vorgang fortsetzen", _
            False, "", frmEdit) = vbNo Then
            GoTo exithandler
        End If
    End If
    'StepArray = Split(frmEdit.GetCurrentStep, ".")                  ' Akt. Schtitt in Haupt & Teilschritt ausspalten
        szSQL = "SELECT TOP 1 Action006 FROM Workflow006 " & _
                " WHERE STEP006 = " & frmEdit.GetCurrentStep _
                & " ORDER BY STEP006"                               ' SQL Statement für Aktions Schritt zusammensetzen
    szAction = objTools.checknull( _
            DBConn.GetValueFromSQL(szSQL), "")                      ' SQL Statement ausführen
    If szAction <> "" Then                                          ' Wenn Ergebnis gefunden
        If InStr(UCase(szAction), "EDIT ") > 0 Then                 ' ActionDetails ist Rootkey für Edit Form ohne ID
            szActionDetails = Replace(szAction, "EDIT", "")         ' ActionDetails extrahieren
            szActionDetails = Trim(szActionDetails)
            Call GetIDCollection(frmEdit, PersID, StellenID, AusschrID) ' Relevante ID ermitteln
            Select Case UCase(szActionDetails)
            Case UCase("Stelle")
                If StellenID <> "" Then Call EditDS(szActionDetails, StellenID)
            Case UCase("Personenkartei")
                If PersID <> "" Then Call EditDS(szActionDetails, PersID)
            Case UCase("AUSSCHREIBUNG")
                If AusschrID <> "" Then Call EditDS(szActionDetails, AusschrID)
            Case Else
            
            End Select
            
        End If
        If InStr(UCase(szAction), "NEW ") > 0 Then                  ' ActionDetails ist Rootkey für Edit Form Mit ID
            szActionDetails = Replace(szAction, "NEW", "")          ' ActionDetails extrahieren
            szActionDetails = Trim(szActionDetails)
            Call GetIDCollection(frmEdit, PersID, StellenID, AusschrID) ' Relevante ID ermitteln
            Select Case UCase(szActionDetails)
            Case UCase("Bewerbungen")
                Call EditDS(szActionDetails, ";" & StellenID & ";" & PersID)
            Case Else
                Call NewDS(szActionDetails, False)                  ' Neuen DS
            End Select
            
        End If
        If InStr(UCase(szAction), "WORD ") > 0 Then                 ' ActionDetails ist Vorlagenname
            szActionDetails = Replace(szAction, "WORD", "")         ' ActionDetails extrahieren
            szActionDetails = Trim(szActionDetails)
            
            Call GetIDCollection(frmEdit, PersID, StellenID, AusschrID) ' Relevante ID ermitteln
            'szSQL = "SELECT FK010013 FROM STELLEN LEFT JOIN BEWERB013 ON ID012 = FK012013 WHERE RANG013 = 1"
            'PersID = objTools.checknull(DBConn.GetValueFromSQL(szSQL), "")
            'PersID = ShowSearch(DBConn , "PersonenNachStellen" , "Nachname", _
                , StellenID, Optional SuchTitel As String) As String
        
            Call WriteWord(szActionDetails, _
                    PersID, StellenID, AusschrID, True, frmEdit)    ' Word mit Vorlage und IDs starten
        End If
        If InStr(UCase(szAction), "TAB ") > 0 Then                  ' ActionDetails ist Tabname der ausgewählt werden soll
            szActionDetails = Replace(szAction, "TAB", "")          ' ActionDetails extrahieren
            szActionDetails = Trim(szActionDetails)
            Call frmEdit.SelectTabByName(szActionDetails)
        End If
    End If
    
    DoWorkflowAction = True
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "DoWorkflowAction", errNr, errDesc)
    Resume exithandler
End Function

Public Sub ShowWorkflowValidateMSG(frmEdit As Form)
' Zeigt Faild Massage füe Neuen Workflow Schritt an

    Dim szSQL As String                                             ' SQL Statement
    Dim DBConn As Object                                            ' Aktuelle DB Verbindung
    Dim szMSG As String                                             ' MessageText
    Dim szAktStep As String
    
    Dim StepArray() As String                                       ' Array mit Haupt & Teilschritt
    
On Error GoTo Errorhandler

    Set DBConn = frmEdit.GetDBConn                                  ' DB Verbindung Holen
    szAktStep = frmEdit.GetCurrentStep
    'StepArray = Split(frmEdit.GetCurrentStep, ".")                  ' Akt. Schtitt in HAupt & Teilschritt ausspalten
    szMSG = "Sie können diesen Schritt noch nicht abschliessen. " & vbCrLf ' Standart Meldung setzen
    
    szSQL = "SELECT CONDITIONFAILDMSG006 FROM Workflow006 " & _
            " WHERE STEP006 = " & szAktStep                        ' SQL Statement zusammen setzen
    szMSG = szMSG & objTools.checknull(DBConn.GetValueFromSQL(szSQL), "") ' Meldungstext Setzen
    
    'Call MsgBox(szMSG, , vbInformation, "Vorgang fortsetzen")       ' Meldung anzeigen
    Call objError.ShowErrMsg(szMSG, vbInformation, "Vorgang fortsetzen", _
            False, "", frmEdit)                                     ' Meldung anzeigen
            
exithandler:
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ShowWorkflowValidateMSG", errNr, errDesc)
    Resume exithandler
End Sub

Public Function GetWorkflowCurrentStep(frmEdit As Form, _
        szWorkFlowField As String) As String

    Dim szSQL As String                                             ' SQL Statement
    Dim DBConn As Object                                            ' Aktuelle DB Verbindung
    Dim AktStep As String
    
On Error GoTo Errorhandler

    AktStep = objTools.checknull(frmEdit.Adodc1.Recordset.Fields( _
            szWorkFlowField).Value, "")                             ' Aktuellen Step ermitteln
    
    If AktStep = "" Then                                            ' Wenn Akt Schritt Leer
        AktStep = GetWorkflowMinStep(frmEdit)
    End If
    
    GetWorkflowCurrentStep = AktStep                                ' Aktuelle Schritt zurück

exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetWorkflowCurrentStep", errNr, errDesc)
    Resume exithandler
End Function
                                                                    
Public Function GetWorkflowMinStep(frmEdit As Form) As String

    Dim szSQL As String                                             ' SQL Statement
    Dim DBConn As Object                                            ' Aktuelle DB Verbindung
    Dim szRootkey As String                                         ' Welches Form
    
On Error GoTo Errorhandler

    szRootkey = frmEdit.GetRootkey
    
    szSQL = "SELECT CAst(Min(STEP006) as varchar(5)) " & _
            " FROM Workflow006 WHERE Rootkey006 = '" & szRootkey & "'"    ' Min Schritt ermitteln
    Set DBConn = frmEdit.GetDBConn                                  ' DB Verbindung Holen
    GetWorkflowMinStep = objTools.checknull(DBConn.GetValueFromSQL(szSQL), "")
        
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetWorkflowMinStep", errNr, errDesc)
    Resume exithandler
End Function

Public Function GetWorkflowNextStep(frmEdit As Form)

    Dim szSQL As String                                             ' SQL Statement
    Dim DBConn As Object                                            ' Aktuelle DB Verbindung
    'Dim StepArray() As String                                       ' Array mit Haupt & Teilschritt
    Dim szAktStep As String
    Dim szRootkey As String
    
On Error GoTo Errorhandler

    Set DBConn = frmEdit.GetDBConn                                  ' DB Verbindung Holen
    szAktStep = frmEdit.GetCurrentStep                    ' Akt. Schtitt in HAupt & Teilschritt ausspalten
    szRootkey = frmEdit.GetRootkey
    
    szSQL = "SELECT TOP 1 STEP006  " & _
            " FROM Workflow006 WHERE ROOTKEY006 = '" & szRootkey & "' AND STEP006 > " & szAktStep _
            & " ORDER BY STEP006"                                   ' Nächsten Schritt ermitteln
    GetWorkflowNextStep = objTools.checknull( _
            DBConn.GetValueFromSQL(szSQL), "")                      ' Nächsten Schritt zurück
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetWorkflowNextStep", errNr, errDesc)
    Resume exithandler
End Function

Public Function GetWorkflowPreStep(frmEdit As Form, Optional lngLevel As Integer)

    Dim szSQL As String                                             ' SQL Statement
    Dim DBConn As Object                                            ' Aktuelle DB Verbindung
    'Dim StepArray() As String                                       ' Array mit Haupt & Teilschritt
    Dim szAktStep As String
    Dim szRootkey As String
    
On Error GoTo Errorhandler

    Set DBConn = frmEdit.GetDBConn                                  ' DB Verbindung Holen
    szAktStep = frmEdit.GetCurrentStep                    ' Akt. Schtitt in HAupt & Teilschritt ausspalten
    szRootkey = frmEdit.GetRootkey

    szSQL = "SELECT STEP006 " & _
            " FROM Workflow006 WHERE ROOTKEY006 = '" & szRootkey & "' AND " & _
            " STEP006 < " & szAktStep                        ' Vorherigen Schritt ermitteln
            
    GetWorkflowPreStep = objTools.checknull( _
            DBConn.GetValueFromSQL(szSQL), "")                      ' Vorherigen Schritt zurück
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetWorkflowPreStep", errNr, errDesc)
    Resume exithandler
End Function

Public Sub ShowWorkflowSteps(frmEdit As Form, LV As ListView)

    Dim szSQL As String                                             ' SQL Statement
    'Dim StepArray() As String                                       ' Array mit Haupt & Teilschritt
    Dim szAktStep As String
    Dim szRootkey As String
On Error GoTo Errorhandler

    szAktStep = frmEdit.GetCurrentStep
    szRootkey = frmEdit.GetRootkey
    'StepArray = Split(frmEdit.GetCurrentStep, ".")                  ' CurrentStep aus Form holen
    
    
    szSQL = "SELECT STEP006, STEPTITLE006 AS Vorgang FROM WORKFLOW006 WHERE ROOTKEY006 = '" & szRootkey & _
            "' ORDER BY STEP006"

    Call InitWorkflowListView(frmEdit, LV, szSQL, "STEP006", CLng(szAktStep))
    Call ShowWorkflowCurentStep(frmEdit, LV, szAktStep)          ' Akt. Schritt hervorheben
    
exithandler:
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ShowWorkflowSteps", errNr, errDesc)
    Resume exithandler
End Sub
                                                                    
Public Sub ShowWorkflowCurentStep(frmEdit As Form, LV As ListView, StepValue As String)
' Setzt LVItem Icon für Aktuellen Schritt

    Dim i As Integer                                                ' Counter
    'Dim StepArray() As String                                       ' Array mit Haupt & Teilschritt
    
On Error Resume Next
    
    For i = 1 To LV.ListItems.Count
        If Right(LV.ListItems(i).Key, 2) = StepValue Then
            LV.ListItems(i).SmallIcon = 27
            Call SelectLVItem(LV, LV.ListItems(i).Key)
            Exit For
        End If
    Next

    Err.Clear
End Sub
                                                                    
                                                                    ' *****************************************
                                                                    ' Workflow Gehampel (ALT)
Public Sub CheckWorkflowButtons(frmEdit As Form, _
        CurrentStep As String, _
        NextStep As String, _
        PrevStep As String)

On Error Resume Next

    If CurrentStep = "" Then
        frmEdit.cmdNextStep.Enabled = False
    Else
        frmEdit.cmdNextStep.Enabled = True
    End If
'    If NextStep = "" Then
'        frmEdit.cmdNextStep.Enabled = False
'    Else
'        frmEdit.cmdNextStep.Enabled = True
'    End If
    
    If PrevStep = "" Then
        frmEdit.cmdPrevStep.Enabled = False
    Else
        frmEdit.cmdPrevStep.Enabled = True
    End If
    Err.Clear
End Sub

                                                                    
'Public Function GetWorkflowCurrentStep(frmEdit As Form, _
'        szWorkFlowField As String, _
'        Optional lngLevel As Integer) As String
'
'    Dim szSQL As String                                             ' SQL Statement
'    Dim DBConn As Object                                            ' Aktuelle DB Verbindung
'    Dim AktStep As String
'
'On Error GoTo Errorhandler
'
'    AktStep = objTools.checknull(frmEdit.Adodc1.Recordset.Fields( _
'            szWorkFlowField).Value, "")                             ' Aktuellen Step ermitteln
'    If lngLevel = 0 Then lngLevel = 1
'    If AktStep = "" Then                                            ' Wenn Akt Schritt Leer
'        AktStep = GetWorkflowMinStep(frmEdit, lngLevel)
''        If szTabname = "" Then
''            szSQL = "SELECT CAst(Min(Order006) as varchar(5)) + '.' + Cast(Min(Step006) as varchar(5)) " & _
''                    " FROM Workflow006"                             ' Min Schritt ermitteln
''        Else
''            szSQL = "SELECT CAst(Min(Order006) as varchar(5)) + '.' + Cast(Min(Step006) as varchar(5)) " & _
''                    " FROM Workflow006 WHERE TABNAME006 = '" _
''                    & szTabname & "'"                               ' Min Schritt ermitteln
''        End If
''        Set dbConn = frmEdit.GetDBConn                              ' DB Verbindung Holen
''        AktStep = objTools.checknull(dbConn.GetValueFromSQL(szSQL), "")
'    End If
'
'    GetWorkflowCurrentStep = AktStep                                ' Aktuelle Schritt zurück
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
'    Call objError.Errorhandler(MODULNAME, "GetWorkflowCurrentStep", errNr, errDesc)
'    Resume exithandler
'End Function

'Public Function GetWorkflowNextStep(frmEdit As Form, Optional lngLevel As Integer)
'
'    Dim szSQL As String                                             ' SQL Statement
'    Dim DBConn As Object                                            ' Aktuelle DB Verbindung
'    Dim StepArray() As String                                       ' Array mit Haupt & Teilschritt
'
'On Error GoTo Errorhandler
'
'    Set DBConn = frmEdit.GetDBConn                                  ' DB Verbindung Holen
'    StepArray = Split(frmEdit.GetCurrentStep, ".")                  ' Akt. Schtitt in HAupt & Teilschritt ausspalten
'    If lngLevel = 0 Then lngLevel = 1
'
'    szSQL = "SELECT TOP 1 CAST(ORDER006 as varchar(5)) + '.' + CAST(Step006 as varchar(5)) " & _
'            " FROM Workflow006 WHERE LEVEL006 = " & lngLevel & " AND Step006 > " & StepArray(1) _
'            & " ORDER BY STEP006"                                   ' Nächsten Schritt ermitteln
'
''    If szTabName = "" Then
''        szSQL = "SELECT TOP 1 CAST(ORDER006 as varchar(5)) + '.' + CAST(Step006 as varchar(5)) " & _
''                " FROM Workflow006 WHERE Step006 > " & StepArray(1) _
''                & " ORDER BY STEP006"                               ' Nächsten Schritt ermitteln
''    Else
''        szSQL = "SELECT TOP 1 CAST(ORDER006 as varchar(5)) + '.' + CAST(Step006 as varchar(5)) " & _
''                " FROM Workflow006 WHERE TABNAME006 = '" & szTabName & "' AND Step006 > " & StepArray(1) _
''                & " ORDER BY STEP006"                               ' Nächsten Schritt ermitteln
''    End If
'    GetWorkflowNextStep = objTools.checknull( _
'            DBConn.GetValueFromSQL(szSQL), "")                      ' Nächsten Schritt zurück
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
'    Call objError.Errorhandler(MODULNAME, "GetWorkflowNextStep", errNr, errDesc)
'    Resume exithandler
'End Function

'Public Function GetWorkflowMinStep(frmEdit As Form, Optional lngLevel As Integer) As String
'
'    Dim szSQL As String                                             ' SQL Statement
'    Dim DBConn As Object                                            ' Aktuelle DB Verbindung
'
'On Error GoTo Errorhandler
'
'    If lngLevel = 0 Then lngLevel = 1
''    If szTabName = "" Then
''        szSQL = "SELECT CAst(Min(Order006) as varchar(5)) + '.' + Cast(Min(Step006) as varchar(5)) " & _
''                    " FROM Workflow006"                             ' Min Schritt ermitteln
''    Else
''        szSQL = "SELECT CAst(Min(Order006) as varchar(5)) + '.' + Cast(Min(Step006) as varchar(5)) " & _
''                    " FROM Workflow006 WHERE TABNAME006 = '" _
''                    & szTabName & "'"                               ' Min Schritt ermitteln
''    End If
'    szSQL = "SELECT CAst(Min(Order006) as varchar(5)) + '.' + Cast(Min(Step006) as varchar(5)) " & _
'            " FROM Workflow006 WHERE Level006 = " & lngLevel        ' Min Schritt ermitteln
'    Set DBConn = frmEdit.GetDBConn                                  ' DB Verbindung Holen
'    GetWorkflowMinStep = objTools.checknull(DBConn.GetValueFromSQL(szSQL), "")
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
'    Call objError.Errorhandler(MODULNAME, "GetWorkflowMinStep", errNr, errDesc)
'    Resume exithandler
'End Function

'Public Function GetWorkflowPreStep(frmEdit As Form, Optional lngLevel As Integer)
'
'    Dim szSQL As String                                             ' SQL Statement
'    Dim DBConn As Object                                            ' Aktuelle DB Verbindung
'    Dim StepArray() As String                                       ' Array mit Haupt & Teilschritt
'
'On Error GoTo Errorhandler
'
'    Set DBConn = frmEdit.GetDBConn                                  ' DB Verbindung Holen
'    StepArray = Split(frmEdit.GetCurrentStep, ".")                  ' Akt. Schtitt in Haupt & Teilschritt ausspalten
'    If lngLevel = 0 Then lngLevel = 1
'
''    If szTabName = "" Then
''        szSQL = "SELECT TOP 1 CAST(ORDER006 as varchar(5)) + '.' + CAST(Step006 as varchar(5)) " & _
''                " FROM Workflow006 WHERE Step006 < " & StepArray(1) _
''                & " ORDER BY STEP006"                                   ' Vorherigen Schritt ermitteln
''    Else
''        szSQL = "SELECT TOP 1 CAST(ORDER006 as varchar(5)) + '.' + CAST(Step006 as varchar(5)) " & _
''                " FROM Workflow006 WHERE TABNAME006 = '" & szTabName & "' AND Step006 < " & StepArray(1) _
''                & " ORDER BY STEP006"                                   ' Vorherigen Schritt ermitteln
''    End If
'    szSQL = "SELECT CAst(Order006 as varchar(5)) + '.' + Cast(Step006 as varchar(5)) " & _
'            " FROM Workflow006 WHERE Level006 = " & lngLevel & _
'            " AND Step006 < " & StepArray(1)                        ' Vorherigen Schritt ermitteln
'
'    GetWorkflowPreStep = objTools.checknull( _
'            DBConn.GetValueFromSQL(szSQL), "")                      ' Vorherigen Schritt zurück
'exithandler:
'
'Exit Function
'Errorhandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.Errorhandler(MODULNAME, "GetWorkflowPreStep", errNr, errDesc)
'    Resume exithandler
'End Function

'Public Function WorkflowNextStep(frmEdit As Form, szWorkFlowField As String, _
'        CurrentStep As String, _
'        NextStep As String) As Boolean
'
'On Error GoTo Errorhandler
'
'    If ValidateWorkflowCurrentStep(frmEdit) Then                    ' Wenn Evtl. Bedingungen erfolgreich
'                                                                    ' Dann nächster Schritt
'        'PrevStep = CurrentStep                                      ' Akt Schritt zu Prev Schritt
'        CurrentStep = NextStep                                      ' Next Schritt zu Akt Schritt
'        'NextStep = ""                                               ' Net Schritt wird in RefreshFrameWorkflow geholt
'        frmEdit.Adodc1.Recordset.Fields(szWorkFlowField).Value = CurrentStep  ' Wert ins RS
'        frmEdit.Adodc1.Recordset.Update                             ' Speichern
'
'        Call frmEdit.RefreshFrameWorkflow(frmEdit.FrameWorkflow.Visible)     ' Ansicht Aktualisieren
'        WorkflowNextStep = True
'    Else                                                            ' Sonst Meldung warum nicht
'        Call ShowWorkflowValidateMSG(frmEdit)                       ' Meldung
'        WorkflowNextStep = False
'    End If
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
'    Call objError.Errorhandler(MODULNAME, "WorkflowNextStep", errNr, errDesc)
'    Resume exithandler
'End Function

'Public Function DoWorkflowAction(frmEdit As Form, Optional bAskFor As Boolean) As Boolean
'
'    Dim szSQL As String                                             ' SQL Statement
'    Dim DBConn As Object                                            ' Aktuelle DB Verbindung
'    Dim StepArray() As String                                       ' Array mit Haupt & Teilschritt
'    Dim szAction As String                                          ' Aktion als String
'    Dim szActionDetails As String
'    Dim PersID As String
'    Dim StellenID As String
'    Dim AusschrID As String
'
'    On Error GoTo Errorhandler
'
'    Set DBConn = frmEdit.GetDBConn                                  ' DB Verbindung Holen
'    StepArray = Split(frmEdit.GetCurrentStep, ".")                  ' Akt. Schtitt in Haupt & Teilschritt ausspalten
'        szSQL = "SELECT TOP 1 Action006 " & _
'                " FROM Workflow006 WHERE Step006 = " & StepArray(1) _
'                & " ORDER BY STEP006"                               ' SQL Statement für Aktions Schritt zusammensetzen
'    szAction = objTools.checknull( _
'            DBConn.GetValueFromSQL(szSQL), "")                      ' SQL Statement ausführen
'    If szAction <> "" Then                                          ' Wenn Ergebnis gefunden
'        If InStr(UCase(szAktion), "EDIT") > 0 Then                  ' ActionDetails ist Rootkey für Edit Form ohne ID
'            szActionDetails = Replace(szAction, "EDIT", "")         ' ActionDetails extrahieren
'            szActionDetails = Trim(szActionDetails)
'            Call GetIDCollection(frmEdit, PersID, StellenID, AusschrID) ' Relevante ID ermitteln
'            Select Case UCase(szActionDetails)
'            Case UCase("Stelle")
'                If StellenID <> "" Then Call EditDS(szActionDetails, StellenID)
'            Case UCase("Bewerber")
'                If PersID <> "" Then Call EditDS(szActionDetails, PersID)
'            Case UCase("AUSSCHREIBUNG")
'                If AusschrID <> "" Then Call EditDS(szActionDetails, AusschrID)
'            Case Else
'
'            End Select
'
'        End If
'        If InStr(UCase(szAktion), "NEW") > 0 Then                   ' ActionDetails ist Rootkey für Edit Form Mit ID
'            szActionDetails = Replace(szAction, "NEW", "")          ' ActionDetails extrahieren
'            szActionDetails = Trim(szActionDetails)
'            Call NewDS(szActionDetails, False)                      ' Neuen DS
'        End If
'        If InStr(UCase(szAktion), "WORD") > 0 Then                  ' ActionDetails ist Vorlagenname
'            szActionDetails = Replace(szAction, "WORD", "")         ' ActionDetails extrahieren
'            szActionDetails = Trim(szActionDetails)
'            Call GetIDCollection(frmEdit, PersID, StellenID, AusschrID) ' Relevante ID ermitteln
'            Call WriteWord(objOptions.GetOptionByName(OPTION_TEMPLATES) & "\" & szActionDetails, _
'                    PersID, StellenID, AusschrID, True, Me)         ' Word mit Vorlage und IDs starten
'        End If
'    End If
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
'    Call objError.Errorhandler(MODULNAME, "DoWorkflowAction", errNr, errDesc)
'    Resume exithandler
'End Function

Public Sub SetWorkflowDescription(frmEdit As Form, Optional Step As String)
' Läd die Beschreibung des Aktuellen Teilschritts

    Dim DBConn As Object                                            ' Aktuelle DB Verbindung
    Dim StepArray() As String                                       ' Array mit Haupt & Teilschritt
    Dim szSQL As String                                             ' SQL Statement
    Dim szAktStep As String
    
On Error GoTo Errorhandler

    Set DBConn = frmEdit.GetDBConn                                  ' DB Verbindung Holen
    szAktStep = frmEdit.GetCurrentStep
    
    'StepArray = Split(frmEdit.GetCurrentStep, ".")                  ' Akt. Schtitt in HAupt & Teilschritt ausspalten
    
    If Step = "" Then Step = szAktStep                           ' Wenn Kein Step angegeben Akt Step holen
    'If SubStep = "" Then SubStep = StepArray(1)                     ' Wenn Kein SubStep angegeben akt SubStep holen
    
    szSQL = "SELECT DESC006 FROM Workflow006 WHERE STEP006 = " & Step                    ' SQL Statement zusammen setzen
    frmEdit.lblStepDesc.Caption = objTools.checknull( _
            DBConn.GetValueFromSQL(szSQL), "")                      ' Ausführen und als LableCaption setzen
    
exithandler:
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SetWorkflowDescription", errNr, errDesc)
    Resume exithandler
End Sub

'Public Sub ShowWorkflowValidateMSG(frmEdit As Form)
'' Zeigt Faild Massage füe Neuen Workflow Schritt an
'
'    Dim szSQL As String                                             ' SQL Statement
'    Dim DBConn As Object                                            ' Aktuelle DB Verbindung
'    Dim szMSG As String                                             ' MessageText
'    Dim StepArray() As String                                       ' Array mit Haupt & Teilschritt
'
'On Error GoTo Errorhandler
'
'    Set DBConn = frmEdit.GetDBConn                                  ' DB Verbindung Holen
'    StepArray = Split(frmEdit.GetCurrentStep, ".")                  ' Akt. Schtitt in HAupt & Teilschritt ausspalten
'    szMSG = "Sie können diesen Schritt noch nicht abschliessen. " & vbCrLf ' Standart Meldung setzen
'
'    szSQL = "SELECT FALIEDMSG006 FROM Workflow006 WHERE ORDER006 = " _
'            & StepArray(0) & " AND STep006 = " & StepArray(1)       ' SQL Statement zusammen setzen
'    szMSG = szMSG & objTools.checknull(DBConn.GetValueFromSQL(szSQL), "") ' Meldungstext Setzen
'
'    'Call MsgBox(szMSG, , vbInformation, "Vorgang fortsetzen")       ' Meldung anzeigen
'    Call objError.ShowErrMsg(szMSG, vbInformation, "Vorgang fortsetzen", _
'            False, "", frmEdit)                                     ' Meldung anzeigen
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
'    Call objError.Errorhandler(MODULNAME, "ShowWorkflowValidateMSG", errNr, errDesc)
'    Resume exithandler
'End Sub

'Public Function ValidateWorkflowCurrentStep(frmEdit As Form) As Boolean
'
'    Dim DBConn As Object                                            ' Aktuelle DB Verbindung
'    Dim StepArray() As String                                       ' Array mit Haupt & Teilschritt
'    Dim NextSteparray() As String
'    Dim szSQL As String                                             ' SQL Statement
'    Dim szCondition As String                                       ' Prüfbedingung
'    Dim szTabName As String                                         ' zu prüfende Tabelle
'    Dim ConditionValue As String                                    ' Prüfwert
'    Dim lngCurStelpLevel As Integer
'    Dim lngNextStepLevel As Integer
'
'On Error GoTo Errorhandler
'
'    Set DBConn = frmEdit.GetDBConn                                  ' DB Verbindung Holen
'    StepArray = Split(frmEdit.GetCurrentStep, ".")                  ' Akt. Schrtitt in Haupt & Teilschritt ausspalten
'    NextSteparray = Split(frmEdit.GetNextStep, ".")                 ' nächsten Schrtitt in Haupt & Teilschritt ausspalten
'
'    szSQL = "SELECT LEVEL006 FROM WORKFLOW006 WHERE STEP006 =" & StepArray(1)
'    lngCurStelpLevel = objTools.checknull(DBConn.GetValueFromSQL(szSQL), "") ' Akt StepLevel ermitteln
'
'    szSQL = "SELECT LEVEL006 FROM WORKFLOW006 WHERE STEP006 =" & NextSteparray(1)
'    lngNextStepLevel = objTools.checknull(DBConn.GetValueFromSQL(szSQL), "") ' NextStepLevel ermitteln
'    If lngNextStepLevel > lngCurStelpLevel Then
'
'    End If
'    szSQL = "SELECT CONDITION006 FROM Workflow006 WHERE ORDER006 = " & StepArray(0) & " AND STep006 = " & StepArray(1)
'    szCondition = objTools.checknull(DBConn.GetValueFromSQL(szSQL), "")
'    szSQL = "SELECT CONDITIONValue006 FROM Workflow006 WHERE ORDER006 = " & StepArray(0) & " AND STep006 = " & StepArray(1)
'    ConditionValue = objTools.checknull(DBConn.GetValueFromSQL(szSQL), "")
'    szSQL = "SELECT TABNAME006  FROM Workflow006 WHERE ORDER006 = " & StepArray(0) & " AND STep006 = " & StepArray(1)
'    szTabName = objTools.checknull(DBConn.GetValueFromSQL(szSQL), "")
'
'    If szCondition <> "" And szTabName <> "" Then
'        szSQL = "IF (SELECT " & ConditionValue & " FROM " & szTabName & _
'                " WHERE ID" & Right(szTabName, 3) & " = '" & frmEdit.ID & "') " & _
'                szCondition & " SELECT 'TRUE' ELSE SELECT 'FALSE' "
'        ValidateWorkflowCurrentStep = CBool(objTools.checknull(DBConn.GetValueFromSQL(szSQL), False))
'    End If
'exithandler:
'
'Exit Function
'Errorhandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.Errorhandler(MODULNAME, "ValidateWorkflowCurrentStep", errNr, errDesc)
'    Resume exithandler
'End Function

'Public Sub ShowWorkflowSteps(frmEdit As Form, LV As ListView, Optional lngLevel As Integer)
'
'    Dim szSQL As String                                             ' SQL Statement
'    Dim StepArray() As String                                       ' Array mit Haupt & Teilschritt
'
'On Error GoTo Errorhandler
'
'    StepArray = Split(frmEdit.GetCurrentStep, ".")                  ' CurrentStep aus Form holen
'    If lngLevel = 0 Then lngLevel = 1
'    szSQL = "SELECT  ORDER006 ,Caption006 AS Vorgang FROM WORKFLOW006 WHERE LEVEL006 = " _
'                & lngLevel & " Group by ORDER006,  Caption006 ORDER BY ORDER006"
'
'    Call InitWorkflowListView(frmEdit, LV, szSQL, "ORDER006", CLng(StepArray(0)))
'    Call ShowWorkflowCurentStep(frmEdit, LV, StepArray(0))          ' Akt. Schritt hervorheben
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
'    Call objError.Errorhandler(MODULNAME, "ShowWorkflowSteps", errNr, errDesc)
'    Resume exithandler
'End Sub

Public Sub ShowWorkflowSubSteps(frmEdit As Form, LV As ListView, Optional Step As String, Optional lngLevel As Integer)

    Dim szSQL As String                                             ' SQL Statement
    Dim StepArray() As String                                       ' Array mit Haupt & Teilschritt
    Dim i As Integer
    
On Error GoTo Errorhandler

    StepArray = Split(frmEdit.GetCurrentStep, ".")                  ' CurrentStep aus Form holen

    If Step = "" Then Step = StepArray(0)
    If lngLevel = 0 Then lngLevel = 1
    szSQL = "SELECT STEP006, STEPTITLE006 AS Teilschritt, SubStep006, REPETITION006, " & _
            " ACTION006, CONDITION006, DESCR006 FROM WORKFLOW006 " _
            & " WHERE LEVEL006 = " & lngLevel & " AND ORDER006 = " _
            & Step & "  ORDER BY SubStep006"
    
    Call InitWorkflowListView(frmEdit, LV, szSQL, "STEP006", CLng(StepArray(1)))
    Call ShowWorkflowCurentStep(frmEdit, LV, StepArray(1))          ' Akt. Schritt hervorheben
    
exithandler:
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ShowWorkflowSteps", errNr, errDesc)
    Resume exithandler
End Sub

'Public Sub ShowWorkflowCurentStep(frmEdit As Form, LV As ListView, StepValue As String)
'' Setzt LVItem Icon für Aktuellen Schritt
'
'    Dim i As Integer                                                ' Counter
'    Dim StepArray() As String                                       ' Array mit Haupt & Teilschritt
'
'On Error Resume Next
'
'    For i = 1 To LV.ListItems.Count
'        If Right(LV.ListItems(i).Key, 2) = StepValue Then
'            LV.ListItems(i).SmallIcon = 27
'            Call SelectLVItem(LV, LV.ListItems(i).Key)
'            Exit For
'        End If
'    Next
'
'    Err.Clear
'End Sub

Public Function InitWorkflowListView(frmEdit As Form, LV As ListView, szSQL As String, _
        AltImgField As String, AltImgValue As Variant) As ADODB.Recordset

    Dim DBConn As Object                                            ' Aktuelle DB Verbindung
    Dim StepArray() As String                                       ' Array mit Haupt & Teilschritt
    Dim RS As ADODB.Recordset                                       ' rs mit schritt daten
    Dim i As Integer                                                ' Counter
    Dim szAktStep As Variant
On Error GoTo Errorhandler

    Set DBConn = frmEdit.GetDBConn                                  ' DB Verbindung Holen
    szAktStep = frmEdit.GetCurrentStep
    'StepArray = Split(frmEdit.GetCurrentStep, ".")                  ' Akt. Schtitt in HAupt & Teilschritt ausspalten

    Set RS = DBConn.fillrs(szSQL, True)                             ' RS füllen
    Call FillLVByRS(LV, "", RS, False, 28, "", False, True, "", _
              26, AltImgField, AltImgValue, "<")                      ' Daten is LV einlesen
    Call SetColumnWidth(LV, 1, LV.Width - 80)                       ' 1. Spalte auf LV breite einstelln
    For i = 2 To LV.ColumnHeaders.Count                             ' Alle weiteren Spalten ausblenden
        Call SetColumnWidth(LV, i, 0)
    Next i
    
    Set InitWorkflowListView = RS                                   ' RS zurück
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitWorkflowListView", errNr, errDesc)
    Resume exithandler
End Function

Public Function GetIDCollection(frmEdit As Form, _
        ByRef PersID As String, _
        ByRef StellenID As String, _
        ByRef AusschrID As String)
' Ermittelt die Möglichste kombination aus Ausschreibung,  Stelle und Person (Bewerber)
                
    Dim szCurRootKey As String                                      ' Gibt an welches Editform offen ist
    Dim szSQL As String                                             ' SQL Statement
    Dim DBConn As Object                                            ' Aktuelle DB Verbindung
    
On Error GoTo Errorhandler

    Set DBConn = frmEdit.GetDBConn
    szCurRootKey = frmEdit.GetRootkey
    If szCurRootKey = "" Then GoTo exithandler
    Select Case UCase(szCurRootKey)
    Case UCase("Ausschreibung")
        AusschrID = frmEdit.ID                                      ' Akt ID ermitteln
    Case UCase("Ausgeschriebene Stellen")
        StellenID = frmEdit.ID                                      ' Akt ID ermitteln
        If StellenID <> "" Then                                     ' Wenn StellenID vorhanden
            szSQL = "SELECT FK020012 FROM STELLEN012 WHERE ID012 = '" & StellenID & "'"
            AusschrID = objTools.checknull(DBConn.GetValueFromSQL(szSQL), "")   ' Ausschreibungs ID ermitteln
        End If
    Case UCase("Personenkartei")
        PersID = frmEdit.ID                                         ' Akt ID ermitteln
        If PersID <> "" Then                                        ' Wenn PersID vorhanden
            szSQL = "SELECT FK012013 FROM BEWERB013 INNER JOIN RA010 ON ID010 = FK010013 " & _
                    " WHERE ID010 = '" & PersID & "'"
            StellenID = objTools.checknull(DBConn.GetValueFromSQL(szSQL), "")   ' Stellen ID ermitteln
            If StellenID <> "" Then                                 ' Wenn StellenID vorhanden
                szSQL = "SELECT FK020012 FROM STELLEN012 WHERE ID012 = '" & StellenID & "'"
                AusschrID = objTools.checknull(DBConn.GetValueFromSQL(szSQL), "")   ' Ausschreibungs ID ermitteln
            End If
        End If
    Case Else
    
    End Select
    
    
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitWorkflowListView", errNr, errDesc)
    Resume exithandler
End Function

