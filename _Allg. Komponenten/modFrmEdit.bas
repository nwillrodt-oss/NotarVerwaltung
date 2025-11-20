Attribute VB_Name = "modFrmEdit"
Option Explicit

Private Const MODULNAME = "modFrmEdit"

Public Sub CheckUpdate(frmEdit As Form)
    frmEdit.cmdUpdate.Enabled = frmEdit.bDirty
End Sub

Public Sub HandleEditLVDoubleClick(frmEdit As Form, lv As ListView)
    Dim RootKey As String
    Dim DetailKey As String
    RootKey = lv.Tag
    DetailKey = lv.SelectedItem.Tag
    If RootKey <> "" And DetailKey <> "" Then
        Call frmMDIMain.OpenEditForm(RootKey, DetailKey)
    End If
End Sub

Public Sub EditFormLoad(frmEdit As Form, szRootkey As String)

    Dim ctl As Control
    Dim szDetails As String
    
On Error GoTo Errorhandler

    szDetails = "Formname: " & frmEdit.Name & vbCrLf & "ID: " & frmEdit.ID

    Call SetEditFormIcon(frmEdit, szRootkey) ' Form Icon setzen
    frmEdit.cmdUpdate.Enabled = False       ' Button Übernehmen erstmmal disablen
    frmEdit.Adodc1.Visible = False          ' Daten Verbindungs Control ausblenden
    
     For Each ctl In frmEdit.Controls
        If Left(ctl.Name, 2) = "LV" Then
            'ctl.Icons = frmDB.ILTree
            'ctl.SmallIcons = frmDB.ILTree
        End If
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
    Call frmMDIMain.CloseEditForm(frmEdit)
End Sub

Public Sub InitFrameInfo(frmEdit As Form)

    Dim szDetails As String
    
On Error GoTo Errorhandler

    szDetails = "Formname: " & frmEdit.Name & vbCrLf & "ID: " & frmEdit.ID
    
    Call frmEdit.FrameInfo.Move(frmEdit.lngFrameLeftPos, frmEdit.lngFrametopPos, _
            frmEdit.lngFrameWidth, frmEdit.lngFrameHeight)
    
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

Public Function HandleKeyDown(frmEdit As Form, ctl As Control, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 And Shift = 0 Then   ' ESC
        ' Form Schliessen ohne speichern
        Unload frmEdit
    End If
    
End Function

Public Function InitAdoDC(frmEdit As Form, DBCon As Object, szSQL As String, szWhere As String)

On Error GoTo Errorhandler
    
    If szSQL = "" Then GoTo exithandler
    If Not frmEdit.bNew And szWhere <> "" Then szSQL = objSQLTools.AddWhereInFullSQL(szSQL, szWhere)
    
On Error Resume Next
    frmEdit.Adodc1.ConnectionString = DBCon.GetConnectString
    Err.Clear
On Error GoTo Errorhandler
    frmEdit.Adodc1.CommandType = adCmdText
    frmEdit.Adodc1.RecordSource = szSQL
    frmEdit.Adodc1.Refresh

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

Public Function FillRelLV(frmEdit As Form, _
        DBCon As Object, _
        IniPath As String, _
        lv As ListView, _
        szRootkey As String, _
        szRelKey As String) As ADODB.Recordset

    Dim szSQLMain As String
    Dim szWhere As String
    Dim szImgIndex As String
    
On Error GoTo Errorhandler

    If IniPath = "" Then GoTo exithandler
    If szRelKey = "" Then GoTo exithandler
    
    szSQLMain = objTools.GetINIValue(IniPath, INI_RELATIONS, szRootkey & szRelKey)
    If szSQLMain = "" Then GoTo exithandler
    
    szWhere = objTools.GetINIValue(IniPath, INI_RELATIONS, "WHERE" & szRootkey & szRelKey)
    szImgIndex = objTools.GetINIValue(IniPath, INI_IMAGE, szRootkey & szRelKey)
    If szWhere <> "" Then
        szWhere = szWhere & "'" & frmEdit.ID & "'"  '"CAST('" & frmEdit.ID & "' as uniqueidentifier)"
        szSQLMain = objSQLTools.AddWhereInFullSQL(szSQLMain, szWhere)
    End If
    If szImgIndex = "" Then szImgIndex = "0"
    Set FillRelLV = FillLVBySQL(lv, szSQLMain, DBCon, , CLng(szImgIndex))
        
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "FillRelLV", errNr, errDesc)
    Resume exithandler
End Function

Public Function ValidateEditForm(frmEdit As Form) As Boolean

    Dim szDetails As String
    
On Error GoTo Errorhandler

    szDetails = "Formname: " & frmEdit.Name & vbCrLf & "ID: " & frmEdit.ID
    
    If frmEdit.bNew Then        ' Neuer DS -> Insert
    
    Else                ' Update
    
    End If
    
    ValidateEditForm = True
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ValidateEditForm", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Public Function UpdateEditForm(frmEdit As Form)

    Dim szDetails As String

On Error GoTo Errorhandler

    szDetails = "Formname: " & frmEdit.Name & vbCrLf & "ID: " & frmEdit.ID

    If Not frmEdit.bDirty Then GoTo exithandler                  ' Keine Änderungen -> Raus

    If frmEdit.bNew Then ' Bei neuen Datensatz
        'frmEdit.txtCreateFrom.Text = objObjectBag.GetUserName    ' Benutzer eintragen
        ' Erstellt datum wird über standartwert in der tabelle geregelt
    End If

    'frmEdit.txtModify.Text = Now()                               ' Änderungsdatum eintragen
    'frmEdit.txtModifyFrom.Text = objObjectBag.GetUserName        ' Benutzer eintragen


    If Not ValidateEditForm(frmEdit) Then GoTo exithandler       ' Prüfe ob Änderungen Zulässig

    If frmEdit.bNew Then        ' Neuer DS -> Insert
        'Adodc1.Recordset.AddNew
        frmEdit.Adodc1.Recordset.Update
    Else                ' Update
        frmEdit.Adodc1.Recordset.Update
    End If

    frmEdit.bDirty = False                                        ' Damit ist das Form nicht mehr Dirty
    Call CheckUpdate(frmEdit)                                     ' Übernehmen button disablen

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

Public Function DelRelationinLV(frmEdit As Form, _
        RootKey As String, _
        DBConn As Object, _
        lv As ListView, _
        rs As ADODB.Recordset, _
        RelIDField As String, _
        RelationTable As String)
    
    Dim RelID As String         ' Relation ID
    Dim szDetails As String     ' Zusatz für Fehlermeldung
    Dim i As Integer            ' Counter
    Dim szSQL As String         ' SQL Statement
    
On Error GoTo Errorhandler

    szDetails = "Formname: " & frmEdit.Name & vbCrLf & "ID: " & frmEdit.ID
    If lv.SelectedItem = "" Then GoTo exithandler  ' Nur wenn ein DS ausgewählt
    rs.MoveFirst
    rs.Find (lv.ColumnHeaders(1).Text & " = '" & lv.SelectedItem & "'")
    
    If Not rs.EOF Or Not rs.BOF Then
        RelID = rs.Fields(RelIDField).Value
    
        If RelID = "" Then GoTo exithandler
    
        szSQL = "DELETE  FROM " & RelationTable & " WHERE " & RelIDField & "='" & RelID & "'"
        Call DBConn.execsql(szSQL)

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

Public Function SetRelationinLV(frmEdit As Form, _
        RootKey As String, _
        SearchField As String, _
        DBConn As Object, _
        lv As ListView, _
        rs As ADODB.Recordset, _
        EntityIDField As String, _
        RelationIDField As String)
    
    Dim RelID As String         ' Relation ID
    'Dim szSQLInsert As String
    Dim szDetails As String     ' Zusatz für Fehlermeldung
    Dim i As Integer            ' Counter
    
On Error GoTo Errorhandler

    szDetails = "Formname: " & frmEdit.Name & vbCrLf & "ID: " & frmEdit.ID
    
    RelID = ShowSearch(DBConn, RootKey, SearchField)
    If RelID = "" Then GoTo exithandler
        If lv.ListItems.Count > 0 Then
            For i = 1 To lv.ListItems.Count
                If lv.ListItems(i).Text = frmEdit.ID Then
                    'schon drin
                    lv.ListItems(i).Selected = True
                Else
                    ' insert
                    rs.AddNew
                    rs.Fields(RelationIDField).Value = RelID
                    rs.Fields(EntityIDField).Value = frmEdit.ID
                    rs.Update
                End If
            Next i
        Else
            rs.AddNew
            rs.Fields(RelationIDField).Value = RelID
            rs.Fields(EntityIDField).Value = frmEdit.ID
            rs.Update
        End If
        
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SetRelationinLV", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Public Function SetEditFormCaption(frmEdit As Form, szRootkey As String, _
                    Optional szAddCaption As String)

    Dim szDetails As String
    
On Error GoTo Errorhandler

    szDetails = "Formname: " & frmEdit.Name & vbCrLf & "ID: " & frmEdit.ID
    
    If frmEdit.bNew Then
        frmEdit.Caption = szRootkey & ": Neuer Datensatz"                    ' Caption Setzen
    Else
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

Public Function SetEditFormIcon(frmEdit As Form, szRootkey As String)
   
On Error Resume Next
    frmEdit.Icon = LoadPicture("images\" & szRootkey & ".ico")           ' Form Icon setzen
    Err.Clear
    
End Function

Public Function PosFrameAndListView(frmEdit As Form, CurFrame As Frame, _
            bWithBorder As Boolean, Optional lv As ListView)
        
    Dim szDetails As String
    
On Error GoTo Errorhandler

    szDetails = "Formname: " & frmEdit.Name & vbCrLf & " ID: " & frmEdit.ID & vbCrLf _
            & " Frame: " & CurFrame.Name
            
    If bWithBorder Then
        CurFrame.BorderStyle = vbFixedSingle
    Else
        CurFrame.BorderStyle = vbBSNone
    End If
    Call CurFrame.Move(frmEdit.lngFrameLeftPos, frmEdit.lngFrametopPos, _
                    frmEdit.lngFrameWidth, frmEdit.lngFrameHeight)
    If Not lv Is Nothing Then
        Call FillFrameWithLV(CurFrame, lv)
        Call lv.Move(120, 240, frmEdit.lngFrameWidth - 240, frmEdit.lngFrameHeight - 360)
        Call InitDefaultListViewResult(lv)
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
 
 
'Private Function GetEditSQL(frmEdit As Form, _
'            DBCon As Object, _
'            szIniSection As String, _
'            szIniKey As String, _
'            bCancel As Boolean, _
'            Optional szOptFrameTitle As String, _
'            Optional bRel As Boolean) As Recordset
'
'Dim szSQL As String             ' SQL Statement
'Dim szWhere As String           ' Where statement
'Dim rsList As ADODB.Recordset   ' Ergebniss Recordet
'
'On Error GoTo ErrorHandler
'
'    szOptFrameTitle = Trim(szOptFrameTitle)
'    szSQL = objTools.GetINIValue(szIniFilePath, szIniSection, szIniKey & szOptFrameTitle)
'    If szSQL = "" Then GoTo cancelhandler               ' Kein SQL -> fertig
'    szWhere = objTools.GetINIValue(szIniFilePath, INI_SQL, "WHERE" & szIniKey)
'
''    szUpdateTable = objSQLTools.GetTableFromSQL(szSQL)  ' Update Table aus dem fertigen SQL statement extrahieren
'
'    If frmEdit.bNew Then    ' Wenn Neuer DS leeres Recordset holen
'        szSQL = szSQL & " " & objSQLTools.AddWhere("", szWhere & "'" & NEWDS & "'")
'    Else
'        If szDetailKey <> "" Then ' Wenn szIniWhereKey (in ausnahmen gibt es kein) dann Where Bed. zusammen bauen
'            szSQL = szSQL & " " & objSQLTools.AddWhere("", szWhere & "'" & szDetailKey & "'")
'        End If
'    End If
'
'    If szSQL = "" Then GoTo exithandler                 ' Kein SQL Statement -> Fertig
'    Set rsList = DBCon.fillrs(szSQL, True)          ' Daten holen
'
'    If rsList Is Nothing Then GoTo cancelhandler        ' Keine Daten -> Fertig
'    Set GetEditSQL = rsList
'
'exithandler:
'
'Exit Function
'cancelhandler:
'    bCancel = True
'    GoTo exithandler
'
'Exit Function
'ErrorHandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.ErrorHandler(MODULNAME, "GetEditSQL", errNr, errDesc)
'    Resume exithandler
'End Function



