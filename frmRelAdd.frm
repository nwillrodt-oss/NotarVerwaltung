VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRelAdd 
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.ListView ListView2 
      Height          =   3495
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ILRelEdit"
      SmallIcons      =   "ILRelEdit"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   ">"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "<"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ILRelEdit"
      SmallIcons      =   "ILRelEdit"
      ColHdrIcons     =   "ILRelEdit"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ILRelEdit 
      Left            =   3120
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelAdd.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelAdd.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelAdd.frx":00BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelAdd.frx":011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelAdd.frx":0178
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelAdd.frx":01D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelAdd.frx":0234
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelAdd.frx":0292
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRelAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULNAME = "frmRelEdit"

Dim szRootkey As String
Dim szRootTable As String
Dim szRootIDField As String

Dim szSecRelKey As String
Dim szSecRelTable As String
Dim szSecRelIDField As String

Dim szIndexTable As String
Dim szRootFKField As String
Dim szSecRelFKField As String

Dim szDetailKey  As String

Dim szUpdateTable As String
Dim rsRelList As ADODB.Recordset    ' RS mit Exist. Relationen
 
Private ThisDBCon As Object     ' Aktuelle DB Verbindung

Public Function InitRelEditForm(DBCon As Object, _
        RootKey As String, _
        DetailKey As String, _
        SecRelKey As String, _
        IndexTable As String, _
        RootTable As String, _
        SecRelTable As String)
    Set ThisDBCon = DBCon
    szRootkey = RootKey
    szRootTable = RootTable
    szRootIDField = "ID" & Right(szRootTable, 3)
    
    szSecRelKey = SecRelKey
    szSecRelTable = SecRelTable
    If szSecRelTable <> "" Then szSecRelIDField = "ID" & Right(szSecRelTable, 3)
    
    szIndexTable = IndexTable
    szRootFKField = "FK" & Right(szRootTable, 3) & Right(szIndexTable, 3)
    szSecRelFKField = "FK" & Right(szSecRelTable, 3) & Right(szIndexTable, 3)
    
    szDetailKey = DetailKey
    
End Function

Private Sub cmdAdd_Click()
Dim i As Integer
Dim LVItem As ListItem
    For i = 1 To Me.ListView2.ListItems.count
        If i > Me.ListView2.ListItems.count Then Exit For
        If Me.ListView2.ListItems(i).Selected Then
            Set LVItem = AddListViewItem(ListView1, ListView2.ListItems(i), ListView2.ListItems(i), ListView2.ListItems(i).SubItems(1))
            'Set LVItem = Me.ListView2.ListItems(i)
            'Me.ListView2.ListItems(i).SubItems.Count
            
            'Me.ListView1.ListItems.Add , LVItem.Key, LVItem.Text, LVItem.Icon, LVItem.SmallIcon
           Me.ListView2.ListItems.Remove (i)
            Me.ListView1.Refresh
            Me.ListView2.Refresh
        End If
    Next
'    If Me.ListView2.SelectedItem <> "" Then
'        Stop
'    End If
End Sub

Private Sub cmdDel_Click()
Dim i As Integer
Dim LVItem As ListItem
    For i = 1 To Me.ListView1.ListItems.count
        If i > Me.ListView1.ListItems.count Then Exit For
        If Me.ListView1.ListItems(i).Selected Then
            Set LVItem = AddListViewItem(ListView2, ListView1.ListItems(i), ListView1.ListItems(i), ListView1.ListItems(i).SubItems(1))
            'Set LVItem = Me.ListView2.ListItems(i)
            'Me.ListView2.ListItems(i).SubItems.Count
            
            'Me.ListView1.ListItems.Add , LVItem.Key, LVItem.Text, LVItem.Icon, LVItem.SmallIcon
           Me.ListView1.ListItems.Remove (i)
            Me.ListView1.Refresh
            Me.ListView2.Refresh
        End If
    Next
End Sub

Private Sub cmdEsc_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    
    'Call AddRelations
    
    'Call DelRelations
    
    Me.Hide
End Sub

Private Sub Form_Load()

    Call FillExistRelations(szSecRelKey)
    
    Call FillPosRelations(szSecRelKey)
    
End Sub

Private Function AddRelations()

Dim mark As Variant
Dim i As Integer
Dim IDField As String
Dim count As Integer
Dim szSQL As String

    IDField = rsRelList.Fields(0).Name
    
    For i = 1 To Me.ListView1.ListItems.count
        'rsRelList.MoveFirst
        If Not rsRelList.BOF Then
            rsRelList.MoveFirst
            rsRelList.Find (IDField & "=" & Me.ListView1.ListItems(i).Text)
            If rsRelList.EOF Then
                ' Nix gefunden -> Eintragen
                Debug.Print "NEU " & Me.ListView1.ListItems(i).SubItems(1)
            
                'rsRelList.AddNew
                szSQL = "INSERT INTO " & szIndexTable & " ( " & szRootFKField & ", " & szSecRelFKField _
                        & " )  Values ( '" & szDetailKey & "', '" & Me.ListView1.ListItems(i) & "')"
                Call ThisDBCon.ExecSQL(szSQL)
            
            'rsRelList.Fields("persID").Value  =
            Else
                ' Gefunden -> Nix zu Tun
                
            End If
        Else ' Keiner drin -> einfach hinzufügen
             Debug.Print "NEU " & Me.ListView1.ListItems(i).SubItems(1)
            
                'rsRelList.AddNew
                szSQL = "INSERT INTO " & szIndexTable & " ( " & szRootFKField & ", " & szSecRelFKField _
                        & " )  Values ( '" & szDetailKey & "', '" & Me.ListView1.ListItems(i) & "')"
                Call ThisDBCon.ExecSQL(szSQL)
        End If
    Next
    
End Function


Private Function DelRelations()

Dim mark As Variant
Dim i As Integer
Dim IDField As String
Dim count As Integer
Dim szSQL As String

    IDField = rsRelList.Fields(0).Name
    
    For i = 1 To Me.ListView2.ListItems.count
        
        If Not rsRelList.BOF Then
            rsRelList.MoveFirst
            rsRelList.Find (IDField & "=" & Me.ListView2.ListItems(i).Text)
            If Not rsRelList.EOF Then
                ' Nix gefunden -> Eintragen
                Debug.Print "Gelöscht " & Me.ListView2.ListItems(i).SubItems(1)
            
                'rsRelList.AddNew
                szSQL = "DELETE FROM " & szIndexTable & " WHERE " & szRootFKField & " = '" & szDetailKey & "' AND " & szSecRelFKField _
                        & " ='" & Me.ListView2.ListItems(i) & "'"
                Call ThisDBCon.ExecSQL(szSQL)
            
            'rsRelList.Fields("persID").Value  =
            Else
                ' Gefunden -> Nix zu Tun
                
            End If
        End If
    Next
    
End Function

Private Sub FillExistRelations(SecRelKey As String)
  
    Dim szSQL As String                 ' SQL Statement
    Dim i As Integer                    ' counter
    Dim szTmpName As String
    Dim szTmpID As String
    Dim LVItem As ListItem
   
    
On Error GoTo ErrorHandler

    Call ClearListView(Me.ListView1)    ' Items und columns löschen
'    Me.ListView1.ListItems.Clear          ' Evtl. vorhandene Items löschen
'    Me.ListView1.ColumnHeaders.Clear      ' Evtl. vorhandene Spaltenköpfe löschen
    
    ' SQL Basis Statement holen
    szSQL = objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_RELATIONS, szRootkey & SecRelKey)
    ' Where anhängen
    szSQL = szSQL & " " & objSQLTools.AddWhere("", _
                objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_SQL, "WHERE" & szRootkey) _
            & "'" & szDetailKey & "'")

    If szSQL = "" Then GoTo exithandler ' Kein SQL Statement -> Fertig
    ' Update Table aus dem fertigen SQL statement extrahieren
    'szUpdateTable = objSQLTools.GetTableFromSQL(szSQL)
    
    Set rsRelList = ThisDBCon.fillrs(szSQL, True)                   ' Daten holen
    If rsRelList Is Nothing Then GoTo exithandler                   ' Keine Daten -> Fertig
    
    'For i = 0 To rsRelList.Fields.Count - 1
    For i = 0 To 1
        Call AddLVColumn(Me.ListView1, rsRelList.Fields(i).Name)    ' Colum hinzufügen
    Next i
    
    Call HideColumn(Me.ListView1, 1)                                ' Erste Spalte ausblenden
    Call SetColumnWidth(Me.ListView1, 2, Me.ListView1.Width - 50)   ' optimale breite einstelln
    'Me.ListView1.ColumnHeaders(1).Width = 0                         ' Erste Spalte ausblenden
    'Me.ListView1.ColumnHeaders(2).Width = Me.ListView1.Width - 50   ' optimale breite einstelln
    
    Do While Not rsRelList.EOF              ' Für jeden DS einen eintrag
        szTmpName = Trim(objTools.checknull(rsRelList.Fields(1).Value, ""))
        szTmpID = Trim(objTools.checknull(rsRelList.Fields(0).Value, ""))
        If szTmpName <> "" And szTmpID <> "" Then
            Set LVItem = AddListViewItem(Me.ListView1, szTmpID, szTmpID, , GetImageIndexByRootNode(SecRelKey))
            For i = 1 To 1 'rsRelList.Fields.Count - 1
                ' Für jedes Feld ein SubItem
                Call AddListViewSubItem(LVItem, Trim(objTools.checknull(rsRelList.Fields(i).Value, "")))
            Next i
            szTmpName = ""
        End If  ' szTmpName <> "" And szTmpID <> ""
        rsRelList.MoveNext
    Loop
    
exithandler:
On Error Resume Next

Exit Sub
ErrorHandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.ErrorHandler(MODULNAME, "FillExistRelations", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub FillPosRelations(SecRelKey As String)
    
    Dim szSQL As String                 ' SQL Statement
    Dim szSQLNotIn As String            ' Not In SQL Statement
    Dim szWhere As String               ' Where Statement
    Dim i As Integer                    ' counter
    Dim szTmpName As String
    Dim szTmpID As String
    Dim LVItem As ListItem
    Dim rsPosRelList As ADODB.Recordset
    Dim szSecRelTable As String
    
On Error GoTo ErrorHandler

    'Me.ListView2.ListItems.Clear          ' Evtl. vorhandene Items löschen
    'Me.ListView2.ColumnHeaders.Clear      ' Evtl. vorhandene Spaltenköpfe löschen
    Call ClearListView(Me.ListView2)            ' Items und columns löschen
    
    ' Not In Statement holen ( ist identisch mit szSQL in FillExistRelations)
    szSQLNotIn = objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_RELATIONS, szRootkey & SecRelKey)
    ' Where anhängen
    szSQLNotIn = szSQLNotIn & " " & objSQLTools.AddWhere("", _
                objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_SQL, "WHERE" & szRootkey) _
            & "'" & szDetailKey & "'")
    If szSQLNotIn = "" Then GoTo exithandler    ' Kein SQL Statement -> Fertig
    
    szSQLNotIn = objSQLTools.GetOneFieldSQL(szSQLNotIn)
    
    ' Update Table aus dem fertigen SQL statement extrahieren
    szUpdateTable = objSQLTools.GetTableFromSQL(szSQLNotIn)
    
    szSQL = objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_SQL, SecRelKey)
    If szSQL = "" Then GoTo exithandler         ' Kein SQL Statement -> Fertig
    szSecRelTable = objSQLTools.GetTableFromSQL(szSQL)
    
    szWhere = objSQLTools.GetIDField(szSecRelTable) & " NOT IN ( " & szSQLNotIn & ")"
    szSQL = szSQL & " " & objSQLTools.AddWhere("", szWhere)
    
    
    
    Set rsPosRelList = ThisDBCon.fillrs(szSQL, True)                   ' Daten holen
    If rsPosRelList Is Nothing Then GoTo exithandler                   ' Keine Daten -> Fertig
    
    'For i = 0 To rsPosRelList.Fields.Count - 1
    For i = 0 To 1
        Call AddLVColumn(Me.ListView2, rsPosRelList.Fields(i).Name)    ' Colum hinzufügen
    Next i
    
    Me.ListView2.ColumnHeaders(1).Width = 0                         ' Erste Spalte ausblenden
    Me.ListView2.ColumnHeaders(2).Width = Me.ListView2.Width - 50   ' optimale breite einstelln
    
    Do While Not rsPosRelList.EOF              ' Für jeden DS einen eintrag
        szTmpName = Trim(objTools.checknull(rsPosRelList.Fields(1).Value, ""))
        szTmpID = Trim(objTools.checknull(rsPosRelList.Fields(0).Value, ""))
        If szTmpName <> "" And szTmpID <> "" Then
            Set LVItem = AddListViewItem(Me.ListView2, szTmpID, szTmpID, , GetImageIndexByRootNode(SecRelKey))
            For i = 1 To 1 'rsPosRelList.Fields.Count - 1
                ' Für jedes Feld ein SubItem
                Call AddListViewSubItem(LVItem, Trim(objTools.checknull(rsPosRelList.Fields(i).Value, "")))
            Next i
            szTmpName = ""
        End If  ' szTmpName <> "" And szTmpID <> ""
        rsPosRelList.MoveNext
    Loop
exithandler:
On Error Resume Next

Exit Sub
ErrorHandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.ErrorHandler(MODULNAME, "FillPosRelations", errNr, errDesc)
    Resume exithandler
End Sub



Private Sub ListView1_GotFocus()
    'Me.ListView2.SelectedItem = ""
End Sub


Private Sub ListView2_GotFocus()
    'Me.ListView1.SelectedItem = ""
End Sub
