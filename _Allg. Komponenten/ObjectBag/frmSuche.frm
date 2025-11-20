VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSuche 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Suchen"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdGetResult 
      Caption         =   "Ergebnis Übernehmen"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.ListView LVResult 
      Height          =   4995
      Left            =   0
      TabIndex        =   7
      Top             =   1440
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8811
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ILSuche"
      SmallIcons      =   "ILSuche"
      ColHdrIcons     =   "ILSuche"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdUp 
      Height          =   315
      Left            =   480
      Picture         =   "frmSuche.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   6
      Top             =   600
      Width           =   315
   End
   Begin VB.CommandButton cmddown 
      Height          =   315
      Left            =   120
      Picture         =   "frmSuche.frx":038A
      Style           =   1  'Grafisch
      TabIndex        =   4
      Top             =   600
      Width           =   315
   End
   Begin VB.TextBox txtSuchbegriff 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox cmbSuchenNach 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin MSComctlLib.ImageList ILSuche 
      Left            =   1080
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":0714
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":0A2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":0D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":12E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":187C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":2556
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":3230
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":37CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":3D64
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":40FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":4498
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":44F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":4554
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":4AEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":5088
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":5622
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":5BBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":6156
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSuche.frx":66F0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblSucheNach 
      Caption         =   "Suchen nach:"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmSuche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MODULNAME = "frmSuche"                                ' Modulname für Fehlerbehandlung

Const INI_SEARCH = "Search"
Const IMG_SORTDOWN = 9
Const IMG_SORTUP = 10

Private szSQLDefault As String                                      ' SQl Statement ohne suchbegriff
Private szSQL As String                                             ' SQl Statement mit suchbegriff
Private szIniFilePath As String                                     ' Pfad der Ini datei

Private objObjectBag As clsObjectBag                                ' ObjectBag object
Private objError As Object
Private objSQLTools As Object
Private objTools As Object
Private objOptions As Object
Private objRegTools As Object
Private DBConn As Object                                            ' Akt DB Verbindung

Private bResult As Boolean                                          ' True wenn ergebiss ListView sichtbar
Private bErweitert As Boolean                                       ' True wenn erweiterte suche sichtbar

Private rsResult As Object                                          ' Ergebniss Recordset
Private szOptWhereID As String

Public szRootkey As String                                          ' Wonach wird gesucht
Public szSearchField As String                                      ' In welchem Feld wird gesucht
Public szResultID As String                                         ' ID des Ergebniss Datensatzes
Private szSearchTitel As String                                     ' Optionaler Fenster Titel der Suche

Private lngHeightErweitert As Integer
Private lngHeightResult As Integer
Private lngHeightDefault As Integer
Private lngHeightErweitertResult As Integer

Private Type SearchInfo
    szDefaultSQL As String
    szSQL As String
    szWhere As String
    OptWhere As String
    Field As String
    Suchname As String
    bInUserList As Boolean
End Type


Private Sub Form_Load()
    Call ShowErweitert
    'txtSuchbegriff.SetFocus
End Sub

Public Function InitSuche(OBag As Object, _
        ThisDBConn As Object, _
        Rootkey As String, SearchField As String, _
        Optional szSearchtext As String, _
        Optional OptWhereID As String, _
        Optional szTitel As String) As Boolean

    Dim szDefaultSearch As String                                   ' Name der evtl standart suche
    Dim szDefaultSearchfield As String                              ' Wenn SearchField ="" erstes mögliches
        
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    lngHeightDefault = 1425                                         ' Suchfenster größe erstmal Statisch
    lngHeightErweitert = 1800
    lngHeightResult = 6465
    lngHeightErweitertResult = 6800

    bResult = False                                                 ' Wir haben gerade angefangen daher kein ergebnis
    Set objObjectBag = OBag                                         ' Benötigte Objekte holen
    Set objError = objObjectBag.GetErrorObj()                       ' Error Obj
    Set objSQLTools = objObjectBag.GetSQLToolsObj()
    Set objTools = objObjectBag.GetToolsObj()
    Set objOptions = objObjectBag.GetOptionsObj()
    Set objRegTools = objObjectBag.GetRegToolsObj()
    Set DBConn = ThisDBConn
    
    szOptWhereID = OptWhereID
    szRootkey = Trim(Rootkey)                                       ' rootkey gibt an was wir suchen
    szSearchField = Trim(SearchField)                               ' Serachfield gibt an wonach
    szIniFilePath = objObjectBag.GetAppDir & objObjectBag.GetXMlFile ' XML inifile festlegen
    szSearchTitel = szTitel
    szDefaultSearchfield = FillSearchlist(szRootkey)                ' Liste der Möglichen Suchen erlitteln bzw. 1. Mögliches Searchfield
    If szSearchField = "" Then szSearchField = szDefaultSearchfield
    
    If szRootkey <> "" And szSearchField <> "" Then Me.cmbSuchenNach.Text = szRootkey & " " & szSearchField
    
    If Me.cmbSuchenNach.Text = "" Then                              ' wenn kein Rootkey und kein Searchfield
        bErweitert = True                                           ' Dann erweiterter Dialog zur auswahl
        Call ShowErweitert
        szDefaultSearch = objOptions.GetOptionByName("DefaultSearch")   ' Standart suche in Optionen eingetragen?
                                                                    ' (evtl. später auch zuletzt ausgewählte Suche???)
        If szDefaultSearch <> "" Then
            Me.cmbSuchenNach.Text = szDefaultSearch
            Call GetRootKey                                         ' Caption Setzten und Glob Var festlegen
        Else
            Me.Caption = "Suche " & szSearchTitel
        End If
    Else
        bErweitert = False                                          ' erweiterter Dialog ausblenden
        Call ShowErweitert
        Me.cmbSuchenNach.Text = szRootkey & " " & szSearchField
        Call GetRootKey                                             ' Caption Setzten und Glob Var festlegen
        'Me.Caption = "Suche " & szRootkey & " nach " & szSearchField
    End If
    
    txtSuchbegriff.Text = szSearchtext                              ' Mitgelieferter Statischer Suche Titel
    If txtSuchbegriff.Text <> "" Then Call Suchen                   ' Suchtest Schon da -> sofort lossuchen
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitSuche", errNr, errDesc)
    Resume exithandler
End Function

Private Function FillSearchlist(Optional Key As String) As String
' Auswahlliste der möglichen Suchen füllen
' Gibt ersten möglichen wert aus default zurück

    Dim sztmp As String
    Dim listArray()  As String                                      ' Array mit listen einträgen
    Dim i As Integer                                                ' Counter
    Dim szDefaultSearchfield As String
    
On Error GoTo Errorhandler

    cmbSuchenNach.Clear                                             ' evtl. alte Liste löschen
    
    'sztmp = objTools.GetINIValue(szIniFilePath, INI_SEARCH, "SearchList")
    sztmp = objTools.GetSearchListFromXML(szIniFilePath)
    If sztmp <> "" Then
        listArray = Split(sztmp, ",")
        
        For i = 0 To UBound(listArray)                              ' Array duchlaufen
            Me.cmbSuchenNach.AddItem listArray(i)                   ' Werte eintragen
            If Key <> "" And szDefaultSearchfield = "" Then
                If InStr(listArray(i), Key) > 0 Then
                    szDefaultSearchfield = Replace(listArray(i), Key, "")
                    szDefaultSearchfield = Replace(szDefaultSearchfield, " ", "")
                End If
            End If
        Next
        
        FillSearchlist = szDefaultSearchfield                       ' 1. mögliches Searchfield zurück geben
    End If
    
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "FillSearchlist", errNr, errDesc)
    Resume exithandler
End Function

Private Sub GetRootKey()
    
    Dim tmparray() As String

On Error Resume Next

    If cmbSuchenNach.Text <> "" Then
        tmparray = Split(cmbSuchenNach.Text, " ")

        szRootkey = tmparray(0)
        szSearchField = tmparray(1)
        If Err <> 0 Then
            Err.Clear
'            GoTo exithandler
        End If
        If szSearchTitel = "" Then
            Me.Caption = "Suche " & szRootkey & " nach " & szSearchField
        Else
            Me.Caption = szSearchTitel
        End If
    End If
    
End Sub

Private Sub Suchen()
' Starte die eigentliche suche und liefert ein ergebniss LV
    Dim szSuchtext As String                                        ' Such text
    Dim szRealSearchField As String                                 ' Reales Feld in dem gesucht wird
    Dim szWhere As String                                           ' Where Statement
    Dim SInfo As SearchInfo
On Error GoTo Errorhandler                                          ' Feherbehandlung aktivieren
    szSQLDefault = ""                                               ' Hier Steht evtl. noch Dreck aus anderen
    szSQL = ""                                                      ' Suchgängen drin - > Löschen
    txtSuchbegriff.Text = Trim(txtSuchbegriff.Text)                 ' Suchbegriff Trimmen
    If cmbSuchenNach.Text = "" Then GoTo exithandler                ' Wissen Wir onach wir suchen ?
    Call GetRootKey
    szSuchtext = txtSuchbegriff.Text                                ' Suchtext eingabe merken
    If Right(szSuchtext, 1) <> "*" Then                             ' Prüfen ob Wildcard am schluss
        szSuchtext = szSuchtext & "*"                               ' Wenn nicht ranhängen
    End If
    szSuchtext = objSQLTools.ReplaceWidcarts(DBConn, szSuchtext)    ' Evtl. Widcard füre SQL oder Access anpassen
'    szSQLDefault = objTools.GetSearchSQLFromXML(szIniFilePath, szRootkey)
'    If szSQLDefault = "" Then GoTo exithandler
    With SInfo
        Call objTools.GetSearchInfoFromXML(szIniFilePath, szRootkey, szSearchField, _
            .szDefaultSQL, _
            .OptWhere, _
            .Suchname, _
            .Field, _
            .bInUserList)                                           ' Such informationen aus XML lesen
'    szRealSearchField = objTools.GetRealSearchFieldFromXML(szIniFilePath, szRootkey, szSearchField)
'    If szRealSearchField <> "" Then
        If .Field <> "" Then
            .szWhere = .Field & " like '" & szSuchtext & "'"
            .szSQL = objSQLTools.AddWhereInFullSQL(.szDefaultSQL, .szWhere)
        End If
    
        If szOptWhereID <> "" And .OptWhere <> "" Then
        'szWhere = objTools.GetINIValue(szIniFilePath, INI_SEARCH, "WHERE" & cmbSuchenNach.Text)
            'If szWhere <> "" Then
            'szWhere = szWhere & "'" & szOptWhereID & "'"
                .szSQL = objSQLTools.AddWhereInFullSQL(.szSQL, .OptWhere & "'" & szOptWhereID & "'")
            'End If
        End If
    
        Set rsResult = DBConn.fillrs(.szSQL)
        If rsResult Is Nothing Then GoTo exithandler
    
    End With
        
    Me.LVResult.Tag = cmbSuchenNach.Text
    Call ShowResult
    
exithandler:
On Error Resume Next

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "Suchen", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub ShowResult()

    Dim i As Integer        ' counter
    Dim LVItem As ListItem
    
On Error GoTo Errorhandler

    If rsResult Is Nothing Then GoTo exithandler
    Me.LVResult.ListItems.Clear
    Me.LVResult.ColumnHeaders.Clear
    
    For i = 0 To rsResult.Fields.Count - 1    ' Erste Spalte ist ID -> auslassen
        Me.LVResult.ColumnHeaders.Add i + 1, rsResult.Fields(i).Name, rsResult.Fields(i).Name
        'Call AddLVColumn(LVResult, rsResult.Fields(i).Name) ' Colum hinzufügen
    Next i
    
    'Call SetColumnWidth(LVResult, 1, 0)
    LVResult.ColumnHeaders(1).Width = 0
    
    If Not rsResult.EOF Then
        rsResult.MoveFirst
        While Not rsResult.EOF
            Set LVItem = AddListViewItem(Me.LVResult, rsResult.Fields(0), rsResult.Fields(0), rsResult.Fields(1))
            For i = 2 To rsResult.Fields.Count - 1
                ' Für jedes Feld ein SubItem
                Call AddListViewSubItem(LVItem, Trim(objTools.checknull(rsResult.Fields(i).Value, "")))
            Next i
            rsResult.MoveNext
        Wend
        Me.cmdGetResult.Visible = True
    End If
    bResult = True
    Call ShowErweitert
    Call LoadColumnWidth(LVResult, "SEARCH")
    
exithandler:
On Error Resume Next
    LVResult.ListItems(1).Selected = True
    If bResult Then LVResult.SetFocus
    If Err Then Err.Clear
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ShowResult", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub ShowErweitert()
    
On Error GoTo Errorhandler

    lngHeightDefault = 1425
    lngHeightErweitert = 1800
    lngHeightResult = 6465
    lngHeightErweitertResult = 6800

    
    Me.cmbSuchenNach.Visible = bErweitert
    Me.lblSucheNach.Visible = bErweitert
    Me.cmdUp.Enabled = bErweitert
    Me.cmddown.Enabled = Not bErweitert
    If bErweitert Then
        Me.cmddown.Top = 960
        Me.cmdUp.Top = 960
        Me.cmdOK.Top = 960
        Me.cmdGetResult.Top = 960
        Me.cmdEsc.Top = 960

    Else
        Me.Height = lngHeightDefault
        Me.cmddown.Top = 600
        Me.cmdUp.Top = 600
        Me.cmdOK.Top = 600
        Me.cmdGetResult.Top = 600
        Me.cmdEsc.Top = 600
    End If
    
    Me.LVResult.Visible = bResult
    If bResult Then
        If bErweitert Then
            Me.LVResult.Top = 1440
            Me.Height = lngHeightErweitertResult
        Else
            Me.LVResult.Top = 1080
            Me.Height = lngHeightResult
        End If
    Else
        Me.Height = lngHeightDefault
        If bErweitert Then Me.Height = lngHeightErweitert
    End If
        
    
exithandler:
On Error Resume Next

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ShowErweitert", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub GetResult()
    szResultID = LVResult.SelectedItem.Text
    Call CloseSearch
End Sub

Private Sub CloseSearch()
    If bResult Then Call SaveColumnWidth(LVResult, "SEARCH")
    Me.Hide
    Unload Me
End Sub

Private Sub LVResult_DblClick()
    'Stop
    Call GetResult
End Sub

Private Sub LVResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SetColumnOrder(LVResult, ColumnHeader)
End Sub
                                                                    ' *****************************************
                                                                    ' Botton Events
Private Sub cmdOK_Click()
    Call Suchen
End Sub

Private Sub cmdEsc_Click()
    szResultID = ""
    Call CloseSearch
End Sub

Private Sub cmdGetResult_Click()
    Call GetResult                                                  ' Ergebnis übernehmen
End Sub

Private Sub cmddown_Click()
    bErweitert = True
    Call ShowErweitert
End Sub

Private Sub cmdUp_Click()
    bErweitert = False
    Call ShowErweitert
End Sub

Private Sub cmbSuchenNach_Validate(Cancel As Boolean)
    Call GetRootKey
    Me.LVResult.Tag = cmbSuchenNach.Text
End Sub
                                                                    ' *****************************************
                                                                    ' Key Events
Private Sub LVResult_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 And Shift = 0 And LVResult.SelectedItem <> "" Then ' Enter
        szResultID = LVResult.SelectedItem.Text
        If bResult Then Call SaveColumnWidth(LVResult, "SEARCH")
        Me.Hide
    End If
End Sub

Private Sub txtSuchbegriff_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Then Unload Me                                  ' ESC
    If KeyCode = 13 And Shift = 0 Then ' Enter
        Call Suchen
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me                                  ' ESC
End Sub

Public Function SetColumnOrder(LV As ListView, ColumnHeader As MSComctlLib.ColumnHeader)
' Sortiert die angegebene Spalte ja nach eltl forhandener Sortierung entgegengesetzt
' Auch Datm und Zahlen werden korrekt sortiert, wenn im Tag des Headers
' die konstaten adInteger oder ad Date eingetragen sind

    Dim i As Integer                                                ' Counter (listItems)
    Dim X As Integer                                                ' noch ein Counter (Column header)
    Dim NewSub As Long                                              ' Index der Dummy Spalte bei Datum oder Zahl
    Dim sFormat As String                                           ' Sortierbares Datumsformat
    Dim Li As ListItem                                              ' LI Current Listview Item

On Error GoTo Errorhandler

    ' Sort Order 0 Aufsteigend
    ' Sort Order 1 Absteigend

    sFormat = "yyyy.mm.dd hh:mm:ss"                                 ' Sortierbares Datumsformat setzen
    i = 0

    With LV
        .Visible = False                                            ' ListView ruhig halten, Sichtbarkeit bleibt trotzdem erhalten

        For X = 1 To .ColumnHeaders.Count                           ' Erst alle icons ausblenden
            .ColumnHeaders(X).Icon = 0
        Next X

        .SortKey = ColumnHeader.Index - 1                           ' zu sortierende Spalte bestimmen
        .ColumnHeaders.Add , , "Dummy", 0                           ' Dummy-Spalte einfügen mit Breite 0
        NewSub = .ColumnHeaders.Count - 1                           ' Nummer der Dummy-Spalte

        If ColumnHeader.Tag = 7 Then                                ' abfragen auf Spalte mit Datum (adDate=7)

            For i = .ListItems.Count To 1 Step -1                   ' Sortiere nach Datum
                Set Li = .ListItems(i)
                'Dummy-Spalte mit sortierfähigem Datum belegen
                Li.SubItems(NewSub) = Format(CDate(Li.SubItems(ColumnHeader.Index - 1)), sFormat)
            Next i
            .SortKey = NewSub                                       ' zu sortierende Spalte umbiegen
        ElseIf ColumnHeader.Tag = 3 Then                            ' abfragen auf Spalte mit Zahlen (adInteger=3)
            For i = .ListItems.Count To 1 Step -1                   ' Sortiere nach Zahlen
                Set Li = .ListItems(i)
                'Dummy-Spalte mit sortierfähiger Zahl belegen
                Li.SubItems(NewSub) = Right(Space(20) & Li.SubItems(ColumnHeader.Index - 1), 20)
            Next i
            .SortKey = NewSub                                       ' zu sortierende Spalte umbiegen
        End If

        If .SortOrder = 0 Then                                      ' SortOrder bestimmen Asc oder Desc
            .SortOrder = 1
            .ColumnHeaders(ColumnHeader.Index).Icon = IMG_SORTUP
        Else
            .SortOrder = 0
            .ColumnHeaders(ColumnHeader.Index).Icon = IMG_SORTDOWN
        End If
        .Sorted = True                                              ' Sort anstossen
        .ColumnHeaders.Remove .ColumnHeaders.Count                  ' Dummy-Spalte entfernen
        .ListItems(1).Selected = True                               ' Zeiger auf 1. Zeile und scrollen
        .ListItems(1).EnsureVisible
        .Visible = True                                             ' sichtbar machen
    End With

exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SetColumnOrder", errNr, errDesc)
End Function

Public Sub SaveColumnWidth(LV As ListView, Optional TagPreFix As String, Optional bForce As Boolean)

    Dim i As Integer            ' Counter
    Dim szRegKey As String
    Dim szRegValue As String
    
On Error GoTo Errorhandler
    
    If LV.Tag = "" Then GoTo exithandler
    szRegKey = "SOFTWARE\" & objObjectBag.GetAppTitle & "\Columns"
    If LV.ColumnHeaders.Count > 2 Or bForce Then
        For i = 1 To LV.ColumnHeaders.Count
            szRegValue = szRegValue & LV.ColumnHeaders(i).Width & ";"
        Next i
        
        If szRegValue <> "" Then Call objRegTools.WriteRegValue("HKCU", szRegKey, TagPreFix & LV.Tag, Left(szRegValue, Len(szRegValue) - 1))
    End If
exithandler:

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SaveColumnWidth", errNr, errDesc)
    Resume exithandler
End Sub

Public Sub LoadColumnWidth(LV As ListView, Optional TagPreFix As String, Optional bForce As Boolean)

    Dim i As Integer        ' Counter
    Dim szRegKey As String
    Dim szRegValue As String
    Dim ColArray() As String
    
On Error GoTo Errorhandler
    
    If LV.Tag = "" Then GoTo exithandler
    szRegKey = "SOFTWARE\" & objObjectBag.GetAppTitle & "\Columns"
    szRegValue = objRegTools.ReadRegValue("HKCU", szRegKey, TagPreFix & LV.Tag)
    If szRegValue <> "" Then
        If LV.ColumnHeaders.Count > 2 Or bForce Then
            ColArray = Split(szRegValue, ";")
            For i = 0 To UBound(ColArray)
                If i < LV.ColumnHeaders.Count Then
                    Call SetColumnWidth(LV, i + 1, CLng(ColArray(i)))
                End If
            Next i
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
    Call objError.Errorhandler(MODULNAME, "LoadColumnWidth", errNr, errDesc)
    Resume exithandler
End Sub

Public Sub SetColumnWidth(ctlListView As ListView, ColIndex As Integer, Width As Integer)
    ctlListView.ColumnHeaders(ColIndex).Width = Width   ' optimale breite einstelln
End Sub
