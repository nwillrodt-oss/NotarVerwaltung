VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmEdit 
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   LinkMode        =   1  'Quelle
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows-Standard
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   5040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Übernehmen"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame FrameEdit 
      BorderStyle     =   0  'Kein
      Height          =   4815
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cmdCMB 
         Caption         =   "..."
         Height          =   310
         Index           =   0
         Left            =   4200
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txt01 
         Height          =   315
         Index           =   0
         Left            =   1900
         TabIndex        =   2
         Top             =   120
         Width           =   3495
      End
      Begin VB.ComboBox cmb01 
         Height          =   315
         Index           =   0
         ItemData        =   "frmEdit.frx":0000
         Left            =   1800
         List            =   "frmEdit.frx":0002
         TabIndex        =   1
         ToolTipText     =   "Hallo Welt"
         Top             =   240
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label lbl01 
         Caption         =   "Label1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList ILEdit 
      Left            =   5160
      Top             =   0
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
            Picture         =   "frmEdit.frx":0004
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":039E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":06F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":0A8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":0DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":1176
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":14C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":1862
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULNAME = "frmEdit"
Private bInit As Boolean            ' Wird True gesetzt wenn Alle werte geladen
Public bDirty As Boolean            ' Wird True gesetzt wenn Daten verändert wurden
Public bNew  As Boolean             ' Wird gesetzt wenn neuer DS sonst Update

Private szSQL As String             ' SQL für a_register
Private szWhere As String           ' Where Klausel
Private szIniFilePath As String     ' Pfad der Ini datei
Private lngImage As Integer         ' Imagiendex

'Private rsUnits As ADODB.Recordset  ' RS aller DS aus szSQLUnit
'Private frmParent As Form          ' Aufrufendes DB form
'Private rsSecure As ADODB.Recordset ' RS aller DS aus szSQLSecure

Private rsBereich As ADODB.Recordset    ' RS für Werteliste der Gerichtsbereiche
Private frmParent As Form           ' Aufrufendes DB form

Private szRootkey As String         ' = z.b. Register
Private szDetailKey As String       ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)

Public ID As String                 ' a_Register.register
Private ThisDBcon As Object         ' Aktuelle DB Verbindung

Public lngFrametopPos As Integer
Public lngFrameLeftPos As Integer
Public lngFrameWidth As Integer
Public lngFrameHeight As Integer

Private Const lngCtlHeight = 315
Private Const lngCtlDiff = 50

Private Sub Form_Load()
    
    Call EditFormLoad(Me, szRootkey)
    
    cmb01(0).Visible = False        ' erstes Feld nie Combo
    'cmb01(0).Left = txt01(0).Left   '
    Me.cmdUpdate.Enabled = False
    
    lngFrametopPos = Me.Top + 360
    lngFrameLeftPos = 120
    lngFrameWidth = Me.Width - 240
    lngFrameHeight = Me.Height - 480
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call EditFormUnload(Me)
End Sub

Public Function InitEditForm(parentform As Form, dbCon As Object, RootKey As String, DetailKey As String)
    
    Dim i As Integer                ' counter
    Dim szWhere2 As String
    Dim DetailKeyArray() As String
    
On Error GoTo Errorhandler

    bInit = True
    Set ThisDBcon = dbCon           ' Aktuelle DB Verbindung übernehmen
    szRootkey = RootKey             ' Welcher TreeRootschlüssel, Legt fest welches Select Statement geholt wird
    
    szDetailKey = DetailKey         ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)
    If szDetailKey = "" Then bNew = True ' Neuer Datensatz
    
    'szIniFilePath = objObjectBag.Getappdir & objObjectBag.GetINIFile
    szIniFilePath = objObjectBag.Getappdir & objObjectBag.GetXMLFile
    'GetEditInfoFromXML(ByVal XMLDocPath As String, NodeName As String, _
            szSQL As String, Optional szIDorder As String)
    Call objTools.GetEditInfoFromXML(szIniFilePath, szRootkey, szSQL, szWhere)
    'szSQL = objTools.GetINIValue(szIniFilePath, INI_EDITSQL, szRootkey)
    'szWhere = objTools.GetINIValue(szIniFilePath, INI_EDITSQL, "WHERE" & szRootkey)
    
'    If InStr(szDetailKey, ";") > 0 Then
'        szWhere2 = objTools.GetINIValue(szIniFilePath, INI_EDITSQL, "WHERE" & szRootkey & "2")
'    End If
    Set ThisDBcon = dbCon           ' Aktuelle DB Verbindung übernehmen
            
    If Not bNew Then
        If szWhere2 <> "" Then
            DetailKeyArray = Split(DetailKey, ";")
            szWhere = szWhere & "'" & DetailKeyArray(0) & "' AND " & szWhere2 & "'" & DetailKeyArray(1) & "'"
        Else
            szWhere = szWhere & "'" & szDetailKey & "'"
        End If

    
    End If
    
    Call InitAdoDC(Me, ThisDBcon, szSQL, szWhere)
    Me.Refresh                      ' Form aktualisieren
    
    If bNew Then                    ' Wenn Neuer DS
        Adodc1.Recordset.AddNew     ' ans RS einen neuen anhängen
        Me.Caption = szRootkey & " - Neuer Datensatz"
    Else                            ' Wenn bestehender DS
         'ID = txtID           ' register als ID merken
         Me.Caption = szRootkey & " - Bearbeiten"
    End If
    
    Call InitEditFrame      ' Reiter 'Algemein' initialisiren
    
    ' Prüfen auf relation Reiter
    
    
exithandler:
    bInit = False
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitEditForm", errNr, errDesc)
    Resume exithandler
End Function

Private Function InitEditFrame()

    'Dim szSQL As String                 ' SQL Statement
    'Dim szSQLValueList As String        ' SQL einer Werteliste (Also CMB feld)
    'Dim szValueListWHERE As String      ' Where statement einer Valueliste
    'Dim szValueListWHEREField As String
    'Dim szDefaultValue As String
    Dim i As Integer                    ' counter
    Dim szFieldName As String           ' Feldname
    Dim szText As String                ' Textwert eines Datenfeldes
    Dim szToolTip As String             ' Tooltip
    Dim szTag As String
    
    Dim lngTop As Integer               ' Obere Pos des Controls
'    Dim szFieldAlias As String
'    Dim rsDefValue As New ADODB.Recordset   ' Recordset
'    Dim ParentFrame As Frame
'    Dim bNotEmpty As Boolean
'    Dim CTLTagArray() As String
    
On Error GoTo Errorhandler


    If Adodc1.Recordset Is Nothing Then GoTo exithandler
    Adodc1.Recordset.MoveFirst
        
    ' 1. Wert holen da sonderbehandlung fürs erste Feld
    ' (sollte immer ein Pflichtfeld ohne werteliste und defaultwert sein)
    szFieldName = Adodc1.Recordset.Fields(0).Name
    If Not bNew Then szText = objTools.checknull(Adodc1.Recordset.Fields(0).Value, "")
    szToolTip = Adodc1.Recordset.Fields(0).Name & " (" & Adodc1.Recordset.Fields(0).DefinedSize & " Zeichen)"
    lngTop = 0
      'szTag = BulidCTLTag(szUpdateTable, objSQLTools.GetFieldFromRSName(szSQL, szFieldName), szFieldName, szText)
'    CTLTagArray = Split(szTag, ";")
    'szTag = objSQLTools.GetFieldFromRSName(szSQL, rsList.Fields(0).Name)
'    bNotEmpty = IsNotEmptyField(CTLTagArray(1))


     ' 1. Feld behandeln
    'Call InitLabel(lbl01(0), szFieldName, bNotEmpty)
    Call InitLabel(lbl01(0), szFieldName, False)
    Call InitTextBox(txt01(0), szText, lngTop, szToolTip, szTag, 0)
    
    ' Dann weitere clonen
    For i = 1 To Adodc1.Recordset.Fields.Count - 1
        ' Werte holen
        szFieldName = Adodc1.Recordset.Fields(i).Name
        If Not bNew Then szText = objTools.checknull(Adodc1.Recordset.Fields(i).Value, "")
        szToolTip = Adodc1.Recordset.Fields(i).Name & " (" & Adodc1.Recordset.Fields(i).DefinedSize & " Zeichen)"
        lngTop = lbl01(i - 1).Top + lngCtlHeight + lngCtlDiff
            Load lbl01(i)
            Load txt01(i)
            Load cmdCMB(i)
'            Call InitLabel(lbl01(i), szFieldName, bNotEmpty, lngTop)   ' Label initialisieren
            Call InitLabel(lbl01(i), szFieldName, False, lngTop)   ' Label initialisieren
            Call InitTextBox(txt01(i), szText, lngTop, szToolTip, szTag)    ' Textbox mit wert füllen und Feldname als Tag setzen
            cmdCMB(i).Visible = False
            'Adodc1.Recordset.MoveNext
    Next i

exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitEditFrame", errNr, errDesc)
    Resume exithandler
End Function

Public Function InitLabel(lblCTL As Label, szCaption As String, bNotEmpty As Boolean, _
            Optional lngTop As Integer)

On Error GoTo Errorhandler
    
    'Set lblCTL.Parent = FrameEdit(0)
    If bNotEmpty Then
        lblCTL.Caption = szCaption & " *"
    Else
        lblCTL.Caption = szCaption
    End If
    lblCTL.Height = lngCtlHeight
    lblCTL.Visible = True
    If lngTop > 0 Then
        lblCTL.Top = lngTop
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
    Call objError.Errorhandler(MODULNAME, "InitLabel", errNr, errDesc)
    Resume exithandler
End Function

Public Function InitTextBox(txtCTL As TextBox, szText As String, _
        Optional lngTop As Integer, _
        Optional szToolTip As String, _
        Optional szTag As String, _
        Optional lngFIndex As Integer)

On Error GoTo Errorhandler
    
    txtCTL.Text = szText
    txtCTL.Height = lngCtlHeight
    txtCTL.ToolTipText = szToolTip
    txtCTL.Tag = szTag
    If lngTop > 0 Then
        txtCTL.Top = lngTop
    End If
    txtCTL.Visible = True
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitTextBox", errNr, errDesc)
    Resume exithandler
End Function

Public Function InitComboBox(Index As Integer, szSQLValueList As String, szText As String, _
        Optional lngTop As Integer, _
        Optional szToolTip As String, _
        Optional szTag As String)

On Error GoTo Errorhandler

    ' "ComboBox" ist hier nur ein mit einem zusätzlichen Button
    Call InitTextBox(txt01(Index), szText, lngTop, szToolTip, szTag)    ' Textbox mit wert füllen und Feldname als Tag setzen
    ' Button behandeln
    cmdCMB(Index).Top = lngTop
    cmdCMB(Index).ToolTipText = szToolTip
    cmdCMB(Index).Tag = szSQLValueList
    cmdCMB(Index).Left = txt01(Index).Left + txt01(Index).Width
    cmdCMB(Index).Visible = True

exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitComboBox", errNr, errDesc)
    Resume exithandler
End Function

'Private Function InitValuelist(szSQL As String, TBox As TextBox)
'
'Dim rsValList As ADODB.Recordset
'Dim i As Integer ' counter
'Dim MaxListheight As Integer
''Dim bListUp As Boolean
'
'Dim ValulistItem As ListItem
'
'Dim frmValueListEdit As Form
'
'On Error GoTo Errorhandler
'
' ' Kein SQL Statement -> Fertig
'    If szSQL = "" Then GoTo exithandler
'
'Set frmValueListEdit = New frmValuelist
'
'    frmValueListEdit.txtValue.Text = TBox.Text
'    frmValueListEdit.CtlIndex = TBox.Index
'    Set frmValueListEdit.fEdit = Me
'    frmValueListEdit.Tag = TBox.Tag
'
'    ' Daten holen
'    Set rsValList = ThisDBCon.fillrs(szSQL)
'    rsValList.MoveFirst
'    i = 0
'
'    frmValueListEdit.LVValuelist.ColumnHeaders.Clear
'    If rsValList.Fields.Count > 1 Then
'        frmValueListEdit.LVValuelist.ColumnHeaders.Add , , , 300
'        frmValueListEdit.LVValuelist.ColumnHeaders.Add , , , frmValueListEdit.LVValuelist.Width - 600
'    Else
'        frmValueListEdit.LVValuelist.ColumnHeaders.Add , , , frmValueListEdit.LVValuelist.Width - 300
'    End If
'
'    While Not rsValList.EOF
'        If rsValList.Fields.Count > 1 Then
'            Set ValulistItem = AddListViewItem(frmValueListEdit.LVValuelist, objTools.checknull(rsValList.Fields(0).Value, 0))
'
'            Call AddListViewSubItem(ValulistItem, objTools.checknull(rsValList.Fields(1).Value, ""))
'        Else
'            Set ValulistItem = AddListViewItem(frmValueListEdit.LVValuelist, objTools.checknull(rsValList.Fields(0).Value, ""))
'        End If
'        i = i + 1
'        rsValList.MoveNext
'    Wend
'
'    MaxListheight = Me.Height
'    If (i * 255) + 55 > MaxListheight Then
'        frmValueListEdit.LVValuelist.Height = MaxListheight
'    Else
'        frmValueListEdit.LVValuelist.Height = (i * 225) + 55
'    End If
'
'    frmValueListEdit.Height = frmValueListEdit.txtValue.Top + _
'        frmValueListEdit.txtValue.Height + 50 + frmValueListEdit.LVValuelist.Height + 500
''    If frmValueListEdit.LVValuelist.Top > FrameEdit(0).Height - (frmValueListEdit.LVValuelist.Top + txt01(0).Height) Then
''        bListUp = True
''        MaxListheight = frmValueListEdit.LVValuelist.Top
''    Else
''        MaxListheight = FrameEdit(0).Height - (frmValueListEdit.LVValuelist.Top + txt01(0).Height)
''    End If
''
''    If (i * 200) + 55 > MaxListheight Then
''        frmValueListEdit.LVValuelist.Height = MaxListheight
''    Else
''        frmValueListEdit.LVValuelist.Height = (i * 255) + 55
''    End If
'
'    frmValueListEdit.Show 1, Me
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
'    Call objError.Errorhandler(MODULNAME, "InitValuelist", errNr, errDesc)
'    Resume exithandler
'End Function

'Private Function InitRelationlist(szSQL As String)
'
'Dim rsValList As ADODB.Recordset
'Dim i As Integer ' counter
'Dim MaxListheight As Integer
''Dim bListUp As Boolean
'
'Dim ValulistItem As ListItem
'
'Dim frmValueListEdit As Form
'
'On Error GoTo Errorhandler
'
' ' Kein SQL Statement -> Fertig
'    If szSQL = "" Then GoTo exithandler
'
'Set frmValueListEdit = New frmValuelist
'
'    'frmValueListEdit.txtValue.Text = TBox.Text
'    'frmValueListEdit.CtlIndex = TBox.Index
'    Set frmValueListEdit.fEdit = Me
'    'frmValueListEdit.Tag = TBox.Tag
'
'    ' Daten holen
'    Set rsValList = ThisDBCon.fillrs(szSQL)
'    rsValList.MoveFirst
'    i = 0
'
'    frmValueListEdit.LVValuelist.ColumnHeaders.Clear
'    If rsValList.Fields.Count > 1 Then
'        frmValueListEdit.LVValuelist.ColumnHeaders.Add , , , 300
'        frmValueListEdit.LVValuelist.ColumnHeaders.Add , , , frmValueListEdit.LVValuelist.Width - 600
'    Else
'        frmValueListEdit.LVValuelist.ColumnHeaders.Add , , , frmValueListEdit.LVValuelist.Width - 300
'    End If
'
'    While Not rsValList.EOF
'        If rsValList.Fields.Count > 1 Then
'            Set ValulistItem = AddListViewItem(frmValueListEdit.LVValuelist, objTools.checknull(rsValList.Fields(0).Value, 0))
'
'            Call AddListViewSubItem(ValulistItem, objTools.checknull(rsValList.Fields(1).Value, ""))
'        Else
'            Set ValulistItem = AddListViewItem(frmValueListEdit.LVValuelist, objTools.checknull(rsValList.Fields(0).Value, ""))
'        End If
'        i = i + 1
'        rsValList.MoveNext
'    Wend
'
'    MaxListheight = Me.Height
'    If (i * 255) + 55 > MaxListheight Then
'        frmValueListEdit.LVValuelist.Height = MaxListheight
'    Else
'        frmValueListEdit.LVValuelist.Height = (i * 225) + 55
'    End If
'
'    frmValueListEdit.Height = frmValueListEdit.txtValue.Top + _
'        frmValueListEdit.txtValue.Height + 50 + frmValueListEdit.LVValuelist.Height + 500
'
'    frmValueListEdit.Show 1, Me
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
'    Call objError.Errorhandler(MODULNAME, "InitRelationlist", errNr, errDesc)
'    Resume exithandler
'End Function

'Private Function UpdateEditForm()
'
'    Dim szSQL As String         ' SQL Statement
'    Dim szValuelist As String   ' Update Werte liste
'    Dim szWhere As String       ' Where Part des SQL Statements
'    Dim i As Integer            ' Counter
'    Dim szFieldAlias As String
'On Error GoTo Errorhandler
'
'' !!!  Wichtig wird das Erste feld geändert müß der Tree im DB Form refresht werden da sonst der Nodename nicht mehr stimmt
'
'    If Not bDirty Then GoTo exithandler             ' Keine Änderungen -> Raus
'
'    If Not ValidateEditForm Then GoTo exithandler   ' Prüfe ob Änderungen Zulässig
'
'    If bNew Then        ' Neuer DS -> Insert
'        rsList.AddNew                               ' Neuen DS an RS anfügen
'
'    Else                ' bestehender DS -> Update
''        For i = 0 To lbl01.Count - 1    ' Alle Felder durchgehen
''            szFieldAlias = GetAliasFromTag(txt01(i).Tag)
''            If szFieldAlias <> "" Then rsList.Fields(szFieldAlias).Value = txt01(i).Text
'
''            If InStr(szLockedFiledList, txt01(i).Tag) = 0 Then      ' Gesperrte Felder Auslassen
''                If Trim(txt01(i).Text) <> "" Then                   ' Steht überhaupt was in Textfeld
''                    If IsNumeric(txt01(i).Text) Then                ' Zahlen ohne Hochkomma
''                        szValuelist = szValuelist & txt01(i).Tag & "= " & txt01(i).Text & ","
''                    Else                                            ' Sonst mit hochkomma
''                        szValuelist = szValuelist & txt01(i).Tag & "= '" & txt01(i).Text & "',"
''                    End If
''                End If
''            End If
''        Next i
'
''        szValuelist = objTools.cutlastchar(szValuelist, ",")        ' Letztes Komma abschneiden
''        If szDetailKey <> "" Then szWhere = " " & objSQLTools.AddWhere("", _
''                objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_SQL, "WHERE" & szRootKey) _
''                & "'" & szDetailKey & "'")
''        szSQL = objSQLTools.BuildUpdateSQL(szUpdateTable, szValuelist, szWhere) ' SQL Statement zusammen setzen
''        Call ThisDBCon.execsql(szSQL)       ' Ausführen
'
'    End If
'
'    For i = 0 To lbl01.Count - 1                    ' Alle Felder durchgehen
'        szFieldAlias = GetAliasFromTag(txt01(i).Tag)
'        If szFieldAlias <> "" Then
'            rsList.Fields(szFieldAlias).Value = txt01(i).Text
'        End If
'    Next i
'
'    rsList.Update                                   ' Recordset Updaten
'
'    bDirty = False                                  ' Damit ist das Form nicht mehr Dirty
'    Call CheckUpdate                                ' Übernehmen button disablen
'
'exithandler:
'On Error Resume Next
'
'Exit Function
'Errorhandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.Errorhandler(MODULNAME, "UpdateEditForm", errNr, errDesc)
'    Resume exithandler
'End Function

Private Function IsNotEmptyField(szFieldName) As Boolean
    
'    If InStr(szNotEmptyList, UCase(szFieldName)) > 0 Then
'        IsNotEmptyField = True
'    Else
'        IsNotEmptyField = False
'    End If
    
End Function

Private Sub CheckUpdate()
    Me.cmdUpdate.Enabled = bDirty
End Sub

Private Sub HandleKeyDown(frmEdit As Form, KeyCode As Integer, Shift As Integer)
    
On Error Resume Next
    
    Call HandleKeyDownEdit(Me, KeyCode, Shift)
    Call frmParent.HandleGlobalKeyCodes(KeyCode, Shift)

End Sub

Private Sub cmdOK_Click()
    Call UpdateEditForm(Me, szRootkey)
    Call frmParent.RefreshListView                      ' ListView aktualisieren
    Me.Hide
End Sub

Public Sub cmdUpdate_Click()
    Call UpdateEditForm(Me, szRootkey)
    Call frmParent.RefreshListView                      ' ListView aktualisieren
End Sub

Private Sub cmdCMB_Click(Index As Integer)
    'Call InitValuelist(cmdCMB(Index).Tag, txt01(Index))
End Sub

Private Sub cmb01_Change(Index As Integer)
    If bInit Then Exit Sub
    bDirty = True
    Call CheckUpdate
End Sub

Private Sub cmb01_Validate(Index As Integer, Cancel As Boolean)
    If bInit Then Exit Sub
    bDirty = True
    Call CheckUpdate
End Sub

Private Sub cmdESC_Click()
    Me.Hide
End Sub

Private Sub txt01_Change(Index As Integer)
    If bInit Then Exit Sub
    bDirty = True
    Call CheckUpdate
End Sub

Private Sub txt01_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

'    If LVcmb.Visible Then
'        If LVcmb.Tag <> Index Then
'            LVcmb.Visible = False
'        End If
'    End If
End Sub

Private Function BulidCTLTag(szTabName As String, _
        szFieldName As String, _
        szFieldAlias As String, _
        szValue As String) As String

    BulidCTLTag = szTabName & ";" & szFieldName & ";" & szFieldAlias & ";" & szValue
End Function

Public Function GetValueFromTag(szTag As String) As String

    Dim szTagArray() As String
    Dim i As Integer            ' Counter
    
On Error GoTo Errorhandler

    If szTag = "" Then GoTo exithandler
    
    szTagArray = Split(szTag, ";")
    GetValueFromTag = szTagArray(3)
    
exithandler:
    bInit = False
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetValueFromTag", errNr, errDesc)
    Resume exithandler
End Function

Public Function GetAliasFromTag(szTag As String) As String

    Dim szTagArray() As String
    Dim i As Integer            ' Counter
    
On Error GoTo Errorhandler

    If szTag = "" Then GoTo exithandler
    
    szTagArray = Split(szTag, ";")
    GetAliasFromTag = szTagArray(2)
    
exithandler:
    bInit = False
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetAliasFromTag", errNr, errDesc)
    Resume exithandler
End Function

Public Function GetFieldFromTag(szTag As String) As String

    Dim szTagArray() As String
    Dim i As Integer            ' Counter
    
On Error GoTo Errorhandler

    If szTag = "" Then GoTo exithandler
    
    szTagArray = Split(szTag, ";")
    GetFieldFromTag = szTagArray(1)
    
exithandler:
    'bInit = False
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetFieldFromTag", errNr, errDesc)
    Resume exithandler
End Function


