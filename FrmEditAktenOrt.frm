VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form FrmEditAktenOrt 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Aktenbewegung"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6870
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   0
      Picture         =   "FrmEditAktenOrt.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Datensatz speichern"
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   375
      Left            =   480
      Picture         =   "FrmEditAktenOrt.frx":058A
      Style           =   1  'Grafisch
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Datensatz löschen"
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdWord 
      Height          =   375
      Left            =   960
      Picture         =   "FrmEditAktenOrt.frx":0914
      Style           =   1  'Grafisch
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Neues Anschreiben"
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txtPerson 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdPersonrSuchen 
      Caption         =   "Suchen"
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtIDPers 
      DataField       =   "fk010017"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame FrameAktenort 
      Caption         =   "Aktenort"
      Height          =   1935
      Left            =   120
      TabIndex        =   22
      Top             =   1200
      Width           =   6375
      Begin VB.TextBox txtEingetragenVon 
         BackColor       =   &H80000000&
         DataField       =   "CFROM017"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1680
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtEingetragenAm 
         BackColor       =   &H80000000&
         DataField       =   "CREATE017"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1680
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtBewID 
         DataField       =   "FK013017"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   4800
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdStelle 
         Caption         =   "Suchen"
         Height          =   315
         Left            =   3840
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAZ 
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox cmbAktenort 
         DataField       =   "Aktenort017"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblEingetragenVon 
         Caption         =   "Eingetragen von"
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblEingetragenAm 
         Caption         =   "Eingetragen am"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblAZ 
         Caption         =   "Aktenzeichen"
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblAktenort 
         Caption         =   "Aktenort"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame FrameInfo 
      Caption         =   "Datensatz Informationen"
      Height          =   2415
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   6375
      Begin VB.TextBox txtID 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "ID017"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   5295
      End
      Begin VB.TextBox txtCreateFrom 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "CFROM017"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   720
         Width           =   5295
      End
      Begin VB.TextBox txtCreate 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "CREATE017"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Width           =   5295
      End
      Begin VB.TextBox txtModifyFrom 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "MFROM017"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1440
         Width           =   5295
      End
      Begin VB.TextBox txtModify 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "MODIFY017"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1800
         Width           =   5295
      End
      Begin VB.Label lblID 
         Caption         =   "Datensatz ID"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblCreateFrom 
         Caption         =   "erstellt von"
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblCreate 
         Caption         =   "erstellt am"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblModifyFrom 
         Caption         =   "geändert von"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblModify 
         Caption         =   "geändert am"
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   975
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3135
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5530
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Aktenort"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Info"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Übernehmen"
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   3720
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   3720
      Visible         =   0   'False
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
End
Attribute VB_Name = "FrmEditAktenOrt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MODULNAME = "frmEditAktenOrt"                         ' Modulname für Fehlerbehandlung

Private bInit As Boolean                                            ' Wird True gesetzt wenn Alle werte geladen
Private bDirty As Boolean                                           ' Wird True gesetzt wenn Daten verändert wurden
Private bNew  As Boolean                                            ' Wird gesetzt wenn neuer DS sonst Update
Private bModal As Boolean                                           ' Ist Modal Geöffnet
Private szID As String                                              ' DS ID
Private ThisDBcon As Object                                         ' Aktuelle DB Verbindung
Private frmParent As Form                                           ' Aufrufendes DB form
Private szIDField As String
Private ThisFramePos As FramePos                                    ' Standart Frame Position

Private szSQL As String                                             ' SQL für a_personen
Private szWhere As String                                           ' Where Klausel
Private szIniFilePath As String                                     ' Pfad der Ini datei
Private lngImage As Integer                                         ' Imagiendex

Private szRootkey As String                                         ' = Aktenort
Private szDetailKey As String                                       ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)

Private CurrentStep As String                                       ' Aktueller WorkflowSchritt
Private OldCmbValue As String                                        'Alter Combo wert

Private Pers_ID As String
Private Bew_ID As String

Private Type FramePos                                               ' Positions Datentyp
    Top As Single                                                   ' Top position (oben)
    Left As Single                                                  ' Left Position (Links)
    Height As Single                                                ' Height (Höhe)
    Width As Single                                                 ' Width (Breite)
End Type

Private Sub Form_Activate()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If bInit Or bDirty Then Exit Sub                                ' Nicht bei initialisierung
    Call RefreshEditForm                                            ' Form daten aktualisieren
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub Form_Load()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call EditFormLoad(Me, szRootkey)                                ' Allg. Formload Aufrufen
    Call InitEditButtonMenue(Me, True, False, False)                ' Buttonleiste initialisieren
    With ThisFramePos
        Call GetTabStrimClientPos(TabStrip1, .Top, .Left, _
                .Height, .Width)                                    ' Frame Positionen aus TabStrip ermitteln
    End With
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If bDirty Then szID = ""                                        ' Id löschen wenn Dirty
    If bModal Then                                                  ' Modal ?
        Me.Hide                                                     ' Dann ausblenden
    Else                                                            ' Sonst
        Call EditFormUnload(Me)                                     ' Entladen
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Public Function InitEditForm(parentform As Form, dbCon As Object, DetailKey As String, Optional bDialog As Boolean)

'    Dim i As Integer                                                ' counter
    Dim tmpArray() As String
        
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    Set frmParent = parentform                                      ' Aufrufendes Form Übergeben
    bInit = True                                                    ' Wir initialisieren das Form-> andere vorgänge nicht ausführen
    Set ThisDBcon = dbCon                                           ' Aktuelle DB Verbindung übernehmen

    szRootkey = "Aktenort"                                          ' für Caption
    szIDField = "ID017"
    If InStr(DetailKey, ";") Then
        tmpArray = Split(DetailKey, ";")
    On Error Resume Next
        szDetailKey = tmpArray(0)
        Bew_ID = tmpArray(1)
        Pers_ID = tmpArray(2)
        Err.Clear
    Else
        szDetailKey = DetailKey ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)
    End If
    bModal = bDialog                                                ' Als Dialog anzeigen
    
    szIniFilePath = objObjectBag.Getappdir & objObjectBag.GetXMLFile ' XML inifile festlegen
    Call objTools.GetEditInfoFromXML(szIniFilePath, szRootkey, szSQL, szWhere, lngImage)
    
    Me.Icon = frmParent.ILTree.ListImages(lngImage).Picture         ' Form Icon Setzen
    
    If szDetailKey = "" Then bNew = True                            ' Neuer Datensatz
    If szDetailKey <> "" Then szWhere = szWhere & "CAST('" & szDetailKey & "' as uniqueidentifier)"
    
     ' Liste für Combo AktenOrt
    'Call FillCmbListWithSQL(cmbAktenort, "SELECT 'Fristenfach' As Aktenort UNION SELECT Nachname001 + ', ' + Vorname001 As Aktenort FROM User001", ThisDBcon)
    Call FillCmbListWithSQL(cmbAktenort, "SELECT VALUE015 As Aktenort FROM VALUES015 WHERE Fieldname015 = 'Aktenort017' UNION SELECT Nachname001 + ', ' + Vorname001 As Aktenort FROM User001", ThisDBcon)
    
    Call InitAdoDC(Me, ThisDBcon, szSQL, szWhere)                   ' ADODC Initialisieren
    Me.Refresh                                                      ' Form Aktualisieren
    
    If bNew Then                                                    ' Wenn DS Neu
        Adodc1.Recordset.AddNew                                     ' Neuen DS an RS anhängen
        txtID.Text = ThisDBcon.GetValueFromSQL("SELECT NewID()")    ' Neue ID (Guid) ermitteln
        txtEingetragenAm.Text = Format(CDate(Now()), "dd.mm.yyyy")
        txtEingetragenVon.Text = objObjectBag.GetUserName()
        Call GetDefaultValues(Me, szRootkey, szIniFilePath)
        
        Adodc1.Recordset.Fields("ID017").Value = txtID.Text
        Adodc1.Recordset.Fields("Create017").Value = txtEingetragenAm.Text
        Adodc1.Recordset.Fields("CFROM017").Value = txtEingetragenVon.Text
        szID = txtID
        
        txtIDPers = Pers_ID
        txtBewID = Bew_ID
        
        bDirty = True                                               ' Dirty da Neu
        'Me.Refresh
    Else
        szID = DetailKey
        txtEingetragenAm.Text = Format(CDate(txtCreate.Text), "dd.mm.yyyy")
        txtEingetragenVon.Text = txtCreateFrom.Text
    End If
    
    Call GetLockedControls(Me)                                      ' Gelockte controls finden
    Call HiglightThisMustFields(Not (bNew Or bDirty))               ' IndexFelder hervorheben
    Call RefreshRelFields                                           ' Relations felder Aktualisieren
    Call RefreshFrameAktenort(True)                                 ' Frame aktenort Aktualisieren
    Call InitFrameInfo(Me)                                          ' Info Frame initialisieren
    Call SetEditFormCaption(Me, szRootkey, "")                      ' Form Caption setzen
    Call CheckUpdate(Me)                                            ' Evtl Übernehmen disablen
    
exithandler:
    bInit = False                                                   ' Initialisierung dieses Forms beendet
    'If bNew Then bDirty = True
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

Public Sub HiglightThisMustFields(Optional bDeHiglight As Boolean)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call HiglightMustFields(Me, bDeHiglight)                        ' Alle PK ind IndexFields entfärben
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub RefreshFrameAktenort(Optional bVisible As Boolean)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call PosFrameAndListView(Me, FrameAktenort, True)               ' Frame Aktenort Positionieren
    FrameAktenort.Visible = bVisible                                ' Sichtbar ?
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Function RefreshRelFields()
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    Call RefreshRelField(Me, txtPerson, txtIDPers, _
            "SELECT TOP 1 NACHNAME010 + ', ' + VORNAME010 FROM RA010", _
            "ID010 =", False)
    
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "RefreshRelFields", errNr, errDesc)
    Resume exithandler
End Function

Private Sub ShowKontextMenu(Menuename As String)
    ' Zeigt das Menü mit MenueName als Kontext (Popup) Menü an
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren

    Select Case Menuename
    Case ""
'        PopupMenu kmnuLVFortbildungen
    Case Else
    
    End Select
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub HandleMenueKlick(szMenueName As String, Optional szCaption As String)
     
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    Dim BewID As String
    
    If HandleLVkmnuNew(Me, szCaption) Then GoTo exithandler
    
    Select Case szMenueName
    Case ""
'        'Call SetRelationinLV(Me, "Fortbildungen", "Thema", _
'                ThisDBCon, LVFortbildungen, rsFortbildungen, _
'                "FK010014", "FK011014")
'        Call RefereshFortbildungen(True)
        
    Case Else
    
    End Select
    
exithandler:
On Error Resume Next

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "HandleMenueKlick", errNr, errDesc)
End Sub
    
Private Sub HandleKeyDown(frmEdit As Form, KeyCode As Integer, Shift As Integer)
' Behandelt KeyDownEvents im Edit Form
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call HandleKeyDownEdit(Me, KeyCode, Shift)                      ' Spezielle KeyDon Events dieses Forms
    Call frmParent.HandleGlobalKeyCodes(KeyCode, Shift)             ' Key Down Events der Anwendung
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub HandleTabClick(TS As TabStrip)
' Behandelt Tab Klicks

On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    Call HandleTabClickNew(Me, TS)                                  ' Wenn bNew dan nur 1. Tab zulassen
    
    If TS.SelectedItem = "Info" Then
        FrameAktenort.Visible = False
        FrameInfo.Visible = True
        Exit Sub
    End If
    
    Select Case TS.SelectedItem.Index
    Case 1
        FrameAktenort.Visible = True
        FrameInfo.Visible = False
    Case 2
        FrameAktenort.Visible = True
        FrameInfo.Visible = True
    Case Else
    
    End Select

exithandler:
On Error Resume Next
    Me.Refresh
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "HandleTabClick", errNr, errDesc)
End Sub

Private Function ValidateEditForm() As Boolean

    Dim szMSG As String                                             ' MessageText
    Dim szTitle As String                                           ' Message Titel
    Dim FocusCTL As Control                                         ' Control das den Focus erhält
    Dim bValidationFaild As Boolean                                 ' Validation nicht erfolgreich
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    szTitle = "Unvollständige Daten"                                ' Meldungstitel setzen
    
    bValidationFaild = ValidateTxtFieldOnEmpty(cmbAktenort, "Aktenort", _
            szMSG, FocusCTL)                                        ' Aktenort auf Leer prüfen
    
'    bValidationFaild = ValidateTxtFieldOnEmpty(cmbAktenort, "Aktenort", _
'            szMsg, FocusCTL)                                        ' Aktenort auf Leer prüfen
            
    If bValidationFaild Then                                        ' Wenn Validierung Gescheitert
        Call objError.ShowErrMsg(szMSG, vbInformation, szTitle)     ' Hinweis meldung anzeigen
        FocusCTL.SetFocus                                           ' Fokus setzen
        ValidateEditForm = False                                    ' Ruckgabewert setzen
    Else
        ValidateEditForm = True
    End If
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ValidateEditForm", errNr, errDesc)
    Resume exithandler
End Function

Private Function SaveEditForm() As Boolean
' Speichert den Datensatz nach Validierung der eingaben
    Dim bNewBeforSave As Boolean                                    ' DS vom speichern Neu
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    bNewBeforSave = bNew                                            ' DS vom speichern Neu
    If Not ValidateEditForm Then GoTo exithandler                   ' Eingaben Validieren
    If UpdateEditForm(Me, szRootkey) Then                           ' Speichern
        bNew = False                                                ' DS nicht mehr neu
        Call HiglightThisMustFields(True)                           ' Hervorhebung Pflichfelder abschalten
    End If
    
    SaveEditForm = True                                             ' erfolg zurück
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SaveEditForm", errNr, errDesc)
    Resume exithandler
End Function

Private Function RefreshEditForm()
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    Call Me.Adodc1.Refresh                                          ' AdoDC aktualisieren
    Me.Refresh                                                      ' Form aktualisieren
    Call RefreshRelFields                                           ' Relations Felder Aktualisieren
    Call RefreshFrameAktenort(Me.FrameAktenort.Visible)             ' Frame aktenort Refreshen
    
exithandler:
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "RefreshEditForm", errNr, errDesc)
    Resume exithandler
End Function
                                                                    ' *****************************************
                                                                    ' TabSrip Events
Private Sub TabStrip1_Click()
    Call HandleTabClick(TabStrip1)                                  ' Tab Klick behandeln
End Sub
                                                                    ' *****************************************
                                                                    ' Button Events
Private Sub cmdESC_Click()
On Error Resume Next                                                ' Fehlerbehandlung deaktivierenn
    Unload Me
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub cmdOK_Click()

On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    If SaveEditForm Then                                            ' Dieses Form Speichern
        Call CheckUpdate(Me)                                        ' Evtl Übernehmen disablen
        Unload Me
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
    Call objError.Errorhandler(MODULNAME, "cmdOK_Click", errNr, errDesc)
    Resume exithandler
End Sub

Public Sub cmdUpdate_Click()

On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    If SaveEditForm Then                                            ' Dieses Form Speichern
        Call CheckUpdate(Me)                                        ' Evtl Übernehmen disablen
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
    Call objError.Errorhandler(MODULNAME, "cmdUpdate_Click", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub cmdDelete_Click()

On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
        
    Call DeleteDS(szRootkey, ID)                                    ' diesen Datensatz Löschen
    Unload Me                                                       ' Dieses form schliessen
    
exithandler:
On Error Resume Next
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "cmdDelete_Click", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub cmdSave_Click()

On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    If SaveEditForm Then                                            ' Dieses Form Speichern
        Call CheckUpdate(Me)                                        ' Evtl Übernehmen disablen
    End If
    
exithandler:
On Error Resume Next
    Me.MousePointer = vbDefault                                     ' Mauszeiger wieder normal
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "cmdSave_Click", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub cmdWord_Click()

'    Dim StellenID As String                                         ' Stellen ID
'
'On Error Resume Next                                                ' Erstmal ohne Fehlerbehandlung
'
'    StellenID = rsBewerbungen.Fields("FK012013").Value              ' evtl. Stellen ID Ermitteln
'    Err.Clear
'
'On Error GoTo Errorhandler
'
'    Call WriteWord("", ID, StellenID)                               ' SAT aufrufen
'    Call RefereshDokumente(True)                                    ' LV Dokumente Aktualisieren
        
exithandler:
On Error Resume Next
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "cmdWord_Click", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub cmdPersonrSuchen_Click()
 Dim NewID As String
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    NewID = ShowSearch(ThisDBcon, "Personen", "Nachname")
    If NewID <> txtIDPers.Text And NewID <> "" Then
        txtIDPers = NewID
        Adodc1.Recordset.Fields(txtIDPers.DataField).Value = NewID
        Call RefreshRelFields
        bDirty = True
        Call CheckUpdate(Me)
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Sub
                                                                    ' *****************************************
                                                                    ' Mouse Events
                                                                    



                                                                    ' *****************************************
                                                                    ' Menue Events
                                                                    
                                                                    ' *****************************************
                                                                    ' Change Events
Private Sub txtPerson_Change()
    If Not bInit Then Call StandartTextChange(Me, txtPerson)
End Sub

Private Sub txtIDPers_Change()
    If Not bInit Then Call StandartTextChange(Me, txtIDPers)
End Sub

Private Sub txtBewID_Change()
    If Not bInit Then Call StandartTextChange(Me, txtBewID)
End Sub

Private Sub txtAZ_Change()
    If Not bInit Then Call StandartTextChange(Me, txtAZ)
End Sub

Private Sub cmbAktenort_DropDown()                                  ' Liste ausklappen
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub
    OldCmbValue = cmbAktenort.Text                                  ' Auswahl beginnt
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub cmbAktenort_Click()                                     ' Änderung duch Liste auswahl
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                          ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> cmbAktenort Then bDirty = True                ' Dirty Nur wenn combo <> oldValue
    Call CheckUpdate(Me)
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub cmbAktenort_Change()                                    ' Änderung duch Texteingabe
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                          ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> cmbAktenort Then bDirty = True                ' Dirty Nur wenn combo <> oldValue
    Call CheckUpdate(Me)
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub cmbAktenort_Validate(Cancel As Boolean)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                          ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> cmbAktenort Then bDirty = True                ' Dirty Nur wenn combo <> oldValue
    Call CheckUpdate(Me)
    OldCmbValue = ""                                                ' Auswahl beendet
    Call FirstCharUp(Me, Me.cmbAktenort)                            ' 1. Zeichen Groß schreiben
    Err.Clear                                                       ' Evtl. Error clearen
End Sub
                                                                    ' *****************************************
                                                                    ' Key Down Events
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)                          ' Standart Keydown Handeln
End Sub

Private Sub dtEingang_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)                          ' Standart Keydown Handeln
End Sub

Private Sub txtBewerber_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)                          ' Standart Keydown Handeln
End Sub

Private Sub txtStelle_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)                          ' Standart Keydown Handeln
End Sub
'Private Sub txtPLZ_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call HandleKeyDown(Me, Me.ActiveControl, KeyCode, Shift)
'End Sub
                                                                    ' *****************************************
                                                                    ' Fokus Events
Private Sub txtAZ_GotFocus()
    Call HiglightCurentField(Me, txtAZ, False)                      ' Hervorhebung Aktiv anschalten
End Sub

Private Sub txtAZ_LostFocus()
    Call HiglightCurentField(Me, txtAZ, True)                       ' Hervorhebung Aktiv abschalten
End Sub

Private Sub txtEingetragenAm_GotFocus()
    Call HiglightCurentField(Me, txtEingetragenAm, False)           ' Hervorhebung Aktiv anschalten
End Sub

Private Sub txtEingetragenAm_LostFocus()
    Call HiglightCurentField(Me, txtEingetragenAm, True)            ' Hervorhebung Aktiv abschalten
End Sub

Private Sub txtEingetragenVon_GotFocus()
    Call HiglightCurentField(Me, txtEingetragenVon, False)          ' Hervorhebung Aktiv anschalten
End Sub

Private Sub txtEingetragenVon_LostFocus()
    Call HiglightCurentField(Me, txtEingetragenVon, True)           ' Hervorhebung Aktiv abschalten
End Sub

Private Sub cmbAktenort_GotFocus()
    Call HiglightCurentField(Me, cmbAktenort, False)                ' Hervorhebung Aktiv anschalten
End Sub

Private Sub cmbAktenort_LostFocus()
    Call HiglightCurentField(Me, cmbAktenort, True)                 ' Hervorhebung Aktiv abschalten
End Sub
                                                                    ' *****************************************
                                                                    ' Properties
Public Property Get IsNew() As Boolean
    IsNew = bNew
End Property

Public Property Get IDField() As String
    IDField = szIDField
End Property

Public Property Get ID() As String
    ID = szID
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = bDirty
End Property

Public Property Let SetDirty(Dirty As Boolean)
    bDirty = Dirty
End Property

Public Property Get GetDBConn() As Object
    Set GetDBConn = ThisDBcon
End Property

Public Property Get GetXMLPath() As String
    GetXMLPath = szIniFilePath
End Property

Public Property Get GetCurrentStep() As String
    GetCurrentStep = CurrentStep
End Property

Public Property Get GetRootkey() As String
    GetRootkey = szRootkey
End Property

Public Property Get GetFrameTop() As Single
    GetFrameTop = ThisFramePos.Top                                  ' Gibt die Top Pos. der Standartframes zurück
End Property

Public Property Get GetFrameLeft() As Single
    GetFrameLeft = ThisFramePos.Left                                ' Gibt die Left Pos. der Standartframes zurück
End Property

Public Property Get GetFrameHeigth() As Single
    GetFrameHeigth = ThisFramePos.Height                            ' Gibt die Height der Standartframes zurück
End Property

Public Property Get GetFrameWidth() As Single
    GetFrameWidth = ThisFramePos.Width                              ' Gibt die Width der Standartframes zurück
End Property

