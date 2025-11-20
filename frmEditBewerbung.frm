VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEditBewerbung 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Berwerbungen"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6495
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   0
      Picture         =   "frmEditBewerbung.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Datensatz speichern"
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   375
      Left            =   480
      Picture         =   "frmEditBewerbung.frx":058A
      Style           =   1  'Grafisch
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Datensatz löschen"
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton cmdWord 
      Height          =   375
      Left            =   960
      Picture         =   "frmEditBewerbung.frx":0914
      Style           =   1  'Grafisch
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Neues Anschreiben"
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox txtEingang 
      DataField       =   "EINGANG013"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd.MM.yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1031
         SubFormatType   =   3
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame FrameBewerbung 
      Caption         =   "Bewerbungsdaten"
      Height          =   1935
      Left            =   120
      TabIndex        =   25
      Top             =   840
      Width           =   6255
      Begin VB.CommandButton cmdNewStelle 
         Caption         =   "Neu"
         Height          =   315
         Left            =   5400
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdNewBewerber 
         Caption         =   "Neu"
         Height          =   315
         Left            =   5400
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtBewerber 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox txtStelle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   3615
      End
      Begin VB.CommandButton cmdBewSuchen 
         Height          =   315
         Left            =   5040
         Picture         =   "frmEditBewerbung.frx":0C9E
         Style           =   1  'Grafisch
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdStelleSuchen 
         Height          =   315
         Left            =   5040
         Picture         =   "frmEditBewerbung.frx":1228
         Style           =   1  'Grafisch
         TabIndex        =   7
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtBem 
         DataField       =   "BEM013"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   960
         Width           =   4815
      End
      Begin VB.TextBox txtIDBewerber 
         DataField       =   "FK010013"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   4680
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtIDStelle 
         DataField       =   "FK012013"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   4800
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblMsg 
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   5055
      End
      Begin VB.Label lblBewerber 
         Caption         =   "Bewerber"
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblStelle 
         Caption         =   "Stelle"
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblBemerkung 
         Caption         =   "Bemerkung"
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame FrameInfo 
      Caption         =   "Datensatz Informationen"
      Height          =   2175
      Left            =   480
      TabIndex        =   13
      Top             =   840
      Width           =   5775
      Begin VB.TextBox txtID 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "ID013"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   4000
      End
      Begin VB.TextBox txtCreateFrom 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "CFROM013"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   720
         Width           =   4000
      End
      Begin VB.TextBox txtCreate 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "CREATE013"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1080
         Width           =   4000
      End
      Begin VB.TextBox txtModifyFrom 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "MFROM013"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1440
         Width           =   4000
      End
      Begin VB.TextBox txtModify 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "MODIFY013"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1800
         Width           =   4000
      End
      Begin VB.Label lblID 
         Caption         =   "Datensatz ID"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblCreateFrom 
         Caption         =   "erstellt von"
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblCreate 
         Caption         =   "erstellt am"
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblModifyFrom 
         Caption         =   "geändert von"
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblModify 
         Caption         =   "geändert am"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   975
      End
   End
   Begin MSComCtl2.DTPicker dtEingang 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd.MM.yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1031
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   57409537
      CurrentDate     =   39260
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   3240
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
      Left            =   5280
      TabIndex        =   12
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2655
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4683
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bewerbung"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Info"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblEingang 
      Caption         =   "Datum d. Bewerbung"
      Height          =   315
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmEditBewerbung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MODULNAME = "frmBewerbung"                            ' Modulname für Fehlerbehandlung

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

Private szRootkey As String                                         ' = Bewerbung
Private szDetailKey As String                                       ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)

Private CurrentStep As String                                       ' Aktueller WorkflowSchritt
Private OldCmbValue As String                                        'Alter Combo wert

Private StellenID As String
Private PersID As String

Private rsBewerbung As ADODB.Recordset
Private rsPerson As ADODB.Recordset
Private rsStelle As ADODB.Recordset

Private Type FramePos                                               ' Positions Datentyp
    Top As Single                                                   ' Top position (oben)
    Left As Single                                                  ' Left Position (Links)
    Height As Single                                                ' Height (Höhe)
    Width As Single                                                 ' Width (Breite)
End Type

'Private szSQLBew As String
'Private szSQLStelle As String
'Private szSQLPers As String

Private Sub Form_Activate()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If bInit Or bDirty Then Exit Sub                                ' Nicht bei initialisierung
    Call RefreshEditForm                                            ' Form daten aktualisieren
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub Form_Load()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call EditFormLoad(Me, szRootkey)                                ' Allg. Formload Aufrufen
    Call InitEditButtonMenue(Me, True, True, False)                 ' Buttonleiste initialisieren
    With ThisFramePos
        Call GetTabStrimClientPos(TabStrip1, .Top, .Left, _
                .Height, .Width)                                    ' Frame Positionen aus TabStrip ermitteln
    End With
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next                                                ' Fehlerbehandlung deaktiviern
    rsBewerbung.Close                                               ' RS Bewerbung schliessen
    rsPerson.Close                                                  ' RS Person schliessen
    rsStelle.Close                                                  ' RS Stellen schliessen
    If bDirty Then szID = ""                                        ' Wenn ungespeichert ID Leeren
    If bModal Then                                                  ' Wenn Modal
        Me.Hide                                                     ' dann ausblenden
    Else                                                            ' Sonst
        Call EditFormUnload(Me)                                     ' AUs Edit Form Array entfernen
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Public Function InitEditForm(parentform As Form, dbCon As Object, DetailKey As String, Optional bDialog As Boolean)

'    Dim i As Integer                                                ' counter
    Dim tmpArray() As String                                        ' Array mit Zerlegtem Detailkey
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
        
    Set frmParent = parentform
    bInit = True                                                    ' Wir initialisieren das Form-> andere vorgänge nicht ausführen
    Set ThisDBcon = dbCon                                           ' Aktuelle DB Verbindung übernehmen
    szRootkey = "Bewerbungen"                                       ' für Caption
    szIDField = "ID013"
                                                                    ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)
    If InStr(DetailKey, ";") Then                                   ' DetailKey evtl. aufspalten
        tmpArray = Split(DetailKey, ";")
    On Error Resume Next
        szDetailKey = tmpArray(0)
        StellenID = tmpArray(1)
        PersID = tmpArray(2)
        Err.Clear
    Else
        szDetailKey = DetailKey
    End If
    bModal = bDialog                                                ' Als Dialog anzeigen
    
    szIniFilePath = objObjectBag.Getappdir & objObjectBag.GetXMLFile ' XML inifile festlegen
    Call objTools.GetEditInfoFromXML(szIniFilePath, szRootkey, szSQL, szWhere, lngImage)
    
    Me.Icon = frmParent.ILTree.ListImages(lngImage).Picture         ' Form Icon Setzen
    
    If szDetailKey = "" Then bNew = True                            ' Neuer Datensatz
    If szDetailKey <> "" Then szWhere = szWhere & "CAST('" & szDetailKey & "' as uniqueidentifier)"
    
    Call InitAdoDC(Me, ThisDBcon, szSQL, szWhere)                   ' ADODC Initialisieren
    Me.Refresh
    
    If bNew Then                                                    ' Wenn DS Neu
        Adodc1.Recordset.AddNew                                     ' Neuen DS an RS anhängen
        txtID.Text = ThisDBcon.GetValueFromSQL("SELECT NewID()")    ' Neue ID (Guid) ermitteln
        szID = txtID
        If StellenID <> "" Then
            txtIDStelle = StellenID
        Else
            StellenID = GetIDFromNode()
            txtIDStelle = StellenID
        End If
        
        txtIDStelle.Text = StellenID
        txtIDBewerber.Text = PersID
        Call FormatDTPicker(Me, dtEingang, Now())
        Call dtEingang_Change
        bDirty = True                                               ' Dirty da Neu
    Else
        'szID = szDetailKey
        szID = Me.txtID
    End If
       
    Call GetLockedControls(Me)                                      ' Gelockte controls finden
    Call HiglightThisMustFields(Not (bNew Or bDirty))               ' IndexFelder hervorheben
    Call InitFrameBewerbungsDaten(True)                             ' Frame Bewerbungdaten Initialisieren
    Call InitFrameInfo(Me)                                          ' Info Frame initialisieren
    Call RefreshRelFields                                           ' zu den FK IDs relevate Werte holen
    Call SetEditFormCaption(Me, szRootkey, txtBewerber)             ' Formular Caption abhängig von den Daten setzen
    Call CheckUpdate(Me)                                            ' Evtl Übernehmen disablen
    
    If txtIDBewerber.Text = "" Then
        txtBewerber.Locked = False
    Else
        txtBewerber.Locked = True
    End If
    
    If txtIDStelle.Text = "" Then
        txtStelle.Locked = False
    Else
        txtStelle.Locked = True
    End If
    
exithandler:
    bInit = False                                                   ' Initialisierung dieses Forms beendet
    
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
    Call HiglightMustField(Me, txtBewerber, bDeHiglight)
    Call HiglightMustField(Me, txtStelle, bDeHiglight)
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub InitFrameBewerbungsDaten(Optional bVisible As Boolean)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call PosFrameAndListView(Me, FrameBewerbung, True)              ' Frame Positionieren
    FrameBewerbung.Visible = bVisible                               ' Sichtbar ?
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Function RefreshRelFields()

On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    Call RefreshRelField(Me, txtStelle, txtIDStelle, _
            "SELECT TOP 1 BEZIRK012 + ' ' + CONVERT(varchar(20),FRIST012,104) FROM STELLEN012", _
            "ID012 =", True)
    Call RefreshRelField(Me, txtBewerber, txtIDBewerber, _
            "SELECT TOP 1 NACHNAME010 + ', ' + ISNULL(VORNAME010,'') FROM RA010", _
            "ID010 =", True)
            
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

Private Sub HandleKeyDown(frmEdit As Form, KeyCode As Integer, Shift As Integer)
' Behandelt KeyDownEvents im Edit Form
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call HandleKeyDownEdit(Me, KeyCode, Shift)                      ' Spezielle KeyDon Events dieses Forms
    Call frmParent.HandleGlobalKeyCodes(KeyCode, Shift)             ' Key Down Events der Anwendung
    Err.Clear                                                       ' Evtl. error clearen
End Sub

Private Sub HandleTabClick(TS As TabStrip)
' Behandelt Tab Klicks

On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    Call HandleTabClickNew(Me, TS)                                  ' Wenn bNew dan nur 1. Tab zulassen
    
    If TS.SelectedItem = "Info" Then
        FrameBewerbung.Visible = False
        FrameInfo.Visible = True
        Exit Sub
    End If
    
    Select Case TS.SelectedItem.Index
    Case 1
        FrameBewerbung.Visible = True
        FrameInfo.Visible = False
    Case 2
        FrameBewerbung.Visible = True
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
    
    bValidationFaild = ValidateTxtFieldOnEmpty(txtBewerber, "Bewerber", _
            szMSG, FocusCTL)                                        ' txtBewerber auf Leer prüfen

    bValidationFaild = ValidateTxtFieldOnEmpty(txtStelle, "Stelle", _
            szMSG, FocusCTL)                                        ' txtStelle auf Leer prüfen
    
    bValidationFaild = ValidateTxtFieldOnEmpty(txtEingang, "Datum d. Bewerbung", _
            szMSG, FocusCTL)                                        ' txtEingang auf Leer prüfen
    
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
    
    bNewBeforSave = bNew                                            ' DS war neu vom speichern
    If Not ValidateEditForm Then GoTo exithandler                   ' Eingaben Validieren
    If UpdateEditForm(Me, szRootkey) Then                           ' Speichern
        bNew = False                                                ' nicht mehr neu
        Call HiglightThisMustFields(True)                           ' Hervorhebung von Indexfeldern abschalten
    End If
    SaveEditForm = True                                             ' Erfolg zurück
    
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
    Call RefreshRelFields                                           ' zu den FK IDs relevate Werte holen
    
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
                                                                    ' Botton Events
Private Sub cmdESC_Click()
On Error Resume Next                                                ' fehlerbehandlung deaktivieren
    Unload Me                                                       ' Form entladen
    Err.Clear                                                       ' Evtl. error clearen
End Sub

Private Sub cmdOK_Click()

On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    If SaveEditForm Then                                            ' Dieses Form Speichern
        Call CheckUpdate(Me)                                        ' Evtl Übernehmen disablen
        Unload Me
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
    Me.MousePointer = vbDefault                                     ' Mauszeiger wieder normal
    
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

Private Sub cmdNewBewerber_Click()
    txtIDBewerber.Text = frmMain.OpenEditForm("Bewerber", "", Me, True) ' Leeres form Bewerber zum eingeben öffnen
    If txtIDBewerber <> "" Then Call RefreshRelFields                   ' evtl. Rükgabewert eintragen
End Sub

Private Sub cmdNewStelle_Click()
    txtIDStelle.Text = frmMain.OpenEditForm("Stellen", "", Me, True)    ' Leeres form Bewerber zum eingeben öffnen
    If txtIDStelle <> "" Then Call RefreshRelFields                     ' evtl. Rükgabewert eintragen
End Sub

Private Sub cmdBewSuchen_Click()
    Dim NewID As String
    Dim szSuchtext As String
    
    If txtIDBewerber.Text = "" Then szSuchtext = txtBewerber
    NewID = ShowSearch(ThisDBcon, "Bewerber", "Nachname", szSuchtext)
    If NewID <> txtIDBewerber.Text And NewID <> "" Then
        txtIDBewerber = NewID
        Adodc1.Recordset.Fields(txtIDBewerber.DataField).Value = NewID
        Call RefreshRelFields
        bDirty = True
        Call CheckUpdate(Me)
    End If
End Sub

Private Sub cmdStelleSuchen_Click()
    Dim NewID As String
    Dim szSuchtext As String
    
    If txtIDStelle.Text = "" Then szSuchtext = txtStelle
    NewID = ShowSearch(ThisDBcon, "Stellen", "Bezirk")
    If NewID <> txtIDStelle.Text And NewID <> "" Then
        txtIDStelle.Text = NewID
        Adodc1.Recordset.Fields(txtIDStelle.DataField).Value = NewID
        Call RefreshRelFields
        bDirty = True
        Call CheckUpdate(Me)
    End If
End Sub
                                                                    ' *****************************************                                                                                                                                        ' *****************************************
                                                                    ' Mouse Events
Private Sub txtBewerber_DblClick()
    Dim RootKey As String
    Dim DetailKey As String
    RootKey = "Personen"
    DetailKey = txtIDBewerber.Text
    If RootKey <> "" And DetailKey <> "" Then
        Call frmMain.OpenEditForm(RootKey, DetailKey, frmParent)
    End If
End Sub

Private Sub txtStelle_DblClick()
    Dim RootKey As String
    Dim DetailKey As String
    RootKey = "Ausgeschriebene Stellen"
    DetailKey = txtIDStelle.Text
    If RootKey <> "" And DetailKey <> "" Then
        Call frmMain.OpenEditForm(RootKey, DetailKey, frmParent)
    End If
End Sub
                                                                    ' *****************************************
                                                                    ' Menue Events
                                                                    
                                                                    ' *****************************************
                                                                    ' Change Events
Private Sub txtEingang_Change()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If Len(txtEingang.Text) < 8 Then Exit Sub                       ' Text zur kurz -> nicht weitermachen
    If IsDate(txtEingang.Text) Then                                 ' Ist Text Datum?
        dtEingang.Value = txtEingang.Text                           ' Datum an DT übergeben
        If bInit Then Exit Sub                                      ' Wenn Form initialisiert -> fertig
        bDirty = True                                               ' DS ist ungespeichert
        Call CheckUpdate(Me)                                        ' Evtl. Buttons dis/enablen
    End If
    Err.Clear                                                       ' Evtl. error clearen
End Sub

Private Sub dtEingang_Change()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    txtEingang.Text = dtEingang.Value                               ' Datum an txt Feld übergeben
    If bInit Then Exit Sub                                          ' Wenn Form initialisiert -> fertig
    Me.Adodc1.Recordset.Fields(txtEingang.DataField).Value _
            = Format(Me.dtEingang.Value, "dd.mm.yyyy")              ' Bei DTPicker Wert nochmal übernehmen da Sonst die Datenbindung nicht funzt
    Err.Clear                                                       ' Evtl. error clearen
End Sub

'Private Sub txtAz_Change()
'    If Not bInit Then Call StandartTextChange(Me, txtAz_Chang)
'End Sub

Private Sub txtBem_Change()
    If Not bInit Then Call StandartTextChange(Me, txtBem)
End Sub

Private Sub txtIDBewerber_Change()
    If Not bInit Then Call StandartTextChange(Me, txtIDBewerber)
End Sub

Private Sub txtIDStelle_Change()
    If Not bInit Then Call StandartTextChange(Me, txtIDStelle)
End Sub
                                                                    ' *****************************************
                                                                    ' Fokus Events
Private Sub txtBem_GotFocus()
    Call HiglightCurentField(Me, txtBem, False)                     ' Hervorhebung Aktiv anschalten
End Sub

Private Sub txtBem_LostFocus()
    Call HiglightCurentField(Me, txtBem, True)                      ' Hervorhebung Aktiv abschalten
End Sub

Private Sub txtStelle_GotFocus()
    Call HiglightCurentField(Me, txtStelle, False)                  ' Hervorhebung Aktiv anschalten
End Sub

Private Sub txtStelle_LostFocus()
    Call HiglightCurentField(Me, txtStelle, True)                   ' Hervorhebung Aktiv abschalten
End Sub

Private Sub txtBewerber_GotFocus()
    Call HiglightCurentField(Me, txtBewerber, False)                ' Hervorhebung Aktiv anschalten
End Sub

Private Sub txtBewerber_LostFocus()
    Call HiglightCurentField(Me, txtBewerber, True)                 ' Hervorhebung Aktiv abschalten
End Sub

Private Sub txtEingang_GotFocus()
    Call HiglightCurentField(Me, txtEingang, False)                 ' Hervorhebung Aktiv anschalten
End Sub

Private Sub txtEingang_LostFocus()
    Call HiglightCurentField(Me, txtEingang, True)                  ' Hervorhebung Aktiv abschalten
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
    If KeyCode = 13 Then Call cmdBewSuchen_Click                    ' ENTER -> Stelle Suchen
End Sub

Private Sub txtStelle_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)                          ' Standart Keydown Handeln
    If KeyCode = 13 Then Call cmdStelleSuchen_Click                 ' ENTER -> Stelle Suchen
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

