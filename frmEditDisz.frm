VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEditDisz 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6780
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame FrameInfo 
      Caption         =   "Datensatz Informationen"
      Height          =   2535
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   6375
      Begin VB.TextBox txtModify 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "MODIFY019"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1800
         Width           =   5295
      End
      Begin VB.TextBox txtModifyFrom 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "MFROM019"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1440
         Width           =   5295
      End
      Begin VB.TextBox txtCreate 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "CREATE019"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1080
         Width           =   5295
      End
      Begin VB.TextBox txtCreateFrom 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "CFROM019"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   720
         Width           =   5295
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "ID019"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label lblModify 
         Caption         =   "geändert am"
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblModifyFrom 
         Caption         =   "geändert von"
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblCreate 
         Caption         =   "erstellt am"
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblCreateFrom 
         Caption         =   "erstellt von"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblID 
         Caption         =   "Datensatz ID"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdWord 
      Height          =   375
      Left            =   960
      Picture         =   "frmEditDisz.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   34
      ToolTipText     =   "Neues Anschreiben"
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   375
      Left            =   480
      Picture         =   "frmEditDisz.frx":038A
      Style           =   1  'Grafisch
      TabIndex        =   33
      ToolTipText     =   "Datensatz löschen"
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   0
      Picture         =   "frmEditDisz.frx":0714
      Style           =   1  'Grafisch
      TabIndex        =   32
      ToolTipText     =   "Datensatz speichern"
      Top             =   5280
      Width           =   375
   End
   Begin VB.Frame FrameDisz 
      Caption         =   "Disziplinarmaßnahmen"
      Height          =   4215
      Left            =   1080
      TabIndex        =   25
      Top             =   840
      Width           =   6495
      Begin VB.TextBox txtDatum 
         DataField       =   "DATUM019"
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
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtErgebniss 
         DataField       =   "ERGEBNISS019"
         DataSource      =   "Adodc1"
         Height          =   630
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   3480
         Width           =   6255
      End
      Begin VB.TextBox txtEUmstaende 
         DataField       =   "EUMSTAENDE019"
         DataSource      =   "Adodc1"
         Height          =   630
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   2520
         Width           =   6255
      End
      Begin VB.TextBox txtMUmstaende 
         DataField       =   "MUMSTAENDE019"
         DataSource      =   "Adodc1"
         Height          =   630
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1560
         Width           =   6255
      End
      Begin VB.TextBox txtVerstoss 
         DataField       =   "VERSTOSS019"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   960
         Width           =   5055
      End
      Begin VB.TextBox txtMassnahme 
         DataField       =   "MASSNAHME019"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   600
         Width           =   5055
      End
      Begin MSComCtl2.DTPicker dtDatum 
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
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57147393
         CurrentDate     =   39280
      End
      Begin VB.Label lblErgebniss 
         Caption         =   "Ergebnis vorhergehender Notarprüfung"
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   3240
         Width           =   3855
      End
      Begin VB.Label lblEUmstaende 
         Caption         =   "Erschwerende Umstände"
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label lblMUmstaende 
         Caption         =   "Mildernde Umstände"
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblVerstoss 
         Caption         =   "Verstoß"
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1000
      End
      Begin VB.Label lblMassnahme 
         Caption         =   "Maßnahme"
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1000
      End
      Begin VB.Label lblDatum 
         Caption         =   "Datum"
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1000
      End
   End
   Begin VB.TextBox txtPersID 
      DataField       =   "FK010019"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5400
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
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
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Übernehmen"
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   5280
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4695
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8281
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Disziplinarmaßnahmen"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Info"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   5280
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
   Begin VB.Label lblPerson 
      Caption         =   "Person"
      Height          =   315
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditDisz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MODULNAME = "frmEditDisz"                             ' Modulname für Fehlerbehandlung

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

Private szRootkey As String                                         ' = Disziplinarmaßnahmen
Private szDetailKey As String                                       ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)

Private CurrentStep As String                                       ' Aktueller WorkflowSchritt
Private OldCmbValue As String                                        'Alter Combo wert

Private Pers_ID As String

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
    Call InitEditButtonMenue(Me, True, True, False)                 ' Buttonleiste initialisieren
    With ThisFramePos
        Call GetTabStrimClientPos(TabStrip1, .Top, .Left, _
                .Height, .Width)                                    ' Frame Positionen aus TabStrip ermitteln
    End With
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next                                                ' Fehlerbehandlung deaktiviern
    If bDirty Then szID = ""                                        ' Wenn ungespeichert ID Leeren
    If bModal Then                                                  ' Wenn Modal
        Me.Hide                                                     ' dann ausblenden
    Else                                                            ' Sonst
        Call EditFormUnload(Me)                                     ' AUs Edit Form Array entfernen
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Sub


Public Function InitEditForm(parentform As Form, dbCon As Object, DetailKey As String, Optional bDialog As Boolean)

    Dim i As Integer                                                ' counter
    Dim tmpArray() As String
        
On Error GoTo Errorhandler
    
    Set frmParent = parentform
    bInit = True                                                    ' Wir initialisieren das Form-> andere vorgänge nicht ausführen
    Set ThisDBcon = dbCon                                           ' Aktuelle DB Verbindung übernehme
    szRootkey = "Disziplinarmaßnahmen"                              ' für Caption
    szIDField = "ID019"
    
    If InStr(DetailKey, ";") Then
        tmpArray = Split(DetailKey, ";")
    On Error Resume Next
        szDetailKey = tmpArray(0)
        Pers_ID = tmpArray(1)
        Err.Clear
    Else
        szDetailKey = DetailKey                                     ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)
    End If
    bModal = bDialog                                                ' Als Dialog anzeigen
    
    szIniFilePath = objObjectBag.Getappdir & objObjectBag.GetXMLFile ' XML inifile festlegen
    Call objTools.GetEditInfoFromXML(szIniFilePath, szRootkey, szSQL, szWhere, lngImage)
    
    Me.Icon = frmParent.ILTree.ListImages(lngImage).Picture         ' Form Icon Setzen
    
    If szDetailKey = "" Then bNew = True                            ' Neue Datensatz
    If szDetailKey <> "" Then szWhere = szWhere & "CAST('" & szDetailKey & "' as uniqueidentifier)"
    
     ' Liste für Combo AktenOrt
    'Call FillCmbListWithSQL(cmbAktenort, "SELECT 'Fristenfach' As Aktenort UNION SELECT Nachname001 + ', ' + Vorname001 As Aktenort FROM User001", ThisDBCon)
    
    Call InitAdoDC(Me, ThisDBcon, szSQL, szWhere)                   ' ADODC Initialisieren
    Me.Refresh
    
    If bNew Then                                                    ' Wenn DS Neu
       
        Adodc1.Recordset.AddNew                                     ' Neuen DS an RS anhängen
         Call FormatDTPicker(Me, dtDatum, Now())
        txtDatum.Text = dtDatum.Value
        txtID.Text = ThisDBcon.GetValueFromSQL("SELECT NewID()")    ' Neue ID (Guid) ermitteln
        szID = txtID
        'Adodc1.Recordset.Fields("DATUM019").Value = txtDatum.Text
        bDirty = True                                               ' Dirty da Neu
    Else
        szID = txtID
        Pers_ID = txtPersID.Text
    End If
    txtPersID = Pers_ID
    
    Call GetLockedControls(Me)                                      ' Gelockte controls finden
    Call HiglightThisMustFields(Not (bNew Or bDirty))               ' IndexFelder hervorheben
    Call RefreshRelFields                                           ' Relation felder aktualisieren
    Call InitFrameDiszip
    Call InitFrameInfo(Me)                                          ' Infoframe initialisieren
    Call SetEditFormCaption(Me, szRootkey, "")                      ' Form Caption setzen
    Call CheckUpdate(Me)                                            ' Evtl Übernehmen disablen
    
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
On Error Resume Next
    Call HiglightMustFields(Me, bDeHiglight)                        ' Alle PK ind IndexFields entfärben
'    Call HiglightMustField(Me, txtBewerber, bDeHiglight)
'    Call NoHiglight(Me, txtStelle, bDeHiglight)

End Sub

Private Sub InitFrameDiszip()
    
On Error GoTo Errorhandler

    Call PosFrameAndListView(Me, FrameDisz, True)
    
exithandler:
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitFrameDiszip", errNr, errDesc)
    Resume exithandler
End Sub

Private Function RefreshRelFields()

'    Dim szSQLPers  As String
    
On Error GoTo Errorhandler

     Call RefreshRelField(Me, txtPerson, txtPersID, _
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
On Error Resume Next

    Select Case Menuename
    Case ""
'        PopupMenu kmnuLVFortbildungen
    Case Else
    
    End Select
End Sub

Private Sub HandleMenueKlick(szMenueName As String, Optional szCaption As String)
     
On Error GoTo Errorhandler
    
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
On Error Resume Next
    Call HandleKeyDownEdit(Me, KeyCode, Shift)                      ' Spezielle KeyDon Events dieses Forms
    Call frmParent.HandleGlobalKeyCodes(KeyCode, Shift)             ' Key Down Events der Anwendung
    Err.Clear
End Sub

Private Function ValidateEditForm() As Boolean

    Dim szMSG As String                                             ' MessageText
    Dim szTitle As String                                           ' Message Titel
    Dim FocusCTL As Control                                         ' Control das den Focus erhält
    Dim bValidationFaild As Boolean                                 ' Validation nicht erfolgreich
    
On Error GoTo Errorhandler
    
    szTitle = "Unvollständige Daten"                                ' Meldungstitel setzen
    
'    If Trim(txtEingang.Text) = "" Then
'        bValidationFaild = True                                     ' Validierung Gescheitert
'        szMSG = "Das Feld Eingang d. Bewerbung darf nicht leer sein!" ' Meldungstext setzen
'        Set FocusCTL = txtEingang                                   ' Focus Control setzen
'    End If
       
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
    
On Error GoTo Errorhandler
    
    bNewBeforSave = bNew
    
    If Not ValidateEditForm Then GoTo exithandler                   ' Eingaben Validieren
        
    If UpdateEditForm(Me, szRootkey) Then                           ' Speichern
        bNew = False                                                ' nicht mehr neu
        Call HiglightThisMustFields(True)                           ' Hervorhebung abschalten
    End If
    
    SaveEditForm = True
    
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
    Call RefreshRelFields                                           ' Relation felder aktualisieren
    
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

    Call HandleTabClickNew(Me, TabStrip1)                           ' Wenn bNew dan nur 1. Tab zulassen
    
    If TabStrip1.SelectedItem = "Info" Then
        FrameDisz.Visible = False
        FrameInfo.Visible = True
        Exit Sub
    End If
    
    Select Case TabStrip1.SelectedItem.Index
    Case 1
        FrameDisz.Visible = True
        FrameInfo.Visible = False
    Case 2
        FrameDisz.Visible = False
        FrameInfo.Visible = True
    Case Else
    
    End Select
    
    Me.Refresh
End Sub
                                                                    ' *****************************************
                                                                    ' Button Events
Private Sub cmdESC_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

On Error GoTo Errorhandler

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

On Error GoTo Errorhandler

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

On Error GoTo Errorhandler
        
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

    On Error GoTo Errorhandler

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
                                                                    ' *****************************************
                                                                    ' Mouse Events
                                                                    
                                                                    ' *****************************************
                                                                    ' Menue Events
                                                                    
                                                                    ' *****************************************
                                                                    ' Change Events
Private Sub txtDatum_Change()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If Len(txtDatum.Text) < 8 Then Exit Sub
    If IsDate(txtDatum.Text) Then
        dtDatum.Value = txtDatum.Text
        If bInit Then Exit Sub
        bDirty = True
        Call CheckUpdate(Me)
    End If
End Sub

Private Sub dtDatum_Change()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    txtDatum.Text = dtDatum.Value
    If bInit Then Exit Sub
    Me.Adodc1.Recordset.Fields(txtDatum.DataField).Value _
            = Format(Me.dtDatum.Value, "dd.mm.yyyy")                ' Bei DTPicker Wert nochmal übernehmen da Sonst die Datenbindung nicht funzt
End Sub

Private Sub txtErgebniss_Change()
    If Not bInit Then Call StandartTextChange(Me, txtErgebniss)
End Sub

Private Sub txtEUmstaende_Change()
    If Not bInit Then Call StandartTextChange(Me, txtEUmstaende)
End Sub

Private Sub txtMassnahme_Change()
    If Not bInit Then Call StandartTextChange(Me, txtMassnahme)
End Sub

Private Sub txtMUmstaende_Change()
    If Not bInit Then Call StandartTextChange(Me, txtMUmstaende)
End Sub

Private Sub txtVerstoss_Change()
    If Not bInit Then Call StandartTextChange(Me, txtVerstoss)
End Sub
                                                                    ' *****************************************
                                                                    ' Fokus Events

Private Sub txtErgebniss_GotFocus()
    Call HiglightCurentField(Me, txtErgebniss, False)
End Sub

Private Sub txtErgebniss_LostFocus()
    Call HiglightCurentField(Me, txtErgebniss, True)
End Sub

Private Sub txtDatum_GotFocus()
    Call HiglightCurentField(Me, txtDatum, False)
End Sub

Private Sub txtDatum_LostFocus()
    Call HiglightCurentField(Me, txtDatum, True)
End Sub

Private Sub txtEUmstaende_GotFocus()
    Call HiglightCurentField(Me, txtEUmstaende, False)
End Sub

Private Sub txtEUmstaende_LostFocus()
    Call HiglightCurentField(Me, txtEUmstaende, True)
End Sub

Private Sub txtMassnahme_GotFocus()
    Call HiglightCurentField(Me, txtMassnahme, False)
End Sub

Private Sub txtMassnahme_LostFocus()
    Call HiglightCurentField(Me, txtMassnahme, True)
End Sub

Private Sub txtMUmstaende_GotFocus()
    Call HiglightCurentField(Me, txtMUmstaende, False)
End Sub

Private Sub txtMUmstaende_LostFocus()
    Call HiglightCurentField(Me, txtMUmstaende, True)
End Sub

Private Sub txtPerson_GotFocus()
    Call HiglightCurentField(Me, txtPerson, False)
End Sub

Private Sub txtPerson_LostFocus()
    Call HiglightCurentField(Me, txtPerson, True)
End Sub

Private Sub txtVerstoss_GotFocus()
    Call HiglightCurentField(Me, txtVerstoss, False)
End Sub

Private Sub txtVerstoss_LostFocus()
    Call HiglightCurentField(Me, txtVerstoss, True)
End Sub


                                                                    ' *****************************************
                                                                    ' Key Down Events
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtDatum_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtEUmstaende_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtErgebniss_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtMassnahme_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtMUmstaende_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtVerstoss_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub dtDatum_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
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

