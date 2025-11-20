VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEditOLG 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Stammdaten"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6300
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame FrameGericht 
      Caption         =   "Gerichtsdaten"
      Height          =   2295
      Left            =   240
      TabIndex        =   26
      Top             =   840
      Width           =   5775
      Begin VB.ComboBox cmbBezirk 
         Height          =   315
         Left            =   2760
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtBezirkID 
         DataField       =   "BezirkID"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   4200
         TabIndex        =   34
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtBezirk 
         Height          =   315
         Left            =   1920
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Frame FrameAdresse 
         Caption         =   "Adresse"
         Height          =   1695
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   3975
         Begin VB.TextBox txtStr 
            DataField       =   "Str"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   960
            TabIndex        =   3
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox txtPLZ 
            DataField       =   "Plz"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   960
            TabIndex        =   4
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtOrt 
            DataField       =   "ort"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   1800
            TabIndex        =   5
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtTel 
            DataField       =   "tel"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   960
            TabIndex        =   6
            Top             =   960
            Width           =   2895
         End
         Begin VB.TextBox txtFax 
            DataField       =   "fax"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   960
            TabIndex        =   7
            Top             =   1320
            Width           =   2895
         End
         Begin VB.Label lblStr 
            Caption         =   "Strasse"
            Height          =   315
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lbPLZOrt 
            Caption         =   "PLZ / Ort"
            Height          =   315
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblTel 
            Caption         =   "Tel.:"
            Height          =   315
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblFax 
            Caption         =   "Fax.:"
            Height          =   315
            Left            =   120
            TabIndex        =   28
            Top             =   1320
            Width           =   735
         End
      End
      Begin VB.Label lblBezirk 
         Caption         =   "Bezirk"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame FrameInfo 
      Caption         =   "Datensatz Informationen"
      Height          =   2535
      Left            =   360
      TabIndex        =   14
      Top             =   840
      Width           =   5655
      Begin VB.TextBox txtModify 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "MODIFY"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1800
         Width           =   4000
      End
      Begin VB.TextBox txtModifyFrom 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "MFROM"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1440
         Width           =   4000
      End
      Begin VB.TextBox txtCreate 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "DATECREATE"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1080
         Width           =   4000
      End
      Begin VB.TextBox txtCreateFrom 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "CFROM"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   4000
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "ID"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   360
         Width           =   4000
      End
      Begin VB.Label lblModify 
         Caption         =   "geändert am"
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblModifyFrom 
         Caption         =   "geändert von"
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblCreate 
         Caption         =   "erstellt am"
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblCreateFrom 
         Caption         =   "erstellt von"
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblID 
         Caption         =   "Datensatz ID"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   0
      Picture         =   "frmEditOLG.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   11
      ToolTipText     =   "Datensatz speichern"
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   375
      Left            =   480
      Picture         =   "frmEditOLG.frx":058A
      Style           =   1  'Grafisch
      TabIndex        =   12
      ToolTipText     =   "Datensatz löschen"
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton cmdWord 
      Height          =   375
      Left            =   960
      Picture         =   "frmEditOLG.frx":0914
      Style           =   1  'Grafisch
      TabIndex        =   13
      ToolTipText     =   "Neues Anschreiben"
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Übernehmen"
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtGericht 
      DataField       =   "Name"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5741
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Daten"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Info"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblGericht 
      Caption         =   "Gericht"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmEditOLG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                                     ' Variaben Deklaration erzwingen
Private Const MODULNAME = "frmEditOLG"                              ' Modulname für Fehlerbehandlung

Private bInit As Boolean                                            ' Wird True gesetzt wenn Alle werte geladen
Private bDirty As Boolean                                           ' Wird True gesetzt wenn Daten verändert wurden
Private bNew  As Boolean                                            ' Wird gesetzt wenn neuer DS sonst Update
Private bModal As Boolean                                           ' Ist Modal Geöffnet
Private szID As String                                              ' DS ID
Private ThisDBCon As Object                                         ' Aktuelle DB Verbindung
Private frmParent As Form                                           ' Aufrufendes DB form
Private szIDField As String
Private ThisFramePos As FramePos                                    ' Standart Frame Position

Private szSQL As String                                             ' SQL Statement
Private szWhere As String                                           ' Where Klausel
Private szIniFilePath As String                                     ' Pfad der Ini datei
Private lngImage As Integer                                         ' Imagiendex

Private szRootkey As String                                         ' =
Private szDetailKey As String                                       ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)

Private CurrentStep As String                                       ' Aktueller WorkflowSchritt
Private OldCmbValue As String                                        'Alter Combo wert

Private Type FramePos                                               ' Positions Datentyp
    Top As Single                                                   ' Top position (oben)
    Left As Single                                                  ' Left Position (Links)
    Height As Single                                                ' Height (Höhe)
    Width As Single                                                 ' Width (Breite)
End Type

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
On Error Resume Next                                                ' Fehlerbehandlung deaktiviern
    If bDirty Then szID = ""                                        ' Wenn ungespeichert ID Leeren
    If bModal Then                                                  ' Wenn Modal
        Me.Hide                                                     ' dann ausblenden
    Else                                                            ' Sonst
        Call EditFormUnload(Me)                                     ' AUs Edit Form Array entfernen
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Public Function InitEditForm(parentform As Form, dbCon As Object, RootKey As String, _
        DetailKey As String, Optional bDialog As Boolean)

    Dim i As Integer                                                ' counter
    Dim tmpArray() As String
    
On Error GoTo Errorhandler

    Set frmParent = parentform
    bInit = True                                                    ' Wir initialisieren das Form-> andere vorgänge nicht ausführen
    Set ThisDBCon = dbCon                                           ' Aktuelle DB Verbindung übernehmen
    
    Select Case UCase(RootKey)                                      ' Nur Bestimmte RootKey erlauben
    Case UCase("Landgerichte")
        szRootkey = RootKey
        szIDField = "ID003"
        lblBezirk.Caption = "Oberlandesgerichtsbezirk"
        lblGericht.Caption = "Landgericht"
        ' Liste für Combo Anrede füllen
        Call FillCmbListWithSQL(cmbBezirk, "SELECT OLGNAME002 FROM OLG002 ORDER BY OLGNAME002", ThisDBCon)
    Case UCase("Amtsgerichte")
        szRootkey = RootKey
        szIDField = "ID004"
        lblBezirk.Caption = "Landgerichtsbezirk"
        lblGericht.Caption = "Amtsgericht"
        Call FillCmbListWithSQL(cmbBezirk, "SELECT LGNAME003 FROM LG003 ORDER BY LGNAME003", ThisDBCon)
    Case UCase("Oberlandesgerichte")
        szRootkey = RootKey
        szIDField = "ID002"
        lblGericht.Caption = "Oberlandesgericht"
        lblBezirk.Visible = False
        cmbBezirk.Visible = False
        txtBezirk.Visible = False
    Case Else
        Unload Me
        GoTo exithandler
    End Select
        
    szDetailKey = DetailKey                                         ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)

    szIniFilePath = objObjectBag.Getappdir & objObjectBag.GetXMLFile ' XML inifile festlegen
    Call objTools.GetEditInfoFromXML(szIniFilePath, szRootkey, szSQL, szWhere, lngImage)
    
    Me.Icon = frmParent.ILTree.ListImages(lngImage).Picture         ' Form Icon Setzen

    If szDetailKey = "" Then bNew = True                            ' Neuer Datensatz
    If szDetailKey <> "" Then szWhere = szWhere & "CAST('" & szDetailKey & "' as uniqueidentifier)"
    bModal = bDialog                                                ' Form wird als Dialog geöffnet
    
    Call InitAdoDC(Me, ThisDBCon, szSQL, szWhere)                   ' ADODC Initialisieren
    Me.Refresh
    
    If bNew Then                                                    ' Wenn DS neu
        Adodc1.Recordset.AddNew                                     ' Neuen DS an RS anhängen
        txtID.Text = ThisDBCon.GetValueFromSQL("SELECT NewID()")    ' Neue ID (Guid) ermitteln
        szID = txtID
        bDirty = True                                               ' Dirty da Neu
    Else
        szID = txtID
    End If
    
    Call GetLockedControls(Me)                                      ' Gelockte controls finden
    Call HiglightThisMustFields(Not (bNew Or bDirty))               ' IndexFelder hervorheben
    Call RefreshRelFields
    Call InitFrameGerichtsDaten(True)                               ' Frame Gericht initialisieren
    Call InitFrameInfo(Me)                                          ' Info Frame initialisieren
    Call SetEditFormCaption(Me, szRootkey, txtGericht)              ' Caption mit Gerichtsnamen
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

Private Sub InitFrameGerichtsDaten(Optional bVisible As Boolean)
    
On Error GoTo Errorhandler
    
    With ThisFramePos
        Call FrameGericht.Move(.Left, .Top, _
                .Width, .Height)                                    ' Frame Positionieren
    End With
    FrameGericht.Visible = bVisible                                 ' Sichtbar ?
    
exithandler:
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitFramePersonenDaten", errNr, errDesc)
    Resume exithandler
End Sub

Private Function RefreshRelFields()

    Dim szSQL As String                                             ' SQL Statement
    Dim szWhere As String                                           ' Where Klausel
    
On Error GoTo Errorhandler

    Select Case UCase(szRootkey)
    Case UCase("Landgerichte")
        szSQL = "SELECT TOP 1 OLGNAME002 FROM OLG002"
        szWhere = "ID002="
    Case UCase("Amtsgerichte")
        szSQL = "SELECT TOP 1  LGNAME003 FROM LG003"
        szWhere = "ID003="
    Case UCase("Oberlandesgerichte")
        szSQL = ""
    Case Else
        szSQL = ""
    End Select

    If szSQL <> "" Then
        Call RefreshRelField(Me, cmbBezirk, txtBezirkID, _
                szSQL, szWhere, True)
    End If
    
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
On Error Resume Next
    Call HandleKeyDownEdit(Me, KeyCode, Shift)                      ' Spezielle KeyDon Events dieses Forms
    Call frmParent.HandleGlobalKeyCodes(KeyCode, Shift)             ' Key Down Events der Anwendung
    Err.Clear
End Sub

Private Sub HandleTabClick(TS As TabStrip)
' Behandelt Tab Klicks

On Error GoTo Errorhandler

    Call HandleTabClickNew(Me, TS)                                  ' Wenn bNew dan nur 1. Tab zulassen
    
    If TS.SelectedItem = "Info" Then                                ' Info
        FrameGericht.Visible = False
        FrameInfo.Visible = True
        Exit Sub
    End If
    
    Select Case TS.SelectedItem.Index
    Case 1                                                          ' Gericht
        FrameGericht.Visible = True
        FrameInfo.Visible = False
    Case 2                                                          ' Info
        FrameGericht.Visible = False
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

                                                                    ' *****************************************
                                                                    ' TabSrip Events
Private Sub TabStrip1_Click()
    Call HandleTabClick(TabStrip1)                                  ' Tab Klick behandeln
End Sub
                                                                    ' *****************************************
                                                                    ' Button Events
Private Sub cmdEsc_Click()
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
                                                                    ' Change Events
Private Sub txtBezirk_Change()
    If Not bInit Then Call StandartTextChange(Me, txtBezirk)
End Sub

Private Sub txtfax_Change()
    If Not bInit Then Call StandartTextChange(Me, txtfax)
End Sub

Private Sub txtGericht_Change()
    If Not bInit Then Call StandartTextChange(Me, txtGericht)
End Sub

'Private Sub txtGericht_Change()
'    If Not bInit Then Call StandartTextChange(Me, txtGericht)
'End Sub

Private Sub txtOrt_Change()
    If Not bInit Then Call StandartTextChange(Me, txtOrt)
End Sub

Private Sub txtPLZ_Change()
    If Not bInit Then Call StandartTextChange(Me, txtPLZ)
End Sub

Private Sub txtStr_Change()
    If Not bInit Then Call StandartTextChange(Me, txtStr)
End Sub

Private Sub txtTel_Change()
    If Not bInit Then Call StandartTextChange(Me, txtTel)
End Sub
'Private Sub txtCreate_Change()
'    If bInit Then Exit Sub
'    bDirty = True
'    Call CheckUpdate(Me)
'End Sub
'
'Private Sub txtCreateFrom_Change()
'    If bInit Then Exit Sub
'    bDirty = True
'    Call CheckUpdate(Me)
'End Sub

'Private Sub txtID_Change()
'    If bInit Then Exit Sub
'    bDirty = True
'    Call CheckUpdate(Me)
'End Sub
'
'Private Sub txtModify_Change()
'    If bInit Then Exit Sub
'    bDirty = True
'    Call CheckUpdate(Me)
'End Sub
'
'Private Sub txtModifyFrom_Change()
'    If bInit Then Exit Sub
'    bDirty = True
'    Call CheckUpdate(Me)
'End Sub
                                                                    ' *****************************************
                                                                    ' Fokus Events
Private Sub cmbBezirk_GotFocus()
    Call HiglightCurentField(Me, cmbBezirk, False)
End Sub

Private Sub cmbBezirk_LostFocus()
    Call HiglightCurentField(Me, cmbBezirk, True)
End Sub

Private Sub txtfax_GotFocus()
    Call HiglightCurentField(Me, txtfax, False)
End Sub

Private Sub txtfax_LostFocus()
    Call HiglightCurentField(Me, txtfax, True)
End Sub

Private Sub txtGericht_GotFocus()
    Call HiglightCurentField(Me, txtGericht, False)
End Sub

Private Sub txtGericht_LostFocus()
    Call HiglightCurentField(Me, txtGericht, True)
End Sub

Private Sub txtOrt_GotFocus()
    Call HiglightCurentField(Me, txtOrt, False)
End Sub

Private Sub txtOrt_LostFocus()
    Call HiglightCurentField(Me, txtOrt, True)
End Sub

Private Sub txtPLZ_GotFocus()
    Call HiglightCurentField(Me, txtPLZ, False)
End Sub

Private Sub txtPLZ_LostFocus()
    Call HiglightCurentField(Me, txtPLZ, True)
End Sub

Private Sub txtStr_GotFocus()
    Call HiglightCurentField(Me, txtStr, False)
End Sub

Private Sub txtStr_LostFocus()
    Call HiglightCurentField(Me, txtStr, True)
End Sub

Private Sub txtTel_GotFocus()
    Call HiglightCurentField(Me, txtTel, False)
End Sub

Private Sub txtTel_LostFocus()
    Call HiglightCurentField(Me, txtTel, True)
End Sub

                                                                    ' *****************************************
                                                                    ' Key Down Events
Private Sub txtGericht_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtPLZ_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtStr_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtTel_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtOrt_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtOLG_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtCreate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtCreateFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtfax_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtModify_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtModifyFrom_KeyDown(KeyCode As Integer, Shift As Integer)
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
    Set GetDBConn = ThisDBCon
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

