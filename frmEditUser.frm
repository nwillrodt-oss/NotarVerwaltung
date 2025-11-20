VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEditUser 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Benutzer"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5565
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   0
      Picture         =   "frmEditUser.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Datensatz speichern"
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   375
      Left            =   480
      Picture         =   "frmEditUser.frx":058A
      Style           =   1  'Grafisch
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Datensatz löschen"
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton cmdWord 
      Height          =   375
      Left            =   960
      Picture         =   "frmEditUser.frx":0914
      Style           =   1  'Grafisch
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Neues Anschreiben"
      Top             =   3600
      Width           =   375
   End
   Begin VB.Frame FrameUserdaten 
      Caption         =   "Benutzerdaten"
      Height          =   2415
      Left            =   120
      TabIndex        =   23
      Top             =   840
      Width           =   5295
      Begin VB.TextBox txtVorname 
         DataField       =   "VORNAME001"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   3195
      End
      Begin VB.TextBox txtNachname 
         DataField       =   "NACHNAME001"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   600
         Width           =   3195
      End
      Begin VB.TextBox txtTel 
         DataField       =   "TEL001"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   960
         Width           =   3195
      End
      Begin VB.TextBox txtfax 
         DataField       =   "FAX001"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Top             =   1320
         Width           =   3195
      End
      Begin VB.TextBox txtemail 
         DataField       =   "EMAIL001"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   1680
         Width           =   3195
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Systemverwalter"
         DataField       =   "SYSTEM001"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Label lblVorname 
         Caption         =   "Vorname"
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label lblNachname 
         Caption         =   "Nachname"
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1605
      End
      Begin VB.Label lblTel 
         Caption         =   "Telefon"
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label lblFax 
         Caption         =   "Telefax"
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   1605
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email"
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   1605
      End
   End
   Begin VB.Frame FrameInfo 
      Caption         =   "Datensatz Informationen"
      Height          =   2535
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   5295
      Begin VB.TextBox txtModify 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "MODIFY001"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1800
         Width           =   3000
      End
      Begin VB.TextBox txtModifyFrom 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "MFROM001"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1440
         Width           =   3000
      End
      Begin VB.TextBox txtCreate 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "CREATE001"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Width           =   3000
      End
      Begin VB.TextBox txtCreateFrom 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "CFROM001"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   3000
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "ID001"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   360
         Width           =   3000
      End
      Begin VB.Label lblModify 
         Caption         =   "geändert am"
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   1800
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
      Begin VB.Label lblCreate 
         Caption         =   "erstellt am"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblCreateFrom 
         Caption         =   "erstellt von"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblID 
         Caption         =   "Datensatz ID"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5318
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Benutzerdaten"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Info"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtUsername 
      DataField       =   "USERNAME001"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   2600
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Übernehmen"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   3600
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
   Begin VB.Label lblUsername 
      Caption         =   "Benutzername"
      Height          =   315
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   1600
   End
End
Attribute VB_Name = "frmEditUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                                     ' Variaben Deklaration erzwingen
Private Const MODULNAME = "frmEditUser"                             ' Modulname für Fehlerbehandlung

Private bInit As Boolean                                            ' Wird True gesetzt wenn Alle werte geladen
Private bDirty As Boolean                                           ' Wird True gesetzt wenn Daten verändert wurden
Private bNew  As Boolean                                            ' Wird gesetzt wenn neuer DS sonst Update
Private bModal As Boolean                                           ' Ist Modal Geöffnet
Private szID As String                                              ' DS ID
Private ThisDBCon As Object                                         ' Aktuelle DB Verbindung
Private frmParent As Form                                           ' Aufrufendes DB form
Private szIDField As String
Private ThisFramePos As FramePos                                    ' Standart Frame Position

Private szSQL As String                                             ' SQL für USER001
Private szWhere As String                                           ' Where Klausel
Private szIniFilePath As String                                     ' Pfad der Ini datei
Private lngImage As Integer                                         ' Imagiendex

'Private rsUser As ADODB.Recordset                                   ' RS mit User Daten

Private szRootkey As String                                         ' = Benutzerverwaltung
Private szDetailKey As String                                       ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)

Private CurrentStep As String                                       ' Aktueller WorkflowSchritt

Private Type FramePos                                               ' Positions Datentyp
    Top As Single                                                   ' Top position (oben)
    Left As Single                                                  ' Left Position (Links)
    Height As Single                                                ' Height (Höhe)
    Width As Single                                                 ' Width (Breite)
End Type

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
        
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    Set frmParent = parentform
    bInit = True                                                    ' Wir initialisieren das Form-> andere vorgänge nicht ausführen
    Set ThisDBCon = dbCon                                           ' Aktuelle DB Verbindung übernehmen
    szRootkey = "Benutzerverwaltung"                                ' für Caption
    szIDField = "ID001"
    szDetailKey = DetailKey                                         ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)
    
    szIniFilePath = objObjectBag.Getappdir & objObjectBag.GetXMLFile ' XML inifile festlegen
    Call objTools.GetEditInfoFromXML(szIniFilePath, szRootkey, szSQL, szWhere, lngImage)
    
    Me.Icon = frmParent.ILTree.ListImages(lngImage).Picture         ' Form Icon Setzen
    
    If szDetailKey = "" Then bNew = True                            ' Neuer Datensatz
    If szDetailKey <> "" Then szWhere = szWhere & "CAST('" & szDetailKey & "' as uniqueidentifier)"
    bModal = bDialog                                                ' Als Dialog anzeigen
    
    Call InitAdoDC(Me, ThisDBCon, szSQL, szWhere)                   ' ADODC Initialisieren
    Me.Refresh
    
    If bNew Then                                                    ' Wenn DS Neu
        Adodc1.Recordset.AddNew                                     ' Neuen DS an RS anhängen
        txtID.Text = ThisDBCon.GetValueFromSQL("SELECT NewID()")    ' Neue ID (Guid) ermitteln
        szID = txtID
        bDirty = True                                               ' Dirty da Neu
    Else
        szID = txtID
    End If
    
    Call GetLockedControls(Me)                                      ' Gelockte controls finden
    Call HiglightThisMustFields(Not (bNew Or bDirty))               ' IndexFelder hervorheben
    Call InitFrameInfo(Me)                                          ' Info Frame initialisieren
    Call SetEditFormCaption(Me, szRootkey, txtUsername)             ' Form Caption setzen mit Usernamen
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
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call HiglightMustFields(Me, bDeHiglight)                        ' Alle PK ind IndexFields entfärben
'    Call HiglightMustField(Me, txtAZ, bDeHiglight)
    Err.Clear                                                       ' Evtl. error clearen
End Sub

Private Sub InitFrameUserDaten()
On Error Resume Next                                                ' Fehlerbehandlung deaktiviere
    With ThisFramePos
        Call FrameUserdaten.Move(.Left, .Top, .Width, .Height)      ' Frame Positionieren
    End With
    Err.Clear                                                       ' Evtl. error clearen
End Sub

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
    
    If TS.SelectedItem = "Info" Then                                ' Info
        FrameUserdaten.Visible = False
        FrameInfo.Visible = True
        Exit Sub
    End If
    
    Select Case TS.SelectedItem.Index
    Case 1                                                          ' Benutzerdaten
        FrameUserdaten.Visible = True
        FrameInfo.Visible = False
    Case 2                                                          ' Info
        FrameUserdaten.Visible = False
        FrameInfo.Visible = True
    Case Else
    
    End Select
    
exithandler:
On Error Resume Next
    Me.Refresh                                                      ' Form Aktualisieren
    
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
    
    If bValidationFaild Then                                        ' Wenn Validierung Gescheitert
        Call objError.ShowErrMsg(szMSG, vbInformation, szTitle)     ' Hinweis meldung anzeigen
        FocusCTL.SetFocus                                           ' Fokus setzen
        ValidateEditForm = False                                    ' Ruckgabewert setzen
    Else
        ValidateEditForm = True                                     ' Erfolg zurück
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
                                                                    ' *****************************************
                                                                    ' Change Events
Private Sub chkSystem_Validate(Cancel As Boolean)
    If bInit Then Exit Sub
    bDirty = True
    Call CheckUpdate(Me)
End Sub

Private Sub txtemail_Change()
    If Not bInit Then Call StandartTextChange(Me, txtemail)
End Sub

Private Sub txtfax_Change()
    If Not bInit Then Call StandartTextChange(Me, txtfax)
End Sub

Private Sub txtNachname_Change()
    If Not bInit Then Call StandartTextChange(Me, txtNachname)
End Sub

Private Sub txtTel_Change()
    If Not bInit Then Call StandartTextChange(Me, txtTel)
End Sub

Private Sub txtUsername_Change()
    If Not bInit Then Call StandartTextChange(Me, txtUsername)
End Sub

Private Sub txtVorname_Change()
    If Not bInit Then Call StandartTextChange(Me, txtVorname)
End Sub

                                                                    ' *****************************************
                                                                    ' Fokus Events
Private Sub txtemail_GotFocus()
    Call HiglightCurentField(Me, txtemail, False)
End Sub

Private Sub txtemail_LostFocus()
    Call HiglightCurentField(Me, txtemail, True)
End Sub

Private Sub txtfax_GotFocus()
    Call HiglightCurentField(Me, txtfax, False)
End Sub

Private Sub txtfax_LostFocus()
    Call HiglightCurentField(Me, txtfax, True)
End Sub

Private Sub txtNachname_GotFocus()
    Call HiglightCurentField(Me, txtNachname, False)
End Sub

Private Sub txtNachname_LostFocus()
    Call HiglightCurentField(Me, txtNachname, True)
End Sub

Private Sub txtTel_GotFocus()
    Call HiglightCurentField(Me, txtTel, False)
End Sub

Private Sub txtTel_LostFocus()
    Call HiglightCurentField(Me, txtTel, True)
End Sub

Private Sub txtUsername_GotFocus()
    Call HiglightCurentField(Me, txtUsername, False)
End Sub

Private Sub txtUsername_LostFocus()
    Call HiglightCurentField(Me, txtUsername, True)
End Sub

Private Sub txtVorname_GotFocus()
    Call HiglightCurentField(Me, txtVorname, False)
End Sub

Private Sub txtVorname_LostFocus()
    Call HiglightCurentField(Me, txtVorname, True)
End Sub
                                                                    ' *****************************************
                                                                    ' Key Down Events
Private Sub txtemail_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtfax_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtNachname_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtTel_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtUsername_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtVorname_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub
                                                                    ' *****************************************
                                                                    ' Adodc Events
Private Sub Adodc1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    If bInit Then fCancelDisplay = True
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

