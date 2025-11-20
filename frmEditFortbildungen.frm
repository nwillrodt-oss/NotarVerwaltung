VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEditFortbildungen 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Fortbildungen"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6765
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame FrameInfo 
      Caption         =   "Datensatz Informationen"
      Height          =   2535
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   7815
      Begin VB.TextBox txtModify 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "MODIFY011"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1800
         Width           =   4000
      End
      Begin VB.TextBox txtModifyFrom 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "MFROM011"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1440
         Width           =   4000
      End
      Begin VB.TextBox txtCreate 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "CREATE011"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1080
         Width           =   4000
      End
      Begin VB.TextBox txtCreateFrom 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "CFROM011"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   720
         Width           =   4000
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "ID011"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   4000
      End
      Begin VB.Label lblModify 
         Caption         =   "geändert am"
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblModifyFrom 
         Caption         =   "geändert von"
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblCreate 
         Caption         =   "erstellt am"
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   1080
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
      Begin VB.Label lblID 
         Caption         =   "Datensatz ID"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CheckBox chkAnerkannt 
      Caption         =   "Anerkannt"
      DataField       =   "ANERKANNT011"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   3240
      TabIndex        =   4
      Top             =   840
      Width           =   2535
   End
   Begin VB.Frame FrameTeilnehmer 
      Caption         =   "Teilnehmer"
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1455
      Begin MSComctlLib.ListView LVTeilnehmer 
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.TextBox txtHalbtage 
      DataField       =   "ANZHT011"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtThema 
      DataField       =   "THEMA011"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
   Begin VB.ComboBox cmbVeranstalter 
      DataField       =   "VERANSTALTER011"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   4800
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   4680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "FORT011"
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Übernehmen"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtDatum 
      DataField       =   "DATUM011"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   57212929
      CurrentDate     =   39216
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3375
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5953
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Teilnehmer"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Info"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblHalbtage 
      Caption         =   "Halbtage"
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1395
   End
   Begin VB.Label lblThema 
      Caption         =   "Thema"
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1395
   End
   Begin VB.Label lblVeranstalter 
      Caption         =   "Veranstalter"
      Height          =   315
      Left            =   3240
      TabIndex        =   10
      Top             =   480
      Width           =   1395
   End
   Begin VB.Label lblDatum 
      Caption         =   "Anfangsdatum"
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1395
   End
   Begin VB.Menu kmnuLVTeilnehmer 
      Caption         =   "KontextmenueLVTeilnehmer"
      Visible         =   0   'False
      Begin VB.Menu kmnuAddTeilnehmer 
         Caption         =   "Teilnehmer hinzufügen"
      End
      Begin VB.Menu kmnuDelTeilnehmer 
         Caption         =   "Teilnehmer entfernen"
      End
   End
End
Attribute VB_Name = "frmEditFortbildungen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULNAME = "frmEditFortbildungen"

Private bInit As Boolean            ' Wird True gesetzt wenn Alle werte geladen
Public bDirty As Boolean            ' Wird True gesetzt wenn Daten verändert wurden
Public bNew  As Boolean             ' Wird gesetzt wenn neuer DS sonst Update

Private frmParent As Form           ' Aufrufendes DB form

Private szSQL As String             ' SQL für FORT011
Private szWhere As String           ' Where Klausel
Private szIniFilePath As String     ' Pfad der Ini datei
Private lngImage As Integer         ' Imagiendex

Private szRootkey As String         ' = Fortbildungen
Private szDetailKey As String       ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)

Public ID As String

Private ThisDBcon As Object     ' Aktuelle DB Verbindung

Public lngFrametopPos As Integer
Public lngFrameLeftPos As Integer
Public lngFrameWidth As Integer
Public lngFrameHeight As Integer

Private rsTeilnehmer As ADODB.Recordset


Private Sub Form_Load()
    
On Error GoTo Errorhandler

    lngFrametopPos = TabStrip1.Top + 360
    lngFrameLeftPos = 120
    lngFrameWidth = TabStrip1.Width - 240
    lngFrameHeight = TabStrip1.Height - 480
    
    Call EditFormLoad(Me, szRootkey)

exithandler:
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "Form_Load", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveColumnWidth(LVTeilnehmer, szRootkey & "LV", True)
    Call EditFormUnload(Me)
End Sub

Public Function InitEditForm(parentform As Form, dbCon As Object, DetailKey As String)

    Dim i As Integer ' counter
    Dim tmpArray() As String
        
On Error GoTo Errorhandler
    
    Set frmParent = parentform
    bInit = True                    ' Wir initialisieren das Form-> andere vorgänge nicht ausführen
    Set ThisDBcon = dbCon           ' Aktuelle DB Verbindung übernehmen
    szRootkey = "Fortbildungen"     ' für Caption
    szDetailKey = DetailKey         ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)
    
    szIniFilePath = objObjectBag.Getappdir & objObjectBag.GetXMLFile
    Call objTools.GetEditInfoFromXML(szIniFilePath, szRootkey, szSQL, szWhere, lngImage)
    
    Me.Icon = frmParent.ILTree.ListImages(lngImage).Picture
      
    If szDetailKey = "" Then bNew = True ' Neue Datensatz
    If szDetailKey <> "" Then szWhere = szWhere & "'" & szDetailKey & "'"
    
    Call InitAdoDC(Me, ThisDBcon, szSQL, szWhere)
    
    If bNew Then
    
    On Error Resume Next
        dtDatum.DataField = ""
        dtDatum.Value = Format(Now(), "dd.mm.yyyy")
        dtDatum.DataField = "DATUM011"
        Err.Clear
    On Error GoTo Errorhandler
        
        Adodc1.Recordset.AddNew
    Else
         ID = Me.txtID
    End If
    
    Call InitFrameFortbildungDaten              ' Frame Benutzer informationen Initialisieren
    Call InitFrameInfo(Me)
    Call RefreshFrameTeilnehmer(True)

    
    Call SetEditFormCaption(Me, szRootkey, txtThema)
    
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

Private Sub InitFrameFortbildungDaten()
    
On Error GoTo Errorhandler
    
    Call PosFrameAndListView(Me, FrameTeilnehmer, True)
    
    'Call FramePersonenDaten.Move(lngFrameLeftPos, lngFrametopPos, lngFrameWidth, lngFrameHeight)
    
exithandler:
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitFrameFortbildungDaten", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub RefreshFrameTeilnehmer(Optional bVisible As Boolean)

On Error Resume Next

    'Set rsTeilnehmer = InitLVFrame(Me,  ThisDBcon, szIniFilePath, szRootkey, szDetailKey, FrameTeilnehmer, LVTeilnehmer, "Fortbildungen\*\Teilnehmer")
    FrameTeilnehmer.Visible = bVisible
    
End Sub

Private Sub HandleMenueKlick(szMenueName As String, Optional szCaption As String)
     
On Error GoTo Errorhandler
    
    If HandleLVkmnuNew(Me, szCaption) Then GoTo exithandler
    
    Select Case szMenueName
    Case "kmnuAddTeilnehmer"
        
        'Call oprn
'        Call SetRelationinLV(Me, "Personen", "Nachname", _
'                ThisDBcon, LVTeilnehmer, rsTeilnehmer, _
'                "FK011014", "FK010014")
        Call RefreshFrameTeilnehmer(True)
    Case "kmnuDelTeilnehmer"
        Call DelRelationinLV(Me, "Teilnehmer", ThisDBcon, LVTeilnehmer, rsTeilnehmer, "ID014", "AFORT014")
       Call RefreshFrameTeilnehmer(True)
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

Private Sub TabStrip1_Click()

    Call HandleTabClickNew(Me, TabStrip1)        ' Wenn bNew dan nur 1. Tab zulassen
    
    If TabStrip1.SelectedItem = "Info" Then
        FrameTeilnehmer.Visible = False
        FrameInfo.Visible = True
        Exit Sub
    End If
    
    Select Case TabStrip1.SelectedItem.Index
    Case 1
        FrameTeilnehmer.Visible = True
        FrameInfo.Visible = False
    Case 2
       
    Case Else
    
    End Select
    
    Me.Refresh
End Sub

Private Sub ShowKontextMenu()
    PopupMenu kmnuLVTeilnehmer
End Sub


Private Sub LVTeilnehmer_DblClick()
    Call HandleEditLVDoubleClick(Me, Me.LVTeilnehmer, frmParent)
End Sub

Private Sub kmnuAddTeilnehmer_Click()
    Call HandleMenueKlick("kmnuAddTeilnehmer", kmnuAddTeilnehmer.Caption)
End Sub

Private Sub kmnuDelTeilnehmer_Click()
    Call HandleMenueKlick("kmnuDelTeilnehmer", kmnuDelTeilnehmer.Caption)
End Sub

Private Sub LVTeilnehmer_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then Call ShowKontextMenu
End Sub

Private Sub LVTeilnehmer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SetColumnOrder(LVTeilnehmer, ColumnHeader)
End Sub

Private Sub Adodc1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    If bInit Then fCancelDisplay = True
End Sub

'******************************************************************* Botton Events
Private Sub cmdESC_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call UpdateEditForm(Me)
    Call frmParent.RefreshListView                      ' ListView aktualisieren
    Unload Me
End Sub

Public Sub cmdUpdate_Click()
    Call UpdateEditForm(Me)
    Call frmParent.RefreshListView                      ' ListView aktualisieren
End Sub

'******************************************************************* Combo Events
Private Sub cmbVeranstalter_Validate(Cancel As Boolean)
    If bInit Then Exit Sub
    bDirty = True
    Call CheckUpdate(Me)
End Sub

Private Sub cmbVeranstalter_Change()
    If Not bInit Then Call StandartTextChange(Me, txtHalbtage)
End Sub

'******************************************************************* Change Events
Private Sub dtDatum_Change()
    If bInit Then Exit Sub
    bDirty = True
    ' Bei DTPicker Wert nochmal übernehmen da Sonst die Datenbindung nicht funzt
    Me.Adodc1.Recordset.Fields("DATUM011").Value = Me.dtDatum.Value
    Me.dtDatum.Refresh
    Call CheckUpdate(Me)
End Sub

Private Sub chkAnerkannt_Validate(Cancel As Boolean)
    If bInit Then Exit Sub
    bDirty = True
    Call CheckUpdate(Me)
End Sub

Private Sub txtHalbtage_Change()
    If Not bInit Then Call StandartTextChange(Me, txtHalbtage)
End Sub

Private Sub txtThema_Change()
    If Not bInit Then Call StandartTextChange(Me, txtThema)
End Sub
