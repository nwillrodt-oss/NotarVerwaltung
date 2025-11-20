VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEditStellen 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Stellen"
   ClientHeight    =   6450
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   7020
   StartUpPosition =   1  'Fenstermitte
   Begin VB.ComboBox cmbAusschreibung 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Frame FrameWorkflow 
      Caption         =   "Vorgang"
      Height          =   4215
      Left            =   120
      TabIndex        =   36
      Top             =   1560
      Width           =   6735
      Begin VB.CommandButton cmdNextStep 
         Caption         =   "Weiter >>"
         Height          =   375
         Left            =   5400
         TabIndex        =   39
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrevStep 
         Caption         =   "<< Zurück"
         Height          =   375
         Left            =   4080
         TabIndex        =   38
         Top             =   3720
         Width           =   1215
      End
      Begin MSComctlLib.ListView LVStep 
         Height          =   2295
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ILTree"
         SmallIcons      =   "ILTree"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblStepDesc 
         Caption         =   "Schritt Beschreibung"
         Height          =   855
         Left            =   120
         TabIndex        =   40
         Top             =   2640
         Width           =   6255
      End
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   0
      Picture         =   "frmEditStellen.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   35
      ToolTipText     =   "Datensatz speichern"
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   375
      Left            =   480
      Picture         =   "frmEditStellen.frx":058A
      Style           =   1  'Grafisch
      TabIndex        =   34
      ToolTipText     =   "Datensatz löschen"
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton cmdWord 
      Height          =   375
      Left            =   960
      Picture         =   "frmEditStellen.frx":0914
      Style           =   1  'Grafisch
      TabIndex        =   33
      ToolTipText     =   "Neues Anschreiben"
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton cmdAusSuchen 
      Height          =   315
      Left            =   3600
      Picture         =   "frmEditStellen.frx":0C9E
      Style           =   1  'Grafisch
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtIDAus 
      DataField       =   "FK020012"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   2640
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtFrist 
      DataField       =   "FRIST012"
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
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame FrameDokumente 
      Caption         =   "Dokumente"
      Height          =   855
      Left            =   120
      TabIndex        =   29
      Top             =   3480
      Width           =   1575
      Begin MSComctlLib.ListView LVDokumente 
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ILTree"
         SmallIcons      =   "ILTree"
         ColHdrIcons     =   "ILTree"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.TextBox txtAZ 
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame FrameInfo 
      Caption         =   "Datensatz Informationen"
      Height          =   2535
      Left            =   480
      TabIndex        =   15
      Top             =   1680
      Width           =   6735
      Begin VB.TextBox txtModify 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "MODIFY012"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1800
         Width           =   5295
      End
      Begin VB.TextBox txtModifyFrom 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "MFROM012"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1440
         Width           =   5295
      End
      Begin VB.TextBox txtCreate 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "CREATE012"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1080
         Width           =   5295
      End
      Begin VB.TextBox txtCreateFrom 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "CFROM012"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   720
         Width           =   5295
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "ID012"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   5295
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
   Begin VB.Frame FrameBewerber 
      Caption         =   "Bewerbungen"
      Height          =   735
      Left            =   120
      TabIndex        =   26
      Top             =   1560
      Width           =   1575
      Begin MSComctlLib.ListView LVBewerber 
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         Icons           =   "ILTree"
         SmallIcons      =   "ILTree"
         ColHdrIcons     =   "ILTree"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   6000
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Height          =   4695
      Left            =   0
      TabIndex        =   7
      Top             =   1200
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8281
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bewerbungen"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dokumente"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Vorgang"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Info"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbBezirk 
      DataField       =   "BEZIRK012"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtAnzStellen 
      DataField       =   "ANZ012"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   6360
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin MSComCtl2.DTPicker DTFrist 
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
      Left            =   5760
      TabIndex        =   6
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   57933825
      CurrentDate     =   39217
   End
   Begin VB.TextBox txtBeschreibung 
      DataField       =   "BESCH012"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   5655
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Übernehmen"
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   6000
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ILTree 
      Left            =   2280
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   28
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":1542
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":185C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":1DF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":2390
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":306A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":3D44
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":42DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":4878
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":4C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":4FAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":500A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":5068
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":5602
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":5B9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":6136
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":66D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":6C6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":7204
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":779E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":8848
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":8DE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":937C
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":9916
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":9EB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":A24A
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":A7E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditStellen.frx":AD7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblAZ 
      Caption         =   "AZ"
      Height          =   315
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   1140
   End
   Begin VB.Label lblAnzStellen 
      Caption         =   "Anz. Stellen"
      Height          =   315
      Left            =   4800
      TabIndex        =   14
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label lblFrist 
      Caption         =   "Bewerbungsfrist"
      Height          =   315
      Left            =   4440
      TabIndex        =   13
      Top             =   120
      Width           =   1260
   End
   Begin VB.Label lblBezirk 
      Caption         =   "Bezirk"
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1140
   End
   Begin VB.Label lblBeschreibung 
      Caption         =   "Beschreibung"
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1140
   End
   Begin VB.Menu kmnuLVBewerber 
      Caption         =   "KontextMenueLVBewerber"
      Visible         =   0   'False
      Begin VB.Menu kmnuLVBewerberOpen 
         Caption         =   "Bewerbung anzeigen"
      End
      Begin VB.Menu kmnuLVBewerberAdd 
         Caption         =   "Bewerbung hinzufügen"
      End
      Begin VB.Menu kmnuLVBewerberDel 
         Caption         =   "Bewerbung entfernen"
      End
      Begin VB.Menu kmnuLVBewerberPersonOpen 
         Caption         =   "Personendaten anzeigen"
      End
   End
   Begin VB.Menu kmnuLVDocument 
      Caption         =   "KontextMenueLVDokumente"
      Visible         =   0   'False
      Begin VB.Menu kmnuLVDokumentAdd 
         Caption         =   "Neues Dokument erstellen"
      End
      Begin VB.Menu kmnuLVDokumentOpen 
         Caption         =   "Dokument anzeigen"
      End
      Begin VB.Menu kmnuLVDokumentDel 
         Caption         =   "Dokument löschen"
      End
      Begin VB.Menu kmnuLVDokumentImport 
         Caption         =   "Dokument Importieren"
      End
   End
End
Attribute VB_Name = "frmEditStellen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                                     ' Variaben Deklaration erzwingen
Private Const MODULNAME = "frmEditStellen"                          ' Modulname für Fehlerbehandlung

Private bInit As Boolean                                            ' Wird True gesetzt wenn Alle werte geladen
Private bDirty As Boolean                                           ' Wird True gesetzt wenn Daten verändert wurden
Private bNew  As Boolean                                            ' Wird gesetzt wenn neuer DS sonst Update
Private bModal As Boolean                                           ' Ist Modal Geöffnet
Private szID As String                                              ' DS ID
Private ThisDBCon As Object                                         ' Aktuelle DB Verbindung
Private frmParent As Form                                           ' Aufrufendes DB form
Private szIDField As String
Private ThisFramePos As FramePos                                    ' Standart Frame Position

Private szSQL As String                                             ' SQL für STELLEN012
Private szWhere As String                                           ' Where Klausel
Private szIniFilePath As String                                     ' Pfad der Ini datei
Private lngImage As Integer                                         ' Imagiendex

Private szRootkey As String                                         ' = Stellen
Private szDetailKey As String                                       ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)

Private CurrentStep As String                                       ' Aktueller WorkflowSchritt
Private NextStep As String                                          ' Nächster Schritt
Private PrevStep As String                                          ' Voheriger Schritt
Private lngWorkflowLevel As Integer                                 ' Workflow Ebene

Private OldCmbValue As String                                       ' Alter Combo wert

Private AusschreibungsID As String
Private szJahr As String
Private szAZ As String

Private rsBewerber As ADODB.Recordset                               ' RS mit bewerber Daten
Private rsDokumente As ADODB.Recordset                              ' RS mit Dokumenten

Private Type FramePos                                               ' Positions Datentyp
    Top As Single                                                   ' Top position (oben)
    Left As Single                                                  ' Left Position (Links)
    Height As Single                                                ' Height (Höhe)
    Width As Single                                                 ' Width (Breite)
End Type

Private Sub Form_Activate()
On Error Resume Next                                                ' Fehlerbehandlung deaktiviern
    If bInit Or bDirty Then Exit Sub                                ' Nicht bei initialisierung
    Call RefreshEditForm
    Err.Clear                                                       ' evtl. Error clearen
End Sub

Private Sub Form_Load()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call EditFormLoad(Me, szRootkey)                                ' Allg. Formload Aufrufen
    Call InitEditButtonMenue(Me, True, True, True)                  ' Buttonleiste initialisieren
    Call RemoveTabByCaption(TabStrip1, "Vorgang")                   ' bis auf weiteres ausblenden
    With ThisFramePos
        Call GetTabStrimClientPos(TabStrip1, .Top, .Left, _
                .Height, .Width)                                    ' Frame Positionen aus TabStrip ermitteln
    End With
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next                                                ' Fehlerbehandlung deaktiviern
    Call SaveColumnWidth(LVBewerber, szRootkey & "LV", True)        ' Spaltenbreiten für LV Speichern
    Call SaveColumnWidth(LVDokumente, szRootkey & "LV", True)       ' Spaltenbreiten für LV Speichern
    rsBewerber.Close                                                ' RS Bewerber schliessen
    rsDokumente.Close                                               ' RS Dokumente schliessen
    If bDirty Then szID = ""                                        ' Wenn ungespeichert ID Leeren
    If bModal Then                                                  ' Wenn Modal
        Me.Hide                                                     ' dann ausblenden
    Else                                                            ' Sonst
        Call EditFormUnload(Me)                                     ' AUs Edit Form Array entfernen
    End If
    Err.Clear                                                       ' Evtl. Error clearen
End Sub

Public Function InitEditForm(parentform As Form, dbCon As Object, DetailKey As String, _
        Optional bDialog As Boolean)

    Dim i As Integer                                                ' counter
    Dim tmpArray() As String                                        ' Array für zusammengesetzten DetailKey
        
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    Set frmParent = parentform                                      ' Aufrufendes Form merken
    bInit = True                                                    ' Wir initialisieren das Form
                                                                    ' -> andere vorgänge nicht ausführen
    Set ThisDBCon = dbCon                                           ' Aktuelle DB Verbindung übernehmen
    szRootkey = "Ausgeschriebene Stellen"                           ' für Caption
    szIDField = "ID012"
    'lngWorkflowLevel = 2                                            ' Für Workflow
                                                                    ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)
    If InStr(DetailKey, ";") Then                                   ' DetailKey evtl. aufspalten
        tmpArray = Split(DetailKey, ";")
    On Error Resume Next
        szDetailKey = tmpArray(0)
        AusschreibungsID = tmpArray(1)
        szJahr = tmpArray(2)
        szAZ = tmpArray(3)
        Err.Clear
    Else
        szDetailKey = DetailKey
    End If
    bModal = bDialog                                                ' Form wird als Dialog geöffnet
    
    szIniFilePath = objObjectBag.Getappdir & objObjectBag.GetXMLFile ' XML inifile festlegen
    Call objTools.GetEditInfoFromXML(szIniFilePath, szRootkey, szSQL, szWhere, lngImage)
    
    Me.Icon = ILTree.ListImages(lngImage).Picture                   ' Form Icon Setzen
    
    If szDetailKey = "" Then bNew = True                            ' Neuer Datensatz
    
    If szDetailKey <> "" Then
        szID = szDetailKey
        szWhere = szWhere & "CAST('" & szDetailKey & "' as uniqueidentifier)"
    End If
    
    Call FillCmbListWithSQL(cmbBezirk, _
            "SELECT AGNAME004 FROM AG004", ThisDBCon)               ' Liste für Combo Bezirk füllen
    Call FillCmbListWithSQL(cmbAusschreibung, "SELECT AZ020 + ' (' + CAST(JAHR020 as varchar(4)) + ')' FROM AUSSCHREIBUNG020", _
            ThisDBCon)                                              ' Liste für Combo Ausschreibung füllen
            
    Call InitAdoDC(Me, ThisDBCon, szSQL, szWhere)                   ' ADODC Initialisieren
    
    If bNew Then                                                    ' Wenn DS neu
        Adodc1.Recordset.AddNew                                     ' Neuen DS an RS anhängen
        txtID.Text = ThisDBCon.GetValueFromSQL("SELECT NewID()")    ' Neue ID (Guid) ermitteln
        szID = txtID                                                ' ID eintragen
        If AusschreibungsID <> "" Then
            txtIDAus = AusschreibungsID
        Else
            AusschreibungsID = GetIDFromNode()
            txtIDAus = AusschreibungsID
        End If
        If szJahr <> "" And szAZ <> "" Then txtAZ = szAZ & " (" & szJahr & ")" ' Wenn Jahr und AZ Übergeben
        
        
        Call FormatDTPicker(Me, DTFrist, Now())
        Call DTFrist_Change
        bDirty = True                                               ' Dirty da Neu
        txtAZ.Enabled = True                                        ' AZ kan eingegeben werden
    Else
        AusschreibungsID = txtIDAus.Text
        szID = Me.txtID
'        txtAZ.Enabled = False
    End If
    If AusschreibungsID <> "" Then
        Call RefreshRelFields
    End If
    Call GetLockedControls(Me)                                      ' Gelockte controls finden
    Call HiglightThisMustFields(Not (bNew Or bDirty))               ' IndexFelder hervorheben
    Call InitFrameInfo(Me)                                          ' Frame Info Initialisieren
    Call RefreshFrameBewerber(True)                                 ' Frame & LV Bewerber Initialisieren & anzeigen
    Call RefreshFrameDokumente(False)                               ' Frame & LV Dokumente Initialisieren
    Call RefreshFrameWorkflow(False)                                ' Frame Workflow initialisieren
    Call SetEditFormCaption(Me, szRootkey, cmbBezirk)               ' Formm Caption mit Bezirk setzen
    Call CheckUpdate(Me)                                            ' Buttons evtl. enablen / disablen
    
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
    Call HiglightMustField(Me, cmbAusschreibung, bDeHiglight)
    Err.Clear                                                       ' Evtl. Error Clearen
End Sub

Private Sub RefreshFrameBewerber(Optional bVisible As Boolean)

    Dim i As Integer                                                ' Counter für Bewerber
    Dim x As Integer                                                ' Counter für Items
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    LVBewerber.Tag = "Ausgeschriebene Stellen\StellenJahr\*\Bewerbungen"
    Set rsBewerber = RefreshFrame(Me, FrameBewerber, LVBewerber, "Ausgeschriebene Stellen", "Bewerbungen", bVisible)
    
    If rsBewerber Is Nothing Then Exit Sub
    If rsBewerber.RecordCount = 0 Then Exit Sub
    rsBewerber.MoveFirst
    For i = 0 To rsBewerber.RecordCount - 1
        For x = 1 To LVBewerber.ListItems.Count
            If LVBewerber.ListItems(x).Text = rsBewerber.Fields("Bewerber").Value Then
                If rsBewerber.Fields("Zusage").Value = "Ja" Then
                    LVBewerber.ListItems(x).Checked = True
                Else
                    LVBewerber.ListItems(x).Checked = False
                End If
            End If
        Next x
        rsBewerber.MoveNext
    Next i
    
exithandler:

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "RefreshFrameBewerber", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub RefreshFrameDokumente(Optional bVisible As Boolean)

On Error GoTo Errorhandler

    LVDokumente.Tag = "Ausgeschriebene Stellen\StellenJahr\*\Dokumente"
    Set rsDokumente = RefreshFrame(Me, FrameDokumente, LVDokumente, "Ausgeschriebene Stellen", "Dokumente", bVisible)

exithandler:

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "RefreshFrameDokumente", errNr, errDesc)
    Resume exithandler
End Sub

Public Sub RefreshFrameWorkflow(Optional bVisible As Boolean)
' Initialisiert und Aktualisiert den Workflow Frame
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    With ThisFramePos
        Call FrameWorkflow.Move(.Left, .Top, .Width, .Height)       ' Frame Positionieren
    End With
    FrameWorkflow.Visible = bVisible                                ' Frame ausblenden
    'CurrentStep = GetWorkflowCurrentStep(Me, "WORKFLOW012", lngWorkflowLevel)  ' Aktuellen Step ermitteln
    CurrentStep = GetWorkflowCurrentStep(Me, "WORKFLOW012")  ' Aktuellen Step ermitteln
'    NextStep = GetWorkflowNextStep(Me, lngWorkflowLevel)            ' Nächsten Schritt ermitteln
'    PrevStep = GetWorkflowPreStep(Me, lngWorkflowLevel)             ' Voherigen Schritt ermitten
    NextStep = GetWorkflowNextStep(Me)            ' Nächsten Schritt ermitteln
    PrevStep = GetWorkflowPreStep(Me)             ' Voherigen Schritt ermitten
    Call ShowWorkflowSteps(Me, LVStep)            ' LV Hauptschritte initialisieren
    'Call ShowWorkflowSteps(Me, LVStep, lngWorkflowLevel)            ' LV Hauptschritte initialisieren
    'Call ShowWorkflowSubSteps(Me, LVSubStep, , lngWorkflowLevel)    ' LV Teilschritte initialisieren
    Call SetWorkflowDescription(Me)                                 ' Teilschritt beschreibung anzeigen
    Call CheckWorkflowButtons(Me, CurrentStep, NextStep, PrevStep)  ' Prüffen ob Buttons Enabled
  
exithandler:

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "RefreshFrameWorkflow", errNr, errDesc)
    Resume exithandler
End Sub

Private Function RefreshRelFields()

On Error GoTo Errorhandler

    Call RefreshRelField(Me, cmbAusschreibung, txtIDAus, _
            "SELECT TOP 1 AZ020 + ' (' + CAST(JAHR020 as varchar(4)) + ')' FROM AUSSCHREIBUNG020", _
            "ID020 =", True)
'    Call RefreshRelField(Me, txtBewerber, txtIDBewerber, _
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

Private Function RefreshIDFields()

On Error GoTo Errorhandler

    Call RefreshRelField(Me, txtIDAus, cmbAusschreibung, _
            "SELECT TOP 1 ID020 FROM AUSSCHREIBUNG020", _
            "AZ020 + ' (' + CAST(JAHR020 as varchar(4)) + ')' =", True)
'    Call RefreshRelField(Me, txtBewerber, txtIDBewerber, _
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

Private Sub ShowKontextMenu(Menuename As String)
    ' Zeigt das Menü mit MenueName als Kontext (Popup) Menü an
On Error Resume Next                                                ' Fehlerbehandlung deaktiviern

    Select Case Menuename
    Case "kmnuLVBewerber"
        PopupMenu kmnuLVBewerber
        
    Case "kmnuLVDocument"
        PopupMenu kmnuLVDocument
    Case ""
    
    Case Else
    
    End Select
    
End Sub

Private Sub HandleMenueKlick(szMenueName As String, Optional szCaption As String)
' Behandlet Kontextmenü klicks in den Listviews

    Dim szItemKeyArray() As String                                  ' Key Array eines List items
    Dim szID As String                                              ' DS ID eines ListItems
    Dim AusschrID As String                                         ' ID der Ausschreibung zu dieser Stelle
    
On Error GoTo Errorhandler
    
    If HandleLVkmnuNew(Me, szCaption) Then GoTo exithandler
    
    Select Case szMenueName
    Case "kmnuLVBewerberAdd"                                        ' Bewerbung hinzufügen
        Call EditDS("Bewerbung", ";" & ID & ";", True)              ' DS Bearbeiten
        'Call frmParent.OpenEditForm("Bewerbung", ";" & ID & ";", frmParent, True)
        Call RefreshFrameBewerber(True)                             ' LV Bewerber Aktualisieren

    Case "kmnuLVBewerberDel"                                        ' Bewerbung löschen
        Call DelRelationinLV(Me, "Bewerbung", ThisDBCon, LVBewerber, rsBewerber, "ID013", "BEWERB013")
        Call RefreshFrameBewerber(True)                             ' LV Bewerber Aktualisieren
    
    Case "kmnuLVBewerberOpen"                                       ' Bewerbung Bearbeiten
        Call HandleEditLVDoubleClick(Me, LVBewerber, frmParent)
        
    Case "kmnuLVBewerberPersonOpen"                                 ' Personen Daten der Bewerbung anzeigen
    On Error Resume Next
        szID = LVBewerber.SelectedItem.SubItems(6)
        Err.Clear
    On Error GoTo Errorhandler
        If szID <> "" Then
            Call EditDS("Personen", szID)                           ' DS Bearbeiten
            'Call frmParent.OpenEditForm("Personen", szID, frmParent)
        End If
    Case "kmnuLVDokumentAdd"                                        ' Neues Dokument
        Call GetIDCollection(Me, "", "", AusschrID)                 ' ID für Auschreibung ermitteln
        Call WriteWord("", "", ID, AusschrID)                       ' SAT Aufrufen
        Call RefreshFrameDokumente(True)                            ' LV Dokumente Aktualisieren
    
    Case "kmnuLVDokumentOpen"                                       ' Dokument anzeigen
        Call HandleEditLVDoubleClick(Me, LVDokumente)
    
    Case "kmnuLVDokumentImport"
        Call GetIDCollection(Me, "", "", AusschrID)                 ' ID für Auschreibung ermitteln
        Call ImportWordDoc(ThisDBCon, "", ID, AusschrID)            ' Dok importieren
        
    Case "kmnuLVDokumentDel"                                        ' Dokument löschen
        szID = GetRelLVSelectedID(LVDokumente)                      ' Doc ID Aus LV ermitteln
        If szID <> "" Then
            Call DeleteDS("Dokumente", szID)                         'DS Löschen
            Call RefreshFrameDokumente(True)                        ' LV Dokumente Aktualisieren
        End If
        
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
On Error Resume Next                                                ' Fehlerbehandlung deaktiviern
    Call HandleKeyDownEdit(Me, KeyCode, Shift)                      ' Spezielle KeyDon Events dieses Forms
    Call frmParent.HandleGlobalKeyCodes(KeyCode, Shift)             ' Key Down Events der Anwendung
    Err.Clear
End Sub

Private Sub HandleTabClick(TS As TabStrip)
' Behandelt Tab Klicks

On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    Call HandleTabClickNew(Me, TS)                                  ' Wenn bNew dan nur 1. Tab zulassen
    
    If TS.SelectedItem = "Info" Then                                ' InfoTabklick behandeln
        FrameBewerber.Visible = False
        FrameDokumente.Visible = False
        FrameWorkflow.Visible = False
        FrameInfo.Visible = True
        Exit Sub
    End If
    
    Select Case TS.SelectedItem.Index
    Case 1                                                          ' Bewerbungen
        FrameBewerber.Visible = True
        FrameDokumente.Visible = False
        FrameWorkflow.Visible = False
        FrameInfo.Visible = False
    Case 2                                                          ' Dokumente
        FrameBewerber.Visible = False
        FrameDokumente.Visible = True
        FrameWorkflow.Visible = False
        FrameInfo.Visible = False
    Case 3                                                          ' Workflow
        FrameBewerber.Visible = False
        FrameDokumente.Visible = False
        FrameWorkflow.Visible = True
        FrameInfo.Visible = False
    Case 4                                                          ' Info
    
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
    
    bValidationFaild = ValidateTxtFieldOnEmpty(txtFrist, "Bewerbungsfrist", _
            szMSG, FocusCTL)                                        ' Bewerbungsfrist auf Leer prüfen
    
    bValidationFaild = ValidateTxtFieldOnEmpty(txtAZ, "Aktenzeichen", _
            szMSG, FocusCTL)                                        ' Aktenzeichen auf Leer prüfen
    
    bValidationFaild = ValidateTxtFieldOnEmpty(txtAnzStellen, "Anz. Stellen", _
            szMSG, FocusCTL)                                        ' Anz. Stellen auf Leer prüfen
    
    bValidationFaild = ValidateTxtFieldOnEmpty(cmbBezirk, "Bezirk", _
            szMSG, FocusCTL)                                        ' Bezirk auf Leer prüfen
    
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
    
    bNewBeforSave = bNew
    
    If Not ValidateEditForm Then GoTo exithandler                   ' Eingaben Validieren
        
    If UpdateEditForm(Me, szRootkey) Then                           ' Speichern
        bNew = False                                                ' nicht mehr neu
        If bNewBeforSave Then
            If txtAnzStellen.Text = "" Then txtAnzStellen.Text = "1"
'            szSQL = "SELECT ID012 FROM STELLEN012 WHERE Bezirk012='" & cmbBezirk & "' AND Az012 ='" & txtAZ _
'                    & "' AND ANZ012 = " & txtAnzStellen & " AND DATEDIFF(n ,create012 , '" & txtCreate & "')< 2"
'            szID = ThisDBcon.GetValueFromSQL(szSQL)
'            If szID <> "" Then Call InitEditForm(frmParent, ThisDBcon, szID, bModal)
        End If
        Call HiglightThisMustFields(True)                           ' Hervorhebung abschalten
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
    Call RefreshRelFields                                           ' Relations felde raktualisieren
    Call RefreshFrameBewerber(FrameBewerber.Visible)                ' Frame Bewerber aktualisieren
    Call RefreshFrameDokumente(FrameDokumente.Visible)              ' Frame Dokumente aktualisieren
    
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

Public Sub SelectTabByName(szTabName As String)

    Dim i As Integer                                                ' Counter
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    For i = 1 To TabStrip1.Tabs.Count
        If UCase(TabStrip1.Tabs(i).Caption) = UCase(szTabName) Then
             TabStrip1.Tabs(i).Selected = True
        End If
        Exit Sub
    Next i
exithandler:
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SelectTabByName", errNr, errDesc)
    Resume exithandler
End Sub
                                                                    ' *****************************************
                                                                    ' TabSrip Events
Private Sub TabStrip1_Click()
    Call HandleTabClick(TabStrip1)                                  ' Tab Klick handlen
End Sub
                                                                    ' *****************************************
                                                                    ' Button Events
Private Sub cmdNextStep_Click()
    Call WorkflowNextStep(Me, "WORKFLOW012", True)                  ' Nächsten WorkflowSchritt
End Sub

Private Sub cmdEsc_Click()
    Unload Me                                                       ' Bearbeiten abbrechen
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

    Dim szID As String
    
On Error Resume Next
        szID = LVBewerber.SelectedItem.SubItems(6)
        Err.Clear
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    Call WriteWord("", szID, ID, AusschreibungsID)                      ' SAT aufrufen
    Call RefreshFrameDokumente(FrameDokumente.Visible)              ' LV Dokumente Aktualisieren
        
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

Private Sub txtCreateFrom_Click()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If txtModifyFrom.Tag <> "" Then                                 ' Tag vorhanden
        Call AskUserAboutThisDS(txtCreateFrom, "Wegen Stelle in " _
                & cmbBezirk)                                        ' Email an User vorbereiten
    End If
    Err.Clear                                                       ' Evtl. Error Clearen
End Sub

Private Sub txtModifyFrom_Click()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If txtModifyFrom.Tag <> "" Then                                 ' Tag vorhanden
        Call AskUserAboutThisDS(txtModifyFrom, "Wegen Stelle in " _
                & cmbBezirk)                                        ' Email an User vorbereiten
    End If
    Err.Clear                                                       ' Evtl. Error Clearen
End Sub
                                                                    ' *****************************************
                                                                    ' Mouse Events
Private Sub txtModifyFrom_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If txtModifyFrom.Tag <> "" Then                                 ' Tag vorhanden
        Call MousePointerLink(Me, txtModifyFrom)                    ' Mouspointer Hyperlink setzen
    End If
    Err.Clear                                                       ' Evtl. Error Clearen
End Sub

Private Sub txtCreateFrom_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If txtCreateFrom.Tag <> "" Then                                 ' Tag vorhanden
        Call MousePointerLink(Me, txtCreateFrom)                    ' Mouspointer Hyperlink setzen
    End If
    Err.Clear                                                       ' Evtl. Error Clearen
End Sub

Private Sub LVStep_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim StepArray() As String
    
On Error Resume Next                                                ' Fehlerbehandlung deaktiviern
    'Call ShowWorkflowSteps(Me, LVStep)                              ' LV Hauptschritte initialisieren
'    Call ShowWorkflowSubSteps(Me, LVSubStep, Right(LVStep.SelectedItem.Key, 2), lngWorkflowLevel)  ' LB Teilschritte aktualisiern
'    Call SelectLVItem(LVSubStep, LVSubStep.ListItems(0).Key)
'    StepArray = Split(CurrentStep, ".")
'
'    Call SetWorkflowDescription(Me, Right(LVStep.SelectedItem.Key, 2), Right(LVSubStep.SelectedItem.Key, 2))                                  ' Teilschritt beschreibung anzeigen
Err.Clear

End Sub

Private Sub LVSubStep_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'Call SetWorkflowDescription(Me, Right(LVStep.SelectedItem.Key, 2), Right(LVSubStep.SelectedItem.Key, 2))                                  ' Teilschritt beschreibung anzeigen
End Sub

Private Sub LVBewerber_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then Call ShowKontextMenu("kmnuLVBewerber")
End Sub

Private Sub LVDokumente_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then Call ShowKontextMenu("kmnuLVDocument")
End Sub
                                                                    ' *****************************************
                                                                    ' Menue Events
Private Sub kmnuLVBewerberAdd_Click()
    Call HandleMenueKlick("kmnuLVBewerberAdd")                      ' KontextMenüKlick im LV Bewerber behandeln
End Sub

Private Sub kmnuLVBewerberDel_Click()
    Call HandleMenueKlick("kmnuLVBewerberDel")                      ' KontextMenüKlick im LV Bewerber behandeln
End Sub

Private Sub kmnuLVBewerberOpen_Click()
    Call HandleMenueKlick("kmnuLVBewerberOpen")                     ' KontextMenüKlick im LV Bewerber behandeln
End Sub

Private Sub kmnuLVBewerberPersonOpen_Click()
    Call HandleMenueKlick("kmnuLVBewerberPersonOpen")               ' KontextMenüKlick im LV Bewerber behandeln
End Sub

Private Sub kmnuLVDokumentAdd_Click()
    Call HandleMenueKlick("kmnuLVDokumentAdd")                      ' KontextMenüKlick im LV Dokumente behandeln
End Sub

Private Sub kmnuLVDokumentOpen_Click()
    Call HandleMenueKlick("kmnuLVDokumentOpen")                     ' KontextMenüKlick im LV Dokumente behandeln
End Sub

Private Sub kmnuLVDokumentDel_Click()
    Call HandleMenueKlick("kmnuLVDokumentDel")                      ' KontextMenüKlick im LV Dokumente behandeln
End Sub

Private Sub kmnuLVDokumentImport_Click()
    Call HandleMenueKlick("kmnuLVDokumentImport")                   ' KontextMenüKlick im LV Dokumente behandeln
End Sub
                                                                    ' *****************************************
                                                                    ' Change Events
Private Sub cmbAusschreibung_DropDown()                             ' Liste ausklappen
    If bInit Then Exit Sub                                          ' Wenn Form Initialisiert wird -> fertig
    OldCmbValue = cmbAusschreibung.Text                             ' Auswahl beginnt
End Sub

Private Sub cmbAusschreibung_Click()                                ' Änderung duch Liste auswahl
    If bInit Then Exit Sub                                          ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> "" And OldCmbValue <> cmbAusschreibung Then bDirty = True ' Dirty Nur wenn combo <> oldValue
    Call RefreshIDFields
    Call CheckUpdate(Me)
End Sub

Private Sub cmbAusschreibung_Change()                               ' Änderung duch Texteingabe
    If bInit Then Exit Sub                                          ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> "" And OldCmbValue <> cmbAusschreibung Then bDirty = True ' Dirty Nur wenn combo <> oldValue
    Call RefreshIDFields
    Call CheckUpdate(Me)
End Sub

Private Sub cmbBezirk_Click()
    If bInit Then Exit Sub                                          ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> "" And OldCmbValue <> cmbBezirk Then bDirty = True ' Dirty Nur wenn combo <> oldValue
    OldCmbValue = ""
    'Call RefreshIDFields
    Call CheckUpdate(Me)
End Sub

Private Sub cmbBezirk_DropDown()
    If bInit Then Exit Sub                                          ' Wenn Form Initialisiert wird -> fertig
    OldCmbValue = cmbBezirk.Text                                    ' Auswahl beginnt
End Sub

Private Sub cmbBezirk_Change()
If bInit Then Exit Sub                                              ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> "" And OldCmbValue <> cmbBezirk Then bDirty = True ' Dirty Nur wenn combo <> oldValue
    'Call RefreshIDFields
    Call CheckUpdate(Me)
End Sub

Private Sub cmbBezirk_Validate(Cancel As Boolean)
    If bInit Then Exit Sub                                          ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> "" And OldCmbValue <> cmbBezirk Then bDirty = True  ' Dirty Nur wenn combo <> oldValue
    Call RefreshIDFields
    Call CheckUpdate(Me)
End Sub

Private Sub txtFrist_Change()
    On Error Resume Next
    If Len(txtFrist.Text) < 10 Then Exit Sub
    If IsDate(txtFrist.Text) Then
        DTFrist.Value = txtFrist.Text
        If bInit Then Exit Sub
        bDirty = True
        Call CheckUpdate(Me)
    End If
End Sub

Private Sub DTFrist_Change()
On Error Resume Next
    txtFrist.Text = DTFrist.Value
    If bInit Then Exit Sub
    Me.Adodc1.Recordset.Fields(txtFrist.DataField).Value _
            = Format(Me.DTFrist.Value, "dd.mm.yyyy")                ' Bei DTPicker Wert nochmal übernehmen da Sonst die Datenbindung nicht funzt
End Sub

Private Sub txtAnzStellen_Change()
    If Not bInit Then Call StandartTextChange(Me, Me.txtAnzStellen)
End Sub

Private Sub txtAZ_Change()
    If Not bInit Then Call StandartTextChange(Me, Me.txtAZ)
End Sub

Private Sub txtBeschreibung_Change()
    If Not bInit Then Call StandartTextChange(Me, Me.txtBeschreibung)
End Sub
                                                                    ' *****************************************
                                                                    ' Fokus Events
Private Sub cmbAusschreibung_GotFocus()
    Call HiglightCurentField(Me, cmbAusschreibung, False)
End Sub

Private Sub cmbAusschreibung_LostFocus()
    Call HiglightCurentField(Me, cmbAusschreibung, True)
End Sub

Private Sub cmbBezirk_GotFocus()
    Call HiglightCurentField(Me, cmbBezirk, False)
End Sub

Private Sub cmbBezirk_LostFocus()
    Call HiglightCurentField(Me, cmbBezirk, True)
End Sub

Private Sub txtAnzStellen_GotFocus()
    Call HiglightCurentField(Me, txtAnzStellen, False)
End Sub

Private Sub txtAnzStellen_LostFocus()
    Call HiglightCurentField(Me, txtAnzStellen, True)
End Sub

Private Sub txtAZ_GotFocus()
    Call HiglightCurentField(Me, txtAZ, False)
End Sub

Private Sub txtAZ_LostFocus()
    Call HiglightCurentField(Me, txtAZ, True)
End Sub

Private Sub txtBeschreibung_GotFocus()
    Call HiglightCurentField(Me, txtBeschreibung, False)
End Sub

Private Sub txtBeschreibung_LostFocus()
    Call HiglightCurentField(Me, txtBeschreibung, True)
End Sub

Private Sub txtFrist_GotFocus()
    Call HiglightCurentField(Me, txtFrist, False)
End Sub

Private Sub txtFrist_LostFocus()
    Call HiglightCurentField(Me, txtFrist, True)
End Sub
                                                                    ' *****************************************
                                                                    ' Key Events
Private Sub txtFrist_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub LVBewerber_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub LVDokumente_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub DTFrist_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub cmbBezirk_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtAnzStellen_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtAZ_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtBeschreibung_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub
                                                                    ' *****************************************
                                                                    ' List View Events

Private Sub LVBewerber_GotFocus()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    'Call RefreshFrameBewerber(True)
    Err.Clear                                                       ' Evt. Error clearen
End Sub

Private Sub LVDokumente_GotFocus()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
'    Call RefreshFrameDokumente(True)
    Err.Clear                                                       ' Evt. Error clearen
End Sub

Private Sub LVBewerber_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    
    If rsBewerber.RecordCount = 0 Then Exit Sub
    rsBewerber.MoveFirst
    For i = 0 To rsBewerber.RecordCount
            If Item.Text = rsBewerber.Fields("Bewerber").Value Then
                rsBewerber.Fields("Zusage013").Value = Item.Checked
                rsBewerber.Update
                Exit For
'                If rsBewerber.Fields("Zusage").Value = "Ja" Then
'                    LVBewerber.ListItems(x).Checked = True
'                Else
'                    LVBewerber.ListItems(x).Checked = False
'                End If
            End If
            rsBewerber.MoveNext
        
    Next i
End Sub

Private Sub LVBewerber_DblClick()
    Call HandleEditLVDoubleClick(Me, Me.LVBewerber, frmParent)
End Sub

Private Sub LVBewerber_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SetColumnOrder(LVBewerber, ColumnHeader)
End Sub

Private Sub LVDokumente_DblClick()
    Call HandleEditLVDoubleClick(Me, Me.LVDokumente, frmParent)
End Sub

Private Sub LVDokumente_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SetColumnOrder(LVDokumente, ColumnHeader)
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

Public Property Get GetCurrentStep() As Variant
    GetCurrentStep = CurrentStep
End Property

Public Property Get GetNextStep() As String
    GetNextStep = NextStep
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

