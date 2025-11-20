VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmEditAusschreibung 
   Caption         =   "Ausschreibung"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   7005
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame FrameWorkflow 
      Caption         =   "Vorgang"
      Height          =   4815
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Width           =   6735
      Begin MSComctlLib.ListView LVStep 
         Height          =   3015
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
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
      Begin VB.CommandButton cmdPrevStep 
         Caption         =   "<< Zurück"
         Height          =   375
         Left            =   4080
         TabIndex        =   30
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdNextStep 
         Caption         =   "Weiter >>"
         Height          =   375
         Left            =   5400
         TabIndex        =   29
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label lblStepDesc 
         Caption         =   "Schritt Beschreibung"
         Height          =   855
         Left            =   120
         TabIndex        =   32
         Top             =   3360
         Width           =   6255
      End
   End
   Begin VB.Frame FrameDokumente 
      Caption         =   "Dokumente"
      Height          =   735
      Left            =   120
      TabIndex        =   26
      Top             =   2520
      Width           =   2055
      Begin MSComctlLib.ListView LVDokumente 
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   0
      Picture         =   "frmAusschreibungEdit.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Datensatz speichern"
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   375
      Left            =   480
      Picture         =   "frmAusschreibungEdit.frx":058A
      Style           =   1  'Grafisch
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Datensatz löschen"
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton cmdWord 
      Height          =   375
      Left            =   960
      Picture         =   "frmAusschreibungEdit.frx":0914
      Style           =   1  'Grafisch
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Neues Anschreiben"
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Übernehmen"
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox txtJahr 
      DataField       =   "JAHR020"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Frame FrameInfo 
      Caption         =   "Datensatz Informationen"
      Height          =   2535
      Left            =   2040
      TabIndex        =   11
      Top             =   1080
      Width           =   6735
      Begin VB.TextBox txtID 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "ID020"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   5295
      End
      Begin VB.TextBox txtCreateFrom 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "CFROM020"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   720
         Width           =   5295
      End
      Begin VB.TextBox txtCreate 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "CREATE020"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Width           =   5295
      End
      Begin VB.TextBox txtModifyFrom 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "MFROM020"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1440
         Width           =   5295
      End
      Begin VB.TextBox txtModify 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "MODIFY020"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1200
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
   Begin VB.Frame FrameBewerbungen 
      Caption         =   "Bewerbungen"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
      Begin MSComctlLib.ListView LVBewerbungen 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
   Begin VB.Frame FrameStellen 
      Caption         =   "Ausgeschriebene Stellen"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   2055
      Begin MSComctlLib.ListView LVStellen 
         Height          =   375
         Left            =   120
         TabIndex        =   4
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
   Begin VB.TextBox txtAZ 
      DataField       =   "AZ020"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5295
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9340
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Stellen"
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
   Begin MSComctlLib.ImageList ILTree 
      Left            =   1680
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
            Picture         =   "frmAusschreibungEdit.frx":0C9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":0FB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":12D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":186C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":1E06
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":2AE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":37BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":3D54
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":42EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":4688
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":4A22
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":26554
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":48086
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":48620
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":48BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":49154
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":496EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":49C88
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":4A222
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":4A7BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":4B866
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":4BE00
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":4C39A
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":4C934
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":4CECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":4D268
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":4D802
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusschreibungEdit.frx":4DD9C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblJahr 
      Caption         =   "Jahr"
      Height          =   255
      Left            =   4080
      TabIndex        =   22
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblAz 
      Caption         =   "Aktenzeichen"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu kmnuLVStellen 
      Caption         =   "KontextmenuLVStellen"
      Visible         =   0   'False
      Begin VB.Menu kmnuLVStellenNew 
         Caption         =   "Neue Stelle ausschreiben"
      End
      Begin VB.Menu kmnuLVStellenDel 
         Caption         =   "Ausgeschreibene Stelle löschen"
      End
      Begin VB.Menu kmnuLVStellenEdit 
         Caption         =   "Ausgeschriebene Stelle bearbeiten"
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
      Begin VB.Menu kmnuDokumentImport 
         Caption         =   "Dokument Importieren"
      End
   End
End
Attribute VB_Name = "frmEditAusschreibung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MODULNAME = "frmEditAusschreibung"                    ' Modulname für Fehlerbehandlung

Private bInit As Boolean                                            ' Wird True gesetzt wenn Alle werte geladen
Private bDirty As Boolean                                           ' Wird True gesetzt wenn Daten verändert wurden
Private bNew  As Boolean                                            ' Wird gesetzt wenn neuer DS sonst Update
Private bModal As Boolean                                           ' Ist Modal Geöffnet
Private szID As String                                              ' DS ID
Private ThisDBcon As Object                                         ' Aktuelle DB Verbindung
Private frmParent As Form                                           ' Aufrufendes DB form
Private szIDField As String
Private ThisFramePos As FramePos                                    ' Standart Frame Position

Private szSQL As String                                             ' SQL für Ausschreibung020
Private szWhere As String                                           ' Where Klausel
Private szIniFilePath As String                                     ' Pfad der Ini datei
Private lngImage As Integer                                         ' Imagiendex

Private szRootkey As String                                         ' = Ausschreibung
Private szDetailKey As String                                       ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)

Private CurrentStep As String                                       ' Aktueller WorkflowSchritt
Private NextStep As String                                          ' Nächster Schritt
Private PrevStep As String                                          ' Voheriger Schritt
Private lngWorkflowLevel As Integer                                 ' Workflow Ebene

Private OldCmbValue As String                                        'Alter Combo wert

'Private rsBewerber As ADODB.Recordset                               ' RS mit bewerber Daten
Private rsStellen As ADODB.Recordset                                ' RS mit Ausgeschriebenen Stellen
Private rsDokumente As ADODB.Recordset                              ' RS mit Dokumenten

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
    Call SaveColumnWidth(LVStellen, szRootkey & "LV", True)         ' Spaltenbreiten für LV Speichern
    Call SaveColumnWidth(LVDokumente, szRootkey & "LV", True)       ' Spaltenbreiten für LV Speichern
    rsStellen.Close                                                 ' RS Stellen schliessen
    rsDokumente.Close                                               ' RS Dokumente schliessen
    If bDirty Then szID = ""                                        ' Wenn ungespeichert ID Leeren
    If bModal Then                                                  ' Wenn Modal
        Me.Hide                                                     ' dann ausblenden
    Else                                                            ' Sonst
        Call EditFormUnload(Me)                                     ' AUs Edit Form Array entfernen
    End If
    Err.Clear                                                       ' Evtl. Error clearen                                                  ' Evtl Err Clearen
End Sub


Public Function InitEditForm(parentform As Form, dbCon As Object, DetailKey As String, _
        Optional bDialog As Boolean)

    Dim i As Integer                                                ' counter
    'Dim tmpArray() As String
        
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    Set frmParent = parentform
    bInit = True                                                    ' Wir initialisieren das Form
                                                                    ' -> andere vorgänge nicht ausführen
    Set ThisDBcon = dbCon                                           ' Aktuelle DB Verbindung übernehmen
    szRootkey = "Ausschreibung"                                     ' für Caption
    szIDField = "ID020"
    
    'lngWorkflowLevel = 1                                            ' Für Workflow
    szDetailKey = DetailKey                                         ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)
    bModal = bDialog                                                ' Als Dialog anzeigen
    
    szIniFilePath = objObjectBag.Getappdir & objObjectBag.GetXMLFile ' XML inifile festlegen
    Call objTools.GetEditInfoFromXML(szIniFilePath, szRootkey, szSQL, szWhere, lngImage)
    
    Me.Icon = ILTree.ListImages(lngImage).Picture                   ' Form Icon Setzen
    
    If szDetailKey = "" Then bNew = True                            ' Neuer Datensatz
    If szDetailKey <> "" Then
        szID = DetailKey
        szWhere = szWhere & "CAST('" & szDetailKey & "' as uniqueidentifier)"
    End If
    
    Call InitAdoDC(Me, ThisDBcon, szSQL, szWhere)                   ' ADODC Initialisieren
    
    If bNew Then                                                    ' Wenn DS neu
        Adodc1.Recordset.AddNew                                     ' Neuen DS an RS anhängen
        txtID.Text = ThisDBcon.GetValueFromSQL("SELECT NewID()")    ' Neue ID (Guid) ermitteln
        szID = txtID
        bDirty = True                                               ' Dirty da Neu
    Else
        szID = Me.txtID
    End If
    
    Call GetLockedControls(Me)                                      ' Gelockte controls finden
    Call HiglightThisMustFields(Not (bNew Or bDirty))               ' IndexFelder hervorheben
    Call InitFrameInfo(Me)                                          ' Info Frame initialisieren
    Call RefreshFrameStellen(True)                                  ' Frame Stellen initialisieren
    Call RefreshFrameDokumente(False)                               ' Frame Dokumente initialisieren
    Call RefreshFrameWorkflow(False)                                ' Frame Workflow initialisieren
    Call SetEditFormCaption(Me, szRootkey, txtAz & " (" & txtJahr & ")") ' Form Caption setzen
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
On Error Resume Next                                                ' Fehlerbehandlung deaktiviern
    Call HiglightMustFields(Me, bDeHiglight)                        ' Alle PK ind IndexFields entfärben
    Err.Clear
End Sub

Private Sub RefreshFrameStellen(Optional bVisible As Boolean)

On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    LVStellen.Tag = "Ausschreibungen\StellenJahr"                   ' Tag setzen
    Set rsStellen = RefreshFrame(Me, FrameStellen, LVStellen, "Ausschreibung", "Ausgeschriebene Stellen", bVisible)
    
exithandler:

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "RefreshFrameStellen", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub RefreshFrameDokumente(Optional bVisible As Boolean)

On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    LVDokumente.Tag = "Ausschreibungen\Dokumente"                   ' Tag setzen
    Set rsDokumente = RefreshFrame(Me, FrameDokumente, LVDokumente, "Ausschreibung", "Dokumente", bVisible)
    
exithandler:

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "RefreshFrameStellen", errNr, errDesc)
    Resume exithandler
End Sub

Public Sub RefreshFrameWorkflow(Optional bVisible As Boolean)
' Initialisiert und Aktualisiert den Workflow Frame

On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

    FrameWorkflow.Visible = False
    Exit Sub
    With ThisFramePos
        Call FrameWorkflow.Move(.Left, .Top, .Width, .Height)       ' Frame Positionieren
    End With
    FrameWorkflow.Visible = bVisible                                ' Frame ausblenden
    'CurrentStep = GetWorkflowCurrentStep(Me, "WORKFLOW020", lngWorkflowLevel)  ' Aktuellen Step ermitteln
    CurrentStep = GetWorkflowCurrentStep(Me, "WORKFLOW020")  ' Aktuellen Step ermitteln
    NextStep = GetWorkflowNextStep(Me)            ' Nächsten Schritt ermitteln
    PrevStep = GetWorkflowPreStep(Me)             ' Voherigen Schritt ermitten
    
    'NextStep = GetWorkflowNextStep(Me, lngWorkflowLevel)            ' Nächsten Schritt ermitteln
    'PrevStep = GetWorkflowPreStep(Me, lngWorkflowLevel)             ' Voherigen Schritt ermitten
'    If bInit Then
'        If CurrentStep = GetWorkflowMinStep(Me, lngWorkflowLevel) Then
'            Call WorkflowNextStep(Me, "WORKFLOW020", CurrentStep, NextStep)  ' Nächsten WorkflowSchritt
'        End If
'    End If
'    Call ShowWorkflowSteps(Me, LVStep, lngWorkflowLevel)            ' LV Hauptschritte initialisieren
    Call ShowWorkflowSteps(Me, LVStep)            ' LV Hauptschritte initialisieren
'    Call ShowWorkflowSubSteps(Me, LVSubStep, , lngWorkflowLevel)    ' LV Teilschritte initialisieren
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

Private Sub ShowKontextMenu(Menuename As String)
    ' Zeigt das Menü mit MenueName als Kontext (Popup) Menü an
On Error Resume Next                                                ' Fehlerbehandlung deaktiviern

    Select Case Menuename
    Case "kmnuLVBewerber"
        'PopupMenu kmnuLVBewerber
    Case "kmnuLVDocument"
        PopupMenu kmnuLVDocument                                    ' Menü kmnuLVDocument öffnen
    Case "kmnuLVStellen"
        PopupMenu kmnuLVStellen                                     ' Menü kmnuLVStellen öffnen
    Case ""
    
    Case Else
    
    End Select
    
End Sub

Private Sub HandleMenueKlick(szMenueName As String, Optional szCaption As String)
    
    Dim szItemKeyArray() As String                                  ' Key Array eines List items
    Dim szID As String                                              ' DS ID eines ListItems
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    If HandleLVkmnuNew(Me, szCaption) Then GoTo exithandler         ' Wenn DS Neu -> keine Aktion
    
    Select Case szMenueName
    Case "kmnuLVStellenNew"                                         ' Stelle hinzufügen
                                                                    ' StellenID;AusschribungsID;Jahr;Az als Detailkey übergeben
        Call EditDS("Stellen", ";" & ID & ";" & txtJahr & ";" & txtAz, True)
        'Call frmParent.OpenEditForm("Stellen", ";" & ID & ";" & txtJahr & ";" & txtAZ, frmParent, True)
        Call RefreshFrameStellen(True)                              ' ListView LVStellen aktualisieren
        
    Case "kmnuLVStellenDel"                                         ' Stelle löschen
        Call DelRelationinLV(Me, "Stellen", ThisDBcon, LVStellen, rsStellen, "ID012", "STELLEN012")
        Call RefreshFrameStellen(True)                              ' ListView LVStellen aktualisieren
    
    Case "kmnuLVStellenEdit"                                        ' Stelle Bearbeiten
        Call HandleEditLVDoubleClick(Me, LVStellen, frmParent)
    
    Case "kmnuLVDokumentAdd"                                        ' Neues Dokument zur Ausschreibung
        Call WriteWord("", "", "", ID)                              ' SAT Starten
        Call RefreshFrameDokumente(True)                            ' LV Dokumente aktualisieren
    
    Case "kmnuLVDokumentDel"                                        ' Dokument löschen
        szID = GetRelLVSelectedID(LVDokumente)                      ' Doc ID Aus LV ermitteln
'        szItemKeyArray = Split(LVDokumente.SelectedItem.Key, TV_KEY_SEP)
'        szID = szItemKeyArray(UBound(szItemKeyArray))
        If szID <> "" Then
            Call DeleteDS("Dokumente", szID)                        ' DS Löschen
            Call RefreshFrameDokumente(True)                        ' LV Dokumente aktualisieren
        End If
        
    Case "kmnuLVDokumentImport"                                     ' Dokument Importieren
        Call ImportWordDoc(ThisDBcon, "", "", ID)                   ' Dokument Importieren
        Call RefreshFrameDokumente(True)                            ' LV Dokumente aktualisieren
        
    Case "kmnuLVDokumentOpen"                                       ' Dokument öffnen
        Call HandleEditLVDoubleClick(Me, LVDokumente)
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
    
    If TS.SelectedItem = "Info" Then
        FrameStellen.Visible = False
        FrameBewerbungen.Visible = False
        FrameDokumente.Visible = False
        FrameWorkflow.Visible = False
        FrameInfo.Visible = True
        Exit Sub
    End If
    
    Select Case TS.SelectedItem.Index
    Case 1                                                          ' Stellen
        FrameStellen.Visible = True
        FrameBewerbungen.Visible = False
        FrameDokumente.Visible = False
        FrameWorkflow.Visible = False
        FrameInfo.Visible = False
    Case 2                                                          ' Dokumente
        FrameStellen.Visible = False
        FrameBewerbungen.Visible = False
        FrameDokumente.Visible = True
        FrameWorkflow.Visible = False
        FrameInfo.Visible = False
    Case 3                                                          ' Workflow
        FrameStellen.Visible = False
        FrameBewerbungen.Visible = False
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
    Dim szTmpValue
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    szTitle = "Unvollständige Daten"                                ' Meldungstitel setzen
    
    bValidationFaild = ValidateTxtFieldOnEmpty(txtAz, "Aktenzeichen", _
            szMSG, FocusCTL)                                        ' Aktenzeichen auf Leer prüfen
            
    bValidationFaild = ValidateTxtFieldOnEmpty(txtJahr, "Jahr", _
            szMSG, FocusCTL)                                        ' Jahr auf Leer prüfen
            
    If IsNew() Then
        szTmpValue = ThisDBcon.GetValueFromSQL("SELECT Jahr020 FROM AUSSCHREIBUNG020 WHERE Jahr020 = " & txtJahr.Text)
        If szTmpValue <> "" Then bValidationFaild = True
        szMSG = "Sie haben für das Jahr " & txtJahr.Text & " bereits eine Auschreibung angelegt." & vbCrLf _
            & "Es kann pro Jahr nur eine Ausschreibung geben."
        Set FocusCTL = txtJahr
    End If
    
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
    
    bNewBeforSave = bNew                                            ' Neuer Datensatz
    
    If Not ValidateEditForm Then GoTo exithandler                   ' Eingaben Validieren
        
    If UpdateEditForm(Me, szRootkey) Then                           ' Speichern
        bNew = False                                                ' nicht mehr neu
        Call HiglightThisMustFields(True)                           ' Hervorhebung von Indexfeldern abschalten
    End If
        
    'Call WorkflowNextStep(Me, "WORKFLOW020", True)
    
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
    Call RefreshFrameDokumente(FrameDokumente.Visible)              ' Frame Dokumente Refreschen
    Call RefreshFrameStellen(FrameStellen.Visible)                  ' Frame Stellen Refreshen
    
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

    Dim i As Integer
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    For i = 0 To TabStrip1.Tabs.Count - 1
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
    Call HandleTabClick(TabStrip1)                                  ' Tab Klick behandeln
End Sub
                                                                    ' *****************************************
                                                                    ' Button Events
Private Sub cmdNextStep_Click()
    Call WorkflowNextStep(Me, "WORKFLOW020", True)                  ' Nächsten WorkflowSchritt
End Sub

Private Sub cmdESC_Click()
On Error Resume Next                                                ' Fehlerbehandlung deaktiviern
    Unload Me
End Sub

Private Sub cmdOK_Click()

On Error GoTo Errorhandler

    If SaveEditForm Then                                            ' Dieses Form Speichern
        Call CheckUpdate(Me)                                        ' Evtl Übernehmen disablen
        Unload Me                                                   ' Form Schliessen
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

On Error GoTo Errorhandler

    Call WriteWord("", "", "", ID)                                  ' SAT aufrufen
    Call RefreshFrameDokumente(True)                                ' LV Dokumente Aktualisieren
        
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
        Call AskUserAboutThisDS(txtCreateFrom, "Wegen " _
                & txtAz & " (" & txtJahr & ")")                     ' Email an User vorbereiten
    End If
    Err.Clear                                                       ' Evtl. Error Clearen
End Sub

Private Sub txtModifyFrom_Click()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If txtModifyFrom.Tag <> "" Then                                 ' Tag vorhanden
        Call AskUserAboutThisDS(txtModifyFrom, "Wegen " _
                & txtAz & " (" & txtJahr & ")")                     ' Email an User vorbereiten
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
    'Call ShowWorkflowSubSteps(Me, LVSubStep, Right(LVStep.SelectedItem.Key, 2)) ' LB Teilschritte aktualisiern
    'Call SelectLVItem(LVSubStep, LVSubStep.ListItems(0).Key)
    StepArray = Split(CurrentStep, ".")
    
    'Call SetWorkflowDescription(Me, Right(LVStep.SelectedItem.Key, 2), Right(LVSubStep.SelectedItem.Key, 2))                                 ' Teilschritt beschreibung anzeigen
Err.Clear

End Sub

Private Sub LVSubStep_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'Call SetWorkflowDescription(Me, Right(LVStep.SelectedItem.Key, 2), Right(LVSubStep.SelectedItem.Key, 2))                                ' Teilschritt beschreibung anzeigen
End Sub

Private Sub LVDokumente_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then Call ShowKontextMenu("kmnuLVDocument")
End Sub

Private Sub LVBewerber_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then Call ShowKontextMenu("kmnuLVBewerber")
End Sub

Private Sub LVStellen_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then Call ShowKontextMenu("kmnuLVStellen")
End Sub
                                                                    ' *****************************************
                                                                    ' Menue Events
Private Sub kmnuLVStellenNew_Click()
    Call HandleMenueKlick("kmnuLVStellenNew")                       ' KontextMenüKlick im LV Stellen behandeln
End Sub

Private Sub kmnuLVStellenDel_Click()
     Call HandleMenueKlick("kmnuLVStellenDel")                      ' KontextMenüKlick im LV Stellen behandeln
End Sub

Private Sub kmnuLVStellenEdit_Click()
    Call HandleMenueKlick("kmnuLVStellenEdit")                      ' KontextMenüKlick im LV Stellen behandeln
End Sub

Private Sub kmnuLVDokumentAdd_Click()
    Call HandleMenueKlick("kmnuLVDokumentAdd")                      ' KontextMenüKlick im LV Dokumente behandeln
End Sub

Private Sub kmnuLVDokumentDel_Click()
    Call HandleMenueKlick("kmnuLVDokumentDel")                      ' KontextMenüKlick im LV Dokumente behandeln
End Sub

Private Sub kmnuLVDokumentImport_Click()
    Call HandleMenueKlick("kmnuLVDokumentImport")                   ' KontextMenüKlick im LV Dokumente behandeln
End Sub

Private Sub kmnuLVDokumentOpen_Click()
    Call HandleMenueKlick("kmnuLVDokumentOpen")                     ' KontextMenüKlick im LV Dokumente behandeln
End Sub
                                                                    ' *****************************************
                                                                    ' Change Events
Private Sub txtAZ_Change()
    If Not bInit Then Call StandartTextChange(Me, Me.txtAz)
End Sub

Private Sub txtJahr_Change()
    If Not bInit Then Call StandartTextChange(Me, Me.txtJahr)
End Sub

'Private Sub cmbBezirk_Validate(Cancel As Boolean)
'    If bInit Then Exit Sub
'    bDirty = True
'    Call CheckUpdate(Me)
'End Sub

'Private Sub txtFrist_Change()
'    On Error Resume Next
'    If Len(txtFrist.Text) < 10 Then Exit Sub
'    If IsDate(txtFrist.Text) Then
'        DTFrist.Value = txtFrist.Text
'        If bInit Then Exit Sub
'        bDirty = True
'        Call CheckUpdate(Me)
'    End If
'End Sub
'
'Private Sub DTFrist_Change()
'On Error Resume Next
'    txtFrist.Text = DTFrist.Value
'    If bInit Then Exit Sub
'    Me.Adodc1.Recordset.Fields(txtFrist.DataField).Value _
'            = Format(Me.DTFrist.Value, "dd.mm.yyyy")                ' Bei DTPicker Wert nochmal übernehmen da Sonst die Datenbindung nicht funzt
'End Sub
'
'Private Sub txtAnzStellen_Change()
'    If Not bInit Then Call StandartTextChange(Me, Me.txtAnzStellen)
'End Sub
'
'Private Sub txtAz_Change()
'    If Not bInit Then Call StandartTextChange(Me, Me.txtAZ)
'End Sub
'
'Private Sub txtBeschreibung_Change()
'    If Not bInit Then Call StandartTextChange(Me, Me.txtBeschreibung)
'End Sub
                                                                    ' *****************************************
                                                                    ' Fokus Events

Private Sub txtJahr_GotFocus()
    Call HiglightCurentField(Me, txtJahr, False)
End Sub

Private Sub txtJahr_LostFocus()
    Call HiglightCurentField(Me, txtJahr, True)
End Sub

Private Sub txtAZ_GotFocus()
    Call HiglightCurentField(Me, txtAz, False)
End Sub

Private Sub txtAZ_LostFocus()
    Call HiglightCurentField(Me, txtAz, True)
End Sub
                                                                    ' *****************************************
                                                                    ' Key Events
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub LVBewerbungen_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub LVStellen_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtAZ_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtJahr_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub
                                                                    ' *****************************************
                                                                    ' ListView Events
Private Sub LVDokumente_DblClick()
    Call HandleEditLVDoubleClick(Me, Me.LVDokumente, frmParent)
End Sub

Private Sub LVDokumente_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SetColumnOrder(LVDokumente, ColumnHeader)
End Sub

Private Sub LVStellen_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SetColumnOrder(LVStellen, ColumnHeader)
End Sub

Private Sub LVStellen_DblClick()
    Call HandleEditLVDoubleClick(Me, Me.LVStellen, frmParent)
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
    Set GetDBConn = ThisDBcon
End Property

Public Property Get GetXMLPath() As String
    GetXMLPath = szIniFilePath
End Property

Public Property Get GetCurrentStep() As String
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


