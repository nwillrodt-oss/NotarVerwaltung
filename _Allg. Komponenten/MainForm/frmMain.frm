VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   12795
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   12795
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.ImageList ILToolBar 
      Left            =   2640
      Top             =   4680
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
            Picture         =   "frmMain.frx":223E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23D70
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":458A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":673D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":88F06
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":892A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8963A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":899D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ILDisToolbar 
      Left            =   3480
      Top             =   4680
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
            Picture         =   "frmMain.frx":89D6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AB8A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CD3D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EEF04
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":110A36
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":110DD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11116A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":111504
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolbarMain 
      Align           =   1  'Oben ausrichten
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ILToolBar"
      DisabledImageList=   "ILDisToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "tbBack"
            Object.Tag             =   "tbBack"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "tbForward"
            Object.Tag             =   "tbForward"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbNew"
            Object.ToolTipText     =   "Neuer Datensatz"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbNewBewerber"
                  Text            =   "Neuer Bewerber"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbNewBewerbung"
                  Text            =   "Neue Bewerbung"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbNewStelle"
                  Text            =   "Neue Stelle"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbNewDokument"
                  Text            =   "Neues Dokument"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbRefresh"
            Object.ToolTipText     =   "Anzeige Aktualisieren"
            Object.Tag             =   "tbRefresh"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbDocNew"
            Object.Tag             =   "tbDocNew"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbSearch"
            Description     =   "Suchen"
            Object.ToolTipText     =   "Suchen"
            Object.Tag             =   "tbSearch"
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbSearchPerson"
                  Text            =   "Suche Person"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbSearchDoc"
                  Text            =   "Suche Dokument"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbHelp"
            Object.ToolTipText     =   "Hilfe zur Notarverwaltung"
            Object.Tag             =   "tbHelp"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbInfo"
            Object.Tag             =   "tbInfo"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicList 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   4335
      Left            =   3120
      ScaleHeight     =   4335
      ScaleWidth      =   5295
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5295
      Begin MSComctlLib.ListView LVMain 
         Height          =   3975
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   7011
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
   End
   Begin VB.PictureBox PicTree 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4575
      ScaleWidth      =   2535
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   2535
      Begin MSComctlLib.ImageList ILTree 
         Left            =   960
         Top             =   2520
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   22
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":11189E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":111BB8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":111ED2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":11246C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":112A06
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1136E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1143BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":114954
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":114EEE
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":115288
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":115622
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":137154
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":158C86
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":159220
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1597BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":159D54
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":15A2EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":15A888
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":15AE22
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":15B3BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":15C466
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":15CA00
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TVMain 
         Height          =   3735
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   6588
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   707
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ILTree"
         Appearance      =   1
      End
   End
   Begin MSComctlLib.StatusBar StatusBarMain 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6150
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Picture         =   "frmMain.frx":15CF9A
            Object.ToolTipText     =   "Heutiges Datum"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Picture         =   "frmMain.frx":15D534
            Object.ToolTipText     =   "Angemeldeter Benutzer"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Picture         =   "frmMain.frx":15DACE
            Object.ToolTipText     =   "Aktueller Datenbankserver"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   9816
            Object.ToolTipText     =   "Anzahl der angezeigten Datensätze"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuDatei 
      Caption         =   "&Datei"
      Begin VB.Menu mnuDateiNew 
         Caption         =   "&Neu"
         Begin VB.Menu mnuDateiNewBewerber 
            Caption         =   "Neuer Bewerbe&r"
         End
         Begin VB.Menu mnuDateiNewBewerbung 
            Caption         =   "Neue &Bewerbung"
         End
         Begin VB.Menu mnuDateiNewStelle 
            Caption         =   "Neue &Stelle"
         End
         Begin VB.Menu mnuDateiNewDoc 
            Caption         =   "Neues &Dokument"
         End
      End
      Begin VB.Menu mnuDateiSuchen 
         Caption         =   "&Suchen"
      End
      Begin VB.Menu mnuDateiExit 
         Caption         =   "&Beenden"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Bearbeiten"
      Begin VB.Menu mnuEditAusschreibungNew 
         Caption         =   "Neue Ausschreibung"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSelectTemplate 
         Caption         =   "&Schreiben erstellen"
      End
      Begin VB.Menu mnuEditWorkflow 
         Caption         =   "Vorgang starten/fortsetzen"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuExtras 
      Caption         =   "&Extras"
      Begin VB.Menu mnuExtrasOptions 
         Caption         =   "&Optionen"
      End
      Begin VB.Menu mnuExtrasChangePWD 
         Caption         =   "&Kennwort ändern"
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "&?"
      Begin VB.Menu mnuInfoAbout 
         Caption         =   "&Info"
      End
      Begin VB.Menu mnuInfoHelp 
         Caption         =   "Hilfe"
      End
   End
   Begin VB.Menu kmnuListDefault 
      Caption         =   "KontextListDefault"
      Visible         =   0   'False
      Begin VB.Menu kmnuListNew 
         Caption         =   "Neu"
      End
      Begin VB.Menu kmnuListEdit 
         Caption         =   "Bearbeiten"
      End
      Begin VB.Menu kmnuListDel 
         Caption         =   "Löschen"
      End
   End
   Begin VB.Menu kmnuListPersonen 
      Caption         =   "KontextPersonen"
      Visible         =   0   'False
      Begin VB.Menu kmnuListNewPerson 
         Caption         =   "Neue Person anlegen"
      End
      Begin VB.Menu kmnuListEditPerson 
         Caption         =   "Person Bearbeiten"
      End
      Begin VB.Menu kmnuListNewDocPerson 
         Caption         =   "Neues Anschreiben"
      End
      Begin VB.Menu kmnuListDelPerson 
         Caption         =   "Person Löschen"
      End
      Begin VB.Menu kmnuListSearchPerson 
         Caption         =   "Person Suchen"
      End
   End
   Begin VB.Menu kmnuListNotare 
      Caption         =   "KontextNotare"
      Visible         =   0   'False
      Begin VB.Menu kmnuListNewNotar 
         Caption         =   "Neuen Notar anlegen"
      End
      Begin VB.Menu kmnuListEditNotar 
         Caption         =   "Notar Bearbeiten"
      End
      Begin VB.Menu kmnuListNewDocNotar 
         Caption         =   "Neues Anschreiben"
      End
      Begin VB.Menu kmnuListDelNotar 
         Caption         =   "Notar Löschen"
      End
      Begin VB.Menu kmnuListSearchNotar 
         Caption         =   "Notar Suchen"
      End
   End
   Begin VB.Menu kmnuListBewerber 
      Caption         =   "KontextBewerber"
      Visible         =   0   'False
      Begin VB.Menu kmnuListNewBewerber 
         Caption         =   "Neuen Bewerber anlegen"
      End
      Begin VB.Menu kmnuListEditBewerber 
         Caption         =   "Bewerber Bearbeiten"
      End
      Begin VB.Menu kmnuListNewDocBewerber 
         Caption         =   "Neues Anschreiben"
      End
      Begin VB.Menu kmnuListDelBewerber 
         Caption         =   "Bewerber Löschen"
      End
      Begin VB.Menu kmnuListSearchBewerber 
         Caption         =   "Bewerber Suchen"
      End
   End
   Begin VB.Menu kmnuListUser 
      Caption         =   "KontextBenutzer"
      Visible         =   0   'False
      Begin VB.Menu kmnuListNewUser 
         Caption         =   "Neuen Benutzer anlegen"
      End
      Begin VB.Menu kmnuListEditUser 
         Caption         =   "Benutzer bearbeiten"
      End
      Begin VB.Menu kmnuListDelUser 
         Caption         =   "Benutzer Löschen"
      End
      Begin VB.Menu kmnuListChangePWD 
         Caption         =   "Kennwort zurücksetzen"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MODULNAME = "frmMain"

Private curlngSplitposProz As Single                                ' Aktueller wert der splitter pos.
Private SplitFlag As Boolean                                        ' True wenn List bzw. TreeView größe verändert wird
Private EditFormArray() As Object                                   ' Auflistung aller geöffneten Edit Formulare
Private NavTreeNodeArray() As String                                ' Array enthält alle selectierten nodes (max 20?)
Private NoKontextMenueList As String                                ' Liste der Listviews/TreeViewe elemente ohne kontextmenue
Private lngNavIndex As Integer                                      ' Aktuelle pos im nav Array

Private ThisDBCon As Object                                         ' Diese Datenbank verbindung

Private Sub Form_Load()
    
    Dim objError As Object

On Error GoTo Errorhandler
    
    Set objError = objObjectBag.GetErrorObj()                       ' Error Ob für Fehlerbehandlung Initialisieren
    
    Call InitMainForm                                               ' Hauptform initialisieren
    
exithandler:
On Error Resume Next

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
On Error Resume Next
    ' Optionen des Hauptforms Speichern
    Call objOptions.SetOptionByName(OPTION_MAINSTATE, Me.WindowState)
    Call objOptions.SetOptionByName(OPTION_MAINSIZE, Me.Width & "/" & Me.Height)
    Call objOptions.SetOptionByName(OPTION_LASTNODE, Me.TVMain.SelectedItem.Key)
    Call objOptions.SetOptionByName(OPTION_SPLIT, curlngSplitposProz)   ' Spliter pos in optionen speichern
    Call AppExit                                                    ' Application beenden
End Sub

Private Function InitMainForm()
    
    Dim szSize As String                                            ' SizeValue aus Reg
    Dim szSizeArray() As String                                     ' (0) = Width , (1) = Height
    Dim szWinState As String                                        ' Window State (min, max, normal)
    Dim szSplitpos As String                                        ' Pos des Splitters

On Error GoTo Errorhandler

    Me.Caption = SZ_APPTITLE                                        ' Form Caption Setzten

    ' Option aus lesen
    szSplitpos = objOptions.GetOptionByName(OPTION_SPLIT)           ' Spliter pos
    If szSplitpos <> "" Then
        ' Hier noch saubere behandlung wenn kein wert vorhanden
        curlngSplitposProz = CSng(szSplitpos)
    Else
        curlngSplitposProz = 0.3                                    ' Default Spliter pos setzen
    End If
    
    szSize = objOptions.GetOptionByName(OPTION_MAINSIZE)            ' Option WindowSize auslesen
    If szSize <> "" Then
        szSizeArray = Split(szSize, "/")                            ' Value aufspliten
        Me.Width = szSizeArray(0)                                   ' (0) = Width
        Me.Height = szSizeArray(1)                                  ' (1) = Height
    End If
    
    Call RepaintMainForm(curlngSplitposProz)                        ' Form neuzeichnen
    Call InitStatusBarMain                                          ' Statusbar initialisieren
    Call InitLV                                                     ' ListView initialisieren
    Call InitTree                                                   ' TreeView initialisieren
   
    'Call InitGrid

exithandler:
On Error Resume Next
    Me.Refresh
    Call objObjectBag.ShowMSGForm(False, "")
    Call ShowSplash(False)

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitMainForm", errNr, errDesc)
    Resume exithandler
End Function

Private Function InitTree()
' Tree View Initialisieren

    Dim i As Integer                                                ' Counter
    Dim cnode As node                                               ' Aktueller Node
    Dim szTmpNodeName As String                                     ' Temp Node Name
    Dim RootNodeArray() As String                                   ' Array mit RootNodenamen
    Dim szRootNodeList As String                                    ' Nodelist als String
    Dim TVNode As TreeViewNodeInfo                                  ' TreeNode Informationen
    Dim szLastNode As String                                        ' Lezter Node Als String aus Reg
    Dim bStartWithlastNode As Boolean                               ' Möchte der Anwender auf dem Letzen Knoten starten
    
On Error GoTo Errorhandler
    Me.MousePointer = vbHourglass                                   ' Sanduhr
                                                                    ' Liste der Rootknoten holen
    szRootNodeList = objTools.GetRootNodeListFromXML(App.Path & "\" & INI_XMLFILE)
    RootNodeArray = Split(szRootNodeList, ";")                      ' Rootnode list in array
    
    For i = 0 To UBound(RootNodeArray)                              ' Alle RootNodes durchlaufen
        szTmpNodeName = RootNodeArray(i)                            ' Akt Nodename holen
        If Not User.System Then                                     ' Wenn NichtUser systemverwalter
            ' Konten Stammdaten und benutzerverwaltung ausblenden
            If (szTmpNodeName = "Stammdaten") Or (szTmpNodeName = "Benutzerverwaltung") Then
                GoTo Skip                                           ' Rest überspringen
            End If
        End If
        With TVNode
            ' Node daten holen
            Call objTools.GetTVNodeInfofromXML(App.Path & "\" & INI_XMLFILE, szTmpNodeName, _
                    .szTag, .szText, .szKey, .bShowSubnodes, .szSQL, .szWhere, .lngImage)
            If .szTag <> "" And .szKey <> "" And .szText <> "" Then ' Wenn Tag und Key und Textvorhanden
                'If .lngImage = "" Then .lngImage = "1"                ' Gegebenenfalls Default image holen
                ' Tree Node anlegen
                Call AddTreeNode_New(TVMain, "", .szKey, .szTag, .szText, ThisDBCon, CLng(.lngImage), Not .bShowSubnodes)
            End If
        End With
Skip:
    Next i
    
'    ' Liste der Einträge ohne Kontextmenu ermitteln (für LV und TV)
'    NoKontextMenueList = objTools.GetNoKontextListFromXML(App.Path & "\" & INI_XMLFILE)
    bStartWithlastNode = objOptions.GetOptionByName(OPTION_STARTLASTNODE) ' auf den letzten node springen?
    
    If bStartWithlastNode Then
        szLastNode = objOptions.GetOptionByName(OPTION_LASTNODE)    ' Letzten Konten aus Reg holen
        If szLastNode <> "" Then                                    ' Wenn Reg Wert vorhanden
            Set cnode = GetNodeByKey(TVMain, szLastNode)            ' entsprechenden TreeNode suchen
            If Not cnode Is Nothing Then                            ' Wenn Tree Node existiert
                Call SelectTreeNode(TVMain, cnode)                  ' Node auswählen
                Call HandleNodeClick(cnode, False)                  ' Node Klick ausführen
                GoTo exithandler                                    ' Fertig
            End If
        End If
    End If
    
    Set cnode = GetNodeByKey(TVMain, SZ_TREENODE_MAIN)              ' Sonst Ersten Konten auswählen
    Call SelectTreeNode(TVMain, cnode)                              ' Node auswählen
    Call HandleNodeClick(cnode, False)                              ' Node Klick ausführen
        
exithandler:
On Error Resume Next
    Me.TVMain.Enabled = True                                        ' TV Enablen
    Me.MousePointer = vbDefault                                     ' Mauszeiger wieder normal
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitTree", errNr, errDesc)
    Resume exithandler
End Function

Private Sub InitLV()
    ' List View Initialisieren
On Error GoTo Errorhandler

    Call ClearListView(Me.LVMain)                                   ' evtl. List Items Löschen
    LVMain.Icons = ILTree                                           ' Verweis auf Image List
    LVMain.SmallIcons = ILTree
    
    'Call ShowLV                                                    ' Zuerst ListView anzeigen, Grid ausblenden
    
exithandler:
On Error Resume Next
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitLV", errNr, errDesc)
    Resume exithandler
End Sub

Private Function InitStatusBarMain()

On Error GoTo Errorhandler
    
    StatusBarMain.Panels(1).Alignment = sbrCenter
    StatusBarMain.Panels(1).Text = Left(CStr(Now()), 10)            ' Datum anzeigen
    
    StatusBarMain.Panels(2).Alignment = sbrLeft
    StatusBarMain.Panels(2).Text = User.UserName                    ' Angemelderter User
    
    StatusBarMain.Panels(3).Alignment = sbrLeft
    If objObjectBag.bUserIsAdmin Then                               ' Admin
        StatusBarMain.Panels(3).Text = "(Admin)"
    Else
        StatusBarMain.Panels(3).Text = "(Benutzer)"
    End If
        
    StatusBarMain.Panels(4).Alignment = sbrLeft
    StatusBarMain.Panels(4).Text = objDBconn.getDBtext              ' Db info
    
    StatusBarMain.Panels(5).Alignment = sbrRight                    ' Listitems Count
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitStatusBarMain", errNr, errDesc)
    Resume exithandler
End Function

Public Function DeleteDS(szRootkey As String, szID As String)

    Dim szSQL As String                                             ' SQL Statement
    Dim szMsg As String                                             ' Meldungtest Für User Nachfrage
    Dim szValue As String                                           ' DS Kennung (z.b. Name) damit der User weiss was er löscht
    Dim szTitle As String                                           ' Meldungs titel
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    
On Error GoTo Errorhandler
    
    If szID = "" Or szRootkey = "" Then GoTo exithandler            ' Keine DS ID -> fertig
    szDetails = "DS ID: " & szID & vbCrLf
    
    LVMain.MousePointer = vbHourglass
    
    Select Case UCase(szRootkey)
    Case UCase("Personen"), UCase("Bewerber"), UCase("Notare")
        Call DeletePerson(szID)                                     ' Detail Daten und Dokumente löschen daher sonderwurst
        GoTo exithandler                                            ' Fetig
        
    Case UCase("Fortbildungen")
        szSQL = "SELECT Thema011 FROM FORT011 WHERE ID011 ='" & szID & "'" '"CAST('" & ID & "' as uniqueidentifier)"
        szValue = objDBconn.GetValueFromSQL(szSQL)
        szMsg = "Möchten Sie die Fortbildung " & szValue & " wirklich löschen?"
        szTitle = "Fortbildung löschen"
        szSQL = "DELETE FROM FORT011 WHERE ID011 ='" & szID & "' "
        szSQL = szSQL & " DELETE FROM AFORT014 WHERE FK011014='" & szID & "' "
        
    Case UCase("Ausgeschriebene Stellen"), UCase("Stellen")
        szSQL = "SELECT Bezirk012 + ' ' + Cast(Frist012 as varchar(20)) FROM STELLEN012 WHERE ID012 ='" & szID & "'" '"CAST('" & ID & "' as uniqueidentifier)"
        szValue = objDBconn.GetValueFromSQL(szSQL)
        szMsg = "Möchten Sie die Ausgeschriebene Stelle " & szValue & " wirklich löschen?"
        szTitle = "Ausgeschriebene Stelle löschen"
        szSQL = "DELETE FROM STELLEN012 WHERE ID012 ='" & szID & "' "
        szSQL = szSQL & " DELETE FROM BEWERB013 WHERE FK012013='" & szID & "' "
        
    Case UCase("Bewerbung"), UCase("Bewerbungen")
        szSQL = "SELECT AZ013 FROM BEWERB013 WHERE ID013 ='" & szID & "'" '"CAST('" & ID & "' as uniqueidentifier)"
        szValue = objDBconn.GetValueFromSQL(szSQL)
        szMsg = "Möchten Sie die Bewerbung " & szValue & " wirklich löschen?"
        szTitle = "Bewerbung löschen"
        szSQL = "DELETE FROM BEWERB013 WHERE ID013 ='" & szID & "' "
        szSQL = szSQL & " DELETE FROM BEWERB013 WHERE FK012013='" & szID & "' "
        
    Case UCase("Stammdaten")
        
    Case UCase("Benutzerverwaltung"), UCase("Benutzer")
        szSQL = "SELECT USERNAME001 FROM USER001 WHERE ID001 ='" & szID & "'" '"CAST('" & ID & "' as uniqueidentifier)"
        szValue = objDBconn.GetValueFromSQL(szSQL)
        szMsg = "Möchten Sie den Benutzer " & szValue & " wirklich löschen?"
        szTitle = "Benutzer löschen"
        ' Hier evtl 2. Fragen
        szSQL = "DELETE FROM USER001 WHERE ID001 ='" & szID & "' "
    Case UCase("Dokumente"), UCase("Letzte Woche"), UCase("Letzter Monat")
        Call DeleteDokument(szID)                                   ' Dokument auch im Filesystem löschen dewegen eigene Fkt
        GoTo exithandler
        
    Case UCase("Aktenort")
        ' bei akten ort kein löschen vorgesehen
    'Case UCase("Vorgang")
        
    Case Else
    
    End Select
    
    If szValue = "" Then GoTo exithandler                           ' Kein DS gefunden -> Raus
    szDetails = szDetails & "Wert: " & szValue & vbCrLf
    
    If objError.ShowErrMsg(szMsg, vbOKCancel + vbQuestion, szTitle) <> vbCancel Then
        Call objDBconn.execsql(szSQL)                               ' Delete Statement ausführen
    End If
        
exithandler:
On Error Resume Next
    LVMain.MousePointer = vbDefault                                 ' Maus Zeiger wieder normal
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "DeleteDS", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Public Function DeletePerson(szPersID As String)
    ' Person löschen eigene Funktion da Sokumente und Detail daten
    Dim szSQL As String                                             ' SQL Statement
    Dim rsDoc As ADODB.Recordset                                    ' RS mit Dokument DS
    Dim szMsg As String                                             ' Meldungtest Für User Nachfrage
    Dim szValue As String                                           ' DS Kennung (z.b. Name) damit der User weiss was er löscht
    Dim szTitle As String                                           ' Meldungs titel
    Dim szPath As String                                            ' Dokumenten Pfad
    Dim szAblagePath As String                                      ' Pfad der Datei ablage
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    
On Error GoTo Errorhandler

    If szPersID = "" Then GoTo exithandler                          ' Ohne Pers ID -> fertig
    szDetails = "PersID: " & szPersID
    
    szAblagePath = objOptions.GetOptionByName(OPTION_ABLAGE) & "\"
    
    ' Pers Name für Nachfrage Holen
    szSQL = "SELECT Nachname010 + ', ' + ISNULL(Vorname010,'') FROM RA010 WHERE ID010 ='" & szPersID & "'" '"CAST('" & ID & "' as uniqueidentifier)"
    szValue = objDBconn.GetValueFromSQL(szSQL)
    szMsg = "Möchten Sie die Person " & szValue & " wirklich löschen?"
    szTitle = "Person löschen"
    ' Lösch SQL Staements Festlegen
    szSQL = "DELETE FROM AFORT014 WHERE FK010014 ='" & szPersID & "' "
    szSQL = szSQL & " DELETE FROM BEWERB013 WHERE FK010013='" & szPersID & "'"
    szSQL = szSQL & " DELETE FROM DOC018 WHERE FK010018='" & szPersID & "'"
    szSQL = szSQL & " DELETE FROM FORD022 WHERE FK010022 ='" & szPersID & "'"
    szSQL = szSQL & " DELETE FROM AKTENORT017 WHERE FK010017 ='" & szPersID & "'"
    szSQL = szSQL & " DELETE FROM RA010 WHERE ID010 ='" & szPersID & "'"
    
    If szValue = "" Then GoTo exithandler                           ' Keine Person gefunden -> fertig
    szDetails = "PersID: " & szPersID & vbCrLf & "Name: " & szValue
    
    szPath = objOptions.GetOptionByName(OPTION_ABLAGE) & "\" & szPath
     
    If objError.ShowErrMsg(szMsg, vbOKCancel + vbQuestion, szTitle) <> vbCancel Then
    
        ' Erst Dok Echt im Verz löschen
        Set rsDoc = ThisDBCon.fillrs("SELECT * FROM DOC018 WHERE FK010018='" & szPersID & "'")
        If Not rsDoc Is Nothing Then
            If rsDoc.RecordCount > 0 Then rsDoc.MoveFirst
            While Not rsDoc.EOF                                     ' Alle Docs durchlaufen
                szPath = rsDoc.Fields("DOCPATH018").Value
                szPath = szAblagePath & szPath                      ' Pfad zusammen setzen
                If objTools.FileDelete(szPath, True) Then           ' Wenn Doc gelöscht
                    'Stop        ' zum Debuggen
                End If
                rsDoc.MoveNext
            Wend
        End If
        Call objDBconn.execsql(szSQL)                               ' Dann Pers eintrag in Tabelle löschen
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
    Call objError.Errorhandler(MODULNAME, "DeletePerson", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Public Function DeleteDokument(szDocID As String)
    ' Dokument löschen eigene Funktion da Auch dateien im Dateisystem
    Dim szSQL As String                                             ' SQL Statement
    Dim szMsg As String                                             ' Meldungtest Für User Nachfrage
    Dim szValue As String                                           ' DS Kennung (z.b. Name) damit der User weiss was er löscht
    Dim szTitle As String                                           ' Meldungs titel
    Dim szPath As String                                            ' Dokumenten Pfad
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    
On Error GoTo Errorhandler

    If szDocID = "" Then GoTo exithandler                           ' Keine Doc ID -> Fertig
    szDetails = "DocID: " & szDocID & vbCrLf
    
    szSQL = "SELECT DOCNAME018 FROM DOC018 WHERE ID018 ='" & szDocID & "'" '"CAST('" & ID & "' as uniqueidentifier)"
    szValue = objDBconn.GetValueFromSQL(szSQL)                      ' Dokumenten namen ermitteln
    szDetails = szDetails & "DocName: " & szValue & vbCrLf
    
    szSQL = "SELECT DOCPATH018 FROM DOC018 WHERE ID018 ='" & szDocID & "'" '"CAST('" & ID & "' as uniqueidentifier)"
    szPath = objDBconn.GetValueFromSQL(szSQL)                       ' Dokumenten Pfad ermitteln
    szDetails = szDetails & "DocPath: " & szValue & vbCrLf
    
    szMsg = "Möchten Sie das Dokument " & szValue & " wirklich löschen?"
    szTitle = "Dokument löschen"
    szSQL = "DELETE FROM DOC018 WHERE ID018 ='" & szDocID & "' "
    
    If szValue = "" Or szPath = "" Then GoTo exithandler            ' Kein Doc gefunden -> Fertig
    
    szPath = objOptions.GetOptionByName(OPTION_ABLAGE) & "\" & szPath
    
    If objError.ShowErrMsg(szMsg, vbOKCancel + vbQuestion, szTitle) <> vbCancel Then
        ' Erst Dok Echt im Verz löschen
        If objTools.FileDelete(szPath, True) Then                   ' Wenn Doc gelöscht
            'Stop        ' zum Debuggen
        End If
        Call objDBconn.execsql(szSQL)                               ' Dann Doc eintrag in Tabelle löschen
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
    Call objError.Errorhandler(MODULNAME, "DeleteDokument", errNr, errDesc, szDetails)
    Resume exithandler
End Function

Private Sub NewDS(szRootkey As String, _
        Optional parentform As Form, Optional bdialog As Boolean)
    ' Öffnet leeres form für Neuen DS
On Error GoTo Errorhandler
                    
    Call OpenEditForm(szRootkey, "", parentform, bdialog)
    
exithandler:
On Error Resume Next
    LVMain.MousePointer = vbDefault
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "NewDS", errNr, errDesc)
    Resume exithandler
End Sub

Public Function OpenEditForm(szRootkey As String, DetailKey As String, _
        Optional parentform As Form, Optional bdialog As Boolean) As String
' Öffnet ein DS Form (edit) zum anzeigen und bearbeiten von DS

    Dim NewFrmEdit As Form                                          ' Neues DS (edit) Form
    Dim i As Integer                                                ' Counter
    Dim lngEditFormCount  As Integer                                ' Anzahl der Edit Forms
    Dim ID As String                                                ' evtl. DS ID
    Dim Detailarray() As String                                     ' Array aus evtl zusammegestzten IDs
    Dim szTmpImageIndex As String                                   ' Image Index des Edit Forms
    
On Error Resume Next
    lngEditFormCount = UBound(EditFormArray)                        ' Anz. Edit form ermitteln
    If Err.Number <> 0 Then                                         ' Errorhandling Deak. da Array evtl. leer
        lngEditFormCount = -1
        Err.Clear                                                   ' Fehler Resetten
    End If
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung wieder aktivieren

    If InStr(DetailKey, ";") > 0 Then                               ' ID Zusammengesetzt
        Detailarray = Split(DetailKey, ";")                         ' Wenn JA aufspalten
        ID = Detailarray(0)
    Else
        ID = DetailKey
    End If
    
    If ID <> "" And lngEditFormCount > -1 Then                      ' ID und offenes Edit Form vorhanden ?
        For i = 0 To lngEditFormCount                               ' Durch FormsArray laufen und überprüfen ob schon ein mit ID offen ist
            If Not EditFormArray(i) Is Nothing Then
                If EditFormArray(i).ID = ID Then                    ' Prüfen ob gleiches Edit form erneut geöffnet werden soll
                    EditFormArray(i).Show                           ' Wenn Ja anzeigen
                    GoTo exithandler                                ' Fertig
                End If
            End If
        Next i
    End If
    
On Error Resume Next                                                ' Fehler behandlung deaktivieren

    ReDim Preserve EditFormArray(UBound(EditFormArray) + 1)         ' Sonst Form Array erweitern
    If Err.Number <> 0 Then                                         ' Wenn Fehler
        ReDim EditFormArray(0)                                      ' 1. Form
        Err.Clear                                                   ' Fehler Resetten
    End If

On Error GoTo Errorhandler                                          ' Fehlerbehandlung wieder aktivieren
    
    ' Und neues form öffnen
    Select Case UCase(szRootkey)                                    ' Form aus Rootkey ermitteln
    Case "Ausschreibung"                                            ' Ausschreibung (zur Zeit nicht benutzt)
'        Set NewFrmEdit = New frmAusschreibungEdit
'        Call NewFrmEdit.InitEditForm(objDBconn, DetailKey)
        
    Case UCase("Personen"), UCase("Bewerber"), UCase("Teilnehmer"), _
                UCase("Notare"), UCase("Notare bestellt"), UCase("Notare ausgeschieden")
        Set NewFrmEdit = New frmEditPersonen
        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey, bdialog)
        If DetailKey = "" Then
            If UCase(szRootkey) = "BEWERBER" Then NewFrmEdit.cmbStatus.Text = "Bewerber"
            If Left(UCase(szRootkey), 5) = "NOTAR" Then NewFrmEdit.cmbStatus.Text = "Notar"
        End If
    Case UCase("Fortbildungen")                                     ' Fortbildungen
        Set NewFrmEdit = New frmEditFortbildungen
        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey)
    Case UCase("Ausgeschriebene Stellen"), UCase("Stellen"), UCase("StellenJahr")   ' Stellen
        Set NewFrmEdit = New frmEditStellen
        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey, bdialog)
    Case UCase("Bewerbung"), UCase("Bewerbungen")                   ' Bewerbungen
        Set NewFrmEdit = New frmEditBewerbung
        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey)
    Case UCase("Stammdaten"), UCase("Landgerichte"), UCase("Amtsgerichte")  ' Stammdaten Gerichte
        Set NewFrmEdit = New frmEditOLG
        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey)
    Case UCase("Benutzerverwaltung"), UCase("Benutzer")             ' Stammdaten Benutzer
        Set NewFrmEdit = New frmEditUser
        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey)
    Case UCase("Dokumente"), UCase("Letzte Woche"), UCase("Letzter Monat")  ' Dokumente
        Call ShowWordDoc(DetailKey)
        GoTo exithandler
    Case UCase("Aktenort")                                          ' Aktenort
        Set NewFrmEdit = New FrmEditAktenOrt
        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey)
    Case UCase("Disziplinarmaßnahmen")                              ' Disziplinarmaßnahmen
        Set NewFrmEdit = New frmEditDisz
        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey)
    Case UCase("Vorgang")                                           ' Vorgang (zur Zeit nicht benutzt)
        'Set NewFrmEdit = New frmVorgangSelect
        'Call NewFrmEdit.InitEditForm(objDBconn, DetailKey)
    Case UCase("Forderungen")                                       ' Forderungen
        Set NewFrmEdit = New frmEditForderungen
        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey)
        
    Case Else                                                       ' Sonstiges Form
        Set NewFrmEdit = New frmEdit
        Call NewFrmEdit.InitEditForm(parentform, ThisDBCon, szRootkey, DetailKey)
        Select Case UCase(szRootkey)
        Case UCase("Amtsgerichte")
            lngImageIndex = 1
        Case UCase("Landgerichte")
            lngImageIndex = 1
        Case Else
            lngImageIndex = 1
        End Select
    End Select
    
    ' Protoklieren
    Call objError.WriteProt("OpenEditForm - RootKey: " & szRootkey & vbCrLf & "DetailKey: " & DetailKey)

     ' Form icon Laden
    'szTmpImageIndex = objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_IMAGE, szRootkey)
    'If IsNumeric(szTmpImageIndex) Then NewFrmEdit.Icon = ILTree.ListImages(CLng(szTmpImageIndex)).Picture
    
    Set EditFormArray(UBound(EditFormArray)) = NewFrmEdit           ' Form in EditFormArray
    
    If bdialog Then                                                 ' Form anzeigen
        NewFrmEdit.Show 1, Me
        OpenEditForm = NewFrmEdit.ID                                ' ID aus form übernehmen (nur dialog)
        Call EditFormUnload(NewFrmEdit)                             ' Form Schliessen
    Else
        NewFrmEdit.Show                                             ' einfach anzeigen (kein  dialog)
    End If
    
'    If ParentForm Is Nothing Then
'        NewFrmEdit.Show
'    Else
'        NewFrmEdit.Show 1, ParentForm                              ' DB Form anzeigen
'    End If
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "OpenEditForm", errNr, errDesc)
    Resume exithandler
End Function

Public Function CloseEditForm(frmEdit As Form)
    ' Schliesst ein Edit Form
    
    Dim i As Integer                                                ' Counter
    Dim lngEditFormCount  As Integer                                ' Anzahl der Edit Forms
    
On Error Resume Next

    lngEditFormCount = UBound(EditFormArray)                        ' Anz. Edit form ermitteln
    If Err.Number <> 0 Then                                         ' Errorhandling Deak. da Array evtl. leer
        lngEditFormCount = -1
        Err.Clear                                                   ' Fehler Resetten
    End If
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung wieder aktivieren

    If lngEditFormCount > -1 Then                                   ' wenn EditForms Array nicht leer
        For i = 0 To UBound(EditFormArray)                          ' Array duchlaufen
            If Not EditFormArray(i) Is Nothing Then
                If EditFormArray(i).ID = frmEdit.ID Then            ' Form anhand ID ermitteln
                    Set EditFormArray(i) = Nothing                  ' Nothing setzen
                    Exit For
                End If
            End If
        Next
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
    Call objError.Errorhandler(MODULNAME, "CloseEditForm", errNr, errDesc)
    Resume exithandler
End Function

Public Function RefreshListView(LV As ListView, TV As TreeView)
' Aktualisiert nur das Listview

    Dim cnode As node                                               ' Akt. Tree Node
    Dim szLvItemKey As String                                       ' List view Item Key

On Error Resume Next                                                ' Errorhandling deak. da selectedItem evtl .leer

    szLvItemKey = LV.SelectedItem.Key
    Err.Clear
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung wieder aktivieren
    
    Set cnode = GetNodeByKey(TV, TV.SelectedItem.Key)               ' Akt. Node ermitteln
    If Not cnode Is Nothing Then Call HandleNodeClick(cnode, True)  ' Duch Nodeclick Listvieew neuaufbauen
    
    If szLvItemKey <> "" Then Call SelectLVItem(LV, szLvItemKey)    ' Selectierten eintrag im LV wider auswählen
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "RefreshListView", errNr, errDesc)
End Function

Public Function RefreshTreeView(TV As TreeView, Optional nodekey As String, Optional bParent As Boolean)
' Aktualisiert TreeView

    Dim cnode As node                                               ' Akt. TV Node
    Dim cParentNode As node                                         ' Übergeordneter node (Parent)

On Error GoTo Errorhandler
    
    If nodekey <> "" Then                                           ' Wenn Node Key angegeben
        Set cnode = GetNodeByKey(TV, nodekey)                       ' Node aus szKey ermitteln
        If cnode Is Nothing Then                                    ' Wenn Node Nicht existiert
            Set cnode = GetNodeByKey(TV, TV.SelectedItem.Key)       ' Node aus Selected Item ermitteln
        End If
    Else
        Set cnode = GetNodeByKey(TV, TV.SelectedItem.Key)           ' Node aus Selected Item ermitteln
    End If
    
    If Not cnode Is Nothing Then
        If Not cnode.Parent Is Nothing Then
            Set cParentNode = cnode.Parent
            Call DelSubTreeNodes(TV, cParentNode)                   ' Alle unter knoten löschen
            Call HandleNodeClick(cParentNode, True)                 ' Unterknoten anlegen
        End If
        'Call HandleNodeClick(cNode, True)                          ' Unterknoten anlegen
    End If
    Set cnode = GetNodeByKey(TV, nodekey)                           ' Akt Node ermitteln
    If Not cnode Is Nothing Then                                    ' Wenn Node existiert
        Call SelectTreeNode(TV, cnode)                              ' Node Selecten
        Call HandleNodeClick(cnode, True)                           ' Node Klick behandeln
    End If
    
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "RefreshTreeView", errNr, errDesc)
End Function

Private Function RepaintMainForm(lngSplitpos As Single)
    ' Positioniert die Steuerelemente auf dem Main form
    
    Const lngSpliterWidth = 70                                      ' Spliter breite
    Const lngStatusHeight = 300                                     ' Statusbar höhe
    Dim CTLTop  As Integer                                          ' Max CTL Top Pos.
    Dim CTLLeft As Integer                                          ' Max CTL Left Pos.
    Dim CTLWidth As Integer                                         ' Max CTL Breite
    Dim CTLHeight As Integer                                        ' Max CTL Höhe
    
On Error GoTo Errorhandler

    If Me.WindowState = vbMinimized Then GoTo exithandler           ' Mainform min. -> Fertig
    
    If Me.Width < 3500 Then Me.Width = 3500                         ' Min. Breite nicht unterschreiten
    If Me.Height < 3000 Then Me.Height = 3000                       ' Min Höhe nicht unterschreiten
    
    If Me.ScaleWidth = 0 Or Me.ScaleHeight = 0 Then GoTo exithandler
    If lngSplitpos = 0 Then lngSplitpos = 3000                      'Min Spliter Pos
    
    'If Me.ScaleTop <= 1000 Then Me.ScaleTop = 1000
    CTLTop = Me.ScaleTop + Me.ToolbarMain.Height                    ' Max Top pos. ermitteln
    'If Me.ScaleLeft <= 2000 Then Me.ScaleLeft = 2000
    
    CTLLeft = Me.ScaleLeft                                          ' Max Left Pos. ermitteln
    'If Me.ScaleHeight < 2000 Then Me.ScaleHeight = 2000
    CTLHeight = Me.ScaleHeight - lngStatusHeight - Me.ToolbarMain.Height ' Max höhe ermitteln
    'If Me.ScaleWidth <= 3000 Then Me.ScaleWidth = 3000
    
    CTLWidth = Me.ScaleWidth                                        ' Max breite ermitteln
    
    Call PicTree.Move(0, CTLTop, (CTLWidth * lngSplitpos), CTLHeight) ' PicTree ausrichten (Splitter)
    
    Call TVMain.Move(0, 0, PicTree.Width, PicTree.Height)           ' Tree an PicTree ausrichten
    
    ' PicList Ausrichten
    Call PicList.Move(PicTree.Width + lngSpliterWidth, CTLTop, CTLWidth - PicTree.Width - lngSpliterWidth, CTLHeight)
    
    Call LVMain.Move(0, 0, PicList.Width, PicList.Height)           ' ListView an PicList ausrichten
    
    curlngSplitposProz = lngSplitpos                                ' Akt. Spliter pos merken
    Call objOptions.SetOptionByName(OPTION_SPLIT, curlngSplitposProz)   ' Spliter pos in optionen speichern
    LVMain.Refresh                                                  ' LV Neuzeichnen
    
exithandler:

Exit Function
Errorhandler:
'    Debug.Print "Me.ScaleWidth: " & Me.ScaleWidth
'    Debug.Print "Me.ScaleHeight: " & Me.ScaleHeight
'    Debug.Print "lngSplitpos: " & lngSplitpos
    
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "RepaintMainForm", errNr, errDesc)
    Resume exithandler
End Function

Public Function HandleNodeClick(ByVal node As MSComctlLib.node, Optional bExpand As Boolean)
' Behandelt den Node Klick im TreeView

    Dim szTmpName As String
    Dim szKeyArray() As String                                      ' Node Key in array aufgespalten
    Dim szTagArray() As String                                      ' Node Tag in array aufgespalten
    Dim sztmp As String                                             ' Hilfsvariable
    Dim ID As String                                                ' Evtl ID des Detaildatensatzes
    Dim bValueList As Boolean                                       ' Darstellung des ListView als Value list
    Dim i As Integer                                                ' Counter
    Dim LVInfo As ListViewInfo                                      ' Infos zum LV Handling aus XML
    
On Error GoTo Errorhandler
        
    szTagArray = Split(node.Tag, "\")                               ' Node Tag aufspalten
    szKeyArray = Split(node.Key, "\")                               ' Node Key aufspalten

    Me.MousePointer = vbHourglass                                   ' Mousepointer auf Sanduhr
    DoEvents                                                        ' Evtl. ander aktionen zulassen
    
    ' Evtl Subnodes hinzufügen
    If node.Children = 0 Then Call AddSubTreeNodes(TVMain, node, ThisDBCon, node.Image, True)

    If Not bNotExpand And Not node.Expanded Then node.Expanded = True   ' Evtl. Knoten auffalten
    
    Call SaveColumnWidth(LVMain)                                    ' Splaten breite speichern
    
    If szTagArray(UBound(szTagArray)) = "*" Then                    ' detaildatensatz
        ID = GetLastKey(node.Key, TV_KEY_SEP)                       ' ID aus Key ermitten
    Else
        If InStr(node.Tag, "*") Then                                ' Statischer unterknoten eines Detaildatensatzes
            sztmp = ""
            i = UBound(szTagArray) + 1
            While sztmp <> "*"                                      ' Tag bis * rückwärts durchlaufen
                i = i - 1
                sztmp = szTagArray(i)
                ID = szKeyArray(i)                                  ' ID am gleicher stelle aus Tag
            Wend
        End If
         If ID = "" Then ID = GetLastKey(node.Key, TV_KEY_SEP)
    End If
    
    With LVInfo
        ' LV infos aus mxl datei holen
        Call objTools.GetLVInfoFromXML(App.Path & "\" & INI_XMLFILE, node.Tag, .szSQL, .szTag, .szWhere, _
                .lngImage, .bValueList, .bListSubNodes)
        Call ListLVByTag(LVMain, ThisDBCon, node.Tag, ID, .bValueList, node.Image)  ' Listitems anzeigen
        ' Subnodes auch als list items anzeigen?
        If .bListSubNodes Then Call ListLVFromSubNodes(LVMain, TVMain, node) ' Subnodes im LV anzeigen
        
        If Not .bValueList Then
            Call CountLVItems(LVMain)                               ' Anzahl der listitem in statusbar anzeigen
        Else
            Call CountLVItems(LVMain, 1)                            ' Wenn .bValueList anzahl ist immer 1
        End If
    End With
    
exithandler:
On Error Resume Next
    Me.MousePointer = vbDefault                                     ' Mauszeiger wieder normal
    DoEvents                                                        ' Andere Events zulassen
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "HandleNodeClick", errNr, errDesc)
End Function

Private Function GetPersIDFormLV() As String
    Dim szRootkey As String
    Dim szPersID As String
    
On Error GoTo Errorhandler

    Call GetKontextRoot(szRootkey, szPersID, "")
    
    Select Case UCase(szRootkey)
    Case UCase("Personen"), UCase("Bewerber"), UCase("Teilnehmer"), _
                UCase("Notare"), UCase("Notare bestellt"), UCase("Notare ausgeschieden")
        GetPersIDFormLV = szPersID
    Case Else
    
    End Select
    
exithandler:
On Error Resume Next
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetPersIDFormLV", errNr, errDesc)
End Function

Private Function DoNav(bBack As Boolean)
    ' Navigiert durch die Liste der gespeicherten Tree nodes
    
    Dim NavIndexMax As Integer                                      ' Index Obergrenze des nav Arrays
    Dim cnode As node                                               ' Akt TV Node
    Dim szKey As String                                             ' Node key
    
On Error Resume Next                                                ' Errorhandling deak. da Array evtl. leer

    NavIndexMax = UBound(NavTreeNodeArray)                          ' Array überprüfen
    If Err.Number <> 0 Then
        NavIndexMax = -1
        Err.Clear
    End If
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung wieder aktivieren
    
    If NavIndexMax = -1 Then GoTo exithandler                       ' Kein vor oder zurück möglich
    
    If bBack Then                                                   ' Rückwärts
        If lngNavIndex < 0 Then GoTo exithandler                    ' Kein zurück möglich
        If lngNavIndex > 0 Then lngNavIndex = lngNavIndex - 1
    Else                                                            ' Vorwärts
        If lngNavIndex = NavIndexMax Then GoTo exithandler          ' Kein vor möglich
        lngNavIndex = lngNavIndex + 1
    End If
    
    szKey = NavTreeNodeArray(lngNavIndex)                           ' Key aus array ermitteln
    If szKey = "" Then GoTo exithandler                             ' Kein Key fertig
    
    Set cnode = GetNodeByKey(TVMain, szKey)                         ' Node mit Key Ermitteln
    If Not cnode Is Nothing Then                                    ' Wenn Knoten gefunden
        Call HandleNodeClick(cnode)                                 ' Click behandeln
        Call SelectTreeNode(TVMain, cnode)
    End If
    
    ToolbarMain.Buttons(2).Enabled = Not (lngNavIndex >= NavIndexMax) ' evtl. Button Vor disablen
    'If lngNavIndex = NavIndexMax Then ToolbarDB.Buttons(2).Enabled = False
    ToolbarMain.Buttons(1).Enabled = Not (lngNavIndex = 0)          ' evtl Button Zurück disablen
    'If lngNavIndex = 0 Then ToolbarDB.Buttons(1).Enabled = False
    ToolbarMain.Refresh                                             ' Toolbar neuzeichnen
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "AddNodeToNavList", errNr, errDesc)
    Resume exithandler
End Function

Private Function AddNodeToNavList(szKey As String)
' speichert szKey (NodePath im Tree) in array für Navigation

On Error GoTo Errorhandler

    If szKey = "" Then GoTo exithandler                             ' kein Key -> Fertig
    
On Error Resume Next                                                ' Errorhandling deak. da Array evtl. leer

    ReDim Preserve NavTreeNodeArray(UBound(NavTreeNodeArray) + 1)   ' Array Prüfen
    If Err.Number <> 0 Then                                         ' Array Leer
        ReDim NavTreeNodeArray(0)
        NavTreeNodeArray(UBound(NavTreeNodeArray)) = szKey          ' Key anfügen
        Err.Clear
        GoTo exithandler                                            ' Fertig
    End If
        
On Error GoTo Errorhandler                                          ' Fehlerbehandlung wieder aktivieren

    If UBound(NavTreeNodeArray) >= 1 Then                           ' array nicht leer
        If NavTreeNodeArray(UBound(NavTreeNodeArray) - 1) = szKey Then GoTo exithandler
    End If
    NavTreeNodeArray(UBound(NavTreeNodeArray)) = szKey              ' Key anfügen
    lngNavIndex = UBound(NavTreeNodeArray)
    ToolbarMain.Buttons(1).Enabled = True                           ' Button Vor enablen
    ToolbarMain.Refresh                                             ' Toolbar neuzeichnen
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "AddNodeToNavList", errNr, errDesc)
    Resume exithandler
End Function

Public Sub HandleLVItemDblKlick()
' Behandelt den Doppelklick in Listview

    Dim szRootkey As String                                         ' TVNode /LV Key als Kontext
    Dim szDetailKey As String                                       ' evtl. Datensatz ID
    Dim szAction As String                                          ' Für diesen Node vorgesehene Aktion
    Dim cnode As node                                               ' Akt. TreeNode
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    
On Error GoTo Errorhandler
    
    Call GetKontextRoot(szRootkey, szDetailKey, szAction)           ' Kontext ermitteln
    szDetails = "RootKey: " & szRootkey & vbCrLf & "ID: " & szDetailKey & vbCrLf & "Aktion: " & szAction
    
    Select Case szAction
    Case "Edit"
        Call frmMain.OpenEditForm(szRootkey, szDetailKey, Me)       ' DS zum bearbeiten öffnen
    Case "SelectNode"
            Set cnode = GetNodeByKey(TVMain, LVMain.SelectedItem.Key)   ' TVNode ermitteln
            If Not cnode Is Nothing Then                            ' Wenn Node Exitsiert
                Call SelectTreeNode(TVMain, cnode, True)            ' Node auswählen
                Call HandleNodeClick(cnode)                         ' Node Klick behandeln
'            Else
'                Call frmMain.OpenEditForm(szRootkey, szDetailKey, Me)
            End If
    Case ""
        'Call frmMDIMain.OpenEditForm(szRootkey, szDetailKey)       ' DS zum bearbeiten öffnen
            Set cnode = GetNodeByKey(TVMain, LVMain.SelectedItem.Key)   ' TVNode ermitteln
            If Not cnode Is Nothing Then                            ' Wenn Node Exitsiert
                Call SelectTreeNode(TVMain, cnode, True)            ' Node auswählen
                Call HandleNodeClick(cnode)                         ' Node Klick behandeln
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
    Call objError.Errorhandler(MODULNAME, "HandleLVItemDblKlick", errNr, errDesc, szDetails)
End Sub

Private Sub HandleLVKontextNew()
' Öffent neuen DS aus Kontextmenü

    Dim szKeyArray() As String                                      ' Array mit Key elementen
    Dim szRootkey As String                                         ' Gibt an was für ein DS neu angelegt werden soll
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    
On Error GoTo Errorhandler
    
    Call GetKontextRoot(szRootkey, "", "")                          ' Rootkey ermitteln
    Call OpenEditForm(szRootkey, "", Me)                            ' DS form für neuen DS öffnen
    
exithandler:
On Error Resume Next
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "HandleLVKontextNew", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub HandleLVKontextDel()
' Löscht DS aus Kontextmenü

    Dim szRootkey  As String                                        ' Gibt an was für ein DS neu angelegt werden soll
    Dim szID As String                                              ' ID des Datensatzes

On Error GoTo Errorhandler
    
    Call GetKontextRoot(szRootkey, szID, "")                        ' ID und Rootkey ermitteln
    If szID = "" Then GoTo exithandler                              ' Keine ID -> Fertig
    Call DeleteDS(szRootkey, szID)                                  ' DS Löschen
    Call RefreshListView                                            ' DS Aktualisieren
    
exithandler:
On Error Resume Next
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "HandleLVKontextDel", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub HandleLVKontextEdit()
' Öffent DS aus Kontextmenü

    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    Dim szRootkey As String                                         ' Gibt an was für ein DS neu angelegt werden soll
    Dim szDetailKey As String                                       ' DS Id
    Dim szAction As String                                          ' Aktion (evtl. edit nicht zulassig
    
On Error Resume Next                                                ' Errorhandling deakt. da SelectedItem evtl. nicht ex.
    szDetails = "LVTag: " & LVDB.Tag & vbCrLf & "SelectedItem.tag: " & LVDB.SelectedItem.Key
    Err.Clear
On Error GoTo Errorhandler                                          ' Fehlerbehandlung wieder aktivieren
    
    Call GetKontextRoot(szRootkey, szDetailKey, szAction)           ' ID, Rootkey und Aktion ermitteln
    
    If szDetailKey = "" Then GoTo exithandler                       ' Keine ID -> Fertig
    Call frmMain.OpenEditForm(szRootkey, szDetailKey, Me)           ' Form zum bearbeuten öffnen
    
exithandler:
On Error Resume Next
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "HandleKontextEdit", errNr, errDesc, szDetails)
    Resume exithandler
End Sub

Private Sub HandleLVKontextNewDoc()
' Öffnet neues Dokument aus Kontextmenü

    Dim szPersID As String                                          ' Empfänger ID
    
On Error GoTo Errorhandler

    szPersID = GetPersIDFormLV()                                    ' Empfäger ID ermitteln
    If szPersID = "" Then GoTo exithandler                          ' Ohne Empfänger fertig
    Call WriteWord("", szPersID)                                    ' Dok erstellen
    
exithandler:
On Error Resume Next
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "HandleLVKontextNewDoc", errNr, errDesc, szDetails)
    Resume exithandler
End Sub

'Public Function GetKontextRoot(bTV As Boolean,szRootkey As String, szDetailKey As String, Optional szAction As String)
Public Function GetKontextRoot(szRootkey As String, szDetailKey As String, Optional szAction As String)
' ermittelt Rootkey und DS ID sowie mögliche Aktionen aus den Kontext (ListView) und XML
    
    Dim szItemTagArray() As String                                  ' Array mit ListView Item Tag elementen
    Dim szItemKeyArray() As String                                  ' Array mit ListView Item Key elementen
    Dim szLVTagArray() As String                                    ' Array mit ListView Tag elementen
    Dim szItemTag As String                                         ' Tag des ListView Items (* statt ID)
    Dim szItemKey As String                                         ' Key des ListView Items (enthält ID)
    Dim szLVTag As String                                           ' Tag des Listviews
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    Dim TVNode As TreeViewNodeInfo                                  ' Infos über TreeNode
    Dim LVInfo As ListViewInfo                                      ' Infos über ListView
     
On Error Resume Next                                                ' Errorhandling deakt. da SelectedItem evtl. Nothing
    szItemTag = LVMain.SelectedItem.Tag                             ' Tag des akt. ListView Items ermitteln
    szItemKey = LVMain.SelectedItem.Key                             ' Key des akt. ListView Items ermitteln
    szLVTag = LVMain.Tag                                            ' Tag des ListViews ermitteln
    
    'szDetails = "LVTag: " & lvmain.Tag & vbCrLf & "SelectedItem.tag: " & lvmain.SelectedItem.Key
    Err.Clear
On Error GoTo Errorhandler                                          ' Fehlerbehandlung wieder aktivieren

    With LVInfo                                                     ' ListViewInfo aus XML füllen
        Call objTools.GetLVInfoFromXML(App.Path & "\" & INI_XMLFILE, LVMain.Tag, _
                .szSQL, .szTag, .szWhere, .lngImage, .bValueList, .bListSubNodes, .bEdit, .bSelectNode)
        If .bEdit Then                                              ' Edit zulässig
            szAction = "Edit"
        End If
        If (Not .bEdit) And .bSelectNode Then szAction = "SelectNode"  ' Select zulässig
    End With
    
    If szItemKey = "" Or szItemTag = "" Then
        szRootkey = LVInfo.szTag
        GoTo exithandler
    End If
    
    szItemTagArray = Split(szItemTag, TV_KEY_SEP)                   ' ListView Item Tag aufspalten
    szItemKeyArray = Split(szItemKey, TV_KEY_SEP)                   ' ListView Item Key aufspalten
    szLVTagArray = Split(szLVTag, TV_KEY_SEP)                       ' ListView Tag aufspalten
    
    If UBound(szItemKeyArray) = UBound(szItemTagArray) Then
    
        If szItemTagArray(UBound(szItemTagArray)) <> "*" Then       ' Case 3 SubNode in Valuelist
            ' Ubound(szItemKeyArray ) = Ubound(szItemTagArray) / key mit ID / tag mit *
        Else                                                        ' Case 1 Detail SubNode in ListView
        ' Ubound(szItemKeyArray ) = Ubound(szItemTagArray) / key mit ID / tag mit *
        ' -> Select Node and/or Edit
            szDetailKey = szItemKeyArray(UBound(szItemKeyArray))    ' ID aus ListItem.Key ermitten
            szRootkey = szItemKeyArray(UBound(szItemKeyArray) - 1)  ' szRootkey aus Item Key ermitteln
        '    szAction = "Edit"
            If szRootkey = "*" Then                                 ' Case 2a einzelner DetailsDS (nicht in Valuelist)
                szRootkey = szItemKeyArray(UBound(szItemKeyArray) - 2)  ' Plan b ( dürft nicht vorkommen)
            End If
        End If
    ' case 4 Relation Deatilnode eines Detailnodes in Liste
    ' Ubound(szItemKeyArray ) = Ubound(szItemTagArray) / key mit ID / tag mit *
    ' szLVTagArray(UBound(szLVTagArray)) <>"*"
    ' Select Node (Edit?)
    
    Else                                                            ' Case 2 einzelner DetailsDS (Valuelist)
    ' Ubound(szLVTagArray) = Ubound(szItemTagArray) / LVTag mit ID / Itemtag mit *
        szDetailKey = szItemTagArray(UBound(szItemTagArray))        ' ID aus ListItem.Tag ermitten
        szRootkey = szItemTagArray(UBound(szItemTagArray) - 1)      ' szRootkey aus Item Tag ermitteln
        'szAction = "Edit"
    End If
    
    If szRootkey = "*" Or szRootkey = "" Then szAction = "SelectNode"
    
exithandler:
On Error Resume Next
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetKontextRoot", errNr, errDesc)
    Resume exithandler
End Function

Private Function CountLVItems(LV As ListView, Optional Count As Integer, Optional bNoShowInStatusBar As Boolean) As Integer

    Dim lngCount As Integer                                         ' Anzahl der angezeigten LV Items
    
On Error GoTo Errorhandler
    
    lngCount = LV.ListItems.Count                                   ' Durchzählen
    If Count > 0 Then lngCount = Count
    If Not bNoShowInStatusBar Then StatusBarMain.Panels(5).Text = lngCount & " DS"  ' In Statusbar anzeigen
    
exithandler:
On Error Resume Next
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "CountLVItems", errNr, errDesc)
    Resume exithandler
End Function

Private Sub HandleMenueKlick(szMenueName As String)
' Behandelt Menü Kontextmenü und Toolbar Klicks

    Dim ID As String                                                ' DS ID
    Dim Key As String                                               ' RootKey gibt an was für ein DS behandelt wird
    Dim PersID As String                                            ' ID eines Personen DS
    Dim tmpKey As String
    Dim szAction As String
    
On Error GoTo Errorhandler

    Select Case szMenueName
                                                                    ' *****************************************
                                                                    ' Main Menü
    Case "mnuDateiExit"                                             ' Anwendung Beenden
        Call AskForExit                                             ' Nachfragen ob Beenden
'        Call Unload(Me)                                            ' Main Form schliessen
'        Call AppExit                                               ' Application beenden
    Case "mnuDateiNewBewerber"
        Call NewDS("Bewerber", Me, False)                           ' Neuen Bewerber anlegen
    Case "mnuDateiNewBewerbung"
        Call NewDS("Bewerbung", Me, False)                          ' Neue Bewerbung anlegen
    Case "mnuDateiNewStelle"
        Call NewDS("Stellen", Me, False)                            ' Neue Stelle anlegen
    Case "mnuDateiNewDoc"
        Call WriteWord                                              ' Neues Dokument ohne vorgabe der Vorlage oder empfänger
    Case "mnuDateiSuchen"                                           ' Suchen
        ID = ShowSearch(objDBconn, Key, "")                         ' Such Dialog aufrufen
        If ID <> "" Then Call OpenEditForm(Key, ID)                 ' Erg anzeigen
'    Case "mnuEditAusschreibungNew"
'        Call OpenEditForm("Ausschreibung", "", Me)
    Case "mnuEditSelectTemplate"                                    ' Anschreiben
        'Call GetKontextRoot("", PersID, "")
        Call WriteWord(, PersID)                                    ' Starte Sat
'    Case "mnuEditWorkflow"
'        Call OpenEditForm("Vorgang", "", Me)
    Case "mnuExtrasOptions"                                         ' Optionen
        Call ShowOptions
    Case "mnuExtrasChangePWD"                                       ' Kennwort änderung
        Call ThisDBCon.UserChangePWD(User.NTUsername)
    Case "mnuInfoAbout"                                             ' Info Dialog
        Call ShowAbout
    Case "mnuInfoHelp"                                              ' Online Hilfe
        Call ShowHelp
                                                                    ' *****************************************
                                                                    ' Kontext menü
    Case "kmnuListNewUser", "kmnuListNewPerson", "kmnuListNew"
        Call HandleLVKontextNew                                     ' Neuer DS
    Case "kmnuListEditUser", "kmnuListEditPerson", "kmnuListEdit"
        Call HandleLVKontextEdit                                    ' Edit DS
    Case "kmnuListNewDocPerson", "kmnuListNewDocNotar", "kmnuListNewDocBewerber"
        Call HandleLVKontextNewDoc                                  ' Neues Doc für
    Case "kmnuListDelPerson", "kmnuListDelNotar", "kmnuListDelBewerber", _
            "kmnuListDelUser", "kmnuListDel"
        Call HandleLVKontextDel                                     ' DS Löschen
                                                                    ' *****************************************
                                                                    ' Toolbar ( szMenueName ist hier ButtonKey)
    Case TB_LEFT
        Call DoNav(False)                                           ' Vorwärts navigieren
    Case TB_RIGHT
        Call DoNav(True)                                            ' Rückwärts navigieren
    Case TB_REFRESH                                                 ' Refresh
        tmpKey = LVMain.SelectedItem.Key
        Call RefreshTreeView(TVMain, TVMain.SelectedItem.Key)
        Call SelectLVItem(LVMain, tmpKey)
    Case TB_NEW                                                     ' Neuer Ds
        Call GetKontextRoot(Key, "", szAction)                      ' Kontext ermitteln
        If Key <> "" And szAction <> "" Then
            Call NewDS(Key, Me, False)                              ' Neuer Datensatz
        End If
    Case TB_NEWBEWERBUNG
        Call NewDS("Bewerbung", Me, False)                          ' Neue Bewerbung anlegen
    Case TB_NEWBEWERBER
        Call NewDS("Bewerber", Me, False)                           ' Neuen Bewerber anlegen
    Case TB_NEWSTELLE
        Call NewDS("Stellen", Me, False)                            ' Neue Stelle anlegen
    Case TB_NEWDOC                                                  ' Starte SAT
        Call WriteWord                                              ' Neues Dokument ohne vorgabe des Vorlage oder empfänger
    Case TB_DOCNEW                                                  ' Starte SAT
        'Call WriteWord(, GetPersIDFormLV())
        'Call GetKontextRoot("", PersID, "")
        PersID = GetPersIDFormLV                                    ' Empfänger aus kontext ermitteln
        Call WriteWord(, PersID)                                    ' Neues Dokument mit empfänger
    Case TB_SEARCH                                                  ' Suchen
        Call GetKontextRoot(Key, "")                                ' Suchkontext ermitteln
        ID = ShowSearch(objDBconn, Key, "")
        If ID <> "" Then Call OpenEditForm(Key, ID)
    Case TB_SEARCH_PERS                                             ' Suche Person
        ID = ShowSearch(objDBconn, "Personen", "Nachname")
        If ID <> "" Then Call OpenEditForm("Personen", ID)
    Case TB_SEARCH_DOC                                              ' Suche Dokument
        ID = ShowSearch(objDBconn, "Dokumente", "Empfänger")
        If ID <> "" Then Call OpenEditForm("Dokumente", ID)
    Case TB_HELP                                                    ' Zeige Hilfe
        Call ShowHelp
    Case TB_INFO                                                    ' Zeige Info
        Call ShowAbout
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

Private Sub HandleGlobalKeyCodes(KeyCode As Integer, Shift As Integer)

    Dim szID As String                                              ' Evtl. ID des Datensatzes
    Dim szKey As String                                             ' TVNode/LV Key zur Kontext ermittlung
    
On Error GoTo Errorhandler

    'If KeyCode = 13 Then                                           ' Enter (wie Klick)
    If KeyCode = 112 Then Call ShowHelp                             ' F1    (Hilfe)
    
    If Shift = 2 Then                                               ' Strg
        If KeyCode = 70 Then                                        ' Strg + F (Suchen)
            Call GetKontextRoot(szKey, szID)                        ' Suchkontext ermitteln
            szID = ShowSearch(objDBconn, szKey, "")                 ' Suchen
            If szKey <> "" And szID <> "" Then
                Call OpenEditForm(szKey, szID, Me)                  ' DS anzeigen
            End If
        End If
        If KeyCode = 78 Then                                        ' Strg + N (Neu)
            Call GetKontextRoot(szKey, szID)                        ' Kontext ermitteln
            If szKey <> "" And szID <> "" Then
                Call OpenEditForm(szKey, "", Me)                    ' Neuen DS anzeigen
            End If
        End If
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
    Call objError.Errorhandler(MODULNAME, "HandleGlobalKeyCodes", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub HandleToolbarClick(ButtonKey As String)
    Call HandleMenueKlick(ButtonKey)                                ' Weiterreichen an Handle menueklick
End Sub

Private Sub ShowKontextMenu(Optional bTV As Boolean)
' Zeigt Kontext menue an
' bTV True -> TV sonst LV

   Dim szKeyArray() As String                                       ' Array mit Key/Tag elementen
    
On Error GoTo Errorhandler

    If bTV Then                                                     ' entsprechenden Key/Tag holen
        szKeyArray = Split(TVMain.SelectedItem.Key, TV_KEY_SEP)     ' Key von TV
    Else
        szKeyArray = Split(LVMain.Tag, TV_KEY_SEP)                  ' Key vom LV
    End If
    
    If InStr(NoKontextMenueList, szKeyArray(0)) > 0 Then GoTo exithandler  ' Prüfen ob Kontext menü gewünscht
    
    Select Case UCase(szKeyArray(0))
    Case UCase("Personen")                                          ' Personen
        PopupMenu kmnuListPersonen
    Case UCase("notare")                                            ' Notare
        PopupMenu kmnuListNotare
    Case UCase("bewerber")                                          ' Bewerber
        PopupMenu kmnuListBewerber
    Case UCase("Benutzerverwaltung")                                ' Benutzerveraltung
        PopupMenu kmnuListUser
'    Case SZ_TREENODE_USERS
'        PopupMenu kmnuLVUser
'    Case SZ_TREENODE_UNITS
'        PopupMenu kmnuLVUnit
'    Case SZ_TREENODE_REGISTERS
'        PopupMenu kmnuLVRegister
'    Case SZ_TREENODE_VERFGEGEN
'        PopupMenu kmnuLVVerfg
    Case Else
        PopupMenu kmnuListDefault                                   ' Default kontextmenü
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
    Call objError.Errorhandler(MODULNAME, "ShowKontextMenu", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub LVMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Me.MousePointer = vbHourglass                                   ' Stundenglass
    DoEvents                                                        ' Andere Events zulassen
    Call SetColumnOrder(LVMain, ColumnHeader)                       ' Spalten sortieren
    
exithandler:
On Error Resume Next
    Me.MousePointer = vbDefault                                     ' Mauszeiger wieder Normal
    DoEvents                                                        ' Andere Events zulassen
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "LVMain_ColumnClick", errNr, errDesc)
End Sub

Private Sub LVMain_DblClick()
    Call HandleLVItemDblKlick
End Sub

Private Sub TVMain_NodeClick(ByVal node As MSComctlLib.node)
    Call AddNodeToNavList(node.Key)                                 ' Knoten in Navlist aufnehmen
    Call HandleNodeClick(node)                                      ' Handling Knoten klick im TreeView
End Sub
                                                                    ' *****************************************
                                                                    ' Key Events

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Call AskForExit                            ' ESC
    Call HandleGlobalKeyCodes(KeyCode, Shift)                       ' Globale Key down ereignisse abarbeiten
End Sub

Private Sub LVMain_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim LVItem As ListItem                                          ' Akt. ListView Item

On Error GoTo Errorhandler

    Set LVItem = GetSelectListItem(LVMain)                          ' Akt. ListView Item ermitteln
    If Not LVItem Is Nothing Then
        If KeyCode = 13 Then Call HandleLVItemDblKlick              ' Enter (wie Klick)
        If KeyCode = 46 Then Call HandleLVKontextDel                ' Entf (löschen)
        If KeyCode = 27 Then Call AskForExit                        ' ESC
        If Shift = 2 Then                                           ' Strg
    
        End If
    End If
    
    Call HandleGlobalKeyCodes(KeyCode, Shift)                       ' Globale Key down ereignisse abarbeiten
    
exithandler:
On Error Resume Next
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "LVMain_KeyDown", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub TVMain_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim cnode As node                                               ' Aktueller TreeNode

On Error GoTo Errorhandler

    Set cnode = GetSelectTreeNode(TVMain)                           ' Aktueller TreeNode ermitteln
    If Not cnode Is Nothing Then
        If KeyCode = 13 Then Call HandleNodeClick(cnode)            ' Enter (wie Klick)
        'If KeyCode = 46 then                                       ' Entf (löschen)
        If KeyCode = 27 Then Call AskForExit                        ' ESC
        If Shift = 2 Then                                           ' Strg
        
        End If
    End If
    
    Call HandleGlobalKeyCodes(KeyCode, Shift)                       ' Globale Key down ereignisse abarbeiten
    
exithandler:
On Error Resume Next
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "TVMain_KeyDown", errNr, errDesc)
    Resume exithandler
End Sub
                                                                    ' *****************************************
                                                                    ' Kontext Menü Events
Private Sub kmnuListEdit_Click()
    Call HandleMenueKlick("kmnuListEdit")                           ' Menü Klick behandeln
End Sub

Private Sub kmnuListNew_Click()
    Call HandleMenueKlick("kmnuListNew")                            ' Menü Klick behandeln
End Sub

Private Sub kmnuListDel_Click()
    Call HandleMenueKlick("kmnuListDel")                            ' Menü Klick behandeln
End Sub

Private Sub kmnuListDelUser_Click()
    Call HandleMenueKlick("kmnuListDelUser")                        ' Menü Klick behandeln
End Sub

Private Sub kmnuListDelBewerber_Click()
    Call HandleMenueKlick("kmnuListDelBewerber")                    ' Menü Klick behandeln
End Sub

Private Sub kmnuListDelNotar_Click()
    Call HandleMenueKlick("kmnuListDelNotar")                       ' Menü Klick behandeln
End Sub

Private Sub kmnuListDelPerson_Click()
    Call HandleMenueKlick("kmnuListDelPerson")                      ' Menü Klick behandeln
End Sub

Private Sub kmnuListNewDocBewerber_Click()
    Call HandleMenueKlick("kmnuListNewDocBewerber")                 ' Menü Klick behandeln
End Sub

Private Sub kmnuListNewDocNotar_Click()
    Call HandleMenueKlick("kmnuListNewDocNotar")                    ' Menü Klick behandeln
End Sub

Private Sub kmnuListNewDocPerson_Click()
    Call HandleMenueKlick("kmnuListNewDocPerson")                   ' Menü Klick behandeln
End Sub

Private Sub kmnuListEditPerson_Click()
    Call HandleMenueKlick("kmnuListEditPerson")                     ' Menü Klick behandeln
End Sub

Private Sub kmnuListEditUser_Click()
    Call HandleMenueKlick("kmnuListEditUser")                       ' Menü Klick behandeln
End Sub

Private Sub kmnuListNewPerson_Click()
    Call HandleMenueKlick("kmnuListNewPerson")                      ' Menü Klick behandeln
End Sub

Private Sub kmnuListNewUser_Click()
    Call HandleMenueKlick("kmnuListNewUser")                        ' Menü Klick behandeln
End Sub
                                                                    ' *****************************************
                                                                    ' Menü Events
Private Sub mnuDateiExit_Click()
    Call HandleMenueKlick("mnuDateiExit")                           ' Menü Klick behandeln
End Sub

Private Sub mnuDateiNew_Click()
    'Call HandleMenueKlick("mnuDateiNew")                           ' Menü Klick behandeln
End Sub

Private Sub mnuDateiNewBewerber_Click()
    Call HandleMenueKlick("mnuDateiNewBewerber")                    ' Menü Klick behandeln
End Sub

Private Sub mnuDateiNewBewerbung_Click()
    Call HandleMenueKlick("mnuDateiNewBewerbung")                   ' Menü Klick behandeln
End Sub

Private Sub mnuDateiNewDoc_Click()
    Call HandleMenueKlick("mnuDateiNewDoc")                         ' Menü Klick behandeln
End Sub

Private Sub mnuDateiNewStelle_Click()
    Call HandleMenueKlick("mnuDateiNewStelle")                      ' Menü Klick behandeln
End Sub

Private Sub mnuDateiSuchen_Click()
     Call HandleMenueKlick("mnuDateiSuchen")                        ' Menü Klick behandeln
End Sub

Private Sub mnuEditAusschreibungNew_Click()
    Call HandleMenueKlick("mnuEditAusschreibungNew")                ' Menü Klick behandeln
End Sub

Private Sub mnuEditSelectTemplate_Click()
    Call HandleMenueKlick("mnuEditSelectTemplate")                  ' Menü Klick behandeln
End Sub

Private Sub mnuEditWorkflow_Click()
    Call HandleMenueKlick("mnuEditWorkflow")                        ' Menü Klick behandeln
End Sub

Private Sub mnuExtrasChangePWD_Click()
    Call HandleMenueKlick("mnuExtrasChangePWD")                     ' Menü Klick behandeln
End Sub

Private Sub mnuExtrasOptions_Click()
    Call HandleMenueKlick("mnuExtrasOptions")                       ' Menü Klick behandeln
End Sub

Private Sub mnuInfoAbout_Click()
    Call HandleMenueKlick("mnuInfoAbout")                           ' Menü Klick behandeln
End Sub

Private Sub mnuInfoHelp_Click()
    Call HandleMenueKlick("mnuInfoHelp")                            ' Menü Klick behandeln
End Sub
                                                                    ' *****************************************
                                                                    ' Toolbar Events
Private Sub ToolbarMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call HandleMenueKlick(Button.Key)                               ' Menü Klick behandeln
End Sub

Private Sub ToolbarMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Call HandleMenueKlick(ButtonMenu.Key)                           ' Menü Klick behandeln
End Sub
                                                                    ' *****************************************
                                                                    ' Mouse Events
Private Sub LVMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = 2 Then Call ShowKontextMenu                        ' Kontextmenü anzeigen
End Sub

Private Sub LVMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = vbDefault                                     ' Mauszeiger wieder Normal
End Sub

Private Sub StatusBarMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = vbDefault                                     ' Mauszeiger wieder Normal
End Sub

Private Sub ToolbarMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = vbDefault                                     ' Mauszeiger wieder normal
End Sub

Private Sub TVMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = vbDefault                                     ' Mauszeiger wieder normal
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SplitFlag = True                                                ' Verschieben des Splitters akttivieren
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
On Error GoTo Errorhandler

    Me.MousePointer = vbSizeWE                                      ' Mauszeiger für Größen änderung
    If SplitFlag Then                                               ' Wenn Spliter Verschoben wird
        curlngSplitposProz = X / ScaleWidth                         ' Neue Pos bestimmen
        If curlngSplitposProz < 0.065 Then curlngSplitposProz = 0.065   ' Min Pos prüfen
        If curlngSplitposProz > 0.94 Then curlngSplitposProz = 0.94 ' Max Pos prüfen
    End If
    
exithandler:
On Error Resume Next
    Call RepaintMainForm(curlngSplitposProz)                        ' Form neu arangieren
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "Form_MouseMove", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
On Error GoTo Errorhandler

    If SplitFlag Then                                               ' Wenn Spliter Verschoben wird (wurde)
        SplitFlag = False                                           ' Verschieben beendet
        Call RepaintMainForm(curlngSplitposProz)                    ' Form neu arangieren
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
    Call objError.Errorhandler(MODULNAME, "Form_MouseUp", errNr, errDesc)
End Sub

Private Sub Form_Resize()
    Call RepaintMainForm(curlngSplitposProz)                        ' Form neu arangieren
End Sub
                                                                    ' *****************************************
                                                                    ' Properties
Public Property Get GetDBConn() As Object
    Set GetDBConn = ThisDBCon
End Property

Public Property Set SetDBConn(DBCon As Object)
    Set ThisDBCon = DBCon
End Property

