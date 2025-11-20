VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11685
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   11685
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox PicNavBar 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'Kein
      Height          =   255
      Left            =   3000
      ScaleHeight     =   255
      ScaleWidth      =   5295
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   5295
      Begin VB.Label lblNavBarArrow 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblNavBar 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   0
         Left            =   360
         MouseIcon       =   "frmMain.frx":223E
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   7
         Top             =   0
         Width           =   585
      End
   End
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
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":239C
            Key             =   ""
            Object.Tag             =   "Word"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2736
            Key             =   ""
            Object.Tag             =   "Help"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AD0
            Key             =   ""
            Object.Tag             =   "Info"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E6A
            Key             =   ""
            Object.Tag             =   "Search"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3404
            Key             =   ""
            Object.Tag             =   "Refresh"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":379E
            Key             =   ""
            Object.Tag             =   "Back"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B38
            Key             =   ""
            Object.Tag             =   "Forward"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3ED2
            Key             =   ""
            Object.Tag             =   "Add"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":426C
            Key             =   ""
            Object.Tag             =   "Print"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolbarMain 
      Align           =   1  'Oben ausrichten
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ILToolBar"
      DisabledImageList=   "ILToolBarDis"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
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
               NumButtonMenus  =   5
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
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbNewAusschreibung"
                  Text            =   "Neue Ausschreibung"
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
            Key             =   "tbPrint"
            Object.ToolTipText     =   "Drucken"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbDocNew"
            Object.ToolTipText     =   "Neues Dokument"
            Object.Tag             =   "tbDocNew"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbHelp"
            Object.ToolTipText     =   "Hilfe zur Notarverwaltung"
            Object.Tag             =   "tbHelp"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbInfo"
            Object.ToolTipText     =   "Info zur Notarverwaltung"
            Object.Tag             =   "tbInfo"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicList 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   3735
      Left            =   3120
      ScaleHeight     =   3735
      ScaleWidth      =   5295
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   5295
      Begin MSComctlLib.ListView LVMain 
         Height          =   2895
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   5106
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
            NumListImages   =   28
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4806
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4B20
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4E3A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":53D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":596E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6648
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7322
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":78BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7E56
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":81F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":858A
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8924
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8CBE
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":9258
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":97F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":9D8C
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":A326
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":A8C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":AE5A
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":B3F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":C49E
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":CA38
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":CFD2
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":D56C
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":DB06
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":DEA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":E43A
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":E9D4
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
      Top             =   6540
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Picture         =   "frmMain.frx":ED6E
            Object.ToolTipText     =   "Heutiges Datum"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Picture         =   "frmMain.frx":F308
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
            Picture         =   "frmMain.frx":F6A2
            Object.ToolTipText     =   "Aktueller Datenbankserver"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   5662
            MinWidth        =   5292
            Object.ToolTipText     =   "Aktuell anzeigt"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   2134
            MinWidth        =   1764
            Object.ToolTipText     =   "Anzahl der angezeigten Datensätze"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ILToolBarDis 
      Left            =   3600
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FA3C
            Key             =   ""
            Object.Tag             =   "Word"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FDD6
            Key             =   ""
            Object.Tag             =   "Help"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10170
            Key             =   ""
            Object.Tag             =   "Info"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1050A
            Key             =   ""
            Object.Tag             =   "Search"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10AA4
            Key             =   ""
            Object.Tag             =   "Refresh"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":111D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11572
            Key             =   ""
            Object.Tag             =   "Add"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1190C
            Key             =   ""
            Object.Tag             =   "Print"
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
         Begin VB.Menu mnuDateiNewAusschreibung 
            Caption         =   "Neue &Ausschreibung"
         End
         Begin VB.Menu mnuDateiNewDoc 
            Caption         =   "Neues &Dokument"
         End
      End
      Begin VB.Menu mnuDateiSuchen 
         Caption         =   "&Suchen"
      End
      Begin VB.Menu mnuDateiPrint 
         Caption         =   "&Drucken"
         Shortcut        =   ^P
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
      Begin VB.Menu mnuEditExcelExport 
         Caption         =   "Excel &Export"
      End
      Begin VB.Menu mnuEditDocImport 
         Caption         =   "Dokument &Import"
      End
      Begin VB.Menu mnuEditVerzImport 
         Caption         =   "Verzeichnis Import"
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
         Caption         =   "Inf&o"
      End
      Begin VB.Menu mnuInfoHelp 
         Caption         =   "Hilfe"
      End
      Begin VB.Menu mnuInfoReadMe 
         Caption         =   "ReadMe"
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
      Begin VB.Menu kmnuListPrint 
         Caption         =   "Drucken"
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
      Begin VB.Menu kmnuListPrintPerson 
         Caption         =   "Liste Drucken"
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
      Begin VB.Menu kmnuListPrintNotar 
         Caption         =   "Liste Drucken"
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
      Begin VB.Menu kmnuListPrintBewerber 
         Caption         =   "Liste Drucken"
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
      Begin VB.Menu kmnuListPrintUser 
         Caption         =   "Liste Drucken"
      End
   End
   Begin VB.Menu kmnuListAusschreibung 
      Caption         =   "KontextAusschreibung"
      Visible         =   0   'False
      Begin VB.Menu kmnuListNewAusschreibung 
         Caption         =   "Neue Ausschreibung anlegen"
      End
      Begin VB.Menu kmnuListEditAusschreibung 
         Caption         =   "Ausschreibung bearbeiten"
      End
      Begin VB.Menu kmnuListDelAusschreibung 
         Caption         =   "Ausschreibung Löschen"
      End
      Begin VB.Menu kmnuListPrintAusschreibung 
         Caption         =   "Liste Drucken"
      End
   End
   Begin VB.Menu kmnuListBewerbung 
      Caption         =   "KontextBewerbung"
      Visible         =   0   'False
      Begin VB.Menu kmnuListNewBewerbung 
         Caption         =   "Neue Bewerbung eintragen"
      End
      Begin VB.Menu kmnuListEditBewerbung 
         Caption         =   "Bewerbung bearbeiten"
      End
      Begin VB.Menu kmnuListDelBewerbung 
         Caption         =   "Bewerbung löschen"
      End
      Begin VB.Menu kmnuListEditBewerberPerson 
         Caption         =   "Personendaten anzeigen"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                                         ' Variaben Deklaration erzwingen
Const MODULNAME = "frmMain"                                             ' Modulname für Fehlerbehandlung

Private curlngSplitposProz As Single                                    ' Aktueller wert der splitter pos.
Private SplitFlag As Boolean                                            ' True wenn List bzw. TreeView größe verändert wird
Private NavBarLinkHover As Boolean                                      ' True wenn in der Navbar ein LinkLable behovert wird
Private EditFormArray() As Object                                       ' Auflistung aller geöffneten Edit Formulare
Private NavTreeNodeArray() As String                                    ' Array enthält alle selectierten nodes (max 20?)
Private NoKontextMenueList As String                                    ' Liste der Listviews/TreeViewe elemente ohne kontextmenue
Private lngNavIndex As Integer                                          ' Aktuelle pos im nav Array

Private ThisDBCon As Object                                             ' Diese Datenbank verbindung

Private Sub Form_Load()
    Dim objError As Object                                              ' Error object
    Dim cNode As node                                                   ' Tree Node Object
    Dim szMsgText As String                                             ' Meldungtext
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    Set objError = objObjectBag.GetErrorObj()                           ' Error Ob für Fehlerbehandlung Initialisieren
    Call InitMainForm                                                   ' Hauptform initialisieren
    If CheckTodayDeadlins(ThisDBCon) Then                               ' Prüfen ob heute Fristen fällig sind
        szMsgText = "Es sind Heute fristen fällig. Möchten Sie zu den heutigen Fristen wechseln?" ' Meldungstext Festlegen
        If objError.ShowErrMsg(szMsgText, vbYesNo + vbQuestion, _
                "Fällige Fristen", , , Me) = vbYes Then                 ' Wenn Antwort Ja
            Set cNode = GetNodeByKey(Me.TVMain, "Fristen\Heute")        ' entsprechenden TreeNode suchen
            If Not cNode Is Nothing Then                                ' Wenn Tree Node existiert
                Call SelectTreeNode(Me.TVMain, cNode)                   ' Node auswählen
                Call HandleNodeClick(cNode, False)                      ' Node Klick ausführen
                GoTo exithandler                                        ' Fertig
            End If
        End If
    End If
exithandler:
Exit Sub                                                                ' Function Beenden
Errorhandler:
    Dim errNr As String                                                 ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "Form_Load", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren (hat eh keinen Zweg mehr)
                                                                        ' Optionen des Hauptforms Speichern
    Call objOptions.SetOptionByName(OPTION_MAINSTATE, _
            Me.WindowState)                                             ' WindowState in optionen Speichern
    If frmMain.WindowState = vbNormal Then                              ' Wenn WindowState Normal
        Call objOptions.SetOptionByName(OPTION_MAINSIZE, _
                Me.Width & "/" & Me.Height)                             ' WindowSize in Optionen speichern
    End If
    Call objOptions.SetOptionByName(OPTION_LASTNODE, _
            Me.TVMain.SelectedItem.Key)                                 ' Akt Treenode in Optionen Speichern
    Call objOptions.SetOptionByName(OPTION_SPLIT, _
            curlngSplitposProz)                                         ' Spliter pos in optionen speichern
    Call AppExit                                                        ' Application beenden
End Sub

Private Function InitMainForm()
    Dim szSize As String                                                ' SizeValue aus Reg
    Dim szSizeArray() As String                                         ' (0) = Width , (1) = Height
    Dim szWinState As String                                            ' Window State (min, max, normal)
    Dim szSplitpos As String                                            ' Pos des Splitters
    Dim szTranzRate As String
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    Me.Caption = SZ_APPTITLE                                            ' Form Caption Setzten
    If objObjectBag.getexpert Then                                      ' Sind Wir im Experten Modus
        Me.Caption = SZ_APPTITLE & " " & SZ_EXPERT                      ' Caption erweitern
' MW 01.09.11 {
'        szTranzRate = objOptions.GetOptionByName(OPTION_TRANZ)          ' Wert Tranzparenz aus Options lesen
'        If szTranzRate <> "" Then                                       ' Nur wenn wert vorhanden
'            If CSng(szTranzRate) < 1 Then                               ' Nur wenn was zu tun ist
'                Call SetWindowTransparency(Me, CSng(szTranzRate))       ' Form Transparent setzen
'            End If
'        End If
' MW 01.09.11 }
    End If
                                                                        ' Option aus lesen
    szSplitpos = objOptions.GetOptionByName(OPTION_SPLIT)               ' Spliter pos
    If szSplitpos <> "" Then                                            ' Spliter pos. vorhanden
        curlngSplitposProz = CSng(szSplitpos)                           ' Spliter Pos setzen
    Else
        curlngSplitposProz = 0.3                                        ' Default Spliter pos setzen
    End If
    szSize = objOptions.GetOptionByName(OPTION_MAINSIZE)                ' Option WindowSize auslesen
    Call SetWindowSizeFromString(Me, szSize)                            ' WindowSize setzen
    szWinState = objOptions.GetOptionByName(OPTION_MAINSTATE)           ' Option WindowState auslesen
    Call SetWindowStateFromString(Me, szWinState)                       ' Windowstate setzen
    Call RepaintMainForm(curlngSplitposProz)                            ' Form neuzeichnen
    Call InitStatusBarMain                                              ' Statusbar initialisieren
    Call InitLV                                                         ' ListView initialisieren
    Call InitTree                                                       ' TreeView initialisieren
    Call objObjectBag.CheckFormStyle(Me)
    'Call InitGrid
'    Call objObjectBag.CheckFormStyle(Me)
exithandler:
On Error Resume Next
    Me.Refresh
    Call objObjectBag.ShowMSGForm(False, "")                            ' MSG Form ausblenden
    Call ShowSplash(False)                                              ' Splash ausblenden
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
    Dim i As Integer                                                    ' Counter
    Dim cNode As node                                                   ' Aktueller Node
    Dim szTmpNodeName As String                                         ' Temp Node Name
    Dim RootNodeArray() As String                                       ' Array mit RootNodenamen
    Dim szRootNodeList As String                                        ' Nodelist als String
    Dim TVNode As TreeViewNodeInfo                                      ' TreeNode Informationen
    Dim szLastNode As String                                            ' Lezter Node Als String aus Reg
    Dim bStartWithlastNode As Boolean                                   ' Möchte der Anwender auf dem Letzen Knoten starten
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    Call MousePointerHourglas(Me)                                       ' Sanduhr anzeigen
    szRootNodeList = objTools.GetRootNodeListFromXML(App.Path _
            & "\" & INI_XMLFILE)                                        ' Liste der Rootknoten holen
    RootNodeArray = Split(szRootNodeList, ";")                          ' Rootnode list in array
    For i = 0 To UBound(RootNodeArray)                                  ' Alle RootNodes durchlaufen
        szTmpNodeName = RootNodeArray(i)                                ' Akt Nodename holen
        If Not User.System Then                                         ' Wenn NichtUser systemverwalter
            If (szTmpNodeName = "Stammdaten") Or (szTmpNodeName = _
                    "Benutzerverwaltung") Then                          ' Konten Stammdaten und benutzerverwaltung ausblenden
                GoTo Skip                                               ' Rest überspringen
            End If
        End If
        With TVNode
            Call objTools.GetTVNodeInfofromXML(App.Path & "\" & INI_XMLFILE, szTmpNodeName, _
                    .szTag, .szText, .szKey, .bShowSubnodes, .szSQL, _
                    .szWhere, .lngImage)                                ' Node daten holen
            If .szTag <> "" And .szKey <> "" And .szText <> "" Then     ' Wenn Tag und Key und Textvorhanden
                'If .lngImage = "" Then .lngImage = "1"                ' Gegebenenfalls Default image holen
                Call AddTreeNode_New(TVMain, "", .szKey, .szTag, .szText, ThisDBCon, _
                        CLng(.lngImage), Not .bShowSubnodes)            ' Neuen Node anlegen
            End If
        End With
Skip:
    Next i                                                              ' Nächster Rootnode
'    ' Liste der Einträge ohne Kontextmenu ermitteln (für LV und TV)
'    NoKontextMenueList = objTools.GetNoKontextListFromXML(App.Path & "\" & INI_XMLFILE)
    bStartWithlastNode = objOptions.GetOptionByName(OPTION_STARTLASTNODE) ' auf den letzten node springen?
    If bStartWithlastNode Then
        szLastNode = objOptions.GetOptionByName(OPTION_LASTNODE)        ' Letzten Konten aus Reg holen
        If szLastNode <> "" Then                                        ' Wenn Reg Wert vorhanden
            Set cNode = GetNodeByKey(TVMain, szLastNode)                ' entsprechenden TreeNode suchen
            If Not cNode Is Nothing Then                                ' Wenn Tree Node existiert
                Call SelectTreeNode(TVMain, cNode)                      ' Node auswählen
                Call HandleNodeClick(cNode, False)                      ' Node Klick ausführen
                GoTo exithandler                                        ' Fertig
            End If
        End If
    End If
    
    Set cNode = GetNodeByKey(TVMain, SZ_TREENODE_MAIN)                  ' Sonst Ersten Konten auswählen
    Call SelectTreeNode(TVMain, cNode)                                  ' Node auswählen
    Call HandleNodeClick(cNode, False)                                  ' Node Klick ausführen
        
exithandler:
On Error Resume Next
    Me.TVMain.Enabled = True                                            ' TV Enablen
    Call MousePointerDefault(Me)                                        ' Mauszeiger wieder normal
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
    Call ClearLV(Me.LVMain)                                             ' evtl. List Items Löschen
    LVMain.Icons = ILTree                                               ' Verweis auf Image List
    LVMain.SmallIcons = ILTree
    'Call ShowLV                                                        ' Zuerst ListView anzeigen, Grid ausblenden
exithandler:
    
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
    
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    
    StatusBarMain.Panels(1).Alignment = sbrCenter
    StatusBarMain.Panels(1).Text = Left(CStr(Now()), 10)                ' Datum anzeigen
    
    StatusBarMain.Panels(2).Alignment = sbrLeft
    StatusBarMain.Panels(2).Text = User.Username                        ' Angemelderter User
    
    StatusBarMain.Panels(3).Alignment = sbrLeft
    If objObjectBag.bUserIsAdmin Then                                   ' Admin
        StatusBarMain.Panels(3).Text = "(Admin)"
    Else
        StatusBarMain.Panels(3).Text = "(Benutzer)"
    End If
        
    StatusBarMain.Panels(4).Alignment = sbrLeft
    StatusBarMain.Panels(4).Text = objDBconn.getDBtext                  ' Db info
    
    StatusBarMain.Panels(5).Alignment = sbrRight                        ' Listitems Count

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
' Löscht DS nach nachfrage oder ruft spez. Losch fkt auf
    Dim szSQL As String                                                 ' SQL Statement
    Dim szMSG As String                                                 ' Meldungtest Für User Nachfrage
    Dim szValue As String                                               ' DS Kennung (z.b. Name) damit der User weiss was er löscht
    Dim szTitle As String                                               ' Meldungs titel
    Dim szDetails As String                                             ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                              ' Fehlerbehandlung löschen
    If szID = "" Or szRootkey = "" Then GoTo exithandler                ' Keine DS ID -> fertig
    szDetails = "DS ID: " & szID & vbCrLf
    Call MousePointerHourglas(Me)                                       ' Sanduhr anzeigen
    Select Case UCase(szRootkey)
    Case UCase("Ausschreibung"), UCase("Ausschreibungen")               ' Ausschreibung Löschen
        szSQL = "SELECT AZ020 + ' (' + CAST(Jahr020 as Varchar(4)) + ')' " & _
                "FROM AUSSCHREIBUNG020 WHERE ID020 ='" & szID & "'"     ' AZ für Meldung ermitteln
        szValue = objDBconn.GetValueFromSQL(szSQL)
        szMSG = "Möchten Sie die Ausschreibung " & szValue & " und alle dazu Ausgeschriebenen Stellen" & _
                ", sowie alle Bewerber wirklich löschen?"               ' Meldungstext Festlegen
        szTitle = "Ausschreibung löschen"                               ' Meldungstitel festlegen
        szSQL = " DELETE BEWERB013 FROM BEWERB013 INNER JOIN STELLEN012 ON ID012 = FK012013 INNER JOIN AUSSCHREIBUNG020 " & _
                " ON FK020012 = ID020 WHERE ID020='" & szID & "' "      ' Bewerbeungen Löschen
        szSQL = szSQL & " DELETE FROM STELLEN012 WHERE FK020012='" _
                & szID & "' "                                           ' Stellen Löschen
        szSQL = szSQL & " DELETE FROM AUSSCHREIBUNG020 WHERE ID020 =    '" _
                & szID & "' "                                       ' Ausschreibung löschen
        ' FK012018 in DOC018 Updaten
        
    Case UCase("Personen"), UCase("Bewerber"), UCase("Notare"), UCase("Personenkartei")
        Call DeletePerson(szID)                                         ' Detail Daten und Dokumente löschen daher sonderwurst
        GoTo exithandler                                                ' Fetig
    Case UCase("Fortbildungen")                                         ' Fortbildungs DS löschen
        szSQL = "SELECT Thema011 FROM FORT011 WHERE ID011 ='" _
                & szID & "'"                                            ' Fortbildungen löschen
        szValue = objDBconn.GetValueFromSQL(szSQL)
        szMSG = "Möchten Sie die Fortbildung " & szValue & _
                " wirklich löschen?"                                    ' Meldungstext Festlegen
        szTitle = "Fortbildung löschen"                                 ' Meldungstitel festlegen
        szSQL = "DELETE FROM FORT011 WHERE ID011 ='" & szID & "' "
        szSQL = szSQL & " DELETE FROM AFORT014 WHERE FK011014='" & szID & "' "
        
    Case UCase("Ausgeschriebene Stellen"), UCase("Stellen"), _
            UCase("StellenJahr")                                        ' Stellen DS löschen
        szSQL = "SELECT Bezirk012 + ' ' + Cast(Frist012 as varchar(20)) " & _
                "FROM STELLEN012 WHERE ID012 ='" & szID & "'"
        szValue = objDBconn.GetValueFromSQL(szSQL)
        szMSG = "Möchten Sie die Ausgeschriebene Stelle " & szValue & " und alle " & _
                "eingetragenen Bewerbungen wirklich löschen?"           ' Meldungstext Festlegen
        szTitle = "Ausgeschriebene Stelle löschen"                      ' Meldungstitel festlegen
        szSQL = " DELETE FROM BEWERB013 WHERE FK012013='" _
                & szID & "' "                                           ' Und Alle BewerbungsDS löschen
        szSQL = szSQL & " DELETE FROM STELLEN012 WHERE ID012 ='" _
                & szID & "'"                                            ' Alle Stellen Daten Löschen
        
        ' FK012018 in DOC018 Updaten
' MW 10.01.11 Fristen Löschen {
    Case UCase("Fristen"), UCase("Heute"), UCase("Abgelaufen")          ' Fristen
        szSQL = "SELECT Frist024.Frist024 " _
                & "FROM Frist024 WHERE ID024 = '" & szID & "'"
        szValue = objDBconn.GetValueFromSQL(szSQL)
        szSQL = "SELECT ISNULL(Nachname010,'') + ', ' + ISNULL(Vorname010,'') " _
                & "FROM Frist024 LEFT JOIN RA010 ON FK010024 = ID010 " _
                & " WHERE ID024 = '" & szID & "'"
        szValue = szValue & " für " & objDBconn.GetValueFromSQL(szSQL)
        szMSG = "Möchten Sie die Frist am " & szValue & _
                " wirklich löschen?"                                    ' Meldungstext Festlegen
        szSQL = "DELETE FROM Frist024 WHERE ID024 ='" & szID & "' "
' MW 10.01.11 }
    Case UCase("Bewerbung"), UCase("Bewerbungen")                       ' Bewerbungs DS Löschen
        szSQL = "SELECT ISNULL(Nachname010,'') + ', ' + ISNULL(Vorname010,'')  FROM BEWERB013 Left Join RA010 ON FK010013 = ID010 " & _
                " WHERE ID013 ='" & szID & "'" '"CAST('" & ID & "' as uniqueidentifier)"
        szValue = objDBconn.GetValueFromSQL(szSQL)
        szSQL = "SELECT BEZIRK012  FROM BEWERB013 Left Join STELLEN012 ON FK012013 = ID012 " & _
                " WHERE ID013 ='" & szID & "'" '"CAST('" & ID & "' as uniqueidentifier)"
        szValue = szValue & " im Bezirk " & objDBconn.GetValueFromSQL(szSQL)
        szMSG = "Möchten Sie die Bewerbung von " & szValue & _
                " wirklich löschen?"                                    ' Meldungstext Festlegen
        szTitle = "Bewerbung löschen"                                   ' Meldungstitel festlegen
        szSQL = "DELETE FROM BEWERB013 WHERE ID013 ='" & szID & "' "
        szSQL = szSQL & " DELETE FROM BEWERB013 WHERE FK012013='" & szID & "' "
        ' FK012018 in DOC018 Updaten
        
    Case UCase("Stammdaten")                                            ' Stammdaten Löschen
        
    Case UCase("Benutzerverwaltung"), UCase("Benutzer")                 ' User Datensatz Löschen
    ' Vorher nach User fragen der die Daten Übernimmt und damit CFrom & MFrom felder Updaten
        szSQL = "SELECT USERNAME001 FROM USER001 WHERE ID001 ='" & szID & "'" '"CAST('" & ID & "' as uniqueidentifier)"
        szValue = objDBconn.GetValueFromSQL(szSQL)
        szMSG = "Möchten Sie den Benutzer " & szValue & _
                " wirklich löschen?"                                    ' Meldungstext fetlegen
        szTitle = "Benutzer löschen"                                    ' Meldungstitel festlegen
        ' Hier evtl 2. Fragen
        szSQL = "DELETE FROM USER001 WHERE ID001 ='" & szID & "' "      ' Benutzer löschen
        
    Case UCase("Dokumente"), UCase("Letzte Woche"), _
            UCase("Letzter Monat")                                      ' Dokument DS Löschen
        Call DeleteDokument(szID)                                       ' Dokument auch im Filesystem löschen dewegen eigene Fkt
        GoTo exithandler                                                ' Fertig
        
    Case UCase("Aktenort")                                              ' Akten Ort Löschen
        ' bei akten ort kein löschen vorgesehen
    Case UCase("Amtsgerichte")
        szSQL = "SELECT AGNAME004 FROM AG004 WHERE ID004 ='" & szID & "'" '"CAST('" & ID & "' as uniqueidentifier)"
        szValue = objDBconn.GetValueFromSQL(szSQL)
        szMSG = "Möchten Sie das Amtsgericht " & szValue & _
                " wirklich löschen?"                                    ' Meldungstext fetlegen
        szTitle = "Gericht löschen"                                     ' Meldungstitel festlegen
        szSQL = "DELETE FROM AG004 WHERE ID004 ='" & szID & "' "        ' AG  löschen
    Case UCase("Landgerichte")
        szSQL = "SELECT LGNAME003 FROM LG003 WHERE ID003 ='" & szID & "'" '"CAST('" & ID & "' as uniqueidentifier)"
        szValue = objDBconn.GetValueFromSQL(szSQL)
        szMSG = "Möchten Sie das Landgericht " & szValue & _
                " wirklich löschen?"                                    ' Meldungstext fetlegen
        szTitle = "Gericht löschen"                                     ' Meldungstitel festlegen
        szSQL = "DELETE FROM LG003 WHERE ID003 ='" & szID & "' "        ' AG  löschen
    Case UCase("Oberlandesgerichte")
    
    Case Else
    
    End Select
    
    If szValue = "" Then GoTo exithandler                               ' Kein DS gefunden -> Raus
    szDetails = szDetails & "Wert: " & szValue & vbCrLf
    
    If objError.ShowErrMsg(szMSG, vbOKCancel + vbQuestion, szTitle) <> vbCancel Then
        Call objDBconn.execSql(szSQL)                                   ' Delete Statement ausführen
        Call RefreshTreeView                                            ' Tree & List View aktuaisieren
    End If
    
exithandler:
On Error Resume Next
    Call MousePointerDefault(Me)                                    ' Maus Zeiger wieder normal
        
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
    Dim szSQL As String                                                 ' SQL Statement
    Dim rsDoc As ADODB.Recordset                                        ' RS mit Dokument DS
    Dim szMSG As String                                                 ' Meldungtest Für User Nachfrage
    Dim szValue As String                                               ' DS Kennung (z.b. Name) damit der User weiss was er löscht
    Dim szTitle As String                                               ' Meldungs titel
    Dim szPath As String                                                ' Dokumenten Pfad
    Dim szAblagePath As String                                          ' Pfad der Datei ablage
    Dim szDetails As String                                             ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    If szPersID = "" Then GoTo exithandler                              ' Ohne Pers ID -> fertig
    szDetails = "PersID: " & szPersID
    
    szAblagePath = objOptions.GetOptionByName(OPTION_ABLAGE) & "\"      ' Ablage Pfad aus optionen holen
    
    ' Pers Name für Nachfrage Holen
    szSQL = "SELECT Nachname010 + ', ' + ISNULL(Vorname010,'') FROM RA010 WHERE ID010 ='" & szPersID & "'" '"CAST('" & ID & "' as uniqueidentifier)"
    szValue = objDBconn.GetValueFromSQL(szSQL)
    szMSG = "Möchten Sie die Person " & szValue & " wirklich löschen?"
    szTitle = "Person löschen"
    ' Lösch SQL Staements Festlegen
    szSQL = "DELETE FROM AFORT014 WHERE FK010014 ='" & szPersID & "' "
    szSQL = szSQL & " DELETE FROM BEWERB013 WHERE FK010013='" & szPersID & "'"
    szSQL = szSQL & " DELETE FROM Frist024 WHERE FK010024 ='" & szPersID & "' "
    szSQL = szSQL & " DELETE FROM DOC018 WHERE FK010018='" & szPersID & "'"
    szSQL = szSQL & " DELETE FROM FORD022 WHERE FK010022 ='" & szPersID & "'"
    szSQL = szSQL & " DELETE FROM AKTENORT017 WHERE FK010017 ='" & szPersID & "'"
    szSQL = szSQL & " DELETE FROM RA010 WHERE ID010 ='" & szPersID & "'"
    
    If szValue = "" Then GoTo exithandler                               ' Keine Person gefunden -> fertig
    szDetails = "PersID: " & szPersID & vbCrLf & "Name: " & szValue
    szPath = objOptions.GetOptionByName(OPTION_ABLAGE) & "\" & szPath
    If objError.ShowErrMsg(szMSG, vbOKCancel + vbQuestion, szTitle) <> vbCancel Then
    
        ' Erst Dok Echt im Verz löschen
        Set rsDoc = ThisDBCon.fillrs("SELECT * FROM DOC018 WHERE FK010018='" & szPersID & "'")
        If Not rsDoc Is Nothing Then
            If rsDoc.RecordCount > 0 Then rsDoc.MoveFirst
            While Not rsDoc.EOF                                         ' Alle Docs durchlaufen
                szPath = rsDoc.Fields("DOCPATH018").Value
                szPath = szAblagePath & szPath                          ' Pfad zusammen setzen
                If objTools.FileDelete(szPath, True) Then               ' Wenn Doc gelöscht
                    'Stop        ' zum Debuggen
                End If
                rsDoc.MoveNext
            Wend
        End If
        If objDBconn.execSql(szSQL) Then                                ' Dann Pers eintrag in Tabelle löschen
            Call RefreshTreeView                                        ' Tree & List View aktuaisieren
            'Call RefreshListView(LVMain, TVMain)                        ' Listview aktualisieren
        End If
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
    Dim szSQL As String                                                 ' SQL Statement
    Dim szMSG As String                                                 ' Meldungtest Für User Nachfrage
    Dim szValue As String                                               ' DS Kennung (z.b. Name) damit der User weiss was er löscht
    Dim szTitle As String                                               ' Meldungs titel
    Dim szPath As String                                                ' Dokumenten Pfad
    Dim szDetails As String                                             ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                              ' Fehler behandlung aktivieren
    If szDocID = "" Then GoTo exithandler                               ' Keine Doc ID -> Fertig
    szDetails = "DocID: " & szDocID & vbCrLf
    szSQL = "SELECT DOCNAME018 FROM DOC018 WHERE ID018 ='" & szDocID & "'" '"CAST('" & ID & "' as uniqueidentifier)"
    szValue = objDBconn.GetValueFromSQL(szSQL)                          ' Dokumenten namen ermitteln
    szDetails = szDetails & "DocName: " & szValue & vbCrLf
    szSQL = "SELECT DOCPATH018 FROM DOC018 WHERE ID018 ='" & szDocID & "'" '"CAST('" & ID & "' as uniqueidentifier)"
    szPath = objDBconn.GetValueFromSQL(szSQL)                           ' Dokumenten Pfad ermitteln
    szDetails = szDetails & "DocPath: " & szValue & vbCrLf
    szMSG = "Möchten Sie das Dokument " & szValue & " wirklich löschen?"
    szTitle = "Dokument löschen"
    szSQL = "DELETE FROM DOC018 WHERE ID018 ='" & szDocID & "' "
    If szValue = "" Or szPath = "" Then GoTo exithandler                ' Kein Doc gefunden -> Fertig
    szPath = objOptions.GetOptionByName(OPTION_ABLAGE) & "\" & szPath
    
    If objError.ShowErrMsg(szMSG, vbOKCancel + vbQuestion, szTitle) <> vbCancel Then ' Nachfragen
        ' Erst Dok Echt im Verz löschen
        If objTools.FileDelete(szPath, True, True) Then                 ' Wenn Doc (imVerz.) gelöscht
            'Stop        ' zum Debuggen
        End If
       If objDBconn.execSql(szSQL) Then                                 ' Dann Doc eintrag in Tabelle löschen
            Call RefreshTreeView                                        ' Tree & List View aktuaisieren
       End If
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
        Optional parentform As Form, Optional bDialog As Boolean)
' Öffnet leeres form für Neuen DS
On Error GoTo Errorhandler
    Call OpenEditForm(szRootkey, "", parentform, bDialog)               ' OpenEditForm ohne DetailKey (DS ID) aufrufeb
exithandler:
On Error Resume Next
    Call MousePointerDefault(Me)                                        ' Mousepointer Normal
    
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
        Optional parentform As Form, Optional bDialog As Boolean) As String
' Öffnet ein DS Form (edit) zum anzeigen und bearbeiten von DS
    Dim NewFrmEdit As Form                                              ' Neues DS (edit) Form
    Dim i As Integer                                                    ' Counter
    Dim lngEditFormCount  As Integer                                    ' Anzahl der Edit Forms
    Dim ID As String                                                    ' evtl. DS ID
    Dim Detailarray() As String                                         ' Array aus evtl zusammegestzten IDs
    Dim szTmpImageIndex As String                                       ' Image Index des Edit Forms
    Dim PersID As String
    Dim StellenID As String
    Dim AusschrID As String
On Error Resume Next                                                    ' Hier erstmal keine Fehlerbehndlung
    lngEditFormCount = UBound(EditFormArray)                            ' Anz. Edit form ermitteln
    If Err.Number <> 0 Then                                             ' Errorhandling Deak. da Array evtl. leer
        lngEditFormCount = -1
        Err.Clear                                                       ' Fehler Resetten
    End If
On Error GoTo Errorhandler                                              ' Fehlerbehandlung wieder aktivieren
    If InStr(DetailKey, ";") > 0 Then                                   ' ID Zusammengesetzt
        Detailarray = Split(DetailKey, ";")                             ' Wenn JA aufspalten
        ID = Detailarray(0)
    Else                                                                ' Sonst
        ID = DetailKey
    End If
    If ID <> "" And lngEditFormCount > -1 Then                          ' ID und offenes Edit Form vorhanden ?
        For i = 0 To lngEditFormCount                                   ' Durch FormsArray laufen und überprüfen ob schon ein mit ID offen ist
            If Not EditFormArray(i) Is Nothing Then
                If EditFormArray(i).ID = ID Then                        ' Prüfen ob gleiches Edit form erneut geöffnet werden soll
                    EditFormArray(i).Show                               ' Wenn Ja anzeigen
                    GoTo exithandler                                    ' Fertig
                End If
            End If
        Next i                                                          ' Nächstes Form Array item
    End If
On Error Resume Next                                                    ' Fehler behandlung deaktivieren
    ReDim Preserve EditFormArray(UBound(EditFormArray) + 1)             ' Sonst Form Array erweitern
    If Err.Number <> 0 Then                                             ' Wenn Fehler
        ReDim EditFormArray(0)                                          ' 1. Form
        Err.Clear                                                       ' Fehler Resetten
    End If
On Error GoTo Errorhandler                                              ' Fehlerbehandlung wieder aktivieren
    ' Und neues form öffnen
    Select Case UCase(szRootkey)                                        ' Form aus Rootkey ermitteln
    Case UCase("Personen"), UCase("Bewerber"), UCase("Teilnehmer"), _
                UCase("Notare"), UCase("Notare bestellt"), UCase("Notare ausgeschieden")
        Set NewFrmEdit = New frmEditPersonen
        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey, bDialog)
        If DetailKey = "" Then
            If UCase(szRootkey) = "BEWERBER" Then NewFrmEdit.cmbStatus.Text = "Bewerber"
            If Left(UCase(szRootkey), 5) = "NOTAR" Then NewFrmEdit.cmbStatus.Text = "Notar"
        End If
'    Case UCase("Fortbildungen")                                        ' Fortbildungen
'        Set NewFrmEdit = New frmEditFortbildungen
'        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey)
    Case UCase("Ausgeschriebene Stellen"), UCase("Stellen"), UCase("StellenJahr")   ' Stellen
        Set NewFrmEdit = New frmEditStellen
        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey, bDialog)
    Case UCase("Ausschreibungen"), UCase("Ausschreibung")
        Set NewFrmEdit = New frmEditAusschreibung
        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey, bDialog)
' MW 30.11.10 {
    Case UCase("Fristen"), UCase("Heute"), UCase("Abgelaufen"), UCase("Morgen") ' Fristen
         Set NewFrmEdit = New frmEditFrist
        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey)
' MW 30.11.10 }
    Case UCase("Bewerbung"), UCase("Bewerbungen")                       ' Bewerbungen
        Set NewFrmEdit = New frmEditBewerbung
        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey)
    Case UCase("Oberlandesgerichte")
         Set NewFrmEdit = New frmEditOLG
         Call NewFrmEdit.InitEditForm(parentform, objDBconn, szRootkey, DetailKey)
    Case UCase("Landgerichte")
         Set NewFrmEdit = New frmEditOLG
         Call NewFrmEdit.InitEditForm(parentform, objDBconn, szRootkey, DetailKey)
    Case UCase("Amtsgerichte")
         Set NewFrmEdit = New frmEditOLG
         Call NewFrmEdit.InitEditForm(parentform, objDBconn, szRootkey, DetailKey)
    Case UCase("Benutzerverwaltung"), UCase("Benutzer")                 ' Stammdaten Benutzer
        Set NewFrmEdit = New frmEditUser
        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey)
    Case UCase("Dokumente"), UCase("Letzte Woche"), UCase("Letzter Monat")  ' Dokumente
        If DetailKey = "" Then
            PersID = GetPersIDFormLV()
            StellenID = GetStellenIDFormLV
            AusschrID = GetAusschrIDFormLV
            Call WriteWord("", PersID, StellenID, AusschrID)            ' Starte Sat
            GoTo exithandler
        Else
            Call ShowWordDoc(DetailKey)
            GoTo exithandler
        End If
    Case UCase("Aktenort")                                              ' Aktenort
        Set NewFrmEdit = New FrmEditAktenOrt
        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey)
    Case UCase("Disziplinarmaßnahmen")                                  ' Disziplinarmaßnahmen
        Set NewFrmEdit = New frmEditDisz
        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey)
    Case UCase("Vorgang")                                               ' Vorgang (zur Zeit nicht benutzt)
        'Set NewFrmEdit = New frmVorgangSelect
        'Call NewFrmEdit.InitEditForm(objDBconn, DetailKey)
    Case UCase("Forderungen")                                           ' Forderungen
        Set NewFrmEdit = New frmEditForderungen
        Call NewFrmEdit.InitEditForm(parentform, objDBconn, DetailKey)
    Case Else                                                           ' Sonstiges Form
        Set NewFrmEdit = New frmEdit
        Call NewFrmEdit.InitEditForm(parentform, ThisDBCon, szRootkey, DetailKey)
'        Select Case UCase(szRootkey)
'        Case UCase("Amtsgerichte")
'            lngImageIndex = 1
'        Case UCase("Landgerichte")
'            lngImageIndex = 1
'        Case Else
'            lngImageIndex = 1
'        End Select
    End Select
    Call objError.WriteProt("OpenEditForm - RootKey: " & szRootkey & vbCrLf _
            & "DetailKey: " & DetailKey)                                ' Protoklieren
    Set EditFormArray(UBound(EditFormArray)) = NewFrmEdit               ' Form in EditFormArray
    If bDialog Then                                                     ' Form anzeigen
        NewFrmEdit.Show 1, Me
        OpenEditForm = NewFrmEdit.ID                                    ' ID aus form übernehmen (nur dialog)
        Call EditFormUnload(NewFrmEdit)                                 ' Form Schliessen
    Else
        NewFrmEdit.Show                                                 ' einfach anzeigen (kein  dialog)
    End If
    
exithandler:

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
    Dim i As Integer                                                    ' Counter
    Dim lngEditFormCount  As Integer                                    ' Anzahl der Edit Forms
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    lngEditFormCount = UBound(EditFormArray)                            ' Anz. Edit form ermitteln
    If Err.Number <> 0 Then                                             ' Errorhandling Deak. da Array evtl. leer
        lngEditFormCount = -1                                           ' Neg. Result für nix gefunden
        Err.Clear                                                       ' Fehler Resetten
    End If
On Error GoTo Errorhandler                                              ' Fehlerbehandlung wieder aktivieren
    If lngEditFormCount > -1 Then                                       ' wenn EditForms Array nicht leer
        For i = 0 To UBound(EditFormArray)                              ' Array duchlaufen
            If Not EditFormArray(i) Is Nothing Then
                If EditFormArray(i).ID = frmEdit.ID Then                ' Form anhand ID ermitteln
                    Set EditFormArray(i) = Nothing                      ' Nothing setzen
                    Exit For                                            ' Fertig
                End If
            End If
        Next                                                            ' Nachstes Form
    End If
    
exithandler:

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

Public Sub DeHoverAll()
    Dim i As Integer                                                    ' Counter
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    For i = 0 To lblNavBar.Count - 1                                    ' Alle Navbarlabled durchlaufen
        Call HoverLabel(lblNavBar(i), False)                            ' Dehovern
    Next i                                                              ' Nächstes Navbar Label
    Err.Clear                                                           ' Evtl Error clearen
End Sub

Public Function RefreshNavBar(TV As TreeView, bText As Boolean, Optional szFullPath As String)
    Dim szPathArray() As String                                         ' TV.Node.Fullpath in Array
    Dim i As Integer                                                    ' Counter
    Dim n As Integer                                                    ' nochn Counter
    Dim szTextPath As String                                            ' Fullpath als pur Text
On Error Resume Next                                                    ' Fehlerbehandlung erstmal Deaktiviert
    If szFullPath = "" Then                                             ' Wenn kein Fulpath angegeben
        szFullPath = TV.SelectedItem.FullPath                           ' Vom Akt Node holen
        Err.Clear                                                       ' Evtl err. Clearen
    End If
On Error GoTo Errorhandler                                              ' Fehlerbehandlung Aktivieren
    If szFullPath = "" Then GoTo exithandler                            ' Immer noch kein FullPath -> Fertig
    szPathArray = Split(szFullPath, TV_KEY_SEP)                         ' Fullpath in Array aufspalten
    lblNavBar(0).Caption = szPathArray(0)                               ' 1. Caption Setzen
    lblNavBar(0).Tag = szPathArray(0)                                   ' 1. Tag Setzen
    szTextPath = szPathArray(0)                                         ' TextPath Starten
    Call HoverLabel(lblNavBar(0), False)                                ' Hover effekt abschalten (dehovern)
    For i = 1 To UBound(szPathArray)                                    ' Array durchlaufen
        szTextPath = szTextPath & TV_KEY_SEP & szPathArray(i)           ' TextPath fortsetzen
        If lblNavBar.Count <= i Then
            Load lblNavBar(i)                                           ' Neues Label laden
            Load lblNavBarArrow(i)
        End If
        lblNavBar(i).Caption = szPathArray(i)                           ' Caption Setzen
        lblNavBar(i).Tag = szTextPath                                   ' Tag Setzen
        
        lblNavBarArrow(i).Left = lblNavBar(i - 1).Left + lblNavBar(i - 1).Width
        lblNavBar(i).Left = lblNavBarArrow(i).Left + lblNavBarArrow(i).Width
        lblNavBar(i).Visible = True                                     ' Sichtbar
        lblNavBarArrow(i).Visible = True
        Call HoverLabel(lblNavBar(i), False)                            ' dehovern
    Next i                                                              ' Nächstes Array Item
    For n = i To lblNavBar.Count - 1                                    ' Alle sonstigen Labels
        lblNavBar(n).Visible = False                                    ' Ausblenden
        lblNavBarArrow(n).Visible = False
        Call HoverLabel(lblNavBar(n), False)                            ' dehovern
    Next n                                                              ' Nächstes Navbar Lable
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "RefreshNavBar", errNr, errDesc)
    Resume exithandler
End Function

Public Function RefreshListView(Optional LV As ListView, Optional TV As TreeView, _
        Optional cNode As node, Optional ID As String)
' Aktualisiert nur das Listview
    Dim szLvItemKey As String                                           ' List view Item Key
    Dim LVInfo As ListViewInfo                                          ' Infos zum LV Handling aus XML
On Error Resume Next                                                    ' Errorhandling deak. da selectedItem evtl .leer
    szLvItemKey = LV.SelectedItem.Key                                   ' Key des Akt. select. LvItem holen
    Err.Clear                                                           ' Evtl Error Clearen
On Error GoTo Errorhandler                                              ' Fehlerbehandlung wieder aktivieren
    If LV Is Nothing Then Set LV = GetLV                                ' LV Prop. abfragen
    If TV Is Nothing Then Set TV = GetTV                                ' TV Prop. abfragen
    Call MousePointerHourglas(Me)                                       ' Sanduhr anzeigen
    If cNode Is Nothing Then                                            ' Kein Node angegeben
        Set cNode = GetSelectTreeNode(TV)                               ' Dann Akt Node holen
    End If
    If cNode Is Nothing Then GoTo exithandler                           ' Immer noch kein Konten dann fertig
    If ID = "" Then ID = GetIDFromNode(TV, cNode)                       ' Akt ID aus Node.key / Tag ermitteln
    With LVInfo
        Call objTools.GetLVInfoFromXML(App.Path & "\" & INI_XMLFILE, cNode.Tag, .szSQL, .szTag, .szWhere, _
                .lngImage, .bValueList, .bListSubNodes)                 ' LV infos aus mxl datei holen
        Call ListLVByTag(LV, ThisDBCon, cNode.Tag, ID, .bValueList, _
                cNode.Image)                                            ' Listitems anzeigen
        If .bListSubNodes Then Call ListLVFromSubNodes(LV, _
                TV, cNode)                                              ' Subnodes im LV anzeigen
        If Not .bValueList Then
            Call CountLVItems(LV)                                       ' Anzahl der listitem in statusbar anzeigen
        Else
            Call CountLVItems(LV, 1)                                    ' Wenn .bValueList anzahl ist immer 1
        End If
    End With
    If szLvItemKey <> "" Then Call SelectLVItem(LV, szLvItemKey)        ' Selectierten eintrag im LV wider auswählen
    
exithandler:
On Error Resume Next
    Call MousePointerDefault(Me)                                        ' Mousepointer wieder normal
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "RefreshListView", errNr, errDesc)
End Function

Public Function RefreshTreeView(Optional TV As TreeView, Optional nodekey As String, Optional bParent As Boolean)
' Aktualisiert TreeView & List View
    Dim cNode As node                                                   ' Akt. TV Node
    Dim cParentNode As node                                             ' Übergeordneter node (Parent)
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    
    If TV Is Nothing Then Set TV = GetTV                                ' TV Prop. abfragen
    If nodekey <> "" Then                                               ' Wenn Node Key angegeben
        Set cNode = GetNodeByKey(TV, nodekey)                           ' Node aus szKey ermitteln
    Else
        Set cNode = GetSelectTreeNode(TV)                               ' Akt Node holen
        nodekey = cNode.Key
    End If
    If cNode Is Nothing Then GoTo exithandler                           ' Kein node -> fertig
        
    If Not cNode Is Nothing Then                                        ' Wenn Node vorhanden
        If Not cNode.Parent Is Nothing Then                             ' Wenn Parent node existiert
            Set cParentNode = cNode.Parent                              ' Parent Node ermitten
            Call DelSubTreeNodes(TV, cParentNode)                       ' Alle unter knoten löschen
            Call HandleNodeClick(cParentNode, True)                     ' Unterknoten anlegen
        End If
        'Call HandleNodeClick(cNode, True)                              ' Unterknoten anlegen
    End If
    Set cNode = GetNodeByKey(TV, nodekey)                               ' Akt Node ermitteln
    If Not cNode Is Nothing Then                                        ' Wenn Node existiert
        Call SelectTreeNode(TV, cNode)                                  ' Node Selecten
        Call HandleNodeClick(cNode, True)                               ' Node Klick behandeln
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
    Const lngSpliterWidth = 70                                          ' Spliter breite
    Const lngStatusHeight = 300                                         ' Statusbar höhe
    Dim CTLTop  As Integer                                              ' Max CTL Top Pos.
    Dim CTLLeft As Integer                                              ' Max CTL Left Pos.
    Dim CTLWidth As Integer                                             ' Max CTL Breite
    Dim CTLHeight As Integer                                            ' Max CTL Höhe
On Error GoTo Errorhandler                                              ' Fehler behandlung aktivieren
    If Me.WindowState = vbMinimized Then GoTo exithandler               ' Mainform min. -> Fertig
    If Me.Width < 4500 Then Me.Width = 4500                             ' Min. Breite nicht unterschreiten
    If Me.Height < 4000 Then Me.Height = 4000                           ' Min Höhe nicht unterschreiten
    If Me.ScaleWidth = 0 Or Me.ScaleHeight = 0 Then GoTo exithandler
    If lngSplitpos = 0 Then lngSplitpos = 3000                          ' Min Spliter Pos
    CTLTop = Me.ScaleTop + Me.ToolbarMain.Height                        ' Max Top pos. ermitteln
    CTLLeft = Me.ScaleLeft                                              ' Max Left Pos. ermitteln
    CTLHeight = Me.ScaleHeight - lngStatusHeight - Me.ToolbarMain.Height ' Max höhe ermitteln
    CTLWidth = Me.ScaleWidth                                            ' Max breite ermitteln
    Call PicTree.Move(0, CTLTop, (CTLWidth * lngSplitpos), CTLHeight)   ' PicTree ausrichten (Splitter)
    Call TVMain.Move(0, 0, PicTree.Width, PicTree.Height)               ' Tree an PicTree ausrichten
    Call PicNavBar.Move(PicTree.Width + lngSpliterWidth, CTLTop, _
        CTLWidth - PicTree.Width - lngSpliterWidth, 255)                ' PicNavbar ausrichten
    Call PicList.Move(PicTree.Width + lngSpliterWidth, CTLTop + PicNavBar.Height, _
            CTLWidth - PicTree.Width - lngSpliterWidth, _
            CTLHeight - PicNavBar.Height)                               ' PicList Ausrichten
    Call LVMain.Move(0, 0, PicList.Width, PicList.Height)               ' ListView an PicList ausrichten
    curlngSplitposProz = lngSplitpos                                    ' Akt. Spliter pos merken
    Call objOptions.SetOptionByName(OPTION_SPLIT, curlngSplitposProz)   ' Spliter pos in optionen speichern
    LVMain.Refresh                                                      ' LV Neuzeichnen
    
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

Public Function HandleNodeClick(ByVal node As MSComctlLib.node, Optional bNotExpand As Boolean)
' Behandelt den Node Klick im TreeView
    Dim ID As String                                                    ' Evtl ID des Detaildatensatzes
    Dim bValueList As Boolean                                           ' Darstellung des ListView als Value list
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren

    Call MousePointerHourglas(Me)                                       ' Mousepointer auf Sanduhr
    DoEvents                                                            ' Evtl. ander aktionen zulassen
    If node.Children = 0 Then Call AddSubTreeNodes(TVMain, node, _
            ThisDBCon, node.Image, True)                                ' Evtl Subnodes hinzufügen
    Call AddNodeToNavList(node.Key)                                     ' Knoten in Navlist aufnehmen
    Call RefreshNavBar(TVMain, True)                                    ' NavBar aktualisieren
    If Not bNotExpand And Not node.Expanded Then node.Expanded = True ' Evtl. Knoten auffalten
    Call SaveColumnWidth(LVMain)                                        ' Spalten breite speichern
    ID = GetIDFromNode(TVMain, node)                                    ' Akt ID aus Node.key / Tag ermitteln
    Call RefreshListView(LVMain, TVMain, node, ID)                      ' ListView füllen
    Call DiscribeTreeNode(TVMain)                                       ' Welche Liste in Status bar
    
exithandler:
On Error Resume Next
    Call MousePointerDefault(Me)                                        ' Mauszeiger wieder normal
    DoEvents                                                            ' Andere Events zulassen
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "HandleNodeClick", errNr, errDesc)
End Function

'Private Function GetIDFromNode(cNode As node) As String
'' Ermitteli DS ID aus cNode.Tag & Key
'    Dim szKeyArray() As String                                      ' Node Key in array aufgespalten
'    Dim szTagArray() As String                                      ' Node Tag in array aufgespalten
'    Dim szTmp As String                                             ' Hilfsvariable
'    Dim ID As String                                                ' Evtl ID des Detaildatensatzes
'    Dim i As Integer                                                ' Counter
'
'On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
'
'    szTagArray = Split(cNode.Tag, "\")                              ' Node Tag aufspalten
'    szKeyArray = Split(cNode.Key, "\")                              ' Node Key aufspalten
'
'    If szTagArray(UBound(szTagArray)) = "*" Then                    ' detaildatensatz
'        ID = GetLastKey(cNode.Key, TV_KEY_SEP)                      ' ID aus Key ermitten
'    Else                                                            ' Wenn kein Detail Datensatz
'        If InStr(cNode.Tag, "*") Then                               ' Statischer unterknoten eines Detaildatensatzes
'            szTmp = ""                                              ' Tmp Leeren
'            i = UBound(szTagArray) + 1                              ' Max Arra Index festlegen
'            While szTmp <> "*"                                      ' Tag bis * rückwärts durchlaufen
'                i = i - 1                                           ' arrayindex herunterzählen
'                szTmp = szTagArray(i)                               ' array wert merken
'                ID = szKeyArray(i)                                  ' ID am gleicher stelle aus Tag
'            Wend
'        End If
'        If ID = "" Then ID = GetLastKey(cNode.Key, TV_KEY_SEP)      ' sonst mit gewalt
'    End If
'
'exithandler:
'    GetIDFromNode = ID                                              ' ID Zurück geben
'Exit Function
'Errorhandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.Errorhandler(MODULNAME, "GetIDFromNode", errNr, errDesc)
'End Function

Public Function DiscribeTreeNode(TV As TreeView, Optional szDesc As String, Optional bNoShowInStatusBar As Boolean) As String
' ermittelt die Beschreibung des Aktuellen Tree nodes und setz diese gegebenebfalls in die Statusbar
    Dim Nodeinfo As TreeViewNodeInfo                                    ' Nodeinfos aus XML
    Dim cNode As node                                                   ' Aktuell ausgewählter Node
On Error GoTo Errorhandler                                              ' Fehlerbehandlung akt.
    Set cNode = GetSelectTreeNode(TV)                                   ' aktuellen Tree node ermitteln
    If Not cNode Is Nothing Then
        With Nodeinfo
            Call objTools.GetTVNodeInfofromXML(App.Path & "\" & INI_XMLFILE, _
                    cNode.Tag, .szTag, .szText, .szKey, .bShowSubnodes, .szSQL, .szWhere, _
                    .lngImage, .bShowKontextMenue, .szDesc)             ' Tree node informationen aus XML laden
            szDesc = .szDesc
        End With
    Else
        szDesc = ""                                                     ' keine Beschreibung
    End If
    If Not bNoShowInStatusBar Then StatusBarMain.Panels(5).Text = szDesc ' In Statusbar anzeigen

exithandler:
    DiscribeTreeNode = szDesc
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "DiscribeTreeNode", errNr, errDesc)
    Resume exithandler
End Function

Private Function GetPersIDFormLV() As String
' Ermittelt eine ID einer Person aus den Daten des Aktuellen Listview im frmMain
    Dim szRootkey As String                                             ' RootKey der ListViewDaten
    Dim DSID As String                                                  ' Datensatz ID zum Rootkey
    Dim szPersID As String                                              ' Personen ID
    Dim szSQL As String                                                 ' SQL Statement
On Error GoTo Errorhandler                                              ' Feherbehandlung aktivieren
    Call GetKontextRoot(szRootkey, DSID, "")                            ' Rootkey und Akt DS ID ermitteln
    Select Case UCase(szRootkey)
    Case UCase("Personen"), UCase("Bewerber"), UCase("Teilnehmer"), UCase("Notare"), _
            UCase("Notare bestellt"), UCase("Notare ausgeschieden")     ' RootKey ist ein Personen DS
        szPersID = DSID                                                 ' Personen ID setzen
    Case UCase("Bewerbungen")                                           ' Rootkey ist Bewerbung
        If DSID <> "" Then
            szSQL = "SELECT FK010013 FROM BEWERB013 WHERE ID013 = '" & DSID & "'"
            szPersID = ThisDBCon.GetValueFromSQL(szSQL)                 ' Person aus Bewerbung ermitteln
        End If
    
    Case Else
    
    End Select
    
    GetPersIDFormLV = szPersID                                          ' ergebniss zurück
    
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

Private Function GetStellenIDFormLV() As String
' Ermittelt eine ID einer Stelle aus den Daten des Aktuellen Listview im frmMain
    Dim szRootkey As String                                             ' RootKey der ListViewDaten
    Dim DSID As String                                                  ' Datensatz ID zum Rootkey
    Dim szStellenID As String                                           ' Stellen ID
    Dim szSQL As String                                                 ' SQL Statement
On Error GoTo Errorhandler
    Call GetKontextRoot(szRootkey, DSID, "")                            ' Rootkey und Akt DS ID ermitteln
    Select Case UCase(szRootkey)
    Case UCase("Ausgeschriebene Stellen"), UCase("StellenJahr")         ' Rootkey ist eine Stelle
        szStellenID = DSID                                              ' Stellen ID rurück liefern
    Case UCase("Personen"), UCase("Bewerber"), UCase("Teilnehmer"), UCase("Notare"), _
            UCase("Notare bestellt"), UCase("Notare ausgeschieden")     ' Rootkey ist eine Person
        If DSID <> "" Then
            szSQL = "SELECT ID012 FROM Stellen012 INNER JOIN BEWERB013 ON FK012013 = ID012 " & _
                " INNER JOIN RA010 ON ID010 = FK010013 WHERE ID010 ='" & DSID & "'"
            szStellenID = ThisDBCon.GetValueFromSQL(szSQL)              ' Stellen ID aus Person ermitteln
        End If
    Case UCase("Bewerbungen")                                           ' Rootkey ist eine Bewerbung
        If DSID <> "" Then
            szSQL = "SELECT FK012013 FROM BEWERB013 WHERE ID013 = '" & DSID & "'"
            szStellenID = ThisDBCon.GetValueFromSQL(szSQL)              ' Stellen ID aus Bewerbung ermmitteln
        End If
    Case Else
    
    End Select
    
    GetStellenIDFormLV = szStellenID                                    ' Ergebnis Zurück
exithandler:
On Error Resume Next
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetStellenIDFormLV", errNr, errDesc)
End Function

Private Function GetAusschrIDFormLV() As String
' Ermittelt eine ID einer Ausschreibung aus den Daten des Aktuellen Listview im frmMain
    Dim szRootkey As String                                             ' RootKey der ListViewDaten
    Dim DSID As String                                                  ' Datensatz ID zum Rootkey
    Dim szAusschrID As String                                           ' Ausschreibungs ID
    Dim szSQL As String                                                 ' SQL Statement
On Error GoTo Errorhandler
    Call GetKontextRoot(szRootkey, DSID, "")                            ' Rootkey und Akt DS ID ermitteln
    Select Case UCase(szRootkey)
    Case UCase("Ausschreibung")                                         ' Rootkey ist Ausschreibung
        szAusschrID = DSID                                              ' DSID ist Ausschreibungs ID
    Case UCase("Ausgeschriebene Stellen"), UCase("StellenJahr")         ' RootKey ist Stelle
        If DSID <> "" Then
            szSQL = "SELECT FK020012 FROM STELLEN012 WHERE ID012='" & DSID & "'"
            szAusschrID = ThisDBCon.GetValueFromSQL(szSQL)              ' Ausschreibungs ID aus Stelle ermitteln
        End If
    Case UCase("Personen"), UCase("Bewerber"), UCase("Teilnehmer"), UCase("Notare"), _
            UCase("Notare bestellt"), UCase("Notare ausgeschieden")     ' RootKey ist Person
        If DSID <> "" Then
            szSQL = "SELECT FK020012 FROM STELLEN012 INNER JOIN BEWERB013 ON FK012013 = ID012 " & _
                " INNER JOIN RA010 ON ID010 = FK010013 WHERE ID010 ='" & DSID & "'"
            szAusschrID = ThisDBCon.GetValueFromSQL(szSQL)              ' Ausschreibungs ID aus Person ermitteln
        End If
    Case UCase("Bewerbungen")                                           ' RootKey ist Bewerbung
        If DSID <> "" Then
            szSQL = "SELECT FK020012 FROM STELLEN012 INNER JOIN BEWERB013 ON FK012013 = ID012 " & _
                " WHERE ID013 = '" & DSID & "'"
            szAusschrID = ThisDBCon.GetValueFromSQL(szSQL)              ' Ausschreibungs ID aus Bewerbung ermitteln
        End If
    Case Else
    
    End Select
    
    GetAusschrIDFormLV = szAusschrID                                    ' Ergebnis Zurück
exithandler:
On Error Resume Next
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetAusschrIDFormLV", errNr, errDesc)
End Function

Private Function DoNav(bBack As Boolean)
    ' Navigiert durch die Liste der gespeicherten Tree nodes
    Dim NavIndexMax As Integer                                          ' Index Obergrenze des nav Arrays
    Dim cNode As node                                                   ' Akt TV Node
    Dim szKey As String                                                 ' Node key
On Error Resume Next                                                    ' Errorhandling deak. da Array evtl. leer
    NavIndexMax = UBound(NavTreeNodeArray)                              ' Array überprüfen
    If Err.Number <> 0 Then
        NavIndexMax = -1
        Err.Clear
    End If
On Error GoTo Errorhandler                                              ' Fehlerbehandlung wieder aktivieren
    If NavIndexMax = -1 Then GoTo exithandler                           ' Kein vor oder zurück möglich
    If bBack Then                                                       ' Rückwärts
        If lngNavIndex < 0 Then GoTo exithandler                        ' Kein zurück möglich
        If lngNavIndex > 0 Then lngNavIndex = lngNavIndex - 1
    Else                                                                ' Vorwärts
        If lngNavIndex = NavIndexMax Then GoTo exithandler              ' Kein vor möglich
        lngNavIndex = lngNavIndex + 1
    End If
    szKey = NavTreeNodeArray(lngNavIndex)                               ' Key aus array ermitteln
    If szKey = "" Then GoTo exithandler                                 ' Kein Key fertig
    Set cNode = GetNodeByKey(TVMain, szKey)                             ' Node mit Key Ermitteln
    If Not cNode Is Nothing Then                                        ' Wenn Knoten gefunden
        Call HandleNodeClick(cNode)                                     ' Click behandeln
        Call SelectTreeNode(TVMain, cNode)
    End If
    ToolbarMain.Buttons(2).Enabled = Not (lngNavIndex >= NavIndexMax)   ' evtl. Button Vor disablen
    'If lngNavIndex = NavIndexMax Then ToolbarDB.Buttons(2).Enabled = False
    ToolbarMain.Buttons(1).Enabled = Not (lngNavIndex = 0)              ' evtl Button Zurück disablen
    'If lngNavIndex = 0 Then ToolbarDB.Buttons(1).Enabled = False
    ToolbarMain.Refresh                                                 ' Toolbar neuzeichnen
    
exithandler:

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
    If szKey = "" Then GoTo exithandler                                 ' kein Key -> Fertig
On Error Resume Next                                                    ' Errorhandling deak. da Array evtl. leer
    ReDim Preserve NavTreeNodeArray(UBound(NavTreeNodeArray) + 1)       ' Array Prüfen
    If Err.Number <> 0 Then                                             ' Array Leer
        ReDim NavTreeNodeArray(0)
        NavTreeNodeArray(UBound(NavTreeNodeArray)) = szKey              ' Key anfügen
        Err.Clear
        GoTo exithandler                                                ' Fertig
    End If
On Error GoTo Errorhandler                                              ' Fehlerbehandlung wieder aktivieren
    If UBound(NavTreeNodeArray) >= 1 Then                               ' array nicht leer
        If NavTreeNodeArray(UBound(NavTreeNodeArray) - 1) = szKey Then GoTo exithandler
    End If
    NavTreeNodeArray(UBound(NavTreeNodeArray)) = szKey                  ' Key anfügen
    lngNavIndex = UBound(NavTreeNodeArray)
    ToolbarMain.Buttons(1).Enabled = True                               ' Button Vor enablen
    ToolbarMain.Refresh                                                 ' Toolbar neuzeichnen
    
exithandler:

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
    Dim szRootkey As String                                             ' TVNode /LV Key als Kontext
    Dim szDetailKey As String                                           ' evtl. Datensatz ID
    Dim szAction As String                                              ' Für diesen Node vorgesehene Aktion
    Dim cNode As node                                                   ' Akt. TreeNode
    Dim szDetails As String                                             ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    Call GetKontextRoot(szRootkey, szDetailKey, szAction)               ' Kontext ermitteln
    szDetails = "RootKey: " & szRootkey & vbCrLf & "ID: " & szDetailKey & vbCrLf & "Aktion: " & szAction
    If InStr(UCase(szAction), "EDIT") > 0 Then
        Call frmMain.OpenEditForm(szRootkey, szDetailKey, Me)           ' DS zum bearbeiten öffnen
    End If
'    Case "SelectNode"
    If InStr(UCase(szAction), UCase("SelectNode")) > 0 Then
        Set cNode = GetNodeByKey(TVMain, LVMain.SelectedItem.Key)       ' TVNode ermitteln
        If Not cNode Is Nothing Then                                    ' Wenn Node Exitsiert
            Call SelectTreeNode(TVMain, cNode, True)                    ' Node auswählen
            Call HandleNodeClick(cNode)                                 ' Node Klick behandeln
        End If
    End If
exithandler:
Exit Sub                                                                ' Function Beenden
Errorhandler:
    Dim errNr As String                                                 ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "HandleLVItemDblKlick", errNr, errDesc, szDetails) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Sub

Private Sub HandleLVKontextNew()
' Öffent neuen DS aus Kontextmenü
    Dim szKeyArray() As String                                          ' Array mit Key elementen
    Dim szRootkey As String                                             ' Gibt an was für ein DS neu angelegt werden soll
    Dim szDetails As String                                             ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    Call GetKontextRoot(szRootkey, "", "")                              ' Rootkey ermitteln
    Call NewDS(szRootkey, Me)                                           ' DS form für neuen DS öffnen
    'Call OpenEditForm(szRootkey, "", Me)
exithandler:
Exit Sub                                                                ' Function Beenden
Errorhandler:
    Dim errNr As String                                                 ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "HandleLVKontextNew", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Sub

Private Sub HandleLVKontextDel()
' Löscht DS aus Kontextmenü
    Dim szRootkey  As String                                            ' Gibt an was für ein DS neu angelegt werden soll
    Dim szID As String                                                  ' ID des Datensatzes
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    Call GetKontextRoot(szRootkey, szID, "")                            ' ID und Rootkey ermitteln
    If szID = "" Then GoTo exithandler                                  ' Keine ID -> Fertig
    Call DeleteDS(szRootkey, szID)                                      ' DS Löschen
    'Call RefreshListView(LVMain, TVMain)                               ' DS Aktualisieren
exithandler:
Exit Sub                                                                ' Function Beenden
Errorhandler:
    Dim errNr As String                                                 ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "HandleLVKontextDel", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Sub

Private Sub HandleLVKontextEdit()
' Öffent DS aus Kontextmenü
    Dim szDetails As String                                             ' Details für Fehlerbehandlung
    Dim szRootkey As String                                             ' Gibt an was für ein DS neu angelegt werden soll
    Dim szDetailKey As String                                           ' DS Id
    Dim szAction As String                                              ' Aktion (evtl. edit nicht zulassig
On Error Resume Next                                                    ' Errorhandling deakt. da SelectedItem evtl. nicht ex.
    szDetails = "LVTag: " & LVMain.Tag & vbCrLf & "SelectedItem.tag: " & LVMain.SelectedItem.Key
    Err.Clear
On Error GoTo Errorhandler                                              ' Fehlerbehandlung wieder aktivieren
    Call GetKontextRoot(szRootkey, szDetailKey, szAction)               ' ID, Rootkey und Aktion ermitteln
    If szDetailKey = "" Then GoTo exithandler                           ' Keine ID -> Fertig
    Call EditDS(szRootkey, szDetailKey, False, Me)                      ' Form zum bearbeiten öffnen
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
    Dim szPersID As String                                              ' Empfänger ID
    Dim szStellenID As String                                           ' Welche Stellen ausschreibung
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    szPersID = GetPersIDFormLV()                                        ' Empfäger ID ermitteln
    szStellenID = GetStellenIDFormLV()
    If szPersID = "" Then GoTo exithandler                              ' Ohne Empfänger fertig
    Call WriteWord("", szPersID, szStellenID)                           ' Dok erstellen
exithandler:
Exit Sub                                                                ' Function Beenden
Errorhandler:
    Dim errNr As String                                                 ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "HandleLVKontextNewDoc", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Sub

'Public Function GetKontextRoot(bTV As Boolean,szRootkey As String, szDetailKey As String, Optional szAction As String)
Public Function GetKontextRoot(szRootkey As String, szDetailKey As String, Optional szAction As String)
' ermittelt Rootkey und DS ID sowie mögliche Aktionen aus den Kontext (ListView) und XML
    Dim szItemTagArray() As String                                      ' Array mit ListView Item Tag elementen
    Dim szItemKeyArray() As String                                      ' Array mit ListView Item Key elementen
    Dim szLVTagArray() As String                                        ' Array mit ListView Tag elementen
    Dim szItemTag As String                                             ' Tag des ListView Items (* statt ID)
    Dim szItemKey As String                                             ' Key des ListView Items (enthält ID)
    Dim szLVTag As String                                               ' Tag des Listviews
    Dim szDetails As String                                             ' Details für Fehlerbehandlung
    Dim TVNode As TreeViewNodeInfo                                      ' Infos über TreeNode
    Dim LVInfo As ListViewInfo                                          ' Infos über ListView
On Error Resume Next                                                    ' Errorhandling deakt. da SelectedItem evtl. Nothing
    szItemTag = LVMain.SelectedItem.Tag                                 ' Tag des akt. ListView Items ermitteln
    szItemKey = LVMain.SelectedItem.Key                                 ' Key des akt. ListView Items ermitteln
    szLVTag = LVMain.Tag                                                ' Tag des ListViews ermitteln
    'szDetails = "LVTag: " & lvmain.Tag & vbCrLf & "SelectedItem.tag: " & lvmain.SelectedItem.Key
    Err.Clear
On Error GoTo Errorhandler                                              ' Fehlerbehandlung wieder aktivieren
    With LVInfo                                                         ' ListViewInfo aus XML füllen
        Call objTools.GetLVInfoFromXML(App.Path & "\" & INI_XMLFILE, LVMain.Tag, _
                .szSQL, .szTag, .szWhere, .lngImage, .bValueList, .bListSubNodes, .bEdit, .bNew, .bSelectNode)
        If .bEdit Then                                                  ' Edit zulässig
            szAction = "Edit"
        End If
        If .bSelectNode Then szAction = szAction & "SelectNode"
        'If (Not .bEdit) And .bSelectNode Then szAction = "SelectNode"  ' Select zulässig
        If Not .bEdit And Not .bSelectNode Then szAction = "NoAction"
    End With
    If szItemKey = "" Or szItemTag = "" Then
        szRootkey = LVInfo.szTag
        If szItemKey = "" And szItemTag = "" Then GoTo exithandler
    End If
    szItemTagArray = Split(szItemTag, TV_KEY_SEP)                       ' ListView Item Tag aufspalten
    szItemKeyArray = Split(szItemKey, TV_KEY_SEP)                       ' ListView Item Key aufspalten
    szLVTagArray = Split(szLVTag, TV_KEY_SEP)                           ' ListView Tag aufspalten
    If UBound(szItemKeyArray) = UBound(szItemTagArray) Then
        If szItemTagArray(UBound(szItemTagArray)) <> "*" Then           ' Case 3 SubNode in Valuelist
            ' Ubound(szItemKeyArray ) = Ubound(szItemTagArray) / key mit ID / tag mit *
            szAction = "SelectNode"
            szRootkey = szItemTagArray(UBound(szItemTagArray))
        Else                                                            ' Case 1 Detail SubNode in ListView
        ' Ubound(szItemKeyArray ) = Ubound(szItemTagArray) / key mit ID / tag mit *
        ' -> Select Node and/or Edit
            szDetailKey = szItemKeyArray(UBound(szItemKeyArray))        ' ID aus ListItem.Key ermitten
            szRootkey = szItemKeyArray(UBound(szItemKeyArray) - 1)      ' szRootkey aus Item Key ermitteln
        '    szAction = "Edit"
            If szRootkey = "*" Then                                     ' Case 2a einzelner DetailsDS (nicht in Valuelist)
                szRootkey = szItemKeyArray(UBound(szItemKeyArray) - 2)  ' Plan b ( dürft nicht vorkommen)
            End If
        End If
    ' case 4 Relation Deatilnode eines Detailnodes in Liste
    ' Ubound(szItemKeyArray ) = Ubound(szItemTagArray) / key mit ID / tag mit *
    ' szLVTagArray(UBound(szLVTagArray)) <>"*"
    ' Select Node (Edit?)
    
    Else                                                                ' Case 2 einzelner DetailsDS (Valuelist)
    ' Ubound(szLVTagArray) = Ubound(szItemTagArray) / LVTag mit ID / Itemtag mit *
        szDetailKey = szItemTagArray(UBound(szItemTagArray))            ' ID aus ListItem.Tag ermitten
        szRootkey = szItemTagArray(UBound(szItemTagArray) - 1)          ' szRootkey aus Item Tag ermitteln
        'szAction = "Edit"
    End If
    If szRootkey = "*" Or szRootkey = "" Then szAction = "SelectNode"
exithandler:
Exit Function                                                           ' Function Beenden
Errorhandler:
    Dim errNr As String                                                 ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "GetKontextRoot", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Function

Private Function CountLVItems(LV As ListView, Optional Count As Integer, Optional bNoShowInStatusBar As Boolean) As Integer
    Dim lngCount As Integer                                             ' Anzahl der angezeigten LV Items
On Error GoTo Errorhandler
    lngCount = LV.ListItems.Count                                       ' Durchzählen
    If Count > 0 Then lngCount = Count
    If Not bNoShowInStatusBar Then StatusBarMain.Panels(6).Text = lngCount & " DS"  ' In Statusbar anzeigen
    
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

Public Function ChangePWD(ID As String)
    Dim szUsername As String                                            ' Benutzername
    'Dim fChangePwd As frmChangePWD                                      ' Form zur pwd eigabe
    Dim szNewPWD As String                                              ' Neues Password
    Dim szSQL As String                                                 ' SQL Statement
    Dim dChangeDate As Date
    Dim szPersID As String                                              ' a_personen.pers_ID
    Dim bCancel As Boolean                                              ' Abbruch bed.
    Dim bChangeAtNextLogin As Boolean
On Error GoTo Errorhandler
    If ID = "" Then GoTo exithandler                                    ' Kein ID dann Raus
    If InStr(ID, Chr(1)) > 0 Then
        'szPersID = GetIDFromClusterID(ID, "pers_id")
    Else
        szPersID = ID
    End If
    If szPersID = "" Then GoTo exithandler                              ' Keine pers_ ID dann Raus
    szSQL = "SELECT USERNAME001 from USER001 " & _
            " WHERE ID001 = '" & szPersID & "'"                         ' Benutzernamen zur id holen
    szUsername = ThisDBCon.GetValueFromSQL(szSQL)                       ' SQL Statement ausführen
    If szUsername = "" Then GoTo exithandler                            ' Kein Username dann Raus
    szNewPWD = ShowChangePWDForm(szUsername, bCancel, False, bChangeAtNextLogin)
    If bCancel Then GoTo exithandler                                    ' Wenn abbruch (durch User) dann raus
    If bChangeAtNextLogin Then
'        dChangeDate = Now()
    Else
'        dChangeDate = DateAdd("d", -200, Now())
    End If
    
    szSQL = "UPDATE USER001 SET PWD001 = '" & szNewPWD & "' WHERE ID001 = '" & szPersID & "'"
    Call ThisDBCon.execSql(szSQL)                                   ' SQL Statement ausführen
    
exithandler:
On Error Resume Next
    
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ChangePWD", errNr, errDesc)
    Resume exithandler
End Function

Private Sub HandleNavbarClick(szFullPath As String)
    Dim cNode As node                                                   ' zu Selectender Tree Node
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    Set cNode = GetNodeByFullPath(TVMain, szFullPath)                   ' Node aus fullpath ermitteln
    If Not cNode Is Nothing Then                                        ' Node gefunden
        Call SelectTreeNode(TVMain, cNode)                              ' Node  selecten
        Call HandleNodeClick(cNode)                                     ' NodeKlick behandeln
    End If
exithandler:
    
Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "HandleNavbarClick", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub HandleMenueKlick(szMenueName As String)
' Behandelt Menü Kontextmenü und Toolbar Klicks
    Dim ID As String                                                    ' DS ID
    Dim Key As String                                                   ' RootKey gibt an was für ein DS behandelt wird
    Dim PersID As String                                                ' ID eines Personen DS
    Dim StellenID As String                                             ' ID Einer Stelle
    Dim AusschrID As String                                             ' ID einer Ausschreibung
    Dim tmpKey As String
    Dim szAction As String
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    Select Case szMenueName
                                                                        ' *****************************************
                                                                        ' Main Menü
    Case "mnuDateiExit"                                                 ' Anwendung Beenden
        Call AskForExit                                                 ' Nachfragen ob Beenden
'        Call Unload(Me)                                                 ' Main Form schliessen
'        Call AppExit                                                    ' Application beenden
    Case "mnuDateiPrint"                                                ' Akt Listview Ansicht über Excel Drucken
        Call SetLVDataInWorkSheet(LVMain, False, True, True)            ' Daten Ausdrucken (über Excel)
    Case "mnuDateiNewBewerber"
        Call NewDS("Bewerber", Me, False)                               ' Neuen Bewerber anlegen
    Case "mnuDateiNewBewerbung"
        Call NewDS("Bewerbung", Me, False)                              ' Neue Bewerbung anlegen
    Case "mnuDateiNewStelle"
        Call NewDS("Stellen", Me, False)                                ' Neue Stelle anlegen
    Case "mnuDateiNewAusschreibung"
        Call NewDS("Ausschreibung", Me, False)                          ' Neue Ausschreibung anlegen
    Case "mnuDateiNewDoc"
        Call WriteWord                                                  ' Neues Dokument ohne vorgabe der Vorlage oder empfänger
    Case "mnuDateiSuchen"                                               ' Suchen
        ID = ShowSearch(objDBconn, Key, "")                             ' Such Dialog aufrufen
        If ID <> "" Then Call OpenEditForm(Key, ID, Me)                 ' Erg anzeigen
'    Case "mnuEditAusschreibungNew"
'        Call OpenEditForm("Ausschreibung", "", Me)
    Case "mnuEditSelectTemplate"                                        ' Anschreiben
        'Call GetKontextRoot("", PersID, "")
        Call WriteWord(, PersID)                                        ' Starte Sat
    Case "mnuEditDocImport"
        Call ImportWordDoc(ThisDBCon, PersID, StellenID)                ' Starte Doc import (einz. Doc)
    Case "mnuEditVerzImport"
        Call ImportWordDocFolder(ThisDBCon, PersID, StellenID)          ' Starte verz import
    Case "mnuEditExcelExport"
        Call SetLVDataInWorkSheet(LVMain, False, False, True)           ' Akt ListView daten nach excel exportieren
'    Case "mnuEditWorkflow"
'        Call OpenEditForm("Vorgang", "", Me)
    Case "mnuExtrasOptions"                                             ' Optionen
        Call ShowOptions
    Case "mnuExtrasChangePWD"                                           ' Kennwort änderung
        Call ThisDBCon.UserChangePWD(User.NTUsername)
    Case "mnuInfoAbout"                                                 ' Info Dialog
        Call ShowAbout("", True)
    Case "mnuInfoHelp"                                                  ' Online Hilfe
        Call ShowHelp
    Case "mnuInfoReadMe"                                                ' ReadMe anzeigen
        Call ShowReadMe
                                                                        ' *****************************************
                                                                        ' Kontext menü
    Case "kmnuListNewUser", "kmnuListNewPerson", "kmnuListNew", "kmnuListNewAusschreibung", _
        "kmnuListNewNotar", "kmnuListNewBewerber", "kmnuListNewBewerbung"
        Call HandleLVKontextNew                                         ' Neuer DS
        
    Case "kmnuListEditUser", "kmnuListEditPerson", "kmnuListEdit", "kmnuListEditAusschreibung", _
        "kmnuListEditNotar", "kmnuListEditBewerber", "kmnuListEditBewerbung"
        Call HandleLVKontextEdit                                        ' Edit DS
    Case "kmnuListNewDocPerson", "kmnuListNewDocNotar", "kmnuListNewDocBewerber"
        Call HandleLVKontextNewDoc                                      ' Neues Doc für
    Case "kmnuListDelPerson", "kmnuListDelNotar", "kmnuListDelBewerber", _
            "kmnuListDelUser", "kmnuListDel", "kmnuListDelAusschreibung", "kmnuListDelBewerbung"
        Call HandleLVKontextDel                                         ' DS Löschen
    Case "kmnuListPrint", "kmnuListPrintBewerber", "kmnuListPrintNotar", _
            "kmnuListPrintPerson", "kmnuListPrintUser", "kmnuListPrintAusschreibung"
        Call SetLVDataInWorkSheet(LVMain, False, True, True)            ' Daten Ausdrucken (über Excel)
    Case "kmnuListChangePWD"                                            ' Kennwort ändern
        Call GetKontextRoot(Key, ID, "")
        If ID <> "" And Key = "Benutzerverwaltung" Then Call ChangePWD(ID)
    Case "kmnuListSearchPerson"                                         ' Person Suchen
        ID = ShowSearch(objDBconn, "Personen", "Nachname")
        If ID <> "" Then Call OpenEditForm("Personen", ID, Me)
    Case "kmnuListSearchNotar"
        ID = ShowSearch(objDBconn, "Notar", "Nachname")                 ' Notar Suchen
        If ID <> "" Then Call OpenEditForm("Personen", ID, Me)
    Case "kmnuListSearchBewerber"                                       ' Bewerber Suchen
        ID = ShowSearch(objDBconn, "Bewerber", "Nachname")
        If ID <> "" Then Call EditDS("Personen", ID, False)
    Case "kmnuListEditBewerberPerson"                                   ' Personen Daten der Bewerbung anzeigen
        ID = GetPersIDFormLV()                                          ' Personen ID ermitteln
        If ID = "" Then GoTo exithandler                                ' Keine ID -> Fertig
        Call EditDS("Personen", ID, False, Me)                          ' Form zum bearbeiten öffnen
                                                                        ' *****************************************
                                                                        ' Toolbar ( szMenueName ist hier ButtonKey)
    Case TB_LEFT
        Call DoNav(False)                                               ' Vorwärts navigieren
    Case TB_RIGHT
        Call DoNav(True)                                                ' Rückwärts navigieren
    Case TB_REFRESH                                                     ' Refresh
        tmpKey = LVMain.SelectedItem.Key
        Call RefreshTreeView(TVMain, TVMain.SelectedItem.Key)
        Call SelectLVItem(LVMain, tmpKey)
    Case TB_NEW                                                         ' Neuer Ds
        Call GetKontextRoot(Key, "", szAction)                          ' Kontext ermitteln
        If Key <> "" And szAction <> "" Then
            Call NewDS(Key, Me, False)                                  ' Neuer Datensatz
        End If
    Case TB_PRINT
        Call SetLVDataInWorkSheet(LVMain, False, True, True)            ' Daten Ausdrucken (über Excel)
    Case TB_NEWBEWERBUNG
        Call NewDS("Bewerbung", Me, False)                              ' Neue Bewerbung anlegen
    Case TB_NEWBEWERBER
        Call NewDS("Bewerber", Me, False)                               ' Neuen Bewerber anlegen
    Case TB_NEWSTELLE
        Call NewDS("Stellen", Me, False)                                ' Neue Stelle anlegen
    Case TB_NEWAUSSCR
        Call NewDS("Ausschreinung", Me, False)                          ' Neue Ausschreibung
    Case TB_NEWDOC                                                      ' Starte SAT
        Call WriteWord                                                  ' Neues Dokument ohne vorgabe des Vorlage oder empfänger
    Case TB_DOCNEW                                                      ' Starte SAT
        PersID = GetPersIDFormLV                                        ' Empfänger aus kontext ermitteln
        StellenID = GetStellenIDFormLV
        AusschrID = GetAusschrIDFormLV
        Call WriteWord(, PersID, StellenID)                             ' Neues Dokument mit empfänger
    Case TB_SEARCH                                                      ' Suchen
        Call GetKontextRoot(Key, "")                                    ' Suchkontext ermitteln
        ID = ShowSearch(objDBconn, Key, "")
        If ID <> "" Then Call EditDS(Key, ID, False)
    Case TB_SEARCH_PERS                                                 ' Suche Person
        ID = ShowSearch(objDBconn, "Personen", "Nachname")
        If ID <> "" Then Call EditDS("Personen", ID, False)
    Case TB_SEARCH_DOC                                                  ' Suche Dokument
        ID = ShowSearch(objDBconn, "Dokumente", "Empfänger")
        If ID <> "" Then Call EditDS("Dokumente", ID, False)
    Case TB_HELP                                                        ' Zeige Hilfe
        Call ShowHelp
    Case TB_INFO                                                        ' Zeige Info
        Call ShowAbout
    Case Else
    
    End Select
    
exithandler:

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
            If szKey <> "" And szID <> "" Then                      ' Ergebnis liegt vor
                Call EditDS(szKey, szID, , Me)                      ' DS zum bearbeiten anzeigen
                'Call OpenEditForm(szKey, szID, Me)                  ' DS anzeigen
            End If
        End If
        If KeyCode = 78 Then                                        ' Strg + N (Neu)
            Call GetKontextRoot(szKey, szID)                        ' Kontext ermitteln
            If szKey <> "" And szID <> "" Then
                Call NewDS(szKey, Me)                               ' Neuen DS anzeigen
                'Call OpenEditForm(szKey, "", Me)                    ' Neuen DS anzeigen
            End If
        End If
    End If

exithandler:
    
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
    Call HandleMenueKlick(ButtonKey)                                    ' Weiterreichen an Handle menueklick
End Sub

Private Sub ShowKontextMenu(Optional bTV As Boolean)
' Zeigt Kontext menue an
' bTV True -> TV sonst LV
    Dim szKeyArray() As String                                          ' Array mit Key/Tag elementen
    Dim TVNode As TreeViewNodeInfo                                      ' TreeNode Infos
    Dim szDetails As String                                             ' Details für Fehlerbehandlung
    Dim LVInfo As ListViewInfo                                          ' ListView infos
    Dim szRootkey As String
    Dim szDetailKey As String
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    Call GetKontextRoot(szRootkey, szDetailKey, "")                     ' Kontext ermitteln
    If bTV Then                                                         ' entsprechenden Key/Tag holen
        szKeyArray = Split(TVMain.SelectedItem.Key, TV_KEY_SEP)         ' Key von TV
        With TVNode                                                     ' TVNode infos einlesen
            Call objTools.GetTVNodeInfofromXML(App.Path & "\" & INI_XMLFILE, TVMain.SelectedItem.Tag, _
                .szTag, .szText, .szKey, .bShowSubnodes, .szSQL, .szWhere, .lngImage, .bShowKontextMenue)
            If Not .bShowKontextMenue Then GoTo exithandler             ' Kein kontextmenü
        End With
    Else
        szKeyArray = Split(LVMain.Tag, TV_KEY_SEP)                      ' Key vom LV
        With LVInfo                                                     ' LV infos einlesen
            Call objTools.GetLVInfoFromXML(App.Path & "\" & INI_XMLFILE, LVMain.Tag, _
                .szSQL, .szTag, .szWhere, .lngImage, .bValueList, .bListSubNodes, _
                .bEdit, .bNew, .bSelectNode, .AltImage, .AltImgField, .AltImgValue, _
                .DelFlagField, .bShowKontextMenue, .bDelete)
            If Not .bShowKontextMenue Then GoTo exithandler             ' Kein kontextmenü
        End With
    End If
    Select Case UCase(szRootkey)                                        ' Kontextmenü aus Rootkey ermitteln
    Case UCase("Ausschreibungen")                                       ' Ausschreibung
        PopupMenu kmnuListAusschreibung                                 ' Kontext Menü Ausschreibungen anzeigen
    Case UCase("Personen"), UCase("Teilnehmer")                         ' Personen
        PopupMenu kmnuListPersonen                                      ' Kontext Menü Personen anzeigen
    Case UCase("Bewerber")                                              ' Bewerber
        PopupMenu kmnuListBewerber                                      ' Kontext Menü Bewerber anzeigen
    Case UCase("Notare"), UCase("Notare bestellt"), _
            UCase("Notare ausgeschieden")                               ' Notare
        PopupMenu kmnuListNotare                                        ' Kontext Menü Notare anzeigen
    Case UCase("Bewerbung"), UCase("Bewerbungen")                       ' Bewerbungen
        PopupMenu kmnuListBewerbung
    Case UCase("Benutzerverwaltung"), UCase("Benutzer")                 ' Stammdaten Benutzer
        PopupMenu kmnuListUser                                          ' Kontext Menü Benutzerveraltung anzeigen
    'Case UCase("Stammdaten"), UCase("Landgerichte"), _
            UCase("Amtsgerichte")                                        ' Stammdaten Gerichte
'    Case UCase("Fortbildungen")                                         ' Fortbildungen
    'Case UCase("Ausgeschriebene Stellen"), _
            UCase("Stellen"), UCase("StellenJahr")                       ' Stellen
    'Case UCase("Dokumente"), UCase("Letzte Woche"), _
            UCase("Letzter Monat")                                       ' Dokumente
    'Case UCase("Aktenort")                                              ' Aktenort
    'Case UCase("Disziplinarmaßnahmen")                                  ' Disziplinarmaßnahmen
'    Case UCase("Vorgang")                                               ' Vorgang (zur Zeit nicht benutzt)
    'Case UCase("Forderungen")                                           ' Forderungen
    Case Else                                                           ' Sonstiges
        kmnuListNew.Visible = LVInfo.bNew                               ' evtl. Neu disablen
        kmnuListEdit.Visible = LVInfo.bEdit                             ' evtl. Bearbeiten disablen
        kmnuListDel.Visible = LVInfo.bDelete                            ' evtl. Löschen disablen = Bearbeiten (erstmal)
        PopupMenu kmnuListDefault                                       ' Default kontextmenü
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

On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Call MousePointerHourglas(Me)                                   ' Stundenglass
    DoEvents                                                        ' Andere Events zulassen
    Call SetColumnOrder(LVMain, ColumnHeader)                       ' Spalten sortieren
    
exithandler:
On Error Resume Next
    Call MousePointerDefault(Me)                                    ' Stundenglass
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
    Call HandleLVItemDblKlick                                       ' ListItem DoppelKlick behandeln
End Sub

Private Sub TVMain_NodeClick(ByVal node As MSComctlLib.node)
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

On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren

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
    
    Dim cNode As node                                               ' Aktueller TreeNode

On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren


    Set cNode = GetSelectTreeNode(TVMain)                           ' Aktueller TreeNode ermitteln
    If Not cNode Is Nothing Then
        If KeyCode = 13 Then Call HandleNodeClick(cNode)            ' Enter (wie Klick)
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
    Call HandleMenueKlick("kmnuListEdit")                           ' Default Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListNew_Click()
    Call HandleMenueKlick("kmnuListNew")                            ' Default Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListDel_Click()
    Call HandleMenueKlick("kmnuListDel")                            ' Default Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListPrint_Click()
    Call HandleMenueKlick("kmnuListPrint")                          ' Default Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListNewBewerbung_Click()
    Call HandleMenueKlick("kmnuListNewBewerbung")                   ' Bewerbungen Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListDelBewerbung_Click()
    Call HandleMenueKlick("kmnuListDelBewerbung")                   ' Bewerbungen Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListEditBewerberPerson_Click()
    Call HandleMenueKlick("kmnuListEditBewerberPerson")             ' Bewerbungen Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListEditBewerbung_Click()
    Call HandleMenueKlick("kmnuListEditBewerbung")                  ' Bewerbungen Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListDelUser_Click()
    Call HandleMenueKlick("kmnuListDelUser")                        ' Benutzerverwaltung Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListChangePWD_Click()
    Call HandleMenueKlick("kmnuListChangePWD")                      ' Benutzerverwaltung Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListEditUser_Click()
    Call HandleMenueKlick("kmnuListEditUser")                       ' Benutzerverwaltung Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListNewUser_Click()
    Call HandleMenueKlick("kmnuListNewUser")                        ' Benutzerverwaltung Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListPrintUser_Click()
    Call HandleMenueKlick("kmnuListPrintUser")                      ' Benutzerverwaltung Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListDelBewerber_Click()
    Call HandleMenueKlick("kmnuListDelBewerber")                    ' Bewerber Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListNewDocBewerber_Click()
    Call HandleMenueKlick("kmnuListNewDocBewerber")                 ' Bewerber Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListPrintBewerber_Click()
    Call HandleMenueKlick("kmnuListPrintBewerber")                  ' Bewerber Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListEditBewerber_Click()
    Call HandleMenueKlick("kmnuListEditBewerber")                   ' Bewerber Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListNewBewerber_Click()
    Call HandleMenueKlick("kmnuListNewBewerber")                    ' Bewerber Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListSearchBewerber_Click()
    Call HandleMenueKlick("kmnuListSearchBewerber")                 ' Bewerber Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListEditNotar_Click()
    Call HandleMenueKlick("kmnuListEditNotar")                      ' Notar Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListNewNotar_Click()
    Call HandleMenueKlick("kmnuListNewNotar")                       ' Notar Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListSearchNotar_Click()
    Call HandleMenueKlick("kmnuListSearchNotar")                    ' Notar Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListDelNotar_Click()
    Call HandleMenueKlick("kmnuListDelNotar")                       ' Notar Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListNewDocNotar_Click()
    Call HandleMenueKlick("kmnuListNewDocNotar")                    ' Notar Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListPrintNotar_Click()
    Call HandleMenueKlick("kmnuListPrintNotar")                     ' Notar Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListDelPerson_Click()
    Call HandleMenueKlick("kmnuListDelPerson")                      ' Personen Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListNewDocPerson_Click()
    Call HandleMenueKlick("kmnuListNewDocPerson")                   ' Personen Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListEditPerson_Click()
    Call HandleMenueKlick("kmnuListEditPerson")                     ' Personen Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListNewPerson_Click()
    Call HandleMenueKlick("kmnuListNewPerson")                      ' Personen Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListPrintPerson_Click()
    Call HandleMenueKlick("kmnuListPrintPerson")                    ' Personen Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListSearchPerson_Click()
    Call HandleMenueKlick("kmnuListSearchPerson")                   ' Personen Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListDelAusschreibung_Click()
    Call HandleMenueKlick("kmnuListDelAusschreibung")               ' Ausschreibung Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListEditAusschreibung_Click()
    Call HandleMenueKlick("kmnuListEditAusschreibung")              ' Ausschreibung Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListNewAusschreibung_Click()
    Call HandleMenueKlick("kmnuListNewAusschreibung")               ' Ausschreibung Kontextmenü Klick behandeln
End Sub

Private Sub kmnuListPrintAusschreibung_Click()
    Call HandleMenueKlick("kmnuListPrintAusschreibung")             ' Ausschreibung Kontextmenü Klick behandeln
End Sub

                                                                    ' *****************************************
                                                                    ' Menü Events
Private Sub mnuDateiPrint_Click()
    Call HandleMenueKlick("mnuDateiPrint")                          ' Menü Klick behandeln
End Sub

Private Sub mnuDateiExit_Click()
    Call HandleMenueKlick("mnuDateiExit")                           ' Menü Klick behandeln
End Sub

Private Sub mnuDateiNew_Click()
    'Call HandleMenueKlick("mnuDateiNew")                           ' Menü Klick behandeln
End Sub

Private Sub mnuDateiNewAusschreibung_Click()
     Call HandleMenueKlick("mnuDateiNewAusschreibung")              ' Menü Klick behandeln
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

Private Sub mnuEditExcelExport_Click()
     Call HandleMenueKlick("mnuEditExcelExport")                    ' Menü Klick behandeln
End Sub

Private Sub mnuEditDocImport_Click()
     Call HandleMenueKlick("mnuEditDocImport")                      ' Menü Klick behandeln
End Sub

Private Sub mnuEditVerzImport_Click()
     Call HandleMenueKlick("mnuEditVerzImport")                      ' Menü Klick behandeln
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

Private Sub mnuInfoReadMe_Click()
    Call HandleMenueKlick("mnuInfoReadMe")                          ' Menü Klick behandeln
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
Private Sub lblNavBar_Click(Index As Integer)
    Call HandleNavbarClick(lblNavBar(Index).Tag)                    ' Navbar Click behandeln
End Sub

Private Sub lblNavBar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call DeHoverAll                                                 ' Alle NavBar labels dehovern
    Call HoverLabel(lblNavBar(Index), True)                         ' Aktives Label Hovern
    
    'lblNavBar(Index).MouseIcon = LoadPicture(App.Path & "\Hand.cur")
End Sub

Private Sub LVMain_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
     If Button = 2 Then Call ShowKontextMenu                        ' Kontextmenü anzeigen
End Sub

Private Sub LVMain_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Me.MousePointer = vbHourglass Then Exit Sub
    Call DeHoverAll                                                 ' Navbar enthovern
    Me.MousePointer = vbDefault                                     ' Mauszeiger wieder Normal
End Sub

Private Sub StatusBarMain_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Me.MousePointer = vbHourglass Then Exit Sub
    Call DeHoverAll                                                 ' Navbar enthovern
    Me.MousePointer = vbDefault                                     ' Mauszeiger wieder Normal
End Sub

Private Sub ToolbarMain_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Me.MousePointer = vbHourglass Then Exit Sub
    Call DeHoverAll                                                 ' Navbar enthovern
    Me.MousePointer = vbDefault                                     ' Mauszeiger wieder normal
End Sub

Private Sub TVMain_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Me.MousePointer = vbHourglass Then Exit Sub
    Call DeHoverAll                                                 ' Navbar enthovern
    Me.MousePointer = vbDefault                                     ' Mauszeiger wieder normal
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Me.MousePointer = vbHourglass Then Exit Sub
    SplitFlag = True                                                ' Verschieben des Splitters akttivieren
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
On Error GoTo Errorhandler

    If Me.MousePointer = vbHourglass Then Exit Sub
    Call DeHoverAll                                                 ' Navbar enthovern
    Me.MousePointer = vbSizeWE                                      ' Mauszeiger für Größen änderung
    If SplitFlag Then                                               ' Wenn Spliter Verschoben wird
        curlngSplitposProz = x / ScaleWidth                         ' Neue Pos bestimmen
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

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
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
Public Property Get GetLV() As ListView
    Set GetLV = Me.LVMain
End Property

Public Property Get GetTV() As TreeView
    Set GetTV = Me.TVMain
End Property

Public Property Get GetDBConn() As Object
    Set GetDBConn = ThisDBCon
End Property

Public Property Set SetDBConn(dbCon As Object)
    Set ThisDBCon = dbCon
End Property

