VERSION 5.00
Begin VB.Form frmDocImport 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Dokumenten Import"
   ClientHeight    =   2610
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6105
   Icon            =   "frmDocImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNewPerson 
      Caption         =   "Neu"
      Height          =   315
      Left            =   5280
      TabIndex        =   6
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdNewStelle 
      Caption         =   "Neu"
      Height          =   315
      Left            =   5280
      TabIndex        =   5
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdNewAusschreibung 
      Caption         =   "Neu"
      Height          =   315
      Left            =   5280
      TabIndex        =   17
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtAusschreibung 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton cmdSuchenAusschreibung 
      Height          =   315
      Left            =   4920
      Picture         =   "frmDocImport.frx":0D22
      Style           =   1  'Grafisch
      TabIndex        =   15
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdSuchenStelle 
      Height          =   315
      Left            =   4920
      Picture         =   "frmDocImport.frx":12AC
      Style           =   1  'Grafisch
      TabIndex        =   10
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton cmdSuchenPers 
      Height          =   315
      Left            =   4920
      Picture         =   "frmDocImport.frx":1836
      Style           =   1  'Grafisch
      TabIndex        =   9
      Top             =   1800
      Width           =   375
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
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1440
      Width           =   3255
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
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1800
      Width           =   3255
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "..."
      Height          =   315
      Left            =   5640
      TabIndex        =   3
      Top             =   360
      Width           =   315
   End
   Begin VB.TextBox txtFile 
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   5535
   End
   Begin VB.CommandButton cmdESC 
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtIDPerson 
      DataField       =   "FK010013"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   3840
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtIDStelle 
      DataField       =   "FK012013"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   3840
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtIDAusschreibung 
      DataField       =   "FK010013"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   3840
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblRelation 
      Caption         =   "Geben Sie Bitte an mit welchen daten das Dokument verknüpft werden soll."
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   720
      Width           =   5775
   End
   Begin VB.Label lblAusschreibung 
      Caption         =   "Ausschreibung"
      Height          =   315
      Left            =   120
      TabIndex        =   19
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblStelle 
      Caption         =   "Stelle"
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblPerson 
      Caption         =   "Person"
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblFile 
      Caption         =   "Wählen Sie eine Datei zum Importieren aus"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmDocImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                                     ' Variaben Deklaration erzwingen
Private Const MODULNAME = "frmDocImport"                            ' Modulname für Fehlerbehandlung
Private bInit As Boolean                                            ' Wird True gesetzt wenn Alle werte geladen
'Private bDirty As Boolean                                           ' Wird True gesetzt wenn Daten verändert wurden
'Private bNew  As Boolean                                            ' Wird gesetzt wenn neuer DS sonst Update
Private bModal As Boolean                                           ' Ist Modal Geöffnet
Private bVerzImport As Boolean                                      ' Soll ein ganzer Ordner importiert werden

Private frmParent As Object                                         ' Aufrufendes Form
Private ThisDBCon As Object                                         ' Akt DB Verbindung

Public bCancel As Boolean                                           ' Abbruchvariable
Public szDefFolder As String
Public szImportFile As String
Public szImportFolder As String

Private Sub Form_Load()
    bCancel = False
    
End Sub

Public Sub InitForm(frmParent As Object, dbCon As Object, _
        Optional szDetailKey As String, _
        Optional bDialog As Boolean, _
        Optional bFolder As Boolean)

On Error GoTo Errorhandler

    bVerzImport = bFolder
    Set frmParent = frmParent                                       ' Aufrufendes Form Übergeben
    bInit = True                                                    ' Wir initialisieren das Form-> andere vorgänge nicht ausführen
    Set ThisDBCon = dbCon                                           ' Aktuelle DB Verbindung übernehmen
    
    Call RefreshRelFields                                           ' evtl.  Werte zu Relations ID holen
    Call CheckSelect                                                ' Prüfen og OK enabled
    
exithandler:
On Error Resume Next

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitForm", errNr, errDesc)
    Resume exithandler
End Sub

Private Function RefreshRelFields()
    
On Error GoTo Errorhandler

    Call RefreshRelField(Me, txtStelle, txtIDStelle, _
            "SELECT TOP 1 BEZIRK012 + ' ' + CONVERT(varchar(20),FRIST012,104) FROM STELLEN012", _
            "ID012 =", True)
    Call RefreshRelField(Me, txtPerson, txtIDPerson, _
            "SELECT TOP 1 NACHNAME010 + ', ' + ISNULL(VORNAME010,'') FROM RA010", _
            "ID010 =", True)
    Call RefreshRelField(Me, txtAusschreibung, txtIDAusschreibung, _
            "SELECT TOP 1 AZ020 + ' (' + CAST(Jahr020 as varchar(5)) + ')' FROM AUSSCHREIBUNG020", _
            "ID020 =", True)
            
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

Private Function CheckSelect() As Boolean

On Error Resume Next

    If txtFile.Text <> "" Then                                      ' Wenn File angegeben
        If txtIDPerson.Text <> "" Or txtIDStelle.Text <> "" _
                Or txtIDAusschreibung <> "" Then
            CheckSelect = True
        End If
    End If
    cmdOK.Enabled = CheckSelect                                     ' OK Button enablen
    Err.Clear
    
End Function

Private Sub HandleKeyDown(frmEdit As Form, KeyCode As Integer, Shift As Integer)
    
On Error Resume Next

    Call HandleKeyDownEdit(Me, KeyCode, Shift)
    Call frmParent.HandleGlobalKeyCodes(KeyCode, Shift)

End Sub

                                                                    ' *****************************************
                                                                    ' Botton Events
Private Sub cmdFile_Click()
    Dim szTmp As String

On Error Resume Next

    If bVerzImport Then                                             ' Wenn Ordner Import dann Ordbnerauswahl zeigen
        szTmp = objObjectBag.SelectFolder("Ordner wählen", "")
        
    Else                                                            ' Sonst Datei auswahl zeigen
        szTmp = objObjectBag.OpenFile("", Me, "")
    End If
    
    If szTmp <> "" Then
        txtFile.Text = szTmp
        Call CheckSelect
    End If
    Err.Clear
    
End Sub

Private Sub cmdNewPerson_Click()
    txtIDPerson.Text = frmMain.OpenEditForm("Personen", "", Me, True) ' Leeres form Bewerber zum eingeben öffnen
    If txtIDPerson <> "" Then Call RefreshRelFields                 ' evtl. Rükgabewert eintragen
End Sub

Private Sub cmdNewStelle_Click()
    txtIDStelle.Text = frmMain.OpenEditForm("Stellen", "", Me, True) ' Leeres form Stelle zum eingeben öffnen
    If txtIDStelle <> "" Then Call RefreshRelFields                 ' evtl. Rükgabewert eintragen
End Sub

Private Sub cmdNewAusschreibung_Click()
    txtIDAusschreibung.Text = frmMain.OpenEditForm("Ausschreibung", "", Me, True) ' Leeres form Ausschreibung zum eingeben öffnen
    If txtIDStelle <> "" Then Call RefreshRelFields                 ' evtl. Rükgabewert eintragen
End Sub

Private Sub cmdSuchenAusschreibung_Click()

    Dim NewID As String                                             ' Evtl. gefundene ID
    Dim szSuchtext As String                                        ' Evtl. Suchbegriff
    
On Error GoTo Errorhandler

    If txtIDAusschreibung.Text = "" Then szSuchtext = txtAusschreibung ' Evtl Eingabe als Suchbegriff übernehmen
    NewID = ShowSearch(ThisDBCon, "Ausschreibungen", "AZ", szSuchtext) ' Suchen
    If NewID <> txtIDAusschreibung.Text And NewID <> "" Then        ' Wenn was Gefunden
        txtIDAusschreibung = NewID                                  ' Neue ID übernehmen
        Call RefreshRelFields                                       ' Inhalt Anzeige zu ID holen
        Call CheckSelect                                            ' Prüfen OK enabled
    End If
exithandler:

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "cmdSuchenAusschreibung_Click", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub cmdSuchenPers_Click()
    
    Dim NewID As String                                             ' Evtl. gefundene ID
    Dim szSuchtext As String                                        ' Evtl. Suchbegriff
    Dim SearchKey As String                                         ' Name der zu verwendenen Suche
    Dim FieldKey As String                                          ' Name des Suchfeldes
    Dim SearchTitle As String                                       ' Alternativer Suchtitel
    
On Error GoTo Errorhandler

     If txtIDStelle <> "" Then                                      ' Suche durch Stelle einschränken
        SearchKey = "PersonenNachStellen"
        FieldKey = "Nachname"
        SearchTitle = "Suche Stelle in Auschreibung " & txtAusschreibung
    Else                                                            ' Alle Stellen Durchsuchen
        SearchKey = "Personen"
        FieldKey = "Nachname"
    End If
    
    If txtIDPerson.Text = "" Then szSuchtext = txtPerson            ' Evtl Eingabe als Suchbegriff übernehmen
    NewID = ShowSearch(ThisDBCon, SearchKey, FieldKey, szSuchtext, txtIDStelle) ' Suchen
    If NewID <> txtIDPerson.Text And NewID <> "" Then               ' Wenn was Gefunden
        txtIDPerson = NewID                                         ' Neue ID übernehmen
        Call RefreshRelFields                                       ' Inhalt Anzeige zu ID holen
        Call CheckSelect                                            ' Prüfen OK enabled
    End If
    
exithandler:

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "cmdSuchenPers_Click", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub cmdSuchenStelle_Click()

    Dim NewID As String                                             ' Evtl. gefundene ID
    Dim szSuchtext As String                                        ' Evtl. Suchbegriff
    Dim SearchKey As String                                         ' Name der zu verwendenen Suche
    Dim FieldKey As String                                          ' Name des Suchfeldes
    Dim SearchTitle As String                                       ' Alternativer Suchtitel
    
On Error GoTo Errorhandler

    If txtIDAusschreibung <> "" Then                                ' Suche durch Ausschreibung einschränken
        SearchKey = "StellenNachAusschreibung"
        FieldKey = "Bezirk"
        SearchTitle = "Suche Stelle in Auschreibung " & txtAusschreibung
    Else                                                            ' Alle Stellen Durchsuchen
        SearchKey = "Stellen"
        FieldKey = "Bezirk"
    End If
    
    If txtIDStelle.Text = "" Then szSuchtext = txtStelle            ' Evtl Eingabe als Suchbegriff übernehmen
    NewID = ShowSearch(ThisDBCon, SearchKey, FieldKey, szSuchtext, txtIDAusschreibung, SearchTitle) ' Suchen
    If NewID <> txtIDStelle.Text And NewID <> "" Then               ' Wenn was Gefunden
        txtIDStelle.Text = NewID                                    ' Neue ID übernehmen
        Call RefreshRelFields                                       ' Inhalt Anzeige zu ID holen
        Call CheckSelect                                            ' Prüfen OK enabled
    End If
exithandler:

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "cmdSuchenPers_Click", errNr, errDesc)
    Resume exithandler
End Sub

Private Sub cmdEsc_Click()
    bCancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub
                                                                    ' *****************************************
                                                                    ' KeyDown Events
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)                          ' Standart Keydown Handlen
End Sub

Private Sub txtFile_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)                          ' Standart Keydown Handlen
    If KeyCode = 13 Then Call cmdFile_Click                         ' ENTER -> Datei auswählen
End Sub

Private Sub txtAusschreibung_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)                          ' Standart Keydown Handlen
    If KeyCode = 13 Then Call cmdSuchenAusschreibung_Click          ' ENTER -> Ausschreibung Suchen
End Sub

Private Sub txtPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then                                            ' Enter gedrückt
        Call cmdSuchenPers_Click                                    ' ENTER -> Person Suchen
        Exit Sub
    End If
    Call HandleKeyDown(Me, KeyCode, Shift)                          ' Standart Keydown Handlen
End Sub

Private Sub txtIDPerson_LostFocus()
    If txtIDPerson = "" Then Call cmdSuchenPers_Click               ' ENTER -> Person Suchen
End Sub

Private Sub txtStelle_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)                          ' Standart Keydown Handlen
    If KeyCode = 13 Then Call cmdSuchenStelle_Click                 ' ENTER -> Stelle Suchen
End Sub


                                                                    ' *****************************************
                                                                    ' Properties
'Public Property Get IsNew() As Boolean
'    IsNew = bNew
'End Property
'
'Public Property Get ID() As String
'    ID = szID
'End Property

'Public Property Get IsDirty() As Boolean
'    IsDirty = bDirty
'End Property
'
'Public Property Let SetDirty(Dirty As Boolean)
'    bDirty = Dirty
'End Property

Public Property Get GetDBConn() As Object
    Set GetDBConn = ThisDBCon
End Property

Public Property Let SetOrdnerImport(OrdnerImport As Boolean)
    bVerzImport = OrdnerImport
End Property

'Public Property Get GetXMLPath() As String
'    GetXMLPath = szIniFilePath
'End Property



