VERSION 5.00
Begin VB.Form frmVorgangSelect 
   Caption         =   "Form1"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   6690
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4455
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   6615
      Begin VB.OptionButton optStep 
         Caption         =   "Option1"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6255
      End
      Begin VB.CheckBox chkStep 
         Caption         =   "Check1"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.TextBox txtStelle 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblStelle 
      Caption         =   "Stelle"
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmVorgangSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULNAME = "frmVorgangSelect"

Private bInit As Boolean       ' Wird True gesetzt wenn Alle werte geladen
Public bDirty As Boolean       ' Wird True gesetzt wenn Daten verändert wurden
Public bNew  As Boolean        ' Wird gesetzt wenn neuer DS sonst Update

Private szSQL As String        ' SQL
Private szWhere As String      ' Where Klausel
Private szIniFilePath As String ' Pfad der Ini datei

Private szRootkey As String         ' = Benutzer
Private szDetailKey As String       ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)

Private rsWorkflow As ADODB.Recordset
Private rsVorgang As ADODB.Recordset

Public lngFrametopPos As Integer
Public lngFrameLeftPos As Integer
Public lngFrameWidth As Integer
Public lngFrameHeight As Integer

Private lngStep As Integer

Public ID As String
Private ThisDBCon As Object     ' Aktuelle DB Verbindung

Private Sub Form_Load()

On Error GoTo Errorhandler

 '   Call EditFormLoad(Me, szRootkey)
    
'    lngFrametopPos = TabStrip1.Top + 360
'    lngFrameLeftPos = 120
'    lngFrameWidth = TabStrip1.Width - 240
'    lngFrameHeight = TabStrip1.Height - 480

    'FrameInfo.Visible = False
    
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
    Call EditFormUnload(Me)
End Sub

Public Function InitEditForm(DBCon As Object, DetailKey As String)

    Dim i As Integer ' counter
    Dim tmpArray() As String
    
On Error GoTo Errorhandler
    
    bInit = True                    ' Wir initialisieren das Form-> andere vorgänge nicht ausführen
    Set ThisDBCon = DBCon           ' Aktuelle DB Verbindung übernehmen
    szRootkey = "Ausschreibung"            ' für Caption
    
    If InStr(DetailKey, ";") Then
        tmpArray = Split(DetailKey, ";")
    On Error Resume Next
        szDetailKey = tmpArray(0)
    Else
        szDetailKey = DetailKey ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)
    End If
    
    szIniFilePath = objObjectBag.Getappdir & objObjectBag.GetINIFile
    szSQL = "SELECT * FROM WORKFLOW006 WHERE Caption006 = '" & szRootkey & "' ORDER BY TABNAME006, STEP006"
    Set rsWorkflow = ThisDBCon.fillrs(szSQL, False)
    
    'szSQL = objTools.GetINIValue(szIniFilePath, INI_EDITSQL, szRootkey)
    'szWhere = objTools.GetINIValue(szIniFilePath, INI_EDITSQL, "WHERE" & szRootkey)
    
    'If szDetailKey <> "" Then szWhere = szWhere & "CAST('" & szDetailKey & "' as uniqueidentifier)"
    
     ' Liste für Combo AktenOrt
    'Call FillCmbListWithSQL(cmbAktenort, "SELECT 'Fristenfach' As Aktenort UNION SELECT Nachname001 + ', ' + Vorname001 As Aktenort FROM User001", ThisDBCon)
    
    'Call InitAdoDC(Me, ThisDBCon, szSQL, szWhere)
    
    Call InitSteps
    
    Me.Refresh
    
    
'    txtIDPers = Pers_ID
'    txtBewID = Bew_ID
'
'    If bNew Then
'        Adodc1.Recordset.AddNew
'        txtEingetragenAm.Text = Format(CDate(Now()), "dd.mm.yy")
'        txtEingetragenVon.Text = objObjectBag.getusername()
'    Else
'        ID = DetailKey
'        txtEingetragenAm.Text = Format(CDate(txtCreate.Text), "dd.mm.yy")
'        txtEingetragenVon.Text = txtCreateFrom.Text
'    End If
    
    
    
'    Call RefreshRelFields
'    Call InitFramePersonenDaten              ' Frame Benutzer informationen Initialisieren
'    Call InitFrameFortbildungen
'    Call InitFrameBewerbungen
'    Call InitFrameBewerberDaten
'    Call InitFrameAktenort
'    Call InitFrameInfo(Me)

'    ' Liste der Pflichtfelder holen
'    szNotEmptyList = UCase(objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_NONEMPTY, szRootKey))
'    ' Liste der gesperrten Felder holen
'    szLockedFiledList = UCase(objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_NONEDIT, szRootKey))
'    ' Liste der ausgeblendeten Felder holen
'    szNonVisibleFieldList = UCase(objTools.GetINIValue(App.Path & "\" & INI_FILENAME, INI_NONVIS, szRootKey))
    
    Call SetEditFormCaption(Me, szRootkey, "")
    
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

Private Sub InitSteps()

    Dim i As Integer

On Error GoTo Errorhandler
    
    If rsWorkflow Is Nothing Then GoTo exithandler
    If rsWorkflow.RecordCount = 0 Then GoTo exithandler
    rsWorkflow.MoveFirst
'    chkStep(0).Caption = rsWorkflow.Fields("STEPTITLE006").Value
'    chkStep(0).Tag = rsWorkflow.Fields("STEP006").Value
'    If chkStep(0).Tag < lngStep Then chkStep(0).Value = True
    optStep(0).Caption = rsWorkflow.Fields("STEPTITLE006").Value
    optStep(0).Tag = rsWorkflow.Fields("STEP006").Value
    If optStep(0).Tag < lngStep Then optStep(0).Value = True
    
    
    rsWorkflow.MoveNext
    For i = 1 To rsWorkflow.RecordCount - 1
'        Load chkStep(i)
        Load optStep(i)
'        chkStep(i).Top = chkStep(i - 1).Top + chkStep(i - 1).Height + 50
'        chkStep(i).Caption = rsWorkflow.Fields("STEPTITLE006").Value
'        chkStep(i).Tag = rsWorkflow.Fields("STEP006").Value
'        chkStep(i).Visible = True
'        If chkStep(i).Tag < lngStep Then chkStep(i).Value = True
        optStep(i).Top = optStep(i - 1).Top + optStep(i - 1).Height + 50
        optStep(i).Caption = rsWorkflow.Fields("STEPTITLE006").Value
        optStep(i).Tag = rsWorkflow.Fields("STEP006").Value
        optStep(i).Visible = True
        If optStep(i).Tag < lngStep Then optStep(i).Value = True
        rsWorkflow.MoveNext
    Next i
    
exithandler:

Exit Sub
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitEditForm", errNr, errDesc)
    Resume exithandler
End Sub
