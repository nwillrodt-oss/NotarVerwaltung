VERSION 5.00
Begin VB.Form frmDBConnProperties 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Datenbank Verbindung"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   Icon            =   "frmDBConnectParams.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ComboBox cmbSavedCon 
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Text            =   "Wählen Sie eine gespeicherte Verbindung aus der Liste"
      Top             =   360
      Width           =   4815
   End
   Begin VB.CommandButton cmdDBPath 
      Caption         =   "..."
      Height          =   300
      Left            =   4560
      TabIndex        =   13
      Top             =   1440
      Width           =   375
   End
   Begin VB.OptionButton OptAcc 
      Caption         =   "MS Access"
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   720
      Width           =   2415
   End
   Begin VB.OptionButton optMSSQL 
      Caption         =   "MS SQL Server"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   1935
   End
   Begin VB.CheckBox chkNT 
      Caption         =   "NT Authentifikation"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   4455
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtPWD 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox txtDBUser 
      Height          =   300
      Left            =   1920
      TabIndex        =   6
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox txtDBName 
      Height          =   300
      Left            =   1920
      TabIndex        =   5
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtServer 
      Height          =   300
      Left            =   1920
      TabIndex        =   4
      Top             =   1050
      Width           =   2655
   End
   Begin VB.Label lblSavedCon 
      Caption         =   "Gespeicherte Verbindungen"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblPWD 
      Caption         =   "Passwort:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblDBUser 
      Caption         =   "Benutzername:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblDBName 
      Caption         =   "Datenbankname:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblServer 
      Caption         =   "Servername:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "frmDBConnProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                                     ' Variaben Deklaration erzwingen
Option Compare Text                                                 ' Sortierreihenfolge festlegen
Private Const MODULNAME = "frmDBConnProperties"                     ' Modulname für Fehlerbehandlung

Public bCancel As Boolean                                           ' Abbruch bedingung

Private objObjectBag As Object                                      ' Objectbag class
Private objError As Object                                          ' Error Class
Private objOptions As Object                                        ' Optionen Class
Private objConn As Object                                           ' Connection Object

Private ConnectionsArray()

Dim szOldSQLServer As String
Dim szOldSQLDB As String
Dim szOldSQLUser As String
Dim szOldSQLPwd As String
Dim bOldNt As Boolean
Dim szOldAccDBPath As String
Dim szOldAccUser As String
Dim szOldAccPwd As String

Private Sub Form_Load()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Me.cmdOK.Enabled = False                                        ' OK Button erstmal disablen
    Err.Clear                                                       ' Evtl. error clearen
End Sub

Private Sub Form_Activate()
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    'Call OnTop(Me)                                                  ' Fenster OnTop setzen
    Call CheckOKEnabled                                             ' Prüfen ob OK Enablet werden kann
exithandler:
Exit Sub                                                            ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "Form_Activate", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Sub

Public Function InitObjectBag(objOb As Object)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Set objObjectBag = objOb
    If Not objObjectBag Is Nothing Then
        Set objError = objObjectBag.GetErrorObj
        Set objOptions = objObjectBag.GetOptionsObj()
        Set objConn = objObjectBag.GetDBConObj()
        Call objObjectBag.CheckFormStyle(Me)
        Call FillCmbWithSavedCons
    End If
    Err.Clear                                                       ' Evtl. error clearen
End Function

Public Sub ChangeDBType()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If Me.OptAcc.Value = True Then ' Access
        'Me.OptAcc.Value = True
        Me.lblServer.Visible = False
        Me.optMSSQL.Value = False
        Me.txtServer.Visible = False
        Me.chkNT.Visible = False
        Me.lblDBName.Caption = "Datenbank:"
        Me.cmdDBPath.Visible = True
        szOldSQLServer = Me.txtServer
        szOldSQLDB = Me.txtDBName
        szOldSQLUser = Me.txtDBUser
        szOldSQLPwd = Me.txtPWD
        Me.txtDBName = szOldAccDBPath
        Me.txtDBUser = szOldAccUser
        Me.txtPWD = szOldAccPwd
        
    Else    ' SQL Server
        'Me.OptAcc.Value = False
        Me.lblServer.Visible = True
        Me.optMSSQL.Value = True
        Me.txtServer.Visible = True
        Me.chkNT.Visible = True
        Me.lblDBName.Caption = "Datenbankname:"
        Me.cmdDBPath.Visible = False
        szOldAccDBPath = Me.txtDBName
        szOldAccUser = Me.txtDBUser
        szOldAccPwd = Me.txtPWD
        Me.txtServer = szOldSQLServer
        Me.txtDBName = szOldSQLDB
        Me.txtDBUser = szOldSQLUser
        Me.txtPWD = szOldSQLPwd
    End If
    Err.Clear                                                       ' Evtl Error Cearen
End Sub

Private Sub chkNT_Click()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Me.txtDBUser.Enabled = Not CBool(chkNT.Value)
    Me.txtPWD.Enabled = Not CBool(chkNT.Value)
    Call CheckOKEnabled
    Err.Clear                                                       ' Evtl Error Cearen
End Sub

Private Sub cmdDBPath_Click()
    Dim Filter As String
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Filter = "Access Datenbanken (*.mdb , *.mde)" & Chr$(0) & "*.mdb;*.mde" & Chr$(0) & Chr$(0)
    Me.txtDBName.Text = objObjectBag.OpenFile(Filter, Me)
    Err.Clear                                                       ' Evtl Error Cearen
End Sub

Private Sub cmdEsc_Click()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    bCancel = True                                                  ' Abbruch Flag setzen
    Me.Hide                                                         ' Form Ausblenden
    Err.Clear                                                       ' Evtl Error Cearen
End Sub

Private Sub cmdOK_Click()
    Dim lngType As Integer
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Me.Hide
    If Me.OptAcc.Value = True Then
        lngType = 1
        If objConn.GetADODBConn(lngType, "", txtDBName.Text, txtDBUser.Text, txtPWD.Text) Then
            
        Else
                ' Keine Verbindung
        End If
    Else
        lngType = 2
        If objConn.GetADODBConn(lngType, txtServer.Text, txtDBName.Text, txtDBUser.Text, txtPWD.Text, CBool(chkNT)) Then
            
        Else
            ' Keine Verbindung
        End If
    End If
exithandler:
Exit Sub                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "cmdOK_Click", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Sub

Private Function GetConnFromCmb(Index As Integer)
'Private Function GetConnFromCmb(szReadableCon As String)
    Dim szConnect  As String
    Dim lngConTyp As Integer
    Dim szDBName As String
    Dim szServer As String
    Dim szUser As String
    Dim szPWD As String
    Dim bNtAut As Boolean
    szConnect = objConn.ReadConnectionFromList(Index)
'    szConnect = objConn.FindReadableConInReg(szReadableCon)
    If szConnect <> "" Then
        If objConn.GetConParamsFromRegString(szConnect, lngConTyp, szServer, _
                szDBName, szUser, szPWD, bNtAut) Then
            If lngConTyp = 1 Then Me.OptAcc.Value = True
            If lngConTyp = 2 Then Me.optMSSQL.Value = True
            Me.txtServer = szServer
            Me.txtDBName = szDBName
            Me.txtDBUser = szUser
            Me.txtPWD = szPWD
            Me.chkNT = bNtAut
        End If
    End If
End Function

Private Sub cmbSavedCon_Change()
'    Call GetConnFromCmb(cmbSavedCon)
    Call GetConnFromCmb(cmbSavedCon.ListIndex)
End Sub

Private Sub cmbSavedCon_Click()
'    Call GetConnFromCmb(cmbSavedCon)
    Call GetConnFromCmb(cmbSavedCon.ListIndex)
End Sub

Private Sub OptAcc_Click()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call ChangeDBType
    Err.Clear                                                       ' Evtl. error clearen
End Sub

Private Sub optMSSQL_Click()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call ChangeDBType
    Err.Clear                                                       ' Evtl. error clearen
End Sub

Private Sub txtDBName_Change()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call CheckOKEnabled                                             ' Prüfen ob OK Button enabeld werden kann
    Err.Clear                                                       ' Evtl. error clearen
End Sub

Private Sub txtDBUser_Change()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call CheckOKEnabled                                             ' Prüfen ob OK Button enabeld werden kann
    Err.Clear                                                       ' Evtl. error clearen
End Sub

Private Sub txtPWD_Change()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call CheckOKEnabled                                             ' Prüfen ob OK Button enabeld werden kann
    Err.Clear                                                       ' Evtl. error clearen
End Sub

Private Sub txtServer_Change()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Call CheckOKEnabled                                             ' Prüfen ob OK Button enabeld werden kann
    Err.Clear                                                       ' Evtl. error clearen
End Sub

Private Sub FillCmbWithSavedCons(Optional bNoSQL As Boolean, Optional bNoAccess As Boolean)
    Dim i As Integer                                                ' Counter
    Dim szTmpValue As String
    Dim szConnArray() As String                                     ' Array mit den einzelnen Verbindungs Parametern
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Me.cmbSavedCon.Clear                                            ' Evtl. Alte einträge löschen
    For i = 0 To objConn.GetMaxRegCon
        szTmpValue = objConn.ReadConnectionFromList(i)
        If szTmpValue <> "" Then                                    ' Wenn Verbingungsert gefunden
            Call cmbSavedCon.AddItem(objConn.GetReadableConnect(szTmpValue))
        End If
    Next i                                                          ' Nächste Verbindung suchen
exithandler:
Exit Sub
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "FillCmbWithSavedCons", errNr, errDesc)  ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Sub

Private Sub CheckOKEnabled()
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If Me.OptAcc.Value = True Then
        If txtDBName.Text <> "" Then
            Me.cmdOK.Enabled = True
        End If
    Else
        If txtServer.Text <> "" And txtDBName.Text <> "" Then
            If (txtDBUser.Text <> "") Or CBool(chkNT.Value) Then
                Me.cmdOK.Enabled = True
            End If
        End If
        If txtDBName.Text <> "" And txtDBUser.Text <> "" And txtPWD.Text <> "" And txtServer.Text <> "" Then
        
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
    Call objError.Errorhandler(MODULNAME, "CheckOKEnabled", errNr, errDesc)
    Resume exithandler
End Sub
'Private Function GetLastConnection()
'    Dim szRegKey As String
'    Dim szConName As String
'    Dim szConnArray() As String
'On Error GoTo ErrorHandler                                          ' Fehlerbehandlung aktivieren
'
'    'szRegKey = "SOFTWARE\" & objObjectBag.GetAppTitle()
'    'szConName = objRegTools.ReadRegValue("HKCU", szRegKey, OPTION_LASTCON)
'    'szConName = objOptions.GetOptionByName(OPTION_LASTCON)
'    If szConName <> "" Then
'        szConnArray = Split(szConName, ":")
'
'        Me.txtServer.Text = szConnArray(0)
'        Me.txtDBName.Text = szConnArray(1)
'        Me.txtDBUser.Text = szConnArray(2)
'        Me.txtPWD.Text = szConnArray(3)                             ' Passwort Ver und entschlüsseln
'    End If
'
'exithandler:
'On Error Resume Next
'
'Exit Function
'ErrorHandler:
'    Dim errNr As String
'    Dim errDesc As String
'    errNr = Err.Number
'    errDesc = Err.Description
'    Err.Clear
'    Call objError.ErrorHandler(MODULNAME, "GetLastConnection", errNr, errDesc)
'    Resume exithandler
'End Function


