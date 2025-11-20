VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmValuelist 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Werteliste"
   ClientHeight    =   2985
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5610
   Icon            =   "frmValuelist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView LVValuelist 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   405
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmValuelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public fEdit As frmEdit
Public CtlIndex As Integer
Public szSecRelation As String

Private Const MODULNAME = "frmValuelist"

Private Sub cmdEsc_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    fEdit.txt01(CtlIndex) = txtValue.Text
    Me.Hide
End Sub

Private Sub LVValuelist_DblClick()
    txtValue.Text = LVValuelist.SelectedItem.Text
    Call cmdOK_Click
End Sub

Private Sub LVValuelist_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
    txtValue.Text = Item.Text
End Sub
