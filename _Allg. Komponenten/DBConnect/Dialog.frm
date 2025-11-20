VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Name des Dialogfeldes"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox txtDBName 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   750
      Width           =   2655
   End
   Begin VB.TextBox txtDBUser 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1230
      Width           =   2655
   End
   Begin VB.TextBox txtPWD 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1710
      Width           =   2655
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblServer 
      Caption         =   "Servername:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   270
      Width           =   1575
   End
   Begin VB.Label lblDBName 
      Caption         =   "Datenbankname:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   750
      Width           =   1575
   End
   Begin VB.Label lblDBUser 
      Caption         =   "Benutzername:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1230
      Width           =   1575
   End
   Begin VB.Label lblPWD 
      Caption         =   "Passwort:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1710
      Width           =   1575
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
