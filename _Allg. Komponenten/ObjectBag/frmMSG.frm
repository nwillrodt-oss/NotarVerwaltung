VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMSG 
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   1545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame FrameMSG 
      Height          =   1455
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin MSComctlLib.ProgressBar ProgressBarMSG 
         Height          =   375
         Left            =   120
         Negotiate       =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Min             =   1e-4
         Scrolling       =   1
      End
      Begin VB.Label lblAction 
         Caption         =   "Aktion"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   4815
      End
      Begin VB.Label lblMSG 
         BackStyle       =   0  'Transparent
         Caption         =   "Message"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULNAME = "frmMSG"

Private Sub Form_Activate()
    'Call OnTop(Me)
    'Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
End Sub

