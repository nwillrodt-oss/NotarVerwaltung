VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEditPersonen 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Personenkartei"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9645
   FillColor       =   &H80000002&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   9645
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame FrameFristen 
      Caption         =   "Fristen"
      Height          =   615
      Left            =   240
      TabIndex        =   156
      Top             =   5160
      Width           =   1215
      Begin MSComctlLib.ListView LVFristen 
         Height          =   375
         Left            =   120
         TabIndex        =   157
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         View            =   3
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
   Begin VB.Frame FrameBewerberDaten 
      Caption         =   "Punkteberechnung"
      Height          =   4815
      Left            =   720
      TabIndex        =   47
      Top             =   1200
      Width           =   8175
      Begin VB.TextBox txtPunkte 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   151
         TabStop         =   0   'False
         Top             =   4305
         Width           =   615
      End
      Begin VB.TextBox txtFaktor 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         Height          =   255
         Index           =   13
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   150
         TabStop         =   0   'False
         Top             =   4305
         Width           =   495
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Rechts
         Height          =   285
         Index           =   13
         Left            =   5760
         TabIndex        =   82
         Top             =   4305
         Width           =   615
      End
      Begin VB.TextBox txtPunkte 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   4020
         Width           =   615
      End
      Begin VB.TextBox txtFaktor 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         Height          =   255
         Index           =   12
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   146
         TabStop         =   0   'False
         Top             =   4020
         Width           =   495
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Rechts
         Height          =   285
         Index           =   12
         Left            =   5760
         TabIndex        =   81
         Top             =   4020
         Width           =   615
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Rechts
         Height          =   285
         Index           =   11
         Left            =   5760
         TabIndex        =   80
         Top             =   3735
         Width           =   615
      End
      Begin VB.TextBox txtFaktor 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         Height          =   255
         Index           =   11
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   142
         TabStop         =   0   'False
         Top             =   3735
         Width           =   495
      End
      Begin VB.TextBox txtPunkte 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   141
         TabStop         =   0   'False
         Top             =   3735
         Width           =   615
      End
      Begin VB.TextBox txtPunkte 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   3450
         Width           =   615
      End
      Begin VB.TextBox txtFaktor 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         Height          =   255
         Index           =   10
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   3450
         Width           =   495
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Rechts
         Height          =   285
         Index           =   10
         Left            =   5760
         TabIndex        =   79
         Top             =   3450
         Width           =   615
      End
      Begin VB.TextBox txtPunkte 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   3165
         Width           =   615
      End
      Begin VB.TextBox txtFaktor 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         Height          =   255
         Index           =   9
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   3165
         Width           =   495
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Rechts
         Height          =   285
         Index           =   9
         Left            =   5760
         TabIndex        =   78
         Top             =   3165
         Width           =   615
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Rechts
         Height          =   285
         Index           =   8
         Left            =   5760
         TabIndex        =   77
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtFaktor 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         Height          =   255
         Index           =   8
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox txtPunkte 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Rechts
         Height          =   285
         Index           =   7
         Left            =   5760
         TabIndex        =   76
         Top             =   2595
         Width           =   615
      End
      Begin VB.TextBox txtFaktor 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         Height          =   255
         Index           =   7
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   2595
         Width           =   495
      End
      Begin VB.TextBox txtPunkte 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   2595
         Width           =   615
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Rechts
         Height          =   285
         Index           =   6
         Left            =   5760
         TabIndex        =   75
         Top             =   2310
         Width           =   615
      End
      Begin VB.TextBox txtFaktor 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         Height          =   255
         Index           =   6
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   2310
         Width           =   495
      End
      Begin VB.TextBox txtPunkte 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   2310
         Width           =   615
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Rechts
         Height          =   285
         Index           =   5
         Left            =   5760
         TabIndex        =   74
         Top             =   2025
         Width           =   615
      End
      Begin VB.TextBox txtFaktor 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         Height          =   255
         Index           =   5
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   2025
         Width           =   495
      End
      Begin VB.TextBox txtPunkte 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   2025
         Width           =   615
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Rechts
         Height          =   285
         Index           =   4
         Left            =   5760
         TabIndex        =   73
         Top             =   1740
         Width           =   615
      End
      Begin VB.TextBox txtFaktor 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         Height          =   255
         Index           =   4
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   1740
         Width           =   495
      End
      Begin VB.TextBox txtPunkte 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   1740
         Width           =   615
      End
      Begin VB.TextBox txtPunkte 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   1455
         Width           =   615
      End
      Begin VB.TextBox txtPunkte 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   1170
         Width           =   615
      End
      Begin VB.TextBox txtPunkte 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   885
         Width           =   615
      End
      Begin VB.TextBox txtFaktor 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         Height          =   255
         Index           =   3
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   1455
         Width           =   495
      End
      Begin VB.TextBox txtFaktor 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         Height          =   255
         Index           =   2
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   1170
         Width           =   495
      End
      Begin VB.TextBox txtFaktor 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         Height          =   255
         Index           =   1
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   885
         Width           =   495
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Rechts
         Height          =   285
         Index           =   3
         Left            =   5760
         TabIndex        =   72
         Top             =   1455
         Width           =   615
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Rechts
         Height          =   285
         Index           =   2
         Left            =   5760
         TabIndex        =   71
         Top             =   1170
         Width           =   615
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Rechts
         Height          =   285
         Index           =   1
         Left            =   5760
         TabIndex        =   70
         Top             =   885
         Width           =   615
      End
      Begin VB.TextBox txtPunkte 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtFaktor 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         Height          =   255
         Index           =   0
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Rechts
         Height          =   285
         Index           =   0
         Left            =   5760
         TabIndex        =   69
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtIDBew 
         Height          =   285
         Left            =   6600
         TabIndex        =   67
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtIDStelle 
         Height          =   285
         Left            =   2880
         TabIndex        =   66
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cmbStelle 
         Height          =   315
         Left            =   4080
         TabIndex        =   65
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lbltext 
         Caption         =   "Label2"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   153
         Top             =   4305
         Width           =   5640
      End
      Begin VB.Label lblX 
         Caption         =   "x"
         Height          =   255
         Index           =   13
         Left            =   6480
         TabIndex        =   152
         Top             =   4305
         Width           =   135
      End
      Begin VB.Label lbltext 
         Caption         =   "Label2"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   149
         Top             =   4020
         Width           =   5640
      End
      Begin VB.Label lblX 
         Caption         =   "x"
         Height          =   255
         Index           =   12
         Left            =   6480
         TabIndex        =   148
         Top             =   4020
         Width           =   135
      End
      Begin VB.Label lblX 
         Caption         =   "x"
         Height          =   255
         Index           =   11
         Left            =   6480
         TabIndex        =   144
         Top             =   3735
         Width           =   135
      End
      Begin VB.Label lbltext 
         Caption         =   "Label2"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   143
         Top             =   3735
         Width           =   5640
      End
      Begin VB.Label lbltext 
         Caption         =   "Label2"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   128
         Top             =   3450
         Width           =   5640
      End
      Begin VB.Label lblX 
         Caption         =   "x"
         Height          =   255
         Index           =   10
         Left            =   6480
         TabIndex        =   127
         Top             =   3450
         Width           =   135
      End
      Begin VB.Label lbltext 
         Caption         =   "Label2"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   124
         Top             =   3165
         Width           =   5640
      End
      Begin VB.Label lblX 
         Caption         =   "x"
         Height          =   255
         Index           =   9
         Left            =   6480
         TabIndex        =   123
         Top             =   3165
         Width           =   135
      End
      Begin VB.Label lblX 
         Caption         =   "x"
         Height          =   255
         Index           =   8
         Left            =   6480
         TabIndex        =   118
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label lbltext 
         Caption         =   "Label2"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   115
         Top             =   2880
         Width           =   5640
      End
      Begin VB.Label lblX 
         Caption         =   "x"
         Height          =   255
         Index           =   7
         Left            =   6480
         TabIndex        =   114
         Top             =   2595
         Width           =   135
      End
      Begin VB.Label lbltext 
         Caption         =   "Label2"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   111
         Top             =   2595
         Width           =   5640
      End
      Begin VB.Label lblX 
         Caption         =   "x"
         Height          =   255
         Index           =   6
         Left            =   6480
         TabIndex        =   110
         Top             =   2310
         Width           =   135
      End
      Begin VB.Label lbltext 
         Caption         =   "Label2"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   107
         Top             =   2310
         Width           =   5640
      End
      Begin VB.Label lblX 
         Caption         =   "x"
         Height          =   255
         Index           =   5
         Left            =   6480
         TabIndex        =   106
         Top             =   2025
         Width           =   135
      End
      Begin VB.Label lbltext 
         Caption         =   "Label2"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   103
         Top             =   2025
         Width           =   5640
      End
      Begin VB.Label lblX 
         Caption         =   "x"
         Height          =   255
         Index           =   4
         Left            =   6480
         TabIndex        =   102
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label lbltext 
         Caption         =   "Label2"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   99
         Top             =   1740
         Width           =   5640
      End
      Begin VB.Label lblX 
         Caption         =   "x"
         Height          =   255
         Index           =   3
         Left            =   6480
         TabIndex        =   95
         Top             =   1455
         Width           =   135
      End
      Begin VB.Label lblX 
         Caption         =   "x"
         Height          =   255
         Index           =   2
         Left            =   6480
         TabIndex        =   94
         Top             =   1170
         Width           =   135
      End
      Begin VB.Label lblX 
         Caption         =   "x"
         Height          =   255
         Index           =   1
         Left            =   6480
         TabIndex        =   93
         Top             =   885
         Width           =   135
      End
      Begin VB.Label lblX 
         Caption         =   "x"
         Height          =   255
         Index           =   0
         Left            =   6480
         TabIndex        =   92
         Top             =   600
         Width           =   135
      End
      Begin VB.Label lbltext 
         Caption         =   "Label2"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   88
         Top             =   1455
         Width           =   5640
      End
      Begin VB.Label lbltext 
         Caption         =   "Label2"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   87
         Top             =   1170
         Width           =   5640
      End
      Begin VB.Label lbltext 
         Caption         =   "Label2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   86
         Top             =   885
         Width           =   5640
      End
      Begin VB.Label lbltext 
         Caption         =   "Label2"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   85
         Top             =   600
         Width           =   5640
      End
      Begin VB.Label lblStelle 
         Caption         =   "für Bewerbung auf folgende Stelle:"
         Height          =   315
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame FramePersonenDaten 
      BorderStyle     =   0  'Kein
      Caption         =   "Daten"
      Height          =   3975
      Left            =   360
      TabIndex        =   39
      Top             =   1200
      Width           =   8655
      Begin VB.CheckBox chkVerstorben 
         Caption         =   "Verstorben"
         DataField       =   "VERSTORBEN010"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   7440
         TabIndex        =   22
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Frame FrameAmtssitz 
         Caption         =   "Amtssitz"
         Height          =   735
         Left            =   4320
         TabIndex        =   154
         Top             =   480
         Width           =   4095
         Begin VB.TextBox txtAmtsPLZ 
            DataField       =   "AMTPLZ010"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   960
            TabIndex        =   15
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtAmtsOrt 
            DataField       =   "AMTORT010"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   1800
            TabIndex        =   16
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "PLZ / Ort"
            Height          =   315
            Left            =   120
            TabIndex        =   155
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox txtAusgeschieden 
         DataField       =   "AUSGESCH010"
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
         Left            =   5880
         TabIndex        =   21
         Top             =   2400
         Width           =   1190
      End
      Begin VB.TextBox txtBestelt 
         DataField       =   "BESTELLT010"
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
         Left            =   1680
         TabIndex        =   20
         Top             =   2400
         Width           =   1190
      End
      Begin VB.TextBox txtAnwaltSeit 
         DataField       =   "ANWALTSEIT010"
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
         Left            =   4320
         TabIndex        =   8
         Top             =   120
         Width           =   1190
      End
      Begin VB.TextBox txtGeb 
         DataField       =   "GEB010"
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
         Left            =   1320
         TabIndex        =   7
         Top             =   120
         Width           =   1190
      End
      Begin MSComCtl2.DTPicker DTAusgeschieden 
         Height          =   315
         Left            =   5880
         TabIndex        =   33
         Top             =   2400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   78118913
         CurrentDate     =   39392
      End
      Begin VB.TextBox txtAz 
         DataField       =   "AZ010"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   5880
         TabIndex        =   19
         Top             =   2040
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTBestellt 
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
         Left            =   1680
         TabIndex        =   32
         Top             =   2400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   78118913
         CurrentDate     =   39392
      End
      Begin VB.ComboBox cmbLG 
         DataField       =   "LG010"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   5880
         TabIndex        =   18
         Top             =   1680
         Width           =   1815
      End
      Begin VB.ComboBox cmbAg 
         DataField       =   "AG010"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   5880
         TabIndex        =   17
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtExaErgebnis 
         DataField       =   "EXANOTE010"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   7800
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComCtl2.DTPicker DTAnwaltSeit 
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
         Left            =   4320
         TabIndex        =   31
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   78118913
         CurrentDate     =   39223
      End
      Begin VB.TextBox txtBem 
         DataField       =   "BEM010"
         DataSource      =   "Adodc1"
         Height          =   915
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   2760
         Width           =   6735
      End
      Begin MSComCtl2.DTPicker DTGeb 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "d.MM.yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   30
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   78118913
         CurrentDate     =   39223.4371759259
      End
      Begin VB.Frame Framekanzlei 
         Caption         =   "Kanzlei"
         Height          =   1815
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   4095
         Begin VB.TextBox txtKanzleiFax 
            DataField       =   "KFAX010"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   960
            TabIndex        =   14
            Top             =   1320
            Width           =   3015
         End
         Begin VB.TextBox txtKanzeliTel 
            DataField       =   "KTEL010"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   960
            TabIndex        =   13
            Top             =   960
            Width           =   3015
         End
         Begin VB.TextBox txtKanzleiOrt 
            DataField       =   "KORT010"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   1800
            TabIndex        =   12
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txtkanzeliPLZ 
            DataField       =   "KPLZ010"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   960
            TabIndex        =   11
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtkanzleiStr 
            DataField       =   "KSTR010"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   960
            TabIndex        =   10
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label lblKanzleiFax 
            Caption         =   "Fax.:"
            Height          =   315
            Left            =   120
            TabIndex        =   44
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label lblKanzelTel 
            Caption         =   "Tel.:"
            Height          =   315
            Left            =   120
            TabIndex        =   43
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lbKanzleiPLZOrt 
            Caption         =   "PLZ / Ort"
            Height          =   315
            Left            =   120
            TabIndex        =   42
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblKanzleiStr 
            Caption         =   "Strasse"
            Height          =   315
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Label lblEntlassen 
         Caption         =   "Ausgeschieden am"
         Height          =   255
         Left            =   4320
         TabIndex        =   145
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblAZ 
         Caption         =   "Az (VI)"
         Height          =   255
         Left            =   4320
         TabIndex        =   138
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblLG 
         Caption         =   "am LG"
         Height          =   255
         Left            =   5280
         TabIndex        =   133
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblAG 
         Caption         =   "am AG"
         Height          =   255
         Left            =   5280
         TabIndex        =   132
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblBestellt 
         Caption         =   "Bestellt am"
         Height          =   315
         Left            =   120
         TabIndex        =   131
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblZugelassen 
         Caption         =   "Zugelassen"
         Height          =   255
         Left            =   4320
         TabIndex        =   130
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblExaErgebnis 
         Caption         =   "Ergebnis Staatsprüfung"
         Height          =   255
         Index           =   0
         Left            =   5880
         TabIndex        =   68
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblAnwaltSeit 
         Caption         =   "Eintr. in Liste d. Rechtsanwälte"
         Height          =   375
         Left            =   3000
         TabIndex        =   63
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label lblBem 
         Caption         =   "Bemerkungen"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label lblGeb 
         Caption         =   "Geburtsdatum"
         Height          =   315
         Left            =   120
         TabIndex        =   45
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   0
      Picture         =   "frmEditPersonen.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Datensatz speichern"
      Top             =   6120
      Width           =   375
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   375
      Left            =   480
      Picture         =   "frmEditPersonen.frx":058A
      Style           =   1  'Grafisch
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Datensatz löschen"
      Top             =   6120
      Width           =   375
   End
   Begin VB.CommandButton cmdWord 
      Height          =   375
      Left            =   960
      Picture         =   "frmEditPersonen.frx":0914
      Style           =   1  'Grafisch
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Neues Anschreiben"
      Top             =   6120
      Width           =   375
   End
   Begin VB.Frame FrameInfo 
      Caption         =   "Datensatz Informationen"
      Height          =   2535
      Left            =   1080
      TabIndex        =   50
      Top             =   1920
      Width           =   7815
      Begin VB.TextBox txtID 
         Appearance      =   0  '2D
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "ID010"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   360
         Width           =   5295
      End
      Begin VB.TextBox txtCreateFrom 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "CFROM010"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   720
         Width           =   5295
      End
      Begin VB.TextBox txtCreate 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "CREATE010"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   1080
         Width           =   5295
      End
      Begin VB.TextBox txtModifyFrom 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "MFROM010"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   1440
         Width           =   5295
      End
      Begin VB.TextBox txtModify 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'Kein
         DataField       =   "MODIFY010"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   1800
         Width           =   5295
      End
      Begin VB.Label lblID 
         Caption         =   "Datensatz ID"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblCreateFrom 
         Caption         =   "erstellt von"
         Height          =   315
         Left            =   120
         TabIndex        =   59
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblCreate 
         Caption         =   "erstellt am"
         Height          =   315
         Left            =   120
         TabIndex        =   58
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblModifyFrom 
         Caption         =   "geändert von"
         Height          =   315
         Left            =   120
         TabIndex        =   57
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblModify 
         Caption         =   "geändert am"
         Height          =   315
         Left            =   120
         TabIndex        =   56
         Top             =   1800
         Width           =   975
      End
   End
   Begin VB.Frame FrameForderungen 
      Caption         =   "Foderungen"
      Height          =   615
      Left            =   240
      TabIndex        =   139
      Top             =   4560
      Width           =   975
      Begin MSComctlLib.ListView LVForderungen 
         Height          =   375
         Left            =   120
         TabIndex        =   140
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         View            =   3
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
   Begin VB.Frame FrameDizip 
      Caption         =   "Disziplinarmaßnahmen"
      Height          =   735
      Left            =   120
      TabIndex        =   136
      Top             =   4080
      Width           =   1335
      Begin MSComctlLib.ListView LVDizip 
         Height          =   375
         Left            =   120
         TabIndex        =   137
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
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
   Begin VB.Frame FrameDokumente 
      Caption         =   "Dokumente"
      Height          =   735
      Left            =   120
      TabIndex        =   134
      Top             =   2640
      Width           =   1335
      Begin MSComctlLib.ListView LVDokumente 
         Height          =   375
         Left            =   120
         TabIndex        =   135
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
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
   Begin VB.ComboBox cmbStatus 
      DataField       =   "STATUS010"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "frmEditPersonen.frx":0C9E
      Left            =   8160
      List            =   "frmEditPersonen.frx":0CA0
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.Frame FrameAktenort 
      Caption         =   "Aktenstandort"
      Height          =   735
      Left            =   120
      TabIndex        =   119
      Top             =   1920
      Width           =   1215
      Begin MSComctlLib.ListView LVAktenort 
         Height          =   375
         Left            =   120
         TabIndex        =   120
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
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
   Begin VB.Frame FrameFortbildungen 
      Caption         =   "Fortbildungsveranstaltungen"
      Height          =   735
      Left            =   120
      TabIndex        =   48
      Top             =   3360
      Width           =   1335
      Begin MSComctlLib.ListView LVFortbildungen 
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
   Begin VB.Frame FrameBewerbungen 
      Caption         =   "Bewerbungen"
      Height          =   735
      Left            =   120
      TabIndex        =   61
      Top             =   1200
      Width           =   1455
      Begin MSComctlLib.ListView LVBewerbungen 
         Height          =   375
         Left            =   120
         TabIndex        =   62
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
      Height          =   330
      Left            =   3720
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.ComboBox cmbTitel 
      DataField       =   "TITEL010"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   4680
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtNamensZusatz 
      DataField       =   "NAMEZUS010"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   3480
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox cmbAnrede 
      DataField       =   "ANREDE010"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "frmEditPersonen.frx":0CA2
      Left            =   0
      List            =   "frmEditPersonen.frx":0CAC
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtNachname 
      DataField       =   "NACHNAME010"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox txtVorname 
      DataField       =   "VORNAME010"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   5760
      TabIndex        =   4
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Übernehmen"
      Height          =   375
      Left            =   8400
      TabIndex        =   26
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   7080
      TabIndex        =   25
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   5760
      TabIndex        =   24
      Top             =   6120
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip TabStripPerson 
      Height          =   5295
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9340
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   10
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Daten"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dokumente"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Punkteberechnung"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bewerbungen"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fortbildungen"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Disziplinarm."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ford-Verz."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Aktenort"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fristen"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPersonen.frx":0CBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPersonen.frx":0FD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPersonen.frx":12F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPersonen.frx":188A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPersonen.frx":1E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPersonen.frx":2AFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPersonen.frx":37D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPersonen.frx":3D72
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPersonen.frx":430C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPersonen.frx":46A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPersonen.frx":4A40
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPersonen.frx":4A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPersonen.frx":4AFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPersonen.frx":5096
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPersonen.frx":5630
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPersonen.frx":5BCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPersonen.frx":6164
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPersonen.frx":66FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   8160
      TabIndex        =   129
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblNamenszusatz 
      Caption         =   "Namenszusatz"
      Height          =   195
      Left            =   3480
      TabIndex        =   38
      Top             =   120
      Width           =   1065
   End
   Begin VB.Label lblTitel 
      Caption         =   "Titel"
      Height          =   195
      Left            =   4680
      TabIndex        =   37
      Top             =   120
      Width           =   945
   End
   Begin VB.Label lblAnrede 
      Caption         =   "Anrede"
      Height          =   195
      Left            =   0
      TabIndex        =   36
      Top             =   120
      Width           =   945
   End
   Begin VB.Label lblNachname 
      Caption         =   "Nachname"
      Height          =   195
      Left            =   1080
      TabIndex        =   35
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label lblVorname 
      Caption         =   "Vorname"
      Height          =   255
      Left            =   5760
      TabIndex        =   34
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu kmnuLVFortbildungen 
      Caption         =   "KontextMenueLVFortbildungen"
      Visible         =   0   'False
      Begin VB.Menu kmnuLVFortbildungAdd 
         Caption         =   "Fortbildung hinzufügen"
      End
      Begin VB.Menu kmnuLVFortbildungDel 
         Caption         =   "Fortbildung entfernen"
      End
   End
   Begin VB.Menu kmnuLVBewerbungen 
      Caption         =   "KontextMenueLVBewerbungen"
      Visible         =   0   'False
      Begin VB.Menu kmnuLVBewerbungAdd 
         Caption         =   "Bewerbung hinzufügen"
      End
      Begin VB.Menu kmnuLVBewerbungDel 
         Caption         =   "Bewerbung entfernen"
      End
   End
   Begin VB.Menu kmnuLVDokumente 
      Caption         =   "KontextmenueLVDokumente"
      Visible         =   0   'False
      Begin VB.Menu kmnuLVDokumentOpen 
         Caption         =   "Dokument anzeigen"
      End
      Begin VB.Menu kmnuLVDokumentNew 
         Caption         =   "Neues Dokument erstellen"
      End
      Begin VB.Menu kmnuLVDokumentDel 
         Caption         =   "Dokument löschen"
      End
      Begin VB.Menu kmnuLVDokumentImport 
         Caption         =   "Dokument Importieren"
      End
   End
   Begin VB.Menu kmnuLVFristen 
      Caption         =   "KontextmenueLVFristen"
      Visible         =   0   'False
      Begin VB.Menu kmnuLVFristenNew 
         Caption         =   "Neue Frist eintragen"
      End
      Begin VB.Menu kmnuLVFristDel 
         Caption         =   "Frist löschen"
      End
   End
   Begin VB.Menu kmnuLVAktenort 
      Caption         =   "KontextmenueLVAktenort"
      Visible         =   0   'False
      Begin VB.Menu kmnuLVAktenortNew 
         Caption         =   "Neuen Aktenort angeben"
      End
   End
   Begin VB.Menu kmnuLVDisz 
      Caption         =   "KontextmenueLVDisz"
      Visible         =   0   'False
      Begin VB.Menu kmnuLVDiszNew 
         Caption         =   "Neue Maßnahme"
      End
   End
   Begin VB.Menu kmnuLVForderungen 
      Caption         =   "KontextMenueLVForderungen"
      Visible         =   0   'False
      Begin VB.Menu kmnuLVForderungNew 
         Caption         =   "Neue Forderung"
      End
   End
End
Attribute VB_Name = "frmEditPersonen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                                         ' Variaben Deklaration erzwingen
Private Const MODULNAME = "frmEditPersonen"                             ' Modulname für Fehlerbehandlung

Private ThisPerson As New clsPerson                                     ' Personen Klasse (test)

Private bInit As Boolean                                                ' Wird True gesetzt wenn Alle werte geladen
Private bDirty As Boolean                                               ' Wird True gesetzt wenn Daten verändert wurden
Private bNew  As Boolean                                                ' Wird gesetzt wenn neuer DS sonst Update
Private bModal As Boolean                                               ' Ist Modal Geöffnet
Private szID As String                                                  ' DS ID
Private ThisDBCon As Object                                             ' Aktuelle DB Verbindung
Private frmParent As Form                                               ' Aufrufendes DB form
Private szIDField As String
Private ThisFramePos As FramePos                                        ' Standart Frame Position
Private bInitTabMSG As Boolean                                          ' Merker ob Hinweismeldung in TabClick bereits ausgeführt

Private szSQL As String                                                 ' SQL für a_personen
Private szWhere As String                                               ' Where Klausel
Private szIniFilePath As String                                         ' Pfad der Ini datei
Private lngImage As Integer                                             ' Imagiendex

Private szRootkey As String                                             ' = Benutzer
Private szDetailKey As String                                           ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)

Private CurrentStep As String                                           ' Aktueller WorkflowSchritt
Private OldCmbValue As String                                           ' Alter Combo wert

Private SummenIndex As Integer
                                                                        ' Recordsets mit Detaildaten
Private rsFortbildungen As ADODB.Recordset                              ' Daten Forderungen (nur Notare)
Private rsBewerbungen As ADODB.Recordset                                ' Daten Bewerbungen (Nur Bewerber)
Private rsAktenort As ADODB.Recordset                                   ' Daten Aktenort
Private rsDokumente As ADODB.Recordset                                  ' Daten Dokumente
Private rsDiszip As ADODB.Recordset                                     ' Daten Disziplinarmaßnahmen (nur Notare)
Private rsForderungen As ADODB.Recordset                                ' Daten Forderungen auß Diz. (nur Notare)
Private rsFristen As ADODB.Recordset                                    ' Daten Fristen

Private Type FramePos                                                   ' Positions Datentyp
    Top As Single                                                       ' Top position (oben)
    Left As Single                                                      ' Left Position (Links)
    Height As Single                                                    ' Height (Höhe)
    Width As Single                                                     ' Width (Breite)
End Type
' MW 09.08.11 {
Private Const TAB_NAME_DATEN = "Daten"
Private Const TAB_NAME_DOC = "Dokumente"
Private Const TAB_NAME_BEW = "Bewerbungen"
'Private Const TAB_NAME_BEWDAT = "Bewerberdaten"
Private Const TAB_NAME_BEWDAT = "Punkteberechnung"
Private Const TAB_NAME_DISZ = "Disziplinarm."
Private Const TAB_NAME_FORD = "Ford-Verz."
Private Const TAB_NAME_FORT = "Fortbildungen"
Private Const TAB_NAME_AKTE = "Aktenort"
Private Const TAB_NAME_FRIST = "Fristen"
Private Const TAB_NAME_INFO = "Info"
' MW 09.08.11 }

Private Sub Form_Activate()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bInit Or bDirty Then Exit Sub                                    ' Nicht bei initialisierung
    Call RefreshEditForm                                                ' Form daten aktualisieren
    Err.Clear                                                           ' Evtl. Error clearen
End Sub

Private Sub Form_Load()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    Call EditFormLoad(Me, szRootkey)                                    ' Allg. Formload Aufrufen
    Call InitEditButtonMenue(Me, True, True, True)                      ' Buttonleiste initialisieren
' MW 09.08.11 {
'    Call RemoveTabByCaption(TabStripPerson, "Fortbildungen")            ' bis auf weiteres ausblenden
    Call RemoveTabByCaption(TabStripPerson, TAB_NAME_FORT)              ' bis auf weiteres ausblenden
' MW 09.08.11 }
    FrameBewerberDaten.Visible = False
    With ThisFramePos
        Call GetTabStrimClientPos(TabStripPerson, .Top, .Left, _
                .Height, .Width)                                        ' Frame Positionen aus TabStrip ermitteln
    End With
    Err.Clear                                                           ' Evtl. Error clearen
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    Call SaveColumnWidth(LVAktenort, szRootkey & "LV", True)            ' Spaltenbreiten LV Aktenort sichern
    Call SaveColumnWidth(LVBewerbungen, szRootkey & "LV", True)         ' Spaltenbreiten LV Bewerbungen sichern
    Call SaveColumnWidth(LVDokumente, szRootkey & "LV", True)           ' Spaltenbreiten LV Dokumente sichern
'    Call SaveColumnWidth(LVFortbildungen, szRootkey & "LV", True)       ' Spaltenbreiten LV Fortbildungen sichern
    Call SaveColumnWidth(LVDizip, szRootkey & "LV", True)               ' Spaltenbreiten LV Diziplinarmaßn. sichern
    Call SaveColumnWidth(LVForderungen, szRootkey & "LV", True)         ' Spaltenbreiten LV Forderungen sichern
' MW 30.11.10 {
    Call SaveColumnWidth(LVFristen, szRootkey & "LV", True)             ' Spaltenbreiten LV Fristen sichern
' MW 30.11.10 }
    rsFortbildungen.Close                                               ' RS Fortbildungen schliessen
    rsBewerbungen.Close                                                 ' RS Bewerbungen schliessen
    rsAktenort.Close                                                    ' RS Aktenort schliessen
    rsDokumente.Close                                                   ' RS Dokumente schliessen
    rsDiszip.Close                                                      ' RS Diziplinarmaßn schliessen
    rsForderungen.Close                                                 ' RS Forderungen schliessen
' MW 30.11.10 {
    rsFristen.Close                                                     ' RS Fristen schliessen
' MW 30.11.10 }
    If bDirty Then szID = ""                                            ' Wenn ungespeichert ID Leeren
    If bModal Then                                                      ' Wenn Modal
        Me.Hide                                                         ' dann ausblenden
    Else                                                                ' Sonst
        Call EditFormUnload(Me)                                         ' AUs Edit Form Array entfernen
    End If
    Err.Clear                                                           ' Evtl. Error clearen
End Sub

Public Function InitEditForm(parentform As Form, dbCon As Object, DetailKey As String, Optional bDialog As Boolean)
    Dim i As Integer                                                    ' counter
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    Set frmParent = parentform
    bInit = True                                                        ' Wir initialisieren das Form-> andere vorgänge nicht ausführen
    Set ThisDBCon = dbCon                                               ' Aktuelle DB Verbindung übernehmen
    'Call ThisPerson.Init(ThisDBCon, Me.Adodc1, DetailKey)
    szRootkey = "Personenkartei"                                        ' für Caption
    szIDField = "ID010"
    szDetailKey = DetailKey                                             ' Welcher DS genau wird bearbeitet (Bedingung für Where klausel)
    bModal = bDialog                                                    ' Als Dialog anzeigen
    szIniFilePath = objObjectBag.Getappdir & objObjectBag.GetXMLFile    ' XML inifile festlegen
    Call objTools.GetEditInfoFromXML(szIniFilePath, szRootkey, szSQL, szWhere, lngImage)
    Me.Icon = ILTree.ListImages(lngImage).Picture                       ' Form Icon Setzen
    If szDetailKey = "" Then bNew = True                                ' Neuer Datensatz
    If szDetailKey <> "" Then szWhere = szWhere & "CAST('" & szDetailKey & "' as uniqueidentifier)"
     ' Liste für Combo Anrede füllen
    Call FillCmbListWithSQL(cmbAnrede, "SELECT VALUE015 FROM VALUES015 WHERE Fieldname015 = 'ANREDE010' ORDER BY ORDER015", ThisDBCon)
    ' Liste für Combo Titel füllen
    Call FillCmbListWithSQL(cmbTitel, "SELECT VALUE015 FROM VALUES015 WHERE Fieldname015 = 'TITEL010' ORDER BY ORDER015", ThisDBCon)
    ' Liste für Combo Status füllen
    Call FillCmbListWithSQL(cmbStatus, "SELECT VALUE015 FROM VALUES015 WHERE Fieldname015 = 'STATUS010' ORDER BY ORDER015", ThisDBCon)
    ' Liste für Combo LG füllen
    Call FillCmbListWithSQL(cmbLG, "SELECT LGNAME003 FROM LG003 ", ThisDBCon)
    ' Liste für Combo AG füllen
    Call FillCmbListWithSQL(cmbAg, "SELECT AGNAME004 FROM AG004 ", ThisDBCon)
    Call InitAdoDC(Me, ThisDBCon, szSQL, szWhere)                       ' ADODC Initialisieren
    Me.Refresh                                                          ' formular aktualisieren
    If bNew Then                                                        ' Wenn DS Neu
        Call FormatDTPicker(Me, DTGeb, Now())
        Call FormatDTPicker(Me, DTAnwaltSeit, Now())
        Call FormatDTPicker(Me, DTBestellt, Now())
        Adodc1.Recordset.AddNew                                         ' Neuen DS an RS anhängen
        txtID.Text = ThisDBCon.GetValueFromSQL("SELECT NewID()")        ' Neue ID (Guid) ermitteln
        szID = txtID                                                    ' ID merken
        Call GetDefaultValues(Me, szRootkey, szIniFilePath)             ' Defaultwerte holen
        bDirty = True                                                   ' Dirty da Neu
    Else
        szID = DetailKey
        Call ChangeStatus                                               ' Evtl. Status änderung (Notar / Bewerber ... ( behandeln
    End If
    Call GetLockedControls(Me)                                          ' Gelockte controls finden
    Call HiglightThisMustFields(Not (bNew Or bDirty))                   ' IndexFelder hervorheben
    Call InitFramePersonenDaten(True)                                   ' Frame Benutzer informationen Initialisieren
    Call RefereshFortbildungen(False)                                   ' Frame fortbildungen initialisieren
    Call RefereshBewerbungen(False)                                     ' Frame Bewerbungen initialisieren
    Call RefereshAktenort(False)                                        ' Frame Aktenort initialisieren
    Call RefereshDokumente(False)                                       ' Frame Dokumente initialisieren
    Call RefereshDizip(False)                                           ' Frame Diziplinarmaßnahmen initialisieren
' MW 30.11.10 {
    Call RefreshFristen(False)                                          ' Frame Fristen inistialisieren
' MW 30.11.10 }
    If Not bNew Then Call PunkteBerechnungNeu                           ' Wenn neuer DS -> Punkteberechnung initialisieren
    Call InitFrameBewerberDatenNeu                                      ' Frame Bewerberdaten initialisieren
    Call RefereshFoderungen(True)                                       ' frame Forderungen initialisieren
    Call InitFrameInfo(Me)                                              ' Info Frame initialisieren
    Call SetEditFormCaption(Me, szRootkey, txtNachname & ", " & txtVorname) ' Form Caption setzen mit Nach & Vorname
    Call CheckUpdate(Me)                                                ' Buttons evtl. enablen / disablen
exithandler:
On Error Resume Next                                                    ' Hier keine fehler mehr
    bInit = False                                                       ' Initialisierung dieses Forms beendet
    Err.Clear                                                           ' Evtl Error claren
    Exit Function                                                       ' Function beenden
Errorhandler:
    Dim errNr As Long                                                   ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "InitEditForm", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Function

Public Sub HiglightThisMustFields(Optional bDeHiglight As Boolean)
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    Call HiglightMustFields(Me, bDeHiglight)                            ' Alle PK ind IndexFields entfärben
    Call HiglightMustField(Me, txtAZ, bDeHiglight)                      ' Aktenzeichen hervorheben
    Call HiglightMustField(Me, cmbAnrede, bDeHiglight)                  ' Anrede
    Err.Clear                                                           ' evtl. error clearen
End Sub

Private Sub InitFramePersonenDaten(Optional bVisible As Boolean)
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    With ThisFramePos
        Call FramePersonenDaten.Move(.Left, .Top, .Width, .Height)      ' Frame Positionieren
    End With
    FramePersonenDaten.Visible = bVisible                               ' Sichtbar ?
    Err.Clear                                                           ' evtl. error clearen
End Sub

Private Sub RefereshBewerbungen(Optional bVisible As Boolean)
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    LVBewerbungen.Tag = "Personen\Bewerber\*\Bewerbungen"
    Set rsBewerbungen = RefreshFrame(Me, FrameBewerbungen, LVBewerbungen, szRootkey, _
            "Bewerbungen", bVisible)                                    ' Frame Positionieren und Daten holen
    Err.Clear                                                           ' evtl. error clearen
End Sub

Private Sub RefereshFortbildungen(Optional bVisible As Boolean)
' Zur zeit werden keine Fortbildungen verwendet. Dehalb hier nur visible
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    FrameFortbildungen.Visible = bVisible
    Err.Clear                                                           ' evtl. error clearen
End Sub

Private Sub RefereshDokumente(Optional bVisible As Boolean)
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    LVDokumente.Tag = "Personen\Bewerber\*\Dokumente"
    Set rsDokumente = RefreshFrame(Me, FrameDokumente, LVDokumente, szRootkey, _
            "Dokumente", bVisible)                                      ' Frame Positionieren und Daten holen
    Err.Clear                                                           ' evtl. error clearen
End Sub

Private Sub RefereshAktenort(Optional bVisible As Boolean)
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    LVAktenort.Tag = "Personen\Bewerber\*\Aktenort"
    Set rsAktenort = RefreshFrame(Me, FrameAktenort, LVAktenort, szRootkey, _
            "Aktenort", bVisible)                                       ' Frame Positionieren und Daten holen
    Err.Clear                                                           ' evtl. error clearen
End Sub

Private Sub RefereshDizip(Optional bVisible As Boolean)
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    LVDizip.Tag = "Personen\Bewerber\*\Disziplinarmaßnahmen"
    Set rsDiszip = RefreshFrame(Me, FrameDizip, LVDizip, szRootkey, _
            "Disziplinarmaßnahmen", bVisible)                           ' Frame Positionieren und Daten holen
    Err.Clear                                                           ' evtl. error clearen
End Sub

Private Sub RefereshFoderungen(Optional bVisible As Boolean)
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    LVForderungen.Tag = "Personen\Bewerber\*\Forderungen"
    Set rsForderungen = RefreshFrame(Me, FrameForderungen, LVForderungen, szRootkey, _
            "Forderungen", bVisible)                                    ' Frame Positionieren und Daten holen
    Err.Clear                                                           ' evtl. error clearen
End Sub

' MW 30.11.10 {
Private Sub RefreshFristen(Optional bVisible As Boolean)
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    LVFristen.Tag = "Personen\Bewerber\*\Fristen"
    Set rsFristen = RefreshFrame(Me, FrameFristen, LVFristen, szRootkey, _
            "Fristen", bVisible)                                        ' Frame Positionieren und Daten holen
    Err.Clear                                                           ' evtl. error clearen
End Sub
' MW 30.11.10 }

Private Sub ChangeStatus()
    Dim i As Integer                                                    ' Counter
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    Select Case UCase(cmbStatus.Text)
    Case UCase("NOTAR")                                                 ' Notar (aktiv)
        Me.txtAZ.Visible = True                                         ' Text Feld AZ sichtbar
        Me.lblAz.Visible = True                                         ' Beschriftung AZ sichtbar
        Me.DTAusgeschieden.Visible = True                               ' DatumsFeld Ausgeschieden sichtbar
        Me.txtAusgeschieden.Visible = True                              ' Beschriftung Ausgeschieden sichtbar
        Me.chkVerstorben.Visible = True
        Me.lblEntlassen.Visible = True
' MW 09.08.11 {
'        Call RemoveTabByCaption(TabStripPerson, "Bewerberdaten")        ' Tab Bewerberdaten ausblenden
'        Call RemoveTabByCaption(TabStripPerson, "Bewerbungen")          ' Tab Bewerbungen ausblenden
        Call RemoveTabByCaption(TabStripPerson, TAB_NAME_BEWDAT)        ' Tab Bewerberdaten ausblenden
'        Call RemoveTabByCaption(TabStripPerson, TAB_NAME_BEW)           ' Tab Bewerbungen ausblenden
' MW 09.08.11 }
        If Not IsDate(DTBestellt.Value) Then
            Call FormatDTPicker(Me, DTBestellt, Now())
        End If
    Case UCase("BEWERBER")                                              ' Bewerber auf Notarstelle
' MW 09.08.11 {
'        Call RemoveTabByCaption(TabStripPerson, "Disziplinarm.")        ' Tab Diziplinarm. ausblenden
'        Call RemoveTabByCaption(TabStripPerson, "Ford-Verz.")           ' Tab Forderungen ausblenden
'        Call RemoveTabByCaption(TabStripPerson, "Aktenort.")            ' Tab Aktenort ausblenden
        Call RemoveTabByCaption(TabStripPerson, TAB_NAME_DISZ)          ' Tab Diziplinarm. ausblenden
        Call RemoveTabByCaption(TabStripPerson, TAB_NAME_FORD)          ' Tab Forderungen ausblenden
        Call RemoveTabByCaption(TabStripPerson, TAB_NAME_AKTE)          ' Tab Aktenort ausblenden
' MW 09.08.11 }
        Me.txtAZ.Visible = False                                        ' Text Feld AZ unsichtbar
        Me.lblAz.Visible = False                                        ' Beschriftung AZ unsichtbar
        Me.chkVerstorben.Visible = False
        Me.lblEntlassen.Visible = False
        Me.DTAusgeschieden.Visible = False
        Me.txtAusgeschieden.Visible = False
    Case UCase("ausgeschieden")                                         ' Ausgeschiedener Notar
        Me.txtAZ.Visible = True                                         ' Text Feld AZ sichtbar
        Me.lblAz.Visible = True                                         ' Beschriftung AZ sichtbar
        Me.DTAusgeschieden.Visible = True                               ' DatumsFeld Ausgeschieden sichtbar
        Me.txtAusgeschieden.Visible = True                              ' Beschriftung Ausgeschieden sichtbar
        Me.chkVerstorben.Visible = True
        Me.lblEntlassen.Visible = True
        Call FormatDTPicker(Me, DTAusgeschieden, Now())
    Case UCase("a. D.")                                                 ' Notar a. D. (wie ausgeschieden)
        Me.txtAZ.Visible = True                                         ' Text Feld AZ sichtbar
        Me.lblAz.Visible = True                                         ' Beschriftung AZ sichtbar
        Me.DTAusgeschieden.Visible = True                               ' DatumsFeld Ausgeschieden sichtbar
        Me.txtAusgeschieden.Visible = True                              ' Beschriftung Ausgeschieden sichtbar
        Me.chkVerstorben.Visible = True
        Me.lblEntlassen.Visible = True
        Call FormatDTPicker(Me, DTAusgeschieden, Now())
    Case Else
' MW 09.08.11 {
'        Call RemoveTabByCaption(TabStripPerson, "Bewerberdaten")        ' Tab Bewerberdaten ausblenden
'        Call RemoveTabByCaption(TabStripPerson, "Bewerbungen")          ' Tab Bewerbungen ausblenden
'        Call RemoveTabByCaption(TabStripPerson, "Disziplinarm.")        ' Tab Diziplinarm. ausblenden
'        Call RemoveTabByCaption(TabStripPerson, "Ford-Verz.")           ' Tab Forderungen ausblenden
        Call RemoveTabByCaption(TabStripPerson, TAB_NAME_BEWDAT)        ' Tab Bewerberdaten ausblenden
'        Call RemoveTabByCaption(TabStripPerson, TAB_NAME_BEW)           ' Tab Bewerbungen ausblenden
        Call RemoveTabByCaption(TabStripPerson, TAB_NAME_DISZ)          ' Tab Diziplinarm. ausblenden
        Call RemoveTabByCaption(TabStripPerson, TAB_NAME_FORD)          ' Tab Forderungen ausblenden
' MW 09.08.11 }
    End Select
exithandler:
On Error Resume Next                                                    ' Hier keine fehler mehr
    Err.Clear                                                           ' Evtl Error claren
    Exit Sub                                                            ' Function beenden
Errorhandler:
    Dim errNr As Long                                                   ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ChangeStatus", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Sub

Private Sub InitFrameBewerberDatenNeu()
    Dim szSQL As String                                                 ' SQL Statement
    Dim rsBerechnungen As ADODB.Recordset                               ' RS mit Berechnungsvorschriften
    'Dim szValue As String
    Dim i As Integer                                                    ' Counter
    Dim szDetails As String
    Dim szField As String
    Dim szMaxWert As Double
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    If bNew Then GoTo exithandler                                       ' Nicht für neue Daten
    'bInit = True
    'If rsBewerbungen.RecordCount = 0 Then GoTo exithandler              ' Keine Bewerbungen -> Raus
    Call PosFrameAndListView(Me, FrameBewerberDaten, True)              ' Gesamten Frame Positionieren
    If cmbStelle.ListCount = 0 Then                                     ' Combo füllen
        cmbStelle.Clear
        If rsBewerbungen.RecordCount > 0 Then
            rsBewerbungen.MoveFirst
            While Not rsBewerbungen.EOF
                cmbStelle.AddItem rsBewerbungen.Fields("BewerbungsFrist").Value & "  " & rsBewerbungen.Fields("Bezirk").Value
                rsBewerbungen.MoveNext
            Wend
        End If
    End If
    If cmbStelle.Text = "" Then                                         ' Nichts ausgewählt
        If rsBewerbungen.RecordCount > 0 Then
            rsBewerbungen.MoveFirst                                     ' Ersten wählen
            cmbStelle.Text = rsBewerbungen.Fields("BewerbungsFrist").Value & "  " & rsBewerbungen.Fields("Bezirk").Value
            txtIDStelle.Text = rsBewerbungen.Fields("FK012013").Value
            txtIDBew.Text = rsBewerbungen.Fields("ID013").Value
        End If
    End If
    If rsBewerbungen.RecordCount > 0 Then                               ' Wenn Bewerbungen vorhanden
        rsBewerbungen.MoveFirst                                         ' zur ersten springen
        For i = 0 To rsBewerbungen.RecordCount                          ' Alle duchlaufen
            If rsBewerbungen.Fields("FK010013").Value = txtID And _
                rsBewerbungen.Fields("FK012013").Value = txtIDStelle Then ' Wenn dies die Aktuell Ausgewählte Bewerbung ist
                Exit For                                                ' fertig
            End If
            rsBewerbungen.MoveNext                                      ' Sonst nächste Bewerbung auswählen
        Next i
    End If
    szSQL = "SELECT * FROM BERECHNUNGEN016 ORDER BY ORDER016"
    Set rsBerechnungen = ThisDBCon.fillrs(szSQL)                        ' Berechnungsvorschriften holen
    szSQL = ""
    If rsBerechnungen Is Nothing Then GoTo exithandler                  ' Ohne Berechnungsvorschriften keine Punkte berechnung
    If rsBerechnungen.RecordCount = 0 Then GoTo exithandler             ' dito
    rsBerechnungen.MoveFirst                                            ' Erste Berechnung
    For i = 0 To lbltext.Count - 1                                      ' Alle Steuerelemente Duchgehen
        If rsBerechnungen.EOF Then                                      ' Keine Berechnung mehr
            lbltext(i).Visible = False
            txtFaktor(i).Visible = False
            txtValue(i).Visible = False
            lblX(i).Visible = False
            txtPunkte(i).Visible = False
        Else
            ' Daten für Beschreibungsfeld Holen
            szDetails = objTools.checknull(rsBerechnungen.Fields("CAPTION016").Value, "")
            szSQL = objTools.checknull(rsBerechnungen.Fields("CAPTIONSQL016").Value, "")
            If szSQL <> "" Then
                szSQL = objSQLTools.AddWhereInFullSQL(szSQL, szWhere)
                lbltext(i).Caption = objTools.checknull(ThisDBCon.GetValueFromSQL(szSQL), "")
            Else
                lbltext(i).Caption = objTools.checknull(rsBerechnungen.Fields("CAPTION016").Value, "")
            End If
            szSQL = ""
            ' Daten Für Faktoren holen
            txtFaktor(i).Text = rsBerechnungen.Fields("FAKTOR016").Value
            txtFaktor(i).Locked = True
            txtFaktor(i).TabStop = False
            If txtFaktor(i).Text = "0" Then
                lblX(i).Visible = False
                txtFaktor(i).Visible = False
                txtPunkte(i).Visible = False
            End If
            ' Value Feld Gesperrt
            txtValue(i).Tag = objTools.checknull(rsBerechnungen.Fields("SAVEFIELD016").Value, "")
            If rsBerechnungen.Fields("LOCKED016").Value = True Then
                txtValue(i).Locked = True
                txtValue(i).TabStop = False
                txtValue(i).BorderStyle = vbBSNone
                txtValue(i).BackColor = -2147483648#
            End If
            If rsBewerbungen.RecordCount > 0 And cmbStelle.Text <> "" Then
                ' Berechnete Punkte ermitteln Wenn Daten vorliegen
                szSQL = objTools.checknull(rsBerechnungen.Fields("VALUESQL016").Value, "")
                If szSQL <> "" Then
                    szSQL = objSQLTools.AddWhereInFullSQL(szSQL, szWhere)
                    txtValue(i).Text = objTools.checknull(ThisDBCon.GetValueFromSQL(szSQL), 0)
                    szMaxWert = objTools.checknull(rsBerechnungen.Fields("MAXVALUE016").Value, -1)
                    If szMaxWert > 0 And szMaxWert < CLng(txtValue(i).Text) Then
                        txtValue(i).Text = szMaxWert
                    End If
                Else
                    txtValue(i).Text = 0
                End If ' szSQL <> ""
                szSQL = ""
                
                 ' Evtbenutzer eingaben prüfen
                If txtValue(i).Locked = False And txtValue(i).Tag <> "" Then
                    szField = txtValue(i).Tag
                    If InStr(szField, ".") > 0 Then szField = Right(szField, Len(szField) - InStr(szField, "."))
                    txtValue(i).Text = rsBewerbungen.Fields(szField).Value
                End If
            Else
                txtValue(i).Text = 0
            End If  ' rsBewerbungen.RecordCount > 0
            
            
            txtPunkte(i).Tag = objTools.checknull(rsBerechnungen.Fields("PUNKTESAVEFIELD016").Value, "")
            ' Evtl. Maxwert in tag von txtPunkte eintragen
            'txtPunkte(i).Tag = objTools.checknull(rsBerechnungen.Fields("MAXWERT016").Value, "")
            txtPunkte(i).Locked = True
            txtPunkte(i).TabStop = False
            If txtPunkte(i).Tag = "" Then
                'txtPunkte(i).Text = txtValue(i).Text * txtFaktor(i).Text
                txtPunkte(i).Text = Format(txtValue(i).Text * txtFaktor(i).Text, "#0.00") ' Formatieren
            Else
                If rsBewerbungen.RecordCount > 0 Then
                    'txtPunkte(i).Text = rsBewerbungen.Fields(txtPunkte(i).Tag).Value
                    txtPunkte(i).Text = Format(rsBewerbungen.Fields(txtPunkte(i).Tag).Value, "#0.00")
                End If
            End If
            ' Sonderfall summe eintragen
            If objTools.checknull(rsBerechnungen.Fields("CAPTION016").Value, "") = "Summe" Then
'                txtPunkte(i).Tag = "Summe"
                SummenIndex = i
                lblX(i).Visible = False
                txtValue(i).Visible = False
                txtFaktor(i).Visible = False
                txtPunkte(i).Visible = True
            End If
             rsBerechnungen.MoveNext
        End If ' f rsBerechnungen.EOF
        
       
    Next i
exithandler:
On Error Resume Next                                                    ' Hier keine fehler mehr
    Err.Clear                                                           ' Evtl Error claren
    Exit Sub                                                            ' Function beenden
Errorhandler:
    Dim errNr As Long                                                   ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "InitFrameBewerberDatenNeu", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Sub

Private Sub PunkteBerechnungNeu(Optional szPersID As String, Optional szStellID As String)
    Dim szSQL As String                                                 ' SQL Statement
    Dim rsBerechnungen As ADODB.Recordset                               ' RS mit Berechnungsvorschriften
    Dim rsMitbwerber As ADODB.Recordset                                 ' RS mit Mitbewerber Daten
    Dim lngSumme As Double                                              ' Punkte Summe
    Dim szValue As String
    Dim szValue2 As String
    Dim szFaktor As String
    Dim szValueType As String                                           ' DatenTyp der Punkte (Teilweise Float , Teilweise Int)
    Dim lngPunkte As Double                                             ' Punkte Wert
    Dim szField As String
    Dim szDetails As String
    Dim szMaxWert As String
    'Dim szWhere2 As String
    Dim szSummeField As String
    Dim i As Integer                                                    ' Counter
    Dim lngAnzStellen As Integer                                        ' Anzahl der Ausgeschreibenen Stellen
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    If bNew Then GoTo exithandler                                       ' Nicht für Neue Datensätze
    rsBewerbungen.Requery
    If rsBewerbungen.RecordCount = 0 Then GoTo exithandler              ' Ohne Bewerbung keine Punkte berechnung
    szSQL = "SELECT * FROM BERECHNUNGEN016 ORDER BY ORDER016"           ' SQL Statement für Berechnungsvorschriften
    Set rsBerechnungen = ThisDBCon.fillrs(szSQL)                        ' Berechnungsvorschriften holen
    szSQL = ""                                                          ' Variable leeren -> Brauchen wir noch
    If rsBerechnungen Is Nothing Then GoTo exithandler                  ' Ohne Berechnungsvorschriften keine Punkte berechnung
    If rsBerechnungen.RecordCount = 0 Then GoTo exithandler             ' dito
' MW 04.11.11 {
    If txtIDBew = "" Then GoTo exithandler
' MW 04.11.11 }
    rsBewerbungen.MoveFirst                                             ' Erste Bewerbung
' MW 12.12.11 {
'    For i = 0 To rsBewerbungen.RecordCount                              ' Alle duchlaufen
'        If rsBewerbungen.Fields("FK010013").Value = txtIDBew Then       ' Wenn dies die Aktuell Ausgewählte Bewerbung ist
    For i = 0 To rsBewerbungen.RecordCount - 1                          ' Alle duchlaufen
        If rsBewerbungen.Fields("ID013").Value = txtIDBew Then          ' Wenn dies die Aktuell Ausgewählte Bewerbung ist
' MW 12.12.11 }
            Exit For                                                    ' fertig
        End If
        rsBewerbungen.MoveNext                                          ' Sonst nächste Bewerbung auswählen
    Next i
            
'    While Not rsBewerbungen.EOF                                         ' für jede bewerbung Punkte berechnen
        lngSumme = 0                                                    ' Summe initialisieren
        rsBerechnungen.MoveFirst                                        ' Erste Berechnung
        While Not rsBerechnungen.EOF                                    ' Jede Berechnung ausführen
            szDetails = objTools.checknull(rsBerechnungen.Fields("CAPTION016").Value, "")
            If szDetails = "Summe" Then
                szSummeField = objTools.checknull(rsBerechnungen.Fields("PUNKTESAVEFIELD016").Value, "")
                If InStr(szSummeField, ".") > 0 Then szSummeField = Right(szSummeField, Len(szSummeField) - InStr(szSummeField, "."))
            End If
            szMaxWert = objTools.checknull(rsBerechnungen.Fields("MAXVALUE016").Value, 0) ' Maximal wert holen für den Value
            szSQL = objTools.checknull(rsBerechnungen.Fields("VALUESQL016").Value, "")  ' Rechenbasis holen
            szField = objTools.checknull(rsBerechnungen.Fields("SAVEFIELD016").Value, "") ' Speicherfeld holen
            If szSQL <> "" Then                                         ' SQL angegeben
                szSQL = objSQLTools.AddWhereInFullSQL(szSQL, szWhere)   ' Where Statement ins SQL integrieren
                szValue = objTools.checknull(ThisDBCon.GetValueFromSQL(szSQL), 0) ' Wert holen
                If szMaxWert > 0 And szMaxWert <> "" Then               ' Wenn Maximal wert vorhanden
                    If CDbl(szValue) > CDbl(szMaxWert) Then szValue = szMaxWert ' Maximal wert prüfen
                End If
            Else                                                        ' Kein SQL Angegeben
                If szField <> "" Then                                   ' Alternativen wert (evtl durch eingabe)
                    If InStr(szField, ".") > 0 Then szField = Right(szField, Len(szField) - InStr(szField, "."))
                    szValue = objTools.checknull(rsBewerbungen.Fields(szField).Value, 0) ' Wert holen
                    If szMaxWert > 0 And szMaxWert <> "" Then           ' Wenn Maximal wert vorhanden
                        If CDbl(szValue) > CDbl(szMaxWert) Then szValue = szMaxWert ' Maximal wert prüfen
                    End If
                Else
                    szValue = 0                                         ' Kann nur 0 sein
                End If
            End If
            If szField <> "" Then                                       ' Wenn Speicherfeld angegeben
                rsBewerbungen.Fields(szField).Value = CDbl(szValue)     ' Wert zurückschreiben
            End If
            szSQL = ""                                                  ' SQL initialisieren -> Brauchen wir weiterhin
            szField = ""                                                ' Feld initialisieren -> Brauchen wir noch
            szMaxWert = ""                                              ' Maxwert init.  -> brauchen wir noch
            szFaktor = rsBerechnungen.Fields("FAKTOR016").Value         ' Faktor holen
            If szFaktor = "" Then szFaktor = 0                          ' Sicherstellen das Zahl
            
            szMaxWert = objTools.checknull(rsBerechnungen.Fields("MAXWERT016").Value, 0) ' Maximal wert holen für die Punkte
            szValueType = objTools.checknull(rsBerechnungen.Fields("VALUETYPE016").Value, "float")  ' Value TYpe ermitteln
            lngPunkte = CDbl(szFaktor) * CDbl(szValue)                  ' berechnen (Faktor mal wert = Punkte
            If szMaxWert > 0 Then
                If lngPunkte > CDbl(szMaxWert) Then lngPunkte = CDbl(szMaxWert) ' Prüfen ob Max wert überschritten
            End If
            Select Case UCase(szValueType)                              ' Value Typ prüfen
            Case "INT"
                    lngPunkte = Int(lngPunkte)                          ' Nachkommas abschneiden
            Case "FLOAT"
                                                                        ' NOOP
            Case Else
            End Select
            
            lngSumme = lngSumme + lngPunkte                             ' Summe Speichern
            szMaxWert = ""                                              ' Maxwert init.  -> brauchen wir noch
            szField = objTools.checknull(rsBerechnungen.Fields("PUNKTESAVEFIELD016").Value, "") ' Speicherfeld für Punkte holen
            If szField <> "" Then                                       ' Wenn Speicherfeld angegeben
                If InStr(szField, ".") > 0 Then szField = Right(szField, Len(szField) - InStr(szField, "."))
                If szDetails = "Summe" Then
' MW 11.08.11 {
'                    rsBewerbungen.Fields("Punkte").Value = lngSumme     ' Summe Speichern
                    rsBewerbungen.Fields("PUNKTESUM013").Value = lngSumme ' Summe Speichern
' MW 11.08.11 }
                Else
                    rsBewerbungen.Fields(szField).Value = lngPunkte     ' Punkte Speichern
                End If
                szField = ""                                            ' Feld für nächste runde initialisieren
            End If
            lngPunkte = 0                                               ' Punkte initialisieren
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
            rsBewerbungen.Update                                        ' Werte in Bewerbung speichern sonst haut nächste berechnung nicht hin
            rsBerechnungen.MoveNext                                     ' Nächste Berechnung
            Err.Clear                                                   ' Evtl. error clearen
On Error GoTo Errorhandler                                              ' Fehlerbehandlung wider aktivieren
        Wend                                                            ' Not rsBerechnungen.EOF
' MW 11.08.11 {
'        rsBewerbungen.Fields("Punkte").Value = lngSumme                 ' Summe Speichern
        rsBewerbungen.Fields("PUNKTESUM013").Value = lngSumme           ' Summe speichern
' MW 11.08.11 }
        rsBewerbungen.Update
        lngSumme = 0                                                    ' Summe initialisieren
                                                                        ' Rang eintragen + Zusage eintragen
        txtIDStelle.Text = rsBewerbungen.Fields("FK012013").Value
        lngAnzStellen = ThisDBCon.GetValueFromSQL( _
                "SELECT ANZ012 FROM STELLEN012 WHERE ID012 = '" & txtIDStelle.Text & "'")
        If txtIDStelle.Text = "" Then GoTo exithandler
        szSQL = "SELECT ID013, PunkteSUM013,Rang013, Zusage013 " & _
            " FROM STELLEN012 LEFT JOIN BEWERB013 ON ID012 = FK012013 " & _
            " WHERE ID012 = '" & txtIDStelle & "' ORDER BY PUNKTESUM013 DESC"
        Set rsMitbwerber = ThisDBCon.fillrs(szSQL, True)
        If rsMitbwerber Is Nothing Then GoTo exithandler
        If rsMitbwerber.RecordCount < 1 Then GoTo exithandler
        rsMitbwerber.MoveFirst                                          ' Zum 1. Mitbewerber springen
        For i = 1 To rsMitbwerber.RecordCount                           ' Alle Mitbewerber durchlaufen
            rsMitbwerber.Fields("RANG013").Value = i
            If i <= lngAnzStellen Then
                rsMitbwerber.Fields("Zusage013").Value = -1
            Else
                rsMitbwerber.Fields("Zusage013").Value = 0
            End If
            rsMitbwerber.Update
            rsMitbwerber.MoveNext
        Next i                                                          ' Nächster Mirbewerber
        rsBewerbungen.MoveNext
'    Wend ' Not rsBewerbungen.EOF

exithandler:
On Error Resume Next                                                    ' Hier keine fehler mehr
    Err.Clear                                                           ' Evtl Error claren
    Me.Refresh
    Exit Sub                                                            ' Function beenden
Errorhandler:
    Dim errNr As Long                                                   ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "PunkteBerechnungNeu", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Sub

Private Sub ShowKontextMenu(Menuename As String)
    ' Zeigt das Menü mit MenueName als Kontext (Popup) Menü an
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    Select Case Menuename
    Case "KontextMenueLVFortbildungen"
        PopupMenu kmnuLVFortbildungen
    Case "KontextMenueLVBewerbungen"
        PopupMenu kmnuLVBewerbungen
    Case "KontextMenueLVAktenort"
        PopupMenu kmnuLVAktenort
    Case "KontextMenueLVDokumente"
        PopupMenu kmnuLVDokumente
    Case "KontextMenueLVDisz"
        PopupMenu kmnuLVDisz
    Case "KontextMenueLVForderungen"
        PopupMenu kmnuLVForderungen
' MW 30.11.10 {
    Case "KontextmenueLVFristen"
        PopupMenu kmnuLVFristen
' MW 30.11.10 }
    Case Else
    
    End Select
    Err.Clear                                                           ' Evtl. error clearen
End Sub

Private Sub HandleMenueKlick(szMenueName As String, Optional szCaption As String)
    Dim BewID As String                                                 ' Bewerbungs ID
    Dim StellenID As String                                             ' Stellen ID
    Dim AusschrID As String                                             ' Ausschreibungs ID
    Dim ItemKey As String
    Dim szItemKeyArray() As String
    Dim szDocID As String                                               ' Dokument ID
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    If HandleLVkmnuNew(Me, szCaption) Then GoTo exithandler             ' bei neuem DS hier keine weitere Aktion
    Select Case szMenueName
    Case "kmnuLVFortbildungAdd"                                         ' Neue Fortbildung
    
    Case "kmnuLVFortbildungDel"                                         ' Fortbildung entfernen
'        Call DelRelationinLV(Me, "Fortbildungen", ThisDBcon, LVFortbildungen, rsFortbildungen, "ID014", "AFORT014")
'        Call RefereshFortbildungen(True)
        
    Case "kmnuLVBewerbungAdd"                                           ' Neue Berwerbung Hinzufügen
        Call EditDS("Bewerbung", ";;" & ID, True)
        Call RefereshBewerbungen(True)                                  ' LV Bewerber Aktualisieren
        Call InitFrameBewerberDatenNeu
        If Not bNew Then Call PunkteBerechnungNeu                       ' Evtl Punkte neu berechnen
        
    Case "kmnuLVBewerbungDel"                                           ' Bewerbung entfernen
        Call DelRelationinLV(Me, "Bewerbung", ThisDBCon, LVBewerbungen, rsBewerbungen, "ID013", "BEWERB013")
        Call RefereshBewerbungen(True)                                  ' LV Bewerber Aktualisieren
        Call InitFrameBewerberDatenNeu
        If Not bNew Then Call PunkteBerechnungNeu                       ' Evtl Punkte neu berechnen
        
    Case "kmnuLVAktenortNew"                                            ' Neuer Aktenort
        If rsBewerbungen.RecordCount = 1 Then                           ' Wenn mehere Bewerbungen Dann erste auswählen
            rsBewerbungen.MoveFirst
            BewID = rsBewerbungen.Fields("ID013").Value
        End If
        Call EditDS("Aktenort", ";" & BewID & ";" & ID, True)
        Call RefereshAktenort(True)                                     ' LV Aktenort aktualisieren
        
    Case "kmnuLVDokumentNew"                                            ' Neues Dokument
        Call GetIDCollection(Me, "", StellenID, AusschrID)              ' ID für Stelle und Auschreibung ermitteln
        Call WriteWord("", ID, StellenID, AusschrID)                    ' Word erstellung aufrufen
        Call RefereshDokumente(True)                                    ' LV Dokumente aktualisieren
        
    Case "kmnuLVDokumentOpen"                                           ' Dokument Öffnen
        Call HandleEditLVDoubleClick(Me, LVDokumente)
        
    Case "kmnuLVDokumentImport"                                         ' Dokument Importieren
        Call GetIDCollection(Me, "", StellenID, AusschrID)              ' ID für Stelle und Auschreibung ermitteln
        Call ImportWordDoc(ThisDBCon, ID, StellenID, AusschrID)         ' Dokument Importieren
        Call RefereshDokumente(True)                                    ' LV Dokumente aktualisieren
        
    Case "kmnuLVDokumentDel"                                            ' Dokument Löschen
        szDocID = GetRelLVSelectedID(LVDokumente)                       ' Doc ID Aus LV ermitteln
'On Error Resume Next
'        ItemKey = LVDokumente.SelectedItem.Key
'        Err.Clear
'        If ItemKey = "" Then GoTo exithandler
'On Error GoTo Errorhandler
'        szItemKeyArray = Split(LVDokumente.SelectedItem.Key, TV_KEY_SEP)
'        szDocID = szItemKeyArray(UBound(szItemKeyArray))
        If szDocID <> "" Then
            Call DeleteDS("Dokumente", szDocID)                         ' DS Löschen
            Call RefereshDokumente(True)                                ' LV Dokumente aktualisieren
        End If
        
    Case "kmnuLVDiszNew"                                                ' Neue Diziplinarmaßname
        Call EditDS("Disziplinarmaßnahmen", "" & ";" & ID, True)
        Call RefereshDizip(True)                                        ' LV Diziplinarmaßnamen aktualisieren
    
    Case "kmnuLVForderungNew"                                           ' Neue Foderung
        Call EditDS("Forderungen", "" & ";" & ID, True)
        Call RefereshFoderungen(True)
    Case "kmnuLVFristenNew"                                             ' Neue Frist
        Call GetIDCollection(Me, ID, StellenID, AusschrID)
        Call EditDS("Fristen", ";" & StellenID & ";" & ID, True)
        Call RefreshFristen(True)                                       ' LV Fristen Aktualisieren
    Case "kmnuLVFristenDel"                                             ' Frist Löschen
        Call DelRelationinLV(Me, "Fristen", ThisDBCon, LVFristen, rsFristen, "ID024", "FRIST024")
        Call RefreshFristen(True)                                       ' LV Fristen Aktualisieren
    Case Else
    
    End Select
exithandler:
On Error Resume Next                                                    ' Hier keine fehler mehr
    Err.Clear                                                           ' Evtl Error claren
    Exit Sub                                                            ' Function beenden
Errorhandler:
    Dim errNr As Long                                                   ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "HandleMenueKlick", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Sub

Private Sub HandleKeyDown(frmEdit As Form, KeyCode As Integer, Shift As Integer)
' Behandelt KeyDownEvents im Edit Form
On Error Resume Next
    Call HandleKeyDownEdit(Me, KeyCode, Shift)                          ' Spezielle KeyDown Events dieses Forms
    Call frmParent.HandleGlobalKeyCodes(KeyCode, Shift)                 ' Key Down Events der Anwendung
    Err.Clear
End Sub

Private Sub HandleTabClick(TS As TabStrip)
' Behandelt Tab Klicks
    Dim bNoBewerbung As Boolean
    Dim szMSG As String
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    If Not bInitTabMSG Then
        bInitTabMSG = True
        If Not HandleTabClickNew(Me, TS) Then                           ' Wenn bNew dan nur 1. Tab zulassen
            TS.Tabs(1).Selected = True
        End If
        bInitTabMSG = False
    End If
' MW 09.08.11 {
'    If TS.SelectedItem = "Info" Then                                    ' Info
    If TS.SelectedItem = TAB_NAME_INFO Then                             ' Info
' MW 09.08.11 }
        FrameBewerberDaten.Visible = False
        FrameDokumente.Visible = False
        FrameFortbildungen.Visible = False
        FramePersonenDaten.Visible = False
        FrameBewerbungen.Visible = False
        FrameDizip.Visible = False
        FrameAktenort.Visible = False
        FrameFristen.Visible = False                                    ' MW 30.11.10
        FrameInfo.Visible = True
        Exit Sub
    End If
    Select Case TS.SelectedItem
' MW 09.08.11 {
'    Case "Daten"                                                        ' Personen daten
    Case TAB_NAME_DATEN                                                 ' Personen daten
' MW 09.08.11 }
        FrameBewerberDaten.Visible = False
        FrameDokumente.Visible = False
        FrameFortbildungen.Visible = False
        FramePersonenDaten.Visible = True
        FrameInfo.Visible = False
        FrameBewerbungen.Visible = False
        FrameDizip.Visible = False
        FrameAktenort.Visible = False
        FrameForderungen.Visible = False
        FrameFristen.Visible = False                                    ' MW 30.11.10
' MW 09.08.11 {
'    Case "Dokumente"                                                    ' Dokumente
    Case TAB_NAME_DOC                                                   ' Dokumente
' MW 09.08.11 }
        FrameBewerberDaten.Visible = False
        FrameDokumente.Visible = True
        FrameFortbildungen.Visible = False
        FramePersonenDaten.Visible = False
        FrameBewerbungen.Visible = False
        FrameDizip.Visible = False
        FrameInfo.Visible = False
        FrameAktenort.Visible = False
        FrameForderungen.Visible = False
        FrameFristen.Visible = False                                    ' MW 30.11.10
' MW 09.08.11 {
'    Case "Bewerberdaten"                                                ' Bewerberdaten
    Case TAB_NAME_BEWDAT                                               ' Bewerberdaten"
' MW 09.08.11 }
'        If Not bNew Then Call PunkteBerechnungNeu
'        Call InitFrameBewerberDatenNeu
        If rsBewerbungen Is Nothing Then
            bNoBewerbung = True
        Else
            If rsBewerbungen.RecordCount = 0 Then bNoBewerbung = True
        End If
        If bNoBewerbung Then
            szMSG = "Sie müssen erst eine Bewerbung für diesen Bewerber eintragen bevor Sie die Bewerberdaten bearbeiten können."
            Call objError.ShowErrMsg(szMSG, vbInformation, "Hinweis", False, "", Me)  ' Meldung ausgeben
            TS.Tabs(4).Selected = True                                  ' 4. Tab Selecten (Bewerbungen)
            Exit Sub
        End If
        FrameBewerberDaten.Visible = True
        FrameDokumente.Visible = False
        FrameFortbildungen.Visible = False
        FramePersonenDaten.Visible = False
        FrameInfo.Visible = False
        FrameBewerbungen.Visible = False
        FrameDizip.Visible = False
        FrameAktenort.Visible = False
        FrameForderungen.Visible = False
        FrameFristen.Visible = False                                    ' MW 30.11.10
'        Call PunkteBerechnung
' MW 09.08.11 {
'    Case "Bewerbungen"                                                  ' Bewerbungen
    Case TAB_NAME_BEW                                                   ' Bewerbungen
' MW 09.08.11 }
        FrameBewerberDaten.Visible = False
        FrameDokumente.Visible = False
        FrameFortbildungen.Visible = False
        FramePersonenDaten.Visible = False
        FrameInfo.Visible = False
        FrameBewerbungen.Visible = True
        FrameDizip.Visible = False
        FrameAktenort.Visible = False
        FrameForderungen.Visible = False
        FrameFristen.Visible = False                                    ' MW 30.11.10
' MW 09.08.11 {
'    Case "Fortbildungen"                                                ' Fortbildungen
    Case TAB_NAME_FORT                                                  ' Fortbildungen
' MW 09.08.11 {
        FrameBewerberDaten.Visible = False
        FrameDokumente.Visible = False
        FrameFortbildungen.Visible = True
        FramePersonenDaten.Visible = False
        FrameInfo.Visible = False
        FrameBewerbungen.Visible = False
        FrameDizip.Visible = False
        FrameAktenort.Visible = False
        FrameForderungen.Visible = False
        FrameFristen.Visible = False                                    ' MW 30.11.10
' MW 09.08.11 {
'    Case "Disziplinarm."                                                ' Diziplinarmaßnamen
    Case TAB_NAME_DISZ                                                  ' Diziplinarmaßnamen
' MW 09.08.11 }
        FrameBewerberDaten.Visible = False
        FrameDokumente.Visible = False
        FrameFortbildungen.Visible = False
        FramePersonenDaten.Visible = False
        FrameBewerbungen.Visible = False
        FrameDizip.Visible = True
        FrameAktenort.Visible = False
        FrameInfo.Visible = False
        FrameForderungen.Visible = False
        FrameFristen.Visible = False                                    ' MW 30.11.10
' MW 09.08.11 {
'    Case "Ford-Verz."                                                   ' Forderungen
    Case TAB_NAME_FORD                                                  ' Forderungen
' MW 09.08.11 }
        FrameInfo.Visible = False
         FrameBewerberDaten.Visible = False
         FrameDokumente.Visible = False
        FrameFortbildungen.Visible = False
        FramePersonenDaten.Visible = False
        FrameBewerbungen.Visible = False
        FrameDizip.Visible = False
        FrameInfo.Visible = False
        FrameAktenort.Visible = False
        FrameForderungen.Visible = True
        FrameFristen.Visible = False                                    ' MW 30.11.10
' MW 09.08.11 {
'    Case "Aktenort"                                                     ' Aktenstandort
    Case TAB_NAME_AKTE                                                  ' Aktenstandort
' MW 09.08.11 }
        FramePersonenDaten.Visible = False
        FrameDokumente.Visible = False
        FrameInfo.Visible = False
        FrameBewerberDaten.Visible = False
        FrameBewerbungen.Visible = False
        FrameFortbildungen.Visible = False
        FrameDizip.Visible = False
        FrameAktenort.Visible = True
        FrameInfo.Visible = False
        FrameForderungen.Visible = False
' MW 30.11.10 {
        FrameFristen.Visible = False
' MW 09.08.11 {
'    Case "Fristen"
    Case TAB_NAME_FRIST                                                 ' Fristen
' MW 09.08.11 }
        FramePersonenDaten.Visible = False
        FrameDokumente.Visible = False
        FrameInfo.Visible = False
        FrameBewerberDaten.Visible = False
        FrameBewerbungen.Visible = False
        FrameFortbildungen.Visible = False
        FrameDizip.Visible = False
        FrameAktenort.Visible = False
        FrameInfo.Visible = False
        FrameForderungen.Visible = False
        FrameFristen.Visible = True
' MW 30.11.10 }
    Case Else
    
    End Select
exithandler:
On Error Resume Next                                                    ' Hier keine fehler mehr
    Me.Refresh
    Err.Clear                                                           ' Evtl Error claren
    Exit Sub                                                            ' Function beenden
Errorhandler:
    Dim errNr As Long                                                   ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "HandleTabClick", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Sub

Private Function ValidateEditForm() As Boolean
    Dim szMSG As String                                                 ' MessageText
    Dim szTitle As String                                               ' Message Titel
    Dim FocusCTL As Control                                             ' Control das den Focus erhält
    Dim bValidationFaild As Boolean                                     ' Validation nicht erfolgreich
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    szTitle = "Unvollständige Daten"                                    ' Meldungstitel setzen
    If cmbStatus.Text = "Notar" Then                                    ' Nur bei Notaren
        bValidationFaild = ValidateTxtFieldOnEmpty(txtAZ, "Aktenzeichen", _
                szMSG, FocusCTL)                                        ' Aktenzeichen auf Leer prüfen
    End If
    If txtAmtsOrt.Text = "" And txtKanzleiOrt.Text <> "" Then
        szMSG = "Sie haben keinen Amtssitz eingetragen. Soll der Kanzlei Ort auch der Amtsitz sein?"
        If objError.ShowErrMsg(szMSG, vbQuestion + vbOKCancel, "Fehlende Eingabe") = vbOK Then
            txtAmtsOrt.Text = txtKanzleiOrt.Text
            Adodc1.Recordset.Fields(txtAmtsOrt.DataField).Value = txtAmtsOrt.Text
            txtAmtsPLZ.Text = txtkanzeliPLZ.Text
            Adodc1.Recordset.Fields(txtAmtsPLZ.DataField).Value = txtAmtsPLZ.Text
        End If
    End If
    bValidationFaild = ValidateTxtFieldOnEmpty(cmbStatus, "Status", _
            szMSG, FocusCTL)                                            ' Status auf Leer prüfen
    bValidationFaild = ValidateTxtFieldOnEmpty(cmbAnrede, "Anrede", _
            szMSG, FocusCTL)                                            ' Anrede auf Leer prüfen
    bValidationFaild = ValidateTxtFieldOnEmpty(txtNachname, "Nachname", _
            szMSG, FocusCTL)                                            ' Nachname auf Leer prüfen
    If bValidationFaild Then                                            ' Wenn Validierung Gescheitert
        Call objError.ShowErrMsg(szMSG, vbInformation, szTitle)         ' Hinweis meldung anzeigen
        FocusCTL.SetFocus                                               ' Fokus setzen
        ValidateEditForm = False                                        ' Rückgabewert setzen (Kein Erfolg)
    Else                                                                ' Sonst
        ValidateEditForm = True                                         ' Rückgabewert setzen (Erfolg)
    End If
exithandler:
On Error Resume Next                                                    ' Hier keine fehler mehr
    Exit Function                                                       ' Function beenden
Errorhandler:
    Dim errNr As Long                                                   ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "ValidateEditForm", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Function

Private Function SaveEditForm() As Boolean
' Speichert den Datensatz nach Validierung der eingaben
    Dim bNewBeforSave As Boolean                                        ' DS vom speichern Neu
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    bNewBeforSave = bNew                                                ' DS war vom speichern neu
    If Not ValidateEditForm Then GoTo exithandler                       ' Eingaben Validieren
    If UpdateEditForm(Me, szRootkey) Then                               ' Speichern
        bNew = False                                                    ' nicht mehr neu
        Call HiglightThisMustFields(True)                               ' Hervorhebung abschalten
        If Not bNew Then Call PunkteBerechnungNeu                       ' Evtl Punkte neu berechnen
    End If
    SaveEditForm = True                                                 ' Erfolg zurück liefern
exithandler:
On Error Resume Next                                                    ' Hier keine fehler mehr
    Exit Function                                                       ' Function beenden
Errorhandler:
    Dim errNr As Long                                                   ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "SaveEditForm", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Function

Private Function RefreshEditForm()
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    Call Me.Adodc1.Refresh                                              ' AdoDC aktualisieren
    Me.Refresh                                                          ' Form aktualisieren
    Call RefereshAktenort(Me.FrameAktenort.Visible)                     ' Frame aktenort Refreshen
    Call RefereshBewerbungen(Me.FrameBewerbungen.Visible)               ' Frame Bewerbungen Refreshen
    Call RefereshDizip(Me.FrameDizip.Visible)                           ' Frame Diziplinarmaßnahmen refreshen
    Call RefereshDokumente(Me.FrameDokumente.Visible)                   ' Frame Dokumente Refreschen
    Call RefereshFoderungen(Me.FrameForderungen.Visible)                ' Frame Forderungen refreshen
    Call RefreshFristen(Me.FrameFristen.Visible)                        ' Frame Fristen aktualisieren

exithandler:
On Error Resume Next                                                    ' Hier keine fehler mehr
    Exit Function                                                       ' Function beenden
Errorhandler:
    Dim errNr As Long                                                   ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "RefreshEditForm", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Function
                                                                        ' *****************************************
                                                                        ' TabSrip Events
Private Sub TabStripPerson_Click()
    Call HandleTabClick(TabStripPerson)                                 ' Tab Klick behandeln
End Sub
                                                                        ' *****************************************
                                                                        ' Button Events
Private Sub cmdESC_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    If SaveEditForm Then                                                ' Dieses Form Speichern
        Call CheckUpdate(Me)                                            ' Evtl Übernehmen disablen
        Unload Me                                                       ' Form Schliessen
    End If
exithandler:
On Error Resume Next                                                    ' Hier keine fehler mehr
    Me.MousePointer = vbDefault                                         ' Mauszeiger wieder normal
    Err.Clear                                                           ' Evtl Error claren
    Exit Sub                                                            ' Function beenden
Errorhandler:
    Dim errNr As Long                                                   ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "cmdOK_Click", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Sub

Public Sub cmdUpdate_Click()
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    If SaveEditForm Then                                                ' Dieses Form Speichern
        Call CheckUpdate(Me)                                            ' Evtl Übernehmen disablen
    End If
exithandler:
On Error Resume Next                                                    ' Hier keine fehler mehr
    Me.MousePointer = vbDefault                                         ' Mauszeiger wieder normal
    Err.Clear                                                           ' Evtl Error claren
    Exit Sub                                                            ' Function beenden
Errorhandler:
    Dim errNr As Long                                                   ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "cmdUpdate_Click", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Sub

Private Sub cmdDelete_Click()

On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    Call DeleteDS(szRootkey, ID)                                        ' diese Person Löschen
    Unload Me                                                           ' Dieses form schliessen
exithandler:
On Error Resume Next                                                    ' Hier keine fehler mehr
    Err.Clear                                                           ' Evtl Error claren
    Exit Sub                                                            ' Function beenden
Errorhandler:
    Dim errNr As Long                                                   ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "cmdDelete_Click", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Sub

Private Sub cmdSave_Click()

On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    If SaveEditForm Then                                                ' Dieses Form Speichern
        Call CheckUpdate(Me)                                            ' Evtl Übernehmen disablen
    End If
    
exithandler:
On Error Resume Next                                                    ' Hier keine fehler mehr
    Me.MousePointer = vbDefault                                         ' Mauszeiger wieder normal
    Err.Clear                                                           ' Evtl Error claren
    Exit Sub                                                            ' Function beenden
Errorhandler:
    Dim errNr As Long                                                   ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "cmdSave_Click", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Sub

Private Sub cmdWord_Click()

    Dim StellenID As String                                             ' Stellen ID
On Error Resume Next                                                    ' Erstmal ohne Fehlerbehandlung
    rsBewerbungen.MoveFirst
    StellenID = rsBewerbungen.Fields("FK012013").Value                  ' evtl. Stellen ID Ermitteln
    Err.Clear
On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren
    Call WriteWord("", ID, StellenID)                                   ' SAT aufrufen
    Call RefereshDokumente(True)                                        ' LV Dokumente Aktualisieren
exithandler:
On Error Resume Next                                                    ' Hier keine fehler mehr
    Err.Clear                                                           ' Evtl Error claren
    Exit Sub                                                            ' Function beenden
Errorhandler:
    Dim errNr As Long                                                   ' Fehlernummer
    Dim errDesc As String                                               ' Fehler beschreibung
    errNr = Err.Number                                                  ' Fehlernummer auslesen
    errDesc = Err.Description                                           ' Fehler beschreibung auslesen
    Err.Clear                                                           ' Fehler Clearen
On Error Resume Next                                                    ' Keinen Fehler in der Fehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "cmdWord_Click", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                                  ' Weiter mit Exithandler
End Sub

Private Sub txtCreateFrom_Click()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If txtModifyFrom.Tag <> "" Then                                     ' Tag vorhanden
        Call AskUserAboutThisDS(txtCreateFrom, "Wegen " _
                & txtNachname & ", " & txtVorname)                      ' Email an User vorbereiten
    End If
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub txtModifyFrom_Click()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If txtModifyFrom.Tag <> "" Then                                     ' Tag vorhanden
        Call AskUserAboutThisDS(txtModifyFrom, "Wegen " _
                & txtNachname & ", " & txtVorname)                      ' Email an User vorbereiten
    End If
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub
                                                                        ' *****************************************
                                                                        ' Mouse Events
Private Sub txtModifyFrom_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If txtModifyFrom.Tag <> "" Then                                     ' Tag vorhanden
        Call MousePointerLink(Me, txtModifyFrom)                        ' Mouspointer Hyperlink setzen
    End If
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub txtCreateFrom_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If txtCreateFrom.Tag <> "" Then                                     ' Tag vorhanden
        Call MousePointerLink(Me, txtCreateFrom)                        ' Mouspointer Hyperlink setzen
    End If
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub LVBewerbungen_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then Call ShowKontextMenu("KontextMenueLVBewerbungen") ' Bei Rechtsklik Kontextmenü anzeigen
End Sub

Private Sub LVFortbildungen_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then Call ShowKontextMenu("KontextMenueLVFortbildungen") ' Bei Rechtsklik Kontextmenü anzeigen
End Sub

Private Sub LVAktenort_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then Call ShowKontextMenu("KontextMenueLVAktenort")   ' Bei Rechtsklik Kontextmenü anzeigen
End Sub

Private Sub LVDokumente_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then Call ShowKontextMenu("KontextMenueLVDokumente")  ' Bei Rechtsklik Kontextmenü anzeigen
End Sub

Private Sub LVDizip_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then Call ShowKontextMenu("KontextMenueLVDisz")       ' Bei Rechtsklik Kontextmenü anzeigen
End Sub

Private Sub LVForderungen_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
     If Button = 2 Then Call ShowKontextMenu("KontextMenueLVForderungen") ' Bei Rechtsklik Kontextmenü anzeigen
End Sub
' MW 30.11.10 {
Private Sub LVFristen_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
     If Button = 2 Then Call ShowKontextMenu("KontextmenueLVFristen")   ' Bei Rechtsklik Kontextmenü anzeigen
End Sub
' MW 30.11.10 }
                                                                        ' *****************************************
                                                                        ' Menue Events
Private Sub kmnuLVAktenortNew_Click()
    Call HandleMenueKlick("kmnuLVAktenortNew")                          ' KontextMenüKlick im LV Aktenort behandeln
End Sub

Private Sub kmnuLVDokumentNew_Click()
    Call HandleMenueKlick("kmnuLVDokumentNew")                          ' KontextMenüKlick im LV Dokumente behandeln
End Sub

Private Sub kmnuLVDokumentOpen_Click()
    Call HandleMenueKlick("kmnuLVDokumentOpen")                         ' KontextMenüKlick im LV Dokumente behandeln
End Sub

Private Sub kmnuLVDokumentImport_Click()
    Call HandleMenueKlick("kmnuLVDokumentImport")                       ' KontextMenüKlick im LV Dokumente behandeln
End Sub

Private Sub kmnuLVDokumentDel_Click()
     Call HandleMenueKlick("kmnuLVDokumentDel")                         ' KontextMenüKlick im LV Dokumente behandeln
End Sub

Private Sub kmnuLVDiszNew_Click()
    Call HandleMenueKlick("kmnuLVDiszNew")                              ' KontextMenüKlick im LV Disziplinarmaßnamen behandeln
End Sub

Private Sub kmnuLVBewerbungAdd_Click()
    Call HandleMenueKlick("kmnuLVBewerbungAdd")                         ' KontextMenüKlick im LV Bewerbungen behandeln
End Sub

Private Sub kmnuLVBewerbungDel_Click()
    Call HandleMenueKlick("kmnuLVBewerbungDel")                         ' KontextMenüKlick im LV Bewerbungen behandeln
End Sub

Private Sub kmnuLVForderungNew_Click()
    Call HandleMenueKlick("kmnuLVForderungNew")                         ' KontextMenüKlick im LV Forderungen behandeln
End Sub

Private Sub kmnuLVFortbildungAdd_Click()
    Call HandleMenueKlick("kmnuLVFortbildungAdd")
End Sub

Private Sub kmnuLVFortbildungDel_Click()
    Call HandleMenueKlick("kmnuLVFortbildungDel")
End Sub
' MW 30.11.10 {
Private Sub kmnuLVFristDel_Click()
    Call HandleMenueKlick("kmnuLVFristenDel")
End Sub

Private Sub kmnuLVFristenNew_Click()
    Call HandleMenueKlick("kmnuLVFristenNew")
End Sub
' MW 30.11.10 }
                                                                        ' *****************************************
                                                                        ' Change Events
Private Sub cmbStelle_Click()

On Error GoTo Errorhandler                                              ' Fehlerbehandlung aktivieren

If cmbStelle.DataChanged Then
    If rsBewerbungen.RecordCount > 0 Then
        Dim Valuearray() As String
        Valuearray = Split(cmbStelle, "  ")
        rsBewerbungen.MoveFirst
        While Not rsBewerbungen.EOF
            If rsBewerbungen.Fields("bezirk").Value = Valuearray(1) _
                And rsBewerbungen.Fields("BewerbungsFrist").Value = Valuearray(0) Then
                    txtIDStelle.Text = rsBewerbungen.Fields("FK012013").Value
                    txtIDBew.Text = rsBewerbungen.Fields("ID013").Value
                    bInit = True
                    Call InitFrameBewerberDatenNeu
                    bInit = False
                    Call PunkteBerechnungNeu
                    Exit Sub
            Else
                rsBewerbungen.MoveNext
            End If
        Wend
    End If
End If
Errorhandler:
    'Stop
bInit = False
End Sub

Private Sub cmbStelle_Validate(Cancel As Boolean)
'On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
'    Call InitFrameBewerberDaten
'    Call PunkteBerechnung
'    Call InitFrameBewerberDatenNeu
    'Call PunkteBerechnungNeu
End Sub

Private Sub chkVerstorben_Click()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                              ' Bei initialisierung nicht weitermachen
    ' Wenn Status Notar dann fragen ob statusänderung
    'If UCase(cmbStatus.Text) = UCase("Notar") Then
        cmbStatus.Text = "ausgeschieden"                                ' Status auf ausgeschieden setzen
        If txtAusgeschieden = "" Then txtAusgeschieden.Text = Now()     ' Ausgeschieden datum (heute) setzen
        bDirty = True                                                   ' jetzt ist der DS ungespeichert
        Call CheckUpdate(Me)                                            ' Evtl. Buttons Dis/enablen
    'End If
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub txtValue_Change(Index As Integer)
    Dim szSQL As String                                                 ' SQL Statement
    Dim szValue As String
    Dim lngSelstart As Integer
    Dim bUpdateOK As Boolean
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If Not bInit And txtValue(Index).Locked = False Then
        lngSelstart = txtValue(Index).SelStart
        ' Bewerberdaten Speichern
        If txtValue(Index).Text = "-0" Then Exit Sub
        If txtValue(Index).Text = "" Then txtValue(Index).Text = "0"
        If Right(txtValue(Index).Text, 1) = "," Or _
                Right(txtValue(Index).Text, 1) = "." Then
            'txtValue(Index).Text = txtValue(Index).Text & "0"
            Exit Sub
        End If
    
        
        If txtValue(Index).Tag <> "" And txtValue(Index).Text <> "" And txtIDBew.Text <> "" Then
            If IsNumeric(txtValue(Index)) Then
                szValue = Replace(txtValue(Index).Text, ",", ".")
'On Error Resume Next
'                rsBewerbungen.Fields(txtValue(Index).Tag).Value = szValue
'                rsBewerbungen.Update
                
                szSQL = "UPDATE BEWERB013 SET " & txtValue(Index).Tag & " = " & szValue & _
                        " WHERE ID013 = '" & txtIDBew & "'"
                If ThisDBCon.execSql(szSQL) Then
'                If Err.Number <> 0 Then bUpdateOK = True
'                Err.Clear
'On Error GoTo errorhandler
'                If bUpdateOK Then
                    'Me.Refresh
                    Call PunkteBerechnungNeu
                    bInit = True
                    Call RefereshBewerbungen(FrameBewerbungen.Visible)
                    Call InitFrameBewerberDatenNeu
                    bInit = False
                    'txtValue(Index).SelStart = Len(txtValue(Index).Text)
                    txtValue(Index).SelStart = lngSelstart
                End If
               
            Else
                
            End If
        End If
    End If
Errorhandler:

    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub cmbAg_DropDown()                                            ' Liste ausklappen
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                              ' Wenn Form Initialisiert wird -> fertig
    OldCmbValue = cmbAg.Text                                            ' Auswahl beginnt
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub cmbAg_Click()                                               ' Änderung duch Liste auswahl
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                              ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> cmbAg Then bDirty = True                          ' Dirty Nur wenn combo <> oldValue
    Call CheckUpdate(Me)                                                ' Evtl. Buttons Dis/enablen
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub cmbAg_Change()                                              ' Änderung duch Texteingabe
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                              ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> cmbAg Then bDirty = True                          ' Dirty Nur wenn combo <> oldValue
    Call CheckUpdate(Me)                                                ' Evtl. Buttons Dis/enablen
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub cmbLG_DropDown()                                            ' Liste ausklappen
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                              ' Wenn Form Initialisiert wird -> fertig
    OldCmbValue = cmbLG.Text                                            ' Auswahl beginnt
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub cmbLG_Click()                                               ' Änderung duch Liste auswahl
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                              ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> cmbLG Then bDirty = True                          ' Dirty Nur wenn combo <> oldValue
    Call CheckUpdate(Me)                                                ' Evtl. Buttons Dis/enablen
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub cmbLG_Change()                                              ' Änderung duch Texteingabe
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                              ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> cmbLG Then bDirty = True                          ' Dirty Nur wenn combo <> oldValue
    Call CheckUpdate(Me)                                                ' Evtl. Buttons Dis/enablen
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub cmbStatus_DropDown()                                        ' Liste ausklappen
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                              ' Wenn Form Initialisiert wird -> fertig
    OldCmbValue = cmbStatus.Text                                        ' Auswahl beginnt
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub cmbStatus_Click()                                           ' Änderung duch Liste auswahl
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                              ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> cmbStatus Then bDirty = True                      ' Dirty Nur wenn combo <> oldValue
    Call CheckUpdate(Me)                                                ' Evtl. Buttons Dis/enablen
    Call ChangeStatus                                                   ' Status änderung -> sonder behandlung
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub cmbStatus_Change()                                          ' Änderung duch Texteingabe
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                              ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> cmbStatus Then bDirty = True                      ' Dirty Nur wenn combo <> oldValue
    Call CheckUpdate(Me)                                                ' Evtl. Buttons Dis/enablen
    Call ChangeStatus                                                   ' Status änderung -> sonder behandlung
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub cmbAnrede_DropDown()                                        ' Liste ausklappen
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                              ' Wenn Form Initialisiert wird -> fertig
    OldCmbValue = cmbAnrede.Text                                        ' Auswahl beginnt
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub cmbAnrede_Click()                                           ' Änderung duch Liste auswahl
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                              ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> cmbAnrede Then bDirty = True                      ' Dirty Nur wenn combo <> oldValue
    Call CheckUpdate(Me)                                                ' Evtl. Buttons Dis/enablen
    Call FirstCharUp(Me, Me.cmbAnrede)                                  ' 1. Zeichen Groß schreiben
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub cmbAnrede_Change()                                          ' Änderung duch Texteingabe
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                              ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> cmbAnrede Then bDirty = True                      ' Dirty Nur wenn combo <> oldValue
    Call CheckUpdate(Me)                                                ' Evtl. Buttons Dis/enablen
    Call FirstCharUp(Me, Me.cmbAnrede)                                  ' 1. Zeichen Groß schreiben
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub cmbTitel_DropDown()                                         ' Liste ausklappen
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                              ' Wenn Form Initialisiert wird -> fertig
    OldCmbValue = cmbTitel.Text                                         ' Auswahl beginnt
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub cmbTitel_Click()                                            ' Änderung duch Liste auswahl
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                              ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> cmbTitel Then bDirty = True                       ' Dirty Nur wenn combo <> oldValue
    Call CheckUpdate(Me)                                                ' Evtl. Buttons Dis/enablen
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub cmbTitel_Change()                                           ' Änderung duch Texteingabe
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                              ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> cmbTitel Then bDirty = True                       ' Dirty Nur wenn combo <> oldValue
    Call CheckUpdate(Me)                                                ' Evtl. Buttons Dis/enablen
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub cmbTitel_Validate(Cancel As Boolean)
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If bInit Then Exit Sub                                              ' Wenn Form Initialisiert wird -> fertig
    If OldCmbValue <> cmbTitel Then bDirty = True                       ' Dirty Nur wenn combo <> oldValue
    Call CheckUpdate(Me)                                                ' Evtl. Buttons Dis/enablen
    OldCmbValue = ""                                                    ' Auswahl beendet
    Call FirstCharUp(Me, Me.cmbTitel)                                   ' 1. Zeichen Groß schreiben
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub txtAusgeschieden_Change()
    Dim szMSG As String                                                 ' Meldungs text
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If IsDate(txtAusgeschieden.Text) Then
        DTAusgeschieden.Value = txtAusgeschieden.Text
        If bInit Then Exit Sub                                          ' Bei initialisierung nicht weitermachen
        bDirty = True                                                   ' jetzt ist der DS ungespeichert
        ' Wenn Status Notar dann fragen ob statusänderung
        If UCase(cmbStatus.Text) = UCase("Notar") Then
            szMSG = "Sie haben das Datum 'Ausgeschieden am' geändert. Möchen Sie den Status dieses Notars auf 'ausgeschieden' verändern?"
            If objError.ShowErrMsg(szMSG, vbQuestion + vbOKCancel, "Statusänderung") = vbOK Then
                cmbStatus.Text = "ausgeschieden"
                bDirty = True                                           ' jetzt ist der DS ungespeichert
                Call CheckUpdate(Me)                                    ' Evtl. Buttons Dis/enablen
            Else
                Me.txtAusgeschieden.Text = Null
            End If
        End If
    End If
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub DTAusgeschieden_Change()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    txtAusgeschieden.Text = DTAusgeschieden.Value
    If bInit Then Exit Sub                                              ' Bei initialisierung nicht weitermachen
    Me.Adodc1.Recordset.Fields(txtAusgeschieden.DataField).Value _
            = Format(Me.DTAusgeschieden.Value, "dd.mm.yyyy")            ' Bei DTPicker Wert nochmal übernehmen da Sonst die Datenbindung nicht funzt
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub txtBestelt_Change()
    Dim szMSG As String                                                 ' Meldungstext
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If Len(txtBestelt.Text) < 8 Then Exit Sub                           ' Text länge < 8 kein datum (oder unvollständig)
    If IsDate(txtBestelt.Text) Then                                     ' Ist txt Feld ein datum?
        DTBestellt.Value = txtBestelt.Text                              ' Datum in DT übernehmen
        If bInit Then Exit Sub                                          ' Bei initialisierung nicht weitermachen
        bDirty = True                                                   ' jetzt ist der DS ungespeichert
        Call CheckUpdate(Me)                                            ' Evtl. Buttons Dis/enablen
        If UCase(cmbStatus.Text) = UCase("Bewerber") Then               ' Wenn Status Bewerber dann fragen ob statusänderung
            szMSG = "Sie haben das Datum 'Bestellt am' geändert. Möchen Sie den Status dieses " & _
                    "Bewerbers zum 'Notar' verändern?"                  ' Meldungstext festlegen
            If objError.ShowErrMsg(szMSG, vbQuestion + vbOKCancel, "Statusänderung") = vbOK Then    ' Wenn Meldung ok
                cmbStatus.Text = "Notar"                                ' Statusänderm
                bDirty = True                                           ' jetzt ist der DS ungespeichert
                Call CheckUpdate(Me)                                    ' Evtl. Buttons Dis/enablen
            Else
                Me.txtBestelt.Text = Null                               ' Wenn Meldung nein -> eingabe zurücksetzen
            End If
        End If
    End If
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub DTBestellt_Change()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    txtBestelt.Text = DTBestellt.Value                                  ' DT DAtum in TXT Feld übernehmen
    If bInit Then Exit Sub                                              ' Bei initialisierung nicht weitermachen
    Me.Adodc1.Recordset.Fields(txtBestelt.DataField).Value _
            = Format(Me.DTBestellt.Value, "dd.mm.yyyy")                 ' Bei DTPicker Wert nochmal übernehmen da Sonst die Datenbindung nicht funzt
    Err.Clear                                                           ' Evtl. error clearen
End Sub

Private Sub txtGeb_Change()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    If Len(txtGeb.Text) < 8 Then Exit Sub                               ' Text länge < 8 kein datum (oder unvollständig)
    If IsDate(txtGeb.Text) Then                                         ' Ist Textfeld Datum ?
        DTGeb.Value = txtGeb.Text                                       ' Datum in DT sezten
        If bInit Then Exit Sub                                          ' Bei initialisierung nicht weitermachen
        bDirty = True                                                   ' jetzt ist der DS ungespeichert
        Call CheckUpdate(Me)                                            ' Evtl. Buttons Dis/enablen
    End If
    Err.Clear                                                           ' Evtl. error clearen
End Sub

Private Sub DTGeb_Change()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
    txtGeb.Text = DTGeb.Value                                           ' Datum aus DT übernehmen
    If bInit Then Exit Sub                                              ' Bei initialisierung nicht weitermachen
    Me.Adodc1.Recordset.Fields(txtGeb.DataField).Value _
            = Format(Me.DTGeb.Value, "dd.mm.yyyy")                      ' Bei DTPicker Wert nochmal übernehmen da Sonst die Datenbindung nicht funzt
    Err.Clear                                                           ' Evtl. error clearen
End Sub


Private Sub txtAnwaltSeit_Change()
On Error Resume Next                                                    ' Fehlerbehandlung aktivieren
    If Len(txtAnwaltSeit.Text) < 8 Then Exit Sub
    If IsDate(txtAnwaltSeit.Text) Then
        DTAnwaltSeit.Value = txtAnwaltSeit.Text
        If bInit Then Exit Sub                                          ' Bei initialisierung nicht weitermachen
        bDirty = True                                                   ' jetzt ist der DS ungespeichert
        Call CheckUpdate(Me)                                            ' Evtl. Buttons Dis/enablen
        Call PunkteBerechnungNeu                                        ' Für die Punkte berechnung relevant
    End If
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub DTAnwaltSeit_Change()
On Error Resume Next                                                    ' Fehlerbehandlung aktivieren
    txtAnwaltSeit.Text = DTAnwaltSeit.Value
    If bInit Then Exit Sub                                              ' Bei initialisierung nicht weitermachen
    Me.Adodc1.Recordset.Fields(txtAnwaltSeit.DataField).Value _
            = Format(Me.DTAnwaltSeit.Value, "dd.mm.yyyy")               ' Bei DTPicker Wert nochmal übernehmen da Sonst die Datenbindung nicht funzt
    Err.Clear                                                           ' Evtl. Error Clearen
End Sub

Private Sub txtExaErgebnis_Change()
    If Not bInit Then                                                   ' Bei initialisierung nicht weitermachen
        Call StandartTextChange(Me, txtExaErgebnis)
'        Call PunkteBerechnungNeu                                        ' Für die Punkte berechnung relevant
    End If
End Sub

Private Sub txtVorname_Change()
    If Not bInit Then Call StandartTextChange(Me, txtVorname)
    Call FirstCharUp(Me, Me.txtVorname)                                 ' 1. Zeichen Groß schreiben
End Sub

Private Sub txtNachname_Change()
    If Not bInit Then Call StandartTextChange(Me, txtNachname)
    Call FirstCharUp(Me, Me.txtNachname)                                ' 1. Zeichen Groß schreiben
End Sub

Private Sub txtAmtsOrt_Change()
    If Not bInit Then Call StandartTextChange(Me, txtAmtsOrt)
    Call FirstCharUp(Me, Me.txtAmtsOrt)                                 ' 1. Zeichen Groß schreiben
End Sub

Private Sub txtAmtsPLZ_Change()
    If Not bInit Then Call StandartTextChange(Me, txtAmtsPLZ)
End Sub

Private Sub txtAZ_Change()
    If Not bInit Then Call StandartTextChange(Me, txtAZ)
End Sub

Private Sub txtBem_Change()
    If Not bInit Then Call StandartTextChange(Me, txtBem)
End Sub

Private Sub txtkanzeliPLZ_Change()
    If Not bInit Then Call StandartTextChange(Me, txtkanzeliPLZ)
End Sub

Private Sub txtKanzeliTel_Change()
    If Not bInit Then Call StandartTextChange(Me, txtKanzeliTel)
End Sub

Private Sub txtKanzleiFax_Change()
    If Not bInit Then Call StandartTextChange(Me, txtKanzleiFax)
End Sub

Private Sub txtKanzleiOrt_Change()
    If Not bInit Then Call StandartTextChange(Me, txtKanzleiOrt)
    Call FirstCharUp(Me, Me.txtKanzleiOrt)                              ' 1. Zeichen Groß schreiben
End Sub

Private Sub txtkanzleiStr_Change()
    If Not bInit Then Call StandartTextChange(Me, txtkanzleiStr)
    Call FirstCharUp(Me, Me.txtkanzleiStr)                              ' 1. Zeichen Groß schreiben
End Sub

Private Sub txtNamensZusatz_Change()
    If Not bInit Then Call StandartTextChange(Me, txtNamensZusatz)
End Sub
                                                                        ' *****************************************
                                                                        ' Focus Events
Private Sub cmbStelle_GotFocus()
    Call HiglightCurentField(Me, cmbStelle, False)
End Sub

Private Sub cmbStelle_LostFocus()
    Call HiglightCurentField(Me, cmbStelle, True)
End Sub

Private Sub txtValue_GotFocus(Index As Integer)
    Call HiglightCurentField(Me, txtValue(Index), False)
End Sub

Private Sub txtValue_LostFocus(Index As Integer)
    Call HiglightCurentField(Me, txtValue(Index), True)
End Sub

Private Sub cmbAg_GotFocus()
    Call HiglightCurentField(Me, cmbAg, False)
End Sub

Private Sub cmbAg_LostFocus()
    Call HiglightCurentField(Me, cmbAg, True)
End Sub

Private Sub cmbAnrede_GotFocus()
    Call HiglightCurentField(Me, cmbAnrede, False)
End Sub

Private Sub cmbAnrede_LostFocus()
    Call HiglightCurentField(Me, cmbAnrede, True)
End Sub

Private Sub cmbLG_GotFocus()
    Call HiglightCurentField(Me, cmbLG, False)
End Sub

Private Sub cmbLG_LostFocus()
    Call HiglightCurentField(Me, cmbLG, True)
End Sub

Private Sub cmbStatus_GotFocus()
    Call HiglightCurentField(Me, cmbStatus, False)
End Sub

Private Sub cmbStatus_LostFocus()
    Call HiglightCurentField(Me, cmbStatus, True)
End Sub

Private Sub cmbTitel_GotFocus()
    Call HiglightCurentField(Me, cmbTitel, False)
End Sub

Private Sub cmbTitel_LostFocus()
    Call HiglightCurentField(Me, cmbTitel, True)
End Sub

Private Sub txtkanzeliPLZ_GotFocus()
    Call HiglightCurentField(Me, txtkanzeliPLZ, False)
End Sub

Private Sub txtkanzeliPLZ_LostFocus()
    Call HiglightCurentField(Me, txtkanzeliPLZ, True)
End Sub

Private Sub txtKanzeliTel_GotFocus()
    Call HiglightCurentField(Me, txtKanzeliTel, False)
End Sub

Private Sub txtKanzeliTel_LostFocus()
    Call HiglightCurentField(Me, txtKanzeliTel, True)
End Sub

Private Sub txtKanzleiFax_GotFocus()
    Call HiglightCurentField(Me, txtKanzleiFax, False)
End Sub

Private Sub txtKanzleiFax_LostFocus()
    Call HiglightCurentField(Me, txtKanzleiFax, True)
End Sub

Private Sub txtKanzleiOrt_GotFocus()
    Call HiglightCurentField(Me, txtKanzleiOrt, False)
End Sub

Private Sub txtKanzleiOrt_LostFocus()
    Call HiglightCurentField(Me, txtKanzleiOrt, True)
End Sub

Private Sub txtkanzleiStr_GotFocus()
    Call HiglightCurentField(Me, txtkanzleiStr, False)
End Sub

Private Sub txtkanzleiStr_LostFocus()
    Call HiglightCurentField(Me, txtkanzleiStr, True)
End Sub

Private Sub txtNachname_GotFocus()
    Call HiglightCurentField(Me, txtNachname, False)
End Sub

Private Sub txtNachname_LostFocus()
    Call HiglightCurentField(Me, txtNachname, True)
End Sub

Private Sub txtNamensZusatz_GotFocus()
    Call HiglightCurentField(Me, txtNamensZusatz, False)
End Sub

Private Sub txtNamensZusatz_LostFocus()
    Call HiglightCurentField(Me, txtNamensZusatz, True)
End Sub

Private Sub txtVorname_GotFocus()
    Call HiglightCurentField(Me, txtVorname, False)
End Sub

Private Sub txtVorname_LostFocus()
    Call HiglightCurentField(Me, txtVorname, True)
End Sub

Private Sub txtAmtsOrt_GotFocus()
    Call HiglightCurentField(Me, txtAmtsOrt, False)
End Sub

Private Sub txtAmtsOrt_LostFocus()
    Call HiglightCurentField(Me, txtAmtsOrt, True)
End Sub

Private Sub txtAmtsPLZ_GotFocus()
    Call HiglightCurentField(Me, txtAmtsPLZ, False)
End Sub

Private Sub txtAmtsPLZ_LostFocus()
    Call HiglightCurentField(Me, txtAmtsPLZ, True)
End Sub

Private Sub txtAnwaltSeit_GotFocus()
    Call HiglightCurentField(Me, txtAnwaltSeit, False)
End Sub

Private Sub txtAnwaltSeit_LostFocus()
    Call HiglightCurentField(Me, txtAnwaltSeit, True)
End Sub

Private Sub txtAusgeschieden_GotFocus()
    Call HiglightCurentField(Me, txtAusgeschieden, False)
End Sub

Private Sub txtAusgeschieden_LostFocus()
    Call HiglightCurentField(Me, txtAusgeschieden, True)
End Sub

Private Sub txtAZ_GotFocus()
    Call HiglightCurentField(Me, txtAZ, False)
End Sub

Private Sub txtAZ_LostFocus()
    Call HiglightCurentField(Me, txtAZ, True)
End Sub

Private Sub txtBem_GotFocus()
    Call HiglightCurentField(Me, txtBem, False)
End Sub

Private Sub txtBem_LostFocus()
    Call HiglightCurentField(Me, txtBem, True)
End Sub

Private Sub txtBestelt_GotFocus()
    Call HiglightCurentField(Me, txtBestelt, False)
End Sub

Private Sub txtBestelt_LostFocus()
    Call HiglightCurentField(Me, txtBestelt, True)
End Sub

Private Sub txtExaErgebnis_GotFocus()
    Call HiglightCurentField(Me, txtExaErgebnis, False)
End Sub

Private Sub txtExaErgebnis_LostFocus()
    Call HiglightCurentField(Me, txtExaErgebnis, True)
End Sub

Private Sub txtGeb_GotFocus()
    Call HiglightCurentField(Me, txtGeb, False)
End Sub

Private Sub txtGeb_LostFocus()
    Call HiglightCurentField(Me, txtGeb, True)
End Sub
                                                                        ' *****************************************
                                                                        ' Key Events
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub LVBewerbungen_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub LVDizip_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub LVDokumente_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub LVForderungen_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub LVFortbildungen_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub cmbAnrede_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub cmbStelle_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub cmbTitel_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub DTAnwaltSeit_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub DTGeb_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtBem_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtCreate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtCreateFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub txtExaErgebnis_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyDown(Me, KeyCode, Shift)
End Sub
                                                                        ' *****************************************
                                                                        ' List View Events

Private Sub LVAktenort_GotFocus()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
'    Call RefereshAktenort(True)
    Err.Clear                                                           ' Evtl. Error clearen
End Sub

Private Sub LVBewerbungen_GotFocus()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
'    Call RefereshBewerbungen(True)
    Err.Clear                                                           ' Evtl. Error clearen
End Sub

Private Sub LVDizip_GotFocus()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
'    Call RefereshDizip(True)
    Err.Clear                                                           ' Evtl. Error clearen
End Sub

Private Sub LVDokumente_GotFocus()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
'    Call RefereshDokumente(True)
    Err.Clear                                                           ' Evtl. Error clearen
End Sub

Private Sub LVForderungen_GotFocus()
On Error Resume Next                                                    ' Fehlerbehandlung deaktivieren
'    Call RefereshFoderungen(True)
    Err.Clear                                                           ' Evtl. Error clearen
End Sub

Private Sub LVDokumente_DblClick()
    Call HandleEditLVDoubleClick(Me, Me.LVDokumente)
End Sub

Private Sub LVAktenort_DblClick()
    Call HandleEditLVDoubleClick(Me, Me.LVAktenort, frmParent)
    Call RefereshAktenort(True)
End Sub

Private Sub LVDizip_DblClick()
    Call HandleEditLVDoubleClick(Me, Me.LVDizip, frmParent)
    Call RefereshDizip(True)
End Sub

Private Sub LVBewerbungen_DblClick()
    Call HandleEditLVDoubleClick(Me, Me.LVBewerbungen, frmParent)
    Call RefereshBewerbungen(True)
End Sub

Private Sub LVFortbildungen_DblClick()
    Call HandleEditLVDoubleClick(Me, Me.LVFortbildungen, frmParent)
    Call RefereshFortbildungen(True)
End Sub

Private Sub LVForderungen_DblClick()
    Call HandleEditLVDoubleClick(Me, Me.LVForderungen, frmParent)
    Call RefereshFoderungen(True)
End Sub

' MW 30.11.10 {
Private Sub LVFristen_DblClick()
    Call HandleEditLVDoubleClick(Me, Me.LVFristen, frmParent)
    Call RefreshFristen(True)
End Sub

Private Sub LVFristen_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SetColumnOrder(LVFristen, ColumnHeader)
End Sub
' MW 30.11.10 }

Private Sub LVForderungen_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SetColumnOrder(LVForderungen, ColumnHeader)
End Sub

Private Sub LVAktenort_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SetColumnOrder(LVAktenort, ColumnHeader)
End Sub

Private Sub LVBewerbungen_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SetColumnOrder(LVBewerbungen, ColumnHeader)
End Sub

Private Sub LVDizip_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SetColumnOrder(LVDizip, ColumnHeader)
End Sub

Private Sub LVDokumente_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SetColumnOrder(LVDokumente, ColumnHeader)
End Sub

Private Sub LVFortbildungen_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SetColumnOrder(LVFortbildungen, ColumnHeader)
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

Public Property Get GetCurrentStep() As String
    GetCurrentStep = CurrentStep
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

