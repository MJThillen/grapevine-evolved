VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWerewolfSheet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Werewolf Character"
   ClientHeight    =   6165
   ClientLeft      =   210
   ClientTop       =   405
   ClientWidth     =   9060
   Icon            =   "frmWerewolfSheet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   9060
   Tag             =   "C"
   Begin VB.CommandButton cmdRename 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1020
      TabIndex        =   134
      Top             =   150
      Width           =   975
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   0
      Left            =   2160
      TabIndex        =   12
      Top             =   1200
      Width           =   6615
      Begin VB.TextBox txtUserField 
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   28
         Top             =   3840
         Width           =   2175
      End
      Begin VB.CheckBox chkNPC 
         Alignment       =   1  'Right Justify
         Caption         =   "NPC"
         Height          =   375
         Left            =   3735
         TabIndex        =   37
         Top             =   2400
         Width           =   660
      End
      Begin VB.TextBox txtUserField 
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   32
         Top             =   960
         Width           =   2175
      End
      Begin MSComCtl2.UpDown updTemper 
         Height          =   315
         Index           =   1
         Left            =   5775
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   3375
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         Value           =   3
         Max             =   1000
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updTemper 
         Height          =   315
         Index           =   2
         Left            =   5775
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3855
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         Max             =   1000
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updTemper 
         Height          =   315
         Index           =   0
         Left            =   5775
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   2895
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         Value           =   3
         Max             =   1000
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTemperFloat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " F "
         Height          =   195
         Left            =   5520
         TabIndex        =   144
         Top             =   3000
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lbShowXPLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XP Unspent / Earned"
         Height          =   375
         Left            =   3240
         TabIndex        =   35
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label lblShowXP 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4200
         TabIndex        =   36
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   14
         Tag             =   "Tribe"
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tribe"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   510
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   960
         TabIndex        =   26
         Tag             =   "Position"
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   25
         Top             =   3390
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   960
         TabIndex        =   20
         Tag             =   "Archetypes"
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nature"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   19
         Top             =   1950
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   960
         TabIndex        =   22
         Tag             =   "Archetypes"
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Demeanor"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   21
         Top             =   2430
         Width           =   855
      End
      Begin VB.Label lblTemper 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   45
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lblTemperLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Willpower"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   44
         Top             =   3870
         Width           =   855
      End
      Begin VB.Label lblTemper 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   42
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label lblTemperLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gnosis"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   41
         Top             =   3390
         Width           =   855
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   255
         Index           =   8
         Left            =   3480
         TabIndex        =   33
         Top             =   1470
         Width           =   615
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Player"
         Height          =   255
         Index           =   7
         Left            =   3480
         TabIndex        =   29
         Top             =   510
         Width           =   615
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   8
         Left            =   4200
         TabIndex        =   34
         Tag             =   "Status, Character"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   4200
         TabIndex        =   30
         Tag             =   "?PL"
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Auspice"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   15
         Top             =   990
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   16
         Tag             =   "Auspice"
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Breed"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   17
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   18
         Tag             =   "Breed"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblTemperLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rage"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   38
         Top             =   2910
         Width           =   855
      End
      Begin VB.Label lblTemper 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   39
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label lblUserField 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pack"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   27
         Top             =   3885
         Width           =   855
      End
      Begin VB.Label lblUserField 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Character ID"
         Height          =   255
         Index           =   2
         Left            =   3195
         TabIndex        =   31
         Top             =   1005
         Width           =   900
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rank"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   23
         Top             =   2910
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   960
         TabIndex        =   24
         Tag             =   "Rank"
         Top             =   2880
         Width           =   2175
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4785
      Index           =   6
      Left            =   2160
      TabIndex        =   115
      Top             =   1200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdEstimate 
         Height          =   315
         Left            =   2535
         Picture         =   "frmWerewolfSheet.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   1470
         Width           =   585
      End
      Begin VB.TextBox txtExperience 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   121
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtExperience 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   118
         Top             =   495
         Width           =   1575
      End
      Begin VB.TextBox txtUserField 
         Height          =   375
         Index           =   3
         Left            =   4200
         TabIndex        =   128
         Top             =   960
         Width           =   2175
      End
      Begin MSComctlLib.ListView lvwHistory 
         Height          =   2430
         Left            =   105
         TabIndex        =   132
         Tag             =   "?XP"
         Top             =   2145
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   4286
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Date"
            Text            =   "Date"
            Object.Width           =   1720
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Change"
            Text            =   "Change"
            Object.Width           =   1905
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "Reason"
            Text            =   "Reason"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "Unspent"
            Text            =   "Unspent"
            Object.Width           =   873
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "Earned"
            Text            =   "Earned"
            Object.Width           =   873
         EndProperty
      End
      Begin MSComCtl2.UpDown updExperience 
         Height          =   315
         Index           =   1
         Left            =   2535
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   990
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         OrigLeft        =   3480
         OrigTop         =   840
         OrigRight       =   3915
         OrigBottom      =   1125
         Max             =   999
         Min             =   -999
         Orientation     =   1
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updExperience 
         Height          =   315
         Index           =   0
         Left            =   2535
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   525
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         OrigLeft        =   3480
         OrigTop         =   840
         OrigRight       =   3915
         OrigBottom      =   1125
         Max             =   999
         Min             =   -999
         Orientation     =   1
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin VB.Label lblXPLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Point Counting Aid..."
         Height          =   255
         Index           =   4
         Left            =   -120
         TabIndex        =   123
         Top             =   1530
         Width           =   2535
      End
      Begin VB.Label lblXPLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Experience &History"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   131
         Top             =   1920
         Width           =   6375
      End
      Begin VB.Label lblXPLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Unspent"
         Height          =   375
         Index           =   0
         Left            =   -120
         TabIndex        =   117
         Top             =   540
         Width           =   975
      End
      Begin VB.Label lblXPLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Earned"
         Height          =   375
         Index           =   1
         Left            =   -120
         TabIndex        =   120
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label lblXPLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Experience"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   116
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblUserField 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Date"
         Height          =   375
         Index           =   3
         Left            =   3360
         TabIndex        =   127
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   9
         Left            =   4200
         TabIndex        =   126
         Tag             =   "?NR"
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Assigned Narrator"
         Height          =   375
         Index           =   9
         Left            =   3360
         TabIndex        =   125
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblModified 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4200
         TabIndex        =   130
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblModifiedLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Modified"
         Height          =   375
         Left            =   3360
         TabIndex        =   129
         Top             =   1440
         Width           =   735
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   4
      Left            =   2160
      TabIndex        =   80
      Top             =   1200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtPerm 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   98
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox txtTemp 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   2
         Left            =   1747
         TabIndex        =   99
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox txtPerm 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   94
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txtTemp 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   1
         Left            =   1747
         TabIndex        =   95
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txtTemp 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   0
         Left            =   1747
         TabIndex        =   91
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtPerm 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   90
         Top             =   2880
         Width           =   735
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   9
         ItemData        =   "frmWerewolfSheet.frx":0B14
         Left            =   120
         List            =   "frmWerewolfSheet.frx":0B16
         TabIndex        =   82
         Tag             =   "Honor"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   10
         ItemData        =   "frmWerewolfSheet.frx":0B18
         Left            =   2280
         List            =   "frmWerewolfSheet.frx":0B1A
         TabIndex        =   84
         Tag             =   "Glory"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   11
         ItemData        =   "frmWerewolfSheet.frx":0B1C
         Left            =   4440
         List            =   "frmWerewolfSheet.frx":0B1E
         TabIndex        =   86
         Tag             =   "Wisdom"
         Top             =   480
         Width           =   2055
      End
      Begin MSComCtl2.UpDown updNumber 
         Height          =   315
         Index           =   0
         Left            =   5895
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   3855
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         Value           =   6
         OrigLeft        =   2535
         OrigTop         =   1335
         OrigRight       =   3120
         OrigBottom      =   1650
         Max             =   1000
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updPermTemp 
         Height          =   315
         Index           =   0
         Left            =   2535
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   2895
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         OrigLeft        =   6015
         OrigTop         =   1815
         OrigRight       =   6600
         OrigBottom      =   2130
         Max             =   1000
         Min             =   -1000
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updPermTemp 
         Height          =   315
         Index           =   1
         Left            =   2535
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   3375
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         OrigLeft        =   6015
         OrigTop         =   1815
         OrigRight       =   6600
         OrigBottom      =   2130
         Max             =   1000
         Min             =   -1000
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updPermTemp 
         Height          =   315
         Index           =   2
         Left            =   2535
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   3855
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         OrigLeft        =   6015
         OrigTop         =   1815
         OrigRight       =   6600
         OrigBottom      =   2130
         Max             =   1000
         Min             =   -1000
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin VB.Label lblPermTemp 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Wisdom"
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   97
         Top             =   3870
         Width           =   855
      End
      Begin VB.Label lblPermTemp 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Glory"
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   93
         Top             =   3390
         Width           =   855
      End
      Begin VB.Label lblTemp 
         Alignment       =   2  'Center
         Caption         =   "Temp."
         Height          =   255
         Left            =   1740
         TabIndex        =   88
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lblPerm 
         Alignment       =   2  'Center
         Caption         =   "Perm."
         Height          =   255
         Left            =   960
         TabIndex        =   87
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Camp"
         Height          =   255
         Index           =   11
         Left            =   3360
         TabIndex        =   138
         Top             =   3390
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   11
         Left            =   4320
         TabIndex        =   137
         Tag             =   "Camps"
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   10
         Left            =   4320
         TabIndex        =   102
         Tag             =   "Totems"
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Totem"
         Height          =   255
         Index           =   10
         Left            =   3360
         TabIndex        =   101
         Top             =   2910
         Width           =   855
      End
      Begin VB.Label lblPermTemp 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Honor"
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   89
         Top             =   2910
         Width           =   855
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Honor"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   81
         Tag             =   "Honor"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Glory"
         Height          =   255
         Index           =   10
         Left            =   2280
         TabIndex        =   83
         Tag             =   "Glory"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Wisdom"
         Height          =   255
         Index           =   11
         Left            =   4440
         TabIndex        =   85
         Tag             =   "Wisdom"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblNumberLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Spirit Notoriety"
         Height          =   495
         Index           =   0
         Left            =   3240
         TabIndex        =   103
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label lblNumber 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   4320
         TabIndex        =   104
         Top             =   3840
         Width           =   1575
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   5
      Left            =   2160
      TabIndex        =   106
      Top             =   1200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtMemo 
         Height          =   1035
         Index           =   1
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   110
         Top             =   480
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   1035
         Index           =   19
         ItemData        =   "frmWerewolfSheet.frx":0B20
         Left            =   120
         List            =   "frmWerewolfSheet.frx":0B22
         TabIndex        =   108
         Tag             =   "?LO"
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtMemo 
         Height          =   2400
         Index           =   0
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   114
         Top             =   2160
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   2400
         Index           =   18
         ItemData        =   "frmWerewolfSheet.frx":0B24
         Left            =   120
         List            =   "frmWerewolfSheet.frx":0B26
         TabIndex        =   112
         Tag             =   "?EQ"
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Biography"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   109
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Favorite Locations"
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   107
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   113
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Equipment / Fetishes"
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   111
         Top             =   1920
         Width           =   3135
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   3
      Left            =   2160
      TabIndex        =   71
      Top             =   1200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.ListBox lstTraits 
         Height          =   1035
         Index           =   17
         ItemData        =   "frmWerewolfSheet.frx":0B28
         Left            =   120
         List            =   "frmWerewolfSheet.frx":0B2A
         TabIndex        =   73
         Tag             =   "Features"
         Top             =   480
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   1035
         Index           =   12
         ItemData        =   "frmWerewolfSheet.frx":0B2C
         Left            =   3360
         List            =   "frmWerewolfSheet.frx":0B2E
         TabIndex        =   75
         Tag             =   "Health Levels"
         Top             =   480
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   2400
         Index           =   16
         ItemData        =   "frmWerewolfSheet.frx":0B30
         Left            =   3360
         List            =   "frmWerewolfSheet.frx":0B32
         TabIndex        =   79
         Tag             =   "Rites"
         Top             =   2160
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   2400
         Index           =   15
         ItemData        =   "frmWerewolfSheet.frx":0B34
         Left            =   120
         List            =   "frmWerewolfSheet.frx":0B36
         TabIndex        =   77
         Tag             =   "Gifts"
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Other Features"
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   72
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rites"
         Height          =   255
         Index           =   16
         Left            =   3360
         TabIndex        =   78
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gifts"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   76
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Health Levels"
         Height          =   255
         Index           =   12
         Left            =   3360
         TabIndex        =   74
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   2
      Left            =   2160
      TabIndex        =   60
      Top             =   1200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   6
         ItemData        =   "frmWerewolfSheet.frx":0B38
         Left            =   120
         List            =   "frmWerewolfSheet.frx":0B3A
         TabIndex        =   62
         Tag             =   "Abilities"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   7
         ItemData        =   "frmWerewolfSheet.frx":0B3C
         Left            =   4440
         List            =   "frmWerewolfSheet.frx":0B3E
         TabIndex        =   66
         Tag             =   "Influences"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   13
         ItemData        =   "frmWerewolfSheet.frx":0B40
         Left            =   120
         List            =   "frmWerewolfSheet.frx":0B42
         TabIndex        =   68
         Tag             =   "Merits"
         Top             =   3120
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   14
         ItemData        =   "frmWerewolfSheet.frx":0B44
         Left            =   3360
         List            =   "frmWerewolfSheet.frx":0B46
         TabIndex        =   70
         Tag             =   "Flaws"
         Top             =   3120
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   8
         ItemData        =   "frmWerewolfSheet.frx":0B48
         Left            =   2280
         List            =   "frmWerewolfSheet.frx":0B4A
         TabIndex        =   64
         Tag             =   "Backgrounds"
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Abilities"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   61
         Tag             =   "Abilities|Ability"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Influences"
         Height          =   255
         Index           =   7
         Left            =   4440
         TabIndex        =   65
         Tag             =   "Influences|Influence"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Merits"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   67
         Top             =   2880
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Flaws"
         Height          =   255
         Index           =   14
         Left            =   3360
         TabIndex        =   69
         Top             =   2880
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Backgrounds"
         Height          =   255
         Index           =   8
         Left            =   2280
         TabIndex        =   63
         Tag             =   "Backgrounds|Background"
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   1
      Left            =   2160
      TabIndex        =   47
      Top             =   1200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdTraitMax 
         Caption         =   "  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   140
         Top             =   2520
         Width           =   255
      End
      Begin VB.CommandButton cmdTraitMax 
         Caption         =   "+"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   139
         Top             =   2520
         Width           =   255
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   0
         ItemData        =   "frmWerewolfSheet.frx":0B4C
         Left            =   120
         List            =   "frmWerewolfSheet.frx":0B4E
         TabIndex        =   49
         Tag             =   "Physical"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   1
         ItemData        =   "frmWerewolfSheet.frx":0B50
         Left            =   2280
         List            =   "frmWerewolfSheet.frx":0B52
         TabIndex        =   51
         Tag             =   "Social"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   2
         ItemData        =   "frmWerewolfSheet.frx":0B54
         Left            =   4440
         List            =   "frmWerewolfSheet.frx":0B56
         TabIndex        =   53
         Tag             =   "Mental"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   3
         ItemData        =   "frmWerewolfSheet.frx":0B58
         Left            =   120
         List            =   "frmWerewolfSheet.frx":0B5A
         TabIndex        =   55
         Tag             =   "Physical, Negative"
         Top             =   3120
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   4
         ItemData        =   "frmWerewolfSheet.frx":0B5C
         Left            =   2280
         List            =   "frmWerewolfSheet.frx":0B5E
         TabIndex        =   57
         Tag             =   "Social, Negative"
         Top             =   3120
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   5
         ItemData        =   "frmWerewolfSheet.frx":0B60
         Left            =   4440
         List            =   "frmWerewolfSheet.frx":0B62
         TabIndex        =   59
         Tag             =   "Mental, Negative"
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CommandButton cmdIncrement 
         Caption         =   "+"
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdDecrement 
         Caption         =   "  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdDescend 
         Caption         =   "â"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdAscend 
         Caption         =   "á"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   2520
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblTraitMax 
         Alignment       =   2  'Center
         Caption         =   "Max 11"
         Height          =   255
         Index           =   2
         Left            =   5040
         TabIndex        =   143
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblTraitMax 
         Alignment       =   2  'Center
         Caption         =   "Max 11"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   142
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblTraitMax 
         Alignment       =   2  'Center
         Caption         =   "Max 11"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   141
         Top             =   2535
         Width           =   855
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Physical"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Tag             =   "Physical"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Social"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   50
         Tag             =   "Social"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Mental"
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   52
         Tag             =   "Mental"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Negative Physical"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   54
         Tag             =   "Negative Physical"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Negative Social"
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   56
         Tag             =   "Negative Social"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Negative Mental"
         Height          =   255
         Index           =   5
         Left            =   4440
         TabIndex        =   58
         Tag             =   "Negative Mental"
         Top             =   2880
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdESCClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   136
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdCustom 
      Caption         =   "Custom &Entry"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdNote 
      Caption         =   "Add &Note to Entry"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<- Re&move"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add ->"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   1695
   End
   Begin VB.ListBox lstMenu 
      Height          =   2010
      ItemData        =   "frmWerewolfSheet.frx":0B64
      Left            =   120
      List            =   "frmWerewolfSheet.frx":0B66
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtUserField 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   135
      Top             =   120
      Width           =   6855
   End
   Begin MSComctlLib.TabStrip tabTabStrip 
      Height          =   5175
      Left            =   2040
      TabIndex        =   0
      Top             =   840
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9128
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identity"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "P S M"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Abl Bkg Inf Mer Fl"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "OF HL Gifts Rites"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Renown"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bio Eqp Note"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  XP"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   480
      Picture         =   "frmWerewolfSheet.frx":0B68
      Top             =   185
      Width           =   480
   End
   Begin VB.Label lblMenuItem 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblUserField 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   133
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblMenuTitle 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   915
      Width           =   1695
   End
End
Attribute VB_Name = "frmWerewolfSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Constants by which specific labels are indexed. (fi = Field Index)
Private Const fiTribe = 0
Private Const fiAuspice = 1
Private Const fiBreed = 2
Private Const fiNature = 3
Private Const fiDemeanor = 4
Private Const fiRank = 5
Private Const fiPosition = 6
Private Const fiPlayer = 7
Private Const fiStatus = 8
Private Const fiNarrator = 9
Private Const fiTotem = 10
Private Const fiCamp = 11

'Constants by which specific number labels are indexed (ni = Number Index)
Private Const niNotoriety = 0

'Constants by which specific temper labels are indexed (pi = Point Index)
Private Const piRage = 0
Private Const piGnosis = 1
Private Const piWillpower = 2

' Constants by which specific list boxes are indexed. (ti = Trait Index)
Private Const tiPhysical = 0
Private Const tiSocial = 1
Private Const tiMental = 2
Private Const tiPhysicalNeg = 3
Private Const tiSocialNeg = 4
Private Const tiMentalNeg = 5
Private Const tiAbilities = 6
Private Const tiInfluences = 7
Private Const tiBackgrounds = 8
Private Const tiHonor = 9
Private Const tiGlory = 10
Private Const tiWisdom = 11
Private Const tiHealth = 12
Private Const tiMerits = 13
Private Const tiFlaws = 14
Private Const tiGifts = 15
Private Const tiRites = 16
Private Const tiFeatures = 17
Private Const tiEquipment = 18
Private Const tiLocations = 19

' Constants by which specific text boxes are indexed. (xi = Text Index)
Private Const xiName = 0
Private Const xiPack = 1
Private Const xiID = 2
Private Const xiStartDate = 3

' Constants by which specific perm/temp fields are indexed. (di = Decimal Index)
Private Const diHonor = 0
Private Const diGlory = 1
Private Const diWisdom = 2

' Constants by which specific memo fields are indexed. (mi = Memo Index)
Private Const miNotes = 0
Private Const miBiography = 1

' Constant by which to reference the index of the XP frame and text boxes
Private Const xpFrame = 6
Private Const xpUnspent = 0
Private Const xpEarned = 1

Private Werewolf As WerewolfClass     'The Werewolf character
Private CharSheetEngine As CharacterSheetEngineClass    'Handles common functions
Private Populating As Boolean                           'defuses some events when characters are loaded

Public Sub ShowWerewolf(Character As WerewolfClass)
'
' Name:         ShowWerewolf
' Description:  Show and initialize a new instance of the form.
' Arguments:    The WerewolfClass whose data the form modifies.
'

    Dim DataState As Boolean

    Populating = True

    Set Werewolf = Nothing
    Set Werewolf = Character
    DataState = Game.DataChanged

    txtUserField(xiName) = Werewolf.Name
    Me.Caption = Werewolf.Name
    
    lblField(fiTribe) = Werewolf.Tribe
    lblField(fiAuspice) = Werewolf.Auspice
    lblField(fiBreed) = Werewolf.Breed
    lblField(fiPosition) = Werewolf.Position
    lblField(fiNature) = Werewolf.Nature
    lblField(fiDemeanor) = Werewolf.Demeanor
    lblField(fiRank) = Werewolf.Rank
    lblField(fiTotem) = Werewolf.Totem
    lblField(fiCamp) = Werewolf.Camp
    lblField(fiStatus) = Werewolf.Status
    lblField(fiPlayer) = Werewolf.Player
    lblField(fiNarrator) = Werewolf.Narrator
        
    lblTraitMax(tiPhysical) = "Max " & CStr(Werewolf.PhysicalMax)
    lblTraitMax(tiSocial) = "Max " & CStr(Werewolf.SocialMax)
    lblTraitMax(tiMental) = "Max " & CStr(Werewolf.MentalMax)
    
    txtUserField(xiID) = Werewolf.ID
    txtUserField(xiPack) = Werewolf.Pack
    txtUserField(xiStartDate) = Format(Werewolf.StartDate, "mmmm d, yyyy")
    
    txtMemo(miBiography) = Werewolf.Biography
    txtMemo(miNotes) = Werewolf.Notes

    Call ChangeTemper(piRage, 0)
    Call ChangeTemper(piGnosis, 0)
    Call ChangeTemper(piWillpower, 0)
    
    CharSheetEngine.RefreshTraitList lstTraits(tiPhysical), Werewolf.PhysicalList
    CharSheetEngine.RefreshTraitList lstTraits(tiSocial), Werewolf.SocialList
    CharSheetEngine.RefreshTraitList lstTraits(tiMental), Werewolf.MentalList
    CharSheetEngine.RefreshTraitList lstTraits(tiPhysicalNeg), Werewolf.PhysicalNegList
    CharSheetEngine.RefreshTraitList lstTraits(tiSocialNeg), Werewolf.SocialNegList
    CharSheetEngine.RefreshTraitList lstTraits(tiMentalNeg), Werewolf.MentalNegList
    CharSheetEngine.RefreshTraitList lstTraits(tiAbilities), Werewolf.AbilityList
    CharSheetEngine.RefreshTraitList lstTraits(tiInfluences), Werewolf.InfluenceList
    CharSheetEngine.RefreshTraitList lstTraits(tiBackgrounds), Werewolf.BackgroundList
    CharSheetEngine.RefreshTraitList lstTraits(tiHonor), Werewolf.HonorList
    CharSheetEngine.RefreshTraitList lstTraits(tiGlory), Werewolf.GloryList
    CharSheetEngine.RefreshTraitList lstTraits(tiWisdom), Werewolf.WisdomList
    CharSheetEngine.RefreshTraitList lstTraits(tiHealth), Werewolf.HealthList
    
    CharSheetEngine.RefreshTraitList lstTraits(tiMerits), Werewolf.MeritList
    CharSheetEngine.RefreshTraitList lstTraits(tiFlaws), Werewolf.FlawList
    CharSheetEngine.RefreshTraitList lstTraits(tiGifts), Werewolf.GiftList
    CharSheetEngine.RefreshTraitList lstTraits(tiRites), Werewolf.RiteList
    CharSheetEngine.RefreshTraitList lstTraits(tiFeatures), Werewolf.FeatureList
    CharSheetEngine.RefreshTraitList lstTraits(tiEquipment), Werewolf.EquipmentList
    CharSheetEngine.RefreshTraitList lstTraits(tiLocations), Werewolf.HangoutList
    
    updNumber(niNotoriety).Value = Werewolf.Notoriety
    
    txtPerm(diHonor).Text = CStr(Werewolf.Honor)
    txtTemp(diHonor).Text = CStr(Werewolf.TempHonor)
    txtPerm(diGlory).Text = CStr(Werewolf.Glory)
    txtTemp(diGlory).Text = CStr(Werewolf.TempGlory)
    txtPerm(diWisdom).Text = CStr(Werewolf.Wisdom)
    txtTemp(diWisdom).Text = CStr(Werewolf.TempWisdom)
    
    lblModified.Caption = Format(Werewolf.LastModified, "mmmm d, yyyy")
    chkNPC.Value = IIf(Werewolf.IsNPC, vbChecked, vbUnchecked)
    
    Me.Show
    
    Game.DataChanged = DataState
    Populating = False

End Sub

Public Sub ShowXP()
'
' Name:         ShowXP
' Description:  Jump to the XP tab.
'

    Set tabTabStrip.SelectedItem = tabTabStrip.Tabs(xpFrame + 1)
    Call tabTabStrip_Click

End Sub

Public Sub SetDefaultOutput()
'
' Name:         SetDefaultOutput
' Description:  Initilize the OutputEngineClass with default output settings.
'
    With OutputEngine
        .Template = IIf(fraTab(xpFrame).Visible, tnXPHistory, tnCharacterSheets)
        .SelectSet(osCharacters).Clear
        .SelectSet(osCharacters).Add Werewolf.Name
        .GameDate = 0
    End With
    
End Sub

Private Sub chkNPC_Click()
'
' Name:         chkNPC_Click
' Description:  Toggle whether the character is an NPC.
'

    If Not Populating Then
        Werewolf.IsNPC = (chkNPC.Value = vbChecked)
        mdiMain.AnnounceChanges Me, atCharacters
        SetDataChanged
    End If

End Sub

Private Sub cmdAdd_Click()
'
' Name:         cmdAdd_Click
' Description:  If a selection is active, have the CharSheetEngine add to
'               the menu, OR add a new XP history entry.
'

    If lstMenu.ListIndex <> -1 Then
        
        If CharSheetEngine.TargetType = ttXPHistory Then
            If CharSheetEngine.AddXPEntry(lvwHistory, Werewolf.Experience) Then
                RefreshXP
                SetDataChanged
                lvwHistory.SetFocus
            End If
        Else
            CharSheetEngine.AddSelected
            SetDataChanged
        End If
    
    End If
    
End Sub

Private Sub cmdCustom_Click()
'
' Name:         cmdCustom_Click
' Description:  Have the CharSheetEngine add a custom entry to the target., OR
'               clear the XP history.

    If CharSheetEngine.TargetType = ttXPHistory Then
        If MsgBox("Are you sure you want to TOTALLY clear this history?", vbYesNo, _
                "Clear History") = vbYes Then
            Werewolf.Experience.Clear
            SetDataChanged
            RefreshXP
        End If
    Else
        CharSheetEngine.AddCustom
        SetDataChanged
    End If
    
End Sub

Private Sub cmdAscend_Click()
'
' Name:         cmdAscend_Click
' Description:  Move the selected entry down.
'

    If cmdAscend.Visible Then
        CharSheetEngine.MoveEntryUp
        CharSheetEngine.TargetList.SetFocus
        SetDataChanged
    End If
    
End Sub

Private Sub cmdDescend_Click()
'
' Name:         cmdDescend_Click
' Description:  Move the selected entry down.
'

    If cmdDescend.Visible Then
        CharSheetEngine.MoveEntryDown
        CharSheetEngine.TargetList.SetFocus
        SetDataChanged
    End If
    
End Sub

Private Sub cmdDecrement_Click()
'
' Name:         cmdDecrement_Click
' Description:  Decrement the selected entry.
'

    If cmdDecrement.Visible Then
        CharSheetEngine.DecrementEntry
        CharSheetEngine.TargetList.SetFocus
        SetDataChanged
    End If
    
End Sub

Private Sub cmdEstimate_Click()
'
' Name:         cmdEstimate_Click
' Description:  Show the point counting tool.
'
    frmCalculator.ShowCalculator Werewolf

End Sub

Private Sub cmdIncrement_Click()
'
' Name:         cmdIncrement_Click
' Description:  Increment the selected entry.
'

    If cmdIncrement.Visible Then
        CharSheetEngine.IncrementEntry
        CharSheetEngine.TargetList.SetFocus
        SetDataChanged
    End If
    
End Sub

Private Sub cmdESCClose_Click()
'
' Name:         cmdESCClose_Click
' Description:  Close the window when the user presses ESC.
'

    Unload Me

End Sub

Private Sub cmdNote_Click()
'
' Name:         cmdNote_Click
' Description:  Have the CharSheetEngine add a note to the selected target
'               entry.
'
    
    If CharSheetEngine.TargetType = ttXPHistory Then
        If CharSheetEngine.EditXPEntry(lvwHistory, Werewolf.Experience) Then
            RefreshXP
            SetDataChanged
            lvwHistory.SetFocus
        End If
    Else
        CharSheetEngine.AddNote
        SetDataChanged
    End If

End Sub

Private Sub cmdRemove_Click()
'
' Name:         cmdRemove_Click
' Description:  Have the CharSheetEngine remove a label or list entry, OR
'               remove an XP history entry.
'
    
    If CharSheetEngine.TargetType = ttXPHistory Then
        If CharSheetEngine.RemoveXPEntry(lvwHistory, Werewolf.Experience) Then
            RefreshXP
            SetDataChanged
            lvwHistory.SetFocus
        End If
    Else
        CharSheetEngine.RemoveEntry
        SetDataChanged
    End If
    
End Sub

Private Sub cmdRename_Click()
'
' Name:         cmdRename_Click
' Description:  Rename the character.
'

    Dim NewName As String
    
    NewName = InputBox("Enter a new name for the character.", "Rename Character", txtUserField(xiName).Text)
    NewName = Trim(NewName)
    
    If Not (NewName = "" Or NewName = txtUserField(xiName).Text) Then
        CharacterList.MoveTo NewName
        If Not CharacterList.Off Then
            MsgBox "The name """ & NewName & _
                    """ is already in use.  Please use a different name.", _
                    vbOKOnly Or vbExclamation, "Duplicate Name"
        Else
            Game.Rename qiCharacters, txtUserField(xiName).Text, NewName
            txtUserField(xiName).Text = NewName
            mdiMain.AnnounceChanges Me, atCharacters
        End If
    End If

End Sub

Private Sub cmdTraitMax_Click(Index As Integer)
'
' Name:         cmdTraitMax_Click
' Description:  Increment or decrement the trait max for the P/S/M traitlists.
'

    Dim TargetIndex As Integer
    
    TargetIndex = -1
    If CharSheetEngine.TargetType = ttListBox Then TargetIndex = CharSheetEngine.TargetList.Index

    If TargetIndex = tiPhysical Or Game.LinkTraitMaxes Then
        Werewolf.PhysicalMax = Werewolf.PhysicalMax + IIf(Index = 1, 1, -1)
        lblTraitMax(tiPhysical).Caption = "Max " & Werewolf.PhysicalMax
    End If
    
    If TargetIndex = tiSocial Or Game.LinkTraitMaxes Then
        Werewolf.SocialMax = Werewolf.SocialMax + IIf(Index = 1, 1, -1)
        lblTraitMax(tiSocial).Caption = "Max " & Werewolf.SocialMax
    End If
    
    If TargetIndex = tiMental Or Game.LinkTraitMaxes Then
        Werewolf.MentalMax = Werewolf.MentalMax + IIf(Index = 1, 1, -1)
        lblTraitMax(tiMental).Caption = "Max " & Werewolf.MentalMax
    End If
    
    SetDataChanged
    
End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  Update the experience total in case it changed elsewhere.  Re-Acquaint
'               the CharacterSheetEngine with the character.
'
    
    If fraTab(xpFrame).Visible Then RefreshXP
    
    Populating = True
    
    lblModified.Caption = Format(Werewolf.LastModified, "mmmm d, yyyy")
    lblShowXP = " " & CStr(Werewolf.Experience.Unspent) & _
            " / " & CStr(Werewolf.Experience.Earned)

    CharSheetEngine.RefreshTraitList lstTraits(tiEquipment), Werewolf.EquipmentList
    CharSheetEngine.RefreshTraitList lstTraits(tiLocations), Werewolf.HangoutList
    lblField(fiPlayer).Caption = Werewolf.Player
    lblField(fiNarrator).Caption = Werewolf.Narrator
    
    If mdiMain.CheckForChanges(Me, atTempers) Then
        Call ChangeTemper(piRage, 0)
        Call ChangeTemper(piGnosis, 0)
        Call ChangeTemper(piWillpower, 0)
        txtPerm(diHonor).Text = CStr(Werewolf.Honor)
        txtTemp(diHonor).Text = CStr(Werewolf.TempHonor)
        txtPerm(diGlory).Text = CStr(Werewolf.Glory)
        txtTemp(diGlory).Text = CStr(Werewolf.TempGlory)
        txtPerm(diWisdom).Text = CStr(Werewolf.Wisdom)
        txtTemp(diWisdom).Text = CStr(Werewolf.TempWisdom)
    End If

    Populating = False
    
End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Checks to make sure that a character is loaded, which happens only
'               when ShowWerewolf is the means of loading the form.  Initializes the
'               MenuStack linked list and the Werewolf Menus.
'

    If Werewolf Is Nothing Then
        MsgBox "Character sheet loaded improperly!"
    Else
        
        Set CharSheetEngine = New CharacterSheetEngineClass
        
        CharSheetEngine.RegisterSheet "Werewolf", lstMenu, lblMenuItem, lblMenuTitle
    
        CharSheetEngine.RegisterTraitList tiPhysical, Werewolf.PhysicalList
        CharSheetEngine.RegisterTraitList tiSocial, Werewolf.SocialList
        CharSheetEngine.RegisterTraitList tiMental, Werewolf.MentalList
        CharSheetEngine.RegisterTraitList tiPhysicalNeg, Werewolf.PhysicalNegList
        CharSheetEngine.RegisterTraitList tiSocialNeg, Werewolf.SocialNegList
        CharSheetEngine.RegisterTraitList tiMentalNeg, Werewolf.MentalNegList
        CharSheetEngine.RegisterTraitList tiAbilities, Werewolf.AbilityList
        CharSheetEngine.RegisterTraitList tiInfluences, Werewolf.InfluenceList
        CharSheetEngine.RegisterTraitList tiBackgrounds, Werewolf.BackgroundList
        CharSheetEngine.RegisterTraitList tiHonor, Werewolf.HonorList
        CharSheetEngine.RegisterTraitList tiGlory, Werewolf.GloryList
        CharSheetEngine.RegisterTraitList tiWisdom, Werewolf.WisdomList
        CharSheetEngine.RegisterTraitList tiHealth, Werewolf.HealthList
        
        CharSheetEngine.RegisterTraitList tiMerits, Werewolf.MeritList
        CharSheetEngine.RegisterTraitList tiFlaws, Werewolf.FlawList
        CharSheetEngine.RegisterTraitList tiGifts, Werewolf.GiftList
        CharSheetEngine.RegisterTraitList tiRites, Werewolf.RiteList
        CharSheetEngine.RegisterTraitList tiFeatures, Werewolf.FeatureList
        CharSheetEngine.RegisterTraitList tiEquipment, Werewolf.EquipmentList
        CharSheetEngine.RegisterTraitList tiLocations, Werewolf.HangoutList
    
    End If
    mdiMain.OrientForm Me
    
End Sub

Private Sub Form_Deactivate()
'
' Name:         Form_Deactivate
' Description:  Save the text.
' Arguments:
' Returns:
'

    ValidateControls
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Unload
' Description:  Save the text and destroy the MenuStack.
'

    ValidateControls
    Set CharSheetEngine = Nothing
    
End Sub

Private Sub lblField_Change(Index As Integer)
'
' Name:         lblField_Change
' Description:  Store the new value in the appropriate property of the character.
'

    Dim Value As String
    
    If Not Populating Then
        Value = lblField(Index).Caption
        SetDataChanged
        Select Case Index
            Case fiTribe
                Werewolf.Tribe = Value
                mdiMain.AnnounceChanges Me, atCharacters
            Case fiAuspice
                Werewolf.Auspice = Value
                mdiMain.AnnounceChanges Me, atCharacters
            Case fiBreed
                Werewolf.Breed = Value
            Case fiNature
                Werewolf.Nature = Value
            Case fiDemeanor
                Werewolf.Demeanor = Value
            Case fiRank
                Werewolf.Rank = Value
            Case fiPosition
                Werewolf.Position = Value
            Case fiTotem
                Werewolf.Totem = Value
            Case fiCamp
                Werewolf.Camp = Value
            Case fiPlayer
                Werewolf.Player = Value
            Case fiStatus
                Werewolf.Status = Value
                mdiMain.AnnounceChanges Me, atCharacters
            Case fiNarrator
                Werewolf.Narrator = Value
        End Select
    End If
    
End Sub

Private Sub lblField_Click(Index As Integer)
'
' Name:         lblField_Click
' Description:  Appoint a new menu, fill the list box.
'

    If Not (CharSheetEngine.TargetType = ttLabel And _
            CharSheetEngine.TargetLabel Is lblField(Index)) Then
        
        cmdNote.Caption = "Add &Note to Entry"
        cmdCustom.Caption = "&Custom Entry"
        lblTemperFloat.Visible = False
        
        If CharSheetEngine.TargetType = ttListBox Then
            CharSheetEngine.TargetList.ListIndex = -1
            cmdIncrement.Visible = False
            cmdDecrement.Visible = False
            cmdAscend.Visible = False
            cmdDescend.Visible = False
            cmdTraitMax(0).Visible = Game.LinkTraitMaxes
            cmdTraitMax(1).Visible = Game.LinkTraitMaxes
        End If
        
        lblMenuTitle.Caption = lblFieldLabel(Index).Caption
        CharSheetEngine.PopulateMenu lblField(Index).Tag
        CharSheetEngine.TargetType = ttLabel
        Set CharSheetEngine.TargetLabel = lblField(Index)

        lstMenu.SetFocus

    End If
    
End Sub

Private Sub lblField_DblClick(Index As Integer)
'
' Name:         lblField_DblClick
' Description:  Cross-Reference this field on a doubleclick.
'

    
    CharSheetEngine.CrossReference

End Sub

Private Sub lblShowXP_Click()
'
' Name:         lblShowXP_Click
' Description:  Jump to the XP tab.
'

    Call ShowXP

End Sub

Private Sub lblTemper_Click(Index As Integer)
'
' Name:         lblTemper_Click
' Description:  Toggle Perm/Temp editing.
'

    If Not lblTemperFloat.Visible Or Not lblTemperFloat.Tag = CStr(Index) Then
        lblTemperFloat.Top = lblTemper(Index).Top + 30
        lblTemperFloat.Left = lblTemper(Index).Left + lblTemper(Index).Width _
                              - lblTemperFloat.Width - 30
        Select Case Index
            Case piWillpower:   lblTemperFloat = " " & CStr(Werewolf.TempWillpower) & " "
            Case piGnosis:      lblTemperFloat = " " & CStr(Werewolf.TempGnosis) & " "
            Case piRage:        lblTemperFloat = " " & CStr(Werewolf.TempRage) & " "
        End Select
        lblTemperFloat.Height = lblTemper(Index).Height - 60
        lblTemperFloat.Visible = True
        lblTemperFloat.Tag = CStr(Index)
    Else
        lblTemperFloat.Visible = False
    End If

End Sub

Private Sub lblTemperFloat_Click()
'
' Name:         lblTemperFloat_Click
' Description:  Deactivate temporary temper editing (Hide the label).
'
    
    lblTemperFloat.Visible = False

End Sub

Private Sub lstTraits_DblClick(Index As Integer)
'
' Name:         lstTraits_DblClick
' Description:  Cross-reference a trait list that's double-clicked,
'               for the sake of items and rotes and regents.
'

    
    CharSheetEngine.CrossReference

End Sub

Private Sub lstMenu_Click()
'
' Name:         lstMenu_Click
' Description:  Show the selection below the list.
'

    lblMenuItem.Caption = lstMenu.Text

End Sub

Private Sub lstMenu_DblClick()
'
' Name:         lstMenu_DblClick
' Description:  See cmdAdd_Click
'

    cmdAdd_Click

End Sub

Private Sub lstMenu_KeyPress(KeyAscii As Integer)
'
' Name:         lstMenu_KeyPress
' Description:  See cmdAdd_Click
'
    
    If KeyAscii = vbKeyReturn Then cmdAdd_Click

End Sub

Private Sub lstTraits_GotFocus(Index As Integer)
'
' Name:         lstTraits_GotFocus
' Description:  Attach the Increment/Decrement buttons, shift focus, populate the menus
'

    Dim OrderTop As Integer
    
    If Not (CharSheetEngine.TargetType = ttListBox And _
            CharSheetEngine.TargetList Is lstTraits(Index)) Then
        
        cmdNote.Caption = "Add &Note to Entry"
        cmdCustom.Caption = "&Custom Entry"
        lblTemperFloat.Visible = False
        
        If CharSheetEngine.TargetType = ttListBox Then _
                CharSheetEngine.TargetList.ListIndex = -1
    
        If CharSheetEngine.CanAdjust(Index) Then
            With lstTraits(Index)
                Set cmdDecrement.Container = .Container
                Set cmdIncrement.Container = .Container
                cmdDecrement.Move .Left, .Top - cmdDecrement.Height
                cmdIncrement.Move .Left + .Width - cmdIncrement.Width, .Top - cmdIncrement.Height
                OrderTop = .Top + .Height
            End With
            cmdIncrement.Visible = True
            cmdDecrement.Visible = True
        Else
            cmdIncrement.Visible = False
            cmdDecrement.Visible = False
            OrderTop = lstTraits(Index).Top - cmdAscend.Height
        End If
        
        If CharSheetEngine.CanOrder(Index) Then
            With lstTraits(Index)
                Set cmdDescend.Container = .Container
                Set cmdAscend.Container = .Container
                cmdDescend.Move .Left, OrderTop
                cmdAscend.Move .Left + .Width - cmdAscend.Width, OrderTop
            End With
            cmdDescend.Visible = True
            cmdAscend.Visible = True
        Else
            cmdDescend.Visible = False
            cmdAscend.Visible = False
        End If
        
        If Not Game.LinkTraitMaxes Then
            If Index = tiPhysical Or Index = tiSocial Or Index = tiMental Then
                cmdTraitMax(0).Left = lblTraitMax(Index).Left - cmdTraitMax(0).Width
                cmdTraitMax(1).Left = lblTraitMax(Index).Left + lblTraitMax(1).Width
                cmdTraitMax(0).Visible = True
                cmdTraitMax(1).Visible = True
            Else
                cmdTraitMax(0).Visible = False
                cmdTraitMax(1).Visible = False
            End If
        End If
        
        CharSheetEngine.UpdateMenuTitleFromTraitLabel lblTraits(Index)
        CharSheetEngine.PopulateMenu lstTraits(Index).Tag
        CharSheetEngine.TargetType = ttListBox
        Set CharSheetEngine.TargetList = lstTraits(Index)
        
    End If
        
End Sub

Private Sub lstTraits_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'
' Name:         lstTraits_KeyDown
' Description:  Keyboard shortcuts

    Select Case KeyCode
        Case vbKeyDelete, vbKeyBack
            cmdRemove_Click
        Case vbKeyLeft
            cmdDecrement_Click
            KeyCode = 0
        Case vbKeyRight
            cmdIncrement_Click
            KeyCode = 0
    End Select
    
End Sub

Private Sub lstTraits_KeyPress(Index As Integer, KeyAscii As Integer)
'
' Name:         lstTraits_KeyPress
' Description:  Catch a Delete; kill the current selection.

    Select Case KeyAscii
        Case Asc("-"), Asc("_")
            cmdDecrement_Click
        Case Asc("+"), Asc("=")
            cmdIncrement_Click
    End Select

End Sub

Private Sub lstTraits_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'
' Name:         lstTraits_MouseDown
' Description:  Bring up a context menu.
'

    If Button = vbRightButton Then
        With CharSheetEngine
            If .TargetList Is lstTraits(Index) And .TargetType = ttListBox Then
                .PopUpTraitListMenu Me, lstTraits(Index)
                .TargetType = ttNothing
                Call lstTraits_GotFocus(Index)
                SetDataChanged
            End If
        End With
    End If

End Sub

Private Sub lvwHistory_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'
' Name:         lvwHistory_ColumnClick
' Description:  Change the sort order when the Date column header is clicked.
'
    If ColumnHeader.Index = 1 Then
        lvwHistory.SortOrder = IIf(lvwHistory.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        RefreshXP
    End If
    
End Sub

Private Sub lvwHistory_DblClick()
'
' Name:         lvwHistory_DblClick
' Description:  Edit selected entry.
'
    
    Call cmdNote_Click

End Sub

Private Sub lvwHistory_GotFocus()
'
' Name:         lvwHistory_GotFocus
' Description:  Shift focus to XP History editing
'

    If Not CharSheetEngine.TargetType = ttXPHistory Then
    
        cmdNote.Caption = "&Edit Entry"
        cmdCustom.Caption = "&Clear History"
        
        If CharSheetEngine.TargetType = ttListBox Then
            CharSheetEngine.TargetList.ListIndex = -1
            cmdIncrement.Visible = False
            cmdDecrement.Visible = False
            cmdAscend.Visible = False
            cmdDescend.Visible = False
            cmdTraitMax(0).Visible = Game.LinkTraitMaxes
            cmdTraitMax(1).Visible = Game.LinkTraitMaxes
        End If
        
        lblMenuTitle.Caption = "Experience History"
        CharSheetEngine.PopulateMenu lvwHistory.Tag
        lstMenu.ListIndex = 0
        CharSheetEngine.TargetType = ttXPHistory
    
    End If
    
End Sub

Private Sub tabTabStrip_Click()
'
' Name:         tabTabStrip_Click
' Description:  Clear the menu and targets.  Display correct frame.
'

    If Not fraTab(tabTabStrip.SelectedItem.Index - 1).Visible Then
        
        CharSheetEngine.DeselectMenus
        
        If CharSheetEngine.TargetType = ttListBox Then
            CharSheetEngine.TargetList.ListIndex = -1
            cmdIncrement.Visible = False
            cmdDecrement.Visible = False
            cmdAscend.Visible = False
            cmdDescend.Visible = False
        End If
        cmdTraitMax(0).Visible = Game.LinkTraitMaxes
        cmdTraitMax(1).Visible = Game.LinkTraitMaxes
        
        CharSheetEngine.TargetType = ttNothing
        
        Dim fTab As Frame
        For Each fTab In fraTab
            fTab.Visible = (fTab.Index = tabTabStrip.SelectedItem.Index - 1)
        Next fTab
        
        lblTemperFloat.Visible = False
        Set lblTemperFloat.Container = fraTab(tabTabStrip.SelectedItem.Index - 1)
        lblTemperFloat.ZOrder
        
        If fraTab(xpFrame).Visible Then
            RefreshXP
            lvwHistory.SetFocus
        Else
            cmdNote.Caption = "Add &Note to Entry"
            cmdCustom.Caption = "&Custom Entry"
        End If
        
    End If

End Sub

Private Sub txtExperience_Change(Index As Integer)
'
' Name:         txtExperience_Change
' Description:  Ensure a valid value and save the change to the character's
'               experience.
'
    
    If Not (Populating Or Game.EnforceHistory) And IsNumeric(txtExperience(Index).Text) Then
        Select Case Index
            Case xpUnspent
                Werewolf.Experience.Unspent = CSng(txtExperience(xpUnspent))
            Case xpEarned
                Werewolf.Experience.Earned = CSng(txtExperience(xpEarned))
        End Select
    End If
    
End Sub

Private Sub txtExperience_GotFocus(Index As Integer)
'
' Name:         txtExperience_GotFocus
' Description:  Select the Text.
'

    Call lvwHistory_GotFocus
    SelectText txtExperience(Index)

End Sub

Private Sub txtExperience_KeyPress(Index As Integer, KeyAscii As Integer)
'
' Name:         txtExperience_KeyPress
' Description:  Shortcut the return key.
'

    If KeyAscii = vbKeyReturn Then KeyAscii = 0

End Sub

Private Sub updExperience_DownClick(Index As Integer)
'
' Name:         updExperience_DownClick
' Description:  Update the label and store the new value.
'

    txtExperience(xpUnspent).Text = CStr(Val(txtExperience(xpUnspent).Text) - 1)
    If Index = xpEarned Then
        txtExperience(xpEarned).Text = CStr(Val(txtExperience(xpEarned).Text) - 1)
    End If

End Sub

Private Sub updExperience_GotFocus(Index As Integer)
'
' Name:         updExperience_GotFocus
' Description:  Prepare for XP History editing.
'
    Call lvwHistory_GotFocus

End Sub

Private Sub updExperience_UpClick(Index As Integer)
'
' Name:         updExperience_UpClick
' Description:  Update the label and store the new value.
'

    txtExperience(xpUnspent).Text = CStr(Val(txtExperience(xpUnspent).Text) + 1)
    If Index = xpEarned Then
        txtExperience(xpEarned).Text = CStr(Val(txtExperience(xpEarned).Text) + 1)
    End If

End Sub

Private Sub txtMemo_Validate(Index As Integer, Cancel As Boolean)
'
' Name:         txtMemo_Change
' Description:  Record changes to the memo field.
'

    Select Case Index
        Case miBiography
            If Werewolf.Biography <> txtMemo(miBiography).Text Then
                SetDataChanged
                Werewolf.Biography = TrimWhiteSpace(txtMemo(miBiography))
            End If
        Case miNotes
            If Werewolf.Notes <> txtMemo(miNotes).Text Then
                SetDataChanged
                Werewolf.Notes = TrimWhiteSpace(txtMemo(miNotes))
            End If
    End Select

End Sub

Private Sub RefreshXP()
'
' Name:         RefreshXP
' Description:  Refresh the display of XP values and histories.
'

    Populating = True
    txtExperience(xpUnspent).Text = CStr(Werewolf.Experience.Unspent)
    txtExperience(xpEarned).Text = CStr(Werewolf.Experience.Earned)
    txtExperience(xpUnspent).Locked = Game.EnforceHistory
    txtExperience(xpEarned).Locked = Game.EnforceHistory
    updExperience(xpUnspent).Visible = Not Game.EnforceHistory
    updExperience(xpEarned).Visible = Not Game.EnforceHistory
    lblShowXP = " " & CStr(Werewolf.Experience.Unspent) & _
            " / " & CStr(Werewolf.Experience.Earned)
    Populating = False
        
    CharSheetEngine.RefreshXP lvwHistory, Werewolf.Experience
    
End Sub

Private Sub SetDataChanged()
'
' Name:         SetDataChanged
' Description:  Tell the game its data has changed and update the character's
'               Last Modified date.
'
        
    If Not Populating Then
        Game.DataChanged = True
        Werewolf.LastModified = Now
        lblModified.Caption = Format(Date, "mmmm d, yyyy")
    End If
    
End Sub

Private Sub txtUserField_Change(Index As Integer)
'
' Name:         txtUserField_Change
' Description:  Store a new value in the appropriate space and set the game as
'               changed.
'

    If Not Populating Then

        SetDataChanged

        Select Case Index
            Case xiName
                ' Name changed through cmdRename_Click
                Me.Caption = txtUserField(Index).Text
            Case xiPack
                Werewolf.Pack = Trim(txtUserField(Index))
            Case xiID
                Werewolf.ID = Trim(txtUserField(Index))
            Case xiStartDate
                If IsDate(txtUserField(Index)) Then Werewolf.StartDate = CDate(txtUserField(Index))
        End Select
        
    End If

End Sub

Private Sub txtUserField_GotFocus(Index As Integer)
'
' Name:         txtUserField_GotFocus
' Description:  Highlight text on click.
'

    SelectText txtUserField(Index)

End Sub

Private Sub txtUserField_KeyPress(Index As Integer, KeyAscii As Integer)
'
' Name:         txtUserField_KeyPress
' Description:  Nullify carriage returns.
'
    
    If KeyAscii = vbKeyReturn Then KeyAscii = 0

End Sub

Private Sub updNumber_Change(Index As Integer)
'
' Name:         updNumber_Change
' Description:  Update the label and store the new value.
'

    lblNumber(Index) = " " & CStr(updNumber(Index).Value)
    
    Select Case Index
        Case niNotoriety
            Werewolf.Notoriety = updNumber(niNotoriety).Value
    End Select
    SetDataChanged
    
End Sub

Private Sub ChangeTemper(Index As Integer, Change As Integer)
'
' Name:         ChangeTemper
' Parameters:   Index           index of temper to change
'               Change          amount to change
'               OnlyTemp        edit only the temp value not both temp and perm
' Description:  Change the perm/temp value of the given temper.
'
    
    Dim Both As Boolean
    Dim Perm As Integer
    Dim Temp As Integer
    
    Both = Not lblTemperFloat.Visible

    With Werewolf
        Select Case Index
            Case piWillpower
                .TempWillpower = .TempWillpower + Change
                If Both Then .Willpower = .Willpower + Change
                If .TempWillpower < 0 Then .TempWillpower = 0
                If .Willpower < 0 Then .Willpower = 0
                Temp = .TempWillpower
                Perm = .Willpower
            Case piGnosis
                .TempGnosis = .TempGnosis + Change
                If Both Then .Gnosis = .Gnosis + Change
                If .TempGnosis < 0 Then .TempGnosis = 0
                If .Gnosis < 0 Then .Gnosis = 0
                Temp = .TempGnosis
                Perm = .Gnosis
            Case piRage
                .TempRage = .TempRage + Change
                If Both Then .Rage = .Rage + Change
                If .TempRage < 0 Then .TempRage = 0
                If .Rage < 0 Then .Rage = 0
                Temp = .TempRage
                Perm = .Rage
        End Select
    End With
    
    If Both Then
        If Change > 0 Then
            CharSheetEngine.Purchase lblTemperLabel(Index).Caption, 3
        ElseIf Change < 0 Then
            CharSheetEngine.Refund lblTemperLabel(Index).Caption, 3
        End If
    End If
        
    lblTemper(Index) = " " & CStr(Perm) & " " & CharSheetEngine.DisplayTemper(Perm, Temp)
    lblTemperFloat = " " & CStr(Temp) & " "
    lblTemperFloat.Height = lblTemper(Index).Height - 60
    If Change <> 0 Then mdiMain.AnnounceChanges Me, atTempers
    SetDataChanged
    
End Sub

Private Sub updTemper_DownClick(Index As Integer)
'
' Name:         updTemper_DownClick
' Description:  Change the associated temper.
'
    If lblTemperFloat.Tag <> CStr(Index) Then lblTemperFloat.Visible = False
    Call ChangeTemper(Index, -1)

End Sub

Private Sub updTemper_UpClick(Index As Integer)
'
' Name:         updTemper_UpClick
' Description:  Change the associated temper.
'
    If lblTemperFloat.Tag <> CStr(Index) Then lblTemperFloat.Visible = False
    Call ChangeTemper(Index, 1)

End Sub

Private Sub txtPerm_Change(Index As Integer)
'
' Name:         txtPerm_Change
' Description:  Ensure a valid value and save the change to the character's
'               experience.
'
    
    If Not Populating Then
    
        If IsNumeric(txtPerm(Index).Text) Then
            
            Dim FixVal As Double
            
            txtPerm(Index).ForeColor = vbWindowText

            FixVal = Int(Val(txtPerm(Index).Text))
            If FixVal > 9999 Then FixVal = 9999
            If FixVal < -9999 Then FixVal = -9999
    
            Select Case Index
                Case diHonor
                    Werewolf.Honor = CInt(FixVal)
                Case diGlory
                    Werewolf.Glory = CInt(FixVal)
                Case diWisdom
                    Werewolf.Wisdom = CInt(FixVal)
            End Select
            mdiMain.AnnounceChanges Me, atTempers
            SetDataChanged
            
        Else
            txtPerm(Index).ForeColor = vbHighlight
        End If
        
    End If
    
End Sub

Private Sub txtPerm_GotFocus(Index As Integer)
'
' Name:         txtPerm_GotFocus
' Description:  Select the Text.
'

    SelectText txtPerm(Index)

End Sub

Private Sub txtPerm_KeyPress(Index As Integer, KeyAscii As Integer)
'
' Name:         txtPerm_KeyPress
' Description:  Shortcut the return key.
'

    If KeyAscii = vbKeyReturn Then KeyAscii = 0

End Sub

Private Sub txtTemp_Change(Index As Integer)
'
' Name:         txtTemp_Change
' Description:  Ensure a valid value and save the change to the character's
'               experience.
'
    
    If Not Populating Then
    
        If IsNumeric(txtTemp(Index).Text) Then
            
            txtTemp(Index).ForeColor = vbWindowText
    
            Select Case Index
                Case diHonor
                    Werewolf.TempHonor = Val(txtTemp(diHonor))
                Case diGlory
                    Werewolf.TempGlory = Val(txtTemp(diGlory))
                Case diWisdom
                    Werewolf.TempWisdom = Val(txtTemp(diWisdom))
            End Select
            mdiMain.AnnounceChanges Me, atTempers
            SetDataChanged
            
        Else
            txtTemp(Index).ForeColor = vbHighlight
        End If
        
    End If
    
End Sub

Private Sub txtTemp_GotFocus(Index As Integer)
'
' Name:         txtTemp_GotFocus
' Description:  Select the Text.
'

    SelectText txtTemp(Index)

End Sub

Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
'
' Name:         txtTemp_KeyPress
' Description:  Shortcut the return key.
'

    If KeyAscii = vbKeyReturn Then KeyAscii = 0

End Sub

Private Sub updPermTemp_DownClick(Index As Integer)
'
' Name:         updPermTemp_DownClick
' Description:  Update the text and store the new value.
'

    If Me.ActiveControl Is txtPerm(Index) Then
        txtPerm(Index).Text = CStr(Val(txtPerm(Index).Text) - 1)
    Else
        txtTemp(Index).Text = CStr(Val(txtTemp(Index).Text) - 1)
    End If
    
End Sub

Private Sub updPermTemp_UpClick(Index As Integer)
'
' Name:         updPermTemp_UpClick
' Description:  Update the text and store the new value.
'

    If Me.ActiveControl Is txtPerm(Index) Then
        txtPerm(Index).Text = CStr(Val(txtPerm(Index).Text) + 1)
    Else
        txtTemp(Index).Text = CStr(Val(txtTemp(Index).Text) + 1)
    End If

End Sub

