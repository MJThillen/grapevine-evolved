VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChangelingSheet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Changeling Character"
   ClientHeight    =   6165
   ClientLeft      =   210
   ClientTop       =   405
   ClientWidth     =   9060
   Icon            =   "frmChangelingSheet.frx":0000
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
      TabIndex        =   109
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
      Begin VB.CheckBox chkNPC 
         Alignment       =   1  'Right Justify
         Caption         =   "NPC"
         Height          =   375
         Left            =   3735
         TabIndex        =   36
         Top             =   2400
         Width           =   660
      End
      Begin VB.TextBox txtUserField 
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   32
         Top             =   960
         Width           =   2175
      End
      Begin MSComCtl2.UpDown updTemper 
         Height          =   315
         Index           =   1
         Left            =   5775
         TabIndex        =   42
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
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   3855
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         Value           =   2
         Max             =   1000
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updTemper 
         Height          =   315
         Index           =   0
         Left            =   5775
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2895
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         Value           =   4
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
         TabIndex        =   118
         Top             =   3000
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblShowXP 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4200
         TabIndex        =   117
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
         Tag             =   "Kith"
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kith"
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
         Index           =   1
         Left            =   960
         TabIndex        =   16
         Tag             =   "Seeming"
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Seeming"
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
         Index           =   2
         Left            =   960
         TabIndex        =   18
         Tag             =   "Seelie"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Seelie Legacy "
         Height          =   435
         Index           =   2
         Left            =   30
         TabIndex        =   17
         Top             =   1410
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
         Tag             =   "Unseelie"
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Unseelie Legacy "
         Height          =   495
         Index           =   3
         Left            =   30
         TabIndex        =   19
         Top             =   1860
         Width           =   855
      End
      Begin VB.Label lblTemper 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   41
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label lblTemperLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Banality"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   40
         Top             =   3390
         Width           =   855
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   255
         Index           =   4
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
         Index           =   5
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
         Index           =   4
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
         Index           =   5
         Left            =   4200
         TabIndex        =   30
         Tag             =   "?PL"
         Top             =   480
         Width           =   2175
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
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Court"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   21
         Top             =   2430
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   960
         TabIndex        =   22
         Tag             =   "Court"
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "House"
         Height          =   255
         Index           =   7
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
         Index           =   7
         Left            =   960
         TabIndex        =   24
         Tag             =   "House"
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Musing/ Ravaging Threshold "
         Height          =   615
         Index           =   9
         Left            =   30
         TabIndex        =   27
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   9
         Left            =   960
         TabIndex        =   28
         Tag             =   "Threshold"
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         Height          =   255
         Index           =   8
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
         Index           =   8
         Left            =   960
         TabIndex        =   26
         Tag             =   "Title"
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label lblTemperLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Glamour"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   37
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
         TabIndex        =   38
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label lblUserField 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Character ID"
         Height          =   255
         Index           =   1
         Left            =   3195
         TabIndex        =   31
         Top             =   1005
         Width           =   900
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4785
      Index           =   5
      Left            =   2160
      TabIndex        =   90
      Top             =   1200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdEstimate 
         Height          =   315
         Left            =   2535
         Picture         =   "frmChangelingSheet.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   1470
         Width           =   585
      End
      Begin VB.TextBox txtUserField 
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   103
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtExperience 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   93
         Top             =   495
         Width           =   1575
      End
      Begin VB.TextBox txtExperience 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   96
         Top             =   960
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvwHistory 
         Height          =   2430
         Left            =   105
         TabIndex        =   107
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
         TabIndex        =   97
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
         TabIndex        =   94
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
         TabIndex        =   98
         Top             =   1530
         Width           =   2535
      End
      Begin VB.Label lblModifiedLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Modified"
         Height          =   375
         Left            =   3360
         TabIndex        =   104
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblModified 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4200
         TabIndex        =   105
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Assigned Narrator"
         Height          =   375
         Index           =   10
         Left            =   3360
         TabIndex        =   100
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   10
         Left            =   4200
         TabIndex        =   101
         Tag             =   "?NR"
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblUserField 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Date"
         Height          =   375
         Index           =   2
         Left            =   3360
         TabIndex        =   102
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblXPLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Experience"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   91
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblXPLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Earned"
         Height          =   375
         Index           =   1
         Left            =   -120
         TabIndex        =   95
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label lblXPLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Unspent"
         Height          =   375
         Index           =   0
         Left            =   -120
         TabIndex        =   92
         Top             =   540
         Width           =   975
      End
      Begin VB.Label lblXPLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Experience &History"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   106
         Top             =   1920
         Width           =   6375
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   4
      Left            =   2160
      TabIndex        =   81
      Top             =   1200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtMemo 
         Height          =   1035
         Index           =   1
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   85
         Top             =   480
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   1035
         Index           =   16
         ItemData        =   "frmChangelingSheet.frx":0B14
         Left            =   120
         List            =   "frmChangelingSheet.frx":0B16
         TabIndex        =   83
         Tag             =   "?LO"
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtMemo 
         Height          =   2400
         Index           =   2
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   89
         Top             =   2160
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   2400
         Index           =   15
         ItemData        =   "frmChangelingSheet.frx":0B18
         Left            =   120
         List            =   "frmChangelingSheet.frx":0B1A
         TabIndex        =   87
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
         TabIndex        =   84
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Favorite Locations"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   82
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   88
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Equip. / Chimera / Treasure"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   86
         Top             =   1920
         Width           =   3135
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   3
      Left            =   2160
      TabIndex        =   70
      Top             =   1200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.ListBox lstTraits 
         Height          =   1035
         Index           =   10
         ItemData        =   "frmChangelingSheet.frx":0B1C
         Left            =   4440
         List            =   "frmChangelingSheet.frx":0B1E
         TabIndex        =   76
         Tag             =   "Health Levels"
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtMemo 
         Height          =   1035
         Index           =   0
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   74
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   1035
         Index           =   9
         ItemData        =   "frmChangelingSheet.frx":0B20
         Left            =   120
         List            =   "frmChangelingSheet.frx":0B22
         TabIndex        =   72
         Tag             =   "Status"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2400
         Index           =   13
         ItemData        =   "frmChangelingSheet.frx":0B24
         Left            =   120
         List            =   "frmChangelingSheet.frx":0B26
         TabIndex        =   78
         Tag             =   "Arts"
         Top             =   2160
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   2400
         Index           =   14
         ItemData        =   "frmChangelingSheet.frx":0B28
         Left            =   3360
         List            =   "frmChangelingSheet.frx":0B2A
         TabIndex        =   80
         Tag             =   "Realms"
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Health Levels"
         Height          =   255
         Index           =   10
         Left            =   4440
         TabIndex        =   75
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Oaths"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   73
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Status Traits"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   71
         Tag             =   "Status Traits|Status Trait"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Arts"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   77
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Realms"
         Height          =   255
         Index           =   14
         Left            =   3360
         TabIndex        =   79
         Top             =   1920
         Width           =   3135
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   1
      Left            =   2160
      TabIndex        =   46
      Top             =   1200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdTraitMax 
         Caption         =   "+"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   114
         Top             =   2520
         Width           =   255
      End
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
         TabIndex        =   113
         Top             =   2520
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
         Left            =   120
         TabIndex        =   10
         Top             =   2520
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
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   255
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
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   0
         ItemData        =   "frmChangelingSheet.frx":0B2C
         Left            =   120
         List            =   "frmChangelingSheet.frx":0B2E
         TabIndex        =   48
         Tag             =   "Physical"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   1
         ItemData        =   "frmChangelingSheet.frx":0B30
         Left            =   2280
         List            =   "frmChangelingSheet.frx":0B32
         TabIndex        =   50
         Tag             =   "Social"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   2
         ItemData        =   "frmChangelingSheet.frx":0B34
         Left            =   4440
         List            =   "frmChangelingSheet.frx":0B36
         TabIndex        =   52
         Tag             =   "Mental"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   3
         ItemData        =   "frmChangelingSheet.frx":0B38
         Left            =   120
         List            =   "frmChangelingSheet.frx":0B3A
         TabIndex        =   54
         Tag             =   "Physical, Negative"
         Top             =   3120
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   4
         ItemData        =   "frmChangelingSheet.frx":0B3C
         Left            =   2280
         List            =   "frmChangelingSheet.frx":0B3E
         TabIndex        =   56
         Tag             =   "Social, Negative"
         Top             =   3120
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   5
         ItemData        =   "frmChangelingSheet.frx":0B40
         Left            =   4440
         List            =   "frmChangelingSheet.frx":0B42
         TabIndex        =   58
         Tag             =   "Mental, Negative"
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label lblTraitMax 
         Alignment       =   2  'Center
         Caption         =   "Max 11"
         Height          =   255
         Index           =   2
         Left            =   5040
         TabIndex        =   116
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblTraitMax 
         Alignment       =   2  'Center
         Caption         =   "Max 11"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   115
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblTraitMax 
         Alignment       =   2  'Center
         Caption         =   "Max 11"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   112
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
         TabIndex        =   47
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
         TabIndex        =   49
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
         TabIndex        =   51
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
         TabIndex        =   53
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
         TabIndex        =   55
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
         TabIndex        =   57
         Tag             =   "Negative Mental"
         Top             =   2880
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdESCClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   111
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
      ItemData        =   "frmChangelingSheet.frx":0B44
      Left            =   120
      List            =   "frmChangelingSheet.frx":0B46
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
      TabIndex        =   110
      Top             =   120
      Width           =   6855
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   2
      Left            =   2160
      TabIndex        =   59
      Top             =   1200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   6
         ItemData        =   "frmChangelingSheet.frx":0B48
         Left            =   120
         List            =   "frmChangelingSheet.frx":0B4A
         TabIndex        =   61
         Tag             =   "Abilities"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   7
         ItemData        =   "frmChangelingSheet.frx":0B4C
         Left            =   4440
         List            =   "frmChangelingSheet.frx":0B4E
         TabIndex        =   65
         Tag             =   "Influences"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   11
         ItemData        =   "frmChangelingSheet.frx":0B50
         Left            =   120
         List            =   "frmChangelingSheet.frx":0B52
         TabIndex        =   67
         Tag             =   "Merits"
         Top             =   3120
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   12
         ItemData        =   "frmChangelingSheet.frx":0B54
         Left            =   3360
         List            =   "frmChangelingSheet.frx":0B56
         TabIndex        =   69
         Tag             =   "Flaws"
         Top             =   3120
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   8
         ItemData        =   "frmChangelingSheet.frx":0B58
         Left            =   2280
         List            =   "frmChangelingSheet.frx":0B5A
         TabIndex        =   63
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
         TabIndex        =   60
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
         TabIndex        =   64
         Tag             =   "Influences|Influence"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Merits"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   66
         Top             =   2880
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Flaws"
         Height          =   255
         Index           =   12
         Left            =   3360
         TabIndex        =   68
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
         TabIndex        =   62
         Tag             =   "Backgrounds|Background"
         Top             =   240
         Width           =   2055
      End
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
         NumTabs         =   6
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
            Caption         =   "Sta Oath HL Art Realm"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Loc Bio Eqp Note"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  XP"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   480
      Picture         =   "frmChangelingSheet.frx":0B5C
      Top             =   180
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
      TabIndex        =   108
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
Attribute VB_Name = "frmChangelingSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Constants by which specific labels are indexed. (fi = Field Index)
Private Const fiKith = 0
Private Const fiSeeming = 1
Private Const fiSeelieLegacy = 2
Private Const fiUnseelieLegacy = 3
Private Const fiStatus = 4
Private Const fiPlayer = 5
Private Const fiCourt = 6
Private Const fiHouse = 7
Private Const fiTitle = 8
Private Const fiThreshold = 9
Private Const fiNarrator = 10

'Constants by which specific temper labels are indexed (pi = Point Index)
Private Const piGlamour = 0
Private Const piBanality = 1
Private Const piWillpower = 2

' Constants by which specific list boxes are indexed.
Private Const tiPhysical = 0
Private Const tiSocial = 1
Private Const tiMental = 2
Private Const tiPhysicalNeg = 3
Private Const tiSocialNeg = 4
Private Const tiMentalNeg = 5
Private Const tiAbilities = 6
Private Const tiInfluences = 7
Private Const tiBackgrounds = 8
Private Const tiStatus = 9
Private Const tiHealth = 10
Private Const tiMerits = 11
Private Const tiFlaws = 12
Private Const tiArts = 13
Private Const tiRealms = 14
Private Const tiEquipment = 15
Private Const tiLocations = 16

' Constants by which specific text boxes are indexed. (xi = Text Index)
Private Const xiName = 0
Private Const xiID = 1
Private Const xiStartDate = 2

' Constants by which specific memo fields are indexed. (mi = Memo Index)
Private Const miOaths = 0
Private Const miBiography = 1
Private Const miNotes = 2

' Constant by which to reference the index of the XP frame and text boxes
Private Const xpFrame = 5
Private Const xpUnspent = 0
Private Const xpEarned = 1

Private Changeling As ChangelingClass                   'The Changeling character
Private CharSheetEngine As CharacterSheetEngineClass    'Handles common functions
Private Populating As Boolean                           'defuses some events when characters are loaded

Public Sub ShowChangeling(Character As ChangelingClass)
'
' Name:         ShowChangeling
' Description:  Show and initialize a new instance of the form.
' Arguments:    The ChangelingClass whose data the form modifies.
'

    Dim DataState As Boolean

    Populating = True

    Set Changeling = Nothing
    Set Changeling = Character
    DataState = Game.DataChanged

    txtUserField(xiName) = Changeling.Name
    Me.Caption = Changeling.Name

    lblField(fiKith) = Changeling.Kith
    lblField(fiSeeming) = Changeling.Seeming
    lblField(fiSeelieLegacy) = Changeling.SeelieLegacy
    lblField(fiUnseelieLegacy) = Changeling.UnseelieLegacy
    lblField(fiStatus) = Changeling.Status
    lblField(fiPlayer) = Changeling.Player
    lblField(fiCourt) = Changeling.Court
    lblField(fiHouse) = Changeling.House
    lblField(fiTitle) = Changeling.Title
    lblField(fiThreshold) = Changeling.Threshold
    lblField(fiNarrator) = Changeling.Narrator
        
    lblTraitMax(tiPhysical) = "Max " & CStr(Changeling.PhysicalMax)
    lblTraitMax(tiSocial) = "Max " & CStr(Changeling.SocialMax)
    lblTraitMax(tiMental) = "Max " & CStr(Changeling.MentalMax)

    Call ChangeTemper(piGlamour, 0)
    Call ChangeTemper(piBanality, 0)
    Call ChangeTemper(piWillpower, 0)
    
    txtUserField(xiID) = Changeling.ID
    txtUserField(xiStartDate) = Format(Changeling.StartDate, "mmmm d, yyyy")
    
    txtMemo(miOaths) = Changeling.Oaths
    txtMemo(miBiography) = Changeling.Biography
    txtMemo(miNotes) = Changeling.Notes
    
    CharSheetEngine.RefreshTraitList lstTraits(tiPhysical), Changeling.PhysicalList
    CharSheetEngine.RefreshTraitList lstTraits(tiSocial), Changeling.SocialList
    CharSheetEngine.RefreshTraitList lstTraits(tiMental), Changeling.MentalList
    CharSheetEngine.RefreshTraitList lstTraits(tiPhysicalNeg), Changeling.PhysicalNegList
    CharSheetEngine.RefreshTraitList lstTraits(tiSocialNeg), Changeling.SocialNegList
    CharSheetEngine.RefreshTraitList lstTraits(tiMentalNeg), Changeling.MentalNegList
    CharSheetEngine.RefreshTraitList lstTraits(tiAbilities), Changeling.AbilityList
    CharSheetEngine.RefreshTraitList lstTraits(tiInfluences), Changeling.InfluenceList
    CharSheetEngine.RefreshTraitList lstTraits(tiBackgrounds), Changeling.BackgroundList
    CharSheetEngine.RefreshTraitList lstTraits(tiStatus), Changeling.StatusList
    CharSheetEngine.RefreshTraitList lstTraits(tiHealth), Changeling.HealthList
    
    CharSheetEngine.RefreshTraitList lstTraits(tiMerits), Changeling.MeritList
    CharSheetEngine.RefreshTraitList lstTraits(tiFlaws), Changeling.FlawList
    CharSheetEngine.RefreshTraitList lstTraits(tiArts), Changeling.ArtList
    CharSheetEngine.RefreshTraitList lstTraits(tiRealms), Changeling.RealmList
    CharSheetEngine.RefreshTraitList lstTraits(tiEquipment), Changeling.EquipmentList
    CharSheetEngine.RefreshTraitList lstTraits(tiLocations), Changeling.HangoutList
    
    lblModified.Caption = Format(Changeling.LastModified, "mmmm d, yyyy")
    chkNPC.Value = IIf(Changeling.IsNPC, vbChecked, vbUnchecked)
    
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
        .SelectSet(osCharacters).Add Changeling.Name
        .GameDate = 0
    End With
    
End Sub

Private Sub chkNPC_Click()
'
' Name:         chkNPC_Click
' Description:  Toggle whether the character is an NPC.
'

    If Not Populating Then
        Changeling.IsNPC = (chkNPC.Value = vbChecked)
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
            If CharSheetEngine.AddXPEntry(lvwHistory, Changeling.Experience) Then
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
            Changeling.Experience.Clear
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
    frmCalculator.ShowCalculator Changeling

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
'               entry, OR edit a history entry.
'
    
    If CharSheetEngine.TargetType = ttXPHistory Then
        If CharSheetEngine.EditXPEntry(lvwHistory, Changeling.Experience) Then
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
        If CharSheetEngine.RemoveXPEntry(lvwHistory, Changeling.Experience) Then
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
        Changeling.PhysicalMax = Changeling.PhysicalMax + IIf(Index = 1, 1, -1)
        lblTraitMax(tiPhysical).Caption = "Max " & Changeling.PhysicalMax
    End If
    
    If TargetIndex = tiSocial Or Game.LinkTraitMaxes Then
        Changeling.SocialMax = Changeling.SocialMax + IIf(Index = 1, 1, -1)
        lblTraitMax(tiSocial).Caption = "Max " & Changeling.SocialMax
    End If
    
    If TargetIndex = tiMental Or Game.LinkTraitMaxes Then
        Changeling.MentalMax = Changeling.MentalMax + IIf(Index = 1, 1, -1)
        lblTraitMax(tiMental).Caption = "Max " & Changeling.MentalMax
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
    
    lblModified.Caption = Format(Changeling.LastModified, "mmmm d, yyyy")
    lblShowXP = " " & CStr(Changeling.Experience.Unspent) & _
            " / " & CStr(Changeling.Experience.Earned)
    
    CharSheetEngine.RefreshTraitList lstTraits(tiEquipment), Changeling.EquipmentList
    CharSheetEngine.RefreshTraitList lstTraits(tiLocations), Changeling.HangoutList
    lblField(fiPlayer).Caption = Changeling.Player
    lblField(fiNarrator).Caption = Changeling.Narrator
    
    If mdiMain.CheckForChanges(Me, atTempers) Then
        Call ChangeTemper(piGlamour, 0)
        Call ChangeTemper(piBanality, 0)
        Call ChangeTemper(piWillpower, 0)
    End If

    Populating = False

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

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Checks to make sure that a character is loaded, which happens only
'               when ShowChangeling is the means of loading the form.  Initializes the
'               MenuStack linked list and the Changeling Menus.
'

    If Changeling Is Nothing Then
        MsgBox "Character sheet loaded improperly!"
    Else
                
        Set CharSheetEngine = New CharacterSheetEngineClass
        
        CharSheetEngine.RegisterSheet "Changeling", lstMenu, lblMenuItem, lblMenuTitle
                
        CharSheetEngine.RegisterTraitList tiPhysical, Changeling.PhysicalList
        CharSheetEngine.RegisterTraitList tiSocial, Changeling.SocialList
        CharSheetEngine.RegisterTraitList tiMental, Changeling.MentalList
        CharSheetEngine.RegisterTraitList tiPhysicalNeg, Changeling.PhysicalNegList
        CharSheetEngine.RegisterTraitList tiSocialNeg, Changeling.SocialNegList
        CharSheetEngine.RegisterTraitList tiMentalNeg, Changeling.MentalNegList
        CharSheetEngine.RegisterTraitList tiAbilities, Changeling.AbilityList
        CharSheetEngine.RegisterTraitList tiInfluences, Changeling.InfluenceList
        CharSheetEngine.RegisterTraitList tiBackgrounds, Changeling.BackgroundList
        CharSheetEngine.RegisterTraitList tiStatus, Changeling.StatusList
        CharSheetEngine.RegisterTraitList tiHealth, Changeling.HealthList
    
        CharSheetEngine.RegisterTraitList tiMerits, Changeling.MeritList
        CharSheetEngine.RegisterTraitList tiFlaws, Changeling.FlawList
        CharSheetEngine.RegisterTraitList tiArts, Changeling.ArtList
        CharSheetEngine.RegisterTraitList tiRealms, Changeling.RealmList
        CharSheetEngine.RegisterTraitList tiEquipment, Changeling.EquipmentList
        CharSheetEngine.RegisterTraitList tiLocations, Changeling.HangoutList
        
    End If
    mdiMain.OrientForm Me
    
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
            Case fiKith
                Changeling.Kith = Value
                mdiMain.AnnounceChanges Me, atCharacters
            Case fiSeeming
                Changeling.Seeming = Value
                mdiMain.AnnounceChanges Me, atCharacters
            Case fiSeelieLegacy
                Changeling.SeelieLegacy = Value
            Case fiUnseelieLegacy
                Changeling.UnseelieLegacy = Value
            Case fiPlayer
                Changeling.Player = Value
            Case fiStatus
                Changeling.Status = Value
                mdiMain.AnnounceChanges Me, atCharacters
            Case fiCourt
                Changeling.Court = Value
            Case fiHouse
                Changeling.House = Value
            Case fiTitle
                Changeling.Title = Value
            Case fiThreshold
                Changeling.Threshold = Value
            Case fiNarrator
                Changeling.Narrator = Value
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
            Case piWillpower:   lblTemperFloat = " " & CStr(Changeling.TempWillpower) & " "
            Case piGlamour:     lblTemperFloat = " " & CStr(Changeling.TempGlamour) & " "
            Case piBanality:    lblTemperFloat = " " & CStr(Changeling.TempBanality) & " "
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
                Changeling.Experience.Unspent = CSng(txtExperience(xpUnspent))
            Case xpEarned
                Changeling.Experience.Earned = CSng(txtExperience(xpEarned))
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
            If Changeling.Biography <> txtMemo(miBiography).Text Then
                SetDataChanged
                Changeling.Biography = TrimWhiteSpace(txtMemo(miBiography))
            End If
        Case miNotes
            If Changeling.Notes <> txtMemo(miNotes).Text Then
                SetDataChanged
                Changeling.Notes = TrimWhiteSpace(txtMemo(miNotes))
            End If
        Case miOaths
            If Changeling.Oaths <> txtMemo(miOaths).Text Then
                SetDataChanged
                Changeling.Oaths = TrimWhiteSpace(txtMemo(miOaths))
            End If
    End Select

End Sub

Private Sub RefreshXP()
'
' Name:         RefreshXP
' Description:  Refresh the display of XP values and histories.
'

    Populating = True
    txtExperience(xpUnspent).Text = CStr(Changeling.Experience.Unspent)
    txtExperience(xpEarned).Text = CStr(Changeling.Experience.Earned)
    txtExperience(xpUnspent).Locked = Game.EnforceHistory
    txtExperience(xpEarned).Locked = Game.EnforceHistory
    updExperience(xpUnspent).Visible = Not Game.EnforceHistory
    updExperience(xpEarned).Visible = Not Game.EnforceHistory
    lblShowXP = " " & CStr(Changeling.Experience.Unspent) & _
            " / " & CStr(Changeling.Experience.Earned)
    Populating = False
        
    CharSheetEngine.RefreshXP lvwHistory, Changeling.Experience
    
End Sub

Private Sub SetDataChanged()
'
' Name:         SetDataChanged
' Description:  Tell the game its data has changed and update the character's
'               Last Modified date.
'
        
    If Not Populating Then
        Game.DataChanged = True
        Changeling.LastModified = Now
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
            Case xiID
                Changeling.ID = Trim(txtUserField(Index))
            Case xiStartDate
                If IsDate(txtUserField(Index)) Then Changeling.StartDate = CDate(txtUserField(Index))
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

    With Changeling
        Select Case Index
            Case piWillpower
                .TempWillpower = .TempWillpower + Change
                If Both Then .Willpower = .Willpower + Change
                If .TempWillpower < 0 Then .TempWillpower = 0
                If .Willpower < 0 Then .Willpower = 0
                Temp = .TempWillpower
                Perm = .Willpower
            Case piGlamour
                .TempGlamour = .TempGlamour + Change
                If Both Then .Glamour = .Glamour + Change
                If .TempGlamour < 0 Then .TempGlamour = 0
                If .Glamour < 0 Then .Glamour = 0
                Temp = .TempGlamour
                Perm = .Glamour
            Case piBanality
                .TempBanality = .TempBanality + Change
                If Both Then .Banality = .Banality + Change
                If .TempBanality < 0 Then .TempBanality = 0
                If .Banality < 0 Then .Banality = 0
                Temp = .TempBanality
                Perm = .Banality
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
