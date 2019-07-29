VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWraithSheet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wraith Character"
   ClientHeight    =   6165
   ClientLeft      =   210
   ClientTop       =   405
   ClientWidth     =   9060
   Icon            =   "frmWraithSheet.frx":0000
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
      TabIndex        =   125
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
         TabIndex        =   35
         Top             =   2400
         Width           =   660
      End
      Begin VB.TextBox txtUserField 
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   30
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox cboEthnos 
         Height          =   315
         ItemData        =   "frmWraithSheet.frx":058A
         Left            =   960
         List            =   "frmWraithSheet.frx":0597
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   480
         Width           =   2175
      End
      Begin MSComCtl2.UpDown updTemper 
         Height          =   315
         Index           =   1
         Left            =   5775
         TabIndex        =   41
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
         TabIndex        =   44
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
         TabIndex        =   38
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
         TabIndex        =   133
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
         TabIndex        =   33
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label lblShowXP 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4200
         TabIndex        =   34
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
         TabIndex        =   16
         Tag             =   "Guild"
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Guild"
         Height          =   255
         Index           =   0
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
         TabIndex        =   18
         Tag             =   "Archetypes"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nature"
         Height          =   255
         Index           =   1
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
         TabIndex        =   20
         Tag             =   "Archetypes"
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Demeanor"
         Height          =   255
         Index           =   2
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
         Index           =   3
         Left            =   960
         TabIndex        =   22
         Tag             =   "Faction"
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Faction"
         Height          =   255
         Index           =   3
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   40
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label lblTemperLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Corpus"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   39
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
         TabIndex        =   31
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
         TabIndex        =   27
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
         TabIndex        =   32
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
         TabIndex        =   28
         Tag             =   "?PL"
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Legion"
         Height          =   255
         Index           =   6
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
         Index           =   6
         Left            =   960
         TabIndex        =   24
         Tag             =   "Legion"
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rank"
         Height          =   255
         Index           =   7
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
         Index           =   7
         Left            =   960
         TabIndex        =   26
         Tag             =   "Rank"
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label lblTemperLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pathos"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   36
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
         TabIndex        =   37
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
         TabIndex        =   29
         Top             =   1005
         Width           =   900
      End
      Begin VB.Label lblEthnos 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ethnos"
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   540
         Width           =   855
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4785
      Index           =   6
      Left            =   2160
      TabIndex        =   107
      Top             =   1200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdEstimate 
         Height          =   315
         Left            =   2535
         Picture         =   "frmWraithSheet.frx":05B3
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   1470
         Width           =   585
      End
      Begin VB.TextBox txtExperience 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   113
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtExperience 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   110
         Top             =   495
         Width           =   1575
      End
      Begin VB.TextBox txtUserField 
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   120
         Top             =   960
         Width           =   2175
      End
      Begin MSComctlLib.ListView lvwHistory 
         Height          =   2430
         Left            =   105
         TabIndex        =   124
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
         TabIndex        =   114
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
         TabIndex        =   111
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
         TabIndex        =   115
         Top             =   1530
         Width           =   2535
      End
      Begin VB.Label lblXPLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Experience &History"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   123
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
         TabIndex        =   109
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
         TabIndex        =   112
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
         TabIndex        =   108
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblUserField 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Date"
         Height          =   375
         Index           =   2
         Left            =   3360
         TabIndex        =   119
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   10
         Left            =   4200
         TabIndex        =   118
         Tag             =   "?NR"
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Assigned Narrator"
         Height          =   375
         Index           =   10
         Left            =   3360
         TabIndex        =   117
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblModified 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4200
         TabIndex        =   122
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblModifiedLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Modified"
         Height          =   375
         Left            =   3360
         TabIndex        =   121
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
      Begin VB.TextBox txtMemo 
         Height          =   1035
         Index           =   5
         Left            =   4440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   86
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtMemo 
         Height          =   1035
         Index           =   2
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   84
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtMemo 
         Height          =   1035
         Index           =   1
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   82
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2400
         Index           =   11
         ItemData        =   "frmWraithSheet.frx":0B3D
         Left            =   4440
         List            =   "frmWraithSheet.frx":0B3F
         TabIndex        =   98
         Tag             =   "Thorns"
         Top             =   2145
         Width           =   2055
      End
      Begin VB.TextBox txtMemo 
         Height          =   2400
         Index           =   6
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   96
         Top             =   2145
         Width           =   2055
      End
      Begin MSComCtl2.UpDown updTemper 
         Height          =   375
         Index           =   3
         Left            =   1590
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   3840
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   661
         _Version        =   393216
         Value           =   4
         Max             =   1000
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Regret"
         Height          =   255
         Index           =   5
         Left            =   4440
         TabIndex        =   85
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Death"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   83
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Life"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   81
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Thorns"
         Height          =   255
         Index           =   11
         Left            =   4440
         TabIndex        =   97
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dark Passions"
         Height          =   255
         Index           =   6
         Left            =   2280
         TabIndex        =   95
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label lblTemper 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   93
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label lblTemperLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Angst"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   92
         Top             =   3630
         Width           =   2055
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Player"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   90
         Top             =   2910
         Width           =   2055
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   91
         Tag             =   "?PL"
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Archetype"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   88
         Top             =   2190
         Width           =   2055
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   89
         Tag             =   "Shadow Archetypes"
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Line linLine 
         X1              =   6480
         X2              =   120
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label lblShadow 
         BackStyle       =   0  'Transparent
         Caption         =   "Shadow"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   1800
         Width           =   735
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   5
      Left            =   2160
      TabIndex        =   99
      Top             =   1200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.ListBox lstTraits 
         Height          =   1035
         Index           =   15
         ItemData        =   "frmWraithSheet.frx":0B41
         Left            =   120
         List            =   "frmWraithSheet.frx":0B43
         TabIndex        =   101
         Tag             =   "?LO"
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtMemo 
         Height          =   2400
         Index           =   8
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   105
         Top             =   2160
         Width           =   6375
      End
      Begin VB.ListBox lstTraits 
         Height          =   1035
         Index           =   14
         ItemData        =   "frmWraithSheet.frx":0B45
         Left            =   3360
         List            =   "frmWraithSheet.frx":0B47
         TabIndex        =   103
         Tag             =   "?EQ"
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Favorite Locations"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   100
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   104
         Top             =   1920
         Width           =   6375
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Equipment / Relics"
         Height          =   255
         Index           =   14
         Left            =   3360
         TabIndex        =   102
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   3
      Left            =   2160
      TabIndex        =   69
      Top             =   1200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   10
         ItemData        =   "frmWraithSheet.frx":0B49
         Left            =   120
         List            =   "frmWraithSheet.frx":0B4B
         TabIndex        =   71
         Tag             =   "Arcanoi"
         Top             =   480
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   8
         ItemData        =   "frmWraithSheet.frx":0B4D
         Left            =   3360
         List            =   "frmWraithSheet.frx":0B4F
         TabIndex        =   73
         Tag             =   "Status"
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtMemo 
         Height          =   1425
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   75
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox txtMemo 
         Height          =   1425
         Index           =   3
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   77
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox txtMemo 
         Height          =   1425
         Index           =   4
         Left            =   4440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   79
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Arcanoi"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Status Traits"
         Height          =   255
         Index           =   8
         Left            =   3360
         TabIndex        =   72
         Tag             =   "Status Traits|Status Trait"
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Passions"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   74
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fetters"
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   76
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Haunt"
         Height          =   255
         Index           =   4
         Left            =   4440
         TabIndex        =   78
         Top             =   2880
         Width           =   2055
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   2
      Left            =   2160
      TabIndex        =   58
      Top             =   1200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   12
         ItemData        =   "frmWraithSheet.frx":0B51
         Left            =   120
         List            =   "frmWraithSheet.frx":0B53
         TabIndex        =   66
         Tag             =   "Merits"
         Top             =   3120
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   13
         ItemData        =   "frmWraithSheet.frx":0B55
         Left            =   3360
         List            =   "frmWraithSheet.frx":0B57
         TabIndex        =   68
         Tag             =   "Flaws"
         Top             =   3120
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   6
         ItemData        =   "frmWraithSheet.frx":0B59
         Left            =   120
         List            =   "frmWraithSheet.frx":0B5B
         TabIndex        =   60
         Tag             =   "Abilities"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   7
         ItemData        =   "frmWraithSheet.frx":0B5D
         Left            =   2280
         List            =   "frmWraithSheet.frx":0B5F
         TabIndex        =   62
         Tag             =   "Backgrounds"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   9
         ItemData        =   "frmWraithSheet.frx":0B61
         Left            =   4440
         List            =   "frmWraithSheet.frx":0B63
         TabIndex        =   64
         Tag             =   "Influences"
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Merits"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   65
         Top             =   2880
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Flaws"
         Height          =   255
         Index           =   13
         Left            =   3360
         TabIndex        =   67
         Top             =   2880
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Abilities"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   59
         Tag             =   "Abilities|Ability"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Backgrounds"
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   61
         Tag             =   "Backgrounds|Background"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Influences"
         Height          =   255
         Index           =   9
         Left            =   4440
         TabIndex        =   63
         Tag             =   "Influences|Influence"
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdESCClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   127
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
      Caption         =   "<- &Remove"
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
      ItemData        =   "frmWraithSheet.frx":0B65
      Left            =   120
      List            =   "frmWraithSheet.frx":0B67
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
      TabIndex        =   126
      Top             =   120
      Width           =   6855
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   1
      Left            =   2160
      TabIndex        =   45
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
         TabIndex        =   129
         Top             =   2520
         Width           =   255
      End
      Begin VB.CommandButton cmdTraitMax 
         Caption         =   "+"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   128
         Top             =   2520
         Width           =   255
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   0
         ItemData        =   "frmWraithSheet.frx":0B69
         Left            =   120
         List            =   "frmWraithSheet.frx":0B6B
         TabIndex        =   47
         Tag             =   "Physical"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   1
         ItemData        =   "frmWraithSheet.frx":0B6D
         Left            =   2280
         List            =   "frmWraithSheet.frx":0B6F
         TabIndex        =   49
         Tag             =   "Social"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   2
         ItemData        =   "frmWraithSheet.frx":0B71
         Left            =   4440
         List            =   "frmWraithSheet.frx":0B73
         TabIndex        =   51
         Tag             =   "Mental"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   3
         ItemData        =   "frmWraithSheet.frx":0B75
         Left            =   120
         List            =   "frmWraithSheet.frx":0B77
         TabIndex        =   53
         Tag             =   "Physical, Negative"
         Top             =   3120
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   4
         ItemData        =   "frmWraithSheet.frx":0B79
         Left            =   2280
         List            =   "frmWraithSheet.frx":0B7B
         TabIndex        =   55
         Tag             =   "Social, Negative"
         Top             =   3120
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   5
         ItemData        =   "frmWraithSheet.frx":0B7D
         Left            =   4440
         List            =   "frmWraithSheet.frx":0B7F
         TabIndex        =   57
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
         TabIndex        =   132
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblTraitMax 
         Alignment       =   2  'Center
         Caption         =   "Max 11"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   131
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblTraitMax 
         Alignment       =   2  'Center
         Caption         =   "Max 11"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   130
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
         TabIndex        =   46
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
         TabIndex        =   48
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
         TabIndex        =   50
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
         TabIndex        =   52
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
         TabIndex        =   54
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
         TabIndex        =   56
         Tag             =   "Negative Mental"
         Top             =   2880
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
            Caption         =   "Arc Sta Pas Fet"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "LDR Shadow"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Eqp Note"
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
      Picture         =   "frmWraithSheet.frx":0B81
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
      TabIndex        =   106
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
Attribute VB_Name = "frmWraithSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Constants by which specific labels are indexed. (fi = Field Index)
Private Const fiGuild = 0
Private Const fiNature = 1
Private Const fiDemeanor = 2
Private Const fiFaction = 3
Private Const fiStatus = 4
Private Const fiPlayer = 5
Private Const fiLegion = 6
Private Const fiRank = 7
Private Const fiShadowArchetype = 8
Private Const fiShadowPlayer = 9
Private Const fiNarrator = 10

'Constants by which specific temper labels are indexed (pi = Point Index)
Private Const piPathos = 0
Private Const piCorpus = 1
Private Const piWillpower = 2
Private Const piAngst = 3

' Constants by which specific list boxes are indexed. (ti = Trait Index)
Private Const tiPhysical = 0
Private Const tiSocial = 1
Private Const tiMental = 2
Private Const tiPhysicalNeg = 3
Private Const tiSocialNeg = 4
Private Const tiMentalNeg = 5
Private Const tiAbilities = 6
Private Const tiBackgrounds = 7
Private Const tiStatus = 8
Private Const tiInfluences = 9
Private Const tiArcanoi = 10
Private Const tiThorns = 11
Private Const tiMerits = 12
Private Const tiFlaws = 13
Private Const tiEquipment = 14
Private Const tiLocations = 15

' Constants by which specific text boxes are indexed. (xi = Text Index)
Private Const xiName = 0
Private Const xiID = 1
Private Const xiStartDate = 2

' Constants by which specific memo fields are indexed. (mi = Memo Index)
Private Const miPassions = 0
Private Const miLife = 1
Private Const miDeath = 2
Private Const miFetters = 3
Private Const miHaunt = 4
Private Const miRegret = 5
Private Const miDarkPassions = 6
Private Const miNotes = 8

' Constant by which to reference the index of the XP frame and text boxes
Private Const xpFrame = 6
Private Const xpUnspent = 0
Private Const xpEarned = 1

Private Wraith As WraithClass 'The Wraith character
Private CharSheetEngine As CharacterSheetEngineClass    'Handles common functions
Private Populating As Boolean                           'defuses some events when characters are loaded

Public Sub ShowWraith(Character As WraithClass)
'
' Name:         ShowWraith
' Description:  Show and initialize a new instance of the form.
' Arguments:    The WraithClass whose data the form modifies.
'

    Dim DataState As Boolean

    Populating = True

    Set Wraith = Nothing
    Set Wraith = Character
    DataState = Game.DataChanged

    txtUserField(xiName) = Wraith.Name
    Me.Caption = Wraith.Name

    lblField(fiGuild) = Wraith.Guild
    lblField(fiNature) = Wraith.Nature
    lblField(fiDemeanor) = Wraith.Demeanor
    lblField(fiFaction) = Wraith.Faction
    lblField(fiLegion) = Wraith.Legion
    lblField(fiRank) = Wraith.Rank
    lblField(fiShadowArchetype) = Wraith.ShadowArchetype
    lblField(fiShadowPlayer) = Wraith.ShadowPlayer
    lblField(fiStatus) = Wraith.Status
    lblField(fiPlayer) = Wraith.Player
    lblField(fiNarrator) = Wraith.Narrator
        
    lblTraitMax(tiPhysical) = "Max " & CStr(Wraith.PhysicalMax)
    lblTraitMax(tiSocial) = "Max " & CStr(Wraith.SocialMax)
    lblTraitMax(tiMental) = "Max " & CStr(Wraith.MentalMax)

    Call ChangeTemper(piPathos, 0)
    Call ChangeTemper(piCorpus, 0)
    Call ChangeTemper(piWillpower, 0)
    Call ChangeTemper(piAngst, 0)

    txtUserField(xiID) = Wraith.ID
    txtUserField(xiStartDate) = Format(Wraith.StartDate, "mmmm d, yyyy")

    txtMemo(miPassions) = Wraith.Passions
    txtMemo(miLife) = Wraith.Life
    txtMemo(miDeath) = Wraith.Death
    txtMemo(miFetters) = Wraith.Fetters
    txtMemo(miHaunt) = Wraith.Haunt
    txtMemo(miRegret) = Wraith.Regret
    txtMemo(miDarkPassions) = Wraith.DarkPassions
    txtMemo(miNotes) = Wraith.Notes

    CharSheetEngine.RefreshTraitList lstTraits(tiPhysical), Wraith.PhysicalList
    CharSheetEngine.RefreshTraitList lstTraits(tiSocial), Wraith.SocialList
    CharSheetEngine.RefreshTraitList lstTraits(tiMental), Wraith.MentalList
    CharSheetEngine.RefreshTraitList lstTraits(tiPhysicalNeg), Wraith.PhysicalNegList
    CharSheetEngine.RefreshTraitList lstTraits(tiSocialNeg), Wraith.SocialNegList
    CharSheetEngine.RefreshTraitList lstTraits(tiMentalNeg), Wraith.MentalNegList
    CharSheetEngine.RefreshTraitList lstTraits(tiAbilities), Wraith.AbilityList
    CharSheetEngine.RefreshTraitList lstTraits(tiBackgrounds), Wraith.BackgroundList
    CharSheetEngine.RefreshTraitList lstTraits(tiStatus), Wraith.StatusList
    CharSheetEngine.RefreshTraitList lstTraits(tiInfluences), Wraith.InfluenceList

    CharSheetEngine.RefreshTraitList lstTraits(tiArcanoi), Wraith.ArcanoiList
    CharSheetEngine.RefreshTraitList lstTraits(tiThorns), Wraith.ThornList
    CharSheetEngine.RefreshTraitList lstTraits(tiMerits), Wraith.MeritList
    CharSheetEngine.RefreshTraitList lstTraits(tiFlaws), Wraith.FlawList
    CharSheetEngine.RefreshTraitList lstTraits(tiEquipment), Wraith.EquipmentList
    CharSheetEngine.RefreshTraitList lstTraits(tiLocations), Wraith.HangoutList

    cboEthnos.ListIndex = CInt(Wraith.Ethnos)
    lblModified.Caption = Format(Wraith.LastModified, "mmmm d, yyyy")
    chkNPC.Value = IIf(Wraith.IsNPC, vbChecked, vbUnchecked)
        
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
        .SelectSet(osCharacters).Add Wraith.Name
        .GameDate = 0
    End With
    
End Sub

Private Sub cboEthnos_Click()
'
' Name:         cboEthnos_Click
' Description:  Set the wraith to Wraith, Risen or Spectre; change labels and
'               menu identification tags accordingly.
'

    Select Case cboEthnos.ListIndex
        
        Case CInt(etWraith), CInt(etRisen)
            
            tabTabStrip.Tabs.Item(5).Caption = "Mer Flaw Shad"
            lblFieldLabel(fiGuild) = "Guild"
            lblField(fiGuild).Tag = "Guild"
            lblField(fiNature).Tag = "Archetypes"
            lblField(fiDemeanor).Tag = "Archetypes"
            lblTemperLabel(piPathos) = "Pathos"
            lblTemperLabel(piWillpower) = "Willpower"
            lblTraits(tiArcanoi) = "Arcanoi"
            lstTraits(tiArcanoi).Tag = "Arcanoi"
            lblMemo(miPassions) = "Passions"
            If cboEthnos.ListIndex = CInt(etWraith) Then
                lblMemo(miHaunt) = "Haunt"
            Else
                lblMemo(miHaunt) = "Conduit"
            End If
            lblShadow.Caption = "Shadow"
            lblField(fiShadowArchetype).Tag = "Shadow Archetypes"
            lblTemperLabel(piAngst) = "Angst"
            lblMemo(miDarkPassions) = "Dark Passions"
            lblTraits(tiThorns) = "Thorns"
            lstTraits(tiThorns).Tag = "Thorns"
        
        Case CInt(etSpectre)
            tabTabStrip.Tabs.Item(4).Caption = "Pas Fet LDHR"
            tabTabStrip.Tabs.Item(5).Caption = "Mer Flaw Psyc"
            lblFieldLabel(fiGuild) = "Spectre"
            lblField(fiGuild).Tag = "Spectre"
            lblField(fiNature).Tag = "Shadow Archetypes"
            lblField(fiDemeanor).Tag = "Shadow Archetypes"
            lblTemperLabel(piPathos) = "Angst"
            lblTemperLabel(piWillpower) = "Being"
            lblTraits(tiArcanoi) = "Dark Arcanoi"
            lstTraits(tiArcanoi).Tag = "Dark Arcanoi"
            lblMemo(miPassions) = "Dark Passions"
            lblMemo(miHaunt) = "Haunt"
            lblShadow.Caption = "Psyche"
            lblField(fiShadowArchetype).Tag = "Psyche"
            lblTemperLabel(piAngst) = "Composure"
            lblMemo(miDarkPassions) = "Passions"
            lblTraits(tiThorns) = "Fronds"
            lstTraits(tiThorns).Tag = "Fronds"
        
    End Select

    If Not Populating Then
        Wraith.Ethnos = cboEthnos.ListIndex
        SetDataChanged
        mdiMain.AnnounceChanges Me, atCharacters
    End If

End Sub

Private Sub chkNPC_Click()
'
' Name:         chkNPC_Click
' Description:  Toggle whether the character is an NPC.
'

    If Not Populating Then
        Wraith.IsNPC = (chkNPC.Value = vbChecked)
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
            If CharSheetEngine.AddXPEntry(lvwHistory, Wraith.Experience) Then
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
            Wraith.Experience.Clear
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
    frmCalculator.ShowCalculator Wraith

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
        If CharSheetEngine.EditXPEntry(lvwHistory, Wraith.Experience) Then
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
        If CharSheetEngine.RemoveXPEntry(lvwHistory, Wraith.Experience) Then
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
        Wraith.PhysicalMax = Wraith.PhysicalMax + IIf(Index = 1, 1, -1)
        lblTraitMax(tiPhysical).Caption = "Max " & Wraith.PhysicalMax
    End If
    
    If TargetIndex = tiSocial Or Game.LinkTraitMaxes Then
        Wraith.SocialMax = Wraith.SocialMax + IIf(Index = 1, 1, -1)
        lblTraitMax(tiSocial).Caption = "Max " & Wraith.SocialMax
    End If
    
    If TargetIndex = tiMental Or Game.LinkTraitMaxes Then
        Wraith.MentalMax = Wraith.MentalMax + IIf(Index = 1, 1, -1)
        lblTraitMax(tiMental).Caption = "Max " & Wraith.MentalMax
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
    
    lblModified.Caption = Format(Wraith.LastModified, "mmmm d, yyyy")
    lblShowXP = " " & CStr(Wraith.Experience.Unspent) & _
            " / " & CStr(Wraith.Experience.Earned)

    CharSheetEngine.RefreshTraitList lstTraits(tiEquipment), Wraith.EquipmentList
    CharSheetEngine.RefreshTraitList lstTraits(tiLocations), Wraith.HangoutList
    lblField(fiPlayer).Caption = Wraith.Player
    lblField(fiNarrator).Caption = Wraith.Narrator
    lblField(fiShadowPlayer).Caption = Wraith.ShadowPlayer
    
    If mdiMain.CheckForChanges(Me, atTempers) Then
        Call ChangeTemper(piCorpus, 0)
        Call ChangeTemper(piAngst, 0)
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
'               when ShowWraith is the means of loading the form.  Initializes the
'               MenuStack linked list and the Wraith Menus.
'

    If Wraith Is Nothing Then
        MsgBox "Character sheet loaded improperly!"
    Else
        
        Set CharSheetEngine = New CharacterSheetEngineClass
        
        CharSheetEngine.RegisterSheet "Wraith", lstMenu, lblMenuItem, lblMenuTitle
                
        CharSheetEngine.RegisterTraitList tiPhysical, Wraith.PhysicalList
        CharSheetEngine.RegisterTraitList tiSocial, Wraith.SocialList
        CharSheetEngine.RegisterTraitList tiMental, Wraith.MentalList
        CharSheetEngine.RegisterTraitList tiPhysicalNeg, Wraith.PhysicalNegList
        CharSheetEngine.RegisterTraitList tiSocialNeg, Wraith.SocialNegList
        CharSheetEngine.RegisterTraitList tiMentalNeg, Wraith.MentalNegList
        CharSheetEngine.RegisterTraitList tiAbilities, Wraith.AbilityList
        CharSheetEngine.RegisterTraitList tiBackgrounds, Wraith.BackgroundList
        CharSheetEngine.RegisterTraitList tiStatus, Wraith.StatusList
        CharSheetEngine.RegisterTraitList tiInfluences, Wraith.InfluenceList
        
        CharSheetEngine.RegisterTraitList tiArcanoi, Wraith.ArcanoiList
        CharSheetEngine.RegisterTraitList tiThorns, Wraith.ThornList
        CharSheetEngine.RegisterTraitList tiMerits, Wraith.MeritList
        CharSheetEngine.RegisterTraitList tiFlaws, Wraith.FlawList
        CharSheetEngine.RegisterTraitList tiEquipment, Wraith.EquipmentList
        CharSheetEngine.RegisterTraitList tiLocations, Wraith.HangoutList
    
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
            Case fiGuild
                Wraith.Guild = Value
                mdiMain.AnnounceChanges Me, atCharacters
            Case fiNature
                Wraith.Nature = Value
            Case fiDemeanor
                Wraith.Demeanor = Value
            Case fiFaction
                Wraith.Faction = Value
            Case fiPlayer
                Wraith.Player = Value
            Case fiStatus
                Wraith.Status = Value
                mdiMain.AnnounceChanges Me, atCharacters
            Case fiLegion
                Wraith.Legion = Value
            Case fiRank
                Wraith.Rank = Value
            Case fiShadowArchetype
                Wraith.ShadowArchetype = Value
            Case fiShadowPlayer
                Wraith.ShadowPlayer = Value
            Case fiNarrator
                Wraith.Narrator = Value
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
            Case piWillpower:   lblTemperFloat = " " & CStr(Wraith.TempWillpower) & " "
            Case piCorpus:      lblTemperFloat = " " & CStr(Wraith.TempCorpus) & " "
            Case piPathos:      lblTemperFloat = " " & CStr(Wraith.TempPathos) & " "
            Case piAngst:       lblTemperFloat = " " & CStr(Wraith.TempAngst) & " "
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
                Wraith.Experience.Unspent = CSng(txtExperience(xpUnspent))
            Case xpEarned
                Wraith.Experience.Earned = CSng(txtExperience(xpEarned))
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
        Case miPassions
            If Wraith.Passions <> txtMemo(miPassions).Text Then
                SetDataChanged
                Wraith.Passions = TrimWhiteSpace(txtMemo(miPassions))
            End If
        Case miLife
            If Wraith.Life <> txtMemo(miLife).Text Then
                SetDataChanged
                Wraith.Life = TrimWhiteSpace(txtMemo(miLife))
            End If
        Case miDeath
            If Wraith.Death <> txtMemo(miDeath).Text Then
                SetDataChanged
                Wraith.Death = TrimWhiteSpace(txtMemo(miDeath))
            End If
        Case miHaunt
            If Wraith.Haunt <> txtMemo(miHaunt).Text Then
                SetDataChanged
                Wraith.Haunt = TrimWhiteSpace(txtMemo(miHaunt))
            End If
        Case miRegret
            If Wraith.Regret <> txtMemo(miRegret).Text Then
                SetDataChanged
                Wraith.Regret = TrimWhiteSpace(txtMemo(miRegret))
            End If
        Case miFetters
            If Wraith.Fetters <> txtMemo(miFetters).Text Then
                SetDataChanged
                Wraith.Fetters = TrimWhiteSpace(txtMemo(miFetters))
            End If
        Case miDarkPassions
            If Wraith.DarkPassions <> txtMemo(miDarkPassions).Text Then
                SetDataChanged
                Wraith.DarkPassions = TrimWhiteSpace(txtMemo(miDarkPassions))
            End If
        Case miNotes
            If Wraith.Notes <> txtMemo(miNotes).Text Then
                SetDataChanged
                Wraith.Notes = TrimWhiteSpace(txtMemo(miNotes))
            End If
    End Select

End Sub

Private Sub RefreshXP()
'
' Name:         RefreshXP
' Description:  Refresh the display of XP values and histories.
'

    Populating = True
    txtExperience(xpUnspent).Text = CStr(Wraith.Experience.Unspent)
    txtExperience(xpEarned).Text = CStr(Wraith.Experience.Earned)
    txtExperience(xpUnspent).Locked = Game.EnforceHistory
    txtExperience(xpEarned).Locked = Game.EnforceHistory
    updExperience(xpUnspent).Visible = Not Game.EnforceHistory
    updExperience(xpEarned).Visible = Not Game.EnforceHistory
    lblShowXP = " " & CStr(Wraith.Experience.Unspent) & _
            " / " & CStr(Wraith.Experience.Earned)
    Populating = False
        
    CharSheetEngine.RefreshXP lvwHistory, Wraith.Experience
    
End Sub

Private Sub SetDataChanged()
'
' Name:         SetDataChanged
' Description:  Tell the game its data has changed and update the character's
'               Last Modified date.
'
        
    If Not Populating Then
        Game.DataChanged = True
        Wraith.LastModified = Now
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
                Wraith.ID = Trim(txtUserField(Index))
            Case xiStartDate
                If IsDate(txtUserField(Index)) Then Wraith.StartDate = CDate(txtUserField(Index))
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

    With Wraith
        Select Case Index
            Case piWillpower
                .TempWillpower = .TempWillpower + Change
                If Both Then .Willpower = .Willpower + Change
                If .TempWillpower < 0 Then .TempWillpower = 0
                If .Willpower < 0 Then .Willpower = 0
                Temp = .TempWillpower
                Perm = .Willpower
            Case piCorpus
                .TempCorpus = .TempCorpus + Change
                If Both Then .Corpus = .Corpus + Change
                If .TempCorpus < 0 Then .TempCorpus = 0
                If .Corpus < 0 Then .Corpus = 0
                Temp = .TempCorpus
                Perm = .Corpus
            Case piPathos
                .TempPathos = .TempPathos + Change
                If Both Then .Pathos = .Pathos + Change
                If .TempPathos < 0 Then .TempPathos = 0
                If .Pathos < 0 Then .Pathos = 0
                Temp = .TempPathos
                Perm = .Pathos
            Case piAngst
                .TempAngst = .TempAngst + Change
                If Both Then .Angst = .Angst + Change
                If .TempAngst < 0 Then .TempAngst = 0
                If .Angst < 0 Then .Angst = 0
                Temp = .TempAngst
                Perm = .Angst
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

