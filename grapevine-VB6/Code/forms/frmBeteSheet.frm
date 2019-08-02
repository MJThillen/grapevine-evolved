VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frmBeteSheet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bete Character"
   ClientHeight    =   6165
   ClientLeft      =   210
   ClientTop       =   405
   ClientWidth     =   9060
   Icon            =   "frmBeteSheet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   9060
   Tag             =   "C"
   Begin VB.CommandButton cmdESCClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdCustom 
      Caption         =   "Custom &Entry"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdNote 
      Caption         =   "Add &Note to Entry"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<- Re&move"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add ->"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   1695
   End
   Begin VB.ListBox lstMenu 
      Height          =   2010
      ItemData        =   "frmBeteSheet.frx":08CA
      Left            =   120
      List            =   "frmBeteSheet.frx":08CC
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
      TabIndex        =   1
      Top             =   120
      Width           =   6855
   End
   Begin TabDlg.SSTab tabTabs 
      Height          =   5175
      Left            =   2040
      TabIndex        =   0
      Top             =   840
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Identity"
      TabPicture(0)   =   "frmBeteSheet.frx":08CE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblField(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblFieldLabel(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblField(6)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblFieldLabel(6)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblField(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblFieldLabel(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblField(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblFieldLabel(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblTemper(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblTemperLabel(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblTemper(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblTemperLabel(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblFieldLabel(8)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblFieldLabel(7)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblField(8)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblField(7)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblFieldLabel(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblField(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblFieldLabel(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblField(2)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblTemperLabel(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblTemper(0)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblUserField(1)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblExperienceLabel"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lblUserField(2)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lblFieldLabel(5)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblField(5)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "updExperience"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "updTemper(0)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "updTemper(2)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "updTemper(1)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtUserField(1)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtExperience"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "chkNPC"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtUserField(2)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cmdChange"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).ControlCount=   36
      TabCaption(1)   =   "Phy. Soc. Men."
      TabPicture(1)   =   "frmBeteSheet.frx":08EA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAscend"
      Tab(1).Control(1)=   "cmdDescend"
      Tab(1).Control(2)=   "cmdDecrement"
      Tab(1).Control(3)=   "cmdIncrement"
      Tab(1).Control(4)=   "lstTraits(5)"
      Tab(1).Control(5)=   "lstTraits(4)"
      Tab(1).Control(6)=   "lstTraits(3)"
      Tab(1).Control(7)=   "lstTraits(2)"
      Tab(1).Control(8)=   "lstTraits(1)"
      Tab(1).Control(9)=   "lstTraits(0)"
      Tab(1).Control(10)=   "lblTraits(5)"
      Tab(1).Control(11)=   "lblTraits(4)"
      Tab(1).Control(12)=   "lblTraits(3)"
      Tab(1).Control(13)=   "lblTraits(2)"
      Tab(1).Control(14)=   "lblTraits(1)"
      Tab(1).Control(15)=   "lblTraits(0)"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Ab. Inf. Bac. Mer. Fl."
      TabPicture(2)   =   "frmBeteSheet.frx":0906
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstTraits(8)"
      Tab(2).Control(1)=   "lstTraits(14)"
      Tab(2).Control(2)=   "lstTraits(13)"
      Tab(2).Control(3)=   "lstTraits(7)"
      Tab(2).Control(4)=   "lstTraits(6)"
      Tab(2).Control(5)=   "lblTraits(8)"
      Tab(2).Control(6)=   "lblTraits(14)"
      Tab(2).Control(7)=   "lblTraits(13)"
      Tab(2).Control(8)=   "lblTraits(7)"
      Tab(2).Control(9)=   "lblTraits(6)"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "HL OF Gifts Rites"
      TabPicture(3)   =   "frmBeteSheet.frx":0922
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblTraits(12)"
      Tab(3).Control(1)=   "lblTraits(15)"
      Tab(3).Control(2)=   "lblTraits(16)"
      Tab(3).Control(3)=   "lblTraits(17)"
      Tab(3).Control(4)=   "lstTraits(15)"
      Tab(3).Control(5)=   "lstTraits(16)"
      Tab(3).Control(6)=   "lstTraits(12)"
      Tab(3).Control(7)=   "lstTraits(17)"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "Renown"
      TabPicture(4)   =   "frmBeteSheet.frx":093E
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblNumber(0)"
      Tab(4).Control(1)=   "lblNumberLabel(0)"
      Tab(4).Control(2)=   "lblTraits(11)"
      Tab(4).Control(3)=   "lblTraits(10)"
      Tab(4).Control(4)=   "lblTraits(9)"
      Tab(4).Control(5)=   "lblDecimal(0)"
      Tab(4).Control(6)=   "lblDecimal(1)"
      Tab(4).Control(7)=   "lblDecimal(2)"
      Tab(4).Control(8)=   "lblFieldLabel(10)"
      Tab(4).Control(9)=   "lblField(10)"
      Tab(4).Control(10)=   "updDecimal(2)"
      Tab(4).Control(11)=   "updDecimal(1)"
      Tab(4).Control(12)=   "updDecimal(0)"
      Tab(4).Control(13)=   "updNumber(0)"
      Tab(4).Control(14)=   "lstTraits(11)"
      Tab(4).Control(15)=   "lstTraits(10)"
      Tab(4).Control(16)=   "lstTraits(9)"
      Tab(4).Control(17)=   "txtDecimal(0)"
      Tab(4).Control(18)=   "txtDecimal(1)"
      Tab(4).Control(19)=   "txtDecimal(2)"
      Tab(4).ControlCount=   20
      TabCaption(5)   =   "Eq. Not."
      TabPicture(5)   =   "frmBeteSheet.frx":095A
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblModifiedLabel"
      Tab(5).Control(1)=   "lblModified"
      Tab(5).Control(2)=   "lblTraits(18)"
      Tab(5).Control(3)=   "lblFieldLabel(9)"
      Tab(5).Control(4)=   "lblField(9)"
      Tab(5).Control(5)=   "lblUserField(3)"
      Tab(5).Control(6)=   "lblMemo(0)"
      Tab(5).Control(7)=   "lstTraits(18)"
      Tab(5).Control(8)=   "txtUserField(3)"
      Tab(5).Control(9)=   "txtMemo(0)"
      Tab(5).ControlCount=   10
      Begin VB.CommandButton cmdChange 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   21.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5895
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   2280
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtDecimal 
         Height          =   375
         Index           =   2
         Left            =   -73920
         TabIndex        =   105
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox txtDecimal 
         Height          =   375
         Index           =   1
         Left            =   -73920
         TabIndex        =   102
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox txtDecimal 
         Height          =   375
         Index           =   0
         Left            =   -73920
         TabIndex        =   99
         Top             =   3240
         Width           =   1575
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   9
         ItemData        =   "frmBeteSheet.frx":0976
         Left            =   -74760
         List            =   "frmBeteSheet.frx":0978
         TabIndex        =   92
         Tag             =   "Honor"
         Top             =   840
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   10
         ItemData        =   "frmBeteSheet.frx":097A
         Left            =   -72600
         List            =   "frmBeteSheet.frx":097C
         TabIndex        =   91
         Tag             =   "Glory"
         Top             =   840
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   11
         ItemData        =   "frmBeteSheet.frx":097E
         Left            =   -70440
         List            =   "frmBeteSheet.frx":0980
         TabIndex        =   90
         Tag             =   "Wisdom"
         Top             =   840
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   1035
         Index           =   17
         ItemData        =   "frmBeteSheet.frx":0982
         Left            =   -71520
         List            =   "frmBeteSheet.frx":0984
         TabIndex        =   86
         Tag             =   "Features"
         Top             =   840
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   1035
         Index           =   12
         ItemData        =   "frmBeteSheet.frx":0986
         Left            =   -74760
         List            =   "frmBeteSheet.frx":0988
         TabIndex        =   85
         Tag             =   "Health Levels"
         Top             =   840
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   2400
         Index           =   16
         ItemData        =   "frmBeteSheet.frx":098A
         Left            =   -71520
         List            =   "frmBeteSheet.frx":098C
         TabIndex        =   84
         Tag             =   "Rites"
         Top             =   2520
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   2400
         Index           =   15
         ItemData        =   "frmBeteSheet.frx":098E
         Left            =   -74760
         List            =   "frmBeteSheet.frx":0990
         TabIndex        =   83
         Tag             =   "Gifts"
         Top             =   2520
         Width           =   3135
      End
      Begin VB.TextBox txtMemo 
         Height          =   4065
         Index           =   0
         Left            =   -71520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   77
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtUserField 
         Height          =   375
         Index           =   3
         Left            =   -73800
         TabIndex        =   76
         Top             =   4050
         Width           =   2175
      End
      Begin VB.ListBox lstTraits 
         Height          =   2595
         Index           =   18
         ItemData        =   "frmBeteSheet.frx":0992
         Left            =   -74760
         List            =   "frmBeteSheet.frx":0994
         TabIndex        =   72
         Tag             =   "Equipment"
         Top             =   840
         Width           =   3135
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
         Left            =   -72960
         TabIndex        =   71
         Top             =   2880
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
         Left            =   -74640
         TabIndex        =   70
         Top             =   2880
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
         Left            =   -74640
         TabIndex        =   69
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdIncrement 
         Caption         =   "+"
         Height          =   255
         Left            =   -72960
         TabIndex        =   68
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtUserField 
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   8
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CheckBox chkNPC 
         Alignment       =   1  'Right Justify
         Caption         =   "NPC"
         Height          =   375
         Left            =   3840
         TabIndex        =   11
         Top             =   2760
         Width           =   660
      End
      Begin VB.TextBox txtExperience 
         Height          =   375
         Left            =   4320
         TabIndex        =   9
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtUserField 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   7
         Top             =   4200
         Width           =   2175
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   8
         ItemData        =   "frmBeteSheet.frx":0996
         Left            =   -70440
         List            =   "frmBeteSheet.frx":0998
         TabIndex        =   23
         Tag             =   "Backgrounds"
         Top             =   840
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   14
         ItemData        =   "frmBeteSheet.frx":099A
         Left            =   -71520
         List            =   "frmBeteSheet.frx":099C
         TabIndex        =   25
         Tag             =   "Flaws"
         Top             =   3480
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   13
         ItemData        =   "frmBeteSheet.frx":099E
         Left            =   -74760
         List            =   "frmBeteSheet.frx":09A0
         TabIndex        =   24
         Tag             =   "Merits"
         Top             =   3480
         Width           =   3135
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   7
         ItemData        =   "frmBeteSheet.frx":09A2
         Left            =   -72600
         List            =   "frmBeteSheet.frx":09A4
         TabIndex        =   22
         Tag             =   "Influences"
         Top             =   840
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   6
         ItemData        =   "frmBeteSheet.frx":09A6
         Left            =   -74760
         List            =   "frmBeteSheet.frx":09A8
         TabIndex        =   21
         Tag             =   "Abilities"
         Top             =   840
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   5
         ItemData        =   "frmBeteSheet.frx":09AA
         Left            =   -70440
         List            =   "frmBeteSheet.frx":09AC
         TabIndex        =   20
         Tag             =   "Mental, Negative"
         Top             =   3480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   4
         ItemData        =   "frmBeteSheet.frx":09AE
         Left            =   -72600
         List            =   "frmBeteSheet.frx":09B0
         TabIndex        =   19
         Tag             =   "Social, Negative"
         Top             =   3480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   1425
         Index           =   3
         ItemData        =   "frmBeteSheet.frx":09B2
         Left            =   -74760
         List            =   "frmBeteSheet.frx":09B4
         TabIndex        =   18
         Tag             =   "Physical, Negative"
         Top             =   3480
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   2
         ItemData        =   "frmBeteSheet.frx":09B6
         Left            =   -70440
         List            =   "frmBeteSheet.frx":09B8
         TabIndex        =   17
         Tag             =   "Mental"
         Top             =   840
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   1
         ItemData        =   "frmBeteSheet.frx":09BA
         Left            =   -72600
         List            =   "frmBeteSheet.frx":09BC
         TabIndex        =   16
         Tag             =   "Social"
         Top             =   840
         Width           =   2055
      End
      Begin VB.ListBox lstTraits 
         Height          =   2010
         Index           =   0
         ItemData        =   "frmBeteSheet.frx":09BE
         Left            =   -74760
         List            =   "frmBeteSheet.frx":09C0
         TabIndex        =   15
         Tag             =   "Physical"
         Top             =   840
         Width           =   2055
      End
      Begin ComCtl2.UpDown updTemper 
         Height          =   315
         Index           =   1
         Left            =   5895
         TabIndex        =   13
         Top             =   3735
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   327681
         Value           =   3
         Max             =   20
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown updTemper 
         Height          =   315
         Index           =   2
         Left            =   5895
         TabIndex        =   14
         Top             =   4215
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   327681
         Value           =   1
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown updTemper 
         Height          =   315
         Index           =   0
         Left            =   5895
         TabIndex        =   12
         Top             =   3255
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   327681
         Value           =   3
         Max             =   20
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown updExperience 
         Height          =   315
         Left            =   5895
         TabIndex        =   10
         Top             =   2295
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   327681
         OrigLeft        =   6015
         OrigTop         =   1815
         OrigRight       =   6600
         OrigBottom      =   2130
         Max             =   999
         Min             =   -999
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown updNumber 
         Height          =   315
         Index           =   0
         Left            =   -68985
         TabIndex        =   93
         Top             =   3735
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   327681
         Value           =   6
         OrigLeft        =   2535
         OrigTop         =   1335
         OrigRight       =   3120
         OrigBottom      =   1650
         Max             =   100
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown updDecimal 
         Height          =   315
         Index           =   0
         Left            =   -72345
         TabIndex        =   100
         Top             =   3255
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   327681
         OrigLeft        =   6015
         OrigTop         =   1815
         OrigRight       =   6600
         OrigBottom      =   2130
         Max             =   1000
         Min             =   -1000
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown updDecimal 
         Height          =   315
         Index           =   1
         Left            =   -72345
         TabIndex        =   103
         Top             =   3735
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   327681
         OrigLeft        =   6015
         OrigTop         =   1815
         OrigRight       =   6600
         OrigBottom      =   2130
         Max             =   1000
         Min             =   -1000
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown updDecimal 
         Height          =   315
         Index           =   2
         Left            =   -72345
         TabIndex        =   106
         Top             =   4215
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   327681
         OrigLeft        =   6015
         OrigTop         =   1815
         OrigRight       =   6600
         OrigBottom      =   2130
         Max             =   1000
         Min             =   -1000
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   10
         Left            =   -70560
         TabIndex        =   109
         Tag             =   "Totems"
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Totem"
         Height          =   255
         Index           =   10
         Left            =   -71520
         TabIndex        =   108
         Top             =   3270
         Width           =   855
      End
      Begin VB.Label lblDecimal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Wisdom, Cleverness"
         Height          =   375
         Index           =   2
         Left            =   -74880
         TabIndex        =   107
         Top             =   4230
         Width           =   855
      End
      Begin VB.Label lblDecimal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Glory, Fer., Suc., Cun."
         Height          =   375
         Index           =   1
         Left            =   -74880
         TabIndex        =   104
         Top             =   3750
         Width           =   855
      End
      Begin VB.Label lblDecimal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Honor, Humor, Ob."
         Height          =   375
         Index           =   0
         Left            =   -74880
         TabIndex        =   101
         Top             =   3270
         Width           =   855
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Honor/Hum./Ob."
         Height          =   255
         Index           =   9
         Left            =   -74760
         TabIndex        =   98
         Tag             =   "Honor/Hum./Ob."
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Glory/F/S/C"
         Height          =   255
         Index           =   10
         Left            =   -72600
         TabIndex        =   97
         Tag             =   "Glory/F/S/C"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Wisdom/Cleverness"
         Height          =   255
         Index           =   11
         Left            =   -70440
         TabIndex        =   96
         Tag             =   "Wisdom/Cleverness"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblNumberLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Spirit Notoriety"
         Height          =   495
         Index           =   0
         Left            =   -71640
         TabIndex        =   95
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label lblNumber 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   -70560
         TabIndex        =   94
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Other Features"
         Height          =   255
         Index           =   17
         Left            =   -71520
         TabIndex        =   89
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rites"
         Height          =   255
         Index           =   16
         Left            =   -71520
         TabIndex        =   88
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gifts"
         Height          =   255
         Index           =   15
         Left            =   -74760
         TabIndex        =   87
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Health Levels"
         Height          =   255
         Index           =   12
         Left            =   -74760
         TabIndex        =   82
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         Height          =   255
         Index           =   0
         Left            =   -71520
         TabIndex        =   81
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label lblUserField 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Date"
         Height          =   375
         Index           =   3
         Left            =   -74640
         TabIndex        =   80
         Top             =   4050
         Width           =   735
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   9
         Left            =   -73800
         TabIndex        =   79
         Tag             =   "Narrators"
         Top             =   3570
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Assigned Narrator"
         Height          =   375
         Index           =   9
         Left            =   -74640
         TabIndex        =   78
         Top             =   3570
         Width           =   735
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Equipment / Fetishes"
         Height          =   255
         Index           =   18
         Left            =   -74760
         TabIndex        =   75
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label lblModified 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -73800
         TabIndex        =   74
         Top             =   4530
         Width           =   2175
      End
      Begin VB.Label lblModifiedLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Modified"
         Height          =   375
         Left            =   -74640
         TabIndex        =   73
         Top             =   4530
         Width           =   735
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   5
         Left            =   1080
         TabIndex        =   67
         Tag             =   "Rank"
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rank"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   66
         Top             =   3270
         Width           =   855
      End
      Begin VB.Label lblUserField 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Character ID"
         Height          =   255
         Index           =   2
         Left            =   3315
         TabIndex        =   64
         Top             =   1365
         Width           =   900
      End
      Begin VB.Label lblExperienceLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XP Unspent / Earned"
         Height          =   375
         Left            =   3360
         TabIndex        =   63
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label lblUserField 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pack"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   40
         Top             =   4245
         Width           =   855
      End
      Begin VB.Label lblTemper 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   4320
         TabIndex        =   46
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label lblTemperLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rage/Blood"
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   45
         Top             =   3270
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   1080
         TabIndex        =   33
         Tag             =   "Breed"
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Breed"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   1830
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   31
         Tag             =   "Auspice"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Auspice/ Pryio "
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Backgrounds"
         Height          =   255
         Index           =   8
         Left            =   -70440
         TabIndex        =   59
         Tag             =   "Backgrounds|Background"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   7
         Left            =   4320
         TabIndex        =   42
         Tag             =   "Players"
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   8
         Left            =   4320
         TabIndex        =   44
         Tag             =   "Status, Character"
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Player"
         Height          =   255
         Index           =   7
         Left            =   3600
         TabIndex        =   41
         Top             =   870
         Width           =   615
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   255
         Index           =   8
         Left            =   3600
         TabIndex        =   43
         Top             =   1830
         Width           =   615
      End
      Begin VB.Label lblTemperLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gnosis"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   47
         Top             =   3750
         Width           =   855
      End
      Begin VB.Label lblTemper 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   48
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label lblTemperLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Willpower"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   49
         Top             =   4230
         Width           =   855
      End
      Begin VB.Label lblTemper 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   50
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Demeanor"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   36
         Top             =   2790
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   4
         Left            =   1080
         TabIndex        =   37
         Tag             =   "Archetypes"
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nature"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   34
         Top             =   2310
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   3
         Left            =   1080
         TabIndex        =   35
         Tag             =   "Archetypes"
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   38
         Top             =   3750
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   6
         Left            =   1080
         TabIndex        =   39
         Tag             =   "Position"
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bete"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   870
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   29
         Tag             =   "Bete"
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Flaws"
         Height          =   255
         Index           =   14
         Left            =   -71520
         TabIndex        =   61
         Top             =   3240
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Merits"
         Height          =   255
         Index           =   13
         Left            =   -74760
         TabIndex        =   60
         Top             =   3240
         Width           =   3135
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Influences"
         Height          =   255
         Index           =   7
         Left            =   -72600
         TabIndex        =   58
         Tag             =   "Influences|Influence"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Abilities"
         Height          =   255
         Index           =   6
         Left            =   -74760
         TabIndex        =   57
         Tag             =   "Abilities|Ability"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Negative Mental"
         Height          =   255
         Index           =   5
         Left            =   -70440
         TabIndex        =   56
         Tag             =   "Negative Mental"
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Negative Social"
         Height          =   255
         Index           =   4
         Left            =   -72600
         TabIndex        =   55
         Tag             =   "Negative Social"
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Negative Physical"
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   54
         Tag             =   "Negative Physical"
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Mental"
         Height          =   255
         Index           =   2
         Left            =   -70440
         TabIndex        =   53
         Tag             =   "Mental"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Social"
         Height          =   255
         Index           =   1
         Left            =   -72600
         TabIndex        =   52
         Tag             =   "Social"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Physical"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   51
         Tag             =   "Physical"
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   480
      Picture         =   "frmBeteSheet.frx":09C2
      Top             =   185
      Width           =   480
   End
   Begin VB.Label lblMenuItem 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   120
      TabIndex        =   27
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
      TabIndex        =   26
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblMenuTitle 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   62
      Top             =   915
      Width           =   1695
   End
End
Attribute VB_Name = "frmBeteSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Constants by which specific labels are indexed. (fi = Field Index)
Private Const fiBete = 0
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

' Constants by which specific text boxes are indexed. (xi = Text Index)
Private Const xiName = 0
Private Const xiPack = 1
Private Const xiID = 2
Private Const xiStartDate = 3

' Constants by which specific decimal fields are indexed. (di = Memo Index)
Private Const diHonor = 0
Private Const diGlory = 1
Private Const diWisdom = 2

' Constants by which specific memo fields are indexed. (mi = Memo Index)
Private Const miNotes = 0

Private LocalTargetType As TargetClassType  'The type of control that has focus
Private LocalTargetLabel As Label           'The label that has focus
Private LocalTargetList As ListBox          'The list box that has focus

Private Bete As BeteClass                   'The Bete character
Private MenuStack As LinkedList             'Menu stack
Private SubMenuLabelStack As LinkedList     'Submenu Label stack
Private Populating As Boolean  'defuses some events when characters are loaded

Public Sub ShowBete(Character As BeteClass)
'
' Name:         ShowBete
' Description:  Show and initialize a new instance of the form.
' Arguments:    The BeteClass whose data the form modifies.
'

    Dim DataState As Boolean

    Populating = True

    Set Bete = Nothing
    Set Bete = Character
    DataState = Game.DataChanged

    txtUserField(xiName) = Bete.Name
    Me.Caption = Bete.Name
    
    lblField(fiBete) = Bete.Bete
    lblField(fiAuspice) = Bete.Auspice
    lblField(fiBreed) = Bete.Breed
    lblField(fiPosition) = Bete.Position
    lblField(fiNature) = Bete.Nature
    lblField(fiDemeanor) = Bete.Demeanor
    lblField(fiRank) = Bete.Rank
    lblField(fiTotem) = Bete.Totem
    lblField(fiStatus) = Bete.Status
    lblField(fiPlayer) = Bete.Player
    lblField(fiNarrator) = Bete.Narrator
    
    txtUserField(xiID) = Bete.ID
    txtUserField(xiPack) = Bete.Pack
    txtUserField(xiStartDate) = Bete.StartDate
    
    txtMemo(miNotes) = Bete.Notes

    updTemper(piRage).Value = Bete.Rage
    updTemper(piGnosis).Value = Bete.Gnosis
    updTemper(piWillpower).Value = Bete.Willpower
    
    CharSheetEngine.RefreshTraitList lstTraits(tiPhysical), Bete.PhysicalList
    CharSheetEngine.RefreshTraitList lstTraits(tiSocial), Bete.SocialList
    CharSheetEngine.RefreshTraitList lstTraits(tiMental), Bete.MentalList
    CharSheetEngine.RefreshTraitList lstTraits(tiPhysicalNeg), Bete.PhysicalNegList
    CharSheetEngine.RefreshTraitList lstTraits(tiSocialNeg), Bete.SocialNegList
    CharSheetEngine.RefreshTraitList lstTraits(tiMentalNeg), Bete.MentalNegList
    CharSheetEngine.RefreshTraitList lstTraits(tiAbilities), Bete.AbilityList
    CharSheetEngine.RefreshTraitList lstTraits(tiInfluences), Bete.InfluenceList
    CharSheetEngine.RefreshTraitList lstTraits(tiBackgrounds), Bete.BackgroundList
    CharSheetEngine.RefreshTraitList lstTraits(tiHonor), Bete.HonorList
    CharSheetEngine.RefreshTraitList lstTraits(tiGlory), Bete.GloryList
    CharSheetEngine.RefreshTraitList lstTraits(tiWisdom), Bete.WisdomList
    CharSheetEngine.RefreshTraitList lstTraits(tiHealth), Bete.HealthList
    
    CharSheetEngine.RefreshTraitList lstTraits(tiMerits), Bete.MeritList
    CharSheetEngine.RefreshTraitList lstTraits(tiFlaws), Bete.FlawList
    CharSheetEngine.RefreshTraitList lstTraits(tiGifts), Bete.GiftList
    CharSheetEngine.RefreshTraitList lstTraits(tiRites), Bete.RiteList
    CharSheetEngine.RefreshTraitList lstTraits(tiFeatures), Bete.FeatureList
    CharSheetEngine.RefreshTraitList lstTraits(tiEquipment), Bete.EquipmentList
    
    updNumber(niNotoriety).Value = Bete.Notoriety
    
    txtDecimal(diHonor).Text = " " & CStr(Bete.Honor)
    txtDecimal(diGlory).Text = " " & CStr(Bete.Glory)
    txtDecimal(diWisdom).Text = " " & CStr(Bete.Wisdom)
    
    lblModified.Caption = Format(Bete.LastModified, "mmmm d, yyyy")
    chkNPC.Value = IIf(Bete.IsNPC, vbChecked, vbUnchecked)
    
    Me.Show
    
    Game.DataChanged = DataState
    Populating = False

End Sub

Private Sub chkNPC_Click()
'
' Name:         chkNPC_Click
' Description:  Toggle whether the character is an NPC.
'

    If Not Populating Then
        Bete.IsNPC = (chkNPC.Value = vbChecked)
        EntryChangedChars = True
        SetDataChanged
    End If
    
End Sub

Private Sub cmdAdd_Click()
'
' Name:         cmdAdd_Click
' Description:  If a selection is active, have the CharSheetEngine add to
'               the menu.
'

    If lstMenu.ListIndex <> -1 Then
    
        CharSheetEngine.AddSelected

    End If

End Sub

Private Sub cmdCustom_Click()
'
' Name:         cmdCustom_Click
' Description:  Have the CharSheetEngine add a custom entry to the target.
'

    CharSheetEngine.AddCustom

End Sub

Private Sub cmdAscend_Click()
'
' Name:         cmdAscend_Click
' Description:  Move the selected entry down.
'

    If cmdAscend.Visible Then
        CharSheetEngine.MoveEntryUp
        CharSheetEngine.TargetList.SetFocus
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
    End If
    
End Sub

Private Sub cmdIncrement_Click()
'
' Name:         cmdIncrement_Click
' Description:  Increment the selected entry.
'

    If cmdIncrement.Visible Then
        CharSheetEngine.IncrementEntry
        CharSheetEngine.TargetList.SetFocus
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
    
    CharSheetEngine.AddNote

End Sub

Private Sub cmdRemove_Click()
'
' Name:         cmdRemove_Click
' Description:  Have the CharSheetEngine remove a label or list entry.
'
    
    CharSheetEngine.RemoveEntry
    
End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  Update the experience total in case it changed elsewhere.
'
    Dim DataState
    
    DataState = Game.DataChanged
    Populating = True
    txtExperience = " " & CStr(Bete.Experience.Unspent) & _
            " / " & CStr(Bete.Experience.Earned)
    txtExperience.Locked = Game.EnforceHistory
    cmdChange.Visible = Game.EnforceHistory
    updExperience.Visible = Not Game.EnforceHistory
    Populating = False
    Game.DataChanged = DataState

    CharSheetEngine.RegisterSheet "Bete", MenuStack, lstMenu, lblMenuItem, _
        lblMenuTitle, SubMenuLabelStack

    CharSheetEngine.RegisterTraitList tiPhysical, Bete.PhysicalList
    CharSheetEngine.RegisterTraitList tiSocial, Bete.SocialList
    CharSheetEngine.RegisterTraitList tiMental, Bete.MentalList
    CharSheetEngine.RegisterTraitList tiPhysicalNeg, Bete.PhysicalNegList
    CharSheetEngine.RegisterTraitList tiSocialNeg, Bete.SocialNegList
    CharSheetEngine.RegisterTraitList tiMentalNeg, Bete.MentalNegList
    CharSheetEngine.RegisterTraitList tiAbilities, Bete.AbilityList
    CharSheetEngine.RegisterTraitList tiInfluences, Bete.InfluenceList
    CharSheetEngine.RegisterTraitList tiBackgrounds, Bete.BackgroundList
    CharSheetEngine.RegisterTraitList tiHonor, Bete.HonorList
    CharSheetEngine.RegisterTraitList tiGlory, Bete.GloryList
    CharSheetEngine.RegisterTraitList tiWisdom, Bete.WisdomList
    CharSheetEngine.RegisterTraitList tiHealth, Bete.HealthList

    CharSheetEngine.RegisterTraitList tiMerits, Bete.MeritList
    CharSheetEngine.RegisterTraitList tiFlaws, Bete.FlawList
    CharSheetEngine.RegisterTraitList tiGifts, Bete.GiftList
    CharSheetEngine.RegisterTraitList tiRites, Bete.RiteList
    CharSheetEngine.RegisterTraitList tiFeatures, Bete.FeatureList
    CharSheetEngine.RegisterTraitList tiEquipment, Bete.EquipmentList

    CharSheetEngine.TargetType = LocalTargetType
    Set CharSheetEngine.TargetLabel = LocalTargetLabel
    Set CharSheetEngine.TargetList = LocalTargetList
    
End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Checks to make sure that a character is loaded, which happens only
'               when ShowBete is the means of loading the form.  Initializes the
'               MenuStack linked list and the Bete Menus.
'

    If Bete Is Nothing Then
        MsgBox "Character sheet loaded improperly!"
    Else
        Set MenuStack = New LinkedList
        Set SubMenuLabelStack = New LinkedList
        CharSheetEngine.TargetType = ttNothing
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
    LocalTargetType = CharSheetEngine.TargetType
    Set LocalTargetLabel = CharSheetEngine.TargetLabel
    Set LocalTargetList = CharSheetEngine.TargetList

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Unload
' Description:  Save the text and destroy the MenuStack.
'

    ValidateControls
    MenuStack.Clear
    Set MenuStack = Nothing
    Set SubMenuLabelStack = Nothing

End Sub

Private Sub lblField_Change(Index As Integer)
'
' Name:         lblField_Change
' Description:  Store the new value in the appropriate property of the character.
'

    Dim Value As String
    
    If Not Populating Then
        Value = lblField(Index).Caption
        Select Case Index
            Case fiBete
                Bete.Bete = Value
            Case fiAuspice
                Bete.Auspice = Value
            Case fiBreed
                Bete.Breed = Value
            Case fiNature
                Bete.Nature = Value
            Case fiDemeanor
                Bete.Demeanor = Value
            Case fiRank
                Bete.Rank = Value
            Case fiPosition
                Bete.Position = Value
            Case fiTotem
                Bete.Totem = Value
            Case fiPlayer
                Bete.Player = Value
            Case fiStatus
                Bete.Status = Value
                EntryChangedChars = True
                EntryChangedExp = True
            Case fiNarrator
                Bete.Narrator = Value
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
        
        If CharSheetEngine.TargetType = ttListBox Then
            CharSheetEngine.TargetList.ListIndex = -1
            cmdIncrement.Visible = False
            cmdDecrement.Visible = False
            cmdAscend.Visible = False
            cmdDescend.Visible = False
        End If
        
        lblMenuTitle.Caption = lblFieldLabel(Index).Caption
        CharSheetEngine.PopulateMenu lblField(Index).Tag
        CharSheetEngine.TargetType = ttLabel
        Set CharSheetEngine.TargetLabel = lblField(Index)
        
        lstMenu.SetFocus
        
    End If
    
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

    Dim OrderTop As Integer
    
    If lstTraits(Index).Left > 0 And _
            Not (CharSheetEngine.TargetType = ttListBox And _
            CharSheetEngine.TargetList Is lstTraits(Index)) Then
        
        If CharSheetEngine.TargetType = ttListBox Then _
                CharSheetEngine.TargetList.ListIndex = -1
    
        If CharSheetEngine.CanAdjust(Index) Then
            With lstTraits(Index)
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
                cmdDescend.Move .Left, OrderTop
                cmdAscend.Move .Left + .Width - cmdAscend.Width, OrderTop
            End With
            cmdDescend.Visible = True
            cmdAscend.Visible = True
        Else
            cmdDescend.Visible = False
            cmdAscend.Visible = False
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

Private Sub tabTabs_Click(PreviousTab As Integer)
'
' Name:         tabTabs_Click
' Description:  Clear the menu and targets.
'

    If tabTabs.Tab <> PreviousTab Then
        
        lstMenu.Clear
        lstMenu.Tag = ""
        lblMenuTitle.Caption = ""
        lblMenuItem.Caption = ""
        MenuStack.Clear
        SubMenuLabelStack.Clear
        
        If CharSheetEngine.TargetType = ttListBox Then
            CharSheetEngine.TargetList.ListIndex = -1
            cmdIncrement.Visible = False
            cmdDecrement.Visible = False
            cmdAscend.Visible = False
            cmdDescend.Visible = False
        End If
        
        CharSheetEngine.TargetType = ttNothing
        
    End If

End Sub

Private Sub txtMemo_Validate(Index As Integer, Cancel As Boolean)
'
' Name:         txtMemo_Change
' Description:  Record changes to the memo field.
'

    Select Case Index
        Case miNotes
            If Bete.Notes <> txtMemo(miNotes).Text Then
                SetDataChanged
                Bete.Notes = TrimWhiteSpace(txtMemo(miNotes))
            End If
    End Select

End Sub

Private Sub SetDataChanged()
'
' Name:         SetDataChanged
' Description:  Tell the game its data has changed and update the character's
'               Last Modified date.
'
        
    Game.DataChanged = True
    Bete.LastModified = Date
    lblModified.Caption = Format(Date, "mmmm d, yyyy")

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
                
                CharacterList.MoveTo Trim(txtUserField(Index))
                If Not CharacterList.Off Then
                    If Not CharacterList.Item Is Bete Then
                        MsgBox "The name """ & Trim(txtUserField(Index)) & _
                                """ is already in use.  Please use a different name.", _
                                vbOKOnly Or vbExclamation, "Duplicate Name"
                        Exit Sub
                    End If
                End If
                EntryChangedChars = True
                EntryChangedExp = True
                NamesChangedInfUse = True
                Bete.Name = Trim(txtUserField(Index))
                Me.Caption = Trim(txtUserField(Index))
            
            Case xiPack
                Bete.Pack = Trim(txtUserField(Index))
            Case xiID
                Bete.ID = Trim(txtUserField(Index))
            Case xiStartDate
                Bete.StartDate = Trim(txtUserField(Index))
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
            Bete.Notoriety = updNumber(niNotoriety).Value
    End Select
    SetDataChanged
    
End Sub

Private Sub cmdChange_Click()
'
' Name:         cmdChange_Click
' Description:  Add a new entry to the experience history, and adjust
'               the XP display accordingly.
'

    Populating = True
    CharSheetEngine.ChangeExperience Bete.Experience
    txtExperience = " " & CStr(Bete.Experience.Unspent) & _
            " / " & CStr(Bete.Experience.Earned)
    Populating = False

End Sub

Private Sub txtExperience_Change()
'
' Name:         txtExperience_Change
' Description:  Ensure a valid value and save the change to the character's
'               experience.
'
    
    If Not Populating Then
        
        Dim Slash As Integer
        Dim Estr As String
        Dim Ustr As String
        
        Slash = InStr(txtExperience.Text, "/")
        
        If Slash > 0 Then
            
            Ustr = Trim(Left(txtExperience.Text, Slash - 1))
            Estr = Trim(Mid(txtExperience.Text, Slash + 1))
            
            If (IsNumeric(Ustr) Or Ustr = "") And _
               (IsNumeric(Estr) Or Estr = "") Then
                Bete.Experience.Unspent = Val(Ustr)
                Bete.Experience.Earned = Val(Estr)
                txtExperience.ForeColor = vbWindowText
                SetDataChanged
            Else
                txtExperience.ForeColor = vbHighlight
            End If
            
        Else
            txtExperience.Text = " " & CStr(Bete.Experience.Unspent) & _
                    " / " & CStr(Bete.Experience.Earned)
        End If
        
    End If
    
End Sub

Private Sub txtExperience_GotFocus()
'
' Name:         txtExperience_GotFocus
' Description:  Select the Text.
'

    SelectText txtExperience

End Sub

Private Sub txtExperience_DblClick()
'
' Name:         txtExperience_DblClick()
' Description:  Highlight all text.
'

    SelectText txtExperience

End Sub

Private Sub txtExperience_KeyPress(KeyAscii As Integer)
'
' Name:         txtExperience_KeyPress
' Description:  Shortcut the return key.
'

    If KeyAscii = vbKeyReturn Then KeyAscii = 0

End Sub

Private Sub updExperience_DownClick()
'
' Name:         updExperience_DownClick
' Description:  Update the label and store the new value.
'

    Dim EditBoth As Boolean
    Dim EditUnspent As Boolean
    Dim SaveStart As Integer
    Dim SaveLen As Integer
    
    Populating = True
    
    With txtExperience
    
        SaveStart = .SelStart
        SaveLen = .SelLength
        If SaveLen = Len(.Text) Then SaveLen = SaveLen + 2
        EditBoth = InStr(.SelText, "/") > 0 Or Not Me.ActiveControl Is txtExperience
        EditUnspent = SaveStart <= InStr(.Text, "/")
        
        With Bete.Experience
            If EditBoth Or EditUnspent Then   'take from unspent
                .Unspent = .Unspent - 1
            Else                               'take from earned
                .Earned = .Earned - 1
            End If
                
            txtExperience.Text = " " & CStr(.Unspent) & " / " & CStr(.Earned)
                    
        End With
        
        .SelStart = SaveStart
        .SelLength = SaveLen
        
    End With
        
    Populating = False
    
    SetDataChanged

End Sub

Private Sub updExperience_UpClick()
'
' Name:         updExperience_UpClick
' Description:  Update the label and store the new value.
'

    Dim EditBoth As Boolean
    Dim EditUnspent As Boolean
    Dim SaveStart As Integer
    Dim SaveLen As Integer
    
    Populating = True
    
    With txtExperience
    
        SaveStart = .SelStart
        SaveLen = .SelLength
        If SaveLen = Len(.Text) Then SaveLen = SaveLen + 2
        EditBoth = InStr(.SelText, "/") > 0 Or Not Me.ActiveControl Is txtExperience
        EditUnspent = SaveStart <= InStr(.Text, "/")
        
        With Bete.Experience
            If EditBoth Then   'add to both
                .Unspent = .Unspent + 1
                .Earned = .Earned + 1
            ElseIf EditUnspent Then  'add to unspent
                .Unspent = .Unspent + 1
            Else                'add to earned
                .Earned = .Earned + 1
            End If
                
            txtExperience.Text = " " & CStr(.Unspent) & " / " & CStr(.Earned)
                    
        End With
        
        .SelStart = SaveStart
        .SelLength = SaveLen
        
    End With
        
    Populating = False
    
    SetDataChanged

End Sub

Private Sub updTemper_Change(Index As Integer)
'
' Name:         updTemper_Change
' Description:  Update the label and store the new value.
'

    lblTemper(Index).Caption = " " & CStr(updTemper(Index).Value) _
        & " " & String(updTemper(Index).Value, "o")

    Select Case Index
        Case piRage
            Bete.Rage = updTemper(piRage).Value
        Case piGnosis
            Bete.Gnosis = updTemper(piGnosis).Value
        Case piWillpower
            Bete.Willpower = updTemper(piWillpower).Value
    End Select
    SetDataChanged

End Sub

Private Sub txtDecimal_Change(Index As Integer)
'
' Name:         txtDecimal_Change
' Description:  Ensure a valid value and save the change to the character's
'               experience.
'
    
    If Not Populating Then
        Select Case Val(txtDecimal(Index)) * 10
            Case Is < updDecimal(Index).Min
                txtDecimal(Index) = " " & CStr(updDecimal(Index).Min)
            Case Is > updDecimal(Index).Max
                txtDecimal(Index) = " " & CStr(updDecimal(Index).Max)
        End Select
        updDecimal(Index).Value = Int(10 * Val(txtDecimal(Index)))
        
        txtDecimal(Index).ForeColor = IIf(IsNumeric(txtDecimal(Index)), _
                vbWindowText, vbHighlight)
        
        Select Case Index
            Case diHonor
                Bete.Honor = Val(txtDecimal(diHonor))
            Case diGlory
                Bete.Glory = Val(txtDecimal(diGlory))
            Case diWisdom
                Bete.Wisdom = Val(txtDecimal(diWisdom))
        End Select
        SetDataChanged
    End If
    
End Sub

Private Sub txtDecimal_GotFocus(Index As Integer)
'
' Name:         txtDecimal_GotFocus
' Description:  Select the Text.
'

    SelectText txtDecimal(Index)

End Sub

Private Sub txtDecimal_KeyPress(Index As Integer, KeyAscii As Integer)
'
' Name:         txtDecimal_KeyPress
' Description:  Shortcut the return key.
'

    If KeyAscii = vbKeyReturn Then KeyAscii = 0

End Sub

Private Sub updDecimal_DownClick(Index As Integer)
'
' Name:         updDecimal_DownClick
' Description:  Update the label and store the new value.
'

    txtDecimal(Index) = " " & CStr(Val(txtDecimal(Index)) - 0.1)

End Sub

Private Sub updDecimal_UpClick(Index As Integer)
'
' Name:         updDecimal_UpClick
' Description:  Update the label and store the new value.
'

    txtDecimal(Index) = " " & CStr(Val(txtDecimal(Index)) + 0.1)

End Sub



