VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGameInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Settings"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   Icon            =   "frmGameInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9030
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   8295
      Begin MSComCtl2.UpDown updAutoSave 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   5220
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3150
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         _Version        =   393216
         Value           =   10
         BuddyControl    =   "txtInfo(12)"
         BuddyDispid     =   196610
         BuddyIndex      =   12
         OrigLeft        =   5400
         OrigTop         =   3120
         OrigRight       =   5655
         OrigBottom      =   3315
         Max             =   1440
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   12
         Left            =   4800
         TabIndex        =   50
         Text            =   "10"
         Top             =   3150
         Width           =   420
      End
      Begin VB.CheckBox chkEnforce 
         Caption         =   "&Enforce the use of Experience Point histories and Player Point histories"
         Height          =   195
         Left            =   1680
         TabIndex        =   82
         Top             =   1320
         Width           =   6495
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   11
         Left            =   3000
         TabIndex        =   43
         Text            =   "[/ST]"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   10
         Left            =   1680
         TabIndex        =   41
         Text            =   "[ST]"
         Top             =   2160
         Width           =   855
      End
      Begin VB.CheckBox chkPreferences 
         Caption         =   "Edit Physical, Social and Mental &Trait Maximums separately"
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   39
         Top             =   1680
         Width           =   6495
      End
      Begin VB.CommandButton cmdRegisterTypes 
         Caption         =   "&Re-Associate File Types"
         Height          =   375
         Left            =   6120
         TabIndex        =   53
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CommandButton cmdGrapevineMenus 
         Caption         =   "Grapevine &Menus..."
         Height          =   375
         Left            =   6120
         TabIndex        =   47
         Top             =   2640
         Width           =   2055
      End
      Begin VB.OptionButton optExtended 
         Caption         =   "Create characters with &Extended (seven) health levels"
         Height          =   255
         Left            =   1680
         TabIndex        =   34
         Top             =   240
         Value           =   -1  'True
         Width           =   4410
      End
      Begin VB.OptionButton optAbbreviated 
         Caption         =   "Create characters with Abbre&viated (four) health levels"
         Height          =   375
         Left            =   1680
         TabIndex        =   35
         Top             =   480
         Width           =   4410
      End
      Begin VB.CommandButton cmdExtended 
         Caption         =   "Convert to E&xtended"
         Height          =   375
         Left            =   6120
         TabIndex        =   36
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdAbbreviated 
         Caption         =   "Convert to A&bbreviated"
         Height          =   375
         Left            =   6120
         TabIndex        =   37
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox chkPreferences 
         Caption         =   "Save a backup copy of the game every"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   49
         Top             =   3120
         Width           =   3375
      End
      Begin VB.Label lblLabels 
         Caption         =   " minutes   ( ~Autosave.gv3 )"
         Height          =   255
         Index           =   21
         Left            =   5760
         TabIndex        =   52
         Top             =   3195
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "&Autosave"
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   48
         Top             =   3195
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Experience Histories"
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   83
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Caption         =   "from appearing in output"
         Height          =   255
         Index           =   20
         Left            =   3960
         TabIndex        =   44
         Top             =   2205
         Width           =   2055
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "and"
         Height          =   255
         Index           =   19
         Left            =   2520
         TabIndex        =   42
         Top             =   2205
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Hide Te&xt between"
         Height          =   375
         Index           =   18
         Left            =   120
         TabIndex        =   40
         Top             =   2205
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Trait Maximums"
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   38
         Top             =   1755
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Grapevine Menu File"
         Height          =   375
         Index           =   11
         Left            =   0
         TabIndex        =   45
         Top             =   2670
         Width           =   1575
      End
      Begin VB.Label lblMenuFile 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1680
         TabIndex        =   46
         Top             =   2640
         Width           =   4335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Health Levels"
         Height          =   315
         Index           =   8
         Left            =   285
         TabIndex        =   33
         Top             =   240
         Width           =   1290
      End
   End
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   3
      Left            =   480
      TabIndex        =   54
      Top             =   1200
      Visible         =   0   'False
      Width           =   8175
      Begin VB.CheckBox chkRumors 
         Caption         =   "Copy &Unused Values from Previous Action"
         Height          =   375
         Index           =   9
         Left            =   2160
         TabIndex        =   60
         Top             =   960
         Width           =   3375
      End
      Begin VB.CheckBox chkRumors 
         Caption         =   "Always add Co&mmon Actions"
         Height          =   255
         Index           =   8
         Left            =   2160
         TabIndex        =   59
         Top             =   720
         Width           =   2415
      End
      Begin VB.ListBox lstBackgrounds 
         Height          =   1425
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   67
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox txtPersonal 
         Height          =   285
         Left            =   2175
         TabIndex        =   56
         Text            =   "0"
         Top             =   240
         Width           =   540
      End
      Begin VB.TextBox txtValue 
         Height          =   285
         Left            =   4680
         TabIndex        =   64
         Text            =   "0"
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdAddBackground 
         Caption         =   "&Add"
         Height          =   375
         Left            =   4680
         TabIndex        =   68
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton cmdDeleteBackground 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4680
         TabIndex        =   69
         Top             =   3480
         Width           =   1815
      End
      Begin MSComCtl2.UpDown updValue 
         Height          =   285
         Left            =   5280
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1800
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtValue"
         BuddyDispid     =   196624
         OrigLeft        =   5341
         OrigTop         =   1200
         OrigRight       =   5536
         OrigBottom      =   1485
         Max             =   9999
         Min             =   -9999
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updPersonal 
         Height          =   285
         Left            =   2715
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   240
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtPersonal"
         BuddyDispid     =   196623
         OrigLeft        =   2655
         OrigTop         =   240
         OrigRight       =   3090
         OrigBottom      =   525
         Max             =   9999
         Min             =   -9999
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComctlLib.ListView lvwLevels 
         Height          =   1455
         Left            =   2160
         TabIndex        =   62
         Top             =   1440
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Level"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   1640
         EndProperty
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Detailed Action Settings"
         Height          =   255
         Index           =   1
         Left            =   -120
         TabIndex        =   58
         Top             =   735
         Width           =   2175
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Total &Personal Action Value"
         Height          =   255
         Index           =   2
         Left            =   -120
         TabIndex        =   55
         Top             =   285
         Width           =   2175
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Action Values per &Level of Influence or Background"
         Height          =   495
         Index           =   3
         Left            =   -120
         TabIndex        =   61
         Top             =   1485
         Width           =   2175
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Backgrounds &with Action Values"
         Height          =   495
         Index           =   5
         Left            =   -120
         TabIndex        =   66
         Top             =   3045
         Width           =   2175
      End
      Begin VB.Label lblLabel 
         Caption         =   "Action &Value"
         Height          =   255
         Index           =   4
         Left            =   4680
         TabIndex        =   63
         Top             =   1560
         Width           =   1695
      End
   End
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4575
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   8295
      Begin MSComCtl2.MonthView mvwCalendar 
         Height          =   2370
         Left            =   2280
         TabIndex        =   24
         Top             =   480
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   47054849
         CurrentDate     =   37497
      End
      Begin VB.Frame fraDateFields 
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   5160
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   3015
         Begin VB.TextBox txtInfo 
            Height          =   1695
            Index           =   9
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            Top             =   2520
            Width           =   3015
         End
         Begin VB.TextBox txtInfo 
            Height          =   375
            Index           =   7
            Left            =   0
            TabIndex        =   28
            Top             =   600
            Width           =   3015
         End
         Begin VB.TextBox txtInfo 
            Height          =   855
            Index           =   8
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   1320
            Width           =   3015
         End
         Begin VB.Label lblDate 
            Alignment       =   2  'Center
            Caption         =   "January 31, 1978"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   3210
         End
         Begin VB.Label lblLabels 
            Caption         =   "&Notes"
            Height          =   255
            Index           =   17
            Left            =   0
            TabIndex        =   31
            Top             =   2280
            Width           =   3015
         End
         Begin VB.Label lblLabels 
            Caption         =   "P&lace"
            Height          =   255
            Index           =   16
            Left            =   0
            TabIndex        =   29
            Top             =   1080
            Width           =   3015
         End
         Begin VB.Label lblLabels 
            Caption         =   "&Time"
            Height          =   255
            Index           =   15
            Left            =   0
            TabIndex        =   27
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.CommandButton cmdDeleteOld 
         Caption         =   "Delete &Older Dates"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   4080
         Width           =   2055
      End
      Begin VB.CommandButton cmdDeleteDate 
         Caption         =   "D&elete Date"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   3600
         Width           =   2055
      End
      Begin VB.ListBox lstDates 
         Height          =   2985
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Click a date to &schedule a game."
         Height          =   255
         Index           =   14
         Left            =   2280
         TabIndex        =   23
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "&Game Dates"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.TextBox txtInfo 
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
      Left            =   1920
      TabIndex        =   80
      Top             =   120
      Width           =   6855
   End
   Begin VB.CommandButton cmdESCClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   8295
      Begin VB.TextBox txtInfo 
         Height          =   1815
         Index           =   6
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   2640
         Width           =   6615
      End
      Begin VB.ListBox lstStaff 
         Height          =   840
         Left            =   1560
         TabIndex        =   12
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtInfo 
         Height          =   855
         Index           =   5
         Left            =   5520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtInfo 
         Height          =   375
         Index           =   4
         Left            =   5520
         TabIndex        =   14
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtInfo 
         Height          =   375
         Index           =   3
         Left            =   1560
         TabIndex        =   10
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtInfo 
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   8
         Top             =   720
         Width           =   6615
      End
      Begin VB.TextBox txtInfo 
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "De&scription"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   17
         Top             =   2685
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Staff"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   11
         Top             =   1710
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "&Usual Place"
         Height          =   255
         Index           =   5
         Left            =   4320
         TabIndex        =   15
         Top             =   1725
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Usual &Time"
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   13
         Top             =   1245
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "&Phone"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   1245
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "&E-Mail"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   765
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "&Web Page"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   4
      Left            =   360
      TabIndex        =   70
      Top             =   1200
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CheckBox chkRumors 
         Caption         =   "&Personal rumors"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   73
         Top             =   600
         Width           =   4215
      End
      Begin VB.CheckBox chkRumors 
         Caption         =   "&Group rumors (clan, tribe, etc.)"
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   75
         Top             =   1320
         Width           =   4215
      End
      Begin VB.CheckBox chkRumors 
         Caption         =   "&Subgroup rumors (sect, auspice, etc.)"
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   76
         Top             =   1680
         Width           =   4095
      End
      Begin VB.CheckBox chkRumors 
         Caption         =   "&Racial rumors (Vampire, Werewolf, etc.)"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   74
         Top             =   960
         Width           =   3855
      End
      Begin VB.CheckBox chkRumors 
         Caption         =   "Public &Knowledge rumors"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   72
         Top             =   240
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.CheckBox chkRumors 
         Caption         =   "I&nfluence rumors"
         Height          =   255
         Index           =   5
         Left            =   2280
         TabIndex        =   77
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkRumors 
         Caption         =   "&All additional rumor types from the previous session"
         Height          =   255
         Index           =   6
         Left            =   2280
         TabIndex        =   78
         Top             =   2400
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.CheckBox chkRumors 
         Caption         =   "C&opy all rumor text from the previous session"
         Height          =   375
         Index           =   7
         Left            =   2520
         TabIndex        =   79
         Top             =   2640
         Width           =   3735
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Standard Rumors include the following:"
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   71
         Top             =   240
         Width           =   2175
      End
   End
   Begin MSComctlLib.TabStrip tabTabs 
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8916
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Information"
            Key             =   "Info"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dates"
            Key             =   "Dates"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General Settings"
            Key             =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Action Settings"
            Key             =   "Actions"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Rumor Settings"
            Key             =   "Rumors"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   240
      Picture         =   "frmGameInfo.frx":058A
      Top             =   187
      Width           =   480
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&Chronicle Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   81
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmGameInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const tiTitle = 0
Private Const tiWebPage = 1
Private Const tiEMail = 2
Private Const tiPhone = 3
Private Const tiUsualTime = 4
Private Const tiUsualPlace = 5
Private Const tiDescription = 6
Private Const tiDateTime = 7
Private Const tiDatePlace = 8
Private Const tiDateNotes = 9
Private Const tiHideStart = 10
Private Const tiHideEnd = 11
Private Const tiAutosave = 12

Private Const ciPublic = 0
Private Const ciPersonal = 1
Private Const ciRace = 2
Private Const ciGroup = 3
Private Const ciSubgroup = 4
Private Const ciInfluence = 5
Private Const ciPrevious = 6
Private Const ciCopy = 7
Private Const ciCommon = 8
Private Const ciUnused = 9

Private Const ciTraitMax = 0
Private Const ciAutosave = 1

Private Populating As Boolean           'are we populating fields?

Public Sub ShowWith(TabKey As String)
'
' Name:         ShowWith
' Parameters:   TabKey          key of the tab to show
' Description:  Show the form with the given tab showing.
'

    tabTabs.Tabs(TabKey).Selected = True
    Me.Show
    Me.SetFocus

End Sub

Private Sub RefreshStaffList()
'
' Name:         RefreshStaffList
' Description:  Populate the staff list.
'
    
    Dim Role As String
    Dim STCount As Integer
    
    PlayerList.First
    lstStaff.Clear
    
    Do Until PlayerList.Off
        If PlayerList.Item.Status = ActiveStatus Then
            Role = PlayerList.Item.Position
            If Role = "Storyteller" Or InStr(Role, "ST") > 0 Then
                lstStaff.AddItem PlayerList.Item.Name & ", " & Role, STCount
                STCount = STCount + 1
            ElseIf Role <> "Player" Then
                lstStaff.AddItem PlayerList.Item.Name & ", " & Role, _
                        lstStaff.ListCount
            End If
        End If
        PlayerList.MoveNext
    Loop

End Sub

Public Sub SetDefaultOutput()
'
' Name:         SetDefaultOutput
' Description:  Initilize the OutputEngineClass with default output settings.
'
    With OutputEngine
        .Template = tnGameCalendar
        .GameDate = 0
    End With
    
End Sub

Private Sub chkEnforce_Click()
'
' Name:         chkEnforce_Click
' Description:  Set the state of DataChanged, EnforceHistory.
'

    If Not Populating Then
        
        Populating = True
        If chkEnforce.Value = vbChecked Then
            
            Dim Reply As Integer
            Dim Explanation As String
            
            Explanation = "Selecting this option forces experience points and player points to " & _
                          "accurately reflect their histories of earnings and expenditures." & _
                          vbCrLf & vbCrLf & "Would you like to bring all such points into line " & _
                          "with their history entries now?" & vbCrLf & vbCrLf & _
                          "Clicking YES changes all points to their histories' latest totals." & vbCrLf & _
                          "Clicking NO changes the histories to bring their totals in line with current " _
                          & "points." & vbCrLf & "Clicking CANCEL turns off this option."
            Reply = MsgBox(Explanation, vbYesNoCancel + vbQuestion, "Enforce Experience Histories")
                        
            If Reply <> vbCancel Then
            
                Dim XP As ExperienceClass
                Dim HistoryEarned As Single
                Dim HistoryUnspent As Single
                
                CharacterList.First
                Do Until CharacterList.Off
                    Set XP = CharacterList.Item.Experience
                    
                    If XP.IsEmpty Then
                        HistoryEarned = 0
                        HistoryUnspent = 0
                    Else
                        XP.Last
                        HistoryEarned = XP.EntryEarned
                        HistoryUnspent = XP.EntryUnspent
                    End If
                    
                    If Reply = vbYes Then
                        XP.Earned = HistoryEarned
                        XP.Unspent = HistoryUnspent
                    Else
                        If XP.Earned <> HistoryEarned Then
                            XP.Insert XP.Earned, ecSetEarned, Date, "Enforced History: Changed from " _
                                      & CStr(HistoryEarned) & " to " & CStr(XP.Earned) & "."
                        End If
                        If XP.Unspent <> HistoryUnspent Then
                            XP.Insert XP.Unspent, ecSetUnspent, Date, "Enforced History: Changed from " _
                                      & CStr(HistoryUnspent) & " to " & CStr(XP.Unspent) & "."
                        End If
                    End If
                    
                    CharacterList.MoveNext
                Loop
                Game.DataChanged = True
                Game.EnforceHistory = True
                
            Else
            
                chkEnforce.Value = vbUnchecked
            
            End If
            
        Else
            Game.DataChanged = True
            Game.EnforceHistory = False
        End If
        Populating = False
        
    End If

End Sub

Private Sub chkPreferences_Click(Index As Integer)
'
' Name:         chkGuess_Click
' Description:  Turn on/off the chosen game preferences.
'

    If Not Populating Then
        Select Case Index
            Case ciTraitMax
                Game.LinkTraitMaxes = Not (chkPreferences(ciTraitMax).Value = vbChecked)
                Game.DataChanged = True
            Case ciAutosave
                If chkPreferences(ciAutosave).Value = vbChecked Then
                    mdiMain.AutosaveTime = Int(Val(txtInfo(tiAutosave).Text))
                Else
                    mdiMain.AutosaveTime = 0
                End If
        End Select
    End If

End Sub

Private Sub cmdGrapevineMenus_Click()
'
' Name:         cmdGrapevineMenus_Click
' Description:  Display the Grapevine Menu Editor window.
'

    frmMenuEditor.Show

End Sub

Private Sub cmdDeleteDate_Click()
'
' Name:         cmdDeleteDate_Click
' Description:  Delete the selected game date.
'
  
    If lstDates.ListIndex > -1 Then
        If MsgBox("Deleting this game date will also delete data associated with the date," & _
                  " such as Actions and Rumors.  Are you sure you want to do this?", vbYesNo + _
                  vbQuestion, "Delete Date") = vbYes Then
            
            Dim StorePos As Integer
            Dim NormForm As Form
            
            For Each NormForm In Forms
                If NormForm.Tag = "U" Or NormForm.Tag = "A" Then
                    Unload NormForm
                End If
            Next NormForm
            
            Game.Calendar.MoveTo CDate(lstDates.Text)
            If Not Game.Calendar.Off Then
                Game.Calendar.Remove
                Game.DataChanged = True
                Game.Calendar.LastModified = Now
                mdiMain.AnnounceChanges Me, atDates
            End If
            
            Game.APREngine.DeleteDate CDate(lstDates.Text)
            
            StorePos = lstDates.ListIndex
            lstDates.RemoveItem StorePos
            If lstDates.ListCount = StorePos Then StorePos = StorePos - 1
            lstDates.ListIndex = StorePos
            
            If StorePos = -1 Then
                fraDateFields.Visible = False
                mvwCalendar.Value = Date
            End If
            
        End If
    End If
    
End Sub

Private Sub cmdDeleteOld_Click()
'
' Name:         cmdDeleteOld_Click
' Description:  Delete all dates older than the selected date.
'

    If lstDates.ListIndex > -1 Then
        If MsgBox("You're about to delete every game date older than the selected date. " & _
                  "This will also delete associated data such as Actions and Rumors " & _
                  "from all these dates.  Are you sure you want to do this?", vbYesNo + _
                  vbQuestion, "Delete Old Dates") = vbYes Then
            
            Dim NormForm As Form
            
            For Each NormForm In Forms
                If NormForm.Tag = "U" Or NormForm.Tag = "A" Then
                    Unload NormForm
                End If
            Next NormForm
            
            With Game.Calendar
                
                .MoveTo CDate(lstDates.Text)
                .MovePrevious
                Do Until .Off
                    Game.APREngine.DeleteDate .GetGameDate
                    .MovePrevious
                Loop
            
                .MoveTo CDate(lstDates.Text)
                Do Until lstDates.ListIndex = lstDates.ListCount - 1
                    .MovePrevious
                    .Remove
                    lstDates.RemoveItem lstDates.ListIndex + 1
                Loop
                
            End With
            
            Game.DataChanged = True
            Game.Calendar.LastModified = Now
            mdiMain.AnnounceChanges Me, atDates

        End If
    End If

End Sub

Private Sub cmdESCClose_Click()
'
' Name:         cmdESCClose_Click
' Description:  Close the window when the user presses ESC.
'

    Unload Me

End Sub

Private Sub ConvertHealth(Extend As Boolean)
'
' Name:         ConvertHealth
' Arguments:    Extend (Boolean) -- TRUE if conversion to extended health levels,
'               FALSE if to abbreviated health levels
' Description:  Convert all the characters in the database to a given base set of
'               health levels
'

    Dim CharForm As Form
    Dim Race As RaceType
    Dim HealthList As LinkedTraitList
    
    '
    ' Unload the visible character sheets
    '
    For Each CharForm In Forms
        If CharForm.Tag = "C" Then Unload CharForm
    Next CharForm
    
    CharacterList.First
    Do Until CharacterList.Off
        
        Race = CharacterList.Item.RaceCode
        
        If Not (Race = gvRaceWraith Or Race = gvRaceVarious) Then
            
            Set HealthList = CharacterList.Item.HealthList
            
            HealthList.MoveTo StdHealth(1)
            
            '
            ' Check the characters by their "Bruised" levels.  Less than 3,
            ' They can be extended; 3 or more, they can be abbreviated.
            '
            If HealthList.Trait.Number < 3 And Extend Then
            
                HealthList.Insert StdHealth(0)
                HealthList.Insert StdHealth(1), "2"
                HealthList.Insert StdHealth(2)
                
            ElseIf HealthList.Trait.Number > 2 And Not Extend Then
                
                HealthList.MoveTo StdHealth(0)
                HealthList.Remove
                HealthList.MoveTo StdHealth(1)
                HealthList.Remove
                HealthList.Remove
                HealthList.MoveTo StdHealth(2)
                HealthList.Remove
            
            End If
               
        End If
        
        CharacterList.MoveNext
        
    Loop

    Game.DataChanged = True

End Sub

Private Sub cmdAbbreviated_Click()
'
' Name:         cmdAbbreviated_Click
' Description:  Convert to abbreviated health levels
'
    If MsgBox("This will PERMANENTLY REMOVE the health levels " & vbCrLf & _
            StdHealth(0) & " x1, " & StdHealth(1) & " x2, and " & StdHealth(2) & _
            " x1" & vbCrLf & "from characters that have extended health levels." & vbCrLf & _
            "Wraiths and Various-type characters will not be affected." & vbCrLf & _
            vbCrLf & "Are you sure you want to continue?", vbYesNo, "Convert to Abbreviated Health") _
            = vbYes Then
                        
            Screen.MousePointer = vbHourglass
            ConvertHealth False
            Screen.MousePointer = vbDefault
            
    End If

End Sub

Private Sub cmdExtended_Click()
'
' Name:         cmdExtended_Click
' Description:  Convert to extended health levels
'
    If MsgBox("This will PERMANENTLY ADD the health levels " & vbCrLf & _
            StdHealth(0) & " x1, " & StdHealth(1) & " x2, and " & StdHealth(2) & _
            " x1 to characters without them." & vbCrLf & _
            "Wraiths and Various-type characters will not be affected." & vbCrLf & _
            vbCrLf & "Are you sure you want to continue?", vbYesNo, "Convert to Extended Health") _
            = vbYes Then
            
            Screen.MousePointer = vbHourglass
            ConvertHealth True
            Screen.MousePointer = vbDefault
            
    End If

End Sub

Private Sub cmdRegisterTypes_Click()
'
' Name:         cmdRegisterTypes_Click
' Description:  Register file types with this version of Grapevine,
'               in the event of multiple installations
'

    CreateGVAssociation SlashPath(App.Path)
    MsgBox "Grapevine file type associations created.  You may have to reboot before you see the effects.", _
            vbOKOnly, "File Types Renewed"

End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  Refresh the staff list if needed.
'
    
    If mdiMain.CheckForChanges(Me, atPlayers) Then RefreshStaffList

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Initialize the game settings.
'

    Dim Role As String
    Dim I As Integer
    Dim NewItem As ListItem
    
    Populating = True
    
    txtInfo(tiTitle).Text = Game.ChronicleTitle
    txtInfo(tiWebPage).Text = Game.Website
    txtInfo(tiEMail).Text = Game.EMail
    txtInfo(tiPhone).Text = Game.Phone
    txtInfo(tiUsualPlace).Text = Game.UsualPlace
    txtInfo(tiUsualTime).Text = Game.UsualTime
    txtInfo(tiDescription).Text = Game.Description
    txtInfo(tiHideStart).Text = Game.STCommentStart
    txtInfo(tiHideEnd).Text = Game.STCommentEnd
    
    RefreshStaffList
    
    optExtended.Value = Game.ExtendedHealth
    optAbbreviated.Value = Not Game.ExtendedHealth
    chkEnforce.Value = IIf(Game.EnforceHistory, vbChecked, vbUnchecked)
    chkPreferences(ciTraitMax).Value = IIf(Game.LinkTraitMaxes, vbUnchecked, vbChecked)
    chkPreferences(ciAutosave).Value = IIf(mdiMain.AutosaveTime > 0, vbChecked, vbUnchecked)
    txtInfo(tiAutosave).Text = CStr(IIf(mdiMain.AutosaveTime > 0, mdiMain.AutosaveTime, 10))
    
    lblMenuFile.Caption = Game.MenuSet.FileName
    mvwCalendar.Value = Date
    
    fraDateFields.Left = mvwCalendar.Left + mvwCalendar.Width + 105
    fraDateFields.Width = 8175 - fraDateFields.Left
    lblDate.Width = fraDateFields.Width
    lblLabels(15).Width = fraDateFields.Width
    lblLabels(16).Width = fraDateFields.Width
    lblLabels(17).Width = fraDateFields.Width
    txtInfo(7).Width = fraDateFields.Width
    txtInfo(8).Width = fraDateFields.Width
    txtInfo(9).Width = fraDateFields.Width
    
    With Game.Calendar
        .Last
        Do Until .Off
            lstDates.AddItem Format(.GetGameDate, "mmmm d, yyyy")
            .MovePrevious
        Loop
        If .Count > 0 Then lstDates.ListIndex = 0
    End With
    
    With Game.APREngine
    
        For I = 1 To 10
            .ActionsPerLevel.MoveTo CStr(I)
            If Not .ActionsPerLevel.Off Then
                Set NewItem = lvwLevels.ListItems.Add(I, , "Influence x" & CStr(I))
                NewItem.ListSubItems.Add , , .ActionsPerLevel.Trait.Total & " Actions"
            End If
        Next I
    
        If lvwLevels.ListItems.Count > 0 Then
            Set lvwLevels.SelectedItem = lvwLevels.ListItems(1)
            Call lvwLevels_ItemClick(lvwLevels.SelectedItem)
        End If
        
        .BackgroundActions.First
        Do Until .BackgroundActions.Off
            lstBackgrounds.AddItem .BackgroundActions.Trait.Name
            .BackgroundActions.MoveNext
        Loop
    
        txtPersonal.Text = .PersonalActions
        
        Populating = True
    
        chkRumors(ciPublic).Value = IIf(.PublicRumors, vbChecked, vbUnchecked)
        chkRumors(ciPersonal).Value = IIf(.PersonalRumors, vbChecked, vbUnchecked)
        chkRumors(ciRace).Value = IIf(.RaceRumors, vbChecked, vbUnchecked)
        chkRumors(ciGroup).Value = IIf(.GroupRumors, vbChecked, vbUnchecked)
        chkRumors(ciSubgroup).Value = IIf(.SubgroupRumors, vbChecked, vbUnchecked)
        chkRumors(ciInfluence).Value = IIf(.InfluenceRumors, vbChecked, vbUnchecked)
        chkRumors(ciPrevious).Value = IIf(.PreviousRumors, vbChecked, vbUnchecked)
        chkRumors(ciCopy).Value = IIf(.CopyPrevious, vbChecked, vbUnchecked)
        chkRumors(ciCommon).Value = IIf(.AddCommon, vbChecked, vbUnchecked)
        chkRumors(ciUnused).Value = IIf(.CarryUnused, vbChecked, vbUnchecked)
    
    End With

    mdiMain.OrientForm Me

    Populating = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Unload
' Description:  Remove me from memory.
'

    ValidateControls
    Unload Me

End Sub

Private Sub lstDates_Click()
'
' Name:         lstDates_Click
' Description:  Select a game date.  Populate fields for it.
'

    If Not lstDates.ListIndex = -1 Then
        
        Populating = True
    
        lblDate.Caption = lstDates.Text
        
        With Game.Calendar
            .MoveTo CDate(lstDates.Text)
            If Not .Off Then
                txtInfo(tiDateTime).Text = .GetGameTime
                txtInfo(tiDatePlace).Text = .GetGamePlace
                txtInfo(tiDateNotes).Text = .GetGameNotes
                mvwCalendar.Value = .GetGameDate
                fraDateFields.Visible = True
            End If
        End With
        
        Populating = False
    
    End If
    
End Sub

Private Sub mvwCalendar_DateClick(ByVal DateClicked As Date)
'
' Name:         mvwCalendar_DateClick
' Parameters:   DateClicked     the date clicked by the user
' Description:  Enter selected date into the calendar, or select it if it
'               already exists.
'

    Dim StringDate As String
    Dim DateIndex As Integer
    
    mvwCalendar.Value = DateClicked
    DateIndex = 0
    StringDate = Format(DateClicked, "mmmm d, yyyy")
    Game.Calendar.MoveTo DateClicked

    If Game.Calendar.Off Then
    
        Game.Calendar.Insert DateClicked, Game.UsualTime, Game.UsualPlace, ""
        
        Do Until DateIndex >= lstDates.ListCount
            If CDate(lstDates.List(DateIndex)) < DateClicked Then Exit Do
            DateIndex = DateIndex + 1
        Loop
        lstDates.AddItem StringDate, DateIndex
        lstDates.ListIndex = DateIndex
        Game.DataChanged = True
        mdiMain.AnnounceChanges Me, atDates
        
    Else

        Do Until DateIndex = lstDates.ListCount
            If StringDate = lstDates.List(DateIndex) Then Exit Do
            DateIndex = DateIndex + 1
        Loop
        lstDates.ListIndex = DateIndex
    
    End If
        
End Sub

Private Sub mvwCalendar_KeyPress(KeyAscii As Integer)
'
' Name:         mvwCalendar_KeyPress
' Parameters:   KeyAscii        ascii value of key pressed
' Description:  Keyboard trigger for date selection.
'

    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
        Call mvwCalendar_DateClick(mvwCalendar.Value)
    End If

End Sub

Private Sub optAbbreviated_Click()
'
' Name:         optAbbreviated_Click
' Description:  Set the state of ExtendedHealth, DataChanged.
'

    If Not Populating Then
        Game.DataChanged = True
        Game.ExtendedHealth = False
    End If
    
End Sub

Private Sub optExtended_Click()
'
' Name:         optExtended_Click
' Description:  Set the state of ExtendedHealth, DataChanged.
'

    If Not Populating Then
        Game.DataChanged = True
        Game.ExtendedHealth = True
    End If

End Sub

Private Sub tabTabs_Click()
'
' Name:         tabTabs_Click()
' Description:  Show the needed frame.
'

    Dim F As Frame
    
    For Each F In fraFrame
        F.Visible = (F.Index = (tabTabs.SelectedItem.Index - 1))
    Next F

End Sub

Private Sub txtInfo_Change(Index As Integer)
'
' Name:         txtInfo_Change
' Description:  Trigger the Validate logic if needed.
'

    If Not Populating Then
        Dim Cancel As Boolean
        If Index = tiAutosave Then Call txtInfo_Validate(Index, Cancel)
    End If
    
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
'
' Name:         txtInfo_GotFocus
' Description:  Highlight the text.
'

    SelectText txtInfo(Index)

End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
'
' Name:         txtInfo_Validate
' Description:  Store the information entered.
'

    Dim Value As String
    Value = TrimWhiteSpace(txtInfo(Index))
    txtInfo(Index).Text = Value

    Select Case Index
        Case tiTitle
            If Value <> Game.ChronicleTitle Then
                Game.ChronicleTitle = Value
                If Value = "" Then
                    mdiMain.Caption = GrapevineCaption
                Else
                    mdiMain.Caption = GrapevineCaption & ": " & Value
                End If
                Game.DataChanged = True
            End If
        Case tiWebPage
            If Game.Website <> Value Then
                Game.Website = Value
                Game.DataChanged = True
            End If
        Case tiEMail
            If Game.EMail <> Value Then
                Game.EMail = Value
                Game.DataChanged = True
            End If
        Case tiPhone
            If Game.Phone <> Value Then
                Game.Phone = Value
                Game.DataChanged = True
            End If
        Case tiUsualTime
            If Game.UsualTime <> Value Then
                Game.UsualTime = Value
                Game.DataChanged = True
            End If
        Case tiUsualPlace
            If Game.UsualPlace <> Value Then
                Game.UsualPlace = Value
                Game.DataChanged = True
            End If
        Case tiDescription
            If Game.Description <> Value Then
                Game.Description = Value
                Game.DataChanged = True
            End If
        Case tiDateTime
            Game.Calendar.MoveTo CDate(lblDate.Caption)
            If Game.Calendar.GetGameTime <> Value Then
                Game.Calendar.SetGameTime Value
                Game.Calendar.LastModified = Now
                Game.DataChanged = True
            End If
        Case tiDatePlace
            Game.Calendar.MoveTo CDate(lblDate.Caption)
            If Game.Calendar.GetGamePlace <> Value Then
                Game.Calendar.SetGamePlace Value
                Game.Calendar.LastModified = Now
                Game.DataChanged = True
            End If
        Case tiDateNotes
            Game.Calendar.MoveTo CDate(lblDate.Caption)
            If Game.Calendar.GetGameNotes <> Value Then
                Game.Calendar.SetGameNotes Value
                Game.Calendar.LastModified = Now
                Game.DataChanged = True
            End If
        Case tiHideStart
            If Game.STCommentStart <> Value Then
                Game.STCommentStart = Value
                Game.DataChanged = True
            End If
        Case tiHideEnd
            If Game.STCommentEnd <> Value Then
                Game.STCommentEnd = Value
                Game.DataChanged = True
            End If
        Case tiAutosave
            If Int(Val(Value)) < 1 Then txtInfo(tiAutosave).Text = "1"
            If chkPreferences(ciAutosave).Value Then
                    mdiMain.AutosaveTime = Int(Val(txtInfo(tiAutosave).Text))
            End If
    End Select
            
End Sub

Private Sub chkRumors_Click(Index As Integer)
'
' Name:         chkRumors_Click
' Description:  Record the new standard rumor settings.
'

    If Not Populating Then
    
        With Game.APREngine
        
            .PublicRumors = (chkRumors(ciPublic).Value = vbChecked)
            .PersonalRumors = (chkRumors(ciPersonal).Value = vbChecked)
            .RaceRumors = (chkRumors(ciRace).Value = vbChecked)
            .GroupRumors = (chkRumors(ciGroup).Value = vbChecked)
            .SubgroupRumors = (chkRumors(ciSubgroup).Value = vbChecked)
            .InfluenceRumors = (chkRumors(ciInfluence).Value = vbChecked)
            .PreviousRumors = (chkRumors(ciPrevious).Value = vbChecked)
            .CopyPrevious = (chkRumors(ciCopy).Value = vbChecked)
            .AddCommon = (chkRumors(ciCommon).Value = vbChecked)
            .CarryUnused = (chkRumors(ciUnused).Value = vbChecked)
        
        End With
    
        Game.DataChanged = True
    
    End If

End Sub

Private Sub cmdAddBackground_Click()
'
' Name:         cmdAddBackground
' Description:  Add a new background to the list of actionable backgrounds.
'

    Dim NewBackground As String

    NewBackground = InputBox("Enter a background that characters can use for actions:", _
                    "Actionable Background")
    
    NewBackground = Trim(NewBackground)
    
    If NewBackground <> "" Then
    
        lstBackgrounds.AddItem NewBackground
        Game.APREngine.BackgroundActions.Insert NewBackground
        Game.DataChanged = True
        
    End If
    
End Sub

Private Sub cmdDeleteBackground_Click()
'
' Name:         cmdDeleteBackground
' Description:  Delete a background from the list of actionable backgrounds.
'

    If lstBackgrounds.ListIndex > -1 Then
    
        Game.APREngine.BackgroundActions.MoveTo lstBackgrounds.Text
        Game.APREngine.BackgroundActions.RemoveTrait
        lstBackgrounds.RemoveItem lstBackgrounds.ListIndex
        If lstBackgrounds.ListCount > 0 Then lstBackgrounds.ListIndex = 0
        Game.DataChanged = True
        
    End If

End Sub

Private Sub lvwLevels_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
' Name:         lvwLevels_ItemClick
' Description:  Move to a new action level, making it available to edit.
'
    
    With Game.APREngine.ActionsPerLevel
        .MoveToPlace (Item.Index - 1)
        If Not .Off Then
            Populating = True
            txtValue.Text = .Trait.Total
            Populating = False
        End If
    End With

End Sub

Private Sub txtPersonal_Change()
'
' Name:         txtPersonal_Change
' Description:  Record the new number of personal actions.
'

    If Not Populating Then
        Game.APREngine.PersonalActions = Val(txtPersonal.Text)
        Game.DataChanged = True
    End If
    
End Sub

Private Sub txtPersonal_GotFocus()
'
' Name:         txtPersonal_GotFocus
' Description:  Select the text upon receiving focus.
'
    SelectText txtPersonal

End Sub

Private Sub txtValue_Change()
'
' Name:         txtValue_Change
' Description:  Record the new action value for this level of influence.
'
    
    If Not (Populating Or lvwLevels.SelectedItem Is Nothing) Then
    
        With Game.APREngine.ActionsPerLevel
            .MoveToPlace (lvwLevels.SelectedItem.Index - 1)
            If Not .Off Then
                .Trait.Total = Val(txtValue.Text)
                lvwLevels.SelectedItem.ListSubItems(1).Text = .Trait.Total & " Actions"
            End If
        End With

    End If

End Sub

Private Sub txtValue_GotFocus()
'
' Name:         txtValue_GotFocus
' Description:  Select the text upon receiving focus.
'
    SelectText txtValue
    
End Sub
