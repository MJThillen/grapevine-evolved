VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOutput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print"
   ClientHeight    =   6120
   ClientLeft      =   210
   ClientTop       =   570
   ClientWidth     =   9060
   Icon            =   "frmOutput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   9060
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Height          =   4335
      Index           =   1
      Left            =   360
      TabIndex        =   33
      Top             =   1320
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CheckBox chkAttach 
         Caption         =   "Send HTML and Text output as Attachments"
         Height          =   255
         Left            =   360
         TabIndex        =   90
         Top             =   3960
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.CheckBox chkShellPrint 
         Caption         =   "Save output to file and print with the associated Windows program"
         Height          =   255
         Left            =   360
         TabIndex        =   48
         Top             =   3960
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.CheckBox chkAllowHTML 
         Caption         =   "Allow HTML tags in text"
         Height          =   375
         Left            =   360
         TabIndex        =   47
         Top             =   3600
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.CheckBox chkSTData 
         Caption         =   "Conceal all text marked ST Only"
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   3360
         Value           =   1  'Checked
         Width           =   5055
      End
      Begin VB.OptionButton optAscend 
         Caption         =   "Reverse Order"
         Height          =   375
         Index           =   1
         Left            =   3240
         TabIndex        =   53
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton optAscend 
         Caption         =   "C&hronological Order"
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   51
         Top             =   960
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Rumors"
         Height          =   375
         Index           =   6
         Left            =   3960
         TabIndex        =   50
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Actions"
         Height          =   375
         Index           =   5
         Left            =   2160
         TabIndex        =   49
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Locations"
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   44
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Rotes"
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   43
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Equipment"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   42
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "XP Histories"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   41
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Notes"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   40
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.ComboBox cboRange 
         Height          =   315
         ItemData        =   "frmOutput.frx":058A
         Left            =   1440
         List            =   "frmOutput.frx":05A5
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblLabels 
         Caption         =   "Formatting"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   45
         Top             =   3000
         Width           =   5415
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Date Range"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   52
         Top             =   645
         Width           =   975
      End
      Begin VB.Label lblOptions 
         Caption         =   "&Include these:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1440
         Width           =   5415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Histories, Game Calendar, Plot Developments"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   3
      Left            =   360
      TabIndex        =   54
      Top             =   1320
      Visible         =   0   'False
      Width           =   5895
      Begin VB.ComboBox cboSearch 
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   720
         Width           =   2655
      End
      Begin VB.OptionButton optMatch 
         Caption         =   "List characters that don't match:"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   74
         Top             =   840
         Width           =   2655
      End
      Begin VB.OptionButton optMatch 
         Caption         =   "List characters that match:"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   73
         Top             =   600
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.Frame fraStatistics 
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   0
         TabIndex        =   70
         Top             =   1320
         Width           =   5895
         Begin VB.ComboBox cboKey 
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   900
            Width           =   2655
         End
         Begin VB.OptionButton optGraph 
            Caption         =   "Distribution"
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   78
            ToolTipText     =   "Graph the range of one value for the characters: Clan or Tribe populations, distribution of Trait totals, etc."
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optGraph 
            Caption         =   "Maxima"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   77
            ToolTipText     =   "Graph the highest value of each type of Trait in a given category: Find the highest levels of Influences, Backgrounds, etc."
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox chkZero 
            Caption         =   "Exclude &zero and (none)"
            Height          =   675
            Left            =   4440
            TabIndex        =   76
            Top             =   705
            Width           =   1215
         End
         Begin VB.OptionButton optGraph 
            Caption         =   "Sums"
            Height          =   255
            Index           =   2
            Left            =   4440
            TabIndex        =   75
            ToolTipText     =   "Graph the sum of all the Traits of one category for the characters: Find the total levels of Influence in the game."
            Top             =   480
            Width           =   1215
         End
         Begin VB.Frame fraTraitDistribution 
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   1680
            TabIndex        =   80
            Top             =   1200
            Visible         =   0   'False
            Width           =   2775
            Begin VB.OptionButton optTraitDistribution 
               Caption         =   "Total traits in the &category"
               Height          =   315
               Index           =   0
               Left            =   360
               TabIndex        =   84
               Top             =   45
               Width           =   2295
            End
            Begin VB.OptionButton optTraitDistribution 
               Caption         =   "Distinct traits in the category"
               Height          =   255
               Index           =   1
               Left            =   360
               TabIndex        =   83
               Top             =   360
               Width           =   2295
            End
            Begin VB.TextBox txtTrait 
               Height          =   315
               Left            =   600
               TabIndex        =   82
               Text            =   "(specific trait)"
               Top             =   600
               Width           =   2055
            End
            Begin VB.OptionButton optTraitDistribution 
               Height          =   255
               Index           =   2
               Left            =   360
               TabIndex        =   81
               Top             =   630
               Width           =   255
            End
            Begin VB.Line linLink 
               Index           =   0
               X1              =   495
               X2              =   225
               Y1              =   750
               Y2              =   750
            End
            Begin VB.Line linLink 
               Index           =   1
               X1              =   240
               X2              =   240
               Y1              =   0
               Y2              =   750
            End
            Begin VB.Line linLink 
               Index           =   2
               X1              =   495
               X2              =   240
               Y1              =   180
               Y2              =   180
            End
            Begin VB.Line linLink 
               Index           =   3
               X1              =   495
               X2              =   240
               Y1              =   480
               Y2              =   480
            End
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "&Data to Graph:"
            Height          =   255
            Index           =   13
            Left            =   360
            TabIndex        =   86
            Top             =   945
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Graph &Type:"
            Height          =   255
            Index           =   14
            Left            =   360
            TabIndex        =   85
            Top             =   495
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Statistics"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   71
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.Label lblLabels 
         Caption         =   "Search"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   55
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame fraDestination 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   2
      Left            =   6720
      TabIndex        =   30
      Top             =   2880
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CommandButton cmdEMailSetup 
         Caption         =   "E-Mail &Setup..."
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdOutput 
         Caption         =   "&E-Mail..."
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame fraDestination 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   0
      Left            =   6720
      TabIndex        =   22
      Top             =   2880
      Visible         =   0   'False
      Width           =   1935
      Begin VB.TextBox txtCopies 
         Height          =   285
         Left            =   720
         TabIndex        =   88
         Text            =   "1"
         Top             =   240
         Width           =   540
      End
      Begin VB.CommandButton cmdOutput 
         Caption         =   "&Print"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "Printer &Setup..."
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   1695
      End
      Begin MSComCtl2.UpDown updCopies 
         Height          =   285
         Left            =   1260
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   240
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtCopies"
         BuddyDispid     =   196631
         OrigLeft        =   1260
         OrigTop         =   240
         OrigRight       =   1815
         OrigBottom      =   525
         Max             =   99
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblPrintHTML 
         Alignment       =   2  'Center
         Caption         =   "Grapevine asks your web browser to print HTML.  Choose printer and page setup options there."
         Height          =   855
         Left            =   0
         TabIndex        =   89
         Top             =   1080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblCopies 
         Caption         =   "Copi&es:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   285
         Width           =   525
      End
   End
   Begin VB.Frame fraDestination 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   1
      Left            =   6720
      TabIndex        =   28
      Top             =   2880
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CommandButton cmdOutput 
         Caption         =   "&Save"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   5895
      Begin VB.CommandButton cmdSelectSame 
         Caption         =   "Select Same &Date"
         Height          =   375
         Index           =   0
         Left            =   4080
         TabIndex        =   7
         Top             =   2160
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdSelectSame 
         Caption         =   "Se&lect Same Name"
         Height          =   375
         Index           =   1
         Left            =   4080
         TabIndex        =   8
         Top             =   2640
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkNot 
         Caption         =   "NO&T"
         Height          =   255
         Left            =   4080
         TabIndex        =   6
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ListBox lstList 
         Columns         =   2
         Height          =   3765
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   2
         Top             =   480
         Width           =   3735
      End
      Begin VB.ComboBox cboSelectOnly 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Select &All"
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdSelectOnly 
         Caption         =   "Select &Only:"
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdSelectNone 
         Caption         =   "Select &None"
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   1
         Left            =   3600
         Picture         =   "frmOutput.frx":060E
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Caption         =   "Characters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   1
         Top             =   120
         Width           =   3015
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "frmOutput.frx":0B98
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   2
      Left            =   360
      TabIndex        =   34
      Top             =   1320
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txtFormat 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   2
         Left            =   3855
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   67
         Top             =   1800
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.CommandButton cmdFormat 
         Height          =   375
         Index           =   2
         Left            =   3360
         Picture         =   "frmOutput.frx":1122
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtFormat 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   1
         Left            =   3855
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   64
         Top             =   840
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.CommandButton cmdFormat 
         Height          =   375
         Index           =   1
         Left            =   3360
         Picture         =   "frmOutput.frx":16AC
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtFormat 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   0
         Left            =   3855
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   61
         Top             =   2760
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.CommandButton cmdFormat 
         Height          =   375
         Index           =   0
         Left            =   3360
         Picture         =   "frmOutput.frx":1C36
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   2760
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdAddReport 
         Caption         =   "&Add Report"
         Height          =   375
         Left            =   120
         TabIndex        =   58
         Top             =   3960
         Width           =   1695
      End
      Begin VB.CommandButton cmdDeleteReport 
         Caption         =   "&Delete Report"
         Height          =   375
         Left            =   2040
         TabIndex        =   57
         Top             =   3960
         Width           =   1695
      End
      Begin VB.ListBox lstTemplates 
         Height          =   3180
         Left            =   600
         Sorted          =   -1  'True
         TabIndex        =   56
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Associated Files"
         Height          =   195
         Index           =   11
         Left            =   3360
         TabIndex        =   69
         Top             =   240
         Width           =   2370
      End
      Begin VB.Label lblLabels 
         Caption         =   "Reports and Character Sheets"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   68
         Top             =   240
         Width           =   2730
      End
      Begin VB.Label lblFormatLabel 
         Caption         =   "&HTML Template File"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   65
         Top             =   1560
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblFormatLabel 
         Caption         =   "Rich Te&xt (RTF) Template File"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   62
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblFormatLabel 
         Caption         =   "Plain Text Te&mplate File"
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   59
         Top             =   2520
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.ComboBox cboDate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   35
      Top             =   480
      Width           =   2655
   End
   Begin MSComctlLib.ProgressBar pgbProgress 
      Height          =   285
      Left            =   6630
      TabIndex        =   26
      Top             =   480
      Visible         =   0   'False
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.OptionButton optFormat 
      Caption         =   "Plain Text"
      Height          =   255
      Index           =   0
      Left            =   6840
      TabIndex        =   18
      Top             =   2040
      Width           =   1455
   End
   Begin VB.OptionButton optFormat 
      Caption         =   "HTML"
      Height          =   375
      Index           =   2
      Left            =   6840
      TabIndex        =   17
      Top             =   1680
      Width           =   1455
   End
   Begin VB.OptionButton optFormat 
      Caption         =   "Rich Text (RTF)"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   16
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox cboTemplate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   6840
      TabIndex        =   10
      Top             =   5280
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cmnDialog 
      Left            =   8040
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "txt"
      Filter          =   "Plain Text Files (*.txt)|*.txt|Rich Text Files (*.rtf)|*.rtf|HTML Files (*.html)|*.html|All Files (*.*)|*.*"
   End
   Begin RichTextLib.RichTextBox rtfEasel 
      Height          =   495
      Left            =   6840
      TabIndex        =   21
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmOutput.frx":21C0
   End
   Begin MSComctlLib.TabStrip tabTabs 
      Height          =   4935
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8705
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dates && &Options"
            Key             =   "Options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Templates"
            Key             =   "Templates"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblLabels 
      Caption         =   "For this &Game Date:"
      Height          =   255
      Index           =   6
      Left            =   3720
      TabIndex        =   36
      Top             =   240
      Width           =   2655
   End
   Begin VB.Image imgCue 
      Height          =   480
      Left            =   240
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   " Print "
      Height          =   195
      Index           =   2
      Left            =   6720
      TabIndex        =   19
      Top             =   2670
      Width           =   405
   End
   Begin VB.Label lblLabels 
      Caption         =   " &Format "
      Height          =   195
      Index           =   0
      Left            =   6720
      TabIndex        =   14
      Top             =   1170
      Width           =   570
   End
   Begin VB.Label lblLabels 
      Caption         =   "Select &what to Print:"
      Height          =   255
      Index           =   5
      Left            =   960
      TabIndex        =   13
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Index           =   1
      Left            =   6600
      TabIndex        =   15
      Top             =   1260
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Index           =   3
      Left            =   6600
      TabIndex        =   20
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lblPercent 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100%"
      Height          =   555
      Left            =   6600
      TabIndex        =   27
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LIST_PLAYERS = "Players"
Private Const LIST_CHARACTERS = "Characters"
Private Const LIST_ITEMS = "Items"
Private Const LIST_ROTES = "Rotes"
Private Const LIST_LOCATIONS = "Locations"
Private Const LIST_ACTIONS = "Actions"
Private Const LIST_PLOTS = "Plots"
Private Const LIST_RUMORS = "Rumors"
Private Const LIST_OPTIONS = "Options"
Private Const LIST_TEMPLATES = "Templates"
Private Const LIST_SEARCH = "Search"

Private Const OUTPUT_PRINT = 0
Private Const OUTPUT_SAVE = 1
Private Const OUTPUT_EMAIL = 2

Private Const FRAME_SELECT = 0
Private Const FRAME_OPTION = 1
Private Const FRAME_TEMPLATE = 2
Private Const FRAME_SEARCH = 3

Private Const SAME_DATE = 0
Private Const SAME_NAME = 1

Private Const OPT_NOTES = 0
Private Const OPT_HISTORY = 1
Private Const OPT_ITEMS = 2
Private Const OPT_ROTES = 3
Private Const OPT_LOCATIONS = 4
Private Const OPT_ACTIONS = 5
Private Const OPT_RUMORS = 6

Private Const OPT_MATCH = 0
Private Const OPT_NOT = 1

Private Const OPT_ASCEND = 0
Private Const OPT_DESCEND = 1

Private Const OPT_DIST = 0
Private Const OPT_MAXIMA = 1
Private Const OPT_SUM = 2

Private Const OPT_TOTAL = 0
Private Const OPT_DISTINCT = 1
Private Const OPT_SPECIFIC = 2

Private Const CharSheetIndex = 0

Private CharSheetOptions  As Long
Private OutputIndex As Long
Private OtherTemplateIndex As Long
Private CustomTemplate As TemplateClass

Public Sub ShowOutput(Device As OutputDeviceType)
'
' Name:         ShowOutput
' Parameters:   OFormat     Print, file or E-Mail
'               Template    Name of template to prepare for
' Description:  Prepare this window for the appropriate type of output.
'
    
    Dim FindIndex As Integer
    
    Select Case Device
        Case odPrinter
            Me.Caption = "Print"
            Me.Icon = mdiMain.imlSmallIcons.ListImages("Printer").Picture
            imgCue.Picture = mdiMain.imlIcons.ListImages("Printer").Picture
            lblLabels(2).Caption = "Print"
            lblLabels(5).Caption = "Select &what to Print:"
            lblOptions.Caption = "&Include these with each printout:"
            OutputIndex = OUTPUT_PRINT
        Case odFile
            Me.Caption = "Save to File"
            Me.Icon = mdiMain.imlSmallIcons.ListImages("Document").Picture
            lblLabels(2).Caption = "Save to File"
            imgCue.Picture = mdiMain.imlIcons.ListImages("Document").Picture
            lblLabels(5).Caption = "Select &what to Save to File:"
            lblOptions.Caption = "&Include these with each document:"
            OutputIndex = OUTPUT_SAVE
        Case odEMail
            Me.Caption = "E-Mail"
            Me.Icon = mdiMain.imlSmallIcons.ListImages("EMail").Picture
            lblLabels(2).Caption = "E-Mail"
            imgCue.Picture = mdiMain.imlIcons.ListImages("EMail").Picture
            lblLabels(5).Caption = "Select &what to E-Mail:"
            lblOptions.Caption = "&Include these with each e-mail:"
            OutputIndex = OUTPUT_EMAIL
    End Select

    Select Case CLng(GetSetting(App.Title, "Output", "Format", ofRTF))
        Case ofText:    optFormat(ofText).Value = True
        Case ofRTF:     optFormat(ofRTF).Value = True
        Case ofHTML:    optFormat(ofHTML).Value = True
    End Select
    
    With Game.TemplateList
        FindIndex = -1
        cboTemplate.AddItem "Character Sheets"
        .First
        Do Until .Off
            lstTemplates.AddItem .Item.Name
            If Not .Item.IsCharacterSheet Then cboTemplate.AddItem .Item.Name
            If .Item.Name = OutputEngine.Template Then FindIndex = cboTemplate.NewIndex
            .MoveNext
        Loop
        cboTemplate.AddItem "Choose Template..."
        OtherTemplateIndex = cboTemplate.NewIndex
        If FindIndex = -1 Then FindIndex = CharSheetIndex
        cboTemplate.ListIndex = FindIndex
    End With

    Set tabTabs.SelectedItem = tabTabs.Tabs(1)

    With Game.Calendar
        .First
        Do Until .Off
            cboDate.AddItem Format(.GetGameDate, "mmmm d, yyyy")
            .MoveNext
        Loop
        If OutputEngine.GameDate = 0 Then
            .MoveToCloseGame
            If Not .Off Then OutputEngine.GameDate = .GetGameDate
        End If
        If Not OutputEngine.GameDate = 0 Then
            cboDate.Text = Format(OutputEngine.GameDate, "mmmm d, yyyy")
        End If
    End With

    With Game.QueryEngine.QueryList
        .First
        Do Until .Off
            If .Item.Inventory = qiCharacters Then cboSearch.AddItem .Item.Name
            If .Item.Name = OutputEngine.SearchName Then cboSearch.ListIndex = cboSearch.NewIndex
            .MoveNext
        Loop
    End With

    fraDestination(OutputIndex).Visible = True
    cboRange.ListIndex = CLng(GetSetting(App.Title, "Output", "Range", "0"))
    optAscend(IIf(CBool(GetSetting(App.Title, "Output", "Reverse", True)), OPT_ASCEND, OPT_DESCEND)).Value = True
    optMatch(IIf(OutputEngine.SearchNot, OPT_NOT, OPT_MATCH)).Value = True
    chkZero.Value = IIf(OutputEngine.OKZero, vbUnchecked, vbChecked)
    txtTrait.Text = OutputEngine.StatTrait
    
    chkSTData.Value = CInt(GetSetting(App.Title, "Output", "HideSTData", vbChecked))
    chkAllowHTML.Value = CInt(GetSetting(App.Title, "Output", "AllowHTML", vbUnchecked))
    If Device = odPrinter Then chkShellPrint.Value = CInt(GetSetting(App.Title, "Output", "ShellPrint", vbUnchecked))
    chkShellPrint.Visible = (Device = odPrinter)
    If Device = odEMail Then chkAttach.Value = CInt(GetSetting(App.Title, "Output", "Attach", vbUnchecked))
    chkAttach.Visible = (Device = odEMail)
    CharSheetOptions = CLng(GetSetting(App.Title, "Output", "Options", ooOptNotes))
    CheckOptions CharSheetOptions
    
    Me.Show vbModal, mdiMain
    
End Sub

Private Function GetOptions() As Long
'
' Name:         GetOptions
' Description:  Return the sum of the current selected output options.
'

    GetOptions = IIf(chkOptions(OPT_NOTES).Value = vbChecked, ooOptNotes, ooNone) Or _
                 IIf(chkOptions(OPT_HISTORY).Value = vbChecked, ooOptHistory, ooNone) Or _
                 IIf(chkOptions(OPT_ITEMS).Value = vbChecked, ooOptitems, ooNone) Or _
                 IIf(chkOptions(OPT_ROTES).Value = vbChecked, ooOptRotes, ooNone) Or _
                 IIf(chkOptions(OPT_LOCATIONS).Value = vbChecked, ooOptLocations, ooNone) Or _
                 IIf(chkOptions(OPT_ACTIONS).Value = vbChecked, ooOptActions, ooNone) Or _
                 IIf(chkOptions(OPT_RUMORS).Value = vbChecked, ooOptRumors, ooNone)

End Function

Private Sub CheckOptions(OutOptions As Long)
'
' Name:         CheckOptions
' Description:  Checkmark the selected options.
'

    chkOptions(OPT_NOTES).Value = IIf(OutOptions And ooOptNotes, vbChecked, vbUnchecked)
    chkOptions(OPT_HISTORY).Value = IIf(OutOptions And ooOptHistory, vbChecked, vbUnchecked)
    chkOptions(OPT_ITEMS).Value = IIf(OutOptions And ooOptitems, vbChecked, vbUnchecked)
    chkOptions(OPT_ROTES).Value = IIf(OutOptions And ooOptRotes, vbChecked, vbUnchecked)
    chkOptions(OPT_LOCATIONS).Value = IIf(OutOptions And ooOptLocations, vbChecked, vbUnchecked)
    chkOptions(OPT_ACTIONS).Value = IIf(OutOptions And ooOptActions, vbChecked, vbUnchecked)
    chkOptions(OPT_RUMORS).Value = IIf(OutOptions And ooOptRumors, vbChecked, vbUnchecked)

End Sub

Private Sub SelectList(ByRef Box As ListBox, ToSelect As Boolean)
'
' Name:         Selectlist
' Parameters:   Box         the listbox to (de)select
'               ToSelect    whether to select or deselect
' Description:  Select or deselect all items in a listview.
'

    Dim I As Long
    
    For I = 0 To Box.ListCount - 1
        Box.Selected(I) = ToSelect
    Next I

End Sub

Private Sub StoreSelection()
'
' Name:         StoreSelection
' Description:  Store a selection before moving to a new list.
'

    Dim SetIndex As Long
    
    Select Case lstList.Tag
        Case LIST_PLAYERS:            SetIndex = osPlayers
        Case LIST_CHARACTERS:         SetIndex = osCharacters
        Case LIST_ITEMS:              SetIndex = osItems
        Case LIST_ROTES:              SetIndex = osRotes
        Case LIST_LOCATIONS:          SetIndex = osLocations
        Case LIST_ACTIONS:            SetIndex = osActions
        Case LIST_PLOTS:              SetIndex = osPlots
        Case LIST_RUMORS:             SetIndex = osRumors
        Case Else:                    Exit Sub
    End Select

    OutputEngine.SelectSet(SetIndex).StoreListBox lstList

End Sub

Private Sub cboDate_Change()
'
' Name:         cboDate_Change
' Description:  Set the dates for the outputengine.
'
    If IsDate(cboDate.Text) Then
        OutputEngine.GameDate = CDate(cboDate.Text)
        
        Call cboRange_Click
    End If
    
End Sub

Private Sub cboDate_Click()
'
' Name:         cboDate_Click
' Description:  Set the dates for the outputengine.
'
    Call cboDate_Change
    
End Sub

Private Sub cboKey_Click()
'
' Name:         cboKey_Click
' Description:  Show the trait distribution fields, if needed.
'

    fraTraitDistribution.Visible = (cboKey.ItemData(cboKey.ListIndex) = qtTraitList) And _
        optGraph(OPT_DIST).Value
    If Not fraTraitDistribution.Visible Then optTraitDistribution(OPT_TOTAL).Value = True
    OutputEngine.StatKey = Game.QueryEngine.TitlesToKeys(cboKey.Text)

End Sub

Private Sub cboRange_Click()
'
' Name:         cboRange_Click
' Description:  Update the OutputEngine with new dates.
'

    OutputEngine.EndDate = OutputEngine.GameDate
    If cboRange.ListIndex = 0 Then
        OutputEngine.StartDate = 0
        OutputEngine.EndDate = 0
    ElseIf cboRange.ListIndex > 0 Then
        OutputEngine.StartDate = OutputEngine.EndDate - CLng(cboRange.ItemData(cboRange.ListIndex))
    End If

End Sub

Private Sub cboSearch_Click()
'
' Name:         cboSearch_Click
' Description:  Update the OutputEngine's search setting.
'
    OutputEngine.SearchName = cboSearch.Text

End Sub

Private Sub cboTemplate_Click()
'
' Name:         cboTemplate_Click
' Description:  Choose a new template.  Examine its subject, and create the needed selection tabs.
'

    Dim I As Integer
    Dim OldKey As String
    Dim Subject As Long
    Dim Template As TemplateClass
    Dim OutFormat As OutputFormatType
    Dim NoFile As Boolean
    
    Subject = ooUnknown
    OutFormat = ofText
    If optFormat(ofHTML).Value Then OutFormat = ofHTML
    If optFormat(ofRTF).Value Then OutFormat = ofRTF
        
    If cboTemplate.ListIndex = OtherTemplateIndex Then
            
        With cmnDialog
            .DialogTitle = "Choose Template"
            .FileName = ""
            .DefaultExt = IIf(OutFormat = ofText, "txt", IIf(OutFormat = ofRTF, "rtf", "html"))
            .FilterIndex = OutFormat + 1
            .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
        End With
        
        On Error Resume Next
        cmnDialog.ShowOpen
        If Err = 0 Then
            CustomTemplate.SetFilename ofRTF, cmnDialog.FileName
            CustomTemplate.SetFilename ofHTML, cmnDialog.FileName
            CustomTemplate.SetFilename ofText, cmnDialog.FileName
            If cboTemplate.ListCount = OtherTemplateIndex + 1 Then cboTemplate.AddItem ""
            cboTemplate.List(OtherTemplateIndex + 1) = cmnDialog.FileTitle
            cboTemplate.ListIndex = OtherTemplateIndex + 1
            Select Case cmnDialog.FilterIndex - 1
                Case ofRTF:     optFormat(ofRTF).Value = True
                Case ofHTML:    optFormat(ofHTML).Value = True
                Case Else:      optFormat(ofText).Value = True
            End Select
            Exit Sub
        Else
            cboTemplate.ListIndex = CharSheetIndex
            Exit Sub
        End If
        On Error GoTo 0

    End If
    
    I = 1
    StoreSelection
    Do Until I > tabTabs.Tabs.Count
        If tabTabs.Tabs(I).Key = LIST_OPTIONS Or tabTabs.Tabs(I).Key = LIST_TEMPLATES Then
            I = I + 1
        Else
            tabTabs.Tabs.Remove I
        End If
    Loop
    
    Select Case cboTemplate.ListIndex
        Case CharSheetIndex
            Subject = ooCharacters Or ooOptionMask
            CheckOptions CharSheetOptions
            OutputEngine.Template = cboTemplate.Text
        Case OtherTemplateIndex + 1
            Subject = CustomTemplate.GetSubject(MIN_OUTFORMAT)
            If Not Subject = ooFileError Then CheckOptions Subject
        Case Else
            Game.TemplateList.MoveTo cboTemplate.Text
            If Game.TemplateList.Off Then
                Subject = ooFileError
            Else
                Set Template = Game.TemplateList.Item
                Subject = Template.GetSubject(OutFormat)
                NoFile = (Template.GetFilename(OutFormat) = "")
                If Not Subject = ooFileError Then CheckOptions Subject
                OutputEngine.Template = cboTemplate.Text
            End If
    End Select

    If Subject = ooFileError Then
        
        For I = 0 To lstTemplates.ListCount - 1
            If lstTemplates.List(I) = cboTemplate.Text Then
                lstTemplates.ListIndex = I
                Exit For
            End If
        Next I
        Subject = ooNone
        fraDestination(OutputIndex).Visible = False
        Set tabTabs.SelectedItem = tabTabs.Tabs(2)
        If NoFile Then
            MsgBox "This format is not available for this report.", vbInformation, _
                   "Unavailable Format"
        Else
            MsgBox "Grapevine can't read the file associated with this report and format." & vbCrLf & _
                   "Please make sure the file name and location is correct.", vbInformation, _
                   "Can't Read Template File"
        End If
        
    Else

        If Subject And ooStatistics Then
            tabTabs.Tabs.Add 1, LIST_SEARCH, "Search && Statistics"
            fraStatistics.Visible = True
        ElseIf Subject And ooSearch Then
            tabTabs.Tabs.Add 1, LIST_SEARCH, "Search"
            fraStatistics.Visible = False
        End If
        
        If Subject And ooRumors Then _
            tabTabs.Tabs.Add 1, LIST_RUMORS, LIST_RUMORS
        If Subject And ooPlots Then _
            tabTabs.Tabs.Add 1, LIST_PLOTS, LIST_PLOTS
        If Subject And ooActions Then _
            tabTabs.Tabs.Add 1, LIST_ACTIONS, LIST_ACTIONS
        If Subject And ooLocations Then _
            tabTabs.Tabs.Add 1, LIST_LOCATIONS, LIST_LOCATIONS
        If Subject And ooRotes Then _
            tabTabs.Tabs.Add 1, LIST_ROTES, LIST_ROTES
        If Subject And ooItems Then _
            tabTabs.Tabs.Add 1, LIST_ITEMS, LIST_ITEMS
        If Subject And ooCharacters Then _
            tabTabs.Tabs.Add 1, LIST_CHARACTERS, LIST_CHARACTERS
        If Subject And ooPlayers Then _
            tabTabs.Tabs.Add 1, LIST_PLAYERS, LIST_PLAYERS
        
        fraDestination(OutputIndex).Visible = True
        Set tabTabs.SelectedItem = tabTabs.Tabs(1)
        
    End If

    chkOptions(OPT_NOTES).Visible = (Subject And ooOptNotes)
    chkOptions(OPT_HISTORY).Visible = (Subject And ooOptHistory)
    chkOptions(OPT_ITEMS).Visible = (Subject And ooOptitems)
    chkOptions(OPT_ROTES).Visible = (Subject And ooOptRotes)
    chkOptions(OPT_LOCATIONS).Visible = (Subject And ooOptLocations)
    chkOptions(OPT_ACTIONS).Visible = (Subject And ooOptActions)
    chkOptions(OPT_RUMORS).Visible = (Subject And ooOptRumors)
    lblOptions.Visible = (Subject And ooOptionMask)
    
End Sub

Private Sub cmdAddReport_Click()
'
' Name:         cmdAddReport
' Description:  Add a new template class.
'

    Dim TemplateName As String
    Dim I As Integer
    
    TemplateName = InputBox("Enter a name for the sheet or report:", "New Sheet or Report")
    
    If TemplateName <> "" Then
        Game.TemplateList.MoveTo TemplateName
        If Game.TemplateList.Off Then
            
            Dim Template As TemplateClass
            Set Template = New TemplateClass
            Template.Name = TemplateName
            Game.TemplateList.InsertSorted Template
            lstTemplates.AddItem TemplateName
            lstTemplates.ListIndex = lstTemplates.NewIndex
            For I = CharSheetIndex + 1 To OtherTemplateIndex
                If I = OtherTemplateIndex Or cboTemplate.List(I) > TemplateName Then
                    cboTemplate.AddItem TemplateName, I
                    Exit For
                End If
            Next I
            OtherTemplateIndex = OtherTemplateIndex + 1
            Game.DataChanged = True
            
        Else
            MsgBox "That name is in use -- enter a different one.", , "Name in Use"
        End If
    End If
    
End Sub

Private Sub cmdClose_Click()
'
' Name:         cmdClose_Click
' Description:  Dismiss this window.
'

    Unload Me

End Sub

Private Sub cmdDeleteReport_Click()
'
' Name:         cmdDeleteReport
' Description:  Delete the selected template.
'

    Game.TemplateList.MoveTo lstTemplates.Text
    If Not Game.TemplateList.Off Then
        If MsgBox("Are you sure you want to delete this template from the list?", vbQuestion + vbYesNo, _
                  "Delete Template") = vbYes Then
                  
            Game.TemplateList.Remove
                  
            If cboTemplate.Text = lstTemplates.Text Then
                cboTemplate.RemoveItem cboTemplate.ListIndex
                cboTemplate.ListIndex = CharSheetIndex
                Set tabTabs.SelectedItem = tabTabs.Tabs(LIST_TEMPLATES)
            Else
                Dim I As Integer
                For I = CharSheetIndex + 1 To OtherTemplateIndex - 1
                    If cboTemplate.List(I) = lstTemplates.Text Then cboTemplate.RemoveItem I
                Next I
            End If
            
            OtherTemplateIndex = OtherTemplateIndex - 1
            lstTemplates.ListIndex = lstTemplates.ListIndex - 1
            lstTemplates.RemoveItem lstTemplates.ListIndex + 1
            Game.DataChanged = True
            
        End If
    End If
    
End Sub

Private Sub cmdEMailSetup_Click()
'
' Name:         cmdEMailSetup_Click
' Description:  Show the EMail Setup window to initialize the mailer.
'
    frmEMailSetup.ShowSetup

End Sub

Private Sub cmdFormat_Click(Index As Integer)
'
' Name:         cmdFormat_Click
' Description:  Prompt the user for the filename to associate with this format in this template.
'

    Game.TemplateList.MoveTo lstTemplates.Text
    If Not Game.TemplateList.Off Then
    
        Dim FullName As String
        Dim Path As String
        Dim ShortName As String
        
        FullName = FindFile(Game.TemplateList.Item.GetFilename(CLng(Index)))
        ShortName = ShortFile(FullName)
        Path = Left(FullName, Len(FullName) - Len(ShortName))
        
        With cmnDialog
            .DialogTitle = "Find " & lblFormatLabel(Index).Caption & " Template"
            .FileName = ShortName
            .DefaultExt = IIf(Index = ofText, "txt", IIf(Index = ofRTF, "rtf", "html"))
            .FilterIndex = Index + 1
            .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
            .InitDir = Path
        End With
        
        On Error Resume Next
        cmnDialog.ShowOpen
        If Err = 0 Then
            On Error GoTo 0
            FullName = cmnDialog.FileName
            ShortName = GetRelativeName(FullName, Game.GameFile)
            Game.TemplateList.Item.SetFilename CLng(Index), ShortName
            txtFormat(Index).Text = ShortName
            Game.DataChanged = True
            If cboTemplate.Text = lstTemplates.Text Then
                cboTemplate_Click
                Set tabTabs.SelectedItem = tabTabs.Tabs(LIST_TEMPLATES)
            End If
        End If
        On Error GoTo 0
        
    End If

End Sub

Private Sub cmdOutput_Click(Index As Integer)
'
' Name:         cmdOutput_Click
' Description:  Print the chosen documents.
'

    Dim SendTemplate As TemplateClass
    Dim ESubject As Long
    Dim ESubjLine As String
    Dim SelLabel As String
    
    StoreSelection

    Select Case cboTemplate.ListIndex
        Case CharSheetIndex
            Set SendTemplate = Nothing
            CharSheetOptions = GetOptions
            ESubject = ooCharacters
            ESubjLine = "Character Sheet"
            SaveSetting App.Title, "Output", "Options", CharSheetOptions
        Case OtherTemplateIndex + 1
            Set SendTemplate = CustomTemplate
            ESubject = SendTemplate.GetSubject(OutputEngine.OutputFormat)
            ESubjLine = SendTemplate.Name
        Case Else
            Game.TemplateList.MoveTo cboTemplate.Text
            If Not Game.TemplateList.Off Then
                Set SendTemplate = Game.TemplateList.Item
                ESubject = SendTemplate.GetSubject(OutputEngine.OutputFormat)
                ESubjLine = SendTemplate.Name
            Else
                Exit Sub
            End If
    End Select

    OutputEngine.OutputDevice = Index
    OutputEngine.OutputOptions = GetOptions
    OutputEngine.HideSTData = (chkSTData.Value = vbChecked)
    OutputEngine.AllowHTML = (chkAllowHTML.Value = vbChecked)
    OutputEngine.PrintAfterSave = ((Index = odPrinter) And chkShellPrint.Value = vbChecked)
    OutputEngine.AlwaysAttach = (chkAttach.Value = vbChecked)
    
' Include the SMTP.ocx control in the project and add an instance to mdiMain,
' naming it smtpMailer, in order to re-enable the code below. Note that SMTP.ocx
' is detected (inaccurately) as a virus component by many antivirus programs.

'    If Index = odEMail Then
'
'        If Not mdiMain.InitializeSMTP Then
'            MsgBox "First you need to set up your outgoing E-Mail access.", vbOKOnly, "E-Mail"
'            Call cmdEMailSetup_Click
'            Exit Sub
'        Else
'            OutputEngine.InitializeMessage ESubject, ESubjLine
'            frmEMailAddressing.ShowAddressing OutputEngine.Mailer, ESubject, cboTemplate.Text
'            If OutputEngine.Mailer.Tag <> "" Then Exit Sub
'        End If
'
'    End If
'
'    Set OutputEngine.CDialog = cmnDialog
'    Set OutputEngine.RTFBox = rtfEasel
'
'    lblPercent.Caption = "0%"
'    lblPercent.Visible = True
'    pgbProgress.Visible = True
'    Me.Refresh
'
'    Screen.MousePointer = vbHourglass
'    OutputEngine.Output SendTemplate, lblPercent
'    Screen.MousePointer = vbDefault
'
'    lblPercent.Visible = False
'    pgbProgress.Visible = False
'    SaveSetting App.Title, "Output", "Range", cboRange.ListIndex
'    SaveSetting App.Title, "Output", "Ascend", optAscend(OPT_ASCEND).Value
'    SaveSetting App.Title, "Output", "HideSTData", chkSTData.Value
'    If optFormat(ofHTML).Value Then SaveSetting App.Title, "Output", "AllowHTML", chkAllowHTML.Value
'    If OutputIndex = OUTPUT_PRINT Then SaveSetting App.Title, "Output", "ShellPrint", chkShellPrint.Value
'    If OutputIndex = OUTPUT_EMAIL Then SaveSetting App.Title, "Output", "Attach", chkAttach.Value
'
'    If OutputEngine.ErrorFlag Then
'        MsgBox OutputEngine.ErrorMessage, vbExclamation, "Output Error"
'    End If

End Sub

Private Sub cmdPrintSetup_Click()
'
' Name:         cmdPrintSetup_Click
' Description:  Display the system's Print Setup dialog box.
'
    
    Dim Device As Printer
    
    On Error Resume Next
    
    With cmnDialog
        .Copies = Int(Val(txtCopies.Text))
        .DialogTitle = "Printer Setup"
        .Flags = cdlPDPrintSetup + cdlPDReturnDC + cdlPDNoPageNums + cdlPDNoSelection
        .ShowPrinter
    End With

    If Err.Number = 0 Then
        txtCopies.Text = CStr(cmnDialog.Copies)
        For Each Device In Printers
            If Device.hdc = cmnDialog.hdc Then Set Printer = Device
        Next Device
    End If

    On Error GoTo 0

End Sub

Private Sub cmdSelectAll_Click()
'
' Name:         cmdSelectAll
' Description:  Select all items in the list.
'
    SelectList lstList, True
    
End Sub

Private Sub cmdSelectSame_Click(Index As Integer)
'
' Name:         cmdSelectSame_Click
' Description:  Select all actions or rumors that share their date or name with the
'               currently selected actions or rumors.
'

    Dim ThisList As String
    Dim SSet As StringSet
    Dim AddSet As StringSet
    Dim Source As LinkedList
    Dim Match As String
    
    ThisList = tabTabs.SelectedItem.Key
    
    Select Case ThisList
        Case LIST_ACTIONS:      Set Source = ActionList
        Case LIST_RUMORS:       Set Source = RumorList
        Case Else:              Exit Sub
    End Select
    
    Set SSet = New StringSet
    Set AddSet = New StringSet
    
    SSet.StoreListBox lstList
        
    Source.First
    Do Until Source.Off
        If SSet.Has(Source.Item.Name) Then
            If Index = SAME_DATE Then
                If ThisList = LIST_ACTIONS Then
                    AddSet.Add CStr(Source.Item.ActDate)
                Else
                    AddSet.Add CStr(Source.Item.RumorDate)
                End If
            Else
                If ThisList = LIST_ACTIONS Then
                    AddSet.Add Source.Item.CharName
                Else
                    AddSet.Add Source.Item.Title
                End If
            End If
        End If
        Source.MoveNext
    Loop
    
    Source.First
    Do Until Source.Off
        Match = ""
        If Index = SAME_DATE Then
            If ThisList = LIST_ACTIONS Then
                Match = CStr(Source.Item.ActDate)
            Else
                Match = CStr(Source.Item.RumorDate)
            End If
        Else
            If ThisList = LIST_ACTIONS Then
                Match = Source.Item.CharName
            Else
                Match = Source.Item.Title
            End If
        End If
        If AddSet.Has(Match) Then SSet.Add Source.Item.Name
        Source.MoveNext
    Loop
    
    SSet.SelectListBox lstList, True, False
    
    Set SSet = Nothing
    Set AddSet = Nothing
    
End Sub

Private Sub cmdSelectNone_Click()
'
' Name:         cmdSelectNone
' Description:  Deselect all items in the list.
'
    
    SelectList lstList, False
    
End Sub

Private Sub cmdSelectOnly_Click()
'
' Name:         cmdSelectOnly_Click
' Description:  Select only those names that match a given query.
'

    Dim ThisList As String
    Dim SSet As StringSet
    Dim Q As QueryClass
    
    ThisList = tabTabs.SelectedItem.Key
    If Not (ThisList = LIST_CHARACTERS Or ThisList = LIST_PLAYERS) Then Exit Sub
    
    Set SSet = New StringSet
    
    With Game.QueryEngine
    
        .QueryList.MoveTo cboSelectOnly.Text
        If .QueryList.Off Then
            Set Q = New QueryClass
            Select Case ThisList
                Case LIST_CHARACTERS:   Q.Inventory = qiCharacters
                Case LIST_PLAYERS:      Q.Inventory = qiPlayers
            End Select
        Else
            Set Q = .QueryList.Item
        End If
        
        .MakeQuery Q, , (chkNot.Value = vbChecked)
        
        .Results.First
        Do Until .Results.Off
            SSet.Add .Results.Item.Name
            .Results.MoveNext
        Loop
        
    End With
        
    SSet.SelectListBox lstList, True, False
    
    Set SSet = Nothing
    Set Q = Nothing
    
End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Create objects needed.
'
    Set CustomTemplate = New TemplateClass
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Load
' Description:  Create objects needed.
'
   
    SaveSetting App.Title, "Output", "Format", CStr(IIf(optFormat(ofText).Value, ofText, _
                                               IIf(optFormat(ofRTF).Value, ofRTF, ofHTML)))
    Set CustomTemplate = Nothing

End Sub

Private Sub lblPercent_Change()
'
' Name:         lblPercent_Change
' Description:  Update the progress bar when the text of the label changes.
'
    If lblPercent.Visible Then
        If Val(lblPercent.Caption) <= pgbProgress.Max Then
            pgbProgress.Value = Val(lblPercent.Caption)
        End If
        Refresh
    End If
    
End Sub

Private Sub lstTemplates_Click()
'
' Name:         lstTemplates_Click
' Description:  Populate the format textboxes with the template filenames.
'

    Dim ShowFormats As Boolean
    Dim Template As TemplateClass
    Dim I As Integer
    
    Game.TemplateList.MoveTo lstTemplates.Text
    ShowFormats = Not Game.TemplateList.Off
    
    If ShowFormats Then
        Set Template = Game.TemplateList.Item
        txtFormat(ofText).Text = Template.GetFilename(ofText)
        txtFormat(ofRTF).Text = Template.GetFilename(ofRTF)
        txtFormat(ofHTML).Text = Template.GetFilename(ofHTML)
        cmdDeleteReport.Enabled = Not Template.IsCharacterSheet
    End If
    
    For I = MIN_OUTFORMAT To MAX_OUTFORMAT
        If txtFormat(I).Text = "" Then txtFormat(I).Text = "(none)"
        txtFormat(I).Visible = ShowFormats
        lblFormatLabel(I).Visible = ShowFormats
        cmdFormat(I).Visible = ShowFormats
    Next I

End Sub

Private Sub optAscend_Click(Index As Integer)
'
' Name:         optAscend_Click
' Description:  Update the outputEngine's AscendDate setting.
'
    OutputEngine.AscendDate = (Index = OPT_ASCEND)
    
End Sub

Private Sub optFormat_Click(Index As Integer)
'
' Name:         optFormat_Click
' Description:  Trigger cboTemplates_Click and update the OutputEngine.
'

    If Not (cboTemplate.ListIndex <= CharSheetIndex Or _
            cboTemplate.ListIndex >= OtherTemplateIndex) Then Call cboTemplate_Click
    OutputEngine.OutputFormat = IIf(Index = ofText, ofText, IIf(Index = ofRTF, ofRTF, ofHTML))
    chkAllowHTML.Visible = (Index = ofHTML)
    
    If OutputIndex = OUTPUT_PRINT Then
        lblPrintHTML.Visible = (Index = ofHTML)
        cmdPrintSetup.Visible = Not (Index = ofHTML)
        lblCopies.Visible = cmdPrintSetup.Visible
        txtCopies.Visible = cmdPrintSetup.Visible
        updCopies.Visible = cmdPrintSetup.Visible
    End If

End Sub

Private Sub optGraph_Click(Index As Integer)
'
' Name:         optGraph_Click
' Description:  Reformat cboKeys if needed, since maxima and sums only work with traits.
'
    
    Dim Store As String
    Dim I As Integer
    Dim Key As String
    Dim KeyType As QueryKeyType
    Dim Include As Boolean
    Dim InfIndex As Integer
    
    Store = cboKey.Text
    
    If Index = OPT_DIST Or cboKey.ListCount = 0 Then
        cboKey.Clear
        For I = 1 To Game.QueryEngine.TitlesToKeys.Count
            Key = Game.QueryEngine.TitlesToKeys(I)
            KeyType = Game.QueryEngine.KeysToTypes(Key)
            Include = CLng(Game.QueryEngine.KeysToInventories(Key)) And qiCharacters
            If ((Index = OPT_DIST) Or (KeyType = qtTraitList) Or (KeyType = qtNumber)) And Include Then
                cboKey.AddItem Game.QueryEngine.KeysToTitles(Key)
                cboKey.ItemData(cboKey.NewIndex) = KeyType
                If cboKey.List(cboKey.NewIndex) = Store Then cboKey.ListIndex = cboKey.NewIndex
            End If
        Next I
        OutputEngine.StatType = IIf(optTraitDistribution(OPT_DISTINCT).Value, _
                stDistinctDistribution, IIf(optTraitDistribution(OPT_SPECIFIC).Value, _
                stSpecificDistribution, stDistribution))
                
    Else
        OutputEngine.StatType = IIf(Index = OPT_SUM, stSums, stMaxima)
        I = 0
        Do While I < cboKey.ListCount
            If cboKey.ItemData(I) = qtTraitList Or cboKey.ItemData(I) = qtNumber Then
                I = I + 1
            Else
                cboKey.RemoveItem I
            End If
        Loop
    End If
    
    If cboKey.ListIndex = -1 Then
        InfIndex = 0
        For I = 0 To cboKey.ListCount - 1
            Key = Game.QueryEngine.TitlesToKeys(cboKey.List(I))
            If Key = OutputEngine.StatKey Then
                cboKey.ListIndex = I
                Exit For
            End If
            If Key = qkInfluences Then InfIndex = I
        Next I
    End If
    
    If cboKey.ListIndex = -1 And cboKey.ListCount > 0 Then cboKey.ListIndex = InfIndex
    
    If cboKey.ListIndex > -1 Then
        chkZero.Visible = (Index = OPT_DIST) And Not optTraitDistribution(OPT_DISTINCT).Value
        fraTraitDistribution.Visible = (cboKey.ItemData(cboKey.ListIndex) = qtTraitList) And _
            optGraph(OPT_DIST).Value
    Else
        chkZero.Visible = False
        fraTraitDistribution.Visible = False
    End If
    
End Sub

Private Sub optMatch_Click(Index As Integer)
'
' Name:         optMatch_Click
' Description:  Update the OutputEngine's SearchNot setting
'
    OutputEngine.SearchNot = (Index = OPT_NOT)
    
End Sub

Private Sub optTraitDistribution_Click(Index As Integer)
'
' Name:         optTraitDistribution_Click
' Description:  Update the title or transfer focus to the text box as needed.
'

    chkZero.Visible = optGraph(OPT_DIST).Value And Not (Index = OPT_DISTINCT)
    OutputEngine.StatType = IIf(Index = OPT_DISTINCT, stDistinctDistribution, _
            IIf(Index = OPT_SPECIFIC, stSpecificDistribution, stDistribution))
    
End Sub

Private Sub tabTabs_Click()
'
' Name:         tabTabs_Click
' Description:  Show the needed ListView and format the controls appropriately.
'

    Dim ThisList As String
    Dim IconKey As String
    Dim PopList As LinkedList
    Dim ShowFrame As Integer
    Dim SelectOnlyType As QueryInventoryType
    Dim SetIndex As Long
    Dim F As Frame
    
    Screen.MousePointer = vbHourglass
    StoreSelection
    ThisList = tabTabs.SelectedItem.Key
    lstList.Tag = ThisList
    
    Select Case ThisList
        Case LIST_OPTIONS:     ShowFrame = FRAME_OPTION
        Case LIST_TEMPLATES:   ShowFrame = FRAME_TEMPLATE
        Case LIST_SEARCH
            
            If cboKey.ListCount = 0 Then
                Select Case OutputEngine.StatType
                    Case stDistribution
                        optGraph(OPT_DIST).Value = True
                        optTraitDistribution(OPT_TOTAL).Value = True
                    Case stDistinctDistribution
                        optGraph(OPT_DIST).Value = True
                        optTraitDistribution(OPT_DISTINCT).Value = True
                    Case stSpecificDistribution
                        optGraph(OPT_DIST).Value = True
                        optTraitDistribution(OPT_SPECIFIC).Value = True
                    Case stMaxima
                        optGraph(OPT_MAXIMA).Value = True
                    Case stSums
                        optGraph(OPT_SUM).Value = True
                End Select
            End If
            ShowFrame = FRAME_SEARCH
            
        Case Else
            
            SelectOnlyType = qiNone
            cmdSelectSame(SAME_DATE).Visible = False
            lstList.Columns = 2
            
            Select Case ThisList
                Case LIST_PLAYERS
                    SelectOnlyType = qiPlayers
                    IconKey = "Players"
                    Set PopList = PlayerList
                    SetIndex = osPlayers
                Case LIST_CHARACTERS
                    SelectOnlyType = qiCharacters
                    IconKey = "Masks"
                    Set PopList = CharacterList
                    SetIndex = osCharacters
                Case LIST_ITEMS
                    IconKey = "Stake"
                    Set PopList = ItemList
                    SetIndex = osItems
                Case LIST_ROTES
                    IconKey = "Mage"
                    Set PopList = RoteList
                    SetIndex = osRotes
                Case LIST_LOCATIONS
                    IconKey = "Lantern"
                    Set PopList = LocationList
                    SetIndex = osLocations
                Case LIST_ACTIONS
                    IconKey = "Action"
                    cmdSelectSame(SAME_DATE).Visible = True
                    lstList.Columns = 1
                    Set PopList = ActionList
                    SetIndex = osActions
                Case LIST_PLOTS
                    IconKey = "Plot"
                    Set PopList = PlotList
                    SetIndex = osPlots
                Case LIST_RUMORS
                    IconKey = "Rumor"
                    cmdSelectSame(SAME_DATE).Visible = True
                    lstList.Columns = 1
                    Set PopList = RumorList
                    SetIndex = osRumors
            End Select
            
            lblTitle.Caption = "&" & tabTabs.SelectedItem.Caption
            imgIcon(0).Picture = mdiMain.imlSmallIcons.ListImages(IconKey).Picture
            imgIcon(1).Picture = imgIcon(0).Picture
        
            cmdSelectSame(SAME_NAME).Visible = cmdSelectSame(SAME_DATE).Visible
            
            lstList.Clear
            If Not PopList Is Nothing Then
                With PopList
                    .First
                    Do Until .Off
                        lstList.AddItem .Item.Name
                        lstList.Selected(lstList.NewIndex) = _
                                OutputEngine.SelectSet(SetIndex).Has(.Item.Name)
                        .MoveNext
                    Loop
                End With
            End If
            
            If SelectOnlyType = qiNone Then
                cmdSelectOnly.Visible = False
                cboSelectOnly.Visible = False
                chkNot.Visible = False
            Else
                cboSelectOnly.Clear
                With Game.QueryEngine.QueryList
                    .First
                    Do Until .Off
                        If .Item.Inventory = SelectOnlyType Then
                            cboSelectOnly.AddItem .Item.Name
                            If .Item.Name = OutputEngine.SearchName Then _
                                cboSelectOnly.ListIndex = cboSelectOnly.NewIndex
                        End If
                        .MoveNext
                    Loop
                End With
                cmdSelectOnly.Visible = True
                cboSelectOnly.Visible = True
                chkNot.Visible = True
                chkNot.Value = IIf(OutputEngine.SearchNot, vbChecked, vbUnchecked)
                If cboSelectOnly.ListIndex = -1 Then
                    If cboSelectOnly.ListCount > 0 Then cboSelectOnly.ListIndex = 0
                End If
            End If
        
            ShowFrame = FRAME_SELECT
            
    End Select
    
    For Each F In fraFrame
        F.Visible = (F.Index = ShowFrame)
    Next F

    Screen.MousePointer = vbDefault

End Sub

Private Sub txtCopies_Change()
'
' Name:         txtCopies_Change
' Description:  Adjust the number of copies to print.
'

    OutputEngine.Copies = Int(Val(txtCopies.Text))
    If OutputEngine.Copies < 1 Then OutputEngine.Copies = 1

End Sub

Private Sub txtFormat_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'
' Name:         txtFormat_KeyUp
' Description:  Delete the association with a file when the user presses delete or backspace.
'

    If KeyCode = vbKeyDelete Then
        Game.TemplateList.MoveTo lstTemplates.Text
        If Not Game.TemplateList.Off Then
            Game.TemplateList.Item.SetFilename CLng(Index), ""
            txtFormat(Index).Text = "(none)"
        End If
    End If
    
End Sub

Private Sub txtTrait_Change()
'
' Name:         txtTrait_Change
' Description:  Select this option, as the user fills it out; set the OutputEngine's StatTrait value.
'
    optTraitDistribution(OPT_SPECIFIC).Value = True
    OutputEngine.StatTrait = txtTrait.Text

End Sub

Private Sub txtTrait_GotFocus()
'
' Name:         txtTrait_GotFocus
' Description:  Select the control's text.
'
    SelectText txtTrait

End Sub

