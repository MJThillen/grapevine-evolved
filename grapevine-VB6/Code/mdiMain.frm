VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Grapevine"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   Begin VB.Timer timAutosave 
      Interval        =   60000
      Left            =   9120
      Tag             =   "0"
      Top             =   7200
   End
   Begin MSComctlLib.ProgressBar pgbProgress 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   0
      Top             =   8100
      Visible         =   0   'False
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ImageList imlSmallIcons 
      Left            =   10440
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   51
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":08CA
            Key             =   "Vampire"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0E64
            Key             =   "Wraith"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":13FE
            Key             =   "Changeling"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1998
            Key             =   "Kuei-Jin"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1F32
            Key             =   "Werewolf"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":24CC
            Key             =   "Mortal"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2A66
            Key             =   "Mummy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3000
            Key             =   "Mage"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":359A
            Key             =   "Various"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3B34
            Key             =   "Fera"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":40CE
            Key             =   "Hunter"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4668
            Key             =   "Demon"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4C02
            Key             =   "Masks"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":519C
            Key             =   "Lantern"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5736
            Key             =   "Tickets"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5CD0
            Key             =   "Stake"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":626A
            Key             =   "Players"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6804
            Key             =   "Printer"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6D9E
            Key             =   "Document"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":7338
            Key             =   "EMail"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":78D2
            Key             =   "Earpaper"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":7E6C
            Key             =   "Earpencil"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":8406
            Key             =   "Earbook"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":89A0
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":8F3A
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":94D4
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":9A6E
            Key             =   "Exchange"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":A008
            Key             =   "Calendar"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":A5A2
            Key             =   "Graph"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":AB3C
            Key             =   "New"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":B0D6
            Key             =   "PP"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":B670
            Key             =   "XP"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":BC0A
            Key             =   "Plot"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":C1A4
            Key             =   "Action"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":C73E
            Key             =   "Rumor"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":CCD8
            Key             =   "Device"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":D272
            Key             =   "Correspondence"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":D80C
            Key             =   "Entropy"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":DDA6
            Key             =   "Forces"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":E340
            Key             =   "Life"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":E8DA
            Key             =   "Matter"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":EE74
            Key             =   "Mind"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":F40E
            Key             =   "Prime"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":F9A8
            Key             =   "Spirit"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":FF42
            Key             =   "Time"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":104DC
            Key             =   "GroupRumor"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":10A76
            Key             =   "PersonalRumor"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":11010
            Key             =   "RaceRumor"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":115AA
            Key             =   "InfluenceRumor"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":11B44
            Key             =   "SubgroupRumor"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":120DE
            Key             =   "Harpy"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmnDialog 
      Left            =   11160
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "gv3"
      DialogTitle     =   "Open Game"
      Filter          =   "Grapevine Game Files|*.gv2;*.gv3|All Files|*.*"
      FilterIndex     =   1
      PrinterDefault  =   0   'False
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   9720
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":12678
            Key             =   "Vampire"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":12F52
            Key             =   "Werewolf"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1382C
            Key             =   "Mortal"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":14106
            Key             =   "Changeling"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":149E0
            Key             =   "Wraith"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":152BA
            Key             =   "Fera"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":15B94
            Key             =   "Kuei-Jin"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1646E
            Key             =   "Mummy"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":16D48
            Key             =   "Mage"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":17622
            Key             =   "Various"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":17EFC
            Key             =   "Hunter"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":187D6
            Key             =   "Stake"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":190B0
            Key             =   "Demon"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1998A
            Key             =   "Lantern"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1A264
            Key             =   "Correspondence"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1AB3E
            Key             =   "Device"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1B418
            Key             =   "Entropy"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1BCF2
            Key             =   "Forces"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1C5CC
            Key             =   "Life"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1CEA6
            Key             =   "Matter"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1D780
            Key             =   "Mind"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1E05A
            Key             =   "Prime"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1E934
            Key             =   "Spirit"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1F20E
            Key             =   "Time"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1FAE8
            Key             =   "Plot"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":203C2
            Key             =   "Action"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":20C9C
            Key             =   "Rumor"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":21576
            Key             =   "GroupRumor"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":21E50
            Key             =   "PersonalRumor"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2272A
            Key             =   "RaceRumor"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":23004
            Key             =   "InfluenceRumor"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":238DE
            Key             =   "SubgroupRumor"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":241B8
            Key             =   "Printer"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":24A92
            Key             =   "EMail"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2536C
            Key             =   "Document"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbToolbar 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imlSmallIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   29
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New Game"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open Game"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Game"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exchange"
            Object.ToolTipText     =   "Exchange Data"
            ImageKey        =   "Exchange"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Dates"
            Object.ToolTipText     =   "Game Dates"
            ImageKey        =   "Calendar"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Menus"
            Object.ToolTipText     =   "Grapevine Menu Editor"
            ImageKey        =   "Earbook"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Players"
            Object.ToolTipText     =   "Player Information"
            ImageKey        =   "Players"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Player Points"
            Object.ToolTipText     =   "Player Points"
            ImageKey        =   "PP"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Characters"
            Object.ToolTipText     =   "Character Sheets"
            ImageKey        =   "Masks"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Experience"
            Object.ToolTipText     =   "Experience Points"
            ImageKey        =   "XP"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tempers"
            Object.ToolTipText     =   "Manage Perm/Temp Ratings"
            ImageKey        =   "Tickets"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Harpy"
            Object.ToolTipText     =   "Vampire Boons and Status"
            ImageKey        =   "Vampire"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            Object.ToolTipText     =   "Search for Characters"
            ImageKey        =   "Search"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Statistics"
            Object.ToolTipText     =   "Character Statistics"
            ImageKey        =   "Graph"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Items"
            Object.ToolTipText     =   "Item Cards"
            ImageKey        =   "Stake"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rotes"
            Object.ToolTipText     =   "Rotes"
            ImageKey        =   "Prime"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Locations"
            Object.ToolTipText     =   "Locations"
            ImageKey        =   "Lantern"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Actions"
            Object.ToolTipText     =   "Actions"
            ImageKey        =   "Action"
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Plots"
            Object.ToolTipText     =   "Plots"
            ImageKey        =   "Plot"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rumors"
            Object.ToolTipText     =   "Rumors"
            ImageKey        =   "Rumor"
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print Sheets or Reports"
            ImageKey        =   "Printer"
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Export"
            Object.ToolTipText     =   "Save Sheets or Reports to File"
            ImageKey        =   "Document"
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "EMail"
            Object.ToolTipText     =   "Email Sheets or Reports"
            ImageKey        =   "EMail"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File "
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
      End
      Begin VB.Menu mnuOpenGame 
         Caption         =   "&Open Game..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSaveGame 
         Caption         =   "&Save Game"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveGameAs 
         Caption         =   "Save Game &As..."
      End
      Begin VB.Menu mnuFileExchangeBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExchange 
         Caption         =   "Data &Exchange..."
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuFileMerge 
         Caption         =   "&Merge Games..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRecentFileBar 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuGameMenu 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameInformation 
         Caption         =   "&Information..."
      End
      Begin VB.Menu mnuGameDates 
         Caption         =   "&Dates..."
      End
      Begin VB.Menu mnuGameSettings 
         Caption         =   "&Settings..."
      End
      Begin VB.Menu mnuGameBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGrapevineMenus 
         Caption         =   "Grapevine &Menus..."
      End
   End
   Begin VB.Menu mnuPlayerMenu 
      Caption         =   " &Players "
      Begin VB.Menu mnuPlayers 
         Caption         =   "&Player Information..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPlayerPoints 
         Caption         =   "Player P&oints..."
      End
      Begin VB.Menu mnuPlayerBar1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPlayerEMail 
         Caption         =   "Send E-Mail to Players..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuCharacterMenu 
      Caption         =   " &Characters "
      Begin VB.Menu mnuCharacters 
         Caption         =   "&Character Sheets..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuExperience 
         Caption         =   "E&xperience Points..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuCharBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCharTempers 
         Caption         =   "&Perm/Temp Ratings..."
      End
      Begin VB.Menu mnuCharHarpy 
         Caption         =   "&Vampire Boons && Status..."
      End
      Begin VB.Menu mnuCharBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCharSearch 
         Caption         =   "&Search for Characters..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuCharStats 
         Caption         =   "S&tatistics..."
      End
   End
   Begin VB.Menu mnuWorldMenu 
      Caption         =   " &World "
      Begin VB.Menu mnuWorldItems 
         Caption         =   "&Items..."
      End
      Begin VB.Menu mnuWorldRotes 
         Caption         =   "&Rotes..."
      End
      Begin VB.Menu mnuWorldLocations 
         Caption         =   "&Locations..."
      End
   End
   Begin VB.Menu mnuChronicleMenu 
      Caption         =   " Chr&onicle"
      Begin VB.Menu mnuChronicleActionS 
         Caption         =   "&Actions..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuChroniclePlots 
         Caption         =   "&Plots..."
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuChronicleRumors 
         Caption         =   "&Rumors..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuChronicleBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChroniclePreferences 
         Caption         =   "A&ction Settings..."
         Index           =   0
      End
      Begin VB.Menu mnuChroniclePreferences 
         Caption         =   "R&umor Settings..."
         Index           =   1
      End
   End
   Begin VB.Menu mnuOutput 
      Caption         =   " &Sheets && Reports"
      Begin VB.Menu mnuOutputPrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuOutputSave 
         Caption         =   "&Save to File..."
      End
      Begin VB.Menu mnuOutputEMail 
         Caption         =   "&E-Mail..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   " Wi&ndow "
      NegotiatePosition=   3  'Right
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowToolbar 
         Caption         =   "Show &Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuWindowBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowMaximize 
         Caption         =   "Ma&ximize Current"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuWindowMinimizeAll 
         Caption         =   "&Minimize All"
      End
      Begin VB.Menu mnuWindowRestoreAll 
         Caption         =   "&Restore All"
      End
      Begin VB.Menu mnuWindowCloseAll 
         Caption         =   "Close &All"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   " &Help "
      Begin VB.Menu mnuHelpMainpage 
         Caption         =   "&Grapevine Homepage"
      End
      Begin VB.Menu mnuHelpHelpPage 
         Caption         =   "Grapevine WWW &Help Pages"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuPopupItem 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public AutosaveTime As Long

Private Announcements() As String   'The array of forms that checked the announcements
Private RecentTitles As LinkedList  'The names of recent files
Private RecentFiles As LinkedList   'The paths of recent files
Private PopupChoice As String       'Chosen selection from a popup menu

Public Sub LaunchBrowser(URL As String)
'
' Name:         LaunchBrowser
' Parameters:   URL     the URL to go to
' Description:  Use the system's register browser to go to a URL.
'

    If ShellExecute(Me.hWnd, "open", URL, "", "", 1) = 31 Then
        MsgBox "Install a web browser on your computer to complete this function.", vbOKOnly, "No Browser"
    End If

End Sub

Public Sub ShowCharacterSheet(CharName As String, Optional ShowXP As Boolean = False)
'
' Name:         ShowCharacterSheet
' CharName:     The name of the character whose sheet to show
' Description:  If the character's sheet is loaded, bring it to the front.
'               If not, load it.
'

    Dim PotentialSheet As Form

    For Each PotentialSheet In Forms
        If PotentialSheet.Caption = CharName And PotentialSheet.Tag = "C" Then
            If ShowXP Then Call PotentialSheet.ShowXP
            PotentialSheet.SetFocus
            PotentialSheet.WindowState = vbNormal
            Exit Sub
        End If
    Next PotentialSheet

    CharacterList.MoveTo CharName
    If Not CharacterList.Off Then
        Select Case CharacterList.Item.RaceCode
            Case gvRaceVampire
                Set PotentialSheet = New frmVampireSheet
                PotentialSheet.ShowVampire CharacterList.Item
            Case gvRaceWerewolf
                Set PotentialSheet = New frmWerewolfSheet
                PotentialSheet.ShowWerewolf CharacterList.Item
            Case gvRaceMortal
                Set PotentialSheet = New frmMortalSheet
                PotentialSheet.ShowMortal CharacterList.Item
            Case gvRaceChangeling
                Set PotentialSheet = New frmChangelingSheet
                PotentialSheet.ShowChangeling CharacterList.Item
            Case gvRaceWraith
                Set PotentialSheet = New frmWraithSheet
                PotentialSheet.ShowWraith CharacterList.Item
            Case gvracemage
                Set PotentialSheet = New frmMageSheet
                PotentialSheet.ShowMage CharacterList.Item
            Case gvRaceFera
                Set PotentialSheet = New frmFeraSheet
                PotentialSheet.ShowFera CharacterList.Item
            Case gvRaceVarious
                Set PotentialSheet = New frmVariousSheet
                PotentialSheet.ShowVarious CharacterList.Item
            Case gvRaceMummy
                Set PotentialSheet = New frmMummySheet
                PotentialSheet.ShowMummy CharacterList.Item
            Case gvRaceKueiJin
                Set PotentialSheet = New frmKueiJinSheet
                PotentialSheet.ShowKueiJin CharacterList.Item
            Case gvRaceHunter
                Set PotentialSheet = New frmHunterSheet
                PotentialSheet.ShowHunter CharacterList.Item
            Case gvRaceDemon
                Set PotentialSheet = New frmDemonSheet
                PotentialSheet.ShowDemon CharacterList.Item
        End Select
        If ShowXP Then Call PotentialSheet.ShowXP
        Set PotentialSheet = Nothing
    Else
        MsgBox "Grapevine could not find a character named """ & CharName & """." & _
               vbCrLf & "It may have recently been deleted or renamed.", vbExclamation, _
               "Character Not Found"
    End If

End Sub

Public Sub ShowItemCard(ItemName As String)
'
' Name:         ShowItemCard
' Parameters:   ItemName        The name of the item whose card to show
' Description:  If the item card is loaded, bring it to the front.
'               If not, load it.
'

    Dim PotentialSheet As Form
    Dim ItemCard As frmItemCard

    For Each PotentialSheet In Forms
        If PotentialSheet.Caption = ItemName And PotentialSheet.Tag = "I" Then
            PotentialSheet.SetFocus 'Move to front
            PotentialSheet.WindowState = vbNormal
            Exit Sub
        End If
    Next PotentialSheet

    ItemList.MoveTo ItemName
    If Not ItemList.Off Then
        Set ItemCard = New frmItemCard
        ItemCard.ShowItem ItemList.Item
        Set ItemCard = Nothing
    Else
        MsgBox "Grapevine could not find an item card named """ & ItemName & """." & _
               vbCrLf & "It may have recently been deleted or renamed.", vbExclamation, _
               "Item Card Not Found"
    End If

End Sub

Public Sub ShowRote(RoteName As String)
'
' Name:         ShowRote
' Parameters:   RoteName        The name of the rote whose card to show
' Description:  If the rote is loaded, bring it to the front.
'               If not, load it.
'

    Dim PotentialSheet As Form
    Dim RoteCard As frmRoteCard

    For Each PotentialSheet In Forms
        If PotentialSheet.Caption = RoteName And PotentialSheet.Tag = "R" Then
            PotentialSheet.SetFocus 'Move to front
            PotentialSheet.WindowState = vbNormal
            Exit Sub
        End If
    Next PotentialSheet

    RoteList.MoveTo RoteName
    If Not RoteList.Off Then
        Set RoteCard = New frmRoteCard
        RoteCard.ShowRote RoteList.Item
        Set RoteCard = Nothing
    Else
        MsgBox "Grapevine could not find a rote named """ & RoteName & """." & _
               vbCrLf & "It may have recently been deleted or renamed.", vbExclamation, _
               "Rote Not Found"
    End If

End Sub

Public Sub ShowPlayer(PlayerName As String, Optional ShowPP As Boolean = False)
'
' Name:         ShowPlayer
' Parameters:   PlayerName        The name of the Player whose card to show
'               ShowPP            whether to show the player points
' Description:  If the Player is loaded, bring it to the front.
'               If not, load it.
'

    Dim PotentialSheet As Form
    Dim PlayerCard As frmPlayerCard

    For Each PotentialSheet In Forms
        If PotentialSheet.Caption = PlayerName And PotentialSheet.Tag = "Y" Then
            If ShowPP Then Call PotentialSheet.ShowPP
            PotentialSheet.SetFocus 'Move to front
            PotentialSheet.WindowState = vbNormal
            Exit Sub
        End If
    Next PotentialSheet

    PlayerList.MoveTo PlayerName
    If Not PlayerList.Off Then
        Set PlayerCard = New frmPlayerCard
        PlayerCard.ShowPlayer PlayerList.Item
        If ShowPP Then PlayerCard.ShowPP
        Set PlayerCard = Nothing
    Else
        MsgBox "Grapevine could not find a player named """ & PlayerName & """." & _
               vbCrLf & "It may have recently been deleted or renamed.", vbExclamation, _
               "Player Not Found"
    End If

End Sub

Public Sub ShowLocation(LocationName As String)
'
' Name:         ShowLocation
' Parameters:   LocationName        The name of the Location whose card to show
' Description:  If the Location is loaded, bring it to the front.
'               If not, load it.
'

    Dim PotentialSheet As Form
    Dim LocationCard As frmLocationCard

    For Each PotentialSheet In Forms
        If PotentialSheet.Caption = LocationName And PotentialSheet.Tag = "L" Then
            PotentialSheet.SetFocus 'Move to front
            PotentialSheet.WindowState = vbNormal
            Exit Sub
        End If
    Next PotentialSheet

    LocationList.MoveTo LocationName
    If Not LocationList.Off Then
        Set LocationCard = New frmLocationCard
        LocationCard.ShowLocation LocationList.Item
        Set LocationCard = Nothing
    Else
        MsgBox "Grapevine could not find a location named """ & LocationName & """." & _
               vbCrLf & "It may have recently been deleted or renamed.", vbExclamation, _
               "Location Not Found"
    End If

End Sub

Public Sub ShowAPR(APR As APRType, ShowName As String, ShowDate As Date, _
                   Optional ShowSubitem As String = "")
'
' Name:         ShowAPR
' Parameters:   APR         Whether to show an action, a plot or a rumor
'               ShowName    Name of the item to show
'               ShowDate    Date of the item to show
'               ShowSubitem The Subitem to show, if any
' Description:  If the Action, Plot or Rumor is loaded, bring it to the front.
'               If not, load it.
'

    Dim PotentialSheet As Form
    Dim Item As Object
    Dim ShowCaption As String
    Dim FormTag As String
    
    Select Case APR
        Case aprAction
            ShowCaption = "an action for " & ShowName & " dated " & Format(ShowDate, "Short Date")
            FormTag = "A"
            Game.APREngine.MoveToPair ActionList, ShowDate, ShowName
            If Not ActionList.Off Then Set Item = ActionList.Item
        Case aprPlot
            ShowCaption = "a plot called """ & ShowName & """"
            FormTag = "P"
            PlotList.MoveTo ShowName
            If Not PlotList.Off Then Set Item = PlotList.Item
        Case aprRumor
            ShowCaption = "a rumor called """ & ShowName & """ and dated " & Format(ShowDate, "Short Date")
            FormTag = "U"
            Game.APREngine.MoveToPair RumorList, ShowDate, ShowName
            If Not RumorList.Off Then Set Item = RumorList.Item
    End Select

    If Item Is Nothing Then
    
        MsgBox "Grapevine could not find " & ShowCaption & "." & _
               vbCrLf & "It may have been deleted, renamed, or left behind in a different file.", vbExclamation, _
               "Chronicle Element Not Found"
               
    Else

        ShowCaption = Item.Name
                
        For Each PotentialSheet In Forms
            If PotentialSheet.Caption = ShowCaption And PotentialSheet.Tag = FormTag Then
                PotentialSheet.SetFocus 'Move to front
                PotentialSheet.WindowState = vbNormal
                Select Case APR
                    Case aprAction
                        If ShowSubitem <> "" Then PotentialSheet.ShowSubaction ShowSubitem
                    Case aprPlot
                        If ShowDate <> 0 Then PotentialSheet.ShowDevelopment ShowDate
                    Case aprRumor
                        If ShowSubitem <> "" Then PotentialSheet.ShowLevel CInt(ShowSubitem)
                End Select
                Exit Sub
            End If
        Next PotentialSheet

        Select Case APR
            Case aprAction
                Set PotentialSheet = New frmAction
                PotentialSheet.ShowAction Item
                If ShowSubitem <> "" Then PotentialSheet.ShowSubaction ShowSubitem
            Case aprPlot
                Set PotentialSheet = New frmPlot
                PotentialSheet.ShowPlot Item
                If ShowDate <> 0 Then PotentialSheet.ShowDevelopment ShowDate
            Case aprRumor
                Set PotentialSheet = New frmRumor
                PotentialSheet.ShowRumor Item
                If ShowSubitem <> "" Then PotentialSheet.ShowLevel CInt(ShowSubitem)
        End Select
        
        Set PotentialSheet = Nothing

    End If
    
End Sub

Public Sub EnableMenus(Maybe As Boolean)
'
' Name:         EnableMenus
' Parameters:   Maybe       whether to enable or diable most menus
' Description:  Enable or disable the menus.
'

    mnuSaveGame.Enabled = Maybe
    mnuSaveGameAs.Enabled = Maybe
    mnuFileExchange.Enabled = Maybe
    mnuFileMerge.Enabled = Maybe
    mnuGameMenu.Enabled = Maybe
    mnuPlayerMenu.Enabled = Maybe
    mnuCharacterMenu.Enabled = Maybe
    mnuWorldMenu.Enabled = Maybe
    mnuChronicleMenu.Enabled = Maybe
    mnuOutput.Enabled = Maybe
    
    tlbToolbar.Buttons("Save").Enabled = Maybe
    tlbToolbar.Buttons("Exchange").Enabled = Maybe
    tlbToolbar.Buttons("Dates").Enabled = Maybe
    tlbToolbar.Buttons("Menus").Enabled = Maybe
    tlbToolbar.Buttons("Players").Enabled = Maybe
    tlbToolbar.Buttons("Player Points").Enabled = Maybe
    tlbToolbar.Buttons("Characters").Enabled = Maybe
    tlbToolbar.Buttons("Experience").Enabled = Maybe
    tlbToolbar.Buttons("Tempers").Enabled = Maybe
    tlbToolbar.Buttons("Harpy").Enabled = Maybe
    tlbToolbar.Buttons("Search").Enabled = Maybe
    tlbToolbar.Buttons("Statistics").Enabled = Maybe
    tlbToolbar.Buttons("Items").Enabled = Maybe
    tlbToolbar.Buttons("Rotes").Enabled = Maybe
    tlbToolbar.Buttons("Locations").Enabled = Maybe
    tlbToolbar.Buttons("Actions").Enabled = Maybe
    tlbToolbar.Buttons("Plots").Enabled = Maybe
    tlbToolbar.Buttons("Rumors").Enabled = Maybe
    tlbToolbar.Buttons("Print").Enabled = Maybe
    tlbToolbar.Buttons("Export").Enabled = Maybe
    tlbToolbar.Buttons("EMail").Enabled = Maybe
    
End Sub

Private Sub UnloadForms()
'
' Name:         UnloadForms
' Description:  Dismiss all child windows.  Clear all form announcements.
'

    Dim ExForm As Form
    
    For Each ExForm In Forms()
        If Not ExForm Is Me Then Unload ExForm
    Next ExForm

    ReDim Announcements(MIN_ANNOUNCE To MAX_ANNOUNCE)

End Sub

Private Sub PromptForSave(ByRef Continue As Boolean)
'
' Name:         PromptForSave
' Description:  If the game's data has changed, ask if the user wants to save.  If so,
'               call mnuSave_Click.
'

    Dim Answer As Integer
    
    If Game.DataChanged Then
        Answer = MsgBox("Do you want to save the game first?", _
                         vbYesNoCancel + vbQuestion, "Save Game?")
        Select Case Answer
            Case vbYes
                Call mnuSaveGame_Click
                Continue = Not Game.FileError
            Case vbNo
                Continue = True
            Case vbCancel
                Continue = False
        End Select
    Else
        Continue = True
    End If

    If Continue Then
        If Game.MenuSet.DataChanged Then
            Answer = MsgBox("You have made changes to the menu file." & vbCrLf & _
                            "Would you like to save those changes?", _
                             vbYesNoCancel + vbQuestion, "Save Menus?")
            Select Case Answer
                Case vbYes
                    Game.MenuSet.SaveMenus Game.MenuSet.FilePath
                    Continue = Not Game.MenuSet.FileError
                Case vbNo
                    Continue = True
                Case vbCancel
                    Continue = False
            End Select
        Else
            Continue = True
        End If
    End If
    
End Sub

Private Sub UpdateRecentFiles(NewTitle As String, NewFile As String)
'
' Name:         UpdateRecentFiles
' Parameters:   NewTitle    The new file title to add
'               NewFile     The new file path to add
' Description:  Update the list of four recently opened files.
'

    Dim Index As Integer

    If NewTitle <> "" Then
        
        RecentFiles.First
        RecentTitles.First
        Do Until RecentFiles.Off
            If RecentFiles.Item = NewFile Then Exit Do
            RecentFiles.MoveNext
            RecentTitles.MoveNext
        Loop
        
        If RecentFiles.Off And RecentFiles.Count = 4 Then
            RecentFiles.Last
            RecentTitles.Last
        End If
        
        If Not RecentFiles.Off Then
            RecentFiles.Remove
            RecentTitles.Remove
        End If
                
        RecentFiles.Prepend NewFile
        RecentTitles.Prepend NewTitle
        
        Index = 0
        RecentFiles.First
        RecentTitles.First
        Do Until RecentFiles.Off
            SaveSetting App.Title, "Settings", "RecentTitle" & CStr(Index), RecentTitles.Item
            SaveSetting App.Title, "Settings", "RecentFile" & CStr(Index), RecentFiles.Item
            Index = Index + 1
            RecentTitles.MoveNext
            RecentFiles.MoveNext
        Loop
        
    End If
    
    RecentTitles.First
    Index = 0
    Do Until RecentTitles.Off
        mnuRecentFile(Index).Caption = "&" & CStr(Index + 1) & ". " & RecentTitles.Item
        mnuRecentFile(Index).Visible = True
        Index = Index + 1
        RecentTitles.MoveNext
    Loop
    If Index > 0 Then mnuRecentFileBar.Visible = True

End Sub

Private Sub MDIForm_Load()
'
' Name:         MDIForm_Load
' Description:  Initialize the main window, loading all needed settings.
'

    Dim Index As Integer
    Dim Setting As String
    Dim MidTop As Integer
    Dim MidLeft As Integer
    Dim MidHeight As Integer
    Dim MidWidth As Integer

    Set Game.FileProgress = pgbProgress
    Set Game.MenuSet.FileProgress = pgbProgress

    Me.Caption = GrapevineCaption
    Me.WindowState = GetSetting(App.Title, "Settings", "WindowState", vbMaximized)
    Me.Width = GetSetting(App.Title, "Settings", "WindowWidth", 12000)
    Me.Height = GetSetting(App.Title, "Settings", "WindowHeight", 8925)
    MidTop = (Screen.Height - Me.Height) \ 2
    MidWidth = (Screen.Width - Me.Width) \ 2
    Me.Top = GetSetting(App.Title, "Settings", "WindowTop", MidTop)
    Me.Left = GetSetting(App.Title, "Settings", "WindowLeft", MidLeft)
    mnuWindowToolbar.Checked = GetSetting(App.Title, "Settings", "ShowToolbar", True)
    AutosaveTime = CLng(GetSetting(App.Title, "Settings", "Autosave", 0))
    tlbToolbar.Visible = mnuWindowToolbar.Checked
    
    ' Include the SMTP.ocx control in the project and add an instance to this form,
    ' naming it smtpMailer, in order to re-enable the code below. Note that SMTP.ocx
    ' is detected (inaccurately) as a virus component by many antivirus programs.
    
    ' Set OutputEngine.Mailer = smtpMailer
    
    Set RecentFiles = New LinkedList
    Set RecentTitles = New LinkedList
    ReDim Announcements(MIN_ANNOUNCE To MAX_ANNOUNCE)
        
    For Index = 0 To 3
        Setting = GetSetting(App.Title, "Settings", "RecentTitle" & CStr(Index), "")
        If Setting <> "" Then
            RecentTitles.Append Setting
            mnuRecentFile(Index).Caption = "&" & CStr(Index + 1) & ". " & Setting
            mnuRecentFile(Index).Visible = True
            Setting = GetSetting(App.Title, "Settings", "RecentFile" & CStr(Index), "")
            RecentFiles.Append Setting
            mnuRecentFileBar.Visible = True
        End If
    Next Index

    EnableMenus False

    If Command <> "" Then
        Setting = Replace(Command, """", "")
        If Dir(Setting) <> "" Then
            Call UpdateRecentFiles(ShortFile(Setting), Setting)
            Call mnuRecentFile_Click(0)
        End If
    End If
    
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'
' Name:         MDIForm_OLEDragDrop
' Description:  Allow users to drag and drop files into Grapevine.
'

    Dim Title As String
    Dim File As String

    If Data.Files.Count > 0 Then
        File = Data.Files(1)
        If LCase(Right(File, 3)) = "gex" Then
            UnloadForms
            Game.LoadExchange File
        Else
            Title = Right(File, Len(File) - InStrRev(File, "\"))
            Call UpdateRecentFiles(Title, File)
            Call mnuRecentFile_Click(0)
        End If
    End If

End Sub

Private Sub mnuAbout_Click()
'
' Name:         mnuAbout_Click
' Description:  Display the About window.
'

    frmAbout.Show 1, Me
    Set frmAbout = Nothing

End Sub

Private Sub mnuCharHarpy_Click()
'
' Name:         mnuCharHarpy_Click
' Description:  Show the Status/Boons window.
'
    frmHarpyLedger.Show
    frmHarpyLedger.SetFocus

End Sub

Private Sub mnuCharTempers_Click()
'
' Name:         mnuCharTempers_Click
' Description:  Make visible the temper management window.
'
    frmPermTemp.Show
    frmPermTemp.SetFocus

End Sub

Private Sub mnuChronicleActions_Click()
'
' Name:         mnuChronicleActions_Click
' Description:  Make visible the Actions window.
'

    frmActionList.Show
    frmActionList.SetFocus

End Sub

Private Sub mnuChroniclePlots_Click()
'
' Name:         mnuChroniclePlots_Click
' Description:  Make visible the Plots window.
'

    frmPlotList.Show
    frmPlotList.SetFocus

End Sub

Private Sub mnuChroniclePreferences_Click(Index As Integer)
'
' Name:         mnuChroniclePreferences_Click
' Description:  Make visible the Game Settings window with the chosen tab selected.
'

    frmGameInfo.ShowWith IIf(Index = 0, "Actions", "Rumors")

End Sub

Private Sub mnuChronicleRumors_Click()
'
' Name:         mnuChronicleRumors_Click
' Description:  Make visible the Rumors window.
'

    frmRumorList.Show
    frmRumorList.SetFocus

End Sub

Private Sub mnuFileMerge_Click()
'
' Name:         mnuFileMerge_Click
' Description:  Merge in another file, replacing the old with the new.
'

    cmnDialog.DialogTitle = "Merge Game"
    cmnDialog.InitDir = GetSetting(App.Title, "Files", "GameDir", CurDir)
    cmnDialog.FileName = ""
    cmnDialog.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist
    cmnDialog.DefaultExt = "gv3"
    cmnDialog.Filter = "Grapevine Game Files|*.gv3;*.gv2|" & _
                       "All Files|*.*"
    cmnDialog.FilterIndex = 1
    
    On Error GoTo mnuMergeGame_AnyError
    cmnDialog.ShowOpen
    On Error GoTo 0

    SaveSetting App.Title, "Files", "GameDir", CurDir

    UnloadForms
    Screen.MousePointer = vbHourglass
    
    pgbProgress.Value = 0
    pgbProgress.Max = 101
    pgbProgress.Visible = True
    Game.Merge cmnDialog.FileName
    pgbProgress.Visible = False
    
    Screen.MousePointer = vbDefault
    
    If Not Game.FileError Then
        frmMergeResults.ShowResults Game.MergeResults, Me
    Else
        MsgBox Game.FileErrorMessage, vbExclamation, "Open Game"
    End If
    
    Call mnuCharacters_Click
    
    GoTo mnuMergeGame_Finish
    
mnuMergeGame_AnyError:
    Resume mnuMergeGame_Finish
mnuMergeGame_Finish:

End Sub

Private Sub mnuGrapevineMenus_Click()
'
' Name:         mnuGrapevineMenus_Click
' Description:  Display the Grapevine Menu Editor window.
'

    frmMenuEditor.Show
    frmMenuEditor.SetFocus

End Sub

Private Sub mnuCharSearch_Click()
'
' Name:         mnuCharSearch_Click
' Description:  Display the Character Searches window.
'
    
    frmQuery.Show
    frmQuery.SetFocus
    
End Sub

Private Sub mnuCharStats_Click()
'
' Name:         mnuCharStats_Click
' Description:  Show the Statistics form.
'
    frmStatistics.Show
    frmStatistics.SetFocus
    
End Sub

Private Sub mnuExperience_Click()
'
' Name:         mnuExperience_Click
' Description:  Display the experience point version of the Point Maintenance window.
'

    Dim MaintForm As Form
    
    For Each MaintForm In Forms
        If MaintForm.Tag = pmExperience Then
            MaintForm.SetFocus
            Exit Sub
        End If
    Next MaintForm
    
    Set MaintForm = New frmPointMaintenance
    MaintForm.ShowPointMaintenance pmExperience
    Set MaintForm = Nothing

End Sub

Private Sub mnuFileExchange_Click()
'
' Name:         mnuFileExchange_Click
' Description:  Display the Exchange window.
'

    UnloadForms
    frmExchange.Show 1, Me
    Set frmExchange = Nothing
    Call mnuCharacters_Click

End Sub

Private Sub mnuGameDates_Click()
'
' Name:         mnuGameDates_Click
' Description:  Display the Game Dates window.
'

    frmGameInfo.ShowWith "Dates"

End Sub

Private Sub mnuGameInformation_Click()
'
' Name:         mnuGameInformation_Click
' Description:  Display the Game Information window.
'

    frmGameInfo.ShowWith "Info"

End Sub

Private Sub mnuGameSettings_Click()
'
' Name:         mnuGameSettings_Click
' Description:  Display the Game Preferences window.
'

    frmGameInfo.ShowWith "General"

End Sub

Private Sub mnuHelpHelpPage_Click()
'
' Name:         mnuHelpHelpPage_Click
' Description:  Load the help pages in the system browser.
'

    LaunchBrowser URLHelpPage

End Sub

Private Sub mnuHelpMainpage_Click()
'
' Name:         mnuHelpMainpage_Click
' Description:  Load the main Grapevine homepage in the system browser.
'

    LaunchBrowser URLMainPage

End Sub

Public Function CreatePopup(Captions As StringSet, Source As Form) As String
'
' Name:         CreatePopup
' Parameters:   Captions        a stringset of captions for the menu items
'               Source          Form that causes the popup menu
' Description:  Create a popup menu and return the user selection from it.
'

    PopupChoice = ""

    Captions.First
    If Not Captions.Off Then
    
        Dim MenuIndex As Integer
        
        MenuIndex = 0
        
        Do
            If mnuPopupItem.Count = MenuIndex Then Load mnuPopupItem(MenuIndex)
            mnuPopupItem(MenuIndex).Caption = Captions.StrItem
            mnuPopupItem(MenuIndex).Visible = True
            Captions.MoveNext
            MenuIndex = MenuIndex + 1
        Loop Until Captions.Off
        
        Do Until MenuIndex = mnuPopupItem.Count
            mnuPopupItem(MenuIndex).Visible = False
            MenuIndex = MenuIndex + 1
        Loop
    
        Source.PopupMenu mnuPopup           'changes PopupChoice
    
    End If

    CreatePopup = PopupChoice

End Function

Private Sub mnuOutputEMail_Click()
'
' Name:         mnuOutputEMail
' Description:  Show the email output window
'

    On Error Resume Next
    ActiveForm.SetDefaultOutput
    On Error GoTo 0
    frmOutput.ShowOutput odEMail

End Sub

Private Sub mnuOutputPrint_Click()
'
' Name:         mnuOutputPrint
' Description:  Show the print output window
'

    On Error Resume Next
    ActiveForm.SetDefaultOutput
    On Error GoTo 0
    frmOutput.ShowOutput odPrinter

End Sub

Private Sub mnuOutputSave_Click()
'
' Name:         mnuOutputSave
' Description:  Show the file output window
'

    On Error Resume Next
    ActiveForm.SetDefaultOutput
    On Error GoTo 0
    frmOutput.ShowOutput odFile

End Sub

Private Sub mnuPlayerEMail_Click()
'
' Name:         mnuPlayerEMail_Click
' Description:  Compose and send an e-mail to players.
'

' Include the SMTP.ocx control in the project and add an instance to this form,
' naming it smtpMailer, in order to re-enable the code below. Note that SMTP.ocx
' is detected (inaccurately) as a virus component by many antivirus programs.

'    If Not InitializeSMTP Then
'        MsgBox "First you need to set up your outgoing E-Mail access.", vbOKOnly, "E-Mail"
'        frmEMailSetup.ShowSetup
'    End If
'
'    If InitializeSMTP Then
'        OutputEngine.InitializeMessage ooReport, Game.ChronicleTitle & " Announcement"
'        frmEMailAddressing.ShowAddressing OutputEngine.Mailer, ooReport, "E-Mail Players"
'
'        On Error Resume Next
'        Me.ActiveForm.Refresh
'        On Error GoTo 0
'        Screen.MousePointer = vbHourglass
'
'        With OutputEngine
'            If .Mailer.Tag = "" Then
'                .Mailer.SendTo = .SendTo
'                .Mailer.CC = .CC
'                .Mailer.BCC = .BCC
'                .Mailer.MessageSubject = .MessageSubject & _
'                        IIf(.ReplyTo = "", "", vbCrLf & "Reply-To: " & .ReplyTo)
'                .Mailer.MessageText = .MessageHeader
'                .Mailer.SendEmail
'                If .Mailer.Tag <> "" Then
'                    MsgBox "Error sending e-mail:" & vbCrLf & vbCrLf & .Mailer.Tag, _
'                            vbOKOnly + vbExclamation, "E-Mail Error"
'                End If
'            End If
'        End With
'
'        Screen.MousePointer = vbDefault
'
'    End If

End Sub

Private Sub mnuPopupItem_Click(Index As Integer)
'
' Name:         mnuPopupItem
' Description:  Store the item the user selects from the popup menu.
'

    PopupChoice = mnuPopupItem(Index).Caption

End Sub

Private Sub mnuPlayerPoints_Click()
'
' Name:         mnuPlayerPoints_Click
' Description:  Display the player points version of the Point Maintenance window.
'

    Dim MaintForm As Form
    
    For Each MaintForm In Forms
        If MaintForm.Tag = pmPlayerPoints Then
            MaintForm.SetFocus
            Exit Sub
        End If
    Next MaintForm
    
    Set MaintForm = New frmPointMaintenance
    MaintForm.ShowPointMaintenance pmPlayerPoints
    Set MaintForm = Nothing

End Sub

Private Sub mnuRecentFile_Click(Index As Integer)
'
' Name:         mnuRecentFile_Click
' Description:  Load a file from the list of recently opened files.
'

    Dim Continue As Boolean
    
    ValidateActiveForm
    PromptForSave Continue
    
    If Continue Then
    
        RecentFiles.First
        RecentTitles.First
        Do Until Index = 0
            Index = Index - 1
            RecentFiles.MoveNext
            RecentTitles.MoveNext
        Loop
    
        UnloadForms
        
        Screen.MousePointer = vbHourglass
        pgbProgress.Value = 0
        pgbProgress.Max = 101
        pgbProgress.Visible = True
        Game.OpenGame RecentFiles.Item
        pgbProgress.Visible = False
        Screen.MousePointer = vbDefault
        
        If Not Game.FileError Then
            Me.Caption = GrapevineCaption & ": " & RecentTitles.Item
            EnableMenus True
            Call UpdateRecentFiles(RecentTitles.Item, RecentFiles.Item)
            Call mnuCharacters_Click
        Else
            Me.Caption = GrapevineCaption
            EnableMenus False
            MsgBox Game.FileErrorMessage, vbExclamation, "Open File"
        End If
            
    End If

End Sub

Private Sub mnuWindowMaximize_Click()
'
' Name:         mnuWindowMaximize_Click
' Description:  Close the active child window.
'

    If Not mdiMain.ActiveForm Is Nothing And Not mdiMain.ActiveForm Is Me Then
        If mdiMain.ActiveForm.BorderStyle = vbSizable Then
            mdiMain.ActiveForm.WindowState = vbNormal
            mdiMain.ActiveForm.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        End If
    End If

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
' Name:         MDIForm_QueryUnload
' Description:  Ensure that the user really wants to close the program.
'

    Dim Continue As Boolean
    
    ValidateActiveForm
    PromptForSave Continue
    Cancel = Not Continue

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'
' Name:         MDIForm_Unload
' Description:  Peform final cleanup before the program exits.
'

    Screen.MousePointer = vbHourglass
    UnloadForms
    SaveSetting App.Title, "Settings", "WindowState", Me.WindowState
    SaveSetting App.Title, "Settings", "ShowToolbar", mnuWindowToolbar.Checked
    If Me.WindowState = vbNormal Then
        SaveSetting App.Title, "Settings", "WindowTop", Me.Top
        SaveSetting App.Title, "Settings", "WindowLeft", Me.Left
        SaveSetting App.Title, "Settings", "WindowWidth", Me.Width
        SaveSetting App.Title, "Settings", "WindowHeight", Me.Height
    End If
    SaveSetting App.Title, "Settings", "Autosave", AutosaveTime
    Set RecentTitles = Nothing
    Set RecentFiles = Nothing
    CleanUp
    Screen.MousePointer = vbDefault

End Sub

Private Sub mnuCharacters_Click()
'
' Name:         mnuCharacters_Click
' Description:  Display the Characters window.
'

    frmCharacters.Show
    frmCharacters.SetFocus

End Sub

Private Sub mnuExit_Click()
'
' Name:         mnuExit_Click
' Description:  Exit the program.
'

    Unload Me

End Sub

Private Sub mnuNewGame_Click()
'
' Name:         mnuNewGame_Click
' Description:  Make sure the user wants a new game.  If so, unload all windows
'               and initialize to an empty game.
'

    Dim Continue As Boolean
    
    ValidateActiveForm
    PromptForSave Continue
    
    If Continue Then
        UnloadForms
        Screen.MousePointer = vbHourglass
        pgbProgress.Value = 0
        pgbProgress.Max = 101
        pgbProgress.Visible = True
        Game.NewGame
        pgbProgress.Visible = False
        Screen.MousePointer = vbDefault
        EnableMenus True
        Me.Caption = GrapevineCaption
        Call mnuCharacters_Click
        Call mnuGameInformation_Click
    End If
    
End Sub

Private Sub mnuOpenGame_Click()
'
' Name:         mnuOpenGame_Click
' Description:  Prompt the user for a filename, then open a new game.
'

    Dim Continue As Boolean
    
    ValidateActiveForm
    PromptForSave Continue
    
    If Continue Then
        cmnDialog.DialogTitle = "Open Game"
        cmnDialog.InitDir = GetSetting(App.Title, "Files", "GameDir", CurDir)
        cmnDialog.FileName = ""
        cmnDialog.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist + cdlOFNHideReadOnly
        cmnDialog.DefaultExt = "gv3"
        cmnDialog.Filter = "Grapevine Game Files|*.gv3;*.gv2|" & _
                           "All Files|*.*"
        cmnDialog.FilterIndex = 1
        
        On Error GoTo mnuOpenGame_AnyError
        cmnDialog.ShowOpen
        On Error GoTo 0
    
        SaveSetting App.Title, "Files", "GameDir", CurDir
    
        UnloadForms
        Screen.MousePointer = vbHourglass
        
        On Error Resume Next
        pgbProgress.Value = 0
        pgbProgress.Max = 101
        pgbProgress.Visible = True
        Game.OpenGame cmnDialog.FileName
        pgbProgress.Visible = False
        On Error GoTo 0
        
        Screen.MousePointer = vbDefault
        
        If Not Game.FileError Then
            If Game.ChronicleTitle = "" Then
                Me.Caption = GrapevineCaption & ": " & cmnDialog.FileTitle
                Call UpdateRecentFiles(cmnDialog.FileTitle, cmnDialog.FileName)
            Else
                Me.Caption = GrapevineCaption & ": " & Game.ChronicleTitle
                Call UpdateRecentFiles(Game.ChronicleTitle, cmnDialog.FileName)
            End If
            EnableMenus True
            Call mnuCharacters_Click
        Else
            Me.Caption = GrapevineCaption
            EnableMenus False
            MsgBox Game.FileErrorMessage, vbExclamation, "Open Game"
        End If
    End If
    
    GoTo mnuOpenGame_Finish
    
mnuOpenGame_AnyError:
    Resume mnuOpenGame_Finish
mnuOpenGame_Finish:

End Sub

Private Sub mnuPlayers_Click()
'
' Name:         mnuPlayers_Click
' Description:  Display the Player Information window.
'

    frmPlayerList.Show
    frmPlayerList.SetFocus
    
End Sub

Private Sub mnuSaveGame_Click()
'
' Name:         mnuSaveGame_Click
' Description:  Save the current game, prompting if there is no filename yet.
'

    If Game.GameFile = "" Then
        Call mnuSaveGameAs_Click
    Else
        
        ValidateActiveForm
        
        Screen.MousePointer = vbHourglass
        pgbProgress.Value = 0
        pgbProgress.Max = 101
        pgbProgress.Visible = True
        Game.SaveGame Game.GameFile
        pgbProgress.Visible = False
        Screen.MousePointer = vbDefault
        
        If Game.FileError Then _
            MsgBox Game.FileErrorMessage, vbExclamation, "Save Game"
    
    End If
    
End Sub

Private Sub mnuSaveGameAs_Click()
'
' Name:         mnuSaveGameAs_Click
' Description:  Prompt for a filename and save the curent game.
'

    ValidateActiveForm
    
    cmnDialog.DialogTitle = "Save Game As..."
    cmnDialog.InitDir = GetSetting(App.Title, "Files", "GameDir", CurDir)
    cmnDialog.FileName = Game.GameFile
    cmnDialog.Flags = cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + _
            cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    cmnDialog.DefaultExt = "gv3"
    cmnDialog.Filter = "Grapevine Game File (Binary Format)|*.gv3;*.gv2|" & _
                       "Grapevine Game File (XML Format)|*.gv3;*.gv2|" & _
                       "All Files|*.*"
    cmnDialog.FilterIndex = IIf(Game.FileFormat = gvXML, 2, 1)
    
    On Error GoTo mnuSaveGameAs_CancelError
    cmnDialog.ShowSave
    On Error GoTo 0
    
    SaveSetting App.Title, "Files", "GameDir", CurDir
    Game.FileFormat = IIf(cmnDialog.FilterIndex = 1, gvBinaryGame, gvXML)
    
    Screen.MousePointer = vbHourglass
    pgbProgress.Value = 0
    pgbProgress.Max = 101
    pgbProgress.Visible = True
    Game.SaveGame cmnDialog.FileName
    pgbProgress.Visible = False
    Screen.MousePointer = vbDefault
    
    If Not Game.FileError Then
        If Game.ChronicleTitle = "" Then
            Call UpdateRecentFiles(cmnDialog.FileTitle, cmnDialog.FileName)
            Me.Caption = GrapevineCaption & ": " & cmnDialog.FileTitle
        Else
            Call UpdateRecentFiles(Game.ChronicleTitle, cmnDialog.FileName)
            Me.Caption = GrapevineCaption & ": " & Game.ChronicleTitle
        End If
    Else
        MsgBox Game.FileErrorMessage, vbExclamation, "Save Game"
    End If
    
    GoTo mnuSaveGameAs_Finish
    
mnuSaveGameAs_CancelError:
    Resume mnuSaveGameAs_Finish
mnuSaveGameAs_Finish:
    
End Sub

Private Sub mnuWindowCloseAll_Click()
'
' Name:         mnuWindowCloseAll_Click
' Description:  Dismiss all child windows.
'

    Call UnloadForms

End Sub

Private Sub mnuWindowMinimizeAll_Click()
'
' Name:         mnuWindowMinimizeAll_Click
' Description:  Minimize all child windows.
'
    
    Dim MinForm As Form
    
    For Each MinForm In Forms()
        If Not MinForm Is Me Then MinForm.WindowState = vbMinimized
    Next MinForm

End Sub

Private Sub mnuWindowRestoreAll_Click()
'
' Name:         mnuWindowRestoreAll_Click
' Description:  Restore all child windows.
'
    
    Dim NormForm As Form
    
    For Each NormForm In Forms()
        If Not NormForm Is Me Then NormForm.WindowState = vbNormal
    Next NormForm

End Sub

Private Sub ValidateActiveForm()
'
' Name:         ValidateActiveForm
' Description:  Ensure that the text being edited on the active form is saved.
'

    If Not Me.ActiveForm Is Nothing Then
        Me.ActiveForm.ValidateControls
    End If

End Sub

Public Sub OrientForm(Window As Form)
'
' Name:         OrientForm
' Parameters:   Window      the form to orient within the parent
' Description:  Ensure a window does not fall outside the bounds of
'               the parent window when it is displayed.
'

    If Window.Width + Window.Left > Me.ScaleWidth Then Window.Left = 0
    If Window.Top + Window.Height > Me.ScaleHeight Then Window.Top = 0

End Sub

Public Sub AnnounceChanges(AnnounceForm As Form, GameComponent As AnnounceType)
'
' Name:         AnnounceChanges
' Parameters:   AnnounceForm            form making the announcement
'               GameComponent           what game component changed
' Description:  Post a notice to all forms that the given game component changed.
'
    Announcements(GameComponent) = AnnounceForm.Name & AnnounceForm.Caption & "+"

End Sub

Public Function CheckForChanges(CheckForm As Form, GameComponent As AnnounceType) As Boolean
'
' Name:         CheckForChanges
' Parameters:   CheckForm           the form that's doing the checking
'               GameComponent       the component to check
' Description:  Return TRUE if changes have been announced since last this form checked.
'               No entry in the collection = no announcement.
'

    Dim Sigs As String
    Dim MySig As String
    
    Sigs = Announcements(GameComponent)
    
    If Sigs = "" Then
        CheckForChanges = False         ' No change announcement made yet
    Else
        MySig = CheckForm.Name & CheckForm.Caption & "+"
        CheckForChanges = (InStr(Sigs, MySig) = 0)
    
        If CheckForChanges Then Announcements(GameComponent) = Sigs & MySig
    End If
    
End Function

Private Sub mnuWindowToolbar_Click()
'
' Name:         mnuWindowToolbar_Click
' Description:  Show or hide the Toolbar as needed.
'

    mnuWindowToolbar.Checked = Not mnuWindowToolbar.Checked
    tlbToolbar.Visible = mnuWindowToolbar.Checked

End Sub

Private Sub mnuWorldItems_Click()
'
' Name:         mnuWorldItems_Click
' Description:  Show the Item Cards window.
'

    frmItemList.Show
    frmItemList.SetFocus
    
End Sub

Private Sub mnuWorldLocations_Click()
'
' Name:         mnuWorldLocations_Click
' Description:  Show the Locations window.
'

    frmLocationList.Show
    frmLocationList.SetFocus
    
End Sub

Private Sub mnuWorldRotes_Click()
'
' Name:         mnuWorldRotes_Click
' Description:  Show the Rotes window.
'

    frmRoteList.Show
    frmRoteList.SetFocus
    
End Sub

' Include the SMTP.ocx control in the project and add an instance to this form,
' naming it smtpMailer, in order to re-enable the code below. Note that SMTP.ocx
' is detected (inaccurately) as a virus component by many antivirus programs.

'Private Sub smtpMailer_ErrorSMTP(ByVal Number As Integer, Description As String)
''
'' Name:         smtpMailer_ErrorSMTP
'' Description:  Set the tag (error field) on a bad send.
''
'    smtpMailer.Tag = TrimWhiteSpace(Description)
'
'End Sub
'
'Private Sub smtpMailer_SendSMTP()
''
'' Name:         smtpMailer_SendSMTP
'' Description:  Clear the tag (error field) on a successful send.
''
'    smtpMailer.Tag = ""
'
'End Sub

'Public Function InitializeSMTP() As Boolean
''
'' Name:         InitializeSMTP
'' Parameters:   Mailer      SMTP control to set up
'' Description:  Retrieve the SMTP settings from the registry.
''
'
'    Dim Password As String
'
'    InitializeSMTP = smtpMailer.Server <> "" And smtpMailer.MailFrom Like "*?@?*.?*" And _
'                     Not (smtpMailer.Username = "" Xor smtpMailer.Password = "")
'
'    If Not InitializeSMTP Then
'
'        smtpMailer.Server = GetSetting(App.Title, "EMail", "Server")
'        smtpMailer.Port = Val(GetSetting(App.Title, "EMail", "Port", "25"))
'        smtpMailer.MailFrom = GetSetting(App.Title, "EMail", "Address")
'        OutputEngine.ReplyTo = GetSetting(App.Title, "EMail", "Reply-To")
'        smtpMailer.Username = GetSetting(App.Title, "EMail", "Username")
'
'        If smtpMailer.Username <> "" Then
'            Password = GetSetting(App.Title, "EMail", "Pwd")
'            smtpMailer.Password = XORScramble(xeKey, XORScramble(smtpMailer.Username, Password))
'        End If
'
'        InitializeSMTP = smtpMailer.Server <> "" And smtpMailer.MailFrom Like "*?@?*.?*" And _
'                         Not (smtpMailer.Username = "" Xor smtpMailer.Password = "")
'    End If
'
'End Function

Private Sub timAutosave_Timer()
'
' Name:         timAutosave_Timer
' Description:  Check for an autosave event.
'

    If AutosaveTime > 0 And mnuSaveGame.Enabled Then
        timAutosave.Tag = CStr(CInt(timAutosave.Tag) + 1)
        If CInt(timAutosave.Tag) >= AutosaveTime Then
        
            Dim DataState As Boolean
            Dim GameFile As String
            
            DataState = Game.DataChanged
            GameFile = Game.GameFile
            timAutosave.Tag = "0"
            
            Screen.MousePointer = vbHourglass
            pgbProgress.Value = 0
            pgbProgress.Max = 101
            pgbProgress.Visible = True
            Game.SaveGame SlashPath(App.Path) & BackupFileName
            pgbProgress.Visible = False
            Screen.MousePointer = vbDefault
            
            Game.DataChanged = DataState
            Game.GameFile = GameFile
            
             If Game.FileError Then
                AutosaveTime = 0
                MsgBox "Error during Autosave:" & vbCrLf & _
                       Game.FileErrorMessage, vbExclamation, "Autosave Error"
            End If
            
        End If
    End If

End Sub

Private Sub tlbToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'
' Name:         tlbToolbar_ButtonClick
' Description:  Call the appropriate menu event when its button is clicked.
'

    Select Case Button.Key
        Case "New"
            Call mnuNewGame_Click
        Case "Open"
            Call mnuOpenGame_Click
        Case "Save"
            Call mnuSaveGame_Click
        Case "Exchange"
            Call mnuFileExchange_Click
        Case "Dates"
            Call mnuGameDates_Click
        Case "Menus"
            Call mnuGrapevineMenus_Click
        Case "Players"
            Call mnuPlayers_Click
        Case "Player Points"
            Call mnuPlayerPoints_Click
        Case "Characters"
            Call mnuCharacters_Click
        Case "Experience"
            Call mnuExperience_Click
        Case "Tempers"
            Call mnuCharTempers_Click
        Case "Harpy"
            Call mnuCharHarpy_Click
        Case "Search"
            Call mnuCharSearch_Click
        Case "Statistics"
            Call mnuCharStats_Click
        Case "Items"
            Call mnuWorldItems_Click
        Case "Rotes"
            Call mnuWorldRotes_Click
        Case "Locations"
            Call mnuWorldLocations_Click
        Case "Actions"
            Call mnuChronicleActions_Click
        Case "Plots"
            Call mnuChroniclePlots_Click
        Case "Rumors"
            Call mnuChronicleRumors_Click
        Case "Print"
            Call mnuOutputPrint_Click
        Case "Export"
            Call mnuOutputSave_Click
        Case "EMail"
            Call mnuOutputEMail_Click
    End Select
    
End Sub


