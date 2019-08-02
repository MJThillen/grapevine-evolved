VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLocationCard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Location"
   ClientHeight    =   5100
   ClientLeft      =   210
   ClientTop       =   405
   ClientWidth     =   8340
   Icon            =   "frmLocationCard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   8340
   Tag             =   "L"
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
      TabIndex        =   57
      Top             =   150
      Width           =   975
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   3615
      Index           =   0
      Left            =   2160
      TabIndex        =   7
      Top             =   1200
      Width           =   5910
      Begin VB.TextBox txtMemo 
         Height          =   3000
         Index           =   1
         Left            =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtMemo 
         Height          =   1185
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   2295
         Width           =   2775
      End
      Begin MSComCtl2.UpDown updNumber 
         Height          =   315
         Index           =   0
         Left            =   2175
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   975
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         OrigLeft        =   2535
         OrigTop         =   1335
         OrigRight       =   3120
         OrigBottom      =   1650
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Appearance"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   17
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Where"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Owner"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   13
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   14
         Tag             =   "?CH"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblNumberLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   990
         Width           =   855
      End
      Begin VB.Label lblNumber 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   9
         Tag             =   "Location Types"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   510
         Width           =   855
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   3615
      Index           =   3
      Left            =   2160
      TabIndex        =   46
      Top             =   1200
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txtMemo 
         Height          =   2010
         Index           =   4
         Left            =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   52
         Top             =   480
         Width           =   2775
      End
      Begin VB.CheckBox chkShowOnlyActive 
         Caption         =   "Show only ""Active"" characters"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   2640
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CommandButton cmdShowCharacter 
         Caption         =   "&Show Character Sheet"
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   3000
         Width           =   2775
      End
      Begin VB.ListBox lstRegulars 
         Height          =   2010
         ItemData        =   "frmLocationCard.frx":058A
         Left            =   120
         List            =   "frmLocationCard.frx":0591
         TabIndex        =   48
         Tag             =   "Tempers"
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   51
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblModified 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3840
         TabIndex        =   54
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblModifiedLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Modified"
         Height          =   375
         Left            =   3000
         TabIndex        =   53
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lblRegulars 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Regulars"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   3585
      Index           =   2
      Left            =   2160
      TabIndex        =   30
      Top             =   1320
      Visible         =   0   'False
      Width           =   5895
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
         Left            =   2400
         TabIndex        =   40
         Top             =   1920
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
         Left            =   360
         TabIndex        =   39
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdIncrement 
         Caption         =   "+"
         Height          =   255
         Left            =   2640
         TabIndex        =   41
         Top             =   1920
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
         TabIndex        =   38
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ListBox lstTraits 
         Height          =   1185
         Index           =   0
         IntegralHeight  =   0   'False
         ItemData        =   "frmLocationCard.frx":05A2
         Left            =   120
         List            =   "frmLocationCard.frx":05A4
         TabIndex        =   43
         Tag             =   "?LO"
         Top             =   2175
         Width           =   2775
      End
      Begin VB.TextBox txtMemo 
         Height          =   3000
         Index           =   3
         Left            =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Top             =   360
         Width           =   2775
      End
      Begin MSComCtl2.UpDown updNumber 
         Height          =   315
         Index           =   3
         Left            =   2175
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1335
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         OrigLeft        =   2535
         OrigTop         =   1335
         OrigRight       =   3120
         OrigBottom      =   1650
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Moon Bridges / Trods"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Totem"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   33
         Top             =   870
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   3
         Left            =   960
         TabIndex        =   34
         Tag             =   "Totems"
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Affinity"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   31
         Top             =   390
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   32
         Tag             =   "Affinity"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblNumberLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gauntlet/ Shroud/ Banality"
         Height          =   615
         Index           =   3
         Left            =   0
         TabIndex        =   35
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblNumber 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   3
         Left            =   960
         TabIndex        =   36
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Umbra / Shadowlands / Dreaming"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   44
         Top             =   120
         Width           =   2775
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   1
      Left            =   2160
      TabIndex        =   19
      Top             =   1320
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txtMemo 
         Height          =   3000
         Index           =   2
         Left            =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   360
         Width           =   2775
      End
      Begin MSComCtl2.UpDown updNumber 
         Height          =   315
         Index           =   1
         Left            =   2175
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   855
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         OrigLeft        =   2535
         OrigTop         =   1335
         OrigRight       =   3120
         OrigBottom      =   1650
         Max             =   999
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updNumber 
         Height          =   315
         Index           =   2
         Left            =   2175
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1335
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         OrigLeft        =   2535
         OrigTop         =   1335
         OrigRight       =   3120
         OrigBottom      =   1650
         Max             =   999
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   4
         Left            =   960
         TabIndex        =   21
         Tag             =   "Access"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Access"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   20
         Top             =   390
         Width           =   855
      End
      Begin VB.Label lblNumber 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   26
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblNumberLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Security Retests"
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   25
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblNumber 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   23
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblNumberLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Security Traits"
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   22
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Security"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   28
         Top             =   120
         Width           =   2775
      End
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
      TabIndex        =   58
      Top             =   120
      Width           =   6135
   End
   Begin VB.CommandButton cmdESCClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdCustom 
      Caption         =   "Custom &Entry"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdNote 
      Caption         =   "Add &Note to Entry"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<- Re&move"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add ->"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   1695
   End
   Begin VB.ListBox lstMenu 
      Height          =   1785
      IntegralHeight  =   0   'False
      ItemData        =   "frmLocationCard.frx":05A6
      Left            =   120
      List            =   "frmLocationCard.frx":05A8
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin MSComctlLib.TabStrip tabTabStrip 
      Height          =   4095
      Left            =   2040
      TabIndex        =   6
      Top             =   840
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7223
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Basics"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Security"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Supernatural"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Regulars, Notes"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMenuItem 
      Height          =   495
      Left            =   240
      TabIndex        =   59
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   480
      Picture         =   "frmLocationCard.frx":05AA
      Top             =   185
      Width           =   480
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
      Left            =   960
      TabIndex        =   56
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblMenuTitle 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmLocationCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Name:         frmLocationCard
' Description:  The screen from which to manipulate Location data.
'
Option Explicit

'Constants by which specific labels are indexed. (fi = Field Index)
Private Const fiType = 0
Private Const fiOwner = 1
Private Const fiAffinity = 2
Private Const fiTotem = 3
Private Const fiAccess = 4

' Constants by which specific list boxes are indexed.
Private Const tiLinks = 0

'Constants by which specific number labels are indexed (ni = Number Index)
Private Const niLevel = 0
Private Const niSecTraits = 1
Private Const niSecRetests = 2
Private Const niGauntlet = 3

' Constants by which specific text boxes are indexed. (xi = Text Index)
Private Const xiName = 0

' Constants by which specific memo fields are indexed. (mi = Memo Index)
Private Const miWhere = 0
Private Const miAppearance = 1
Private Const miSecurity = 2
Private Const miUmbra = 3
Private Const miNotes = 4

'Constants to index important tabs
Private Const tbRegulars = 3

Private Location As LocationClass                   'The Location in question
Private CharSheetEngine As CharacterSheetEngineClass    'Handles common functions
Private Populating As Boolean                           'defuses some events when characters are loaded

Public Sub ShowLocation(Card As LocationClass)
'
' Name:         ShowLocation
' Parameter:    Card        the LocationClass this form displays and modifies.
' Description:  Show and initialize a new instance of the form.
'

    Dim DataState As Boolean

    Populating = True

    Set Location = Nothing
    Set Location = Card
    DataState = Game.DataChanged

    txtUserField(xiName) = Location.Name
    Me.Caption = Location.Name

    updNumber(niLevel) = Location.Level
    lblNumber(niLevel) = " " & CStr(Location.Level) & " " & String(Location.Level, "o")
    updNumber(niSecTraits) = Location.SecTraits
    lblNumber(niSecTraits) = " " & CStr(Location.SecTraits)
    updNumber(niSecRetests) = Location.SecRetests
    lblNumber(niSecRetests) = " " & CStr(Location.SecRetests)
    updNumber(niGauntlet) = Location.Gauntlet
    lblNumber(niGauntlet) = " " & CStr(Location.Gauntlet)
    
    lblField(fiType) = Location.LocType
    lblField(fiOwner) = Location.Owner
    lblField(fiAffinity) = Location.Affinity
    lblField(fiTotem) = Location.Totem
    lblField(fiAccess) = Location.Access
    
    imgIcon.Picture = mdiMain.imlIcons.ListImages(Location.IconKey).Picture
    
    txtMemo(miWhere) = Location.Where
    txtMemo(miAppearance) = Location.Appearance
    txtMemo(miSecurity) = Location.Security
    txtMemo(miUmbra) = Location.Umbra
    txtMemo(miNotes) = Location.Notes
    
    CharSheetEngine.RefreshTraitList lstTraits(tiLinks), Location.LinkList
        
    lblModified.Caption = Format(Location.LastModified, "mmmm d, yyyy")
    
    Me.Show
    
    Game.DataChanged = DataState
    Populating = False

End Sub

Public Sub FindRegulars()
'
' Name:         FindRegulars
' Description:  Populate lstRegulars with characters who frequent this Location.
'

    Dim LocQuery As QueryClass
    Dim LocationText As String
    Dim FList As LinkedTraitList
    Dim StoreCursor As Integer
    
    Screen.MousePointer = vbHourglass
    Set LocQuery = New QueryClass
    StoreCursor = lstRegulars.ListIndex
    
    lstRegulars.Clear
    LocQuery.Inventory = qiCharacters
    LocQuery.MatchAll = True
    
    If chkShowOnlyActive.Value = vbChecked Then _
        LocQuery.AddClause qkPlayStatus, "active", 0, qcEquals, False
        
    LocQuery.AddClause qkLocations, Location.Name, 0, qcContains, False
    
    With Game.QueryEngine
        .MakeQuery LocQuery
    
        .Results.First
        Do Until .Results.Off
            
            Set FList = .Results.Item.HangoutList
            FList.MoveTo Location.Name
            If Not FList.Off Then
                lstRegulars.AddItem .Results.Item.Name
            End If
            
            .Results.MoveNext
        Loop
    End With
    
    lblRegulars.Caption = CStr(lstRegulars.ListCount) & " Regulars"
    
    If StoreCursor >= lstRegulars.ListCount Then StoreCursor = lstRegulars.ListCount - 1
    lstRegulars.ListIndex = StoreCursor
    
    Set LocQuery = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Sub ManageLink(Increment As Boolean)
'
' Name:         ManageLink
' Parameters:   Increment       Add to other LinkList iff TRUE, else remove from other LinkList
' Description:  If a location is being added to LinkList, add it to the other location's LinkList also.
'               If a location is being removed from LinkList, remove it frlom the other location's
'               LinkList also.
'

    If (CharSheetEngine.TargetType = ttListBox) And _
       (CharSheetEngine.TargetList Is lstTraits(tiLinks)) Then
             
        Dim LinkName As String
        Dim LList As LinkedTraitList
         
        Location.LinkList.MoveToPlace lstTraits(tiLinks).ListIndex
        If Location.LinkList.Off Then Exit Sub
        LinkName = Location.LinkList.Trait.Name
        
        If LinkName <> Location.Name Then
    
            LocationList.MoveTo LinkName
            If Not LocationList.Off Then
                Set LList = LocationList.Item.LinkList
                If Increment Then
                    LList.Insert Location.Name
                Else
                    LList.MoveTo Location.Name
                    LList.Remove
                End If
                LocationList.Item.LastModified = Now
            End If
        
        End If
        
    End If

End Sub

Public Sub SetDefaultOutput()
'
' Name:         SetDefaultOutput
' Description:  Initilize the OutputEngineClass with default output settings.
'
    With OutputEngine
        .Template = tnLocationCards
        .SelectSet(osLocations).Clear
        .SelectSet(osLocations).Add Location.Name
        .GameDate = 0
    End With
    
End Sub

Private Sub chkShowOnlyActive_Click()
'
' Name:         chkShowOnlyActive_Click
' Description:  Refresh the items in play list.
'

    FindRegulars

End Sub

Private Sub cmdAdd_Click()
'
' Name:         cmdAdd_Click
' Description:  If a selection is active, have the CharSheetEngine add to
'               the menu.
'

    If lstMenu.ListIndex <> -1 Then
    
        If fraTab(tbRegulars).Visible Then               ' add this card to the selected character
        
            CharacterList.MoveTo lstMenu.Text
            If Not CharacterList.Off Then
                CharacterList.Item.HangoutList.Insert Location.Name
                CharacterList.Item.LastModified = Now
                Game.DataChanged = True
                FindRegulars
            ElseIf Right(lstMenu.Text, 1) = ":" Then
                CharSheetEngine.TargetType = ttNothing
                CharSheetEngine.AddSelected
            End If
        
        Else
            CharSheetEngine.AddSelected
            ManageLink Increment:=True
            SetDataChanged
        End If
        
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

Private Sub cmdCustom_Click()
'
' Name:         cmdCustom_Click
' Description:  Have the CharSheetEngine add a custom entry to the target.
'

    CharSheetEngine.AddCustom
    ManageLink Increment:=True
    SetDataChanged
    
End Sub

Private Sub cmdDecrement_Click()
'
' Name:         cmdDecrement_Click
' Description:  Decrement the selected entry.
'

    If cmdDecrement.Visible Then
        If fraTab(tbRegulars).Visible Then
            Call cmdRemove_Click
            lstRegulars.SetFocus
        Else
            ManageLink Increment:=False
            CharSheetEngine.DecrementEntry
            CharSheetEngine.TargetList.SetFocus
        End If
        SetDataChanged
    End If
    
End Sub

Private Sub cmdIncrement_Click()
'
' Name:         cmdIncrement_Click
' Description:  Increment the selected entry.
'

    If cmdIncrement.Visible Then
        If fraTab(tbRegulars).Visible And lstRegulars.ListIndex <> -1 Then
            
            CharacterList.MoveTo lstRegulars.Text
            If Not CharacterList.Off Then
                CharacterList.Item.EquipmentList.Insert Location.Name
                CharacterList.Item.LastModified = Now
                Game.DataChanged = True
                FindRegulars
            End If
            lstRegulars.SetFocus
            
        Else
            CharSheetEngine.IncrementEntry
            ManageLink Increment:=True
            CharSheetEngine.TargetList.SetFocus
        End If
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
    
    CharSheetEngine.AddNote
    SetDataChanged

End Sub

Private Sub cmdRemove_Click()
'
' Name:         cmdRemove_Click
' Description:  Have the CharSheetEngine remove a label or list entry.
'
    
    If fraTab(tbRegulars).Visible Then   ' Remove this item from the character
    
        If lstRegulars.ListIndex <> -1 Then
    
            CharacterList.MoveTo lstRegulars.Text
            If Not CharacterList.Off Then
                CharacterList.Item.HangoutList.MoveTo Location.Name
                CharacterList.Item.HangoutList.Remove
                CharacterList.Item.LastModified = Now
                Game.DataChanged = True
                FindRegulars
            End If
        
        End If
        
    Else
        
        ManageLink Increment:=False
        CharSheetEngine.RemoveEntry
        SetDataChanged
        
    End If
    
End Sub

Private Sub cmdRename_Click()
'
' Name:         cmdRename_Click
' Description:  Rename the Location.
'

    Dim NewName As String
    
    NewName = InputBox("Enter a new name for the location.", "Rename Location", txtUserField(xiName).Text)
    NewName = Trim(NewName)
    
    If Not (NewName = "" Or NewName = txtUserField(xiName).Text) Then
        LocationList.MoveTo NewName
        If Not LocationList.Off Then
            MsgBox "The name """ & NewName & _
                    """ is already in use.  Please use a different name.", _
                    vbOKOnly Or vbExclamation, "Duplicate Name"
        Else
            Game.Rename qiLocations, txtUserField(xiName).Text, NewName
            txtUserField(xiName).Text = NewName
            mdiMain.AnnounceChanges Me, atLocations
        End If
    End If

End Sub

Private Sub cmdShowCharacter_Click()
'
' Name:         cmdShowCharacter_Click
' Description:  Display the selected possessor of this item.
'

    If lstRegulars.ListIndex > -1 Then
    
        mdiMain.ShowCharacterSheet lstRegulars.Text
    
    End If
    
End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  Find the regular in case they've changed.
'

    lblModified.Caption = Format(Location.LastModified, "mmmm d, yyyy")
    CharSheetEngine.RefreshTraitList lstTraits(tiLinks), Location.LinkList
    If fraTab(tbRegulars).Visible Then FindRegulars
    
End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Checks to make sure that a character is loaded, which happens only
'               when ShowVarious is the means of loading the form.  Initializes the
'               MenuStack linked list and the Various Menus.
'

    If Location Is Nothing Then
        MsgBox "Location Card loaded improperly!"
    Else
        
        Set CharSheetEngine = New CharacterSheetEngineClass
        
        CharSheetEngine.RegisterSheet "Location", lstMenu, lblMenuItem, lblMenuTitle
    
        CharSheetEngine.RegisterTraitList tiLinks, Location.LinkList
    
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
            Case fiType
                Location.LocType = Value
                imgIcon.Picture = mdiMain.imlIcons.ListImages(Location.IconKey).Picture
                mdiMain.AnnounceChanges Me, atLocations
            Case fiOwner
                Location.Owner = Value
            Case fiAffinity
                Location.Affinity = Value
            Case fiTotem
                Location.Totem = Value
            Case fiAccess
                Location.Access = Value
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

Private Sub lblField_DblClick(Index As Integer)
'
' Name:         lblField_DblClick
' Description:  Cross-reference to the selected character.
'

    
    CharSheetEngine.CrossReference

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

Private Sub lstRegulars_DblClick()
'
' Name:         lstRegulars_DblClick
' Description:  Call the cmdShowCharacter button.
'

    Call cmdShowCharacter_Click

End Sub

Private Sub lstTraits_DblClick(Index As Integer)
'
' Name:         lstTraits_DblClick
' Description:  Cross-reference to another item.
'
    CharSheetEngine.CrossReference

End Sub

Private Sub lstTraits_GotFocus(Index As Integer)
'
' Name:         lstTraits_GotFocus
' Description:  Attach the Increment/Decrement buttons, shift focus, populate the menus
'

    Dim OrderTop As Integer

    If Not (CharSheetEngine.TargetType = ttListBox And _
            CharSheetEngine.TargetList Is lstTraits(Index)) Then

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
' Description:  keyboard shortcuts.

    Select Case KeyAscii
        Case Asc("-"), Asc("_")
            cmdDecrement_Click
        Case Asc("+"), Asc("=")
            cmdIncrement_Click
    End Select

End Sub

Private Sub lstRegulars_GotFocus()
'
' Name:         lstRegulars_Click
' Description:  Show Increment/Decrement, Populate the menu with active characters.
'
    
    CharSheetEngine.PopulateMenu "?CH"

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

Private Sub tabTabStrip_Click()
'
' Name:         tabTabStrip_Click
' Description:  Clear the menu and targets.  Display correct frame.
'

    If Not fraTab(tabTabStrip.SelectedItem.Index - 1).Visible Then
        
        CharSheetEngine.DeselectMenus
        
        If CharSheetEngine.TargetType = ttListBox Then
            CharSheetEngine.TargetList.ListIndex = -1
        End If
        lstRegulars.ListIndex = -1
        cmdIncrement.Visible = False
        cmdDecrement.Visible = False
        
        CharSheetEngine.TargetType = ttNothing
        If tabTabStrip.SelectedItem.Index = (tbRegulars + 1) Then
            FindRegulars
        End If
        
        Dim fTab As Frame
        For Each fTab In fraTab
            fTab.Visible = (fTab.Index = tabTabStrip.SelectedItem.Index - 1)
        Next fTab
        
    End If

End Sub

Private Sub txtMemo_Validate(Index As Integer, Cancel As Boolean)
'
' Name:         txtMemo_Change
' Description:  Record changes to the memo field.
'

    Select Case Index
        Case miWhere
            If Location.Where <> txtMemo(miWhere).Text Then
                SetDataChanged
                Location.Where = TrimWhiteSpace(txtMemo(miWhere))
            End If
        Case miAppearance
            If Location.Appearance <> txtMemo(miAppearance).Text Then
                SetDataChanged
                Location.Appearance = TrimWhiteSpace(txtMemo(miAppearance))
            End If
        Case miSecurity
            If Location.Security <> txtMemo(miSecurity).Text Then
                SetDataChanged
                Location.Security = TrimWhiteSpace(txtMemo(miSecurity))
            End If
        Case miUmbra
            If Location.Umbra <> txtMemo(miUmbra).Text Then
                SetDataChanged
                Location.Umbra = TrimWhiteSpace(txtMemo(miUmbra))
            End If
        Case miNotes
            If Location.Notes <> txtMemo(miNotes).Text Then
                SetDataChanged
                Location.Notes = TrimWhiteSpace(txtMemo(miNotes))
            End If
    End Select

End Sub

Private Sub SetDataChanged()
'
' Name:         SetDataChanged
' Description:  Tell the game its data has changed and update the character's
'               Last Modified date.
'
        
    If Not Populating Then
        Game.DataChanged = True
        Location.LastModified = Now
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

    If Not Populating Then
        Select Case Index
            Case niLevel
                Location.Level = updNumber(niLevel).Value
                lblNumber(niLevel).Caption = " " & CStr(Location.Level) & " " & String(Location.Level, "o")
                mdiMain.AnnounceChanges Me, atLocations
            Case niSecTraits
                Location.SecTraits = updNumber(niSecTraits).Value
                lblNumber(niSecTraits).Caption = " " & CStr(Location.SecTraits)
            Case niSecRetests
                Location.SecRetests = updNumber(niSecRetests).Value
                lblNumber(niSecRetests).Caption = " " & CStr(Location.SecRetests)
            Case niGauntlet
                Location.Gauntlet = updNumber(niGauntlet).Value
                lblNumber(niGauntlet).Caption = " " & CStr(Location.Gauntlet)
        End Select
        SetDataChanged
    End If
    
End Sub
