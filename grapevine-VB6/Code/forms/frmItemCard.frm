VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemCard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Card"
   ClientHeight    =   5100
   ClientLeft      =   210
   ClientTop       =   405
   ClientWidth     =   8340
   Icon            =   "frmItemCard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   8340
   Tag             =   "I"
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
      TabIndex        =   49
      Top             =   150
      Width           =   975
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
      TabIndex        =   50
      Top             =   120
      Width           =   6135
   End
   Begin VB.CommandButton cmdESCClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   47
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
      ItemData        =   "frmItemCard.frx":058A
      Left            =   120
      List            =   "frmItemCard.frx":058C
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   3615
      Index           =   0
      Left            =   2160
      TabIndex        =   7
      Top             =   1200
      Width           =   5910
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
         Left            =   3360
         TabIndex        =   53
         Top             =   840
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
         Left            =   5640
         TabIndex        =   52
         Top             =   840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdIncrement 
         Caption         =   "+"
         Height          =   255
         Left            =   5640
         TabIndex        =   24
         Top             =   1080
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
         Left            =   3360
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ListBox lstTraits 
         Height          =   840
         Index           =   0
         ItemData        =   "frmItemCard.frx":058E
         Left            =   3840
         List            =   "frmItemCard.frx":0590
         TabIndex        =   26
         Tag             =   "Tempers"
         Top             =   480
         Width           =   1815
      End
      Begin VB.ListBox lstTraits 
         Height          =   840
         Index           =   2
         ItemData        =   "frmItemCard.frx":0592
         Left            =   3840
         List            =   "frmItemCard.frx":0594
         TabIndex        =   30
         Tag             =   "Negative Traits"
         Top             =   2400
         Width           =   1815
      End
      Begin VB.ListBox lstTraits 
         Height          =   840
         Index           =   1
         ItemData        =   "frmItemCard.frx":0596
         Left            =   3840
         List            =   "frmItemCard.frx":0598
         TabIndex        =   28
         Tag             =   "Abilities"
         Top             =   1440
         Width           =   1815
      End
      Begin MSComCtl2.UpDown updNumber 
         Height          =   315
         Index           =   0
         Left            =   2175
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1455
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
      Begin MSComCtl2.UpDown updFieldNumber 
         Height          =   315
         Index           =   1
         Left            =   2175
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2415
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         OrigLeft        =   2535
         OrigTop         =   1335
         OrigRight       =   3120
         OrigBottom      =   1650
         Max             =   99
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updNumber 
         Height          =   315
         Index           =   1
         Left            =   2175
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1935
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         OrigLeft        =   2535
         OrigTop         =   1335
         OrigRight       =   3120
         OrigBottom      =   1650
         Max             =   99
         Min             =   -99
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin VB.Label lblNumber 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   16
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblNumberLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   15
         Top             =   1950
         Width           =   855
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Subtype"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   10
         Top             =   990
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   3
         Left            =   960
         TabIndex        =   11
         Tag             =   "Item Subtypes"
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblTraits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Negatives"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   29
         Top             =   2430
         Width           =   855
      End
      Begin VB.Label lblTraits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Abilities"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   27
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label lblTraits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tempers"
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   25
         Top             =   510
         Width           =   855
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Damage/ Armor"
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   18
         Top             =   2340
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   19
         Tag             =   "Damage"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblNumberLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label lblNumber 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   9
         Tag             =   "Item Types"
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
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Conceal"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   21
         Top             =   2910
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   22
         Tag             =   "Concealability"
         Top             =   2880
         Width           =   1815
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   3615
      Index           =   1
      Left            =   2160
      TabIndex        =   31
      Top             =   1200
      Visible         =   0   'False
      Width           =   5895
      Begin VB.ListBox lstTraits 
         Height          =   1185
         Index           =   3
         IntegralHeight  =   0   'False
         ItemData        =   "frmItemCard.frx":059A
         Left            =   120
         List            =   "frmItemCard.frx":059C
         TabIndex        =   37
         Tag             =   "Availability"
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox txtMemo 
         Height          =   1185
         Index           =   2
         Left            =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Top             =   2175
         Width           =   2775
      End
      Begin VB.TextBox txtMemo 
         Height          =   1185
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   495
         Width           =   2775
      End
      Begin VB.TextBox txtMemo 
         Height          =   1185
         Index           =   1
         Left            =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   495
         Width           =   2775
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Availability"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   36
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   38
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Powers"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Appearance"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   34
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   3615
      Index           =   2
      Left            =   2160
      TabIndex        =   40
      Top             =   1200
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CheckBox chkShowOnlyActive 
         Caption         =   "Show only ""Active"" characters"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   2640
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CommandButton cmdShowCharacter 
         Caption         =   "&Show Character Sheet"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   3000
         Width           =   2775
      End
      Begin VB.ListBox lstOwners 
         Height          =   2010
         ItemData        =   "frmItemCard.frx":059E
         Left            =   120
         List            =   "frmItemCard.frx":05A5
         TabIndex        =   42
         Tag             =   "Tempers"
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label lblModified 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3840
         TabIndex        =   46
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblModifiedLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Modified"
         Height          =   375
         Left            =   3000
         TabIndex        =   45
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblOwners 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "In Play"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   2775
      End
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
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Basics"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Details"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "In Play"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMenuItem 
      Height          =   495
      Left            =   240
      TabIndex        =   51
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   480
      Picture         =   "frmItemCard.frx":05B4
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
      TabIndex        =   48
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
Attribute VB_Name = "frmItemCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Name:         frmItemCard
' Description:  The screen in which to manipulate Item Card data.
'
Option Explicit

'Constants by which specific labels are indexed. (fi = Field Index)
Private Const fiType = 0
Private Const fiDamage = 1
Private Const fiConceal = 2
Private Const fiSubtype = 3

' Constants by which specific list boxes are indexed.
Private Const tiTempers = 0
Private Const tiAbilities = 1
Private Const tiNegatives = 2
Private Const tiAvailability = 3

'Constants by which specific number labels are indexed (ni = Number Index)
Private Const niLevel = 0
Private Const niBonus = 1

' Constants by which specific text boxes are indexed. (xi = Text Index)
Private Const xiName = 0

' Constants by which specific memo fields are indexed. (mi = Memo Index)
Private Const miPowers = 0
Private Const miAppearance = 1
Private Const miNotes = 2

'Constants to index important tabs
Private Const tbOwners = 2

Private Item As ItemClass           'The Item Card
Private CharSheetEngine As CharacterSheetEngineClass    'Handles common functions
Private Populating As Boolean                           'defuses some events when characters are loaded

Public Sub ShowItem(Card As ItemClass)
'
' Name:         ShowItem
' Parameter:    Card        the ItemClass this form displays and modifies.
' Description:  Show and initialize a new instance of the form.
'

    Dim DataState As Boolean

    Populating = True

    Set Item = Nothing
    Set Item = Card
    DataState = Game.DataChanged

    txtUserField(xiName) = Item.Name
    Me.Caption = Item.Name

    updNumber(niLevel) = Item.Level
    updNumber(niBonus) = Item.Bonus
    updFieldNumber(fiDamage) = Item.DamageAmount
    
    imgIcon.Picture = mdiMain.imlIcons.ListImages(Item.IconKey).Picture
    
    lblField(fiType) = Item.ItemType
    lblField(fiSubtype) = Item.ItemSubtype
    lblNumber(niLevel) = " " & CStr(Item.Level) & " " & String(Item.Level, "o")
    lblNumber(niBonus) = IIf(Item.Bonus > 0, " +", " ") & CStr(Item.Bonus)
    lblField(fiDamage) = IIf(Item.DamageType <> "", " " & CStr(Item.DamageAmount) & " ", "") & Item.DamageType
    lblField(fiConceal) = Item.Concealability
    
    txtMemo(miPowers) = Item.Powers
    txtMemo(miAppearance) = Item.Appearance
    txtMemo(miNotes) = Item.Notes
    
    CharSheetEngine.RefreshTraitList lstTraits(tiTempers), Item.TemperList
    CharSheetEngine.RefreshTraitList lstTraits(tiAbilities), Item.AbilityList
    CharSheetEngine.RefreshTraitList lstTraits(tiNegatives), Item.NegativeList
    CharSheetEngine.RefreshTraitList lstTraits(tiAvailability), Item.Availability
        
    lblModified.Caption = Format(Item.LastModified, "mmmm d, yyyy")
    
    Me.Show
    
    Game.DataChanged = DataState
    Populating = False

End Sub

Public Sub FindOwners()
'
' Name:         FindOwners
' Description:  Populat lstOwners with characters who possess this item.
'

    Dim OwnQuery As QueryClass
    Dim ItemText As String
    Dim ItemCount As Integer
    Dim EquipList As LinkedTraitList
    Dim StoreCursor As Integer
    
    Screen.MousePointer = vbHourglass
    Set OwnQuery = New QueryClass
    StoreCursor = lstOwners.ListIndex
        
    lstOwners.Clear
    OwnQuery.Inventory = qiCharacters
    OwnQuery.MatchAll = True
    
    If chkShowOnlyActive.Value = vbChecked Then _
        OwnQuery.AddClause qkPlayStatus, "active", 0, qcEquals, False
        
    OwnQuery.AddClause qkEquipment, Item.Name, 0, qcContains, False
    
    With Game.QueryEngine
        .MakeQuery OwnQuery
    
        ItemCount = 0
        .Results.First
        Do Until .Results.Off
            
            Set EquipList = .Results.Item.EquipmentList
            EquipList.MoveTo Item.Name
            If Not EquipList.Off Then
                lstOwners.AddItem .Results.Item.Name & " x" & EquipList.Trait.Total
                ItemCount = ItemCount + EquipList.Trait.Number
            End If
            
            .Results.MoveNext
        Loop
    
    End With
    
    lblOwners.Caption = CStr(ItemCount) & " in Play"
    
    If StoreCursor >= lstOwners.ListCount Then StoreCursor = lstOwners.ListCount - 1
    lstOwners.ListIndex = StoreCursor
   
    Set OwnQuery = Nothing
    Screen.MousePointer = vbDefault

End Sub

Public Sub SetDefaultOutput()
'
' Name:         SetDefaultOutput
' Description:  Initilize the OutputEngineClass with default output settings.
'
    With OutputEngine
        .Template = tnItemCards
        .SelectSet(osItems).Clear
        .SelectSet(osItems).Add Item.Name
        .GameDate = 0
    End With
    
End Sub

Private Sub chkShowOnlyActive_Click()
'
' Name:         chkShowOnlyActive_Click
' Description:  Refresh the items in play list.
'

    FindOwners

End Sub

Private Sub cmdAdd_Click()
'
' Name:         cmdAdd_Click
' Description:  If a selection is active, have the CharSheetEngine add to
'               the menu.
'

    If lstMenu.ListIndex <> -1 Then
    
        If fraTab(tbOwners).Visible Then               ' add this card to the selected character
        
            CharacterList.MoveTo lstMenu.Text
            If Not CharacterList.Off Then
                CharacterList.Item.EquipmentList.Insert Item.Name, , Item.CombatNote
                CharacterList.Item.LastModified = Now
                Game.DataChanged = True
                FindOwners
            ElseIf Right(lstMenu.Text, 1) = ":" Then
                CharSheetEngine.TargetType = ttNothing
                CharSheetEngine.AddSelected
            End If
        
        Else
            CharSheetEngine.AddSelected
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
    SetDataChanged

End Sub

Private Sub cmdDecrement_Click()
'
' Name:         cmdDecrement_Click
' Description:  Decrement the selected entry.
'

    If cmdDecrement.Visible Then
        If fraTab(tbOwners).Visible Then
            Call cmdRemove_Click
            lstOwners.SetFocus
        Else
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
        If fraTab(tbOwners).Visible And lstOwners.ListIndex <> -1 Then
            
            Dim TrimName As String
    
            TrimName = lstOwners.Text
            TrimName = Left(TrimName, InStrRev(TrimName, " x") - 1)
            CharacterList.MoveTo TrimName
            If Not CharacterList.Off Then
                CharacterList.Item.EquipmentList.Insert Item.Name
                CharacterList.Item.LastModified = Now
                Game.DataChanged = True
                FindOwners
            End If
            lstOwners.SetFocus
            
        Else
            CharSheetEngine.IncrementEntry
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
    
    If fraTab(tbOwners).Visible Then   ' Remove this item from the character
    
        If lstOwners.ListIndex <> -1 Then
    
            Dim TrimName As String
    
            TrimName = lstOwners.Text
            TrimName = Left(TrimName, InStrRev(TrimName, " x") - 1)
            CharacterList.MoveTo TrimName
            If Not CharacterList.Off Then
                CharacterList.Item.EquipmentList.MoveTo Item.Name
                CharacterList.Item.EquipmentList.Remove
                CharacterList.Item.LastModified = Now
                Game.DataChanged = True
                FindOwners
            End If
        
        End If
        
    Else
    
        CharSheetEngine.RemoveEntry
        SetDataChanged
    
    End If
    
End Sub

Private Sub cmdRename_Click()
'
' Name:         cmdRename_Click
' Description:  Rename the item.
'

    Dim NewName As String
    
    NewName = InputBox("Enter a new name for the item.", "Rename Item", txtUserField(xiName).Text)
    NewName = Trim(NewName)
    
    If Not (NewName = "" Or NewName = txtUserField(xiName).Text) Then
        ItemList.MoveTo NewName
        If Not ItemList.Off Then
            MsgBox "The name """ & NewName & _
                    """ is already in use.  Please use a different name.", _
                    vbOKOnly Or vbExclamation, "Duplicate Name"
        Else
            Game.Rename qiItems, txtUserField(xiName).Text, NewName
            txtUserField(xiName).Text = NewName
            mdiMain.AnnounceChanges Me, atItems
        End If
    End If

End Sub

Private Sub cmdShowCharacter_Click()
'
' Name:         cmdShowCharacter_Click
' Description:  Display the selected possessor of this item.
'

    Dim TrimName As String
    
    If lstOwners.ListIndex > -1 Then
    
        TrimName = lstOwners.Text
        TrimName = Left(TrimName, InStrRev(TrimName, " x") - 1)
        mdiMain.ShowCharacterSheet TrimName
    
    End If
    
End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  Find the owners in case they've changed.
'
       
    If fraTab(tbOwners).Visible Then FindOwners
    
End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Checks to make sure that a character is loaded, which happens only
'               when ShowVarious is the means of loading the form.  Initializes the
'               MenuStack linked list and the Various Menus.
'

    If Item Is Nothing Then
        MsgBox "Item Card loaded improperly!"
    Else
                
        Set CharSheetEngine = New CharacterSheetEngineClass
        
        CharSheetEngine.RegisterSheet "Item", lstMenu, lblMenuItem, lblMenuTitle
    
        CharSheetEngine.RegisterTraitList tiTempers, Item.TemperList
        CharSheetEngine.RegisterTraitList tiAbilities, Item.AbilityList
        CharSheetEngine.RegisterTraitList tiNegatives, Item.NegativeList
        CharSheetEngine.RegisterTraitList tiAvailability, Item.Availability
    
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
                Item.ItemType = Value
                imgIcon.Picture = mdiMain.imlIcons.ListImages(Item.IconKey).Picture
                mdiMain.AnnounceChanges Me, atItems
            Case fiSubtype
                Item.ItemSubtype = Value
                mdiMain.AnnounceChanges Me, atItems
            Case fiDamage
                Item.DamageType = Value
                If Value <> "" Then
                    Populating = True
                    lblField(Index).Caption = " " & CStr(Item.DamageAmount) & " " & Value
                    Populating = False
                End If
            Case fiConceal
                Item.Concealability = Value
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

Private Sub lstOwners_GotFocus()
'
' Name:         lstOwners_Click
' Description:  Show Increment/Decrement, Populate the menu with active characters.
'
    
    With lstOwners
        Set cmdDecrement.Container = .Container
        Set cmdIncrement.Container = .Container
        cmdDecrement.Move .Left, .Top - cmdDecrement.Height
        cmdIncrement.Move .Left + .Width - cmdIncrement.Width, .Top + -cmdIncrement.Height
    End With
    cmdIncrement.Visible = True
    cmdDecrement.Visible = True
    
    CharSheetEngine.PopulateMenu "?CH"

End Sub

Private Sub lstOwners_DblClick()
'
' Name:         lstOwners_DblClick
' Description:  Call the cmdShowCharacter button.
'

    Call cmdShowCharacter_Click

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
                If Index = tiAvailability Then
                    cmdDecrement.Move .Left, .Top - cmdDecrement.Height
                    cmdIncrement.Move .Left + .Width - cmdIncrement.Width, .Top - cmdIncrement.Height
                    OrderTop = .Top + .Height
                Else
                    cmdDecrement.Move .Left - cmdDecrement.Width, .Top + .Height - cmdDecrement.Height
                    cmdIncrement.Move .Left + .Width, .Top + .Height - cmdIncrement.Height
                    OrderTop = cmdDecrement.Top - cmdDescend.Height
                End If
            End With
            cmdIncrement.Visible = True
            cmdDecrement.Visible = True
        Else
            cmdIncrement.Visible = False
            cmdDecrement.Visible = False
            If Index = tiAvailability Then
                OrderTop = lstTraits(Index).Top - cmdAscend.Height
            Else
                OrderTop = lstTraits(Index).Top + lstTraits(Index).Height - cmdDescend.Height
            End If
        End If
        
        If CharSheetEngine.CanOrder(Index) Then
            With lstTraits(Index)
                Set cmdDescend.Container = .Container
                Set cmdAscend.Container = .Container
                If Index = tiAvailability Then
                    cmdDescend.Move .Left, OrderTop
                    cmdAscend.Move .Left + .Width - cmdAscend.Width, OrderTop
                Else
                    cmdDescend.Move .Left - cmdDescend.Width, OrderTop
                    cmdAscend.Move .Left + .Width, OrderTop
                End If
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
        lstOwners.ListIndex = -1
        cmdIncrement.Visible = False
        cmdDecrement.Visible = False
        cmdAscend.Visible = False
        cmdDescend.Visible = False
        
        CharSheetEngine.TargetType = ttNothing
        If tabTabStrip.SelectedItem.Index = (tbOwners + 1) Then
            FindOwners
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
        Case miPowers
            If Item.Powers <> txtMemo(miPowers).Text Then
                SetDataChanged
                Item.Powers = TrimWhiteSpace(txtMemo(miPowers))
            End If
        Case miAppearance
            If Item.Appearance <> txtMemo(miAppearance).Text Then
                SetDataChanged
                Item.Appearance = TrimWhiteSpace(txtMemo(miAppearance))
            End If
        Case miNotes
            If Item.Notes <> txtMemo(miNotes).Text Then
                SetDataChanged
                Item.Notes = TrimWhiteSpace(txtMemo(miNotes))
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
        Item.LastModified = Now
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

Private Sub updFieldNumber_Change(Index As Integer)
'
' Name:         updNumber_Change
' Description:  Update the label and store the new value.
'

    If Not Populating Then
        Select Case Index
            Case fiDamage
                Populating = True
                Item.DamageAmount = updFieldNumber(fiDamage).Value
                lblField(fiDamage).Caption = _
                    IIf(Item.DamageType <> "", " " & CStr(Item.DamageAmount) & " ", "") & Item.DamageType
                Populating = False
        End Select
        SetDataChanged
    End If
    
End Sub

Private Sub updNumber_Change(Index As Integer)
'
' Name:         updNumber_Change
' Description:  Update the label and store the new value.
'

    If Not Populating Then
        Select Case Index
            Case niLevel
                Item.Level = updNumber(niLevel).Value
                lblNumber(niLevel).Caption = " " & CStr(Item.Level) & " " & String(Item.Level, "o")
                mdiMain.AnnounceChanges Me, atItems
            Case niBonus
                Item.Bonus = updNumber(niBonus).Value
                lblNumber(niBonus).Caption = IIf(Item.Bonus > 0, " +", " ") & CStr(Item.Bonus)
        End Select
        SetDataChanged
    End If
    
End Sub

