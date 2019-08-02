VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHarpyLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vampire Boon and Status Management"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "frmHarpyLedger.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   9060
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Height          =   5385
      Index           =   0
      Left            =   2160
      TabIndex        =   13
      Top             =   600
      Width           =   6615
      Begin VB.CommandButton cmdAddBoon 
         Caption         =   "Add Ne&w Boon"
         Height          =   375
         Left            =   3000
         TabIndex        =   14
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdDeleteBoon 
         Caption         =   "&Delete Boon"
         Height          =   375
         Left            =   4800
         TabIndex        =   15
         Top             =   120
         Width           =   1695
      End
      Begin MSComctlLib.ListView lvwBoons 
         Height          =   2055
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3625
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   2073
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Due From"
            Object.Width           =   3237
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Owed To"
            Object.Width           =   3237
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   2073
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "SortDate"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Frame fraBoons 
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   0
         TabIndex        =   40
         Top             =   2640
         Width           =   6615
         Begin VB.TextBox txtMemo 
            Height          =   2055
            Index           =   0
            Left            =   3360
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   25
            Top             =   480
            Width           =   3135
         End
         Begin VB.Label lblField 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   960
            TabIndex        =   42
            Tag             =   "?DT"
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   41
            Top             =   1950
            Width           =   855
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   18
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
            TabIndex        =   19
            Tag             =   "Boons"
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Due from"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   20
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
            TabIndex        =   21
            Tag             =   "?BB"
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Owed to"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   22
            Top             =   1470
            Width           =   855
         End
         Begin VB.Label lblField 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   960
            TabIndex        =   23
            Tag             =   "?BB"
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lblMemo 
            Alignment       =   2  'Center
            Caption         =   "De&scription"
            Height          =   255
            Index           =   0
            Left            =   3360
            TabIndex        =   24
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Label lblBoonCount 
         BackStyle       =   0  'Transparent
         Caption         =   "0 B&oons"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Height          =   5385
      Index           =   1
      Left            =   2160
      TabIndex        =   34
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Frame fraStatus 
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   0
         TabIndex        =   38
         Top             =   2760
         Width           =   6615
         Begin VB.CommandButton cmdShow 
            Caption         =   "S&how Character"
            Height          =   375
            Left            =   960
            TabIndex        =   6
            Top             =   1320
            Width           =   2175
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
            TabIndex        =   9
            Top             =   120
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton cmdIncrement 
            Caption         =   "+"
            Height          =   255
            Left            =   6240
            TabIndex        =   10
            Top             =   120
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
            Left            =   6240
            TabIndex        =   12
            Top             =   2280
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
            Left            =   3360
            TabIndex        =   11
            Top             =   2280
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.ListBox lstTraits 
            Height          =   2010
            Index           =   0
            Left            =   3360
            TabIndex        =   8
            Tag             =   "Status"
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   2
            Top             =   390
            Width           =   855
         End
         Begin VB.Label lblName 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   960
            TabIndex        =   3
            Tag             =   "Title"
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label lblTraits 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0 &Status"
            Height          =   255
            Index           =   0
            Left            =   3360
            TabIndex        =   7
            Tag             =   "Status"
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Title"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   4
            Top             =   870
            Width           =   855
         End
         Begin VB.Label lblField 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   5
            Tag             =   "Title"
            Top             =   840
            Width           =   2175
         End
      End
      Begin VB.ComboBox cboSearch 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   240
         Width           =   2415
      End
      Begin VB.OptionButton optNot 
         Caption         =   "Vampires n&ot among"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   36
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton optNot 
         Caption         =   "Vampires among"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   35
         Top             =   120
         Value           =   -1  'True
         Width           =   1815
      End
      Begin MSComctlLib.ListView lvwVampires 
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3625
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   3228
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   1217
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Title"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Clan"
            Object.Width           =   2037
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Sect"
            Object.Width           =   1587
         EndProperty
      End
      Begin VB.Label lblVampireCount 
         BackStyle       =   0  'Transparent
         Caption         =   "0 &Vampire Characters"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdESCClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.ListBox lstMenu 
      Height          =   2010
      ItemData        =   "frmHarpyLedger.frx":014A
      Left            =   120
      List            =   "frmHarpyLedger.frx":014C
      TabIndex        =   27
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add ->"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<- Re&move"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdNote 
      Caption         =   "Add &Note to Entry"
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdCustom 
      Caption         =   "Custom &Entry"
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   5640
      Width           =   1695
   End
   Begin MSComctlLib.TabStrip tabTabStrip 
      Height          =   5775
      Left            =   2040
      TabIndex        =   33
      Top             =   240
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   10186
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  &Boons  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  Stat&us  "
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   960
      Picture         =   "frmHarpyLedger.frx":014E
      ToolTipText     =   "It's a harpy, can't you tell?"
      Top             =   185
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   480
      Picture         =   "frmHarpyLedger.frx":0A18
      Top             =   185
      Width           =   480
   End
   Begin VB.Label lblMenuTitle 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblMenuItem 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   120
      TabIndex        =   28
      Top             =   3240
      Width           =   1695
   End
End
Attribute VB_Name = "frmHarpyLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Constants by which specific labels are indexed. (fi = Field Index)
Private Const fiTitle = 0
Private Const fiType = 1
Private Const fiDueFrom = 2
Private Const fiOwedTo = 3
Private Const fiDate = 4

' Constants by which specific list boxes are indexed.
Private Const tiStatus = 0

' Constants by which specific multiline text boxes are indexed.
Private Const miDescription = 0

' Constants by which selection controls are indexed
Private Const OPT_MATCH = 0
Private Const OPT_NO_MATCH = 1

Private Const FRAME_BOONS = 0
Private Const FRAME_STATUS = 1

Private Vampire As VampireClass                         'The Vampire character selected for Status manipulation
Private OwedTo As VampireClass                          'Vampire Character to whom a boon is owed
Private DueFrom As VampireClass                         'Vampire character from whom a boon is due
Private OwedBoon As BoonClass                           'Boon object owed to the other
Private DueBoon As BoonClass                            'Boon object due from the other
Private BoonIDSet As StringSet                          'Set of all Boon IDs

Private Const keyOWE = "owedto"
Private Const keyDUE = "duefrom"
Private Const keyDATE = "date"
Private Const keySORT = "sort"

Private CharSheetEngine As CharacterSheetEngineClass    'Handles common functions
Private Populating As Boolean                           'defuses some events when characters are loaded

Private Sub RefreshVampires()
'
' Name:         RefreshVampires
' Description:  Preserving the current selection, this refills the list box from the list of
'               vampires according to the chosen search.
'

    Dim StoreSelKey As String
    Dim Search As QueryClass
    Dim Leech As VampireClass
    Dim NewItem As ListItem
    Dim HighStat As Integer
    Dim NoKey As Boolean
    
    Screen.MousePointer = vbHourglass
    HighStat = -1
    
    If Not (lvwVampires.SelectedItem Is Nothing) Then _
            StoreSelKey = lvwVampires.SelectedItem.Key
    NoKey = (StoreSelKey = "")
    
    lvwVampires.ListItems.Clear
    
    With Game.QueryEngine
        
        .QueryList.MoveTo cboSearch.Text
        If Not .QueryList.Off Then
            Set Search = .QueryList.Item
        Else
            Set Search = New QueryClass
            Search.Inventory = qiCharacters
        End If
    
        .MakeQuery Search, , optNot(OPT_NO_MATCH).Value
    
        .Results.First
        Do Until .Results.Off
            If .Results.Item.RaceCode = gvRaceVampire Then
                Set Leech = .Results.Item
                Set NewItem = lvwVampires.ListItems.Add(Key:="k" & Leech.Name, Text:=Leech.Name)
                NewItem.ListSubItems.Add Text:=Leech.StatusList.Count
                NewItem.ListSubItems.Add Text:=Leech.Title
                NewItem.ListSubItems.Add Text:=Leech.Clan
                NewItem.ListSubItems.Add Text:=Leech.Sect
                If Leech.StatusList.Count > HighStat Then
                    If NoKey Then StoreSelKey = "k" & Leech.Name
                    HighStat = Leech.StatusList.Count
                End If
            End If
            .Results.MoveNext
        Loop
        
    End With
    
    lblVampireCount.Caption = CStr(lvwVampires.ListItems.Count) & " &Vampire Characters"
    
    On Error Resume Next
    Set lvwVampires.SelectedItem = lvwVampires.ListItems(StoreSelKey)
    If lvwVampires.SelectedItem Is Nothing And lvwVampires.ListItems.Count > 0 Then _
        Set lvwVampires.SelectedItem = lvwVampires.GetFirstVisible
    lvwVampires.SelectedItem.EnsureVisible
    On Error GoTo 0

    lvwVampires_ItemClick lvwVampires.SelectedItem

    Screen.MousePointer = vbDefault

    Set Search = Nothing

End Sub

Private Sub RefreshSearches(Default As String)
'
' Name:         RefreshSearches
' Parameters:   Default         Default search to use
' Description:  Fill cboSearch from the QueryList.  Exclude searches that are only on Race.
'               This will force either a cboSearch_Click event (and thus a RefreshVampires) or
'               a RefreshVampires directly.
'

    Dim I As Integer
    Dim Q As QueryClass
    
    cboSearch.Clear
    
    With Game.QueryEngine.QueryList
        .First
        Do Until .Off
            Set Q = .Item
            Q.First
            If Q.Inventory = qiCharacters Then
                If Q.ClauseCount = 1 Then
                    If Q.Clause.Key <> qkRace Then cboSearch.AddItem Q.Name
                Else
                    cboSearch.AddItem Q.Name
                End If
            End If
            .MoveNext
        Loop
    End With
    
    For I = 0 To cboSearch.ListCount - 1
        If cboSearch.List(I) = Default Then
            cboSearch.ListIndex = I                     'Triggers cboSearch_Click,
            Exit For                                    'which populates the list.
        End If
    Next I

    If Not cboSearch.ListIndex >= 0 Then RefreshVampires 'Force the list populous with all characters

End Sub

Private Sub RefreshCurrentVampire()
'
' Name:         RefreshCurrentVampire
' Description:  When a vampire's info changes, change his list entry.
'

    If Not (lvwVampires.SelectedItem Is Nothing Or Vampire Is Nothing) Then
        lvwVampires.SelectedItem.SubItems(1) = Vampire.StatusList.Count
        lvwVampires.SelectedItem.SubItems(2) = Vampire.Title
    End If

End Sub

Private Sub RefreshBoons()
'
' Name:         RefreshBoons
' Description:  Refresh the list of boons.
'

    Dim Leech As VampireClass
    Dim Boon As BoonClass
    Dim ID As String
    Dim NewItem As ListItem
    Dim StoreSelKey As String
    Dim TestNum As Long
    
    Screen.MousePointer = vbHourglass
    
    If Not (lvwBoons.SelectedItem Is Nothing) Then StoreSelKey = lvwBoons.SelectedItem.Key
    
    BoonIDSet.Clear
    lvwBoons.ListItems.Clear
    TestNum = 0
    
    CharacterList.First
    Do Until CharacterList.Off
        If CharacterList.Item.RaceCode = gvRaceVampire Then
            Set Leech = CharacterList.Item
            Leech.BoonList.First
            Do Until Leech.BoonList.Off
                Set Boon = Leech.BoonList.Item
                ID = Boon.BoonID(Leech.Name)
                If Not BoonIDSet.Has(ID) Then
                    BoonIDSet.Add ID
                    Set NewItem = lvwBoons.ListItems.Add(Text:=Boon.BoonType)
                    NewItem.Tag = ID
                    If Boon.IsOwed Then
                        NewItem.ListSubItems.Add Text:=Leech.Name, Key:=keyDUE
                        NewItem.ListSubItems.Add Text:=Boon.CharName, Key:=keyOWE
                    Else
                        NewItem.ListSubItems.Add Text:=Boon.CharName, Key:=keyDUE
                        NewItem.ListSubItems.Add Text:=Leech.Name, Key:=keyOWE
                    End If
                    NewItem.ListSubItems.Add Text:=Format(Boon.BoonDate, "Short Date"), Key:=keyDATE
                    NewItem.ListSubItems.Add Text:=Format(Boon.BoonDate, "yyyy-mm-dd"), Key:=keySORT
                Else
                    TestNum = TestNum + 1
                End If
                Leech.BoonList.MoveNext
            Loop
        End If
        CharacterList.MoveNext
    Loop

    lblBoonCount.Caption = CStr(lvwBoons.ListItems.Count) & " Boons"
    If TestNum <> BoonIDSet.Count Then lblBoonCount.Caption = lblBoonCount.Caption & "."

    On Error Resume Next
    Set lvwBoons.SelectedItem = lvwBoons.ListItems(StoreSelKey)
    If lvwBoons.SelectedItem Is Nothing And lvwBoons.ListItems.Count > 0 Then _
        Set lvwBoons.SelectedItem = lvwBoons.GetFirstVisible
    lvwBoons.SelectedItem.EnsureVisible
    On Error GoTo 0

    lvwBoons_ItemClick lvwBoons.SelectedItem

    Screen.MousePointer = vbDefault

End Sub

Public Sub SetDefaultOutput()
'
' Name:         SetDefaultOutput
' Description:  Initilize the OutputEngineClass with default output settings.
'
    With OutputEngine
        .Template = tnVampireStatus
        .SelectSet(osCharacters).Clear
        .SelectSet(osCharacters).StoreListView lvwVampires, True
        .SearchName = cboSearch.Text
        .SearchNot = optNot(OPT_NO_MATCH).Value
        .GameDate = 0
    End With

End Sub

Private Sub cboSearch_Click()
'
' Name:         cboSearch_Click
' Description:  The user has chosen a new query, so populate the list.
'
    RefreshVampires
    
End Sub

Private Sub cmdAdd_Click()
'
' Name:         cmdAdd_Click
' Description:  If a selection is active, have the CharSheetEngine add to
'               the menu.
'

    If Not (lstMenu.ListIndex = -1 Or Vampire Is Nothing) Then
        
        CharSheetEngine.AddSelected
        SetDataChanged
    
    End If
    
End Sub

Private Sub cmdAddBoon_Click()
'
' Name:         cmdAddBoon_Click
' Description:  Add a new boon to the list.
'

    Set lvwBoons.SelectedItem = lvwBoons.ListItems.Add(Text:="New Boon")
    lvwBoons.SelectedItem.ListSubItems.Add Text:="", Key:=keyDUE
    lvwBoons.SelectedItem.ListSubItems.Add Text:="", Key:=keyOWE
    lvwBoons.SelectedItem.ListSubItems.Add Text:=Format(Now, "Short Date"), Key:=keyDATE
    lvwBoons.SelectedItem.ListSubItems.Add Text:=Format(Now, "yyyy-mm-dd"), Key:=keySORT

    Set OwedTo = Nothing
    Set DueFrom = Nothing
    Set OwedBoon = Nothing
    Set DueBoon = Nothing
    
    lblBoonCount.Caption = CStr(lvwBoons.ListItems.Count) & " Boons"

    lvwBoons_ItemClick lvwBoons.SelectedItem
    If Not lvwBoons.SelectedItem Is Nothing Then Call lblField_Click(fiType)
    
End Sub

Private Sub cmdAscend_Click()
'
' Name:         cmdAscend_Click
' Description:  Move the selected entry down.
'

    If cmdAscend.Visible And Not Vampire Is Nothing Then
        CharSheetEngine.MoveEntryUp
        CharSheetEngine.TargetList.SetFocus
        SetDataChanged
    End If
    
End Sub

Private Sub cmdDeleteBoon_Click()
'
' Name:         cmdDeleteBoon_Click
' Description:  Delete the selected boon from the game.
'

    If Not lvwBoons.SelectedItem Is Nothing Then

        If MsgBox("Are you sure you want to delete this boon?", _
                  vbYesNo + vbQuestion, "Delete Boon") = vbYes Then
                  
            Dim Index As Long
            Dim ID As String
            
            Index = lvwBoons.SelectedItem.Index
            ID = lvwBoons.SelectedItem.Tag
            
            lvwBoons.ListItems.Remove Index
            
            If Not DueFrom Is Nothing Then
                If Not DueBoon Is Nothing Then
                    With DueFrom.BoonList
                        .First
                        Do Until .Off
                            If ID = .Item.BoonID(DueFrom.Name) Then Exit Do
                            .MoveNext
                        Loop
                        If Not .Off Then .Remove
                    End With
                    Set DueBoon = Nothing
                End If
                DueFrom.RefreshBoonTraits
                Set DueFrom = Nothing
            End If
            
            If Not OwedTo Is Nothing Then
                If Not OwedBoon Is Nothing Then
                    With OwedTo.BoonList
                        .First
                        Do Until .Off
                            If ID = .Item.BoonID(OwedTo.Name) Then Exit Do
                            .MoveNext
                        Loop
                        If Not .Off Then .Remove
                    End With
                    Set OwedBoon = Nothing
                End If
                OwedTo.RefreshBoonTraits
                Set OwedTo = Nothing
            End If
            
            lblBoonCount.Caption = CStr(lvwBoons.ListItems.Count) & " Boons"
            If Index > lvwBoons.ListItems.Count Then Index = 1
            If lvwBoons.ListItems.Count > 0 Then Set lvwBoons.SelectedItem = lvwBoons.ListItems(Index)
            lvwBoons_ItemClick lvwBoons.SelectedItem
            
        End If
    
    End If
    
End Sub

Private Sub cmdDescend_Click()
'
' Name:         cmdDescend_Click
' Description:  Move the selected entry down.
'

    If cmdDescend.Visible And Not Vampire Is Nothing Then
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

    If cmdDecrement.Visible And Not Vampire Is Nothing Then
        CharSheetEngine.DecrementEntry
        CharSheetEngine.TargetList.SetFocus
        SetDataChanged
    End If
    
End Sub

Private Sub cmdIncrement_Click()
'
' Name:         cmdIncrement_Click
' Description:  Increment the selected entry.
'

    If cmdIncrement.Visible And Not Vampire Is Nothing Then
        CharSheetEngine.IncrementEntry
        CharSheetEngine.TargetList.SetFocus
        SetDataChanged
    End If
    
End Sub

Private Sub cmdCustom_Click()
'
' Name:         cmdCustom_Click
' Description:  Have the CharSheetEngine add a custom entry to the target.

    If Not Vampire Is Nothing Then
        CharSheetEngine.AddCustom
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
    
    If Not Vampire Is Nothing Then
        CharSheetEngine.AddNote
        SetDataChanged
    End If
    
End Sub

Private Sub cmdRemove_Click()
'
' Name:         cmdRemove_Click
' Description:  Have the CharSheetEngine remove a label or list entry.
'
    
    If Not Vampire Is Nothing Then
        CharSheetEngine.RemoveEntry
        SetDataChanged
    End If
    
End Sub

Private Sub cmdShow_Click()
'
' Name:         cmdShow_Click
' Description:  Show the character sheet being edited.
'

    mdiMain.ShowCharacterSheet lblName.Caption

End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  If the characters have changed, refresh the list.
'

    Dim QChange As Boolean
    Dim CChange As Boolean
    
    QChange = mdiMain.CheckForChanges(Me, atQueries)
    CChange = mdiMain.CheckForChanges(Me, atCharacters) Or mdiMain.CheckForChanges(Me, atStatus)
    
    If QChange Then
        RefreshSearches cboSearch.Text
    Else
        If CChange Then
            RefreshVampires
        Else
            lvwVampires_ItemClick lvwVampires.SelectedItem
        End If
    End If

    If CChange Then RefreshBoons

End Sub

Private Sub Form_Deactivate()
'
' Name:         Form_Deactivate
' Description:  Save the text.
'

    ValidateControls

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Initializes the character sheet engine and fills the needed lists.
'

    Set CharSheetEngine = New CharacterSheetEngineClass
    Set BoonIDSet = New StringSet
    
    CharSheetEngine.RegisterSheet "Vampire", lstMenu, lblMenuItem, lblMenuTitle
    
    RefreshSearches "All Characters"   'Will trigger RefreshVampires and lstVampires_ItemClick
    RefreshBoons
    
    mdiMain.OrientForm Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Unload
' Description:  Save the text and destroy the character sheet engine.
'

    ValidateControls
    Set CharSheetEngine = Nothing
    Set BoonIDSet = Nothing
    
End Sub

Private Sub lblBoonCount_DblClick()
'
' Name:         lblBoonCount_DblClick
' Description:  Recount the boons for testing purposes.
'

    RefreshBoons

End Sub

Private Sub lblField_Change(Index As Integer)
'
' Name:         lblField_Change
' Description:  Store the new value in the appropriate property of the character.
'

    Dim Value As String
    Dim BItem As ListItem
    
    If Not Populating Then
        
        Value = lblField(Index).Caption
        Set BItem = lvwBoons.SelectedItem
        
        Select Case Index
            
            Case fiTitle
                If Not Vampire Is Nothing Then Vampire.Title = Value
                
            Case fiType
                If Not DueBoon Is Nothing Then
                    DueBoon.BoonType = Value
                    BItem.Tag = DueBoon.BoonID(DueFrom.Name)
                End If
                If Not OwedBoon Is Nothing Then
                    OwedBoon.BoonType = Value
                    BItem.Tag = OwedBoon.BoonID(OwedTo.Name)
                End If
                BItem.Text = Value
            
            Case fiDueFrom
                BItem.ListSubItems(keyDUE).Text = Value
                If Not OwedBoon Is Nothing Then OwedBoon.CharName = Value
                If Not DueFrom Is Nothing Then
                    With DueFrom.BoonList
                        .First
                        Do Until .Off
                            If BItem.Tag = .Item.BoonID(DueFrom.Name) Then Exit Do
                            .MoveNext
                        Loop
                        If Not .Off Then .Remove
                    End With
                    DueFrom.LastModified = Now
                    DueFrom.RefreshBoonTraits
                End If
                CharacterList.MoveTo Value
                If Not CharacterList.Off Then
                    Set DueFrom = CharacterList.Item
                    If DueBoon Is Nothing Then
                        Set DueBoon = New BoonClass
                        DueBoon.BoonType = BItem.Text
                        DueBoon.BoonDate = CDate(BItem.ListSubItems(keyDATE))
                        DueBoon.CharName = BItem.ListSubItems(keyOWE)
                        DueBoon.IsOwed = True
                        DueBoon.Description = TrimWhiteSpace(txtMemo(miDescription).Text)
                    End If
                    DueFrom.BoonList.InsertSorted DueBoon
                    BItem.Tag = DueBoon.BoonID(DueFrom.Name)
                Else
                    Set DueBoon = Nothing
                    Set DueFrom = Nothing
                    BItem.Tag = ""
                    If Not (OwedTo Is Nothing Or OwedBoon Is Nothing) Then
                        BItem.Tag = OwedBoon.BoonID(OwedTo.Name)
                    End If
                End If
                
            Case fiOwedTo
                BItem.ListSubItems(keyOWE).Text = Value
                If Not DueBoon Is Nothing Then DueBoon.CharName = Value
                If Not OwedTo Is Nothing Then
                    With OwedTo.BoonList
                        .First
                        Do Until .Off
                            If BItem.Tag = .Item.BoonID(OwedTo.Name) Then Exit Do
                            .MoveNext
                        Loop
                        If Not .Off Then .Remove
                    End With
                    OwedTo.LastModified = Now
                    OwedTo.RefreshBoonTraits
                End If
                CharacterList.MoveTo Value
                If Not CharacterList.Off Then
                    Set OwedTo = CharacterList.Item
                    If OwedBoon Is Nothing Then
                        Set OwedBoon = New BoonClass
                        OwedBoon.BoonType = BItem.Text
                        OwedBoon.BoonDate = CDate(BItem.ListSubItems(keyDATE))
                        OwedBoon.CharName = BItem.ListSubItems(keyDUE)
                        OwedBoon.IsOwed = False
                        OwedBoon.Description = TrimWhiteSpace(txtMemo(miDescription).Text)
                    End If
                    OwedTo.BoonList.InsertSorted OwedBoon
                    BItem.Tag = OwedBoon.BoonID(OwedTo.Name)
                Else
                    Set OwedBoon = Nothing
                    Set OwedTo = Nothing
                    BItem.Tag = ""
                    If Not (DueFrom Is Nothing Or DueBoon Is Nothing) Then
                        BItem.Tag = DueBoon.BoonID(DueFrom.Name)
                    End If
                End If
                
            Case fiDate
                
                If IsDate(Value) Then
                    Dim DateVal As Date
                    DateVal = CDate(Value)
                    If Not DueBoon Is Nothing Then
                        DueBoon.BoonDate = DateVal
                        lvwBoons.SelectedItem.Tag = DueBoon.BoonID(DueFrom.Name)
                    End If
                    If Not OwedBoon Is Nothing Then
                        OwedBoon.BoonDate = DateVal
                        lvwBoons.SelectedItem.Tag = OwedBoon.BoonID(OwedTo.Name)
                    End If
                    lvwBoons.SelectedItem.ListSubItems(keyDATE).Text = Format(DateVal, "Short Date")
                    lvwBoons.SelectedItem.ListSubItems(keySORT).Text = Format(DateVal, "yyyy-mm-dd")
                End If
                
        End Select
        SetDataChanged
    
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

Private Sub lvwBoons_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'
' Name:         lvwBoons_ColumnClick
' Description:  Change the key by which the entries are sorted, or the sort order on a second click.
'
    
    Dim I As Integer
    
    I = ColumnHeader.Index - 1
    If I = 4 Then I = 5
    
    If lvwBoons.SortKey = I Then
        lvwBoons.SortOrder = IIf(lvwBoons.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwBoons.SortKey = I
    End If

End Sub

Private Sub lvwBoons_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
' Name:         lvwBoons_ItemClick
' Description:  Update the current boon info.
'

    Dim ID As String
    Dim FromName As String
    Dim ToName As String
    Dim ShowBoon As BoonClass
    Dim BoonDate As Date
    Dim BoonType As String
    
    Set OwedTo = Nothing
    Set DueFrom = Nothing
    Set OwedBoon = Nothing
    Set DueBoon = Nothing
    
    If Not lvwBoons.SelectedItem Is Nothing Then
        
        ID = lvwBoons.SelectedItem.Tag
        FromName = lvwBoons.SelectedItem.ListSubItems(keyDUE).Text
        ToName = lvwBoons.SelectedItem.ListSubItems(keyOWE).Text
        BoonType = lvwBoons.SelectedItem.Text
        BoonDate = CDate(lvwBoons.SelectedItem.ListSubItems(keyDATE).Text)
        
        With CharacterList
            
            If FromName <> "" Then
                .MoveTo FromName
                If Not .Off Then
                    If .Item.RaceCode = gvRaceVampire Then
                        Set DueFrom = .Item
                        With DueFrom.BoonList
                            .First
                            Do Until .Off
                                If .Item.BoonID(FromName) = ID Then Exit Do
                                .MoveNext
                            Loop
                            If Not .Off Then Set DueBoon = .Item
                        End With
                    End If
                End If
            End If
            
            If ToName <> "" Then
                .MoveTo ToName
                If Not .Off Then
                    If .Item.RaceCode = gvRaceVampire Then
                        Set OwedTo = .Item
                        With OwedTo.BoonList
                            .First
                            Do Until .Off
                                If .Item.BoonID(ToName) = ID Then Exit Do
                                .MoveNext
                            Loop
                            If Not .Off Then Set OwedBoon = .Item
                        End With
                    End If
                End If
            End If
            
        End With
        
        Populating = True
        lblField(fiType).Caption = BoonType
        lblField(fiDueFrom).Caption = FromName
        lblField(fiOwedTo).Caption = ToName
        lblField(fiDate).Caption = Format(BoonDate, "mmmm d, yyyy")
        If Not DueBoon Is Nothing Then
            txtMemo(miDescription).Text = DueBoon.Description
        ElseIf Not OwedBoon Is Nothing Then
            txtMemo(miDescription).Text = OwedBoon.Description
        Else
            txtMemo(miDescription).Text = ""
        End If
        Populating = False
        
        fraBoons.Visible = True
    Else
        fraBoons.Visible = False
    End If

End Sub

Private Sub lvwVampires_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'
' Name:         lvwVampires_ColumnClick
' Description:  Change the key by which the entries are sorted, or the sort order on a second click.
'
    
    If lvwVampires.SortKey = ColumnHeader.Index - 1 Then
        lvwVampires.SortOrder = IIf(lvwVampires.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwVampires.SortKey = ColumnHeader.Index - 1
    End If

End Sub

Private Sub lvwVampires_DblClick()
'
' Name:         lstCharacters_DblClick
' Description:  See cmdShow_Click.
'

    Call cmdShow_Click

End Sub

Private Sub lvwVampires_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
' Name:         lstCharacters_ItemClick
' Description:  Find the character and display the appropriate information at right.
'

    Set Vampire = Nothing
    If Not (Item Is Nothing) Then
        CharacterList.MoveTo Item.Text
        If Not CharacterList.Off Then
            Set Vampire = CharacterList.Item
            Populating = True
            lblName.Caption = Vampire.Name
            lblField(fiTitle).Caption = Vampire.Title
            CharSheetEngine.RegisterTraitList tiStatus, Vampire.StatusList
            CharSheetEngine.RefreshTraitList lstTraits(tiStatus), Vampire.StatusList
            Populating = False
        End If
    End If
    fraStatus.Visible = Not (Vampire Is Nothing)

End Sub

Private Sub optNot_Click(Index As Integer)
'
' Name:         optNot_Click
' Description:  The user has inverted the query, so populate the list.
'
    RefreshVampires
    
End Sub

Private Sub tabTabStrip_Click()
'
' Name:         tabTabStrip_Click
' Description:  Clear the menu and targets.  Display correct frame.
'

    If Not fraFrame(tabTabStrip.SelectedItem.Index - 1).Visible Then
        
        CharSheetEngine.DeselectMenus
        
        If CharSheetEngine.TargetType = ttListBox Then
            CharSheetEngine.TargetList.ListIndex = -1
            cmdIncrement.Visible = False
            cmdDecrement.Visible = False
            cmdAscend.Visible = False
            cmdDescend.Visible = False
        End If
        
        CharSheetEngine.TargetType = ttNothing
        
        Dim fTab As Frame
        For Each fTab In fraFrame
            fTab.Visible = (fTab.Index = tabTabStrip.SelectedItem.Index - 1)
        Next fTab
        
    End If

End Sub

Private Sub SetDataChanged()
'
' Name:         SetDataChanged
' Description:  Tell the game its data has changed and update the character's
'               Last Modified date.
'
        
    If Not Populating Then
        Game.DataChanged = True
        If fraFrame(FRAME_STATUS).Visible Then
            Call mdiMain.AnnounceChanges(Me, atStatus)
            If Not Vampire Is Nothing Then
                RefreshCurrentVampire
                Vampire.LastModified = Now
            End If
        Else
            If Not DueFrom Is Nothing Then
                DueFrom.LastModified = Now
                DueFrom.RefreshBoonTraits
            End If
            If Not OwedTo Is Nothing Then
                OwedTo.LastModified = Now
                OwedTo.RefreshBoonTraits
            End If
        End If
    End If
    
End Sub

Private Sub txtMemo_Validate(Index As Integer, Cancel As Boolean)
'
' Name:         txtMemo_Change
' Description:  Record changes to the memo field.
'

    Dim DidChange As Boolean
    Dim Value As String

    Select Case Index
        Case miDescription
            Value = TrimWhiteSpace(txtMemo(Index).Text)
            txtMemo(Index).Text = Value
            If Not DueBoon Is Nothing Then
                DidChange = (DueBoon.Description <> Value)
                DueBoon.Description = Value
            End If
            If Not OwedBoon Is Nothing Then
                DidChange = DidChange Or (OwedBoon.Description <> Value)
                OwedBoon.Description = Value
            End If
            If DidChange Then SetDataChanged
    
    End Select

End Sub
