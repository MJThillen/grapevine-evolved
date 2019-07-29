VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCharacters 
   Caption         =   "Characters"
   ClientHeight    =   6165
   ClientLeft      =   1875
   ClientTop       =   750
   ClientWidth     =   9030
   Icon            =   "frmCharacters.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   9030
   Begin VB.Frame fraBottom 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   5400
      Width           =   6375
      Begin VB.ComboBox cboSearch 
         Height          =   315
         Left            =   3525
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   75
         Width           =   2055
      End
      Begin VB.OptionButton optNot 
         Caption         =   "List characters that d&on't match the search:"
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   3390
      End
      Begin VB.OptionButton optNot 
         Caption         =   "List characters that &match the search:"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Value           =   -1  'True
         Width           =   3390
      End
      Begin VB.Frame fraSortOrder 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   5595
         TabIndex        =   11
         Top             =   75
         Width           =   735
         Begin VB.OptionButton optSortOrder 
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
            Height          =   315
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   0
            Value           =   -1  'True
            Width           =   315
         End
         Begin VB.OptionButton optSortOrder 
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
            Height          =   315
            Index           =   1
            Left            =   315
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   0
            Width           =   315
         End
      End
   End
   Begin VB.Frame fraRight 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   6720
      TabIndex        =   2
      Top             =   480
      Width           =   2055
      Begin VB.CommandButton cmdChangeRace 
         Caption         =   "Copy to New &Race"
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   4455
         Width           =   2055
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Character"
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   3960
         Width           =   2055
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Show Character"
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "&Add New Character"
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   3465
         Width           =   2055
      End
      Begin VB.Label lblForeName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   15
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label lblBackName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   420
         Index           =   3
         Left            =   135
         TabIndex        =   17
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label lblBackName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   420
         Index           =   2
         Left            =   105
         TabIndex        =   18
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label lblBackName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   420
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   555
         Width           =   1815
      End
      Begin VB.Label lblBackName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   420
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   525
         Width           =   1815
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   787
         Tag             =   "11"
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblPlayer 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Experience:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblExperience 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Last Modified:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Player:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   2535
         Index           =   2
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   2055
      End
   End
   Begin MSComctlLib.ListView lvwCharacters 
      Height          =   4815
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8493
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Text            =   "Race, Name"
         Object.Width           =   4075
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Group"
         Text            =   "Group"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Subgroup"
         Text            =   "Subgroup"
         Object.Width           =   1826
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   "NPC"
         Text            =   "NPC"
         Object.Width           =   900
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Status"
         Text            =   "Status"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Race"
         Text            =   "Race"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdESCClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      Caption         =   "0 &Characters"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "frmCharacters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' Name:         frmCharacters
' Description:  Form that lists characters; allows you to view, add, and delete characters.
'               Displayed only from the menu.
'
Private ShiftDown As Boolean

Private Const FORM_START_HEIGHT = 6570
Private Const FORM_START_WIDTH = 9150
Private Const FORM_MIN_SCALEHEIGHT = 6165
Private Const FORM_MIN_SCALEWIDTH = 6855
Private Const BOTTOM_MARGIN = 765
Private Const RIGHT_MARGIN = 2310
Private Const HORIZONTAL_GAP = 105
Private Const VERTICAL_GAP = 225

Private Const OPT_MATCH = 0
Private Const OPT_NO_MATCH = 1
Private Const OPT_ASCEND = 0
Private Const OPT_DESCEND = 1

Private Sub RefreshList()
'
' Name:         RefreshList
' Description:  Preserving the current selection, this refills the list box from the list of
'               characters according to the chosen search.
'

    Dim StoreSelKey As String
    Dim Search As QueryClass
    Dim Character As Object
    Dim NewItem As ListItem
    
    Screen.MousePointer = vbHourglass
    
    If Not (lvwCharacters.SelectedItem Is Nothing) Then _
            StoreSelKey = lvwCharacters.SelectedItem.Key
    
    lvwCharacters.ListItems.Clear
    
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
            Set Character = .Results.Item
            Set NewItem = lvwCharacters.ListItems.Add(, _
                          "key" & Character.Name, Character.Name, , Character.Race)
            NewItem.ListSubItems.Add , "Group", Character.Group
            NewItem.ListSubItems.Add , "Subgroup", Character.Subgroup
            NewItem.ListSubItems.Add , "NPC", IIf(Character.IsNPC, "X", "")
            NewItem.ListSubItems.Add , "Status", Character.Status
            NewItem.ListSubItems.Add , "Race", Character.Race
            .Results.MoveNext
        Loop
        
    End With
    
    lblCount.Caption = CStr(lvwCharacters.ListItems.Count) & " &Characters" & _
            IIf(cboSearch.Text = "All Characters" Or cboSearch.Text = "", "", _
                " (" & IIf(optNot(OPT_NO_MATCH).Value, "Not ", "") & cboSearch.Text & ")")
    
    On Error Resume Next
    Set lvwCharacters.SelectedItem = lvwCharacters.ListItems(StoreSelKey)
    If lvwCharacters.SelectedItem Is Nothing And lvwCharacters.ListItems.Count > 0 Then _
        Set lvwCharacters.SelectedItem = lvwCharacters.GetFirstVisible
    lvwCharacters.SelectedItem.EnsureVisible
    On Error GoTo 0

    lvwCharacters_ItemClick lvwCharacters.SelectedItem

    Screen.MousePointer = vbDefault

    Set Search = Nothing

End Sub

Public Sub SetDefaultOutput()
'
' Name:         SetDefaultOutput
' Description:  Initilize the OutputEngineClass with default output settings.
'
    With OutputEngine
        .Template = tnCharacterSheets
        .GameDate = 0
        .SearchName = cboSearch.Text
        .SearchNot = optNot(OPT_NO_MATCH).Value
    End With
    
End Sub

Private Sub cboSearch_Click()
'
' Name:         cboSearch_Click
' Description:  The user has chosen a new query, so populate the list.
'
    RefreshList
    
End Sub

Private Sub cmdAddNew_Click()
'
' Name:         cmdAddNew_Click
' Description:  Calls on frmAddNewCharacter to display itself and return a name and race in
'               its CharacterName and Race attributes.  Creates the appropriate character and
'               adds it to the linked list and the list box, selecting it.
'

    Dim NewChar As Object
    Dim HasInList As Boolean
    
    frmAddNewCharacter.GetCharacter
    Select Case frmAddNewCharacter.Race
        Case gvRaceVampire
            Set NewChar = New VampireClass
        Case gvRaceWerewolf
            Set NewChar = New WerewolfClass
        Case gvRaceMortal
            Set NewChar = New MortalClass
        Case gvRaceChangeling
            Set NewChar = New ChangelingClass
        Case gvRaceWraith
            Set NewChar = New WraithClass
        Case gvracemage
            Set NewChar = New MageClass
        Case gvRaceFera
            Set NewChar = New FeraClass
        Case gvRaceVarious
            Set NewChar = New VariousClass
        Case gvRaceMummy
            Set NewChar = New MummyClass
        Case gvRaceKueiJin
            Set NewChar = New KueiJinClass
        Case gvRaceHunter
            Set NewChar = New HunterClass
        Case gvRaceDemon
            Set NewChar = New DemonClass
    End Select

    If frmAddNewCharacter.Race <> gvRaceNone Then
        
        NewChar.Name = frmAddNewCharacter.CharacterName
        If frmAddNewCharacter.RandomGen Then Game.MenuSet.GenerateRandomTraits NewChar
        CharacterList.InsertSorted NewChar
        
        mdiMain.AnnounceChanges Me, atCharacters
        Game.DataChanged = True
        RefreshList
        On Error Resume Next
        Set lvwCharacters.SelectedItem = _
                lvwCharacters.ListItems("key" & NewChar.Name)
        lvwCharacters.SelectedItem.EnsureVisible
        HasInList = (NewChar.Name = lvwCharacters.SelectedItem.Text)
        On Error GoTo 0
    
        If HasInList Then
            lvwCharacters.SetFocus
        Else
            mdiMain.ShowCharacterSheet NewChar.Name
        End If
    
    End If

End Sub

Private Sub cmdChangeRace_Click()
'
' Name:         cmdChangeRace_Click
' Description:  Change the race of a given character, preserving as much data as possible
'

    Dim NormForm As Form
    Dim NewChar As Object
    Dim OldChar As Object
    
    If Not (lvwCharacters.SelectedItem Is Nothing) Then
        CharacterList.MoveTo lvwCharacters.SelectedItem.Text
        If Not CharacterList.Off Then
            
            Set OldChar = CharacterList.Item
            
            frmAddNewCharacter.GetCharacter OldChar.Name & " II"
            
            Select Case frmAddNewCharacter.Race
                Case gvRaceVampire
                    Set NewChar = New VampireClass
                Case gvRaceWerewolf
                    Set NewChar = New WerewolfClass
                Case gvRaceMortal
                    Set NewChar = New MortalClass
                Case gvRaceChangeling
                    Set NewChar = New ChangelingClass
                Case gvRaceWraith
                    Set NewChar = New WraithClass
                Case gvracemage
                    Set NewChar = New MageClass
                Case gvRaceFera
                    Set NewChar = New FeraClass
                Case gvRaceVarious
                    Set NewChar = New VariousClass
                Case gvRaceMummy
                    Set NewChar = New MummyClass
                Case gvRaceKueiJin
                    Set NewChar = New KueiJinClass
                Case gvRaceHunter
                    Set NewChar = New HunterClass
                Case gvRaceDemon
                    Set NewChar = New DemonClass
            End Select
            
            If frmAddNewCharacter.Race <> gvRaceNone Then
                
                NewChar.Name = frmAddNewCharacter.CharacterName
                
                On Error Resume Next
                NewChar.Nature = OldChar.Nature
                NewChar.Demeanor = OldChar.Demeanor
                NewChar.Willpower = OldChar.Willpower
                NewChar.TempWillpower = OldChar.TempWillpower
                NewChar.PhysicalMax = OldChar.PhysicalMax
                NewChar.SocialMax = OldChar.SocialMax
                NewChar.MentalMax = OldChar.MentalMax
                NewChar.Player = OldChar.Player
                NewChar.ID = OldChar.ID
                NewChar.Status = OldChar.Status
                NewChar.IsNPC = OldChar.IsNPC
                NewChar.Narrator = OldChar.Narrator
                NewChar.StartDate = OldChar.StartDate
                NewChar.Notes = OldChar.Notes
                NewChar.Experience.Copy OldChar.Experience
                NewChar.PhysicalList.Copy OldChar.PhysicalList
                NewChar.SocialList.Copy OldChar.SocialList
                NewChar.MentalList.Copy OldChar.MentalList
                NewChar.PhysicalNegList.Copy OldChar.PhysicalNegList
                NewChar.SocialNegList.Copy OldChar.SocialNegList
                NewChar.MentalNegList.Copy OldChar.MentalNegList
                NewChar.AbilityList.Copy OldChar.AbilityList
                NewChar.BackgroundList.Copy OldChar.BackgroundList
                NewChar.InfluenceList.Copy OldChar.InfluenceList
                NewChar.MeritList.Copy OldChar.MeritList
                NewChar.FlawList.Copy OldChar.FlawList
                NewChar.EquipmentList.Copy OldChar.EquipmentList
                NewChar.HangoutList.Copy OldChar.HangoutList
                On Error GoTo 0
                
                CharacterList.InsertSorted NewChar
                
                mdiMain.AnnounceChanges Me, atCharacters
                Game.DataChanged = True
                RefreshList
                On Error Resume Next
                Set lvwCharacters.SelectedItem = lvwCharacters.ListItems("key" & NewChar.Name)
                lvwCharacters.SelectedItem.EnsureVisible
                On Error GoTo 0
                lvwCharacters.SetFocus
                
            End If
                
        Else
            MsgBox "Grapevine can't find this character!  Was it renamed or deleted?", vbExclamation
        End If
    End If

End Sub

Private Sub cmdDelete_Click()
'
' Name:         cmdDelete_Click
' Description:  Finds the character and asks confirmation of deletion.  If yes, remove the character
'               and refill the list.
'

    Dim NormForm As Form
    Dim DelName As String
    Dim Answer As Boolean
    
    If Not (lvwCharacters.SelectedItem Is Nothing) Then
        DelName = lvwCharacters.SelectedItem.Text
        CharacterList.MoveTo DelName
        If Not CharacterList.Off Then
            
            Answer = ShiftDown
            If Not Answer Then Answer = (MsgBox("This will PERMANENTLY remove " & DelName & _
                    " from the game. Are you sure you want to delete this character?", _
                    vbQuestion + vbYesNo, "Delete Character") = vbYes)
            If Answer Then
                    
                mdiMain.AnnounceChanges Me, atCharacters
                Game.DataChanged = True
    
                For Each NormForm In Forms()
                    If (NormForm.Caption = DelName And NormForm.Tag = "C") Or NormForm Is frmCalculator Then
                        Unload NormForm
                    End If
                Next NormForm
                
                CharacterList.Remove
                RefreshList
                'Call lstCharacters_Click
                
            End If
        Else
            MsgBox "Grapevine can't find this character!  Was it renamed or deleted?", vbExclamation
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

Private Sub cmdShow_Click()
'
' Name:         cmdShow_Click
' Description:  Asks the parent form to create a character sheet screen for the selected character.
'

    If Not (lvwCharacters.SelectedItem Is Nothing) Then _
        mdiMain.ShowCharacterSheet lvwCharacters.SelectedItem.Text

End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  If the characters have changed, refresh the list.
'

    Dim QChange As Boolean
    Dim CChange As Boolean
    
    QChange = mdiMain.CheckForChanges(Me, atQueries)
    CChange = mdiMain.CheckForChanges(Me, atCharacters)
    
    If QChange Then
        PopulateSearches cboSearch.Text
    Else
        If CChange Then
            RefreshList
        Else
            lvwCharacters_ItemClick lvwCharacters.SelectedItem
        End If
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
' Name:         Form_KeyDown
' Description:  Record the state of the Shift key for deletions.
'

    If KeyCode = vbKeyShift Then ShiftDown = True

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'
' Name:         Form_KeyDown
' Description:  Record the state of the Shift key for deletions.
'

    If KeyCode = vbKeyShift Then ShiftDown = False
    If KeyCode = vbKeyDelete Then Call cmdDelete_Click

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Fill the list and select the first character.
'

    Dim LastSearchName As String

    Me.Height = GetSetting(App.Title, "Settings", "CharHeight", FORM_START_HEIGHT)
    Me.Width = GetSetting(App.Title, "Settings", "CharWidth", FORM_START_WIDTH)
    Me.Top = GetSetting(App.Title, "Settings", "CharTop", 0)
    Me.Left = GetSetting(App.Title, "Settings", "CharLeft", 0)
        
    LastSearchName = GetSetting(App.Title, "Settings", "CharList", "All Characters")
    lvwCharacters.SortKey = GetSetting(App.Title, "Settings", "CharSort", 0)
    lvwCharacters.SortOrder = GetSetting(App.Title, "Settings", "CharSortOrder", lvwAscending)

    If lvwCharacters.SortOrder = lvwAscending Then
        optSortOrder(OPT_ASCEND).Value = True
    Else
        optSortOrder(OPT_DESCEND).Value = True
    End If
    
    Set lvwCharacters.SmallIcons = mdiMain.imlSmallIcons

    PopulateSearches LastSearchName
    
End Sub

Private Sub Form_Resize()
'
' Name:         Form_Resize
' Description:  Position the controls appropriately on a resized form.
'

    Dim SH As Integer
    Dim SW As Integer
    
    If Me.WindowState <> vbMinimized Then
        
        SH = Me.ScaleHeight
        SW = Me.ScaleWidth
    
        If SH < FORM_MIN_SCALEHEIGHT Then SH = FORM_MIN_SCALEHEIGHT
        If SW < FORM_MIN_SCALEWIDTH Then SW = FORM_MIN_SCALEWIDTH
    
        fraBottom.Top = SH - BOTTOM_MARGIN
        fraRight.Left = SW - RIGHT_MARGIN
        lvwCharacters.Height = SH - lvwCharacters.Top - BOTTOM_MARGIN - HORIZONTAL_GAP
        lvwCharacters.Width = SW - lvwCharacters.Left - RIGHT_MARGIN - VERTICAL_GAP
        lblCount.Width = lvwCharacters.Width

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Unload
' Description:  Destroy the objects created by this form.
'

    If Not Me.WindowState = vbMinimized Then
        SaveSetting App.Title, "Settings", "CharTop", Me.Top
        SaveSetting App.Title, "Settings", "CharLeft", Me.Left
        SaveSetting App.Title, "Settings", "CharWidth", Me.Width
        SaveSetting App.Title, "Settings", "CharHeight", Me.Height
    End If
    SaveSetting App.Title, "Settings", "CharList", cboSearch.Text
    SaveSetting App.Title, "Settings", "CharSort", lvwCharacters.SortKey
    SaveSetting App.Title, "Settings", "CharSortOrder", lvwCharacters.SortOrder

End Sub

Private Sub lvwCharacters_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'
' Name:         lvwCharacters_ColumnClick
' Description:  Change the key by which the entries are sorted.
'
    
    If ColumnHeader.Key = "Name" Then
        If lvwCharacters.SortKey = 0 Then
            lvwCharacters.SortKey = 5
        Else
            lvwCharacters.SortKey = 0
        End If
    Else
        lvwCharacters.SortKey = ColumnHeader.Index - 1
    End If

End Sub

Private Sub lvwCharacters_DblClick()
'
' Name:         lstCharacters_DblClick
' Description:  See cmdShow_Click.
'

    Call cmdShow_Click

End Sub

Private Sub lvwCharacters_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
' Name:         lstCharacters_ItemClick
' Description:  Find the character and display the appropriate information at right.
'

    If Not (Item Is Nothing) Then
        CharacterList.MoveTo Item.Text
        If Not CharacterList.Off Then
            imgIcon.Picture = mdiMain.imlIcons.ListImages(CharacterList.Item.Race).Picture
            imgIcon.Visible = True
            lblForeName.Caption = Item.Text
            lblBackName(0).Caption = lblForeName.Caption
            lblBackName(1).Caption = lblForeName.Caption
            lblBackName(2).Caption = lblForeName.Caption
            lblBackName(3).Caption = lblForeName.Caption
            lblPlayer.Caption = CharacterList.Item.Player
            lblExperience.Caption = CStr(CharacterList.Item.Experience.Unspent) & " / " & _
                                    CStr(CharacterList.Item.Experience.Earned)
            lblDate.Caption = Format(CharacterList.Item.LastModified, "Short Date")
        Else
            MsgBox "Grapevine can't find this character!  Was it renamed or deleted?", vbExclamation
        End If
    Else
        imgIcon.Visible = False
        lblForeName.Caption = ""
        lblBackName(0).Caption = ""
        lblBackName(1).Caption = ""
        lblBackName(2).Caption = ""
        lblBackName(3).Caption = ""
        lblExperience.Caption = ""
        lblPlayer.Caption = ""
        lblDate.Caption = ""
    End If

End Sub

Private Sub optNot_Click(Index As Integer)
'
' Name:         optNot_Click
' Description:  The user has inverted the query, so populate the list.
'
    RefreshList
    
End Sub

Private Sub optSortOrder_Click(Index As Integer)
'
' Name:         optSortOrder_Click
' Description:  Change the sorting order of the character list, as needed.
'
    lvwCharacters.SortOrder = IIf(Index = OPT_ASCEND, lvwAscending, lvwDescending)

End Sub

Private Sub PopulateSearches(Default As String)
'
' Name:         PopulateSearches
' Parameters:   Default         Default search to use
' Description:  Fill cboSearch from the QueryList.  This will force either a
'               cboSearch_Click event (and thus a RefreshList) or a RefreshList
'               directly.
'

    Dim I As Integer
    
    cboSearch.Clear
    
    With Game.QueryEngine.QueryList
        .First
        Do Until .Off
            If .Item.Inventory = qiCharacters Then
                cboSearch.AddItem .Item.Name
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

    If Not cboSearch.ListIndex >= 0 Then RefreshList    'Force the list populous with all characters

End Sub
