VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlayerList 
   Caption         =   "Players"
   ClientHeight    =   5070
   ClientLeft      =   1875
   ClientTop       =   750
   ClientWidth     =   7605
   Icon            =   "frmPlayerList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5070
   ScaleWidth      =   7605
   Begin VB.Frame fraBottom 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   4440
      Width           =   5175
      Begin VB.ComboBox cboSearch 
         Height          =   315
         Left            =   2445
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   75
         Width           =   2055
      End
      Begin VB.OptionButton optNot 
         Caption         =   "List players that d&on't match:"
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   2430
      End
      Begin VB.OptionButton optNot 
         Caption         =   "List players that &match:"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Value           =   -1  'True
         Width           =   2430
      End
      Begin VB.Frame fraSortOrder 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   4515
         TabIndex        =   10
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
            TabIndex        =   11
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
            TabIndex        =   12
            Top             =   0
            Width           =   315
         End
      End
   End
   Begin VB.Frame fraRight 
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   5640
      TabIndex        =   2
      Top             =   480
      Width           =   1695
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Player"
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Show Player"
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "&Add New Player"
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   3105
         Width           =   1695
      End
      Begin VB.Label lblForeName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Valentino Monterrey"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1455
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   607
         Picture         =   "frmPlayerList.frx":058A
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
         TabIndex        =   18
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Player Points:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1455
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
         TabIndex        =   16
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Last Modified:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Index           =   2
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   1695
      End
   End
   Begin MSComctlLib.ListView lvwPlayers 
      Height          =   3855
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6800
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Group"
         Text            =   "Position"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Subgroup"
         Text            =   "ID"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   "NPC"
         Text            =   "Status"
         Object.Width           =   1676
      EndProperty
   End
   Begin VB.CommandButton cmdESCClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      Caption         =   "0 &Players"
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
      Width           =   5175
   End
End
Attribute VB_Name = "frmPlayerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Name:         frmPlayerList
' Description:  Form that lists players; allows you to view, add, and delete players.
'
Private ShiftDown As Boolean

Private Const FORM_START_HEIGHT = 5475
Private Const FORM_START_WIDTH = 7740
Private Const FORM_MIN_SCALEHEIGHT = 5070
Private Const FORM_MIN_SCALEWIDTH = 4005
Private Const RIGHT_MARGIN = 1980
Private Const VERTICAL_GAP = 225
Private Const LIST_SCROLL_WIDTH = 300
Private Const BOTTOM_MARGIN = 645
Private Const HORIZONTAL_GAP = 105

Private Const OPT_MATCH = 0
Private Const OPT_NO_MATCH = 1
Private Const OPT_ASCEND = 0
Private Const OPT_DESCEND = 1

Private Sub RefreshList()
'
' Name:         RefreshList
' Description:  Preserving the current selection, this refills the list box from the list of
'               Players according to the chosen search.
'

    Dim StoreSelKey As String
    Dim Search As QueryClass
    Dim Player As PlayerClass
    Dim NewItem As ListItem
    
    Screen.MousePointer = vbHourglass
    
    If Not (lvwPlayers.SelectedItem Is Nothing) Then _
            StoreSelKey = lvwPlayers.SelectedItem.Key
    
    lvwPlayers.ListItems.Clear
    
    With Game.QueryEngine.QueryList
        .MoveTo cboSearch.Text
        If Not .Off Then
            Set Search = .Item
        Else
            Set Search = New QueryClass
            Search.Inventory = qiPlayers
        End If
    End With
    
    With Game.QueryEngine
        .MakeQuery Search, , optNot(OPT_NO_MATCH).Value
    
        .Results.First
        Do Until .Results.Off
            Set Player = .Results.Item
            Set NewItem = lvwPlayers.ListItems.Add(Key:="key" & Player.Name, Text:=Player.Name)
            NewItem.ListSubItems.Add Text:=Player.Position
            NewItem.ListSubItems.Add Text:=Player.ID
            NewItem.ListSubItems.Add Text:=Player.Status
            .Results.MoveNext
        Loop
    End With
    
    lblCount.Caption = CStr(lvwPlayers.ListItems.Count) & " &Players" & _
            IIf(cboSearch.Text = "All Players" Or cboSearch.Text = "", "", _
                " (" & IIf(optNot(OPT_NO_MATCH).Value, "Not ", "") & cboSearch.Text & ")")
    
    On Error Resume Next
    Set lvwPlayers.SelectedItem = lvwPlayers.ListItems(StoreSelKey)
    If lvwPlayers.SelectedItem Is Nothing And lvwPlayers.ListItems.Count > 0 Then _
        Set lvwPlayers.SelectedItem = lvwPlayers.GetFirstVisible
    lvwPlayers.SelectedItem.EnsureVisible
    On Error GoTo 0

    lvwPlayers_ItemClick lvwPlayers.SelectedItem

    Screen.MousePointer = vbDefault

    Set Search = Nothing

End Sub

Public Sub SetDefaultOutput()
'
' Name:         SetDefaultOutput
' Description:  Initilize the OutputEngineClass with default output settings.
'
    With OutputEngine
        .Template = tnPlayerRoster
        .SelectSet(osPlayers).Clear
        .SelectSet(osPlayers).StoreListView lvwPlayers, True
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
    RefreshList
    
End Sub

Private Sub cmdAddNew_Click()
'
' Name:         cmdAddNew_Click
' Description:  Add a new player to the game.
'

    Dim NewName As String
    Dim Player As PlayerClass
    Dim HasInList As Boolean
    
    NewName = InputBox("Enter a name for the new player:", "Add New Player")
    NewName = Trim(NewName)
    
    If NewName <> "" Then
    
        PlayerList.MoveTo NewName
        If PlayerList.Off Then
    
            Set Player = New PlayerClass
            Player.Name = NewName
            PlayerList.InsertSorted Player
            RefreshList
            
            On Error Resume Next
            Set lvwPlayers.SelectedItem = lvwPlayers.ListItems("key" & NewName)
            lvwPlayers.SelectedItem.EnsureVisible
            HasInList = (NewName = lvwPlayers.SelectedItem.Text)
            On Error GoTo 0
        
            If HasInList Then
                lvwPlayers.SetFocus
            Else
                mdiMain.ShowPlayer NewName
            End If

            mdiMain.AnnounceChanges Me, atPlayers
            Game.DataChanged = True
        Else
            MsgBox "The name """ & NewName & """ is already in use.  Please " & _
                    "enter a different name.", vbExclamation + vbOKOnly, "Duplicate Name"
        End If
        
    End If

End Sub

Private Sub cmdDelete_Click()
'
' Name:         cmdDelete_Click
' Description:  Finds the Player and asks confirmation of deletion.  If yes, remove the Player
'               and refill the list.
'

    Dim NormForm As Form
    Dim DelName As String
    Dim Answer As Boolean
    
    If Not (lvwPlayers.SelectedItem Is Nothing) Then
        DelName = lvwPlayers.SelectedItem.Text
        PlayerList.MoveTo DelName
        If Not PlayerList.Off Then
            
            Answer = ShiftDown
            If Not Answer Then Answer = (MsgBox("This will PERMANENTLY remove " & DelName & _
                    " from the game. Are you sure you want to delete this player?", _
                    vbQuestion + vbYesNo, "Delete Player") = vbYes)
            If Answer Then
                    
                mdiMain.AnnounceChanges Me, atPlayers
                Game.DataChanged = True
    
                For Each NormForm In Forms()
                    If NormForm.Caption = DelName Then
                        Unload NormForm
                        Exit For
                    End If
                Next NormForm
                
                PlayerList.Remove
                RefreshList
                
            End If
        Else
            MsgBox "Grapevine can't find this player!  Was it renamed or deleted?", vbExclamation
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
' Description:  Asks the parent form to create a Player sheet screen for the selected Player.
'

    If Not (lvwPlayers.SelectedItem Is Nothing) Then _
        mdiMain.ShowPlayer lvwPlayers.SelectedItem.Text

End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  If the Players have changed, refresh the list.
'

    Dim QChange As Boolean
    Dim PChange As Boolean
    
    QChange = mdiMain.CheckForChanges(Me, atQueries)
    PChange = mdiMain.CheckForChanges(Me, atPlayers)
    
    If QChange Then
        PopulateSearches cboSearch.Text
    Else
        If PChange Then
            RefreshList
        Else
            lvwPlayers_ItemClick lvwPlayers.SelectedItem
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
' Description:  Fill the list and select the first Player.
'

    Dim LastSearchName As String

    Me.Height = FORM_START_HEIGHT
    Me.Width = FORM_START_WIDTH
        
    LastSearchName = GetSetting(App.Title, "Settings", "PlayerList", "All Players")
    lvwPlayers.SortKey = GetSetting(App.Title, "Settings", "PlayerSort", 0)
    lvwPlayers.SortOrder = GetSetting(App.Title, "Settings", "PlayerSortOrder", lvwAscending)

    If lvwPlayers.SortOrder = lvwAscending Then
        optSortOrder(OPT_ASCEND).Value = True
    Else
        optSortOrder(OPT_DESCEND).Value = True
    End If
    
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
        lvwPlayers.Height = SH - lvwPlayers.Top - BOTTOM_MARGIN - HORIZONTAL_GAP
        lvwPlayers.Width = SW - lvwPlayers.Left - RIGHT_MARGIN - VERTICAL_GAP
        lblCount.Width = lvwPlayers.Width

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Unload
' Description:  Destroy the objects created by this form.
'

    SaveSetting App.Title, "Settings", "PlayerList", cboSearch.Text
    SaveSetting App.Title, "Settings", "PlayerSort", lvwPlayers.SortKey
    SaveSetting App.Title, "Settings", "PlayerSortOrder", lvwPlayers.SortOrder

End Sub

Private Sub lvwPlayers_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'
' Name:         lvwPlayers_ColumnClick
' Description:  Change the key by which the entries are sorted.
'
    
    lvwPlayers.SortKey = ColumnHeader.Index - 1

End Sub

Private Sub lvwPlayers_DblClick()
'
' Name:         lstPlayers_DblClick
' Description:  See cmdShow_Click.
'

    Call cmdShow_Click

End Sub

Private Sub lvwPlayers_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
' Name:         lstPlayers_ItemClick
' Description:  Find the Player and display the appropriate information at right.
'

    If Not (Item Is Nothing) Then
        PlayerList.MoveTo Item.Text
        If Not PlayerList.Off Then
            imgIcon.Visible = True
            lblForeName.Caption = Item.Text
            lblExperience.Caption = CStr(PlayerList.Item.Experience.Unspent) & " / " & _
                                    CStr(PlayerList.Item.Experience.Earned)
            lblDate.Caption = Format(PlayerList.Item.LastModified, "Short Date")
        Else
            MsgBox "Grapevine can't find this player!  Was it renamed or deleted?", vbExclamation
        End If
    Else
        imgIcon.Visible = False
        lblForeName.Caption = ""
        lblExperience.Caption = ""
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
' Description:  Change the sorting order of the Player list, as needed.
'
    lvwPlayers.SortOrder = IIf(Index = OPT_ASCEND, lvwAscending, lvwDescending)

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
            If .Item.Inventory = qiPlayers Then
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

    If Not cboSearch.ListIndex >= 0 Then RefreshList    'Force the list populous with all Players

End Sub
