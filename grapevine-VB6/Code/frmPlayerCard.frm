VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlayerCard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player"
   ClientHeight    =   5100
   ClientLeft      =   210
   ClientTop       =   405
   ClientWidth     =   8340
   Icon            =   "frmPlayerCard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   8340
   Tag             =   "Y"
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
      TabIndex        =   39
      Top             =   150
      Width           =   975
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   3615
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   1200
      Width           =   5910
      Begin VB.TextBox txtUserField 
         Height          =   375
         Index           =   3
         Left            =   3840
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtUserField 
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   11
         Top             =   1440
         Width           =   4695
      End
      Begin VB.TextBox txtUserField 
         Height          =   375
         Index           =   1
         Left            =   3840
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtMemo 
         Height          =   1185
         Index           =   1
         Left            =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox txtMemo 
         Height          =   1185
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   2295
         Width           =   2775
      End
      Begin VB.Label lblUserField 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   8
         Top             =   1005
         Width           =   855
      End
      Begin VB.Label lblUserField 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   10
         Top             =   1485
         Width           =   855
      End
      Begin VB.Label lblUserField 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Player ID"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   6
         Top             =   525
         Width           =   855
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   14
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   990
         Width           =   855
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   5
         Tag             =   "Status"
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   3
         Tag             =   "Position"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   510
         Width           =   855
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   3615
      Index           =   1
      Left            =   2160
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   5895
      Begin VB.ListBox lstTraits 
         Height          =   1185
         Index           =   0
         IntegralHeight  =   0   'False
         ItemData        =   "frmPlayerCard.frx":058A
         Left            =   3000
         List            =   "frmPlayerCard.frx":058C
         TabIndex        =   25
         Tag             =   "Spheres"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2775
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
         Left            =   3000
         TabIndex        =   20
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdIncrement 
         Caption         =   "+"
         Height          =   255
         Left            =   5520
         TabIndex        =   24
         Top             =   1080
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
         Left            =   3480
         TabIndex        =   21
         Top             =   1080
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
         Left            =   5160
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdShowCharacter 
         Caption         =   "&Show Character Sheet"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   3000
         Width           =   2775
      End
      Begin VB.ListBox lstCharacters 
         Height          =   2400
         ItemData        =   "frmPlayerCard.frx":058E
         Left            =   120
         List            =   "frmPlayerCard.frx":0595
         TabIndex        =   18
         Tag             =   "Tempers"
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label lblTraits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lblTraits"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   22
         Top             =   1080
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label lblCharacters 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Regulars"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Index           =   2
      Left            =   2160
      TabIndex        =   26
      Top             =   1200
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txtExperience 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   32
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtExperience 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   29
         Top             =   480
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvwHistory 
         Height          =   1710
         Left            =   105
         TabIndex        =   37
         Tag             =   "?PP"
         Top             =   1785
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   3016
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
            Object.Width           =   4022
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
         Left            =   2175
         TabIndex        =   33
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
         Left            =   2175
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   510
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
      Begin VB.Label lblModifiedLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Modified"
         Height          =   375
         Left            =   3120
         TabIndex        =   34
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblModified 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3960
         TabIndex        =   35
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblXPLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Player Point &History"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   36
         Top             =   1560
         Width           =   5685
      End
      Begin VB.Label lblXPLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Unspent"
         Height          =   375
         Index           =   0
         Left            =   -120
         TabIndex        =   28
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
         TabIndex        =   31
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label lblXPLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Player Points"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   27
         Top             =   240
         Width           =   1215
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
      TabIndex        =   40
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
      TabIndex        =   46
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdNote 
      Caption         =   "Add &Note to Entry"
      Height          =   375
      Left            =   120
      TabIndex        =   45
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<- Re&move"
      Height          =   375
      Left            =   120
      TabIndex        =   44
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add ->"
      Height          =   375
      Left            =   120
      TabIndex        =   43
      Top             =   3120
      Width           =   1695
   End
   Begin VB.ListBox lstMenu 
      Height          =   1785
      IntegralHeight  =   0   'False
      ItemData        =   "frmPlayerCard.frx":05A8
      Left            =   120
      List            =   "frmPlayerCard.frx":05AA
      TabIndex        =   42
      Top             =   1200
      Width           =   1695
   End
   Begin MSComctlLib.TabStrip tabTabStrip 
      Height          =   4095
      Left            =   2040
      TabIndex        =   0
      Top             =   840
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7223
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identity"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Characters"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Player Points"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMenuItem 
      Height          =   495
      Left            =   240
      TabIndex        =   48
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   480
      Picture         =   "frmPlayerCard.frx":05AC
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
      TabIndex        =   38
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblMenuTitle 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmPlayerCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Name:         frmPlayerCard
' Description:  The screen from which to manipulate Player data.
'
Option Explicit

'Constants by which specific labels are indexed. (fi = Field Index)
Private Const fiPosition = 0
Private Const fiStatus = 1

' Constants by which specific list boxes are indexed.
Private Const tiNone = 0

' Constants by which specific text boxes are indexed. (xi = Text Index)
Private Const xiName = 0
Private Const xiID = 1
Private Const xiEMail = 2
Private Const xiPhone = 3

' Constants by which specific memo fields are indexed. (mi = Memo Index)
Private Const miAddress = 0
Private Const miNotes = 1

'Constants to index important frames and the PP editing boxes
Private Const tbCharacters = 1
Private Const xpFrame = 2
Private Const xpUnspent = 0
Private Const xpEarned = 1

Private Player As PlayerClass                           'The Player in question
Private CharSheetEngine As CharacterSheetEngineClass    'Handles common functions
Private Populating As Boolean                           'defuses some events when characters are loaded

Public Sub ShowPlayer(Card As PlayerClass)
'
' Name:         ShowPlayer
' Parameter:    Card        the PlayerClass this form displays and modifies.
' Description:  Show and initialize a new instance of the form.
'

    Dim DataState As Boolean

    Populating = True

    Set Player = Nothing
    Set Player = Card
    DataState = Game.DataChanged

    Me.Caption = Player.Name

    txtUserField(xiName) = Player.Name
    txtUserField(xiID) = Player.ID
    txtUserField(xiEMail) = Player.EMail
    txtUserField(xiPhone) = Player.Phone

    lblField(fiPosition) = Player.Position
    lblField(fiStatus) = Player.Status
    
    txtMemo(miAddress) = Player.Address
    txtMemo(miNotes) = Player.Notes
    
    'CharSheetEngine.RefreshTraitList lstTraits(tiNone), Player.SomeList
        
    lblModified.Caption = Format(Player.LastModified, "mmmm d, yyyy")
    
    Me.Show
    
    Game.DataChanged = DataState
    Populating = False

End Sub

Public Sub ShowPP()
'
' Name:         ShowPP
' Description:  Jump to the PP tab.
'

    Set tabTabStrip.SelectedItem = tabTabStrip.Tabs(xpFrame + 1)
    Call tabTabStrip_Click

End Sub

Private Sub FindCharacters()
'
' Name:         FindCharacters
' Description:  Populate lstCharacters with the characters registered to this player.
'

    Dim Query As QueryClass
    Dim Characters As LinkedList
    Dim Values As LinkedList
    Dim PlayerText As String
    Dim FList As LinkedTraitList
    Dim StoreCursor As Integer
    
    Screen.MousePointer = vbHourglass
    
    lstCharacters.Clear
    
    With CharacterList
        .First
        Do Until .Off
            If Player.Name = .Item.Player Then
                lstCharacters.AddItem .Item.Name
            End If
            .MoveNext
        Loop
    End With
    
    lblCharacters.Caption = CStr(lstCharacters.ListCount) & " Characters"
    
    Screen.MousePointer = vbDefault

End Sub

Public Sub SetDefaultOutput()
'
' Name:         SetDefaultOutput
' Description:  Initilize the OutputEngineClass with default output settings.
'
    With OutputEngine
        .Template = tnPPHistory
        .SelectSet(osPlayers).Clear
        .SelectSet(osPlayers).Add Player.Name
        .GameDate = 0
    End With
    
End Sub

Private Sub cmdAdd_Click()
'
' Name:         cmdAdd_Click
' Description:  If a selection is active, have the CharSheetEngine add to
'               the menu, OR add a new XP history entry.
'

    If lstMenu.ListIndex <> -1 Then
        
        If fraTab(tbCharacters).Visible Then
            
            CharacterList.MoveTo lstMenu.Text
            If Not CharacterList.Off Then
                CharacterList.Item.Player = Player.Name
                CharacterList.Item.LastModified = Now
                Game.DataChanged = True
                FindCharacters
            End If
            
        Else
        
            If CharSheetEngine.TargetType = ttXPHistory Then
                If CharSheetEngine.AddXPEntry(lvwHistory, Player.Experience) Then
                    RefreshXP
                    SetDataChanged
                    lvwHistory.SetFocus
                End If
            Else
                CharSheetEngine.AddSelected
                SetDataChanged
            End If
        
        End If
        
    End If
    
End Sub

'Private Sub cmdAscend_Click()
''
'' Name:         cmdAscend_Click
'' Description:  Move the selected entry down.
''
'
'    If cmdAscend.Visible Then
'        CharSheetEngine.MoveEntryUp
'        CharSheetEngine.TargetList.SetFocus
'        SetDataChanged
'    End If
'
'End Sub
'
'Private Sub cmdDescend_Click()
''
'' Name:         cmdDescend_Click
'' Description:  Move the selected entry down.
''
'
'    If cmdDescend.Visible Then
'        CharSheetEngine.MoveEntryDown
'        CharSheetEngine.TargetList.SetFocus
'        SetDataChanged
'    End If
'
'End Sub

Private Sub cmdCustom_Click()
'
' Name:         cmdCustom_Click
' Description:  Have the CharSheetEngine add a custom entry to the target., OR
'               clear the XP history.

    If Not fraTab(tbCharacters).Visible Then
        If CharSheetEngine.TargetType = ttXPHistory Then
            If MsgBox("Are you sure you want to TOTALLY clear this history?", vbYesNo, _
                    "Clear History") = vbYes Then
                Player.Experience.Clear
                SetDataChanged
                RefreshXP
            End If
        Else
            CharSheetEngine.AddCustom
            SetDataChanged
        End If
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
'               entry, OR edit a history entry.
'
    If Not fraTab(tbCharacters).Visible Then
        If CharSheetEngine.TargetType = ttXPHistory Then
            If CharSheetEngine.EditXPEntry(lvwHistory, Player.Experience) Then
                RefreshXP
                SetDataChanged
                lvwHistory.SetFocus
            End If
        Else
            CharSheetEngine.AddNote
            SetDataChanged
        End If
    End If
    
End Sub

Private Sub cmdRemove_Click()
'
' Name:         cmdRemove_Click
' Description:  Have the CharSheetEngine remove a label or list entry, OR
'               remove an XP history entry.
'
    
    If fraTab(tbCharacters).Visible Then
    
        CharacterList.MoveTo lstCharacters.Text
        If Not CharacterList.Off Then
            CharacterList.Item.Player = ""
            CharacterList.Item.LastModified = Now
            Game.DataChanged = True
            FindCharacters
        End If
    
    Else
    
        If CharSheetEngine.TargetType = ttXPHistory Then
            If CharSheetEngine.RemoveXPEntry(lvwHistory, Player.Experience) Then
                RefreshXP
                SetDataChanged
                lvwHistory.SetFocus
            End If
        Else
            CharSheetEngine.RemoveEntry
            SetDataChanged
        End If
    
    End If
    
End Sub

Private Sub cmdRename_Click()
'
' Name:         cmdRename_Click
' Description:  Rename the Player.
'

    Dim NewName As String
    
    NewName = InputBox("Enter a new name for the player.", "Rename Player", txtUserField(xiName).Text)
    NewName = Trim(NewName)
    
    If Not (NewName = "" Or NewName = txtUserField(xiName).Text) Then
        PlayerList.MoveTo NewName
        If Not PlayerList.Off Then
            MsgBox "The name """ & NewName & _
                    """ is already in use.  Please use a different name.", _
                    vbOKOnly Or vbExclamation, "Duplicate Name"
        Else
            Game.Rename qiPlayers, txtUserField(xiName).Text, NewName
            txtUserField(xiName).Text = NewName
            mdiMain.AnnounceChanges Me, atPlayers
        End If
    End If

End Sub

Private Sub cmdShowCharacter_Click()
'
' Name:         cmdShowCharacter_Click
' Description:  Display the selected possessor of this item.
'

    If lstCharacters.ListIndex > -1 Then
    
        mdiMain.ShowCharacterSheet lstCharacters.Text
    
    End If
    
End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  Update the experience total in case it changed elsewhere.  Re-Acquaint
'               the CharacterSheetEngine with the character.
'
    
    If fraTab(xpFrame).Visible Then RefreshXP
    If fraTab(tbCharacters).Visible Then FindCharacters
    lblModified.Caption = Format(Player.LastModified, "mmmm d, yyyy")

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Checks to make sure that a character is loaded, which happens only
'               when ShowVarious is the means of loading the form.  Initializes the
'               MenuStack linked list and the Various Menus.
'

    If Player Is Nothing Then
        MsgBox "Player loaded improperly!"
    Else
        
        Set CharSheetEngine = New CharacterSheetEngineClass
        
        CharSheetEngine.RegisterSheet "Player", lstMenu, lblMenuItem, lblMenuTitle
    
        'CharSheetEngine.RegisterTraitList tiNone, Player.SomeList
    
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
        mdiMain.AnnounceChanges Me, atPlayers
        Select Case Index
            Case fiPosition
                Player.Position = Value
            Case fiStatus
                Player.Status = Value
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

Private Sub lstCharacters_GotFocus()
'
' Name:         lstOwners_Click
' Description:  Hide Increment/Decrement, Populate the menu with active characters.
'
    
    cmdIncrement.Visible = False
    cmdDecrement.Visible = False
    
    CharSheetEngine.PopulateMenu "?CH"

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

Private Sub lstCharacters_DblClick()
'
' Name:         lstCharacters_DblClick
' Description:  Call the cmdShowCharacter button.
'

    Call cmdShowCharacter_Click

End Sub
'
'Private Sub lstTraits_GotFocus(Index As Integer)
''
'' Name:         lstTraits_GotFocus
'' Description:  Attach the Increment/Decrement buttons, shift focus, populate the menus
''
'
'    Dim OrderTop As Integer
'
'    If Not (CharSheetEngine.TargetType = ttListBox And _
'            CharSheetEngine.TargetList Is lstTraits(Index)) Then
'
'        cmdNote.Caption = "Add &Note to Entry"
'        cmdCustom.Caption = "&Custom Entry"
'
'        If CharSheetEngine.TargetType = ttListBox Then _
'                CharSheetEngine.TargetList.ListIndex = -1
'
'        If CharSheetEngine.CanAdjust(Index) Then
'            With lstTraits(Index)
'                Set cmdDecrement.Container = .Container
'                Set cmdIncrement.Container = .Container
'                cmdDecrement.Move .Left, .Top - cmdDecrement.Height
'                cmdIncrement.Move .Left + .Width - cmdIncrement.Width, .Top - cmdIncrement.Height
'                OrderTop = .Top + .Height
'            End With
'            cmdIncrement.Visible = True
'            cmdDecrement.Visible = True
'        Else
'            cmdIncrement.Visible = False
'            cmdDecrement.Visible = False
'            OrderTop = lstTraits(Index).Top - cmdAscend.Height
'        End If
'
'        If CharSheetEngine.CanOrder(Index) Then
'            With lstTraits(Index)
'                Set cmdDescend.Container = .Container
'                Set cmdAscend.Container = .Container
'                cmdDescend.Move .Left, OrderTop
'                cmdAscend.Move .Left + .Width - cmdAscend.Width, OrderTop
'            End With
'            cmdDescend.Visible = True
'            cmdAscend.Visible = True
'        Else
'            cmdDescend.Visible = False
'            cmdAscend.Visible = False
'        End If
'
'        CharSheetEngine.UpdateMenuTitleFromTraitLabel lblTraits(Index)
'        CharSheetEngine.PopulateMenu lstTraits(Index).Tag
'        CharSheetEngine.TargetType = ttListBox
'        Set CharSheetEngine.TargetList = lstTraits(Index)
'
'    End If
'
'End Sub
'
'Private Sub lstTraits_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
''
'' Name:         lstTraits_KeyDown
'' Description:  Keyboard shortcuts
'
'    Select Case KeyCode
'        Case vbKeyDelete, vbKeyBack
'            cmdRemove_Click
'        Case vbKeyLeft
'            cmdDecrement_Click
'            KeyCode = 0
'        Case vbKeyRight
'            cmdIncrement_Click
'            KeyCode = 0
'    End Select
'
'End Sub
'
'Private Sub lstTraits_KeyPress(Index As Integer, KeyAscii As Integer)
''
'' Name:         lstTraits_KeyPress
'' Description:  keyboard shortcuts.
'
'    Select Case KeyAscii
'        Case Asc("-"), Asc("_")
'            cmdDecrement_Click
'        Case Asc("+"), Asc("=")
'            cmdIncrement_Click
'    End Select
'
'End Sub

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
        End If
        
        lblMenuTitle.Caption = "Player Point History"
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
        End If
        lstCharacters.ListIndex = -1
        cmdIncrement.Visible = False
        cmdDecrement.Visible = False
        
        CharSheetEngine.TargetType = ttNothing
        
        Dim fTab As Frame
        For Each fTab In fraTab
            fTab.Visible = (fTab.Index = tabTabStrip.SelectedItem.Index - 1)
        Next fTab
        
        If fraTab(xpFrame).Visible Then
            RefreshXP
            lvwHistory.SetFocus
        Else
            cmdNote.Caption = "Add &Note to Entry"
            cmdCustom.Caption = "&Custom Entry"
        End If
        
        If fraTab(tbCharacters).Visible Then FindCharacters
        
    End If

End Sub

Private Sub txtMemo_Validate(Index As Integer, Cancel As Boolean)
'
' Name:         txtMemo_Change
' Description:  Record changes to the memo field.
'

    Select Case Index
        Case miAddress
            If Player.Address <> txtMemo(miAddress).Text Then
                SetDataChanged
                Player.Address = TrimWhiteSpace(txtMemo(miAddress))
            End If
        Case miNotes
            If Player.Notes <> txtMemo(miNotes).Text Then
                SetDataChanged
                Player.Notes = TrimWhiteSpace(txtMemo(miNotes))
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
        Player.LastModified = Now
        lblModified.Caption = Format(Date, "mmmm d, yyyy")
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
                Player.Experience.Unspent = CSng(txtExperience(xpUnspent))
            Case xpEarned
                Player.Experience.Earned = CSng(txtExperience(xpEarned))
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

Private Sub RefreshXP()
'
' Name:         RefreshXP
' Description:  Refresh the display of XP values and histories.
'

    Populating = True
    txtExperience(xpUnspent).Text = CStr(Player.Experience.Unspent)
    txtExperience(xpEarned).Text = CStr(Player.Experience.Earned)
    txtExperience(xpUnspent).Locked = Game.EnforceHistory
    txtExperience(xpEarned).Locked = Game.EnforceHistory
    updExperience(xpUnspent).Visible = Not Game.EnforceHistory
    updExperience(xpEarned).Visible = Not Game.EnforceHistory
    Populating = False
        
    CharSheetEngine.RefreshXP lvwHistory, Player.Experience
    
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
                Player.ID = Trim(txtUserField(Index))
                mdiMain.AnnounceChanges Me, atPlayers
            Case xiEMail
                Player.EMail = Trim(txtUserField(Index))
            Case xiPhone
                Player.Phone = Trim(txtUserField(Index))
            
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
