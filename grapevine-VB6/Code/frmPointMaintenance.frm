VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPointMaintenance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Experience Points"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9060
   Icon            =   "frmPointMaintenance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   9060
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Height          =   5175
      Index           =   1
      Left            =   360
      TabIndex        =   25
      Top             =   600
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CommandButton cmdShowHistory 
         Caption         =   "Show Character &History"
         Height          =   375
         Left            =   6240
         TabIndex        =   37
         Top             =   4320
         Width           =   1935
      End
      Begin VB.ComboBox cboSearches 
         Height          =   315
         Index           =   1
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   2400
         Width           =   1935
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "&not to View"
         Height          =   255
         Index           =   3
         Left            =   7080
         TabIndex        =   36
         Top             =   2145
         Width           =   1095
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "to &View"
         Height          =   255
         Index           =   2
         Left            =   7080
         TabIndex        =   35
         Top             =   1905
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.ComboBox cboDate 
         Height          =   315
         Index           =   1
         Left            =   6240
         TabIndex        =   30
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox cboAwards 
         Height          =   315
         Index           =   1
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1440
         Width           =   1935
      End
      Begin MSComctlLib.ListView lvwList 
         Height          =   4215
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7435
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   3916
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Award"
            Object.Width           =   3016
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Monthly Earnings"
            Object.Width           =   2857
         EndProperty
      End
      Begin VB.Label lblClue 
         Alignment       =   2  'Center
         Caption         =   "Checkmark a name to record an award for this date."
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   4800
         Width           =   5895
      End
      Begin VB.Label lblAttendance 
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
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Characters"
         Height          =   255
         Index           =   13
         Left            =   6240
         TabIndex        =   33
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Date to View"
         Height          =   255
         Index           =   12
         Left            =   6240
         TabIndex        =   29
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Award to View"
         Height          =   255
         Index           =   11
         Left            =   6240
         TabIndex        =   31
         Top             =   1200
         Width           =   1935
      End
   End
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Height          =   5175
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   8295
      Begin VB.ComboBox cboAwards 
         Height          =   315
         Index           =   0
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   540
         Width           =   3375
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "Select not from search:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   4320
         Width           =   1935
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "&Select from search:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   4080
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.ComboBox cboSearches 
         Height          =   315
         Index           =   0
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   4170
         Width           =   1935
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "&Change Experience"
         Height          =   375
         Left            =   6240
         TabIndex        =   24
         Top             =   4680
         Width           =   1935
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Select &All"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   4680
         Width           =   1935
      End
      Begin VB.CommandButton cmdSelectNone 
         Caption         =   "Select &None"
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   4680
         Width           =   1935
      End
      Begin VB.ListBox lstMany 
         Columns         =   2
         Height          =   3495
         IntegralHeight  =   0   'False
         Left            =   120
         MultiSelect     =   1  'Simple
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox txtReason 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4560
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   3360
         Width           =   3375
      End
      Begin VB.CheckBox chkRecord 
         Alignment       =   1  'Right Justify
         Caption         =   "Record this change in the &histories"
         Height          =   255
         Left            =   5355
         TabIndex        =   23
         Top             =   4200
         Width           =   2820
      End
      Begin VB.TextBox txtNumber 
         Height          =   315
         Index           =   0
         Left            =   6045
         TabIndex        =   15
         Text            =   "1"
         Top             =   1560
         Width           =   570
      End
      Begin VB.ComboBox cboChange 
         Height          =   315
         Index           =   0
         ItemData        =   "frmPointMaintenance.frx":058A
         Left            =   4560
         List            =   "frmPointMaintenance.frx":058C
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ComboBox cboDate 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4560
         TabIndex        =   20
         Top             =   2760
         Width           =   2055
      End
      Begin MSComCtl2.UpDown updNumber 
         Height          =   315
         Index           =   0
         Left            =   6600
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1560
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtNumber(0)"
         BuddyDispid     =   196624
         BuddyIndex      =   0
         OrigLeft        =   6600
         OrigTop         =   1560
         OrigRight       =   7155
         OrigBottom      =   1875
         Max             =   999
         Min             =   -999
         Orientation     =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Standard Experience A&ward"
         Height          =   195
         Index           =   5
         Left            =   4440
         TabIndex        =   9
         Top             =   240
         Width           =   1980
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "C&ustom Change to Experience "
         Height          =   195
         Index           =   1
         Left            =   4440
         TabIndex        =   12
         Top             =   1230
         Width           =   2190
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Characters for whom &to change experience:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4020
      End
      Begin VB.Label lblUpDown 
         BackStyle       =   0  'Transparent
         Caption         =   "&Experience"
         Height          =   375
         Index           =   0
         Left            =   7200
         TabIndex        =   17
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Date"
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   19
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Reason"
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   21
         Top             =   3120
         Width           =   3375
      End
      Begin VB.Label lblExplanation 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   4560
         TabIndex        =   18
         Top             =   1920
         Width           =   3375
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   2655
         Index           =   4
         Left            =   4320
         TabIndex        =   13
         Top             =   1320
         Width           =   3855
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Index           =   6
         Left            =   4320
         TabIndex        =   10
         Top             =   330
         Width           =   3855
      End
   End
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Height          =   5175
      Index           =   2
      Left            =   360
      TabIndex        =   38
      Top             =   600
      Visible         =   0   'False
      Width           =   8295
      Begin VB.ComboBox cboChange 
         Height          =   315
         Index           =   1
         ItemData        =   "frmPointMaintenance.frx":058E
         Left            =   4560
         List            =   "frmPointMaintenance.frx":0590
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtNumber 
         Height          =   315
         Index           =   1
         Left            =   6045
         TabIndex        =   47
         Text            =   "1"
         Top             =   1200
         Width           =   540
      End
      Begin VB.TextBox txtReason 
         Height          =   315
         Index           =   1
         Left            =   4560
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Top             =   1800
         Width           =   3375
      End
      Begin VB.CommandButton cmdAddXPAward 
         Caption         =   "&Add"
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CommandButton cmdDeleteXPAward 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   2160
         TabIndex        =   42
         Top             =   2520
         Width           =   1935
      End
      Begin MSComctlLib.ListView lvwXPAwards 
         Height          =   1935
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Award"
            Object.Width           =   2469
         EndProperty
      End
      Begin MSComCtl2.UpDown updNumber 
         Height          =   315
         Index           =   1
         Left            =   6600
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1200
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtNumber(1)"
         BuddyDispid     =   196624
         BuddyIndex      =   1
         OrigLeft        =   6600
         OrigTop         =   1320
         OrigRight       =   7155
         OrigBottom      =   1635
         Max             =   999
         Min             =   -999
         Orientation     =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblAward 
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
         Height          =   255
         Left            =   4560
         TabIndex        =   44
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "A&ward"
         Height          =   255
         Index           =   9
         Left            =   4560
         TabIndex        =   45
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Reason"
         Height          =   255
         Index           =   8
         Left            =   4560
         TabIndex        =   50
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label lblUpDown 
         BackStyle       =   0  'Transparent
         Caption         =   "Experience"
         Height          =   375
         Index           =   1
         Left            =   7200
         TabIndex        =   49
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Standard Experience Awards"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label lblLabels 
         BorderStyle     =   1  'Fixed Single
         Height          =   1935
         Index           =   10
         Left            =   4320
         TabIndex        =   43
         Top             =   480
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmdESCClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin MSComctlLib.TabStrip tabTabs 
      Height          =   5655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9975
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Group Management"
            Key             =   "Group"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Attendance View"
            Key             =   "View"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Standard Awards"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPointMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FRAME_GROUP = 0
Private Const FRAME_VIEW = 1
Private Const FRAME_AWARDS = 2

Private Const MANAGE_GROUP = 1
Private Const VIEW_ATTENDANCE = 2
Private Const EDIT_AWARDS = 3

Private Const OPT_SELECT = 0
Private Const OPT_NOT = 1
Private Const OPT_VIEW = 2
Private Const OPT_NOTVIEW = 3

Private Const INDEX_GROUP = 0
Private Const INDEX_VIEW = 1
Private Const INDEX_AWARD = 1

Private PointType As String * 1         'A character indicating the kind of maintenance screen this is
Private MaintenanceList As LinkedList   'List of the players or the characters
Private CurrentJob As Integer           'Current frame we're working on
Private Populating As Boolean

Public Sub ShowPointMaintenance(PT As String)
'
' Name:         InitializePointMaintenance
' Parameters:   PT          character indicating what kind of window to load
' Description:  Sets all variables and labels necessary to maintain the
'               points identified by the PT argument.  The screen comes
'               ready to maintain Experience Points; but player points
'               necessitate a reorganization.
'

    PointType = PT
    
    If PointType = pmExperience Then
    
        Set MaintenanceList = CharacterList
    
    Else ' PointType = pmPlayerPoints
    
        Set MaintenanceList = PlayerList
        
        Me.Caption = "Player Points"
        Me.Icon = mdiMain.imlSmallIcons.ListImages("PP").Picture
        lblLabels(0) = "Players for whom &to maintain player points:"
        lblLabels(1) = "C&ustom Change to Player Points"
        lblLabels(5) = "Standard Player Point A&ward"
        lblLabels(7) = "&Standard Player Point A&wards"
        lblLabels(13) = "&Players"
        cmdShowHistory.Caption = "Show Player &History"
        lblUpDown(0) = "&Points"
        lblUpDown(1) = "&Points"
        
        cmdSelectAll.Caption = "&All Players"
        cmdChange.Caption = "Change P&layer Points"
        
    End If

    Me.Tag = PointType
    Me.Show

End Sub

Private Sub RefreshLists()
'
' Name:         RefreshLists
' Description:  Fill the lists of all individuals, preserving selections
'
    
    Select Case CurrentJob
        Case MANAGE_GROUP
            
            Dim Selections As StringSet
            
            Set Selections = New StringSet
            Selections.StoreListBox lstMany
            
            lstMany.Clear
            
            MaintenanceList.First
            Do Until MaintenanceList.Off
                lstMany.AddItem MaintenanceList.Item.Name
                MaintenanceList.MoveNext
            Loop
            
            Selections.SelectListBox lstMany, True, False
            Set Selections = Nothing
        
        Case VIEW_ATTENDANCE
        
            Dim SelKey As String
                        
            If Not lvwList.SelectedItem Is Nothing Then SelKey = lvwList.SelectedItem.Key
            Game.XPAwardList.MoveTo cboAwards(INDEX_VIEW).Text
            lvwList.ListItems.Clear
            
            If IsDate(cboDate(INDEX_VIEW).Text) And Not Game.XPAwardList.Off Then
            
                Dim Results As LinkedList
                Dim NewItem As ListItem
                Dim XP As ExperienceClass
                Dim AwardDate As Date
                Dim Award As ExperienceAwardClass
                Dim ValText As String
                
                lblAttendance.Caption = cboAwards(INDEX_VIEW).Text & ", " & cboDate(INDEX_VIEW).Text
                
                AwardDate = CDate(cboDate(INDEX_VIEW).Text)
                Set Award = Game.XPAwardList.Item
                lvwList.ColumnHeaders(3).Text = Format(AwardDate, "mmmm") & " Total"
                
                Game.QueryEngine.QueryList.MoveTo cboSearches(INDEX_VIEW).Text
                If Game.QueryEngine.QueryList.Off Then
                    Set Results = CharacterList
                Else
                    Game.QueryEngine.MakeQuery Game.QueryEngine.QueryList.Item, _
                                               , optSelect(OPT_NOTVIEW).Value
                    Set Results = Game.QueryEngine.Results
                End If
                
                Results.First
                Do Until Results.Off
                    Set NewItem = lvwList.ListItems.Add(Key:="k" & Results.Item.Name, Text:=Results.Item.Name)
                    Set XP = Results.Item.Experience
                    
                    ValText = "-"
                    XP.MoveToDate AwardDate, True
                    Do Until XP.Off
                        If XP.EntryDate > AwardDate + #11:59:59 PM# Then Exit Do
                        If InStr(1, XP.EntryReason, Award.Reason, vbTextCompare) > 0 Then
                            ValText = CStr(XP.EntryChange) & " " & Award.Name
                            Exit Do
                        End If
                        XP.MoveNext
                    Loop
                    NewItem.ListSubItems.Add Text:=ValText
                    NewItem.ListSubItems.Add Text:=CStr(XP.GetMonthXP(AwardDate))
                    NewItem.Checked = Not (ValText = "-")
                    Results.MoveNext
                Loop
        
                On Error Resume Next
                If SelKey <> "" Then Set lvwList.SelectedItem = lvwList.ListItems(SelKey)
                On Error GoTo 0
            
            Else
            
                lblAttendance.Caption = "Select a date and an award to view characters."
            
            End If
        
    End Select
    
End Sub

Private Sub RefreshDates()
'
' Name:         RefreshDates
' Description:  Fill the combo list of dates, preserving the selection
'
    
    If Not CurrentJob = EDIT_AWARDS Then
        
        Dim Cursor As String
        Dim Index As Integer
        
        Index = IIf(CurrentJob = MANAGE_GROUP, INDEX_GROUP, INDEX_VIEW)
        Cursor = cboDate(Index).Text

        If Not IsDate(Cursor) Then
            Cursor = Format(Now, "mmmm d, yyyy")
            With Game.Calendar
                .MoveToCloseGame
                If Not .Off Then
                    If .GetGameDate > Now Then .MovePrevious
                    If Not .Off Then Cursor = Format(.GetGameDate, "mmmm d, yyyy")
                End If
            End With
        End If
                
        cboDate(Index).Clear
        
        With Game.Calendar
            .First
            Do Until .Off
                cboDate(Index).AddItem Format(.GetGameDate, "mmmm d, yyyy")
                .MoveNext
            Loop
        End With
                
        cboDate(Index).Text = Cursor
    
    End If
        
End Sub

Private Sub RefreshAwards()
'
' Name:         RefreshAwards
' Description:  Refresh the standard awards shown.
'

    Dim Cursor As String
    Dim Index As Integer
    
    Populating = True
    
    Select Case CurrentJob
        Case MANAGE_GROUP, VIEW_ATTENDANCE
     
            Index = IIf(CurrentJob = MANAGE_GROUP, INDEX_GROUP, INDEX_VIEW)
            Cursor = cboAwards(Index).Text
            cboAwards(Index).Clear
            
            With Game.XPAwardList
                .First
                Do Until .Off
                    If .Item.XP = (PointType = pmExperience) Then cboAwards(Index).AddItem .Item.Name
                    If Cursor = .Item.Name Then cboAwards(Index).ListIndex = cboAwards(Index).NewIndex
                    .MoveNext
                Loop
            End With
            
            If CurrentJob = VIEW_ATTENDANCE And cboAwards(Index).ListIndex = -1 _
                And cboAwards(Index).ListCount > 0 Then cboAwards(Index).ListIndex = 0
            
        Case EDIT_AWARDS
        
            Dim NewItem As ListItem
        
            lvwXPAwards.ListItems.Clear
            
            With Game.XPAwardList
                .First
                Do Until .Off
                    If .Item.XP = (PointType = pmExperience) Then
                        Set NewItem = lvwXPAwards.ListItems.Add(Key:="k" & .Item.Name, Text:=.Item.Name)
                        NewItem.ListSubItems.Add Text:=.Item.ChangeTypeText
                    End If
                    .MoveNext
                Loop
                If lvwXPAwards.ListItems.Count > 0 Then
                    Set lvwXPAwards.SelectedItem = lvwXPAwards.ListItems(1)
                    Call lvwXPAwards_ItemClick(lvwXPAwards.SelectedItem)
                End If
            End With

    End Select

    Populating = False

End Sub

Public Sub RefreshSearches()
'
' Name:         RefreshSearches
' Description:  Refresh the dropdown of searches.
'

    Dim Cursor As String
    Dim SearchType As QueryInventoryType
    
    If PointType = pmExperience Then
        SearchType = qiCharacters
    Else
        SearchType = qiPlayers
    End If
    
    Populating = True
    
    Cursor = cboSearches(INDEX_GROUP).Text
    cboSearches(INDEX_GROUP).Clear
    cboSearches(INDEX_VIEW).Clear
    
    With Game.QueryEngine.QueryList
        .First
        Do Until .Off
            If .Item.Inventory = SearchType Then
                cboSearches(INDEX_GROUP).AddItem .Item.Name
                cboSearches(INDEX_VIEW).AddItem .Item.Name
                If Cursor = .Item.Name Then
                    cboSearches(INDEX_GROUP).ListIndex = cboSearches(INDEX_GROUP).NewIndex
                    cboSearches(INDEX_VIEW).ListIndex = cboSearches(INDEX_VIEW).NewIndex
                End If
            End If
            .MoveNext
        Loop
    End With
    
    Populating = False
    
End Sub

Private Sub StoreAward()
'
' Name:         StoreAward
' Description:  Store the award currently specified by the award-edit fields
'

    If Not lvwXPAwards.SelectedItem Is Nothing And Not Populating Then
        With Game.XPAwardList
            .MoveTo lvwXPAwards.SelectedItem.Text
            If Not .Off Then
                .Item.ChangeType = cboChange(INDEX_AWARD).ItemData(cboChange(INDEX_AWARD).ListIndex)
                .Item.Change = CSng(Val(txtNumber(INDEX_AWARD).Text))
                lvwXPAwards.SelectedItem.ListSubItems(1).Text = .Item.ChangeTypeText
                .Item.Reason = Trim(txtReason(INDEX_AWARD).Text)
                Game.DataChanged = True
            End If
        End With
    End If

End Sub

Public Sub SetDefaultOutput()
'
' Name:         SetDefaultOutput
' Description:  Initilize the OutputEngineClass with default output settings.
'
    With OutputEngine
        .Template = tnSignIn
        .SelectSet(osCharacters).Clear
        .SelectSet(osCharacters).StoreListView lvwList, True
        .GameDate = 0
    End With
    
End Sub

Private Sub cboAwards_Click(Index As Integer)
'
' Name:         cboAwards_Click
' Description:  Populate the Change and Reason fields with the chosen award.
'

    Dim I As Integer

    If Not Populating Then

        Select Case Index
            Case INDEX_GROUP
                With Game.XPAwardList
                    .MoveTo cboAwards(INDEX_GROUP).Text
                    If Not .Off Then
                        For I = 0 To cboChange(INDEX_GROUP).ListCount - 1
                            If cboChange(INDEX_GROUP).ItemData(I) = .Item.ChangeType Then
                                cboChange(INDEX_GROUP).ListIndex = I
                            End If
                        Next I
                        txtNumber(INDEX_GROUP).Text = CStr(.Item.Change)
                        txtReason(INDEX_GROUP).Text = .Item.Reason
                    End If
                End With
        
            Case INDEX_VIEW
                RefreshLists
        
        End Select
        
    End If
    
End Sub

Private Sub cboChange_Click(Index As Integer)
'
' Name:         cboChange_Click
' Description:  Explain to the user what he's about to do.  Toggle the visibility
'               of the point field if it's a comment.
'

    Dim ShowNumber As Boolean

    ShowNumber = True

    Select Case cboChange(Index).ItemData(cboChange(Index).ListIndex)
        Case ecEarned
            lblExplanation.Caption = "(Add this number to both Earned and Unspent totals.)"
        Case ecSpent
            lblExplanation.Caption = "(Subtract this number from the Unspent total.)"
        Case ecDeducted
            lblExplanation.Caption = "(Subtract this number from both Earned and Unspent totals.)"
        Case ecUnspent
            lblExplanation.Caption = "(Add this number to the Unspent total.)"
        Case ecSetEarned
            lblExplanation.Caption = "(Set the Earned total to this number.)"
        Case ecSetUnspent
            lblExplanation.Caption = "(Set the Unspent total to this number.)"
        Case ecComment
            lblExplanation.Caption = "(Add a comment to the history without changing any totals.)"
            ShowNumber = False
    End Select

    txtNumber(Index).Visible = ShowNumber
    updNumber(Index).Visible = ShowNumber
    lblUpDown(Index).Visible = ShowNumber

    If Index = INDEX_AWARD Then StoreAward
    
End Sub

Private Sub cboDate_Change(Index As Integer)
'
' Name:         cboDate_Change
' Description:  Sync with other cboDate.
'

    If Not Populating Then
        Populating = True
        cboDate(IIf(Index = INDEX_GROUP, INDEX_VIEW, INDEX_GROUP)).Text = cboDate(Index).Text
        Populating = False
    End If

End Sub

Private Sub cboDate_Click(Index As Integer)
'
' Name:         cboDate_Click
' Description:  Refresh the list if in view attendance mode; sync with other cboDate.
'

    If Not Populating Then
        Populating = True
        cboDate(IIf(Index = INDEX_GROUP, INDEX_VIEW, INDEX_GROUP)).Text = cboDate(Index).Text
        If CurrentJob = VIEW_ATTENDANCE Then RefreshLists
        Populating = False
    End If
    
End Sub

Private Sub cboDate_Validate(Index As Integer, Cancel As Boolean)
'
' Name:         cboDate_Click
' Description:  Refresh the list if in view attendance mode.
'

    If Not Populating Then
        Populating = True
        RefreshLists
        Populating = False
    End If

End Sub

Private Sub cboSearches_Click(Index As Integer)
'
' Name:         cboSearches_Click
' Description:  Select the names found by the search.
'

    If Not Populating Then
        
        Populating = True
        Select Case Index
            Case INDEX_GROUP
                Game.QueryEngine.SelectQueryResults cboSearches(Index).Text, lstMany, optSelect(OPT_NOT).Value
            Case INDEX_VIEW
                RefreshLists
        End Select
        Populating = False
        
    End If

End Sub

Private Sub chkRecord_Click()
'
' Name:         chkRecord_Click
' Description:  Enable history fields.
'

    cboDate(INDEX_GROUP).Enabled = (chkRecord.Value <> vbUnchecked)
    txtReason(INDEX_GROUP).Enabled = cboDate(INDEX_GROUP).Enabled

End Sub

Private Sub cmdAddXPAward_Click()
'
' Name:         cmdAddXPAward_Click
' Description:  Add a new XP award.
'

    Dim NewAward As String
    Dim NewItem As ListItem
    Dim Award As ExperienceAwardClass
    Dim AwardLabel As String
    
    If PointType = pmExperience Then
        AwardLabel = "Enter a name for the new experience award:"
    Else
        AwardLabel = "Enter a name for the new player award:"
    End If
    
    NewAward = InputBox(AwardLabel, "New Award")
    NewAward = Trim(NewAward)
    
    If NewAward <> "" Then
        With Game.XPAwardList
            .MoveTo NewAward
            If .Off Then
                Set Award = New ExperienceAwardClass
                Award.Name = NewAward
                Award.Change = 1
                Award.ChangeType = ecEarned
                Award.XP = (PointType = pmExperience)
                Award.Reason = NewAward
                .InsertSorted Award
                Set NewItem = lvwXPAwards.ListItems.Add(Key:="k" & NewAward, Text:=NewAward)
                NewItem.ListSubItems.Add Text:="earn 1"
                Set lvwXPAwards.SelectedItem = NewItem
                Call lvwXPAwards_ItemClick(NewItem)
                Game.DataChanged = True
            Else
                MsgBox "That name is in use.", vbExclamation + vbOKOnly, "Duplicate Name"
            End If
        End With
    End If

End Sub

Private Sub cmdChange_Click()
'
' Name:         cmdChange_Click
' Description:  Make the specified change in points to the selected characters,
'               recording, if checked, a date and a reason in the histories
'

    Dim Warning As String
    Dim NewReason As String
    Dim PointName As String
    Dim ChangeType As ExperienceChangeType
    
    ChangeType = cboChange(INDEX_GROUP).ItemData(cboChange(INDEX_GROUP).ListIndex)
        
    If PointType = pmExperience Then
        Warning = "The selected characters are about to "
        PointName = "experience"
    Else
        Warning = "The selected players are about to "
        PointName = "player points"
    End If

    Select Case ChangeType
        Case ecEarned
            Warning = Warning & "earn "
        Case ecSpent
            Warning = Warning & "spend "
        Case ecDeducted
            Warning = Warning & "lose "
        Case ecUnspent
            Warning = Warning & "unspend "
        Case ecSetEarned
            Warning = Warning & "set their earned " & PointName & " to "
        Case ecSetUnspent
            Warning = Warning & "set their unspent " & PointName & " to "
        Case ecComment
            Warning = Warning & "have a comment added to their histories."
    End Select
    
    If ChangeType <> ecComment Then _
        Warning = Warning & CStr(Val(txtNumber(INDEX_GROUP).Text)) & " " & PointName & "."
    
    If chkRecord Then
        If IsDate(cboDate(INDEX_GROUP).Text) Then
            Warning = Warning & vbCrLf & vbCrLf & "You will add history entries dated " _
                    & cboDate(INDEX_GROUP).Text & " with "
            NewReason = Replace(txtReason(INDEX_GROUP).Text, vbCrLf, " ")
            NewReason = TrimWhiteSpace(NewReason)
            If NewReason <> "" Then
                Warning = Warning & "the following reason:" & vbCrLf & NewReason
            Else
                Warning = Warning & "NO given reason."
            End If
        Else
            MsgBox "Please supply a legitimate date for the history entry.", vbOKOnly + _
                    vbExclamation, "Date"
            Exit Sub
        End If
    End If
        
    Warning = Warning & vbCrLf & vbCrLf & "Are you sure you want to do this?"
    
    If MsgBox(Warning, vbYesNo + vbQuestion + vbDefaultButton2, "Make Changes") = vbYes Then
    
        Dim Selected As StringSet
        Dim Change As Single
        Dim Changed As Boolean
        Dim XP As ExperienceClass
        
        Screen.MousePointer = vbHourglass
        Changed = False
        Change = Val(txtNumber(INDEX_GROUP).Text)
        Set Selected = New StringSet
        
        Selected.StoreListBox lstMany
        
        MaintenanceList.First
        Do Until MaintenanceList.Off
        
            If Selected.Has(MaintenanceList.Item.Name) Then
            
                Set XP = MaintenanceList.Item.Experience
                If chkRecord.Value <> vbUnchecked Then
                    XP.Insert Change, ChangeType, CDate(cboDate(INDEX_GROUP).Text) + Time, NewReason
                Else
                    Select Case ChangeType
                        Case ecEarned
                            XP.Earned = XP.Earned + Change
                            XP.Unspent = XP.Unspent + Change
                        Case ecDeducted
                            XP.Earned = XP.Earned - Change
                            XP.Unspent = XP.Unspent - Change
                        Case ecSetEarned
                            XP.Earned = Change
                        Case ecSpent
                            XP.Unspent = XP.Unspent - Change
                        Case ecUnspent
                            XP.Unspent = XP.Unspent + Change
                        Case ecSetUnspent
                            XP.Unspent = Change
                        Case ecComment
                            'No change
                    End Select
                End If
                
                MaintenanceList.Item.LastModified = Now
                Changed = True
            
            End If
            
            MaintenanceList.MoveNext
        Loop
    
        If Changed Then RefreshLists
                
        Set Selected = Nothing
        Game.DataChanged = Game.DataChanged Or Changed
        Screen.MousePointer = vbDefault
    
    End If
    
End Sub

Private Sub cmdDeleteXPAward_Click()
'
' Name:         cmdDeleteXPAward
' Description:  Delete the current XP award.
'
    
    If Not lvwXPAwards.SelectedItem Is Nothing Then
        With Game.XPAwardList
            .MoveTo lvwXPAwards.SelectedItem.Text
            If Not .Off Then .Remove
        End With
        lvwXPAwards.ListItems.Remove lvwXPAwards.SelectedItem.Index
        If lvwXPAwards.ListItems.Count > 0 Then
            Set lvwXPAwards.SelectedItem = lvwXPAwards.ListItems(1)
        End If
        Call lvwXPAwards_ItemClick(lvwXPAwards.SelectedItem)
    End If
    
End Sub

Private Sub cmdESCClose_Click()
'
' Name:         cmdESCClose_Click
' Description:  Close the window when the user presses ESC.
'

    Unload Me

End Sub

Private Sub cmdSelectAll_Click()
'
' Name:         cmdSelectAll_Click
' Description:  Select all characters/players.
'

    Dim Entry As Integer
    
    For Entry = 0 To lstMany.ListCount - 1
        lstMany.Selected(Entry) = True
    Next Entry

End Sub

Private Sub cmdSelectNone_Click()
'
' Name:         cmdClearSelect_Click
' Description:  Clear all selected players or characters.
'

    Dim Entry As Integer
    
    For Entry = 0 To lstMany.ListCount - 1
        lstMany.Selected(Entry) = False
    Next Entry

End Sub

Private Sub cmdShowHistory_Click()
'
' Name:         cmdShowHistory_Click
' Description:  Jump to a character's sheet and XP history.
'

    If Not lvwList.SelectedItem Is Nothing Then
        If PointType = pmExperience Then
            mdiMain.ShowCharacterSheet lvwList.SelectedItem.Text, True
        Else
            mdiMain.ShowPlayer lvwList.SelectedItem.Text, True
        End If
    End If

End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  Refresh the lists and experience displays.  Enable/disable
'               features based on experience history preferences.
'

    Screen.MousePointer = vbHourglass

    If CurrentJob = VIEW_ATTENDANCE Then
        RefreshLists
    Else
        If PointType = pmExperience Then
            If mdiMain.CheckForChanges(Me, atCharacters) Then RefreshLists
        Else
            If mdiMain.CheckForChanges(Me, atPlayers) Then RefreshLists
        End If
    End If
    
    If mdiMain.CheckForChanges(Me, atDates) Then RefreshDates
    If mdiMain.CheckForChanges(Me, atQueries) Then RefreshSearches
    
    chkRecord.Value = IIf(Game.EnforceHistory, vbGrayed, _
            IIf(chkRecord.Value = vbUnchecked, vbUnchecked, vbChecked))
    chkRecord.Enabled = Not Game.EnforceHistory
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Initialize data and controls.
'

    Dim DataState As Boolean
    Dim Combo As ComboBox
    
    DataState = Game.DataChanged
    CurrentJob = MANAGE_GROUP
    
    Screen.MousePointer = vbHourglass
    RefreshDates
    RefreshAwards
    RefreshSearches
    RefreshLists
    Screen.MousePointer = vbDefault
    
    Populating = True
    For Each Combo In cboChange
        Combo.AddItem "Earn"
        Combo.ItemData(Combo.NewIndex) = ecEarned
        Combo.AddItem "Spend"
        Combo.ItemData(Combo.NewIndex) = ecSpent
        Combo.AddItem "Lose"
        Combo.ItemData(Combo.NewIndex) = ecDeducted
        Combo.AddItem "Unspend"
        Combo.ItemData(Combo.NewIndex) = ecUnspent
        Combo.AddItem "Set Earned to"
        Combo.ItemData(Combo.NewIndex) = ecSetEarned
        Combo.AddItem "Set Unspent to"
        Combo.ItemData(Combo.NewIndex) = ecSetUnspent
        Combo.AddItem "Comment"
        Combo.ItemData(Combo.NewIndex) = ecComment
        Combo.ListIndex = 0
    Next Combo
    Populating = False
    
    Game.DataChanged = DataState
    
    mdiMain.OrientForm Me
    
End Sub

Private Sub cboDate_GotFocus(Index As Integer)
'
' Name:         cboDate_GotFocus
' Description:  Select this text.
'

    cboDate(Index).SelStart = 0
    cboDate(Index).SelLength = Len(cboDate(Index).Text)

End Sub

Private Sub cboDate_KeyPress(Index As Integer, KeyAscii As Integer)
'
' Name:         cboDate_KeyPress
' Description:  Move to the next field when Return is pressed.
'

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Index = INDEX_GROUP Then
            txtReason(Index).SetFocus
        Else
            Call cboDate_Click(Index)
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Unload
' Description:  Validate remaingin controls.
'
    Me.ValidateControls
    
End Sub

Private Sub lvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'
' Name:         lvwList_ColumnClick
' Description:  Change the order of this sorting.
'

    If lvwList.SortKey = ColumnHeader.Index - 1 Then
        lvwList.SortOrder = IIf(lvwList.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwList.SortKey = ColumnHeader.Index - 1
    End If

End Sub

Private Sub lvwList_DblClick()
'
' Name:         lvwList_DblClick
' Description:  Shortcut to Show Character History
'
    Call cmdShowHistory_Click
    
End Sub

Private Sub lvwList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'
' Name:         lvwList_ItemCheck
' Description:  Add or delete the selected award for the selected character on the selected date.
'

    If Not Item Is Nothing And Not Populating Then
    
        Populating = True
    
        Set lvwList.SelectedItem = Item
    
        MaintenanceList.MoveTo Item.Text
        Game.XPAwardList.MoveTo cboAwards(INDEX_VIEW).Text
        
        If Not (MaintenanceList.Off Or Game.XPAwardList.Off Or _
                Not IsDate(cboDate(INDEX_VIEW).Text)) Then
        
            Dim XP As ExperienceClass
            Dim AwardDate As Date
            Dim Award As ExperienceAwardClass
            
            Set XP = MaintenanceList.Item.Experience
            AwardDate = CDate(cboDate(INDEX_VIEW).Text)
            Set Award = Game.XPAwardList.Item
            
            If Item.Checked Then    'Add Award
            
                XP.Insert Award.Change, Award.ChangeType, AwardDate, Award.Reason
                Item.ListSubItems(1).Text = CStr(Award.Change) & " " & Award.Name
            
            Else                    'Delete Award
                
                XP.MoveToDate AwardDate, True
                Do Until XP.Off
                    If XP.EntryDate > AwardDate + #11:59:59 PM# Then Exit Do
                    If InStr(1, XP.EntryReason, Award.Reason, vbTextCompare) > 0 Then
                        XP.Remove
                    Else
                        XP.MoveNext
                    End If
                Loop
                Item.ListSubItems(1).Text = "-"
                
            End If
            
            Item.ListSubItems(2).Text = CStr(XP.GetMonthXP(AwardDate))
            Game.DataChanged = True
            MaintenanceList.Item.LastModified = Now
            
        End If
        
        Populating = False
        
    End If

End Sub

Private Sub lvwXPAwards_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
' Name:         lvwXPAwards_ItemClick
' Description:  Move to a new XP award, making it available to edit.
'
    
    If Not Item Is Nothing Then
        
        Dim I As Integer
        
        Populating = True
        With Game.XPAwardList
            .MoveTo Item.Text
            If Not .Off Then
                lblAward.Caption = .Item.Name & " Award"
                For I = 0 To cboChange(INDEX_AWARD).ListCount - 1
                    If cboChange(INDEX_AWARD).ItemData(I) = .Item.ChangeType Then _
                        cboChange(INDEX_AWARD).ListIndex = I
                Next I
                txtNumber(INDEX_AWARD).Text = CStr(.Item.Change)
                txtReason(INDEX_AWARD).Text = .Item.Reason
            End If
        End With
        Populating = False
        
    Else
    
        lblAward.Caption = ""
        txtNumber(INDEX_AWARD) = ""
        txtReason(INDEX_AWARD) = ""
    
    End If
    
End Sub

Private Sub optSelect_Click(Index As Integer)
'
' Name:         optSelect_Click
' Description:  Select names found by the current search.
'
    
    If Not Populating Then
        Populating = True
        optSelect((Index + 2) Mod 4).Value = optSelect(Index).Value
        Populating = False
        Call cboSearches_Click(IIf(Index < 2, INDEX_GROUP, INDEX_VIEW))
    End If
    
End Sub

Private Sub tabTabs_Click()
'
' Name:         tabTabs_Click
' Description:  Show the appropriate frame.
'
    
    CurrentJob = tabTabs.SelectedItem.Index
    
    Screen.MousePointer = vbHourglass
    RefreshDates
    RefreshAwards
    RefreshSearches
    RefreshLists
    Screen.MousePointer = vbDefault
    
    fraFrame(FRAME_GROUP).Visible = False
    fraFrame(FRAME_VIEW).Visible = False
    fraFrame(FRAME_AWARDS).Visible = False
    
    Select Case CurrentJob
        Case MANAGE_GROUP:      fraFrame(FRAME_GROUP).Visible = True
        Case VIEW_ATTENDANCE:   fraFrame(FRAME_VIEW).Visible = True
        Case EDIT_AWARDS:       fraFrame(FRAME_AWARDS).Visible = True
    End Select
    
End Sub

Private Sub txtNumber_Change(Index As Integer)
'
' Name:         txtNumber_Change
' Description:  Color the text red and invalidate it if it is not numeric.
'

    txtNumber(Index).ForeColor = IIf(IsNumeric(txtNumber(Index)), _
            vbWindowText, vbHighlight)
    
    If Index = INDEX_AWARD Then StoreAward
    
End Sub

Private Sub txtNumber_GotFocus(Index As Integer)
'
' Name:         txtNumber_GotFocus
' Description:  Select this text.
'

    SelectText txtNumber(Index)

End Sub

Private Sub txtNumber_KeyPress(Index As Integer, KeyAscii As Integer)
'
' Name:         txtNumber_KeyPress
' Description:  Nullify the press of Return and move on.
'

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Index = INDEX_AWARD Then
            txtReason(INDEX_AWARD).SetFocus
        Else
            If cboDate(INDEX_GROUP).Enabled Then cboDate(INDEX_GROUP).SetFocus
        End If
    End If
    
End Sub

Private Sub txtReason_GotFocus(Index As Integer)
'
' Name:         txtReason_GotFocus
' Description:  Select this text.
'

    SelectText txtReason(Index)

End Sub

Private Sub txtReason_KeyPress(Index As Integer, KeyAscii As Integer)
'
' Name:         txtNumber_KeyPress
' Description:  Nullify the press of Return and move on.
'

    If KeyAscii = vbKeyReturn Then KeyAscii = 0

End Sub

Private Sub txtReason_Validate(Index As Integer, Cancel As Boolean)
'
' Name:         txtReason_Validate
' Description:  Store the new reason for the xp award, if needed
'

    If Index = INDEX_AWARD Then StoreAward

End Sub
