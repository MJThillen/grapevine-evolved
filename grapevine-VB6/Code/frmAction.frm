VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Action"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "frmAction.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   9060
   Tag             =   "A"
   Begin VB.CheckBox chkDone 
      Alignment       =   1  'Right Justify
      Caption         =   "A&ll Done"
      Height          =   195
      Left            =   5985
      TabIndex        =   29
      Top             =   720
      Width           =   975
   End
   Begin VB.CheckBox chkAdvanced 
      Caption         =   "Sho&w Adv. Actions"
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   660
      Value           =   2  'Grayed
      Width           =   1695
   End
   Begin VB.Frame fraBody 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   240
      TabIndex        =   14
      Top             =   1080
      Width           =   8535
      Begin VB.TextBox txtAction 
         Height          =   1455
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox txtResults 
         Height          =   1455
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   2040
         Width           =   4935
      End
      Begin VB.ListBox lstLinks 
         Height          =   2295
         Index           =   0
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.ListBox lstLinks 
         Height          =   2295
         Index           =   1
         IntegralHeight  =   0   'False
         Left            =   6840
         TabIndex        =   24
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddLink 
         Caption         =   "&Add"
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   17
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemoveLink 
         Caption         =   "Re&move"
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   18
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddLink 
         Caption         =   "A&dd"
         Height          =   375
         Index           =   1
         Left            =   6840
         TabIndex        =   25
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemoveLink 
         Caption         =   "Remo&ve"
         Height          =   375
         Index           =   1
         Left            =   6840
         TabIndex        =   26
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label lblTitle 
         Caption         =   "A&ction"
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
         Index           =   1
         Left            =   1800
         TabIndex        =   19
         Top             =   0
         Width           =   4935
      End
      Begin VB.Label lblTitle 
         Caption         =   "&Results"
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
         Index           =   2
         Left            =   1800
         TabIndex        =   21
         Top             =   1800
         Width           =   4935
      End
      Begin VB.Label lblLabel 
         Caption         =   "Affected &By:"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label lblLabel 
         Caption         =   "Aff&ects:"
         Height          =   255
         Index           =   6
         Left            =   6840
         TabIndex        =   23
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdEscClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "S&how Character"
      Height          =   375
      Left            =   7080
      TabIndex        =   28
      Top             =   240
      Width           =   1695
   End
   Begin VB.Frame fraDetail 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   240
      TabIndex        =   34
      Top             =   720
      Width           =   8535
      Begin VB.CommandButton cmdAddCommon 
         Caption         =   "Add C&ommon Actions"
         Height          =   375
         Left            =   0
         TabIndex        =   33
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddAction 
         Caption         =   "Add Actio&n"
         Height          =   375
         Left            =   0
         TabIndex        =   31
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdDeleteAction 
         Caption         =   "Delete Act&ion"
         Height          =   375
         Left            =   0
         TabIndex        =   32
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtDetail 
         Height          =   285
         Index           =   0
         Left            =   7440
         TabIndex        =   3
         Top             =   435
         Width           =   600
      End
      Begin VB.TextBox txtDetail 
         Height          =   285
         Index           =   1
         Left            =   7440
         TabIndex        =   6
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox txtDetail 
         Height          =   285
         Index           =   2
         Left            =   7440
         TabIndex        =   9
         Top             =   1005
         Width           =   600
      End
      Begin VB.TextBox txtDetail 
         Height          =   285
         Index           =   3
         Left            =   7440
         TabIndex        =   12
         Top             =   1290
         Width           =   600
      End
      Begin MSComCtl2.UpDown updDetail 
         Height          =   285
         Index           =   0
         Left            =   8040
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   435
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDetail(0)"
         BuddyDispid     =   196625
         BuddyIndex      =   0
         OrigLeft        =   8040
         OrigTop         =   435
         OrigRight       =   8535
         OrigBottom      =   720
         Max             =   9999
         Min             =   -9999
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updDetail 
         Height          =   285
         Index           =   1
         Left            =   8040
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDetail(1)"
         BuddyDispid     =   196625
         BuddyIndex      =   1
         OrigLeft        =   8040
         OrigTop         =   720
         OrigRight       =   8535
         OrigBottom      =   1005
         Max             =   9999
         Min             =   -9999
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updDetail 
         Height          =   285
         Index           =   2
         Left            =   8040
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1005
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDetail(2)"
         BuddyDispid     =   196625
         BuddyIndex      =   2
         OrigLeft        =   8040
         OrigTop         =   1005
         OrigRight       =   8535
         OrigBottom      =   1290
         Max             =   9999
         Min             =   -9999
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updDetail 
         Height          =   285
         Index           =   3
         Left            =   8040
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1290
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDetail(3)"
         BuddyDispid     =   196625
         BuddyIndex      =   3
         OrigLeft        =   8040
         OrigTop         =   1290
         OrigRight       =   8535
         OrigBottom      =   1575
         Max             =   9999
         Min             =   -9999
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComctlLib.ListView lvwActions 
         Height          =   1335
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2355
         View            =   3
         LabelEdit       =   1
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
            Key             =   "Category"
            Text            =   "Category"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Key             =   "X"
            Text            =   "Done"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Key             =   "Level"
            Text            =   "Level"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Key             =   "Unused"
            Text            =   "Unused"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Key             =   "Growth"
            Text            =   "Growth"
            Object.Width           =   1235
         EndProperty
      End
      Begin VB.Label lblLabel 
         Caption         =   "De&tailed Actions"
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   0
         Top             =   0
         Width           =   4935
      End
      Begin VB.Label lblLabel 
         Caption         =   "&Level"
         Height          =   285
         Index           =   0
         Left            =   6840
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblLabel 
         Caption         =   "&Unused"
         Height          =   285
         Index           =   1
         Left            =   6840
         TabIndex        =   5
         Top             =   765
         Width           =   615
      End
      Begin VB.Label lblLabel 
         Caption         =   "o&f Total"
         Height          =   285
         Index           =   2
         Left            =   6840
         TabIndex        =   8
         Top             =   1050
         Width           =   615
      End
      Begin VB.Label lblLabel 
         Caption         =   "&Growth"
         Height          =   285
         Index           =   3
         Left            =   6840
         TabIndex        =   11
         Top             =   1335
         Width           =   615
      End
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   840
      Picture         =   "frmAction.frx":058A
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   0
      Left            =   2040
      TabIndex        =   27
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "frmAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Action As ActionClass               'the action object this window manipulates
Private SubAction As ActionNode             'the subaction object presently under scrutiny
Private ShortDate As String
Private Populating As Boolean

Private Const FULL_HEIGHT = 6540
Private Const FULL_FRAME_TOP = 2400
Private Const SHORT_HEIGHT = 5220
Private Const SHORT_FRAME_TOP = 1080

Private Const liMainTitle = 0
Private Const liActionTitle = 1
Private Const liResultTitle = 2

Private Const tiLevel = 0
Private Const tiUnused = 1
Private Const tiTotal = 2
Private Const tiGrowth = 3

Private Const ceCause = 0
Private Const ceEffect = 1

Public Sub ShowAction(Act As ActionClass)
'
' Name:         ShowAction
' Parameters:   Act         the actionclass object to show
' Description:  Show the given ActionClass object so the user can work with it.
'               Ready the form for its arrival.
'

    Set Action = Act
    
    Me.Caption = Action.Name
    ShortDate = Format(Action.ActDate, "Short Date")
    lblTitle(liMainTitle).Caption = " " & Action.Name
    chkDone.Value = IIf(Action.Done, vbChecked, vbUnchecked)
    
    If Action.Count = 0 Then
        Action.Add BasicSubactionName, 0, Game.APREngine.PersonalActions, Game.APREngine.PersonalActions, 0
    End If
    
    RefreshSubactionList
    
    If Action.Count > 1 Then
        chkAdvanced.Value = vbChecked
    Else
        chkAdvanced.Value = vbUnchecked
    End If

    mdiMain.OrientForm Me

    Me.Show

End Sub

Public Sub ShowSubaction(SubAct As String)
'
' Name:         ShowSubaction
' Parameters:   SubAct      The name of the subaction to show
' Description:  Allow other forms to force this window to show a certain subaction
'

    ValidateControls

    If Not SubAct = BasicSubactionName Then
        chkAdvanced.Value = vbChecked       'triggers chkAdvanced_click
    End If
    
    On Error Resume Next
    Set lvwActions.SelectedItem = lvwActions.ListItems("k" & SubAct)
    On Error GoTo 0

    Call lvwActions_ItemClick(lvwActions.SelectedItem)

End Sub

Public Sub RefreshSubactionList()
'
' Name:         RefreshSubactionList
' Description:  Clear and repopulate the list of subactions
'

    Dim StoreIndex As Integer
    Dim NewItem As ListItem

    StoreIndex = 1
    If Not lvwActions.SelectedItem Is Nothing Then StoreIndex = lvwActions.SelectedItem.Index
    lvwActions.ListItems.Clear

    Action.First
    Do Until Action.Off
        Set NewItem = lvwActions.ListItems.Add(Key:="k" & Action.SubAction.Name, Text:=Action.SubAction.Name)
        NewItem.ListSubItems.Add , "done", IIf(Action.SubAction.IsComplete, "X", "-")
        NewItem.ListSubItems.Add , "level", CStr(Action.SubAction.Level)
        NewItem.ListSubItems.Add , "use", CStr(Action.SubAction.Unused) & "/" & CStr(Action.SubAction.Total)
        NewItem.ListSubItems.Add , "growth", CStr(Action.SubAction.Growth)
        Action.MoveNext
    Loop

    If StoreIndex > lvwActions.ListItems.Count Then StoreIndex = lvwActions.ListItems.Count
    Set lvwActions.SelectedItem = lvwActions.ListItems(StoreIndex)
    lvwActions_ItemClick lvwActions.SelectedItem

End Sub

Public Sub SetDefaultOutput()
'
' Name:         SetDefaultOutput
' Description:  Initilize the OutputEngineClass with default output settings.
'
    With OutputEngine
        .Template = tnActionRumor
        .SelectSet(osActions).Clear
        .SelectSet(osActions).Add Action.Name
        .GameDate = Action.ActDate
    End With

End Sub

Private Sub chkDone_Click()
'
' Name:         chkDone_Click
' Description:  Set the Done status of this action.
'

    If Action.Done <> (chkDone.Value = vbChecked) Then
        Action.Done = (chkDone.Value = vbChecked)
        Game.DataChanged = True
        Action.LastModified = Now
        mdiMain.AnnounceChanges Me, atActions
    End If
    
End Sub

Private Sub cmdAddAction_Click()
'
' Name:         cmdAddAction_Click
' Description:  Add a new subaction with the user-provided name
'

    Dim NewAct As String
    
    NewAct = InputBox("Enter a name for the new subaction:", "Add Action")
    NewAct = Trim(NewAct)
    Action.MoveTo NewAct
    If NewAct <> "" And Action.Off Then
        Action.Add NewAct, 0, 0, 0, 0
        RefreshSubactionList
        Action.LastModified = Now
        Game.DataChanged = True
    End If

End Sub

Private Sub cmdAddCommon_Click()
'
' Name:         cmdAddCommon_Click
' Description:  Add the common subactions to this action.
'
    Action.AddCommonActions
    RefreshSubactionList
    Action.LastModified = Now
    Game.DataChanged = True
    
End Sub

Private Sub cmdAddLink_Click(Index As Integer)
'
' Name:         cmdAddLink_Click
' Description:  Add a cause or effect link to the subaction.
'

    frmSelectLink.SelectLink Action.ActDate, (Index = ceEffect)
    If Not frmSelectLink.ChoiceType = aprNone Then
    
        If Index = ceCause Then
            SubAction.Causes.AddLink frmSelectLink.ChoiceType, frmSelectLink.When, _
                                     frmSelectLink.Item, frmSelectLink.Subitem
        Else
            SubAction.Effects.AddLink frmSelectLink.ChoiceType, frmSelectLink.When, _
                                     frmSelectLink.Item, frmSelectLink.Subitem
        End If
    
        SubAction.Causes.PopulateList lstLinks(ceCause), Action.ActDate
        SubAction.Effects.PopulateList lstLinks(ceEffect), Action.ActDate
        
        Action.LastModified = Now
        Game.DataChanged = True
        
    End If

End Sub

Private Sub cmdDeleteAction_Click()
'
' Name:         cmdDeleteAction_Click
' Description:  Delect the selected subaction.
'

    If Not lvwActions.SelectedItem Is Nothing Then
        
        If lvwActions.SelectedItem.Text = BasicSubactionName Then
            MsgBox "You cannot delete a character's personal actions.", vbOKCancel, "Delete"
            Exit Sub
        End If
    
        If MsgBox("Are you sure you want to delete this subaction?", _
                   vbYesNo + vbQuestion, "Delete Subaction") = vbYes Then
            
            Action.MoveTo lvwActions.SelectedItem.Text
            Action.Remove
            RefreshSubactionList
            Action.LastModified = Now
            Game.DataChanged = True
                 
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

Private Sub cmdRemoveLink_Click(Index As Integer)
'
' Name:         cmdRemoveLink_Click
' Description:  Remove the selected link from the cause or effect list.
'

    If lstLinks(Index).ListIndex > -1 Then
    
        Dim Place As Integer
        
        Place = lstLinks(Index).ItemData(lstLinks(Index).ListIndex)
        
        If Index = ceCause Then
            SubAction.Causes.MoveToPlace Place
            SubAction.Causes.RemoveLink
        Else
            SubAction.Effects.MoveToPlace Place
            SubAction.Effects.RemoveLink
        End If
        
        SubAction.Causes.PopulateList lstLinks(ceCause), Action.ActDate
        SubAction.Effects.PopulateList lstLinks(ceEffect), Action.ActDate
        
        Action.LastModified = Now
        Game.DataChanged = True
        
    End If

End Sub

Private Sub cmdShow_Click()
'
' Name:         cmdShow_Click
' Description:  Show the character associated with this action.
'
    
    mdiMain.ShowCharacterSheet Action.CharName

End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  Refresh the link lists when the form is reactivated.
'

    SubAction.Causes.PopulateList lstLinks(ceCause), Action.ActDate
    SubAction.Effects.PopulateList lstLinks(ceEffect), Action.ActDate
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Unload
' Description:  Validate all controls before unloading.
'
    ValidateControls

End Sub

Private Sub lblTitle_Click(Index As Integer)
'
' Name:         lblTitle_Click
' Description:  Offer the option to reassign characters/dates if clicking the title.
'

    If Index = liMainTitle Then

        Dim Choices As StringSet
        Dim NewDate As Date
        Dim NewName As String
        
        Set Choices = New StringSet
        
        Choices.Add "Change character"
        Choices.Add "Change date"
        
        Select Case mdiMain.CreatePopup(Choices, Me)
            Case "Change character"
                frmGetAPRInfo.GetNewActionChar Action.ActDate
            Case "Change date"
                frmGetAPRInfo.GetNewActionDate Action.CharName
            Case Else
                Exit Sub
        End Select
        
        NewDate = frmGetAPRInfo.NewDate
        NewName = Trim(frmGetAPRInfo.NewItem)
        Unload frmGetAPRInfo
        
        If Not (NewName = "" Or (NewName = Action.CharName And NewDate = Action.ActDate)) Then
        
            Game.APREngine.MoveToPair ActionList, NewDate, NewName
            If ActionList.Off Then
                Game.APREngine.Reassign ActionList, Action.CharName, NewName, Action.ActDate, NewDate
                ShortDate = Format(Action.ActDate, "Short Date")
                Me.Caption = Action.Name
                lblTitle(liMainTitle).Caption = " " & Action.Name
                SubAction.Causes.PopulateList lstLinks(ceCause), Action.ActDate
                SubAction.Effects.PopulateList lstLinks(ceEffect), Action.ActDate
                Action.LastModified = Now
                Game.DataChanged = True
                mdiMain.AnnounceChanges Me, atActions
            Else
                MsgBox "An action already exists for that character and date.", _
                       vbOKOnly + vbExclamation, "Reassign Action"
            End If
                
        End If
        
        Set Choices = Nothing

    End If
    
End Sub

Private Sub lstLinks_Click(Index As Integer)
'
' Name:         lstLinks_Click
' Description:  Set the tooltip of this control to match the selection,
'               helping the user view captions that don't fit.
'

    lstLinks(Index).ToolTipText = lstLinks(Index).Text

End Sub

Private Sub lstLinks_DblClick(Index As Integer)
'
' Name:         lstLinks_Click
' Description:  Show the window associated with the link double-clicked.
'

    If lstLinks(Index).ListIndex > -1 Then
        
        Dim LinkList As CauseEffectList
        
        If Index = ceCause Then
            Set LinkList = SubAction.Causes
        Else
            Set LinkList = SubAction.Effects
        End If
        
        LinkList.MoveToPlace lstLinks(Index).ItemData(lstLinks(Index).ListIndex)
        
        If Not LinkList.Off Then
            With LinkList.Link
                mdiMain.ShowAPR .Target, .Item, .When, .Subitem
            End With
        End If
                
    End If

End Sub

Private Sub lstLinks_KeyPress(Index As Integer, KeyAscii As Integer)
'
' Name:         lstLinks_KeyPress
' Description:  Translate an enter into a double-click.
'
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call lstLinks_DblClick(Index)
    End If
    
End Sub

Private Sub lvwActions_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
' Name:         lvwActions_ItemClick
' Description:  Populate the action window with the data from the selected item.
'

    Action.MoveTo Item.Text
    If Not Action.Off Then
    
        Set SubAction = Action.SubAction
        cmdDeleteAction.Enabled = Not (SubAction.Name = BasicSubactionName)
        
        Populating = True
    
        txtDetail(tiLevel).Text = CStr(SubAction.Level)
        txtDetail(tiUnused).Text = CStr(SubAction.Unused)
        txtDetail(tiTotal).Text = CStr(SubAction.Total)
        txtDetail(tiGrowth).Text = CStr(SubAction.Growth)
        
        Populating = False
        
        SubAction.Causes.PopulateList lstLinks(ceCause), Action.ActDate
        SubAction.Effects.PopulateList lstLinks(ceEffect), Action.ActDate
    
        lblTitle(liActionTitle).Caption = ShortDate & " " & SubAction.Name & " A&ction"
        lblTitle(liResultTitle).Caption = ShortDate & " " & SubAction.Name & " &Results"
        txtAction.Text = SubAction.Action
        txtResults.Text = SubAction.Result

    End If
    
End Sub

Private Sub chkAdvanced_Click()
'
' Name:         chkAdvanced_Click
' Description:  Show or hide the advanced action tools as needed.
'

    If chkAdvanced.Value = vbChecked Then
        Me.Height = FULL_HEIGHT
        fraBody.Top = FULL_FRAME_TOP
        fraDetail.Visible = True
    Else
        Me.Height = SHORT_HEIGHT
        fraBody.Top = SHORT_FRAME_TOP
        fraDetail.Visible = False
    End If
    
    Set lvwActions.SelectedItem = lvwActions.ListItems(1)
    lvwActions_ItemClick lvwActions.SelectedItem

End Sub

Private Sub txtAction_Validate(Cancel As Boolean)
'
' Name:         txtAction_Validate
' Description:  Store the new action text.
'

    If Not (SubAction Is Nothing) Then
    
        If txtAction.Text <> SubAction.Action Then
            SubAction.Action = TrimWhiteSpace(txtAction.Text)
            Game.DataChanged = True
            Action.LastModified = Now
            lvwActions.SelectedItem.ListSubItems("done").Text = IIf(SubAction.IsComplete, "X", "-")
            Action.IfDoneSetDone
            chkDone.Value = IIf(Action.Done, vbChecked, vbUnchecked)
        End If
    
    End If

End Sub

Private Sub txtDetail_Change(Index As Integer)
'
' Name:         txtDetail_Change
' Description:  Apply changes made to the detailed subaction text boxes.
'

    If Not (Populating Or SubAction Is Nothing) Then
    
        Dim Value As Integer
        Value = Val(txtDetail(Index).Text)
    
        Select Case Index
            Case tiLevel
                SubAction.Level = Value
                lvwActions.SelectedItem.ListSubItems("level").Text = CStr(Value)
            Case tiUnused
                SubAction.Unused = Value
                lvwActions.SelectedItem.ListSubItems("use").Text = CStr(Value) & "/" & CStr(SubAction.Total)
            Case tiTotal
                SubAction.Total = Value
                lvwActions.SelectedItem.ListSubItems("use").Text = CStr(SubAction.Unused) & "/" & CStr(Value)
            Case tiGrowth
                SubAction.Growth = Value
                lvwActions.SelectedItem.ListSubItems("growth").Text = CStr(Value)
        End Select
        
        Game.DataChanged = True
        Action.LastModified = Now
        
    End If

End Sub

Private Sub txtDetail_GotFocus(Index As Integer)
'
' Name:         txtDetail_GotFocus
' Description:  Select the number.
'
    SelectText txtDetail(Index)

End Sub

Private Sub txtResults_Validate(Cancel As Boolean)
'
' Name:         txtResults_Validate
' Description:  Store the new results text.
'

    If Not (SubAction Is Nothing) Then
    
        If txtResults.Text <> SubAction.Result Then
            
            Dim OldDone As Boolean
        
            SubAction.Result = TrimWhiteSpace(txtResults.Text)
            Game.DataChanged = True
            Action.LastModified = Now
            lvwActions.SelectedItem.ListSubItems("done").Text = IIf(SubAction.IsComplete, "X", "-")
            
            OldDone = Action.Done
            Action.IfDoneSetDone
            If Not OldDone = Action.Done Then
                chkDone.Value = IIf(Action.Done, vbChecked, vbUnchecked)
                mdiMain.AnnounceChanges Me, atActions
            End If
            
        End If
    
    End If

End Sub

