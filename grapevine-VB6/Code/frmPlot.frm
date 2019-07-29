VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plot"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "frmPlot.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   9060
   Tag             =   "P"
   Begin VB.Frame fraOutline 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   240
      TabIndex        =   30
      Top             =   1440
      Width           =   8535
      Begin VB.CommandButton cmdRemoveCharacter 
         Caption         =   "&Remove Character"
         Height          =   375
         Left            =   0
         TabIndex        =   28
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CommandButton cmdAddCharacter 
         Caption         =   "Add C&haracter..."
         Height          =   375
         Left            =   0
         TabIndex        =   27
         Top             =   2640
         Width           =   2055
      End
      Begin VB.ListBox lstCharacters 
         Height          =   1575
         IntegralHeight  =   0   'False
         Left            =   0
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   960
         Width           =   2055
      End
      Begin VB.ComboBox cboNarrator 
         Height          =   315
         Left            =   4320
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   2055
      End
      Begin VB.ListBox lstDevelopments 
         Height          =   1575
         IntegralHeight  =   0   'False
         Left            =   6480
         TabIndex        =   3
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add New Development..."
         Height          =   375
         Left            =   6480
         TabIndex        =   4
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Development"
         Height          =   375
         Left            =   6480
         TabIndex        =   5
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox txtOutline 
         Height          =   2535
         Left            =   2160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   960
         Width           =   4215
      End
      Begin VB.ComboBox cboEndDate 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cboStartDate 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblCastCount 
         Alignment       =   2  'Center
         Caption         =   "0 &Key Characters"
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblLabel 
         Caption         =   "&Narrator"
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   21
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label lblDevCount 
         Alignment       =   2  'Center
         Caption         =   "0 Plot De&velopments"
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
         Left            =   6480
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblLastModified 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " January 01, 2002"
         Height          =   315
         Left            =   6480
         TabIndex        =   24
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblLabel 
         Caption         =   "Last Modified"
         Height          =   255
         Index           =   3
         Left            =   6480
         TabIndex        =   23
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label lblLabel 
         Caption         =   "&Plot Outline / Notes"
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
         Index           =   4
         Left            =   2160
         TabIndex        =   0
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label lblLabel 
         Caption         =   "&End Date"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   19
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label lblLabel 
         Caption         =   "&Start Date"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdEscClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin MSComctlLib.TabStrip tabTabs 
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   960
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
      TabWidthStyle   =   2
      Style           =   1
      TabFixedWidth   =   1323
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Outline"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraDevelopment 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   240
      TabIndex        =   29
      Top             =   1440
      Visible         =   0   'False
      Width           =   8535
      Begin VB.TextBox txtDevelopment 
         Height          =   3255
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   240
         Width           =   4935
      End
      Begin VB.CommandButton cmdAddLink 
         Caption         =   "&Add"
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemoveLink 
         Caption         =   "&Remove"
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   3120
         Width           =   1695
      End
      Begin VB.ListBox lstLinks 
         Height          =   2220
         Index           =   0
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddLink 
         Caption         =   "A&dd"
         Height          =   375
         Index           =   1
         Left            =   6840
         TabIndex        =   14
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemoveLink 
         Caption         =   "Re&move"
         Height          =   375
         Index           =   1
         Left            =   6840
         TabIndex        =   15
         Top             =   3120
         Width           =   1695
      End
      Begin VB.ListBox lstLinks 
         Height          =   2265
         Index           =   1
         IntegralHeight  =   0   'False
         Left            =   6840
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblLabel 
         Caption         =   "Affe&cts:"
         Height          =   255
         Index           =   7
         Left            =   6840
         TabIndex        =   12
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label lblLabel 
         Caption         =   "A&ffected By:"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label lblDevCaption 
         Caption         =   "01/01/01 Plot De&velopment"
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
         Left            =   1800
         TabIndex        =   10
         Top             =   0
         Width           =   4935
      End
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
      Height          =   495
      Left            =   2040
      TabIndex        =   32
      Top             =   240
      Width           =   6735
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   840
      Picture         =   "frmPlot.frx":058A
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ceCause = 0
Private Const ceEffect = 1

Private Plot As PlotClass
Private Dev As PlotNode
Private Populating As Boolean
Private ShiftDown As Boolean

Public Sub ShowPlot(ShowPlot As PlotClass)
'
' Name:         ShowPlot
' Parameters:   Plot         the PlotClass object to show
' Description:  Show the given PlotClass object so the user can work with it.
'               Ready the form for its arrival.
'

    Set Plot = ShowPlot
    
    Populating = True
    
    Me.Caption = Plot.Name
    lblTitle.Caption = " " & Plot.Name
    txtOutline.Text = Plot.Outline
    lblLastModified.Caption = Format(Plot.LastModified, "mmmm d, yyyy")
    cboNarrator.Text = Plot.Narrator
    
    Plot.CastList.First
    Do Until Plot.CastList.Off
        lstCharacters.AddItem Plot.CastList.Trait.Name
        Plot.CastList.MoveNext
    Loop
    
    Populating = False
    
    RefreshDates
    RefreshDevelopments
    RefreshNarrators
    
    lstDevelopments.ListIndex = lstDevelopments.ListCount - 1
    
End Sub

Public Sub ShowDevelopment(DevDate As Date)
'
' Name:         ShowDevelopment
' Parameter:    DevDate     Date of development to show
' Description:  Allow other forms to force this form to jump to a given plot development.
'

    ValidateControls

    On Error Resume Next
    Set tabTabs.SelectedItem = tabTabs.Tabs("k" & CStr(DevDate))
    On Error GoTo 0
    
    Call tabTabs_Click
    
    Me.SetFocus

End Sub

Private Sub RefreshDates()
'
' Name:         RefreshDates
' Description:  Fill the combo boxes with dates from the calendar and the null entries.
'               Provide the correct values from the plot, if they're not there.  Provide
'               the null entries.  Select the values from the plot.
'

    Dim CurDate As Date
    
    Populating = True

    cboStartDate.Clear
    cboEndDate.Clear

    cboStartDate.AddItem "(none)"
    cboEndDate.AddItem "(none)"
    If Plot.StartDate = 0 Then cboStartDate.ListIndex = 0
    If Plot.EndDate = 0 Then cboEndDate.ListIndex = 0
    
    With Game.Calendar
    
        .First
        Do Until .Off
            CurDate = .GetGameDate
            cboStartDate.AddItem Format(CurDate, "mmmm d, yyyy")
            cboEndDate.AddItem Format(CurDate, "mmmm d, yyyy")
            If CurDate = Plot.StartDate Then cboStartDate.ListIndex = cboStartDate.NewIndex
            If CurDate = Plot.EndDate Then cboEndDate.ListIndex = cboEndDate.NewIndex
            .MoveNext
        Loop
    
    End With

    If cboStartDate.ListIndex = -1 Then
        cboStartDate.AddItem Format(Plot.StartDate, "mmmm d, yyyy")
        cboStartDate.ListIndex = cboStartDate.NewIndex
    End If
    
    If cboEndDate.ListIndex = -1 Then
        cboEndDate.AddItem Format(Plot.EndDate, "mmmm d, yyyy")
        cboEndDate.ListIndex = cboEndDate.NewIndex
    End If
    
    Populating = False

End Sub

Public Sub RefreshNarrators()
'
' Name:         RefreshNarrators
' Description:  Use the Staff query to rebuild the list of potential narrators.
'

    Dim SaveText As String
    Dim StaffQ As QueryClass
    
    Populating = True
    SaveText = cboNarrator.Text
    
    cboNarrator.Clear
    
    With Game.QueryEngine
        .QueryList.MoveTo "Staff"
        If .QueryList.Off Then
            Set StaffQ = New QueryClass
            StaffQ.Inventory = qiPlayers
            StaffQ.AddClause qkPosition, "Player", 0, qcEquals, True
        Else
            Set StaffQ = .QueryList.Item
        End If
        .MakeQuery StaffQ
    End With
    
    With Game.QueryEngine.Results
        .First
        Do Until .Off
            cboNarrator.AddItem .Item.Name
            .MoveNext
        Loop
    End With
    
    cboNarrator.Text = SaveText
    Populating = False

End Sub

Private Sub RefreshDevelopments()
'
' Name:         RefreshDevelopments
' Description:  Refresh the list of developments.  Populate the tabs as needed.
'

    Dim I As Integer
    Dim Store As String

    Store = tabTabs.SelectedItem.Tag

    lstDevelopments.Clear
    For I = tabTabs.Tabs.Count To 2 Step -1
        tabTabs.Tabs.Remove I
    Next I

    Plot.First
    Do Until Plot.Off
        lstDevelopments.AddItem Format(Plot.PlotDev.DevDate, "mmmm d, yyyy")
        Call tabTabs.Tabs.Add(pvKey:="k" & CStr(Plot.PlotDev.DevDate), _
                              pvcaption:=Format(Plot.PlotDev.DevDate, "mmm dd"))
        tabTabs.Tabs("k" & CStr(Plot.PlotDev.DevDate)).Tag = CStr(Plot.PlotDev.DevDate)
        Plot.MoveNext
    Loop

    On Error Resume Next
    Set tabTabs.SelectedItem = tabTabs.Tabs("k" & Store)
    On Error GoTo 0
    
    If tabTabs.SelectedItem Is Nothing Then
        Set tabTabs.SelectedItem = tabTabs.Tabs(1)
        Call tabTabs_Click
    End If
    
    lblDevCount.Caption = CStr(Plot.Count) & " Plot De&velopments"

End Sub

Private Sub SetDataChanged()
'
' Name:         SetDataChanged
' Description:  Do the upkeep associated with the plot's data changing.
'

    Game.DataChanged = True
    Plot.LastModified = Now
    lblLastModified.Caption = Format(Plot.LastModified, "mmmm d, yyyy")

End Sub

Public Sub SetDefaultOutput()
'
' Name:         SetDefaultOutput
' Description:  Initilize the OutputEngineClass with default output settings.
'
    With OutputEngine
        .Template = tnPlot
        .SelectSet(osPlots).Clear
        .SelectSet(osPlots).Add Plot.Name
        .GameDate = 0
    End With
    
End Sub

Private Sub cboEndDate_Click()
'
' Name:         cboEndDate_Click
' Description:  Record a new Ending date for this plot.
'

    If Not Populating Then
        
        If cboEndDate.Text = "(none)" Then
            Plot.EndDate = 0
        Else
            Plot.EndDate = CDate(cboEndDate.Text)
        End If
        SetDataChanged
        mdiMain.AnnounceChanges Me, atPlots
        
    End If

End Sub

Private Sub cboNarrator_Change()
'
' Name:         cboNarrator_Change
' Description:  Store a new Narrator.
'

    If Not Populating Then
        Plot.Narrator = cboNarrator.Text
        SetDataChanged
    End If

End Sub

Private Sub cboNarrator_Click()
'
' Name:         cboNarrator_Click
' Description:  Store a new Narrator.
'
    Call cboNarrator_Change
    
End Sub

Private Sub cboNarrator_GotFocus()
'
' Name:         cboNarrator_GotFocus
' Description:  Select this text.
'
    cboNarrator.SelStart = 0
    cboNarrator.SelLength = Len(cboNarrator.Text)
    
End Sub

Private Sub cboStartDate_Click()
'
' Name:         cboStartDate_Click
' Description:  Record a new starting date for this plot.
'

    If Not Populating Then
        
        If cboStartDate.Text = "(none)" Then
            Plot.StartDate = 0
        Else
            Plot.StartDate = CDate(cboStartDate.Text)
        End If
        SetDataChanged
        mdiMain.AnnounceChanges Me, atPlots
        
    End If
    
End Sub

Private Sub cmdAddCharacter_Click()
'
' Name:         cmdAddCharacter_Click
' Description:  Show a dialog for selecting characters, and add the one chosen.
'
    
    Dim KeyChar As String
    
    KeyChar = frmSelectFromList.ShowSelect(qiCharacters, "Active Characters", "Select Key Character")

    If KeyChar <> "" Then
        Plot.CastList.Insert KeyChar
        lstCharacters.AddItem KeyChar
        lblCastCount.Caption = CStr(lstCharacters.ListCount) & " &Key Characters"
        SetDataChanged
    End If

End Sub

Private Sub cmdAddLink_Click(Index As Integer)
'
' Name:         cmdAddLink_Click
' Description:  Add a cause or effect link to the subaction.
'

    frmSelectLink.SelectLink Dev.DevDate, (Index = ceEffect)
    If Not frmSelectLink.ChoiceType = aprNone Then
    
        If Index = ceCause Then
            Dev.Causes.AddLink frmSelectLink.ChoiceType, frmSelectLink.When, _
                               frmSelectLink.Item, frmSelectLink.Subitem
        Else
            Dev.Effects.AddLink frmSelectLink.ChoiceType, frmSelectLink.When, _
                                frmSelectLink.Item, frmSelectLink.Subitem
        End If
        
        Dev.Causes.PopulateList lstLinks(ceCause), Dev.DevDate
        Dev.Effects.PopulateList lstLinks(ceEffect), Dev.DevDate
        
        SetDataChanged
        
    End If

End Sub

Private Sub cmdDelete_Click()
'
' Name:         cmdDelete_Click
' Description:  Finds the plot development and asks confirmation of deletion.
'               If yes, remove the development and refill the list.
'

    Dim DelDate As Date
    Dim Answer As Boolean
    
    If lstDevelopments.ListIndex > -1 Then
    
        DelDate = CDate(lstDevelopments.Text)
    
        Plot.MoveTo DelDate
        If Not Plot.Off Then
            
            Answer = ShiftDown
            If Not Answer Then Answer = (MsgBox("This will permanently delete the plot development for " & _
                    lstDevelopments.Text & ". Are you sure you want to delete it?", _
                    vbYesNo + vbQuestion, "Delete Development") = vbYes)
            If Answer Then
                
                Plot.Remove
                SetDataChanged
                mdiMain.AnnounceChanges Me, atPlots
    
            End If
        
        End If
        
        RefreshDevelopments
        
    End If

End Sub

Private Sub cmdAdd_Click()
'
' Name:         cmdAdd_Click
' Description:  Calls on frmGetAPRInfo to display itself.  Retrieve a date or
'               from that window and add that date as a plot development.
'
    
    Dim NewDate As Date
    
    frmGetAPRInfo.GetNewPlotDate Plot.Name
    NewDate = frmGetAPRInfo.NewDate
    Unload frmGetAPRInfo
    
    If Not NewDate = 0 Then
        Plot.MoveTo NewDate
        If Plot.Off Then
            
            Plot.Add NewDate, ""
            SetDataChanged
            mdiMain.AnnounceChanges Me, atPlots
            RefreshDevelopments
                
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

Private Sub cmdRemoveCharacter_Click()
'
' Name:         cmdRemoveCharacter_Click
' Description:  Remove a character from the key characters list.
'

    Dim LI As Integer
    
    LI = lstCharacters.ListIndex

    If LI > -1 Then
        
        Plot.CastList.MoveTo lstCharacters.Text
        Plot.CastList.RemoveTrait
        lstCharacters.RemoveItem LI
        
        If LI >= lstCharacters.ListCount Then LI = lstCharacters.ListCount - 1
        lstCharacters.ListIndex = LI
        
        lblCastCount.Caption = CStr(lstCharacters.ListCount) & " &Key Characters"
        SetDataChanged
        
    End If

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
            Dev.Causes.MoveToPlace Place
            Dev.Causes.RemoveLink
        Else
            Dev.Effects.MoveToPlace Place
            Dev.Effects.RemoveLink
        End If
        
        Dev.Causes.PopulateList lstLinks(ceCause), Dev.DevDate
        Dev.Effects.PopulateList lstLinks(ceEffect), Dev.DevDate
        
        SetDataChanged
        
    End If

End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  Repopulate the date lists or developments as needed.
'
    
    If mdiMain.CheckForChanges(Me, atDates) Then RefreshDates
    If mdiMain.CheckForChanges(Me, atPlots) Then RefreshDevelopments
    If mdiMain.CheckForChanges(Me, atPlayers) Then RefreshNarrators
    If Not Dev Is Nothing And fraDevelopment.Visible Then
        Dev.Causes.PopulateList lstLinks(ceCause), Dev.DevDate
        Dev.Effects.PopulateList lstLinks(ceEffect), Dev.DevDate
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

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Unload
' Description:  Validate the controls before the form is unloaded.
'
    ValidateControls

End Sub

Private Sub lblTitle_Click()
'
' Name:         lblTitle_Click
' Description:  Offer to rename the plot if the title is clicked.
'

    Dim Choices As StringSet
    Dim NewName As String
    
    Set Choices = New StringSet
    
    Choices.Add "Change title"
    
    Select Case mdiMain.CreatePopup(Choices, Me)
        Case "Change title"
                    
            NewName = Trim(InputBox("Enter a new title for this plot:", "Plot Title", Plot.Name))
            
            If Not (NewName = "" Or NewName = Plot.Name) Then
                PlotList.MoveTo NewName
                If PlotList.Off Then
                    Game.APREngine.Reassign PlotList, Plot.Name, NewName
                    Me.Caption = Plot.Name
                    lblTitle.Caption = " " & Plot.Name
                    SetDataChanged
                    If Not Dev Is Nothing Then
                        Dev.Causes.PopulateList lstLinks(ceCause), Dev.DevDate
                        Dev.Effects.PopulateList lstLinks(ceEffect), Dev.DevDate
                    End If
                    mdiMain.AnnounceChanges Me, atPlots
                Else
                    MsgBox "A plot already exists by that title.", _
                           vbOKOnly + vbExclamation, "Rename Plot"
                End If
            End If
            
    End Select
    
    Set Choices = Nothing

End Sub

Private Sub lstCharacters_DblClick()
'
' Name:         lstCharacters_DblClick
' Description:  Show the selected character.
'

    If lstCharacters.ListIndex > -1 Then mdiMain.ShowCharacterSheet lstCharacters.Text

End Sub

Private Sub lstDevelopments_DblClick()
'
' Name:         lstDevelopments_DblClick
' Description:  Jump to the chosen plot development
'

    If Not lstDevelopments.ListIndex = -1 Then
        On Error Resume Next
        Set tabTabs.SelectedItem = tabTabs.Tabs("k" & CStr(CDate(lstDevelopments.Text)))
        On Error GoTo 0
        Call tabTabs_Click
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
            Set LinkList = Dev.Causes
        Else
            Set LinkList = Dev.Effects
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

Private Sub tabTabs_Click()
'
' Name:         tabTabs_Click
' Description:  Show the needed frame and populate it with the needed data.
'
    
    If tabTabs.SelectedItem.Index = 1 Then
        fraOutline.Visible = True
        fraDevelopment.Visible = False
    Else
        
        Plot.MoveTo CDate(tabTabs.SelectedItem.Tag)
        If Not Plot.Off Then
        
            Set Dev = Plot.PlotDev
            
            lblDevCaption.Caption = Format(Dev.DevDate, "Short Date") & " &Plot Development"
            txtDevelopment.Text = Dev.Development
            
            Dev.Causes.PopulateList lstLinks(ceCause), Dev.DevDate
            Dev.Effects.PopulateList lstLinks(ceEffect), Dev.DevDate
        
            fraDevelopment.Visible = True
            fraOutline.Visible = False
            
        End If
    
    End If

End Sub

Private Sub txtDevelopment_Validate(Cancel As Boolean)
'
' Name:         txtDevelopment_Validate
' Description:  Store the new plot development
'

    If Not Dev Is Nothing Then
        txtDevelopment.Text = TrimWhiteSpace(txtDevelopment.Text)
        If txtDevelopment.Text <> Dev.Development Then
            Dev.Development = txtDevelopment.Text
            SetDataChanged
            mdiMain.AnnounceChanges Me, atPlots
        End If
    End If
    
End Sub

Private Sub txtOutline_Validate(Cancel As Boolean)
'
' Name:         txtOutline_Validate
' Description:  Store the new plot outline
'

    txtOutline.Text = TrimWhiteSpace(txtOutline.Text)
    If txtOutline.Text <> Plot.Outline Then
        Plot.Outline = txtOutline.Text
        SetDataChanged
    End If

End Sub
