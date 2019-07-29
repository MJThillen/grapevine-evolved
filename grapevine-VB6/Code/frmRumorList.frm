VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRumorList 
   Caption         =   "Rumors"
   ClientHeight    =   5775
   ClientLeft      =   1875
   ClientTop       =   750
   ClientWidth     =   7605
   Icon            =   "frmRumorList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   7605
   Begin VB.CommandButton cmdView 
      Caption         =   "&View by Title"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   5160
      Width           =   1695
   End
   Begin MSComctlLib.ListView lvwTitles 
      Height          =   5100
      Left            =   2040
      TabIndex        =   15
      Top             =   480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   8996
      SortKey         =   2
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "rumor"
         Text            =   "Rumor"
         Object.Width           =   4419
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   "done"
         Text            =   "Done"
         Object.Width           =   1005
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "sortkey"
         Text            =   "Sort Key"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.ListBox lstDates 
      Height          =   4575
      IntegralHeight  =   0   'False
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.Frame fraRight 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   5640
      TabIndex        =   4
      Top             =   480
      Width           =   1725
      Begin VB.CommandButton cmdDeleteAll 
         Caption         =   "Delete A&ll"
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   4680
         Width           =   1725
      End
      Begin VB.CommandButton cmdStandard 
         Caption         =   "Add Standard R&umors"
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   3720
         Width           =   1725
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "&Add New Rumor"
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   3240
         Width           =   1725
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Show Rumor"
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   2280
         Width           =   1725
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Rumor"
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   4200
         Width           =   1725
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Last Modified:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1485
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
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   1485
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   615
         Picture         =   "frmRumorList.frx":058A
         Tag             =   "11"
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblName 
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
         Height          =   765
         Left            =   120
         TabIndex        =   6
         Top             =   660
         Width           =   1485
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Index           =   2
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1725
      End
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
   Begin MSComctlLib.ListView lvwDates 
      Height          =   5100
      Left            =   2040
      TabIndex        =   16
      Top             =   480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   8996
      SortKey         =   2
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "date"
         Text            =   "Date"
         Object.Width           =   4419
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   "done"
         Text            =   "Done"
         Object.Width           =   1005
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "sortkey"
         Text            =   "SortKey"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblCategory 
      Alignment       =   2  'Center
      Caption         =   "Dat&es"
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
      Width           =   1695
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      Caption         =   "0 &Rumors"
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
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmRumorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' Name:         frmRumorList
' Description:  Manage the game's list of rumors.
'
Private ShiftDown As Boolean
Private CategoryDate As Boolean

Private Const FORM_START_HEIGHT = 6180
Private Const FORM_START_WIDTH = 7740
Private Const FORM_MIN_SCALEHEIGHT = 5775
Private Const FORM_MIN_SCALEWIDTH = 5910
Private Const BOTTOM_MARGIN = 255
Private Const RIGHT_MARGIN = 1980
Private Const VERTICAL_GAP = 225
Private Const BUTTON_GAP = 120
Private Const TOP_MARGIN = 480
Private Const RUMOR_LEFT = 2040
Private Const CATEGORY_LEFT = 240
Private Const CATEGORY_WIDTH = 1695
Private Const DONE_WIDTH = 570
Private Const SCROLL_WIDTH = 300

Private Sub RefreshCategories(Optional SelGame As Boolean = False)
'
' Name:         RefreshCategories
' Parameters:   SelGame     if TRUE, select the next/last game date;
'               else preserve current slection
' Description:  Refill the category list with dates/rumor titles.
'

    Dim StoreCategory As Integer
    Dim SelDate As Date
    
    Screen.MousePointer = vbHourglass
    
    If CategoryDate Then
    
        If SelGame Then
            With Game.Calendar
                If .HasNextGame Then
                    SelDate = .NextGameDate
                ElseIf .HasPreviousGame Then
                    SelDate = .PreviousGameDate
                End If
            End With
        End If
    
        StoreCategory = lstDates.ListIndex
        lstDates.Clear
        
        Game.Calendar.Last
        Do Until Game.Calendar.Off
            lstDates.AddItem Format(Game.Calendar.GetGameDate, "mmmm d, yyyy")
            If SelGame Then
                If Game.Calendar.GetGameDate = SelDate Then
                    StoreCategory = lstDates.NewIndex
                End If
            End If
            Game.Calendar.MovePrevious
        Loop
    
        If StoreCategory >= lstDates.ListCount Then StoreCategory = lstDates.ListCount - 1
        If StoreCategory = -1 And lstDates.ListCount > 0 Then StoreCategory = 0
        lstDates.ListIndex = StoreCategory       'this triggers lstDates_Click, which
                                                 'triggers RefreshRumors
    
    Else
    
        Dim I As Integer
        Dim CatSet As StringSet
        Dim NewItem As ListItem
        Dim Rumor As RumorClass
        
        StoreCategory = 1
        If Not lvwTitles.SelectedItem Is Nothing Then _
            StoreCategory = lvwTitles.SelectedItem.Index
        lvwTitles.ListItems.Clear
    
        Set CatSet = New StringSet
        
        RumorList.First
        Do Until RumorList.Off
            Set Rumor = RumorList.Item
            If Not CatSet.Has(Rumor.Title) Then
                CatSet.Add Rumor.Title
                Set NewItem = lvwTitles.ListItems.Add(Key:="k" & Rumor.Title, _
                                             Text:=Rumor.Title, SmallIcon:=Rumor.IconKey)
                Call NewItem.ListSubItems.Add(Text:="")
                Call NewItem.ListSubItems.Add(Text:=(CStr(Rumor.Category) & Rumor.Title))
            End If
            RumorList.MoveNext
        Loop
    
        Set CatSet = Nothing

        On Error Resume Next
        If StoreCategory > lvwTitles.ListItems.Count Then StoreCategory = 1
        Set lvwTitles.SelectedItem = lvwTitles.ListItems(StoreCategory)
        On Error GoTo 0
        
        RefreshRumors

    End If
        
    Screen.MousePointer = vbDefault

End Sub

Private Sub RefreshRumors()
'
' Name:         RefreshRumors
' Description:  Refresh the list of Rumors based on the current date/character.
'

    Dim StoreRumor As Integer
    Dim NewItem As ListItem
    
    If CategoryDate Then
    
        StoreRumor = 1
        If Not lvwTitles.SelectedItem Is Nothing Then _
            StoreRumor = lvwTitles.SelectedItem.Index
        lvwTitles.ListItems.Clear
        lblCount.Caption = "0 &Rumors"
        
        If Not lstDates.Text = "" Then
        
            Dim SelDate As Date
            Dim Rumor As RumorClass
            
            SelDate = CDate(lstDates.Text)
            Game.APREngine.MoveToFirstDate RumorList, SelDate
                
            Do Until RumorList.Off
                Set Rumor = RumorList.Item
                Set NewItem = lvwTitles.ListItems.Add(Key:="k" & Rumor.Title, _
                                             Text:=Rumor.Title, SmallIcon:=Rumor.IconKey)
                Call NewItem.ListSubItems.Add(Text:=IIf(Rumor.Done, "X", ""), Key:="done")
                Call NewItem.ListSubItems.Add(Text:=(CStr(Rumor.Category) & Rumor.Title), Key:="sortkey")
                Game.APREngine.MoveToNextDate RumorList, SelDate
            Loop
                
            lblCount.Caption = CStr(lvwTitles.ListItems.Count) & " &Rumors for " & lstDates.Text
            
            On Error Resume Next
            If StoreRumor >= lvwTitles.ListItems.Count Then StoreRumor = 1
            Set lvwTitles.SelectedItem = lvwTitles.ListItems(StoreRumor)
            On Error GoTo 0
        
        End If
        
    Else

        StoreRumor = 1
        If Not lvwDates.SelectedItem Is Nothing Then _
            StoreRumor = lvwDates.SelectedItem.Index
        lvwDates.ListItems.Clear
        lblCount.Caption = "0 &Rumors"
        
        If Not lvwTitles.SelectedItem Is Nothing Then
                
            Dim SelName As String
            Dim CurDate As Long
            
            SelName = lvwTitles.SelectedItem.Text
            Game.APREngine.MoveToFirstTitle RumorList, SelName
            
            Do Until RumorList.Off
                CurDate = RumorList.Item.RumorDate
                                
                Set NewItem = lvwDates.ListItems.Add(Text:=Format(CurDate, "mmmm d, yyyy"), _
                       Key:="k" & CStr(CurDate))
                Call NewItem.ListSubItems.Add(Text:=IIf(RumorList.Item.Done, "X", ""), Key:="done")
                Call NewItem.ListSubItems.Add(Text:=Format(CurDate, "yyyy-mm-dd"), Key:="sortkey")
                
                Game.APREngine.MoveToNextTitle RumorList, SelName
            Loop
        
            lblCount.Caption = CStr(lvwDates.ListItems.Count) & " &Rumors for " & SelName
                
            On Error Resume Next
            If StoreRumor >= lvwDates.ListItems.Count Then StoreRumor = 1
            Set lvwDates.SelectedItem = lvwDates.ListItems(StoreRumor)
            On Error GoTo 0
            
        End If
        
    End If

    RefreshInfo

End Sub

Private Sub RefreshInfo()
'
' Name:         RefreshInfo
' Description:  Find the character and display the appropriate information at right.
'

    Dim ShowDate As Date
    Dim ShowName As String
    
    ShowName = ""
    ShowDate = 0
    
    If Not (lvwTitles.SelectedItem Is Nothing) Then ShowName = lvwTitles.SelectedItem.Text

    If CategoryDate Then
        If lstDates.ListIndex > -1 Then ShowDate = CDate(lstDates.Text)
    Else
        If Not (lvwDates.SelectedItem Is Nothing) Then ShowDate = CDate(lvwDates.SelectedItem.Text)
    End If

    If Not (ShowDate = 0 Or ShowName = "") Then
        
        Game.APREngine.MoveToPair RumorList, ShowDate, ShowName
        If Not RumorList.Off Then
            Dim Rumor As RumorClass
            Set Rumor = RumorList.Item
            lblName.Caption = Rumor.Name
            lblDate.Caption = Format(Rumor.LastModified, "Short Date")
            imgIcon.Picture = mdiMain.imlIcons.ListImages(Rumor.IconKey).Picture
            imgIcon.Visible = True
        Else
            MsgBox "Grapevine can't find this rumor!  Was it renamed or deleted?", vbExclamation
        End If
    
    Else
        lblName.Caption = ""
        lblDate.Caption = ""
        imgIcon.Visible = False
    End If

End Sub

Public Sub SetDefaultOutput()
'
' Name:         SetDefaultOutput
' Description:  Initilize the OutputEngineClass with default output settings.
'
    
    Dim ShowDate As Date
    Dim ShowName As String
    
    ShowName = ""
    ShowDate = 0
    
    If CategoryDate Then
        If lstDates.ListIndex > -1 Then
            ShowDate = CDate(lstDates.Text)
            Game.APREngine.MoveToFirstDate RumorList, ShowDate
        Else
            Exit Sub
        End If
    Else
        If Not (lvwTitles.SelectedItem Is Nothing) Then
            ShowName = lvwTitles.SelectedItem.Text
            Game.APREngine.MoveToFirstTitle RumorList, ShowName
        Else
            Exit Sub
        End If
    End If

    If Not RumorList.Off Then
    
        With OutputEngine
            .SelectSet(osRumors).Clear
            Do Until RumorList.Off
                .SelectSet(osRumors).Add RumorList.Item.Name
                If CategoryDate Then
                    Game.APREngine.MoveToNextDate RumorList, ShowDate
                Else
                    Game.APREngine.MoveToNextTitle RumorList, ShowName
                End If
            Loop
            .Template = tnMasterRumor
            .GameDate = 0
        End With
    
    End If
    
End Sub

Private Sub cmdAddNew_Click()
'
' Name:         cmdAddNew_Click
' Description:  Calls on frmGetAPRInfo to display itself and return a date and
'               and title to create an Rumor for.
'
    
    Dim NewDate As Date
    Dim NewName As String
    
    If CategoryDate And lstDates.ListIndex > -1 Then
        frmGetAPRInfo.GetNewRumorTitle CDate(lstDates.Text)
    ElseIf Not CategoryDate And Not lvwTitles.SelectedItem Is Nothing Then
        frmGetAPRInfo.GetNewRumorDate lvwTitles.SelectedItem.Text
    Else
        Exit Sub
    End If
        
    NewDate = frmGetAPRInfo.NewDate
    NewName = Trim(frmGetAPRInfo.NewItem)
    
    Unload frmGetAPRInfo
    
    If NewName <> "" Then
    
        Dim NewRumor As RumorClass
        Dim Find As String
        Dim I As Integer
        
        Set NewRumor = New RumorClass
        NewRumor.InitializeQueryRumor NewName, NewDate, rtGeneral
        Game.APREngine.InsertSorted RumorList, NewRumor
        
        mdiMain.AnnounceChanges Me, atRumors
        Game.DataChanged = True
        RefreshRumors
    
        On Error Resume Next
        If CategoryDate Then
            Set lvwTitles.SelectedItem = lvwTitles.ListItems("k" & NewName)
            lvwTitles.SelectedItem.EnsureVisible
            lvwTitles.SetFocus
        Else
            Set lvwDates.SelectedItem = lvwDates.ListItems("k" & CStr(NewDate))
            lvwDates.SelectedItem.EnsureVisible
            lvwDates.SetFocus
        End If
        On Error GoTo 0
        RefreshInfo
        
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
    Dim DelDate As Date
    Dim Answer As Boolean
    
    DelName = ""
    DelDate = 0
    
    If Not (lvwTitles.SelectedItem Is Nothing) Then DelName = lvwTitles.SelectedItem.Text

    If CategoryDate Then
        If lstDates.ListIndex > -1 Then DelDate = CDate(lstDates.Text)
    Else
        If Not (lvwDates.SelectedItem Is Nothing) Then DelDate = CDate(lvwDates.SelectedItem.Text)
    End If
    
    If Not (DelName = "" Or DelDate = 0) Then
    
        Game.APREngine.MoveToPair RumorList, DelDate, DelName
    
        If Not RumorList.Off Then
    
            Answer = ShiftDown
            If Not Answer Then Answer = (MsgBox("This will PERMANENTLY remove this rumor" & _
                    " from the game. Are you sure you want to delete it?", _
                    vbYesNo + vbQuestion, "Delete Rumor") = vbYes)
            If Answer Then
                    
                mdiMain.AnnounceChanges Me, atRumors
                Game.DataChanged = True
    
                For Each NormForm In Forms()
                    If NormForm.Caption = RumorList.Item.Name And NormForm.Tag = "U" Then
                        Unload NormForm
                        Exit For
                    End If
                Next NormForm
                
                RumorList.Remove
                RefreshCategories
                
            End If
        Else
            MsgBox "Grapevine can't find this rumor!  Was it renamed or deleted?", vbExclamation
        End If
        
    End If

End Sub

Private Sub cmdDeleteAll_Click()
'
' Name:         cmdDeleteAll_Click
' Description:  Delete all actions in the current category.
'

    Dim NormForm As Form
    Dim DelName As String
    Dim DelDate As Date
    Dim Answer As Boolean
    
    If CategoryDate And lstDates.ListIndex > -1 Then
        DelDate = CDate(lstDates.Text)
        DelName = lstDates.Text
        Game.APREngine.MoveToFirstDate RumorList, DelDate
    ElseIf Not CategoryDate And Not lvwTitles.SelectedItem Is Nothing Then
        DelName = lvwTitles.SelectedItem.Text
        Game.APREngine.MoveToFirstTitle RumorList, DelName
    Else
        Exit Sub
    End If
    
    If Not RumorList.Off Then

        Answer = (MsgBox("This will PERMANENTLY remove all rumors associated with " & _
                DelName & " from the game. Are you sure you want to delete them?", _
                vbYesNo + vbExclamation, "Delete Rumors") = vbYes)
                
        If Answer Then
                
            mdiMain.AnnounceChanges Me, atRumors
            Game.DataChanged = True

            For Each NormForm In Forms()
                If NormForm.Tag = "U" Then
                    Unload NormForm
                End If
            Next NormForm
            
            Do
                RumorList.Remove
                If CategoryDate Then
                    Game.APREngine.MoveToFirstDate RumorList, DelDate
                Else
                    Game.APREngine.MoveToFirstTitle RumorList, DelName
                End If
            Loop Until RumorList.Off
            
            RefreshCategories
            
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

    Dim ShowDate As Date
    Dim ShowName As String
    
    ShowName = ""
    ShowDate = 0
    
    If Not (lvwTitles.SelectedItem Is Nothing) Then ShowName = lvwTitles.SelectedItem.Text

    If CategoryDate Then
        If lstDates.ListIndex > -1 Then ShowDate = CDate(lstDates.Text)
    Else
        If Not (lvwDates.SelectedItem Is Nothing) Then ShowDate = CDate(lvwDates.SelectedItem.Text)
    End If

    If Not (ShowDate = 0 Or ShowName = "") Then mdiMain.ShowAPR aprRumor, ShowName, ShowDate

End Sub

Private Sub cmdStandard_Click()
'
' Name:         cmdStandard_Click
' Description:  Add the standard rumors for the selected date.
'

    If lstDates.ListIndex > -1 Then
        Game.APREngine.AddStandardRumors CDate(lstDates.Text)
        RefreshRumors
        Game.DataChanged = True
    End If
    
End Sub

Private Sub cmdView_Click()
'
' Name:         cmdView_Click
' Description:  Change the listing to go by the given category.
'

    CategoryDate = Not CategoryDate
    cmdView.Caption = IIf(CategoryDate, "&View by Title", "&View by Date")
    lblCategory.Caption = IIf(CategoryDate, "Dat&es", "Titl&es")
    cmdStandard.Enabled = CategoryDate
    lvwTitles.SortOrder = lvwAscending
    lvwTitles.SortKey = 2
    Call Form_Resize
    RefreshCategories

End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  If the characters have changed, refresh the list.
'

    If mdiMain.CheckForChanges(Me, atRumors) Or mdiMain.CheckForChanges(Me, atDates) Then
        RefreshCategories
    Else
        RefreshInfo
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
    
    CategoryDate = True
    
    Me.Width = FORM_START_WIDTH
    Me.Height = FORM_START_HEIGHT
    
    Set lvwTitles.SmallIcons = mdiMain.imlSmallIcons
    
    RefreshCategories True
    
End Sub

Private Sub Form_Resize()
'
' Name:         Form_Resize
' Description:  Position the controls appropriately on a resized form.
'

    If Me.WindowState <> vbMinimized Then
        
        Dim SH As Integer
        Dim SW As Integer
        Dim CategoryBox As Control
        Dim RumorBox As Control
        
        SH = Me.ScaleHeight
        SW = Me.ScaleWidth
    
        If SH < FORM_MIN_SCALEHEIGHT Then SH = FORM_MIN_SCALEHEIGHT
        If SW < FORM_MIN_SCALEWIDTH Then SW = FORM_MIN_SCALEWIDTH
    
        fraRight.Left = SW - RIGHT_MARGIN
        
        If CategoryDate Then
            
            lstDates.Visible = True
            lvwDates.Visible = False
        
            lvwTitles.HideColumnHeaders = False
            lvwTitles.ColumnHeaders(2).Width = DONE_WIDTH
            lvwTitles.ColumnHeaders(3).Width = 0
        
            Set CategoryBox = lstDates
            Set RumorBox = lvwTitles
        
        Else
            
            lstDates.Visible = False
            lvwDates.Visible = True
        
            lvwTitles.HideColumnHeaders = True
            lvwTitles.ColumnHeaders(2).Width = 0
            lvwTitles.ColumnHeaders(3).Width = 0
            
            Set CategoryBox = lvwTitles
            Set RumorBox = lvwDates
        
        End If
                
        RumorBox.Left = RUMOR_LEFT
        RumorBox.Width = SW - RUMOR_LEFT - RIGHT_MARGIN - VERTICAL_GAP
    
        CategoryBox.Left = CATEGORY_LEFT
        CategoryBox.Width = CATEGORY_WIDTH
                
        RumorBox.Height = SH - TOP_MARGIN - BOTTOM_MARGIN
        CategoryBox.Height = RumorBox.Height - cmdView.Height - BUTTON_GAP
        
        If CategoryDate Then
            lvwTitles.ColumnHeaders(1).Width = lvwTitles.Width - DONE_WIDTH - SCROLL_WIDTH
        Else
            lvwTitles.ColumnHeaders(1).Width = lvwTitles.Width - SCROLL_WIDTH
        End If
        
        cmdView.Top = CategoryBox.Top + CategoryBox.Height + BUTTON_GAP
        lblCount.Width = RumorBox.Width

    End If

End Sub

Private Sub lstDates_Click()
'
' Name:         lstDates_Click
' Description:  Refresh the rumors as needed.
'

    RefreshRumors

End Sub

Private Sub lstDates_DblClick()
'
' Name:         lstDates_DblClick
' Description:  Shortcut to Add or Show, as needed
'

    Call cmdAddNew_Click
    
End Sub

Private Sub lvwDates_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'
' Name:         lvwDates_ColumnClick
' Description:  Sort the list according to the column clicked.
'

    If lvwDates.SortKey + ColumnHeader.Index = 3 Then     ' same header clicked: reverse
        lvwDates.SortOrder = IIf(lvwDates.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwDates.SortKey = IIf(lvwDates.SortKey = 2, 1, 2)
    End If
    If Not lvwDates.SelectedItem Is Nothing Then lvwDates.SelectedItem.EnsureVisible
    
End Sub

Private Sub lvwDates_DblClick()
'
' Name:         lvwDates_DblClick
' Description:  Shortcut to Show
'
    
    Call cmdShow_Click

End Sub

Private Sub lvwDates_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
' Name:         lvwDates_ItemClick
' Description:  Refresh the info
'
    RefreshInfo
    
End Sub

Private Sub lvwTitles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'
' Name:         lvwTitles_ColumnClick
' Description:  Sort the list according to the column clicked.
'

    If lvwTitles.SortKey + ColumnHeader.Index = 3 Then     ' same header clicked: reverse
        lvwTitles.SortOrder = IIf(lvwTitles.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwTitles.SortKey = IIf(lvwTitles.SortKey = 2, 1, 2)
    End If
    If Not lvwTitles.SelectedItem Is Nothing Then lvwTitles.SelectedItem.EnsureVisible

End Sub

Private Sub lvwTitles_DblClick()
'
' Name:         lvwRumors_DblClick
' Description:  Shortcut to Add or Show, as needed
'
    
    If CategoryDate Then
        Call cmdShow_Click
    Else
        Call cmdAddNew_Click
    End If
    
End Sub

Private Sub lvwTitles_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
' Name:         lvwTitles_ItemClick
' Description:  Refresh the rumors or info as needed.
'

    If CategoryDate Then
        RefreshInfo
    Else
        RefreshRumors
    End If

End Sub
