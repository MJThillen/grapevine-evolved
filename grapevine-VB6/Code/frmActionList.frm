VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmActionList 
   Caption         =   "Actions"
   ClientHeight    =   5775
   ClientLeft      =   1875
   ClientTop       =   750
   ClientWidth     =   7575
   Icon            =   "frmActionList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   7575
   Begin VB.CommandButton cmdView 
      Caption         =   "&View by Character"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   5160
      Width           =   1695
   End
   Begin MSComctlLib.ListView lvwActions 
      Height          =   5100
      Left            =   2040
      TabIndex        =   1
      Top             =   465
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
         Key             =   "char"
         Text            =   "Character"
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
   Begin VB.ListBox lstCategory 
      Height          =   4575
      IntegralHeight  =   0   'False
      Left            =   240
      TabIndex        =   13
      Top             =   480
      Width           =   1695
   End
   Begin VB.Frame fraRight 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   5640
      TabIndex        =   2
      Top             =   480
      Width           =   1695
      Begin VB.CommandButton cmdDeleteAll 
         Caption         =   "Delete A&ll Actions"
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   4680
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "&Add New Action"
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   3720
         Width           =   1695
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Show Action"
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Action"
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Last Modified:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
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
         TabIndex        =   6
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   607
         Picture         =   "frmActionList.frx":058A
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
         Height          =   885
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Index           =   2
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1695
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
      TabIndex        =   12
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      Caption         =   "0 A&ctions"
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
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmActionList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' Name:         frmLocationList
' Description:  Manage the game's list of Locations.
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
Private Const COLUMN_MARGIN = 870

Private Sub RefreshList(Optional SelGame As Boolean = False)
'
' Name:         RefreshList
' Parameter:    SelGame         if TRUE, select the next/last game;
'                               else preserve current selection
' Description:  Refill the category list with dates/characters as appropriate.
'

    Dim StoreCategory As Integer
    Dim SelDate As Date
    
    Screen.MousePointer = vbHourglass
    StoreCategory = lstCategory.ListIndex
    
    lstCategory.Clear
    
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
        
        Game.Calendar.Last
        Do Until Game.Calendar.Off
            lstCategory.AddItem Format(Game.Calendar.GetGameDate, "mmmm d, yyyy")
            If SelGame Then
                If SelDate = Game.Calendar.GetGameDate Then
                    StoreCategory = lstCategory.NewIndex
                End If
            End If
            Game.Calendar.MovePrevious
        Loop
    
    Else
    
        Dim CharName As String
        Dim I As Integer
        Dim CatSet As StringSet
        
        Set CatSet = New StringSet
        
        ActionList.First
        Do Until ActionList.Off
            CharName = ActionList.Item.CharName
            If Not CatSet.Has(CharName) Then
                CatSet.Add CharName
                I = 0
                Do Until I = lstCategory.ListCount
                    If lstCategory.List(I) > CharName Then Exit Do
                    I = I + 1
                Loop
                lstCategory.AddItem CharName, I
            End If
            ActionList.MoveNext
        Loop
    
        Set CatSet = Nothing

    End If
    
    If StoreCategory >= lstCategory.ListCount Then StoreCategory = lstCategory.ListCount - 1
    If StoreCategory = -1 And lstCategory.ListCount > 0 Then StoreCategory = 0
    lstCategory.ListIndex = StoreCategory       'this triggers lstCategory_Click
    If StoreCategory = -1 Then Call lstCategory_Click
    
    Screen.MousePointer = vbDefault

End Sub

Public Sub SetDefaultOutput()
'
' Name:         SetDefaultOutput
' Description:  Initilize the OutputEngineClass with default output settings.
'
    
    If Not lstCategory.Text = "" Then
        If CategoryDate Then
            Game.APREngine.MoveToFirstDate ActionList, CDate(lstCategory.Text)
        Else
            Game.APREngine.MoveToFirstTitle ActionList, lstCategory.Text
        End If
        With OutputEngine
            .SelectSet(osActions).Clear
            Do Until ActionList.Off
                .SelectSet(osActions).Add ActionList.Item.Name
                If CategoryDate Then
                    Game.APREngine.MoveToNextDate ActionList, CDate(lstCategory.Text)
                Else
                    Game.APREngine.MoveToNextTitle ActionList, lstCategory.Text
                End If
            Loop
            .Template = tnMasterAction
            .GameDate = 0
        End With
    End If
    
End Sub

Private Sub cmdAddNew_Click()
'
' Name:         cmdAddNew_Click
' Description:  Calls on frmGetAPRInfo to display itself and return a date and
'               and name to create an action for.
'
    
    If lstCategory.ListIndex > -1 Then
    
        Dim NewDate As Date
        Dim NewName As String
        
        If CategoryDate Then
            frmGetAPRInfo.GetNewActionChar CDate(lstCategory.Text)
        Else
            frmGetAPRInfo.GetNewActionDate lstCategory.Text
        End If
        
        NewDate = frmGetAPRInfo.NewDate
        NewName = Trim(frmGetAPRInfo.NewItem)
        Unload frmGetAPRInfo
        
        If NewName <> "" Then
        
            Dim NewAction As ActionClass
            Dim Find As String
            Dim I As Integer
            
            Set NewAction = New ActionClass
            NewAction.Initialize NewName, NewDate
            Game.APREngine.InsertSorted ActionList, NewAction
            
            mdiMain.AnnounceChanges Me, atActions
            Game.DataChanged = True
            Call lstCategory_Click
        
            On Error Resume Next
            Set lvwActions.SelectedItem = lvwActions.ListItems("k" & _
                                          IIf(CategoryDate, NewName, CStr(NewDate)))
            On Error GoTo 0
            lvwActions.SetFocus
            
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
    Dim DelDate As Date
    Dim Answer As Boolean
    
    If Not (lvwActions.SelectedItem Is Nothing) And lstCategory.ListIndex > -1 Then
    
        If CategoryDate Then
            DelDate = CDate(lstCategory.Text)
            DelName = lvwActions.SelectedItem.Text
        Else
            DelDate = CDate(lvwActions.SelectedItem.Text)
            DelName = lstCategory.Text
        End If
        
        Game.APREngine.MoveToPair ActionList, DelDate, DelName
    
        If Not ActionList.Off Then
    
            Answer = ShiftDown
            If Not Answer Then Answer = (MsgBox("This will PERMANENTLY remove this action" & _
                    " from the game. Are you sure you want to delete it?", _
                    vbYesNo + vbQuestion, "Delete Action") = vbYes)
            If Answer Then
                    
                mdiMain.AnnounceChanges Me, atActions
                Game.DataChanged = True
    
                For Each NormForm In Forms()
                    If NormForm.Caption = ActionList.Item.Name And NormForm.Tag = "A" Then
                        Unload NormForm
                        Exit For
                    End If
                Next NormForm
                
                ActionList.Remove
                RefreshList
                
            End If
        Else
            MsgBox "Grapevine can't find this action!  Was it renamed or deleted?", vbExclamation
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
    
    If CategoryDate Then
        If Not IsDate(lstCategory.Text) Then Exit Sub
        DelDate = CDate(lstCategory.Text)
        Game.APREngine.MoveToFirstDate ActionList, DelDate
    Else
        DelName = lstCategory.Text
        Game.APREngine.MoveToFirstTitle ActionList, DelName
    End If
        
    If Not ActionList.Off Then

        Answer = (MsgBox("This will PERMANENTLY remove all actions associated with " & _
                lstCategory.Text & " from the game. Are you sure you want to delete them?", _
                vbYesNo + vbExclamation, "Delete Actions") = vbYes)
        If Answer Then
                
            mdiMain.AnnounceChanges Me, atActions
            Game.DataChanged = True

            For Each NormForm In Forms()
                If NormForm.Tag = "A" Then
                    Unload NormForm
                End If
            Next NormForm
            
            Do
                ActionList.Remove
                If CategoryDate Then
                    Game.APREngine.MoveToFirstDate ActionList, DelDate
                Else
                    Game.APREngine.MoveToFirstTitle ActionList, DelName
                End If
            Loop Until ActionList.Off
            RefreshList
            
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

    If Not (lvwActions.SelectedItem Is Nothing) And lstCategory.ListIndex > -1 Then
        
        Dim ShowDate As Date
        Dim ShowName As String
        
        If CategoryDate Then
            ShowDate = CDate(lstCategory.Text)
            ShowName = lvwActions.SelectedItem.Text
        Else
            ShowDate = CDate(lvwActions.SelectedItem.Text)
            ShowName = lstCategory.Text
        End If
        
        mdiMain.ShowAPR aprAction, ShowName, ShowDate
        
    End If

End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  If the characters have changed, refresh the list.
'

    If mdiMain.CheckForChanges(Me, atActions) Or mdiMain.CheckForChanges(Me, atDates) Then
        RefreshList
    Else
        Call lvwActions_ItemClick(lvwActions.SelectedItem)
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
    
    RefreshList True
    
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
    
        fraRight.Left = SW - RIGHT_MARGIN
        lvwActions.Height = SH - lvwActions.Top - BOTTOM_MARGIN
        lvwActions.Width = SW - lvwActions.Left - RIGHT_MARGIN - VERTICAL_GAP
        lvwActions.ColumnHeaders(1).Width = lvwActions.Width - COLUMN_MARGIN
        lstCategory.Height = lvwActions.Height - cmdView.Height - BUTTON_GAP
        cmdView.Top = lvwActions.Top + lvwActions.Height - cmdView.Height
        lblCount.Width = lvwActions.Width

    End If

End Sub

Private Sub lstCategory_Click()
'
' Name:         lstCategory_Click
' Description:  Refresh the list of actions when a new date/character is chosen.
'

    Dim StoreAction As Integer
    Dim NewItem As ListItem
    
    StoreAction = 1
    If Not lvwActions.SelectedItem Is Nothing Then StoreAction = lvwActions.SelectedItem.Index
    lvwActions.ListItems.Clear
    lvwActions.SortKey = 2
    
    If Not lstCategory.Text = "" Then
        If CategoryDate Then
        
            Dim SelDate As Date
            Dim CurName As String
            
            lvwActions.SortOrder = lvwAscending
            SelDate = CDate(lstCategory.Text)
            Game.APREngine.MoveToFirstDate ActionList, SelDate
            
            Do Until ActionList.Off
                CurName = ActionList.Item.CharName
                
                Set NewItem = lvwActions.ListItems.Add(Text:=CurName, Key:="k" & CurName)
                Call NewItem.ListSubItems.Add(Text:=IIf(ActionList.Item.Done, "X", ""), Key:="done")
                Call NewItem.ListSubItems.Add(Text:=CurName, Key:="sortkey")
                
                Game.APREngine.MoveToNextDate ActionList, SelDate
            Loop
            
        Else
        
            Dim SelName As String
            Dim CurDate As Date
            
            lvwActions.SortOrder = lvwDescending
            SelName = lstCategory.Text
            Game.APREngine.MoveToFirstTitle ActionList, SelName
            
            Do Until ActionList.Off
                CurDate = ActionList.Item.ActDate
                
                Set NewItem = lvwActions.ListItems.Add(Text:=Format(CurDate, "mmmm d, yyyy"), _
                                                       Key:="k" & CStr(CurDate))
                Call NewItem.ListSubItems.Add(Text:=IIf(ActionList.Item.Done, "X", ""), Key:="done")
                Call NewItem.ListSubItems.Add(Text:=Format(CurDate, "yyyy-mm-dd"), Key:="sortkey")
                
                Game.APREngine.MoveToNextTitle ActionList, SelName
            Loop
            
        End If
    End If
    
    If StoreAction > lvwActions.ListItems.Count Then StoreAction = lvwActions.ListItems.Count
    If StoreAction > 0 Then
        Set lvwActions.SelectedItem = lvwActions.ListItems(StoreAction)
        lvwActions.SelectedItem.EnsureVisible
    End If
    Call lvwActions_ItemClick(lvwActions.SelectedItem)
    
    lblCount.Caption = CStr(lvwActions.ListItems.Count) & " A&ctions for " & lstCategory.Text

End Sub

Private Sub lstCategory_DblClick()
'
' Name:         lstCategory_DblClick
' Description:  Shortcut to cmdAddNew_Click.
'
    Call cmdAddNew_Click
    
End Sub

Private Sub lvwActions_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'
' Name:         lvwActions_ColumnClick
' Description:  Sort the list according to the column clicked.
'

    If lvwActions.SortKey + ColumnHeader.Index = 3 Then     ' same header clicked: reverse
        lvwActions.SortOrder = IIf(lvwActions.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwActions.SortKey = IIf(lvwActions.SortKey = 2, 1, 2)
    End If
    If Not lvwActions.SelectedItem Is Nothing Then lvwActions.SelectedItem.EnsureVisible
    
End Sub

Private Sub lvwActions_DblClick()
'
' Name:         lvwActions_DblClick
' Description:  Shortcut to cmdShow_Click
'
    Call cmdShow_Click

End Sub

Private Sub lvwActions_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
' Name:         lvwActions_ItemClick
' Description:  Find the character and display the appropriate information at right.
'

    If Not (Item Is Nothing) And lstCategory.ListIndex > -1 Then
        
        Dim ShowDate As Date
        Dim ShowName As String
        
        If CategoryDate Then
            ShowDate = CDate(lstCategory.Text)
            ShowName = Item.Text
        Else
            ShowDate = CDate(Item.Text)
            ShowName = lstCategory.Text
        End If
        
        Game.APREngine.MoveToPair ActionList, ShowDate, ShowName
        If Not ActionList.Off Then
            lblName.Caption = ActionList.Item.Name
            lblDate.Caption = Format(ActionList.Item.LastModified, "Short Date")
        Else
            MsgBox "Grapevine can't find this action!  Was it renamed or deleted?", vbExclamation
        End If
    
    Else
        lblName.Caption = ""
        lblDate.Caption = ""
    End If

End Sub

Private Sub cmdView_Click()
'
' Name:         cmdView_Click
' Description:  Change the listing to go by the given category.
'

    CategoryDate = Not CategoryDate
    cmdView.Caption = IIf(CategoryDate, "&View by Character", "&View by Date")
    lblCategory.Caption = IIf(CategoryDate, "Dat&es", "Charact&ers")
    lvwActions.ColumnHeaders(1).Text = IIf(CategoryDate, "Character", "Date")
    RefreshList

End Sub
