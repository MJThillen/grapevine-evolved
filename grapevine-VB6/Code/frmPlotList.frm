VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlotList 
   Caption         =   "Plots"
   ClientHeight    =   5790
   ClientLeft      =   1875
   ClientTop       =   750
   ClientWidth     =   7575
   Icon            =   "frmPlotList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   7575
   Begin VB.CommandButton cmdView 
      Caption         =   "&View by Title"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   5160
      Width           =   1695
   End
   Begin MSComctlLib.ListView lvwTitles 
      Height          =   4620
      Left            =   2040
      TabIndex        =   13
      Top             =   450
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   8149
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "plot"
         Text            =   "Plot"
         Object.Width           =   4419
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   "status"
         Text            =   "Status"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   "done"
         Text            =   "Done"
         Object.Width           =   1005
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
      Begin VB.CommandButton cmdDevActive 
         Caption         =   "Develop A&ctive Plots"
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   3720
         Width           =   1725
      End
      Begin VB.CommandButton cmdDeletePlot 
         Caption         =   "D&elete Plot"
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   4200
         Width           =   1725
      End
      Begin VB.CommandButton cmdAddPlot 
         Caption         =   "&Add New Plot"
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   3240
         Width           =   1725
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Show Plot"
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   2280
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
         Picture         =   "frmPlotList.frx":058A
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin MSComctlLib.ListView lvwDates 
      Height          =   4620
      Left            =   3480
      TabIndex        =   14
      Top             =   450
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   8149
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
      Caption         =   "0 Da&tes"
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
      Caption         =   "0 Develop&ments"
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
Attribute VB_Name = "frmPlotList"
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
Private Const FORM_START_WIDTH = 9150
Private Const FORM_MIN_SCALEHEIGHT = 5280
Private Const FORM_MIN_SCALEWIDTH = 5910
Private Const BOTTOM_MARGIN = 705
Private Const RIGHT_MARGIN = 1980
Private Const VERTICAL_GAP1 = 120
Private Const VERTICAL_GAP2 = 225
Private Const BUTTON_GAP = 120

Private Const COL1_MIN_WIDTH = 1255
Private Const STATUS_WIDTH = 800
Private Const DONE_WIDTH = 570
Private Const SCROLL_WIDTH = 300

Private Const OutlineItem = "List All Plots"

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
            Game.Calendar.MoveToCloseGame
            If Not Game.Calendar.Off Then SelDate = Game.Calendar.GetGameDate
        End If
    
        StoreCategory = lstDates.ListIndex
        lstDates.Clear
        
        lstDates.AddItem OutlineItem
        
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
                                                 'triggers RefreshDevelopments
    
        lblCategory.Caption = "Da&tes"
    
    Else
    
        Dim NewItem As ListItem
        Dim Plot As PlotClass
        Dim Status As String
        
        StoreCategory = 1
        If Not lvwTitles.SelectedItem Is Nothing Then _
            StoreCategory = lvwTitles.SelectedItem.Index
        lvwTitles.ListItems.Clear
        
        PlotList.First
        Do Until PlotList.Off
            Set Plot = PlotList.Item
            Set NewItem = lvwTitles.ListItems.Add(Key:="k" & Plot.Name, Text:=Plot.Name)
            Call NewItem.ListSubItems.Add(Text:=Plot.GetStatus)
            PlotList.MoveNext
        Loop

        On Error Resume Next
        If StoreCategory > lvwTitles.ListItems.Count Then StoreCategory = 1
        Set lvwTitles.SelectedItem = lvwTitles.ListItems(StoreCategory)
        On Error GoTo 0
        
        lblCategory.Caption = CStr(lvwTitles.ListItems.Count) & " Plo&ts"
        
        RefreshDevelopments
        RefreshInfo

    End If
        
    Screen.MousePointer = vbDefault

End Sub

Private Sub RefreshDevelopments()
'
' Name:         RefreshRumors
' Description:  Refresh the list of Rumors based on the current date/character.
'

    Dim StorePlot As Integer
    Dim NewItem As ListItem
    Dim Plot As PlotClass
    Dim CurDate As Date
    Dim SelDate As Date
    
    StorePlot = 1
    lblCount = "0 Develop&ments"
    
    If CategoryDate Then
    
        If Not lvwTitles.SelectedItem Is Nothing Then _
            StorePlot = lvwTitles.SelectedItem.Index
        lvwTitles.ListItems.Clear
        lblCount.Caption = "0 Develop&ments"
        
        If lstDates.Text = OutlineItem Then
        
            lvwTitles.ColumnHeaders(1).Width = lvwTitles.Width - STATUS_WIDTH - SCROLL_WIDTH
            lvwTitles.ColumnHeaders(2).Width = STATUS_WIDTH
            lvwTitles.ColumnHeaders(3).Width = 0
        
            PlotList.First
            Do Until PlotList.Off
                Set Plot = PlotList.Item
                Set NewItem = lvwTitles.ListItems.Add(Key:="k" & Plot.Name, Text:=Plot.Name)
                Call NewItem.ListSubItems.Add(Text:=Plot.GetStatus(CurDate))
                Call NewItem.ListSubItems.Add(Text:="")
                PlotList.MoveNext
            Loop
        
            lblCount.Caption = CStr(lvwTitles.ListItems.Count) & " Plots"
            
        ElseIf Not lstDates.Text = "" Then
        
            CurDate = CDate(lstDates.Text)
            
            lvwTitles.ColumnHeaders(1).Width = lvwTitles.Width - DONE_WIDTH - SCROLL_WIDTH
            lvwTitles.ColumnHeaders(2).Width = 0
            lvwTitles.ColumnHeaders(3).Width = DONE_WIDTH
                            
            PlotList.First
            Do Until PlotList.Off
                Set Plot = PlotList.Item
                Plot.MoveTo CurDate
                If Not Plot.Off Then
                    Set NewItem = lvwTitles.ListItems.Add(Key:="k" & Plot.Name, Text:=Plot.Name)
                    Call NewItem.ListSubItems.Add(Text:=Plot.GetStatus(CurDate))
                    Call NewItem.ListSubItems.Add(Text:=IIf(Plot.PlotDev.IsComplete, "X", ""))
                End If
                PlotList.MoveNext
            Loop
        
            lblCount.Caption = CStr(lvwTitles.ListItems.Count) & " Develop&ments for " & lstDates.Text
        
        End If
        
        On Error Resume Next
        If StorePlot >= lvwTitles.ListItems.Count Then StorePlot = 1
        Set lvwTitles.SelectedItem = lvwTitles.ListItems(StorePlot)
        On Error GoTo 0
        
        RefreshInfo
        
    Else

        If Not lvwDates.SelectedItem Is Nothing Then _
            SelDate = CDate(lvwDates.SelectedItem.Index)
        lvwDates.ListItems.Clear
        
        If Not lvwTitles.SelectedItem Is Nothing Then
            
            If lvwTitles.Tag <> lvwTitles.SelectedItem.Text Then
                Game.Calendar.MoveToCloseGame
                If Not Game.Calendar.Off Then SelDate = Game.Calendar.GetGameDate
            End If
            
            lvwTitles.Tag = lvwTitles.SelectedItem.Text
            PlotList.MoveTo lvwTitles.SelectedItem.Text
            
            If Not PlotList.Off Then
                
                Set Plot = PlotList.Item
                Plot.First
                Do Until Plot.Off
                    CurDate = Plot.PlotDev.DevDate
                    Set NewItem = lvwDates.ListItems.Add(Text:=Format(CurDate, "mmmm d, yyyy"), _
                           Key:="k" & CStr(CurDate))
                    Call NewItem.ListSubItems.Add(Text:=IIf(Plot.PlotDev.IsComplete, "X", ""), Key:="done")
                    Call NewItem.ListSubItems.Add(Text:=Format(CurDate, "yyyy-mm-dd"), Key:="sortkey")
                    Plot.MoveNext
                Loop
                
                lblCount.Caption = CStr(lvwDates.ListItems.Count) & " Develop&ments"
                lvwDates.ColumnHeaders(1).Text = Plot.Name
                
            End If
                        
            On Error Resume Next
            Set lvwDates.SelectedItem = lvwDates.ListItems("k" & CStr(SelDate))
            On Error GoTo 0
            
        End If
        
    End If

End Sub

Private Sub RefreshInfo()
'
' Name:         RefreshInfo
' Description:  Find the plot and display the appropriate information at right.
'

    If Not (lvwTitles.SelectedItem Is Nothing) Then
        
        PlotList.MoveTo lvwTitles.SelectedItem.Text
        If Not PlotList.Off Then
            
            lblName.Caption = PlotList.Item.Name
            lblDate.Caption = Format(PlotList.Item.LastModified, "Short Date")
            imgIcon.Visible = True
            
        Else
            MsgBox "Grapevine can't find this plot!  Was it renamed or deleted?", vbExclamation
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
    
    With OutputEngine
        .Template = tnPlot
        .GameDate = 0
        .SelectSet(osPlots).Clear
        If CategoryDate Then
            If IsDate(lstDates.Text) Then .GameDate = CDate(lstDates.Text)
            .SelectSet(osPlots).StoreListView lvwTitles, True
        ElseIf Not lvwTitles.SelectedItem Is Nothing Then
            .SelectSet(osPlots).Add lvwTitles.SelectedItem.Text
        End If
    End With
    
End Sub

Private Sub cmdAddPlot_Click()
'
' Name:         cmdAddPlot_Click
' Description:  Creates a new plot and adds it to the linked list and the
'               list box, selecting it.
'

    Dim NewPlot As PlotClass
    Dim NewName As String
    
    Do
        NewName = InputBox("Enter a name for the plot:", "New Plot")
        NewName = Trim(NewName)
        If NewName = "" Then Exit Do
        PlotList.MoveTo NewName
    Loop Until PlotList.Off
    
    If NewName <> "" Then
    
        Set NewPlot = New PlotClass
        NewPlot.Name = NewName
        PlotList.InsertSorted NewPlot
        
        mdiMain.AnnounceChanges Me, atPlots
        Game.DataChanged = True
        
        If CategoryDate And Not lstDates.ListIndex = 0 Then
            lstDates.ListIndex = 0
        Else
            RefreshCategories
        End If
        
        On Error Resume Next
        Set lvwTitles.SelectedItem = lvwTitles.ListItems("k" & NewName)
        lvwTitles.SelectedItem.EnsureVisible
        On Error GoTo 0
        
        RefreshInfo
        lvwTitles.SetFocus

    End If

End Sub

Private Sub cmdDeletePlot_Click()
'
' Name:         cmdDeleteAll_Click
' Description:  Delete a plot completely.
'

    Dim NormForm As Form
    Dim DelName As String
    Dim Answer As Boolean
    
    If Not lvwTitles.SelectedItem Is Nothing Then
    
        DelName = lvwTitles.SelectedItem.Text
        PlotList.MoveTo DelName
        
        If Not PlotList.Off Then
        
            Answer = ShiftDown
            If Not Answer Then Answer = (MsgBox("This will PERMANENTLY remove the plot " & _
                    DelName & " from the game. Are you sure you want to delete it?", _
                    vbYesNo + vbExclamation, "Delete Plot") = vbYes)
                    
            If Answer Then
                    
                mdiMain.AnnounceChanges Me, atPlots
                Game.DataChanged = True
    
                For Each NormForm In Forms()
                    If NormForm.Tag = "P" And NormForm.Caption = DelName Then
                        Unload NormForm
                        Exit Sub
                    End If
                Next NormForm
                
                PlotList.Remove
                
            End If
    
        End If
    
        RefreshCategories
        
    End If

End Sub

Private Sub cmdDevActive_Click()
'
' Name:         cmdDevActive_Click
' Description:  Add developments for all plots considered Active.
'

    Dim DevDate As Date
    Dim Plot As PlotClass
    
    DevDate = 0
    
    If CategoryDate And lstDates.ListIndex > 0 Then
        DevDate = CDate(lstDates.Text)
    Else
        Game.Calendar.MoveToCloseGame
        If Not Game.Calendar.Off Then DevDate = Game.Calendar.GetGameDate
    End If

    If Not DevDate = 0 Then

        PlotList.First
        Do Until PlotList.Off
    
            Set Plot = PlotList.Item
    
            If Plot.GetStatus(DevDate) = psActive Then
                Plot.MoveTo DevDate
                If Plot.Off Then
                    Plot.Add DevDate, ""
                End If
            End If
            
            PlotList.MoveNext
        Loop
        
        mdiMain.AnnounceChanges Me, atPlots
        RefreshDevelopments

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
' Description:  Asks the parent form to display the selected plot.
'

    If Not (lvwTitles.SelectedItem Is Nothing) Then

        Dim ShowDate As Date
        
        ShowDate = 0
        
        If CategoryDate Then
            If lstDates.ListIndex > 0 Then ShowDate = CDate(lstDates.Text)
        Else
            If Not (lvwDates.SelectedItem Is Nothing) Then ShowDate = CDate(lvwDates.SelectedItem.Text)
        End If
    
        mdiMain.ShowAPR aprPlot, lvwTitles.SelectedItem.Text, ShowDate

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
    lvwTitles.SortOrder = lvwAscending
    lvwTitles.SortKey = 1
    Call Form_Resize
    RefreshCategories

End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  If the data has changed, refresh the list.
'

    If mdiMain.CheckForChanges(Me, atPlots) Or mdiMain.CheckForChanges(Me, atDates) Then
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
    If KeyCode = vbKeyDelete Then Call cmdDeletePlot_Click

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Fill the list and select the first character.
'
    
    CategoryDate = True
    
    Me.Width = FORM_START_WIDTH
    Me.Height = FORM_START_HEIGHT
    
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
        
            lstDates.Height = SH - lstDates.Top - BOTTOM_MARGIN
        
            lvwTitles.Left = lstDates.Left + lstDates.Width + VERTICAL_GAP1
            lvwTitles.Width = SW - lvwTitles.Left - RIGHT_MARGIN - VERTICAL_GAP2
        
            lvwTitles.ColumnHeaders(1).Width = lvwTitles.Width - DONE_WIDTH - SCROLL_WIDTH
            lvwTitles.ColumnHeaders(2).Width = 0
            lvwTitles.ColumnHeaders(3).Width = DONE_WIDTH
                
            lblCategory.Width = lstDates.Width
            lblCount.Left = lvwTitles.Left
            lblCount.Width = lvwTitles.Width
            
        Else
            
            Dim TW As Integer
            Dim DW As Integer
            
            TW = SW - lstDates.Left - RIGHT_MARGIN - VERTICAL_GAP2 - VERTICAL_GAP1
            DW = TW * 0.4
            TW = TW * 0.6
            
            lstDates.Visible = False
            lvwDates.Visible = True
            
            lvwTitles.Left = lstDates.Left
            lvwTitles.Width = TW
        
            lvwDates.Left = lvwTitles.Left + TW + VERTICAL_GAP1
            lvwDates.Width = DW
            lvwDates.Height = SH - lvwDates.Top - BOTTOM_MARGIN

            lblCategory.Width = TW
            lblCount.Left = lvwDates.Left
            lblCount.Width = DW

            TW = TW - STATUS_WIDTH - SCROLL_WIDTH
            If TW < COL1_MIN_WIDTH Then TW = COL1_MIN_WIDTH
            lvwTitles.ColumnHeaders(1).Width = TW
            lvwTitles.ColumnHeaders(2).Width = STATUS_WIDTH
            lvwTitles.ColumnHeaders(3).Width = 0
            
            DW = DW - DONE_WIDTH - SCROLL_WIDTH
            If DW < COL1_MIN_WIDTH Then DW = COL1_MIN_WIDTH
            lvwDates.ColumnHeaders(1).Width = DW
            lvwDates.ColumnHeaders(2).Width = DONE_WIDTH
            lvwDates.ColumnHeaders(3).Width = 0
        
        End If
        
        lvwTitles.Height = SH - lvwTitles.Top - BOTTOM_MARGIN
        
        cmdView.Top = SH - BOTTOM_MARGIN + BUTTON_GAP
                
    End If

End Sub

Private Sub lstDates_Click()
'
' Name:         lstDates_Click
' Description:  Refresh the rumors as needed.
'

    RefreshDevelopments

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

    If lvwTitles.SortKey = ColumnHeader.Index - 1 Then    ' same header clicked: reverse
        lvwTitles.SortOrder = IIf(lvwTitles.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwTitles.SortKey = ColumnHeader.Index - 1
    End If
    If Not lvwTitles.SelectedItem Is Nothing Then lvwTitles.SelectedItem.EnsureVisible

End Sub

Private Sub lvwTitles_DblClick()
'
' Name:         lvwRumors_DblClick
' Description:  Shortcut to Add or Show, as needed
'
    
    Call cmdShow_Click
    
End Sub

Private Sub lvwTitles_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
' Name:         lvwTitles_ItemClick
' Description:  Refresh the rumors or info as needed.
'

    If CategoryDate Then
        RefreshInfo
    Else
        RefreshDevelopments
        Set lvwDates.SelectedItem = Nothing
    End If

End Sub
