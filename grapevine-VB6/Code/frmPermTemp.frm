VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmPermTemp 
   Caption         =   "Permanent/Temporary Ratings Management"
   ClientHeight    =   6165
   ClientLeft      =   1875
   ClientTop       =   750
   ClientWidth     =   9030
   Icon            =   "frmPermTemp.frx":0000
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
      TabIndex        =   23
      Top             =   5400
      Width           =   4815
      Begin VB.ComboBox cboSearch 
         Height          =   315
         Left            =   2685
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   75
         Width           =   2055
      End
      Begin VB.OptionButton optNot 
         Caption         =   "List characters that &don't match:"
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   25
         Top             =   240
         Width           =   2670
      End
      Begin VB.OptionButton optNot 
         Caption         =   "List characters that &match:"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Value           =   -1  'True
         Width           =   2670
      End
   End
   Begin VB.Frame fraRight 
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   6720
      TabIndex        =   2
      Top             =   240
      Width           =   2055
      Begin VB.Frame fraCharacter 
         Height          =   2295
         Left            =   -120
         TabIndex        =   13
         Top             =   3480
         Visible         =   0   'False
         Width           =   2295
         Begin VB.TextBox txtChar 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   17
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txtChar 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   20
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton cmdShow 
            Caption         =   "&Show Character"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   1680
            Width           =   2055
         End
         Begin ComCtl2.UpDown updChar 
            Height          =   285
            Index           =   1
            Left            =   1575
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            _Version        =   327681
            BuddyControl    =   "txtChar(1)"
            BuddyDispid     =   196614
            BuddyIndex      =   1
            OrigLeft        =   1440
            OrigTop         =   4680
            OrigRight       =   1935
            OrigBottom      =   4965
            Max             =   9999
            Orientation     =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown updChar 
            Height          =   285
            Index           =   0
            Left            =   1575
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   840
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            _Version        =   327681
            BuddyControl    =   "txtChar(0)"
            BuddyDispid     =   196614
            BuddyIndex      =   0
            OrigLeft        =   1440
            OrigTop         =   4320
            OrigRight       =   1935
            OrigBottom      =   4605
            Max             =   9999
            Orientation     =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lblTemper 
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
            TabIndex        =   15
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "&Temporary"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   19
            Top             =   1245
            Width           =   855
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "&Permanent"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   16
            Top             =   885
            Width           =   855
         End
         Begin VB.Label lblName 
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
            TabIndex        =   14
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "&Apply the Change"
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CheckBox chkCap 
         Caption         =   "&Go no higher than the permanent rating"
         Height          =   495
         Left            =   0
         TabIndex        =   11
         Top             =   2400
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkRandomly 
         Caption         =   "&Randomly"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   2160
         Width           =   1095
      End
      Begin ComCtl2.UpDown updChange 
         Height          =   285
         Left            =   975
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1800
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         _Version        =   327681
         OrigLeft        =   960
         OrigTop         =   1800
         OrigRight       =   1455
         OrigBottom      =   2085
         Max             =   9999
         Min             =   -9999
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtChange 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   480
         TabIndex        =   8
         Text            =   "1"
         Top             =   1800
         Width           =   495
      End
      Begin VB.ComboBox cboAct 
         Height          =   315
         ItemData        =   "frmPermTemp.frx":058A
         Left            =   0
         List            =   "frmPermTemp.frx":059A
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   2055
      End
      Begin VB.ComboBox cboTemper 
         Height          =   315
         Left            =   0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblGuide 
         Alignment       =   2  'Center
         Caption         =   "&its temporary rating to"
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   1470
         Width           =   2055
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "For each character chec&kmarked,"
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   660
         Width           =   2055
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "&What to Manage"
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
         Index           =   0
         Left            =   0
         TabIndex        =   3
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
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Text            =   "Name"
         Object.Width           =   4498
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   "Perm"
         Text            =   "Perm"
         Object.Width           =   1085
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   "Temper"
         Text            =   "(Temper)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   "Temp"
         Text            =   "Temp"
         Object.Width           =   1085
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Key             =   "Diff"
         Text            =   "Diff"
         Object.Width           =   1085
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdESCClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   27
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
Attribute VB_Name = "frmPermTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' Name:         frmPermTemp
' Description:  Form that manages perm/temp ratings for characters.
'

Private Const FORM_START_HEIGHT = 6570
Private Const FORM_START_WIDTH = 9150
Private Const FORM_MIN_SCALEHEIGHT = 6165
Private Const FORM_MIN_SCALEWIDTH = 7485
Private Const BOTTOM_MARGIN = 765
Private Const RIGHT_MARGIN = 2310
Private Const HORIZONTAL_GAP = 105
Private Const VERTICAL_GAP = 225

Private Const OPT_MATCH = 0
Private Const OPT_NO_MATCH = 1
Private Const ADJ_PERM = 0
Private Const ADJ_TEMP = 1

Private Const ACT_INCREASE = "Increase"
Private Const ACT_DECREASE = "Decrease"
Private Const ACT_SET = "Set"
Private Const ACT_FILL = "Fill"

Private DeselectSet As StringSet
Private TemperKeys As Collection

Private TempKey As String
Private PermKey As String
Private TemperName As String
Private CharTemp As Variant
Private CharPerm As Variant
Private Populating As Boolean

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
    
    lvwCharacters.ColumnHeaders("Temper").Text = TemperName
    
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
            Character.GetValue PermKey, CharPerm
            Character.GetValue TempKey, CharTemp
            If Not (IsNull(CharPerm) Or IsNull(CharTemp)) Then
                    
                Set NewItem = lvwCharacters.ListItems.Add(, "k" & Character.Name, Character.Name)
                NewItem.ListSubItems.Add , , CharPerm
                NewItem.ListSubItems.Add , , DisplayTemper(CInt(CharPerm), CInt(CharTemp))
                NewItem.ListSubItems.Add , , CharTemp
                NewItem.ListSubItems.Add , , CharTemp - CharPerm
                NewItem.ListSubItems.Add , , Format(CharPerm, "000")
                NewItem.ListSubItems.Add , , Format(CharTemp, "000")
                NewItem.ListSubItems.Add , , Format(CharTemp - CharPerm + 500, "000")
                NewItem.Checked = Not DeselectSet.Has(Character.Name)
                
            End If
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
        .SelectSet(osCharacters).Clear
        .SelectSet(osCharacters).StoreListView lvwCharacters, True
        .GameDate = 0
        .SearchName = cboSearch.Text
        .SearchNot = optNot(OPT_NO_MATCH).Value
    End With
    
End Sub

Private Function DisplayTemper(Perm As Integer, Temp As Integer) As String
'
' Name:         DisplayTemper
' Parameters:   Perm        permanent dots of temper
'               Temp        temporary dots of tempet
' Description:  Return a temper string.
'

    Const P = "o"
    Const PPPPP = "ooooo"
    Const P5 = "O"
    Const T = "õ"
    Const TTTTT = "õõõõõ"
    Const T5 = "Õ"
    Const E = "ø"
    Const EEEEE = "øøøøø"
    Const E5 = "Ø"

    If Temp >= Perm Then
        DisplayTemper = String(Perm, P) & String(Temp - Perm, T)
    Else
        DisplayTemper = String(Temp, P) & String(Perm - Temp, E)
    End If
    
    Do Until Len(DisplayTemper) <= 10
        If InStr(DisplayTemper, EEEEE) > 0 Then
            DisplayTemper = StrReverse(DisplayTemper)
            DisplayTemper = Replace(DisplayTemper, EEEEE, E5, 1, 1)
            DisplayTemper = StrReverse(DisplayTemper)
        ElseIf InStr(DisplayTemper, PPPPP) > 0 Then
            DisplayTemper = Replace(DisplayTemper, PPPPP, P5, 1, 1)
        ElseIf InStr(DisplayTemper, TTTTT) > 0 Then
            DisplayTemper = StrReverse(DisplayTemper)
            DisplayTemper = Replace(DisplayTemper, TTTTT, T5, 1, 1)
            DisplayTemper = StrReverse(DisplayTemper)
        Else
            Exit Do
        End If
    Loop
    
End Function

Private Function LowEnd(Range As String) As String
'
' Name:         LowEnd
' Description:  Find the lesser (left) end of the range expressed in the string.
'

    Dim I As Integer

    If Range <> "" Then
        For I = 1 To Len(Range)
            If Not Mid(Range, I, 1) Like "#" Then Exit For
        Next I
        LowEnd = Left(Range, I - 1)
    Else
        LowEnd = "0"
    End If
    
End Function

Private Function HighEnd(Range As String) As String
'
' Name:         HighEnd
' Description:  Find the higher (right) end of the range expressed in the string.
'

    Dim I As Integer

    If Range <> "" Then
        For I = Len(Range) To 1 Step -1
            If Not Mid(Range, I, 1) Like "#" Then Exit For
        Next I
        HighEnd = Mid(Range, I + 1)
    Else
        HighEnd = "0"
    End If
    
End Function

Private Sub cboAct_Click()
'
' Name:         cboAct_Click
' Description:  Adjust the controls based on the selected act.
'

    Dim NoFill As Boolean
    
    NoFill = Not (cboAct.Text = ACT_FILL)

    txtChange.Visible = NoFill
    updChange.Visible = NoFill
    chkRandomly.Visible = NoFill
    chkCap.Visible = NoFill
    
    If NoFill Then
        If cboAct.Text = ACT_SET Then
            lblGuide.Caption = "&its temporary rating to"
        Else
            lblGuide.Caption = "&its temporary rating by"
        End If
    Else
        lblGuide.Caption = "its temporary rating up to its permanent rating."
    End If

End Sub

Private Sub cboSearch_Click()
'
' Name:         cboSearch_Click
' Description:  The user has chosen a new query, so populate the list.
'
    DeselectSet.Clear
    RefreshList
    
End Sub

Private Sub cboTemper_Click()
'
' Name:         cboTemper_Click
' Description:  Change the listed temper.
'

    On Error Resume Next
    
    TemperName = cboTemper.Text
    lblTemper = TemperName
    PermKey = TemperKeys(TemperName)
    TempKey = "temp" & PermKey
    RefreshList
    
    On Error GoTo 0

End Sub

Private Sub chkRandomly_Click()
'
' Name:         chkRandomly_Click
' Description:  Adjust the txtChange display.
'

    If chkRandomly.Value = vbChecked Then
        txtChange.Text = "0 - " & txtChange.Text
    Else
        txtChange.Text = CStr(Int(Val(HighEnd(txtChange.Text))))
    End If

End Sub

Private Sub cmdChange_Click()
'
' Name:         cmdChange_Click
' Description:  Apply the changes chosen by the user to those characters checkmarked.
'

    Dim SelSet As StringSet
    Dim I As Long
    Dim Character As Object
    Dim ChangeVal As Long
    Dim LowVal As Long
    
    Set SelSet = New StringSet
    
    With lvwCharacters.ListItems
        For I = 1 To .Count
            If .Item(I).Checked Then SelSet.Add .Item(I).Text
        Next I
    End With
    
    If chkRandomly.Value = vbChecked Then
        LowVal = Int(Val(LowEnd(txtChange.Text)))
        I = Int(Val(HighEnd(txtChange.Text))) - LowVal
    Else
        ChangeVal = Int(Val(txtChange.Text))
    End If
    
    CharacterList.First
    Do Until CharacterList.Off
        Set Character = CharacterList.Item
        If SelSet.Has(Character.Name) Then
            Character.GetValue PermKey, CharPerm
            Character.GetValue TempKey, CharTemp
            If chkRandomly.Value = vbChecked Then
                ChangeVal = Int(Rnd * (I + 1)) + LowVal
            End If
            Select Case cboAct.Text
                Case ACT_INCREASE:  CharTemp = CharTemp + ChangeVal
                Case ACT_DECREASE:  CharTemp = CharTemp - ChangeVal
                Case ACT_SET:       CharTemp = ChangeVal
                Case ACT_FILL:      CharTemp = CharPerm
            End Select
            If CharTemp < 0 Then CharTemp = 0
            If chkCap.Value = vbChecked And CharTemp > CharPerm Then CharTemp = CharPerm
            Character.SetValue TempKey, CharTemp
            Character.LastModified = Now
        End If
        CharacterList.MoveNext
    Loop
    
    If Not SelSet.Count = 0 Then
        Game.DataChanged = True
        mdiMain.AnnounceChanges Me, atTempers
        RefreshList
    End If
    
    Set SelSet = Nothing

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
    Dim TChange As Boolean
    
    QChange = mdiMain.CheckForChanges(Me, atQueries)
    CChange = mdiMain.CheckForChanges(Me, atCharacters)
    TChange = mdiMain.CheckForChanges(Me, atTempers)
    
    If QChange Then
        PopulateSearches cboSearch.Text
    Else
        If CChange Or TChange Then RefreshList
    End If

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Fill the list and select the first character.
'

    Dim Find As Variant
    Dim I As Long
    
    Set DeselectSet = New StringSet
    Set TemperKeys = New Collection
    
    For Each Find In Game.QueryEngine.TitlesToKeys
        If Left(Find, 4) = "temp" And Game.QueryEngine.KeysToTypes(Find) = qtNumber Then
            
            TemperName = ""
            On Error Resume Next
            TemperName = CStr(Game.QueryEngine.KeysToTitles(Mid(Find, 5)))
            If TemperName <> "" Then    'It's a temper!
                TempKey = Find
                PermKey = Mid(Find, 5)
                I = InStr(TemperName, " (")
                If I > 0 Then TemperName = Left(TemperName, I - 1)
                cboTemper.AddItem TemperName
                TemperKeys.Add PermKey, TemperName
            End If
            On Error GoTo 0
            
        End If
    Next Find

    For I = cboTemper.ListCount - 1 To 0 Step -1
        If cboTemper.List(I) = "Willpower" Then
            cboTemper.ListIndex = I
            Exit For
        End If
    Next I
    
    cboAct.ListIndex = 0
    
    Me.Width = FORM_START_WIDTH
    Me.Height = FORM_START_HEIGHT
    
    PopulateSearches "Active Characters"
    
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
    Set DeselectSet = Nothing
    Set TemperKeys = Nothing
    
End Sub

Private Sub lvwCharacters_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'
' Name:         lvwCharacters_ColumnClick
' Description:  Change the key by which the entries are sorted, or the sort order on a second click.
'
    
    Dim Sort As Integer
    Select Case ColumnHeader.Index
        Case 2: Sort = 5
        Case 4: Sort = 6
        Case 5: Sort = 7
        Case Else: Sort = ColumnHeader.Index - 1
    End Select
    
    If lvwCharacters.SortKey = Sort Then
        lvwCharacters.SortOrder = IIf(lvwCharacters.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwCharacters.SortKey = Sort
    End If

End Sub

Private Sub lvwCharacters_DblClick()
'
' Name:         lstCharacters_DblClick
' Description:  See cmdShow_Click.
'

    If Not lvwCharacters.SelectedItem Is Nothing Then
        lvwCharacters.SelectedItem.Checked = Not lvwCharacters.SelectedItem.Checked
        Call lvwCharacters_ItemCheck(lvwCharacters.SelectedItem)
    End If

End Sub

Private Sub lvwCharacters_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'
' Name:         lvwCharacters_ItemCheck
' Description:  Add or remove the name from the unselected set.
'

    Set lvwCharacters.SelectedItem = Item
    If Item.Checked Then
        DeselectSet.Remove Item.Text
    Else
        DeselectSet.Add Item.Text
    End If

End Sub

Private Sub lvwCharacters_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
' Name:         lvwCharacters_ItemClick
' Description:  Select the character and display its info at the lower right.
'

    Dim ShowChar As Boolean

    ShowChar = False
    If Not Item Is Nothing Then
        CharacterList.MoveTo Item.Text
        If Not CharacterList.Off Then
            CharacterList.Item.GetValue PermKey, CharPerm
            CharacterList.Item.GetValue TempKey, CharTemp
            If Not (IsNull(CharPerm) Or IsNull(CharTemp)) Then
                Populating = True
                lblName = Item.Text
                txtChar(ADJ_PERM).Text = CharPerm
                txtChar(ADJ_TEMP).Text = CharTemp
                ShowChar = True
                Populating = False
            End If
        End If
    End If
    fraCharacter.Visible = ShowChar

End Sub

Private Sub optNot_Click(Index As Integer)
'
' Name:         optNot_Click
' Description:  The user has inverted the query, so populate the list.
'
    DeselectSet.Clear
    RefreshList
    
End Sub

Private Sub PopulateSearches(Default As String)
'
' Name:         PopulateSearches
' Parameters:   Default         Default search to use
' Description:  Fill cboSearch from the QueryList.  This will force either a
'               cboSearch_Click event (and thus a RefreshList) or a RefreshList
'               directly.

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

Private Sub txtChange_GotFocus()
'
' Name:         txtChange_GotFocus
' Description:  Select the text.
'
    SelectText txtChange

End Sub

Private Sub txtChar_Change(Index As Integer)
'
' Name:         txtChar_Change
' Description:  Adjust the character's temper.
'

    If Not Populating Then
    
        CharacterList.MoveTo lvwCharacters.SelectedItem.Text
        If Not CharacterList.Off Then
        
            Dim WhichKey As String
            Dim NewNum As Integer
            Dim OtherNum As Integer
            
            Populating = True
            WhichKey = IIf(Index = ADJ_PERM, PermKey, TempKey)
            NewNum = Int(Val(txtChar(Index).Text))
            If NewNum < 0 Then NewNum = 0
            If NewNum > 9999 Then NewNum = 9999
            txtChar(Index).Text = CStr(NewNum)
            Populating = False
            
            CharacterList.Item.SetValue WhichKey, NewNum
            CharacterList.Item.LastModified = Now
            Game.DataChanged = True
            mdiMain.AnnounceChanges Me, atTempers
                                                
            CharacterList.Item.GetValue PermKey, CharPerm
            CharacterList.Item.GetValue TempKey, CharTemp
            If Not (IsNull(CharPerm) Or IsNull(CharTemp)) Then
                    
                With lvwCharacters.SelectedItem
                    .ListSubItems(1).Text = CharPerm
                    .ListSubItems(2).Text = DisplayTemper(CInt(CharPerm), CInt(CharTemp))
                    .ListSubItems(3).Text = CharTemp
                    .ListSubItems(4).Text = CharTemp - CharPerm
                    .ListSubItems(5).Text = Format(CharPerm, "000")
                    .ListSubItems(6).Text = Format(CharTemp, "000")
                    .ListSubItems(7).Text = Format(CharTemp - CharPerm + 500, "000")
                End With
                
            End If
            
        End If
    
    End If

End Sub

Private Sub txtChar_GotFocus(Index As Integer)
'
' Name:         txtChar_GotFocus
' Description:  Select the text.
'
    SelectText txtChar(Index)

End Sub

Private Sub updChange_DownClick()
'
' Name:         updChange_DownClick
' Description:  Adjust txtChange.
'

    If chkRandomly.Value = vbChecked Then
        txtChange.Text = LowEnd(txtChange.Text) & " - " & CStr(Int(Val(HighEnd(txtChange.Text))) - 1)
    Else
        txtChange.Text = CStr(Int(Val(txtChange.Text)) - 1)
    End If
    
End Sub

Private Sub updChange_UpClick()
'
' Name:         updChange_UpClick
' Description:  Adjust txtChange.
'

    If chkRandomly.Value = vbChecked Then
        txtChange.Text = LowEnd(txtChange.Text) & " - " & CStr(Int(Val(HighEnd(txtChange.Text))) + 1)
    Else
        txtChange.Text = CStr(Int(Val(txtChange.Text)) + 1)
    End If
    
End Sub
