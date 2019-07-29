VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQuery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search for Character"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   Icon            =   "frmQuery.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5445
   ScaleMode       =   0  'User
   ScaleWidth      =   9030
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   240
      TabIndex        =   15
      Top             =   4320
      Width           =   8535
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve This Search..."
         Height          =   375
         Left            =   6480
         TabIndex        =   19
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Perform Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   6480
         TabIndex        =   18
         Top             =   0
         Width           =   2055
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add &New Term"
         Height          =   375
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   2055
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove Term"
         Height          =   375
         Left            =   2160
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Frame fraSortOrder 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   3240
      Width           =   2055
      Begin VB.OptionButton optSortOrder 
         Caption         =   "Descend"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   10
         Top             =   0
         Width           =   975
      End
      Begin VB.OptionButton optSortOrder 
         Caption         =   "As&cend"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "S&how Character"
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete This Search"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.ComboBox cboSearches 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   6720
      TabIndex        =   30
      Top             =   240
      Width           =   2055
   End
   Begin VB.ComboBox cboSort 
      Height          =   315
      Left            =   4560
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2880
      Width           =   2055
   End
   Begin VB.OptionButton optAnyAll 
      Caption         =   "All"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   13
      Top             =   3915
      Width           =   615
   End
   Begin VB.OptionButton optAnyAll 
      Caption         =   "&Any"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   12
      Top             =   3675
      Value           =   -1  'True
      Width           =   615
   End
   Begin MSComctlLib.ListView lvwResults 
      Height          =   2055
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3625
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
         Key             =   "Name"
         Text            =   "Name"
         Object.Width           =   5344
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Match"
         Text            =   "Matching Value"
         Object.Width           =   5344
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Sort"
         Text            =   "Sort Value"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame fraTerm 
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   4320
      Width           =   8655
      Begin MSComCtl2.UpDown updNumber 
         Height          =   315
         Index           =   0
         Left            =   7335
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   0
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         _Version        =   393216
         Orientation     =   1
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtNumber 
         Height          =   315
         Index           =   0
         Left            =   6720
         TabIndex        =   28
         Text            =   "1"
         Top             =   0
         Width           =   615
      End
      Begin VB.TextBox txtFind 
         Height          =   315
         Index           =   0
         Left            =   4440
         TabIndex        =   26
         Top             =   0
         Width           =   2055
      End
      Begin VB.CheckBox chkNot 
         Caption         =   "NOT"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   24
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox cboCompare 
         Height          =   315
         Index           =   0
         ItemData        =   "frmQuery.frx":0742
         Left            =   2280
         List            =   "frmQuery.frx":0749
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   0
         Width           =   2055
      End
      Begin VB.ComboBox cboKey 
         Height          =   315
         Index           =   0
         Left            =   360
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblJoin 
         Caption         =   "OR"
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
         Left            =   360
         TabIndex        =   22
         Top             =   420
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblTermNum 
         AutoSize        =   -1  'True
         Caption         =   "&1."
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lblX 
         Caption         =   "x"
         Height          =   255
         Index           =   0
         Left            =   6570
         TabIndex        =   27
         Top             =   45
         Width           =   135
      End
   End
   Begin VB.Label lblFound 
      Caption         =   "0 characters found."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2910
      Width           =   3375
   End
   Begin VB.Label lblNoTerms 
      Alignment       =   1  'Right Justify
      Caption         =   "(Add search terms, or the search will match all characters)"
      Height          =   255
      Left            =   4560
      TabIndex        =   31
      Top             =   3795
      Width           =   4215
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   240
      Picture         =   "frmQuery.frx":0764
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblLabel 
      Caption         =   "&Search:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   840
      TabIndex        =   0
      Top             =   300
      Width           =   855
   End
   Begin VB.Label lblLabel 
      Caption         =   "Sort &by:"
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   6
      Top             =   2940
      Width           =   615
   End
   Begin VB.Label lblLabel 
      Caption         =   "of the following terms:"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   14
      Top             =   3795
      Width           =   1665
   End
   Begin VB.Label lblLabel 
      Caption         =   "List Characters that match"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   3795
      Width           =   1935
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FIRST_TERM_LEFT = 4440
Private Const SECOND_TERM_LEFT = 6720
Private Const UPDOWN_OFFSET = 615
Private Const TWO_COLUMN_WIDTH = 3030
Private Const THREE_COLUMN_WIDTH = 2025

Private Const OPT_ANY = 0
Private Const OPT_ALL = 1

Private Const OPT_ASCEND = 0
Private Const OPT_DESCEND = 1

Private Const NAME_INDEX = 0
Private Const MATCH_INDEX = 1
Private Const DIVIDER_INDEX = 2

Private UserQuery As QueryClass             'Current User-defined Query
Private Populating As Boolean               'Flag to disable code when populating fields
Private TermCount As Integer                'Current number of terms in a query

Public Sub SetDefaultOutput()
'
' Name:         SetDefaultOutput
' Description:  Initilize the OutputEngineClass with default output settings.
'
    With OutputEngine
        .Template = tnSearch
        .SearchName = RecentSearchName
        .SearchNot = False
        .GameDate = 0
    End With
    
End Sub

Private Sub cboCompare_Click(Index As Integer)
'
' Name:         cboCompare_Click
' Description:  Show or hide the input boxes as needed according to the type
'               of query and comparison.  This is only needed for Trait List queries:
'               all others handle this function when the subject of the query is
'               selected (cboKey_Click).
'

    If cboKey(Index).Tag = CStr(qtTraitList) Then ShowInputFields Index

End Sub

Private Sub cboKey_Click(Index As Integer)
'
' Name:         cboKey_Click
' Description:  A new key has been selected: populate cboComparison with the
'               comparisons that can be paired with it.
'

    Dim Key As String
    Dim KeyType As QueryKeyType
    
    Key = Game.QueryEngine.TitlesToKeys(cboKey(Index).Text)
    KeyType = Game.QueryEngine.KeysToTypes(Key)
    
    If cboKey(Index).Tag <> CStr(KeyType) Or cboCompare(Index).ListCount = 0 Then
        
        cboKey(Index).Tag = CStr(KeyType)
        
        With cboCompare(Index)
            .Clear
            
            Select Case KeyType
        
                Case qtField
                    .AddItem "contains", 0
                    .ItemData(0) = qcContains
                    .AddItem "equals", 1
                    .ItemData(1) = qcEquals
                    .ListIndex = 0
                    ShowInputFields Index
                Case qtNumber
                    .AddItem "is less than", 0
                    .ItemData(0) = qcLess
                    .AddItem "is no more than", 1
                    .ItemData(1) = qcNoMore
                    .AddItem "equals", 2
                    .ItemData(2) = qcEquals
                    .AddItem "is at least", 3
                    .ItemData(3) = qcAtLeast
                    .AddItem "is greater than", 4
                    .ItemData(4) = qcGreater
                    .ListIndex = 2
                    ShowInputFields Index
                Case qtTraitList
                    .AddItem "total less than", 0
                    .ItemData(0) = qcTotalsLess
                    .AddItem "total no more than", 1
                    .ItemData(1) = qcTotalsNoMore
                    .AddItem "total", 2
                    .ItemData(2) = qcTotals
                    .AddItem "total at least", 3
                    .ItemData(3) = qcTotalsAtLeast
                    .AddItem "total more than", 4
                    .ItemData(4) = qcTotalsMore
                    .AddItem "contain", 5
                    .ItemData(5) = qcContains
                    .AddItem "contain the note", 6
                    .ItemData(6) = qcContainsNote
                    .AddItem "contain less than", 7
                    .ItemData(7) = qcContainsLess
                    .AddItem "contain no more than", 8
                    .ItemData(8) = qcContainsNoMore
                    .AddItem "contain exactly", 9
                    .ItemData(9) = qcContainsExactly
                    .AddItem "contain at least", 10
                    .ItemData(10) = qcContainsAtLeast
                    .AddItem "contain more than", 11
                    .ItemData(11) = qcContainsMore
                    .ListIndex = 5
                Case qtDate
                    .AddItem "is earlier than", 0
                    .ItemData(0) = qcLess
                    .AddItem "is no later than", 1
                    .ItemData(1) = qcNoMore
                    .AddItem "equals", 2
                    .ItemData(2) = qcEquals
                    .AddItem "is no earlier than", 3
                    .ItemData(3) = qcAtLeast
                    .AddItem "is later than", 4
                    .ItemData(4) = qcGreater
                    .ListIndex = 2
                    ShowInputFields Index
                    txtFind(Index).Text = Format(Now, "Short Date")
                Case qtBoolean
                    .AddItem "is true", 0
                    .ItemData(0) = qcIsTrue
                    .AddItem "is false", 1
                    .ItemData(1) = qcIsFalse
                    .ListIndex = 0
                    ShowInputFields Index
            End Select
        End With
        
    End If

End Sub

Private Sub cboSearches_Click()
'
' Name:         cboSearches_Click
' Description:  Copy the selected query into the UserQuery.  Create and
'               populate its search terms.
'

    If Not Populating Then

        Game.QueryEngine.QueryList.MoveTo cboSearches.Text
        If Not Game.QueryEngine.QueryList.Off Then
        
            Dim Search As QueryClass
            Dim SortTitle As String
            Dim I As Integer
            
            Set Search = Game.QueryEngine.QueryList.Item
        
            UserQuery.Clear
            UserQuery.Name = RecentSearchName
            Do Until TermCount = 0
                RemoveTerm
            Loop
            
            With Search
            
                UserQuery.MatchAll = .MatchAll
                UserQuery.SortKey = .SortKey
                UserQuery.Inventory = .Inventory
                
                optAnyAll(OPT_ANY).Value = Not .MatchAll
                optAnyAll(OPT_ALL).Value = .MatchAll
                
                Populating = True
                optSortOrder(IIf(.SortDescend, OPT_DESCEND, OPT_ASCEND)).Value = True
                Select Case .SortKey
                    Case qkName
                        cboSort.ListIndex = NAME_INDEX
                    Case ""
                        cboSort.ListIndex = MATCH_INDEX
                    Case Else
                        For I = 3 To cboSort.ListCount - 1
                            If Game.QueryEngine.TitlesToKeys(cboSort.List(I)) = .SortKey Then
                                cboSort.ListIndex = I
                                Exit For
                            End If
                        Next I
                End Select
                Populating = False
                
                .First
                Do Until .Off
                    UserQuery.AddClause .Clause.Key, .Clause.Find, .Clause.Number, _
                                        .Clause.Comparison, .Clause.CompNot
                    AddTerm .Clause
                    Search.MoveNext
                Loop
        
            End With
        
            cmdSearch_Click
        
        End If

    End If

End Sub

Private Sub cboSort_Click()
'
' Name:         cboSort_Click
' Description:  Format the ListView to fit the sorting criteria;
'               then sort the list by it.
'

    If cboSort.ListIndex = 2 Then cboSort.ListIndex = 0
    
    If cboSort.ListIndex < 2 Then   'Sort by name or matching values
        
        lvwResults.ColumnHeaders(1).Width = TWO_COLUMN_WIDTH
        lvwResults.ColumnHeaders(2).Width = TWO_COLUMN_WIDTH
        lvwResults.ColumnHeaders(3).Width = 0
    
        lvwResults.SortKey = cboSort.ListIndex
        lvwResults.Sorted = True
        
        UserQuery.SortKey = IIf(cboSort.ListIndex = 0, qkName, "")
        
    Else
    
        lvwResults.ColumnHeaders(1).Width = THREE_COLUMN_WIDTH
        lvwResults.ColumnHeaders(2).Width = THREE_COLUMN_WIDTH
        lvwResults.ColumnHeaders(3).Width = THREE_COLUMN_WIDTH
        
        lvwResults.ColumnHeaders(3).Text = cboSort.Text
        
        lvwResults.Sorted = False
        
        If Not Populating Then cmdSearch_Click
    
    End If

End Sub

Private Sub cmdAdd_Click()
'
' Name:         cmdAdd_Click
' Description:  Add a new term to the user query.
'

    UserQuery.AddClause qkName, "", 0, qcContains, False
    UserQuery.Last
    AddTerm UserQuery.Clause

End Sub

Private Sub cmdClose_Click()
'
' Name:         cmdClose_Click
' Description:  Hide this window.
'

    Me.Hide

End Sub

Private Sub cmdDelete_Click()
'
' Name:         cmdDelete_Click
' Description:  Delect the currently selected query.
'

    Game.QueryEngine.QueryList.MoveTo cboSearches.Text

    If Not Game.QueryEngine.QueryList.Off Then
    
        If MsgBox("Are you sure you want to delete the search """ & _
                    cboSearches.Text & """?", vbYesNo, "Delete Search") = vbYes Then
                    
            Game.QueryEngine.QueryList.Remove
            cboSearches.RemoveItem cboSearches.ListIndex
            mdiMain.AnnounceChanges Me, atQueries
            Game.DataChanged = True
                    
        End If
    
    End If

End Sub

Private Sub cmdRemove_Click()
'
' Name:         cmdRemove_Click
' Description:  Remove the last term of the current query.
'

    RemoveTerm

End Sub

Private Sub cmdSave_Click()
'
' Name:         cmdSave_Click
' Description:  If this is a valid query, save it in the query list.
'

    If ValidateTerms Then
    
        Dim Name As String
        Dim X As Integer
        
        Name = Trim(InputBox("Enter a name to describe this search:", _
                     "Saved Search Name", cboSearches.Text))
        
        If Name <> "" And Name <> RecentSearchName Then
        
            Populating = True
        
            Game.QueryEngine.QueryList.MoveTo Name
            If Not Game.QueryEngine.QueryList.Off Then
            
                If MsgBox("A search named """ & Name & """ already exists." & vbCrLf & _
                           "Are you sure you want to replace it?", vbYesNo, "Replace Search") _
                           = vbNo Then Exit Sub
                Game.QueryEngine.QueryList.Remove
                For X = 0 To cboSearches.ListCount - 1
                    If cboSearches.List(X) = Name Then cboSearches.ListIndex = X
                Next X
                
            Else
            
                cboSearches.AddItem Name
                cboSearches.ListIndex = cboSearches.NewIndex
            
            End If
                        
            UserQuery.Name = Name
            Game.QueryEngine.AddQueryCopy UserQuery
            UserQuery.Name = RecentSearchName
            mdiMain.AnnounceChanges Me, atQueries
            Game.DataChanged = True
            
            Populating = False
        
        End If
    
    End If

End Sub

Private Sub cmdSearch_Click()
'
' Name:         cmdSearch_Click
' Description:  Perform a search.  First validate the terms, then run the query,
'               then populate the list.
'
        
    If ValidateTerms Then
        
        Dim NewItem As ListItem
        
        Screen.MousePointer = vbHourglass
        
        With Game.QueryEngine
        
            .MakeQuery UserQuery, Not (UserQuery.SortKey = "" Or UserQuery.SortKey = qkName)
        
            .Results.First
            .Values.First
            .SortList.First
            lvwResults.ListItems.Clear
        
            lblFound.Caption = CStr(.Results.Count) & " character" & _
                               IIf(.Results.Count = 1, "", "s") & " found."
            
            Do Until .Results.Off
                Set NewItem = lvwResults.ListItems.Add(, , .Results.Item.Name)
                NewItem.SmallIcon = .Results.Item.Race
                NewItem.ListSubItems.Add , "Match", .Values.Item
                If Not .SortList.Off Then
                    NewItem.ListSubItems.Add , "Sort", .SortList.Item
                    .SortList.MoveNext
                End If
                .Results.MoveNext
                .Values.MoveNext
            Loop
        
        End With
        
        Screen.MousePointer = vbDefault

    End If

End Sub

Private Sub cmdShow_Click()
'
' Name:         cmdShow_Click
' Description:  Show a character sheet for the selected character.
'

    If Not lvwResults.SelectedItem Is Nothing Then
    
        mdiMain.ShowCharacterSheet lvwResults.SelectedItem.Text
    
    End If

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Initialize all the controls for the window.
'

    Dim Key As Variant
    
    'No Terms Selected

    fraTerm(0).Visible = False
    fraButtons.Top = fraTerm(0).Top
    cmdRemove.Value = False
    lblNoTerms.Visible = True
    lvwResults.SmallIcons = mdiMain.imlSmallIcons
    
    ' Add "Most Recent Search" Query, if it's not there; then
    ' fill list of saved queries
    
    With Game.QueryEngine.QueryList
    
        .First
        Do Until .Off
            Set UserQuery = .Item
            If UserQuery.Inventory = qiCharacters And UserQuery.Name <> RecentSearchName Then
                cboSearches.AddItem UserQuery.Name
            End If
            .MoveNext
        Loop
        
        Set UserQuery = Nothing
        .MoveTo RecentSearchName
        If .Off Then
            Set UserQuery = New QueryClass
            UserQuery.Name = RecentSearchName
            UserQuery.Inventory = qiCharacters
            .InsertSorted UserQuery
            mdiMain.AnnounceChanges Me, atQueries
        Else
            Set UserQuery = .Item
            UserQuery.Clear
            UserQuery.Name = RecentSearchName
        End If
        
    End With
       
    'Populate cboKey(0) and cboSort
    
    With Game.QueryEngine
    
        For Each Key In .TitlesToKeys
        
            If (CInt(.KeysToInventories(Key)) And qiCharacters) Then
                cboKey(0).AddItem CStr(.KeysToTitles(Key))
                cboSort.AddItem CStr(.KeysToTitles(Key))
            End If
            
        Next Key
    
    End With
    
    '
    ' Populating cboSort will trigger the sort-formatting of cboSort_Click
    '
    cboSort.AddItem "Name", NAME_INDEX
    cboSort.AddItem "Matching Values", MATCH_INDEX
    cboSort.AddItem "------------------------------", DIVIDER_INDEX
    cboSort.ListIndex = NAME_INDEX
        
    TermCount = 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
' Name:         Form_QueryUnload
' Description:  If the user has closed the form from its control menu (or little X button)
'               then just hide the form, don't unload it.

    If UnloadMode = vbFormControlMenu Then
        Me.Hide
        Cancel = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Unload
' Description:  Unload all the extra controls when the form is unloaded.
'

    Dim X As Integer
    
    For X = fraTerm.Count - 1 To 1 Step -1
        Unload cboKey(X)
        Unload cboCompare(X)
        Unload chkNot(X)
        Unload txtFind(X)
        Unload lblX(X)
        Unload txtNumber(X)
        Unload updNumber(X)
        Unload lblTermNum(X)
        Unload lblJoin(X)
        Unload fraTerm(X)
    Next X
    
    Set UserQuery = Nothing
        
End Sub

Private Sub lvwResults_DblClick()
'
' Name:         lvwResults_DblClick
' Description:  Show the selected character.
'
    Call cmdShow_Click

End Sub

Private Sub optAnyAll_Click(Index As Integer)
'
' Name:         optAnyAll_Click
' Parameters:   Index: 0 - Any, 1 - All
' Description:  Set whether the user query finds characters matching any or all terms.
'

    Dim LJoin As Label

    UserQuery.MatchAll = (Index = 1)
    
    For Each LJoin In lblJoin
        LJoin.Caption = IIf(Index = OPT_ANY, "OR", "AND")
    Next LJoin
    
End Sub

Private Sub AddTerm(Clause As QueryClauseClass)
'
' Name:         AddTerm
' Parameters:   Clause      the query clause whose pieces populate the new term
' Description:  Add a new frame with the controls with which to build a query.
'

    Dim I As Integer
    Dim X As Integer
    Dim LoadNew As Boolean
    Dim KeyTitle As String
    
    lblNoTerms.Visible = False
    cmdRemove.Visible = True
    Me.Height = Me.Height + fraTerm(0).Height
    fraButtons.Top = fraButtons.Top + fraTerm(0).Height
    
    '
    ' Create new Controls if needed.
    '
    
    LoadNew = True
    For I = 0 To fraTerm.Count - 1
        If Not fraTerm(I).Visible Then
            LoadNew = False
            Exit For
        End If
    Next I
    
    If LoadNew Then
        
        I = fraTerm.Count
        Load fraTerm(I)
        Load lblTermNum(I)
        Load lblJoin(I)
        Load cboKey(I)
        Load cboCompare(I)
        Load chkNot(I)
        Load txtFind(I)
        Load lblX(I)
        Load txtNumber(I)
        Load updNumber(I)
            
        Set lblTermNum(I).Container = fraTerm(I)
        Set lblJoin(I).Container = fraTerm(I)
        Set cboKey(I).Container = fraTerm(I)
        Set cboCompare(I).Container = fraTerm(I)
        Set chkNot(I).Container = fraTerm(I)
        Set txtFind(I).Container = fraTerm(I)
        Set lblX(I).Container = fraTerm(I)
        Set txtNumber(I).Container = fraTerm(I)
        Set updNumber(I).Container = fraTerm(I)
        
        fraTerm(I).Visible = True
        lblTermNum(I).Visible = True
        cboKey(I).Visible = True
        cboCompare(I).Visible = True
        chkNot(I).Visible = True
        txtFind(I).Visible = True
        lblX(I).Visible = True
        txtNumber(I).Visible = True
        updNumber(I).Visible = True
        
    End If

    '
    ' Load the contents for the dropdown Key List
    '
    If LoadNew Then
        For X = 0 To cboKey(0).ListCount - 1
            cboKey(I).AddItem cboKey(0).List(X)
        Next X
    End If
    
    lblTermNum(I).Caption = "&" & CStr(I + 1) & "."
    lblJoin(I).Caption = lblJoin(0).Caption
    If I > 0 Then lblJoin(I - 1).Visible = True

    'Now, fill all this out according to the current clause of the query

    On Error Resume Next            ' Just in case there's no match for the key in the collection
    KeyTitle = qkName
    KeyTitle = Game.QueryEngine.KeysToTitles(Clause.Key)
    On Error GoTo 0
    
    For X = 0 To cboKey(I).ListCount - 1
        If cboKey(I).List(X) = KeyTitle Then
            cboKey(I).ListIndex = X
            Exit For
        End If
    Next X

    ' Setting the ListIndex above triggers a click event that populates
    ' cboCompare, enabling the selection below
    
    For X = 0 To cboCompare(I).ListCount - 1
        If cboCompare(I).ItemData(X) = Clause.Comparison Then
            cboCompare(I).ListIndex = X
            Exit For
        End If
    Next X

    txtFind(I).Text = Clause.Find
    txtNumber(I).Text = Clause.Number
    chkNot(I).Value = IIf(Clause.CompNot, vbChecked, vbUnchecked)
    
    fraTerm(I).Top = fraTerm(0).Top + (fraTerm(0).Height * I)
    fraTerm(I).Visible = True
    TermCount = TermCount + 1
    
    If TermCount <> UserQuery.ClauseCount Then Stop
    
End Sub

Private Sub RemoveTerm()
'
' Name:         RemoveTerm
' Description:  Remove the last term of the query.
'

    If TermCount > 0 Then
    
        TermCount = TermCount - 1
        
        UserQuery.RemoveLast
        fraButtons.Top = fraTerm(TermCount).Top
        Me.Height = Me.Height - fraTerm(0).Height
        
        fraTerm(TermCount).Visible = False
        If TermCount > 0 Then lblJoin(TermCount - 1).Visible = False
        cmdRemove.Visible = (TermCount > 0)
        lblNoTerms.Visible = (TermCount = 0)
    
    End If

End Sub

Private Sub ShowInputFields(Index As Integer)
'
' Name:         ShowInputFields
' Parameters:   Index       Index of the term to format
' Description:  Format the visibility and alignment of txtFind, lblX,
'               txtNumber and updNumber according to the query and comparison
'               selected.
'
    
    Dim ShowFind As Boolean
    Dim ShowNumber As Boolean
    
    Select Case CLng(cboKey(Index).Tag)
        
        Case qtField
            ShowFind = True
            ShowNumber = False
        Case qtNumber
            ShowFind = False
            ShowNumber = True
        Case qtTraitList
            Select Case cboCompare(Index).ItemData(cboCompare(Index).ListIndex)
                Case qcContains, qcContainsNote
                    ShowFind = True
                    ShowNumber = False
                Case qcTotals, qcTotalsLess, qcTotalsNoMore, qcTotalsMore, qcTotalsAtLeast
                    ShowFind = False
                    ShowNumber = True
                Case Else
                    ShowFind = True
                    ShowNumber = True
            End Select
        Case qtDate
            ShowFind = True
            ShowNumber = False
        Case qtBoolean
            ShowFind = False
            ShowNumber = False
    End Select
    
    txtFind(Index).Visible = ShowFind
    lblX(Index).Visible = ShowFind And ShowNumber
    txtNumber(Index).Visible = ShowNumber
    updNumber(Index).Visible = ShowNumber
        
    If ShowNumber And Not ShowFind Then
        txtNumber(Index).Left = FIRST_TERM_LEFT
    ElseIf ShowNumber And ShowFind Then
        txtNumber(Index).Left = SECOND_TERM_LEFT
    End If
    updNumber(Index).Left = txtNumber(Index).Left + UPDOWN_OFFSET
    
End Sub

Private Sub optSortOrder_Click(Index As Integer)
'
' Name:         optSortOrder_Click
' Description:  Set the sorting order of the search.
'

    If lvwResults.Sorted Then
        lvwResults.SortOrder = IIf(Index = OPT_ASCEND, lvwAscending, lvwDescending)
        UserQuery.SortDescend = (Index = OPT_DESCEND)
    Else
        If Not Populating Then Call cmdSearch_Click
    End If

End Sub

Private Sub txtFind_GotFocus(Index As Integer)
'
' Name:         txtFind_GotFocus
' Description:  Select the text when this field gets the focus.
'
    SelectText txtFind(Index)

End Sub

Private Sub txtNumber_GotFocus(Index As Integer)
'
' Name:         txtNumber_GotFocus
' Description:  Select the text when this field gets the focus.
'
    SelectText txtNumber(Index)

End Sub

Private Sub txtNumber_Validate(Index As Integer, Cancel As Boolean)
'
' Name:         txtNumber_Validate
' Description:  If the value isn't a number, make it one.
'
    txtNumber(Index).Text = CStr(Val(txtNumber(Index).Text))

End Sub

Private Sub updNumber_DownClick(Index As Integer)
'
' Name:         updNumber_DownClick
' Description:  Decrement the number field.
'
    txtNumber(Index).Text = CStr(Val(txtNumber(Index).Text) - 1)

End Sub

Private Sub updNumber_UpClick(Index As Integer)
'
' Name:         updNumber_UpClick
' Description:  Increment the number field.
'
    txtNumber(Index).Text = CStr(Val(txtNumber(Index).Text) + 1)

End Sub

Private Function ValidateTerms() As Boolean
'
' Name:         ValidateTerms
' Returns:      TRUE iff all terms are good.
' Description:  Ensure the terms are valid; supply the UserQuery with
'               their values.
'

    Dim I As Integer
    Dim Key As String
    
    UserQuery.Inventory = qiCharacters
    UserQuery.MatchAll = optAnyAll(OPT_ALL).Value
    UserQuery.SortDescend = optSortOrder(OPT_DESCEND).Value
    
    On Error Resume Next
    UserQuery.SortKey = ""
    If cboSort.ListIndex = 0 Then
        UserQuery.SortKey = qkName
    Else
        UserQuery.SortKey = Game.QueryEngine.TitlesToKeys(cboSort.Text)
    End If
    On Error GoTo 0
    
    UserQuery.First
    For I = 0 To UserQuery.ClauseCount - 1
    
        On Error Resume Next
        Key = "(none)"
        Key = Game.QueryEngine.TitlesToKeys(cboKey(I).Text)
        On Error GoTo 0
        
        If Key = "(none)" Then
            MsgBox "Grapevine Error -- no query key is associated with """ & _
                    cboKey(I).Text & """!", vbCritical, "Query Error"
            Exit For
        End If
        
        If cboKey(I).Tag = CStr(qtDate) And Not IsDate(txtFind(I).Text) Then
            MsgBox "One of your search terms needs you to enter a valid date.", _
                    vbExclamation, "Enter Date"
            txtFind(I).SetFocus
            Exit For
        End If
        
        With UserQuery.Clause
            .Key = Key
            .Comparison = cboCompare(I).ItemData(cboCompare(I).ListIndex)
            .CompNot = (chkNot(I).Value = vbChecked)
            .Find = txtFind(I).Text
            .Number = Val(txtNumber(I).Text)
        End With
        
        UserQuery.MoveNext
        
    Next I

    ValidateTerms = UserQuery.Off

End Function
