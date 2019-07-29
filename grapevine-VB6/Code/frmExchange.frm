VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExchange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player and Character Exchange"
   ClientHeight    =   6165
   ClientLeft      =   210
   ClientTop       =   570
   ClientWidth     =   9060
   Icon            =   "frmExchange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   9060
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkHideST 
      Alignment       =   1  'Right Justify
      Caption         =   "Don't save text marked ST Only"
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   8295
      Begin VB.CommandButton cmdSelectSame 
         Caption         =   "Select Same &Date"
         Height          =   375
         Index           =   0
         Left            =   4080
         TabIndex        =   10
         Top             =   3120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdSelectSame 
         Caption         =   "Select Same Na&me"
         Height          =   375
         Index           =   1
         Left            =   4080
         TabIndex        =   11
         Top             =   3600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkNot 
         Caption         =   "NO&T"
         Height          =   255
         Left            =   4080
         TabIndex        =   7
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ListBox lstList 
         Columns         =   2
         Height          =   3960
         Index           =   0
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.ComboBox cboSelectOnly 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtReport 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         Height          =   3810
         Left            =   6000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "frmExchange.frx":058A
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton cmdSelectAssoc 
         Caption         =   "S&elect Associated"
         Height          =   375
         Left            =   4080
         TabIndex        =   8
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Select &All"
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdSelectOnly 
         Caption         =   "Select &Only:"
         Height          =   375
         Left            =   4080
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdSelectNone 
         Caption         =   "Select &None"
         Height          =   375
         Left            =   4080
         TabIndex        =   12
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   1
         Left            =   3600
         Picture         =   "frmExchange.frx":0702
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Caption         =   "&Characters"
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
         Left            =   480
         TabIndex        =   2
         Top             =   120
         Width           =   3015
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "frmExchange.frx":0C8C
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblAssociated 
         Alignment       =   2  'Center
         Caption         =   "(Player, Items, Rotes, Locations, Actions)"
         Height          =   405
         Left            =   4080
         TabIndex        =   9
         Top             =   2625
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Data..."
      Height          =   375
      Left            =   4440
      TabIndex        =   16
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load Data..."
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   5520
      Width           =   1695
   End
   Begin MSComctlLib.TabStrip tabTabs 
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8916
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   10
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Game Settings"
            Key             =   "game"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Players"
            Key             =   "players"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Characters"
            Key             =   "characters"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Items"
            Key             =   "items"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Rotes"
            Key             =   "rotes"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Locations"
            Key             =   "locations"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Actions"
            Key             =   "actions"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Plots"
            Key             =   "plots"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Rumors"
            Key             =   "rumors"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Searches"
            Key             =   "searches"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   7080
      TabIndex        =   17
      Top             =   5520
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cmnDialog 
      Left            =   6360
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "gex"
      Filter          =   "Player/Character Exchange Files (*.gex)|*.gex|All Files (*.*)|*.*"
   End
End
Attribute VB_Name = "frmExchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LIST_GAME = 1
Private Const LIST_PLAYERS = 2
Private Const LIST_CHARACTERS = 3
Private Const LIST_ITEMS = 4
Private Const LIST_ROTES = 5
Private Const LIST_LOCATIONS = 6
Private Const LIST_ACTIONS = 7
Private Const LIST_PLOTS = 8
Private Const LIST_RUMORS = 9
Private Const LIST_SEARCHES = 10
Private Const MIN_LIST = 1
Private Const MAX_LIST = 10

Private Const SAME_DATE = 0
Private Const SAME_NAME = 1

Private ThisList As Integer                         'Current selected list

Private Sub PopulateList(Index As Integer)
'
' Name:         PopulateList
' Parameters:   Index       ID of the list to populate
' Description:  Populate a list box with the appropriate contents.
'

    Dim List As LinkedList
    
    lstList(Index).Clear
    
    Select Case Index
        Case LIST_GAME
            lstList(Index).AddItem "Game Calendar"
            lstList(Index).AddItem "Action/Rumor Settings"
            lstList(Index).AddItem "XP and PP Awards"
            lstList(Index).AddItem "Template Settings"
        Case LIST_PLAYERS:      Set List = PlayerList
        Case LIST_CHARACTERS:   Set List = CharacterList
        Case LIST_ITEMS:        Set List = ItemList
        Case LIST_ROTES:        Set List = RoteList
        Case LIST_LOCATIONS:    Set List = LocationList
        Case LIST_ACTIONS:      Set List = ActionList
        Case LIST_PLOTS:        Set List = PlotList
        Case LIST_RUMORS:       Set List = RumorList
        Case LIST_SEARCHES:     Set List = Game.QueryEngine.QueryList
    End Select
    
    If Not List Is Nothing Then
        With List
            .First
            Do Until .Off
                If Not (.Item.Name = RecentSearchName And Index = LIST_SEARCHES) Then
                    lstList(Index).AddItem .Item.Name
                End If
                .MoveNext
            Loop
        End With
    End If
    
End Sub

Private Sub ReportCounts()
'
' Name:         ReportCounts
' Description:  Report the count of selections in txtReport.
'

    txtReport.Alignment = 0
    txtReport.Text = "Selected:" & vbCrLf
    txtReport.Text = txtReport.Text & "     " & CStr(lstList(1).SelCount) & " Game Settings" & vbCrLf
    txtReport.Text = txtReport.Text & "     " & CStr(lstList(2).SelCount) & " Players" & vbCrLf
    txtReport.Text = txtReport.Text & "     " & CStr(lstList(3).SelCount) & " Characters" & vbCrLf
    txtReport.Text = txtReport.Text & "     " & CStr(lstList(4).SelCount) & " Items" & vbCrLf
    txtReport.Text = txtReport.Text & "     " & CStr(lstList(5).SelCount) & " Rotes" & vbCrLf
    txtReport.Text = txtReport.Text & "     " & CStr(lstList(6).SelCount) & " Locations" & vbCrLf
    txtReport.Text = txtReport.Text & "     " & CStr(lstList(7).SelCount) & " Actions" & vbCrLf
    txtReport.Text = txtReport.Text & "     " & CStr(lstList(8).SelCount) & " Plots" & vbCrLf
    txtReport.Text = txtReport.Text & "     " & CStr(lstList(9).SelCount) & " Rumors" & vbCrLf
    txtReport.Text = txtReport.Text & "     " & CStr(lstList(10).SelCount) & " Searches"

End Sub

Private Function CreateSelectedList(Index As Integer) As LinkedList
'
' Name:         CreateSelectedList
' Parameter:    Index           ID of the list to create
' Description:  Return a new list full of the items selected from the given list.
'

    Dim InvList As LinkedList
    Dim SSet As StringSet
    
    Set SSet = New StringSet
    Set CreateSelectedList = New LinkedList

    Select Case Index
        Case LIST_GAME
            CreateSelectedList.Append IIf(lstList(LIST_GAME).Selected(0), 1, 0)
            CreateSelectedList.Append IIf(lstList(LIST_GAME).Selected(1), 1, 0)
            CreateSelectedList.Append IIf(lstList(LIST_GAME).Selected(2), 1, 0)
            CreateSelectedList.Append IIf(lstList(LIST_GAME).Selected(3), 1, 0)
        Case LIST_PLAYERS:      Set InvList = PlayerList
        Case LIST_CHARACTERS:   Set InvList = CharacterList
        Case LIST_ITEMS:        Set InvList = ItemList
        Case LIST_ROTES:        Set InvList = RoteList
        Case LIST_LOCATIONS:    Set InvList = LocationList
        Case LIST_ACTIONS:      Set InvList = ActionList
        Case LIST_PLOTS:        Set InvList = PlotList
        Case LIST_RUMORS:       Set InvList = RumorList
        Case LIST_SEARCHES:     Set InvList = Game.QueryEngine.QueryList
    End Select

    If Not InvList Is Nothing Then
        SSet.StoreListBox lstList(Index)
        InvList.First
        Do Until InvList.Off
            If SSet.Has(InvList.Item.Name) Then CreateSelectedList.Append InvList.Item
            InvList.MoveNext
        Loop
    End If

    Set SSet = Nothing

End Function

Private Sub SelectList(ByRef Box As ListBox, ToSelect As Boolean)
'
' Name:         Selectlist
' Parameters:   Box         the listbox to (de)select
'               ToSelect    whether to select or deselect
' Description:  Select or deselect all items in a listview.
'

    Dim I As Long
    
    For I = 0 To Box.ListCount - 1
        Box.Selected(I) = ToSelect
    Next I

End Sub

Private Sub cmdClose_Click()
'
' Name:         cmdClose_Click
' Description:  Dismiss this window.
'

    Unload Me

End Sub

Private Sub cmdLoad_Click()
'
' Name:         cmdLoad_Click
' Description:  Prompt the user for an exchange file, then load data from it.
'

    Dim FileError As Boolean
    Dim I As Integer
    Dim J As Integer
    
    cmnDialog.DialogTitle = "Load Exchange File"
    cmnDialog.InitDir = GetSetting(App.Title, "Files", "ExchangeDir", CurDir)
    cmnDialog.FileName = ""
    cmnDialog.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist
    cmnDialog.Filter = "Grapevine Exchange Files (*.gex)|*.gex|All Files|*.*"
    cmnDialog.FilterIndex = 1
    
    On Error GoTo cmdLoad_AnyError
    cmnDialog.ShowOpen
    On Error GoTo 0

    Screen.MousePointer = vbHourglass
    
    SaveSetting App.Title, "Files", "ExchangeDir", CurDir
    Game.LoadExchange cmnDialog.FileName
    
    If Game.FileError Then
        MsgBox Game.FileErrorMessage, vbExclamation, "Load Exchange File"
    End If
    
    For I = MIN_LIST To MAX_LIST
        PopulateList I
    Next I
    
    ReportCounts
    
    Screen.MousePointer = vbDefault
    
    If Not Game.FileError Then
        frmMergeResults.ShowResults Game.MergeResults, Me
    End If
    
    GoTo cmdLoad_Finish
    
cmdLoad_AnyError:
    Resume cmdLoad_Finish
cmdLoad_Finish:

End Sub

Private Sub cmdSave_Click()
'
' Name:         cmdSave_Click
' Description:  Prompt the user for a filename, then save the selected players
'               and characters to it.
'

    Dim MainList As LinkedList

    cmnDialog.DialogTitle = "Save Exchange File As..."
    cmnDialog.InitDir = GetSetting(App.Title, "Files", "ExchangeDir", CurDir)
    cmnDialog.FileName = "Exchange.gex"
    cmnDialog.Flags = cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNOverwritePrompt
    cmnDialog.Filter = "Grapevine Exchange File [Binary Format] (*.gex)|*.gex|" & _
                       "Grapevine Exchange File [XML Format] (*.gex)|*.gex|All Files|*.*"
    cmnDialog.FilterIndex = IIf(Game.FileFormat = gvXML, 2, 1)
    
    On Error GoTo cmdSave_AnyError
    cmnDialog.ShowSave
    On Error GoTo 0
    
    Screen.MousePointer = vbHourglass
    
    SaveSetting App.Title, "Files", "ExchangeDir", CurDir
    SaveSetting App.Title, "Settings", "FilterExchange", chkHideST.Value
    
    Set MainList = New LinkedList

    MainList.Append CreateSelectedList(LIST_GAME)
    MainList.Append CreateSelectedList(LIST_PLAYERS)
    MainList.Append CreateSelectedList(LIST_CHARACTERS)
    MainList.Append CreateSelectedList(LIST_SEARCHES)
    MainList.Append CreateSelectedList(LIST_ITEMS)
    MainList.Append CreateSelectedList(LIST_ROTES)
    MainList.Append CreateSelectedList(LIST_LOCATIONS)
    MainList.Append CreateSelectedList(LIST_ACTIONS)
    MainList.Append CreateSelectedList(LIST_PLOTS)
    MainList.Append CreateSelectedList(LIST_RUMORS)

    Game.SaveExchange cmnDialog.FileName, IIf(cmnDialog.FilterIndex = 1, gvBinaryExchange, gvXML), _
                      MainList, (chkHideST.Value = vbChecked)
    
    MainList.Clear
    Set MainList = Nothing
    
    Screen.MousePointer = vbDefault
    
    If Game.FileError Then _
            MsgBox Game.FileErrorMessage, vbExclamation, "Save Exchange File"
    
    GoTo cmdSave_Finish
    
cmdSave_AnyError:
    Resume cmdSave_Finish
cmdSave_Finish:

End Sub

Private Sub cmdSelectAll_Click()
'
' Name:         cmdSelectAll
' Description:  Select all items in the list.
'
    SelectList lstList(ThisList), True
    ReportCounts
    
End Sub

Private Sub cmdSelectAssoc_Click()
'
' Name:         cmdSelectAssoc_Click
' Description:  Select the entities associated with these.
'

    Dim FirstSet As StringSet
    Dim SecondSet As StringSet
    Dim ThirdSet As StringSet
    Dim FourthSet As StringSet
    Dim FifthSet As StringSet
    Dim SixthSet As StringSet
    Dim TempList As LinkedTraitList
    
    Screen.MousePointer = vbHourglass
    
    Set FirstSet = New StringSet
    Set SecondSet = New StringSet
    Set ThirdSet = New StringSet
    Set FourthSet = New StringSet
    Set FifthSet = New StringSet
    Set SixthSet = New StringSet
    
    FirstSet.StoreListBox lstList(ThisList)
    
    Select Case ThisList
        
        Case LIST_PLAYERS
            CharacterList.First
            Do Until CharacterList.Off
                If FirstSet.Has(CharacterList.Item.Player) Then SecondSet.Add CharacterList.Item.Name
                CharacterList.MoveNext
            Loop
            SecondSet.SelectListBox lstList(LIST_CHARACTERS), True, False
        
        Case LIST_CHARACTERS
            
            CharacterList.First
            Do Until CharacterList.Off
                If FirstSet.Has(CharacterList.Item.Name) Then
                    SecondSet.Add CharacterList.Item.Player
                    Set TempList = CharacterList.Item.EquipmentList
                    TempList.First
                    Do Until TempList.Off
                        ThirdSet.Add TempList.Trait.Name
                        TempList.MoveNext
                    Loop
                    Set TempList = CharacterList.Item.HangoutList
                    TempList.First
                    Do Until TempList.Off
                        FourthSet.Add TempList.Trait.Name
                        TempList.MoveNext
                    Loop
                    If CharacterList.Item.RaceCode = gvracemage Then
                        Set TempList = CharacterList.Item.RoteList
                        TempList.First
                        Do Until TempList.Off
                            FifthSet.Add TempList.Trait.Name
                            TempList.MoveNext
                        Loop
                    End If
                    Game.APREngine.MoveToFirstTitle ActionList, CharacterList.Item.Name
                    Do Until ActionList.Off
                        SixthSet.Add ActionList.Item.Name
                        Game.APREngine.MoveToNextTitle ActionList, CharacterList.Item.Name
                    Loop
                End If
                CharacterList.MoveNext
            Loop
            SecondSet.SelectListBox lstList(LIST_PLAYERS), True, False
            ThirdSet.SelectListBox lstList(LIST_ITEMS), True, False
            FourthSet.SelectListBox lstList(LIST_LOCATIONS), True, False
            FifthSet.SelectListBox lstList(LIST_ROTES), True, False
            SixthSet.SelectListBox lstList(LIST_ACTIONS), True, False
            
        Case LIST_ACTIONS
        
            Dim Action As ActionClass
            ActionList.First
            Do Until ActionList.Off
                If FirstSet.Has(ActionList.Item.Name) Then
                    Set Action = ActionList.Item
                    SecondSet.Add Action.Name
                    Action.First
                    Do Until Action.Off
                        With Action.SubAction.Causes
                            .First
                            Do Until .Off
                                Select Case .Link.Target
                                    Case aprAction: SecondSet.Add .Link.TargetName
                                    Case aprPlot:   ThirdSet.Add .Link.TargetName
                                    Case aprRumor:  FourthSet.Add .Link.TargetName
                                End Select
                                .MoveNext
                            Loop
                        End With
                        With Action.SubAction.Effects
                            .First
                            Do Until .Off
                                Select Case .Link.Target
                                    Case aprAction: SecondSet.Add .Link.TargetName
                                    Case aprPlot:   ThirdSet.Add .Link.TargetName
                                    Case aprRumor:  FourthSet.Add .Link.TargetName
                                End Select
                                .MoveNext
                            Loop
                        End With
                        Action.MoveNext
                    Loop
                End If
                ActionList.MoveNext
            Loop
            SecondSet.SelectListBox lstList(LIST_ACTIONS), True, False
            ThirdSet.SelectListBox lstList(LIST_PLOTS), True, False
            FourthSet.SelectListBox lstList(LIST_RUMORS), True, False
        
        Case LIST_PLOTS
            
            Dim Plot As PlotClass
            PlotList.First
            Do Until PlotList.Off
                If FirstSet.Has(PlotList.Item.Name) Then
                    Set Plot = PlotList.Item
                    ThirdSet.Add Plot.Name
                    Plot.First
                    Do Until Plot.Off
                        With Plot.PlotDev.Causes
                            .First
                            Do Until .Off
                                Select Case .Link.Target
                                    Case aprAction: SecondSet.Add .Link.TargetName
                                    Case aprPlot:   ThirdSet.Add .Link.TargetName
                                    Case aprRumor:  FourthSet.Add .Link.TargetName
                                End Select
                                .MoveNext
                            Loop
                        End With
                        With Plot.PlotDev.Effects
                            .First
                            Do Until .Off
                                Select Case .Link.Target
                                    Case aprAction: SecondSet.Add .Link.TargetName
                                    Case aprPlot:   ThirdSet.Add .Link.TargetName
                                    Case aprRumor:  FourthSet.Add .Link.TargetName
                                End Select
                                .MoveNext
                            Loop
                        End With
                        Plot.MoveNext
                    Loop
                End If
                PlotList.MoveNext
            Loop
            SecondSet.SelectListBox lstList(LIST_ACTIONS), True, False
            ThirdSet.SelectListBox lstList(LIST_PLOTS), True, False
            FourthSet.SelectListBox lstList(LIST_RUMORS), True, False
        
        Case LIST_RUMORS
            
            Dim Rumor As RumorClass
            RumorList.First
            Do Until RumorList.Off
                If FirstSet.Has(RumorList.Item.Name) Then
                    Set Rumor = RumorList.Item
                    FourthSet.Add Rumor.Name
                    Rumor.First
                    Do Until Rumor.Off
                        With Rumor.SubRumor.Causes
                            .First
                            Do Until .Off
                                Select Case .Link.Target
                                    Case aprAction: SecondSet.Add .Link.TargetName
                                    Case aprPlot:   ThirdSet.Add .Link.TargetName
                                    Case aprRumor:  FourthSet.Add .Link.TargetName
                                End Select
                                .MoveNext
                            Loop
                        End With
                        Rumor.MoveNext
                    Loop
                End If
                RumorList.MoveNext
            Loop
            SecondSet.SelectListBox lstList(LIST_ACTIONS), True, False
            ThirdSet.SelectListBox lstList(LIST_PLOTS), True, False
            FourthSet.SelectListBox lstList(LIST_RUMORS), True, False
            
    End Select
    
    ReportCounts
    
    Screen.MousePointer = vbDefault
    
    Set FirstSet = Nothing
    Set SecondSet = Nothing
    Set ThirdSet = Nothing
    Set FourthSet = Nothing
    Set FifthSet = Nothing
    Set SixthSet = Nothing
    
End Sub

Private Sub cmdSelectSame_Click(Index As Integer)
'
' Name:         cmdSelectSame_Click
' Description:  Select all actions or rumors that share their date or name with the
'               currently selected actions or rumors.
'

    Dim SSet As StringSet
    Dim AddSet As StringSet
    Dim Source As LinkedList
    Dim Match As String
    
    Set SSet = New StringSet
    Set AddSet = New StringSet
    
    SSet.StoreListBox lstList(ThisList)
    
    Select Case ThisList
        Case LIST_ACTIONS:      Set Source = ActionList
        Case LIST_RUMORS:       Set Source = RumorList
    End Select
    
    Source.First
    Do Until Source.Off
        If SSet.Has(Source.Item.Name) Then
            If Index = SAME_DATE Then
                If ThisList = LIST_ACTIONS Then
                    AddSet.Add CStr(Source.Item.ActDate)
                Else
                    AddSet.Add CStr(Source.Item.RumorDate)
                End If
            Else
                If ThisList = LIST_ACTIONS Then
                    AddSet.Add Source.Item.CharName
                Else
                    AddSet.Add Source.Item.Title
                End If
            End If
        End If
        Source.MoveNext
    Loop
    
    Source.First
    Do Until Source.Off
        Match = ""
        If Index = SAME_DATE Then
            If ThisList = LIST_ACTIONS Then
                Match = CStr(Source.Item.ActDate)
            Else
                Match = CStr(Source.Item.RumorDate)
            End If
        Else
            If ThisList = LIST_ACTIONS Then
                Match = Source.Item.CharName
            Else
                Match = Source.Item.Title
            End If
        End If
        If AddSet.Has(Match) Then SSet.Add Source.Item.Name
        Source.MoveNext
    Loop
    
    SSet.SelectListBox lstList(ThisList), True, False
    
    Set SSet = Nothing
    Set AddSet = Nothing
    
End Sub

Private Sub cmdSelectNone_Click()
'
' Name:         cmdSelectNone
' Description:  Deselect all items in the list.
'
    
    SelectList lstList(ThisList), False
    ReportCounts
    
End Sub

Private Sub cmdSelectOnly_Click()
'
' Name:         cmdSelectOnly_Click
' Description:  Select only those names that match a given query.
'

    Dim SSet As StringSet
    Dim Q As QueryClass
    
    Set SSet = New StringSet
    
    With Game.QueryEngine
        .QueryList.MoveTo cboSelectOnly.Text
        If .QueryList.Off Then
            Set Q = New QueryClass
            Select Case ThisList
                Case LIST_CHARACTERS:   Q.Inventory = qiCharacters
                Case LIST_PLAYERS:      Q.Inventory = qiPlayers
            End Select
        Else
            Set Q = .QueryList.Item
        End If
        
        .MakeQuery Q, , (chkNot.Value = vbChecked)
        
        .Results.First
        Do Until .Results.Off
            SSet.Add .Results.Item.Name
            .Results.MoveNext
        Loop
    End With
    
    SSet.SelectListBox lstList(ThisList), True, False
    ReportCounts
    
    Set SSet = Nothing
    Set Q = Nothing
    
End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Populate the lists.
'

    Load lstList(LIST_GAME)
    Load lstList(LIST_PLAYERS)
    Load lstList(LIST_CHARACTERS)
    Load lstList(LIST_ITEMS)
    Load lstList(LIST_ROTES)
    Load lstList(LIST_LOCATIONS)
    Load lstList(LIST_ACTIONS)
    Load lstList(LIST_PLOTS)
    Load lstList(LIST_RUMORS)
    Load lstList(LIST_SEARCHES)

    lstList(LIST_ACTIONS).Columns = 1
    lstList(LIST_RUMORS).Columns = 1

    PopulateList LIST_GAME
    PopulateList LIST_PLAYERS
    PopulateList LIST_CHARACTERS
    PopulateList LIST_ITEMS
    PopulateList LIST_ROTES
    PopulateList LIST_LOCATIONS
    PopulateList LIST_ACTIONS
    PopulateList LIST_PLOTS
    PopulateList LIST_RUMORS
    PopulateList LIST_SEARCHES

    Set tabTabs.SelectedItem = tabTabs.Tabs(LIST_CHARACTERS)
    lstList(LIST_CHARACTERS).Visible = True
    ThisList = LIST_CHARACTERS
    
    chkHideST.Value = GetSetting(App.Title, "Settings", "FilterExchange", vbUnchecked)
    
End Sub

Private Sub lstList_Click(Index As Integer)
'
' Name:         lstList_Click
' Description:  Adjust the count when an item is selected or unselected.
'

    ReportCounts

End Sub

Private Sub tabTabs_Click()
'
' Name:         tabTabs_Click
' Description:  Show the needed ListView and format the controls appropriately.
'

    Dim SelectOnlyType As QueryInventoryType
    Dim AssocStr As String
    Dim IconKey As String
    
    lstList(ThisList).Visible = False
    ThisList = tabTabs.SelectedItem.Index
    lstList(ThisList).Visible = True
    
    SelectOnlyType = qiNone
    
    cmdSelectSame(SAME_DATE).Visible = False
    
    Select Case ThisList
        Case LIST_GAME
            IconKey = "Calendar"
        Case LIST_PLAYERS
            AssocStr = "(Characters)"
            SelectOnlyType = qiPlayers
            IconKey = "Players"
        Case LIST_CHARACTERS
            AssocStr = "(Player, Items, Rotes, Locations, Actions)"
            SelectOnlyType = qiCharacters
            IconKey = "Masks"
        Case LIST_ITEMS
            IconKey = "Stake"
        Case LIST_ROTES
            IconKey = "Mage"
        Case LIST_LOCATIONS
            IconKey = "Lantern"
        Case LIST_ACTIONS
            AssocStr = "(Plots, Rumors)"
            IconKey = "Action"
            cmdSelectSame(SAME_DATE).Visible = True
        Case LIST_PLOTS:
            AssocStr = "(Actions, Rumors)"
            IconKey = "Plot"
        Case LIST_RUMORS:
            AssocStr = "(Actions, Plots)"
            IconKey = "Rumor"
            cmdSelectSame(SAME_DATE).Visible = True
        Case LIST_SEARCHES
            IconKey = "Search"
    End Select

    lblTitle.Caption = "&" & tabTabs.SelectedItem.Caption
    imgIcon(0).Picture = mdiMain.imlSmallIcons.ListImages(IconKey).Picture
    imgIcon(1).Picture = imgIcon(0).Picture

    cmdSelectAssoc.Visible = Not (AssocStr = "")
    lblAssociated.Caption = AssocStr
    lblAssociated.Visible = Not (AssocStr = "")
    cmdSelectSame(SAME_NAME).Visible = cmdSelectSame(SAME_DATE).Visible

    If SelectOnlyType = qiNone Then
        cmdSelectOnly.Visible = False
        cboSelectOnly.Visible = False
        chkNot.Visible = False
    Else
        cboSelectOnly.Clear
        With Game.QueryEngine.QueryList
            .First
            Do Until .Off
                If .Item.Inventory = SelectOnlyType Then cboSelectOnly.AddItem .Item.Name
                .MoveNext
            Loop
        End With
        cmdSelectOnly.Visible = True
        cboSelectOnly.Visible = True
        chkNot.Visible = True
        If cboSelectOnly.ListCount > 0 Then cboSelectOnly.ListIndex = 0
    End If

End Sub
