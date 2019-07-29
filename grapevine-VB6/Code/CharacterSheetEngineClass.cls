VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CharacterSheetEngineClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum TargetClassType 'For indicating the present target of the engine
    ttNothing = 0
    ttListBox = 1
    ttLabel = 2
    ttXPHistory = 3
End Enum

Private Const gvMinListIndex = 0
Private Const gvMaxListIndex = 20

Public TargetType As TargetClassType    'The type of display to adjust
Public TargetLabel As Label             'a target label control
Public TargetList As ListBox            'a target list box

Private CharacterRace As String         'the character's race
Private MenuStack As LinkedList         'a stack of the menus traversed
Private MenuList As ListBox             'the list box that holds menus
Private MenuCaption As Label            'the caption label under that list box
Private MenuTitle As Label              'the title label above that list box
Private SubMenuLabelStack As LinkedList 'menu labels corresponding to the menu stack

Private PurchaseList As LinkedTraitList 'List of all purchases made
Private RefundList As LinkedTraitList   'List of all refunds made
Private CostEstimate As Single          'The estimated cost of these transactions
Private RandNum As Long                 'User-specified random number

'
' Arrays of references to Trait list boxes and 'Regular' List boxes
'
Private TraitListWithIndex(gvMinListIndex To gvMaxListIndex) As LinkedTraitList

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Create the stacks.
'
    
    Set MenuStack = New LinkedList
    Set SubMenuLabelStack = New LinkedList
    Set PurchaseList = New LinkedTraitList
    Set RefundList = New LinkedTraitList
    PurchaseList.Initialize "", True, False, False, ldMultiplier
    RefundList.Initialize "", True, False, False, ldMultiplier
    RandNum = 1
    
End Sub

Public Sub RegisterSheet(CRace As String, MList As ListBox, MCaption As Label, MTitle As Label)
'
' Name:         RegisterSheet
' Parameters:   CRace           the race of this character
'               MList           the listbox that displays menus
'               MCaptions       the label that captions the menu box
'               MTitle          the label that titles the menu box
' Description:  Supply data to this engine that will enable it to work with a
'               character sheet window.
'

    CharacterRace = CRace
    Set MenuList = MList
    Set MenuCaption = MCaption
    Set MenuTitle = MTitle
    
End Sub

Public Sub RegisterTraitList(Index As Integer, TraitList As LinkedTraitList)
'
' Name:         RegisterTraitList
' Parameters:   Index           index of this trait list
'               TraitList       a character class's LinkedTraitList
' Description:  Associate an index with a list of trait data from a
'               character class, such as a PhysicalTraitList or
'               AbilityList.
'

    Set TraitListWithIndex(Index) = TraitList

End Sub

Public Sub RefreshTraitList(List As ListBox, TraitList As LinkedTraitList)
'
' Name:         RefreshTraitList
' Parameters:   List            a ListBox to refresh
'               TraitList       the list of trait data with which to refresh it
' Description:  Clears a list box and refills it from the appropriate trait
'               list.
'

    Dim StorePos As Integer
    Dim ListPos As Integer
    
    StorePos = List.ListIndex
    ListPos = 0
    
    TraitList.First
    Do Until TraitList.Off Or ListPos = List.ListCount
        List.List(ListPos) = TraitList.DisplayTrait
        ListPos = ListPos + 1
        TraitList.MoveNext
    Loop

    If TraitList.Off Then   'remove extra list items
        Do Until List.ListCount = TraitList.EntryCount
            List.RemoveItem ListPos
        Loop
    Else                    'add more list items
        Do Until TraitList.Off
            List.AddItem TraitList.DisplayTrait
            TraitList.MoveNext
        Loop
    End If

    ' This is a hack.  A little hack, but a hack.  I know these lists are always
    ' associated with a corresponding lblTraits() label.
    UpdateTraitLabel List.Parent.lblTraits(List.Index), TraitList.Count
    
    If StorePos >= List.ListCount Then StorePos = List.ListCount - 1
    List.ListIndex = StorePos

End Sub

Private Sub RemoveFromList(List As ListBox)
'
' Name:         RemoveFromList
' Parameters:   List            the list from which to remove an item
' Description:  Removes the presently selected item from the character
'               list and updates the list box
'

    Dim TraitList As LinkedTraitList

    Set TraitList = TraitListWithIndex(List.Index)
    TraitList.MoveToPlace List.ListIndex
    Refund TraitList.Trait.Name, IIf(TraitList.Atomic, TraitList.Trait.Number, 1) _
                                     * IIf(TraitList.Negative, -1, 1)
    TraitList.Remove
    RefreshTraitList List, TraitList

End Sub

Private Sub AddToList(List As ListBox, AddText As String, AddNumber As String, AddNote As String)
'
' Name:         AddToList
' Parameters:   List        a ListBox to which to add an item
'               AddText     the item to add
'               AddNumber   its number / cost
'               AddNote     its note
' Description:  Adds a given item to a list box and records it in
'               the character class.
'

    Dim TraitList As LinkedTraitList
    Dim ListPos As Integer
    Dim FullText As String
    Dim MustSort As Boolean
    
    Set TraitList = TraitListWithIndex(List.Index)
        
    FullText = ""
    SubMenuLabelStack.First
    Do Until SubMenuLabelStack.Off
        FullText = FullText & SubMenuLabelStack.Item & ": "
        If FullText = ": " Then FullText = ""
        SubMenuLabelStack.MoveNext
    Loop
    FullText = FullText & AddText
    
    '
    ' We must sort only if 1) the list is neither alphabetized nor empty and
    ' 2) it's either atomic OR doesn't contain the item.
    '
    MustSort = Not (TraitList.IsAlphabetized Or TraitList.IsEmpty)
    If MustSort Then
        TraitList.MoveTo FullText
        MustSort = TraitList.Atomic Or TraitList.Off
    End If
    
    If MustSort Then
    
        Dim OrderText As String
        Dim Menu As LinkedMenuList
        Dim Absent As Boolean
        Dim PreText As String
        Dim Suffix As String
        
        Absent = False
        SubMenuLabelStack.First
        MenuStack.First
        TraitList.First
        
        '
        ' Example:      Inserting "Thaumaturgy: Path of Blood: Blood of Potency"
        '               MenuStack and Traitlist are guaranteed non-empty.
        '
        '      MenuStack       SubmenuLabelStack   Traitlist
        '       Disciplines     Thaumaturgy         Auspex: Heightened Senses
        '       Thaumaturgy     Path of Blood       Auspex: Aura Perception
        '       Path of Blood                       Thaumaturgy: Movement of the Mind: Force Bolt
        '                                           Thaumaturgy: Path of Blood: Taste of Blood
        '                                           Thaumaturgy: Path of Blood: Blood Rage
        '                                           Thaumaturgy: Path of Blood: Theft of Vitae
        '                                           Thaumaturgy: Lure of Flames: Hand of Flame
        '                                           Vissicitude: Malleable Visage
        '
        
        Do Until MenuStack.Off                                  'At least once.
            
            If SubMenuLabelStack.Off Then                       'For the last item,
                OrderText = AddText                             ' "Blood of Potency" is the text.
                Suffix = ""
            Else                                                'Otherwise it's the name of
                OrderText = SubMenuLabelStack.Item              ' the submenu -- "Thaumaturgy"
                SubMenuLabelStack.MoveNext                      ' or "Path of Blood".
                Suffix = ": *"
            End If
            
            Set Menu = MenuStack.Item
            Menu.First
            Do Until Menu.ItemName = OrderText                  'Look through the menu until the
                                                                'order text is found.
                                                                
                Do Until Not TraitList.Trait.Name Like _
                            PreText & Menu.ItemName & Suffix    'For each nonmatching menu item,
                    TraitList.MoveNext                          ' move past similar traits.
                    If TraitList.Off Then                       ' if you leave the trait list this
                        Absent = True                           ' way, you're done.
                        Exit Do
                    End If
                Loop
                
                Menu.MoveNext
                Absent = Absent Or Menu.Off
                
                If Absent Then Exit Do

            Loop
                        
            MenuStack.MoveNext
            If Absent Then Exit Do
            PreText = PreText & _
                    OrderText & ": "
            
        Loop
    
    End If
    
    TraitList.Insert FullText, AddNumber, AddNote
    Purchase FullText, Val(AddNumber) * IIf(TraitList.Negative, -1, 1)
    RefreshTraitList List, TraitList
    
    TraitList.MoveTo FullText
    For ListPos = 0 To List.ListCount - 1
        If List.List(ListPos) = TraitList.DisplayTrait Then
            List.ListIndex = ListPos
            Exit For
        End If
    Next ListPos
    
End Sub

Public Sub UpdateTraitLabel(TraitLabel As Label, Total As Double)
'
' Name:         UpdateTraitLabel
' Parameters:   TraitLabel      The trait label to change
'               Total           Number of Traits to report
' Description:  Puts the correct number and plural form into the caption
'               of a trait list box.  The singular and plural forms are
'               stored in the .Tag of the label, in the form
'               "plural|singular".
'

    Dim PipeLoc As Integer

    ' Do not operate on labels without .Tags; their captions are static.
    If TraitLabel.Tag <> "" Then
        
        PipeLoc = InStr(TraitLabel.Tag, "|")
        TraitLabel.Caption = CStr(Total) & " "
        
        ' If there is a pipe in the tag, there is a plural form.
        If PipeLoc > 0 Then
            
            ' If the singular is needed, grab everything right of the pipe.
            ' Otherwise grab everything left of it.
            If Total = 1 Then
                TraitLabel.Caption = TraitLabel.Caption & _
                        Right(TraitLabel.Tag, Len(TraitLabel.Tag) - PipeLoc)
            Else
                TraitLabel.Caption = TraitLabel.Caption & _
                        Left(TraitLabel.Tag, PipeLoc - 1)
            End If
            
        Else
            ' No special plural form.
            TraitLabel.Caption = TraitLabel.Caption & TraitLabel.Tag
        End If
    
    End If

End Sub

Public Function DisplayTemper(Perm As Integer, Temp As Integer) As String
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

Public Sub UpdateMenuTitleFromTraitLabel(TraitLabel As Label)
'
' Name:         UpdateMenuTitleFromTraitLabel
' Parameters:   TraitList       the trait label to base the change upon
' Description:  Set the menu title based on a label's .Tag and .Caption.
'

    Dim PipeLoc As Integer

    If TraitLabel.Tag = "" Then
    
        ' Without a tag, the menu title gets the same caption as the label
        MenuTitle.Caption = TraitLabel.Caption
    
    Else
        
        ' Otherwise, get the plural form from the tag.  The whole tag is used
        ' when there is no plural/singular syntax.
        PipeLoc = InStr(TraitLabel.Tag, "|")
        If PipeLoc = 0 Then PipeLoc = Len(TraitLabel.Tag) + 1
        MenuTitle.Caption = Left(TraitLabel.Tag, PipeLoc - 1)
           
    End If

End Sub

Public Sub DecrementEntry()
'
' Name:         DecrementEntry
' Description:  Decrement the current entry of the current target list.
'               For Trait Lists this is like RemoveEntry; for regular lists this
'               will drop the cost of a merit.

    If TargetType = ttListBox Then
        If TargetList.ListIndex > -1 Then
                
            Dim TraitList As LinkedTraitList

            Set TraitList = TraitListWithIndex(TargetList.Index)
            TraitList.MoveToPlace TargetList.ListIndex
            
            If Not TraitList.Atomic Then Refund TraitList.Trait.Name, IIf(TraitList.Negative, -1, 1)
            
            TraitList.Decrement
            RefreshTraitList TargetList, TraitList
            Game.DataChanged = True
            
        End If
    End If
    
End Sub

Public Sub IncrementEntry()
'
' Name:         IncrementEntry
' Description:  Increment the current entry of the current target list.
'               For Trait Lists this is like remove; for regular lists this
'               will drop the cost of a merit.

    If TargetType = ttListBox Then
        If TargetList.ListIndex > -1 Then
                
            Dim TraitList As LinkedTraitList

            Set TraitList = TraitListWithIndex(TargetList.Index)
            TraitList.MoveToPlace TargetList.ListIndex
            
            If Not TraitList.Atomic Then Purchase TraitList.Trait.Name, IIf(TraitList.Negative, -1, 1)
            
            TraitList.Increment
            TargetList.List(TargetList.ListIndex) = TraitList.DisplayTrait
            UpdateTraitLabel TargetList.Parent.lblTraits(TargetList.Index), TraitList.Count
            Game.DataChanged = True
            
        End If
    End If

End Sub

Public Sub PopulateMenu(MenuName As String, Optional SubMenu As Boolean = False, _
        Optional SubMenuLabel = "")
'
' Name:         PopulateMenu
' Parameters:   MenuName    the name of the menu
'               SubMenu     whether this is itself a submenu
' Description:  Fill the menu list box with the items from the named menu.
'

    Dim SameMenu As Boolean

    SameMenu = Not MenuStack.IsEmpty
    If SameMenu Then
        MenuStack.Last
        SameMenu = (MenuName = MenuStack.Item.Name)
    End If
    
    If Not SameMenu Then
        
        Dim Menu As LinkedMenuList
        
        ' Find the correct menu: First try special menus, then try race-specific menus,
        ' then last try the menu by literal name.
        
        Set Menu = Game.MenuSet.GetSpecialMenu(MenuName)
        If Menu Is Nothing Then
            If Not SubMenu Then                 'Only top-Level Menus can be Race-Specific
                Set Menu = Game.MenuSet.GetCompleteMenu(MenuName & ", " & CharacterRace)
            End If
            If Menu Is Nothing Then
                Set Menu = Game.MenuSet.GetCompleteMenu(MenuName)
            End If
        Else
            SubMenuLabel = ""
        End If
        
        If Menu Is Nothing Then
            If Not SubMenu Then MenuList.Clear
            MenuCaption.Caption = "The menu """ & MenuName & """ doesn't exist!"
        Else
                
            MenuList.Clear
            MenuCaption.Caption = ""
            
            If SubMenu Then
                SubMenuLabelStack.Append SubMenuLabel
                MenuList.AddItem "(back)"
            Else
                SubMenuLabelStack.Clear
                MenuStack.Clear
            End If
        
            MenuStack.Append Menu
            Menu.First
            Do Until Menu.Off
                MenuList.AddItem Menu.DisplayItem
                Menu.MoveNext
            Loop

            Set Menu = Nothing
            
        End If

    End If

End Sub

Public Sub AddSelected()
'
' Name:         AddSelected
' Description:  Add a trait to the target list or label if the selection is a string;
'               or fill the menu if the selection is a list;
'               or return to the previous menu if the selection is (back).
'

    Dim Menu As LinkedMenuList
    Dim TreeText As String
    
    If MenuList.Text = "(back)" Then
    
        MenuStack.Last
        MenuStack.Remove
        SubMenuLabelStack.Last
        
        MenuStack.Last
        Set Menu = MenuStack.Item
        
        MenuList.Clear
        If MenuStack.Count > 1 Then MenuList.AddItem "(back)"
        
        Menu.First
        Do Until Menu.Off
            MenuList.AddItem Menu.DisplayItem
            If Menu.ItemName = SubMenuLabelStack.Item Then _
                    MenuList.ListIndex = MenuList.NewIndex
            Menu.MoveNext
        Loop
    
        SubMenuLabelStack.Remove
    
    Else
    
        MenuStack.Last
        Set Menu = MenuStack.Item
        
        If MenuStack.Count > 1 Then
            Menu.MoveToPlace MenuList.ListIndex - 1
        Else
            Menu.MoveToPlace MenuList.ListIndex
        End If
        
        If Menu.ItemCost = ":" Then  'It's a submenu!
            
            PopulateMenu Menu.ItemNote, True, Menu.ItemName
            
        Else

            Select Case TargetType
                Case ttLabel

                    TreeText = ""
                    SubMenuLabelStack.First
                    Do Until SubMenuLabelStack.Off
                        TreeText = TreeText & SubMenuLabelStack.Item & ": "
                        If TreeText = ": " Then TreeText = ""
                        SubMenuLabelStack.MoveNext
                    Loop
                    
                    If Menu.Autonote And Menu.ItemNote <> "" Then
                        TargetLabel.Caption = TreeText & Menu.ItemName & _
                                            " (" & Menu.ItemNote & ")"
                    Else
                        TargetLabel.Caption = TreeText & Menu.ItemName
                    End If
                
                Case ttListBox
                    If Menu.Autonote And Menu.ItemNote <> "" Then
                        AddToList TargetList, Menu.ItemName, _
                                    Menu.ItemCost, Menu.ItemNote
                    Else
                        AddToList TargetList, Menu.ItemName, _
                                    Menu.ItemCost, ""
                    End If
            
            End Select
        
            Game.DataChanged = True
        
        End If
      
    End If

End Sub

Public Sub AddCustom()
'
' Name:         AddCustom
' Description:  Add a custom entry to the target list or label.  Custom text added to
'               lists may not contain the patterns "* (*)" or "* x*#".
'
    
    Dim CustomName As String
    Dim CustomNumber As String
    Dim CustomNote As String
    Dim RecordInMenu As Boolean
    
    Dim IsTrait As Boolean
    Dim Display As ListDisplayType
    Dim AllowAdd As Boolean
    
    Dim CustomText As String
    Dim TraitList As LinkedTraitList
    Dim NewTrait As TraitClass
    
    If TargetType <> ttNothing Then
    
        If TargetType = ttLabel Then
            IsTrait = False
            Display = ldNoteOnly
            AllowAdd = Left(TargetLabel.Tag, 1) <> "?"
        Else
            IsTrait = True
            Display = TraitListWithIndex(TargetList.Index).Display
            AllowAdd = Left(TargetList.Tag, 1) <> "?"
        End If
            
        frmCustom.GetCustom IsTrait, Display, AllowAdd, CustomName, CustomNumber, CustomNote, RecordInMenu
        
        If CustomName <> "" Then
            
            Select Case TargetType
                
                Case ttLabel
                    CustomText = CustomName
                    If CustomNote <> "" Then CustomText = CustomText & " (" & CustomNote & ")"
                    TargetLabel.Caption = CustomText
                    Select Case TargetLabel.Tag
                        Case "?EQ", "?RO", "?LO", "?PL", "?NR"
                            Call CrossReference(True)
                    End Select
                    
                Case ttListBox
                    
                    AddToList TargetList, CustomName, CustomNumber, CustomNote
            
                    Select Case TargetList.Tag
                        Case "?EQ", "?RO", "?LO", "?PL", "?NR"
                            Call CrossReference(True)
                    End Select
                        
            End Select
            
            If RecordInMenu Then
                
                Dim Menu As LinkedMenuList
                
                MenuStack.Last
                Set Menu = Game.MenuSet.GetMenu(MenuStack.Item.Name)
                
                If Not Menu Is Nothing Then
                    
                    Menu.Append CustomName, CStr(CustomNumber), CustomNote
                    Game.MenuSet.DataChanged = True
                    
                    If Menu.HasInclusion Then
                        Set Menu = MenuStack.Item
                        Menu.Append CustomName, CStr(CustomNumber), CustomNote
                    End If
                    
                    MenuList.Clear
                    If MenuStack.Count > 1 Then MenuList.AddItem "(back)"
                    
                    Menu.First
                    Do Until Menu.Off
                        MenuList.AddItem Menu.DisplayItem
                        Menu.MoveNext
                    Loop
                    
                End If
            End If
            
            Game.DataChanged = True
        
        End If
        
    End If

End Sub


Public Sub AddNote()
'
' Name:         AddNote
' Description:  Add a note to a label or list entry.  For labels and regular lists, this
'               is string manipulation; for traits, the note is added to the trait's
'               .Note property.
'
    
    Dim Note As String
    Dim Original As String
    Dim MainPart As String
    Dim OldNote As String
    Dim ParenPlace As Integer
    
    Dim StorePos As Integer
    Dim Entry As String
    Dim TraitList As LinkedTraitList
    
    If TargetType <> ttNothing Then
        
        Select Case TargetType
            
            Case ttLabel
                Original = TargetLabel.Caption
                ParenPlace = InStr(Original, " (")
                If ParenPlace = 0 Then
                    ParenPlace = Len(Original) + 1
                    OldNote = ""
                Else
                    OldNote = Right(Original, Len(Original) - ParenPlace - 1)
                    OldNote = Left(OldNote, Len(OldNote) - 1)
                End If
                MainPart = Left(Original, ParenPlace - 1)
            
            Case ttListBox
            
                StorePos = TargetList.ListIndex
                If StorePos = -1 Then Exit Sub
                
                Set TraitList = TraitListWithIndex(TargetList.Index)
                TraitList.MoveToPlace StorePos
                OldNote = TraitList.Trait.Note
                TraitList.Trait.Note = ""
                MainPart = TraitList.DisplayTrait
                TraitList.Trait.Note = OldNote
            
        End Select
        
        Note = InputBox("Enter the note for """ & _
                MainPart & """, or press Cancel for no note.", _
                "Add Note to Entry", OldNote)
        Note = Trim(Note)
        
        If Note <> OldNote Then
        
            Select Case TargetType
            
                Case ttLabel
                    
                    If Note <> "" Then
                        TargetLabel.Caption = MainPart & " (" & Note & ")"
                    Else
                        TargetLabel.Caption = MainPart
                    End If
                    Game.DataChanged = True
                    
                Case ttListBox
                
                    TraitList.Trait.Note = Note
                    TargetList.List(StorePos) = TraitList.DisplayTrait
                    Game.DataChanged = True
        
            End Select
    
        End If

    End If

End Sub

Public Sub RemoveEntry()
'
' Name:         RemoveEntry
' Description:  Remove a label or list entry.
'
    
    Dim StorePos As Integer
    
    Select Case TargetType
        Case ttLabel
            TargetLabel.Caption = ""
            Game.DataChanged = True
        Case ttListBox
            StorePos = TargetList.ListIndex
            If StorePos <> -1 Then RemoveFromList TargetList
            If StorePos >= TargetList.ListCount Then StorePos = TargetList.ListCount - 1
            TargetList.ListIndex = StorePos
            Game.DataChanged = True
    End Select

End Sub

Public Function CanAdjust(Index As Integer) As Boolean
'
' Name:         CanAdjust
' Parameters:   Index       the list we're interested in adjusting
' Returns:      TRUE iff the corresponding list has a visible Total;
'               i.e., its Display isn't Simple or NoteOnly.
'

    With TraitListWithIndex(Index)
        CanAdjust = Not (.Display = ldSimple Or .Display = ldNoteOnly)
    End With
    
End Function

Public Function CanOrder(Index As Integer) As Boolean
'
' Name:         CanOrder
' Parameters:   Index       the list we're interested in ordering
' Returns:      TRUE iff the corresponding list is not alphabetized.
'

    CanOrder = Not TraitListWithIndex(Index).IsAlphabetized
    
End Function

Public Sub MoveEntryUp()
'
' Name:         MoveEntryUp
' Description:  Move the selected entry up in the list.
'

    Dim StorePos As Integer

    If TargetType = ttListBox Then
        StorePos = TargetList.ListIndex
        If StorePos > -1 Then
                
            Dim TraitList As LinkedTraitList

            Set TraitList = TraitListWithIndex(TargetList.Index)
            If Not TraitList.IsAlphabetized Then
                TraitList.MoveToPlace TargetList.ListIndex
                TraitList.SwapFrontward
                RefreshTraitList TargetList, TraitList
                If StorePos > 0 Then TargetList.ListIndex = StorePos - 1
                Game.DataChanged = True
            End If
            
        End If
    End If
    
End Sub

Public Sub MoveEntryDown()
'
' Name:         MoveEntryDown
' Description:  Move the selected entry down in the list.
'

    Dim StorePos As Integer

    If TargetType = ttListBox Then
        StorePos = TargetList.ListIndex
        If StorePos > -1 Then
                
            Dim TraitList As LinkedTraitList

            Set TraitList = TraitListWithIndex(TargetList.Index)
            If Not TraitList.IsAlphabetized Then
                TraitList.MoveToPlace TargetList.ListIndex
                TraitList.SwapBackward
                RefreshTraitList TargetList, TraitList
                If StorePos < TargetList.ListCount - 1 Then _
                    TargetList.ListIndex = StorePos + 1
                Game.DataChanged = True
            End If
            
        End If
    End If
    
End Sub

Public Sub RefreshXP(HistoryList As ListView, XP As ExperienceClass)
'
' Name:         RefreshXP
' Parameters:   HistoryList         listview of the XP history
'               XP                  The character's Experience object
' Description:  Repopulate the XP History list.
'

    Dim StringDate As String
    Dim NewItem As ListItem
    Dim SaveLoc As Integer
    
    With HistoryList
        SaveLoc = -1
        If Not .SelectedItem Is Nothing Then SaveLoc = .SelectedItem.Index
        .ListItems.Clear
    End With
    
    With XP
        If HistoryList.SortOrder = lvwDescending Then
            .Last
        Else
            .First
        End If
        Do Until .Off
        
            StringDate = Format(.EntryDate, "Short Date")
            Set NewItem = HistoryList.ListItems.Add(, , StringDate)
            NewItem.ListSubItems.Add , "Change", .EntryChangeText
            NewItem.ListSubItems.Add , "Reason", .EntryReason
            NewItem.ListSubItems.Add , "Unspent", .EntryUnspent
            NewItem.ListSubItems.Add , "Earned", .EntryEarned
            
            If HistoryList.SortOrder = lvwDescending Then
                .MovePrevious
            Else
                .MoveNext
            End If
        Loop
    End With
    
    With HistoryList
        If SaveLoc > -1 And .ListItems.Count > 0 Then
            If SaveLoc > .ListItems.Count Then SaveLoc = .ListItems.Count
            Set .SelectedItem = .ListItems(SaveLoc)
        End If
    End With
    
End Sub

Public Function AddXPEntry(HistoryList As ListView, XP As ExperienceClass) As Boolean
'
' Name:         AddXPEntry
' Parameters:   HistoryList         listview of the XP history
'               XP                  The character's Experience object
' Description:  Add an XP entry from the current menu, returning TRUE iff addition
'               was not canceled.
'

    Dim CType As ExperienceChangeType
    
    Select Case MenuList.Text
        Case "New Entry"
            frmHistoryEntry.MakeEntry XP, True, "Add History Entry"
        Case "Recent Purchases"
            CType = IIf(CostEstimate < 0, ecUnspent, ecSpent)
            frmHistoryEntry.MakeEntry XP, True, "Add History Entry", _
                                      GetTransactionReason, CType, Abs(CostEstimate), True
            If Not frmHistoryEntry.Canceled Then ClearTransactions
        Case Else
        
            With Game.XPAwardList
                .MoveTo MenuList.Text
                If Not .Off Then
                    frmHistoryEntry.MakeEntry XP, True, "Add History Entry", _
                        .Item.Reason, .Item.ChangeType, .Item.Change
                End If
            End With
            
    End Select
        
    AddXPEntry = Not frmHistoryEntry.Canceled

End Function

Public Function EditXPEntry(HistoryList As ListView, XP As ExperienceClass) As Boolean
'
' Name:         EditXPEntry
' Parameters:   HistoryList         listview of the XP history
'               XP                  The character's Experience object
' Description:  Edit an XP entry, returning TRUE iff editing was not canceled.
'

    EditXPEntry = False
    If Not HistoryList.SelectedItem Is Nothing Then
        If HistoryList.SortOrder = lvwDescending Then
            XP.MoveToPlace HistoryList.ListItems.Count - HistoryList.SelectedItem.Index
        Else
            XP.MoveToPlace HistoryList.SelectedItem.Index - 1
        End If
        frmHistoryEntry.MakeEntry XP, False, "Edit History Entry"
        EditXPEntry = Not frmHistoryEntry.Canceled
    End If

End Function

Public Function RemoveXPEntry(HistoryList As ListView, XP As ExperienceClass) As Boolean
'
' Name:         RemoveXPEntry
' Parameters:   HistoryList         listview of the XP history
'               XP                  The character's Experience object
' Description:  Confirm and remove an experience history entry, returning TRUE iff not canceled.
'

    RemoveXPEntry = False
    If Not HistoryList.SelectedItem Is Nothing Then
        If MsgBox("Are you sure you want to remove this history entry?", vbYesNo, _
                "Remove Entry") = vbYes Then
                                            
            If HistoryList.SortOrder = lvwDescending Then
                XP.MoveToPlace HistoryList.ListItems.Count - HistoryList.SelectedItem.Index
            Else
                XP.MoveToPlace HistoryList.SelectedItem.Index - 1
            End If
            XP.Remove
            RemoveXPEntry = True
            
        End If
    End If
        
End Function

Public Sub CrossReference(Optional CreateOnly As Boolean = False)
'
' Name:         CrossReference
' Parameters:   CreateOnly      if TRUE, don't display a card unless the user wants to create one
' Description:  Cross-Reference items or characters double-clicked on the current target.
'

    Dim Reference As String
    Dim Tag As String
    Dim TraitList As LinkedTraitList
    
    If TargetType = ttLabel Then
        Reference = TargetLabel.Caption
        Tag = TargetLabel.Tag
    ElseIf TargetType = ttListBox Then
        If TargetList.ListIndex > -1 Then
            Tag = TargetList.Tag
            Set TraitList = TraitListWithIndex(TargetList.Index)
            TraitList.MoveToPlace TargetList.ListIndex
            Reference = TraitList.Trait.Name
        End If
    End If
    
    If Reference <> "" Then
        Select Case Tag
        
            Case "?BB", "?GX", "?CH"
                If Not CreateOnly Then mdiMain.ShowCharacterSheet Reference
                
            Case "?EQ"
                ItemList.MoveTo Reference
                If Not ItemList.Off Then
                    If Not CreateOnly Then mdiMain.ShowItemCard Reference
                Else
                    If MsgBox("You don't have an item card for """ _
                              & Reference & """." & vbCrLf & _
                              "Would you like to create one?", _
                              vbYesNo + vbQuestion, "New Item Card") = vbYes Then
                        
                        Dim ItemCard As ItemClass
                        Set ItemCard = New ItemClass
                        ItemCard.Name = Reference
                        ItemList.InsertSorted ItemCard
                        Game.DataChanged = True
                        mdiMain.ShowItemCard Reference
                        Set ItemCard = Nothing
                        mdiMain.AnnounceChanges mdiMain, atItems
                        
                    End If
                End If
                
            Case "?RO"
                RoteList.MoveTo Reference
                If Not RoteList.Off Then
                    If Not CreateOnly Then mdiMain.ShowRote Reference
                Else
                    If MsgBox("You don't have a rote prepared for """ _
                              & Reference & """." & vbCrLf & _
                              "Would you like to create one?", _
                              vbYesNo + vbQuestion, "New Rote") = vbYes Then
                        
                        Dim RoteCard As RoteClass
                        Set RoteCard = New RoteClass
                        RoteCard.Name = Reference
                        RoteList.InsertSorted RoteCard
                        Game.DataChanged = True
                        mdiMain.ShowRote Reference
                        Set RoteCard = Nothing
                        mdiMain.AnnounceChanges mdiMain, atRotes
                        
                    End If
                End If
            
            Case "?LO"
                LocationList.MoveTo Reference
                If Not LocationList.Off Then
                    If Not CreateOnly Then mdiMain.ShowLocation Reference
                Else
                    If MsgBox("You don't have a location described for """ _
                              & Reference & """." & vbCrLf & _
                              "Would you like to create one?", _
                              vbYesNo + vbQuestion, "New Location") = vbYes Then
                        
                        Dim LocationCard As LocationClass
                        Set LocationCard = New LocationClass
                        LocationCard.Name = Reference
                        LocationList.InsertSorted LocationCard
                        Game.DataChanged = True
                        mdiMain.ShowLocation Reference
                        Set LocationCard = Nothing
                        mdiMain.AnnounceChanges mdiMain, atLocations
                        
                    End If
                End If
            
            Case "?PL", "?NR"
                PlayerList.MoveTo Reference
                If Not PlayerList.Off Then
                    If Not CreateOnly Then mdiMain.ShowPlayer Reference
                Else
                    If MsgBox("You don't have information entered for """ _
                              & Reference & """." & vbCrLf & _
                              "Would you like to add it?", _
                              vbYesNo + vbQuestion, "New Player") = vbYes Then
                        
                        Dim PlayerCard As PlayerClass
                        Set PlayerCard = New PlayerClass
                        PlayerCard.Name = Reference
                        PlayerList.InsertSorted PlayerCard
                        Game.DataChanged = True
                        mdiMain.ShowPlayer Reference
                        Set PlayerCard = Nothing
                        mdiMain.AnnounceChanges mdiMain, atPlayers
                        
                    End If
                End If
                
        End Select
    End If

End Sub

Public Sub DeselectMenus()
'
' Name:         DeselectMenus
' Description:  Empty the stacks and clear the display: nothing is selected.
'

    MenuCaption.Caption = ""
    MenuTitle.Caption = ""
    MenuList.Tag = ""
    MenuList.Clear
    MenuStack.Clear
    SubMenuLabelStack.Clear

End Sub

Public Sub Purchase(Item As String, Cost As Single)
'
' Name:         Purchase
' Parameters:   Item        what is purchased
'               Cost        how much it costs
' Description:  Add an item to the list of recent purchases.
'

    RefundList.MoveTo Item
    If RefundList.Off Then
        PurchaseList.Insert Item, "1"
    Else
        RefundList.Remove
    End If
    CostEstimate = CostEstimate + Cost

End Sub

Public Sub Refund(Item As String, Cost As Single)
'
' Name:         Refund
' Parameters:   Item        what is refunded
'               Cost        how much it costs
' Description:  Add an item to a list of recent refunds.
'

    PurchaseList.MoveTo Item
    If PurchaseList.Off Then
        RefundList.Insert Item, "1"
    Else
        PurchaseList.Remove
    End If
    CostEstimate = CostEstimate - Cost

End Sub

Public Function GetCostEstimate() As Single
'
' Name:         GetCostEstimate
' Description:  Get the current estimate of transaction costs.
'

    GetCostEstimate = CostEstimate

End Function

Public Function GetTransactionReason() As String
'
' Name:         GetTransactionReason
' Description:  Construct a transaction reason from the purchase and refund lists.
'

    Dim Reason As String

    If Not PurchaseList.IsEmpty Then
    
        Reason = "Purchased "
        PurchaseList.First
        Do Until PurchaseList.Off
            Reason = Reason & PurchaseList.DisplayTrait
            PurchaseList.MoveNext
            If Not PurchaseList.Off Then Reason = Reason & ", "
        Loop
        
        Reason = Reason & "." & IIf(RefundList.IsEmpty, "", " ")
        
    End If
    
    If Not RefundList.IsEmpty Then
        
        Reason = Reason & "Removed "
        RefundList.First
        Do Until RefundList.Off
            Reason = Reason & RefundList.DisplayTrait
            RefundList.MoveNext
            If Not RefundList.Off Then Reason = Reason & ", "
        Loop
        Reason = Reason & "."
        
    End If

    GetTransactionReason = Reason

End Function

Public Sub ClearTransactions()
'
' Name:         ClearTransactions
' Description:  Empty the purchase and refund lists.
'

    PurchaseList.Clear
    RefundList.Clear
    CostEstimate = 0

End Sub

Public Sub PopUpTraitListMenu(Source As Form, MyList As ListBox)
'
' Name:         PopUpTraitListMenu
' Parameters:   Source          Source Form
'               MyList          List being targeted for the popup
' Description:  Pop Up a menu of extra options for the current trait list.
'

    If TargetType = ttListBox And Not TargetList Is Nothing Then
    
        Dim PopupList As StringSet
        Dim Response As String
        Dim MyTraitList As LinkedTraitList
        Dim I As Long
        
        Set PopupList = New StringSet
        Set MyTraitList = TraitListWithIndex(MyList.Index)
        
        PopupList.Add "Add random entries"
        If CanAdjust(MyList.Index) Then PopupList.Add "Randomize these entries"
        PopupList.Add "Change display"
        PopupList.Add "Turn alphabetization " & IIf(MyTraitList.IsAlphabetized, "off", "on")
        PopupList.Add "Make entries " & IIf(MyTraitList.Atomic, "", "in") & "divisible"
        PopupList.Add "Remove all"
        
        Response = mdiMain.CreatePopup(PopupList, Source)
        
        Select Case Response
            Case "Change display"
                PopupList.Clear
                PopupList.Add "0 - Simple"
                PopupList.Add "1 - Multiplier, Note"
                PopupList.Add "2 - Multiplier, Dots, Note"
                PopupList.Add "3 - Dots, Note"
                PopupList.Add "4 - Cost, Note"
                PopupList.Add "5 - Note Only"
                PopupList.Add "6 - Cost Only"
                PopupList.Add "(Current Display: " & CStr(MyTraitList.Display) & ")"
                Response = mdiMain.CreatePopup(PopupList, Source)
                If Left(Response, 1) Like "#" Then
                    MyTraitList.Display = CLng(Left(Response, 1))
                    RefreshTraitList MyList, MyTraitList
                End If
                
            Case "Add random entries"
                Response = InputBox("How many random entries would you like to add?", "Add Random", RandNum)
                I = Int(Val(Response))
                If I > 0 Then
                    RandNum = I
                    Dim TempList As LinkedTraitList
                    Set TempList = New LinkedTraitList
                    TempList.Atomic = MyTraitList.Atomic
                    Game.MenuSet.AddToListRandomly RandNum, TempList, MyList.Tag, CharacterRace
                    With TempList
                        .First
                        Do Until .Off
                            With .Trait
                                MyTraitList.Insert .Name, .Total, .Note
                                If MyTraitList.Atomic Then
                                    Purchase .Name, .Number * IIf(MyTraitList.Negative, -1, 1)
                                Else
                                    For I = 1 To .Number
                                        Purchase .Name, 1 * IIf(MyTraitList.Negative, -1, 1)
                                    Next I
                                End If
                            End With
                            .MoveNext
                        Loop
                    End With
                    Set TempList = Nothing
                    RefreshTraitList MyList, MyTraitList
                End If
                
            Case "Randomize these entries"
                Response = InputBox("How many total traits would you like the list to have?", _
                           "Randomize Entries", MyTraitList.TraitCount)
                I = Int(Val(Response))
                If I > 0 Then
                    MyTraitList.RandomizeCurrentTraits I
                    RefreshTraitList MyList, MyTraitList
                End If
                
            Case "Turn alphabetization on", "Turn alphabetization off"
                MyTraitList.SetAlphabetized (Not MyTraitList.IsAlphabetized)
                RefreshTraitList MyList, MyTraitList
                
            Case "Make entries indivisible", "Make entries divisible"
                MyTraitList.Atomic = (Not MyTraitList.Atomic)
                
            Case "Remove all"
                If MsgBox("Are you sure you want to remove all traits from this list?", _
                          vbYesNo, "Clear List") = vbYes Then
                    With MyTraitList
                        .First
                        Do Until .IsEmpty
                            If .Atomic Then
                                Refund .Trait.Name, .Trait.Number * IIf(.Negative, -1, 1)
                            Else
                                For I = 1 To .Trait.Number
                                    Refund .Trait.Name, IIf(.Negative, -1, 1)
                                Next I
                            End If
                            .RemoveTrait
                        Loop
                    End With
                    RefreshTraitList MyList, MyTraitList
                End If
                
        End Select
        
        Set PopupList = Nothing
    
    End If

End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Destroy the stacks.
'

    Set MenuStack = Nothing
    Set SubMenuLabelStack = Nothing
    Set PurchaseList = Nothing
    Set RefundList = Nothing
    
End Sub
