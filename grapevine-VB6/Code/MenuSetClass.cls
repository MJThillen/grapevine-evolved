VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuSetClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FilePath As String                   'Full Path of this menu's file
Public FileName As String                   'Short Filename of this menu's file
Public Description As String                'Description of this file
Public DataChanged As Boolean               'whether data has changed since last save

Public FileFormat As FileFormatType         'Current format of the menu file
Public FileError As Boolean                 'Whether a file error has occurred
Public FileErrorMessage As String           'Description of the error

Public FileProgress As ProgressBar          'Control describing progress of load

Private MenuCollection As Collection        'Collection (Hash Set) of LinkedMenuLists
Private Cursor As Long                      'location in collection for traversal

Private Sub AddFileProgress(Addition As Single)
'
' Name:         AddFileProgress
' Parameter:    Addition        amount to add to the FileProgress bar
' Description:  Increment FileProgress and make sure it doesn't go over its maximum and cause an
'               error.
'

    FileProgress.Value = IIf(FileProgress.Value + Addition > FileProgress.Max, _
                             FileProgress.Max, FileProgress.Value + Addition)

End Sub

Public Function GetMenu(Name As String) As LinkedMenuList
'
' Name:         GetMenu
' Parameters:   Name        name of the menu to get
' Description:  Return a reference to the named menu.  This does not
'               move the cursor -- it bypasses the linked implementation of
'               the MenuSet by using a hashed set that indexes all the
'               menus by name.
'
    
    On Error Resume Next
    Set GetMenu = Nothing
    Set GetMenu = MenuCollection.Item(Name)

End Function

Public Function Count() As Integer
'
' Name:         Count
' Description:  Returns the number of menus in the menu set
' Returns:      the number of menus in the menu set
'

    Count = MenuCollection.Count

End Function

Public Function AddNewMenu(MenuName As String) As LinkedMenuList
'
' Name:         AddNewMenu
' Parameters:   MenuName        the name of the menu to insert.
' Description:  create a menu and add it to the menu set.
' Returns:      a reference to the new menu.
'

    Dim NewMenu As LinkedMenuList
    
    Set NewMenu = New LinkedMenuList
    NewMenu.Name = MenuName
    AddMenu NewMenu
    Set AddNewMenu = NewMenu

End Function

Public Sub AddMenu(NewMenu As LinkedMenuList)
'
' Name:         AddMenu
' Parameters:   NewMenu        the menu to insert.
' Description:  insert a menu alphabetically by key into the menu set.  If this
'               menu's name is in use, this method does nothing.
'

    Dim I As Integer
    
    I = MenuCollection.Count
    
    If I = 0 Then
        MenuCollection.Add NewMenu, NewMenu.Name
    Else
    
        Do Until I = 0
            If LCase(NewMenu.Name) > LCase(MenuCollection.Item(I).Name) Then Exit Do
            I = I - 1
        Loop
        
        On Error Resume Next    'Do not add duplicate menus
        If I = 0 Then
            MenuCollection.Add NewMenu, NewMenu.Name, 1
        Else
            MenuCollection.Add NewMenu, NewMenu.Name, , I
        End If
    End If

End Sub

Public Sub RemoveMenu(Name As String)
'
' Name:         RemoveMenu
' Parameter:    Name        Name of menu to remove
' Description:  Totally remove the given menu.
'

    On Error Resume Next
    MenuCollection.Remove Name

End Sub

Public Sub SetMenuName(OldName As String, NewName As String)
'
' Name:         SetMenuName
' Parameters:   OldName             the old name
'               NewName             the new name for the menu
' Description:  Change a menu's name.  Nothing happens if the new name is in use already.
'
    
    Dim OldMenu As LinkedMenuList
    
    Set OldMenu = GetMenu(OldName)
    
    If GetMenu(NewName) Is Nothing And Not OldMenu Is Nothing Then
        MenuCollection.Remove OldName
        OldMenu.Name = NewName
        AddMenu OldMenu
    End If
    
End Sub

Public Function IsEmpty() As Boolean
'
' Name:         IsEmpty
' Description:  Tells whether the list is empty or not.
' Returns:      TRUE if the list is empty, FALSE otherwise.
'

    IsEmpty = (MenuCollection.Count = 0)

End Function

Public Sub MoveTo(MenuName As String)
'
' Name:         MoveTo
' Parameters:   MenuName        a menu name to which to move
' Description:  Move the cursor to the menu whose name matches
'               the parameter.  Move off if no such item is found.
'

    Dim Found As Boolean

    Found = False
    First

    Do Until Off Or Found
        Found = (Menu.Name = MenuName)
        If Not Found Then MoveNext
    Loop

End Sub

Public Sub Clear()
'
' Name:         Clear
' Description:  Empty the list of all items.
'

    Do Until MenuCollection.Count = 0
        MenuCollection.Remove 1
    Loop

End Sub

Public Sub SaveMenus(FileName As String)
'
' Name:         SaveMenus
' Parameters:   FileName        the name under which to save the file
' Description:  Save the menus under a given filename.
'

    Select Case FileFormat
        Case gvBinaryMenu
            SaveMenusBinary FileName
        Case gvXML
            SaveMenusXML FileName
        Case Else
            FileError = True
            FileErrorMessage = "Can't save in file format " & CStr(FileFormat)
    End Select

End Sub

Private Sub SaveMenusBinary(FileName As String)
'
' Name:         SaveMenusBinary
' Parameters:   FileName        name of the file to save
' Description:  Save menu data to a binary file.
'

    Dim FileNum As Integer
    Dim aMenu As LinkedMenuList
    Dim OneBit As Single
    
    FileError = True
    FileNum = FreeFile
    
    On Error GoTo SMB_AnyError
    
    FileErrorMessage = "Error deleting old menu file " & ShortFile(FileName) & ":"
    If Dir(FileName) <> "" Then Kill FileName
    
    FileErrorMessage = "Error opening binary menu file " & ShortFile(FileName) & " for writing:"
    Open FileName For Binary As #FileNum
    
    FileErrorMessage = "Error writing binary file header:"
    PutStrB FileNum, BinHeaderMenu
    Put #FileNum, , ThisVersion
    PutStrB FileNum, Description
    Put #FileNum, , CInt(MenuCollection.Count)
            
    OneBit = 100 / (MenuCollection.Count + 1)
    FileProgress.Value = 0
    
    FileErrorMessage = "Error writing binary menus:"
    
    For Each aMenu In MenuCollection
        
        aMenu.OutputToBinary FileNum
        AddFileProgress OneBit
        
    Next aMenu
    
    FileError = False
    DataChanged = False
    FileFormat = gvBinaryMenu
    Me.FilePath = FileName
    Me.FileName = ShortFile(FileName)
    
    GoTo SMB_CloseFile

SMB_AnyError:
    FileErrorMessage = FileErrorMessage & vbCrLf & Err.Description
    Resume SMB_CloseFile

SMB_CloseFile:
    Close #FileNum

End Sub

Private Sub SaveMenusXML(FileName As String)
'
' Name:         SaveMenusXML
' Parameters:   FileName        the name under which to save the file
' Description:  Save the menus under a given filename in XML.
'

    Dim FileNum As Integer
    Dim aMenu As LinkedMenuList
    Dim XML As XMLWriterClass
    Dim OneBit As Single
    
    FileError = True
    FileNum = FreeFile
    Set XML = New XMLWriterClass
    
    On Error GoTo SMX_AnyError
    FileErrorMessage = "Failed to open " & ShortFile(FileName) & ":"
    Open FileName For Output As #FileNum
    
    OneBit = 100 / (MenuCollection.Count + 1)
    FileProgress.Value = 0
    
    With XML
    
        .BeginXML FileNum
    
        .BeginTag "grapevinemenus"
            .WriteAttribute "version", ThisVersion
            .WriteAttribute "size", MenuCollection.Count
            .WriteCDataTag "description", Description
            
            FileErrorMessage = "Error writing the menus to disk:"
            For Each aMenu In MenuCollection
                
                aMenu.OutputToFile XML
                AddFileProgress OneBit
                
            Next aMenu
    
        .EndTag
        
    End With
    
    DataChanged = False
    FileError = False
    FileFormat = gvXML
    Me.FilePath = FileName
    Me.FileName = ShortFile(FileName)
    
    GoTo SMX_CloseFile
    
SMX_AnyError:
    FileErrorMessage = FileErrorMessage & vbCrLf & Err.Description
    Resume SMX_CloseFile

SMX_CloseFile:
    Close #FileNum
    Set XML = Nothing

End Sub

Public Sub MergeMenus(FileName As String, Aggressive As Boolean)
'
' Name:         MergeMenus
' Parameters:   FileName        name of the file to open and merge
'               Aggressive      Whether to do an aggressive or conservative merge
' Description:  Open a second menu file, then reconcile its differences with the current menu.
'               Agressive mean give preference to the newcomer; Conservative means give
'               preference to the tried and true menu.
'

    Dim NewMenus As MenuSetClass
    Dim OneBit As Single
    Dim OldMenu As LinkedMenuList
    Dim Changes As String
    Dim FileNum As Integer
    
    FileError = True
    FileErrorMessage = ""
    
    Set NewMenus = New MenuSetClass
    
    Set NewMenus.FileProgress = FileProgress
    NewMenus.OpenMenus FileName, True
    
    If NewMenus.FileError Then
        FileErrorMessage = NewMenus.FileErrorMessage
    Else
    
        FileError = False
        OneBit = 100 / (NewMenus.Count + 1)
        Changes = "Merging the new file " & ShortFile(FileName) & _
                  " into the existing file " & ShortFile(Me.FileName) & vbCrLf
        Changes = Changes & IIf(Aggressive, "Aggressive", "Conservative") & " Merge: preference given to " & _
                  ShortFile(IIf(Aggressive, FileName, Me.FileName)) & vbCrLf & vbCrLf
        
        Description = Description & vbCrLf & vbCrLf & IIf(Aggressive, "Aggressively", "Conservatively") & _
                      " merged with " & ShortFile(FileName) & ":" & vbCrLf & NewMenus.Description
        
        NewMenus.First
        Do Until NewMenus.Off
            Set OldMenu = GetMenu(NewMenus.Menu.Name)
            If Not OldMenu Is Nothing Then
                If NewMenus.Menu.Category = gvMenuDelete Then
                    If Aggressive Then
                        Changes = Changes & """" & OldMenu.Name & """ Menu:" & vbCrLf & _
                                  "- Deleted Entire Menu" & vbCrLf & vbCrLf
                        RemoveMenu OldMenu.Name
                    End If
                Else
                    OldMenu.Merge NewMenus.Menu, Aggressive, Changes
                End If
            ElseIf Not NewMenus.Menu.Category = gvMenuDelete Then
                AddMenu NewMenus.Menu
                Changes = Changes & """" & NewMenus.Menu.Name & """ Menu:" & vbCrLf & _
                          "+ Added Entire Menu" & vbCrLf & vbCrLf
            End If
            AddFileProgress OneBit
            NewMenus.MoveNext
        Loop
    
        On Error Resume Next
        FileNum = FreeFile
        Open SlashPath(App.Path) & "MergeLog.txt" For Output As #FileNum
        Print #FileNum, Changes;
        Close #FileNum
        On Error GoTo 0
        
    End If
        
    Set NewMenus = Nothing
        
End Sub

Public Sub OpenMenus(FileName As String, ReOpen As Boolean)
'
' Name:         OpenMenus
' Parameters:   FileName        name of the file to open
'               Reopen          Whether to reopen a menu file that's already loaded
' Description:  Find the format of a file and call the appropriate load method.
'

    FileError = True
    FileErrorMessage = ""
    
    On Error GoTo OpenMenus_AnyError

    FileName = FindFile(FileName, SlashPath(App.Path) & DefaultMenuFile)
    If FileName = "" Then
        FileErrorMessage = "No suitable menu file -- not even the default " & DefaultMenuFile & _
               " -- could be found for use with this game."
        GoTo OpenMenus_EndSub
    End If

    If FileName <> FilePath Or ReOpen Then

        Select Case DetectFileFormat(FileName)
            Case gvBinaryMenu
                OpenMenusBinary FileName
            Case gvXML
                OpenMenusXML FileName
            Case gvBinaryGame
                FileErrorMessage = ShortFile(FileName) & _
                        " is a grapevine menu file: open it from File -> Open Game."
            Case gv23Game
                FileErrorMessage = ShortFile(FileName) & _
                        " is a grapevine menu file: open it from File -> Open Game."
            Case gvBinaryExchange
                FileErrorMessage = ShortFile(FileName) & _
                        " is a grapevine exchange file: open it from File -> Data Exchange."
            Case gv23Exchange
                FileErrorMessage = ShortFile(FileName) & _
                        " is a grapevine exchange file: open it from File -> Data Exchange."
            Case gvInvalid
                FileErrorMessage = ShortFile(FileName) & _
                        " is not a grapevine game file."
        End Select
        
    Else  'This menu is already loaded, and we don't intend to reload it
    
        FileError = False
        
    End If
    
    GoTo OpenMenus_EndSub
    
OpenMenus_AnyError:
    FileErrorMessage = "Error opening " & ShortFile(FileName) & ":" & vbCrLf & Err.Description
    Resume OpenMenus_EndSub
    
OpenMenus_EndSub:
    
End Sub

Private Sub OpenMenusBinary(FileName As String)
'
' Name:         OpenMenusBinary
' Parameters:   FileName        name of the file to open
' Description:  Clear all menus and load new data from a binary file.
'

    Dim FileNum As Integer
    Dim Verify As String
    Dim Version As Double
    Dim I As Integer
    Dim OneBit As Single
    Dim NewMenu As LinkedMenuList
    
    FileNum = FreeFile
    FileError = True
    
    On Error GoTo OMB_AnyError
    
    FileErrorMessage = "Error opening binary menu file " & ShortFile(FileName) & ":"
    Open FileName For Binary As #FileNum
    
    FileErrorMessage = "Error reading binary file header:"
    GetStrB FileNum, Verify
    If Verify = BinHeaderMenu Then
                
        Clear
        
        Get #FileNum, , Version
        GetStrB FileNum, Description
        Get #FileNum, , I
        
        OneBit = 100 / (I + 1)
        FileProgress.Value = 0
        
        FileErrorMessage = "Error reading binary menu:"
        
        Do Until I = 0
            
            Set NewMenu = New LinkedMenuList
            NewMenu.InputFromBinary FileNum, Version
            AddMenu NewMenu
            
            AddFileProgress OneBit
            I = I - 1
            
        Loop
    
        Me.FilePath = FileName
        Me.FileName = Mid(FilePath, InStrRev(FilePath, "\") + 1)
        FileFormat = gvBinaryMenu
        DataChanged = False
        FileError = False
    
    Else
        FileErrorMessage = ShortFile(FileName) & " is not a valid Grapevine menu file."
    End If
    
    GoTo OMB_CloseFile

OMB_AnyError:
    FileErrorMessage = FileErrorMessage & vbCrLf & Err.Description
    Resume OMB_CloseFile

OMB_CloseFile:
    Close #FileNum

End Sub

Private Sub OpenMenusXML(FileName As String)
'
' Name:         OpenMenus
' Parameters:   FileName        name of the file to open
' Description:  Clear all menus and load new data from a file.
'

    Dim FileNum As Integer
    Dim Read As String
    Dim Total As Integer
    Dim Version As Double
    Dim XML As XMLReaderClass
    Dim OneBit As Single
    
    Dim NewMenu As LinkedMenuList
    
    FileError = True
    FileNum = FreeFile
    Set XML = New XMLReaderClass
        
    '
    ' Open the file
    '
    On Error GoTo OMX_AnyError
    
    FileErrorMessage = "Failed to open " & ShortFile(FileName) & ":"
    Open FileName For Input As #FileNum
    
    With XML
    
        .BeginReading FileNum
        
        '
        ' Check the File
        '
        If Left(.VersionTag, 5) = "<?xml" Then
    
            .ReadTag
            If .Tag = "grapevinemenus" Then
            
                Clear
        
                Version = .GetDouble("version")
                OneBit = 100 / (.GetInteger("size", 650) + 1)
                FileProgress.Value = 0
                
                Do Until .Tag = "grapevinemenus" And .Terminal
                
                    .ReadTag
                    Select Case .Tag
                    
                        Case "description"
                            '
                            ' Read the description
                            '
                            FileErrorMessage = "Error reading description from file:"
                            Do Until .Tag = "description" And .Terminal
                                .ReadTag
                            Loop
                            Description = .GetTagText
                            
                        Case "menu"
                            '
                            ' Read the menus
                            '
                            FileErrorMessage = "Error reading a menu from file:"
                            Set NewMenu = New LinkedMenuList
                            NewMenu.InputFromFile XML, Version
                            AddMenu NewMenu
                            AddFileProgress OneBit
                            
                    End Select
                
                Loop
        
                Me.FilePath = FileName
                Me.FileName = ShortFile(FileName)
                FileFormat = gvXML
                DataChanged = False
                FileError = False
                
            Else ' Not a grapevinemenus tag
                FileErrorMessage = ShortFile(FileName) & " is not a Grapevine Menu File."
            End If
     
        Else 'not an XML file
            FileErrorMessage = ShortFile(FileName) & " is not a Grapevine Menu File."
        End If
    
    End With
    
    GoTo OMX_CloseFile
    
OMX_AnyError:
    FileErrorMessage = FileErrorMessage & vbCrLf & Err.Description
    Resume OMX_CloseFile

OMX_CloseFile:
    Close #FileNum
    Set XML = Nothing
    
End Sub

Public Function GetCompleteMenu(MenuName As String) As LinkedMenuList
' Name:         GetAssembledMenu
' Parameter:    MenuName        name of the menu to get and complete
' Description:  Return the named menu.  If it has inclusions, construct and return the
'               complete menu with the inclusions.
'

    Set GetCompleteMenu = GetMenu(MenuName)
    
    If Not GetCompleteMenu Is Nothing Then
        If GetCompleteMenu.HasInclusion Then
    
            Dim IncludedList As LinkedList
            Dim MadeMenu As LinkedMenuList
            
            Set MadeMenu = New LinkedMenuList
            Set IncludedList = New LinkedList
            
            MadeMenu.Name = GetCompleteMenu.Name
            MadeMenu.Autonote = GetCompleteMenu.Autonote
            MadeMenu.Display = GetCompleteMenu.Display
            MadeMenu.Negative = GetCompleteMenu.Negative
            MadeMenu.Category = GetCompleteMenu.Category
            MadeMenu.SetAlphabetized GetCompleteMenu.IsAlphabetized
            
            PopulateInclusion GetCompleteMenu, MadeMenu, IncludedList
            
            Set GetCompleteMenu = MadeMenu
            Set MadeMenu = Nothing
            Set IncludedList = Nothing
            
        End If
    End If

End Function

Private Sub PopulateInclusion(Menu As LinkedMenuList, MadeMenu As LinkedMenuList, _
        IncludedList As LinkedList)
'
' Name:         PopulateInclusion
' Parameters:   Menu            the current menu from which we're populating
'               MadeMenu        the linked menu list we're constructing
'               IncludedList    Menus that have come before
' Description:  Populate a menu using menu inclusions.  Make sure no included menu
'               gets in more than once.
'

    Menu.First
    Do Until Menu.Off
        If Menu.ItemCost = "+" Then
        
            IncludedList.MoveTo Menu.ItemNote
            If IncludedList.Off Then
                Dim NextMenu As LinkedMenuList
                Set NextMenu = GetMenu(Menu.ItemNote)
                If Not NextMenu Is Nothing Then
                    IncludedList.Append NextMenu.Name
                    PopulateInclusion NextMenu, MadeMenu, IncludedList
                End If
            End If
        
        Else
            MadeMenu.Append Menu.ItemName, Menu.ItemCost, Menu.ItemNote
        End If
        Menu.MoveNext
    Loop
    
End Sub

Public Function GetSpecialMenu(MenuName As String) As LinkedMenuList
'
' Name:         GetSpecialMenu
' Parameters:   MenuName        name of the menu to populate
' Description:  Creates the special "Players," "Characters," "Narrators"
'               "Items", etc. menus.
' Returns:      The created menu.
'

    If Left(MenuName, 1) = "?" Then
        
        Dim Sort As Boolean
        Dim Query As QueryClass
        Dim Cost As String
        Dim Note As String
        Dim ActivityLink As String
        
        Screen.MousePointer = vbHourglass
    
        Set GetSpecialMenu = New LinkedMenuList
        GetSpecialMenu.Autonote = True
        GetSpecialMenu.SetAlphabetized False
        Note = ""
        ActivityLink = ""
        
        Select Case Left(MenuName, 3)
                
            Case "?PL"
                
                Set Query = New QueryClass
                Query.Inventory = qiPlayers
                ActivityLink = "?PL:X"
                Cost = "0"
                
            Case "?BB"
                
                Set Query = New QueryClass
                Query.Inventory = qiCharacters
                Query.MatchAll = True
                Query.AddClause qkRace, "vampire", 0, qcEquals, False
                ActivityLink = "?BB:X"
                Cost = "1"
                           
            Case "?GX"
                
                Set Query = New QueryClass
                Query.Inventory = qiCharacters
                Query.MatchAll = True
                Query.AddClause qkRace, "kuei-jin", 0, qcEquals, False
                ActivityLink = "?GX:X"
                Cost = "1"
                           
            Case "?EQ"
                
                Dim IC As ItemClass
                Dim FoundTypes As String
                                
                ItemList.First
                FoundTypes = "/"
                Do Until ItemList.Off
                    Set IC = ItemList.Item
                    If InStr(FoundTypes, IC.ItemType & "/") = 0 Then
                        FoundTypes = FoundTypes & IC.ItemType & "/"
                        GetSpecialMenu.Append IC.ItemType, ":", "?E:" & IC.ItemType
                    End If
                    ItemList.MoveNext
                Loop
                                
            Case "?E:"
            
                Set Query = New QueryClass
                Query.Inventory = qiItems
                Query.AddClause qkType, Mid(MenuName, 4), 0, qcEquals, False
                Cost = "1"
                            
            Case "?LO"
            
                Set Query = New QueryClass
                Query.Inventory = qiLocations
                Cost = "1"
                            
            Case "?RO"
                
                Dim SphereMenu As LinkedMenuList
                
                GetSpecialMenu.Append "All Rotes", ":", "?R:"
                
                Set SphereMenu = Game.MenuSet.GetCompleteMenu("Spheres")
                If Not SphereMenu Is Nothing Then
                    With SphereMenu
                        .First
                        Do Until .Off
                            GetSpecialMenu.Append .ItemName, ":", "?R:" & .ItemName
                            .MoveNext
                        Loop
                    End With
                End If
                Set SphereMenu = Nothing

            Case "?R:"

                Set Query = New QueryClass
                Query.Inventory = qiRotes
                Query.SortKey = qkLevel
                Query.AddClause qkSpheres, Mid(MenuName, 4), 0, qcContains, False
                Cost = "1"
                GetSpecialMenu.SetAlphabetized False
                GetSpecialMenu.Display = ldNoteOnly
                Sort = True
            
            Case "?CH"
                
                Set Query = New QueryClass
                Query.Inventory = qiCharacters
                ActivityLink = "?CH:X"
                Cost = "1"
            
            Case "?MG"
                
                Set Query = New QueryClass
                Query.Inventory = qiCharacters
                Query.MatchAll = True
                Query.AddClause qkRace, "mage", 0, qcEquals, False
                ActivityLink = "?MG:X"
                Cost = "1"
            
            Case "?DT"
            
                With Game.Calendar
                    .Last
                    Do Until .Off
                        GetSpecialMenu.Append Format(.GetGameDate, "mmmm d, yyyy"), "", ""
                        .MovePrevious
                    Loop
                End With
            
            Case "?NR"
                                
                Set Query = New QueryClass
                Query.Inventory = qiPlayers
                Query.MatchAll = True
                Query.AddClause qkPosition, "player", 0, qcEquals, True
                ActivityLink = "?NR:X"
                Cost = "0"
            
            Case "?XP"
            
                GetSpecialMenu.Append "New Entry", "", ""
                GetSpecialMenu.Append "Recent Purchases", "", ""
                With Game.XPAwardList
                    .First
                    Do Until .Off
                        If .Item.XP Then GetSpecialMenu.Append .Item.Name, "", ""
                        .MoveNext
                    Loop
                End With
            
            Case "?PP"
            
                GetSpecialMenu.Append "New Entry", "", ""
                With Game.XPAwardList
                    .First
                    Do Until .Off
                        If Not .Item.XP Then GetSpecialMenu.Append .Item.Name, "", ""
                        .MoveNext
                    Loop
                End With
            
            Case Else
                
                Set GetSpecialMenu = Nothing
            
        End Select
                
        If Not Query Is Nothing Then
                
            If ActivityLink <> "" Then
                Query.AddClause qkPlayStatus, ActiveStatus, 0, qcEquals, (Mid(MenuName, 4) = ":X")
            End If
            
            With Game.QueryEngine
                .MakeQuery Query, Sort
            
                .Results.First
                Do Until .Results.Off
                    If Query.Inventory = qiItems Then Note = .Results.Item.CombatNote
                    If Query.Inventory = qiRotes Then Note = "Lv. " & CStr(.Results.Item.Level)
                    GetSpecialMenu.Append .Results.Item.Name, Cost, Note
                    .Results.MoveNext
                Loop
            End With
            
            If ActivityLink <> "" And Not (Mid(MenuName, 4) = ":X") Then
                GetSpecialMenu.Append "Not Active", ":", MenuName & ":X"
            End If
            
        End If
        
        Set Query = Nothing
        
        Screen.MousePointer = vbDefault
        
    End If

End Function

Public Sub AddToListRandomly(ByVal Total As Long, List As LinkedTraitList, MenuName As String, Race As String)
'
' Name:         AddToListRandomly
' Parameters:   Total           Number of Traits to add
'               List            Trait List to add to
'               MenuName        Name of menu to add from
'               Race            Name of character race (for race-specific menus)
' Description:  Randomly add traits from a menu to a list.
'

    Dim StartMenu As LinkedMenuList
    Dim WorkingMenu As LinkedMenuList
    Dim Pick As Integer
    Dim Cost As String
    Dim Cluster As Integer
    Dim Prefix As String
    Dim SubName As String
    
    Set StartMenu = GetSpecialMenu(MenuName)
    If StartMenu Is Nothing Then
        If Race <> "" Then Set StartMenu = GetCompleteMenu(MenuName & ", " & Race)
        If StartMenu Is Nothing Then Set StartMenu = GetCompleteMenu(MenuName)
    End If
    
    If Not StartMenu Is Nothing Then
        Do Until Total <= 0
        
            Set WorkingMenu = StartMenu
            Prefix = ""
            
            Do
                Pick = Int(Rnd * WorkingMenu.Count)
                WorkingMenu.MoveToPlace Pick
                If WorkingMenu.Off Then Exit Do
                If WorkingMenu.ItemCost <> ":" Then Exit Do
                Prefix = Prefix & WorkingMenu.ItemName & ": "
                SubName = WorkingMenu.ItemNote
                Set WorkingMenu = Nothing
                Set WorkingMenu = GetSpecialMenu(SubName)
                If WorkingMenu Is Nothing Then
                    Set WorkingMenu = GetCompleteMenu(SubName)
                Else
                    Prefix = ""
                End If
            Loop Until WorkingMenu Is Nothing
        
            If WorkingMenu Is Nothing Then Exit Do
            If WorkingMenu.Off Then Exit Do
            
            If List.Atomic Then
                Pick = 1
                Cost = WorkingMenu.ItemCost
            Else
                Select Case Int(Rnd * 6)
                    Case Is < 3: Pick = 1
                    Case Is < 5: Pick = 2
                    Case Is < 6: Pick = 3
                End Select
                If Pick > Total Then Pick = Total
                Cost = CStr(Pick)
            End If
            
            List.Insert Prefix & WorkingMenu.ItemName, Cost, _
                        IIf(WorkingMenu.Autonote, WorkingMenu.ItemNote, "")
        
            Total = Total - Pick
        
        Loop
    End If

    Set WorkingMenu = Nothing
    Set StartMenu = Nothing

End Sub

Public Sub GenerateRandomTraits(Char As Object)
'
' Name:         GenerateRandomTraits
' Parameter:    Char        Character object to make traits for
' Description:  Generate Random Traits for the given character,
'               according to the RandomTraits setting.
'

    Const rtPrimary = 0
    Const rtSecondary = 1
    Const rtTertiary = 2
    Const rtNegative = 3
    Const rtAbilities = 4
    Const rtBackgrounds = 5
    Const rtFree = 6

    Dim RT As String
    Dim Traits(6) As Long
    Dim Distribute(6) As Long
    Dim C As Integer
    Dim I As Integer
    Dim J As Integer
    Dim DList As LinkedList
    
    Set DList = New LinkedList
    
    On Error Resume Next

    RT = Game.RandomTraits & ","
    For I = 0 To 6
        C = InStr(RT, ",")
        If C > 0 Then Traits(I) = Int(Val(Mid(RT, 1, C - 1)))
        RT = Mid(RT, C + 1)
    Next I
    
    Traits(rtNegative) = Int(Rnd * (Traits(rtNegative) + 1))    'Adjust negative traits
    
    DList.InsertSorted Int(Rnd * (Traits(rtNegative) + 1))
    DList.InsertSorted Int(Rnd * (Traits(rtNegative) + 1))
    
    DList.First
    
    Distribute(rtPrimary) = CInt(DList.Item)
    I = Distribute(rtPrimary)
    DList.MoveNext
    Distribute(rtSecondary) = CInt(DList.Item) - I
    I = I + Distribute(rtSecondary)
    Distribute(rtTertiary) = Traits(rtNegative) - I
    
    AddToListRandomly Distribute(rtPrimary), Char.PhysicalNegList, "Physical, Negative", Char.Race
    AddToListRandomly Distribute(rtSecondary), Char.SocialNegList, "Social, Negative", Char.Race
    AddToListRandomly Distribute(rtTertiary), Char.MentalNegList, "Mental, Negative", Char.Race
    
    RT = CStr(Traits(rtNegative)) & " Negative and " & CStr(Traits(rtFree)) & " Free Traits"
    
    C = Traits(rtFree) + Traits(rtNegative)     ' Set C equal to total free traits
        
    DList.Clear
        
    DList.InsertSorted Int(Rnd * (C + 1))
    DList.InsertSorted Int(Rnd * (C + 1))
    DList.InsertSorted Int(Rnd * (C + 1))
    DList.InsertSorted Int(Rnd * (C + 1))
        
    DList.First
    
    Distribute(rtPrimary) = CInt(DList.Item)
    I = Distribute(rtPrimary)
    DList.MoveNext
    Distribute(rtSecondary) = CInt(DList.Item) - I
    I = I + Distribute(rtSecondary)
    DList.MoveNext
    Distribute(rtTertiary) = CInt(DList.Item) - I
    I = I + Distribute(rtTertiary)
    DList.MoveNext
    Distribute(rtAbilities) = CInt(DList.Item) - I
    I = I + Distribute(rtAbilities)
    Distribute(rtBackgrounds) = C - I
    
    For I = 0 To 6
        Distribute(I) = Distribute(I) + Traits(I)
    Next I
    
    For I = 0 To 2
        C = Int(Rnd * 3)
        J = Distribute(C)
        Distribute(C) = Distribute(I)
        Distribute(I) = J
    Next I

    AddToListRandomly Distribute(rtPrimary), Char.PhysicalList, "Physical", Char.Race
    AddToListRandomly Distribute(rtSecondary), Char.SocialList, "Social", Char.Race
    AddToListRandomly Distribute(rtTertiary), Char.MentalList, "Mental", Char.Race
    AddToListRandomly Distribute(rtAbilities), Char.AbilityList, "Abilities", Char.Race
    AddToListRandomly Distribute(rtBackgrounds), Char.BackgroundList, "Backgrounds", Char.Race

    RT = "Created " & CStr(Distribute(rtPrimary) & "/" & CStr(Distribute(rtSecondary))) & "/" & _
        CStr(Distribute(rtTertiary)) & " with " & CStr(Distribute(rtAbilities)) & " Abilities and " & _
        CStr(Distribute(rtBackgrounds)) & " Backgrounds, using " & RT

    I = Char.PhysicalMax
    If Char.PhysicalList.Count > I Then I = Char.PhysicalList.Count
    If Char.SocialList.Count > I Then I = Char.SocialList.Count
    If Char.MentalList.Count > I Then I = Char.MentalList.Count
    Char.PhysicalMax = I
    Char.SocialMax = I
    Char.MentalMax = I
    
    Char.Experience.Insert 0, ecComment, Now, RT
    
    Set DList = Nothing
    
End Sub

Public Sub SortListsByMenu(List As LinkedList, MenuName As String)
'
' Name:         SortListsByMenu
' Parameters:   List        a list of traitlists to sort
'               MenuName    name of the menu by which to order the traitlists
' Description:  Attempt to order the items in several lists of Powers according to the
'               order they have in a given menu.  Won't work with menus more than 10
'               deep, or menus with more than 10K items.
'

    Dim MList As LinkedMenuList
    Dim TList As LinkedTraitList
    Dim ID As String * 40
    Dim IDStep As Long
    Dim TName As String
    Dim TLen As Long
    Dim SegStart As Long
    Dim SegLen As Long
    Dim IPlace As Long
    Dim TmpName As String
    Dim StepBack As Long
    Dim StepAhead As Long
    
    Set MList = GetMenu(MenuName)
    
    If Not (MList Is Nothing Or List.IsEmpty) Then          'If we have a menu and some traitlists
    
        List.First
        Do Until List.Off                                   'for each Traitlist
        
            Set TList = List.Item
            If TList.Count > 1 Then                         'if the traitlist needs sorting
                        
                TList.First
                Do Until TList.Off                          'for each item in the list
                    
                    Set MList = GetMenu(MenuName)               'From the top menu
                    ID = String(40, "0")                        'ID string = 0
                    IDStep = 0                                  'current location in ID string
                    SegStart = 1                                'Start at the front of the trait name
                    TName = TList.Trait.Name                    'current trait name
                    TLen = Len(TName)                           'length of trait name
                    
                    Do Until SegStart > TLen Or IDStep = 10     'Until we exhaust the trait segments
                        SegLen = InStr(SegStart + 1, TName, ": ")
                        If SegLen = 0 Then SegLen = TLen + 1
                        SegLen = SegLen - SegStart                  'Find length of next segment
                                                                    'Move to that segment
                        IPlace = MList.GetItemPlace(Mid(TName, SegStart, SegLen))
                        
                        If IPlace = 0 Then                          'not found -- stop looking!
                            Exit Do
                        Else
                            Mid(ID, (IDStep * 4) + 1, 4) = Right("000" & CStr(IPlace), 4)
                            IDStep = IDStep + 1                         'Add to the ID, increment ID step
                            
                            Select Case MList.ItemCost
                                Case ":", "+"                           'it's a submenu link: deeper!
                                    Set MList = GetMenu(MList.ItemNote)
                                    If MList Is Nothing Then Exit Do
                                Case Else                               'it's not a submenu -- stop looking!
                                    TList.Trait.Total = MList.ItemCost
                                    If TList.Trait.Note = "" Then _
                                        TList.Trait.Note = MList.ItemNote
                                    Exit Do
                            End Select
                        End If
                        
                        SegStart = SegStart + SegLen + 2            'Move to next segment
                    Loop
                    
                    TmpName = ID & TName                        'create the ID
                    TList.Trait.Name = TmpName                  'add the ID
                    StepBack = 1
                    TList.MovePrevious
                    Do Until TList.Off                          'swap back until it's sorted in
                        If TmpName < TList.Trait.Name Then
                            TList.SwapBackward
                            TList.MovePrevious
                            TList.MovePrevious
                            StepBack = StepBack + 1
                        Else
                            Exit Do
                        End If
                    Loop
                    
                    If TList.Off Then
                        TList.First               'Restore cursor position
                        StepBack = StepBack - 1
                    End If
                    
                    For StepAhead = 1 To StepBack
                        TList.MoveNext
                    Next StepAhead
                    
                    TList.MoveNext                              'on to the next trait!
                Loop
        
                TList.First
                Do Until TList.Off                              'Clip away all the IDs
                    TList.Trait.Name = Mid(TList.Trait.Name, 41)
                    TList.MoveNext
                Loop
        
            End If
            
            List.MoveNext                                   'on to the next traitlist
        Loop

    End If

End Sub

Public Sub DeepSortListsByMenu(List As LinkedList, MenuName As String)
'
' Name:         DeepSortListsByMenu
' Parameters:   List        a list of traitlists to sort
'               MenuName    name of the menu by which to order the traitlists
' Description:  Attempt to order the items in several lists of Powers according to the
'               order they have in a given menu.  Slower and more brutal than above.
'               Used for powers whose menus changed a great deal between versions.
'

    Dim MList As LinkedMenuList
    Dim Stack As LinkedList
    Dim NameStack As LinkedList
    Dim FullName As String
    Dim TList As LinkedTraitList

    Set MList = GetMenu(MenuName)

    If Not (MList Is Nothing Or List.IsEmpty) Then

        Set Stack = New LinkedList
        Set NameStack = New LinkedList

        Stack.Append MList

        MList.First

        Do Until Stack.IsEmpty

            If MList.ItemCost = ":" Then
                NameStack.Append MList.ItemName
                Set MList = GetMenu(MList.ItemNote)
                If MList Is Nothing Then Exit Do
                Stack.Append MList
                MList.First
            ElseIf MList.ItemCost = "+" Then
                MList.MoveNext
            Else

                FullName = ""
                NameStack.First
                Do Until NameStack.Off
                    FullName = FullName & NameStack.Item & ": "
                    NameStack.MoveNext
                Loop
                FullName = FullName & MList.ItemName

                List.First
                Do Until List.Off

                    Set TList = List.Item

                    If Not (TList.IsEmpty Or TList.IsAlphabetized) Then
                        TList.MoveTo FullName
                        If TList.Off Then TList.MoveTo MList.ItemName

                        If Not TList.Off Then
                            TList.Trait.Name = FullName
                            If TList.Trait.Total = "1" Then TList.Trait.Total = MList.ItemCost
                            If MList.Autonote And TList.Trait.Note = "" Then _
                                    TList.Trait.Note = MList.ItemNote
                            Do Until TList.Off
                                TList.SwapBackward
                            Loop
                        End If
                    End If

                    List.MoveNext

                Loop

                MList.MoveNext
            End If

            Do Until Not MList.Off
                Set MList = Nothing
                NameStack.Last
                NameStack.Remove
                Stack.Last
                Stack.Remove
                Stack.Last
                If Not Stack.IsEmpty Then
                    Set MList = Stack.Item
                    MList.MoveNext
                Else
                    Exit Do
                End If
            Loop

        Loop

        Set Stack = Nothing
        Set NameStack = Nothing

    End If

End Sub

Public Sub AssignMeritsAndFlaws(Character As Object)
'
' Name:         AssignMeritsAndFlaws
' Parameters:   Character       character to give merits and flaws to
' Description:  When loading an old character, assign points to its
'               merits and flaws.
'

    Dim GeneralMenu As LinkedMenuList
    Dim SpecificMenu As LinkedMenuList
    Dim TraitList As LinkedTraitList
    
    Select Case Character.RaceCode
        Case gvRaceWraith:      Exit Sub
        Case gvRaceVarious:     Exit Sub
    End Select
    
    Set TraitList = Character.MeritList
    If Not TraitList.IsEmpty Then
        Set GeneralMenu = GetMenu("Merits")
        Set SpecificMenu = GetMenu("Merits, " & Character.Race)
        If Not GeneralMenu Is Nothing Then
            TraitList.First
            Do Until TraitList.Off
                If TraitList.Trait.Total = "1" Then
                    GeneralMenu.MoveTo TraitList.Trait.Name
                    If Not GeneralMenu.Off Then
                        TraitList.Trait.Total = GeneralMenu.ItemCost
                    ElseIf Not SpecificMenu Is Nothing Then
                        SpecificMenu.MoveTo TraitList.Trait.Name
                        If Not SpecificMenu.Off Then
                            TraitList.Trait.Total = SpecificMenu.ItemCost
                        End If
                    End If
                ElseIf TraitList.Trait.Total = "1 " Then
                    TraitList.Trait.Total = "1"
                End If
                TraitList.MoveNext
            Loop
        End If
    End If
    
    Set TraitList = Character.FlawList
    If Not TraitList.IsEmpty Then
        Set GeneralMenu = GetMenu("Flaws")
        Set SpecificMenu = GetMenu("Flaws, " & Character.Race)
        If Not GeneralMenu Is Nothing Then
            TraitList.First
            Do Until TraitList.Off
                If TraitList.Trait.Total = "1" Then
                    GeneralMenu.MoveTo TraitList.Trait.Name
                    If Not GeneralMenu.Off Then
                        TraitList.Trait.Total = GeneralMenu.ItemCost
                    ElseIf Not SpecificMenu Is Nothing Then
                        SpecificMenu.MoveTo TraitList.Trait.Name
                        If Not SpecificMenu.Off Then
                            TraitList.Trait.Total = SpecificMenu.ItemCost
                        End If
                    End If
                ElseIf TraitList.Trait.Total = "1 " Then
                    TraitList.Trait.Total = "1"
                End If
                TraitList.MoveNext
            Loop
        End If
    End If
    
End Sub

Public Function Off() As Boolean
'
' Name:         Off
' Description:  Tells whether the cursor has moved off the list.
' Returns:      TRUE if the cursor is off the list, FALSE otherwise.
'

    Off = (Cursor < 1 Or Cursor > MenuCollection.Count)

End Function
Public Sub MoveNext()
'
' Name:         MoveNext
' Description:  Move the cursor to the next item.
'

    Cursor = Cursor + 1

End Sub
Public Sub MovePrevious()
'
' Name:         MovePrevious
' Description:  Move the cursor to the previous item.
'

    Cursor = Cursor - 1

End Sub
Public Sub First()
'
' Name:         First
' Description:  Move the cursor to the first item.
'

    Cursor = 1

End Sub
Public Sub Last()
'
' Name:         Last
' Description:  Move the cursor to the last item.
'

    Cursor = MenuCollection.Count

End Sub
Public Function Menu() As LinkedMenuList
'
' Name:         Menu
' Description:  Returns the menu under the cursor.
' Returns:      the menu under the cursor, or Nothing if off.
'

    If Not Off Then Set Menu = MenuCollection.Item(Cursor)
    
End Function

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Create the collection object when the class is created.
'

    Set MenuCollection = New Collection

End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  If a list is destroyed, clear it first.  VB calls
'               this method automatically.
'

    Me.Clear
    Set MenuCollection = Nothing
    
End Sub
