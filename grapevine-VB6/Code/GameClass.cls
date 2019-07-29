VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GameClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PlayerList As LinkedList         'collection of all players
Public CharacterList As LinkedList      'collection of all characters
Public ItemList As LinkedList           'collection of all items
Public RoteList As LinkedList           'collection of all rotes
Public LocationList As LinkedList       'collection of all locations
Public AllRumorLists As LinkedList      'collection of lists of rumors, a list per date
Public InfluenceUseList As LinkedList   'collection of influence use actions
Public MenuSet As MenuSetClass          'collection of menus
Public QueryEngine As QueryEngineClass  'collection, logic of queries
Public APREngine As APREngineClass      'collection, logic of Actions/Plots/Rumors

Public FileFormat As FileFormatType     'Current file type to save as

Public GameFile As String               'full pathname of game file
Public DataChanged As Boolean           '"dirty" flag

Public ChronicleTitle As String         'Title of the Chronicle
Public Website As String                'Game URL
Public EMail As String                  'Main E-Mail address
Public Phone As String                  'Main Phone number
Public UsualPlace As String             'Usual Place to show
Public UsualTime As String              'usual time to show
Public Description As String            'description of your game
Public Calendar As CalendarClass        'game dates
Public ExtendedHealth As Boolean        'whether this game uses extended or abbreviated
                                        'health levels
Public EnforceHistory As Boolean        'whether to enforce the use of XP history
Public LinkTraitMaxes As Boolean        'whether to link Trait Maximums on character sheets
Public RandomTraits As String           'Comma-separated list of random trait options

Public STCommentStart As String         'opening markup of an ST comment
Public STCommentEnd As String           'ending markup of an ST comment

Public XPAwardList As LinkedList        'list of standard XP and PP awards
Public TemplateList As LinkedList       'list of output templates

Public FileError As Boolean             'whether a file error happened during open or save
Public FileErrorMessage As String       'Description of the error
Public MergeResults As String           'vbCR-delimited results of a merge or exchange file load

Public FileProgress As ProgressBar      'Control describing progress of load

Private DuplicateAction As DupAction    'What action to take when duplicating characters
Private DuplicateAll As Boolean         'Whether to take that action in all cases

Private Const CALENDAR_PERCENT = 5      'Percent of progress bar to fill for each loading part
Private Const PLAYER_PERCENT = 30
Private Const CHAR_PERCENT = 65

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Create all needed objects.
'

    Set PlayerList = New LinkedList
    Set CharacterList = New LinkedList
    Set ItemList = New LinkedList
    Set RoteList = New LinkedList
    Set LocationList = New LinkedList
    Set AllRumorLists = New LinkedList
    Set InfluenceUseList = New LinkedList
    Set MenuSet = New MenuSetClass
    Set Calendar = New CalendarClass
    Set QueryEngine = New QueryEngineClass
    Set APREngine = New APREngineClass
    Set XPAwardList = New LinkedList
    Set TemplateList = New LinkedList
    
    ExtendedHealth = True
    EnforceHistory = True
    LinkTraitMaxes = True
    
    RandomTraits = "7,5,3,5,5,5,5"
    STCommentStart = "[ST]"
    STCommentEnd = "[/ST]"
    
End Sub

Private Sub AddCharacterFromFile(Race As String, FileNum As Integer, _
        OldDate As Date, Version As String, Optional ByRef CharAdded As Object = Nothing)
'
' Name:         AddCharacterFromFile
' Parameters:   Race        a string representation of the character's race
'               FileNum     an open file number
'               OldDate     the date of the old file
'               Version     the file format version tag
'               CharAdded   reference to the actual character added to the game, if any
' Description:  Create a new character object of the appropriate type and
'               have it load itself from the file. Used for v2.3 files only --
'               this method survives for backward-compatibility.
'
' REQUIRED GUI CLASS: frmDuplicate
'
        
    Dim NewCharacter As Object
    Dim NewName As String
    
    Select Case Trim(Race)
        Case CStr(gvRaceVampire), "Vampire"
            Set NewCharacter = New VampireClass
        Case CStr(gvRaceWerewolf), "Werewolf"
            Set NewCharacter = New WerewolfClass
        Case CStr(gvRaceMortal), "Mortal"
            Set NewCharacter = New MortalClass
        Case CStr(gvRaceChangeling), "Changeling"
            Set NewCharacter = New ChangelingClass
        Case CStr(gvRaceWraith)
            Set NewCharacter = New WraithClass
        Case CStr(gvracemage)
            Set NewCharacter = New MageClass
        Case CStr(gvRaceFera)
            Set NewCharacter = New FeraClass
        Case CStr(gvRaceVarious)
            Set NewCharacter = New VariousClass
        Case CStr(gvRaceMummy)
            Set NewCharacter = New MummyClass
        Case CStr(gvRaceKueiJin)
            Set NewCharacter = New KueiJinClass
        Case CStr(gvRaceHunter)
            Set NewCharacter = New HunterClass
        Case CStr(gvRaceDemon)
            Set NewCharacter = New DemonClass
        Case Else
            Err.Raise vbObjectError + 9, "Game.AddCharacterFromFile", _
                "Unexpected character type """ & Race & """ read."
    End Select

    On Error GoTo AddChar_Finish
    NewCharacter.OldInputFromFile FileNum, OldDate, Version
    
    On Error GoTo 0
    
    CharacterList.MoveTo NewCharacter.Name
    If CharacterList.Off Then
        CharacterList.InsertSorted NewCharacter
        MergeResults = MergeResults & NewCharacter.Name & vbCr & "(added)" & vbCr
    Else
        InsertDuplicate CharacterList, NewCharacter
    End If

    Set CharAdded = NewCharacter

AddChar_Finish:
    Set NewCharacter = Nothing

End Sub

Private Sub AddDefaultTemplates()
'
' Name:         AddDefaultTemplates
' Description:  Add the default output templates to this game.
'

    Dim Template As TemplateClass
    Dim NameSet As Collection
    Dim Name As Variant
    Dim FileBase As String
    
    Set NameSet = New Collection
    TemplateList.Clear
    
    NameSet.Add "Changeling" & tnCharSheetSuffix
    NameSet.Add "Demon" & tnCharSheetSuffix
    NameSet.Add "Fera" & tnCharSheetSuffix
    NameSet.Add "Kuei-Jin" & tnCharSheetSuffix
    NameSet.Add "Hunter" & tnCharSheetSuffix
    NameSet.Add "Mage" & tnCharSheetSuffix
    NameSet.Add "Mortal" & tnCharSheetSuffix
    NameSet.Add "Mummy" & tnCharSheetSuffix
    NameSet.Add "Vampire" & tnCharSheetSuffix
    NameSet.Add "Various" & tnCharSheetSuffix
    NameSet.Add "Werewolf" & tnCharSheetSuffix
    NameSet.Add "Wraith" & tnCharSheetSuffix
    
    NameSet.Add tnActionRumor
    NameSet.Add tnMasterAction
    NameSet.Add tnMasterRumor
    NameSet.Add tnPlot
    NameSet.Add tnCharacterRoster
    NameSet.Add tnEquipment
    NameSet.Add tnSignIn
    NameSet.Add tnItemCards
    NameSet.Add tnRoteCards
    NameSet.Add tnLocationCards
    NameSet.Add tnXPHistory
    NameSet.Add tnPPHistory
    NameSet.Add tnPlayerRoster
    NameSet.Add tnGameCalendar
    NameSet.Add tnSearch
    NameSet.Add tnStatistics
    NameSet.Add tnVampireStatus
    NameSet.Add tnMFReport
    NameSet.Add tnInfluenceReport
    
    For Each Name In NameSet
        Set Template = New TemplateClass
        Template.Name = Name
        FileBase = Name
        Template.IsCharacterSheet = (InStr(FileBase, tnCharSheetSuffix) > 0)
        If Template.IsCharacterSheet Then
            FileBase = Left(FileBase, InStr(FileBase, " ") - 1)
        End If
        Template.SetFilename ofText, "Templates\Text\" & FileBase & ".txt"
        Template.SetFilename ofRTF, "Templates\RTF\" & FileBase & ".rtf"
        Template.SetFilename ofHTML, "Templates\HTML\" & FileBase & ".html"
        TemplateList.InsertSorted Template
    Next Name
    
    Set NameSet = Nothing

End Sub

Private Sub AddDefaultXPAwards()
'
' Name:         AddDefaultXPAwards
' Description:  Add the default XP and PP awards to this game.
'

    Dim Award As ExperienceAwardClass
    Dim I As Integer
    
    XPAwardList.Clear
    For I = 1 To 6
        Set Award = New ExperienceAwardClass
        Select Case I
            Case 1: Award.Name = "Attendance"
            Case 2: Award.Name = "Costuming"
            Case 3: Award.Name = "Downtime Activities"
            Case 4: Award.Name = "First Night"
            Case 5: Award.Name = "Good Roleplaying"
            Case 6: Award.Name = "Leadership"
        End Select
        Award.XP = True
        Award.ChangeType = ecEarned
        Award.Change = 1
        Award.Reason = Award.Name
        XPAwardList.Append Award
    Next I
    
    For I = 1 To 4
        Set Award = New ExperienceAwardClass
        Select Case I
            Case 1: Award.Name = "Storytelling"
            Case 2: Award.Name = "Bookkeeping"
            Case 3: Award.Name = "Narrating"
            Case 4: Award.Name = "Setup/Cleanup"
        End Select
        Award.XP = False
        Award.ChangeType = ecEarned
        Award.Change = 1
        Award.Reason = Award.Name
        XPAwardList.Append Award
    Next I
    
End Sub

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

Private Sub DeleteDuplicates(List As LinkedList, Warning As String)
'
' Name:         DeleteDuplicates
' Parameters:   List        list from which to delete
'               Warning     Type of object to warn about if duplicates found
' Description:  Traverse a list looking for duplicates and deleting those found.
'

    Dim Found As StringSet
    Dim Warn As Boolean
    
    Set Found = New StringSet
    Warn = False

    List.First
    Do Until List.Off
        If Found.Has(List.Item.Name) Then
            List.Remove
            Warn = True
        Else
            Found.Add List.Item.Name
            List.MoveNext
        End If
    Loop

    If Warn Then MsgBox "Warning -- duplicate " & Warning & " were found in this file." & _
                        vbCrLf & "They have been deleted.", vbExclamation + vbOKOnly, "Duplicates"

    Set Found = Nothing

End Sub

Private Sub EnsureNoDuplicates()
'
' Name:         EnsureNoDuplicates
' Description:  Check all game lists for duplicates, deleting thos found.
'

    DeleteDuplicates CharacterList, "characters"
    DeleteDuplicates PlayerList, "players"
    DeleteDuplicates QueryEngine.QueryList, "searches"
    DeleteDuplicates ItemList, "items"
    DeleteDuplicates RoteList, "rotes"
    DeleteDuplicates LocationList, "locations"
    DeleteDuplicates APREngine.ActionList, "actions"
    DeleteDuplicates APREngine.PlotList, "plots"
    DeleteDuplicates APREngine.RumorList, "rumors"
    DeleteDuplicates XPAwardList, "XP awards"
    DeleteDuplicates TemplateList, "template settings"
    
End Sub

Private Function EntityCount() As Integer
'
' Name:         EntityCount
' Description:  Return the number of entities in this game.
'
    EntityCount = PlayerList.Count + CharacterList.Count + QueryEngine.QueryList.Count + _
                  ItemList.Count + RoteList.Count + LocationList.Count + _
                  APREngine.ActionList.Count + APREngine.PlotList.Count + APREngine.RumorList.Count
           
End Function

Public Sub GetValue(Key As String, GData As Variant)
'
' Name:         GetValue
' Parameters:   Key         a key (string) to the value to retreive
'               GData       the returned value - a string, number, date, etc.
' Description:  Use the given key to return data from the character.  Keys are defined
'               in the PublicQueryKeys module.
'

    GData = ""
    Select Case Key
        
        Case qkTitle:       GData = ChronicleTitle
        Case qkWebSite:     GData = Website
        Case qkEMail:       GData = EMail
        Case qkPhone:       GData = Phone
        Case qkDescription: GData = Description
        Case qkUsualPlace:  GData = UsualPlace
        Case qkUsualTime:   GData = UsualTime
        Case qkNextGame:    If Not Calendar.Off Then GData = Calendar.GetGameDate
        Case qkNextPlace:   If Not Calendar.Off Then GData = Calendar.GetGamePlace
        Case qkNextTime:    If Not Calendar.Off Then GData = Calendar.GetGameTime
        Case qkNextNotes:   If Not Calendar.Off Then GData = Calendar.GetGameNotes
        Case Else:          GData = Null
        
    End Select
    
End Sub

Public Sub InitializeForOutput()
'
' Name:         InitializeForOutput
' Description:  Prepare the game for output by moving to the next game date.
'
    Calendar.MoveToCloseGame

End Sub

Private Sub InsertDuplicate(List As LinkedList, NewEntry As Object, Optional Rename As Boolean = True)
'
' Name:         InsertDuplicate
' Parameters:   List        The list to which to append the duplicate
'               NewEntry    The player or character to append
'               Rename      Whether renaming the object is a viable option
' Description:  Determine what to do with a duplicate entry added to
'               a list, then carry out that action.  If DuplicateAll
'               is set, perform the selected action automatically.
'
' REQUIRED GUI CLASS: frmDuplicate
'

    Dim NewName As String
    Dim ExMod As Date
    Dim InMod As Date
    Dim MouseState As Integer
    
    ExMod = Now
    InMod = ExMod
    On Error Resume Next
    ExMod = List.Item.LastModified
    InMod = NewEntry.LastModified
    On Error GoTo 0
    
    If Not DuplicateAll Or (DuplicateAll And Not Rename And (DuplicateAction = RenameNew Or DuplicateAction = RenameOld)) Then
        Load frmDuplicate
        MouseState = Screen.MousePointer
        Screen.MousePointer = vbDefault
        frmDuplicate.FixDuplicate NewEntry.Name, List, ExMod, InMod, Rename
        Screen.MousePointer = MouseState
        DuplicateAll = frmDuplicate.All
        DuplicateAction = frmDuplicate.Action
        NewName = frmDuplicate.NewName
        Unload frmDuplicate
        Set frmDuplicate = Nothing
    End If
    
    Select Case DuplicateAction
        Case RenameNew
            NewEntry.Name = NewName
            MergeResults = MergeResults & NewEntry.Name & vbCr & "(added)" & vbCr
        Case RenameOld
            List.MoveTo NewEntry.Name
            List.Item.Name = NewName
            MergeResults = MergeResults & NewName & vbCr & "(renamed)" & vbCr
            MergeResults = MergeResults & NewEntry.Name & vbCr & "(added)" & vbCr
        Case ReplaceOld
            List.MoveTo NewEntry.Name
            List.Remove
            MergeResults = MergeResults & NewEntry.Name & vbCr & "(replaced)" & vbCr
        Case Skip
            Set NewEntry = Nothing
            Exit Sub
        Case Else
            If (DuplicateAction = KeepNewer And InMod > ExMod) Or _
               (DuplicateAction = keepolder And InMod < ExMod) Then
                List.MoveTo NewEntry.Name
                List.Remove
                MergeResults = MergeResults & NewEntry.Name & vbCr & "(replaced)" & vbCr
            Else
                Set NewEntry = Nothing
                Exit Sub
            End If
    End Select
    
    If List Is ActionList Or List Is RumorList Then
        APREngine.InsertSorted List, NewEntry
    Else
        List.InsertSorted NewEntry
    End If

End Sub

Public Sub LoadExchange(FileName As String)
'
' Name:         LoadExchange
' Parameters:   FileName        the file to load
' Description:  Merge an exchange file with the current game's
'               data.
'
' REQUIRED GUI CLASS: frmDuplicate
'

    FileError = True
    FileErrorMessage = ""
    
    On Error GoTo LoadExchange_DetectError
    
    Select Case DetectFileFormat(FileName)
        Case gvBinaryExchange
            LoadExchangeBinary FileName
        Case gvXML
            LoadExchangeXML FileName
        Case gv23Exchange
            LoadExchangeV23 FileName
        Case gvBinaryMenu
            FileErrorMessage = ShortFile(FileName) & _
                    " is a grapevine menu file: open it from Game -> Grapevine Menus."
        Case gvBinaryGame
            FileErrorMessage = ShortFile(FileName) & _
                    " is a grapevine Game file: open it from File -> Open Game."
        Case gv23Game
            FileErrorMessage = ShortFile(FileName) & _
                    " is a grapevine Game file: open it from File -> Open Game."
        Case gvInvalid
            FileErrorMessage = ShortFile(FileName) & _
                    " is not a grapevine game file."
    End Select
    
    On Error GoTo 0
    
    EnsureNoDuplicates
    APREngine.ReconnectAll True
    
    GoTo LoadExchange_EndSub
    
LoadExchange_DetectError:
    FileErrorMessage = "Error loading " & ShortFile(FileName) & ":" & vbCrLf & Err.Description
    Resume LoadExchange_EndSub
    
LoadExchange_EndSub:

End Sub

Private Sub LoadExchangeBinary(FileName As String)
'
' Name:         LoadExchangeBinary
' Parameters:   FileName        the name of the file to load
' Description:  Load a selection of data from a binary file.
'
    
    Dim FileNum As Integer
    Dim Verify As String
    Dim NewEntity As Object
    Dim Version As Double
    Dim I As Integer
    Dim RCode As Integer
    
    FileNum = FreeFile
    FileError = True
    DuplicateAll = False
    MergeResults = ""
    
    On Error GoTo LEB_AnyError
    
    FileErrorMessage = "Failed to open binary exchange file " & ShortFile(FileName) & ":"
    Open FileName For Binary As #FileNum
    
    FileErrorMessage = "Error reading binary exchange header:"
    GetStrB FileNum, Verify
    If Not Verify = BinHeaderExchange Then
    
        FileErrorMessage = ShortFile(FileName) & " is not a Grapevine exchange file."
        Exit Sub
        
    End If
    
    Get #FileNum, , Version
    
    If Version >= 2.395 Then
    
        FileErrorMessage = "Error reading calendar:"
        Get #FileNum, , I
        If I > 0 Then
            Set NewEntity = New CalendarClass
            NewEntity.InputFromBinary FileNum, Version
            If MsgBox("Do you want to replace the current game calendar, modified " & _
                      Format(Calendar.LastModified, "Short Date") & ", with the calendar " & _
                      "from this file, modified " & Format(NewEntity.LastModified, "Short Date") & _
                      "?", vbYesNo + vbQuestion, "Load Calendar") = vbYes Then
                Set Calendar = Nothing
                Set Calendar = NewEntity
                MergeResults = MergeResults & "Game Calendar" & vbCr & _
                               "(replaced)" & vbCr
            End If
            Set NewEntity = Nothing
        End If
    
        If Version >= 2.397 Then
        
            FileErrorMessage = "Error reading action/plot/rumor settings:"
            Get #FileNum, , I
            If I > 0 Then
                Set NewEntity = New APREngineClass
                NewEntity.InputFromBinary FileNum, Version
                If MsgBox("Do you want to replace the current action/plot/rumor settings " & _
                          " with the settings from this file?", vbYesNo + vbQuestion, _
                          "Load Settings") = vbYes Then
                    
                    Set NewEntity.ActionList = Nothing
                    Set NewEntity.PlotList = Nothing
                    Set NewEntity.RumorList = Nothing
                    Set NewEntity.ActionList = APREngine.ActionList
                    Set NewEntity.PlotList = APREngine.PlotList
                    Set NewEntity.RumorList = APREngine.RumorList
                    Set APREngine = Nothing
                    Set APREngine = NewEntity
                    Set ActionList = APREngine.ActionList
                    Set PlotList = APREngine.PlotList
                    Set RumorList = APREngine.RumorList
                    
                    MergeResults = MergeResults & "Action/Plot/Rumor Settings" & vbCr & _
                                   "(replaced)" & vbCr
                End If
                Set NewEntity = Nothing
            End If
        
            FileErrorMessage = "Error reading XP/PP awards:"
            Get #FileNum, , I
            Do Until I = 0
                Set NewEntity = New ExperienceAwardClass
                NewEntity.InputFromBinary FileNum, Version
                XPAwardList.MoveTo NewEntity.Name
                If XPAwardList.Off Then
                    MergeResults = MergeResults & NewEntity.Name & vbCr & "(added)" & vbCr
                Else
                    XPAwardList.Remove
                    MergeResults = MergeResults & NewEntity.Name & vbCr & "(replaced)" & vbCr
                End If
                XPAwardList.InsertSorted NewEntity
                I = I - 1
            Loop
            
            FileErrorMessage = "Error reading templates:"
            Get #FileNum, , I
            Do Until I = 0
                Set NewEntity = New TemplateClass
                NewEntity.InputFromBinary FileNum, Version
                TemplateList.MoveTo NewEntity.Name
                If TemplateList.Off Then
                    MergeResults = MergeResults & NewEntity.Name & vbCr & "(added)" & vbCr
                Else
                    TemplateList.Remove
                    MergeResults = MergeResults & NewEntity.Name & vbCr & "(replaced)" & vbCr
                End If
                TemplateList.InsertSorted NewEntity
                I = I - 1
            Loop
                    
        End If
    
    End If
    
    FileErrorMessage = "Error reading players:"
    Get #FileNum, , I
    Do Until I = 0
        Set NewEntity = New PlayerClass
        NewEntity.InputFromBinary FileNum, Version
        PlayerList.MoveTo NewEntity.Name
        If PlayerList.Off Then
            PlayerList.InsertSorted NewEntity
            MergeResults = MergeResults & NewEntity.Name & vbCr & "(added)" & vbCr
        Else
            InsertDuplicate PlayerList, NewEntity
        End If
        I = I - 1
    Loop
    
    FileErrorMessage = "Error reading characters:"
    Get #FileNum, , I
    Do Until I = 0
        Get #FileNum, , RCode
        Select Case RCode
            Case gvRaceVampire
                Set NewEntity = New VampireClass
            Case gvRaceWerewolf
                Set NewEntity = New WerewolfClass
            Case gvracemage
                Set NewEntity = New MageClass
            Case gvRaceChangeling
                Set NewEntity = New ChangelingClass
            Case gvRaceWraith
                Set NewEntity = New WraithClass
            Case gvRaceMortal
                Set NewEntity = New MortalClass
            Case gvRaceMummy
                Set NewEntity = New MummyClass
            Case gvRaceKueiJin
                Set NewEntity = New KueiJinClass
            Case gvRaceFera
                Set NewEntity = New FeraClass
            Case gvRaceVarious
                Set NewEntity = New VariousClass
            Case gvRaceHunter
                Set NewEntity = New HunterClass
            Case gvRaceDemon
                Set NewEntity = New DemonClass
        End Select
        NewEntity.InputFromBinary FileNum, Version
        CharacterList.MoveTo NewEntity.Name
        If CharacterList.Off Then
            CharacterList.InsertSorted NewEntity
            MergeResults = MergeResults & NewEntity.Name & vbCr & "(added)" & vbCr
        Else
            InsertDuplicate CharacterList, NewEntity
        End If
        I = I - 1
    Loop
    
    FileErrorMessage = "Error reading queries:"
    Get #FileNum, , I
    Do Until I = 0
        Set NewEntity = New QueryClass
        NewEntity.InputFromBinary FileNum, Version
        QueryEngine.QueryList.MoveTo NewEntity.Name
        If QueryEngine.QueryList.Off Then
            QueryEngine.QueryList.InsertSorted NewEntity
            MergeResults = MergeResults & NewEntity.Name & vbCr & "(added)" & vbCr
        Else
            InsertDuplicate QueryEngine.QueryList, NewEntity
        End If
        I = I - 1
    Loop
    
    FileErrorMessage = "Error reading items:"
    Get #FileNum, , I
    Do Until I = 0
        Set NewEntity = New ItemClass
        NewEntity.InputFromBinary FileNum, Version
        ItemList.MoveTo NewEntity.Name
        If ItemList.Off Then
            ItemList.InsertSorted NewEntity
            MergeResults = MergeResults & NewEntity.Name & vbCr & "(added)" & vbCr
        Else
            InsertDuplicate ItemList, NewEntity
        End If
        I = I - 1
    Loop
    
    FileErrorMessage = "Error reading rotes:"
    Get #FileNum, , I
    Do Until I = 0
        Set NewEntity = New RoteClass
        NewEntity.InputFromBinary FileNum, Version
        RoteList.MoveTo NewEntity.Name
        If RoteList.Off Then
            RoteList.InsertSorted NewEntity
            MergeResults = MergeResults & NewEntity.Name & vbCr & "(added)" & vbCr
        Else
            InsertDuplicate RoteList, NewEntity
        End If
        I = I - 1
    Loop
    
    FileErrorMessage = "Error reading locations:"
    Get #FileNum, , I
    Do Until I = 0
        Set NewEntity = New LocationClass
        NewEntity.InputFromBinary FileNum, Version
        LocationList.MoveTo NewEntity.Name
        If LocationList.Off Then
            LocationList.InsertSorted NewEntity
            MergeResults = MergeResults & NewEntity.Name & vbCr & "(added)" & vbCr
        Else
            InsertDuplicate LocationList, NewEntity
        End If
        I = I - 1
    Loop
    
    FileErrorMessage = "Error reading actions:"
    Get #FileNum, , I
    Do Until I = 0
        Set NewEntity = New ActionClass
        NewEntity.InputFromBinary FileNum, Version
        ActionList.MoveTo NewEntity.Name
        If ActionList.Off Then
            APREngine.InsertSorted ActionList, NewEntity
            MergeResults = MergeResults & NewEntity.Name & vbCr & "(added)" & vbCr
        Else
            InsertDuplicate ActionList, NewEntity, False
        End If
        I = I - 1
    Loop
    
    FileErrorMessage = "Error reading plots:"
    Get #FileNum, , I
    Do Until I = 0
        Set NewEntity = New PlotClass
        NewEntity.InputFromBinary FileNum, Version
        PlotList.MoveTo NewEntity.Name
        If PlotList.Off Then
            PlotList.InsertSorted NewEntity
            MergeResults = MergeResults & NewEntity.Name & vbCr & "(added)" & vbCr
        Else
            InsertDuplicate PlotList, NewEntity, False
        End If
        I = I - 1
    Loop
    
    FileErrorMessage = "Error reading rumors:"
    Get #FileNum, , I
    Do Until I = 0
        Set NewEntity = New RumorClass
        NewEntity.InputFromBinary FileNum, Version
        RumorList.MoveTo NewEntity.Name
        If RumorList.Off Then
            Game.APREngine.InsertSorted RumorList, NewEntity
            MergeResults = MergeResults & NewEntity.Name & vbCr & "(added)" & vbCr
        Else
            InsertDuplicate RumorList, NewEntity, False
        End If
        I = I - 1
    Loop
    
    FileError = False
    DataChanged = True
    
    GoTo LEB_CloseFile

LEB_AnyError:
    FileErrorMessage = FileErrorMessage & vbCrLf & Err.Description
    Resume LEB_CloseFile

LEB_CloseFile:
    Close #FileNum
    
End Sub

Private Sub LoadExchangeV23(FileName As String)
'
' Name:         LoadExchangeV23
' Parameters:   FileName        the file to load
' Description:  Merge a v2.3 exchange file with the current game's
'               data.
'
' REQUIRED GUI CLASS: frmDuplicate
'

    Dim FileNum As Integer
    Dim Read As String
    Dim Version As String
    Dim NewPlayer As PlayerClass
    Dim OldDate As Date
    Dim CharAdded As Object
    Dim AddedList As LinkedList
    
    Const CharacterTag = "<-Character Tag->"
    Const PlayerTag = "<-Player Tag->"
    Const EndTag = "<-End Tag->"

    FileNum = FreeFile
    FileError = True
    MergeResults = ""
    Set AddedList = New LinkedList
    
    On Error GoTo LEV23_AnyError
    FileErrorMessage = "Failed to open file:"
    OldDate = FileDateTime(FileName)
    Open FileName For Input As #FileNum
    
    Line Input #FileNum, Version
    Line Input #FileNum, Version
        
    Line Input #FileNum, Read
    DuplicateAll = False
    
    Do Until Read = EndTag
    
        Select Case Read

            Case PlayerTag
                
                FileErrorMessage = "Error reading player from disk:"
                Set NewPlayer = New PlayerClass
                NewPlayer.OldInputFromFile FileNum, OldDate, Version
                PlayerList.MoveTo NewPlayer.Name
                If PlayerList.Off Then
                    PlayerList.InsertSorted NewPlayer
                    MergeResults = MergeResults & NewPlayer.Name & vbCr & "(added)" & vbCr
                Else
                    InsertDuplicate PlayerList, NewPlayer
                End If
                
            Case CharacterTag
        
                FileErrorMessage = "Error reading character from disk:"
                Line Input #FileNum, Read
                AddCharacterFromFile Read, FileNum, OldDate, Version, CharAdded
                If Not CharAdded Is Nothing Then AddedList.Append CharAdded
                    
        End Select
    
        Line Input #FileNum, Read
            
    Loop
    
    FileProgress.Value = 0
    FileProgress.Visible = True
    OrderLists AddedList
    FileProgress.Visible = False
    
    FileError = False
    DataChanged = True

    GoTo LEV23_CloseFile
    
LEV23_AnyError:
    FileErrorMessage = FileErrorMessage & vbCrLf & Err.Description
    Resume LEV23_CloseFile

LEV23_CloseFile:
    Close #FileNum
    Set NewPlayer = Nothing
    Set CharAdded = Nothing
    Set AddedList = Nothing
    
End Sub

Private Sub LoadExchangeXML(FileName As String)
'
' Name:         LoadExchangeXML
' Parameters:   FileName        the file to load
' Description:  Merge an XML exchange file with the current game's
'               data.
'
' REQUIRED GUI CLASS: frmDuplicate
'

    Dim FileNum As Integer
    Dim XML As XMLReaderClass
    Dim NewEntity As Object
    Dim NewEntityList As LinkedList
    Dim Version As Double
    Dim Rename As Boolean
    
    FileNum = FreeFile
    FileError = True
    Set XML = New XMLReaderClass
    DuplicateAll = False
    MergeResults = ""
    
    On Error GoTo LEX_AnyError
    FileErrorMessage = "Failed to open " & ShortFile(FileName) & ":"
    Open FileName For Input As #FileNum
    
    XML.BeginReading FileNum
    If Left(XML.VersionTag, 5) = "<?xml" Then   'legal file
    
        With XML
        
            .ReadTag
            If .Tag = "grapevine" Then
            
                Version = .GetDouble("version")
            
                Do Until .Tag = "grapevine" And .Terminal
                
                    Rename = True
                    Set NewEntity = Nothing
                    Set NewEntityList = CharacterList
                    .ReadTag
                    
                    Select Case .Tag
                        Case "calendar"
                            
                            Set NewEntity = New CalendarClass
                            NewEntity.InputFromFile XML, Version
                            
                            If MsgBox("Do you want to replace the current game calendar, modified " & _
                                      Format(Calendar.LastModified, "Short Date") & ", with the calendar " & _
                                      "from this file, modified " & Format(NewEntity.LastModified, "Short Date") & _
                                      "?", vbYesNo + vbQuestion, "Load Calendar") = vbYes Then
                                Set Calendar = Nothing
                                Set Calendar = NewEntity
                                MergeResults = MergeResults & "Game Calendar" & vbCr & _
                                               "(replaced)" & vbCr
                            End If
                            Set NewEntity = Nothing
                            
                        Case "aprsettings"
                            
                            Set NewEntity = New APREngineClass
                            NewEntity.InputFromFile XML, Version
                            
                            If MsgBox("Do you want to replace the current action/plot/rumor settings " & _
                                      " with the settings from this file?", vbYesNo + vbQuestion, _
                                      "Load Settings") = vbYes Then
                                
                                Set NewEntity.ActionList = Nothing
                                Set NewEntity.PlotList = Nothing
                                Set NewEntity.RumorList = Nothing
                                Set NewEntity.ActionList = APREngine.ActionList
                                Set NewEntity.PlotList = APREngine.PlotList
                                Set NewEntity.RumorList = APREngine.RumorList
                                Set APREngine = Nothing
                                Set APREngine = NewEntity
                                Set ActionList = APREngine.ActionList
                                Set PlotList = APREngine.PlotList
                                Set RumorList = APREngine.RumorList
                                
                                MergeResults = MergeResults & "Action/Plot/Rumor Settings" & vbCr & _
                                               "(replaced)" & vbCr
                            End If
                            Set NewEntity = Nothing
                            
                        Case "player"
                            Set NewEntity = New PlayerClass
                            Set NewEntityList = PlayerList
                        Case "vampire"
                            Set NewEntity = New VampireClass
                        Case "werewolf"
                            Set NewEntity = New WerewolfClass
                        Case "mage"
                            Set NewEntity = New MageClass
                        Case "changeling"
                            Set NewEntity = New ChangelingClass
                        Case "wraith"
                            Set NewEntity = New WraithClass
                        Case "mortal"
                            Set NewEntity = New MortalClass
                        Case "mummy"
                            Set NewEntity = New MummyClass
                        Case "kueijin"
                            Set NewEntity = New KueiJinClass
                        Case "fera"
                            Set NewEntity = New FeraClass
                        Case "various"
                            Set NewEntity = New VariousClass
                        Case "hunter"
                            Set NewEntity = New HunterClass
                        Case "demon"
                            Set NewEntity = New DemonClass
                        Case "query"
                            Set NewEntity = New QueryClass
                            Set NewEntityList = QueryEngine.QueryList
                        Case "item"
                            Set NewEntity = New ItemClass
                            Set NewEntityList = ItemList
                        Case "rote"
                            Set NewEntity = New RoteClass
                            Set NewEntityList = RoteList
                        Case "location"
                            Set NewEntity = New LocationClass
                            Set NewEntityList = LocationList
                        Case "action"
                            Set NewEntity = New ActionClass
                            Set NewEntityList = APREngine.ActionList
                            Rename = False
                        Case "plot"
                            Set NewEntity = New PlotClass
                            Set NewEntityList = APREngine.PlotList
                            Rename = False
                        Case "rumor"
                            Set NewEntity = New RumorClass
                            Set NewEntityList = APREngine.RumorList
                            Rename = False
                        Case "award"
                            Set NewEntity = New ExperienceAwardClass
                            Set NewEntityList = XPAwardList
                        Case "template"
                            Set NewEntity = New TemplateClass
                            Set NewEntityList = TemplateList
                    End Select
    
                    If Not NewEntity Is Nothing Then
                        NewEntity.InputFromFile XML, Version
                        NewEntityList.MoveTo NewEntity.Name
                        If NewEntityList.Off Then
                            If NewEntityList Is ActionList Or NewEntityList Is RumorList Then
                                APREngine.InsertSorted NewEntityList, NewEntity
                            Else
                                NewEntityList.InsertSorted NewEntity
                            End If
                            MergeResults = MergeResults & NewEntity.Name & vbCr & "(added)" & vbCr
                        Else
                            If NewEntityList Is XPAwardList Or NewEntityList Is TemplateList Then
                                NewEntityList.Remove
                                NewEntityList.InsertSorted NewEntity
                                MergeResults = MergeResults & NewEntity.Name & vbCr & "(replaced)" & vbCr
                            Else
                                InsertDuplicate NewEntityList, NewEntity, Rename
                            End If
                        End If
                    End If
                    
                Loop
            
                FileError = False
                DataChanged = True
    
            Else 'not a grapevine game file!
            
                FileErrorMessage = ShortFile(FileName) & " is not a Grapevine file."
            
            End If
            
        End With
        
    Else
            
        FileErrorMessage = ShortFile(FileName) & " is not a Grapevine file."
        
    End If

    GoTo LEX_CloseFile
    
LEX_AnyError:
    FileErrorMessage = FileErrorMessage & vbCrLf & Err.Description
    Resume LEX_CloseFile

LEX_CloseFile:
    Set XML = Nothing
    Close #FileNum
       
End Sub

Public Sub Merge(MergeFile As String)
'
' Name:         Merge
' Parameters:   MergeFile       the file to merge with
' Description:  Merge two games, keeping the newest data.
'

    Dim MergeGame As GameClass
    
    Set MergeGame = New GameClass
    Set MergeGame.FileProgress = FileProgress
    MergeResults = ""
    
    MergeGame.OpenGame MergeFile, False  'load the game, but not the menu
    
    If MergeGame.FileError Then
        FileError = True
        FileErrorMessage = MergeGame.FileErrorMessage
    Else
        
        FileProgress.Value = 0
        FileProgress.Max = 100
        
        If Calendar.LastModified < MergeGame.Calendar.LastModified Then
            Set Calendar = Nothing
            Set Calendar = MergeGame.Calendar
            Set MergeGame.Calendar = Nothing
            Set MergeGame.Calendar = New CalendarClass
            MergeResults = "Game Calendar" & vbCr & "(replaced)" & vbCr
        End If
        
        MergeLists PlayerList, MergeGame.PlayerList
        FileProgress.Value = 10
        
        MergeLists CharacterList, MergeGame.CharacterList
        FileProgress.Value = 20
        
        MergeLists QueryEngine.QueryList, MergeGame.QueryEngine.QueryList
        FileProgress.Value = 30
        
        MergeLists ItemList, MergeGame.ItemList
        FileProgress.Value = 40
        
        MergeLists RoteList, MergeGame.RoteList
        FileProgress.Value = 50
        
        MergeLists LocationList, MergeGame.LocationList
        FileProgress.Value = 60
        
        MergeLists APREngine.ActionList, MergeGame.APREngine.ActionList
        FileProgress.Value = 70
        
        MergeLists APREngine.PlotList, MergeGame.APREngine.PlotList
        FileProgress.Value = 80
        
        MergeLists APREngine.RumorList, MergeGame.APREngine.RumorList
        FileProgress.Value = 90
                
        FileProgress.Value = 100
        
    End If
    
    EnsureNoDuplicates
    
    Set MergeGame = Nothing

End Sub

Private Sub MergeLists(ByRef Receiver As LinkedList, ByRef Supplier As LinkedList)
'
' Name:         MergeLists
' Parameters:   Receiver            the list to receive new items
'               Supplier            the list to supply new items
' Description:  Merge newer items from the supplier list into the receiver list.
'               On name conflicts, decide using LastModified.
'

    Supplier.First
    Do Until Supplier.Off
        Receiver.MoveTo Supplier.Item.Name
        If Receiver.Off Then
            MergeResults = MergeResults & Supplier.Item.Name & vbCr & "(added)" & vbCr
            Receiver.InsertSorted Supplier.Item
            Supplier.Remove
        Else
            If Receiver.Item.LastModified < Supplier.Item.LastModified Then
                MergeResults = MergeResults & Supplier.Item.Name & vbCr & _
                               "(replaced)" & vbCr
                Receiver.Remove
                Receiver.InsertSorted Supplier.Item
                Supplier.Remove
            Else
                Supplier.MoveNext
            End If
        End If
    Loop

End Sub

Public Sub NewGame()
'
' Name:         NewGame
' Description:  Destroy all data, initializing a new game.
'
    PlayerList.Clear
    CharacterList.Clear
    ItemList.Clear
    RoteList.Clear
    LocationList.Clear
    AllRumorLists.Clear
    InfluenceUseList.Clear
    Calendar.Clear
    QueryEngine.QueryList.Clear
    APREngine.Initialize
    XPAwardList.Clear

    Calendar.LastModified = Now
    
    GameFile = ""
    ChronicleTitle = ""
    Website = ""
    EMail = ""
    Phone = ""
    UsualPlace = ""
    UsualTime = ""
    Description = ""
    ExtendedHealth = True
    EnforceHistory = True
    LinkTraitMaxes = True
    FileFormat = gvBinaryGame
    
    MenuSet.OpenMenus SlashPath(App.Path) & DefaultMenuFile, False
    If MenuSet.FileError Then _
        MsgBox "Warning: " & MenuSet.FileErrorMessage, _
        vbExclamation, "Open Game: Open Game Menus"
    
    LoadExchange SlashPath(App.Path) & DefaultInventoryFile

    QueryEngine.AddDefaultQueries
    AddDefaultXPAwards
    AddDefaultTemplates
    
    DataChanged = False
    
End Sub

Public Sub OpenGame(FileName As String, Optional LoadMenu As Boolean = True)
'
' Name:         OpenGame
' Parameters:   FileName        name of the file to open
'               LoadMenu        whether or not to load the menu file as well
' Description:  Prepare to open a new game.  Check the file header for its format
'               and call the appropriate routine.
'

    FileError = True
    FileErrorMessage = ""
    
    On Error GoTo OpenGame_DetectError
    
    Select Case DetectFileFormat(FileName)
        Case gvBinaryGame
            OpenGameBinary FileName, LoadMenu
        Case gvXML
            OpenGameXML FileName, LoadMenu
        Case gv23Game
            OpenGameV23 FileName, LoadMenu
        Case gvBinaryMenu
            FileErrorMessage = ShortFile(FileName) & _
                    " is a grapevine menu file: open it from Game -> Grapevine Menus."
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
    
    QueryEngine.AddDefaultQueries
    EnsureNoDuplicates
    APREngine.ReconnectAll True
    If XPAwardList.IsEmpty Then AddDefaultXPAwards
    If TemplateList.IsEmpty Then AddDefaultTemplates
    
    GoTo OpenGame_EndSub
    
OpenGame_DetectError:
    FileErrorMessage = "Error opening " & ShortFile(FileName) & ":" & vbCrLf & Err.Description
    Resume OpenGame_EndSub
    
OpenGame_EndSub:
    
End Sub

Private Sub OpenGameBinary(FileName As String, LoadMenu As Boolean)
'
' Name:         OpenGameBinary
' Parameters:   FileName        name of the file to open.
'               LoadMenu        whether or not to load the menu file as well
' Description:  Open a proprietary-format Grapevine game file.
'
    
    Dim FileNum As Integer
    Dim Verify As String
    Dim NewEntity As Object
    Dim MenuFileName As String
    Dim Version As Double
    Dim I As Integer
    Dim OneBit As Single
    Dim RCode As Integer
    Dim Size As Integer
    Dim TempBool As Boolean
    
    FileNum = FreeFile
    FileError = True
    DuplicateAll = False
    
    On Error GoTo OGB_AnyError
    
    FileErrorMessage = "Failed to open binary game file " & ShortFile(FileName) & ":"
    Open FileName For Binary As #FileNum
    GameFile = FileName
    
    FileErrorMessage = "Error reading binary file header:"
    GetStrB FileNum, Verify
    If Not Verify = BinHeaderGame Then
    
        FileErrorMessage = ShortFile(FileName) & " is not a Grapevine game file."
        Exit Sub
        
    End If
    
    Get #FileNum, , Version
    Get #FileNum, , Size
    FileFormat = gvBinaryGame
    OneBit = 100 / (Size + 1)
    
    PlayerList.Clear
    CharacterList.Clear
    ItemList.Clear
    RoteList.Clear
    LocationList.Clear
    
    AllRumorLists.Clear
    InfluenceUseList.Clear
    Calendar.Clear
    QueryEngine.QueryList.Clear
    APREngine.Initialize
    XPAwardList.Clear
    TemplateList.Clear
    
    GetStrB FileNum, ChronicleTitle
    GetStrB FileNum, Website
    GetStrB FileNum, EMail
    GetStrB FileNum, Phone
    GetStrB FileNum, UsualTime
    GetStrB FileNum, UsualPlace
    GetStrB FileNum, Description
    
    Get #FileNum, , ExtendedHealth
    Get #FileNum, , EnforceHistory
    If Version = 2.396 Then Get #FileNum, , TempBool
    If Version >= 2.397 Then
        GetStrB FileNum, STCommentStart
        GetStrB FileNum, STCommentEnd
        Get #FileNum, , LinkTraitMaxes
        If Version >= 2.399 Then GetStrB FileNum, RandomTraits
    End If
    
    FileErrorMessage = "Error reading/opening menu file:"
    GetStrB FileNum, MenuFileName
    If MenuFileName = "" Then MenuFileName = SlashPath(App.Path) & DefaultMenuFile
    
    FileProgress.Value = 0
    FileProgress.Max = 201
    If LoadMenu Then MenuSet.OpenMenus MenuFileName, False
    If FileProgress.Value = 0 Then
        FileProgress.Max = 101
    Else
        FileProgress.Value = 100
    End If
    If MenuSet.FileError Then _
        MsgBox "Warning: " & MenuSet.FileErrorMessage, _
        vbExclamation, "Open Game: Open Game Menus"
    
    FileErrorMessage = "Error reading calendar:"
    Calendar.InputFromBinary FileNum, Version
    
    If Version >= 2.397 Then
        
        Get #FileNum, , I
        Do Until I = 0
            Set NewEntity = New ExperienceAwardClass
            NewEntity.InputFromBinary FileNum, Version
            XPAwardList.Append NewEntity
            I = I - 1
        Loop
        
        Get #FileNum, , I
        Do Until I = 0
            Set NewEntity = New TemplateClass
            NewEntity.InputFromBinary FileNum, Version
            TemplateList.Append NewEntity
            I = I - 1
        Loop
    
    End If
    
    If Version >= 2.396 Then
        FileErrorMessage = "Error reading action and rumor settings:"
        APREngine.InputFromBinary FileNum, Version
    End If
    
    FileErrorMessage = "Error reading players:"
    Get #FileNum, , I
    Do Until I = 0
        Set NewEntity = New PlayerClass
        NewEntity.InputFromBinary FileNum, Version
        PlayerList.Append NewEntity
        AddFileProgress OneBit
        I = I - 1
    Loop
    
    FileErrorMessage = "Error reading characters:"
    Get #FileNum, , I
    Do Until I = 0
        Get #FileNum, , RCode
        Select Case RCode
            Case gvRaceVampire
                Set NewEntity = New VampireClass
            Case gvRaceWerewolf
                Set NewEntity = New WerewolfClass
            Case gvracemage
                Set NewEntity = New MageClass
            Case gvRaceChangeling
                Set NewEntity = New ChangelingClass
            Case gvRaceWraith
                Set NewEntity = New WraithClass
            Case gvRaceMortal
                Set NewEntity = New MortalClass
            Case gvRaceMummy
                Set NewEntity = New MummyClass
            Case gvRaceKueiJin
                Set NewEntity = New KueiJinClass
            Case gvRaceFera
                Set NewEntity = New FeraClass
            Case gvRaceVarious
                Set NewEntity = New VariousClass
            Case gvRaceHunter
                Set NewEntity = New HunterClass
            Case gvRaceDemon
                Set NewEntity = New DemonClass
        End Select
        NewEntity.InputFromBinary FileNum, Version
        CharacterList.Append NewEntity
        AddFileProgress OneBit
        I = I - 1
    Loop
        
    FileErrorMessage = "Error reading queries:"
    Get #FileNum, , I
    Do Until I = 0
        Set NewEntity = New QueryClass
        NewEntity.InputFromBinary FileNum, Version
        QueryEngine.QueryList.Append NewEntity
        AddFileProgress OneBit
        I = I - 1
    Loop
    
    FileErrorMessage = "Error reading items:"
    Get #FileNum, , I
    Do Until I = 0
        Set NewEntity = New ItemClass
        NewEntity.InputFromBinary FileNum, Version
        ItemList.Append NewEntity
        AddFileProgress OneBit
        I = I - 1
    Loop
    
    FileErrorMessage = "Error reading rotes:"
    Get #FileNum, , I
    Do Until I = 0
        Set NewEntity = New RoteClass
        NewEntity.InputFromBinary FileNum, Version
        RoteList.Append NewEntity
        AddFileProgress OneBit
        I = I - 1
    Loop
    
    FileErrorMessage = "Error reading locations:"
    Get #FileNum, , I
    Do Until I = 0
        Set NewEntity = New LocationClass
        NewEntity.InputFromBinary FileNum, Version
        LocationList.Append NewEntity
        AddFileProgress OneBit
        I = I - 1
    Loop
    
    FileErrorMessage = "Error reading actions:"
    Get #FileNum, , I
    Do Until I = 0
        Set NewEntity = New ActionClass
        NewEntity.InputFromBinary FileNum, Version
        ActionList.Append NewEntity
        AddFileProgress OneBit
        I = I - 1
    Loop
    
    FileErrorMessage = "Error reading plots:"
    Get #FileNum, , I
    Do Until I = 0
        Set NewEntity = New PlotClass
        NewEntity.InputFromBinary FileNum, Version
        APREngine.PlotList.Append NewEntity
        AddFileProgress OneBit
        I = I - 1
    Loop
    
    FileErrorMessage = "Error reading rumors:"
    Get #FileNum, , I
    Do Until I = 0
        Set NewEntity = New RumorClass
        NewEntity.InputFromBinary FileNum, Version
        RumorList.Append NewEntity
        AddFileProgress OneBit
        I = I - 1
    Loop
       
    If Version < 2.395 Then
        LoadExchange SlashPath(App.Path) & DefaultInventoryFile
    End If
    
    FileError = False
    DataChanged = False
    
    GoTo OGB_CloseFile
    
OGB_AnyError:
    FileErrorMessage = FileErrorMessage & vbCrLf & Err.Description
    GameFile = ""
    Resume OGB_CloseFile

OGB_CloseFile:
    Close #FileNum

End Sub

Private Sub OpenGameV23(FileName As String, LoadMenu As Boolean)
'
' Name:         OpenGame
' Parameters:   FileName        name of the file to open
'               LoadMenu        whether or not to load the menu file as well
' Description:  Clear all game data and load new data from an
'               old-format file.
'

    Dim Read As String
    Dim Total As Integer
    Dim X As Integer
    Dim FileNum As Integer
    Dim Version As String
    Dim OneBit As Single
    
    Dim NewPlayer As PlayerClass
    
    Dim OldDate As Date
    
    FileError = True
    FileNum = FreeFile
    
    FileErrorMessage = "Error opening " & ShortFile(FileName) & ":"
    On Error GoTo OGV23_AnyError
    OldDate = FileDateTime(FileName)
    Open FileName For Input As #FileNum
    GameFile = FileName
    
    FileFormat = gvBinaryGame
    Line Input #FileNum, Version

    If Not (Version = GameFileVersionTag0 Or _
            Version = GameFileVersionTag1 Or _
            Version = GameFileVersionTag2 Or _
            Version = GameFileVersionTag3 Or _
            Version = GameFileVersionTag4 Or _
            Version = GameFileVersionTag5) Then

        FileErrorMessage = ShortFile(FileName) & " is not a valid Grapevine game file."
        Exit Sub
        
    End If
    
    PlayerList.Clear
    CharacterList.Clear
    ItemList.Clear
    RoteList.Clear
    LocationList.Clear
    
    AllRumorLists.Clear
    InfluenceUseList.Clear
    Calendar.Clear
    QueryEngine.QueryList.Clear
    APREngine.Initialize
    XPAwardList.Clear
    TemplateList.Clear
    
    ChronicleTitle = ""
    Website = ""
    EMail = ""
    Phone = ""
    UsualPlace = ""
    UsualTime = ""
    Description = ""
    EnforceHistory = True
        
    FileProgress.Value = 0
    FileProgress.Max = 201
    If LoadMenu Then MenuSet.OpenMenus SlashPath(App.Path) & DefaultMenuFile, False
    If FileProgress.Value = 0 Then
        FileProgress.Max = 101
    Else
        FileProgress.Value = 100
    End If
    If MenuSet.FileError Then _
        MsgBox "Warning: " & MenuSet.FileErrorMessage, _
        vbExclamation, "Open Game: Open Game Menus"
    
    '
    ' Read the player list
    '
    FileErrorMessage = "Error reading the player list from disk:"
    Line Input #FileNum, Read
    Total = Val(Read)
    OneBit = 15 / (Total + 1)
    For X = 1 To Total
        Set NewPlayer = New PlayerClass
        NewPlayer.OldInputFromFile FileNum, OldDate, Version
        PlayerList.InsertSorted NewPlayer
        AddFileProgress OneBit
    Next X
    
    If Version >= GameFileVersionTag3 Then
        
        'Discard the Narrator list -- don't need it any more!
        FileErrorMessage = "Error reading the narrator list from disk:"
        Line Input #FileNum, Read
        Total = Val(Read)
        For X = 1 To Total
            Line Input #FileNum, Read
        Next X
    
    End If
    
    '
    ' Read the character list
    '
    FileErrorMessage = "Error reading the character list from disk:"
    Line Input #FileNum, Read
    Total = Val(Read)
    OneBit = 30 / (Total + 1)
    For X = 1 To Total
        Line Input #FileNum, Read
        AddCharacterFromFile Read, FileNum, OldDate, Version
        AddFileProgress OneBit
    Next X
    
    '
    ' Read the rumor list
    '
    FileErrorMessage = "Error reading the rumor list from disk:"
    Line Input #FileNum, Read
    Total = Val(Read)
    OneBit = 5 / (Total + 1)
    For X = 1 To Total
        APREngine.OldInputFromFile FileNum, Version
        AddFileProgress OneBit
    Next X

    FileErrorMessage = "Error reading the influence use list from disk:"
    Line Input #FileNum, Read
    Total = Val(Read)
    OneBit = 5 / (Total + 1)
    For X = 1 To Total
        APREngine.OldInputFromFile FileNum, Version
        AddFileProgress OneBit
    Next X
    
    If Version = GameFileVersionTag5 Then
        FileErrorMessage = "Error reading the health level setting:"
        Line Input #FileNum, Read
        ExtendedHealth = CBool(Read)
    End If

    FileErrorMessage = "Error closing the file:"
    Close #FileNum

    LoadExchange SlashPath(App.Path) & DefaultInventoryFile   'Add default items to the game

    FileError = False
    DataChanged = False
    
    OrderLists CharacterList                'Order the disciplines, etc according to menus
    
    GoTo OGV23_CloseFile
    
OGV23_AnyError:
    FileErrorMessage = FileErrorMessage & vbCrLf & Err.Description
    GameFile = ""
    Resume OGV23_CloseFile

OGV23_CloseFile:
    Close #FileNum

End Sub

Private Sub OpenGameXML(FileName As String, LoadMenu As Boolean)
'
' Name:         OpenGameXML
' Parameters:   FileName        name of the file to open
'               LoadMenu        whether or not to load the menu file as well
' Description:  Clear all game data and load new data from a file.
'

    Dim FileNum As Integer
    Dim XML As XMLReaderClass
    Dim NewEntity As Object
    Dim NewEntityList As LinkedList
    Dim MenuFileName As String
    Dim Version As Double
    Dim OneBit As Single
    
    FileNum = FreeFile
    FileError = True
    DuplicateAll = False
    Set XML = New XMLReaderClass
        
    On Error GoTo OGX_AnyError
    
    FileErrorMessage = "Failed to open " & ShortFile(FileName) & ":"
    Open FileName For Input As #FileNum
    GameFile = FileName

    XML.BeginReading FileNum
    If Left(XML.VersionTag, 5) = "<?xml" Then

        FileFormat = gvXML
        
        PlayerList.Clear
        CharacterList.Clear
        ItemList.Clear
        RoteList.Clear
        LocationList.Clear
        
        AllRumorLists.Clear
        InfluenceUseList.Clear
        Calendar.Clear
        QueryEngine.QueryList.Clear
        APREngine.Initialize
        XPAwardList.Clear
        TemplateList.Clear
        
        With XML

            .ReadTag
            If .Tag = "grapevine" Then
                            
                Version = .GetDouble("version")
                ChronicleTitle = .GetString("chronicle")
                Website = .GetString("website")
                EMail = .GetString("email")
                Phone = .GetString("phone")
                UsualTime = .GetString("usualtime")
                ExtendedHealth = .GetBoolean("extendedhealth", True)
                EnforceHistory = .GetBoolean("enforcehistory", True)
                STCommentStart = .GetString("stcommentstart", "[ST]")
                STCommentEnd = .GetString("stcommentend", "[/ST]")
                LinkTraitMaxes = .GetBoolean("linktraitmax", True)
                RandomTraits = .GetString("randomtraits", RandomTraits)
                
                OneBit = 100 / (.GetInteger("size", 99) + 1)
                
                MenuFileName = .GetString("menupath", SlashPath(App.Path) & DefaultMenuFile)
                
                FileProgress.Value = 0
                FileProgress.Max = 201
                If LoadMenu Then MenuSet.OpenMenus MenuFileName, False
                If FileProgress.Value = 0 Then
                    FileProgress.Max = 101
                Else
                    FileProgress.Value = 100
                End If
                If MenuSet.FileError Then _
                    MsgBox "Warning: " & MenuSet.FileErrorMessage, _
                    vbExclamation, "Open Game: Open Game Menus"
                
                Do Until .Tag = "grapevine" And .Terminal
                    
                    Set NewEntity = Nothing
                    Set NewEntityList = CharacterList
                    .ReadTag
                    
                    Select Case .Tag
                        Case "player"
                            Set NewEntity = New PlayerClass
                            Set NewEntityList = PlayerList
                        Case "vampire"
                            Set NewEntity = New VampireClass
                        Case "werewolf"
                            Set NewEntity = New WerewolfClass
                        Case "mage"
                            Set NewEntity = New MageClass
                        Case "changeling"
                            Set NewEntity = New ChangelingClass
                        Case "wraith"
                            Set NewEntity = New WraithClass
                        Case "mortal"
                            Set NewEntity = New MortalClass
                        Case "mummy"
                            Set NewEntity = New MummyClass
                        Case "kueijin"
                            Set NewEntity = New KueiJinClass
                        Case "fera"
                            Set NewEntity = New FeraClass
                        Case "various"
                            Set NewEntity = New VariousClass
                        Case "hunter"
                            Set NewEntity = New HunterClass
                        Case "demon"
                            Set NewEntity = New DemonClass
                        Case "query"
                            Set NewEntity = New QueryClass
                            Set NewEntityList = QueryEngine.QueryList
                        Case "item"
                            Set NewEntity = New ItemClass
                            Set NewEntityList = ItemList
                        Case "rote"
                            Set NewEntity = New RoteClass
                            Set NewEntityList = RoteList
                        Case "location"
                            Set NewEntity = New LocationClass
                            Set NewEntityList = LocationList
                        Case "action"
                            Set NewEntity = New ActionClass
                            Set NewEntityList = APREngine.ActionList
                        Case "plot"
                            Set NewEntity = New PlotClass
                            Set NewEntityList = APREngine.PlotList
                        Case "rumor"
                            Set NewEntity = New RumorClass
                            Set NewEntityList = APREngine.RumorList
                        Case "calendar"
                            Calendar.InputFromFile XML, Version
                        Case "aprsettings"
                            APREngine.InputFromFile XML, Version
                        Case "award"
                            Set NewEntity = New ExperienceAwardClass
                            Set NewEntityList = XPAwardList
                        Case "template"
                            Set NewEntity = New TemplateClass
                            Set NewEntityList = TemplateList
                        Case "usualplace"
                            If Not .Terminal Then
                                .ReadTag
                                UsualPlace = .GetTagText
                            End If
                        Case "description"
                            If Not .Terminal Then
                                .ReadTag
                                Description = .GetTagText
                            End If
                    End Select
                    
                    If Not NewEntity Is Nothing Then
                        NewEntity.InputFromFile XML, Version
                        If NewEntityList Is ActionList Or NewEntityList Is RumorList Then
                            APREngine.InsertSorted NewEntityList, NewEntity
                        Else
                            NewEntityList.InsertSorted NewEntity
                        End If
                        AddFileProgress OneBit
                    End If
                                        
                Loop
            
                If Version < 2.395 Then
                    LoadExchange SlashPath(App.Path) & DefaultInventoryFile
                End If
            
                DataChanged = False
                FileError = False
                                
            Else 'not a grapevine game file!
            
                FileErrorMessage = ShortFile(FileName) & " is not a Grapevine game file."
                GameFile = ""
                
            End If
        
        End With

    Else

        FileErrorMessage = ShortFile(FileName) & " is not a Grapevine game file."
        GameFile = ""
        
    End If

    GoTo OGX_CloseFile
    
OGX_AnyError:
    FileErrorMessage = FileErrorMessage & vbCrLf & Err.Description
    GameFile = ""
    Resume OGX_CloseFile

OGX_CloseFile:
    Set XML = Nothing
    Close #FileNum

End Sub

Private Sub OrderLists(SourceList As LinkedList)
'
' Name:         OrderLists
' Paramaters:   SourceList          List from which to pull characters for ordering
' Description:  Attempt to order v2.3 characters' disciplines, rituals, merits, etc. by the
'               values in the current menus.
'

    Dim PowerList As LinkedList
    Dim OneBit As Single

    Set PowerList = New LinkedList
        
    OneBit = (FileProgress.Max - FileProgress.Value) / 16
    
    ' Vampire Disciplines
    SourceList.First
    Do Until SourceList.Off
        If SourceList.Item.RaceCode = gvRaceVampire Then
            PowerList.Append SourceList.Item.DisciplineList
        End If
        SourceList.MoveNext
    Loop
    MenuSet.SortListsByMenu PowerList, "Disciplines"
    PowerList.Clear
    AddFileProgress OneBit
    
    ' Vampire Rituals
    SourceList.First
    Do Until SourceList.Off
        If SourceList.Item.RaceCode = gvRaceVampire Then
            PowerList.Append SourceList.Item.RitualList
        End If
        SourceList.MoveNext
    Loop
    MenuSet.DeepSortListsByMenu PowerList, "Rituals"
    PowerList.Clear
    AddFileProgress OneBit
    
    ' Werewolf Gifts
    SourceList.First
    Do Until SourceList.Off
        If SourceList.Item.RaceCode = gvRaceWerewolf Then
            PowerList.Append SourceList.Item.GiftList
        End If
        SourceList.MoveNext
    Loop
    MenuSet.SortListsByMenu PowerList, "Gifts, Werewolf"
    PowerList.Clear
    AddFileProgress OneBit
    
    ' Werewolf Rites
    SourceList.First
    Do Until SourceList.Off
        If SourceList.Item.RaceCode = gvRaceWerewolf Then
            PowerList.Append SourceList.Item.RiteList
        End If
        SourceList.MoveNext
    Loop
    MenuSet.DeepSortListsByMenu PowerList, "Rites, Werewolf"
    PowerList.Clear
    AddFileProgress OneBit
    
    ' Mortal Numina
    SourceList.First
    Do Until SourceList.Off
        If SourceList.Item.RaceCode = gvRaceMortal Then
            PowerList.Append SourceList.Item.NuminaList
        End If
        SourceList.MoveNext
    Loop
    MenuSet.SortListsByMenu PowerList, "Numina"
    PowerList.Clear
    AddFileProgress OneBit
    
    ' Changeling Arts
    SourceList.First
    Do Until SourceList.Off
        If SourceList.Item.RaceCode = gvRaceChangeling Then
            PowerList.Append SourceList.Item.ArtList
        End If
        SourceList.MoveNext
    Loop
    MenuSet.SortListsByMenu PowerList, "Arts"
    PowerList.Clear
    AddFileProgress OneBit
    
    ' Changeling Realms
    SourceList.First
    Do Until SourceList.Off
        If SourceList.Item.RaceCode = gvRaceChangeling Then
            PowerList.Append SourceList.Item.RealmList
        End If
        SourceList.MoveNext
    Loop
    MenuSet.SortListsByMenu PowerList, "Realms"
    PowerList.Clear
    AddFileProgress OneBit
    
    ' Wraith Arcanoi
    SourceList.First
    Do Until SourceList.Off
        If SourceList.Item.RaceCode = gvRaceWraith Then
            PowerList.Append SourceList.Item.ArcanoiList
        End If
        SourceList.MoveNext
    Loop
    MenuSet.SortListsByMenu PowerList, "Arcanoi"
    PowerList.Clear
    AddFileProgress OneBit
    
    ' Fera Gifts
    SourceList.First
    Do Until SourceList.Off
        If SourceList.Item.RaceCode = gvRaceFera Then
            PowerList.Append SourceList.Item.GiftList
        End If
        SourceList.MoveNext
    Loop
    MenuSet.SortListsByMenu PowerList, "Gifts, Fera"
    PowerList.Clear
    AddFileProgress OneBit
    
    ' Fera Rites
    SourceList.First
    Do Until SourceList.Off
        If SourceList.Item.RaceCode = gvRaceFera Then
            PowerList.Append SourceList.Item.RiteList
        End If
        SourceList.MoveNext
    Loop
    MenuSet.DeepSortListsByMenu PowerList, "Rites, Fera"
    PowerList.Clear
    AddFileProgress OneBit
    
    ' Kuei-Jin Disciplines
    SourceList.First
    Do Until SourceList.Off
        If SourceList.Item.RaceCode = gvRaceKueiJin Then
            PowerList.Append SourceList.Item.DisciplineList
        End If
        SourceList.MoveNext
    Loop
    MenuSet.SortListsByMenu PowerList, "Disciplines, Kuei-Jin"
    PowerList.Clear
    AddFileProgress OneBit
    
    ' Kuei-Jin Rites
    SourceList.First
    Do Until SourceList.Off
        If SourceList.Item.RaceCode = gvRaceKueiJin Then
            PowerList.Append SourceList.Item.RiteList
        End If
        SourceList.MoveNext
    Loop
    MenuSet.SortListsByMenu PowerList, "Rites, Kuei-Jin"
    PowerList.Clear
    AddFileProgress OneBit
    
    ' Mummy Hekau
    SourceList.First
    Do Until SourceList.Off
        If SourceList.Item.RaceCode = gvRaceMummy Then
            PowerList.Append SourceList.Item.HekauList
        End If
        SourceList.MoveNext
    Loop
    MenuSet.DeepSortListsByMenu PowerList, "Hekau"
    PowerList.Clear
    AddFileProgress OneBit
    
    ' Mage Spheres -- already good.
    
    ' Various Powers
    SourceList.First
    Do Until SourceList.Off
        If SourceList.Item.RaceCode = gvRaceVarious Then
            PowerList.Append SourceList.Item.PowerList
        End If
        SourceList.MoveNext
    Loop
    MenuSet.DeepSortListsByMenu PowerList, "Powers"
    PowerList.Clear
    AddFileProgress OneBit
    
    Set PowerList = Nothing
    
    '
    ' Assign point values to Merits and Flaws
    '
    SourceList.First
    Do Until SourceList.Off
        MenuSet.AssignMeritsAndFlaws SourceList.Item
        SourceList.MoveNext
    Loop
    AddFileProgress OneBit

End Sub

Public Function OutputID() As Integer
'
' Name:         OutputID
' Description:  Return an ID that tells the OutputEngineClass how to manipulate this object
'
    OutputID = oidNone

End Function

Public Sub Rename(ObjType As QueryInventoryType, OldName As String, NewName As String)
'
' Name:         Rename
' Parameters:   ObjType         Type of object being renamed
'               OldName         Old name of object
'               NewName         New Name of object
' Description:  Rename an object and all references to it in the system.
'

    Dim SourceList As LinkedList
    
    Select Case ObjType
        Case qiCharacters:  Set SourceList = CharacterList
        Case qiPlayers:     Set SourceList = PlayerList
        Case qiItems:       Set SourceList = ItemList
        Case qiLocations:   Set SourceList = LocationList
        Case qiRotes:       Set SourceList = RoteList
    End Select
    
    SourceList.MoveTo OldName
    If Not SourceList.Off Then
    
        Dim Thing As Object
        Set Thing = SourceList.Item
        
        SourceList.Remove
        Thing.Name = NewName
        Thing.LastModified = Now
        SourceList.InsertSorted Thing
        
        Game.DataChanged = True
        
        Select Case ObjType
            Case qiCharacters
            
                CharacterList.First
                Do Until CharacterList.Off
                    Set Thing = CharacterList.Item
                    Select Case Thing.RaceCode
                        Case gvRaceVampire
                            If Thing.Sire = OldName Then Thing.Sire = NewName
                            Thing.BondList.RenameTrait OldName, NewName
                            Thing.BoonList.First
                            Do Until Thing.BoonList.Off
                                If Thing.BoonList.Item.CharName = OldName Then
                                    Thing.BoonList.Item.CharName = NewName
                                End If
                                Thing.BoonList.MoveNext
                            Loop
                            Thing.RefreshBoonTraits
                        Case gvRaceKueiJin
                            Thing.GuanxiList.RenameTrait OldName, NewName
                        Case gvRaceMortal
                            If Thing.Regnant = OldName Then Thing.Regnant = NewName
                    End Select
                    CharacterList.MoveNext
                Loop
                
                LocationList.First
                Do Until LocationList.Off
                    If LocationList.Item.Owner = OldName Then LocationList.Item.Owner = NewName
                    LocationList.MoveNext
                Loop
                
                PlotList.First
                Do Until PlotList.Off
                    PlotList.Item.CastList.RenameTrait OldName, NewName
                    PlotList.MoveNext
                Loop
                            
                APREngine.Reassign ActionList, OldName, NewName
            
            Case qiPlayers
            
                CharacterList.First
                Do Until CharacterList.Off
                    If CharacterList.Item.Player = OldName Then CharacterList.Item.Player = NewName
                    If CharacterList.Item.Narrator = OldName Then CharacterList.Item.Narrator = NewName
                    If CharacterList.Item.RaceCode = gvRaceWraith Then
                        If CharacterList.Item.ShadowPlayer = OldName Then
                            CharacterList.Item.ShadowPlayer = NewName
                        End If
                    End If
                    CharacterList.MoveNext
                Loop
                
                PlotList.First
                Do Until PlotList.Off
                    If PlotList.Item.Narrator = OldName Then PlotList.Item.Narrator = NewName
                    PlotList.MoveNext
                Loop
            
            Case qiItems
                                        
                CharacterList.First
                Do Until CharacterList.Off
                    CharacterList.Item.EquipmentList.RenameTrait OldName, NewName
                    CharacterList.MoveNext
                Loop
            
            Case qiLocations
                            
                CharacterList.First
                Do Until CharacterList.Off
                    CharacterList.Item.HangoutList.RenameTrait OldName, NewName
                    CharacterList.MoveNext
                Loop
                
                LocationList.First
                Do Until LocationList.Off
                    LocationList.Item.LinkList.RenameTrait OldName, NewName
                    LocationList.MoveNext
                Loop
            
            Case qiRotes
                
                CharacterList.First
                Do Until CharacterList.Off
                    If CharacterList.Item.RaceCode = gvracemage Then
                        CharacterList.Item.RoteList.RenameTrait OldName, NewName
                    End If
                    CharacterList.MoveNext
                Loop
                
        End Select
        
    End If
    
End Sub

Public Sub SaveGame(FileName As String)
'
' Name:         SaveGame
' Parameters:   FileName        the name under which to save the file
' Description:  Save the game data under a given filename.
'

    HideSTFromFile = False
    
    XPAwardList.Memorize    'Memorize and Remember: done to preserve game state. Important for Autosaving.
    TemplateList.Memorize
    PlayerList.Memorize
    CharacterList.Memorize
    QueryEngine.QueryList.Memorize
    ItemList.Memorize
    RoteList.Memorize
    LocationList.Memorize
    ActionList.Memorize
    PlotList.Memorize
    RumorList.Memorize
    
    Select Case FileFormat
        Case gvBinaryGame
            SaveGameBinary FileName
        Case gvXML
            SaveGameXML FileName
        Case Else
            FileError = True
            FileErrorMessage = "Can't save in file format " & CStr(FileFormat)
    End Select

    XPAwardList.Remember
    TemplateList.Remember
    PlayerList.Remember
    CharacterList.Remember
    QueryEngine.QueryList.Remember
    ItemList.Remember
    RoteList.Remember
    LocationList.Remember
    ActionList.Remember
    PlotList.Remember
    RumorList.Remember

End Sub

Private Sub SaveGameBinary(FileName As String)
'
' Name:         SaveGameBinary
' Parameters:   FileName        name of the binary file to save
' Description:  Sve a game file in a proprietary binary format.
'
    
    Dim FileNum As Integer
    Dim Size As Integer
    Dim OneBit As Single
    Dim RelativeMenu As String
    
    FileNum = FreeFile
    FileFormat = gvBinaryGame
    FileError = True
    Size = EntityCount
    OneBit = 100 / (Size + 1)
    FileProgress = 0
    
    On Error GoTo SGB_AnyError
    
    FileErrorMessage = "Error deleting old game file " & ShortFile(FileName) & ":"
    If Dir(FileName) <> "" Then Kill FileName
    
    FileErrorMessage = "Error opening binary game file " & ShortFile(FileName) & " for writing:"
    Open FileName For Binary As #FileNum
    
    FileErrorMessage = "Error writing binary file header:"
    PutStrB FileNum, BinHeaderGame
    Put #FileNum, , ThisVersion
    Put #FileNum, , Size
    
    FileErrorMessage = "Error writing initial game data:"
    PutStrB FileNum, ChronicleTitle
    PutStrB FileNum, Website
    PutStrB FileNum, EMail
    PutStrB FileNum, Phone
    PutStrB FileNum, UsualTime
    PutStrB FileNum, UsualPlace
    PutStrB FileNum, Description
    Put #FileNum, , ExtendedHealth
    Put #FileNum, , EnforceHistory
    PutStrB FileNum, STCommentStart
    PutStrB FileNum, STCommentEnd
    Put #FileNum, , LinkTraitMaxes
    PutStrB FileNum, RandomTraits
    
    RelativeMenu = GetRelativeName(MenuSet.FilePath, FileName)
    PutStrB FileNum, RelativeMenu
    
    FileErrorMessage = "Error writing game calendar:"
    Calendar.OutputToBinary FileNum
    
    FileErrorMessage = "Error writing XP awards:"
    Put #FileNum, , CInt(XPAwardList.Count)
    XPAwardList.First
    Do Until XPAwardList.Off
        XPAwardList.Item.OutputToBinary FileNum
        XPAwardList.MoveNext
    Loop
    
    FileErrorMessage = "Error writing Templates:"
    Put #FileNum, , CInt(TemplateList.Count)
    TemplateList.First
    Do Until TemplateList.Off
        TemplateList.Item.OutputToBinary FileNum
        TemplateList.MoveNext
    Loop
    
    FileErrorMessage = "Error writing action/rumor settings:"
    APREngine.OutputToBinary FileNum
    
    FileErrorMessage = "Error writing players:"
    Put #FileNum, , CInt(PlayerList.Count)
    PlayerList.First
    Do Until PlayerList.Off
        PlayerList.Item.OutputToBinary FileNum
        PlayerList.MoveNext
        FileProgress = FileProgress + OneBit
    Loop

    FileErrorMessage = "Error writing characters:"
    Put #FileNum, , CInt(CharacterList.Count)
    CharacterList.First
    Do Until CharacterList.Off
        Put #FileNum, , CInt(CharacterList.Item.RaceCode)
        CharacterList.Item.OutputToBinary FileNum
        CharacterList.MoveNext
        FileProgress = FileProgress + OneBit
    Loop
    
    FileErrorMessage = "Error writing queries:"
    With QueryEngine.QueryList
        .MoveTo RecentSearchName
        If .Off Then
            Put #FileNum, , CInt(.Count)
        Else
            Put #FileNum, , CInt(.Count) - 1
        End If
        .First
        Do Until .Off
            If Not .Item.Name = RecentSearchName Then
                .Item.OutputToBinary FileNum
            End If
            .MoveNext
            FileProgress = FileProgress + OneBit
        Loop
    End With
    
    FileErrorMessage = "Error writing items:"
    Put #FileNum, , CInt(ItemList.Count)
    ItemList.First
    Do Until ItemList.Off
        ItemList.Item.OutputToBinary FileNum
        ItemList.MoveNext
        FileProgress = FileProgress + OneBit
    Loop
    
    FileErrorMessage = "Error writing rotes:"
    Put #FileNum, , CInt(RoteList.Count)
    RoteList.First
    Do Until RoteList.Off
        RoteList.Item.OutputToBinary FileNum
        RoteList.MoveNext
        FileProgress = FileProgress + OneBit
    Loop
    
    FileErrorMessage = "Error writing locations:"
    Put #FileNum, , CInt(LocationList.Count)
    LocationList.First
    Do Until LocationList.Off
        LocationList.Item.OutputToBinary FileNum
        LocationList.MoveNext
        FileProgress = FileProgress + OneBit
    Loop
        
    FileErrorMessage = "Error writing actions:"
    Put #FileNum, , CInt(ActionList.Count)
    ActionList.First
    Do Until ActionList.Off
        ActionList.Item.OutputToBinary FileNum
        ActionList.MoveNext
        FileProgress = FileProgress + OneBit
    Loop
    
    FileErrorMessage = "Error writing plots:"
    Put #FileNum, , CInt(PlotList.Count)
    PlotList.First
    Do Until PlotList.Off
        PlotList.Item.OutputToBinary FileNum
        PlotList.MoveNext
        FileProgress = FileProgress + OneBit
    Loop
    
    FileErrorMessage = "Error writing rumors:"
    Put #FileNum, , CInt(RumorList.Count)
    RumorList.First
    Do Until RumorList.Off
        RumorList.Item.OutputToBinary FileNum
        RumorList.MoveNext
        FileProgress = FileProgress + OneBit
    Loop
        
    DataChanged = False
    FileError = False
    GameFile = FileName
    
    GoTo SGB_CloseFile

SGB_AnyError:
    FileErrorMessage = FileErrorMessage & vbCrLf & Err.Description
    Resume SGB_CloseFile

SGB_CloseFile:
    Close #FileNum

End Sub

Private Sub SaveGameXML(FileName As String)
'
' Name:         SaveGameXML
' Parameters:   FileName        the name under which to save the file
' Description:  Save the game data in XML under a given filename.
'

    Dim FileNum As Integer
    Dim XML As XMLWriterClass
    Dim Size As Integer
    Dim OneBit As Single
    Dim RelativeMenu As String
    
    FileError = True
    FileNum = FreeFile
    Set XML = New XMLWriterClass
    FileFormat = gvXML
    Size = EntityCount
    OneBit = 100 / (Size + 1)
    FileProgress = 0
    
    On Error GoTo SGX_AnyError
    FileErrorMessage = "Failed to open " & ShortFile(FileName) & ":"
    Open FileName For Output As #FileNum
    
    XML.BeginXML FileNum
    
    XML.BeginTag "grapevine"
        XML.WriteAttribute "version", ThisVersion
        
        FileErrorMessage = "Error writing game attributes:"
        XML.WriteAttribute "chronicle", ChronicleTitle
        XML.WriteAttribute "website", Website
        XML.WriteAttribute "email", EMail
        XML.WriteAttribute "phone", Phone
        XML.WriteAttribute "usualtime", UsualTime
        XML.WriteAttribute "extendedhealth", ExtendedHealth, True
        XML.WriteAttribute "enforcehistory", EnforceHistory, True
        XML.WriteAttribute "stcommentstart", STCommentStart, "[ST]"
        XML.WriteAttribute "stcommentend", STCommentEnd, "[/ST]"
        XML.WriteAttribute "linktraitmax", LinkTraitMaxes, True
        XML.WriteAttribute "randomtraits", RandomTraits
        
        RelativeMenu = GetRelativeName(MenuSet.FilePath, FileName)
        XML.WriteAttribute "menupath", RelativeMenu
        XML.WriteAttribute "size", Size
        
        XML.BeginTag "usualplace"
            XML.WriteStringData UsualPlace
        XML.EndTag
        XML.BeginTag "description"
            XML.WriteStringData Description
        XML.EndTag
        
        FileErrorMessage = "Error writing calendar:"
        Calendar.OutputToFile XML
    
        FileErrorMessage = "Error writing XP awards:"
        XPAwardList.First
        Do Until XPAwardList.Off
            XPAwardList.Item.OutputToFile XML
            XPAwardList.MoveNext
        Loop
    
        FileErrorMessage = "Error writing Templates:"
        TemplateList.First
        Do Until TemplateList.Off
            TemplateList.Item.OutputToFile XML
            TemplateList.MoveNext
        Loop
    
        FileErrorMessage = "Error writing action/rumor settings:"
        APREngine.OutputToFile XML
    
        FileErrorMessage = "Error writing players:"
        PlayerList.First
        Do Until PlayerList.Off
            PlayerList.Item.OutputToFile XML
            PlayerList.MoveNext
            FileProgress = FileProgress + OneBit
        Loop
    
        FileErrorMessage = "Error writing characters:"
        CharacterList.First
        Do Until CharacterList.Off
            CharacterList.Item.OutputToFile XML
            CharacterList.MoveNext
            FileProgress = FileProgress + OneBit
        Loop
    
        FileErrorMessage = "Error writing queries:"
        With QueryEngine.QueryList
            .First
            Do Until .Off
                If Not .Item.Name = RecentSearchName Then
                    .Item.OutputToFile XML
                End If
                .MoveNext
                FileProgress = FileProgress + OneBit
            Loop
        End With
    
        FileErrorMessage = "Error writing items:"
        ItemList.First
        Do Until ItemList.Off
            ItemList.Item.OutputToFile XML
            ItemList.MoveNext
            FileProgress = FileProgress + OneBit
        Loop
        
        FileErrorMessage = "Error writing rotes:"
        RoteList.First
        Do Until RoteList.Off
            RoteList.Item.OutputToFile XML
            RoteList.MoveNext
            FileProgress = FileProgress + OneBit
        Loop
        
        FileErrorMessage = "Error writing locations:"
        LocationList.First
        Do Until LocationList.Off
            LocationList.Item.OutputToFile XML
            LocationList.MoveNext
            FileProgress = FileProgress + OneBit
        Loop
        
        FileErrorMessage = "Error writing Actions:"
        ActionList.First
        Do Until ActionList.Off
            ActionList.Item.OutputToFile XML
            ActionList.MoveNext
            FileProgress = FileProgress + OneBit
        Loop
        
        FileErrorMessage = "Error writing Plots:"
        PlotList.First
        Do Until PlotList.Off
            PlotList.Item.OutputToFile XML
            PlotList.MoveNext
            FileProgress = FileProgress + OneBit
        Loop
        
        FileErrorMessage = "Error writing Rumors:"
        RumorList.First
        Do Until RumorList.Off
            RumorList.Item.OutputToFile XML
            RumorList.MoveNext
            FileProgress = FileProgress + OneBit
        Loop
        
    
    XML.EndTag
    
    FileError = False
    DataChanged = False
    GameFile = FileName

    GoTo SGX_CloseFile

SGX_AnyError:
    FileErrorMessage = FileErrorMessage & vbCrLf & Err.Description
    Resume SGX_CloseFile

SGX_CloseFile:

    Close #FileNum
    Set XML = Nothing

End Sub

Public Sub SaveExchange(FileName As String, Format As FileFormatType, _
                        ByRef Master As LinkedList, HideST As Boolean)
' Name:         SaveExchange
' Parameters:   FileName        the name of the file to save as
'               Format          Format to save it in
'               Master          A list of lists of entities to save, in the following order:
'                               1. List of Game settings
'                                   - If element 1 is not nothing, save Calendar.
'                                   - If element 2 is not nothing, save Action/Rumor settings.
'                                   - If element 3 is not nothing, save XP Awards.
'                                   - If element 4 is not nothing, save Templates.
'                               2. List of Players
'                               3. List of Characters
'                               4. List of Queries
'                               5. List of Items
'                               6. List of Rotes
'                               7. List of Locations
'                               8. List of Actions
'                               9. List of Plots
'                               10. List of Rumors
'               HideST          Whether or not to hide ST comments from the exchange file
' Description:  Save a selection of the game's data (players and characters)
'               to a file.
'

    HideSTFromFile = HideST
    
    XPAwardList.Memorize    'Memorize and Remember: done to preserve game state. Important for Autosaving.
    TemplateList.Memorize
    PlayerList.Memorize
    CharacterList.Memorize
    QueryEngine.QueryList.Memorize
    ItemList.Memorize
    RoteList.Memorize
    LocationList.Memorize
    ActionList.Memorize
    PlotList.Memorize
    RumorList.Memorize
    
    Select Case Format
        Case gvBinaryExchange
            SaveExchangeBinary FileName, Master
        Case gvXML
            SaveExchangeXML FileName, Master
        Case Else
            FileError = True
            FileErrorMessage = "Can't save in file format " & CStr(FileFormat)
    End Select

    XPAwardList.Remember
    TemplateList.Remember
    PlayerList.Remember
    CharacterList.Remember
    QueryEngine.QueryList.Remember
    ItemList.Remember
    RoteList.Remember
    LocationList.Remember
    ActionList.Remember
    PlotList.Remember
    RumorList.Remember

End Sub

Private Sub SaveExchangeBinary(FileName As String, ByRef Master As LinkedList)
'
' Name:         SaveExchangeBinary
' Parameters:   FileName        the name of the file to save as
'               Master          a linkedlist of lists with data to save
'                               (See SaveExchange comments)
' Description:  Save a selection of the game's data (players and characters)
'               to a binary file.
'

    Dim FileNum As Integer
    Dim SubList As LinkedList
    Dim L As Integer
    
    FileNum = FreeFile
    FileError = True
    
    On Error GoTo SEB_AnyError
    
    FileErrorMessage = "Error deleting old exchange file " & ShortFile(FileName) & ":"
    If Dir(FileName) <> "" Then Kill FileName
    
    FileErrorMessage = "Error opening binary exchange file " & ShortFile(FileName) & " for writing:"
    Open FileName For Binary As #FileNum
    
    FileErrorMessage = "Error writing binary file header:"
    PutStrB FileNum, BinHeaderExchange
    Put #FileNum, , ThisVersion
    
    FileErrorMessage = "Error writing binary exchange file:"
    
    Master.First
    Set SubList = Master.Item               'deal with the game settings list first
    
    SubList.First
    If SubList.Item = 1 Then     'Game Calendar
        Put #FileNum, , 1
        Calendar.OutputToBinary FileNum
    Else
        Put #FileNum, , 0
    End If
    
    SubList.MoveNext
    If SubList.Item = 1 Then     'APR Settings
        Put #FileNum, , 1
        APREngine.OutputToBinary FileNum
    Else
        Put #FileNum, , 0
    End If
    
    SubList.MoveNext
    If SubList.Item = 1 Then     'XP Awards
        Put #FileNum, , CInt(XPAwardList.Count)
        XPAwardList.First
        Do Until XPAwardList.Off
            XPAwardList.Item.OutputToBinary FileNum
            XPAwardList.MoveNext
        Loop
    Else
        Put #FileNum, , 0
    End If
    
    SubList.MoveNext
    If SubList.Item = 1 Then     'Template settings
        Put #FileNum, , CInt(TemplateList.Count)
        TemplateList.First
        Do Until TemplateList.Off
            TemplateList.Item.OutputToBinary FileNum
            TemplateList.MoveNext
        Loop
    Else
        Put #FileNum, , 0
    End If
    
    
    Master.MoveNext
    L = 2                                   'The third list requires a little extra data before each
                                            'character entity: watch for and perform that when L=3
    Do Until Master.Off
        Set SubList = Master.Item
        Put #FileNum, , CInt(SubList.Count)
        SubList.First
        Do Until SubList.Off
            If L = 3 Then Put #FileNum, , CInt(SubList.Item.RaceCode)
            SubList.Item.OutputToBinary FileNum
            SubList.MoveNext
        Loop
        Master.MoveNext
        L = L + 1
    Loop
        
    FileError = False
    
    GoTo SEB_CloseFile

SEB_AnyError:
    FileErrorMessage = FileErrorMessage & vbCrLf & Err.Description
    Resume SEB_CloseFile

SEB_CloseFile:
    Close #FileNum
    
End Sub

Private Sub SaveExchangeXML(FileName As String, Master As LinkedList)
'
' Name:         SaveExchange
' Parameters:   FileName        the name of the file to save as
'               Master          a linkedlist of lists with data to save
'                               (See SaveExchange comments)
' Description:  Save a selection of the game's data (players and characters)
'               to an XML file.
'

    Dim FileNum As Integer
    Dim XML As New XMLWriterClass
    Dim SubList As LinkedList
    
    FileError = True
    FileNum = FreeFile
    Set XML = New XMLWriterClass
    
    On Error GoTo SEX_AnyError
    
    FileErrorMessage = "Failed to open " & ShortFile(FileName) & ":"
    Open FileName For Output As #FileNum
    
    XML.BeginXML FileNum
    XML.BeginTag "grapevine"
        XML.WriteAttribute "version", ThisVersion
    
        FileErrorMessage = "Error writing XML exchange file:"
        Master.First
        Set SubList = Master.Item               'deal with the game settings list first
        
        SubList.First
        If SubList.Item = 1 Then      'Game Calendar
            Calendar.OutputToFile XML
        End If
        
        SubList.MoveNext
        If SubList.Item = 1 Then     'APR Settings
            APREngine.OutputToFile XML
        End If
        
        SubList.MoveNext
        If SubList.Item = 1 Then     'XP Awards
            XPAwardList.First
            Do Until XPAwardList.Off
                XPAwardList.Item.OutputToFile XML
                XPAwardList.MoveNext
            Loop
        End If
        
        SubList.MoveNext
        If SubList.Item = 1 Then     'Template settings
            TemplateList.First
            Do Until TemplateList.Off
                TemplateList.Item.OutputToFile XML
                TemplateList.MoveNext
            Loop
        End If
                
        Master.MoveNext
        Do Until Master.Off
            Set SubList = Master.Item
            SubList.First
            Do Until SubList.Off
                SubList.Item.OutputToFile XML
                SubList.MoveNext
            Loop
            Master.MoveNext
        Loop
        
    XML.EndTag
    
    FileError = False
    
    GoTo SEX_CloseFile

SEX_AnyError:
    FileErrorMessage = FileErrorMessage & vbCrLf & Err.Description
    Resume SEX_CloseFile

SEX_CloseFile:
    Close #FileNum
    Set XML = Nothing
    
End Sub

Public Function STFilter(ByRef Text As String) As String
'
' Name:         STFilter
' Parameters:   Text        the text to be filtered
' Description:  Filter out text between ST comment tags.
'

    STFilter = Text
    If Not (STCommentStart = "" Or STCommentEnd = "") Then
            
        Dim STLocStart As Long
        Dim STLocEnd As Long
        
        STLocStart = InStr(1, STFilter, STCommentStart, vbTextCompare)
        If Not STLocStart = 0 Then
            Do
                STLocEnd = InStr(STLocStart + Len(STCommentStart), STFilter, STCommentEnd, vbTextCompare)
                If STLocEnd = 0 Then STLocEnd = Len(STFilter) - Len(STCommentEnd)
                STFilter = Left(STFilter, STLocStart - 1) & Mid(STFilter, STLocEnd + Len(STCommentEnd) + 1)
                STLocStart = InStr(1, STFilter, STCommentStart, vbTextCompare)
            Loop Until STLocStart = 0
            STFilter = TrimWhiteSpace(STFilter)
        End If
    
    End If

End Function

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Destroy all game objects.
'

    PlayerList.Clear
    CharacterList.Clear
    ItemList.Clear
    RoteList.Clear
    LocationList.Clear
    AllRumorLists.Clear
    InfluenceUseList.Clear
    MenuSet.Clear
    Calendar.Clear
    QueryEngine.QueryList.Clear
    APREngine.ActionList.Clear
    APREngine.PlotList.Clear
    APREngine.RumorList.Clear
    XPAwardList.Clear
    TemplateList.Clear
    
    Set PlayerList = Nothing
    Set CharacterList = Nothing
    Set ItemList = Nothing
    Set RoteList = Nothing
    Set LocationList = Nothing
    Set AllRumorLists = Nothing
    Set InfluenceUseList = Nothing
    Set MenuSet = Nothing
    Set Calendar = Nothing
    Set QueryEngine = Nothing
    Set APREngine = Nothing
    Set XPAwardList = Nothing
    Set TemplateList = Nothing
    
End Sub
