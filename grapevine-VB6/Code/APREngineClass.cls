VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APREngineClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ActionList As LinkedList
Public PlotList As LinkedList
Public RumorList As LinkedList

Public BackgroundActions As LinkedTraitList
Public ActionsPerLevel As LinkedTraitList

Public PersonalActions As Integer
Public AddCommon As Boolean
Public CarryUnused As Boolean

Public PublicRumors As Boolean
Public PersonalRumors As Boolean
Public RaceRumors As Boolean
Public GroupRumors As Boolean
Public SubgroupRumors As Boolean
Public InfluenceRumors As Boolean
Public PreviousRumors As Boolean
Public CopyPrevious As Boolean

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Create the needed lists.
'

    Set ActionList = New LinkedList
    Set PlotList = New LinkedList
    Set RumorList = New LinkedList
    
    Set BackgroundActions = New LinkedTraitList
    Set ActionsPerLevel = New LinkedTraitList
    
    BackgroundActions.Initialize "Backgrounds", False, False, False, ldSimple
    ActionsPerLevel.Initialize "Actions", False, False, False, ldSimple
    
    Initialize
    
End Sub

Public Sub Initialize()
'
' Name:         Initialize
' Description:  Clear and initialize the APREngine for a new game.
'
    
    ActionList.Clear
    PlotList.Clear
    RumorList.Clear
    
    BackgroundActions.Clear
    ActionsPerLevel.Clear
    
    BackgroundActions.Append "Contacts"
    BackgroundActions.Append "Resources"
    
    ActionsPerLevel.Append "1", "2"
    ActionsPerLevel.Append "2", "4"
    ActionsPerLevel.Append "3", "6"
    ActionsPerLevel.Append "4", "8"
    ActionsPerLevel.Append "5", "10"
    ActionsPerLevel.Append "6", "12"
    ActionsPerLevel.Append "7", "14"
    ActionsPerLevel.Append "8", "16"
    ActionsPerLevel.Append "9", "18"
    ActionsPerLevel.Append "10", "20"
    
    AddCommon = False
    CarryUnused = False
    PublicRumors = True
    PersonalRumors = False
    RaceRumors = False
    GroupRumors = False
    SubgroupRumors = False
    InfluenceRumors = True
    PreviousRumors = True
    CopyPrevious = False

End Sub

Public Function GetLink(Link As CauseEffectNode) As Object
'
' Name:         GetLink
' Parameters:   Link            Node with data about the link to get
' Description:  Locate and return the object pointed to by the given node
'

    Select Case Link.Target
        Case aprAction
            MoveToPair ActionList, Link.When, Link.Item
            If Not ActionList.Off Then
                Dim Act As ActionClass
                Set Act = ActionList.Item
                Act.MoveTo Link.Subitem
                If Not Act.Off Then Set GetLink = Act.SubAction
            End If
        Case aprPlot
            PlotList.MoveTo Link.Item
            If Not PlotList.Off Then
                Dim Plot As PlotClass
                Set Plot = PlotList.Item
                Plot.MoveTo Link.When
                If Not Plot.Off Then Set GetLink = Plot.PlotDev
            End If
        Case aprRumor
            MoveToPair RumorList, Link.When, Link.Item
            If Not RumorList.Off Then
                Dim Rumor As RumorClass
                Set Rumor = RumorList.Item
                Rumor.MoveTo CInt(Link.Subitem)
                If Not Rumor.Off Then Set GetLink = Rumor.SubRumor
            End If
        Case Else
            Set GetLink = Nothing
    End Select

End Function

Public Sub AddStandardRumors(When As Date)
'
' Name:         AddStandardRumors
' Parameters:   When        Date for the rumors
' Description:  Add all the standard rumors to the RumorList for the given date,
'               according tot he preferences the user has specified.
'

    Dim Existing As StringSet
    Dim NewRumor As RumorClass
    Dim Char As Object
    Dim NewTitle As String
    Dim Value As Variant
    
    Screen.MousePointer = vbHourglass
    
    Set Existing = New StringSet
    Existing.Add ""
    
    MoveToFirstDate RumorList, When
    Do Until RumorList.Off
        Existing.Add RumorList.Item.Title
        MoveToNextDate RumorList, When
    Loop
    
    If PreviousRumors Then      'Add copies of the previous date's rumors
    
        Game.Calendar.MoveTo When
        Game.Calendar.MovePrevious
        If Not Game.Calendar.Off Then   ' if there *is* a prior game
        
            Dim PrevDate As Date
            
            PrevDate = Game.Calendar.GetGameDate
            
            MoveToFirstDate RumorList, PrevDate
            Do Until RumorList.Off
                If Not Existing.Has(RumorList.Item.Title) Then
                    Set NewRumor = RumorList.Item.Clone(When, CopyPrevious)
                    Game.APREngine.InsertSorted RumorList, NewRumor
                    Existing.Add NewRumor.Title
                End If
                MoveToNextDate RumorList, PrevDate
            Loop
            
        End If
    
    End If

    If PublicRumors Then
        If Not Existing.Has(PublicRumorTitle) Then
            Set NewRumor = New RumorClass
            NewRumor.InitializeQueryRumor PublicRumorTitle, When, rtGeneral
            Game.APREngine.InsertSorted RumorList, NewRumor
            Existing.Add NewRumor.Title
        End If
    End If

    CharacterList.First
    Do Until CharacterList.Off
        Set Char = CharacterList.Item
        If Char.Status = "Active" Then
        
            NewTitle = Char.Name
            If PersonalRumors And Not Existing.Has(NewTitle) Then
                Set NewRumor = New RumorClass
                NewRumor.InitializeQueryRumor NewTitle, When, rtPersonal
                NewRumor.Query.AddClause qkName, NewTitle, 0, qcEquals, False
                Game.APREngine.InsertSorted RumorList, NewRumor
                Existing.Add NewTitle
            End If
        
            NewTitle = Char.Race
            If RaceRumors And Not Existing.Has(NewTitle) Then
                Set NewRumor = New RumorClass
                NewRumor.InitializeQueryRumor NewTitle, When, rtRace
                NewRumor.Query.AddClause qkRace, NewTitle, 0, qcEquals, False
                Game.APREngine.InsertSorted RumorList, NewRumor
                Existing.Add NewTitle
            End If
        
            NewTitle = Char.Group
            If GroupRumors And Not Existing.Has(NewTitle) Then
                Set NewRumor = New RumorClass
                NewRumor.InitializeQueryRumor NewTitle, When, rtGroup
                NewRumor.Query.AddClause qkGroup, NewTitle, 0, qcEquals, False
                Game.APREngine.InsertSorted RumorList, NewRumor
                Existing.Add NewTitle
            End If
            
            NewTitle = Char.Subgroup
            If SubgroupRumors And Not Existing.Has(NewTitle) Then
                Set NewRumor = New RumorClass
                NewRumor.InitializeQueryRumor NewTitle, When, rtSubgroup
                NewRumor.Query.AddClause qkSubgroup, NewTitle, 0, qcEquals, False
                Game.APREngine.InsertSorted RumorList, NewRumor
                Existing.Add NewTitle
            End If
        
            If InfluenceRumors Then
                
                Dim Influences As LinkedTraitList
                Call Char.GetValue(qkInfluences, Value)
                If Not IsNull(Value) Then
                
                    Set Influences = Value
                    Influences.First
                    Do Until Influences.Off
                        NewTitle = Influences.Trait.Name & " Influence"
                        If Not Existing.Has(NewTitle) Then
                            Set NewRumor = New RumorClass
                            NewRumor.InitializeMultiRumor NewTitle, When, _
                                     qkInfluences, Influences.Trait.Name
                            Game.APREngine.InsertSorted RumorList, NewRumor
                            Existing.Add NewTitle
                        End If
                        Influences.MoveNext
                    Loop
                
                End If
                
            End If
                
        End If
        CharacterList.MoveNext
    Loop

    Set Existing = Nothing
    
    Game.DataChanged = True
    
    Screen.MousePointer = vbDefault

End Sub

Public Sub MoveToFirstDate(List As LinkedList, ToDate As Date)
'
' Name:         MoveToFirstDate
' Parameters:   List        either the actionlist or the rumorlist
'               ToDate      Date to move to
' Description:  Move to the first item in the list matching the given date.
'

    List.First
    If List Is RumorList Then
        Do Until List.Off
            If List.Item.RumorDate = ToDate Then Exit Do
            List.MoveNext
        Loop
    ElseIf List Is ActionList Then
        Do Until List.Off
            If List.Item.ActDate = ToDate Then Exit Do
            List.MoveNext
        Loop
    End If

End Sub

Public Sub MoveToNextDate(List As LinkedList, ToDate As Date)
'
' Name:         MoveToNextDate
' Parameters:   List        either the actionlist or the rumorlist
'               ToDate      Date to move to
' Description:  Move to the next item in the list matching the given date.
'

    If List Is RumorList Then
        Do
            List.MoveNext
            If List.Off Then Exit Do
        Loop Until List.Item.RumorDate = ToDate
    ElseIf List Is ActionList Then
        Do
            List.MoveNext
            If List.Off Then Exit Do
        Loop Until List.Item.ActDate = ToDate
    End If

End Sub

Public Sub MoveToPair(List As LinkedList, ToDate As Date, Title As String)
'
' Name:         MoveToPair
' Parameters:   List        either the actionlist or the rumorlist
'               ToDate      Date to move to
'               Title       Rumor Title or List Character to move to
' Description:  Move to the item in the list matching the given date / title combo.
'

    MoveToFirstDate List, ToDate
    If List Is RumorList Then
        Do Until List.Off
            If List.Item.Title = Title Then Exit Do
            MoveToNextDate List, ToDate
        Loop
    ElseIf List Is ActionList Then
        Do Until List.Off
            If List.Item.CharName = Title Then Exit Do
            MoveToNextDate List, ToDate
        Loop
    End If

End Sub

Public Sub MoveToFirstTitle(List As LinkedList, Title As String)
'
' Name:         MoveToFirstTitle
' Parameters:   List        either the actionlist or the rumorlist
'               Title       Title or character name to move to
' Description:  Move to the first item in the list matching the given title
'

    List.First
    If List Is RumorList Then
        Do Until List.Off
            If List.Item.Title = Title Then Exit Do
            List.MoveNext
        Loop
    ElseIf List Is ActionList Then
        Do Until List.Off
            If List.Item.CharName = Title Then Exit Do
            List.MoveNext
        Loop
    End If

End Sub

Public Sub MoveToNextTitle(List As LinkedList, Title As String)
'
' Name:         MoveToNextTitle
' Parameters:   List        either the actionlist or the rumorlist
'               Title       Title or character name to move to
' Description:  Move to the next item in the list matching the given title
'

    If List Is RumorList Then
        Do
            List.MoveNext
            If List.Off Then Exit Do
        Loop Until List.Item.Title = Title
    ElseIf List Is ActionList Then
        Do
            List.MoveNext
            If List.Off Then Exit Do
        Loop Until List.Item.CharName = Title
    End If

End Sub

Public Sub MoveToEarlierTitle(List As LinkedList, Title As String, EndDate As Date)
'
' Name:         MoveToNextTitle
' Parameters:   List        either the actionlist or the rumorlist
'               Title       Title or character name to move to
'               EndDate     Date to move to the predecessor of
' Description:  Move to the next item in the list matching the given title and
'               whose date is the last one coming before the end date
'

    Dim FoundDate As Date
    Dim TestDate As Date
    
    FoundDate = 0

    MoveToFirstTitle List, Title
    Do Until List.Off
        If List Is RumorList Then
            TestDate = List.Item.RumorDate
        Else
            TestDate = List.Item.ActDate
        End If
        If FoundDate < TestDate And TestDate < EndDate Then FoundDate = TestDate
        MoveToNextTitle List, Title
    Loop

    If FoundDate > 0 Then
        MoveToPair List, FoundDate, Title
    End If
    
End Sub

Public Sub DeleteDate(DelDate As Date)
'
' Name:         DeleteDate
' Parameters:   DelDate     the date to delete
' Description:  Delete all actions and rumors corresponding to the given date;
'               clear the causes and effects for plot developments on that date
'

    Dim Plot As PlotClass

    MoveToFirstDate ActionList, DelDate
    Do Until ActionList.Off
        ActionList.Remove
        MoveToFirstDate ActionList, DelDate
    Loop

    MoveToFirstDate RumorList, DelDate
    Do Until RumorList.Off
        RumorList.Remove
        MoveToFirstDate RumorList, DelDate
    Loop
    
    ReconnectAll False
    
End Sub

Public Sub InsertSorted(List As LinkedList, Entity As Object)
'
' Name:         InsertSorted
' Parameters:   List            ActionList or RumorList to insert into
'               Entity          Action or Rumor object to insert
' Description:  Insert an action or rumor into its list, sorting by
'               descending date then ascending name.
'

    Dim FindDate As Date
    Dim FindName As String
    Dim FindCat As RumorCategoryType
    Dim IsAction As Boolean
    Dim Act As ActionClass
    Dim Rumor As RumorClass
    
    IsAction = List Is ActionList
    
    If IsAction Then
        FindDate = Entity.ActDate
        FindName = Entity.CharName
    Else
        FindDate = Entity.RumorDate
        FindName = Entity.Title
        FindCat = Entity.Category
    End If
    
    List.First
    Do Until List.Off
        If IsAction Then
            Set Act = List.Item
            If Act.ActDate < FindDate Then Exit Do
            If Act.ActDate = FindDate And Act.CharName > FindName Then Exit Do
        Else
            Set Rumor = List.Item
            If Rumor.RumorDate < FindDate Then Exit Do
            If Rumor.RumorDate = FindDate Then
                If Rumor.Category > FindCat Then Exit Do
                If Rumor.Category = FindCat And Rumor.Title > FindName Then Exit Do
            End If
        End If
        List.MoveNext
    Loop
    
    If List.Off Then
        List.Append Entity
    Else
        List.Insert Entity
    End If

End Sub

Public Sub ReconnectAll(CreateDates As Boolean)
'
' Name:         ReconnectAll
' Parameters:   CreateDates     Whether or not to create missing dates
' Description:  When a game or exchange file is loaded, ensure the following:
'               1. All dates appearing in Action and Rumor lists are created.
'               2. All Effects of actions and plot developments create reciprocal Causes.
'                  Effects that don't exist result in a removal of those links.
'

    Dim DateSet As StringSet
    Dim Act As ActionClass
    Dim Plot As PlotClass
    Dim Rumor As RumorClass
    Dim Target As Object
    Dim P1 As Integer
    Dim P2 As Integer
    Dim P3 As Integer
    
    ' Step 1.
    
    If CreateDates Then
    
        Set DateSet = New StringSet
        
        With Game.Calendar
            .First
            Do Until .Off
                DateSet.Add CStr(.GetGameDate)
                .MoveNext
            Loop
        End With
        
        ActionList.First
        Do Until ActionList.Off
            Set Act = ActionList.Item
            If Not DateSet.Has(CStr(Act.ActDate)) Then
                Game.Calendar.Insert Act.ActDate, Game.UsualPlace, Game.UsualPlace, ""
                DateSet.Add CStr(Act.ActDate)
            End If
            ActionList.MoveNext
        Loop
        
        RumorList.First
        Do Until RumorList.Off
            Set Rumor = RumorList.Item
            If Not DateSet.Has(CStr(Rumor.RumorDate)) Then
                Game.Calendar.Insert Rumor.RumorDate, Game.UsualPlace, Game.UsualPlace, ""
                DateSet.Add CStr(Rumor.RumorDate)
            End If
            RumorList.MoveNext
        Loop
        
        Set DateSet = Nothing

    End If

    'Step 2a.  Clear all Cause links from the chronicle items.
    
    ActionList.First
    Do Until ActionList.Off
        Set Act = ActionList.Item
        Act.First
        Do Until Act.Off
            Act.SubAction.Causes.Clear
            Act.MoveNext
        Loop
        ActionList.MoveNext
    Loop

    PlotList.First
    Do Until PlotList.Off
        Set Plot = PlotList.Item
        Plot.First
        Do Until Plot.Off
            Plot.PlotDev.Causes.Clear
            Plot.MoveNext
        Loop
        PlotList.MoveNext
    Loop

    RumorList.First
    Do Until RumorList.Off
        Set Rumor = RumorList.Item
        Rumor.First
        Do Until Rumor.Off
            Rumor.SubRumor.Causes.Clear
            Rumor.MoveNext
        Loop
        RumorList.MoveNext
    Loop

    'Step 2b.  Rebuild the Cause links from the Effect links.
    
    For P1 = 0 To ActionList.Count - 1        'MoveToPlace loop are needed because Getlink messes with cursors
        ActionList.MoveToPlace P1
        Set Act = ActionList.Item
        For P2 = 0 To Act.Count - 1
            Act.MoveToPlace P2
            With Act.SubAction.Effects
                P3 = .Count - 1
                Do Until P3 = -1
                    .MoveToPlace P3
                    Set Target = GetLink(.Link)
                    If Not Target Is Nothing Then
                        Target.Causes.AppendNode Act.SubAction
                    Else
                        Act.LastModified = Now
                        .Remove
                    End If
                    P3 = P3 - 1
                Loop
            End With
        Next P2
        ActionList.MoveNext
    Next P1

    For P1 = 0 To PlotList.Count - 1
        PlotList.MoveToPlace P1
        Set Plot = PlotList.Item
        For P2 = 0 To Plot.Count - 1
            Plot.MoveToPlace P2
            With Plot.PlotDev.Effects
                P3 = .Count - 1
                Do Until P3 = -1
                    .MoveToPlace P3
                    Set Target = GetLink(.Link)
                    If Not Target Is Nothing Then
                        Target.Causes.AppendNode Plot.PlotDev
                    Else
                        Plot.LastModified = Now
                        .Remove
                    End If
                    P3 = P3 - 1
                Loop
            End With
        Next P2
        PlotList.MoveNext
    Next P1

End Sub

Public Sub PrepareRumorOutput(PrepDate As Date)
'
' Name:         PrepareRumorOutput
' Parameters:   PrepDate    Dates of rumors to prepare
' Description:  Tell all the rumors of a given date to re-generate their recipient sets.
'

    MoveToFirstDate RumorList, PrepDate
    Do Until RumorList.Off
        RumorList.Item.PrepareRecipients
        MoveToNextDate RumorList, PrepDate
    Loop

End Sub

Public Sub Reassign(SourceList As LinkedList, OldName As String, NewName As String, _
                    Optional OldDate As Date = 0, Optional NewDate As Date = 0)
'
' Name:         Reassign
' Parameters:   Sourcelist      List holding item to reassign
'               OldName         Old name of item
'               NewName         New name of item
'               OldDate         Old Date of item
'               NewDate         New Date of item
' Description:  Given a new Name/Title and Date to an Action, Plot or Rumor.
'

    Dim APR As Object
    Dim Links As CauseEffectList
    Dim EndLinks As CauseEffectList
    Dim APRCode As APRType
    Dim Entity As Object
    Dim AllDates As Boolean
    
    AllDates = (OldDate = 0) And (NewDate = 0) And Not (SourceList Is PlotList)
    
    If SourceList Is PlotList Then
        SourceList.MoveTo NewName
    Else
        If AllDates Then
            MoveToFirstTitle SourceList, NewName
        Else
            MoveToPair SourceList, NewDate, NewName
        End If
    End If
    If Not SourceList.Off Then Exit Sub

    If SourceList Is PlotList Then
        SourceList.MoveTo OldName
    Else
        If AllDates Then
            MoveToFirstTitle SourceList, OldName
        Else
            MoveToPair SourceList, OldDate, OldName
        End If
    End If
    
    Do Until SourceList.Off
        
        If AllDates Then
            If SourceList Is ActionList Then
                OldDate = SourceList.Item.ActDate
            Else
                OldDate = SourceList.Item.RumorDate
            End If
            NewDate = OldDate
        End If
        
        Set APR = SourceList.Item
        If SourceList Is PlotList Then
            APR.Name = NewName
            APRCode = aprPlot
        ElseIf SourceList Is ActionList Then
            APR.ActDate = NewDate
            APR.CharName = NewName
            APRCode = aprAction
        Else
            APR.RumorDate = NewDate
            APR.Title = NewName
            APRCode = aprRumor
        End If
        
        APR.First
        Do Until APR.Off
            If SourceList Is PlotList Then
                Set Links = APR.PlotDev.Causes
            ElseIf SourceList Is ActionList Then
                Set Links = APR.SubAction.Causes
            Else
                Set Links = APR.SubRumor.Causes
            End If
            Links.First
            Do Until Links.Off
                Set Entity = GetLink(Links.Link)
                If Not Entity Is Nothing Then
                    Entity.Parent.LastModified = Now
                    Set EndLinks = Entity.Effects
                    EndLinks.First
                    Do Until EndLinks.Off
                        If APRCode = EndLinks.Link.Target Then
                            If EndLinks.Link.Item = OldName Then EndLinks.Link.Item = NewName
                            If Not APRCode = aprPlot Then
                                If EndLinks.Link.When = OldDate Then EndLinks.Link.When = NewDate
                            End If
                        End If
                        EndLinks.MoveNext
                    Loop
                Else
                    If Links.Link.Target = APRCode And Links.Link.Item = OldName Then
                        If APRCode = aprPlot Then
                            Links.Link.Item = NewName
                        ElseIf Links.Link.When = OldDate Then
                            Links.Link.Item = NewName
                            Links.Link.When = NewDate
                        End If
                    End If
                End If
                Links.MoveNext
            Loop
            If Not SourceList Is RumorList Then
                If SourceList Is PlotList Then
                    Set Links = APR.PlotDev.Effects
                ElseIf SourceList Is ActionList Then
                    Set Links = APR.SubAction.Effects
                End If
                Links.First
                Do Until Links.Off
                    Set Entity = GetLink(Links.Link)
                    If Not Entity Is Nothing Then
                        Entity.Parent.LastModified = Now
                        Set EndLinks = Entity.Causes
                        EndLinks.First
                        Do Until EndLinks.Off
                            If APRCode = EndLinks.Link.Target Then
                                If EndLinks.Link.Item = OldName Then EndLinks.Link.Item = NewName
                                If Not APRCode = aprPlot Then
                                    If EndLinks.Link.When = OldDate Then EndLinks.Link.When = NewDate
                                End If
                            End If
                            EndLinks.MoveNext
                        Loop
                    Else
                        If Links.Link.Target = APRCode And Links.Link.Item = OldName Then
                            If APRCode = aprPlot Then
                                Links.Link.Item = NewName
                            ElseIf Links.Link.When = OldDate Then
                                Links.Link.Item = NewName
                                Links.Link.When = NewDate
                            End If
                        End If
                    End If
                    Links.MoveNext
                Loop
            End If
            APR.MoveNext
        Loop
    
        If AllDates Then
            MoveToNextTitle SourceList, OldName
        Else
            SourceList.Last
            SourceList.MoveNext
        End If
    
    Loop
    
End Sub

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write the object to an XML file.
'

    With XML
    
        .BeginTag "aprsettings"
    
            .WriteAttribute "personalactions", PersonalActions
            .WriteAttribute "addcommon", AddCommon, False
            .WriteAttribute "carryunused", CarryUnused, False
            .WriteAttribute "publicrumors", PublicRumors, False
            .WriteAttribute "personalrumors", PersonalRumors, False
            .WriteAttribute "racerumors", RaceRumors, False
            .WriteAttribute "grouprumors", GroupRumors, False
            .WriteAttribute "subgrouprumors", SubgroupRumors, False
            .WriteAttribute "influencerumors", InfluenceRumors, False
            .WriteAttribute "previousrumors", PreviousRumors, False
            .WriteAttribute "copyprevious", CopyPrevious, False
            
            ActionsPerLevel.OutputToFile XML
            BackgroundActions.OutputToFile XML
            
        .EndTag
        
    End With

End Sub

Public Sub InputFromFile(XML As XMLReaderClass, Version As Double)
'
' Name:         InputFromFile
' Parameters:   XML         an XMLReaderClass that just read the opening
'                           tag of this object
'               Version     version of the file format
' Description:  Read the object from an XML file.
'

    If XML.Tag = "aprsettings" Then
        
        With XML
        
            PersonalActions = .GetInteger("personalactions")
            AddCommon = .GetBoolean("addcommon")
            CarryUnused = .GetBoolean("carryunused")
            PublicRumors = .GetBoolean("publicrumors")
            PersonalRumors = .GetBoolean("personalrumors")
            RaceRumors = .GetBoolean("racerumors")
            GroupRumors = .GetBoolean("grouprumors")
            SubgroupRumors = .GetBoolean("subgrouprumors")
            InfluenceRumors = .GetBoolean("influencerumors")
            PreviousRumors = .GetBoolean("previousrumors")
            CopyPrevious = .GetBoolean("copyprevious")
            
            Do Until .Tag = "aprsettings" And .Terminal
                .ReadTag
                If .Tag = "traitlist" Then
                    Select Case .GetString("name")
                        Case "Backgrounds": BackgroundActions.InputFromFile XML, Version
                        Case "Actions":     ActionsPerLevel.InputFromFile XML, Version
                    End Select
                End If
            Loop
            
        End With
        
    Else
        Err.Raise vbObjectError + 8, "APR Settings", _
                  "Wrong file input to action/plot/rumor settings"
    End If

End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this object to a binary file.
'

    Put #FileNum, , PersonalActions
    Put #FileNum, , AddCommon
    Put #FileNum, , CarryUnused
    Put #FileNum, , PublicRumors
    Put #FileNum, , PersonalRumors
    Put #FileNum, , RaceRumors
    Put #FileNum, , GroupRumors
    Put #FileNum, , SubgroupRumors
    Put #FileNum, , InfluenceRumors
    Put #FileNum, , PreviousRumors
    Put #FileNum, , CopyPrevious

    BackgroundActions.OutputToBinary FileNum
    ActionsPerLevel.OutputToBinary FileNum
    
End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this object from a binary file.
'

    Get #FileNum, , PersonalActions
    Get #FileNum, , AddCommon
    If Version >= 2.397 Then Get #FileNum, , CarryUnused
    Get #FileNum, , PublicRumors
    Get #FileNum, , PersonalRumors
    Get #FileNum, , RaceRumors
    Get #FileNum, , GroupRumors
    Get #FileNum, , SubgroupRumors
    Get #FileNum, , InfluenceRumors
    Get #FileNum, , PreviousRumors
    Get #FileNum, , CopyPrevious

    BackgroundActions.InputFromBinary FileNum, Version
    ActionsPerLevel.InputFromBinary FileNum, Version
     
End Sub

Public Sub OldInputFromFile(FileNum As Integer, Version As String)
'
' Name:         OldInputFromFile
' Parameters:   FileNum     number of an open file
'               Version     version tag of the file format
'               Actions     TRUE iff reading actions; FALSE iff reading rumors
' Description:  Read a list of rumors or actions from a v2.3 file.
'

    Dim Read As String
    
    Dim DateStamp As Date
    Dim Category As Integer     ' Categories: General = 0, Personal = 1, Group = 2, Influence = 3, Use = 4
    Dim Recipient As String
    Dim Level As Integer
    Dim Text As String
    Dim UsedBy As String
    Dim Exclude As Boolean
    
    Dim Action As ActionClass
    Dim Rumor As RumorClass
    
    Line Input #FileNum, Read
    DateStamp = CDate(Read)
    
    Line Input #FileNum, Read
    Do Until Read = "<-End of LinkedRumorList Items->"
        
        Category = Val(Read)
        Line Input #FileNum, Read
        Recipient = Read
        Line Input #FileNum, Read
        Level = Val(Read)
        Text = ReadLongField(FileNum, "<-End Rumor Text->")
        
        If Version <> GameFileVersionTag0 Then
            Line Input #FileNum, Read
            UsedBy = Read
            Line Input #FileNum, Read
            Exclude = CBool(Read)
        End If
        
        If Category = 4 Then        ' Create Action: Old 'Influence Use' Category
        
            MoveToPair ActionList, DateStamp, UsedBy
            
            If ActionList.Off Then
                Set Action = New ActionClass
                Action.CharName = UsedBy
                Action.ActDate = DateStamp
                InsertSorted ActionList, Action
            Else
                Set Action = ActionList.Item
            End If
        
            Action.MoveTo Recipient
            If Action.Off Then
                Action.Add Recipient, 0, 0, 0, 0
                Action.SubAction.Action = "Level " & CStr(Level) & ": " & Text
            Else
                Action.SubAction.Action = Action.SubAction.Action & vbCrLf & vbCrLf _
                                        & "Level " & CStr(Level) & ": " & Text
            End If
            
            If Level > Action.SubAction.Level Then
                Action.SubAction.Level = Level
                ActionsPerLevel.MoveTo CStr(Level)
                Action.SubAction.Total = ActionsPerLevel.Trait.Number
            End If
            
            Action.Done = True
            
        Else                        ' Create rumor
            
            Select Case Category
                Case 0  ' General Rumor (Public Knowledge)
                    Set Rumor = New RumorClass
                    Rumor.InitializeQueryRumor Recipient, DateStamp, rtRace
                    Rumor.First
                    Rumor.SubRumor.Rumor = Text
                    If Recipient = "Public Knowledge" Then
                        Rumor.Category = rtGeneral
                        Rumor.Title = PublicRumorTitle
                    Else
                        Rumor.Query.AddClause qkRace, Recipient, 0, qcEquals, False
                    End If
                    Rumor.Done = True
                    InsertSorted RumorList, Rumor
                Case 1  ' Personal Rumor
                    Set Rumor = New RumorClass
                    Rumor.InitializeQueryRumor Recipient, DateStamp, rtPersonal
                    Rumor.First
                    Rumor.SubRumor.Rumor = Text
                    Rumor.Query.AddClause qkName, Recipient, 0, qcEquals, False
                    Rumor.Done = True
                    InsertSorted RumorList, Rumor
                Case 2  ' Group Rumor
                    Set Rumor = New RumorClass
                    Rumor.InitializeQueryRumor Recipient, DateStamp, rtGroup
                    Rumor.First
                    Rumor.SubRumor.Rumor = Text
                    Rumor.Query.AddClause qkGroup, Recipient, 0, qcEquals, False
                    Rumor.Done = True
                    InsertSorted RumorList, Rumor
                Case 3  ' Influence Rumor
                    MoveToPair RumorList, DateStamp, Recipient & " Influence"
                    If RumorList.Off Then
                        Set Rumor = New RumorClass
                        Rumor.InitializeMultiRumor Recipient & " Influence", _
                                                   DateStamp, qkInfluences, Recipient
                        InsertSorted RumorList, Rumor
                    Else
                        Set Rumor = RumorList.Item
                    End If
                    Rumor.MoveTo Level
                    If Not Rumor.Off Then Rumor.SubRumor.Rumor = Text
                    Rumor.Done = True
            End Select
                                                                                
        End If
        
        Line Input #FileNum, Read
    
    Loop

End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Destroy the action, plot and rumor lists.
'

    Set ActionList = Nothing
    Set PlotList = Nothing
    Set RumorList = Nothing
    
    Set BackgroundActions = Nothing
    Set ActionsPerLevel = Nothing
    
End Sub
