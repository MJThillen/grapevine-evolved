VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QueryEngineClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public Enum QueryKeyType                        'An enumeration of data types
    qtError = -1                                    'No type: Error
    qtField = 0                                     'String
    qtNumber = 1                                    'Integer or Decimal
    qtTraitList = 2                                 'LinkedTraitList
    qtDate = 3                                      'Date
    qtBoolean = 4                                   'Boolean
End Enum

Public Enum StatisticType                       'An enumeration of the statistics that can be run
    stDistribution = 0                              'Distribution (General)
    stDistinctDistribution = 1                      'Distinct Trait Distribution
    stSpecificDistribution = 2                      'Specific Trait Distribution
    stMaxima = 3                                    'Maxima of trait lists or number data
    stSums = 4                                      'Sums of trait lists or number data
End Enum

Public Enum QueryCompareType                    'An enumeration of comparisons for queries
    qcContains = 0                                  'The field/list contains the data.
    qcEquals = 1                                    'The field/list/number equals the data.
    qcAtLeast = 2                                   'The number is greater than or equal to the data.
    qcGreater = 3                                   'The number is greater than the data.
    qcLess = 4                                      'The number is less than the data.
    qcNoMore = 5                                    'The number is no more than than the data.
    qcContainsExactly = 6                           'The list contains exactly this number of a trait.
    qcContainsAtLeast = 7                           'The list contains at least this number of a trait.
    qcContainsMore = 8                              'The list contains more than this number of a trait.
    qcContainsLess = 9                              'The list contains less than this number of a trait (but not none).
    qcContainsNoMore = 10                           'The list contains no more than this number of a trait (but not none).
    qcTotals = 11                                   'The list totals this amount of traits.
    qcTotalsAtLeast = 12                            'The list totals at least this amount of traits.
    qcTotalsMore = 13                               'The list totals more than this amount of traits.
    qcTotalsNoMore = 14                             'The list totals no more than this amount of traits.
    qcTotalsLess = 15                               'The list totals less than this amount of traits.
    qcContainsNote = 16                             'The list contains any trait with this note.
    qcIsTrue = 17                                   'The boolean value is true.
    qcIsFalse = 18                                  'The boolean value is false.
End Enum

Public Enum QueryInventoryType                  'An enumeration of inventories that one can query.
    qiNone = 0
    qiCharacters = 1                                'The characters
    qiPlayers = 2                                   'The players
    qiItems = 4                                     'The items
    qiRotes = 128                                   'The rotes
    qiLocations = 8                                 'The locations
    qiActions = 16                                  'The actions
    qiPlots = 32                                    'The plots
    qiRumors = 64                                   'The rumors
End Enum

Public TitlesToKeys As Collection               'A collection of keys, themselves keyed by display title
Public KeysToTypes As Collection                'A collection of query keys' data types
Public KeysToTitles As Collection               'A collection of query keys' display titles
Public KeysToInventories As Collection          'A collection of query keys' associated inventories

Public Results As LinkedList                    'The results of running a query
Public Values As LinkedList                     'The matching values of that query
Public SortList As LinkedList                   'The values by which that query is sorted

Public StatResults As LinkedList                'The results of a Statistics run
Public NumberSet As Collection                  'The numbers corresponding to each statistic find
Public MatchSet As Collection                   'A collection of traitlists of names and numbers that were found
Public Total As Double                          'Total to come out of a statistic
Public Maximum As Double                        'Maximum value to come out of a statistic

Public QueryList As LinkedList                  'The list of queries

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Associate keys with titles and data types.  (See PublicQueryKeys.bas.)
'

    Dim FileNum As Integer
    Dim Key As String
    Dim Title As String
    Dim SType As String
    Dim SInv As String
    Dim KType As QueryKeyType
    Dim KInv As QueryInventoryType
    
    Set TitlesToKeys = New Collection
    Set KeysToTypes = New Collection
    Set KeysToTitles = New Collection
    Set KeysToInventories = New Collection
    
    Set Results = New LinkedList
    Set Values = New LinkedList
    Set SortList = New LinkedList
    
    Set StatResults = New LinkedList
    Set NumberSet = New Collection
    Set MatchSet = New Collection
    
    Set QueryList = New LinkedList
    
    FileNum = FreeFile
    
    On Error GoTo QDataError
    Open (App.Path & "\qkdata.gvd") For Input As #FileNum

    Do Until EOF(FileNum)
        
        Input #FileNum, Key, Title, SType, SInv
        
        Select Case SType
            Case "field":       KType = qtField
            Case "list":        KType = qtTraitList
            Case "num":         KType = qtNumber
            Case "date":        KType = qtDate
            Case "bool":        KType = qtBoolean
        End Select
        
        KInv = qiNone
        If InStr(SInv, "char") > 0 Then KInv = KInv + qiCharacters
        If InStr(SInv, "play") > 0 Then KInv = KInv + qiPlayers
        If InStr(SInv, "item") > 0 Then KInv = KInv + qiItems
        If InStr(SInv, "rote") > 0 Then KInv = KInv + qiRotes
        If InStr(SInv, "loc") > 0 Then KInv = KInv + qiLocations
        If InStr(SInv, "plot") > 0 Then KInv = KInv + qiPlots
        If InStr(SInv, "rumor") > 0 Then KInv = KInv + qiRumors
        If InStr(SInv, "act") > 0 Then KInv = KInv + qiActions
        
        TitlesToKeys.Add Key, Title
        KeysToTitles.Add Title, Key
        KeysToTypes.Add KType, Key
        KeysToInventories.Add KInv, Key
        
    Loop
    
    GoTo QDataClose

QDataError:
    
    MsgBox "Grapevine Loading Error: qkdata.gvd" & vbCrLf & Err.Description, vbCritical, "Error"
    
QDataClose:

    Close #FileNum
    On Error GoTo 0

End Sub

Public Sub AddQueryCopy(Q As QueryClass)
'
' Name:             AddQueryCopy
' Parameters:       Q           Query to copy and add
' Description:      Create a new query, copy the parameter into it, and add it to the list.
'

    Dim NewQuery As QueryClass
    
    Set NewQuery = New QueryClass
    
    With Q
        
        NewQuery.Name = .Name
        NewQuery.Inventory = .Inventory
        NewQuery.SortKey = .SortKey
        NewQuery.MatchAll = .MatchAll
        
        .First
        Do Until .Off
            With .Clause
                NewQuery.AddClause .Key, .Find, .Number, .Comparison, .CompNot
            End With
            .MoveNext
        Loop
        
    End With
    
    QueryList.InsertSorted NewQuery

End Sub

Public Sub MakeQuery(Q As QueryClass, Optional Sort As Boolean = False, Optional Invert = False)
'
' Name:         MakeQuery
' Parameters:   Q           QueryType that guides this query
'               Sort        Whether or not to use sorting on this query
'               Invert      Whether to invert the result of this query
' Description:  Populates the Results, Values, and possible SortList lists with results from
'               the query.
'

    Dim Inventory As LinkedList
    Dim UnmatchedSet As LinkedList
    
    Results.Clear
    Values.Clear
    SortList.Clear
    
    Select Case Q.Inventory
        Case qiCharacters
            Set Inventory = CharacterList
        Case qiPlayers
            Set Inventory = PlayerList
        Case qiItems
            Set Inventory = ItemList
        Case qiRotes
            Set Inventory = RoteList
        Case qiLocations
            Set Inventory = LocationList
'        Case qiActions
'        Case qiPlots
'        Case qiRumors
        Case Else
            Exit Sub
    End Select

    If Invert Then                              'Make this query its logical inverse
        Q.MatchAll = Not Q.MatchAll
        Q.First
        Do Until Q.Off
            Q.Clause.CompNot = Not Q.Clause.CompNot
            Q.MoveNext
        Loop
    End If

    Q.First
    
    If Not Q.Off Then                           'This Query has terms
    
        Set UnmatchedSet = New LinkedList
    
        Inventory.First                        'Copy the inventory to the unmatched set
        Do Until Inventory.Off
            UnmatchedSet.Append Inventory.Item
            Inventory.MoveNext
        Loop
       
        Call ProcessClause(Q.Clause, UnmatchedSet, Results, Values, False)
        Q.MoveNext
        
        Do Until Q.Off
        
            If Q.MatchAll Then
                Call ProcessClause(Q.Clause, Results, Nothing, Values, True)
            Else
                Call ProcessClause(Q.Clause, UnmatchedSet, Results, Values, False)
            End If
            Q.MoveNext
            
        Loop

        Set UnmatchedSet = Nothing

    ElseIf Not Invert Then                      'This Query has no terms:
                                                '  Match either everything or nothing.
        Inventory.First                         'Copy the inventory to the result set
        Do Until Inventory.Off
            Results.Append Inventory.Item
            Values.Append ""
            Inventory.MoveNext
        Loop
    
    End If
    
    If Sort Then
        '
        ' Build Sort list, and bubble sort
        '
        Dim CharData As Variant
        
        Results.First
        Values.First
        Do Until Results.Off
            If Q.SortKey = "" Then
                SortList.Append Values.Item
            Else
                Call Results.Item.GetValue(Q.SortKey, CharData)
                If IsNull(CharData) Then
                    CharData = "N/A"
                ElseIf IsObject(CharData) Then
                    CharData = CharData.Count
                End If
                SortList.Append CharData
            End If
            Results.MoveNext
            Values.MoveNext
        Loop
                
        SortList.SyncSort "N/A", Q.SortDescend, Results, Values
        
    End If

    If Invert Then                              'Switch the query back to normal, if needed
        Q.MatchAll = Not Q.MatchAll
        Q.First
        Do Until Q.Off
            Q.Clause.CompNot = Not Q.Clause.CompNot
            Q.MoveNext
        Loop
    End If
                                
End Sub

Private Sub ProcessClause(QC As QueryClauseClass, SetA As LinkedList, SetB As LinkedList, _
                          ValueSet As LinkedList, MatchAll As Boolean)
' Name:         ProcessClause
' Parameters:   QC              Query clause for which to find results
'               SetA            a list to search for matches
'               SetB            a list that MIGHT be used to hold matches (see MatchAll)
'               ValueSet        the list that describes the reason for matches
'               MatchAll        If TRUE, destroy non-matches in SetA and ignore SetB.
'                               If FALSE, move matches from SetA to SetB.
' Description:  Process one clause of a query, returning data that can be used to create the whole
'               of the query.  Used by the Query method.
'

    Dim QType As QueryKeyType
    Dim Applicable As Boolean
    Dim Match As Boolean
    Dim CharData As Variant
    Dim MatchValue As String
    Dim TraitList As LinkedTraitList
    
    On Error Resume Next
    QType = qtError
    QType = KeysToTypes.Item(QC.Key)
    On Error GoTo 0
    
    If QType = qtError Then Exit Sub
    
    SetA.First
    ValueSet.First
    
    Do Until SetA.Off
        
        Applicable = True
        Match = False
        Call SetA.Item.GetValue(QC.Key, CharData)
        MatchValue = ""
        
        If Not IsNull(CharData) Then
            
            Select Case QType
                Case qtField
                
                    MatchValue = CStr(CharData)
                    Select Case QC.Comparison
                        Case qcContains
                            Match = (InStr(CharData, QC.Find) > 0)
                        Case qcEquals
                            Match = (CharData = QC.Find)
                        Case Else
                            Applicable = False
                    End Select
                    
                Case qtNumber
                    
                    MatchValue = CStr(CharData)
                    Select Case QC.Comparison
                        Case qcEquals
                            Match = (CharData = QC.Number)
                        Case qcAtLeast
                            Match = (CharData >= QC.Number)
                        Case qcGreater
                            Match = (CharData > QC.Number)
                        Case qcLess
                            Match = (CharData < QC.Number)
                        Case qcNoMore
                            Match = (CharData <= QC.Number)
                        Case Else
                            Applicable = False
                    End Select
                    
                Case qtTraitList
                    
                    Set TraitList = CharData
                    MatchValue = CStr(TraitList.Count)
                    
                    Select Case QC.Comparison
                        Case qcTotals
                            Match = (TraitList.Count = QC.Number)
                        Case qcTotalsAtLeast
                            Match = (TraitList.Count >= QC.Number)
                        Case qcTotalsMore
                            Match = (TraitList.Count > QC.Number)
                        Case qcTotalsNoMore
                            Match = (TraitList.Count <= QC.Number)
                        Case qcTotalsLess
                            Match = (TraitList.Count < QC.Number)
                        Case qcContainsNote
                            TraitList.First
                            TraitList.MoveToNotePart CStr(QC.Find)
                            Match = Not TraitList.Off
                            If Match Then
                                MatchValue = TraitList.DisplayTrait
                            Else
                                MatchValue = "(none)"
                            End If
                        Case qcContains
                            TraitList.First
                            TraitList.MoveToPart CStr(QC.Find)
                            Match = Not TraitList.Off
                            If Match Then
                                MatchValue = TraitList.DisplayTrait
                            Else
                                MatchValue = "(none)"
                            End If
                        Case Else
                        
                            Dim FromTheTop As Boolean
                        
                            MatchValue = "(none)"
                            TraitList.First
                            FromTheTop = True
                            
                            Do
                                
                                If Not FromTheTop Then TraitList.MoveNext
                                TraitList.MoveToPart CStr(QC.Find)
                                FromTheTop = False
                                
                                If Not TraitList.Off Then
                                   
                                    MatchValue = TraitList.DisplayTrait
                                    Select Case QC.Comparison
                                        Case qcContainsNoMore
                                            Match = (TraitList.Trait.Number <= QC.Number)
                                        Case qcContainsLess
                                            Match = (TraitList.Trait.Number < QC.Number)
                                        Case qcContainsExactly
                                            Match = (TraitList.Trait.Number = QC.Number)
                                        Case qcContainsAtLeast
                                            Match = (TraitList.Trait.Number >= QC.Number)
                                        Case qcContainsMore
                                            Match = (TraitList.Trait.Number > QC.Number)
                                        Case Else
                                            Applicable = False
                                    End Select
                                    
                                End If
                                                                   
                            Loop Until Match Or TraitList.Atomic = False Or TraitList.Off
                                    
                    End Select
                    
                Case qtDate
                    
                    If IsDate(QC.Find) Then
                        MatchValue = Format(CharData, "Short Date")
                        Select Case QC.Comparison
                            Case qcEquals
                                Match = (CharData = CDate(QC.Find))
                            Case qcAtLeast
                                Match = (CharData >= CDate(QC.Find))
                            Case qcGreater
                                Match = (CharData > CDate(QC.Find))
                            Case qcLess
                                Match = (CharData < CDate(QC.Find))
                            Case qcNoMore
                                Match = (CharData <= CDate(QC.Find))
                            Case Else
                                Applicable = False
                        End Select
                    Else
                        Applicable = False
                    End If
                    
                Case qtBoolean
                    
                    MatchValue = CStr(CharData)
                    Select Case QC.Comparison
                        Case qcIsTrue
                            Match = CBool(CharData)
                        Case qcIsFalse
                            Match = Not CBool(CharData)
                        Case Else
                            Applicable = False
                    End Select

                Case Else
                    Applicable = False
            End Select
        
        Else
        
            Applicable = False
            MatchValue = "N/A"
        
        End If
        
        Match = Applicable And (Match Xor QC.CompNot)
    
        If Match Then
        
            If Not MatchAll Then                'move matches from SetA to SetB
                SetB.Append SetA.Item
                ValueSet.Append MatchValue          'create ValueSet
                SetA.Remove
            Else
                SetA.MoveNext                       'adjust ValueSet
                MatchValue = ValueSet.Item & ", " & MatchValue
                ValueSet.Remove
                ValueSet.Insert MatchValue
            End If
            
        Else
        
            If MatchAll Then                    'remove non-matches from SetA
                SetA.Remove
                ValueSet.Remove                     'prune ValueSet
            Else
                SetA.MoveNext
            End If
        
        End If
        
    Loop

End Sub

Public Sub SelectQueryResults(QueryName As String, ResultBox As ListBox, Invert As Boolean)
'
' Name:         SelectQueryResults
' Parameters:   QueryName       Name of the query to use
'               ResultBox       ListBox holding items to select
'               Invert          whether to select the results or the exceptions
' Description:  Perform a query and select items in a list that correspond to the results.
'

    QueryList.MoveTo QueryName
    If Not QueryList.Off Then
    
        Dim NameSet As StringSet
        Dim I As Integer
        
        Set NameSet = New StringSet
        
        MakeQuery QueryList.Item
        
        Results.First
        Do Until Results.Off
            NameSet.Add Results.Item.Name
            Results.MoveNext
        Loop
        
        NameSet.SelectListBox ResultBox, Not Invert, Invert
        
        Set NameSet = Nothing
        
    End If

End Sub

Public Sub AddDefaultQueries()
'
' Name:         DefaultQueries
' Description:  Add the default queries to the list.
'

    Dim Q As QueryClass
    
    Set Q = New QueryClass
    
    With Q

        .Inventory = qiCharacters
        .SortKey = qkName
        .MatchAll = False
        
            .Name = "Various"
            QueryList.MoveTo .Name
            If QueryList.Off Then
                .AddClause qkRace, "various", 0, qcEquals, False
                Call AddQueryCopy(Q)
            End If
            
            .RemoveLast
            .Name = "Mummies"
            QueryList.MoveTo .Name
            If QueryList.Off Then
                .AddClause qkRace, "mummy", 0, qcEquals, False
                Call AddQueryCopy(Q)
            End If
                        
            .RemoveLast
            .Name = "Fera"
            QueryList.MoveTo .Name
            If QueryList.Off Then
                .AddClause qkRace, "fera", 0, qcEquals, False
                Call AddQueryCopy(Q)
            End If
            
            .RemoveLast
            .Name = "Kuei-Jin"
            QueryList.MoveTo .Name
            If QueryList.Off Then
                .AddClause qkRace, "kuei-jin", 0, qcEquals, False
                Call AddQueryCopy(Q)
            End If
            
            .RemoveLast
            .Name = "Hunters"
            QueryList.MoveTo .Name
            If QueryList.Off Then
                .AddClause qkRace, "hunter", 0, qcEquals, False
                Call AddQueryCopy(Q)
            End If
            
            .RemoveLast
            .Name = "Demons"
            QueryList.MoveTo .Name
            If QueryList.Off Then
                .AddClause qkRace, "demon", 0, qcEquals, False
                Call AddQueryCopy(Q)
            End If
            
            .RemoveLast
            .Name = "Mortals"
            QueryList.MoveTo .Name
            If QueryList.Off Then
                .AddClause qkRace, "mortal", 0, qcEquals, False
                Call AddQueryCopy(Q)
            End If
            
            .RemoveLast
            .Name = "Changelings"
            QueryList.MoveTo .Name
            If QueryList.Off Then
                .AddClause qkRace, "changeling", 0, qcEquals, False
                Call AddQueryCopy(Q)
            End If
            
            .RemoveLast
            .Name = "Wraiths"
            QueryList.MoveTo .Name
            If QueryList.Off Then
                .AddClause qkRace, "wraith", 0, qcEquals, False
                Call AddQueryCopy(Q)
            End If
            
            .RemoveLast
            .Name = "Mages"
            QueryList.MoveTo .Name
            If QueryList.Off Then
                .AddClause qkRace, "mage", 0, qcEquals, False
                Call AddQueryCopy(Q)
            End If
            
            .RemoveLast
            .Name = "Werewolves"
            QueryList.MoveTo .Name
            If QueryList.Off Then
                .AddClause qkRace, "werewolf", 0, qcEquals, False
                Call AddQueryCopy(Q)
            End If
            
            .RemoveLast
            .Name = "Vampires"
            QueryList.MoveTo .Name
            If QueryList.Off Then
                .AddClause qkRace, "vampire", 0, qcEquals, False
                Call AddQueryCopy(Q)
            End If
            
            .RemoveLast
            .Name = "NPCs"
            QueryList.MoveTo .Name
            If QueryList.Off Then
                .AddClause qkNPC, "", 0, qcIsTrue, False
                Call AddQueryCopy(Q)
            End If
            
            .RemoveLast
            .Name = "Active Characters"
            QueryList.MoveTo .Name
            If QueryList.Off Then
                .AddClause qkPlayStatus, "active", 0, qcEquals, False
                Call AddQueryCopy(Q)
            End If
            
            .RemoveLast
            .Name = "All Characters"
            QueryList.MoveTo .Name
                If QueryList.Off Then
                Call AddQueryCopy(Q)
            End If
            
        .Inventory = qiPlayers
        
            .Name = "All Players"
            QueryList.MoveTo .Name
                If QueryList.Off Then
                Call AddQueryCopy(Q)
            End If
            
            .Name = "Staff"
            QueryList.MoveTo .Name
            If QueryList.Off Then
                .AddClause qkPosition, "player", 0, qcEquals, True
                Call AddQueryCopy(Q)
            End If
            
            .RemoveLast
            .Name = "Active Players"
            QueryList.MoveTo .Name
            If QueryList.Off Then
                .AddClause qkPlayStatus, ActiveStatus, 0, qcEquals, False
                Call AddQueryCopy(Q)
            Else
                If QueryList.Item.LastModified < #1/21/2003# Then
                    QueryList.Item.Clear
                    QueryList.Item.Name = "Active Players"
                    QueryList.Item.Inventory = qiPlayers
                    QueryList.Item.AddClause qkPlayStatus, ActiveStatus, 0, qcEquals, False
                End If
            End If
            
    End With
    
    Set Q = Nothing

End Sub

Public Sub GetStatistics(Stat As StatisticType, Query As QueryClass, Key As String, _
        Optional OKZero As Boolean = True, Optional Trait As String)
'
' Name:         GetStatistics
' Parameters:   Stat        Type of statistic being requested
'               Query       Query describing the characters to be examined
'               Key         The key of the character data to examine
'               OKZero      Whether to include results of 0 and "" in the counts.
'               Trait       Required for Specific Distribution type -- trait to look for.
' Description:  Run a statistic on the key value for characters returned by Query.
'               Store the values found as strings in StatResults; store the corresponding
'               numbers in NumberSet; store the individual names and numbers found in
'               traitlists in MatchSet.
'
'

    Dim CharData As Variant
    Dim DataList As LinkedTraitList
    Dim ResKey As Variant
    Dim I As Double
    Dim NewNum As Double
    Dim MatchList As LinkedTraitList
    
    Total = 0
    Maximum = 0
    
    StatResults.Clear
    Do Until NumberSet.Count = 0
        NumberSet.Remove 1
    Loop
    Do Until MatchSet.Count = 0
        MatchSet.Remove 1
    Loop
    
    MakeQuery Query                             'Get the set of data to examine
    
    Results.First                               'for each character in the set
    Do Until Results.Off
    
        Call Results.Item.GetValue(Key, CharData)      'Examine the data in its key
        
        If Not IsNull(CharData) Then                    'N/A keys don't count
        
            If IsObject(CharData) Then                  'dealing with a trait list
        
                Set DataList = CharData
        
                Select Case Stat
                    Case stDistribution
                        ResKey = DataList.Count
                    Case stSpecificDistribution
                        DataList.First
                        DataList.MoveToPart Trait
                        If DataList.Off Then
                            ResKey = 0
                        Else
                            ResKey = DataList.Trait.Number
                        End If
                    Case stDistinctDistribution, stMaxima, stSums
                    
                        DataList.First                              'Loop through trait list data
                        Do Until DataList.Off
                            ResKey = DataList.Trait.Name
                            NewNum = DataList.Trait.Number
                            If Stat = stDistinctDistribution Then NewNum = 1
                            Total = Total + NewNum
                            StatResults.MoveTo ResKey
                            If StatResults.Off Then                  'First such entry
                                StatResults.InsertSorted ResKey
                                NumberSet.Add NewNum, CStr(ResKey)
                                MatchSet.Add New LinkedTraitList, CStr(ResKey)
                            Else                                    'Compare such entries
                                I = NumberSet(CStr(ResKey))
                                If Stat = stMaxima Then                 'Replace with higher value
                                    If NewNum > I Then
                                        NumberSet.Remove CStr(ResKey)
                                        NumberSet.Add NewNum, CStr(ResKey)
                                    End If                              'Add to current value
                                Else
                                    NumberSet.Remove CStr(ResKey)
                                    NumberSet.Add I + NewNum, CStr(ResKey)
                                End If
                            End If
                            Set MatchList = MatchSet(CStr(ResKey))
                            MatchList.Insert Results.Item.Name, DataList.Trait.Total
                            DataList.MoveNext
                        Loop
                    
                End Select
        
            Else                                        'not dealing with trait lists
            
                Select Case Stat
                    Case stDistribution
                        ResKey = CharData
                        If IsDate(CharData) Then ResKey = Format(CharData, "Short Date")
                        If ResKey = "" Then ResKey = "(none)"
                    Case stMaxima, stSums
                        ResKey = KeysToTitles(Key)          'Just one possible datum: ResKey
                        NewNum = CDbl(CharData)
                        Total = Total + NewNum
                        If StatResults.IsEmpty Then          'First entry
                            StatResults.Append ResKey
                            NumberSet.Add NewNum, CStr(ResKey)
                            MatchSet.Add New LinkedTraitList, CStr(ResKey)
                        Else                                'Compare entries
                            StatResults.First
                            I = NumberSet(ResKey)
                            If Stat = stMaxima Then                 'Replace with higher value
                                If NewNum > I Then
                                    NumberSet.Remove CStr(ResKey)
                                    NumberSet.Add NewNum, CStr(ResKey)
                                End If                              'Add to current value
                            Else
                                NumberSet.Remove CStr(ResKey)
                                NumberSet.Add I + NewNum, CStr(ResKey)
                            End If
                        End If
                        Set MatchList = MatchSet(CStr(ResKey))
                        MatchList.Insert Results.Item.Name, CStr(NewNum)
                End Select
            
            End If
        
            Select Case Stat
                Case stDistribution, stSpecificDistribution     'Distributions: insert into results
                    If OKZero Or Not (ResKey = 0 Or ResKey = "(none)") Then
                        Total = Total + 1
                        StatResults.MoveTo ResKey
                        If StatResults.Off Then                          'first such entry
                            StatResults.InsertSorted ResKey
                            NumberSet.Add 1, CStr(ResKey)
                            Set MatchList = New LinkedTraitList
                            MatchList.Display = ldSimple
                            MatchSet.Add MatchList, CStr(ResKey)
                        Else                                            'add to previous entry
                            I = NumberSet(CStr(ResKey))
                            NumberSet.Remove CStr(ResKey)
                            NumberSet.Add I + 1, CStr(ResKey)
                        End If
                        Set MatchList = MatchSet(CStr(ResKey))
                        MatchList.Insert Results.Item.Name
                    End If
                Case stDistinctDistribution
                    Total = Results.Count                       'total characters examined
            End Select
        
        End If

        Results.MoveNext
    Loop
    
    StatResults.First
    
    Select Case Stat
        Case stSpecificDistribution                'Reformat the results for specific distribution
            Do Until StatResults.Off
                ResKey = StatResults.Item
                StatResults.Remove
                StatResults.Insert Trait & " x" & CStr(ResKey)
                NumberSet.Add NumberSet(CStr(ResKey)), Trait & " x" & CStr(ResKey)
                NumberSet.Remove CStr(ResKey)
                MatchSet.Add MatchSet(CStr(ResKey)), Trait & " x" & CStr(ResKey)
                MatchSet.Remove CStr(ResKey)
            Loop
        Case stDistribution
            If Not (KeysToTypes(Key) = qtField) Then
                Do Until StatResults.Off                         'Reformat for distribution of numbers
                    ResKey = StatResults.Item
                    StatResults.Remove
                    StatResults.Insert CStr(ResKey) & " " & KeysToTitles(Key)
                    NumberSet.Add NumberSet(CStr(ResKey)), CStr(ResKey) & " " & KeysToTitles(Key)
                    NumberSet.Remove CStr(ResKey)
                    MatchSet.Add MatchSet(CStr(ResKey)), CStr(ResKey) & " " & KeysToTitles(Key)
                    MatchSet.Remove CStr(ResKey)
                Loop
            End If
        Case stMaxima
            Total = 0
            For I = 1 To NumberSet.Count
                Total = Total + NumberSet(I)
            Next I
    End Select
    
    For I = 1 To NumberSet.Count
        If Maximum < NumberSet(I) Then Maximum = NumberSet(I)
    Next I
    
End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Clean up after the class.
'
    Set TitlesToKeys = Nothing
    Set KeysToTitles = Nothing
    Set KeysToTypes = Nothing
    Set KeysToInventories = Nothing
    Set Results = Nothing
    Set Values = Nothing
    Set SortList = Nothing
    Set StatResults = Nothing
    Set NumberSet = Nothing
    Set MatchSet = Nothing
    Set QueryList = Nothing
            
End Sub
