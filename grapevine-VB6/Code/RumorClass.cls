VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RumorClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Title As String
Public RumorDate As Date

Public Category As RumorCategoryType
Public Query As QueryClass
Public MultiKey As String
Public MultiMatch As String

Public Done As Boolean
Public LastModified As Date

Public SubRumor As RumorNode

Private FirstNode As RumorNode
Private LastNode As RumorNode
Private NodeCount As Integer

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Create the list and default Rumor node.
'

    Set FirstNode = Nothing
    Set LastNode = Nothing
    Set SubRumor = Nothing
    NodeCount = 0
    LastModified = Now
    
End Sub

Public Sub InitializeQueryRumor(What As String, When As Date, RType As RumorCategoryType)
'
' Name:         InitializeQueryRumor
' Parameters:   What        Title of the new rumor
'               When        Date of the new rumor
'               RType       Category of the rumor
' Description:  Initialize a simple query-based rumor.
'
    Add 0, ""
    Set Query = New QueryClass
    Query.Inventory = qiCharacters
    Title = What
    RumorDate = When
    Category = RType
    
End Sub

Public Sub InitializeMultiRumor(What As String, When As Date, NewKey As String, NewMatch As String)
'
' Name:         InitializeMultiRumor
' Parameters:   What        Title of the new rumor
'               When        Date of the new rumor
'               NewKey      The list associated with this multirumor
'               NewMatch    The trait associated with this multirumor
' Description:  Initilize a multi-level rumor.
'

    Dim I As Integer
    
    Title = What
    RumorDate = When
    MultiKey = NewKey
    MultiMatch = NewMatch
    Category = rtInfluence
    
    For I = 1 To 10
        Add I, ""
    Next I

End Sub

Public Function OutputID() As Integer
'
' Name:         OutputID
' Description:  Return an ID that tells the OutputEngineClass how to manipulate this object
'
    OutputID = oidRumor

End Function

Public Sub InitializeForOutput()
'
' Name:         InitializeForOutput
' Description:  Position the object in a starting state, ready for output.
'

    First

End Sub

Public Sub GetValue(Key As String, ObjData As Variant)
'
' Name:         GetValue
' Parameters:   Key         a key (string) to the value to retreive
'               ObjData     the returned value
' Description:  Use the given key to return data from the object.  Keys are defined
'               in the PublicQueryKeys module.
'

    ObjData = ""
    Select Case Key
        Case qkTitle:           ObjData = Title
        Case qkDate:            ObjData = RumorDate
        Case qkCount:           ObjData = Count
        Case qkRumor:           If Not Off Then ObjData = SubRumor.Rumor
        Case qkLevel:           If Category = rtInfluence And Not Off Then ObjData = SubRumor.Level
        Case qkType
            If Category = rtInfluence Then
                If Not Off Then ObjData = "Level " & CStr(SubRumor.Level) & " " & Title
            Else
                ObjData = Title
            End If
        Case Else:              ObjData = Null
    End Select

End Sub

Public Function IconKey() As String
'
' Name:         IconKey
' Description:  Return a string key for the icon corresponding to this rumor's category.
'
    
    Select Case Category
        Case rtInfluence:   IconKey = "InfluenceRumor"
        Case rtGroup:       IconKey = "GroupRumor"
        Case rtSubgroup:    IconKey = "SubgroupRumor"
        Case rtRace:        IconKey = "RaceRumor"
        Case rtPersonal:    IconKey = "PersonalRumor"
        Case Else:          IconKey = "Rumor"
    End Select

End Function

Public Function Clone(When As Date, IncludeText As Boolean) As RumorClass
'
' Name:         Clone
' Parameters:   When        the date of the cloned rumor
'               IncludeText whether or not to clone the text from this rumor
' Description:  Return a copy of this Rumor object, with a new date,
'               and possibly without text in its rumors.  Do not copy the
'               causes of the rumor's subrumors.
'

    Set Clone = New RumorClass
    
    Clone.RumorDate = When
    Clone.Title = Title
    Clone.MultiKey = MultiKey
    Clone.MultiMatch = MultiMatch
    Clone.LastModified = Now
    Clone.Category = Category
    
    First
    Do Until Off
        Clone.Add SubRumor.Level, IIf(IncludeText, SubRumor.Rumor, "")
        MoveNext
    Loop

    If Not (Query Is Nothing) Then
        
        Set Clone.Query = New QueryClass
        With Query
            Clone.Query.Inventory = .Inventory
            Clone.Query.MatchAll = .MatchAll
            .First
            Do Until .Off
                With .Clause
                    Clone.Query.AddClause .Key, .Find, .Number, .Comparison, .CompNot
                End With
                .MoveNext
            Loop
        End With
        
    End If
    
End Function

Public Sub IfDoneSetDone(ByVal DoneLevel As Integer, ThisLevel As Integer)
'
' Name:         IfDoneSetDone
' Parameter:    DoneLevel       number of rumors that must be finished for
'                               this to mark itself 'done'
'               ThisLevel       the rumor level just altered
' Description:  If this isn't an influence rumor, Done is ser true iff the subrumor isn't blank.
'               If this is an influence rumor and ThisLevel <= DoneLevel, Done is set true iff
'                   the first (DoneLevel) subrumors are finished.
'               If this is an influence rumor and ThisLevel > DoneLevel, Done is set true iff
'                   all the subrumors are finished.
'

    Done = True
    First
    If Not Category = rtInfluence Then
        Done = SubRumor.Rumor <> ""
    ElseIf ThisLevel <= DoneLevel Then
        Do Until DoneLevel = 0 Or SubRumor Is Nothing
            Done = Done And (SubRumor.Rumor <> "")
            DoneLevel = DoneLevel - 1
            MoveNext
        Loop
    Else
        Do Until SubRumor Is Nothing
            Done = Done And (SubRumor.Rumor <> "")
            MoveNext
        Loop
    End If
    
End Sub

Public Sub PrepareRecipients()
'
' Name:         PrepareRecipients
' Description:  Populate the recipient set of each subrumor with the names of those who
'               will recieve this rumor.
'

    Dim Q As QueryClass
    
    If Category = rtInfluence Then
        Set Q = New QueryClass
        Q.Inventory = qiCharacters
        Q.AddClause MultiKey, MultiMatch, 0, qcContains, False
    Else
        Set Q = Query
    End If
    
    If Not (Q Is Nothing) Then
        
        Game.QueryEngine.MakeQuery Q
        
        First
        Do Until Off
            SubRumor.RecipientSet.Clear
            MoveNext
        Loop
        
        If Category = rtInfluence Then
            
            Dim InfList As LinkedTraitList
            Dim CharData As Variant
            Dim I As Integer
            
            With Game.QueryEngine
                With Game.QueryEngine.Results
                    .First
                    Do Until .Off
                        .Item.GetValue MultiKey, CharData
                        If IsObject(CharData) Then
                            Set InfList = CharData
                            InfList.MoveTo MultiMatch
                            If Not InfList.Off Then
                                First
                                I = InfList.Trait.Number
                                Do Until I <= 0 Or Off
                                    SubRumor.RecipientSet.Add .Item.Name
                                    MoveNext
                                    I = I - 1
                                Loop
                            End If
                        End If
                        .MoveNext
                    Loop
                End With
            End With
            
        Else
            First
            With Game.QueryEngine.Results
                .First
                Do Until .Off
                    SubRumor.RecipientSet.Add .Item.Name
                    .MoveNext
                Loop
            End With
        End If
        
    End If

    Set Q = Nothing

End Sub

Public Function Name() As String
'
' Name:         Name
' Description:  Return a name for this rumor, combining the date and the title.
'

    Name = Format(RumorDate, "Short Date") & " " & Title

End Function

Public Sub Add(Level As Integer, Rumor As String)
'
' Name:         Add
' Parameters:   Level       Level of the Rumor development to add
'               Rumor       text of the Rumor development to add
' Description:  Add a new subrumor.
'

    Dim NewRumor As RumorNode

    Set NewRumor = New RumorNode

    NewRumor.Level = Level
    NewRumor.Rumor = Rumor

    AddNode NewRumor

End Sub

Private Sub AddNode(NewRumor As RumorNode)
'
' Name:         AddNode
' Parameters:   NewRumor    the new rumor node to insert
' Description:  Add a new rumor directly from a created node.
'

    Set NewRumor.Parent = Me
    
    If FirstNode Is Nothing Then
    
        Set FirstNode = NewRumor
        Set LastNode = NewRumor
        Set SubRumor = NewRumor
    
    Else
    
        Set SubRumor = FirstNode
        Do
            If NewRumor.Level < SubRumor.Level Then Exit Do
            Set SubRumor = SubRumor.NextNode
        Loop Until SubRumor Is Nothing
        
        If SubRumor Is Nothing Then
            Set NewRumor.PrevNode = LastNode
            Set LastNode.NextNode = NewRumor
            Set LastNode = NewRumor
        Else
            Set NewRumor.NextNode = SubRumor.NextNode
            Set NewRumor.PrevNode = SubRumor.PrevNode
            If SubRumor Is FirstNode Then
                Set FirstNode = NewRumor
            Else
                Set SubRumor.PrevNode.NextNode = NewRumor
            End If
            Set SubRumor.PrevNode = NewRumor
        End If
        Set SubRumor = NewRumor
       
    End If
    
    NodeCount = NodeCount + 1

End Sub

Public Sub First()
'
' Name:         First
' Description:  Set SubRumor to the first Rumor in the nodelist.
'

    Set SubRumor = FirstNode

End Sub

Public Sub MoveNext()
'
' Name:         MoveNext
' Description:  Move to the next Rumor in the list and set SubRumor to it.
'

    If Not SubRumor Is Nothing Then Set SubRumor = SubRumor.NextNode
    
End Sub

Public Function Off() As Boolean
'
' Name:         Off
' Description:  Return TRUE iff there is no current SubRumor.
'
    Off = (SubRumor Is Nothing)

End Function

Public Sub Remove()
'
' Name:         Remove
' Description:  Remove the current SubRumor.
'
    
    If Not SubRumor Is Nothing Then
        
        Dim TempNode As RumorNode
        
        If SubRumor Is FirstNode Then
            Set FirstNode = SubRumor.NextNode
        Else
            Set SubRumor.PrevNode.NextNode = SubRumor.NextNode
        End If
        
        If SubRumor Is LastNode Then
            Set LastNode = SubRumor.PrevNode
        Else
            Set SubRumor.NextNode.PrevNode = SubRumor.PrevNode
        End If
        
        Set TempNode = SubRumor.NextNode
        Set SubRumor = Nothing
        Set SubRumor = TempNode
    
        NodeCount = NodeCount - 1
    
    End If
    
End Sub

Public Sub MoveTo(Level As Integer)
'
' Name:         MoveTo
' Parameters:   Level           Level of the rumor to move to.
' Description:  Move to the given Rumor level.
'

    Set SubRumor = FirstNode
    Do Until SubRumor Is Nothing
        If SubRumor.Level = Level Then Exit Do
        Set SubRumor = SubRumor.NextNode
    Loop

End Sub

Public Sub MoveToPlace(ByVal Place As Integer)
'
' Name:         MoveToPlace
' Parameters:   Place           the position in the list to move to
' Description:  Move the Link to the given number position in the list.
'               Like listbox listindexes, this is 0-based.
'
    
    Set SubRumor = FirstNode
    Do Until (Place <= 0 Or SubRumor Is Nothing)
        Set SubRumor = SubRumor.NextNode
        Place = Place - 1
    Loop

End Sub

Public Function Count() As Integer
'
' Name:         Count
' Description:  Return the number of SubRumors in this Rumor.
'
    Count = NodeCount

End Function

Public Sub Clear()
'
' Name:         Clear
' Description:  Empty this Rumor of SubRumors.
'

    First
    Do Until SubRumor Is Nothing
        Remove
    Loop

End Sub

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write the object to an XML file.
'

    Dim MemCursor As RumorNode

    Set MemCursor = SubRumor

    With XML
    
        .BeginTag "rumor"
    
            .WriteAttribute "title", Title
            .WriteAttribute "date", RumorDate
            .WriteAttribute "category", Category, 0
            .WriteAttribute "list", MultiKey
            .WriteAttribute "trait", MultiMatch
            .WriteAttribute "done", Done, False
            
            .WriteAttribute "lastmodified", LastModified
            
            If Not (Query Is Nothing) Then Query.OutputToFile XML
            
            First
            Do Until Off
                SubRumor.OutputToFile XML
                MoveNext
            Loop
    
        .EndTag
        
    End With

    Set SubRumor = MemCursor

End Sub

Public Sub InputFromFile(XML As XMLReaderClass, Version As Double)
'
' Name:         InputFromFile
' Parameters:   XML         an XMLReaderClass that just read the opening
'                           tag of this object
'               Version     version of the file format
' Description:  Read the object from an XML file.
'

    Dim NewNode As RumorNode

    If XML.Tag = "rumor" Then
        
        With XML
        
            Clear
        
            Title = .GetString("title")
            RumorDate = .GetDate("date")
            Category = .GetLong("category")
            MultiKey = .GetString("list")
            MultiMatch = .GetString("trait")
            Done = .GetBoolean("done")
            
            LastModified = .GetDate("lastmodified")
            
            Do Until .Tag = "rumor" And .Terminal
                .ReadTag
                Select Case .Tag
                    Case "query"
                        Set Query = Nothing
                        Set Query = New QueryClass
                        Query.InputFromFile XML, Version
                    Case "subrumor"
                        Set NewNode = New RumorNode
                        NewNode.InputFromFile XML, Version
                        AddNode NewNode
                End Select
            Loop
            
        End With
        
    Else
        Err.Raise vbObjectError + 8, "RumorClass", _
                  "Wrong file input to a rumor"
    End If

End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this object to a binary file.
'

    Dim MemCursor As RumorNode

    Set MemCursor = SubRumor

    PutStrB FileNum, Title
    Put #FileNum, , RumorDate
    Put #FileNum, , Category
    PutStrB FileNum, MultiKey
    PutStrB FileNum, MultiMatch
    Put #FileNum, , Done
    Put #FileNum, , LastModified
    
    If MultiKey = "" Then
        Query.OutputToBinary FileNum
    End If
    
    Put #FileNum, , NodeCount
    First
    Do Until Off
        SubRumor.OutputToBinary FileNum
        MoveNext
    Loop

    Set SubRumor = MemCursor

End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this object from a binary file.
'

    Dim I As Integer
    Dim NewNode As RumorNode
    
    Clear

    GetStrB FileNum, Title
    Get #FileNum, , RumorDate
    Get #FileNum, , Category
    GetStrB FileNum, MultiKey
    GetStrB FileNum, MultiMatch
    Get #FileNum, , Done
    Get #FileNum, , LastModified
    
    If MultiKey = "" Then
        Set Query = Nothing
        Set Query = New QueryClass
        Query.InputFromBinary FileNum, Version
    End If
        
    Get #FileNum, , I
    Do Until I = 0
        Set NewNode = New RumorNode
        NewNode.InputFromBinary FileNum, Version
        AddNode NewNode
        I = I - 1
    Loop
 
End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Clean up.
'
    
    Clear

End Sub


