VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QueryClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String                           'Name of this query.
Public Inventory As QueryInventoryType          'The inventory to check against.
Public MatchAll As Boolean                      'Match all the parts of the query.  Alternative: Match any.
Public SortKey As String                        'Field by which to sort ("" to sort by deciding query term)
Public SortDescend As Boolean                   'True, sort descending; false, sort ascending
Public LastModified As Date                     'Query creation date

Private Clauses As LinkedList                   'Linkedlist of query clauses

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write this object to a file.
'

    Dim InvStr As String

    Select Case Inventory
        Case qicharacters:  InvStr = "chars"
        Case qiPlayers:     InvStr = "players"
        Case qiItems:       InvStr = "items"
        Case qiRotes:       InvStr = "rotes"
        Case qiLocations:   InvStr = "locs"
        Case qiActions:     InvStr = "actions"
        Case qiPlots:       InvStr = "plots"
        Case qiRumors:      InvStr = "rumors"
    End Select

    Clauses.Memorize

    With XML
    
        .BeginTag "query"
            .WriteAttribute "name", Name
            .WriteAttribute "table", InvStr
            .WriteAttribute "match", IIf(MatchAll, "all", "any")
            .WriteAttribute "sort", SortKey
            .WriteAttribute "descend", SortDescend, False
            .WriteAttribute "lastmodified", LastModified
            First
            Do Until Off
                Clause.OutputToFile XML
                MoveNext
            Loop
        .EndTag
        
    End With

    Clauses.Remember

End Sub

Public Sub InputFromFile(XML As XMLReaderClass, Version As Double)
'
' Name:         InputFromFile
' Parameters:   XML             XMLReaderClass that has read the tag
'               Version         file format version
' Description:  Reads the object in from an XML file.
'

    Dim InvStr As String
    Dim QClause As QueryClauseClass
    
    With XML
    
        If .Tag = "query" Then
        
            Clauses.Clear
        
            Name = .GetString("name")
            InvStr = .GetString("table")
            MatchAll = (.GetString("match") = "all")
            SortKey = .GetString("sort")
            SortDescend = .GetBoolean("descend")
            LastModified = .GetDate("lastmodified", Now)
            
            Select Case Left(InvStr, 3)
                Case "cha":     Inventory = qicharacters
                Case "pla":     Inventory = qiPlayers
                Case "ite":     Inventory = qiItems
                Case "rot":     Inventory = qiRotes
                Case "loc":     Inventory = qiLocations
                Case "act":     Inventory = qiActions
                Case "plo":     Inventory = qiPlots
                Case "rum":     Inventory = qiRumors
            End Select
            
            Do Until .Tag = "query" And .Terminal
                .ReadTag
                If .Tag = "clause" Then
                    Set QClause = New QueryClauseClass
                    QClause.InputFromFile XML, Version
                    Clauses.Append QClause
                End If
            Loop
        
        End If
    
    End With
    
End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this character's data to a binary file.
'

    Clauses.Memorize

    PutStrB FileNum, Name
    Put #FileNum, , Inventory
    Put #FileNum, , MatchAll
    PutStrB FileNum, SortKey
    Put #FileNum, , SortDescend
    Put #FileNum, , LastModified
    Put #FileNum, , CInt(ClauseCount)
    First
    Do Until Off
        Clause.OutputToBinary FileNum
        MoveNext
    Loop

    Clauses.Remember

End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this character's data from a binary file.
'

    Dim I As Integer
    Dim QClause As QueryClauseClass
        
    Clear
    
    GetStrB FileNum, Name
    Get #FileNum, , Inventory
    Get #FileNum, , MatchAll
    GetStrB FileNum, SortKey
    Get #FileNum, , SortDescend
    
    If Version >= 2.395 Then
        Get #FileNum, , LastModified
    End If

    Get #FileNum, , I
        
    Do Until I = 0
        Set QClause = New QueryClauseClass
        QClause.InputFromBinary FileNum, Version
        Clauses.Append QClause
        I = I - 1
    Loop

End Sub

Public Function ClauseCount() As Integer
'
' Name:         ClauseCount
' Description:  Returns the number of clauses in the query
' Returns:      the number of clauses in the query
'

    ClauseCount = Clauses.Count

End Function

Public Function Clause() As QueryClauseClass
'
' Name:         Clause
' Description:  Return the clause at the current cursor position
' Returns:      the clause at the current cursor position
'

    Set Clause = Clauses.Item
    
End Function

Public Function ClauseDescNext() As String
'
' Name:         ClauseDescNext
' Description:  Return an english description of the current clause, in the
'               context of the query (including AND or OR in the line as needed).
'               Advance the cursor.
'

    If Not Clauses.Off Then
        ClauseDescNext = Clauses.Item.Description
        Clauses.MoveNext
        If Not Clauses.Off Then
            ClauseDescNext = ClauseDescNext & IIf(MatchAll, " AND", " OR")
        End If
    End If

End Function

Public Sub AddClause(Key As String, Find As String, Number As Double, Comp As QueryCompareType, _
                     CompNot As Boolean)
'
' Name:         AddClause
' Parameters:   Key         the key of the data to search
'               Find        a string value to search for
'               Number      a numeric value to compare to
'               Comp        the comparison to make
'               CompNot     whether to NOT the result
' Description:  Add a clause to the end of the query.
'

    Dim QC As QueryClauseClass
    Set QC = New QueryClauseClass
    
    QC.Key = Key
    QC.Find = Find
    QC.Number = Number
    QC.Comparison = Comp
    QC.CompNot = CompNot
    
    Clauses.Append QC

End Sub

Public Sub Remove(ByVal Index As Integer)
'
' Name:         Remove
' Paramaters:   Index           0-based position of clause to remove
' Description:  Remove the clase from the given position.
'

    Clauses.First
    Do Until Index <= 0
        Clauses.MoveNext
        Index = Index - 1
    Loop
    Clauses.Remove

End Sub

Public Sub RemoveLast()
'
' Name:         RemoveLast
' Description:  Remove the last clause from the list.
'

    Clauses.Last
    Clauses.Remove

End Sub

Public Function IsEmpty() As Boolean
'
' Name:         IsEmpty
' Description:  Tells whether this query has clauses or not.
' Returns:      TRUE if the query is empty, FALSE otherwise.
'

    IsEmpty = Clauses.IsEmpty

End Function

Public Function Off() As Boolean
'
' Name:         Off
' Description:  Tells whether the cursor has moved off the clauses.
' Returns:      TRUE if the cursor is off the list, FALSE otherwise.
'

    Off = Clauses.Off

End Function

Public Sub MoveNext()
'
' Name:         MoveNext
' Description:  Move the cursor to the next clause .
'

    Clauses.MoveNext

End Sub

Public Sub First()
'
' Name:         First
' Description:  Move the cursor to the first clause.
'

    Clauses.First

End Sub

Public Sub Last()
'
' Name:         Last
' Description:  Move the cursor to the last clause.
'

    Clauses.Last

End Sub

Public Sub Clear()
'
' Name:         Clear
' Description:  Erase this query.
'

    Clauses.Clear
    MatchAll = False
    Inventory = qicharacters
    SortKey = ""
    Name = ""
    
End Sub

Public Sub MoveToClause(ByVal Index As Integer)
'
' Name:         MoveToClause
' Parameters:   Index           Zero-based position of the clause to move to
' Description:  Move to the indicated clause of the query.
'

    Clauses.First
    Do Until Index = 0 Or Clauses.Off
        Index = Index - 1
        Clauses.MoveNext
    Loop

End Sub

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Create class members.
'

    LastModified = Now
    Set Clauses = New LinkedList
    
End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Clean up class members.
'

    Set Clauses = Nothing

End Sub
