VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalendarClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FirstNode As CalendarNode           'First history entry
Private LastNode As CalendarNode            'Last history entry
Private Cursor As CalendarNode              'Current history node
Private NodeCount As Integer                'Number of history entries

Public LastModified As Date                 'When the calendar was last modified

Public Sub MoveTo(When As Date)
'
' Name:         MoveTo
' Parameters:   When        the date to move to
' Description:  Move the cursor to the given date.
'

    First
    Do Until Cursor Is Nothing
        If Not Cursor.GameDate = When Then
            MoveNext
        Else
            Exit Do
        End If
    Loop

End Sub

Public Function OutputID() As Integer
'
' Name:         OutputID
' Description:  Return an ID that tells the OutputEngineClass how to manipulate this object
'
    OutputID = oidCalendar

End Function

Public Sub GetValue(Key As String, ObjData As Variant)
'
' Name:         GetValue
' Parameters:   Key         a key (string) to the value to retreive
'               ObjData     the returned value - a string, number, date or object
' Description:  Use the given key to return data from the object.
'

    ObjData = ""
    Select Case Key
        Case qkDate:        If Not Off Then ObjData = Cursor.GameDate
        Case qkPlace:       If Not Off Then ObjData = Cursor.Place
        Case qkTime:        If Not Off Then ObjData = Cursor.Time
        Case qkNotes:       If Not Off Then ObjData = Cursor.Notes
        Case Else:          ObjData = Null
    End Select

End Sub

Public Sub InitializeForOutput()
'
' Name:         InitializeForOutput
' Description:  Position the object in a starting state, ready for output.
'

    With OutputEngine
        If .AscendDate Then
            First
            If .StartDate <> 0 Then
                Do Until Off
                    If Cursor.GameDate >= .StartDate Then Exit Sub
                    MoveNext
                Loop
            End If
        Else
            Last
            If .EndDate <> 0 Then
                Do Until Off
                    If Cursor.GameDate <= .EndDate + #11:59:59 PM# Then Exit Sub
                    MovePrevious
                Loop
            End If
        End If
    End With
    
End Sub

Public Function NextGameDate() As Date
'
' Name:         NextGameDate
' Returns:      Date of the next game.
'

    Dim FindNode As CalendarNode
    
    Set FindNode = FirstNode
    
    Do Until Date <= FindNode.GameDate
        Set FindNode = FindNode.NextNode
        If FindNode Is Nothing Then Exit Do
    Loop

    If Not FindNode Is Nothing Then
        NextGameDate = FindNode.GameDate
    Else
        NextGameDate = 0
    End If

End Function

Public Function HasNextGame() As Boolean
'
' Name:         HasNextGame
' Returns:      TRUE iff there is a game after today's date.
'

    HasNextGame = False
    If Not LastNode Is Nothing Then HasNextGame = (Date <= LastNode.GameDate)

End Function

Public Function PreviousGameDate() As Date
'
' Name:         PreviousGameDate
' Returns:      Date of the previous game.
'

    Dim FindNode As CalendarNode
    
    Set FindNode = LastNode
    
    Do Until Date > FindNode.GameDate
        Set FindNode = FindNode.PrevNode
        If FindNode Is Nothing Then Exit Do
    Loop

    If Not FindNode Is Nothing Then
        PreviousGameDate = FindNode.GameDate
    Else
        PreviousGameDate = 0
    End If

End Function

Public Function HasPreviousGame() As Boolean
'
' Name:         HasPreviousGame
' Returns:      TRUE iff there is a game after today's date.
'

    HasPreviousGame = False
    If Not FirstNode Is Nothing Then HasPreviousGame = (Date > FirstNode.GameDate)

End Function

Public Sub MoveToCloseGame()
'
' Name:         MoveToCloseGame
' Description:  Move to the next game in time, or if there is none, the last game.
'
    First
    Do Until Off
        If Date <= Cursor.GameDate Then Exit Sub
        MoveNext
    Loop
    Last

End Sub

Public Function GetGameDate() As Date
'
' Name:         GetGameDate
' Returns:      Date of the game under the cursor.
'

    If Not (Cursor Is Nothing) Then GetGameDate = Cursor.GameDate

End Function

Public Function GetGamePlace() As String
'
' Name:         GetGameDate
' Returns:      Place of the game under the cursor.
'

    If Not (Cursor Is Nothing) Then GetGamePlace = Cursor.Place

End Function

Public Function GetGameTime() As String
'
' Name:         GetGameDate
' Returns:      time of the game under the cursor.
'

    If Not (Cursor Is Nothing) Then GetGameTime = Cursor.Time

End Function

Public Function GetGameNotes() As String
'
' Name:         GetGameDate
' Returns:      notes of the game under the cursor.
'

    If Not (Cursor Is Nothing) Then GetGameNotes = Cursor.Notes

End Function

Public Sub SetGameDate(NewDate As Date)
'
' Name:         SetGameDate
' Description:  Change the date of the game under the cursor.
'

    If Not Off Then
        Dim Place As String
        Dim Time As String
        Dim Notes As String
        
        Place = Cursor.Place
        Time = Cursor.Time
        Notes = Cursor.Notes
               
        Remove
        Insert NewDate, Place, Time, Notes
    End If

End Sub

Public Sub SetGamePlace(NewPlace As String)
'
' Name:         SetGamePlace
' Description:  Change the place of the game under the cursor.
'

    If Not (Cursor Is Nothing) Then Cursor.Place = NewPlace

End Sub

Public Sub SetGameTime(NewTime As String)
'
' Name:         SetGameTime
' Description:  Change the Time of the game under the cursor.
'

    If Not (Cursor Is Nothing) Then Cursor.Time = NewTime

End Sub

Public Sub SetGameNotes(NewNotes As String)
'
' Name:         SetGameNotes
' Description:  Change the Notes of the game under the cursor.
'

    If Not (Cursor Is Nothing) Then Cursor.Notes = NewNotes

End Sub

Public Sub MoveToPlace(ByVal Place As Integer)
'
' Name:         MoveToPlace
' Parameters:   Place           the position in the list to move to
' Description:  Move the cursor to the given number position in the list.
'               Like listbox listindexes, this is 0-based.
'
    
    Set Cursor = FirstNode
    Do Until (Place <= 0 Or Cursor Is Nothing)
        Set Cursor = Cursor.NextNode
        Place = Place - 1
    Loop

End Sub

Public Function Count() As Integer
'
' Name:         Count
' Description:  Return the number of dates in the list.
' Returns:      as above.
'

    Count = NodeCount

End Function

Public Sub Insert(GameDate As Date, Time As String, Place As String, Notes As String)
'
' Name:         Insert
' Parameters:   GameDate        Date of a game
'               Time            the time of the game
'               Place           the place of the game
'               Notes           game notes
' Description:  Insert a new game date into the list, sorted chronologically.
'

    Me.Last
    Do Until Off
        If GameDate < Cursor.GameDate Then
            Me.MovePrevious
        Else
            Exit Do
        End If
    Loop
    
    Dim NewNode As CalendarNode
    Set NewNode = New CalendarNode
    
    Set NewNode.PrevNode = Cursor
    
    If Off Then         'Place at the front
        Set NewNode.NextNode = FirstNode
        Set FirstNode = NewNode
    Else                'Insert after cursor
        Set NewNode.NextNode = Cursor.NextNode
        Set Cursor.NextNode = NewNode
    End If
    
    If NewNode.NextNode Is Nothing Then
        Set LastNode = NewNode
    Else
        Set NewNode.NextNode.PrevNode = NewNode
    End If
    
    Set Cursor = NewNode
    
    NewNode.GameDate = GameDate
    NewNode.Time = Time
    NewNode.Place = Place
    NewNode.Notes = Notes
    NodeCount = NodeCount + 1
    
End Sub

Public Sub Remove()
'
' Name:         Remove
' Description:  Totally remove the entry from the current cursor
'               position.
'

    Dim NextCursor As CalendarNode

    If Not Cursor Is Nothing Then
    
        If Cursor Is FirstNode Then
            Set FirstNode = Cursor.NextNode
        Else
            Set Cursor.PrevNode.NextNode = Cursor.NextNode
        End If
        
        If Cursor Is LastNode Then
            Set LastNode = Cursor.PrevNode
        Else
            Set Cursor.NextNode.PrevNode = Cursor.PrevNode
        End If
        
        Set NextCursor = Cursor.NextNode
        Set Cursor = Nothing
        Set Cursor = NextCursor
        
        NodeCount = NodeCount - 1
    
    End If

End Sub

Public Function IsEmpty() As Boolean
'
' Name:         IsEmpty
' Description:  Tells whether the list is empty or not.
' Returns:      TRUE if the list is empty, FALSE otherwise.
'

    IsEmpty = (NodeCount = 0)

End Function

Public Function Off() As Boolean
'
' Name:         Off
' Description:  Tells whether the cursor has moved off the list.
' Returns:      TRUE if the cursor is off the list, FALSE otherwise.
'

    Off = (Cursor Is Nothing)

End Function

Public Sub MoveNext()
'
' Name:         MoveNext
' Description:  Move the cursor to the next item.
'

    If Not (Cursor Is Nothing) Then Set Cursor = Cursor.NextNode

End Sub

Public Sub MovePrevious()
'
' Name:         MovePrevious
' Description:  Move the cursor to the previous item.
'

    If Not (Cursor Is Nothing) Then Set Cursor = Cursor.PrevNode

End Sub

Public Sub First()
'
' Name:         First
' Description:  Move the cursor to the first item.
'

    Set Cursor = FirstNode

End Sub

Public Sub Last()
'
' Name:         Last
' Description:  Move the cursor to the last item.
'
    
    Set Cursor = LastNode
    
End Sub

Public Sub Clear()
'
' Name:         Clear
' Description:  Empty the list of all items.
'

    Me.First
    Do Until Off()
        Me.Remove
    Loop

End Sub

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write the calendar to a file.
'

    Dim MemCursor As CalendarNode

    Set MemCursor = Cursor

    With XML
    
        .BeginTag "calendar"
        .WriteAttribute "lastmodified", LastModified
        
            First
            Do Until Cursor Is Nothing
                .BeginTag "game"
                    .WriteAttribute "date", Cursor.GameDate
                    .WriteAttribute "time", Cursor.Time
                    .WriteCDataTag "place", Cursor.Place
                    .WriteCDataTag "notes", Cursor.Notes
                .EndTag
                MoveNext
            Loop
        
        .EndTag
        
    End With
    
    Set Cursor = MemCursor
    
End Sub

Public Sub InputFromFile(XML As XMLReaderClass, Version As Double)
'
' Name:         InputFromFile
' Parameters:   XML         XMLReaderClass that has just read the calendar tag
'               Version     version tag of the file format
' Description:  Read a calendar from a file.
'

    Dim GDate As Date
    Dim GTime As String
    Dim GPlace As String
    Dim GNotes As String
    
    With XML

        If .Tag = "calendar" Then

            LastModified = .GetDate("lastmodified", Now)
            
            Clear

            Do Until .Tag = "calendar" And .Terminal
                .ReadTag
                If .Tag = "game" Then
                
                    GDate = .GetDate("date")
                    GTime = .GetString("time")
                    GPlace = ""
                    GNotes = ""
                    Do Until .Tag = "game" And .Terminal
                        .ReadTag
                        Select Case .Tag
                            Case "place"
                                If Not .Terminal Then
                                    .ReadTag
                                    GPlace = .GetTagText
                                End If
                            Case "notes"
                                If Not .Terminal Then
                                    .ReadTag
                                    GNotes = .GetTagText
                                End If
                        End Select
                    Loop
                
                    Insert GDate, GTime, GPlace, GNotes
                
                End If
            Loop

        Else
            Err.Raise vbObjectError + 8, "CalendarClass", _
                      "Wrong file input to calendar"
        End If

    End With

End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this character's data to a binary file.
'

    Dim MemCursor As CalendarNode

    Set MemCursor = Cursor

    Put #FileNum, , LastModified
    Put #FileNum, , CInt(Count)
    
    First
    Do Until Cursor Is Nothing
        Put #FileNum, , Cursor.GameDate
        PutStrB FileNum, Cursor.Time
        PutStrB FileNum, Cursor.Place
        PutStrB FileNum, Cursor.Notes
        MoveNext
    Loop

    Set Cursor = MemCursor

End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this character's data from a binary file.
'

    Dim I As Integer
    Dim GDate As Date
    Dim GTime As String
    Dim GPlace As String
    Dim GNotes As String
    
    If Version >= 2.395 Then Get #FileNum, , LastModified
    
    Clear
    Get #FileNum, , I
        
    Do Until I = 0
        Get #FileNum, , GDate
        GetStrB FileNum, GTime
        GetStrB FileNum, GPlace
        GetStrB FileNum, GNotes
        Insert GDate, GTime, GPlace, GNotes
        I = I - 1
    Loop

End Sub

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Initialize an empty history list.  VB calls this method
'               automatically.
'

    LastModified = Now
    Set Cursor = Nothing
    Set FirstNode = Nothing
    Set LastNode = Nothing
    NodeCount = 0
    
End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  If a list is destroyed, clear it first.  VB calls
'               this method automatically.
'

    Me.Clear

End Sub
