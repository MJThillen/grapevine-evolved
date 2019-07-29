VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExperienceClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FirstNode As ExperienceHistoryNode  'First history entry
Private LastNode As ExperienceHistoryNode   'Last history entry
Private Cursor As ExperienceHistoryNode     'Current history node
Private NodeCount As Integer                'Number of history entries

Public Earned As Single                     'Experience Earned
Public Unspent As Single                    'Experience Unspent
Attribute Unspent.VB_VarMemberFlags = "200"

Public Sub Copy(Other As ExperienceClass)
'
' Name:         Copy
' Parameters:   Other       Experience class to duplicate
' Description:  Make this class a copy of the other one.
'

    Clear
    Other.First
    Do Until Other.Off
        Insert Other.EntryChange, Other.EntryChangeType, Other.EntryDate, Other.EntryReason
        Other.MoveNext
    Loop

    Earned = Other.Earned
    Unspent = Other.Unspent

End Sub

Public Function OutputID() As Integer
'
' Name:         OutputID
' Description:  Return an ID that tells the OutputEngineClass how to manipulate this object
'
    OutputID = oidHistory

End Function

Public Sub GetValue(Key As String, ObjData As Variant)
'
' Name:         GetValue
' Parameters:   Key         a key (string) to the value to retreive
'               CharData    the returned value - a string, number, date or object
' Description:  Use the given key to return data from the object.
'

    ObjData = ""
    Select Case Key
        Case qkXPUnspent:   ObjData = Unspent
        Case qkPPUnspent:   ObjData = Unspent
        Case qkXPEarned:    ObjData = Earned
        Case qkPPEarned:    ObjData = Earned
        Case qkDate:        If Not Off Then ObjData = Cursor.When
        Case qkChange:      If Not Off Then ObjData = Cursor.Change
        Case qkChangeText:  If Not Off Then ObjData = EntryChangeText
        Case qkReason:      If Not Off Then ObjData = Cursor.Reason
        Case qkUnspent:     If Not Off Then ObjData = Cursor.Unspent
        Case qkEarned:      If Not Off Then ObjData = Cursor.Earned
        Case qkCount:       ObjData = Count
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
            If .StartDate = 0 Then
                First
            Else
                MoveToDate .StartDate, True
            End If
        Else
            If .EndDate = 0 Then
                Last
            Else
                MoveToDate .EndDate, False
            End If
        End If
    End With
    
End Sub

Public Function EntryChange() As Single
'
' Name:         EntryChange
' Description:  Return the value of the current history entry's Change data.
' Returns:      as above.
'

    EntryChange = Cursor.Change

End Function

Public Function EntryChangeType() As ExperienceChangeType
'
' Name:         EntryChangeType
' Description:  Return the value of the current history entry's ChangeType.
' Returns:      as above.
'

    EntryChangeType = Cursor.ChangeType

End Function

Public Function EntryChangeTypeText() As String
'
' Name:         EntryChangeTypeText
' Description:  Return a string describing the type of the current entry.
' Returns:      as above.
'

    Select Case Cursor.ChangeType
        Case ecEarned
            EntryChangeTypeText = "earned"
        Case ecSpent
            EntryChangeTypeText = "spent"
        Case ecDeducted
            EntryChangeTypeText = "deducted"
        Case ecUnspent
            EntryChangeTypeText = "unspent"
        Case ecSetEarned
            EntryChangeTypeText = "set earned"
        Case ecSetUnspent
            EntryChangeTypeText = "set unspent"
        Case ecComment
            EntryChangeTypeText = "comment"
    End Select

End Function

Public Function EntryChangeText() As String
'
' Name:         EntryChangeText
' Description:  Return a string describing the change of the current entry.
' Returns:      as above.
'

    If Cursor.ChangeType = ecComment Then
        EntryChangeText = "comment"
    Else
        EntryChangeText = CStr(Cursor.Change) & " " & EntryChangeTypeText
    End If

End Function

Public Function EntryDate() As Date
'
' Name:
' Description:  Return the value of the current history entry's date.
' Returns:      as above.
'

    EntryDate = Cursor.When

End Function

Public Function EntryReason() As String
'
' Name:
' Description:  Return the value of the current history entry's reason.
' Returns:      as above.
'

    EntryReason = Cursor.Reason

End Function

Public Function EntryEarned() As Single
'
' Name:
' Description:  Return the value of the current history entry's Earned exp.
' Returns:      as above.
'

    EntryEarned = Cursor.Earned

End Function

Public Function EntryUnspent() As Single
'
' Name:
' Description:  Return the value of the current history entry's unspent exp.
' Returns:      as above.
'

    EntryUnspent = Cursor.Unspent

End Function

Public Sub MoveToDate(ToDate As Date, Ascend As Boolean)
'
' Name:         MoveToDate
' Parameters:   ToDate      Date to move to
'               Ascend      whether to approach that date from the beginning or the end
' Description:  Move to an instance of the given date: the first instance if Ascend = TRUE,
'               the last if Ascend = False.  If the date does not exist, stop at the next
'               entry is Ascend = True, the previous entry if Ascend = False.
'

    If Ascend Then
        First
        Do Until Off
            If Cursor.When >= ToDate Then Exit Do
            MoveNext
        Loop
    Else
        Last
        Do Until Off
            If Cursor.When <= ToDate + #11:59:59 PM# Then Exit Do
            MovePrevious
        Loop
    End If

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

Public Function GetMonthXP(MonthDate As Date) As Single
'
' Name:         GetMonthXP
' Description:  Return the amount of XP earned in the month of the given date.
'

    Dim Y As Integer
    Dim M As Integer
    
    Y = Year(MonthDate)
    M = Month(MonthDate)
    GetMonthXP = 0
    
    First
    Do Until Cursor Is Nothing
        If Year(Cursor.When) = Y And Month(Cursor.When) = M Then
            If Cursor.ChangeType = ecEarned Then GetMonthXP = GetMonthXP + Cursor.Change
        ElseIf Cursor.When > MonthDate Then
            Exit Do
        End If
        Set Cursor = Cursor.NextNode
    Loop

End Function

Public Function Count() As Integer
'
' Name:         Count
' Description:  Return the number of history entries in the list.
' Returns:      as above.
'

    Count = NodeCount

End Function

Public Sub Insert(Change As Single, ChangeType As ExperienceChangeType, When As Date, _
        Reason As String)
'
' Name:         Insert
' Parameters:   Change          amount of the change
'               ChangeType      type of change
'               When            date of change
'               Reason          reason for change
' Description:  Insert a new history entry into the list, sorted chronologically.
'

    Me.Last
    Do Until Off
        If When < Cursor.When Then
            Me.MovePrevious
        Else
            Exit Do
        End If
    Loop
    
    Dim NewNode As ExperienceHistoryNode
    Set NewNode = New ExperienceHistoryNode
    
    Set NewNode.PrevNode = Cursor
    
    If Off Then         'Place at the front
        Set NewNode.NextNode = FirstNode
        Set FirstNode = NewNode
    Else                'Insert after cursor
        Set NewNode.NextNode = Cursor.NextNode
        Set Cursor.NextNode = NewNode
    End If
    
    If NewNode.NextNode Is Nothing Then 'last node
        Set LastNode = NewNode
    Else
        Set NewNode.NextNode.PrevNode = NewNode
    End If
    
    Set Cursor = NewNode
    
    NewNode.When = When
    NewNode.Reason = Reason
    NewNode.Change = Change
    NewNode.ChangeType = ChangeType
    NewNode.PropogateChange
    NodeCount = NodeCount + 1
    
    If Game.EnforceHistory Then
        Last
        Earned = Cursor.Earned
        Unspent = Cursor.Unspent
    Else
        Select Case ChangeType
            Case ecEarned
                Earned = Earned + Change
                Unspent = Unspent + Change
            Case ecDeducted
                Earned = Earned - Change
                Unspent = Unspent - Change
            Case ecSetEarned
                Earned = Change
            Case ecSpent
                Unspent = Unspent - Change
            Case ecUnspent
                Unspent = Unspent + Change
            Case ecSetUnspent
                Unspent = Change
            Case ecComment
                'No change
        End Select
    End If

End Sub

Public Sub Remove()
'
' Name:         Remove
' Description:  Totally remove the entry from the current cursor
'               position.
'

    Dim NextCursor As ExperienceHistoryNode

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
        
        If Not Cursor Is Nothing Then Cursor.PropogateChange
        NodeCount = NodeCount - 1
    
        If Game.EnforceHistory And Not IsEmpty Then
            Last
            Earned = Cursor.Earned
            Unspent = Cursor.Unspent
        End If
    
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
    Set FirstNode = Nothing
    Do Until Cursor Is Nothing
        Set Cursor.PrevNode = Nothing
        If Not Cursor.NextNode Is Nothing Then
            Set Cursor = Cursor.NextNode
        Else
            Set Cursor = Nothing
            Set LastNode = Nothing
        End If
    Loop
    NodeCount = 0

End Sub

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write the history to a file.
'

    Dim MemCursor As ExperienceHistoryNode

    Set MemCursor = Cursor

    With XML
        
        .BeginTag "experience"
        
            .WriteAttribute "unspent", Unspent
            .WriteAttribute "earned", Earned
        
            First
            Do Until Cursor Is Nothing
                .BeginTag "entry"
                    .WriteAttribute "date", Cursor.When
                    .WriteAttribute "change", Cursor.Change
                    .WriteAttribute "type", Cursor.ChangeType
                    .WriteAttribute "reason", Cursor.Reason
                    .WriteAttribute "earned", Cursor.Earned
                    .WriteAttribute "unspent", Cursor.Unspent
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
' Parameters:   XML         XMLReaderClass that has read the opening tag
'               Version     version tag of the file format
' Description:  Read a ExperienceClass from a file.
'

    Clear
        
    With XML
    
        If .Tag = "experience" Then
        
            Unspent = .GetSingle("unspent")
            Earned = .GetSingle("earned")
        
            Dim NewNode As ExperienceHistoryNode
            'assign properties
            
            Do Until .Tag = "experience" And .Terminal
                .ReadTag
                If .Tag = "entry" Then
                    
                    Set NewNode = New ExperienceHistoryNode
                    
                    If NodeCount = 0 Then
                        Set FirstNode = NewNode
                    Else
                        Set LastNode.NextNode = NewNode
                        Set NewNode.PrevNode = LastNode
                    End If
                    Set LastNode = NewNode
                    NodeCount = NodeCount + 1
                    
                    NewNode.When = .GetDate("date")
                    NewNode.Change = .GetSingle("change")
                    NewNode.ChangeType = .GetLong("type")
                    NewNode.Reason = .GetString("reason")
                    NewNode.Earned = .GetSingle("earned")
                    NewNode.Unspent = .GetSingle("unspent")
                    
                End If
            Loop
        
        Else
            Err.Raise vbObjectError + 8, "ExperienceClass", _
                      "Wrong file input to experience history"
        End If
    
    End With

End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this character's data to a binary file.
'

    Dim MemCursor As ExperienceHistoryNode

    Set MemCursor = Cursor

    Put #FileNum, , Unspent
    Put #FileNum, , Earned

    Put #FileNum, , CInt(Count)

    First
    Do Until Cursor Is Nothing
        Cursor.OutputToBinary FileNum
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
    Dim NewNode As ExperienceHistoryNode

    Clear
        
    Get #FileNum, , Unspent
    Get #FileNum, , Earned
        
    Get #FileNum, , I
    Do Until I = 0
        
        Set NewNode = New ExperienceHistoryNode
        
        If NodeCount = 0 Then
            Set FirstNode = NewNode
        Else
            Set LastNode.NextNode = NewNode
            Set NewNode.PrevNode = LastNode
        End If
        
        Set LastNode = NewNode
        NodeCount = NodeCount + 1
        
        NewNode.InputFromBinary FileNum, Version
        
        I = I - 1
    Loop

End Sub

Public Sub OldInputFromFile(FileNum As Integer, Version As String)
'
' Name:         InputFromFile
' Parameters:   FileNum     number of the file to read from
'               Version     version tag of the file format
' Description:  Read a ExperienceClass from a file.
'
        
    Dim CType As ExperienceChangeType
    Dim Change As Single
    Dim Reason As String
    Dim KeyPhrase As String
    Dim Start As Integer
    Dim PrevEarned As Single
    Dim PrevUnspent As Single
    Dim Read As String
            
    PrevEarned = Earned
    PrevUnspent = Unspent
    Earned = 0
    Unspent = 0
    
    Clear
    
    Do
        
        KeyPhrase = ""
        If Read Like "*:  Added *" Then
            CType = ecEarned
            KeyPhrase = "Added "
        ElseIf Read Like "*:  Subtracted *" Then
            CType = ecSpent
            KeyPhrase = "Subtracted "
        ElseIf Read Like "*:  Set * to *" Then
            CType = ecSetUnspent
            KeyPhrase = " to "
        End If
        
        If KeyPhrase <> "" Then
            Start = InStr(Read, KeyPhrase) + Len(KeyPhrase)
            Change = Val(Mid(Read, Start, 3))
            KeyPhrase = Left(Read, InStr(Read, ":") - 1)
            If IsDate(KeyPhrase) Then
                Line Input #FileNum, Read
                Reason = ""
                Do While Read <> "" And Not Read Like "<-End * History->"
                    Reason = Reason & Read & " "
                    Line Input #FileNum, Read
                Loop
                Reason = Trim(Reason)
                Insert Change, CType, CDate(KeyPhrase), Reason
            Else
                Line Input #FileNum, Read
            End If
        Else
            Line Input #FileNum, Read
        End If
        
    Loop Until Read Like "<-End * History->"

    'Insert PrevEarned, ecSetEarned, Date, "Converted to " & GrapevineCaption
    Insert PrevUnspent, ecSetUnspent, Date, "Unspent XP upon conversion to " & GrapevineCaption
    
End Sub

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Initialize an empty history list.  VB calls this method
'               automatically.
'

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

