VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LinkedRumorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Cursor As LinkedRumorListNode       'Current location in the list
Private FirstNode As LinkedRumorListNode    'First node of the list
Private LastNode As LinkedRumorListNode     'Last node of the list
Private NodeCount As Integer                'number of nodes in the list

Public DateStamp As Date                    'The date of these Rumors

Public Function Name() As String
'
' Name:         Name
' Description:  Return a name for this list -- a string version of its date.
' Returns:      A string version of its date.
'

    Name = CStr(DateStamp)

End Function

Public Function Count() As Integer
'
' Name:         Count
' Description:  Returns the number of items in the list
' Returns:      the number of items in the list
'

    Count = NodeCount
    
End Function

Public Sub MoveTo(RumorType As Integer, Value As String, _
        Level As Integer, Optional UsedBy As String = "")
'
' Name:         MoveTo
' Parameters:   RumorType       the type of rumor to find
'               Value           the recipient to find
'               Level           the level of rumor to find
'               UsedBy          the person who used the level
' Description:  Move the cursor to the rumor that matches the
'               parameters.  Move off if no such rumor is found.
'

    Dim Found As Boolean
    
    Found = False
    First
    
    Do Until Off Or Found
        With Cursor.Data
            Found = (.Category = RumorType _
                    And .Recipient = Value _
                    And .Level = Level)
            If UsedBy <> "" Then Found = _
                    Found And (.UsedBy = UsedBy)
        End With
        If Not Found Then MoveNext
    Loop

End Sub

Public Function Item() As OldRumorClass
'
' Name:         Item
' Description:  Return the rumor at the current cursor position
' Returns:      the rumor at the current cursor position
'

    Debug.Assert Not Cursor Is Nothing
    Set Item = Cursor.Data
    
End Function

Public Sub InsertSorted(Data As OldRumorClass)
'
' Name:         InsertSorted
' Parameters:   Data    the rumor to insert
' Description:  Insert a rumor into the list, sorting first by type,
'               then by recipient, then by UsedBy, then by Level.
'

    Dim NewNode As LinkedRumorListNode
    
    First
    Do
        
        If Off Then 'insert at end
            Append Data
            Exit Do
        End If

        If Cursor.Data.Category > Data.Category Or _
 _
                (Cursor.Data.Category = Data.Category And _
                Cursor.Data.Recipient > Data.Recipient) Or _
 _
                (Cursor.Data.Category = Data.Category And _
                Cursor.Data.Recipient = Data.Recipient And _
                Cursor.Data.UsedBy > Data.UsedBy) Or _
 _
                (Cursor.Data.Category = Data.Category And _
                Cursor.Data.Recipient = Data.Recipient And _
                Cursor.Data.UsedBy = Data.UsedBy And _
                Cursor.Data.Level > Cursor.Data.Level) Then 'Insert before present node
        
            Set NewNode = New LinkedRumorListNode
            Set NewNode.Data = Data
            
            Set NewNode.PrevNode = Cursor.PrevNode
            Set NewNode.NextNode = Cursor
            
            If Cursor.PrevNode Is Nothing Then
                Set FirstNode = NewNode
            Else
                Set Cursor.PrevNode.NextNode = NewNode
            End If
            
            Set Cursor.PrevNode = NewNode
            NodeCount = NodeCount + 1
            Exit Do
            
        Else
            MoveNext
        End If
        
    Loop

End Sub

Public Sub Append(Data As OldRumorClass)
'
' Name:         Append
' Parameters:   Data    the rumor to add to the list
' Description:  Add a rumor to the end of the list.
'

    Dim NewNode As LinkedRumorListNode
    
    Set NewNode = New LinkedRumorListNode
    Set NewNode.Data = Data
    Set NewNode.NextNode = Nothing
    
    If LastNode Is Nothing Then
        Set FirstNode = NewNode
    Else
        Set LastNode.NextNode = NewNode
        Set NewNode.PrevNode = LastNode
    End If
    
    Set LastNode = NewNode
    NodeCount = NodeCount + 1

End Sub
Public Sub Remove()
'
' Name:         Remove
' Description:  Remove the rumor at the current cursor position.
'

    Dim NextCursor As LinkedRumorListNode

    If Not Off() Then
    
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
        Set Cursor.Data = Nothing
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

    IsEmpty = (Count() = 0)

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

    If Not Off() Then Set Cursor = Cursor.NextNode

End Sub
Public Sub MovePrevious()
'
' Name:         MovePrevious
' Description:  Move the cursor to the previous item.
'

    If Not Off() Then Set Cursor = Cursor.PrevNode

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

    First
    Do Until Off
        Remove
    Loop

End Sub

Public Sub InputFromFile(FileNum As Integer, Version As String)
'
' Name:         InputFromFile
' Parameters:   FileNum     number of an open file
'               Version     version tag of the file format
' Description:  Read a list of rumors from a file.
'

    Dim Read As String
    Dim Rumor As OldRumorClass

    Clear
    Line Input #FileNum, Read
    DateStamp = CDate(Read)
    
    Line Input #FileNum, Read
    Do Until Read = "<-End of LinkedRumorList Items->"
        
        Set Rumor = New OldRumorClass
        With Rumor
        
            .Category = Val(Read)
            Line Input #FileNum, Read
            .Recipient = Read
            Line Input #FileNum, Read
            .Level = Val(Read)
            .Text = ReadLongField(FileNum, "<-End Rumor Text->")
            
            If Version <> GameFileVersionTag0 Then
                Line Input #FileNum, Read
                .UsedBy = Read
                Line Input #FileNum, Read
                .Exclude = CBool(Read)
            End If
            
        End With
        
        InsertSorted Rumor
        Line Input #FileNum, Read
    
    Loop

End Sub

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Initialize an empty list.  VB calls this method
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
