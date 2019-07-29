VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LinkedTraitList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Cursor As LinkedTraitListNode       'Current location in the list
Private FirstNode As LinkedTraitListNode    'First node of the list
Private LastNode As LinkedTraitListNode     'Last node of the list

Private NodeCount As Integer                'number of nodes in the list
Private TraitTotal As Double                'number of levels/costs invested in this list
Private Alphabetized As Boolean             'whether or not this list is alphabetized

Public Name As String                       'Name of this list
Public Negative As Boolean                  'whether this is a list of negative traits/flaws
Public Atomic As Boolean                    'whether this list's traits are atomic (add/remove
                                            '   doesn't increment/decrement, duplicates allowed)
Public Display As ListDisplayType           'How to display traits

Public Sub Initialize(myName As String, myAlpha As Boolean, myNeg As Boolean, _
        myAtom As Boolean, myDisp As ListDisplayType)
'
' Name:         Initialize
' Parameters:   myName          a name for the list
'               myAlpha         whether or not it's alphabetized
'               myNeg           whether it's negative
'               myAtom          whether it's atomic
'               myDisp          the display format it uses
' Description:  Clear a list and initialize it with the given parameters.
'

    Me.Clear
    Name = myName
    Alphabetized = myAlpha
    Negative = myNeg
    Atomic = myAtom
    Display = myDisp
          
End Sub

Public Sub Copy(Other As LinkedTraitList)
'
' Name:         Copy
' Parameters:   Other       a LinkedTraitList to copy
' Description:  Make this list an exact copy of the other list.
'

    Name = Other.Name
    Alphabetized = Other.IsAlphabetized
    Negative = Other.Negative
    Atomic = Other.Atomic
    Display = Other.Display
    
    Clear
    Other.First
    Do Until Other.Off
        AppendDirect Other.Trait.Name, Other.Trait.Total, Other.Trait.Note
        Other.MoveNext
    Loop

End Sub

Public Function Count() As Integer
'
' Name:         Count
' Description:  Returns the number of trait levels invested in the list (for non-atomic lists)
'               ot the number of traits in the list (for atomic lists).
' Returns:      the number of trait levels invested in the list
'

    If Atomic Then
        Count = NodeCount
    Else
        Count = TraitTotal
    End If
    
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
        Case qkCount:       ObjData = Count
        Case qkDefault:     If Not Off Then ObjData = DisplayTrait
        Case qkName:        If Not Off Then ObjData = Cursor.Data.Name
        Case qkValue:       If Not Off Then ObjData = Cursor.Data.Total
        Case qkNote:        If Not Off Then ObjData = Cursor.Data.Note
        Case Else:          ObjData = Null
    End Select

End Sub

Public Function DisplayTrait() As String
'
' Name:         DisplayTrait
' Description:  Return the string displaying the current trait --
'               The format is controlled by ListDisplayType
' Returns:      a string displaying the current trait
'

    Dim Dot As String
    Dim MyDisplay As ListDisplayType
    Dim I As Integer

    If OutputEngine.OutputTraits Then
        MyDisplay = IIf(OutputEngine.Display > -1, OutputEngine.Display, Display)
        Dot = OutputEngine.DotSubstitute
    Else
        MyDisplay = Display
        Dot = "o"
    End If

    If Not Cursor.Data Is Nothing Then
    
        With Cursor.Data
            
            'Add the name
            DisplayTrait = .Name
            
            Select Case MyDisplay
                Case ldSimpleNumber
                    DisplayTrait = .Total
                    
                Case ldSimpleNote
                    DisplayTrait = .Note
                    
                Case ldDotSeparate
                
                    DisplayTrait = DisplayTrait & IIf(.Note = "", "", " (" & .Note & ")")
                    Dot = Dot & DisplayTrait
                    
                    I = Int(Val(.Total))
                    Do Until I < 2
                        DisplayTrait = DisplayTrait & Dot
                        I = I - 1
                    Loop
                
                Case Else
                
                    'Add the multiplier
                    If MyDisplay = ldMultiplierDot Or _
                            (MyDisplay = ldMultiplier And .Total <> "1") Then
                        DisplayTrait = DisplayTrait & " x" & .Total
                    End If
                    
                    'Add the dots
                    If MyDisplay = ldMultiplierDot Or MyDisplay = ldDot Or MyDisplay = ldSimpleDots Then
                        DisplayTrait = DisplayTrait & " "
                        If MyDisplay = ldSimpleDots Then DisplayTrait = ""
                        I = Int(Val(.Total))
                        Do Until I < 1
                            DisplayTrait = DisplayTrait & Dot
                            I = I - 1
                        Loop
                    End If
                    
                    'Add the parens, cost if needed, and note if needed
                    If MyDisplay = ldCost Or MyDisplay = ldCostOnly Then
                        
                        DisplayTrait = DisplayTrait & " (" & .Total
                        If .Note <> "" And MyDisplay = ldCost Then
                            DisplayTrait = DisplayTrait & ", " & .Note
                        End If
                        DisplayTrait = DisplayTrait & ")"
                        
                    ElseIf Not (MyDisplay = ldSimple Or MyDisplay = ldSimpleDots Or .Note = "") Then
                        DisplayTrait = DisplayTrait & " (" & .Note & ")"
                    End If
        
            End Select
        
        End With
    
    End If
    
End Function

Public Function IsAlphabetized() As Boolean
'
' Name:         IsAlphabetized
' Description:  Returns whether this trait list is alphabetized or ordered
' Returns:      Whether this trait list is alphabetized or ordered
'

    IsAlphabetized = Alphabetized

End Function

Public Sub SetAlphabetized(Alpha As Boolean)
'
' Name:         SetAlphabetized
' Parameters:   Alpha       TRUE if this list is to be alphabetized; FALSE
'                           otherwise
' Description:  Set whether this list is to be alphabetized.  If it is
'               set TRUE, alphabetize the list.
'

    Dim X As Integer
    Dim Y As Integer
    Dim Cursor1 As LinkedTraitListNode
    Dim Cursor2 As LinkedTraitListNode
    Dim Swap As TraitClass
    
    If Alpha Then
        
        'Alphabetize the list.  Because this is expected to be a rare
        'operation, I'll just use a bubblesort instead of trying to
        'figure out how to quicksort a linked list.
        
        For X = 1 To NodeCount - 1
            Set Cursor1 = FirstNode
            Set Cursor2 = Cursor1.NextNode
            For Y = 1 To X
                If LCase(Cursor1.Data.Name) > LCase(Cursor2.Data.Name) Then
                    Set Swap = Cursor1.Data
                    Set Cursor1.Data = Cursor2.Data
                    Set Cursor2.Data = Swap
                End If
                Set Cursor1 = Cursor2
                Set Cursor2 = Cursor1.NextNode
            Next Y
        Next X
        
    End If
    
    Alphabetized = Alpha

End Sub

Public Sub SwapFrontward()
'
' Name:         SwapFrontward
' Description:  Swap the contents of the current item toward the front.
'               Stop when the cursor is the first node.  Only works
'               on non-alphabetized lists.
'

    Dim SwapItem As TraitClass
    
    If Not (Alphabetized Or Off) Then
        If Cursor Is FirstNode Then
            Set Cursor = Nothing
        Else
            Set SwapItem = Cursor.Data
            Set Cursor.Data = Cursor.PrevNode.Data
            Set Cursor.PrevNode.Data = SwapItem
            Set Cursor = Cursor.PrevNode
        End If
    End If

End Sub

Public Sub SwapBackward()
'
' Name:         SwapBackward
' Description:  Swap the contents of the current item toward the end.
'               Stop when the cursor is the last node.  Only works
'               on non-alphabetized lists.
'

    Dim SwapItem As TraitClass
    
    If Not (Alphabetized Or Off) Then
        If Cursor Is LastNode Then
            Set Cursor = Nothing
        Else
            Set SwapItem = Cursor.Data
            Set Cursor.Data = Cursor.NextNode.Data
            Set Cursor.NextNode.Data = SwapItem
            Set Cursor = Cursor.NextNode
        End If
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

Public Sub MoveTo(ByVal TraitName As String)
'
' Name:         MoveTo
' Parameters:   TraitName   a trait name to which to move
' Description:  Move the cursor to the trait whose name matches
'               the parameter.  Move off if no such item is found.
'               Case insensitive.
'

    Dim Found As Boolean
    
    TraitName = LCase(TraitName)
    Found = False
    Me.First
    
    Do Until (Cursor Is Nothing) Or Found
        Found = (LCase(Cursor.Data.Name) = TraitName)
        If Not Found Then Me.MoveNext
    Loop

End Sub

Public Sub MoveToNotePart(ByVal NotePart As String)
'
' Name:         MoveToNote
' Parameters:   NotePart        a string to partially match to a Note
' Description:  Move the cursor to the trait whose note partially matches
'               the parameter.  Move off if no such item is found.
'               Case insensitive.
'

    Dim Found As Boolean
    
    NotePart = LCase(NotePart)
    Found = False
    Me.First
    
    Do Until (Cursor Is Nothing) Or Found
        Found = (InStr(LCase(Cursor.Data.Note), NotePart) > 0)
        If Not Found Then Me.MoveNext
    Loop

End Sub

Public Sub MoveToPart(ByVal NamePart As String)
'
' Name:         MoveToPart
' Parameters:   NamePart        a string to partially match to a trait name
' Description:  Move the cursor to the trait whose name partially matches
'               the parameter.  Move off if no such item is found.  Note:
'               DO NOT start from the front of the list.  Start from the
'               current cursor position.
'               Case insensitive.
'

    Dim Found As Boolean
    
    Found = False
    NamePart = LCase(NamePart)
    
    Do Until (Cursor Is Nothing) Or Found
        Found = (InStr(LCase(Cursor.Data.Name), NamePart) > 0)
        If Not Found Then Me.MoveNext
    Loop

End Sub

Public Sub Remove()
'
' Name:         Remove
' Description:  If this list is atomic, remove the current trait.  If not,
'               decrement it.
'

    If Atomic Then
        Me.RemoveTrait
    Else
        Me.Decrement
    End If
    
End Sub
Public Sub Decrement()
'
' Name:         Decrement
' Description:  Decrease the number of current Trait by one.  If
'               the whole number is a non-numeric string, look for
'               the first numeric part and set it to that.  If the
'               the new number is zero and the list is not atomic,
'               remove the trait.
'

    If Not Cursor Is Nothing Then
    
        If Cursor.Data.Number = 1 And Not Atomic Then
            Me.RemoveTrait
        Else
            TraitTotal = TraitTotal - Cursor.Data.Number
            Cursor.Data.Decrement
            TraitTotal = TraitTotal + Cursor.Data.Number
        End If
        
    End If

End Sub

Public Sub Increment()
'
' Name:         Increment
' Description:  Increase the number of current Trait by one.  If
'               the whole number is a non-numeric string, look for
'               the last numeric part and set it to that.
'

    If Not Cursor Is Nothing Then
        
        TraitTotal = TraitTotal - Cursor.Data.Number
        Cursor.Data.Increment
        TraitTotal = TraitTotal + Cursor.Data.Number
        
    End If

End Sub

Public Function EntryCount() As Integer
'
' Name:         EntryCount
' Description:  Count the number of Nodes / TraitClasses in this list.
'               This is not the same of the number of traits, which is
'               the sum of the Total fields of the TraitClasses.
' Returns:      the number of TraitClasses
'

    EntryCount = NodeCount

End Function

Public Function TraitCount() As Integer
'
' Name:         TraitCount
' Description:  Count the number of traits/costs invested in this list.
' Returns:      the number of Traits in all traitclasses in the list.
'

    TraitCount = TraitTotal

End Function

Public Function Trait() As TraitClass
'
' Name:         Trait
' Description:  Return the TraitClass from the current cursor position
' Returns:      the TraitClass from the current cursor position
'

    Set Trait = Cursor.Data
    
End Function

Public Sub Insert(TraitName As String, Optional TraitNum = "1", Optional TraitNote = "")
'
' Name:         Insert
' Parameters:   TraitName       name of the trait to insert.
'               TraitNum        its number/cost
'               TraitNote       its note
' Description:  Insert a TraitClass into the list.  If the list is not atomic,
'               duplicate Traits are added to the existing Traits.  If the list is
'               alphabetized, they are added into sorted position; otherwise they
'               are inserted before the cursor.
'
'

    Dim LCTrait As String
    Dim Found As Boolean            'Does the Trait exist already?
    
    Found = False
    LCTrait = LCase(TraitName)
    
    If Alphabetized Then            'Locate the sorted position
    
        Me.First
        Do Until Off
            If LCTrait > LCase(Cursor.Data.Name) Then
                Me.MoveNext
            Else
                Found = (LCTrait = LCase(Cursor.Data.Name))
                Exit Do
            End If
        Loop
    
    ElseIf Not Atomic Then          'See if the Trait exists already
    
        Dim StoreCursor As LinkedTraitListNode
    
        Set StoreCursor = Cursor
        Me.First
        Do Until Off
            If LCTrait = LCase(Cursor.Data.Name) Then
                Found = True
                Exit Do
            End If
            Me.MoveNext
        Loop
        If Not Found Then Set Cursor = StoreCursor
        
    End If
    
    
    If Found And Not Atomic Then    'We found it AND we have to add to it
    
        TraitTotal = TraitTotal - Cursor.Data.Number
        Cursor.Data.Total = CStr(Cursor.Data.Number + Val(TraitNum))
        TraitTotal = TraitTotal + Cursor.Data.Number
    
    Else                            'Doesn't matter if we found it -- we insert.
    
        If Off Then                 'Append to the end
            AppendDirect TraitName, TraitNum, TraitNote
        Else                        'Insert at cursor
        
            Dim NewNode As LinkedTraitListNode
            Set NewNode = New LinkedTraitListNode
            Set NewNode.Data = New TraitClass
    
            NewNode.Data.Name = TraitName
            NewNode.Data.Total = TraitNum
            NewNode.Data.Note = TraitNote

            Set NewNode.NextNode = Cursor
            Set NewNode.PrevNode = Cursor.PrevNode
            
            If Cursor Is FirstNode Then
                Set FirstNode = NewNode
            Else
                Set Cursor.PrevNode.NextNode = NewNode
            End If
            
            Set Cursor.PrevNode = NewNode
            
            NodeCount = NodeCount + 1
            TraitTotal = TraitTotal + NewNode.Data.Number
        
        End If
    
    End If

End Sub
Public Sub Append(TraitName As String, Optional TraitNum = "1", Optional TraitNote = "")
'
' Name:         Append
' Parameters:   TraitName       name of the trait to insert.
'               TraitNum        its number/cost
'               TraitNote       its note
' Description:  Add an instance of TraitClass to the end of the list,
'               after ensuring it belongs there.
'

    
    If Not LastNode Is Nothing Then
        If Not Atomic Then
            Insert TraitName, TraitNum, TraitNote
            Exit Sub
        ElseIf Alphabetized Then
            If LCase(TraitName) < LCase(LastNode.Data.Name) Then
                Insert TraitName, TraitNum, TraitNote
                Exit Sub
            End If
        End If
    End If

    AppendDirect TraitName, TraitNum, TraitNote
    
End Sub

Private Sub AppendDirect(TraitName As String, Optional TraitNum = "1", Optional TraitNote = "")
'
' Name:         AppendDirect
' Parameters:   TraitName       name of the trait to insert.
'               TraitNum        its number/cost
'               TraitNote       its note
' Description:  Add an instance of TraitClass to the end of the list,
'               without checking whether it belongs there.
'

    Dim NewNode As LinkedTraitListNode

    Set NewNode = New LinkedTraitListNode
    Set NewNode.Data = New TraitClass
    
    NewNode.Data.Name = TraitName
    NewNode.Data.Total = TraitNum
    NewNode.Data.Note = TraitNote
    
    Set NewNode.NextNode = Nothing
    Set NewNode.PrevNode = LastNode

    If LastNode Is Nothing Then
        Set FirstNode = NewNode
    Else
        Set LastNode.NextNode = NewNode
    End If

    Set LastNode = NewNode
    
    NodeCount = NodeCount + 1
    TraitTotal = TraitTotal + NewNode.Data.Number

End Sub

Public Sub RemoveTrait()
'
' Name:         RemoveTrait
' Description:  Totally remove the TraitClass from the current cursor
'               position.
'

    Dim NextCursor As LinkedTraitListNode

    If Not Cursor Is Nothing Then
    
        TraitTotal = TraitTotal - Cursor.Data.Number
    
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

Public Sub RenameTrait(OldName As String, NewName As String)
'
' Name:         RenameTrait
' Parameters:   OldName     Trait to rename
'               NewName     New name for the trait
' Description:  Change the name of a trait.
'
    
    MoveTo OldName
    If Not Cursor Is Nothing Then
        If Alphabetized Then
            Dim OldTotal As String
            Dim OldNote As String
            OldTotal = Cursor.Data.Total
            OldNote = Cursor.Data.Note
            RemoveTrait
            Insert NewName, OldTotal, OldNote
        Else
            Cursor.Data.Name = NewName
        End If
    End If
    
End Sub

Public Sub RandomizeCurrentTraits(NewTotal As Long)
'
' Name:         RandomizeCurrentTraits
' Parameters:   NewTotal        new number of traits to aim for
' Description:  Randomize the levels of traits so that they add up to the new total.
'

    Dim LevelList As LinkedList
    Dim I As Integer
    
    Set LevelList = New LinkedList
    
    Me.First
    Do Until Me.Off
        TraitTotal = TraitTotal - Cursor.Data.Number + 1
        Cursor.Data.Total = "1"
        Me.MoveNext
    Loop
    
    If NewTotal > NodeCount Then
    
        NewTotal = NewTotal - NodeCount
    
        For I = 2 To NodeCount
            LevelList.InsertSorted Int(Rnd * (NewTotal + 1))
        Next I
        LevelList.InsertSorted NewTotal
        
        I = 0
        Me.First
        LevelList.First
        Do Until LevelList.Off Or Me.Off
            NewTotal = 1 + Val(LevelList.Item) - I
            TraitTotal = TraitTotal - Cursor.Data.Number
            Cursor.Data.Total = CStr(NewTotal)
            TraitTotal = TraitTotal + Cursor.Data.Number
            I = Val(LevelList.Item)
            Me.MoveNext
            LevelList.MoveNext
        Loop
    
    ElseIf NewTotal < NodeCount Then
    
        Do Until NodeCount = 0 Or NodeCount = NewTotal
        
            I = Int(Rnd * NodeCount)
            Me.First
            Do Until I = 0
                Me.MoveNext
                I = I - 1
            Loop
            Me.RemoveTrait
        
        Loop
        
    End If
    
    Set LevelList = Nothing

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

    First
    Do Until Cursor Is Nothing
        RemoveTrait
    Loop

End Sub

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write all the traits to a file.
'

    Dim CurNode As LinkedTraitListNode

    Set CurNode = Cursor

    With XML
    
        .BeginTag "traitlist"
    
            .WriteAttribute "name", Name
            .WriteAttribute "abc", Alphabetized
            .WriteAttribute "atomic", Atomic, False
            .WriteAttribute "negative", Negative, False
            .WriteAttribute "display", Display
        
            First
            Do Until Off
                
                .BeginTag "trait"
                    .WriteAttribute "name", Trait.Name
                    .WriteAttribute "val", Trait.Total, "1"
                    .WriteAttribute "note", Trait.Note
                .EndTag
                MoveNext
            Loop
    
        .EndTag
        
    End With

    Set Cursor = CurNode

End Sub

Public Sub LLInputFromFile(FileNum As Integer, Version As String)
'
' Name:         LLInputFromFile
' Parameters:   FileNum     number of an open file
'               Version     version tag of the file format
' Description:  Read a LinkedTraitList from a file, converting from the old
'               linked-list format.
'

    Dim Read As String
    Dim Total As String
    Dim Note As String
    
    Clear
    Line Input #FileNum, Read
    Do Until Read = "<-End of LinkedList Items->"
    
        If Trim(Read) <> "" Then
            Note = ""
            Total = "1"
            If Read Like "* (*)" Then
                Note = Mid(Read, InStr(Read, " (") + 2)
                Note = Left(Note, Len(Note) - 1)
                If Val(Note) <> 0 Then
                    Total = CStr(Val(Note))
                    Note = Trim(Mid(Note, Len(Total) + 2))
                    If Total = "1" Then Total = "1 "
                End If
                Read = Left(Read, InStr(Read, " (") - 1)
            End If
            AppendDirect Read, Total, Note
        End If
        Line Input #FileNum, Read
    
    Loop

End Sub

Public Sub OldInputFromFile(FileNum As Integer, Version As String)
'
' Name:         OldInputFromFile
' Parameters:   FileNum     number of an open file
'               Version     version tag of the file format
' Description:  Read a LinkedTraitList from an older Grapevine file,
'               converting from the old trait list format.
'

    Dim Read As String
    Dim ReadNum As String
    Dim ReadNote As String

    Clear
    Line Input #FileNum, Read
    Do Until Read = "<-End of LinkedTraitList Items->"
        Line Input #FileNum, ReadNum
        ReadNum = Trim(ReadNum)
        Line Input #FileNum, ReadNote
        AppendDirect Read, ReadNum, ReadNote
        Line Input #FileNum, Read
    Loop

End Sub

Public Sub InputFromFile(XML As XMLReaderClass, Version As Double)
'
' Name:         InputFromFile
' Parameters:   XML         an XMLReaderClass that just read the opening
'                           tag of this traitlist
'               Version     version of the file format
' Description:  Read a LinkedTraitList from a file.
'

    If XML.Tag = "traitlist" Then
        
        Dim TraitName As String
        
        With XML
        
            Clear
            Name = .GetString("name")
            Alphabetized = .GetBoolean("alphabetized", True) And .GetBoolean("abc", True)
            Atomic = .GetBoolean("atomic", False)
            Negative = .GetBoolean("negative", False)
            Display = .GetLong("display", ldMultiplier)
            
            Do Until .Tag = "traitlist" And .Terminal
                .ReadTag
                If .Tag = "trait" Then
                    TraitName = .GetString("name")
                    If TraitName <> "" Then AppendDirect TraitName, _
                                            .GetString("val", "1"), _
                                            .GetString("note")
                End If
            Loop
            
        End With
        
    Else
        Err.Raise vbObjectError + 8, "LinkedTraitList", _
                  "Wrong file input to a trait list"
    End If

End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this character's data to a binary file.
'

    Dim CurNode As LinkedTraitListNode

    Set CurNode = Cursor

    PutStrB FileNum, Name
    Put #FileNum, , Alphabetized
    Put #FileNum, , Atomic
    Put #FileNum, , Negative
    Put #FileNum, , Display

    Put #FileNum, , CInt(NodeCount)

    First
    Do Until Off
        PutStrB FileNum, Trait.Name
        PutStrB FileNum, Trait.Total
        PutStrB FileNum, Trait.Note
        MoveNext
    Loop

    Set Cursor = CurNode

End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this character's data from a binary file.
'

    Dim TName As String
    Dim TVal As String
    Dim TNote As String
    Dim Alpha As Boolean
    Dim I As Integer
    
    Alphabetized = False
    Clear
            
    GetStrB FileNum, Name
    Get #FileNum, , Alpha
    Get #FileNum, , Atomic
    Get #FileNum, , Negative
    Get #FileNum, , Display
            
    Get #FileNum, , I
            
    Do Until I = 0
        GetStrB FileNum, TName
        GetStrB FileNum, TVal
        GetStrB FileNum, TNote
        AppendDirect TName, TVal, TNote
        I = I - 1
    Loop
            
    SetAlphabetized Alpha
 
End Sub

Public Function OutputID() As Integer
'
' Name:         OutputID
' Description:  Return an ID that tells the OutputEngineClass how to manipulate this object
'
    OutputID = oidTraitList

End Function

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
    TraitTotal = 0
    Negative = False
    Atomic = False
    Alphabetized = True
    Display = ldMultiplier
    
End Sub
Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  If a list is destroyed, clear it first.  VB calls
'               this method automatically.
'

    Me.Clear

End Sub
