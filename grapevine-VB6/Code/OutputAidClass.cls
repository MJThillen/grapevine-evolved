VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OutputAidClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public Enum GameOutputType
    goDebug = 0
    goPrinter = 1
    goFile = 2
End Enum

Public Enum InfluenceRumorSelectType
    irsAll = 0
    irsExcludeUsed = 1
    irsIncludeOnlyUsed = 2
End Enum

Public PageWidth As Integer             'width of a plain text page
Public HalfWidth As Integer             'half of that width
Public Destination As GameOutputType    'Whether to send to file or printer
Public FileLoc As Integer               'number of an open file

'
' Section:  RTF Output
'
'           FilterRTFCharacter
'           KeyReplace
'           FilterRTFRoster
'

Public Sub FilterRTFCharacter(TemplateFile As Integer, FilterObject As Object, _
        Optional IncludeNotes As Boolean = True, Optional IncludeHistory As Boolean = True)
'
' Name:         FilterRTFCharacter
' Parameters:   TemplateFile        number of the open RTF template file
'               FilterObject        object whose KeyReplace method to use
'               IncludeNotes        whether to include the Notes field
'               IncludeHistory      whether to include the History field
' Description:  Read an RTF character sheet template file.  Replace the keywords
'               in that file with appropriate values from the given object, then
'               write the modified file to disk.
'

    Dim ProcessOn As Boolean
    Dim ClosingWord As String
    Dim Segment As String
    Dim Replacement As String
    Dim Repeat As Boolean
    Dim SegmentLen As Integer
    Dim SegmentNum As Integer

    ProcessOn = True
    SegmentNum = 1

    On Error GoTo RTFError

    Do Until EOF(TemplateFile)
    
        Segment = Segment & Input(1, TemplateFile)
        
        On Error GoTo 0
        
        If ProcessOn Then
        
            If EndsWith(Segment, kwComment) Then
                ClosingWord = kwComment
                
            ElseIf EndsWith(Segment, kwHistory) Then
                If Not IncludeHistory Then ClosingWord = kwHistory
                
            ElseIf EndsWith(Segment, kwNotes) Then
                If Not IncludeNotes Then ClosingWord = kwNotes
                
            ElseIf EndsWith(Segment, kwPageBreak) Then
                If Destination = goFile Then Segment = Left(Segment, Len(Segment) - Len(kwPageBreak))
                
            ElseIf EndsWith(Segment, kwLineMarkerSp) Or EndsWith(Segment, kwLineMarkerCR) Then
                
                FilterObject.InitializeForOutput
                
                Do
                    KeyReplace Segment, FilterObject, Replacement, Repeat
                    On Error GoTo RTFError
                    Print #FileLoc, Replacement;
                    On Error GoTo 0
                Loop While Repeat
                
                SegmentNum = SegmentNum + 1
                Segment = ""
                            
            End If
    
            If ClosingWord <> "" Then
                SegmentLen = Len(Segment) - Len(ClosingWord)
                ProcessOn = False
            End If
    
        Else
        
            If EndsWith(Segment, ClosingWord) Then
                
                Segment = Left(Segment, SegmentLen)
                ClosingWord = ""
                ProcessOn = True
            
            End If
            
        End If
    
        On Error GoTo RTFError
    
    Loop
    
    Print #FileLoc, Segment

    On Error GoTo 0
    
    GoTo RTFEnd
    
RTFError:

    MsgBox "Error reading template at line " & SegmentNum & ": " & Err.Description, _
        vbCritical, "Processing Error"
    Resume RTFEnd

RTFEnd:

End Sub

Private Sub KeyReplace(Segment As String, FilterObject As Object, Replacement As String, _
        Repeat As Boolean)
'
' Name:         KeyReplace
' Parameters:   Segment         the string whose keywords are to be replaced
'               FilterObject    the object supplying the replacement values
'               Replacement     the segment with keywords replaces
'               Repeat          whether or not to duplicate this line and repeat the
'                               replacement to get further data from the object
'                               (needed for lists)
' Description:  Replace the keywords from a given string with the values from a
'               given object.  Store the result in the Replacement string.  If
'               some keywords have been replaced by data found while traversing a
'               list, set Repeat TRUE to signal that this method should be called
'               again until the list has been completely traversed.
' Returns:      Replacement and Repeat are, in effect, return values.
'
        
    Dim Char As String * 1
    Dim Keyword As String
    Dim KeyOn As Boolean
    Dim Loc As Integer
    Dim Value As String
    Dim More As Boolean
    
    Replacement = ""
    Repeat = False
    
    For Loc = 1 To Len(Segment)
    
        Char = Mid(Segment, Loc, 1)
        
        Select Case Char
            Case "["
                KeyOn = True
                Keyword = ""
            Case "]"
                
                If KeyOn Then
                    
                    KeyOn = False
                    Value = ""
                    More = False
                    FilterObject.KeyReplace Keyword, Value, More
                    
                    Value = Replace(Value, vbCrLf, vbCrLf & kwLineMarkerSp)
                    
                    Replacement = Replacement & Value
                    Repeat = Repeat Or More
                
                End If
            
            Case Else
                If KeyOn Then
                    Keyword = Keyword & Char
                Else
                    Replacement = Replacement & Char
                End If
        End Select
    
    Next Loc
        
End Sub

Public Sub FilterRTFRoster(TemplateFile As Integer, Roster As RosterType, _
        GivenDate As String, ByVal DelimNames As String, Delim As String)
'
' Name:         FilterRTFRoster
' Parameters:   TemplateFile    number of an open template file
'               Roster          the type of roster to generate
'               GivenDate       given date
'               DelimNames      delimited list of names
'               Delim           delimiter in the list
' Description:  Read an RTF template file, replacing keywords to generate a roster
'               of the given type.  Generate the roster for only the players and
'               characters in the list of names.
'

    Dim Segment As String
    Dim MemberCount As Integer
    Dim NextMember As Boolean
    Dim SegmentNum As Integer
    Dim MemberList As LinkedList
    Dim x As Integer
    Dim RepeatSeg As String
    
    If Roster = roCharacter Or Roster = roAttendance Then
        Set MemberList = CharacterList
    Else
        Set MemberList = PlayerList
    End If
    
    Do
        MemberCount = MemberCount + 1
        x = InStr(x + 1, DelimNames, Delim)
    Loop Until x = 0
    MemberCount = MemberCount - 2

    DelimNames = Right(DelimNames, Len(DelimNames) - Len(Delim))
    MemberList.MoveTo Left(DelimNames, InStr(DelimNames, Delim) - 1)
    
    On Error GoTo MemberRosterError

    Do Until EOF(TemplateFile)
    
        Segment = Segment & Input(1, TemplateFile)
        
        On Error GoTo 0
        
        If EndsWith(Segment, kwLineMarkerSp) Or EndsWith(Segment, kwLineMarkerCR) Then
            
            Segment = Replace(Segment, kwPrintDate, Format(Now, "mmmm d, yyyy"))
            Segment = Replace(Segment, kwDate, GivenDate)
            Segment = Replace(Segment, kwMemberNum, CStr(MemberCount))
                        
            If InStr(Segment, "[" & qkName & "]") > 0 Then
            
                Do Until MemberList.Off Or DelimNames = ""
                                
                    With MemberList.Item
                        RepeatSeg = Replace(Segment, "\", "%&BSMark&%")
                        RepeatSeg = Replace(RepeatSeg, "[" & qkName & "]", .Name)
                        RepeatSeg = Replace(RepeatSeg, "[" & qkID & "]", .ID)
                        If Roster = roCharacter Or Roster = roAttendance Then
                            RepeatSeg = Replace(RepeatSeg, "[" & qkPlayer & "]", .Player)
                            RepeatSeg = Replace(RepeatSeg, "[" & qkRace & "]", .Race)
                            RepeatSeg = Replace(RepeatSeg, "[" & qkGroup & "]", .Group)
                            RepeatSeg = Replace(RepeatSeg, "[" & qkSubgroup & "]", .Subgroup)
                            RepeatSeg = Replace(RepeatSeg, "[" & qkXPEarned & "]", CStr(.Experience.Earned))
                            RepeatSeg = Replace(RepeatSeg, "[" & qkXPUnspent & "]", CStr(.Experience.Unspent))
                            RepeatSeg = Replace(RepeatSeg, "[" & qkStatus & "]", .Status)
                            RepeatSeg = Replace(RepeatSeg, "[" & qkNPC & "]", IIf(.IsNPC, "NPC", ""))
                            RepeatSeg = Replace(RepeatSeg, "[" & qkNarrator & "]", .Narrator)
                            RepeatSeg = Replace(RepeatSeg, "[" & qkStartDate & "]", .StartDate)
                        Else
                            RepeatSeg = Replace(RepeatSeg, "[" & qkPhone & "]", .Phone)
                            RepeatSeg = Replace(RepeatSeg, "[" & qkEMail & "]", .EMail)
                            RepeatSeg = Replace(RepeatSeg, "[" & qkPPEarned & "]", CStr(.Experience.Earned))
                            RepeatSeg = Replace(RepeatSeg, "[" & qkPPUnspent & "]", CStr(.Experience.Unspent))
                            RepeatSeg = Replace(RepeatSeg, "[" & qkPosition & "]", .Position)
                        End If
                        RepeatSeg = Replace(RepeatSeg, "\", "\\")
                        RepeatSeg = Replace(RepeatSeg, "%&BSMark&%", "\")
                    End With
                   
                    DelimNames = Right(DelimNames, Len(DelimNames) - (InStr(DelimNames, Delim) + _
                            (Len(Delim) - 1)))
                    If DelimNames <> "" Then _
                            MemberList.MoveTo Left(DelimNames, InStr(DelimNames, Delim) - 1)
                
                    On Error GoTo MemberRosterError
                    Print #FileLoc, RepeatSeg;
                                
                Loop
            
            Else
                            
                On Error GoTo MemberRosterError
                Print #FileLoc, Segment;
                
            End If
            
            Segment = ""
            SegmentNum = SegmentNum + 1
            
        End If

    Loop

    On Error GoTo MemberRosterError
    
    Print #FileLoc, Segment

    On Error GoTo 0
    
    GoTo MemberRosterEnd
    
MemberRosterError:

    MsgBox "Error reading template at line " & SegmentNum & ": " & Err.Description, _
        vbCritical, "Processing Error"
    Resume MemberRosterEnd

MemberRosterEnd:

End Sub

Private Function EndsWith(Main As String, Tail As String) As Boolean
'
' Name:         EndsWith
' Parameters:   Main    a string to test
'               Tail    a suffix
' Description:  Return TRUE if a string ends with the given suffix.
'               Case-insensitive.
' Returns:      TRUE if Main ends with Tail; FALSE otherwise.
'
    
    EndsWith = LCase(Right(Main, Len(Tail))) = Tail

End Function

'
' Section:  Text Output
'
'           OutputTextRoster
'           SetStandardPageWidth
'           WrapOutput
'           Output
'

Public Sub OutputTextRoster(Roster As RosterType, GivenDate As String, _
        ByVal DelimNames As String, Delim As String)
'
' Name:         OutputTextRoster
' Parameters:   the type of roster to generate
'               GivenDate       given date
'               DelimNames      delimited list of names
'               Delim           delimiter in the list
' Description:  Output a roster of the specified type for the character or
'               player names in the list.
'

    Dim MemberList As LinkedList
    Dim MemberCount As Integer
    Dim ThirdPage As Integer
    Dim x As Integer
    
    ThirdPage = PageWidth \ 3
    
    Do
        MemberCount = MemberCount + 1
        x = InStr(x + 1, DelimNames, Delim)
    Loop Until x = 0
    MemberCount = MemberCount - 2
    
    Select Case Roster
        Case roCharacter
        
            Set MemberList = CharacterList
            Output "Character Roster for " & GivenDate
            Output CStr(MemberCount) & " Characters"
            Output ""
            Output Left("Name" & String(ThirdPage, " "), ThirdPage) & _
                   Left("Player" & String(ThirdPage, " "), ThirdPage) & _
                   Left("Status" & String(ThirdPage, " "), ThirdPage)
            Output String(ThirdPage, " ") & _
                   Left("Race" & String(ThirdPage, " "), ThirdPage) & _
                   Left("Group" & String(ThirdPage, " "), ThirdPage)
            Output String(ThirdPage, " ") & _
                   Left("ID" & String(ThirdPage, " "), ThirdPage) & _
                   Left("Experience" & String(ThirdPage, " "), ThirdPage)
            Output String(PageWidth, "-")
            
        Case roAttendance
        
            Set MemberList = CharacterList
            Output "Attendance List for " & GivenDate
            Output CStr(MemberCount) & " Characters"
            Output ""
            Output "Here? Character Name", "Player"
            Output "----- --------------", "------"

        Case roPlayer
            
            Set MemberList = PlayerList
            Output "Player Roster for " & GivenDate
            Output CStr(MemberCount) & " Players"
            Output ""
            Output Left("Name" & String(ThirdPage, " "), ThirdPage) & _
                   Left("Phone" & String(ThirdPage, " "), ThirdPage) & _
                   Left("E-Mail" & String(ThirdPage, " "), ThirdPage)
            Output String(ThirdPage, " ") & _
                   Left("ID" & String(ThirdPage, " "), ThirdPage) & _
                   Left("Points" & String(ThirdPage, " "), ThirdPage)
            Output String(PageWidth, "-")
            
    End Select

    DelimNames = Right(DelimNames, Len(DelimNames) - Len(Delim))

    For x = 1 To MemberCount

        MemberList.MoveTo Left(DelimNames, InStr(DelimNames, Delim) - 1)

        If Not MemberList.Off Then
            
            With MemberList.Item
            Select Case Roster
                Case roCharacter
                
                    Output Left(.Name & String(ThirdPage, " "), ThirdPage) & _
                           Left(.Player & String(ThirdPage, " "), ThirdPage) & _
                           Left(.Status & String(ThirdPage, " "), ThirdPage)
                    Output String(ThirdPage, " ") & _
                           Left(.Race & String(ThirdPage, " "), ThirdPage) & _
                           Left(.Group & String(ThirdPage, " "), ThirdPage)
                    Output String(ThirdPage, " ") & _
                           Left("ID: " & .ID & String(ThirdPage, " "), ThirdPage) & _
                           Left(CStr(.Experience) & " Exp." & String(ThirdPage, " "), ThirdPage)
                    
                Case roAttendance
                
                    Output "_____ " & .Name, .Player
        
                Case roPlayer
                    
                    Output Left(.Name & String(ThirdPage, " "), ThirdPage) & _
                           Left(.Phone & String(ThirdPage, " "), ThirdPage) & _
                           Left(.EMail & String(ThirdPage, " "), ThirdPage)
                    Output String(ThirdPage, " ") & _
                           Left("ID: " & .ID & String(ThirdPage, " "), ThirdPage) & _
                           Left(CStr(.Points) & " Pts." & String(ThirdPage, " "), ThirdPage)

            End Select
            End With
            
        End If

        DelimNames = Right(DelimNames, Len(DelimNames) - (InStr(DelimNames, Delim) + _
                (Len(Delim) - 1)))

    Next x
    
End Sub

Public Sub SetStandardPageWidth(Optional Cancel As Boolean = False)
'
' Name:         SetStandardPageWidth
' Parameters:   Cancel      whether to cancel the operation
' Description:  Set the standard page width to that offered by the printer object.
'               If it is very small, offer the user the option to cancel.
' Returns:      Whether or not to cancel
'

    Cancel = False
    If Destination = goFile Then
        PageWidth = 70
    Else
        PageWidth = Printer.ScaleWidth \ Printer.TextWidth("_")
        If PageWidth <= 50 Then
            If MsgBox("Wow, you've chosen a really small paper size--only " & CStr(PageWidth) & _
                      "characters wide.  The printout may look funny.  You sure about this?", _
                      vbYesNo + vbQuestion, "Puny Paper") = vbNo Then Cancel = True
        End If
    End If
    HalfWidth = PageWidth \ 2

End Sub

Public Sub WrapOutput(ByVal Out As String)
'
' Name:         WrapOutput
' Parameters:   Out     the string output to wrap
' Description:  Wrap the string to the page width, and send it to output.
'

    Dim BlankLoc As Integer
    Dim OutLeft As String
    Dim RightChar As String * 1

    Do
        '
        ' Strip newlines off the front
        '
        Do While Left(Out, 2) = vbCrLf
            Out = Right(Out, Len(Out) - 2)
        Loop

        '
        ' Grab the string up to the first newline
        '
        BlankLoc = InStr(Out, vbCrLf)
        If BlankLoc = 0 Then
            OutLeft = Out
        Else
            OutLeft = Left(Out, BlankLoc - 1)
        End If

        Out = Right(Out, Len(Out) - Len(OutLeft))

        '
        ' Wrap that line, outputting it
        '
        Do

            BlankLoc = PageWidth
            RightChar = Mid(OutLeft, BlankLoc, 1)
            Do Until RightChar = "" Or RightChar = " "
                BlankLoc = BlankLoc - 1
                If BlankLoc = 0 Then Exit Do
                RightChar = Mid(OutLeft, BlankLoc, 1)
            Loop

            If BlankLoc = 0 Then BlankLoc = PageWidth
            If BlankLoc > Len(OutLeft) Then BlankLoc = Len(OutLeft)

            Output Left(OutLeft, BlankLoc)

            OutLeft = Right(OutLeft, Len(OutLeft) - BlankLoc)

        Loop While Len(OutLeft) > 0

        If Left(Out, 4) = vbCrLf & vbCrLf Then Output ""

    Loop While Len(Out) > 0

    Output ""

End Sub

Public Sub Output(ByVal Out As String, Optional Out2 As String = "")
'
' Name:         Output
' Parameters:   Out     a string to send to printer or file
'               Out2    a second string to send
' Description:  Based on the Destination, send the string to printer or file.
'               If Out2 is specified, give the strings a two-column format.
'

    If Out2 <> "" Then
        Out = Left(Out & Space(HalfWidth), HalfWidth) & Left(Out2 & Space(HalfWidth), HalfWidth)
    End If

    Select Case Destination
        Case goPrinter
            Printer.Print Out
        Case goFile
            Print #FileLoc, Out
        Case goDebug
            Debug.Print Out
    End Select
    
End Sub
