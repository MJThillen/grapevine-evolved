VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OutputEngineClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum OutputDeviceType
    odFile = 0
    odPrinter = 1
    odEMail = 2
End Enum

Public Enum OutputOptionType
    
    ooFileError = -1
    
    ooUnknown = &H0
    ooReport = &H1
    ooPlayers = &H2
    ooCharacters = &H4
    ooItems = &H8
    ooRotes = &H10
    ooLocations = &H20
    ooActions = &H40
    ooPlots = &H80
    ooRumors = &H100
    ooSearch = &H200
    ooStatistics = &H400

    ooOptNotes = &H800
    ooOptHistory = &H1000
    ooOptitems = &H2000
    ooOptRotes = &H4000
    ooOptLocations = &H10000     'Can't use &H8000!
    ooOptActions = &H20000
    ooOptRumors = &H40000
    
    ooNone = &H0
    ooSubjectMask = &H7FF
    ooOptionMask = &H77800
       
End Enum

Public Template As String

Public SearchName As String
Public SearchNot As Boolean
Public StatType As StatisticType
Public StatKey As String
Public StatTrait As String
Public OKZero As Boolean

Public GameDate As Date
Public AscendDate As Boolean
Public StartDate As Date
Public EndDate As Date

Public HideSTData As Boolean
Public AllowHTML As Boolean

Public OutputFormat As OutputFormatType
Public OutputDevice As OutputDeviceType
Public Copies As Integer
Public OutputOptions As OutputOptionType
Public PrintAfterSave As Boolean
Public AlwaysAttach As Boolean

Public CDialog As CommonDialog
Public RTFBox As RichTextBox

' Include the SMTP.ocx control in the project in order to re-enable the code
' below. Note that SMTP.ocx is detected (inaccurately) as a virus component
' by many antivirus programs.

'Public Mailer As SMTP
Public ReplyTo As String
Public SendTo As String
Public CC As String
Public BCC As String
Public MessageSubject As String
Public MessageHeader As String

Public OutputTraits As Boolean
Public Display As ListDisplayType
Public Dot As String
Public DotSubstitute As String

Public ErrorFlag As Boolean
Public ErrorMessage As String

Private TempDot As String
Private EmptyDot As String
Private Divider As String
Private Wrap As Long
Private DateFormat As String
Private StatScale As Double
Private AdvanceList As LinkedList
Private MatchValue As Variant
Private SortValue As Variant
Private CharName As String

Private Const M_LEFT = 0
Private Const M_RIGHT = 1
Private Const M_TOP = 2
Private Const M_BOTTOM = 3

Private Margin(0 To 3) As Long
Private Landscape As Boolean
Private PrintInProgress As Boolean
Private HTMLFileNum As Integer
Private HTMLTempFile As String

Private FinishingTouches As StringSet
Private SelectionSet(osMin To osMax) As StringSet

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Prepare this class to do its work.
'

    Dim I As Integer

    SearchName = "Active Characters"
    SearchNot = False
    StatType = stDistribution
    StatKey = qkInfluences
    StatTrait = "(specific trait)"
    OKZero = False
    
    Set AdvanceList = New LinkedList
    Set FinishingTouches = New StringSet
    
    For I = osMin To osMax
        Set SelectionSet(I) = New StringSet
    Next I
    
    AdvanceList.Clear
    FinishingTouches.Clear
    
    Dot = "O"
    TempDot = "Õ"
    EmptyDot = "Ø"
    DotSubstitute = "$~DS~$"
    Divider = ""
    Display = ldDefault
    Wrap = 0
    DateFormat = "mmmm d, yyyy"
    AscendDate = True
    StartDate = 0
    EndDate = 0
    HideSTData = True
    AllowHTML = False
    PrintAfterSave = False
    SearchName = "All Characters"
    StatType = stDistribution
    StatKey = qkGroup
    StatScale = 100
    Copies = 1
    Landscape = False
    Margin(M_LEFT) = 1440
    Margin(M_RIGHT) = 1440
    Margin(M_TOP) = 1440
    Margin(M_BOTTOM) = 1440
    
End Sub

Public Sub InitializeMessage(OutOpts As Long, Subject As String)
'
' Name:         InitializeMessage
' Parameters:   OutOpts         Output Options -- the contents we will send
'               Subject         Subject line
' Description:  Set up the mailer for use with new output.
'

    Dim I As Long
    Dim MultiMessage As Boolean
        
    MultiMessage = ooReport And Not OutOpts
    If MultiMessage Then MultiMessage = OutOpts And (ooCharacters Or ooPlayers)
    If MultiMessage Then Subject = "[Name] " & Subject

' Include the SMTP.ocx control in the project in order to re-enable the code
' below. Note that SMTP.ocx is detected (inaccurately) as a virus component
' by many antivirus programs.

'    Mailer.MessageHTML = ""
'    Mailer.MessageText = ""
'    BCC = ""
'    CC = ""
'    SendTo = ""
'    MessageHeader = ""
'    MessageSubject = Subject
'
'    Do Until Mailer.Attachments.Count = 0
'        Mailer.Attachments.Remove 1
'    Loop

    If MultiMessage Then SendTo = SendToSelect
    
End Sub

Public Sub ClearSelections()
'
' Name:         ClearSelections
' Description:  Clear all the selection sets of the output engine.
'

    Dim I As Long
    
    For I = osMin To osMax
        SelectionSet(I).Clear
    Next I
    
End Sub

Public Function SelectSet(SetIndex As Long) As StringSet
'
' Name:         SelectSet
' Parameter:    SetIndex      the index of the set to fetch
' Description:  Access one of the selection sets of the output engine.
'
    Set SelectSet = SelectionSet(SetIndex)

End Function

Public Function GetTemplateSubject(FileName As String) As Long
'
' Name:         GetTemplateSubject
' Parameters:   Filename        file to open and examine
' Description:  Read a file and determine what kind of template it is from its content.
'
    
    Dim FileNum As Integer
    Dim Read As String
    Dim StartLoc As Long
    Dim EndLoc As Long
    Dim Keyword As String
    Dim Param As String
    Dim AddVal As Long
    Dim NullStats As Boolean
    
    GetTemplateSubject = ooUnknown
            
    FileNum = FreeFile
    Open FileName For Input As #FileNum
    Do Until EOF(FileNum)
        
        Line Input #FileNum, Read
        StartLoc = 1
        Do While GetNextKeyword(Read, Keyword, Param, StartLoc, EndLoc)
            
            Param = LCase(Param)
            AddVal = 0
            
            Select Case Keyword
                Case "report", "template"
                    
                    If Keyword = "template" Then
                        GetTemplateSubject = GetTemplateSubject And Not ooSubjectMask
                    End If
                    
                    If Keyword = "report" Or Param = "report" Then
                        GetTemplateSubject = GetTemplateSubject Or ooReport
                    End If
                    
                    Select Case Left(Param, 3)
                        Case "cha": AddVal = ooCharacters
                        Case "pla": AddVal = ooPlayers
                        Case "ite": AddVal = ooItems
                        Case "rot": AddVal = ooRotes
                        Case "loc": AddVal = ooLocations
                        Case "act": AddVal = ooActions
                        Case "plo": AddVal = ooPlots
                        Case "rum": AddVal = ooRumors
                        Case "sea": AddVal = ooSearch
                        Case "sta": AddVal = ooStatistics
                    End Select
                
                Case "option"
            
                    AddVal = 0
                    Select Case Left(Param, 3)
                        Case "not": AddVal = ooOptNotes
                        Case "his": AddVal = ooOptHistory
                        Case "ite": AddVal = ooOptitems
                        Case "rot": AddVal = ooOptRotes
                        Case "loc": AddVal = ooOptLocations
                        Case "act": AddVal = ooOptActions
                        Case "rum": AddVal = ooOptRumors
                    End Select
            
                Case "stattype", "statfield"
                    NullStats = True
            
            End Select
            
            GetTemplateSubject = GetTemplateSubject Or AddVal
            StartLoc = StartLoc + 1
        
        Loop
        
        If NullStats Then GetTemplateSubject = (GetTemplateSubject And Not ooStatistics) Or ooSearch
        
    Loop
    Close #FileNum

End Function

Private Function GetNextKeyword(ByRef FromString As String, ByRef Keyword As String, ByRef Param As String, _
                               ByRef StartLoc As Long, ByRef EndLoc As Long) As Boolean
'
' Name:         GetNextKeyword
' Parameters:   FromString          string to get the keyword from
'               Keyword             place to store the keyword
'               Param               place to store its parameters
'               StartLoc            passed as the place to start looking; returned as the place the keyword
'                                   was found.
'               EndLoc              The end of the found keyword.
' Description:  Find the next instance of a keyword [*] in a string.  Return the keyword and its parameters.
'

    Dim SpaceLoc As Long
    
    GetNextKeyword = False

    Keyword = ""
    Param = ""

    StartLoc = InStr(StartLoc, FromString, "[")
    If StartLoc > 0 Then
        
        EndLoc = InStr(StartLoc, FromString, "]")
        If EndLoc > 0 Then
        
            Keyword = Mid(FromString, StartLoc + 1, EndLoc - StartLoc - 1)
            SpaceLoc = InStr(Keyword, " ")
            If SpaceLoc > 0 Then
                Param = Mid(Keyword, SpaceLoc + 1)
                Keyword = Left(Keyword, SpaceLoc - 1)
                If Len(Param) > 1 And Left(Param, 1) = Right(Param, 1) Then
                    If Left(Param, 1) = """" Or Left(Param, 1) = "'" Then
                        Param = Mid(Param, 2, Len(Param) - 2)
                    End If
                End If
            End If
            Keyword = LCase(Keyword)
            GetNextKeyword = True
            
        End If
        
    End If
    
End Function

Private Function FormatControl(ByRef ControlKey As String, ByRef ControlVal As String) As Boolean
'
' Name:         FormatControl
' Parameters:   ControlKey      text with the control command and its parameters
'               ControlVal      value to use with the command
' Description:  If the text given is a control key, apply it and return TRUE
'

    FormatControl = True
    Select Case ControlKey
        Case "dot"
            Dot = ControlVal
        Case "emptydot"
            EmptyDot = ControlVal
        Case "tempdot"
            TempDot = ControlVal
        Case "display"
            If LCase(ControlVal) = "default" Or ControlVal = "" Then
                Display = ldDefault
            Else
                Display = Val(ControlVal)
            End If
        Case "divider"
            Divider = ControlVal
        Case "dateformat"
            DateFormat = ControlVal
        Case "search"
            SearchName = ControlVal
        Case "stattype"
            Select Case LCase(ControlVal)
                Case "distinct":        StatType = stDistinctDistribution
                Case "distribution":    StatType = stDistribution
                Case "maxima", "max":   StatType = stMaxima
                Case "specific":        StatType = stSpecificDistribution
                Case "sums", "sum":     StatType = stSums
            End Select
        Case "statfield"
            StatKey = LCase(ControlVal)
        Case "statmatch"
            StatTrait = ControlVal
        Case "statzero"
            OKZero = CBool(ControlVal)
        Case "statscale"
            StatScale = Val(ControlVal)
        Case "landscape"
            Landscape = True
        Case "marginleft"
            Margin(M_LEFT) = Int(Val(ControlVal) * 1440)
        Case "marginright"
            Margin(M_RIGHT) = Int(Val(ControlVal) * 1440)
        Case "margintop"
            Margin(M_TOP) = Int(Val(ControlVal) * 1440)
        Case "marginbottom"
            Margin(M_BOTTOM) = Int(Val(ControlVal) * 1440)
            
        Case "col", "wrap", "tab", "template", "/?", "/!" 'do nothing with these... yet
        
        Case Else
            FormatControl = False
    End Select

End Function

Private Function MakeTally(Perm As Long, Temp As Long) As String
'
' Name:         MakeTally
' Parameters:   Perm        Permanent temper rating
'               Temp        Temporary temper rating
' Description:  Create a string of permanent vs. temporary ratings, using the current dots and dividers
'

    Dim I As Long
    
    For I = 1 To Perm
        If Not MakeTally = "" Then MakeTally = MakeTally & Divider
        MakeTally = MakeTally & IIf(I <= Temp, Dot, EmptyDot)
    Next I
    
    For I = Perm + 1 To Temp
        If Not MakeTally = "" Then MakeTally = MakeTally & Divider
        MakeTally = MakeTally & TempDot
    Next

End Function

Private Function FormatForOutput(ByRef Text As String) As String
'
' Name:             FormatForOutput
' Parameters:       Text        String to Format
' Description:      Do the text formatting called for by this output format
'

    FormatForOutput = Text

    If HideSTData Then FormatForOutput = Game.STFilter(FormatForOutput)
    
    Select Case OutputFormat
        Case ofRTF
            FormatForOutput = Replace(FormatForOutput, "\", "\\")
            FormatForOutput = Replace(FormatForOutput, vbCrLf, vbCrLf & "\par ")
        Case ofHTML
            If Not AllowHTML Then
                FormatForOutput = Replace(FormatForOutput, "&", "&amp;")
                FormatForOutput = Replace(FormatForOutput, "<", "&lt;")
                FormatForOutput = Replace(FormatForOutput, ">", "&gt;")
                FormatForOutput = Replace(FormatForOutput, """", "&quot;")
                FormatForOutput = Replace(FormatForOutput, "@", "&#064;")
            End If
            FormatForOutput = Replace(FormatForOutput, vbCrLf, "<BR>")
    End Select
    
End Function

Private Function AdvanceObjects() As Boolean
'
' Name:         AdvanceObjects
' Description:  Advance all the objects in the AdvanceList according to their type;
'               then clear the list.
' Return:       TRUE iff there's more left to grab from the objects.
'

    Dim Checkdate As Date
    Dim OutputID As Long
    
    AdvanceObjects = False

    With AdvanceList
        
        .First
        Do Until .Off
        
            OutputID = .Item.OutputID
        
            Select Case OutputID
                
                Case oidTraitList
                    .Item.MoveNext
                    AdvanceObjects = AdvanceObjects Or (Not .Item.Off)
                
                Case oidHistory, oidPlot, oidCalendar
                    If AscendDate Then
                        .Item.MoveNext
                    Else
                        .Item.MovePrevious
                    End If
                    If Not .Item.Off Then
                        Select Case .Item.OutputID
                            Case oidHistory:    Checkdate = .Item.EntryDate
                            Case oidPlot:       Checkdate = .Item.PlotDev.DevDate
                            Case oidCalendar:   Checkdate = .Item.GetGameDate
                        End Select
                        If AscendDate Then
                            AdvanceObjects = AdvanceObjects Or Not (EndDate > 0 And _
                                    Checkdate > (EndDate + #11:59:59 PM#))
                        Else
                            AdvanceObjects = AdvanceObjects Or Not (StartDate > 0 And _
                                             Checkdate < StartDate)
                        End If
                    End If
                    
                Case oidAction
                    .Item.MoveNext
                    Do Until .Item.Off
                        If Not (.Item.SubAction.Action = "" And .Item.SubAction.Result = "") Then Exit Do
                        .Item.MoveNext
                    Loop
                    AdvanceObjects = AdvanceObjects Or (Not .Item.Off)
                
                Case oidRumor
                    .Item.MoveNext
                    Do Until .Item.Off
                        If Not .Item.SubRumor.Rumor = "" Then Exit Do
                        .Item.MoveNext
                    Loop
                    AdvanceObjects = AdvanceObjects Or (Not .Item.Off)
                
            End Select
            AdvanceList.Remove
        Loop
        
    End With
    
End Function

Private Sub WrapString(ByRef S As String)
'
' Name:         WrapString
' Parameters:   S       the string to wrap
' Description:  Wrap the string to the width of the Wrap variable by
'               inserting CR+LF at approriate intervals
'

    If Wrap > 0 Then

        Dim Loc As Long
        Dim EndLine As Long
        Dim SpaceLoc As Long
        
        Loc = 1
        
        Do Until Loc + Wrap > Len(S)
            
            EndLine = InStr(Loc, S, vbCrLf)
            If EndLine = 0 Then EndLine = Len(S)
            
            If (EndLine - Loc) > Wrap Then
            
                SpaceLoc = InStrRev(S, " ", Loc + Wrap - 1)
                
                If SpaceLoc > 0 Then
                    S = Left(S, SpaceLoc - 1) & vbCrLf & Mid(S, SpaceLoc + 1)
                    Loc = SpaceLoc + 2
                Else
                    S = Left(S, Loc + Wrap) & vbCrLf & Mid(S, Loc + Wrap + 1)
                    Loc = Loc + Wrap + 2
                End If
            
            Else
                Loc = EndLine + 2
            End If
            
        Loop
        
    End If

End Sub

Private Sub CompileSubsection(ByRef Section As String, Entity As Object)
'
' Name:         CompileSubsection
' Parameters:   Section         String whose keywords will be replaced
'               Entity          Object that will supply the data
' Description:  Locate each keyword in a string, then replace that keyword with a
'               value retrieved from a data object.
'

    Dim EditStart As Long
    Dim EditEnd As Long
    Dim Keyword As String
    Dim Param As String
    Dim ObjData As Variant
    Dim NewValue As String
    Dim Count As Boolean
    Dim Advance As Boolean
    Dim Tally As Boolean
    Dim Inadvance As Boolean
    Dim TraitList As LinkedTraitList
    Dim Comp As Boolean
    Dim CompNot As Boolean
    Dim OldDisplay As ListDisplayType
    Dim IfSkip As Long
    Dim EntityAppended As Boolean
    
    EditStart = 1                                                           'Initialize values
    Tally = False
    IfSkip = 0
    OutputTraits = True
    EntityAppended = False
    
    Do While GetNextKeyword(Section, Keyword, Param, EditStart, EditEnd)    'While keywords remain...
    
        If Not FormatControl(Keyword, Param) Then                           'If this keyword isn't formatting...
        
            Comp = (Left(Keyword, 1) = "?")                                 'Recognize ? (If-Skip beginning)
            CompNot = (Left(Keyword, 1) = "!")                              'Recognize ! (If-Skip beginning)
            
            If Comp Or CompNot Then Keyword = Mid(Keyword, 2)
            
            Count = (Left(Keyword, 1) = "#")                                'Recognize # (List Count Shortcut)
            Advance = (Left(Keyword, 1) = "+")                              'Recognize + (Advance Shortcut)
            
            If Count Or Advance Then Keyword = Mid(Keyword, 2)
            
            If Keyword = "tally" Then                                       'Recognize the [tally x y] structure
                Dim TalSpc As Long
                Param = LCase(Param)
                TalSpc = InStr(Param, " ")
                If TalSpc > 0 Then
                    Keyword = Left(Param, TalSpc - 1)
                    Param = Mid(Param, TalSpc + 1)
                Else
                    Keyword = Param
                    Param = ""
                End If
                Tally = True
            End If
            
            ObjData = Null                                                  'Initialize the object data
            
            Select Case Keyword                                             'recognize special keywords
                Case "printdate":               ObjData = Now
                Case "gamedate":                ObjData = GameDate
                Case "matchvalue":              ObjData = MatchValue
                Case "sortvalue":               ObjData = SortValue
                Case "charname":                ObjData = CharName
                Case "gametitle":               ObjData = Game.ChronicleTitle
                Case "daterange"
                    If StartDate = 0 And EndDate = 0 Then
                        ObjData = "All Dates"
                    ElseIf StartDate = 0 Then
                        ObjData = "After " & Format(StartDate, DateFormat)
                    ElseIf EndDate = 0 Then
                        ObjData = "Before " & Format(EndDate, DateFormat)
                    ElseIf StartDate = EndDate Then
                        ObjData = StartDate
                    Else
                        ObjData = Format(StartDate, DateFormat) & " to " & Format(EndDate, DateFormat)
                    End If
                Case Else
                    If Not (Entity Is Nothing) Then                         'if we have an object to ask,
                        Entity.GetValue Keyword, ObjData                    'ask the object for its value.
                    End If
            End Select
            
            If Not IsNull(ObjData) Then                                     'if a value was found...

                If Not (IsObject(ObjData) Or EntityAppended Or Entity Is Nothing) Then
                    If Entity.OutputID <> oidNone Then                      'Assume we have cause to advance
                        AdvanceList.Append Entity                           '  this object if it gives us non-
                        EntityAppended = True                               '  traitlist data
                    End If
                End If

                If IsObject(ObjData) Then                                   'if it's an object itself,
                    Set TraitList = ObjData                                 'then it's a trait list.
                    If Advance Then TraitList.MoveNext                      'Advance the trait list if asked
                    If Count Then                                           'if you want to count it, do so
                        NewValue = CStr(TraitList.Count)
                    ElseIf Not TraitList.Off Then                           'else if it's not out of data...
                        OldDisplay = Display                                'store the display parameter
                        If Param <> "" Then Display = Int(Val(Param))
                        Inadvance = False                                   'put it in the advance list if
                        AdvanceList.First                                   '  it's not there already
                        Do Until AdvanceList.Off
                            Inadvance = Inadvance Or (AdvanceList.Item Is TraitList)
                            AdvanceList.MoveNext
                        Loop
                        If Not Inadvance Then AdvanceList.Append TraitList
                        NewValue = FormatForOutput(TraitList.DisplayTrait)  'get its trait
                        NewValue = Replace(NewValue, DotSubstitute, Dot)
                        Display = OldDisplay                                'reset the display
                    Else
                        NewValue = ""                                       'return "" if the list is off
                    End If
                
                ElseIf TypeName(ObjData) = "Date" Then                      'if it's a date, format it
                    If ObjData = 0 Then
                        NewValue = ""
                    Else
                        NewValue = FormatForOutput(Format(ObjData, DateFormat))
                    End If
                    
                ElseIf Tally And IsNumeric(ObjData) Then                    'if it's numeric and a tally,
                    If Param = "" Or Entity Is Nothing Then                 'create the tally
                        NewValue = MakeTally(CLng(ObjData), CLng(ObjData))
                    Else
                        NewValue = CStr(ObjData)
                        Entity.GetValue Param, ObjData
                        If IsNumeric(ObjData) Then NewValue = _
                                MakeTally(CLng(NewValue), CLng(ObjData))
                    End If
                    Tally = False
                    
                Else                                                        'if else, call it a string value
                    NewValue = FormatForOutput(CStr(ObjData))
                End If
                
                If Comp Or CompNot Then                                     'perform an if-skip comparison
                    Comp = (StrComp(NewValue, Param, vbTextCompare) <> 0) Xor CompNot
                    If Comp Then IfSkip = EditStart                         'if Newvalue and param match [?]
                    NewValue = ""                                           'or don't match [!], begin an if-skip
                End If
                
                Section = Left(Section, EditStart - 1) & NewValue & _
                          Mid(Section, EditEnd + 1)                         'add it into the section and proceed
                EditStart = EditStart + Len(NewValue)
            
            Else                                                            '...if a value wasn't found...
                
                EditStart = EditEnd + 1                                     'just move on
                
            End If
        
        Else                                                                'if this was a formatting command,
                                                                            ' or we're if-skipping...
            If Keyword = "col" Or Keyword = "wrap" Then                     'ignore it if it is col or wrap
                EditStart = EditEnd + 1
            ElseIf (Keyword = "/?" Or Keyword = "/!") And IfSkip > 0 Then   'perform an if-skip if needed
                Section = Left(Section, IfSkip - 1) & Mid(Section, EditEnd + 1)
                EditStart = IfSkip
                IfSkip = 0
            ElseIf Keyword = "tab" Then                                     'replace tabs
                Select Case OutputFormat
                    Case ofRTF:     ObjData = "\tab "
                    Case Else:      ObjData = vbTab
                End Select
                Section = Left(Section, EditStart - 1) & ObjData & Mid(Section, EditEnd + 1)
            Else                                                            'erase it otherwise
                Section = Left(Section, EditStart - 1) & Mid(Section, EditEnd + 1)
            End If
            
        End If
        
    Loop                                                                    'move to the next keyword

    OutputTraits = False

End Sub

Private Sub CompileEntitySection(ByRef Section As String, ByRef Entity As Object)
'
' Name:         CompileEntitySection
' Parameters:   Section             the string section to process
'               Entity              the entity from which to pull data for the subsection
' Description:  Populate a string with all the appropriate values from its entity.
'

    Dim EditStart As Long
    Dim RepeatStart As Long
    Dim RepeatEnd As Long
    Dim FirstSection As String
    Dim Subsection As String
    Dim DoneSection As String
    
    EditStart = 1
    RepeatStart = InStr(1, Section, "[repeat]", vbTextCompare)
    RepeatEnd = InStr(RepeatStart + 1, Section, "[/repeat]", vbTextCompare)
    
    Do Until RepeatStart = 0 Or RepeatEnd = 0
    
        If Not (Entity Is Nothing) Then Entity.InitializeForOutput
    
        FirstSection = Mid(Section, EditStart, RepeatStart - EditStart)
        CompileSubsection FirstSection, Entity
        AdvanceList.Clear
        
        DoneSection = ""
        Do
            If Not DoneSection = "" Then DoneSection = DoneSection & Divider
            Subsection = Mid(Section, RepeatStart + 8, RepeatEnd - RepeatStart - 8)
            CompileSubsection Subsection, Entity
            DoneSection = DoneSection & Subsection
        Loop While AdvanceObjects

        Section = Left(Section, EditStart - 1) & FirstSection & DoneSection & _
                  Mid(Section, RepeatEnd + 9)

        EditStart = EditStart - 1 + Len(FirstSection) + Len(DoneSection)
        RepeatStart = InStr(EditStart, Section, "[repeat]", vbTextCompare)
        RepeatEnd = InStr(RepeatStart + 1, Section, "[/repeat]", vbTextCompare)

    Loop

    If Not (Entity Is Nothing) Then Entity.InitializeForOutput
    FirstSection = Mid(Section, EditStart)
    CompileSubsection FirstSection, Entity
    AdvanceList.Clear
    Section = Left(Section, EditStart - 1) & FirstSection

End Sub

Private Sub CompileReportSection(ByRef Section As String, Source As LinkedList, UseCharLogic As Boolean)
'
' Name:         CompileReport
' Parameters:   Section              the string section to repeat into a report
'               Source               the entities to report
'               UseCharLogic         whether to compile each entity as a character or as a generic entity
' Description:  Fill out a report with members of the source list.
'

    Dim Blueprint As String
    Dim DoneSection As String
    
    Blueprint = Section
    
    With Source
        .First
        Do Until .Off
            If UseCharLogic Then
                CompileCharacterSection Section, .Item
            Else
                CompileEntitySection Section, .Item
            End If
            If Not DoneSection = "" Then DoneSection = DoneSection & Divider
            DoneSection = DoneSection & Section
            Section = Blueprint
            .MoveNext
        Loop
    End With

    Section = DoneSection

End Sub

Private Sub CompileCharacterSection(ByRef Section As String, ByRef Character As Object)
'
' Name:         CompileCharacterSection
' Parameters:   Section             the string section to process
'               Character           the character from which to pull data for the section
' Description:  Populate a string with all the appropriate values from the options of a character.
'

    Dim EditStart As Long
    Dim OptionStart As Long
    Dim OptionEnd As Long
    Dim OptionType As String
    Dim FirstSection As String
    Dim Subsection As String
    Dim OptionSection As String
    Dim SourceList As LinkedList
    Dim Count As Integer
    Dim TraitList As LinkedTraitList
    Dim DoneSection As String
    Dim IsPlayer As Boolean
    Dim IsHas As Boolean
    
    EditStart = 1
    OptionStart = InStr(1, Section, "[option ", vbTextCompare)
    OptionEnd = InStr(OptionStart + 1, Section, "[/option]", vbTextCompare)
    Set SourceList = New LinkedList
    If Not Character Is Nothing Then
        IsPlayer = (Character.Race = "Player")
        CharName = Character.Name
    End If
    
    Do Until OptionStart = 0 Or OptionEnd = 0
    
        FirstSection = Mid(Section, EditStart, OptionStart - EditStart)
        CompileEntitySection FirstSection, Character
    
        OptionType = Mid(Section, OptionStart, InStr(OptionStart, Section, "]") - OptionStart)
        OptionType = LCase(Mid(OptionType, 9))
        
        OptionSection = Mid(Section, OptionStart, OptionEnd - OptionStart)
        OptionSection = Mid(OptionSection, 10 + Len(OptionType))
        
        DoneSection = ""
        SourceList.Clear
        
        If Not Character Is Nothing Then
        
            IsHas = (Left(OptionType, 4) = "has ")
            If IsHas Then OptionType = Mid(OptionType, 5)
            
            Select Case Left(OptionType, 3)
                Case "gam"
                    CompileEntitySection OptionSection, Game
                    DoneSection = OptionSection
                Case "not"
                    If OutputOptions And ooOptNotes Then
                        CompileEntitySection OptionSection, Character
                        DoneSection = OptionSection
                    End If
                Case "his"
                    If OutputOptions And ooOptHistory Then
                        CompileEntitySection OptionSection, IIf(IsHas, Character, Character.Experience)
                        DoneSection = OptionSection
                    End If
            End Select
            
            If Not IsPlayer And DoneSection = "" Then
                Select Case Left(OptionType, 3)
                    Case "pla"
                        PlayerList.MoveTo Character.Player
                        If Not PlayerList.Off Then
                            CompileEntitySection OptionSection, PlayerList.Item
                            DoneSection = OptionSection
                        End If
                    Case "ite"
                        If OutputOptions And ooOptitems Then
                            Set TraitList = Character.EquipmentList
                            With TraitList
                                .First
                                Do Until .Off
                                    ItemList.MoveTo .Trait.Name
                                    If Not ItemList.Off Then
                                        For Count = 1 To .Trait.Number
                                            SourceList.Append ItemList.Item
                                        Next Count
                                    End If
                                    .MoveNext
                                Loop
                            End With
                            If Not SourceList.IsEmpty Then
                                If IsHas Then
                                    CompileEntitySection OptionSection, Character
                                Else
                                    CompileReportSection OptionSection, SourceList, False
                                End If
                                DoneSection = OptionSection
                            End If
                        End If
                    Case "boo"
                        If Character.RaceCode = gvRaceVampire Then
                            With Character.BoonList
                                .First
                                Do Until .Off
                                    SourceList.Append .Item
                                    .MoveNext
                                Loop
                            End With
                            If Not SourceList.IsEmpty Then
                                If IsHas Then
                                    CompileEntitySection OptionSection, Character
                                Else
                                    CompileReportSection OptionSection, SourceList, False
                                End If
                                DoneSection = OptionSection
                            End If
                        End If
                    Case "rot"
                        If OutputOptions And ooOptRotes And Character.RaceCode = gvracemage Then
                            Character.RoteList.First
                            Set TraitList = Character.RoteList
                            With TraitList
                                .First
                                Do Until .Off
                                    RoteList.MoveTo .Trait.Name
                                    If Not RoteList.Off Then SourceList.Append RoteList.Item
                                    .MoveNext
                                Loop
                            End With
                            If Not SourceList.IsEmpty Then
                                If IsHas Then
                                    CompileEntitySection OptionSection, Character
                                Else
                                    CompileReportSection OptionSection, SourceList, False
                                End If
                                DoneSection = OptionSection
                            End If
                        End If
                    Case "loc"
                        If OutputOptions And ooOptLocations Then
                            Character.HangoutList.First
                            Set TraitList = Character.HangoutList
                            With TraitList
                                .First
                                Do Until .Off
                                    LocationList.MoveTo .Trait.Name
                                    If Not LocationList.Off Then SourceList.Append LocationList.Item
                                    .MoveNext
                                Loop
                            End With
                            If Not SourceList.IsEmpty Then
                                If IsHas Then
                                    CompileEntitySection OptionSection, Character
                                Else
                                    CompileReportSection OptionSection, SourceList, False
                                End If
                                DoneSection = OptionSection
                            End If
                        End If
                    Case "act"
                        If OutputOptions And ooOptActions Then
                            Game.APREngine.MoveToPair ActionList, GameDate, Character.Name
                            If Not ActionList.Off Then
                                If IsHas Then
                                    CompileEntitySection OptionSection, Character
                                Else
                                    CompileEntitySection OptionSection, ActionList.Item
                                End If
                                DoneSection = OptionSection
                            End If
                        End If
                    Case "rum"
                        If OutputOptions And ooOptRumors Then
                            Dim Rumor As RumorClass
                            RumorList.First
                            Do Until RumorList.Off
                                Set Rumor = RumorList.Item
                                With Rumor
                                    If .RumorDate = GameDate Then
                                        .First
                                        Do Until .Off
                                            If Not (.SubRumor.Rumor = "") And _
                                               .SubRumor.RecipientSet.Has(Character.Name) Then
                                                    SourceList.Append .SubRumor
                                            End If
                                            .MoveNext
                                        Loop
                                    End If
                                End With
                                RumorList.MoveNext
                            Loop
                            If Not SourceList.IsEmpty Then
                                If IsHas Then
                                    CompileEntitySection OptionSection, Character
                                Else
                                    CompileReportSection OptionSection, SourceList, False
                                End If
                                DoneSection = OptionSection
                            End If
                        End If
                End Select
                
            End If
            
        End If
        
        Replace FirstSection, "[#option " & OptionType & "]", CStr(SourceList.Count), , , vbTextCompare
        Section = Left(Section, EditStart - 1) & FirstSection & DoneSection & _
                  Mid(Section, OptionEnd + 9)
        
        EditStart = EditStart - 1 + Len(FirstSection) + Len(DoneSection)
        OptionStart = InStr(EditStart, Section, "[option ", vbTextCompare)
        OptionEnd = InStr(OptionStart + 1, Section, "[/option]", vbTextCompare)
        
    Loop

    FirstSection = Mid(Section, EditStart)
    CompileEntitySection FirstSection, Character
    Section = Left(Section, EditStart - 1) & FirstSection
    Set SourceList = Nothing
    CharName = ""
    
End Sub

Private Sub CompileSearchSection(ByRef Section As String)
'
' Name:         CompileSearchSection
' Parameters:   Section              the string section to repeat into a report
' Description:  Fill out a report with the current search.
'

    With Game.QueryEngine

        .QueryList.MoveTo SearchName
        If Not .QueryList.Off Then
            
            Dim Query As QueryClass
            Dim Blueprint As String
            Dim DoneSection As String
            Dim Inventory As QueryInventoryType
            Dim ShowSort As Boolean
            Dim QueryDesc As String
            
            Set Query = .QueryList.Item
            
            Inventory = Query.Inventory
            ShowSort = Not (Query.SortKey = "" Or Query.SortKey = qkName)
            
            .MakeQuery Query, ShowSort, SearchNot
            
            MatchValue = ""
            SortValue = ""
            
            .Results.First
            .Values.First
            .SortList.First
            
            Do Until .Results.Off
                Blueprint = Section
                If Not .Values.Off Then MatchValue = .Values.Item
                If ShowSort And Not .SortList.Off Then SortValue = .SortList.Item
                CompileEntitySection Section, .Results.Item
                If Not DoneSection = "" Then DoneSection = DoneSection & Divider
                DoneSection = DoneSection & Section
                Section = Blueprint
                .Results.MoveNext
                .Values.MoveNext
                .SortList.MoveNext
            Loop
            
            Query.First
            Do Until Query.Off
                If Not QueryDesc = "" Then QueryDesc = QueryDesc & " "
                QueryDesc = QueryDesc & Query.ClauseDescNext
            Loop
            
            FinishingTouches.Add "#search" & vbTab & CStr(.Results.Count)
            FinishingTouches.Add "searchname" & vbTab & CStr(SearchName)
            FinishingTouches.Add "terms" & vbTab & QueryDesc
            
        End If

    End With

    Section = DoneSection

End Sub

Private Sub CompileStatisticSection(ByRef Section As String)
'
' Name:         CompileSearchSection
' Parameters:   Section              the string section to repeat into a report
' Description:  Fill out a report with the current search.
'

    Game.QueryEngine.QueryList.MoveTo SearchName
    If Not Game.QueryEngine.QueryList.Off Then
        
        Dim Blueprint As String
        Dim TotalLabel As String
        Dim FinishedStatSection As String
        
        Dim EditStart As Long
        Dim RepeatStart As Long
        Dim RepeatEnd As Long
        Dim FirstSection As String
        Dim Subsection As String
        Dim MatchList As LinkedTraitList
        Dim DoneSection As String
        
        With Game.QueryEngine
            
            .GetStatistics StatType, Game.QueryEngine.QueryList.Item, StatKey, OKZero, StatTrait
            FinishedStatSection = ""
            
            If StatType = stMaxima Or StatType = stSums Then
                TotalLabel = "Total Traits"
            Else
                TotalLabel = "Characters Examined"
            End If
            
            .StatResults.First
            Do Until .StatResults.Off
            
                Blueprint = Section
                Set MatchList = .MatchSet(CStr(.StatResults.Item))

                Section = Replace(Section, "[value]", _
                          FormatForOutput(CStr(.StatResults.Item)), , , vbTextCompare)
                Section = Replace(Section, "[number]", _
                          CStr(.NumberSet(.StatResults.Item)), , , vbTextCompare)
                Section = Replace(Section, "[percent]", _
                          Format(.NumberSet(.StatResults.Item) / .Total, "##0.0%"), , , vbTextCompare)
                Section = Replace(Section, "[scale]", _
                          CStr(Round(StatScale * .NumberSet(.StatResults.Item) / .Maximum, 0)), _
                          , , vbTextCompare)
                Section = Replace(Section, "[value]", _
                          FormatForOutput(CStr(.StatResults.Item)), , , vbTextCompare)
                Section = Replace(Section, "[charcount]", _
                          CStr(MatchList.EntryCount), , , vbTextCompare)
                
                If InStr(1, Section, "[textgraph]", vbTextCompare) Then
                    Dim I As Long
                    Dim T As String
                    I = CLng(Round(StatScale * .NumberSet(.StatResults.Item) / .Maximum, 0))
                    T = ""
                    Do Until I = 0
                        If T <> "" Then T = T & Divider & Dot Else T = Dot
                        I = I - 1
                    Loop
                    Section = Replace(Section, "[textgraph]", T, , , vbTextCompare)
                End If
                
                EditStart = 1
                RepeatStart = InStr(1, Section, "[repeat]", vbTextCompare)
                RepeatEnd = InStr(RepeatStart + 1, Section, "[/repeat]", vbTextCompare)

                Do Until RepeatStart = 0 Or RepeatEnd = 0
                
                    FirstSection = Mid(Section, EditStart, RepeatStart - EditStart)
                    CompileSubsection FirstSection, Nothing
                    MatchList.First
                    DoneSection = ""
                    
                    Do Until MatchList.Off
                        If Not DoneSection = "" Then DoneSection = DoneSection & Divider
                        Subsection = Mid(Section, RepeatStart + 8, RepeatEnd - RepeatStart - 8)
                        CharName = MatchList.Trait.Name
                        MatchValue = MatchList.Trait.Total
                        If MatchList.Display = ldSimple Then MatchValue = ""
                        CompileSubsection Subsection, Nothing
                        DoneSection = DoneSection & Subsection
                        MatchList.MoveNext
                    Loop

                    Section = Left(Section, EditStart - 1) & FirstSection & DoneSection & _
                              Mid(Section, RepeatEnd + 9)

                    EditStart = EditStart - 1 + Len(FirstSection) + Len(DoneSection)
                    RepeatStart = InStr(EditStart, Section, "[repeat]", vbTextCompare)
                    RepeatEnd = InStr(RepeatStart + 1, Section, "[/repeat]", vbTextCompare)

                Loop

                FirstSection = Mid(Section, EditStart)
                CompileSubsection FirstSection, Nothing
                Section = Left(Section, EditStart - 1) & FirstSection
                
                If Not FinishedStatSection = "" Then FinishedStatSection = FinishedStatSection & Divider
                FinishedStatSection = FinishedStatSection & Section
                
                Section = Blueprint
                .StatResults.MoveNext
                
            Loop
            
            FinishingTouches.Add "total" & vbTab & CStr(.Total)
            FinishingTouches.Add "totallabel" & vbTab & TotalLabel
            FinishingTouches.Add "maximum" & vbTab & CStr(.Maximum)
            FinishingTouches.Add "count" & vbTab & CStr(.StatResults.Count)
            FinishingTouches.Add "searchname" & vbTab & CStr(SearchName)
        
        End With
        
        Select Case StatType
            Case stDistribution:            Blueprint = "Distribution of total"
            Case stDistinctDistribution:    Blueprint = "Distribution of distinct"
            Case stSpecificDistribution:    Blueprint = "Distribution of " & StatTrait
            Case stMaxima:                  Blueprint = "Maxima of"
            Case stSums:                    Blueprint = "Sums of"
        End Select
        On Error Resume Next
        Blueprint = Blueprint & " " & Game.QueryEngine.KeysToTitles(StatKey) & " for " & SearchName
        On Error GoTo 0
        FinishingTouches.Add "statname" & vbTab & Blueprint
        
    End If

    Section = FinishedStatSection

End Sub

Private Sub CompileGameSection(ByRef Section As String)
'
' Name:         CompileCharacterSection
' Parameters:   Section             the string section to process
' Description:  Populate a string with all the appropriate values from the game and its values.
'

    Dim EditStart As Long
    Dim ReportStart As Long
    Dim ReportEnd As Long
    Dim ReportType As String
    Dim FirstSection As String
    Dim Subsection As String
    Dim ReportSection As String
    Dim SourceList As LinkedList
    Dim DoneSection As String
    Dim SelSet As StringSet
    Dim SelList As LinkedList
    Dim UseCharLogic As Boolean
    
    EditStart = 1
    ReportStart = InStr(1, Section, "[report ", vbTextCompare)
    ReportEnd = InStr(ReportStart + 1, Section, "[/report]", vbTextCompare)
    Set SourceList = New LinkedList
    FinishingTouches.Clear

    Do Until ReportStart = 0 Or ReportEnd = 0
    
        FirstSection = Mid(Section, EditStart, ReportStart - EditStart)
        CompileEntitySection FirstSection, Game
    
        ReportType = Mid(Section, ReportStart, InStr(ReportStart, Section, "]") - ReportStart)
        ReportType = Mid(ReportType, 9)
        
        ReportSection = Mid(Section, ReportStart, ReportEnd - ReportStart)
        ReportSection = Mid(ReportSection, 10 + Len(ReportType))
        
        DoneSection = ""
        
        SourceList.Clear
        UseCharLogic = False
        Set SelSet = Nothing
        
        ReportType = LCase(ReportType)
        
        Select Case Left(ReportType, 3)
            Case "cha"
                Set SelSet = SelectionSet(osCharacters)
                Set SelList = CharacterList
                UseCharLogic = True
                FinishingTouches.Add "#characters" & vbTab & CStr(SelSet.Count)
            Case "pla"
                Set SelSet = SelectionSet(osPlayers)
                Set SelList = PlayerList
                UseCharLogic = True
                FinishingTouches.Add "#players" & vbTab & CStr(SelSet.Count)
            Case "ite"
                Set SelSet = SelectionSet(osItems)
                Set SelList = ItemList
                FinishingTouches.Add "#items" & vbTab & CStr(SelSet.Count)
            Case "rot"
                Set SelSet = SelectionSet(osRotes)
                Set SelList = RoteList
                FinishingTouches.Add "#rotes" & vbTab & CStr(SelSet.Count)
            Case "loc"
                Set SelSet = SelectionSet(osLocations)
                Set SelList = LocationList
                FinishingTouches.Add "#locations" & vbTab & CStr(SelSet.Count)
            Case "act"
                Set SelSet = SelectionSet(osActions)
                Set SelList = ActionList
                FinishingTouches.Add "#actions" & vbTab & CStr(SelSet.Count)
            Case "plo"
                Set SelSet = SelectionSet(osPlots)
                Set SelList = PlotList
                FinishingTouches.Add "#plots" & vbTab & CStr(SelSet.Count)
            Case "rum"
                Set SelSet = SelectionSet(osRumors)
                Set SelList = RumorList
                FinishingTouches.Add "#rumors" & vbTab & CStr(SelSet.Count)
            Case "cal"
                CompileEntitySection ReportSection, Game.Calendar
                DoneSection = ReportSection
            Case "sea"
                CompileSearchSection ReportSection
                DoneSection = ReportSection
            Case "sta"
                CompileStatisticSection ReportSection
                DoneSection = ReportSection
        End Select
                    
        If Not SelSet Is Nothing Then
            SelList.First
            Do Until SelList.Off
                If SelSet.Has(SelList.Item.Name) Then
                    SourceList.Append SelList.Item
                End If
                SelList.MoveNext
            Loop
            CompileReportSection ReportSection, SourceList, UseCharLogic
            DoneSection = ReportSection
        End If
        
        Section = Left(Section, EditStart - 1) & FirstSection & DoneSection & _
                  Mid(Section, ReportEnd + 9)
        
        EditStart = EditStart - 1 + Len(FirstSection) + Len(DoneSection)
        ReportStart = InStr(EditStart, Section, "[report ", vbTextCompare)
        ReportEnd = InStr(ReportStart + 1, Section, "[/report]", vbTextCompare)
        
    Loop

    FirstSection = Mid(Section, EditStart)
    CompileEntitySection FirstSection, Game
    Section = Left(Section, EditStart - 1) & FirstSection
    
    FinishingTouches.First
    Do Until FinishingTouches.Off
        Dim FKey As String
        Dim FVal As String
        FKey = FinishingTouches.StrItem
        FVal = FormatForOutput(Mid(FKey, InStr(FKey, vbTab) + 1))
        FKey = "[" & Left(FKey, InStr(FKey, vbTab) - 1) & "]"
        Section = Replace(Section, FKey, FVal, , , vbTextCompare)
        FinishingTouches.MoveNext
    Loop
    
    Set SourceList = Nothing
    
End Sub

Private Sub LayoutText(ByRef Text As String)
'
' Name:         LayoutText
' Parameters:   Text            String to format
' Description:  Layout a string according to the [wrap] and [col] keywords.
'

    Dim EditStart As Long
    Dim EditEnd As Long
    Dim Keyword As String
    Dim Param As String
    Dim LineStart As Long
    Dim WrapStart As Long
    Dim WrapText As String
    Dim PCol As Long
    
    EditStart = 1
    WrapStart = 1
    
    Do While GetNextKeyword(Text, Keyword, Param, EditStart, EditEnd)
    
        Select Case Keyword
            Case "col"
                PCol = Int(Val(Param))
                If PCol > 0 Then
                    LineStart = InStrRev(Text, vbCrLf, EditStart) + 2
                    If LineStart = 2 Then LineStart = 1
                    Text = Left(Text, LineStart - 1) & _
                           Left(Mid(Text, LineStart, EditStart - LineStart) & _
                                 String(PCol, " "), PCol - 1) & _
                           Mid(Text, EditEnd + 1)
                    EditStart = LineStart + PCol - 1
                Else
                    EditStart = EditStart + 1
                End If
            Case "wrap"
                WrapText = Mid(Text, WrapStart, EditStart - WrapStart)
                Call WrapString(WrapText)
                Wrap = Val(Param)
                Text = Left(Text, WrapStart - 1) & WrapText & Mid(Text, EditEnd + 1)
                WrapStart = EditStart
                EditStart = EditStart + 1
            Case Else
                EditStart = EditEnd + 1
        End Select
        
    Loop

    If Wrap > 0 Then
        WrapText = Mid(Text, WrapStart)
        Call WrapString(WrapText)
        Text = Left(Text, WrapStart - 1) & WrapText
    End If
    
End Sub

Private Sub TextPrint(ByRef OutString As String)
'
' Name:         TextPrint
' Parameters:   OutString       the data to print
' Description:  Send the given data to the printer using the method best suited to
'               either a Text or an RTF OutputFormat.  HTML need not apply.
'

    Dim StartLoc As Long
    Dim EndLoc As Long
    Dim ICopy As Integer
    
    If OutputFormat = ofRTF Then
        RTFBox.TextRTF = OutString
        OutString = RTFBox.Text
        If Landscape Xor (Printer.Orientation = vbPRORLandscape) Then
            If PrintInProgress Then Printer.EndDoc
            Printer.Orientation = IIf(Landscape, vbPRORLandscape, vbPRORPortrait)
        ElseIf PrintInProgress Then
            Printer.NewPage
        End If
    Else
        Printer.FontName = "Courier New"
        Printer.FontSize = 10
        If PrintInProgress Then Printer.NewPage
    End If
        
    For ICopy = 1 To Copies
        
        PrintInProgress = True
        StartLoc = 0
        EndLoc = InStr(1, OutString, "[pagebreak]", vbTextCompare)
        If EndLoc = 0 Then EndLoc = Len(OutString) + 1
        
        Do
        
            If Not StartLoc = 0 Then Printer.NewPage
        
            If OutputFormat = ofText Then
                Printer.Print Mid(OutString, StartLoc + 1, EndLoc - StartLoc - 1)
            Else
                RTFBox.SelStart = StartLoc
                RTFBox.SelLength = EndLoc - StartLoc - 1
                Call PrintRTF(RTFBox, Margin(M_LEFT), Margin(M_TOP), Margin(M_RIGHT), Margin(M_BOTTOM))
            End If
            
            StartLoc = EndLoc + Len("[pagebreak]") - 1
            EndLoc = InStr(StartLoc, OutString, "[pagebreak]", vbTextCompare)
            If EndLoc = 0 Then EndLoc = Len(OutString) + 1
            
        Loop Until StartLoc >= Len(OutString)
    
        If ICopy < Copies Then Printer.NewPage

    Next ICopy

End Sub

Private Sub HTMLPrintPart(ByRef OutString As String)
'
' Name:         HTMLPrintPart
' Parameter:    Outstring           HTML output to write to the file
' Description:  Construct the file that Grapevine asks the web browser to print.
'

    Dim BodyStart As Long
    Dim BodyEnd As Long
    
    OutString = Replace(OutString, "[pagebreak]", _
        "<P STYLE=""page-break-after:always""></P>", , , vbTextCompare)
    
    BodyStart = InStr(1, OutString, "<body", vbTextCompare)
    If BodyStart = 0 Then
        BodyStart = 1
    Else
        BodyStart = InStr(BodyStart + 1, OutString, ">") + 1
    End If
    
    BodyEnd = InStr(BodyStart, OutString, "</body", vbTextCompare) - 1
    If BodyEnd < BodyStart Then BodyEnd = Len(OutString)
    
    If HTMLFileNum = -1 Then
        
        HTMLTempFile = App.Path & "\~GVPrint.html"
        HTMLFileNum = FreeFile
        Open HTMLTempFile For Output As HTMLFileNum
        Print #HTMLFileNum, Left(OutString, BodyEnd)
        
    Else
    
        Print #HTMLFileNum, "<P style=""page-break-after:always""></P>"
        Print #HTMLFileNum, Mid(OutString, BodyStart, BodyEnd - BodyStart + 1)
        
    End If
    
End Sub

Private Sub HTMLPrintFinish(OutString As String)
'
' Name:         HTMLPrintFinish
' Parameter:    Outstring           HTML output to write to the file
' Description:  Finish writing the HTML output file, and .
'

    Dim BodyEnd As Long
    
    BodyEnd = InStr(1, OutString, "</body", vbTextCompare)
    If BodyEnd <> 0 Then
        Print #HTMLFileNum, Mid(OutString, BodyEnd)
    End If
    
    Print #HTMLFileNum, "<SCRIPT Language=""JavaScript1.2""> window.print(); </SCRIPT>"
    Close #HTMLFileNum
    
    mdiMain.LaunchBrowser HTMLTempFile
    
End Sub

Public Sub Output(TInfo As TemplateClass, ProgressLabel As Label)
'
' Name:         Output
' Parameters:   TInfo       Template used to create the output: if TInfo is Nothing
'                           then create character sheets
' Description:  Create Grapevine output from the given template.  Settings on this
'               class determine the device to which the output should be written and
'               the format it is to take.
'

    Dim TSubject As Long
    Dim TFile As String
    Dim TSet As Long
    Dim TList As LinkedList
    Dim IsCharacterSheet As Boolean
    Dim Blueprint As String
    Dim OutString As String
    Dim MoreOutput As Boolean
    Dim AlwaysGetFile As Boolean
    Dim FileNum As Integer
    Dim Extension As String
    Dim OutputDirectory As String
    Dim Increment As Single
    Dim Total As Single
    Dim CheckOverwrite As Boolean
    Dim NoEMail As String
    Dim SaveFile As String
    
    ' Initialize basic values
    
    ErrorFlag = False
    Total = 0
    CheckOverwrite = False
    ProgressLabel.Caption = CStr(Int(Total)) & "%"
    Select Case OutputFormat
        Case ofText:    Extension = ".txt"
        Case ofRTF:     Extension = ".rtf"
        Case ofHTML:    Extension = ".html"
    End Select
    AdvanceList.Clear
    FinishingTouches.Clear
    Dot = "O"
    TempDot = "Õ"
    EmptyDot = "Ø"
    Divider = ""
    Display = ldDefault
    Wrap = 0
    DateFormat = "mmmm d, yyyy"
    StatScale = 100
    PrintInProgress = False
    HTMLFileNum = -1
    If PrintAfterSave And OutputDevice = odPrinter Then OutputDevice = odFile
    
    'Initialize character sheet values if TInfo is nothing
    
    If TInfo Is Nothing Then
        
        TSubject = ooCharacters
        TSet = osCharacters
        Set TList = CharacterList
        IsCharacterSheet = True
        TFile = ""
        AlwaysGetFile = (SelectionSet(TSet).Count = 1)
        If SelectionSet(TSet).Count = 0 Then Exit Sub
        Increment = 100 / SelectionSet(TSet).Count
        
    Else    'Initialize values for the template given
        
        TSubject = TInfo.GetSubject(OutputFormat)
        TFile = TInfo.GetFilename(OutputFormat)
        IsCharacterSheet = False
        If TSubject = ooFileError Then
            ErrorFlag = True
            ErrorMessage = "Grapevine can't find the template file it needs to create this " _
                         & IIf(TInfo.Name = "", "output", TInfo.Name) & "."
            Exit Sub
        End If
        
        ' Initialize values for multiple-output templates
        
        If (Not TSubject) And ooReport Then
        
            If TSubject And ooRumors Then TSet = osRumors
            If TSubject And ooPlots Then TSet = osPlots
            If TSubject And ooActions Then TSet = osActions
            If TSubject And ooLocations Then TSet = osLocations
            If TSubject And ooRotes Then TSet = osRotes
            If TSubject And ooItems Then TSet = osItems
            If TSubject And ooPlayers Then TSet = osPlayers
            If TSubject And ooCharacters Then TSet = osCharacters
            Select Case TSet
                Case osCharacters:      Set TList = CharacterList
                Case osPlayers:         Set TList = PlayerList
                Case osItems:           Set TList = ItemList
                Case osRotes:           Set TList = RoteList
                Case osLocations:       Set TList = LocationList
                Case osActions:         Set TList = ActionList
                Case osPlots:           Set TList = PlotList
                Case osRumors:          Set TList = RumorList
            End Select
            
            AlwaysGetFile = (SelectionSet(TSet).Count = 1)
            If SelectionSet(TSet).Count = 0 Then Exit Sub
            Increment = 100 / SelectionSet(TSet).Count
        
        Else    'Initialize values for single-output (report) templates
            
            AlwaysGetFile = True
            Set TList = Nothing
            MoreOutput = True
            Increment = 100
            
        End If
        
    End If
        
    Select Case OutputDevice
        Case odFile     ' For File output, initialize the CDialog and get the base directory if needed
            OutputDirectory = GetSetting(App.Title, "Output", "Folder", CurDir)
            On Error Resume Next
            ChDrive OutputDirectory
            ChDir OutputDirectory
            On Error GoTo 0
            If Not AlwaysGetFile Then
                Screen.MousePointer = vbDefault
                frmSaveDirectory.GetSaveDirectory "Choose Destination Directory"
                AlwaysGetFile = (frmSaveDirectory.Value = sdIndividual)
                CheckOverwrite = (frmSaveDirectory.chkPrompt.Value = vbChecked)
                Screen.MousePointer = vbHourglass
                If frmSaveDirectory.Value = sdCancel Then
                    Unload frmSaveDirectory
                    Exit Sub
                End If
                Unload frmSaveDirectory
            End If
            With CDialog
                .DefaultExt = Mid(Extension, 2)
                .DialogTitle = "Save To File"
                .Filter = "Text Files (*.txt)|*.txt|Rich Text Files (*.rtf)|*.rtf" & _
                          "|HTML Files (*.html)|*.html|All Files (*.*)|*.*"
                .FilterIndex = OutputFormat + 1
                .Flags = cdlOFNHideReadOnly + cdlOFNNoReadOnlyReturn + _
                         cdlOFNOverwritePrompt + cdlOFNPathMustExist
                .InitDir = CurDir
            End With
            OutputDirectory = CurDir
        Case odPrinter
            Increment = Increment / Copies
    End Select
        
    ' for rumor output, initialize the rumor recipients
    
    If ((TSubject And ooRumors) <> 0) Or _
       ((TSubject And ooCharacters) <> 0 And (OutputOptions And ooOptRumors) <> 0) Then
        Game.APREngine.PrepareRumorOutput GameDate
    End If
    
    ' Initialize TList and MoreOutput
    
    If Not TList Is Nothing Then
        TList.First
        Do Until TList.Off
            If SelectionSet(TSet).Has(TList.Item.Name) Then Exit Do
            TList.MoveNext
        Loop
        MoreOutput = Not TList.Off
    
    End If
    
    ' While there is more output, continue the processing loop
    
    Do While MoreOutput
    
        ' Character Sheet templates must be re-read on each iteration if they've changed
    
        If IsCharacterSheet Then
            Dim NewTFile As String
            Game.TemplateList.MoveTo TList.Item.Race & tnCharSheetSuffix
            If Not Game.TemplateList.Off Then
                NewTFile = Game.TemplateList.Item.GetFilename(OutputFormat)
            Else
                NewTFile = ""
            End If
            If NewTFile <> TFile Then
                Blueprint = ""
                TFile = NewTFile
            End If
        End If
        
        ' If the template hasn't been read yet, read it into the Blueprint
        
        If Blueprint = "" Then
            
            Dim TempFile As String
            Dim NotFound As Boolean
            
            TempFile = FindFile(TFile)
            
            ' If located, read it
                          
            If TempFile <> "" Then
                TFile = TempFile
                FileNum = FreeFile
                ErrorMessage = "Error reading " & TFile & ": "
                On Error GoTo FinalOutputError
                Open TFile For Input As #FileNum
                Do Until EOF(FileNum)
                    Line Input #FileNum, OutString
                    If Blueprint <> "" Then OutString = vbCrLf & OutString
                    Blueprint = Blueprint & OutString
                Loop
                Close #FileNum
                On Error GoTo 0
            Else
                If MsgBox("Grapevine can't find the template " & TFile & _
                          ".  Continue?", vbYesNo + vbExclamation, "Template Error") = vbNo Then
                    Exit Do
                End If
            End If
            
            ' If it'll be printed, reset its print settings
            If OutputDevice = odPrinter Then
                Landscape = False
                Margin(M_LEFT) = 1440
                Margin(M_RIGHT) = 1440
                Margin(M_TOP) = 1440
                Margin(M_BOTTOM) = 1440
            End If
                        
        End If
        
        ' Now if a Blueprint is in our hands, process it
        
        If Blueprint <> "" Then
            
            ' Copy the Blueprint into an Output string, then compile it into its final form
            
            OutString = Blueprint
            If TSubject And ooReport Then
                CompileGameSection OutString
            ElseIf Not TList Is Nothing Then
                If TSet = osCharacters Or TSet = osPlayers Then
                    CompileCharacterSection OutString, TList.Item
                Else
                    CompileEntitySection OutString, TList.Item
                End If
            End If
            If OutputFormat = ofText Then LayoutText OutString
                            
            ' Create a base filename, just in case it's needed
        
            If IsCharacterSheet Then
                SaveFile = TList.Item.Name
            ElseIf Not TList Is Nothing Then
                SaveFile = TList.Item.Name & " " & TInfo.Name
            Else
                SaveFile = TInfo.Name
                If SaveFile = "" Then SaveFile = "Report"
            End If
            SaveFile = ConvertToFileName(SaveFile) & Extension
            
            ' Do with the output string as the output device requires
            
            Select Case OutputDevice
                Case odFile
                
                    ' Use the base filename to get a full filename
                
                    If AlwaysGetFile Then
                        CDialog.FileName = SaveFile
                        On Error Resume Next
                        CDialog.ShowSave
                        If Err.Number <> 0 Then
                            SaveFile = ""
                        Else
                            SaveFile = CDialog.FileName
                        End If
                        On Error GoTo 0
                    Else
                        SaveFile = SlashPath(OutputDirectory) & SaveFile
                        If CheckOverwrite And Dir(SaveFile) <> "" Then
                            If MsgBox("Overwrite existing file """ & ShortFile(SaveFile) & """?", _
                                      vbQuestion + vbYesNo, "Overwrite File") = vbNo Then SaveFile = ""
                        End If
                    End If
                    
                    ' Format the output
                    
                    If OutputFormat = ofHTML Then
                        OutString = Replace(OutString, "[pagebreak]", _
                                    "<P STYLE=""page-break-after:always""></P>", , , vbTextCompare)
                    Else
                        OutString = Replace(OutString, "[pagebreak]", "", , , vbTextCompare)
                    End If
                    
                    ' Write the file
                    
                    If SaveFile <> "" Then
                        FileNum = FreeFile
                        ErrorMessage = "Error saving " & ShortFile(SaveFile) & ": "
                        On Error GoTo FinalOutputError
                        Open SaveFile For Output As #FileNum
                        Print #FileNum, OutString;
                        Close #FileNum
                        On Error GoTo 0
                        
                        If PrintAfterSave Then
                            For FileNum = 1 To Copies
                                ShellExecute mdiMain.hWnd, "print", SaveFile, "", "", 7
                            Next FileNum
                        End If
                        
                    End If

                Case odPrinter
                    
                    ' Send result to the appropriate print routine
                    
                    ErrorMessage = "Error while trying to print:"
                    On Error GoTo FinalOutputError
                    If OutputFormat = ofHTML Then
                        Call HTMLPrintPart(OutString)
                    Else
                        Call TextPrint(OutString)
                    End If
                    On Error GoTo 0

                Case odEMail
                    
                    Dim TRecipient As String
                    
                    If TList Is PlayerList Or TList Is CharacterList Then 'Get current e-mail recipient
                        TRecipient = TList.Item.ExpandedEMail
                        If TRecipient = "" Then
                            NoEMail = NoEMail & vbCrLf & TList.Item.Name
                        End If
                    End If
                    
' Include the SMTP.ocx control in the project in order to re-enable the code
' below. Note that SMTP.ocx is detected (inaccurately) as a virus component
' by many antivirus programs.
                    
'                    With Mailer                                         'Prepare SMTP E-Mail control
'                        .SendTo = Replace(SendTo, SendToSelect, TRecipient) 'Add current recipient
'                        If Left(.SendTo, 1) = "," Then .SendTo = Mid(.SendTo, 2)
'                        If .SendTo <> "" Then                           ' If there is a recipient...
'
'                            .CC = Replace(CC, SendToSelect, TRecipient) 'Prepare CC, BCC, Subject, Reply-To
'                            If Left(.CC, 1) = "," Then .CC = Mid(.CC, 2)
'                            .BCC = Replace(BCC, SendToSelect, TRecipient)
'                            If Left(.BCC, 1) = "," Then .BCC = Mid(.BCC, 2)
'                            .MessageSubject = MessageSubject & _
'                                              IIf(ReplyTo = "", "", vbCrLf & "Reply-To: " & ReplyTo)
'                            If Not TList Is Nothing Then
'                                .MessageSubject = Replace(.MessageSubject, "[Name]", TList.Item.Name, _
'                                                  , , vbTextCompare)
'                            End If
'
'                            If OutputFormat = ofRTF Or AlwaysAttach Then
'
'                                Dim BaseName As String
'                                BaseName = SaveFile
'                                SaveFile = SlashPath(App.Path) & BaseName
'                                Do Until Dir(SaveFile) = ""
'                                    BaseName = "~" & BaseName
'                                    SaveFile = SlashPath(App.Path) & BaseName
'                                Loop
'                                FileNum = FreeFile                  'Save and attach file
'                                ErrorMessage = "Error saving " & ShortFile(SaveFile) & " for attachment: "
'                                On Error GoTo FinalOutputError
'                                Open SaveFile For Output As #FileNum
'                                Print #FileNum, OutString;
'                                Close #FileNum
'                                On Error GoTo 0
'                                If .Attachments.Count > 0 Then
'                                    .Attachments.Add SaveFile, Before:=1
'                                Else
'                                    .Attachments.Add SaveFile
'                                End If
'
'                                .MessageText = MessageHeader        'Add text header
'                                .MessageHTML = ""
'
'                            Else
'
'                                Select Case OutputFormat
'                                    Case ofText                             'Add header to text message
'                                        .MessageHTML = ""
'                                        If MessageHeader <> "" Then
'                                            .MessageText = MessageHeader & vbCrLf & vbCrLf & OutString
'                                        Else
'                                            .MessageText = OutString
'                                        End If
'
'                                    Case ofHTML
'
'                                        If MessageHeader <> "" Then
'                                            Dim I As Long                       'Add header to HTML
'                                            I = InStr(1, OutString, "<body", vbTextCompare)
'                                            I = InStr(I + 1, OutString, ">")
'                                            If I > 0 Then
'                                                OutString = Left(OutString, I) & vbCrLf & _
'                                                            "<P STYLE=""page-break-after: always"">" & _
'                                                            FormatForOutput(MessageHeader) & _
'                                                            "</P>" & Mid(OutString, I + 1)
'                                            End If
'                                        End If
'
'                                        .MessageHTML = OutString            'Add HTML
'                                        .MessageText = ""
'
'                                End Select
'
'                            End If
'
'                            On Error GoTo FinalOutputError
'
'                            Do
'                                .Tag = ""
'                                .SendEmail
'                                If .Tag <> "" Then
'                                    Dim UAct As Integer
'                                    ErrorMessage = "Error sending e-mail:" & vbCrLf & vbCrLf & .Tag & _
'                                                   vbCrLf & vbCrLf & "ABORT cancels all sending." & vbCrLf & _
'                                                   "RETRY tries again to send the message that got the error." & _
'                                                   vbCrLf & "IGNORE resumes sending with the next message in the queue."
'                                    If TList Is Nothing Then
'                                        UAct = vbAbort
'                                        ErrorMessage = "Error sending e-mail:"
'                                    Else
'                                        UAct = MsgBox(ErrorMessage, vbAbortRetryIgnore, "E-Mail Error")
'                                        ErrorMessage = "Aborted after e-mail error:"
'                                    End If
'                                    Select Case UAct
'                                        Case vbAbort
'                                            Err.Raise vbObjectError + 64, "OutputEngine", .Tag
'                                        Case vbIgnore
'                                            .Tag = ""
'                                    End Select
'                                End If
'                            Loop Until .Tag = ""
'
'                            If OutputFormat = ofRTF Or AlwaysAttach Then
'                                ErrorMessage = "Error deleting attachment file " & ShortFile(SaveFile) & ":"
'                                Kill SaveFile
'                                .Attachments.Remove 1
'                            End If
'
'                            On Error GoTo 0
'
'                        Else
'                            ' No Recipients for this message!
'                        End If
'                    End With
            
            End Select
    
        End If
    
        ' Move to the next item for output
    
        If TList Is Nothing Then
            MoreOutput = False
        Else
            TList.MoveNext
            Do Until TList.Off
                If SelectionSet(TSet).Has(TList.Item.Name) Then Exit Do
                TList.MoveNext
            Loop
            MoreOutput = Not TList.Off
        End If
        
        ' Increment the counter
        
        Total = Total + Increment
        ProgressLabel.Caption = CStr(Int(Total)) & "%"
        
    Loop
    
    ErrorMessage = "Finishing the print job:"
    On Error GoTo FinalOutputError
    If PrintInProgress Then Printer.EndDoc
    If HTMLFileNum > -1 Then Call HTMLPrintFinish(OutString)
    If NoEMail <> "" Then
        MsgBox "Unable to find e-mail addresses for:" & vbCrLf & NoEMail & vbCrLf & vbCrLf & _
               "All other deliveries were made.", vbOKOnly + vbExclamation, "No E-Mail Addresses"
    End If
    On Error GoTo 0
    
    GoTo OutputEnd
    
FinalOutputError:
    ErrorFlag = True
    ErrorMessage = ErrorMessage & vbCrLf & vbCrLf & Err.Description
    If OutputDevice = odPrinter Then Printer.KillDoc
    Resume OutputEnd
    
OutputEnd:
    
    If OutputDevice = odFile Then
        SaveSetting App.Title, "Output", "Folder", OutputDirectory
    End If
    If PrintAfterSave And OutputDevice = odFile Then OutputDevice = odPrinter
    
End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Cleanup after using the output engine.
'

    Dim I As Integer
    
    Set AdvanceList = Nothing
    Set FinishingTouches = Nothing
    
    For I = osMin To osMax
        Set SelectionSet(I) = Nothing
    Next I
    
End Sub
