VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLReaderClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Terminal As Boolean          'the current tag is a closing tag
Public Tag As String                'Type of current tag

Public OldFile As Boolean           'Whether this is an old Grapevine file
Public VersionTag As String         'the entire file's first tag

Private FileNum As Integer          'The file number from which to read
Private Attributes As Collection    'The tag's attributes
Private DataStack As LinkedList     'The ordinary data in the current stack of tags
Private XMLThusFar As String        'XML yet to process

Public Sub BeginReading(XMLFile As Integer)
'
' Name:         BeginReading
' Parameters:   XMLFile         the file number from which to read
' Description:  Initialize all values for reading XML.
'

    If FileNum < 0 Then                 'reading no file yet
        
        Dim Read As String
        
        FileNum = XMLFile
        Tag = ""
        Terminal = False
        Do Until Attributes.Count = 0   'clear attribute collection
            Attributes.Remove 1
        Loop
        DataStack.Clear
        
        Line Input #FileNum, Read
        VersionTag = Left(Read, InStr(Read, ">"))
        XMLThusFar = Mid(Read, InStr(Read, ">") + 1)
        
        OldFile = (VersionTag = GameFileVersionTag0 Or VersionTag = GameFileVersionTag1 _
                   Or VersionTag = GameFileVersionTag2 Or VersionTag = GameFileVersionTag3 _
                   Or VersionTag = GameFileVersionTag4 Or VersionTag = GameFileVersionTag5 _
                   Or VersionTag = ExchangeFileVersionTag0 Or VersionTag = ExchangeFileVersionTag1)
        
    Else
        Err.Raise vbObjectError + 1, "XMLReader", "Began reading file while one " _
                  & "was already being read"
    End If
    
End Sub

Public Sub ReadTag()
'
' Name:         ReadTag
' Parameters:   FileNum     file from which to read
' Description:  Read the next XML tag from the file.  Non-tag data goes on the
'               DataStack.
'

    Dim Data As String
    Dim CData As String
    Dim FinalData As String
    Dim WholeTag As String
    Dim Attrib As String
    Dim QChar As String
    Dim Value As String
    
    Dim Read As String
    Dim TagBegin As Integer
    Dim TagEnd As Integer
    Dim AttBegin As Integer
    Dim AttEnd As Integer
    
    If Left(Tag, 1) = "/" Then                          'if previous tag finishes a data-capable
        DataStack.Last                                  'section, pop its data
        DataStack.Remove
    End If

    WholeTag = ""

    Do                                                  'until a tag is found

        TagBegin = InStr(XMLThusFar, "<")               'read lines until a tag is contained
        TagEnd = InStr(XMLThusFar, ">")                 'therein
        Do Until TagBegin > 0 And TagEnd > TagBegin
            If EOF(FileNum) Then
                Err.Raise vbObjectError + 3, "XMLReaderClass", _
                          "Attempted to read a tag past the end of the file"
            End If
            Line Input #FileNum, Read
            XMLThusFar = XMLThusFar & Read & vbCrLf
            TagBegin = InStr(XMLThusFar, "<")
            TagEnd = InStr(XMLThusFar, ">")
        Loop
    
        Data = Left(XMLThusFar, TagBegin - 1)
        Data = TrimWhiteSpace(Data)                     'store data read
        If Data <> "" Then
            Data = FormatFromXML(Data)
            FinalData = FinalData & Data
        End If
        
        WholeTag = Mid(XMLThusFar, TagBegin, TagEnd - TagBegin + 1)
        
        If Left(WholeTag, 9) = "<![CDATA[" Then         'CData tag
        
            TagBegin = TagBegin + 9
            TagEnd = InStr(XMLThusFar, "]]>")           'Read until end of CDATA tag
            Do Until TagEnd > 0
                If EOF(FileNum) Then
                    Err.Raise vbObjectError + 4, "XMLReaderClass", _
                              "Couldn't find the end of a CDATA field starting with: " _
                              & vbCrLf & Mid(XMLThusFar, TagBegin, 16)
                End If
                Line Input #FileNum, Read
                XMLThusFar = XMLThusFar & Read & vbCrLf
                TagEnd = InStr(XMLThusFar, "]]>")
            Loop
            CData = Mid(XMLThusFar, TagBegin, TagEnd - TagBegin)
            FinalData = FinalData & CData               'assign contents to finaldata
            XMLThusFar = Mid(XMLThusFar, TagEnd + 3)
            WholeTag = ""
            
        ElseIf Left(WholeTag, 4) = "<!--" Then          'Comment -- discard
        
            TagEnd = InStr(XMLThusFar, "-->")           'Read until end of comment
            Do Until TagEnd > 0
                If EOF(FileNum) Then
                    Err.Raise vbObjectError + 5, "XMLReaderClass", _
                              "Couldn't find the end of a comment starting with: " _
                              & vbCrLf & Mid(XMLThusFar, TagBegin + 4, 16)
                End If
                Line Input #FileNum, Read
                XMLThusFar = XMLThusFar & Read & vbCrLf
                TagEnd = InStr(XMLThusFar, "-->")
            Loop
            XMLThusFar = Mid(XMLThusFar, TagEnd + 3)
            WholeTag = ""
        
        End If
    
    Loop Until WholeTag <> ""
    
    XMLThusFar = Mid(XMLThusFar, TagEnd + 1)
    
    If InStr(WholeTag, " ") > 0 Then                    'get tag name
        Tag = Mid(WholeTag, 2, InStr(WholeTag, " ") - 2)
    Else
        Tag = Mid(WholeTag, 2, Len(WholeTag) - 2)
    End If
    
    Do Until Attributes.Count = 0                       'clear attribute collection
        Attributes.Remove 1
    Loop
    
    AttBegin = InStr(WholeTag, " ")
    AttEnd = InStr(WholeTag, "=")
    
    Do While AttBegin > 0 And AttEnd > AttBegin         'loop while attributes remain
    
        Attrib = Mid(WholeTag, AttBegin + 1, AttEnd - AttBegin - 1)
        Attrib = TrimWhiteSpace(Attrib)
        WholeTag = TrimWhiteSpace(Mid(WholeTag, AttEnd + 1))
        QChar = Left(WholeTag, 1)
        AttEnd = InStr(2, WholeTag, QChar)
        If AttEnd = 0 Then
            Err.Raise vbObjectError + 6, "XMLReaderClass", _
                      "A <" & Tag & "> tag has a malformed '" & Attrib & "' attribute " _
                      & "whose value begins: " & vbCrLf & Left(WholeTag, 16)
        End If
        Value = Trim(Mid(WholeTag, 2, AttEnd - 2))
        Value = FormatFromXML(Value)
        On Error Resume Next
        Attributes.Add Value, Attrib
        On Error GoTo 0
        WholeTag = Mid(WholeTag, AttEnd + 1)
        AttBegin = InStr(WholeTag, " ")
        AttEnd = InStr(WholeTag, "=")
    
    Loop
    
    If DataStack.IsEmpty Then
        DataStack.Append FinalData
    Else
        If FinalData <> "" Then
            DataStack.Last
            FinalData = DataStack.Item & FinalData
            DataStack.Remove
            DataStack.Append FinalData
        End If
    End If
    
    Terminal = Left(Tag, 1) = "/" Or Right(WholeTag, 2) = "/>"
    If Terminal Then
        If Left(Tag, 1) = "/" Then Tag = Mid(Tag, 2)
        If Right(Tag, 1) = "/" Then Tag = Left(Tag, Len(Tag) - 1)
    End If
    If Not Terminal Then DataStack.Append ""
    
End Sub

Public Function GetTagText() As String
'
' Name:         GetTagText
' Returns:      The ordinary text prior to this tag, possibly including
'               sections extracted from CDATA markup.
'

    DataStack.Last
    GetTagText = DataStack.Item

End Function

Public Function GetFileNum() As Integer
'
' Name:         GetFileNum
' Returns:      the number of the file this reader is reading from.
'

    GetFileNum = FileNum

End Function

Public Function GetSingle(Attrib As String, Optional Default = 0) As Single
'
' Name:         GetSingle
' Parameters:   Attrib      name of the Single to get
'               Default     default value if it's not found
' Description:  Get a Single attribute from this tag's collection.  If not found,
'               return default value.
' Returns:      Value of the attribute.
'

    On Error GoTo SngNotFound
    GetSingle = CSng(Attributes.Item(Attrib))
    Exit Function

SngNotFound:
    GetSingle = CSng(Default)

End Function

Public Function GetDate(Attrib As String, Optional Default) As Date
'
' Name:         GetDate
' Parameters:   Attrib      name of the Date to get
'               Default     default value if it's not found
' Description:  Get a Date attribute from this tag's collection.  If not found,
'               return default value.
' Returns:      Value of the attribute.
'

    On Error GoTo DateNotFound
    GetDate = CDate(Attributes.Item(Attrib))
    Exit Function

DateNotFound:
    If IsMissing(Default) Then GetDate = Date Else GetDate = CDate(Default)

End Function

Public Function GetLong(Attrib As String, Optional Default = 0) As Long
'
' Name:         GetLong
' Parameters:   Attrib      name of the Long to get
'               Default     default value if it's not found
' Description:  Get a Long attribute from this tag's collection.  If not found,
'               return default value.
' Returns:      Value of the attribute.
'

    On Error GoTo LngNotFound
    GetLong = CLng(Attributes.Item(Attrib))
    Exit Function

LngNotFound:
    GetLong = CLng(Default)

End Function

Public Function GetDouble(Attrib As String, Optional Default = 0) As Double
'
' Name:         GetDouble
' Parameters:   Attrib      name of the Double to get
'               Default     default value if it's not found
' Description:  Get a Double attribute from this tag's collection.  If not found,
'               return default value.
' Returns:      Value of the attribute.
'

    On Error GoTo DblNotFound
    GetDouble = CDbl(Attributes.Item(Attrib))
    Exit Function

DblNotFound:
    GetDouble = CDbl(Default)

End Function

Public Function GetInteger(Attrib As String, Optional Default = 0) As Integer
'
' Name:         GetInteger
' Parameters:   Attrib      name of the Integer to get
'               Default     default value if it's not found
' Description:  Get a Integer attribute from this tag's collection.  If not found,
'               return default value.
' Returns:      Value of the attribute.
'

    On Error GoTo IntNotFound
    GetInteger = CInt(Attributes.Item(Attrib))
    Exit Function

IntNotFound:
    GetInteger = CInt(Default)

End Function

Public Function GetBoolean(Attrib As String, Optional Default = False) As Boolean
'
' Name:         GetBoolean
' Parameters:   Attrib      name of the boolean to get
'               Default     default value if it's not found
' Description:  Get a Boolean attribute from this tag's collection.  If not found,
'               return default value.
' Returns:      Value of the attribute.
'

    Dim b As String

    On Error GoTo BoolNotFound
    
    b = LCase(Attributes.Item(Attrib))
    GetBoolean = IIf(b = "yes" Or b = "y" Or b = "true" Or b = "t", True, False)
    Exit Function

BoolNotFound:
    GetBoolean = CBool(Default)

End Function

Public Function GetString(Attrib As String, Optional Default = "") As String
'
' Name:         GetString
' Parameters:   Attrib      name of the string attribute to get
' Description:  Get an attribute from this tag's collection.
' Returns:      Value of the attribute, or "" if not found.
'

    On Error GoTo StrNotFound
    GetString = Attributes.Item(Attrib)
    Exit Function

StrNotFound:
    GetString = CStr(Default)

End Function

Private Function FormatFromXML(Data As String) As String
'
' Name:         FormatFromXML
' Parameters:   Data        the string to format
' Description:  Replace XML markup with ordinary characters.
' Returns:      Formatted string.
'

    FormatFromXML = Replace(Data, "&quot;", """")
    FormatFromXML = Replace(FormatFromXML, "&apos;", "'")
    FormatFromXML = Replace(FormatFromXML, "&lt;", "<")
    FormatFromXML = Replace(FormatFromXML, "&gt;", ">")
    FormatFromXML = Replace(FormatFromXML, "&amp;", "&")
    If InStr(FormatFromXML, vbCrLf) > 0 Then
        FormatFromXML = Replace(FormatFromXML, vbCrLf, " ")
        Do Until InStr(FormatFromXML, "  ") = 0
            FormatFromXML = Replace(FormatFromXML, "  ", " ")
        Loop
    End If
    
End Function

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Create objects needed by the XMLReaderClass.
'

    Set Attributes = New Collection
    Set DataStack = New LinkedList
    FileNum = -1
    
End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Destroy objects needed by the XMLReaderClass.
'

    Set Attributes = Nothing
    Set DataStack = Nothing

End Sub
