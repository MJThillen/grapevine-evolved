VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLWriterClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FileNum As Integer      'number of the file to write to
Private Indent As Integer       'depth of indentation
Private Const Spaces = 2        'standard spaces' worth of indentation levels
Private TagStack As LinkedList  'stack of XML tags currently being written
Private TagHasData As Boolean   'tag has more data than only attributes

Public Sub BeginXML(XMLFile As Integer)
'
' Name:         BeginXML
' Paramaters:   XMLFile     number of the file we're writing to
' Description:  Begin writing an XML file.
'

    If FileNum < 0 Then 'as long as we're not already writing a file
    
        FileNum = XMLFile
        Indent = 0
        TagStack.Clear
        TagHasData = False
        Print #FileNum, "<?xml version=""1.0""?>"
        
    Else                'problem: already writing a file
    
        Err.Raise vbObjectError + 1, "XMLWriterClass", "BeginXML: Already writing a file"
    
    End If

End Sub

Public Sub BeginTag(Tag As String)
'
' Name:         BeginTag
' Parameters:   Tag     XML tag to start writing
' Description:  Begin writing an XML tag.  Add ">" to the previous tag if need be and
'               add indentation.
'

    If Not TagStack.IsEmpty Then    'don't do this the first time
        If Not TagHasData Then Print #FileNum, ">"
        Indent = Indent + Spaces
    End If
    
    Print #FileNum, Spc(Indent);
    Print #FileNum, "<" & Tag;
    TagStack.Append Tag
    TagHasData = False
    
End Sub

Public Sub EndTag()
'
' Name:         EndTag
' Description:  Complete the current tag.
'

    Dim Tag As String
    
    TagStack.Last
    Tag = TagStack.Item
    TagStack.Remove
    
    If TagHasData Then
        'other data has been added
        Print #FileNum, Spc(Indent);
        Print #FileNum, "</" & Tag & ">"
    Else
        'no data (attributes only) has been added
        Print #FileNum, "/>"
    End If
    
    Indent = Indent - Spaces
    TagHasData = True

End Sub

Public Sub WriteAttribute(Attrib As String, Value As Variant, Optional Omit = "")
'
' Name:         WriteAttribute
' Parameters:   Attrib      attribute name
'               Value       attribute value
' Description:  Add an attribute to a tag, unless it matches the Omit value--
'               in which case, add nothing.
'

    If Not TagHasData Then  'still writing attributes
    
        Dim SVal As String
        
        If Value <> Omit Then
            If TypeName(Value) = "Boolean" Then
                SVal = IIf(Value, "yes", "no")
            Else
                SVal = FormatForXML(CStr(Value))
            End If
            Print #FileNum, " " & Attrib & "=""" & SVal & """";
        End If
        
    Else                'problem: already writing complex data
        Err.Raise vbObjectError + 2, "XMLWriter", "WriteAttribute: Can't add attributes when " _
                  & "complex data has already been written"
    End If
    
End Sub

Public Sub WriteCDataTag(Tag As String, Data As String)
'
' Name:         WriteCDataTag
' Parameters:   Tag         the tag to write
'               Data        the CData contents to write
' Description:  If Data is not empty, write the given tag.
'

    If Data <> "" Then
        BeginTag Tag
        WriteStringData Data, True
        EndTag
    End If

End Sub

Public Sub WriteStringData(Data As String, Optional CData = True)
'
' Name:         WriteStringData
' Parameters:   Data        the string data to write
'               CData       whether it is raw
'

    If Not TagHasData Then Print #FileNum, ">"
    TagHasData = True
    If Data <> "" Then
    
        Print #FileNum, Spc(Indent + Spaces);
        If CData Then
            If HideSTFromFile Then
                Print #FileNum, "<![CDATA[" & Replace(Game.STFilter(Data), "]]>", "]] >") & "]]>"
            Else
                Print #FileNum, "<![CDATA[" & Replace(Data, "]]>", "]] >") & "]]>"
            End If
        Else
            Print #FileNum, FormatForXML(Data)
        End If
    
    End If
    
End Sub

Private Function FormatForXML(ByRef SVal As String) As String
'
' Name:         FormatForXML
' Parameters:   SVal        Raw string data to format for XML
' Description:  Replaces characters illegal in XML with XML symbols.
' Returns:      The formatted string.
'

    FormatForXML = SVal
    If HideSTFromFile Then FormatForXML = Game.STFilter(FormatForXML)
    FormatForXML = Replace(FormatForXML, "&", "&amp;")
    FormatForXML = Replace(FormatForXML, "<", "&lt;")
    FormatForXML = Replace(FormatForXML, ">", "&gt;")
    FormatForXML = Replace(FormatForXML, """", "&quot;")

End Function

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Create objects that compose the XMLWriterClass.
'

    FileNum = -1
    Set TagStack = New LinkedList

End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Destroy objects that compose the XMLWriterClass.
'

    Set TagStack = Nothing

End Sub
