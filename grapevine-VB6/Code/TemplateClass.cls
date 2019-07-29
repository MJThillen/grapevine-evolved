VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public IsCharacterSheet As Boolean

Private FileName(MIN_OUTFORMAT To MAX_OUTFORMAT) As String
Private Subject(MIN_OUTFORMAT To MAX_OUTFORMAT) As Long

Public Function GetFilename(OutFormat As OutputFormatType) As String
'
' Name:         GetFilename
' Parameters:   OutFormat       a file format (text, RTF or HTML)
' Description:  Return the filename associated with this template and the given format.
'

    GetFilename = FileName(OutFormat)

End Function

Public Sub SetFilename(OutFormat As OutputFormatType, TName As String)
'
' Name:         SetFilename
' Parameters:   OutFormat       a file format (text, RTF or HTML)
' Description:  Set the filename associated with this template and the given format.
'

    FileName(OutFormat) = TName
    Subject(OutFormat) = ooUnknown
    
End Sub

Public Function GetSubject(OutFormat As OutputFormatType) As Long
'
' Name:         GetSubject
' Parameters:   OutFormat       a file format (text, RTF or HTML)
' Description:  Get the subject of the template for the given format.  If the subject
'               is currently unknown, read the subject from the file itself.
'

    If FileName(OutFormat) = "" Then
        
        Subject(OutFormat) = ooFileError

    ElseIf Subject(OutFormat) = ooUnknown Then
    
        Dim TempFile As String
        TempFile = FindFile(FileName(OutFormat))
        Subject(OutFormat) = ooFileError
        
        If TempFile <> "" Then
            
            On Error Resume Next
            Subject(OutFormat) = OutputEngine.GetTemplateSubject(TempFile)
            On Error GoTo 0
            
        End If
        
    End If

    GetSubject = Subject(OutFormat)

End Function

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write the object to an XML file.
'

    With XML
    
        .BeginTag "template"
    
            .WriteAttribute "name", Name
            .WriteAttribute "sheet", IsCharacterSheet, False
            .WriteAttribute "text", FileName(ofText)
            .WriteAttribute "rtf", FileName(ofRTF)
            .WriteAttribute "html", FileName(ofHTML)

        .EndTag
        
    End With

End Sub

Public Sub InputFromFile(XML As XMLReaderClass, Version As Double)
'
' Name:         InputFromFile
' Parameters:   XML         an XMLReaderClass that just read the opening
'                           tag of this object
'               Version     version of the file format
' Description:  Read the object from an XML file.
'

    If XML.Tag = "template" Then
        
        With XML
        
            Name = .GetString("name")
            IsCharacterSheet = .GetBoolean("sheet")
            FileName(ofText) = .GetString("text")
            FileName(ofRTF) = .GetString("rtf")
            FileName(ofHTML) = .GetString("html")
            
            Do Until .Tag = "template" And .Terminal
                .ReadTag
            Loop
            
        End With
        
        Subject(ofText) = ooUnknown
        Subject(ofRTF) = ooUnknown
        Subject(ofHTML) = ooUnknown
        
    Else
        Err.Raise vbObjectError + 8, "Templates", _
                  "Wrong file input to template descriptions"
    End If

End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this object to a binary file.
'

    PutStrB FileNum, Name
    Put #FileNum, , IsCharacterSheet
    PutStrB FileNum, FileName(ofText)
    PutStrB FileNum, FileName(ofRTF)
    PutStrB FileNum, FileName(ofHTML)
    
End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this object from a binary file.
'

    GetStrB FileNum, Name
    Get #FileNum, , IsCharacterSheet
    GetStrB FileNum, FileName(ofText)
    GetStrB FileNum, FileName(ofRTF)
    GetStrB FileNum, FileName(ofHTML)
     
End Sub
