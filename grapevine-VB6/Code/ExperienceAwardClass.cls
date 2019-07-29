VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExperienceAwardClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public ChangeType As ExperienceChangeType
Public Change As Single
Public Reason As String
Public XP As Boolean

Public Function ChangeTypeText() As String
'
' Name:         ChangeTypeText
' Description:  Return a string describing the effect of the award.
' Returns:      as above.
'

    Select Case ChangeType
        Case ecEarned
            ChangeTypeText = "earn " & CStr(Change)
        Case ecSpent
            ChangeTypeText = "spend " & CStr(Change)
        Case ecDeducted
            ChangeTypeText = "deduct " & CStr(Change)
        Case ecUnspent
            ChangeTypeText = "unspend " & CStr(Change)
        Case ecSetEarned
            ChangeTypeText = "set earned to " & CStr(Change)
        Case ecSetUnspent
            ChangeTypeText = "set unspent to " & CStr(Change)
        Case ecComment
            ChangeTypeText = "comment"
    End Select

End Function

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write the object to an XML file.
'

    With XML
    
        .BeginTag "award"
    
            .WriteAttribute "use", IIf(XP, "XP", "PP")
            .WriteAttribute "name", Name
            .WriteAttribute "type", ChangeType
            .WriteAttribute "change", Change
            .WriteAttribute "reason", Reason
            
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

    If XML.Tag = "award" Then
        
        With XML
        
            XP = (.GetString("use", "XP") = "XP")
            Name = .GetString("name")
            ChangeType = .GetLong("type", ecEarned)
            Change = .GetSingle("change")
            Reason = .GetString("reason")
            
            Do Until .Tag = "award" And .Terminal
                .ReadTag
            Loop
            
        End With
        
    Else
        Err.Raise vbObjectError + 8, "XP/PP Awards", _
                  "Wrong file input to XP/PP award"
    End If

End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this object to a binary file.
'

    Put #FileNum, , XP
    PutStrB FileNum, Name
    Put #FileNum, , ChangeType
    Put #FileNum, , Change
    PutStrB FileNum, Reason
    
End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this object from a binary file.
'

    Get #FileNum, , XP
    GetStrB FileNum, Name
    Get #FileNum, , ChangeType
    Get #FileNum, , Change
    GetStrB FileNum, Reason
     
End Sub

