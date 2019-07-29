VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RumorNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Level As Integer
Public Rumor As String

Public Causes As CauseEffectList
Public RecipientSet As StringSet

Public NextNode As RumorNode
Public PrevNode As RumorNode
Public Parent As RumorClass

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Create the needed cause and effect lists.
'

    Set Causes = New CauseEffectList
    Set Causes.Parent = Me
    Causes.IsEffect = False

    Set RecipientSet = New StringSet

End Sub

Public Function GetAPRType() As APRType
'
' Name:         GetAPRType
' Description:  Return whether this is an action, plot or rumor node.
'               Needed to simulate polymorphism.
'

    GetAPRType = aprRumor

End Function

Public Function IsComplete() As Boolean
'
' Name:         IsComplete
' Description:  Return TRUE iff the action and result are both filled.
'

    IsComplete = Not (Rumor = "")

End Function

Public Function OutputID() As Integer
'
' Name:         OutputID
' Description:  Return an ID that tells the OutputEngineClass how to manipulate this object
'
    OutputID = oidNone

End Function

Public Sub GetValue(Key As String, ObjData As Variant)
'
' Name:         GetValue
' Parameters:   Key         a key (string) to the value to retreive
'               ObjData     the returned value
' Description:  Use the given key to return data from the object.  Keys are defined
'               in the PublicQueryKeys module.
'

    ObjData = ""
    Select Case Key
        Case qkTitle:           ObjData = Parent.Title
        Case qkDate:            ObjData = Parent.RumorDate
        Case qkRumor:           ObjData = Rumor
        Case qkLevel:           If Parent.Category = rtInfluence Then ObjData = Level
        Case qkType
            If Parent.Category = rtInfluence Then
                ObjData = "Level " & CStr(Level) & " " & Parent.Title
            Else
                ObjData = Parent.Title
            End If
        Case Else:              ObjData = Null
    End Select

End Sub

Public Sub InitializeForOutput()
'
' Name:         InitializeForOutput
' Description:  Position the object in a starting state, ready for output.
'

    'Do nothing...!
    
End Sub

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write the object to an XML file.
'

    With XML
    
        .BeginTag "subrumor"
    
            .WriteAttribute "level", Level, 0
            
            .WriteCDataTag "body", Rumor
            
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

    If XML.Tag = "subrumor" Then
        
        With XML
        
            Level = .GetInteger("level")
            
            Do Until .Tag = "subrumor" And .Terminal
                .ReadTag
                Select Case .Tag
                    Case "body"
                        If Not .Terminal Then
                            .ReadTag
                            Rumor = .GetTagText
                        End If
                End Select
            Loop
            
        End With
        
    Else
        Err.Raise vbObjectError + 8, "RumorNode", _
                  "Wrong file input to a subrumor"
    End If

End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this object to a binary file.
'

    Put #FileNum, , Level
    
    PutStrB FileNum, Rumor

End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this object from a binary file.
'

    Get #FileNum, , Level
    
    GetStrB FileNum, Rumor
 
End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Destroy the cause and effect lists.
'

    Set Causes = Nothing
    Set RecipientSet = Nothing
    
End Sub

