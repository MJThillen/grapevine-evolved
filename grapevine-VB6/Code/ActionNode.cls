VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ActionNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Action As String
Public Result As String

Public Level As Integer
Public Unused As Integer
Public Total As Integer
Public Growth As Integer

Public Causes As CauseEffectList
Public Effects As CauseEffectList

Public NextNode As ActionNode
Public PrevNode As ActionNode
Public Parent As ActionClass

Public Function GetAPRType() As APRType
'
' Name:         GetAPRType
' Description:  Return whether this is an action, plot or rumor node.
'               Needed to simulate polymorphism.
'

    GetAPRType = aprAction

End Function

Public Function IsComplete() As Boolean
'
' Name:         IsComplete
' Description:  Return TRUE iff the action and result are both filled.
'

    IsComplete = Not (Action = "" Or Result = "")

End Function

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Create the needed cause and effect lists.
'

    Set Causes = New CauseEffectList
    Set Causes.Parent = Me
    Causes.IsEffect = False
    
    Set Effects = New CauseEffectList
    Set Effects.Parent = Me
    Effects.IsEffect = True
    
End Sub

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write the object to an XML file.
'

    With XML
    
        .BeginTag "subaction"
    
            .WriteAttribute "name", Name
            .WriteAttribute "level", Level
            .WriteAttribute "unused", Unused
            .WriteAttribute "total", Total
            .WriteAttribute "growth", Growth, 0
            
            .WriteCDataTag "act", Action
            .WriteCDataTag "results", Result
            
            Effects.OutputToFile XML
                
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

    If XML.Tag = "subaction" Then
        
        With XML
        
            Name = .GetString("name")
            Level = .GetInteger("level")
            Unused = .GetInteger("unused")
            Total = .GetInteger("total")
            Growth = .GetInteger("growth")
            
            Do Until .Tag = "subaction" And .Terminal
                .ReadTag
                Select Case .Tag
                    Case "linklist"
                        Effects.InputFromFile XML, Version
                    Case "act"
                        If Not .Terminal Then
                            .ReadTag
                            Action = .GetTagText
                        End If
                    Case "results"
                        If Not .Terminal Then
                            .ReadTag
                            Result = .GetTagText
                        End If
                End Select
            Loop
            
        End With
        
    Else
        Err.Raise vbObjectError + 8, "ActionNode", _
                  "Wrong file input to a subaction"
    End If

End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this object to a binary file.
'

    PutStrB FileNum, Name
    Put #FileNum, , Level
    Put #FileNum, , Unused
    Put #FileNum, , Total
    Put #FileNum, , Growth
    
    Effects.OutputToBinary FileNum
    
    PutStrB FileNum, Action
    PutStrB FileNum, Result

End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this object from a binary file.
'

    GetStrB FileNum, Name
    Get #FileNum, , Level
    Get #FileNum, , Unused
    Get #FileNum, , Total
    Get #FileNum, , Growth
    
    Effects.InputFromBinary FileNum, Version
    
    GetStrB FileNum, Action
    GetStrB FileNum, Result
 
End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Destroy the cause and effect lists.
'

    Set Causes = Nothing
    Set Effects = Nothing

End Sub
