VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlotNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DevDate As Date
Public Development As String

Public Causes As CauseEffectList
Public Effects As CauseEffectList

Public NextNode As PlotNode
Public PrevNode As PlotNode
Public Parent As PlotClass

Public Function GetAPRType() As APRType
'
' Name:         GetAPRType
' Description:  Return whether this is an action, plot or rumor node.
'               Needed to simulate polymorphism.
'

    GetAPRType = aprPlot
    
End Function

Public Function IsComplete() As Boolean
'
' Name:         IsComplete
' Description:  Return TRUE iff the action and result are both filled.
'

    IsComplete = Not (Development = "")

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
    
        .BeginTag "development"
    
            .WriteAttribute "date", DevDate
            
            .WriteCDataTag "event", Development
            
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

    If XML.Tag = "development" Then
        
        With XML
        
            DevDate = .GetDate("date")
            
            Do Until .Tag = "development" And .Terminal
                .ReadTag
                Select Case .Tag
                    Case "linklist"
                        Effects.InputFromFile XML, Version
                    Case "event"
                        If Not .Terminal Then
                            .ReadTag
                            Development = .GetTagText
                        End If
                End Select
            Loop
            
        End With
        
    Else
        Err.Raise vbObjectError + 8, "PlotNode", _
                  "Wrong file input to a plot development"
    End If

End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this object to a binary file.
'

    Put #FileNum, , DevDate
    
    Effects.OutputToBinary FileNum
    
    PutStrB FileNum, Development

End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this object from a binary file.
'

    Get #FileNum, , DevDate
    
    Effects.InputFromBinary FileNum, Version
    
    GetStrB FileNum, Development
 
End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Destroy the cause and effect lists.
'

    Set Causes = Nothing
    Set Effects = Nothing

End Sub

