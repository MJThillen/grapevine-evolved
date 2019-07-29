VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlotClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public StartDate As Date
Public EndDate As Date
Public Outline As String
Public Narrator As String

Public LastModified As Date

Public PlotDev As PlotNode
Public CastList As LinkedTraitList

Private FirstNode As PlotNode
Private LastNode As PlotNode
Private NodeCount As Integer

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Create the list and default Plot node.
'

    StartDate = 0
    EndDate = 0
    Set FirstNode = Nothing
    Set LastNode = Nothing
    Set PlotDev = Nothing
    NodeCount = 0
    LastModified = Now
    Set CastList = New LinkedTraitList
    CastList.Initialize "Key Characters", True, False, False, ldSimple
    
End Sub

Public Function OutputID() As Integer
'
' Name:         OutputID
' Description:  Return an ID that tells the OutputEngineClass how to manipulate this object
'
    OutputID = oidPlot

End Function

Public Sub InitializeForOutput()
'
' Name:         InitializeForOutput
' Description:  Position the object in a starting state, ready for output.
'

    With OutputEngine
        If .AscendDate Then
            First
            If .StartDate <> 0 Then
                Do Until Off
                    If PlotDev.DevDate >= .StartDate Then Exit Do
                    MoveNext
                Loop
            End If
        Else
            Last
            If .EndDate <> 0 Then
                Do Until Off
                    If PlotDev.DevDate <= .EndDate + #11:59:59 PM# Then Exit Do
                    MovePrevious
                Loop
            End If
        End If
    End With
    CastList.First
    
End Sub

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
        Case qkTitle:           ObjData = Name
        Case qkStatus:          ObjData = GetStatus
        Case qkStartDate:       If StartDate <> 0 Then ObjData = StartDate
        Case qkEndDate:         If EndDate <> 0 Then ObjData = EndDate
        Case qkOutline:         ObjData = Outline
        Case qkNarrator:        ObjData = Narrator
        Case qkKeyCharacters:   Set ObjData = CastList
        Case qkCount:           ObjData = Count
        Case qkDevDate:         If Not Off Then ObjData = PlotDev.DevDate
        Case qkDevelopment:     If Not Off Then ObjData = PlotDev.Development
        Case Else:              ObjData = Null
    End Select

End Sub

Public Function GetStatus(Optional NowDate As Date = 0) As String
'
' Name:         GetStatus
'               NowDate         The date to compare this plot to.
' Description:  Return if the plot is active, finished or pending.
'

    If NowDate = 0 Then
        Game.Calendar.MoveToCloseGame
        If Game.Calendar.Off Then
            NowDate = Date
        Else
            NowDate = Game.Calendar.GetGameDate
        End If
    End If
    
    If NowDate < StartDate And Not StartDate = 0 Then
        GetStatus = psPending
    Else
        If EndDate = 0 Or NowDate <= EndDate Then
            GetStatus = psActive
        Else
            GetStatus = psFinished
        End If
    End If

End Function

Public Sub Add(DevDate As Date, Development As String)
'
' Name:         Add
' Parameters:   DevDate         date of the plot development to add
'               Development     text of the plot development to add
' Description:  Create and add a new plot development.
'

    Dim NewDev As PlotNode

    Set NewDev = New PlotNode

    NewDev.DevDate = DevDate
    NewDev.Development = Development

    AddNode NewDev

End Sub

Private Sub AddNode(NewDev As PlotNode)
'
' Name:         Add
' Parameters:   NewDev
' Description:  Place a plot development node into the list.
'

    Set NewDev.Parent = Me
    
    If FirstNode Is Nothing Then
    
        Set FirstNode = NewDev
        Set LastNode = NewDev
        Set PlotDev = NewDev
    
    Else
    
        Set PlotDev = FirstNode
        Do
            If NewDev.DevDate < PlotDev.DevDate Then Exit Do
            Set PlotDev = PlotDev.NextNode
        Loop Until PlotDev Is Nothing
        
        If PlotDev Is Nothing Then
            Set NewDev.PrevNode = LastNode
            Set LastNode.NextNode = NewDev
            Set LastNode = NewDev
        Else
            Set NewDev.NextNode = PlotDev
            Set NewDev.PrevNode = PlotDev.PrevNode
            If PlotDev Is FirstNode Then
                Set FirstNode = NewDev
            Else
                Set PlotDev.PrevNode.NextNode = NewDev
            End If
            Set PlotDev.PrevNode = NewDev
        End If
    
    End If
    
    NodeCount = NodeCount + 1

End Sub

Public Sub First()
'
' Name:         First
' Description:  Set PlotDev to the first Plot in the nodelist.
'

    Set PlotDev = FirstNode

End Sub

Public Sub Last()
'
' Name:         Last
' Description:  Set PlotDev to the last Plot in the nodelist.
'

    Set PlotDev = LastNode

End Sub

Public Sub MoveNext()
'
' Name:         MoveNext
' Description:  Move to the next Plot in the list and set PlotDev to it.
'

    If Not PlotDev Is Nothing Then Set PlotDev = PlotDev.NextNode
    
End Sub

Public Sub MovePrevious()
'
' Name:         MovePrevious
' Description:  Move to the previous Plot in the list and set PlotDev to it.
'

    If Not PlotDev Is Nothing Then Set PlotDev = PlotDev.PrevNode
    
End Sub

Public Function Off() As Boolean
'
' Name:         Off
' Description:  Return TRUE iff there is no current PlotDev.
'
    Off = (PlotDev Is Nothing)

End Function

Public Sub Remove()
'
' Name:         Remove
' Description:  Remove the current PlotDev.
'
    
    If Not PlotDev Is Nothing Then
        
        Dim TempNode As PlotNode
        
        If Not PlotDev Is FirstNode Then
            Set PlotDev.PrevNode.NextNode = PlotDev.NextNode
        Else
            Set FirstNode = PlotDev.NextNode
        End If
        
        If Not PlotDev Is LastNode Then
            Set PlotDev.NextNode.PrevNode = PlotDev.PrevNode
        Else
            Set LastNode = PlotDev.PrevNode
        End If
        
        Set TempNode = PlotDev.NextNode
        Set PlotDev = Nothing
        Set PlotDev = TempNode
    
        NodeCount = NodeCount - 1
    
    End If
    
End Sub

Public Sub MoveTo(DevDate As Date)
'
' Name:         MoveTo
' Parameters:   Name        Date to move to.
' Description:  Move to the given plot date.
'

    Set PlotDev = FirstNode
    Do Until PlotDev Is Nothing
        If PlotDev.DevDate = DevDate Then Exit Do
        Set PlotDev = PlotDev.NextNode
    Loop

End Sub

Public Sub MoveToPlace(ByVal Place As Integer)
'
' Name:         MoveToPlace
' Parameters:   Place           the position in the list to move to
' Description:  Move the Link to the given number position in the list.
'               Like listbox listindexes, this is 0-based.
'
    
    Set PlotDev = FirstNode
    Do Until (Place <= 0 Or PlotDev Is Nothing)
        Set PlotDev = PlotDev.NextNode
        Place = Place - 1
    Loop

End Sub

Public Function Count() As Integer
'
' Name:         Count
' Description:  Return the number of PlotDevs in this Plot.
'
    Count = NodeCount

End Function

Public Sub Clear()
'
' Name:         Clear
' Description:  Empty this Plot of PlotDevs.
'

    First
    Do Until PlotDev Is Nothing
        Remove
    Loop

End Sub

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write the object to an XML file.
'

    Dim MemCursor As PlotNode

    Set MemCursor = PlotDev

    With XML
    
        .BeginTag "plot"
    
            .WriteAttribute "name", Name
            .WriteAttribute "start", StartDate, 0
            .WriteAttribute "end", EndDate, 0
            .WriteAttribute "narrator", Narrator
            .WriteAttribute "lastmodified", LastModified
            
            .WriteCDataTag "outline", Outline
            
            CastList.OutputToFile XML
            
            First
            Do Until Off
                PlotDev.OutputToFile XML
                MoveNext
            Loop
    
        .EndTag
        
    End With

    Set PlotDev = MemCursor
    
End Sub

Public Sub InputFromFile(XML As XMLReaderClass, Version As Double)
'
' Name:         InputFromFile
' Parameters:   XML         an XMLReaderClass that just read the opening
'                           tag of this object
'               Version     version of the file format
' Description:  Read the object from an XML file.
'

    Dim NewDev As PlotNode

    If XML.Tag = "plot" Then
        
        With XML
        
            Clear
        
            Name = .GetString("name")
            StartDate = .GetDate("start", 0)
            EndDate = .GetDate("end", 0)
            Narrator = .GetString("narrator")
            LastModified = .GetDate("lastmodified")
            
            Do Until .Tag = "plot" And .Terminal
                .ReadTag
                Select Case .Tag
                    Case "outline"
                        If Not .Terminal Then
                            .ReadTag
                            Outline = .GetTagText
                        End If
                    Case "traitlist"
                        Select Case .GetString("name")
                            Case "Key Characters":  CastList.InputFromFile XML, Version
                        End Select
                    Case "development"
                        Set NewDev = New PlotNode
                        NewDev.InputFromFile XML, Version
                        AddNode NewDev
                End Select
            Loop
            
        End With
        
    Else
        Err.Raise vbObjectError + 8, "PlotClass", _
                  "Wrong file input to a plot"
    End If

End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this object to a binary file.
'

    Dim MemCursor As PlotNode

    Set MemCursor = PlotDev

    PutStrB FileNum, Name
    Put #FileNum, , StartDate
    Put #FileNum, , EndDate
    PutStrB FileNum, Narrator
    Put #FileNum, , LastModified
    PutStrB FileNum, Outline
    
    CastList.OutputToBinary FileNum
    
    Put #FileNum, , NodeCount
    First
    Do Until Off
        PlotDev.OutputToBinary FileNum
        MoveNext
    Loop

    Set PlotDev = MemCursor

End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this object from a binary file.
'

    Dim I As Integer
    Dim NewDev As PlotNode

    GetStrB FileNum, Name
    Get #FileNum, , StartDate
    Get #FileNum, , EndDate
    If Version >= 2.399 Then GetStrB FileNum, Narrator
    Get #FileNum, , LastModified
    GetStrB FileNum, Outline
    
    If Version >= 2.399 Then CastList.InputFromBinary FileNum, Version
    
    Get #FileNum, , I
    Do Until I = 0
        Set NewDev = New PlotNode
        NewDev.InputFromBinary FileNum, Version
        AddNode NewDev
        I = I - 1
    Loop
 
End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Clean up.
'
    
    Set CastList = Nothing
    Clear
    
End Sub

