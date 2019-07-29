VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CauseEffectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum APRType
    aprNone = 0
    aprAction = 1
    aprPlot = 2
    aprRumor = 4
End Enum

Public Link As CauseEffectNode              'Current node of the list

Public IsEffect As Boolean                  'Whether this is an effect list (TRUE) or a cause list (FALSE)
Public Parent As Object                     'The ActionNode, PlotNode or RumorNode parenting this list

Private FirstNode As CauseEffectNode        'First node of the list
Private LastNode As CauseEffectNode         'Last node of the list

Private NodeCount As Integer                'number of nodes in the list
Private CurPlace As Integer                 'current numerical location in the list

Public Function Count() As Integer
'
' Name:         Count
' Description:  Returns the number of items in the list
' Returns:      the number of items in the list
'

    Count = NodeCount

End Function

Public Sub PopulateList(List As ListBox, HideDate As Date)
'
' Name:         PopulateList
' Parameters:   List        List Box to populate
'               HideDate    Date to shorten in the list
' Description:  Populate the given list box with the abbreviated contents of this list.
'

    Dim I As Integer
    
    I = 0
    List.Clear
    List.ToolTipText = ""
    First
    Do Until Off
        List.AddItem Link.ShortDesc(HideDate)
        List.ItemData(List.NewIndex) = I
        I = I + 1
        MoveNext
    Loop

End Sub

Public Function GetPlace() As Integer
'
' Name:         GetPlace
' Description:  Returns the location of the cursor in the list.
'
    GetPlace = CurPlace

End Function

Public Sub MoveToPlace(ByVal Place As Integer)
'
' Name:         MoveToPlace
' Parameters:   Place           the position in the list to move to
' Description:  Move the Link to the given number position in the list.
'               Like listbox listindexes, this is 0-based.
'
    
    Set Link = FirstNode
    CurPlace = 0
    Do Until (CurPlace >= Place Or Link Is Nothing)
        Set Link = Link.NextNode
        CurPlace = CurPlace + 1
    Loop

End Sub

Public Sub RemoveNode(Node As Object)
'
' Name:         RemoveNode
' Parameters:   Node            An ActionNode, PlotNode or RumorNode
' Description:  Move to the link corresponding to the given node, and remove it.
'               Do not ensure consistency.
'

    Dim APR As APRType
    Dim When As Date
    Dim Item As String
    Dim Subitem As String
    Dim Act As ActionNode
    Dim Plot As PlotNode
    Dim Rumor As RumorNode
    
    APR = Node.GetAPRType

    First
    Do Until Off
        If Link.Target = APR Then
            Select Case APR
                Case aprAction
                    Set Act = Node
                    When = Act.Parent.ActDate
                    Item = Act.Parent.CharName
                    Subitem = Act.Name
                Case aprPlot
                    Set Plot = Node
                    When = Plot.DevDate
                    Item = Plot.Parent.Name
                    Subitem = ""
                Case aprRumor
                    Set Rumor = Node
                    When = Rumor.Parent.RumorDate
                    Item = Rumor.Parent.Title
                    Subitem = CStr(Rumor.Level)
            End Select
            If Link.When = When And Link.Item = Item And Link.Subitem = Subitem Then Exit Do
        End If
        MoveNext
    Loop

    If Not Off Then Remove

End Sub

Public Sub AddLink(APR As APRType, When As Date, Item As String, Subitem As String)
'
' Name:         AddLink
' Parameters:   APR         whether the object on the other end of the link is an action, plot or rumor
'               When        the date of the link
'               Item        the name of the link
'               Subitem     the topic of the link
' Description:  Add the appropriate links to and from the referenced item.
'

    Dim Linked As Object

    Append APR, When, Item, Subitem
    Set Linked = Game.APREngine.GetLink(Link)

    If Not Linked Is Nothing Then

        On Error Resume Next
        If IsEffect Then
            Linked.Causes.AppendNode Me.Parent
        Else
            Linked.Effects.AppendNode Me.Parent
        End If
        On Error GoTo 0
            
        Linked.Parent.LastModified = Now
    
    End If
    
End Sub

Public Sub RemoveLink()
'
' Name:         RemoveLink
' Description:  Remove the current link, and the associated link.
'

    If Not Link Is Nothing Then

        Dim Linked As Object
                    
        Set Linked = Game.APREngine.GetLink(Link)
        
        If Not Linked Is Nothing Then
            On Error Resume Next
            If IsEffect Then
                Linked.Causes.RemoveNode Me.Parent
            Else
                Linked.Effects.RemoveNode Me.Parent
            End If
            On Error GoTo 0
            Linked.Parent.LastModified = Now
        End If
            
        Remove
        
    End If
    
End Sub

Public Sub Clear()
'
' Name:         Clear
' Description:  Clear the links, but not associated links.
'

    First
    Do Until Link Is Nothing
        Remove
    Loop

End Sub

Public Sub ClearLinks()
'
' Name:         ClearLinks
' Description:  Clear the links and all associated links.
'

    First
    Do Until Link Is Nothing
        RemoveLink
    Loop

End Sub

Private Sub Append(APR As APRType, When As Date, Item As String, Subitem As String)
'
' Name:         Append
' Parameters:   APR         whether this is Action, Plot or Rumor
'               When        the date of the link
'               Item        the name of the link
'               Subitem     the topic of the link
' Description:  Add an link to the end of the list.
'

    Set Link = New CauseEffectNode
    
    Link.Target = APR
    Link.When = When
    Link.Item = Item
    Link.Subitem = Subitem
    
    Set Link.NextNode = Nothing
    Set Link.PrevNode = LastNode

    If LastNode Is Nothing Then
        Set FirstNode = Link
    Else
        Set LastNode.NextNode = Link
    End If

    Set LastNode = Link
    NodeCount = NodeCount + 1

End Sub

Public Sub AppendNode(Node As Object)
'
' Name:         AppendNode
' Parameters:   Node        the ActionNode, PlotNode or RumorNode to link
' Description:  Append a link to the given node to the list.  Do not ensure consistency.
'

    Select Case Node.GetAPRType
        Case aprAction
            Dim Act As ActionNode
            Set Act = Node
            Append aprAction, Act.Parent.ActDate, Act.Parent.CharName, Act.Name
        Case aprPlot
            Dim Plot As PlotNode
            Set Plot = Node
            Append aprPlot, Plot.DevDate, Plot.Parent.Name, ""
        Case aprRumor
            Dim Rumor As RumorNode
            Set Rumor = Node
            Append aprRumor, Rumor.Parent.RumorDate, Rumor.Parent.Title, CStr(Rumor.Level)
    End Select
    
End Sub

Public Sub Remove()
'
' Name:         Remove
' Description:  Totally remove the item from the current Link
'               position.
'

    Dim NextLink As CauseEffectNode

    If Not Link Is Nothing Then
        
        If Link Is FirstNode Then
            Set FirstNode = Link.NextNode
        Else
            Set Link.PrevNode.NextNode = Link.NextNode
        End If

        If Link Is LastNode Then
            Set LastNode = Link.PrevNode
        Else
            Set Link.NextNode.PrevNode = Link.PrevNode
        End If

        Set NextLink = Link.NextNode
        Set Link = Nothing
        Set Link = NextLink

        NodeCount = NodeCount - 1

    End If

End Sub

Public Function IsEmpty() As Boolean
'
' Name:         IsEmpty
' Description:  Tells whether the list is empty or not.
' Returns:      TRUE if the list is empty, FALSE otherwise.
'

    IsEmpty = (NodeCount = 0)

End Function

Public Function Off() As Boolean
'
' Name:         Off
' Description:  Tells whether the Link has moved off the list.
' Returns:      TRUE if the Link is off the list, FALSE otherwise.
'

    Off = (Link Is Nothing)

End Function

Public Sub MoveNext()
'
' Name:         MoveNext
' Description:  Move the Link to the next item.
'

    If Not Link Is Nothing Then
        Set Link = Link.NextNode
        CurPlace = CurPlace + 1
    End If

End Sub
Public Sub MovePrevious()
'
' Name:         MovePrevious
' Description:  Move the Link to the previous item.
'

    If Not Link Is Nothing Then
        Set Link = Link.PrevNode
        CurPlace = CurPlace - 1
    End If

End Sub

Public Sub First()
'
' Name:         First
' Description:  Move the Link to the first item.
'

    Set Link = FirstNode
    CurPlace = 0
    
End Sub

Public Sub Last()
'
' Name:         Last
' Description:  Move the Link to the last item.
'

    Set Link = LastNode
    CurPlace = NodeCount - 1
    
End Sub


Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write the object to an XML file.
'

    Dim MemCursor As CauseEffectNode

    Set MemCursor = Link

    With XML
    
        .BeginTag "linklist"
    
        First
        Do Until Link Is Nothing
        
            Select Case Link.Target
                Case aprAction: .BeginTag "action"
                Case aprPlot:   .BeginTag "plot"
                Case aprRumor:  .BeginTag "rumor"
            End Select
            
            .WriteAttribute "date", Link.When
            .WriteAttribute "name", Link.Item
            .WriteAttribute "item", Link.Subitem
            .EndTag
            
            MoveNext
        Loop
    
        .EndTag
        
    End With

    Set Link = MemCursor

End Sub

Public Sub InputFromFile(XML As XMLReaderClass, Version As Double)
'
' Name:         InputFromFile
' Parameters:   XML         an XMLReaderClass that just read the opening
'                           tag of this object
'               Version     version of the file format
' Description:  Read the object from an XML file.
'

    Dim APR As APRType

    If XML.Tag = "linklist" Then
        
        Clear
        
        With XML
        
            Do Until .Tag = "linklist" And .Terminal
                
                APR = aprNone
                .ReadTag
                Select Case .Tag
                    Case "action":  APR = aprAction
                    Case "plot":    APR = aprPlot
                    Case "rumor":   APR = aprRumor
                End Select
            
                If APR <> aprNone Then
                    Append APR, .GetDate("date"), .GetString("name"), .GetString("item")
                End If
                
            Loop
            
        End With
        
    Else
        Err.Raise vbObjectError + 8, "CauseEffectList", _
                  "Wrong file input to a Cause/Effect List"
    End If

End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this object to a binary file.
'

    Dim MemCursor As CauseEffectNode

    Set MemCursor = Link

    Put #FileNum, , NodeCount
    First
    Do Until Off
        Put #FileNum, , Link.Target
        Put #FileNum, , Link.When
        PutStrB FileNum, Link.Item
        PutStrB FileNum, Link.Subitem
        MoveNext
    Loop

    Set Link = MemCursor

End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this object from a binary file.
'

    Dim I As Integer
    Dim APR As APRType
    Dim When As Date
    Dim Item As String
    Dim Subitem As String
    
    Clear
    
    Get #FileNum, , I
    Do Until I = 0
        Get #FileNum, , APR
        Get #FileNum, , When
        GetStrB FileNum, Item
        GetStrB FileNum, Subitem
        Append APR, When, Item, Subitem
        I = I - 1
    Loop
 
End Sub

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Initialize an empty list.  VB calls this method
'               automatically.
'

    Set Link = Nothing
    Set FirstNode = Nothing
    Set LastNode = Nothing
    NodeCount = 0
    CurPlace = 0
    
End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  If a list is destroyed, clear it first.  VB calls
'               this method automatically.
'

    If Not Game Is Nothing Then
        ClearLinks
    Else
        Clear
    End If

End Sub

