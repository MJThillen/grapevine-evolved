VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ActionClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CharName As String
Public ActDate As Date
Public SubAction As ActionNode
Public Done As Boolean

Public LastModified As Date

Private FirstNode As ActionNode
Private LastNode As ActionNode
Private NodeCount As Integer

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Create the list and default Action node.
'

    Add BasicSubactionName, 0, Game.APREngine.PersonalActions, Game.APREngine.PersonalActions, 0
    LastModified = Now
    
End Sub

Public Sub Initialize(CName As String, ADate As Date)
'
' Name:         Initialize
' Parameters:   CName       Character whose action this is
'               ADate       Date of the action
' Description:  Initialize this instance as a brand-shiny-new action class.
'

    Dim PrevAct As ActionClass
    
    CharName = CName
    ActDate = ADate
    
    MoveTo BasicSubactionName
    If Off Then
        Add BasicSubactionName, 0, Game.APREngine.PersonalActions, Game.APREngine.PersonalActions, 0
    End If
        
    Game.APREngine.MoveToEarlierTitle ActionList, CharName, ActDate
    If Not ActionList.Off Then
        Set PrevAct = ActionList.Item
        PrevAct.MoveTo BasicSubactionName
        If Not PrevAct.Off Then
            If Game.APREngine.CarryUnused Then SubAction.Unused = PrevAct.SubAction.Unused
            SubAction.Growth = PrevAct.SubAction.Growth
        End If
    End If

    If Game.APREngine.AddCommon Then AddCommonActions

End Sub

Public Function Name() As String
'
' Name:         Name
' Description:  Return an object name for this action, combining the date and the character name.
'

    Name = Format(ActDate, "Short Date") & " " & CharName

End Function

Public Function OutputID() As Integer
'
' Name:         OutputID
' Description:  Return an ID that tells the OutputEngineClass how to manipulate this object
'
    OutputID = oidAction

End Function

Public Sub InitializeForOutput()
'
' Name:         InitializeForOutput
' Description:  Position the object in a starting state, ready for output.
'

    First

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
        Case qkName:            ObjData = CharName
        Case qkDate:            ObjData = ActDate
        Case qkCount:           ObjData = Count
        Case qkType:            If Not Off Then ObjData = SubAction.Name
        Case qkLevel:           If Not Off Then ObjData = SubAction.Level
        Case qkUnused:          If Not Off Then ObjData = SubAction.Unused
        Case qkTotal:           If Not Off Then ObjData = SubAction.Total
        Case qkGrowth:          If Not Off Then ObjData = SubAction.Growth
        Case qkAction:          If Not Off Then ObjData = SubAction.Action
        Case qkResult:          If Not Off Then ObjData = SubAction.Result
        Case Else:              ObjData = Null
    End Select

End Sub

Public Sub AddCommonActions()
'
' Name:         AddCommonActions
' Description:  Add the common subactions to this action:  the character's Influence
'               actions and appropriate Background actions.
'

    Dim PrevAct As ActionClass
    
    Game.APREngine.MoveToEarlierTitle ActionList, CharName, ActDate
    If Not ActionList.Off Then Set PrevAct = ActionList.Item

    MoveTo BasicSubactionName
    If Off Then
        Add BasicSubactionName, 0, Game.APREngine.PersonalActions, Game.APREngine.PersonalActions, 0
    End If
    
    CharacterList.MoveTo CharName
    If Not CharacterList.Off Then
    
        Dim TraitList As LinkedTraitList
        Dim CharData As Variant
        Dim ActValue As Integer
        Dim PrevGrowth As Integer
        Dim PrevUnused As Integer
        
        CharacterList.Item.GetValue qkInfluences, CharData
        If Not IsNull(CharData) Then
        
            Set TraitList = CharData
            TraitList.First
            Do Until TraitList.Off
            
                MoveTo TraitList.Trait.Name
                If Off Then
                
                    ActValue = 2 * TraitList.Trait.Number
                    PrevGrowth = 0
                
                    With Game.APREngine.ActionsPerLevel
                        .MoveTo CStr(TraitList.Trait.Total)
                        If Not .Off Then
                            ActValue = .Trait.Number
                        End If
                    End With
                    
                    PrevUnused = ActValue
                    
                    If Not PrevAct Is Nothing Then
                        PrevAct.MoveTo TraitList.Trait.Name
                        If Not PrevAct.Off Then
                            If Game.APREngine.CarryUnused Then PrevUnused = PrevAct.SubAction.Unused
                            PrevGrowth = PrevAct.SubAction.Growth
                        End If
                    End If
                    
                    Add TraitList.Trait.Name, TraitList.Trait.Number, ActValue, PrevUnused, PrevGrowth
                    
                End If
                TraitList.MoveNext
                
            Loop
        
        End If
    
        CharacterList.Item.GetValue qkBackgrounds, CharData
        If Not IsNull(CharData) Then
        
            Set TraitList = CharData
            TraitList.First
            Do Until TraitList.Off
            
                Game.APREngine.BackgroundActions.MoveTo TraitList.Trait.Name
                If Not Game.APREngine.BackgroundActions.Off Then
            
                    MoveTo TraitList.Trait.Name
                    If Off Then
                    
                        ActValue = 2 * TraitList.Trait.Number
                        PrevGrowth = 0
                    
                        With Game.APREngine.ActionsPerLevel
                            .MoveTo CStr(TraitList.Trait.Total)
                            If Not .Off Then
                                ActValue = .Trait.Number
                            End If
                        End With
                        
                        PrevUnused = ActValue
                        
                        If Not PrevAct Is Nothing Then
                            PrevAct.MoveTo TraitList.Trait.Name
                            If Not PrevAct.Off Then
                                If Game.APREngine.CarryUnused Then PrevUnused = PrevAct.SubAction.Unused
                                PrevGrowth = PrevAct.SubAction.Growth
                            End If
                        End If
                        
                        Add TraitList.Trait.Name, TraitList.Trait.Number, ActValue, PrevUnused, PrevGrowth
                        
                    End If
                    
                End If
                TraitList.MoveNext
                
            Loop
        
        End If
    
    End If

End Sub

Public Sub IfDoneSetDone()
'
' Name:         IfDoneSetDone
' Description:  See if all subactions are filled with actions and results;
'               if so, set Done=True; if not, set Done=False.
'

    Done = True
    
    First
    Do Until SubAction Is Nothing
        Done = Done And (SubAction.Action <> "")
        Done = Done And (SubAction.Result <> "")
        MoveNext
    Loop

End Sub

Public Sub Add(ActName As String, Level As Integer, Total As Integer, Unused As Integer, Growth As Integer)
'
' Name:         Add
' Parameters:   ActName         Name of the subaction to add (usu. "Street", "Contacts", etc.)
'               Level           Associated level of the subaction
'               Total           Total action points to spend
'               Unused       Unused action points to spend
'               Growth          Accumulated action points from last week
' Description:  Add a new SubAction.
'

    Set SubAction = New ActionNode

    SubAction.Name = ActName
    SubAction.Level = Level
    SubAction.Total = Total
    SubAction.Unused = Unused
    SubAction.Growth = Growth
    Set SubAction.Parent = Me
    
    Set SubAction.PrevNode = LastNode
    If Not LastNode Is Nothing Then
        Set LastNode.NextNode = SubAction
    End If
    Set LastNode = SubAction
    If FirstNode Is Nothing Then Set FirstNode = SubAction
    
    NodeCount = NodeCount + 1

End Sub

Public Sub First()
'
' Name:         First
' Description:  Set SubAction to the first action in the nodelist.
'

    Set SubAction = FirstNode

End Sub

Public Sub MoveNext()
'
' Name:         MoveNext
' Description:  Move to the next action in the list and set SubAction to it.
'

    If Not SubAction Is Nothing Then Set SubAction = SubAction.NextNode
    
End Sub

Public Function Off() As Boolean
'
' Name:         Off
' Description:  Return TRUE iff there is no current subaction.
'
    Off = (SubAction Is Nothing)

End Function

Public Sub Remove()
'
' Name:         Remove
' Description:  Remove the current Subaction.
'
    
    If Not SubAction Is Nothing Then
        
        Dim TempNode As ActionNode
        
        If Not SubAction Is FirstNode Then
            Set SubAction.PrevNode.NextNode = SubAction.NextNode
        Else
            Set FirstNode = SubAction.NextNode
        End If
        
        If Not SubAction Is LastNode Then
            Set SubAction.NextNode.PrevNode = SubAction.PrevNode
        Else
            Set LastNode = SubAction.PrevNode
        End If
        
        Set TempNode = SubAction.NextNode
        Set SubAction = Nothing
        Set SubAction = TempNode
    
        NodeCount = NodeCount - 1
    
    End If
    
End Sub

Public Sub MoveTo(Name As String)
'
' Name:         MoveTo
' Parameters:   Name        Name of the SubAction to move to.
' Description:  Move to the given subaction.
'

    Set SubAction = FirstNode
    Do Until SubAction Is Nothing
        If SubAction.Name = Name Then Exit Do
        Set SubAction = SubAction.NextNode
    Loop

End Sub

Public Sub MoveToPlace(ByVal Place As Integer)
'
' Name:         MoveToPlace
' Parameters:   Place           the position in the list to move to
' Description:  Move the Link to the given number position in the list.
'               Like listbox listindexes, this is 0-based.
'
    
    Set SubAction = FirstNode
    Do Until (Place <= 0 Or SubAction Is Nothing)
        Set SubAction = SubAction.NextNode
        Place = Place - 1
    Loop

End Sub

Public Function Count() As Integer
'
' Name:         Count
' Description:  Return the number of subactions in this action.
'
    Count = NodeCount

End Function

Public Sub Clear()
'
' Name:         Clear
' Description:  Empty this action of subactions.
'

    First
    Do Until SubAction Is Nothing
        Remove
    Loop

End Sub

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write the object to an XML file.
'

    Dim MemCursor As ActionNode

    Set MemCursor = SubAction

    With XML
    
        .BeginTag "action"
    
            .WriteAttribute "date", ActDate
            .WriteAttribute "character", CharName
            .WriteAttribute "done", Done, False
            .WriteAttribute "lastmodified", LastModified
            
            First
            Do Until Off
                SubAction.OutputToFile XML
                MoveNext
            Loop
    
        .EndTag
        
    End With

    Set SubAction = MemCursor

End Sub

Public Sub InputFromFile(XML As XMLReaderClass, Version As Double)
'
' Name:         InputFromFile
' Parameters:   XML         an XMLReaderClass that just read the opening
'                           tag of this object
'               Version     version of the file format
' Description:  Read the object from an XML file.
'

    If XML.Tag = "action" Then
        
        With XML
        
            Clear
        
            ActDate = .GetDate("date")
            CharName = .GetString("character")
            Done = .GetBoolean("done")
            LastModified = .GetDate("lastmodified")
            
            Do Until .Tag = "action" And .Terminal
                .ReadTag
                If .Tag = "subaction" Then
                    Add "", 0, 0, 0, 0
                    SubAction.InputFromFile XML, Version
                End If
            Loop
            
        End With
        
    Else
        Err.Raise vbObjectError + 8, "ActionClass", _
                  "Wrong file input to an action"
    End If

End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this object to a binary file.
'

    Dim MemCursor As ActionNode

    Set MemCursor = SubAction

    Put #FileNum, , ActDate
    PutStrB FileNum, CharName
    Put #FileNum, , Done
    Put #FileNum, , LastModified
    
    Put #FileNum, , NodeCount
    First
    Do Until Off
        SubAction.OutputToBinary FileNum
        MoveNext
    Loop

    Set SubAction = MemCursor

End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this object from a binary file.
'

    Dim I As Integer

    Clear

    Get #FileNum, , ActDate
    GetStrB FileNum, CharName
    Get #FileNum, , Done
    Get #FileNum, , LastModified
    
    Get #FileNum, , I
    Do Until I = 0
        Add "", 0, 0, 0, 0
        SubAction.InputFromBinary FileNum, Version
        I = I - 1
    Loop
 
End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Clean up.
'
    
    Clear

End Sub
