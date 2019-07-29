VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RoteClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Level As Integer
Public Duration As String
Public SphereList As LinkedTraitList
Public Description As String
Public Grades As String
Public LastModified As Date
Public IconKey As String

Public Sub InitializeForOutput()
'
' Name:         InitializeForOutput
' Description:  For each list of data in this item,
'               move the cursor to the first item.  Used to
'               help generate output.
'

    SphereList.First
    
End Sub

Public Sub UpdateIconKey()
'
' Name:         UpdateIconKey
' Description:  Use a formula to choose a new icon for the rote.
'

    Dim HighLevel As Integer
    Dim HighName As String
    Dim NewLevel As Integer
    Dim NewName As String

    SphereList.First
    
    Do Until SphereList.Off
        NewLevel = SphereList.Trait.Number
        NewName = SphereList.Trait.Name
        If NewName Like "* Initiate" Or NewName Like "* Adept" Then NewLevel = NewLevel + 1
        If NewLevel > HighLevel Or (NewLevel = HighLevel And NewName > Name) Then
            HighLevel = NewLevel
            If InStr(NewName, ":") > 0 Then
                HighName = Left(NewName, InStr(NewName, ":") - 1)
            Else
                HighName = NewName
            End If
        End If
        SphereList.MoveNext
    Loop
         
    Select Case HighName
        Case "Correspondence":                  IconKey = "Correspondence"
        Case "Entropy":                         IconKey = "Entropy"
        Case "Forces":                          IconKey = "Forces"
        Case "Life":                            IconKey = "Life"
        Case "Matter":                          IconKey = "Matter"
        Case "Mind":                            IconKey = "Mind"
        Case "Spirit", "Dimensional Science":   IconKey = "Spirit"
        Case "Time":                            IconKey = "Time"
        Case Else:                              IconKey = "Prime"
    End Select
         
End Sub

Public Function Race() As String
'
' Name:         Race
' Description:  Return a string representation of this object's type.
' Returns:      "Rote", the object's type
'

    Race = "Rote"

End Function

Public Function OutputID() As Integer
'
' Name:         OutputID
' Description:  Return an ID that tells the OutputEngineClass how to manipulate this object
'
    OutputID = oidNone

End Function

Public Sub GetValue(Key As String, ItemData As Variant)
'
' Name:         GetValue
' Parameters:   Key         a key (string) to the value to retreive
'               CharData    the returned value - a string, number, date or list
' Description:  Use the given key to return data from the rote.  Keys are defined
'               in the PublicQueryKeys module.
'

    Select Case Key
        
        Case qkName:            ItemData = Name
        Case qkLevel:           ItemData = Level
        Case qkDuration:        ItemData = Duration
        Case qkSpheres:         Set ItemData = SphereList
        Case qkDescription:     ItemData = Description
        Case qkGrades:          ItemData = Grades
        Case qkLastModified:    ItemData = LastModified
        Case qkRandom:          ItemData = CInt(Rnd() * 100)
        
        Case Else:              ItemData = Null
        
    End Select

End Sub

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write this item's data to an XML file.
'

    With XML
    
        .BeginTag "rote"
        
            .WriteAttribute "name", Name
            .WriteAttribute "level", Level
            .WriteAttribute "duration", Duration
            .WriteAttribute "lastmodified", LastModified
            
            SphereList.OutputToFile XML
            
            .WriteCDataTag "description", Description
            .WriteCDataTag "grades", Grades
            
        .EndTag

    End With

End Sub

Public Sub InputFromFile(XML As XMLReaderClass, Version As Double)
'
' Name:         InputFromFile
' Parameters:   XML             XMLReaderClass that has read the tag
'               Version         file format version
' Description:  Reads the object in from an XML file.
'

    Dim ListName As String

    With XML
    
        If .Tag = "rote" Then
        
            Name = .GetString("name")
            Level = .GetInteger("level")
            Duration = .GetString("duration")
            LastModified = .GetDate("lastmodified")
            
            Do Until .Tag = "rote" And .Terminal
                .ReadTag
                Select Case .Tag
                    Case "traitlist"
                        ListName = .GetString("name")
                        Select Case ListName
                            Case "Spheres":         SphereList.InputFromFile XML, Version
                        End Select
                    Case "description"
                        If Not .Terminal Then
                            .ReadTag
                            Description = .GetTagText
                        End If
                    Case "grades"
                        If Not .Terminal Then
                            .ReadTag
                            Grades = .GetTagText
                        End If
                End Select
            Loop
        
            UpdateIconKey
        
        Else
            Err.Raise vbObjectError + 8, "RoteClass", _
                      "Wrong file input to a Rote Card"
        End If
    
    End With
        
End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this data to a binary file.
'

    PutStrB FileNum, Name
    Put #FileNum, , Level
    PutStrB FileNum, Duration
    
    SphereList.OutputToBinary FileNum

    PutStrB FileNum, Description
    PutStrB FileNum, Grades
    Put #FileNum, , LastModified
    
End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this data from a binary file.
'

    GetStrB FileNum, Name
    Get #FileNum, , Level
    GetStrB FileNum, Duration
    
    SphereList.InputFromBinary FileNum, Version

    GetStrB FileNum, Description
    GetStrB FileNum, Grades
    Get #FileNum, , LastModified

    UpdateIconKey

End Sub

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Intialize all data to defaults.
'
    
    LastModified = Now
    
    IconKey = "Prime"
    
    Set SphereList = New LinkedTraitList
    SphereList.Initialize "Spheres", False, False, True, ldNoteOnly
    
End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Destroy all objects created in this class.
'

    Set SphereList = Nothing

End Sub



