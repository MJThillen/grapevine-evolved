VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BoonClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CharName As String
Public IsOwed As Boolean
Public BoonType As String
Public BoonDate As Date
Public Description As String

Public Sub InitializeForOutput()
'
' Name:         InitializeForOutput
' Description:  For each list of data in this item,
'               move the cursor to the first item.  Used to
'               help generate output.
'
    'No lists to initialize!
    
End Sub

Public Function Race() As String
'
' Name:         Race
' Description:  Return a string representation of this object's type.
' Returns:      "Boon", the object's type.
'

    Race = "Boon"

End Function

Public Function Name() As String
'
' Name:         Name
' Description:  Make up a name for this object.
'

    Name = BoonType & IIf(IsOwed, " owed to ", " due from ") & CharName

End Function

Public Function BoonID(Holder As String) As String
'
' Name:         BoonID
' Parameter:    Holder      Name of character holding this boon
' Description:  Create an ID that identifies this boon.  The same ID is created whether
'               A owes B or B is owed by A.
'
    If IsOwed Then
        BoonID = Holder & " -> " & CharName
    Else
        BoonID = CharName & " -> " & Holder
    End If
    
    BoonID = BoonID & ": " & BoonType & ", " & CStr(BoonDate)

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
' Description:  Use the given key to return data from the boon.  Keys are defined
'               in the PublicQueryKeys module.
'

    Select Case Key
        
        Case qkName:            ItemData = Name
        Case qkOwed:            ItemData = IsOwed
        Case qkPartner:         ItemData = CharName
        Case qkType:            ItemData = BoonType
        Case qkDate:            ItemData = BoonDate
        Case qkDescription:     ItemData = Description
        
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
    
        .BeginTag "boon"
        
            .WriteAttribute "type", BoonType
            .WriteAttribute "partner", CharName
            .WriteAttribute "owed", IsOwed
            .WriteAttribute "date", BoonDate
            
            .WriteCDataTag "description", Description
            
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

'    Dim ListName As String

    With XML
    
        If .Tag = "boon" Then
        
            BoonType = .GetString("type")
            CharName = .GetString("partner")
            IsOwed = .GetBoolean("owed")
            BoonDate = .GetDate("date")
            
            Do Until .Tag = "boon" And .Terminal
                .ReadTag
                Select Case .Tag
                    Case "description"
                        If Not .Terminal Then
                            .ReadTag
                            Description = .GetTagText
                        End If
                End Select
            Loop
        
        Else
            Err.Raise vbObjectError + 8, "BoonClass", _
                      "Wrong file input to a Boon"
        End If
    
    End With
        
End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this data to a binary file.
'

    PutStrB FileNum, BoonType
    PutStrB FileNum, CharName
    Put #FileNum, , IsOwed
    Put #FileNum, , BoonDate
    
    PutStrB FileNum, Description
    
End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this data from a binary file.
'

    GetStrB FileNum, BoonType
    GetStrB FileNum, CharName
    Get #FileNum, , IsOwed
    Get #FileNum, , BoonDate
    
    GetStrB FileNum, Description

End Sub

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Intialize all data to defaults.
'
    
    BoonDate = Now
    
End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Destroy all objects created in this class.
'

    ' Nothing to clean up!

End Sub
