VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlayerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String       'the player's name
Public ID As String         'the player's ID
Public EMail As String      'e-mail address
Public Phone As String      'phone
Public Position As String   'role in the organization
Public Status As String     'Player's status
Public Address As String    'address
Public Experience As ExperienceClass 'player point tracking
Public Notes As String      'notes on the player
Public LastModified As Date

Public Function OutputID() As Integer
'
' Name:         OutputID
' Description:  Return an ID that tells the OutputEngineClass how to manipulate this object
'
    OutputID = oidNone

End Function

Public Sub InitializeForOutput()
'
' Name:         InitializeForOutput
' Description:  Position the object in a starting state, ready for output.
'

    Experience.InitializeForOutput

End Sub

Public Function ExpandedEMail() As String
'
' Name:         ExpandedEmail
' Description:  Prefix the e-mail address with a name in quotes, and surround it with < and >.
'

    ExpandedEMail = ""
    If EMail <> "" Then
        ExpandedEMail = Replace(Name, ",", "")
        ExpandedEMail = Replace(ExpandedEMail, "<", "(")
        ExpandedEMail = Replace(ExpandedEMail, ">", ")")
        ExpandedEMail = Replace(ExpandedEMail, """", "'")
        ExpandedEMail = """" & ExpandedEMail & """ <" & EMail & ">"
    End If
    
End Function

Public Sub GetValue(Key As String, PData As Variant)
'
' Name:         GetValue
' Parameters:   Key         a key (string) to the value to retreive
'               PData       the returned value - a string, number, date or list
' Description:  Use the given key to return data from the object.  Keys are defined
'               in the PublicQueryKeys module.
'

    Select Case Key
        
        Case qkPosition:        PData = Position
        Case qkName:            PData = Name
        Case qkID:              PData = ID
        Case qkPlayStatus:      PData = Status
        Case qkPPEarned:        PData = Experience.Earned
        Case qkPPUnspent:       PData = Experience.Unspent
        Case qkLastModified:    PData = LastModified
        Case qkEMail:           PData = EMail
        Case qkPhone:           PData = Phone
        Case qkAddress:         PData = Address
        Case qkNotes:           PData = Notes
        Case qkActive:          PData = (Status = ActiveStatus)
        Case qkRandom:          PData = CInt(Rnd() * 100)
        Case Else:              PData = Null
        
    End Select

End Sub

Public Function Race() As String
'
' Name:         Race
' Description:  Return a string representation of this player's faux
'               "character type".
' Returns:      the name of the "character type"
'

    Race = "Player"

End Function

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write this player's data to a file.
'

    With XML
    
        .BeginTag "player"
        
            .WriteAttribute "name", Name
            .WriteAttribute "id", ID
            .WriteAttribute "email", EMail
            .WriteAttribute "phone", Phone
            .WriteAttribute "position", Position
            .WriteAttribute "status", Status
            .WriteAttribute "lastmodified", LastModified
            Experience.OutputToFile XML
            .WriteCDataTag "address", Address
            .WriteCDataTag "notes", Notes
            
        .EndTag
    
    End With

End Sub

Public Sub InputFromFile(XML As XMLReaderClass, Version As Double)
'
' Name:         InputFromFile
' Parameters:   XML             XMLReaderClass that has read the player tag
'               Version         file format version
' Description:  Reads player information from a file.  Version
'               ensures backward-compatibility.
'
    
    With XML
    
        If .Tag = "player" Then
        
            Dim Active As Boolean
        
            Name = .GetString("name")
            ID = .GetString("id")
            EMail = .GetString("email")
            Phone = .GetString("phone")
            Position = .GetString("position", "Player")
            Active = Not .GetBoolean("inactive")    'old version
            Status = .GetString("status", IIf(Active, ActiveStatus, "Inactive"))
            LastModified = .GetDate("lastmodified")
            
            Do Until .Tag = "player" And .Terminal
                .ReadTag
                Select Case .Tag
                    Case "experience"
                        Experience.InputFromFile XML, Version
                    Case "address", "location"
                        .ReadTag
                        Address = .GetTagText
                    Case "notes"
                        .ReadTag
                        Notes = .GetTagText
                End Select
            Loop
        
        Else
        
            Err.Raise vbObjectError + 8, "PlayerClass", _
                      "Wrong file input to a player"
    
        End If
    
    End With
    
End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this character's data to a binary file.
'

    PutStrB FileNum, Name
    PutStrB FileNum, ID
    PutStrB FileNum, EMail
    PutStrB FileNum, Phone
    PutStrB FileNum, Position
    PutStrB FileNum, Status
    Put #FileNum, , LastModified
    Experience.OutputToBinary FileNum
    PutStrB FileNum, Address
    PutStrB FileNum, Notes

End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this character's data from a binary file.
'

    GetStrB FileNum, Name
    GetStrB FileNum, ID
    GetStrB FileNum, EMail
    GetStrB FileNum, Phone
    GetStrB FileNum, Position
    If Version >= 2.397 Then
        GetStrB FileNum, Status
    Else
        Dim Active As Boolean
        Get #FileNum, , Active
        Status = IIf(Active, ActiveStatus, "Inactive")
    End If
    Get #FileNum, , LastModified
    Experience.InputFromBinary FileNum, Version
    GetStrB FileNum, Address
    GetStrB FileNum, Notes

End Sub

Public Sub OldInputFromFile(FileNum As Integer, OldDate As Date, Version As String)
'
' Name:         InputFromFile
' Parameters:   FileNum         number of an open file
'               OldDate         date of the old file
'               Version         file format version
' Description:  Reads player information from a file.  Version
'               ensures backward-compatibility.
'
    
    Dim Read As String

    Line Input #FileNum, Name
    
    If Version >= GameFileVersionTag3 Then
        Line Input #FileNum, ID
        Line Input #FileNum, Read
        Experience.Unspent = Val(Read)
        Experience.Earned = Val(Read)
    End If
    
    Line Input #FileNum, EMail
    Line Input #FileNum, Phone
    
    If Version >= GameFileVersionTag3 Then
        Line Input #FileNum, Position
    End If
    
    Address = ReadLongField(FileNum, "<-End Player Location->")
    
    If Version >= GameFileVersionTag3 Then
        Experience.OldInputFromFile FileNum, Version
    End If
    
    Notes = ReadLongField(FileNum, "<-End Player Notes->")
    LastModified = OldDate

End Sub

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Initialize data to default values.
'
    Position = "Player"
    Status = ActiveStatus
    Set Experience = New ExperienceClass

End Sub
