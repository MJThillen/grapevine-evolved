VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LocationClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public LocType As String
Public Level As Integer
Public Owner As String
Public Where As String
Public Appearance As String
Public SecTraits As Integer
Public SecRetests As Integer
Public Access As String
Public Security As String
Public Gauntlet As Integer
Public Affinity As String
Public Totem As String
Public Umbra As String
Public LinkList As LinkedTraitList
Public Notes As String
Public LastModified As Date

Public Sub InitializeForOutput()
'
' Name:         InitializeForOutput
' Description:  For each list of data in this item,
'               move the cursor to the first item.  Used to
'               help generate output.
'
    LinkList.First
    
End Sub

Public Function IconKey() As String
'
' Name:         IconKey
' Description:  Return the best icon for this location, based on its type.
'
    
    Select Case LocType
        Case "Haven", "Elysium":    IconKey = "Vampire"
        Case "Caern", "Blight", "Wyrmhole", "Hive"
                                    IconKey = "Werewolf"
        Case "Node", "Chantry":     IconKey = "Mage"
        Case "Freehold", "Glen", "Glade", "Trod":
                                    IconKey = "Changeling"
        Case "Haunt":               IconKey = "Wraith"
        Case "Dragon's Nest":       IconKey = "Kuei-Jin"
        Case "Realm":               IconKey = "Various"
        Case Else:                  IconKey = "Lantern"
    End Select
             
End Function

Public Function Race() As String
'
' Name:         Race
' Description:  Return a string representation of this object's type.
' Returns:      "Location", the object's type.
'

    Race = "Location"

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
' Description:  Use the given key to return data from the location.  Keys are defined
'               in the PublicQueryKeys module.
'

    Select Case Key
        
        Case qkName:            ItemData = Name
        Case qkType:            ItemData = LocType
        Case qkLevel:           ItemData = Level
        Case qkLinks:           Set ItemData = LinkList
        Case qkOwner:           ItemData = Owner
        Case qkAffinity:        ItemData = Affinity
        Case qkTotem:           ItemData = Totem
        Case qkWhere:           ItemData = Where
        Case qkAppearance:      ItemData = Appearance
        Case qkAccess:          ItemData = Access
        Case qkSecurity:        ItemData = Security
        Case qkSecTraits:       ItemData = SecTraits
        Case qkSecRetests:      ItemData = SecRetests
        Case qkGauntlet:        ItemData = Gauntlet
        Case qkUmbra:           ItemData = Umbra
        Case qkNotes:           ItemData = Notes
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
    
        .BeginTag "location"
        
            .WriteAttribute "name", Name
            .WriteAttribute "type", LocType
            .WriteAttribute "level", Level
            .WriteAttribute "owner", Owner
            .WriteAttribute "access", Access
            .WriteAttribute "securitytraits", SecTraits
            .WriteAttribute "securityretests", SecRetests
            .WriteAttribute "gauntlet", Gauntlet
            .WriteAttribute "affinity", Affinity
            .WriteAttribute "totem", Totem
            .WriteAttribute "lastmodified", LastModified
            
            LinkList.OutputToFile XML
            
            .WriteCDataTag "where", Where
            .WriteCDataTag "appearance", Appearance
            .WriteCDataTag "security", Security
            .WriteCDataTag "umbra", Umbra
            .WriteCDataTag "notes", Notes
            
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
    
        If .Tag = "location" Then
        
            Name = .GetString("name")
            LocType = .GetString("type")
            Level = .GetInteger("level")
            Owner = .GetString("owner")
            Access = .GetString("access")
            SecTraits = .GetInteger("securitytraits", 7)
            SecRetests = .GetInteger("securityretests")
            Gauntlet = .GetInteger("gauntlet", 7)
            Affinity = .GetString("affinity")
            Totem = .GetString("totem")
            LastModified = .GetDate("lastmodified")
            
            Do Until .Tag = "location" And .Terminal
                .ReadTag
                Select Case .Tag
                    Case "traitlist"
                        ListName = .GetString("name")
                        Select Case ListName
                            Case "Links":         LinkList.InputFromFile XML, Version
                        End Select
                    Case "where"
                        If Not .Terminal Then
                            .ReadTag
                            Where = .GetTagText
                        End If
                    Case "appearance"
                        If Not .Terminal Then
                            .ReadTag
                            Appearance = .GetTagText
                        End If
                    Case "security"
                        If Not .Terminal Then
                            .ReadTag
                            Security = .GetTagText
                        End If
                    Case "umbra"
                        If Not .Terminal Then
                            .ReadTag
                            Umbra = .GetTagText
                        End If
                    Case "notes"
                        If Not .Terminal Then
                            .ReadTag
                            Notes = .GetTagText
                        End If
                End Select
            Loop
        
        Else
            Err.Raise vbObjectError + 8, "LocationClass", _
                      "Wrong file input to a Location Card"
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
    PutStrB FileNum, LocType
    Put #FileNum, , Level
    PutStrB FileNum, Owner
    PutStrB FileNum, Access
    PutStrB FileNum, Affinity
    PutStrB FileNum, Totem
    Put #FileNum, , SecTraits
    Put #FileNum, , SecRetests
    Put #FileNum, , Gauntlet
    
    LinkList.OutputToBinary FileNum
    
    PutStrB FileNum, Where
    PutStrB FileNum, Appearance
    PutStrB FileNum, Security
    PutStrB FileNum, Umbra
    PutStrB FileNum, Notes
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
    GetStrB FileNum, LocType
    Get #FileNum, , Level
    GetStrB FileNum, Owner
    If Version >= 2.399 Then GetStrB FileNum, Access
    GetStrB FileNum, Affinity
    GetStrB FileNum, Totem
    If Version >= 2.396 Then
        Get #FileNum, , SecTraits
        Get #FileNum, , SecRetests
        Get #FileNum, , Gauntlet
        If Version >= 2.399 Then LinkList.InputFromBinary FileNum, Version
    End If
    
    GetStrB FileNum, Where
    GetStrB FileNum, Appearance
    GetStrB FileNum, Security
    If Version >= 2.396 Then GetStrB FileNum, Umbra
    GetStrB FileNum, Notes
    Get #FileNum, , LastModified

End Sub

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Intialize all data to defaults.
'
    
    SecTraits = 7
    Gauntlet = 7
    LastModified = Now
    
    Set LinkList = New LinkedTraitList
    LinkList.Initialize "Links", True, False, True, ldNoteOnly
    
End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Destroy all objects created in this class.
'

    Set LinkList = Nothing

End Sub



