VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ItemClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public ItemType As String
Public ItemSubtype As String
Public Level As Integer
Public Bonus As Integer
Public DamageType As String
Public DamageAmount As Integer
Public Concealability As String
Public TemperList As LinkedTraitList
Public NegativeList As LinkedTraitList
Public AbilityList As LinkedTraitList
Public Powers As String
Public Availability As LinkedTraitList
Public Appearance As String
Public Notes As String
Public LastModified As Date

Public Sub InitializeForOutput()
'
' Name:         InitializeForOutput
' Description:  For each list of data in this item,
'               move the cursor to the first item.  Used to
'               help generate output.
'

    Availability.First
    AbilityList.First
    NegativeList.First
    TemperList.First
    
End Sub

Public Function CombatNote() As String
'
' Name:         CombatNote
' Description:  Return the combat note of this item, if any, in the style:
'               "+3, 2 Agg, Loud, Clumsy"

    If DamageType <> "" Or Bonus > 0 Then
        CombatNote = "+" & CStr(Bonus)
        If DamageType <> "" Then CombatNote = CombatNote & ", " & CStr(DamageAmount) & " " & DamageType
    End If
    
    NegativeList.First
    Do Until NegativeList.Off
        CombatNote = CombatNote & ", " & NegativeList.DisplayTrait
        NegativeList.MoveNext
    Loop

End Function

Public Function IconKey() As String
'
' Name:         IconKey
' Description:  Returns a key that picks an icon from an imagelist based on Type
'

    Select Case ItemType
        Case "Fetish":              IconKey = "Werewolf"
        Case "Talen":               IconKey = "Fera"
        Case "Chimera", "Treasure": IconKey = "Changeling"
        Case "Artifact", "Charm", "Talisman", "Wonder": _
                                    IconKey = "Mage"
        Case "Effigy":              IconKey = "Mummy"
        Case "Relic":               IconKey = "Wraith"
        Case "Device", "Invention", "Gadget": _
                                    IconKey = "Device"
        Case Else:                  IconKey = "Stake"
    End Select
        
End Function

Public Function Race() As String
'
' Name:         Race
' Description:  Return a string representation of this item's type.
' Returns:      "Item", the item's type
'

    Race = "Item"

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
' Description:  Use the given key to return data from the item.  Keys are defined
'               in the PublicQueryKeys module.
'

    Select Case Key
        
        Case qkName:            ItemData = Name
        Case qkType:            ItemData = ItemType
        Case qkSubtype:         ItemData = ItemSubtype
        Case qkLevel:           ItemData = Level
        Case qkConcealability:  ItemData = Concealability
        Case qkAbilities:       Set ItemData = AbilityList
        Case qkTempers:         Set ItemData = TemperList
        Case qkAvailability:    Set ItemData = Availability
        Case qkNegatives:       Set ItemData = NegativeList
        Case qkDamageType:      ItemData = DamageType
        Case qkDamageAmount:    ItemData = DamageAmount
        Case qkBonus:           ItemData = Bonus
        Case qkPowers:          ItemData = Powers
        Case qkAppearance:      ItemData = Appearance
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
    
        .BeginTag "item"
        
            .WriteAttribute "name", Name
            .WriteAttribute "type", ItemType
            .WriteAttribute "subtype", ItemSubtype
            .WriteAttribute "level", Level
            .WriteAttribute "bonus", Bonus
            .WriteAttribute "damage", DamageType
            .WriteAttribute "amount", DamageAmount
            .WriteAttribute "conceal", Concealability
            .WriteAttribute "lastmodified", LastModified
            
            TemperList.OutputToFile XML
            AbilityList.OutputToFile XML
            NegativeList.OutputToFile XML
            Availability.OutputToFile XML
            
            .WriteCDataTag "powers", Powers
            .WriteCDataTag "appearance", Appearance
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
    
        If .Tag = "item" Then
        
            Name = .GetString("name")
            ItemType = .GetString("type")
            ItemSubtype = .GetString("subtype")
            Level = .GetInteger("level")
            Bonus = .GetInteger("bonus")
            DamageType = .GetString("damage")
            DamageAmount = .GetInteger("amount")
            Concealability = .GetString("conceal")
            LastModified = .GetDate("lastmodified")
            
            Do Until .Tag = "item" And .Terminal
                .ReadTag
                Select Case .Tag
                    Case "traitlist"
                        ListName = .GetString("name")
                        Select Case ListName
                            Case "Tempers":         TemperList.InputFromFile XML, Version
                            Case "Abilities":       AbilityList.InputFromFile XML, Version
                            Case "Negatives":       NegativeList.InputFromFile XML, Version
                            Case "Availability":    Availability.InputFromFile XML, Version
                        End Select
                    Case "powers"
                        If Not .Terminal Then
                            .ReadTag
                            Powers = .GetTagText
                        End If
                    Case "appearance"
                        If Not .Terminal Then
                            .ReadTag
                            Appearance = .GetTagText
                        End If
                    Case "notes"
                        If Not .Terminal Then
                            .ReadTag
                            Notes = .GetTagText
                        End If
                End Select
            Loop
        
        Else
            Err.Raise vbObjectError + 8, "ItemClass", _
                      "Wrong file input to an Item Card"
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
    PutStrB FileNum, ItemType
    PutStrB FileNum, ItemSubtype
    Put #FileNum, , Level
    Put #FileNum, , Bonus
    PutStrB FileNum, DamageType
    Put #FileNum, , DamageAmount
    PutStrB FileNum, Concealability
    
    TemperList.OutputToBinary FileNum
    AbilityList.OutputToBinary FileNum
    NegativeList.OutputToBinary FileNum
    Availability.OutputToBinary FileNum
    
    PutStrB FileNum, Powers
    PutStrB FileNum, Appearance
    PutStrB FileNum, Notes
    Put #FileNum, , LastModified
    
End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this character's data from a binary file.
'

    GetStrB FileNum, Name
    GetStrB FileNum, ItemType
    GetStrB FileNum, ItemSubtype
    Get #FileNum, , Level
    Get #FileNum, , Bonus
    GetStrB FileNum, DamageType
    Get #FileNum, , DamageAmount
    GetStrB FileNum, Concealability
    
    TemperList.InputFromBinary FileNum, Version
    AbilityList.InputFromBinary FileNum, Version
    NegativeList.InputFromBinary FileNum, Version
    Availability.InputFromBinary FileNum, Version

    GetStrB FileNum, Powers
    GetStrB FileNum, Appearance
    GetStrB FileNum, Notes
    Get #FileNum, , LastModified

End Sub

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Intialize all data to defaults.
'
    
    LastModified = Now
    
    ItemType = "Equipment"
    
    Set TemperList = New LinkedTraitList
    Set AbilityList = New LinkedTraitList
    Set NegativeList = New LinkedTraitList
    Set Availability = New LinkedTraitList

    TemperList.Initialize "Tempers", True, False, False, ldMultiplierDot
    AbilityList.Initialize "Abilities", True, False, False, ldMultiplier
    NegativeList.Initialize "Negatives", True, True, False, ldMultiplier
    Availability.Initialize "Availability", True, False, False, ldMultiplier
    
End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Destroy all objects created in this class.
'

    Set AbilityList = Nothing
    Set TemperList = Nothing
    Set NegativeList = Nothing
    Set Availability = Nothing

End Sub


