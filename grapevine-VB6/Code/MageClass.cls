VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MageClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Nature As String
Public Demeanor As String
Public Essence As String
Public Tradition As String
Public Faction As String
Public Cabal As String
Public Rank As String

Public Willpower As Integer
Public Arete As Integer
Public Quintessence As Integer
Public Paradox As Integer

Public TempWillpower As Integer
Public TempArete As Integer
Public TempQuintessence As Integer
Public TempParadox As Integer

Public PhysicalList As LinkedTraitList
Public SocialList As LinkedTraitList
Public MentalList As LinkedTraitList
Public PhysicalNegList As LinkedTraitList
Public SocialNegList As LinkedTraitList
Public MentalNegList As LinkedTraitList
Public PhysicalMax As Integer
Public SocialMax As Integer
Public MentalMax As Integer

Public AbilityList As LinkedTraitList
Public InfluenceList As LinkedTraitList
Public BackgroundList As LinkedTraitList
Public HealthList As LinkedTraitList
Public ResonanceList As LinkedTraitList
Public ReputationList As LinkedTraitList

Public SphereList As LinkedTraitList

Public MeritList As LinkedTraitList
Public FlawList As LinkedTraitList

Public Foci As String
Public RoteList As LinkedTraitList

Public EquipmentList As LinkedTraitList
Public HangoutList As LinkedTraitList
Public Biography As String
Public Notes As String

Public Player As String
Public Status As String
Public ID As String
Public StartDate As Date
Public Narrator As String
Public IsNPC As Boolean
Public Experience As ExperienceClass
Public LastModified As Date

Public Sub InitializeForOutput()
'
' Name:         InitializeForOutput
' Description:  For each list of data in this character,
'               move the cursor to the first item.  Used to
'               help generate character sheets.
'

    PhysicalList.First
    SocialList.First
    MentalList.First
    PhysicalNegList.First
    SocialNegList.First
    MentalNegList.First
    AbilityList.First
    InfluenceList.First
    BackgroundList.First
    HealthList.First
    SphereList.First
    ResonanceList.First
    ReputationList.First
    RoteList.First
    MeritList.First
    FlawList.First
    EquipmentList.First
    HangoutList.First
    Experience.InitializeForOutput

End Sub

Public Function OutputID() As Integer
'
' Name:         OutputID
' Description:  Return an ID that tells the OutputEngineClass how to manipulate this object
'
    OutputID = oidNone

End Function

Public Sub GetValue(Key As String, CharData As Variant)
'
' Name:         GetValue
' Parameters:   Key         a key (string) to the value to retreive
'               CharData    the returned value - a string, number, date or list
' Description:  Use the given key to return data from the character.  Keys are defined
'               in the PublicQueryKeys module.
'

    Select Case Key
        
        Case qkPlayStatus:      CharData = Status
        Case qkInfluences:      Set CharData = InfluenceList
        Case qkMerits:          Set CharData = MeritList
        Case qkFlaws:           Set CharData = FlawList
        Case qkBackgrounds:     Set CharData = BackgroundList
        Case qkNPC:             CharData = IsNPC
        Case qkAbilities:       Set CharData = AbilityList
        Case qkEquipment:       Set CharData = EquipmentList
        Case qkLocations:       Set CharData = HangoutList
        Case qkPhysical:        Set CharData = PhysicalList
        Case qkSocial:          Set CharData = SocialList
        Case qkMental:          Set CharData = MentalList
        Case qkPhysicalNeg:     Set CharData = PhysicalNegList
        Case qkSocialNeg:       Set CharData = SocialNegList
        Case qkMentalNeg:       Set CharData = MentalNegList
        Case qkXPEarned:        CharData = Experience.Earned
        Case qkXPUnspent:       CharData = Experience.Unspent
        Case qkNature:          CharData = Nature
        Case qkDemeanor:        CharData = Demeanor
        Case qkRace:            CharData = Race
        Case qkGroup:           CharData = Group
        Case qkSubgroup:        CharData = Subgroup
        Case qkWillpower:       CharData = Willpower
        Case qkTempWillpower:   CharData = TempWillpower
        Case qkLastModified:    CharData = LastModified
        Case qkStartDate:       CharData = StartDate
        Case qkHealth:          Set CharData = HealthList
        Case qkNotes:           CharData = Notes
        Case qkID:              CharData = ID
        Case qkPlayer:          CharData = Player
        Case qkNarrator:        CharData = Narrator
        Case qkName:            CharData = Name
        Case qkEMail:           CharData = EMail
        Case qkRandom:          CharData = CInt(Rnd() * 100)
        
        Case qkSpheres:         Set CharData = SphereList
        Case qkResonance:       Set CharData = ResonanceList
        Case qkReputation:      Set CharData = ReputationList
        Case qkRotes:           Set CharData = RoteList
        Case qkArete:           CharData = Arete
        Case qkTradition:       CharData = Tradition
        Case qkFaction:         CharData = Faction
        Case qkRank:            CharData = Rank
        Case qkQuintessence:    CharData = Quintessence
        Case qkParadox:         CharData = Paradox
        Case qkEssence:         CharData = Essence
        Case qkCabal:           CharData = Cabal
        Case qkFoci:            CharData = Foci
        Case qkBiography:       CharData = Biography
        Case qkPhysicalMax:     CharData = PhysicalMax
        Case qkSocialMax:       CharData = SocialMax
        Case qkMentalMax:       CharData = MentalMax
        Case qkTempArete:       CharData = TempArete
        Case qkTempQuintessence: CharData = TempQuintessence
        Case qkTempParadox:     CharData = TempParadox
        
        Case Else:              CharData = Null
        
    End Select

End Sub

Public Sub SetValue(Key As String, CharData As Variant)
'
' Name:         SetValue
' Parameters:   Key         a key (string) to the value to set
'               CharData    the new value - a string, number, date or list
' Description:  Use the given key to set data for the character.  Keys are defined
'               in the PublicQueryKeys module.
'

    On Error Resume Next
    
    Select Case Key
        
        Case qkWillpower:       Willpower = CInt(CharData)
        Case qkTempWillpower:   TempWillpower = CInt(CharData)
        Case qkArete:           Arete = CInt(CharData)
        Case qkQuintessence:    Quintessence = CInt(CharData)
        Case qkParadox:         Paradox = CInt(CharData)
        Case qkTempArete:       TempArete = CInt(CharData)
        Case qkTempQuintessence: TempQuintessence = CInt(CharData)
        Case qkTempParadox:     TempParadox = CInt(CharData)
        
    End Select

End Sub

Public Sub CompleteEstimateList(Estimate As LinkedTraitList)
'
' Name:         CompleteEstimateList
' Parameter:    Estimate        Traitlist that will hold estimate info
' Description:  Add race-specific entries to a LinkedTraitList for use with the Point Estimate Tool.
'

    Estimate.Append qkMerits
    Estimate.Append qkFlaws
    Estimate.Append qkSpheres, , "30"
    Estimate.Append qkRotes
    Estimate.Append qkArete, "4", "1"
    Estimate.Append qkWillpower, "3", "5"

End Sub

Public Function Race() As String
'
' Name:         Race
' Description:  Return a string representation of this character's
'               character type.
' Returns:      the name of the character type
'

    Race = "Mage"

End Function

Public Function RaceCode() As RaceType
'
' Name:         RaceCode
' Description:  Return the RaceType (enumeration value) of this
'               character type.
' Returns:      the corresponding RaceType
'
    
    RaceCode = gvracemage

End Function

Public Function Group() As String
'
' Name:         Group
' Description:  Return the 'Group' this character belongs to.
'               Clan, Tribe, Kith, Guild, etc.
' Returns:      the group
'

    Group = Tradition

End Function

Public Function Subgroup() As String
'
' Name:         Subgroup
' Description:  Return the 'Subgroup' this character belongs to.
'               Sect, Auspice, Seeming, etc.
' Returns:      the subgroup
'

    Subgroup = Rank

End Function

Public Function EMail() As String
'
' Name:         EMail
' Description:  Find the player of this character and return its email, if any.
'
    
    EMail = ""
    PlayerList.Memorize
    PlayerList.MoveTo Player
    If Not PlayerList.Off Then EMail = PlayerList.Item.EMail
    PlayerList.Remember

End Function

Public Function ExpandedEMail() As String
'
' Name:         ExpandedEMail
' Description:  Find the player of this character and return its Expanded EMail, if any.
'
    
    ExpandedEMail = ""
    PlayerList.Memorize
    PlayerList.MoveTo Player
    If Not PlayerList.Off Then ExpandedEMail = PlayerList.Item.ExpandedEMail
    PlayerList.Remember

End Function

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write this character's data to a file.
'

    With XML
    
        .BeginTag "mage"
        
            .WriteAttribute "name", Name
            .WriteAttribute "nature", Nature
            .WriteAttribute "demeanor", Demeanor
            .WriteAttribute "essence", Essence
            .WriteAttribute "tradition", Tradition
            .WriteAttribute "cabal", Cabal
            .WriteAttribute "rank", Rank
            .WriteAttribute "faction", Faction
            .WriteAttribute "willpower", Willpower
            .WriteAttribute "tempwillpower", TempWillpower, Willpower
            .WriteAttribute "arete", Arete
            .WriteAttribute "temparete", TempArete, Arete
            .WriteAttribute "quintessence", Quintessence
            .WriteAttribute "tempquintessence", TempQuintessence, Quintessence
            .WriteAttribute "paradox", Paradox
            .WriteAttribute "tempparadox", TempParadox, Paradox
            .WriteAttribute "physicalmax", PhysicalMax
            .WriteAttribute "socialmax", SocialMax, PhysicalMax
            .WriteAttribute "mentalmax", MentalMax, PhysicalMax
            .WriteAttribute "player", Player
            .WriteAttribute "status", Status
            .WriteAttribute "id", ID
            .WriteAttribute "startdate", StartDate
            .WriteAttribute "narrator", Narrator
            .WriteAttribute "npc", IsNPC, False
            .WriteAttribute "lastmodified", LastModified
            
            Experience.OutputToFile XML
        
            PhysicalList.OutputToFile XML
            SocialList.OutputToFile XML
            MentalList.OutputToFile XML
            PhysicalNegList.OutputToFile XML
            SocialNegList.OutputToFile XML
            MentalNegList.OutputToFile XML
                
            AbilityList.OutputToFile XML
            InfluenceList.OutputToFile XML
            
            BackgroundList.OutputToFile XML
            HealthList.OutputToFile XML
            ResonanceList.OutputToFile XML
            ReputationList.OutputToFile XML
            
            SphereList.OutputToFile XML
            RoteList.OutputToFile XML
                    
            MeritList.OutputToFile XML
            FlawList.OutputToFile XML
            EquipmentList.OutputToFile XML
            HangoutList.OutputToFile XML
    
            .WriteCDataTag "foci", Foci
            .WriteCDataTag "biography", Biography
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
    
        If .Tag = "mage" Then
        
            Name = .GetString("name")
            Tradition = .GetString("tradition")
            Faction = .GetString("faction")
            Nature = .GetString("nature")
            Demeanor = .GetString("demeanor")
            Essence = .GetString("essence")
            Cabal = .GetString("cabal")
            Rank = .GetString("rank")
            Faction = .GetString("faction")
            Arete = .GetInteger("arete")
            TempArete = .GetInteger("temparete", Arete)
            Quintessence = .GetInteger("quintessence")
            TempQuintessence = .GetInteger("tempquintessence", Quintessence)
            Paradox = .GetInteger("paradox")
            TempParadox = .GetInteger("tempparadox", Paradox)
            Willpower = .GetInteger("willpower")
            TempWillpower = .GetInteger("tempwillpower", Willpower)
            PhysicalMax = .GetInteger("physicalmax", 11)
            SocialMax = .GetInteger("socialmax", PhysicalMax)
            MentalMax = .GetInteger("mentalmax", PhysicalMax)
            Player = .GetString("player")
            Status = .GetString("status")
            ID = .GetString("id")
            StartDate = .GetDate("startdate")
            Narrator = .GetString("narrator")
            IsNPC = .GetBoolean("npc")
            LastModified = .GetDate("lastmodified")
            
            Do Until .Tag = "mage" And .Terminal
                .ReadTag
                Select Case .Tag
                    Case "experience"
                        Experience.InputFromFile XML, Version
                    Case "traitlist"
                        ListName = .GetString("name")
                        Select Case ListName
                            Case "Physical":            PhysicalList.InputFromFile XML, Version
                            Case "Social":              SocialList.InputFromFile XML, Version
                            Case "Mental":              MentalList.InputFromFile XML, Version
                            Case "Negative Physical":   PhysicalNegList.InputFromFile XML, Version
                            Case "Negative Social":     SocialNegList.InputFromFile XML, Version
                            Case "Negative Mental":     MentalNegList.InputFromFile XML, Version
                            Case "Abilities":           AbilityList.InputFromFile XML, Version
                            Case "Influences":          InfluenceList.InputFromFile XML, Version
                            Case "Backgrounds":         BackgroundList.InputFromFile XML, Version
                            Case "Health Levels":       HealthList.InputFromFile XML, Version
                            Case "Resonance":           ResonanceList.InputFromFile XML, Version
                            Case "Reputation":          ReputationList.InputFromFile XML, Version
                            Case "Rotes":               RoteList.InputFromFile XML, Version
                            Case "Spheres":             SphereList.InputFromFile XML, Version
                            Case "Merits":              MeritList.InputFromFile XML, Version
                            Case "Flaws":               FlawList.InputFromFile XML, Version
                            Case "Equipment":           EquipmentList.InputFromFile XML, Version
                            Case "Locations":           HangoutList.InputFromFile XML, Version
                        End Select
                    Case "foci"
                        If Not .Terminal Then
                            .ReadTag
                            Foci = .GetTagText
                        End If
                    Case "biography"
                        If Not .Terminal Then
                            .ReadTag
                            Biography = .GetTagText
                        End If
                    Case "notes"
                        If Not .Terminal Then
                            .ReadTag
                            Notes = .GetTagText
                        End If
                End Select
            Loop

            If Version < 2.397 Then
                If PhysicalList.Count > PhysicalMax Then PhysicalMax = PhysicalList.Count
                If SocialList.Count > PhysicalMax Then PhysicalMax = SocialList.Count
                If MentalList.Count > PhysicalMax Then PhysicalMax = MentalList.Count
                SocialMax = PhysicalMax
                MentalMax = PhysicalMax
            End If

        Else
            Err.Raise vbObjectError + 8, "MageClass", _
                      "Wrong file input to a Mage character"
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
    PutStrB FileNum, Nature
    PutStrB FileNum, Demeanor
    PutStrB FileNum, Essence
    PutStrB FileNum, Tradition
    PutStrB FileNum, Faction
    PutStrB FileNum, Cabal
    PutStrB FileNum, Rank
    Put #FileNum, , Willpower
    Put #FileNum, , TempWillpower
    Put #FileNum, , Arete
    Put #FileNum, , TempArete
    Put #FileNum, , Quintessence
    Put #FileNum, , TempQuintessence
    Put #FileNum, , Paradox
    Put #FileNum, , TempParadox
    Put #FileNum, , PhysicalMax
    Put #FileNum, , SocialMax
    Put #FileNum, , MentalMax
    PutStrB FileNum, Player
    PutStrB FileNum, Status
    PutStrB FileNum, ID
    Put #FileNum, , StartDate
    PutStrB FileNum, Narrator
    Put #FileNum, , IsNPC
    Put #FileNum, , LastModified
    
    Experience.OutputToBinary FileNum

    PhysicalList.OutputToBinary FileNum
    SocialList.OutputToBinary FileNum
    MentalList.OutputToBinary FileNum
    PhysicalNegList.OutputToBinary FileNum
    SocialNegList.OutputToBinary FileNum
    MentalNegList.OutputToBinary FileNum
        
    AbilityList.OutputToBinary FileNum
    InfluenceList.OutputToBinary FileNum
    
    BackgroundList.OutputToBinary FileNum
    HealthList.OutputToBinary FileNum
    ResonanceList.OutputToBinary FileNum
    ReputationList.OutputToBinary FileNum
    
    SphereList.OutputToBinary FileNum
    RoteList.OutputToBinary FileNum
            
    MeritList.OutputToBinary FileNum
    FlawList.OutputToBinary FileNum
    EquipmentList.OutputToBinary FileNum
    HangoutList.OutputToBinary FileNum

    PutStrB FileNum, Foci
    PutStrB FileNum, Biography
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
    GetStrB FileNum, Nature
    GetStrB FileNum, Demeanor
    GetStrB FileNum, Essence
    GetStrB FileNum, Tradition
    GetStrB FileNum, Faction
    GetStrB FileNum, Cabal
    GetStrB FileNum, Rank
    
    If Version >= 2.397 Then
        Get #FileNum, , Willpower
        Get #FileNum, , TempWillpower
        Get #FileNum, , Arete
        Get #FileNum, , TempArete
        Get #FileNum, , Quintessence
        Get #FileNum, , TempQuintessence
        Get #FileNum, , Paradox
        Get #FileNum, , TempParadox
        Get #FileNum, , PhysicalMax
        Get #FileNum, , SocialMax
        Get #FileNum, , MentalMax
    Else
        Get #FileNum, , Willpower
        Get #FileNum, , Arete
        Get #FileNum, , Quintessence
        Get #FileNum, , Paradox
        TempWillpower = Willpower
        TempArete = Arete
        TempQuintessence = Quintessence
        TempParadox = Paradox
    End If
    
    GetStrB FileNum, Player
    GetStrB FileNum, Status
    GetStrB FileNum, ID
    If Version >= 2.397 Then
        Get #FileNum, , StartDate
    Else
        GetStrB FileNum, Narrator 'Use temporarily to convert from old string format
        If IsDate(Narrator) Then StartDate = CDate(Narrator)
    End If
    GetStrB FileNum, Narrator
    Get #FileNum, , IsNPC
    Get #FileNum, , LastModified
    
    Experience.InputFromBinary FileNum, Version

    PhysicalList.InputFromBinary FileNum, Version
    SocialList.InputFromBinary FileNum, Version
    MentalList.InputFromBinary FileNum, Version
    PhysicalNegList.InputFromBinary FileNum, Version
    SocialNegList.InputFromBinary FileNum, Version
    MentalNegList.InputFromBinary FileNum, Version
        
    AbilityList.InputFromBinary FileNum, Version
    InfluenceList.InputFromBinary FileNum, Version
    
    BackgroundList.InputFromBinary FileNum, Version
    HealthList.InputFromBinary FileNum, Version
    ResonanceList.InputFromBinary FileNum, Version
    ReputationList.InputFromBinary FileNum, Version
    
    SphereList.InputFromBinary FileNum, Version
    RoteList.InputFromBinary FileNum, Version
            
    MeritList.InputFromBinary FileNum, Version
    FlawList.InputFromBinary FileNum, Version
    EquipmentList.InputFromBinary FileNum, Version
    If Version >= 2.395 Then HangoutList.InputFromBinary FileNum, Version

    GetStrB FileNum, Foci
    If Version >= 2.397 Then GetStrB FileNum, Biography
    GetStrB FileNum, Notes

    If Version < 2.397 Then
        If PhysicalList.Count > PhysicalMax Then PhysicalMax = PhysicalList.Count
        If SocialList.Count > PhysicalMax Then PhysicalMax = SocialList.Count
        If MentalList.Count > PhysicalMax Then PhysicalMax = MentalList.Count
        SocialMax = PhysicalMax
        MentalMax = PhysicalMax
    End If
    
End Sub

Public Sub OldInputFromFile(FileNum As Integer, OldDate As Date, Version As String)
'
' Name:         OldInputFromFile
' Parameters:   FileNum         number of an open file
'               Version         file format version of the file
' Description:  Read a character from an old-format file.  The version
'               ensures backward-compatibility.
'

    Dim Read As String
    
    ' Race has already been read
    Line Input #FileNum, Name
    Line Input #FileNum, Nature
    Line Input #FileNum, Demeanor
    Line Input #FileNum, Essence
    Line Input #FileNum, Tradition
    Line Input #FileNum, Cabal
    Line Input #FileNum, Rank
    Rank = ""
    
    Line Input #FileNum, Read
    Willpower = Val(Read)
    TempWillpower = Willpower
    Line Input #FileNum, Read
    Arete = Val(Read)
    TempArete = Arete
    Line Input #FileNum, Read
    Quintessence = Val(Read)
    TempQuintessence = Quintessence
    Line Input #FileNum, Read
    Paradox = Val(Read)
    TempParadox = Paradox

    PhysicalList.OldInputFromFile FileNum, Version
    SocialList.OldInputFromFile FileNum, Version
    MentalList.OldInputFromFile FileNum, Version
    PhysicalNegList.OldInputFromFile FileNum, Version
    SocialNegList.OldInputFromFile FileNum, Version
    MentalNegList.OldInputFromFile FileNum, Version

    AbilityList.OldInputFromFile FileNum, Version
    InfluenceList.OldInputFromFile FileNum, Version

    BackgroundList.OldInputFromFile FileNum, Version
    HealthList.OldInputFromFile FileNum, Version
    ResonanceList.OldInputFromFile FileNum, Version
    
    SphereList.OldInputFromFile FileNum, Version
    
    SphereList.First
    Do Until SphereList.Off
        Dim NameSuffix As String
        Dim NoteSuffix As String
        Select Case SphereList.Trait.Number
            Case 1
                NameSuffix = ": Apprentice"
                NoteSuffix = " basic"
            Case 2
                NameSuffix = ": Initiate"
                NoteSuffix = " basic"
            Case 3
                NameSuffix = ": Disciple"
                NoteSuffix = " int."
            Case 4
                NameSuffix = ": Adept"
                NoteSuffix = " int."
            Case 5
                NameSuffix = ": Master"
                NoteSuffix = " adv."
            Case 6
                NameSuffix = ": Archmage"
                NoteSuffix = ""
            Case Else
                NameSuffix = ""
                NoteSuffix = ""
        End Select
        SphereList.Trait.Name = SphereList.Trait.Name & NameSuffix
        SphereList.Trait.Note = Trim(SphereList.Trait.Note & NoteSuffix)
        SphereList.MoveNext
    Loop
    
    MeritList.LLInputFromFile FileNum, Version
    FlawList.LLInputFromFile FileNum, Version

    Foci = ReadLongField(FileNum, "<-End Mage Character Foci->")
    
    Notes = ""
    Read = ReadLongField(FileNum, "<-End Mage Character Rotes->")
    If Read <> "" Then Notes = "2.3 Rotes:" & vbCrLf & Read & vbCrLf
    
    Notes = ""
    Read = ReadLongField(FileNum, "<-End Mage Character Equipment->")
    If Read <> "" Then Notes = "2.3 Equipment:" & vbCrLf & Read & vbCrLf
    Read = ReadLongField(FileNum, "<-End Mage Character Other->")
    If Read <> "" Then Notes = "2.3 Other:" & vbCrLf & Read & vbCrLf
    
    Notes = Notes & ReadLongField(FileNum, "<-End Mage Character Notes->")
    
    If PhysicalList.Count > PhysicalMax Then PhysicalMax = PhysicalList.Count
    If SocialList.Count > PhysicalMax Then PhysicalMax = SocialList.Count
    If MentalList.Count > PhysicalMax Then PhysicalMax = MentalList.Count
    SocialMax = PhysicalMax
    MentalMax = PhysicalMax
    
    Line Input #FileNum, Player
    Line Input #FileNum, Read
    Experience.Unspent = Val(Read)
    Experience.Earned = Val(Read)
    Line Input #FileNum, Status
    Line Input #FileNum, ID
    Line Input #FileNum, Narrator
    If IsDate(Narrator) Then StartDate = CDate(Narrator)
    Line Input #FileNum, Narrator
    Line Input #FileNum, Read
    IsNPC = CBool(Read)
    Experience.OldInputFromFile FileNum, Version
    LastModified = OldDate

End Sub

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Intialize all data to defaults.
'
    
    Willpower = 2
    Arete = 1
    TempWillpower = 2
    TempArete = 1
    PhysicalMax = 11
    SocialMax = 11
    MentalMax = 11
    
    Status = ActiveStatus
    StartDate = Now
    LastModified = Now
    
    Set PhysicalList = New LinkedTraitList
    Set SocialList = New LinkedTraitList
    Set MentalList = New LinkedTraitList
    Set PhysicalNegList = New LinkedTraitList
    Set SocialNegList = New LinkedTraitList
    Set MentalNegList = New LinkedTraitList

    Set AbilityList = New LinkedTraitList
    Set InfluenceList = New LinkedTraitList

    Set BackgroundList = New LinkedTraitList
    Set HealthList = New LinkedTraitList
    Set ResonanceList = New LinkedTraitList
    Set ReputationList = New LinkedTraitList
    Set RoteList = New LinkedTraitList
    Set SphereList = New LinkedTraitList

    Set MeritList = New LinkedTraitList
    Set FlawList = New LinkedTraitList

    Set EquipmentList = New LinkedTraitList
    Set HangoutList = New LinkedTraitList
    Set Experience = New ExperienceClass

    PhysicalList.Initialize "Physical", True, False, False, ldMultiplier
    SocialList.Initialize "Social", True, False, False, ldMultiplier
    MentalList.Initialize "Mental", True, False, False, ldMultiplier
    PhysicalNegList.Initialize "Negative Physical", True, True, False, ldMultiplier
    SocialNegList.Initialize "Negative Social", True, True, False, ldMultiplier
    MentalNegList.Initialize "Negative Mental", True, True, False, ldMultiplier
    
    AbilityList.Initialize "Abilities", True, False, False, ldMultiplier
    InfluenceList.Initialize "Influences", True, False, False, ldMultiplier
    BackgroundList.Initialize "Backgrounds", True, False, False, ldMultiplier
    HealthList.Initialize "Health Levels", False, False, False, ldMultiplier
    
    ResonanceList.Initialize "Resonance", True, False, False, ldMultiplier
    ReputationList.Initialize "Reputation", True, False, False, ldMultiplier
    RoteList.Initialize "Rotes", True, False, True, ldNoteOnly
    SphereList.Initialize "Spheres", False, False, True, ldNoteOnly
    
    MeritList.Initialize "Merits", True, False, True, ldCost
    FlawList.Initialize "Flaws", True, True, True, ldCost
    
    EquipmentList.Initialize "Equipment", True, False, False, ldMultiplier
    HangoutList.Initialize "Locations", True, False, True, ldNoteOnly
    
    With HealthList
        .Append StdHealth(0)
        .Append StdHealth(1)
        .Append StdHealth(2)
        .Append StdHealth(3)
        .Append StdHealth(4)
        If Game.ExtendedHealth Then
            .Append StdHealth(0)
            .Append StdHealth(1)
            .Append StdHealth(1)
            .Append StdHealth(2)
        End If
    End With
    
End Sub

Private Sub Class_Terminate()
'
' Name:         Class_Terminate
' Description:  Destroy all objects created in this class.
'

    Set PhysicalList = Nothing
    Set SocialList = Nothing
    Set MentalList = Nothing
    Set PhysicalNegList = Nothing
    Set SocialNegList = Nothing
    Set MentalNegList = Nothing

    Set AbilityList = Nothing
    Set InfluenceList = Nothing

    Set BackgroundList = Nothing
    Set HealthList = Nothing
    Set ResonanceList = Nothing
    Set ReputationList = Nothing
    Set RoteList = Nothing
    Set SphereList = Nothing

    Set MeritList = Nothing
    Set FlawList = Nothing

    Set EquipmentList = Nothing
    Set HangoutList = Nothing
    Set Experience = Nothing

End Sub
