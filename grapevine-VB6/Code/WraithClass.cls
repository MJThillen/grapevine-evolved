VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WraithClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EthnosType
    etWraith = 0
    etRisen = 1
    etSpectre = 2
End Enum

Public Name As String
Public Ethnos As EthnosType
Public Nature As String
Public Demeanor As String
Public Guild As String
Public Faction As String
Public Legion As String
Public Rank As String

Public Pathos As Integer
Public Corpus As Integer
Public Willpower As Integer

Public TempPathos As Integer
Public TempCorpus As Integer
Public TempWillpower As Integer

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
Public BackgroundList As LinkedTraitList
Public StatusList As LinkedTraitList
Public InfluenceList As LinkedTraitList

Public ArcanoiList As LinkedTraitList
Public MeritList As LinkedTraitList
Public FlawList As LinkedTraitList

Public Passions As String
Public Fetters As String
Public Life As String
Public Death As String
Public Haunt As String
Public Regret As String

Public ShadowArchetype As String
Public ShadowPlayer As String
Public Angst As Integer
Public TempAngst As Integer
Public DarkPassions As String
Public ThornList As LinkedTraitList

Public EquipmentList As LinkedTraitList
Public HangoutList As LinkedTraitList
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
    BackgroundList.First
    StatusList.First
    InfluenceList.First
    ArcanoiList.First
    MeritList.First
    FlawList.First
    ThornList.First
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
        Case qkStatus:          Set CharData = StatusList
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
        Case qkNotes:           CharData = Notes
        Case qkID:              CharData = ID
        Case qkPlayer:          CharData = Player
        Case qkNarrator:        CharData = Narrator
        Case qkName:            CharData = Name
        Case qkEMail:           CharData = EMail
        Case qkRandom:          CharData = CInt(Rnd() * 100)
        
        Case qkArcanoi:         Set CharData = ArcanoiList
        Case qkThorns:          Set CharData = ThornList
        Case qkPathos:          CharData = Pathos
        Case qkCorpus:          CharData = Corpus
        Case qkAngst:           CharData = Angst
        Case qkGuild:           CharData = Guild
        Case qkEthnos:          CharData = Group
        Case qkFaction:         CharData = Faction
        Case qkLegion:          CharData = Legion
        Case qkRank:            CharData = Rank
        Case qkPassions:        CharData = Passions
        Case qkFetters:         CharData = Fetters
        Case qkLife:            CharData = Life
        Case qkDeath:           CharData = Death
        Case qkHaunt:           CharData = Haunt
        Case qkRegret:          CharData = Regret
        Case qkShadowArchetype: CharData = ShadowArchetype
        Case qkShadowPlayer:    CharData = ShadowPlayer
        Case qkDarkPassions:    CharData = DarkPassions
        Case qkBiography:       CharData = Life & vbCrLf & vbCrLf & Death
        Case qkPhysicalMax:     CharData = PhysicalMax
        Case qkSocialMax:       CharData = SocialMax
        Case qkMentalMax:       CharData = MentalMax
        Case qkTempPathos:      CharData = TempPathos
        Case qkTempCorpus:      CharData = TempCorpus
        Case qkTempAngst:       CharData = TempAngst
        
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
        Case qkPathos:          Pathos = CInt(CharData)
        Case qkCorpus:          Corpus = CInt(CharData)
        Case qkAngst:           Angst = CInt(CharData)
        Case qkTempPathos:      TempPathos = CInt(CharData)
        Case qkTempCorpus:      TempCorpus = CInt(CharData)
        Case qkTempAngst:       TempAngst = CInt(CharData)
        
    End Select

End Sub

Public Sub CompleteEstimateList(Estimate As LinkedTraitList)
'
' Name:         CompleteEstimateList
' Parameter:    Estimate        Traitlist that will hold estimate info
' Description:  Add race-specific entries to a LinkedTraitList for use with the Point Estimate Tool.
'

    Estimate.Append qkStatus
    Estimate.Append qkMerits
    Estimate.Append qkFlaws
    Estimate.Append qkArcanoi, , "12"
    Estimate.Append qkPathos, "3", "2"
    Estimate.Append qkWillpower, "3", "2"
    Estimate.Append qkCorpus, "3", Corpus
    Estimate.Append qkAngst, "-2", Angst

End Sub

Public Function Group() As String
'
' Name:         Group
' Description:  Return the 'Group' this character belongs to.
'               Clan, Tribe, Kith, etc.
' Returns:      the group
'

    Select Case Ethnos
        Case etWraith
            Group = "Wraith"
        Case etRisen
            Group = "Risen"
        Case etSpectre
            Group = "Spectre"
    End Select

End Function

Public Function Subgroup() As String
'
' Name:         Subgroup
' Description:  Return the 'Subgroup' this character belongs to.
'               Sect, Auspice, Seeming, etc.
' Returns:      the subgroup
'

    Subgroup = Guild

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

Public Function Race() As String
'
' Name:         Race
' Description:  Return a string representation of this character's
'               character type.
' Returns:      the name of the character type
'

    Race = "Wraith"

End Function

Public Function RaceCode() As RaceType
'
' Name:         RaceCode
' Description:  Return the RaceType (enumeration value) of this
'               character type.
' Returns:      the corresponding RaceType
'
    
    RaceCode = gvRaceWraith

End Function

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write this character's data to a file.
'

    With XML
        
        .BeginTag "wraith"

            .WriteAttribute "name", Name
            Select Case Ethnos
                Case etWraith
                    .WriteAttribute "ethnos", "Wraith"
                Case etRisen
                    .WriteAttribute "ethnos", "Risen"
                Case etSpectre
                    .WriteAttribute "ethnos", "Spectre"
            End Select
            .WriteAttribute "nature", Nature
            .WriteAttribute "demeanor", Demeanor
            .WriteAttribute "guild", Guild
            .WriteAttribute "faction", Faction
            .WriteAttribute "legion", Legion
            .WriteAttribute "rank", Rank
            .WriteAttribute "pathos", Pathos
            .WriteAttribute "temppathos", TempPathos, Pathos
            .WriteAttribute "corpus", Corpus
            .WriteAttribute "tempcorpus", TempCorpus, Corpus
            .WriteAttribute "willpower", Willpower
            .WriteAttribute "tempwillpower", TempWillpower, Willpower
            .WriteAttribute "shadowarchetype", ShadowArchetype
            .WriteAttribute "shadowplayer", ShadowPlayer
            .WriteAttribute "angst", Angst
            .WriteAttribute "tempangst", TempAngst, Angst
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
            BackgroundList.OutputToFile XML
            StatusList.OutputToFile XML
            InfluenceList.OutputToFile XML
            
            .WriteCDataTag "passions", Passions
            .WriteCDataTag "fetters", Fetters
            .WriteCDataTag "life", Life
            .WriteCDataTag "death", Death
            .WriteCDataTag "haunt", Haunt
            .WriteCDataTag "regret", Regret
            
            ArcanoiList.OutputToFile XML
            
            MeritList.OutputToFile XML
            FlawList.OutputToFile XML
            EquipmentList.OutputToFile XML
            HangoutList.OutputToFile XML
    
            .WriteCDataTag "darkpassions", DarkPassions
            ThornList.OutputToFile XML
            
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
    
        If .Tag = "wraith" Then
        
            Name = .GetString("name")
            Select Case .GetString("ethnos")
                Case "Wraith":  Ethnos = etWraith
                Case "Risen":   Ethnos = etRisen
                Case "Spectre": Ethnos = etSpectre
            End Select
            Nature = .GetString("nature")
            Demeanor = .GetString("demeanor")
            Guild = .GetString("guild")
            Faction = .GetString("faction")
            Legion = .GetString("legion")
            Rank = .GetString("rank")
            ShadowArchetype = .GetString("shadowarchetype")
            ShadowPlayer = .GetString("shadowplayer")
            Pathos = .GetInteger("pathos")
            TempPathos = .GetInteger("temppathos", Pathos)
            Corpus = .GetInteger("corpus")
            TempCorpus = .GetInteger("tempcorpus", Corpus)
            Angst = .GetInteger("angst")
            TempAngst = .GetInteger("tempangst", Angst)
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
            
            Do Until .Tag = "wraith" And .Terminal
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
                            Case "Status":              StatusList.InputFromFile XML, Version
                            Case "Arcanoi":             ArcanoiList.InputFromFile XML, Version
                            Case "Thorns":              ThornList.InputFromFile XML, Version
                            Case "Merits":              MeritList.InputFromFile XML, Version
                            Case "Flaws":               FlawList.InputFromFile XML, Version
                            Case "Equipment":           EquipmentList.InputFromFile XML, Version
                            Case "Locations":           HangoutList.InputFromFile XML, Version
                        End Select
                    Case "passions"
                        If Not .Terminal Then
                            .ReadTag
                            Passions = .GetTagText
                        End If
                    Case "fetters"
                        If Not .Terminal Then
                            .ReadTag
                            Fetters = .GetTagText
                        End If
                    Case "haunt"
                        If Not .Terminal Then
                            .ReadTag
                            Haunt = .GetTagText
                        End If
                    Case "life"
                        If Not .Terminal Then
                            .ReadTag
                            Life = .GetTagText
                        End If
                    Case "death"
                        If Not .Terminal Then
                            .ReadTag
                            Death = .GetTagText
                        End If
                    Case "regret"
                        If Not .Terminal Then
                            .ReadTag
                            Regret = .GetTagText
                        End If
                    Case "darkpassions"
                        If Not .Terminal Then
                            .ReadTag
                            DarkPassions = .GetTagText
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
            Err.Raise vbObjectError + 8, "WraithClass", _
                      "Wrong file input to Wraith character"
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
    Put #FileNum, , Ethnos
    PutStrB FileNum, Nature
    PutStrB FileNum, Demeanor
    PutStrB FileNum, Guild
    PutStrB FileNum, Faction
    PutStrB FileNum, Legion
    PutStrB FileNum, Rank
    Put #FileNum, , Pathos
    Put #FileNum, , TempPathos
    Put #FileNum, , Corpus
    Put #FileNum, , TempCorpus
    Put #FileNum, , Willpower
    Put #FileNum, , TempWillpower
    PutStrB FileNum, ShadowArchetype
    PutStrB FileNum, ShadowPlayer
    Put #FileNum, , Angst
    Put #FileNum, , TempAngst
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
    BackgroundList.OutputToBinary FileNum
    StatusList.OutputToBinary FileNum
    InfluenceList.OutputToBinary FileNum
    
    PutStrB FileNum, Passions
    PutStrB FileNum, Fetters
    PutStrB FileNum, Life
    PutStrB FileNum, Death
    PutStrB FileNum, Haunt
    PutStrB FileNum, Regret
    
    ArcanoiList.OutputToBinary FileNum
    
    MeritList.OutputToBinary FileNum
    FlawList.OutputToBinary FileNum
    EquipmentList.OutputToBinary FileNum
    HangoutList.OutputToBinary FileNum

    PutStrB FileNum, DarkPassions
    ThornList.OutputToBinary FileNum
    
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
    Get #FileNum, , Ethnos
    GetStrB FileNum, Nature
    GetStrB FileNum, Demeanor
    GetStrB FileNum, Guild
    GetStrB FileNum, Faction
    GetStrB FileNum, Legion
    GetStrB FileNum, Rank
    
    If Version >= 2.397 Then
        Get #FileNum, , Pathos
        Get #FileNum, , TempPathos
        Get #FileNum, , Corpus
        Get #FileNum, , TempCorpus
        Get #FileNum, , Willpower
        Get #FileNum, , TempWillpower
    Else
        Get #FileNum, , Pathos
        Get #FileNum, , Corpus
        Get #FileNum, , Willpower
        TempPathos = Pathos
        TempCorpus = Corpus
        TempWillpower = Willpower
    End If
    
    GetStrB FileNum, ShadowArchetype
    GetStrB FileNum, ShadowPlayer
    Get #FileNum, , Angst
    If Version >= 2.397 Then
        Get #FileNum, , TempAngst
        Get #FileNum, , PhysicalMax
        Get #FileNum, , SocialMax
        Get #FileNum, , MentalMax
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
    BackgroundList.InputFromBinary FileNum, Version
    StatusList.InputFromBinary FileNum, Version
    InfluenceList.InputFromBinary FileNum, Version
    
    GetStrB FileNum, Passions
    GetStrB FileNum, Fetters
    GetStrB FileNum, Life
    GetStrB FileNum, Death
    GetStrB FileNum, Haunt
    GetStrB FileNum, Regret
    
    ArcanoiList.InputFromBinary FileNum, Version
    
    MeritList.InputFromBinary FileNum, Version
    FlawList.InputFromBinary FileNum, Version
    EquipmentList.InputFromBinary FileNum, Version
    If Version >= 2.395 Then HangoutList.InputFromBinary FileNum, Version

    GetStrB FileNum, DarkPassions
    ThornList.InputFromBinary FileNum, Version
    
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
' Description:  Read a character from an input file.  The version
'               ensures backward-compatibility.
'

    Dim Read As String
    
    ' Race has already been read
    Line Input #FileNum, Name
    Line Input #FileNum, Read
    Ethnos = Val(Read)
    Line Input #FileNum, Nature
    Line Input #FileNum, Demeanor
    Line Input #FileNum, Guild
    Line Input #FileNum, Faction
    Line Input #FileNum, Legion
    Line Input #FileNum, Rank
    
    Line Input #FileNum, Read
    Pathos = Val(Read)
    TempPathos = Pathos
    Line Input #FileNum, Read
    Corpus = Val(Read)
    TempCorpus = Corpus
    Line Input #FileNum, Read
    Willpower = Val(Read)
    TempWillpower = Willpower
    
    PhysicalList.OldInputFromFile FileNum, Version
    SocialList.OldInputFromFile FileNum, Version
    MentalList.OldInputFromFile FileNum, Version
    PhysicalNegList.OldInputFromFile FileNum, Version
    SocialNegList.OldInputFromFile FileNum, Version
    MentalNegList.OldInputFromFile FileNum, Version
    
    AbilityList.OldInputFromFile FileNum, Version
    BackgroundList.OldInputFromFile FileNum, Version
    StatusList.OldInputFromFile FileNum, Version
    InfluenceList.OldInputFromFile FileNum, Version
    
    ArcanoiList.LLInputFromFile FileNum, Version
    
    Passions = ReadLongField(FileNum, "<-End Wraith Character Passions->")
    Fetters = ReadLongField(FileNum, "<-End Wraith Character Fetters->")
    Life = ReadLongField(FileNum, "<-End Wraith Character Life->")
    Death = ReadLongField(FileNum, "<-End Wraith Character Death->")
    Haunt = ReadLongField(FileNum, "<-End Wraith Character Haunt->")
    Regret = ReadLongField(FileNum, "<-End Wraith Character Regret->")
    
    Line Input #1, ShadowArchetype
    Line Input #1, ShadowPlayer
    Line Input #1, Read
    Angst = Val(Read)
    TempAngst = Angst
    DarkPassions = ReadLongField(FileNum, "<-End Wraith Character DarkPassions->")
    ThornList.LLInputFromFile FileNum, Version
    
    Notes = ""
    Read = ReadLongField(FileNum, "<-End Wraith Character Equipment->")
    If Read <> "" Then Notes = "2.3 Equipment:" & vbCrLf & Read & vbCrLf
    
    Notes = Notes & ReadLongField(FileNum, "<-End Wraith Character Notes->")
    
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

    Ethnos = etWraith
    Pathos = 2
    Corpus = 2
    Angst = 2
    Willpower = 2
    TempPathos = 2
    TempCorpus = 2
    TempWillpower = 2
    TempAngst = 2
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
    Set BackgroundList = New LinkedTraitList
    Set StatusList = New LinkedTraitList
    Set InfluenceList = New LinkedTraitList
    
    Set ArcanoiList = New LinkedTraitList
    Set MeritList = New LinkedTraitList
    Set FlawList = New LinkedTraitList
    
    Set ThornList = New LinkedTraitList
    
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
    StatusList.Initialize "Status", True, False, False, ldMultiplier
    
    ArcanoiList.Initialize "Arcanoi", False, False, True, ldNoteOnly
    MeritList.Initialize "Merits", True, False, True, ldCost
    FlawList.Initialize "Flaws", True, True, True, ldCost
    ThornList.Initialize "Thorns", True, True, True, ldCost
    
    EquipmentList.Initialize "Equipment", True, False, False, ldMultiplier
    HangoutList.Initialize "Locations", True, False, True, ldNoteOnly

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
    Set BackgroundList = Nothing
    Set StatusList = Nothing
    Set InfluenceList = Nothing

    Set ArcanoiList = Nothing
    Set MeritList = Nothing
    Set FlawList = Nothing
    
    Set ThornList = Nothing

    Set EquipmentList = Nothing
    Set HangoutList = Nothing
    Set Experience = Nothing

End Sub
