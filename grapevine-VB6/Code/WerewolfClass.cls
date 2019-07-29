VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WerewolfClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Nature As String
Public Demeanor As String
Public Tribe As String
Public Breed As String
Public Auspice As String
Public Rank As String
Public Pack As String
Public Totem As String
Public Camp As String
Public Position As String
Public Notoriety As Integer

Public Rage As Integer
Public Gnosis As Integer
Public Willpower As Integer

Public TempRage As Integer
Public TempGnosis As Integer
Public TempWillpower As Integer

Public Honor As Integer
Public Glory As Integer
Public Wisdom As Integer
Public TempHonor As Single
Public TempGlory As Single
Public TempWisdom As Single

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
Public FeatureList As LinkedTraitList

Public GiftList As LinkedTraitList
Public RiteList As LinkedTraitList

Public HonorList As LinkedTraitList
Public GloryList As LinkedTraitList
Public WisdomList As LinkedTraitList

Public MeritList As LinkedTraitList
Public FlawList As LinkedTraitList

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
    FeatureList.First
    GiftList.First
    RiteList.First
    HonorList.First
    GloryList.First
    WisdomList.First
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
        
        Case qkRank:            CharData = Rank
        Case qkPack:            CharData = Pack
        Case qkHonor:           CharData = Honor
        Case qkGlory:           CharData = Glory
        Case qkWisdom:          CharData = Wisdom
        Case qkTempHonor:       CharData = TempHonor
        Case qkTempGlory:       CharData = TempGlory
        Case qkTempWisdom:      CharData = TempWisdom
        Case qkHonorTraits:     Set CharData = HonorList
        Case qkGloryTraits:     Set CharData = GloryList
        Case qkWisdomTraits:    Set CharData = WisdomList
        Case qkGifts:           Set CharData = GiftList
        Case qkRites:           Set CharData = RiteList
        Case qkTribe:           CharData = Tribe
        Case qkBreed:           CharData = Breed
        Case qkAuspice:         CharData = Auspice
        Case qkRage:            CharData = Rage
        Case qkGnosis:          CharData = Gnosis
        Case qkNotoriety:       CharData = Notoriety
        Case qkPosition:        CharData = Position
        Case qkTotem:           CharData = Totem
        Case qkCamp:            CharData = Camp
        Case qkFeatures:        Set CharData = FeatureList
        Case qkBiography:       CharData = Biography
        Case qkPhysicalMax:     CharData = PhysicalMax
        Case qkSocialMax:       CharData = SocialMax
        Case qkMentalMax:       CharData = MentalMax
        Case qkTempRage:        CharData = TempRage
        Case qkTempGnosis:      CharData = TempGnosis
        
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
        Case qkRage:            Rage = CInt(CharData)
        Case qkGnosis:          Gnosis = CInt(CharData)
        Case qkTempRage:        TempRage = CInt(CharData)
        Case qkTempGnosis:      TempGnosis = CInt(CharData)
        Case qkHonor:           Honor = CInt(CharData)
        Case qkGlory:           Glory = CInt(CharData)
        Case qkWisdom:          Wisdom = CInt(CharData)
        Case qkTempHonor:       TempHonor = CSng(CharData)
        Case qkTempGlory:       TempGlory = CSng(CharData)
        Case qkTempWisdom:      TempWisdom = CSng(CharData)
        
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
    Estimate.Append qkGifts, , "9"
    Estimate.Append qkRites
    Estimate.Append qkRage, "3", "3"
    Estimate.Append qkWillpower, "3", "3"
    Estimate.Append qkGnosis, "3", "3"
    Estimate.Append qkFeatures, "0"
        
End Sub

Public Function Group() As String
'
' Name:         Group
' Description:  Return the 'Group' this character belongs to.
'               Clan, Tribe, Kith, Guild, etc.
' Returns:      the group
'

    Group = Tribe

End Function

Public Function Subgroup() As String
'
' Name:         Subgroup
' Description:  Return the 'Subgroup' this character belongs to.
'               Sect, Auspice, Seeming, etc.
' Returns:      the subgroup
'

    Subgroup = Auspice

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

    Race = "Werewolf"

End Function

Public Function RaceCode() As RaceType
'
' Name:         RaceCode
' Description:  Return the RaceType (enumeration value) of this
'               character type.
' Returns:      the corresponding RaceType
'
    
    RaceCode = gvRaceWerewolf

End Function

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write this character's data to a file.
'

    With XML
    
        .BeginTag "werewolf"
            .WriteAttribute "name", Name
            .WriteAttribute "nature", Nature
            .WriteAttribute "demeanor", Demeanor
            .WriteAttribute "tribe", Tribe
            .WriteAttribute "breed", Breed
            .WriteAttribute "auspice", Auspice
            .WriteAttribute "rank", Rank
            .WriteAttribute "pack", Pack
            .WriteAttribute "totem", Totem
            .WriteAttribute "camp", Camp
            .WriteAttribute "position", Position
            .WriteAttribute "notoriety", Notoriety, 0
            .WriteAttribute "rage", Rage
            .WriteAttribute "temprage", TempRage, Rage
            .WriteAttribute "gnosis", Gnosis
            .WriteAttribute "tempgnosis", TempGnosis, Gnosis
            .WriteAttribute "willpower", Willpower
            .WriteAttribute "tempwillpower", TempWillpower, Willpower
            .WriteAttribute "honor", Honor
            .WriteAttribute "temphonor", TempHonor
            .WriteAttribute "glory", Glory
            .WriteAttribute "tempglory", TempGlory
            .WriteAttribute "wisdom", Wisdom
            .WriteAttribute "tempwisdom", TempWisdom
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
            FeatureList.OutputToFile XML
            
            GiftList.OutputToFile XML
            RiteList.OutputToFile XML
            
            HonorList.OutputToFile XML
            GloryList.OutputToFile XML
            WisdomList.OutputToFile XML
            
            MeritList.OutputToFile XML
            FlawList.OutputToFile XML
            
            EquipmentList.OutputToFile XML
            HangoutList.OutputToFile XML
            
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
    
        If .Tag = "werewolf" Then
        
            Name = .GetString("name")
            Nature = .GetString("nature")
            Demeanor = .GetString("demeanor")
            Tribe = .GetString("tribe")
            Breed = .GetString("breed")
            Auspice = .GetString("auspice")
            Rank = .GetString("rank")
            Pack = .GetString("pack")
            Totem = .GetString("totem")
            Camp = .GetString("camp")
            Position = .GetString("position")
            Notoriety = .GetInteger("notoriety")
            Rage = .GetInteger("rage")
            TempRage = .GetInteger("temprage", Rage)
            Gnosis = .GetInteger("gnosis")
            TempGnosis = .GetInteger("tempgnosis", Gnosis)
            Willpower = .GetInteger("willpower")
            TempWillpower = .GetInteger("tempwillpower", Willpower)
            
            If Version >= 2.395 Then
                Honor = .GetInteger("honor")
                Glory = .GetInteger("glory")
                Wisdom = .GetInteger("wisdom")
                TempHonor = .GetSingle("temphonor")
                TempGlory = .GetSingle("tempglory")
                TempWisdom = .GetSingle("tempwisdom")
            Else
                TempHonor = .GetSingle("honor")
                Honor = Int(TempHonor)
                TempHonor = Round((TempHonor - Honor) * 10, 1)
                TempGlory = .GetSingle("glory")
                Glory = Int(TempGlory)
                TempGlory = Round((TempGlory - Glory) * 10, 1)
                TempWisdom = .GetSingle("wisdom")
                Wisdom = Int(TempWisdom)
                TempWisdom = Round((TempWisdom - Wisdom) * 10, 1)
            End If
            
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
            
            Do Until .Tag = "werewolf" And .Terminal
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
                            Case "Features":            FeatureList.InputFromFile XML, Version
                            Case "Gifts":               GiftList.InputFromFile XML, Version
                            Case "Rites":               RiteList.InputFromFile XML, Version
                            Case "Honor":               HonorList.InputFromFile XML, Version
                            Case "Glory":               GloryList.InputFromFile XML, Version
                            Case "Wisdom":              WisdomList.InputFromFile XML, Version
                            Case "Merits":              MeritList.InputFromFile XML, Version
                            Case "Flaws":               FlawList.InputFromFile XML, Version
                            Case "Equipment":           EquipmentList.InputFromFile XML, Version
                            Case "Locations":           HangoutList.InputFromFile XML, Version
                        End Select
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
            Err.Raise vbObjectError + 8, "WerewolfClass", _
                      "Wrong file input to a werewolf character"
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
    PutStrB FileNum, Tribe
    PutStrB FileNum, Breed
    PutStrB FileNum, Auspice
    PutStrB FileNum, Rank
    PutStrB FileNum, Pack
    PutStrB FileNum, Totem
    PutStrB FileNum, Camp
    PutStrB FileNum, Position
    Put #FileNum, , Notoriety
    Put #FileNum, , Rage
    Put #FileNum, , TempRage
    Put #FileNum, , Gnosis
    Put #FileNum, , TempGnosis
    Put #FileNum, , Willpower
    Put #FileNum, , TempWillpower
    Put #FileNum, , Honor
    Put #FileNum, , Glory
    Put #FileNum, , Wisdom
    Put #FileNum, , TempHonor
    Put #FileNum, , TempGlory
    Put #FileNum, , TempWisdom
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
    FeatureList.OutputToBinary FileNum
    
    GiftList.OutputToBinary FileNum
    RiteList.OutputToBinary FileNum
    
    HonorList.OutputToBinary FileNum
    GloryList.OutputToBinary FileNum
    WisdomList.OutputToBinary FileNum
    
    MeritList.OutputToBinary FileNum
    FlawList.OutputToBinary FileNum
    
    EquipmentList.OutputToBinary FileNum
    HangoutList.OutputToBinary FileNum
    
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
    GetStrB FileNum, Tribe
    GetStrB FileNum, Breed
    GetStrB FileNum, Auspice
    GetStrB FileNum, Rank
    GetStrB FileNum, Pack
    GetStrB FileNum, Totem
    GetStrB FileNum, Camp
    GetStrB FileNum, Position
    Get #FileNum, , Notoriety
    
    If Version >= 2.397 Then
        Get #FileNum, , Rage
        Get #FileNum, , TempRage
        Get #FileNum, , Gnosis
        Get #FileNum, , TempGnosis
        Get #FileNum, , Willpower
        Get #FileNum, , TempWillpower
    Else
        Get #FileNum, , Rage
        Get #FileNum, , Gnosis
        Get #FileNum, , Willpower
        TempRage = Rage
        TempGnosis = Gnosis
        TempWillpower = Willpower
    End If
    
    If Version >= 2.395 Then
        Get #FileNum, , Honor
        Get #FileNum, , Glory
        Get #FileNum, , Wisdom
        Get #FileNum, , TempHonor
        Get #FileNum, , TempGlory
        Get #FileNum, , TempWisdom
    Else
        Get #FileNum, , TempHonor
        Honor = Int(TempHonor)
        TempHonor = Round((TempHonor - Honor) * 10, 1)
        Get #FileNum, , TempGlory
        Glory = Int(TempGlory)
        TempGlory = Round((TempGlory - Glory) * 10, 1)
        Get #FileNum, , TempWisdom
        Wisdom = Int(TempWisdom)
        TempWisdom = Round((TempWisdom - Wisdom) * 10, 1)
    End If
    
    If Version >= 2.397 Then
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
    InfluenceList.InputFromBinary FileNum, Version
    BackgroundList.InputFromBinary FileNum, Version
    
    HealthList.InputFromBinary FileNum, Version
    FeatureList.InputFromBinary FileNum, Version
    
    GiftList.InputFromBinary FileNum, Version
    RiteList.InputFromBinary FileNum, Version
    
    HonorList.InputFromBinary FileNum, Version
    GloryList.InputFromBinary FileNum, Version
    WisdomList.InputFromBinary FileNum, Version
    
    MeritList.InputFromBinary FileNum, Version
    FlawList.InputFromBinary FileNum, Version
    
    EquipmentList.InputFromBinary FileNum, Version
    If Version >= 2.395 Then HangoutList.InputFromBinary FileNum, Version
    
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
    Line Input #FileNum, Tribe
    Line Input #FileNum, Breed
    Line Input #FileNum, Auspice
    If Version >= GameFileVersionTag5 Then
        Line Input #FileNum, Rank
    Else
        Line Input #FileNum, Read
        Select Case Val(Read)
            Case 0: Rank = "Cub"
            Case 1: Rank = "Cliath"
            Case 2: Rank = "Fostern"
            Case 3: Rank = "Adren"
            Case 4: Rank = "Athro"
            Case 5: Rank = "Elder"
            Case 6: Rank = "Legend"
        End Select
    End If
    Line Input #FileNum, Pack
    Line Input #FileNum, Position
    If Version >= GameFileVersionTag5 Then
        Line Input #FileNum, Read
        Notoriety = Val(Read)
    End If
    Line Input #FileNum, Read
    Rage = Val(Read)
    TempRage = Rage
    Line Input #FileNum, Read
    Gnosis = Val(Read)
    TempGnosis = Gnosis
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
    InfluenceList.OldInputFromFile FileNum, Version
    
    BackgroundList.OldInputFromFile FileNum, Version
    If Version >= GameFileVersionTag4 Then
        HealthList.OldInputFromFile FileNum, Version
    End If
    If Version >= GameFileVersionTag5 Then
        FeatureList.LLInputFromFile FileNum, Version
    End If
    
    GiftList.LLInputFromFile FileNum, Version
    RiteList.LLInputFromFile FileNum, Version
    
    HonorList.OldInputFromFile FileNum, Version
    GloryList.OldInputFromFile FileNum, Version
    WisdomList.OldInputFromFile FileNum, Version
    
    Honor = HonorList.Count
    Glory = GloryList.Count
    Wisdom = WisdomList.Count
    
    MeritList.LLInputFromFile FileNum, Version
    FlawList.LLInputFromFile FileNum, Version
    
    Notes = ""
    If Version <> GameFileVersionTag0 Then
        Read = ReadLongField(FileNum, "<-End Werewolf Character Equipment->")
        If Read <> "" Then Notes = "2.3 Equipment:" & vbCrLf & Read & vbCrLf
    End If
    
    Notes = Notes & ReadLongField(FileNum, "<-End Werewolf Character Notes->")
    
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
    If Version >= GameFileVersionTag3 Then
        Line Input #FileNum, ID
        Line Input #FileNum, Narrator
        If IsDate(Narrator) Then StartDate = CDate(Narrator)
        Line Input #FileNum, Narrator
        Line Input #FileNum, Read
        IsNPC = CBool(Read)
    Else
        StartDate = ""
        If Status = "" Then Status = ActiveStatus
    End If
    Experience.OldInputFromFile FileNum, Version
    LastModified = OldDate

End Sub

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Intialize all data to defaults.
'

    Rage = 3
    Gnosis = 3
    Willpower = 3
    TempRage = 3
    TempGnosis = 3
    TempWillpower = 3
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
    Set FeatureList = New LinkedTraitList
    
    Set GiftList = New LinkedTraitList
    Set RiteList = New LinkedTraitList
    
    Set HonorList = New LinkedTraitList
    Set GloryList = New LinkedTraitList
    Set WisdomList = New LinkedTraitList
    
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
    
    FeatureList.Initialize "Features", True, False, True, ldNoteOnly
    GiftList.Initialize "Gifts", False, False, True, ldNoteOnly
    RiteList.Initialize "Rites", False, False, True, ldNoteOnly
    HonorList.Initialize "Honor", True, False, False, ldMultiplier
    GloryList.Initialize "Glory", True, False, False, ldMultiplier
    WisdomList.Initialize "Wisdom", True, False, False, ldMultiplier
    
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
    Set FeatureList = Nothing
    
    Set GiftList = Nothing
    Set RiteList = Nothing

    Set HonorList = Nothing
    Set GloryList = Nothing
    Set WisdomList = Nothing

    Set MeritList = Nothing
    Set FlawList = Nothing

    Set EquipmentList = Nothing
    Set HangoutList = Nothing
    Set Experience = Nothing

End Sub
