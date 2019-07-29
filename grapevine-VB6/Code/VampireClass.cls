VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VampireClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Nature As String
Public Demeanor As String
Public Clan As String
Public Sect As String
Public Generation As Integer
Public Title As String
Public Coterie As String
Public Path As String
Public Sire As String
Public Aura As String
Public AuraBonus As String

Public Blood As Integer
Public Willpower As Integer
Public Conscience As Integer
Public SelfControl As Integer
Public Courage As Integer
Public PathTraits As Integer

Public TempBlood As Integer
Public TempWillpower As Integer
Public TempConscience As Integer
Public TempSelfControl As Integer
Public TempCourage As Integer
Public TempPathTraits As Integer

Public PhysicalList As LinkedTraitList
Public SocialList As LinkedTraitList
Public MentalList As LinkedTraitList
Public PhysicalNegList As LinkedTraitList
Public SocialNegList As LinkedTraitList
Public MentalNegList As LinkedTraitList
Public PhysicalMax As Integer
Public SocialMax As Integer
Public MentalMax As Integer

Public StatusList As LinkedTraitList

Public AbilityList As LinkedTraitList
Public InfluenceList As LinkedTraitList

Public BackgroundList As LinkedTraitList
Public HealthList As LinkedTraitList
Public BondList As LinkedTraitList

Public MiscellaneousList As LinkedTraitList
Public DerangementList As LinkedTraitList

Public DisciplineList As LinkedTraitList
Public RitualList As LinkedTraitList

Public MeritList As LinkedTraitList
Public FlawList As LinkedTraitList

Public EquipmentList As LinkedTraitList
Public HangoutList As LinkedTraitList

Public BoonList As LinkedList
Public BoonTraitList As LinkedTraitList

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
    StatusList.First
    AbilityList.First
    InfluenceList.First
    MiscellaneousList.First
    DerangementList.First
    DisciplineList.First
    RitualList.First
    MeritList.First
    FlawList.First
    BackgroundList.First
    HealthList.First
    BondList.First
    EquipmentList.First
    HangoutList.First
    BoonList.First
    BoonTraitList.First
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
        Case qkHealth:          Set CharData = HealthList
        Case qkNotes:           CharData = Notes
        Case qkID:              CharData = ID
        Case qkPlayer:          CharData = Player
        Case qkNarrator:        CharData = Narrator
        Case qkName:            CharData = Name
        Case qkEMail:           CharData = EMail
        Case qkRandom:          CharData = CInt(Rnd() * 100)
        
        Case qkDerangements:    Set CharData = DerangementList
        Case qkDisciplines:     Set CharData = DisciplineList
        Case qkRituals:         Set CharData = RitualList
        Case qkGeneration:      CharData = Generation
        Case qkCoterie:         CharData = Coterie
        Case qkSire:            CharData = Sire
        Case qkAura:            CharData = Aura
        Case qkAuraBonus:       CharData = AuraBonus
        Case qkPathTraits:      CharData = PathTraits
        Case qkConscience:      CharData = Conscience
        Case qkSelfControl:     CharData = SelfControl
        Case qkCourage:         CharData = Courage
        Case qkClan:            CharData = Clan
        Case qkSect:            CharData = Sect
        Case qkPath:            CharData = Path
        Case qkTitle:           CharData = Title
        Case qkBlood:           CharData = Blood
        Case qkBonds:           Set CharData = BondList
        Case qkMiscellaneous:   Set CharData = MiscellaneousList
        Case qkBoons:           Set CharData = BoonTraitList
        Case qkBiography:       CharData = Biography
        Case qkPhysicalMax:     CharData = PhysicalMax
        Case qkSocialMax:       CharData = SocialMax
        Case qkMentalMax:       CharData = MentalMax
        Case qkTempPathTraits:  CharData = TempPathTraits
        Case qkTempConscience:  CharData = TempConscience
        Case qkTempSelfControl: CharData = TempSelfControl
        Case qkTempCourage:     CharData = TempCourage
        Case qkTempBlood:       CharData = TempBlood
        
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
        Case qkPathTraits:      PathTraits = CInt(CharData)
        Case qkConscience:      Conscience = CInt(CharData)
        Case qkSelfControl:     SelfControl = CInt(CharData)
        Case qkCourage:         Courage = CInt(CharData)
        Case qkBlood:           Blood = CInt(CharData)
        Case qkTempPathTraits:  TempPathTraits = CInt(CharData)
        Case qkTempConscience:  TempConscience = CInt(CharData)
        Case qkTempSelfControl: TempSelfControl = CInt(CharData)
        Case qkTempCourage:     TempCourage = CInt(CharData)
        Case qkTempBlood:       TempBlood = CInt(CharData)
        
    End Select

End Sub

Public Sub CompleteEstimateList(Estimate As LinkedTraitList)
'
' Name:         CompleteEstimateList
' Parameter:    Estimate        Traitlist that will hold estimate info
' Description:  Add race-specific entries to a LinkedTraitList for use with the Point Estimate Tool.
'

    Estimate.Append qkDerangements
    Estimate.Append qkMerits
    Estimate.Append qkFlaws
    Estimate.Append qkDisciplines, , "9"
    Estimate.Append qkRituals
    Estimate.Append qkMiscellaneous
    Estimate.Append qkWillpower, "3", "2"
    Estimate.Append qkPathTraits, "2", PathTraits
    Estimate.Append qkConscience, "2", Conscience
    Estimate.Append qkSelfControl, "2", SelfControl
    Estimate.Append qkCourage, "2", Courage

End Sub

Public Function Group() As String
'
' Name:         Group
' Description:  Return the 'Group' this character belongs to.
'               Clan, Tribe, Kith, Guild, etc.
' Returns:      the group
'

    Group = Clan

End Function

Public Function Subgroup() As String
'
' Name:         Subgroup
' Description:  Return the 'Subgroup' this character belongs to.
'               Sect, Auspice, Seeming, etc.
' Returns:      the subgroup
'

    Subgroup = Sect

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

    Race = "Vampire"

End Function

Public Function RaceCode() As RaceType
'
' Name:         RaceCode
' Description:  Return the RaceType (enumeration value) of this
'               character type.
' Returns:      the corresponding RaceType
'
    
    RaceCode = gvRaceVampire

End Function

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write this character's data to a file.
'

    With XML
        
        .BeginTag "vampire"
    
            .WriteAttribute "name", Name
            .WriteAttribute "nature", Nature
            .WriteAttribute "demeanor", Demeanor
            .WriteAttribute "clan", Clan
            .WriteAttribute "sect", Sect
            .WriteAttribute "coterie", Coterie
            .WriteAttribute "sire", Sire
            .WriteAttribute "generation", Generation
            .WriteAttribute "title", Title
            .WriteAttribute "blood", Blood
            .WriteAttribute "tempblood", TempBlood, Blood
            .WriteAttribute "willpower", Willpower
            .WriteAttribute "tempwillpower", TempWillpower, Willpower
            .WriteAttribute "conscience", Conscience
            .WriteAttribute "tempconscience", TempConscience, Conscience
            .WriteAttribute "selfcontrol", SelfControl
            .WriteAttribute "tempselfcontrol", TempSelfControl, SelfControl
            .WriteAttribute "courage", Courage
            .WriteAttribute "tempcourage", TempCourage, Courage
            .WriteAttribute "path", Path
            .WriteAttribute "pathtraits", PathTraits
            .WriteAttribute "temppathtraits", TempPathTraits, PathTraits
            .WriteAttribute "aura", Aura
            .WriteAttribute "aura", AuraBonus, "+0"
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
    
            StatusList.OutputToFile XML
            
            AbilityList.OutputToFile XML
            InfluenceList.OutputToFile XML
            
            BackgroundList.OutputToFile XML
            HealthList.OutputToFile XML
            BondList.OutputToFile XML
            
            MiscellaneousList.OutputToFile XML
            DerangementList.OutputToFile XML
            
            DisciplineList.OutputToFile XML
            RitualList.OutputToFile XML
            
            MeritList.OutputToFile XML
            FlawList.OutputToFile XML
            
            EquipmentList.OutputToFile XML
            HangoutList.OutputToFile XML
            
            BoonList.Memorize
            BoonList.First
            Do Until BoonList.Off
                BoonList.Item.OutputToFile XML
                BoonList.MoveNext
            Loop
            BoonList.Remember
            
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
    Dim Boon As BoonClass
    
    With XML
    
        If .Tag = "vampire" Then
        
            Name = .GetString("name")
            Nature = .GetString("nature")
            Demeanor = .GetString("demeanor")
            Clan = .GetString("clan")
            Sect = .GetString("sect")
            Coterie = .GetString("coterie")
            Sire = .GetString("sire")
            Title = .GetString("title")
            Generation = .GetInteger("generation")
            Blood = .GetInteger("blood")
            TempBlood = .GetInteger("tempblood", Blood)
            Conscience = .GetInteger("conscience")
            TempConscience = .GetInteger("tempconscience", Conscience)
            SelfControl = .GetInteger("selfcontrol")
            TempSelfControl = .GetInteger("tempselfcontrol", SelfControl)
            Courage = .GetInteger("courage")
            TempCourage = .GetInteger("tempcourage", Courage)
            Path = .GetString("path")
            PathTraits = .GetInteger("pathtraits")
            TempPathTraits = .GetInteger("temppathtraits", PathTraits)
            Aura = .GetString("aura")
            AuraBonus = .GetString("aurabonus", "+0")
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
            
            BoonList.Clear
            BoonTraitList.Clear
            
            Do Until .Tag = "vampire" And .Terminal
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
                            Case "Status":              StatusList.InputFromFile XML, Version
                            Case "Bonds":               BondList.InputFromFile XML, Version
                            Case "Beast Traits":        MiscellaneousList.InputFromFile XML, Version
                            Case "Miscellaneous":       MiscellaneousList.InputFromFile XML, Version
                            Case "Derangements":        DerangementList.InputFromFile XML, Version
                            Case "Disciplines":         DisciplineList.InputFromFile XML, Version
                            Case "Rituals":             RitualList.InputFromFile XML, Version
                            Case "Merits":              MeritList.InputFromFile XML, Version
                            Case "Flaws":               FlawList.InputFromFile XML, Version
                            Case "Equipment":           EquipmentList.InputFromFile XML, Version
                            Case "Locations":           HangoutList.InputFromFile XML, Version
                        End Select
                    Case "boon"
                        Set Boon = New BoonClass
                        Boon.InputFromFile XML, Version
                        BoonList.Append Boon
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

            RefreshBoonTraits

        Else
            Err.Raise vbObjectError + 8, "VampireClass", _
                      "Wrong file input to a Vampire character"
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
    PutStrB FileNum, Clan
    PutStrB FileNum, Sect
    PutStrB FileNum, Coterie
    PutStrB FileNum, Sire
    Put #FileNum, , Generation
    PutStrB FileNum, Title
    Put #FileNum, , Blood
    Put #FileNum, , TempBlood
    Put #FileNum, , Willpower
    Put #FileNum, , TempWillpower
    Put #FileNum, , Conscience
    Put #FileNum, , TempConscience
    Put #FileNum, , SelfControl
    Put #FileNum, , TempSelfControl
    Put #FileNum, , Courage
    Put #FileNum, , TempCourage
    PutStrB FileNum, Path
    Put #FileNum, , PathTraits
    Put #FileNum, , TempPathTraits
    PutStrB FileNum, Aura
    PutStrB FileNum, AuraBonus
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

    StatusList.OutputToBinary FileNum
    
    AbilityList.OutputToBinary FileNum
    InfluenceList.OutputToBinary FileNum
    
    BackgroundList.OutputToBinary FileNum
    HealthList.OutputToBinary FileNum
    BondList.OutputToBinary FileNum
    
    MiscellaneousList.OutputToBinary FileNum
    DerangementList.OutputToBinary FileNum
    
    DisciplineList.OutputToBinary FileNum
    RitualList.OutputToBinary FileNum
    
    MeritList.OutputToBinary FileNum
    FlawList.OutputToBinary FileNum
    
    EquipmentList.OutputToBinary FileNum
    HangoutList.OutputToBinary FileNum
    
    Put #FileNum, , BoonList.Count
    BoonList.Memorize
    BoonList.First
    Do Until BoonList.Off
        BoonList.Item.OutputToBinary FileNum
        BoonList.MoveNext
    Loop
    BoonList.Remember
    
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
    GetStrB FileNum, Clan
    GetStrB FileNum, Sect
    If Version >= 2.395 Then GetStrB FileNum, Coterie
    If Version >= 2.399 Then GetStrB FileNum, Sire
    Get #FileNum, , Generation
    GetStrB FileNum, Title
    
    If Version >= 2.397 Then
        Get #FileNum, , Blood
        Get #FileNum, , TempBlood
        Get #FileNum, , Willpower
        Get #FileNum, , TempWillpower
        Get #FileNum, , Conscience
        Get #FileNum, , TempConscience
        Get #FileNum, , SelfControl
        Get #FileNum, , TempSelfControl
        Get #FileNum, , Courage
        Get #FileNum, , TempCourage
        GetStrB FileNum, Path
        Get #FileNum, , PathTraits
        Get #FileNum, , TempPathTraits
        If Version >= 2.399 Then
            GetStrB FileNum, Aura
            GetStrB FileNum, AuraBonus
        End If
        Get #FileNum, , PhysicalMax
        Get #FileNum, , SocialMax
        Get #FileNum, , MentalMax
    Else
        Get #FileNum, , Blood
        Get #FileNum, , Willpower
        Get #FileNum, , Conscience
        Get #FileNum, , SelfControl
        Get #FileNum, , Courage
        GetStrB FileNum, Path
        Get #FileNum, , PathTraits
        TempBlood = Blood
        TempWillpower = Willpower
        TempConscience = Conscience
        TempSelfControl = SelfControl
        TempCourage = Courage
        TempPathTraits = PathTraits
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

    StatusList.InputFromBinary FileNum, Version
    
    AbilityList.InputFromBinary FileNum, Version
    InfluenceList.InputFromBinary FileNum, Version
    
    BackgroundList.InputFromBinary FileNum, Version
    HealthList.InputFromBinary FileNum, Version
    BondList.InputFromBinary FileNum, Version
    
    MiscellaneousList.InputFromBinary FileNum, Version
    DerangementList.InputFromBinary FileNum, Version
    
    DisciplineList.InputFromBinary FileNum, Version
    RitualList.InputFromBinary FileNum, Version
    
    MeritList.InputFromBinary FileNum, Version
    FlawList.InputFromBinary FileNum, Version
    
    EquipmentList.InputFromBinary FileNum, Version
    If Version >= 2.395 Then HangoutList.InputFromBinary FileNum, Version
    
    If Version >= 2.399 Then
        Dim I As Integer
        Dim Boon As BoonClass
        Get #FileNum, , I
        Do Until I <= 0
            Set Boon = New BoonClass
            Boon.InputFromBinary FileNum, Version
            BoonList.Append Boon
            I = I - 1
        Loop
        RefreshBoonTraits
    End If
    
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
' Description:  Read a character from an input file.  The version
'               ensures backward-compatibility.
'

    Dim Read As String
    
    ' Race has already been read
    Line Input #FileNum, Name
    Line Input #FileNum, Nature
    Line Input #FileNum, Demeanor
    Line Input #FileNum, Clan
    If Version >= GameFileVersionTag4 Then
        Line Input #FileNum, Sect
    End If
    Line Input #FileNum, Read
    Generation = Val(Read)
    Line Input #FileNum, Title
    Line Input #FileNum, Read
    Blood = Val(Read)
    TempBlood = Blood
    Line Input #FileNum, Read
    Willpower = Val(Read)
    TempWillpower = Willpower
    
    If Version >= GameFileVersionTag4 Then
        Line Input #FileNum, Read
        Conscience = Val(Read)
        TempConscience = Conscience
        Line Input #FileNum, Read
        SelfControl = Val(Read)
        TempSelfControl = SelfControl
        Line Input #FileNum, Read
        Courage = Val(Read)
        TempCourage = Courage
        Line Input #FileNum, Path
        Line Input #FileNum, Read
        PathTraits = Val(Read)
        TempPathTraits = PathTraits
    End If
    
    PhysicalList.OldInputFromFile FileNum, Version
    SocialList.OldInputFromFile FileNum, Version
    MentalList.OldInputFromFile FileNum, Version
    PhysicalNegList.OldInputFromFile FileNum, Version
    SocialNegList.OldInputFromFile FileNum, Version
    MentalNegList.OldInputFromFile FileNum, Version
    
    StatusList.OldInputFromFile FileNum, Version
    
    AbilityList.OldInputFromFile FileNum, Version
    InfluenceList.OldInputFromFile FileNum, Version
    
    If Version >= GameFileVersionTag4 Then
        BackgroundList.OldInputFromFile FileNum, Version
        HealthList.OldInputFromFile FileNum, Version
        BondList.OldInputFromFile FileNum, Version
    End If
    
    MiscellaneousList.LLInputFromFile FileNum, Version
    DerangementList.LLInputFromFile FileNum, Version
    DisciplineList.LLInputFromFile FileNum, Version
    RitualList.LLInputFromFile FileNum, Version
    MeritList.LLInputFromFile FileNum, Version
    FlawList.LLInputFromFile FileNum, Version
                
    Notes = ""
    If Version <> GameFileVersionTag0 Then
        Read = ReadLongField(FileNum, "<-End Vampire Character Equipment->")
        If Read <> "" Then Notes = "2.3 Equipment:" & vbCrLf & Read & vbCrLf
    End If
    Notes = Notes & ReadLongField(FileNum, "<-End Vampire Character Notes->")
    
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

Public Sub RefreshBoonTraits()
'
' Name:         RefreshBoonTraits
' Description:  Ensure that the BoonTraitList echoes the content of the BoonList.
'

    Dim Boon As BoonClass
    
    BoonTraitList.Clear
    BoonList.First
    Do Until BoonList.Off
        Set Boon = BoonList.Item
        BoonTraitList.Append Boon.Name
        BoonList.MoveNext
    Loop

End Sub

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Intialize all data to defaults.
'

    Generation = 13
    Blood = 10
    Willpower = 2
    Conscience = 1
    SelfControl = 1
    Courage = 1
    PathTraits = 1
    TempBlood = 10
    TempWillpower = 2
    TempConscience = 1
    TempSelfControl = 1
    TempCourage = 1
    TempPathTraits = 1
    PhysicalMax = 11
    SocialMax = 11
    MentalMax = 11
    AuraBonus = "+0"
    
    Status = ActiveStatus
    StartDate = Now
    LastModified = Now

    Set PhysicalList = New LinkedTraitList
    Set SocialList = New LinkedTraitList
    Set MentalList = New LinkedTraitList
    Set PhysicalNegList = New LinkedTraitList
    Set SocialNegList = New LinkedTraitList
    Set MentalNegList = New LinkedTraitList
    
    Set StatusList = New LinkedTraitList
    
    Set AbilityList = New LinkedTraitList
    Set InfluenceList = New LinkedTraitList
    
    Set BackgroundList = New LinkedTraitList
    Set HealthList = New LinkedTraitList
    Set BondList = New LinkedTraitList
    
    Set MiscellaneousList = New LinkedTraitList
    Set DerangementList = New LinkedTraitList
    
    Set DisciplineList = New LinkedTraitList
    Set RitualList = New LinkedTraitList
    
    Set MeritList = New LinkedTraitList
    Set FlawList = New LinkedTraitList
    
    Set EquipmentList = New LinkedTraitList
    Set HangoutList = New LinkedTraitList
    
    Set BoonList = New LinkedList
    Set BoonTraitList = New LinkedTraitList
    
    Set Experience = New ExperienceClass
    
    PhysicalList.Initialize "Physical", True, False, False, ldMultiplier
    SocialList.Initialize "Social", True, False, False, ldMultiplier
    MentalList.Initialize "Mental", True, False, False, ldMultiplier
    PhysicalNegList.Initialize "Negative Physical", True, True, False, ldMultiplier
    SocialNegList.Initialize "Negative Social", True, True, False, ldMultiplier
    MentalNegList.Initialize "Negative Mental", True, True, False, ldMultiplier
    
    StatusList.Initialize "Status", True, False, False, ldMultiplier
    AbilityList.Initialize "Abilities", True, False, False, ldMultiplier
    InfluenceList.Initialize "Influences", True, False, False, ldMultiplier
    BackgroundList.Initialize "Backgrounds", True, False, False, ldMultiplier
    HealthList.Initialize "Health Levels", False, False, False, ldMultiplier
    BondList.Initialize "Bonds", True, False, False, ldMultiplier
    
    MiscellaneousList.Initialize "Miscellaneous", False, False, False, ldMultiplier
    DerangementList.Initialize "Derangements", True, True, True, ldNoteOnly

    DisciplineList.Initialize "Disciplines", False, False, True, ldNoteOnly
    RitualList.Initialize "Rituals", False, False, True, ldNoteOnly
    
    MeritList.Initialize "Merits", True, False, True, ldCost
    FlawList.Initialize "Flaws", True, True, True, ldCost
    
    EquipmentList.Initialize "Equipment", True, False, False, ldMultiplier
    HangoutList.Initialize "Locations", True, False, True, ldNoteOnly
    
    BoonTraitList.Initialize "Boons", False, False, True, ldNoteOnly
    
    With HealthList
        .Append StdHealth(0)
        .Append StdHealth(1)
        .Append StdHealth(2)
        .Append StdHealth(3)
        .Append "Torpor"
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

    Set StatusList = Nothing

    Set AbilityList = Nothing
    Set InfluenceList = Nothing

    Set BackgroundList = Nothing
    Set HealthList = Nothing
    Set BondList = Nothing
    
    Set MiscellaneousList = Nothing
    Set DerangementList = Nothing

    Set DisciplineList = Nothing
    Set RitualList = Nothing

    Set MeritList = Nothing
    Set FlawList = Nothing

    Set EquipmentList = Nothing
    Set HangoutList = Nothing
    
    Set BoonList = Nothing
    Set BoonTraitList = Nothing
    
    Set Experience = Nothing

End Sub
