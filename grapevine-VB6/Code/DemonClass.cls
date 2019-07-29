VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DemonClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Nature As String
Public Demeanor As String
Public House As String
Public Faction As String

Public Willpower As Integer
Public Torment As Integer
Public Faith As Integer
Public Conscience As Integer
Public Conviction As Integer
Public Courage As Integer

Public TempWillpower As Integer
Public TempTorment As Integer
Public TempFaith As Integer
Public TempConscience As Integer
Public TempConviction As Integer
Public TempCourage As Integer

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

Public LoreList As LinkedTraitList
Public VisageList As LinkedTraitList

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
    LoreList.First
    VisageList.First
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
        
        Case qkLores:           Set CharData = LoreList
        Case qkVisage:          Set CharData = VisageList
        Case qkConscience:      CharData = Conscience
        Case qkConviction:      CharData = Conviction
        Case qkCourage:         CharData = Courage
        Case qkHouse:           CharData = House
        Case qkFaction:         CharData = Faction
        Case qkTorment:         CharData = Torment
        Case qkFaith:           CharData = Faith
        Case qkBiography:       CharData = Biography
        Case qkPhysicalMax:     CharData = PhysicalMax
        Case qkSocialMax:       CharData = SocialMax
        Case qkMentalMax:       CharData = MentalMax
        Case qkTempConscience:  CharData = TempConscience
        Case qkTempConviction:  CharData = TempConviction
        Case qkTempCourage:     CharData = TempCourage
        Case qkTempTorment:     CharData = TempTorment
        Case qkTempFaith:       CharData = TempFaith
        
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
        Case qkConscience:      Conscience = CInt(CharData)
        Case qkConviction:      Conviction = CInt(CharData)
        Case qkCourage:         Courage = CInt(CharData)
        Case qkTorment:         Torment = CInt(CharData)
        Case qkFaith:           Faith = CInt(CharData)
        Case qkTempConscience:  TempConscience = CInt(CharData)
        Case qkTempConviction:  TempConviction = CInt(CharData)
        Case qkTempCourage:     TempCourage = CInt(CharData)
        Case qkTempTorment:     TempTorment = CInt(CharData)
        Case qkTempFaith:       TempFaith = CInt(CharData)
        
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
    Estimate.Append qkLores, , "9"
    Estimate.Append qkVisage, , VisageList.TraitCount
    Estimate.Append qkTorment, "-2", Torment
    Estimate.Append qkWillpower, "3", "4"
    Estimate.Append qkFaith, "3", "3"
    Estimate.Append qkConscience, "2", Conscience
    Estimate.Append qkConviction, "2", Conviction
    Estimate.Append qkCourage, "2", Courage

End Sub

Public Function Race() As String
'
' Name:         Race
' Description:  Return a string representation of this character's
'               character type.
' Returns:      the name of the character type
'

    Race = "Demon"

End Function

Public Function RaceCode() As RaceType
'
' Name:         RaceCode
' Description:  Return the RaceType (enumeration value) of this
'               character type.
' Returns:      the corresponding RaceType
'
    
    RaceCode = gvRaceDemon

End Function

Public Function Group() As String
'
' Name:         Group
' Description:  Return the 'Group' this character belongs to.
'               Clan, Tribe, Kith, Guild, etc.
' Returns:      the group
'

    Group = House

End Function

Public Function Subgroup() As String
'
' Name:         Subgroup
' Description:  Return the 'Subgroup' this character belongs to.
'               Sect, Auspice, Seeming, etc.
' Returns:      the subgroup
'

    Subgroup = Faction

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
        .BeginTag "demon"
        
            .WriteAttribute "name", Name
            .WriteAttribute "house", House
            .WriteAttribute "faction", Faction
            .WriteAttribute "nature", Nature
            .WriteAttribute "demeanor", Demeanor
            .WriteAttribute "torment", Torment
            .WriteAttribute "temptorment", TempTorment, Torment
            .WriteAttribute "faith", Faith
            .WriteAttribute "tempfaith", TempFaith, Faith
            .WriteAttribute "willpower", Willpower
            .WriteAttribute "tempwillpower", TempWillpower, Willpower
            .WriteAttribute "conscience", Conscience
            .WriteAttribute "tempconscience", TempConscience, Conscience
            .WriteAttribute "conviction", Conviction
            .WriteAttribute "tempconviction", TempConviction, Conviction
            .WriteAttribute "courage", Courage
            .WriteAttribute "tempcourage", TempCourage, Courage
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
            
            LoreList.OutputToFile XML
            VisageList.OutputToFile XML
            
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
    
        If .Tag = "demon" Then
        
            Name = .GetString("name")
            House = .GetString("house")
            Faction = .GetString("faction")
            Nature = .GetString("nature")
            Demeanor = .GetString("demeanor")
            Torment = .GetInteger("torment")
            TempTorment = .GetInteger("temptorment", Torment)
            Faith = .GetInteger("faith")
            TempFaith = .GetInteger("tempfaith", Faith)
            Willpower = .GetInteger("willpower")
            TempWillpower = .GetInteger("tempwillpower", Willpower)
            Conscience = .GetInteger("conscience")
            TempConscience = .GetInteger("tempconscience", Conscience)
            Conviction = .GetInteger("conviction")
            TempConviction = .GetInteger("tempconviction", Conviction)
            Courage = .GetInteger("courage")
            TempCourage = .GetInteger("tempcourage", Courage)
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
            
            Do Until .Tag = "demon" And .Terminal
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
                            Case "Lores":               LoreList.InputFromFile XML, Version
                            Case "Apocalyptic Form":    VisageList.InputFromFile XML, Version
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

        Else
            Err.Raise vbObjectError + 8, "DemonClass", _
                      "Wrong file input to a demon character"
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
    PutStrB FileNum, House
    PutStrB FileNum, Faction
    PutStrB FileNum, Nature
    PutStrB FileNum, Demeanor
    Put #FileNum, , Torment
    Put #FileNum, , TempTorment
    Put #FileNum, , Faith
    Put #FileNum, , TempFaith
    Put #FileNum, , Willpower
    Put #FileNum, , TempWillpower
    Put #FileNum, , Conscience
    Put #FileNum, , TempConscience
    Put #FileNum, , Conviction
    Put #FileNum, , TempConviction
    Put #FileNum, , Courage
    Put #FileNum, , TempCourage
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
    
    LoreList.OutputToBinary FileNum
    VisageList.OutputToBinary FileNum
    
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
    GetStrB FileNum, House
    GetStrB FileNum, Faction
    GetStrB FileNum, Nature
    GetStrB FileNum, Demeanor
    Get #FileNum, , Torment
    Get #FileNum, , TempTorment
    Get #FileNum, , Faith
    Get #FileNum, , TempFaith
    Get #FileNum, , Willpower
    Get #FileNum, , TempWillpower
    Get #FileNum, , Conscience
    Get #FileNum, , TempConscience
    Get #FileNum, , Conviction
    Get #FileNum, , TempConviction
    Get #FileNum, , Courage
    Get #FileNum, , TempCourage
    Get #FileNum, , PhysicalMax
    Get #FileNum, , SocialMax
    Get #FileNum, , MentalMax
    GetStrB FileNum, Player
    GetStrB FileNum, Status
    GetStrB FileNum, ID
    Get #FileNum, , StartDate
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
    
    LoreList.InputFromBinary FileNum, Version
    VisageList.InputFromBinary FileNum, Version
    
    MeritList.InputFromBinary FileNum, Version
    FlawList.InputFromBinary FileNum, Version
    
    EquipmentList.InputFromBinary FileNum, Version
    HangoutList.InputFromBinary FileNum, Version
    
    If Version >= 2.397 Then GetStrB FileNum, Biography
    GetStrB FileNum, Notes
    
End Sub

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Intialize all data to defaults.
'
    
    Willpower = 4
    Faith = 3
    Torment = 3
    Conscience = 1
    Conviction = 1
    Courage = 1
    TempWillpower = 4
    TempFaith = 3
    TempTorment = 3
    TempConscience = 1
    TempConviction = 1
    TempCourage = 1
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

    Set LoreList = New LinkedTraitList
    Set VisageList = New LinkedTraitList

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
    
    LoreList.Initialize "Lores", False, False, True, ldNoteOnly
    VisageList.Initialize "Apocalyptic Form", False, False, True, ldNoteOnly
    
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

    Set LoreList = Nothing
    Set VisageList = Nothing

    Set MeritList = Nothing
    Set FlawList = Nothing

    Set EquipmentList = Nothing
    Set HangoutList = Nothing
    Set Experience = Nothing

End Sub
