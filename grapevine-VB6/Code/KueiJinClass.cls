VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KueiJinClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Nature As String
Public Demeanor As String
Public Dharma As String
Public Direction As String
Public Balance As String
Public Station As String
Public PoArchetype As String

Public Hun As Integer
Public Po As Integer
Public YinChi As Integer
Public YangChi As Integer
Public DemonChi As Integer
Public DharmaTraits As Integer
Public Willpower As Integer

Public TempHun As Integer
Public TempPo As Integer
Public TempYinChi As Integer
Public TempYangChi As Integer
Public TempDemonChi As Integer
Public TempDharmaTraits As Integer
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

Public StatusList As LinkedTraitList
Public AbilityList As LinkedTraitList
Public InfluenceList As LinkedTraitList
Public BackgroundList As LinkedTraitList

Public HealthList As LinkedTraitList
Public GuanxiList As LinkedTraitList

Public DisciplineList As LinkedTraitList
Public RiteList As LinkedTraitList

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
    StatusList.First
    AbilityList.First
    InfluenceList.First
    GuanxiList.First
    BackgroundList.First
    HealthList.First
    DisciplineList.First
    RiteList.First
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
        
        Case qkDisciplines:     Set CharData = DisciplineList
        Case qkRites:           Set CharData = RiteList
        Case qkHun:             CharData = Hun
        Case qkPo:              CharData = Po
        Case qkYin:             CharData = YinChi
        Case qkYang:            CharData = YangChi
        Case qkDemonChi:        CharData = DemonChi
        Case qkPathTraits:      CharData = DharmaTraits
        Case qkStation:         CharData = Station
        Case qkDharma:          CharData = Dharma
        Case qkKJBalance:       CharData = Balance
        Case qkDirection:       CharData = Direction
        Case qkPoArchetype:     CharData = PoArchetype
        Case qkGuanxi:          Set CharData = GuanxiList
        Case qkBiography:       CharData = Biography
        Case qkPhysicalMax:     CharData = PhysicalMax
        Case qkSocialMax:       CharData = SocialMax
        Case qkMentalMax:       CharData = MentalMax
        Case qkTempHun:         CharData = TempHun
        Case qkTempPo:          CharData = TempPo
        Case qkTempYin:         CharData = TempYinChi
        Case qkTempYang:        CharData = TempYangChi
        Case qkTempDemonChi:    CharData = TempDemonChi
        Case qkTempPathTraits:  CharData = TempDharmaTraits
        
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
        Case qkHun:             Hun = CInt(CharData)
        Case qkPo:              Po = CInt(CharData)
        Case qkYin:             YinChi = CInt(CharData)
        Case qkYang:            YangChi = CInt(CharData)
        Case qkDemonChi:        DemonChi = CInt(CharData)
        Case qkPathTraits:      DharmaTraits = CInt(CharData)
        Case qkTempHun:         TempHun = CInt(CharData)
        Case qkTempPo:          TempPo = CInt(CharData)
        Case qkTempYin:         TempYinChi = CInt(CharData)
        Case qkTempYang:        TempYangChi = CInt(CharData)
        Case qkTempDemonChi:    TempDemonChi = CInt(CharData)
        Case qkTempPathTraits:  TempDharmaTraits = CInt(CharData)
        
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
    Estimate.Append qkDisciplines, , "9"
    Estimate.Append qkRites
    Estimate.Append qkPathTraits, "2", DharmaTraits
    Estimate.Append qkWillpower, "3", "4"
    Estimate.Append qkHun, "2", Hun
    Estimate.Append qkPo, "-2", Po
    Estimate.Append qkYin, "2", YinChi
    Estimate.Append qkYang, "2", YangChi
    Estimate.Append qkDemonChi, "2", DemonChi
    
End Sub

Public Function Group() As String
'
' Name:         Group
' Description:  Return the 'Group' this character belongs to.
'               Clan, Tribe, Kith, Guild, etc.
' Returns:      the group
'

    Group = Dharma

End Function

Public Function Subgroup() As String
'
' Name:         Subgroup
' Description:  Return the 'Subgroup' this character belongs to.
'               Sect, Auspice, Seeming, etc.
' Returns:      the subgroup
'

    Subgroup = Balance

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

    Race = "Kuei-Jin"

End Function

Public Function RaceCode() As RaceType
'
' Name:         RaceCode
' Description:  Return the RaceType (enumeration value) of this
'               character type.
' Returns:      the corresponding RaceType
'
    
    RaceCode = gvRaceKueiJin

End Function

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write this character's data to a file.
'

    With XML
        
        .BeginTag "kueijin"
    
            .WriteAttribute "name", Name
            .WriteAttribute "nature", Nature
            .WriteAttribute "demeanor", Demeanor
            .WriteAttribute "dharma", Dharma
            .WriteAttribute "balance", Balance
            .WriteAttribute "direction", Direction
            .WriteAttribute "station", Station
            .WriteAttribute "poarchetype", PoArchetype
            .WriteAttribute "hun", Hun
            .WriteAttribute "temphun", TempHun, Hun
            .WriteAttribute "po", Po
            .WriteAttribute "temppo", TempPo, Po
            .WriteAttribute "yinchi", YinChi
            .WriteAttribute "tempyinchi", TempYinChi, YinChi
            .WriteAttribute "yangchi", YangChi
            .WriteAttribute "tempyangchi", TempYangChi, YangChi
            .WriteAttribute "demonchi", DemonChi
            .WriteAttribute "tempdemonchi", TempDemonChi, DemonChi
            .WriteAttribute "dharmatraits", DharmaTraits
            .WriteAttribute "tempdharmatraits", TempDharmaTraits, DharmaTraits
            .WriteAttribute "willpower", Willpower
            .WriteAttribute "tempwillpower", TempWillpower, Willpower
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
            GuanxiList.OutputToFile XML
            
            DisciplineList.OutputToFile XML
            RiteList.OutputToFile XML
            
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
    
        If .Tag = "kueijin" Then
        
            Name = .GetString("name")
            Nature = .GetString("nature")
            Demeanor = .GetString("demeanor")
            Dharma = .GetString("dharma")
            Balance = .GetString("balance")
            Direction = .GetString("direction")
            Station = .GetString("station")
            Hun = .GetInteger("hun")
            TempHun = .GetInteger("temphun", Hun)
            Po = .GetInteger("po")
            TempPo = .GetInteger("temppo", Po)
            PoArchetype = .GetString("poarchetype")
            YinChi = .GetInteger("yinchi")
            TempYinChi = .GetInteger("tempyinchi", YinChi)
            YangChi = .GetInteger("yangchi")
            TempYangChi = .GetInteger("tempyangchi", YangChi)
            DemonChi = .GetInteger("demonchi")
            TempDemonChi = .GetInteger("tempdemonchi", DemonChi)
            DharmaTraits = .GetInteger("dharmatraits")
            TempDharmaTraits = .GetInteger("tempdharmatraits", DharmaTraits)
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
            
            Do Until .Tag = "kueijin" And .Terminal
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
                            Case "Guanxi":              GuanxiList.InputFromFile XML, Version
                            Case "Disciplines":         DisciplineList.InputFromFile XML, Version
                            Case "Rites":               RiteList.InputFromFile XML, Version
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
            Err.Raise vbObjectError + 8, "KueiJinClass", _
                      "Wrong file input to a Kuei-Jin character"
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
    PutStrB FileNum, Dharma
    PutStrB FileNum, Balance
    PutStrB FileNum, Direction
    PutStrB FileNum, Station
    PutStrB FileNum, PoArchetype
    Put #FileNum, , Hun
    Put #FileNum, , TempHun
    Put #FileNum, , Po
    Put #FileNum, , TempPo
    Put #FileNum, , YinChi
    Put #FileNum, , TempYinChi
    Put #FileNum, , YangChi
    Put #FileNum, , TempYangChi
    Put #FileNum, , DemonChi
    Put #FileNum, , TempDemonChi
    Put #FileNum, , DharmaTraits
    Put #FileNum, , TempDharmaTraits
    Put #FileNum, , Willpower
    Put #FileNum, , TempWillpower
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
    GuanxiList.OutputToBinary FileNum
    
    DisciplineList.OutputToBinary FileNum
    RiteList.OutputToBinary FileNum
    
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
    GetStrB FileNum, Dharma
    GetStrB FileNum, Balance
    GetStrB FileNum, Direction
    GetStrB FileNum, Station
    GetStrB FileNum, PoArchetype
    
    If Version >= 2.397 Then
        Get #FileNum, , Hun
        Get #FileNum, , TempHun
        Get #FileNum, , Po
        Get #FileNum, , TempPo
        Get #FileNum, , YinChi
        Get #FileNum, , TempYinChi
        Get #FileNum, , YangChi
        Get #FileNum, , TempYangChi
        Get #FileNum, , DemonChi
        Get #FileNum, , TempDemonChi
        Get #FileNum, , DharmaTraits
        Get #FileNum, , TempDharmaTraits
        Get #FileNum, , Willpower
        Get #FileNum, , TempWillpower
        Get #FileNum, , PhysicalMax
        Get #FileNum, , SocialMax
        Get #FileNum, , MentalMax
    Else
        Get #FileNum, , Hun
        Get #FileNum, , Po
        Get #FileNum, , YinChi
        Get #FileNum, , YangChi
        Get #FileNum, , DemonChi
        Get #FileNum, , DharmaTraits
        Get #FileNum, , Willpower
        TempHun = Hun
        TempPo = Po
        TempYinChi = YinChi
        TempYangChi = YangChi
        TempDemonChi = DemonChi
        TempDharmaTraits = DharmaTraits
        TempWillpower = Willpower
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
    GuanxiList.InputFromBinary FileNum, Version
    
    DisciplineList.InputFromBinary FileNum, Version
    RiteList.InputFromBinary FileNum, Version
    
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
    Line Input #FileNum, Dharma
    Line Input #FileNum, Balance
    Line Input #FileNum, Direction
    Line Input #FileNum, Station
    Line Input #FileNum, PoArchetype
    
    Line Input #FileNum, Read
    Hun = Val(Read)
    TempHun = Hun
    Line Input #FileNum, Read
    Po = Val(Read)
    TempPo = Po
    Line Input #FileNum, Read
    YinChi = Val(Read)
    TempYinChi = YinChi
    Line Input #FileNum, Read
    YangChi = Val(Read)
    TempYangChi = YangChi
    Line Input #FileNum, Read
    DemonChi = Val(Read)
    TempDemonChi = DemonChi
    Line Input #FileNum, Read
    DharmaTraits = Val(Read)
    TempDharmaTraits = DharmaTraits
    Line Input #FileNum, Read
    Willpower = Val(Read)
    TempWillpower = Willpower
        
    PhysicalList.OldInputFromFile FileNum, Version
    SocialList.OldInputFromFile FileNum, Version
    MentalList.OldInputFromFile FileNum, Version
    PhysicalNegList.OldInputFromFile FileNum, Version
    SocialNegList.OldInputFromFile FileNum, Version
    MentalNegList.OldInputFromFile FileNum, Version
    
    StatusList.OldInputFromFile FileNum, Version
    AbilityList.OldInputFromFile FileNum, Version
    InfluenceList.OldInputFromFile FileNum, Version
    BackgroundList.OldInputFromFile FileNum, Version
        
    HealthList.OldInputFromFile FileNum, Version
    GuanxiList.OldInputFromFile FileNum, Version
    
    DisciplineList.LLInputFromFile FileNum, Version
    RiteList.LLInputFromFile FileNum, Version
    MeritList.LLInputFromFile FileNum, Version
    FlawList.LLInputFromFile FileNum, Version
    
    Notes = ""
    Read = ReadLongField(FileNum, "<-End Kuei-Jin Character Equipment->")
    If Read <> "" Then Notes = "2.3 Equipment:" & vbCrLf & Read & vbCrLf
     
    Notes = Notes & ReadLongField(FileNum, "<-End Kuei-Jin Character Notes->")
    
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

    Hun = 1
    Po = 3
    YinChi = 1
    YangChi = 1
    DemonChi = 1
    DharmaTraits = 1
    Willpower = 1
    TempHun = 1
    TempPo = 3
    TempYinChi = 1
    TempYangChi = 1
    TempDemonChi = 1
    TempDharmaTraits = 1
    TempWillpower = 1
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
    
    Set StatusList = New LinkedTraitList
    Set AbilityList = New LinkedTraitList
    Set InfluenceList = New LinkedTraitList
    Set BackgroundList = New LinkedTraitList
    
    Set HealthList = New LinkedTraitList
    Set GuanxiList = New LinkedTraitList
    
    Set DisciplineList = New LinkedTraitList
    Set RiteList = New LinkedTraitList
    
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
    StatusList.Initialize "Status", True, False, False, ldMultiplier
    
    GuanxiList.Initialize "Guanxi", True, False, False, ldMultiplier
    DisciplineList.Initialize "Disciplines", False, False, True, ldNoteOnly
    RiteList.Initialize "Rites", False, False, True, ldNoteOnly
        
    MeritList.Initialize "Merits", True, False, True, ldCost
    FlawList.Initialize "Flaws", True, True, True, ldCost
    
    EquipmentList.Initialize "Equipment", True, False, False, ldMultiplier
    HangoutList.Initialize "Locations", True, False, True, ldNoteOnly
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
    Set GuanxiList = Nothing
    
    Set DisciplineList = Nothing
    Set RiteList = Nothing

    Set MeritList = Nothing
    Set FlawList = Nothing

    Set EquipmentList = Nothing
    Set HangoutList = Nothing
    Set Experience = Nothing

End Sub
