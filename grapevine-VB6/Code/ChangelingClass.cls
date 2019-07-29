VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChangelingClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public SeelieLegacy As String
Public UnseelieLegacy As String
Public Kith As String
Public Seeming As String
Public Court As String
Public House As String
Public Title As String
Public Threshold As String

Public Glamour As Integer
Public TempGlamour As Integer
Public Banality As Integer
Public TempBanality As Integer
Public Willpower As Integer
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
Public InfluenceList As LinkedTraitList

Public BackgroundList As LinkedTraitList
Public HealthList As LinkedTraitList
Public StatusList As LinkedTraitList

Public ArtList As LinkedTraitList
Public RealmList As LinkedTraitList

Public MeritList As LinkedTraitList
Public FlawList As LinkedTraitList

Public Oaths As String
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
    StatusList.First
    ArtList.First
    RealmList.First
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
        
        Case qkArts:            Set CharData = ArtList
        Case qkRealms:          Set CharData = RealmList
        Case qkKith:            CharData = Kith
        Case qkSeeming:         CharData = Seeming
        Case qkSeelie:          CharData = SeelieLegacy
        Case qkUnseelie:        CharData = UnseelieLegacy
        Case qkCourt:           CharData = Court
        Case qkHouse:           CharData = House
        Case qkTitle:           CharData = Title
        Case qkThreshold:       CharData = Threshold
        Case qkGlamour:         CharData = Glamour
        Case qkTempGlamour:     CharData = TempGlamour
        Case qkBanality:        CharData = Banality
        Case qkTempBanality:    CharData = TempBanality
        Case qkOaths:           CharData = Oaths
        Case qkBiography:       CharData = Biography
        Case qkPhysicalMax:     CharData = PhysicalMax
        Case qkSocialMax:       CharData = SocialMax
        Case qkMentalMax:       CharData = MentalMax
        
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
        Case qkGlamour:         Glamour = CInt(CharData)
        Case qkTempGlamour:     TempGlamour = CInt(CharData)
        Case qkBanality:        Banality = CInt(CharData)
        Case qkTempBanality:    TempBanality = CInt(CharData)
        
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
    Estimate.Append qkArts, , "9"
    Estimate.Append qkRealms, , "10"
    Estimate.Append qkGlamour, "3", "4"
    Estimate.Append qkWillpower, "3", "3"
    Estimate.Append qkBanality, "-2", Banality

End Sub

Public Function Group() As String
'
' Name:         Group
' Description:  Return the 'Group' this character belongs to.
'               Clan, Tribe, Kith, Guild, etc.
' Returns:      the group
'

    Group = Kith

End Function

Public Function Subgroup() As String
'
' Name:         Subgroup
' Description:  Return the 'Subgroup' this character belongs to.
'               Sect, Auspice, Seeming, etc.
' Returns:      the subgroup
'

    Subgroup = Seeming

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

    Race = "Changeling"

End Function

Public Function RaceCode() As RaceType
'
' Name:         RaceCode
' Description:  Return the RaceType (enumeration value) of this
'               character type.
' Returns:      the corresponding RaceType
'
    
    RaceCode = gvRaceChangeling

End Function

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write this character's data to a file.
'

    With XML
    
        .BeginTag "changeling"
        
            .WriteAttribute "name", Name
            .WriteAttribute "seelie", SeelieLegacy
            .WriteAttribute "unseelie", UnseelieLegacy
            .WriteAttribute "court", Court
            .WriteAttribute "kith", Kith
            .WriteAttribute "seeming", Seeming
            .WriteAttribute "house", House
            .WriteAttribute "threshold", Threshold
            .WriteAttribute "title", Title
            .WriteAttribute "glamour", Glamour
            .WriteAttribute "tempglamour", TempGlamour, Glamour
            .WriteAttribute "banality", Banality
            .WriteAttribute "tempbanality", TempBanality, Banality
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
                
            AbilityList.OutputToFile XML
            InfluenceList.OutputToFile XML
            
            BackgroundList.OutputToFile XML
            HealthList.OutputToFile XML
            StatusList.OutputToFile XML
            
            ArtList.OutputToFile XML
            RealmList.OutputToFile XML
            
            MeritList.OutputToFile XML
            FlawList.OutputToFile XML
            EquipmentList.OutputToFile XML
            HangoutList.OutputToFile XML
    
            .WriteCDataTag "oaths", Oaths
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
    
        If .Tag = "changeling" Then
        
            Name = .GetString("name")
            SeelieLegacy = .GetString("seelie")
            UnseelieLegacy = .GetString("unseelie")
            Court = .GetString("court")
            Kith = .GetString("kith")
            Seeming = .GetString("seeming")
            House = .GetString("house")
            Threshold = .GetString("threshold")
            Title = .GetString("title")
            Glamour = .GetInteger("glamour")
            TempGlamour = .GetInteger("tempglamour", Glamour)
            Banality = .GetInteger("banality")
            TempBanality = .GetInteger("tempbanality", Banality)
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
            
            Do Until .Tag = "changeling" And .Terminal
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
                            Case "Arts":                ArtList.InputFromFile XML, Version
                            Case "Realms":              RealmList.InputFromFile XML, Version
                            Case "Merits":              MeritList.InputFromFile XML, Version
                            Case "Flaws":               FlawList.InputFromFile XML, Version
                            Case "Equipment":           EquipmentList.InputFromFile XML, Version
                            Case "Locations":           HangoutList.InputFromFile XML, Version
                        End Select
                    Case "notes"
                        If Not .Terminal Then
                            .ReadTag
                            Notes = .GetTagText
                        End If
                    Case "biography"
                        If Not .Terminal Then
                            .ReadTag
                            Biography = .GetTagText
                        End If
                    Case "oaths"
                        If Not .Terminal Then
                            .ReadTag
                            Oaths = .GetTagText
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
            Err.Raise vbObjectError + 8, "ChangelingClass", _
                      "Wrong file input to a changeling character"
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
    PutStrB FileNum, SeelieLegacy
    PutStrB FileNum, UnseelieLegacy
    PutStrB FileNum, Court
    PutStrB FileNum, Kith
    PutStrB FileNum, Seeming
    PutStrB FileNum, House
    PutStrB FileNum, Threshold
    PutStrB FileNum, Title
    Put #FileNum, , Glamour
    Put #FileNum, , TempGlamour
    Put #FileNum, , Banality
    Put #FileNum, , TempBanality
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
        
    AbilityList.OutputToBinary FileNum
    InfluenceList.OutputToBinary FileNum
    
    BackgroundList.OutputToBinary FileNum
    HealthList.OutputToBinary FileNum
    StatusList.OutputToBinary FileNum
    
    ArtList.OutputToBinary FileNum
    RealmList.OutputToBinary FileNum
    
    MeritList.OutputToBinary FileNum
    FlawList.OutputToBinary FileNum
    EquipmentList.OutputToBinary FileNum
    HangoutList.OutputToBinary FileNum
    
    PutStrB FileNum, Oaths
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
    GetStrB FileNum, SeelieLegacy
    GetStrB FileNum, UnseelieLegacy
    GetStrB FileNum, Court
    GetStrB FileNum, Kith
    GetStrB FileNum, Seeming
    GetStrB FileNum, House
    GetStrB FileNum, Threshold
    GetStrB FileNum, Title
    If Version >= 2.397 Then
        Get #FileNum, , Glamour
        Get #FileNum, , TempGlamour
        Get #FileNum, , Banality
        Get #FileNum, , TempBanality
        Get #FileNum, , Willpower
        Get #FileNum, , TempWillpower
        Get #FileNum, , PhysicalMax
        Get #FileNum, , SocialMax
        Get #FileNum, , MentalMax
    Else
        Get #FileNum, , Glamour
        Get #FileNum, , Banality
        Get #FileNum, , Willpower
        TempGlamour = Glamour
        TempBanality = Banality
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
        
    AbilityList.InputFromBinary FileNum, Version
    InfluenceList.InputFromBinary FileNum, Version
    
    BackgroundList.InputFromBinary FileNum, Version
    HealthList.InputFromBinary FileNum, Version
    StatusList.InputFromBinary FileNum, Version
    
    ArtList.InputFromBinary FileNum, Version
    RealmList.InputFromBinary FileNum, Version
    
    MeritList.InputFromBinary FileNum, Version
    FlawList.InputFromBinary FileNum, Version
    EquipmentList.InputFromBinary FileNum, Version
    If Version >= 2.395 Then HangoutList.InputFromBinary FileNum, Version
    
    GetStrB FileNum, Oaths
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
' Description:  Read a character from an old-format input file.  The version
'               ensures backward-compatibility.
'

    Dim Read As String
    
    ' Race has already been read
    Line Input #FileNum, Name
    Line Input #FileNum, SeelieLegacy
    Line Input #FileNum, UnseelieLegacy
    Line Input #FileNum, Court
    Line Input #FileNum, Kith
    Line Input #FileNum, Seeming
    Line Input #FileNum, House
    Line Input #FileNum, Threshold
    Line Input #FileNum, Title
    Line Input #FileNum, Read
    Glamour = Val(Read)
    TempGlamour = Glamour
    Line Input #FileNum, Read
    Banality = Val(Read)
    TempBanality = Banality
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
    StatusList.OldInputFromFile FileNum, Version
    
    ArtList.LLInputFromFile FileNum, Version
    RealmList.LLInputFromFile FileNum, Version
    MeritList.LLInputFromFile FileNum, Version
    FlawList.LLInputFromFile FileNum, Version
    
    Oaths = ReadLongField(FileNum, "<-End Changeling Character Oaths->")
    
    Notes = ""
    Read = ReadLongField(FileNum, "<-End Changeling Character Equipment->")
    If Read <> "" Then Notes = "2.3 Equipment:" & vbCrLf & Read & vbCrLf
    
    Notes = Notes & ReadLongField(FileNum, "<-End Changeling Character Notes->")

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

    Glamour = 4
    TempGlamour = 4
    Banality = 3
    TempBanality = 3
    Willpower = 2
    TempWillpower = 2
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
    Set StatusList = New LinkedTraitList
    
    Set ArtList = New LinkedTraitList
    Set RealmList = New LinkedTraitList
        
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
    
    ArtList.Initialize "Arts", False, False, True, ldNoteOnly
    RealmList.Initialize "Realms", False, False, True, ldNoteOnly
    
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
    Set StatusList = Nothing

    Set ArtList = Nothing
    Set RealmList = Nothing

    Set MeritList = Nothing
    Set FlawList = Nothing

    Set EquipmentList = Nothing
    Set HangoutList = Nothing
    Set Experience = Nothing

End Sub
