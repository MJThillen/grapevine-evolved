VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BeteClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Nature As String
Public Demeanor As String
Public Bete As String
Public Breed As String
Public Auspice As String
Public Rank As String
Public Pack As String
Public Totem As String
Public Position As String
Public Notoriety As Integer

Public Rage As Integer
Public Gnosis As Integer
Public Willpower As Integer
Public Honor As Single
Public Glory As Single
Public Wisdom As Single

Public PhysicalList As LinkedTraitList
Public SocialList As LinkedTraitList
Public MentalList As LinkedTraitList
Public PhysicalNegList As LinkedTraitList
Public SocialNegList As LinkedTraitList
Public MentalNegList As LinkedTraitList

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
Public Notes As String

Public Player As String
Public Status As String
Public ID As String
Public StartDate As String
Public Narrator As String
Public IsNPC As Boolean
Public Experience As ExperienceClass
Public LastModified As Date

Public Sub InitializeLists()
'
' Name:         InitializeLists
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

End Sub

Public Function GetTemplateSheet()
'
' Name:         GetTemplateSheet
' Description:  Return the filename of the character sheet
'               template associated with this character.
' Returns:      template filename
'

    GetTemplateSheet = BeteSheetName

End Function

Public Sub KeyReplace(Keyword As String, Value As String, More As Boolean)
''
'' Name:         KeyReplace
'' Parameters:   Keyword     the key to replace
''               Value       the data from this character to
''                           replace it with
''               More        TRUE if there is more data of
''                           this type left; FALSE otherwise
'' Description:  Given a keyword from a character sheet template,
''               return the value to replace it with and whether
''               or not there are more such values.
''
'
'    Select Case LCase(Keyword)
'        Case "name"
'            Value = Name
'        Case "race"
'            Value = Race
'        Case "nature"
'            Value = Nature
'        Case "demeanor"
'            Value = Demeanor
'        Case "bete"
'            Value = Bete
'        Case "breed"
'            Value = Breed
'        Case "auspice"
'            Value = Auspice
'        Case "rank"
'            Value = Rank
'        Case "pack"
'            Value = Pack
'        Case "position"
'            Value = Position
'        Case "notoriety"
'            Value = Notoriety
'        Case "rage"
'            Value = CStr(Rage) & " " & String(Rage, "o")
'        Case "gnosis"
'            Value = CStr(Gnosis) & " " & String(Gnosis, "o")
'        Case "willpower"
'            Value = CStr(Willpower) & " " & String(Willpower, "o")
'        Case "#physical"
'            Value = CStr(PhysicalList.Count)
'        Case "physical"
'            If Not PhysicalList.Off Then
'                Value = PhysicalList.DisplayTrait
'                PhysicalList.MoveNext
'                More = Not PhysicalList.Off
'            End If
'        Case "#social"
'            Value = CStr(SocialList.Count)
'        Case "social"
'            If Not SocialList.Off Then
'                Value = SocialList.DisplayTrait
'                SocialList.MoveNext
'                More = Not SocialList.Off
'            End If
'        Case "#mental"
'            Value = CStr(MentalList.Count)
'        Case "mental"
'            If Not MentalList.Off Then
'                Value = MentalList.DisplayTrait
'                MentalList.MoveNext
'                More = Not MentalList.Off
'            End If
'        Case "#physicalneg"
'            Value = CStr(PhysicalNegList.Count)
'        Case "physicalneg"
'            If Not PhysicalNegList.Off Then
'                Value = PhysicalNegList.DisplayTrait
'                PhysicalNegList.MoveNext
'                More = Not PhysicalNegList.Off
'            End If
'        Case "#socialneg"
'            Value = CStr(SocialNegList.Count)
'        Case "socialneg"
'            If Not SocialNegList.Off Then
'                Value = SocialNegList.DisplayTrait
'                SocialNegList.MoveNext
'                More = Not SocialNegList.Off
'            End If
'        Case "#mentalneg"
'            Value = CStr(MentalNegList.Count)
'        Case "mentalneg"
'            If Not MentalNegList.Off Then
'                Value = MentalNegList.DisplayTrait
'                MentalNegList.MoveNext
'                More = Not MentalNegList.Off
'            End If
'        Case "#abilities"
'            Value = CStr(AbilityList.Count)
'        Case "abilities"
'            If Not AbilityList.Off Then
'                Value = AbilityList.DisplayTrait
'                AbilityList.MoveNext
'                More = Not AbilityList.Off
'            End If
'        Case "#influences"
'            Value = CStr(InfluenceList.Count)
'        Case "influences"
'            If Not InfluenceList.Off Then
'                Value = InfluenceList.DisplayTrait
'                InfluenceList.MoveNext
'                More = Not InfluenceList.Off
'            End If
'        Case "#backgrounds"
'            Value = CStr(BackgroundList.Count)
'        Case "backgrounds"
'            If Not BackgroundList.Off Then
'                Value = BackgroundList.DisplayTrait
'                BackgroundList.MoveNext
'                More = Not BackgroundList.Off
'            End If
'        Case "#healthlevels"
'            Value = CStr(HealthList.Count)
'        Case "healthlevels"
'            If Not HealthList.Off Then
'                Value = HealthList.DisplayTrait
'                HealthList.MoveNext
'                More = Not HealthList.Off
'            End If
'        Case "#features"
'            Value = CStr(FeatureList.Count)
'        Case "features"
'            If Not FeatureList.Off Then
'                Value = FeatureList.DisplayTrait
'                FeatureList.MoveNext
'                More = Not FeatureList.Off
'            End If
'        Case "#gifts"
'            Value = CStr(GiftList.Count)
'        Case "gifts"
'            If Not GiftList.Off Then
'                Value = GiftList.DisplayTrait
'                GiftList.MoveNext
'                More = Not GiftList.Off
'            End If
'        Case "#rites"
'            Value = CStr(RiteList.Count)
'        Case "rites"
'            If Not RiteList.Off Then
'                Value = RiteList.DisplayTrait
'                RiteList.MoveNext
'                More = Not RiteList.Off
'            End If
'        Case "#honor"
'            Value = CStr(HonorList.Count)
'        Case "honor"
'            If Not HonorList.Off Then
'                Value = HonorList.DisplayTrait
'                HonorList.MoveNext
'                More = Not HonorList.Off
'            End If
'        Case "#glory"
'            Value = CStr(GloryList.Count)
'        Case "glory"
'            If Not GloryList.Off Then
'                Value = GloryList.DisplayTrait
'                GloryList.MoveNext
'                More = Not GloryList.Off
'            End If
'        Case "#wisdom"
'            Value = CStr(WisdomList.Count)
'        Case "wisdom"
'            If Not WisdomList.Off Then
'                Value = WisdomList.DisplayTrait
'                WisdomList.MoveNext
'                More = Not WisdomList.Off
'            End If
'        Case "#merits"
'            Value = CStr(MeritList.Count)
'        Case "merits"
'            If Not MeritList.Off Then
'                Value = MeritList.DisplayTrait
'                MeritList.MoveNext
'                More = Not MeritList.Off
'            End If
'        Case "#flaws"
'            Value = CStr(FlawList.Count)
'        Case "flaws"
'            If Not FlawList.Off Then
'                Value = FlawList.DisplayTrait
'                FlawList.MoveNext
'                More = Not FlawList.Off
'            End If
'        Case "equipment"
'            'Value = Equipment
'        Case "notes"
'            Value = Notes
'        Case "player"
'            Value = Player
'        Case "experience"
'            Value = CStr(Experience.Unspent)
'        Case "status"
'            Value = Status
'        Case "id"
'            Value = ID
'        Case "startdate"
'            Value = StartDate
'        Case "narrator"
'            Value = Narrator
'        Case "npc"
'            Value = IIf(IsNPC, "NPC", "")
'        Case "history"
'            'Value = History
'        Case "printdate"
'            Value = Format(Date, "mmmm d, yyyy")
'        Case "?comment", "?history", "?notes"
'            Value = ""
'        Case Else
'            Value = "[" & Keyword & "]"
'    End Select
'
'    Value = Replace(Value, "\", "\\")

End Sub

Public Sub OutputTextSheet(IncludeNotes As Boolean, IncludeHistory As Boolean, _
        OutAid As OutputAidClass)
''
'' Name:         OutputTextSheet
'' Parameters:   IncludeNotes        whether to include the notes field
''               IncludeHistory      whether to include the history field
''               OutAid              the output destination
'' Description:  Send a text version of this character to the given
''               output object.  Notes and History are optional.
''
'
'    Dim Trait1 As String
'    Dim Trait2 As String
'    Dim PageWidth As Integer
'
'    PageWidth = OutAid.PageWidth
'
'    OutAid.Output String(PageWidth, "-")
'    OutAid.Output Name, "Character Sheet " & Format(Now, "mm/dd/yy")
'    OutAid.Output String(PageWidth, "-")
'    OutAid.Output Race & IIf(IsNPC, " NPC", ""), "Character ID: " & ID
'    OutAid.Output "Player: " & Player, "Experience: " & CStr(Experience.Unspent)
'    OutAid.Output "Status: " & Status, "Narrator: " & Narrator
'    OutAid.Output "", "Start Date: " & StartDate
'
'    OutAid.Output String(PageWidth, "-")
'
'    OutAid.Output "Nature: " & Nature, "Demeanor: " & Demeanor
'    OutAid.Output "Bete: " & Bete, "Breed: " & Breed
'    OutAid.Output "Auspice: " & Auspice, "Rank: " & Rank
'    OutAid.Output "Pack: " & Pack, "Position: " & Position
'    OutAid.Output "Notoriety: " & CStr(Notoriety)
'
'    OutAid.Output String(PageWidth, "-")
'
'    OutAid.Output ""
'    OutAid.Output "Rage/Blood: " & CStr(Rage) & " " & String(Rage, "o"), _
'            "Gnosis: " & CStr(Gnosis) & " " & String(Gnosis, "o")
'    OutAid.Output "Willpower: " & CStr(Willpower) & " " & String(Willpower, "o")
'
'    OutAid.Output ""
'    OutAid.Output CStr(PhysicalList.Count) & " Physical Traits:", _
'           CStr(PhysicalNegList.Count) & " Negative Physical Traits:"
'    OutAid.Output "-------------------", "---------------------------"
'    PhysicalList.First
'    PhysicalNegList.First
'    Do Until PhysicalList.Off And PhysicalNegList.Off
'        Trait1 = ""
'        Trait2 = ""
'        If Not PhysicalList.Off Then
'            Trait1 = PhysicalList.DisplayTrait
'            PhysicalList.MoveNext
'        End If
'        If Not PhysicalNegList.Off Then
'            Trait2 = PhysicalNegList.DisplayTrait
'            PhysicalNegList.MoveNext
'        End If
'        OutAid.Output Trait1, Trait2
'    Loop
'
'    OutAid.Output ""
'    OutAid.Output CStr(SocialList.Count) & " Social Traits:", _
'           CStr(SocialNegList.Count) & " Negative Social Traits:"
'    OutAid.Output "-------------------", "---------------------------"
'    SocialList.First
'    SocialNegList.First
'    Do Until SocialList.Off And SocialNegList.Off
'        Trait1 = ""
'        Trait2 = ""
'        If Not SocialList.Off Then
'            Trait1 = SocialList.DisplayTrait
'            SocialList.MoveNext
'        End If
'        If Not SocialNegList.Off Then
'            Trait2 = SocialNegList.DisplayTrait
'            SocialNegList.MoveNext
'        End If
'        OutAid.Output Trait1, Trait2
'    Loop
'
'    OutAid.Output ""
'    OutAid.Output CStr(MentalList.Count) & " Mental Traits:", _
'           CStr(MentalNegList.Count) & " Negative Mental Traits:"
'    OutAid.Output "-------------------", "---------------------------"
'    MentalList.First
'    MentalNegList.First
'    Do Until MentalList.Off And MentalNegList.Off
'        Trait1 = ""
'        Trait2 = ""
'        If Not MentalList.Off Then
'            Trait1 = MentalList.DisplayTrait
'            MentalList.MoveNext
'        End If
'        If Not MentalNegList.Off Then
'            Trait2 = MentalNegList.DisplayTrait
'            MentalNegList.MoveNext
'        End If
'        OutAid.Output Trait1, Trait2
'    Loop
'
'    OutAid.Output ""
'    OutAid.Output CStr(InfluenceList.Count) & " Influences:", _
'           CStr(AbilityList.Count) & " Abilities:"
'    OutAid.Output "--------------", "-------------"
'    InfluenceList.First
'    AbilityList.First
'    Do Until InfluenceList.Off And AbilityList.Off
'        Trait1 = ""
'        Trait2 = ""
'        If Not InfluenceList.Off Then
'            Trait1 = InfluenceList.DisplayTrait
'            InfluenceList.MoveNext
'        End If
'        If Not AbilityList.Off Then
'            Trait2 = AbilityList.DisplayTrait
'            AbilityList.MoveNext
'        End If
'        OutAid.Output Trait1, Trait2
'    Loop
'
'    OutAid.Output ""
'    OutAid.Output CStr(BackgroundList.Count) & " Backgrounds:", _
'            CStr(HonorList.Count) & " Honor/Humor/Obedience:"
'    OutAid.Output "--------------", "------------------------"
'    BackgroundList.First
'    HonorList.First
'    Do Until BackgroundList.Off And HonorList.Off
'        Trait1 = ""
'        Trait2 = ""
'        If Not BackgroundList.Off Then
'            Trait1 = BackgroundList.DisplayTrait
'            BackgroundList.MoveNext
'        End If
'        If Not HonorList.Off Then
'            Trait2 = HonorList.DisplayTrait
'            HonorList.MoveNext
'        End If
'        OutAid.Output Trait1, Trait2
'    Loop
'
'    OutAid.Output ""
'    OutAid.Output CStr(GloryList.Count) & " Glory/Ferocity/Succor/Cunning:", _
'            CStr(WisdomList.Count) & " Wisdom/Cleverness:"
'    OutAid.Output "--------------------------------", "--------------------"
'    GloryList.First
'    WisdomList.First
'    Do Until GloryList.Off And WisdomList.Off
'        Trait1 = ""
'        Trait2 = ""
'        If Not GloryList.Off Then
'            Trait1 = GloryList.DisplayTrait
'            GloryList.MoveNext
'        End If
'        If Not WisdomList.Off Then
'            Trait2 = WisdomList.DisplayTrait
'            WisdomList.MoveNext
'        End If
'        OutAid.Output Trait1, Trait2
'    Loop
'
'    OutAid.Output ""
'    OutAid.Output "Health Levels:", "Other Features:"
'    OutAid.Output "--------------", "---------------"
'    HealthList.First
'    FeatureList.First
'    Do Until HealthList.Off And FeatureList.Off
'        Trait1 = ""
'        Trait2 = ""
'        If Not HealthList.Off Then
'            Trait1 = HealthList.DisplayTrait
'            HealthList.MoveNext
'        End If
'        If Not FeatureList.Off Then
'            Trait2 = FeatureList.DisplayTrait
'            FeatureList.MoveNext
'        End If
'        OutAid.Output Trait1, Trait2
'    Loop
'
'    OutAid.Output ""
'    OutAid.Output "Gifts:"
'    OutAid.Output "------"
'    GiftList.First
'    Do Until GiftList.Off
'        OutAid.Output GiftList.DisplayTrait
'        GiftList.MoveNext
'    Loop
'
'    OutAid.Output ""
'    OutAid.Output "Rites:"
'    OutAid.Output "------"
'    RiteList.First
'    Do Until RiteList.Off
'        OutAid.Output RiteList.DisplayTrait
'        RiteList.MoveNext
'    Loop
'
'    OutAid.Output ""
'    OutAid.Output "Merits:", "Flaws:"
'    OutAid.Output "-------", "------"
'    MeritList.First
'    FlawList.First
'    Do Until MeritList.Off And FlawList.Off
'        Trait1 = ""
'        Trait2 = ""
'        If Not MeritList.Off Then
'            Trait1 = MeritList.DisplayTrait
'            MeritList.MoveNext
'        End If
'        If Not FlawList.Off Then
'            Trait2 = FlawList.DisplayTrait
'            FlawList.MoveNext
'        End If
'        OutAid.Output Trait1, Trait2
'    Loop
'
'    OutAid.Output ""
'    OutAid.Output "Equipment / Fetishes:"
'    OutAid.Output "---------------------"
'    'OutAid.WrapOutput Equipment
'
'    If IncludeNotes Then
'        OutAid.Output ""
'        OutAid.Output "Notes:"
'        OutAid.Output "------"
'        OutAid.WrapOutput Notes
'    End If
'
'    If IncludeHistory Then
'        OutAid.Output ""
'        OutAid.Output "Experience History:"
'        OutAid.Output "-------------------"
'        'OutAid.WrapOutput History
'    End If
'
'    OutAid.Output String(PageWidth, "-")

End Sub

Public Function Group() As String
'
' Name:         Group
' Description:  Return the 'Group' this character belongs to.
'               Clan, Tribe, Kith, Guild, etc.
' Returns:      the group
'

    Group = ""

End Function
Public Function Race() As String
'
' Name:         Race
' Description:  Return a string representation of this character's
'               character type.
' Returns:      the name of the character type
'

    Race = Bete

End Function

Public Function RaceCode() As RaceType
'
' Name:         RaceCode
' Description:  Return the RaceType (enumeration value) of this
'               character type.
' Returns:      the corresponding RaceType
'
    
    RaceCode = gvRaceBete

End Function

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write this character's data to a file.
'

    With XML
    
        .BeginTag "bete"
    
            .WriteAttribute "name", Name
            .WriteAttribute "nature", Nature
            .WriteAttribute "demeanor", Demeanor
            .WriteAttribute "bete", Bete
            .WriteAttribute "breed", Breed
            .WriteAttribute "auspice", Auspice
            .WriteAttribute "rank", Rank
            .WriteAttribute "pack", Pack
            .WriteAttribute "totem", Totem
            .WriteAttribute "position", Position
            .WriteAttribute "notoriety", Notoriety, 0
            .WriteAttribute "rage", Rage
            .WriteAttribute "gnosis", Gnosis
            .WriteAttribute "willpower", Willpower
            .WriteAttribute "honor", Honor
            .WriteAttribute "glory", Glory
            .WriteAttribute "wisdom", Wisdom
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
            
            .WriteCDataTag "notes", Notes
        
        .EndTag

    End With

End Sub

Public Sub InputFromFile(XML As XMLReaderClass, Version As String)
'
' Name:         InputFromFile
' Parameters:   XML             XMLReaderClass that has read the tag
'               Version         file format version
' Description:  Reads the object in from an XML file.
'

    Dim ListName As String

    With XML
    
        If .Tag = "bete" Then
        
            Name = .GetString("name")
            Nature = .GetString("nature")
            Demeanor = .GetString("demeanor")
            Bete = .GetString("bete")
            Breed = .GetString("breed")
            Auspice = .GetString("auspice")
            Rank = .GetString("rank")
            Pack = .GetString("pack")
            Totem = .GetString("totem")
            Position = .GetString("position")
            Notoriety = .GetInteger("notoriety")
            Rage = .GetInteger("rage")
            Gnosis = .GetInteger("gnosis")
            Willpower = .GetInteger("willpower")
            Honor = .GetSingle("honor")
            Glory = .GetSingle("glory")
            Wisdom = .GetSingle("wisdom")
            Player = .GetString("player")
            Status = .GetString("status")
            ID = .GetString("id")
            StartDate = .GetString("startdate")
            Narrator = .GetString("narrator")
            IsNPC = .GetBoolean("npc")
            LastModified = .GetDate("lastmodified")
            
            Do Until .Tag = "bete" And .Terminal
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
                        End Select
                    Case "notes"
                        If Not .Terminal Then
                            .ReadTag
                            Notes = .GetTagText
                        End If
                End Select
            Loop
        
        Else
            Err.Raise vbObjectError + 8, "BeteClass", _
                      "Wrong file input to a bete character"
        End If
    
    End With
        
End Sub

Public Sub OldInputFromFile(FileNum As Integer, Version As String)
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
    Line Input #FileNum, Bete
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
    Line Input #FileNum, Read
    Gnosis = Val(Read)
    Line Input #FileNum, Read
    Willpower = Val(Read)

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
    If Version >= GameFileVersionTag5 Then
        FeatureList.LLInputFromFile FileNum, Version
    End If
    
    GiftList.LLInputFromFile FileNum, Version
    RiteList.LLInputFromFile FileNum, Version
    
    HonorList.OldInputFromFile FileNum, Version
    GloryList.OldInputFromFile FileNum, Version
    WisdomList.OldInputFromFile FileNum, Version
    
    MeritList.LLInputFromFile FileNum, Version
    FlawList.LLInputFromFile FileNum, Version
    
    Notes = ""
    If Version <> GameFileVersionTag0 Then
        Read = ReadLongField(FileNum, "<-End Bete Character Equipment->")
        If Read <> "" Then Notes = "2.3 Equipment:" & vbCrLf & Read & vbCrLf
    End If
    
    Notes = Notes & ReadLongField(FileNum, "<-End Bete Character Notes->")
    
    Line Input #FileNum, Player
    Line Input #FileNum, Read
    Experience.Unspent = Val(Read)
    Experience.Earned = Val(Read)
    Line Input #FileNum, Status
    Line Input #FileNum, ID
    Line Input #FileNum, StartDate
    Line Input #FileNum, Narrator
    Line Input #FileNum, Read
    IsNPC = CBool(Read)
    Experience.OldInputFromFile FileNum, Version
    LastModified = Now
    
End Sub

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Intialize all data to defaults.
'

    Notoriety = 0
    Rage = 3
    Gnosis = 3
    Willpower = 3
    
    Status = ActiveStatus
    StartDate = Format(Date, "mmmm d, yyyy")
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
    Set Experience = Nothing

End Sub
