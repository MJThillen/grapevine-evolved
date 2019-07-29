Attribute VB_Name = "PublicTypes"
'
' File:             PublicTypes.bas
' Author:           Adam Cerling
' Description:      Declare all the public types and enums needed by Grapevine.
'
Option Explicit

'
' Enum:             FileFormatType
' Description:      Different types of file formats.
' Used In:          GameClass, MenuSetClass, mdiMain, frmMenuEditor, frmExchange
'
Public Enum FileFormatType
    gvInvalid = 0
    gv23Game = 1
    gv23Exchange = 2
    gvXML = 3
    gvBinaryGame = 4
    gvBinaryMenu = 5
    gvBinaryExchange = 6
End Enum

'
' Enum:             RaceType
' Description:      An efficient method of communicating race.
' Used In:          frmAddNewCharacter, frmCharacter, Character classes, GameClass
'
Public Enum RaceType
    gvMenuDelete = -1
    gvRaceNone = 0
    gvRaceAll = 1
    gvRaceVampire = 2
    gvRaceWerewolf = 3
    gvRaceMortal = 4
    gvRaceChangeling = 5
    gvRaceWraith = 6
    gvracemage = 7
    gvRaceFera = 8
    gvRaceVarious = 9
    gvRaceMummy = 10
    gvRaceKueiJin = 11
    gvRaceHunter = 12
    gvRaceDemon = 13
End Enum

'
' Enum:             ListDisplayType
' Description:      The formats in which lists and menus can display items.
' Used In:          LinkedMenuList, LinkedTraitList
'
Public Enum ListDisplayType
    ldDefault = -1
    ldSimple = 0
    ldMultiplier = 1
    ldMultiplierDot = 2
    ldDot = 3
    ldCost = 4
    ldNoteOnly = 5
    ldCostOnly = 6
    ldDotSeparate = 7
    ldSimpleDots = 8
    ldSimpleNumber = 9
    ldSimpleNote = 10
End Enum

'
' Enum:             AnnounceType
' Description:      Values to signal to forms that data elsewhere in the program has changed
' Used In:          Forms that alter or display data
'
Public Enum AnnounceType            'Game components for use with AnnounceChanges and CheckForChanges
    atCharacters = 0                 'Changes to the list of characters (add, delete, name change)
    atPlayers = 1                    'Changes to the list of players (add, delete, name change)
    atQueries = 2                       'Changes to the list of queries
    atItems = 3                         'Changes to the list of items
    atRotes = 4                         'Changes to the list of rotes
    atLocations = 5                     'Changes to the list of locations
    atActions = 6                       'Changes to the list of actions
    atPlots = 7                         'Changes to the list of plots
    atRumors = 8                        'Changes to the list of rumors
    atDates = 9                         'Changes to the game dates
    atTempers = 10                      'Changes to character tempers (Willpower, Blood, etc.)
    atStatus = 11                       'Changes to character tempers (Willpower, Blood, etc.)
End Enum

Public Const MIN_ANNOUNCE = 0           'Highest and lowest announcement values
Public Const MAX_ANNOUNCE = 11


'
' Enum:             ExperienceChangeType
' Description:      The type of changes experience and player points undergo
' Used In:          ExperienceClass, frmPointMaintenance
'
Public Enum ExperienceChangeType
    ecEarned = 0
    ecDeducted = 1
    ecSetEarned = 2
    ecSpent = 3
    ecUnspent = 4
    ecSetUnspent = 5
    ecComment = 6
End Enum

'
' Enum:             RumorCategoryType
' Description:      The type of rumor this is
' Used In:          RumorClass, related forms
'
Public Enum RumorCategoryType
    rtGeneral = 0
    rtInfluence = 1
    rtRace = 2
    rtGroup = 3
    rtSubgroup = 4
    rtPersonal = 5
End Enum

'
' Enum:         OutputFormatType
' Description:  The type of format of a template
' Used in:      OutputEngineClass, TemplateClass, frmOutput
'
Public Enum OutputFormatType
    ofText = 0
    ofRTF = 1
    ofHTML = 2
End Enum

Public Const MIN_OUTFORMAT = 0
Public Const MAX_OUTFORMAT = 2
