Attribute VB_Name = "PublicConstants"
'
' File:             PublicConstants.bas
' Author:           Adam Cerling
' Description:      Declare all the Constants needed by Grapevine.
'
Option Explicit

'
' Constants:        Grapevine Caption
' Description:      Title of current version of Grapevine
' Used In:          mdiMain
'
Public Const GrapevineCaption = "Grapevine 3.0"

'
' Constants:        Current Program Version
' Description:      Used to identify file types and versions
' Used in:          GameClass, MenuSetClass
'
Public Const ThisVersion As Double = 3#

'
' Constants:        APR constants
' Description:      Important strings dealing with actions and rumors
' Used in:          APREngineClass, ActionClass, frmAction
'
Public Const BasicSubactionName = "Personal"
Public Const PublicRumorTitle = "Public Knowledge"

Public Const RecentSearchName = "Most Recent Search"
Public Const BackupFileName = "~Autosave.gv3"

'
' Constants:        Binary File Header Data
' Description:      Used to identify binary file types and versions
' Used in:          Root, GameClass, MenuSetClass
'
Public Const BinHeaderLen As Integer = 4 'length of all binary headers
Public Const BinHeaderGame As String = "GVBG"
Public Const BinHeaderMenu As String = "GVBM"
Public Const BinHeaderExchange As String = "GVBE"

'
' Constants:        WWW Locations
' Description:      Location of Grapevine WWW resources
' Used In:          mdiMain, frmAbout
'
Public Const URLMainPage = "http://www.GrapevineLARP.com/"
Public Const URLHelpPage = "http://www.GrapevineLARP.com/help.shtml"

'
' Constants:        Default filenames
' Description:      Names of the default files
' Used In:          GameClass, MenuSetClass
'
Public Const DefaultMenuFile = "Grapevine Menus.gvm"
Public Const DefaultInventoryFile = "New Game Items.gex"

'
' Constants:        Game File Versions
' Description:      The headers to the evolving file formats of Grapevine
' Used In:          GameClass, various OldInputFromFile routines
'
Public Const GameFileVersionTag0 = "<-Grapevine II Game File->"
Public Const GameFileVersionTag1 = "<-Grapevine 2.0 Game File / Format 1->"
Public Const GameFileVersionTag2 = "<-Grapevine 2.0 Game File / Format 2->"
Public Const GameFileVersionTag3 = "<-Grapevine 2.1 Game File / Format 1->"
Public Const GameFileVersionTag4 = "<-Grapevine 2.2 Game File / Format 1->"
Public Const GameFileVersionTag5 = "<-Grapevine 2.3 Game File / Format 1->"

Public Const ExchangeFileVersionTag0 = "<-Exchange File / Grapevine 2.2 / Format 1 ->"
Public Const ExchangeFileVersionTag1 = "<-Exchange File / Grapevine 2.3 / Format 1 ->"

'
' Constants:        Active Status
' Description:      The Item in the Status menus that means "Active"
' Used In:          frmCharacters, frmOutputSheets, frmOutputRumors, frmPointMaintenance
'
Public Const ActiveStatus = "Active"

'
' Constants:        Point Type tags
' Description:      A value that controls whether an instance of frmPointMaintenance
'                   is tracking experience or player points
' Used In:          frmPointMaintenance
'
Public Const pmExperience = "E"
Public Const pmPlayerPoints = "P"

'
' Constants:        Health Level constants
' Description:      Values of standard health leves
' Used In:          clsCharSheetEngine, Character Classes, Root
'
Public Const hlStdHealth0 = "Healthy"
Public Const hlStdHealth1 = "Bruised"
Public Const hlStdHealth2 = "Wounded"
Public Const hlStdHealth3 = "Incapacitated"
Public Const hlStdHealth4 = "Mortally Wounded"
Public Const hlMinHealth = 0
Public Const hlMaxHealth = 4

'
' Constants:        Plot Status constants
' Description:      Statuses for plot objects
' Used In:          PlotClass, associated forms
'
Public Const psActive = "Active"
Public Const psFinished = "Finished"
Public Const psPending = "Pending"

' Constants:        Output ID Constants
' Description:      Returned by several objects to describe how they are to be
'                   maipulated by the OutputEngineClass
' Used In:          OutputEngineClass, LinkedTraitList, ExperienceClass,
'                   PlotClass, ActionClass, RumorClass
'
Public Const oidNone = -1
Public Const oidTraitList = 0
Public Const oidHistory = 1
Public Const oidPlot = 2
Public Const oidAction = 3
Public Const oidRumor = 4
Public Const oidCalendar = 5

' Constants:        Output Selection Constants
' Description:      Indices of selections of objects for output
' Used In:          OutputEngineClass, output forms
'
Public Const osMin = 1
Public Const osPlayers = 1
Public Const osCharacters = 2
Public Const osItems = 3
Public Const osRotes = 4
Public Const osLocations = 5
Public Const osActions = 6
Public Const osPlots = 7
Public Const osRumors = 8
Public Const osMax = 8
Public Const osSearch = 9
Public Const osStatistics = 10

'
' Constants:        Standard Template Names
' Description:      Names of the standard templates Grapevine expects to see
' Used In:          TemplateClass, OutputEngineClass, output forms, all forms
'                   with a SetDefaultOutput method
'
Public Const tnCharSheetSuffix = " Character Sheet"
Public Const tnActionRumor = "Action and Rumor Report"
Public Const tnMasterAction = "Master Action Report"
Public Const tnMasterRumor = "Master Rumor Report"
Public Const tnPlot = "Plot Report"
Public Const tnCharacterSheets = "Character Sheets"
Public Const tnCharacterRoster = "Character Roster"
Public Const tnEquipment = "Character Equipment"
Public Const tnSignIn = "Sign-In Sheet"
Public Const tnItemCards = "Item Cards"
Public Const tnRoteCards = "Rote Cards"
Public Const tnLocationCards = "Location Cards"
Public Const tnXPHistory = "Experience History"
Public Const tnPPHistory = "Player Point History"
Public Const tnPlayerRoster = "Player Roster"
Public Const tnGameCalendar = "Game Calendar"
Public Const tnSearch = "Search Report"
Public Const tnStatistics = "Statistics Report"
Public Const tnVampireStatus = "Vampire Status Report"
Public Const tnMFReport = "Merits and Flaws Report"
Public Const tnInfluenceReport = "Influence Report"

'
' Constants:        Special E-Mail Recipients
' Description:      Names of special e-mail recipients
' Used In:          frmEMailAddressing, frmOutput, OutputEngineClass
'
Public Const SendToSelect = "[Send_to_Select]"

'
' Constants:        Character Value Keys
' Description:      Keywords used to retrieve values from character classes
' Used In:          Character Classes, Search and Statistics tools, Printing
'
' --- See PublicQueryKeys module

Public Const xeKey = "Gv3.0XK!"
