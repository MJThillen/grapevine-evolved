Attribute VB_Name = "PublicVariables"
'
' File:             PublicVariables.bas
' Author:           Adam Cerling
' Description:      Declare all the public variables needed by Grapevine.
'
Option Explicit

'
' Variable:         Game
' Description:      Reference to the instance of GameClass.
' Used In:          Everything.
'
Public Game As GameClass

'
' Variables:        Utility Lists
' Description:      Refences to Lists within Game.
' Used In:          Everything.
'
Public PlayerList As LinkedList                         'Public Pointers to the game lists
Public CharacterList As LinkedList
Public ItemList As LinkedList
Public RoteList As LinkedList
Public LocationList As LinkedList
Public ActionList As LinkedList
Public PlotList As LinkedList
Public RumorList As LinkedList
Public AllRumorLists As LinkedList
Public InfluenceUseList As LinkedList

'
' Variables:        OutputEngine
' Description:      Instance of utility class that assists with all forms of output.
' Used In:          the printing screens.
'
Public OutputEngine As OutputEngineClass

'
' Variable:         StdHealth
' Description:      Array of the Four standard Health Levels
' Used In:          clsCharSheetEngine, Root, Character Classes
'
Public StdHealth(hlMinHealth To hlMaxHealth) As String

' Variable:         HideSTFromFile
' Description:      Whether to filter out ST comments when saving an Exchange file
' Used in:          XMLWriterClass, PutStrB function, frmExchange, GameClass
'
Public HideSTFromFile As Boolean
