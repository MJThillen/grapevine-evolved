VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OldRumorClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Public Category As Integer      'Category of rumor: General, Group, Influence, Personal
Public Recipient As String      'Recipient of rumor
Public Level As Integer         'Influence Level (if any)
Public Text As String           'main rumor text
Public UsedBy As String         'Character using influence
Public Exclude As Boolean       'whether to exclude rumors for this influence level

