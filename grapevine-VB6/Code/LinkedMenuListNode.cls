VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LinkedMenuListNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String                   'the name of the menu choice
Public Cost As String                   'the cost of the choice, or ":" if it is a link
Public Note As String                   'the note associated with the choice, or the
                                        '  linked menu if it is a link

Public NextNode As LinkedMenuListNode   'the next node
Public PrevNode As LinkedMenuListNode   'the previous node
