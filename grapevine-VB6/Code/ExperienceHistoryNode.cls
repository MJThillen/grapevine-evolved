VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExperienceHistoryNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PrevNode As ExperienceHistoryNode    'Previous history entry
Public NextNode As ExperienceHistoryNode    'Next history entry

Public Change As Single                     'amount of change
Public ChangeType As ExperienceChangeType   'type of change

Public When As Date                         'date of change
Public Reason As String                     'reason for change
Public Earned As Single                     'historic earned total
Public Unspent As Single                    'historic unspent total

Public Sub PropogateChange()
'
' Name:         PropogateChange
' Description:  Update this entry and all etries after it according to its
'               current data.
'

    Dim PrevEarned As Single
    Dim PrevUnspent As Single

    If PrevNode Is Nothing Then
        PrevEarned = 0
        PrevUnspent = 0
    Else
        PrevEarned = PrevNode.Earned
        PrevUnspent = PrevNode.Unspent
    End If
    
    Select Case ChangeType
        Case ecEarned
            Earned = PrevEarned + Change
            Unspent = PrevUnspent + Change
        Case ecDeducted
            Earned = PrevEarned - Change
            Unspent = PrevUnspent - Change
        Case ecSetEarned
            Earned = Change
            Unspent = PrevUnspent
        Case ecSpent
            Unspent = PrevUnspent - Change
            Earned = PrevEarned
        Case ecUnspent
            Unspent = PrevUnspent + Change
            Earned = PrevEarned
        Case ecSetUnspent
            Earned = PrevEarned
            Unspent = Change
        Case ecComment
            Earned = PrevEarned
            Unspent = PrevUnspent
    End Select

    If Not NextNode Is Nothing Then NextNode.PropogateChange

End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this character's data to a binary file.
'

    Put #FileNum, , When
    Put #FileNum, , Change
    Put #FileNum, , ChangeType
    PutStrB FileNum, Reason
    Put #FileNum, , Earned
    Put #FileNum, , Unspent

End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this character's data from a binary file.
'
    
    Get #FileNum, , When
    Get #FileNum, , Change
    Get #FileNum, , ChangeType
    GetStrB FileNum, Reason
    Get #FileNum, , Earned
    Get #FileNum, , Unspent

End Sub

