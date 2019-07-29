VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExperienceClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Earned As Single       'Experience Earned
Public Unspent As Single      'Experience Unspent
Attribute Unspent.VB_VarMemberFlags = "200"
Public History As String      'Experience History (unused -- remove)

Public Sub OutputToFile(FileNum As Integer)
'
' Name:         OutputToFile
' Parameters:   FileNum     number of an open file
' Description:  Write the history to a file.
'

    Print #FileNum, Earned
    Print #FileNum, Unspent
    Print #FileNum, History
    Print #FileNum, "<-End Experience History->"

End Sub

Public Sub InputFromFile(FileNum As Integer, Version As String)
'
' Name:         InputFromFile
' Parameters:   FileNum     number of an open file
'               Version     version tag of the file format
' Description:  Read a ExperienceClass from a file.
'

    Dim Read As String

    If Version >= GameFileVersionTag6 Then
        Line Input #FileNum, Read
        Earned = Val(Read)
        Line Input #FileNum, Read
        Unspent = Val(Read)
        History = ReadLongField(FileNum, "<-End Experience History->")
    Else
        Do Until Read Like "<-End * History->"
            Line Input #FileNum, Read
        Loop
    End If

End Sub
