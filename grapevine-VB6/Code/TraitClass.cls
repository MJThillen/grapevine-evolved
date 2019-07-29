VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TraitClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String       'the name of the trait
Public Total As String       'the Total of this trait
Public Note As String       'optional trait note

Public Function Number() As Double
'
' Name:         Number
' Description:  Returns the numeric value of Total, or zero if it's not
'               numeric.
' Returns:      as above.
'
    
    Number = 0
    If IsNumeric(Total) Then Number = CDbl(Total)

End Function

Public Function HasNumber() As Boolean
'
' Name:         HasNumber
' Description:  Returns whether this trait has a numeric Total field.
' returns:      TRUE if Total is numeric.
'

    HasNumber = IsNumeric(Total)

End Function

Public Sub Decrement()
'
' Name:         Decrement
' Description:  Decrease the Total of this Trait by one.  If
'               the Total is a non-numeric string, look for
'               the first numeric part and set it to that.
' NOTE:         Bypassing the LinkedTraitList's Decrement method and using
'               this directly will result in an inconsistent Count for the
'               list.
'

    If IsNumeric(Total) Then
        Total = CStr(CDbl(Total) - 1)
    Else
    
        Dim I As Integer

        For I = 1 To Len(Total)
            If Not Mid(Total, I, 1) Like "#" Then Exit For
        Next I
        
        Total = Left(Total, I - 1)
        If Total = "" Then Total = "0"

    End If

End Sub

Public Sub Increment()
'
' Name:         Increment
' Description:  Increase the Total of this Trait by one.  If
'               the Total is a non-numeric string, look for
'               the last numeric part and set it to that.
' NOTE:         Bypassing the LinkedTraitList's Increment method and using
'               this directly will result in an inconsistent Count for the
'               list.
'

    If IsNumeric(Total) Then
        Total = CStr(CDbl(Total) + 1)
    Else
    
        Dim I As Integer
    
        For I = Len(Total) To 1 Step -1
            If Not Mid(Total, I, 1) Like "#" Then Exit For
        Next I
        Total = Mid(Total, I + 1)
        If Total = "" Then Total = "0"
        
    End If

End Sub


