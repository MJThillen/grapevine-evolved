VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CauseEffectNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Target As APRType
Public When As Date
Public Item As String
Public Subitem As String

Public NextNode As CauseEffectNode
Public PrevNode As CauseEffectNode

Public Function ShortDesc(CompDate As Date) As String
'
' Name:         ShortDesc
' Description:  Return a Short text description of this link.
'
    
    Select Case Target
        Case aprAction
            
            Dim I As Integer
            Dim Name As String
            
            I = InStr(Item, " ")
            If I = 0 Then I = Len(Item) + 1
            Name = Left(Item, I - 1)
            
            ShortDesc = Name & " " & Subitem & " Action"
            
        Case aprPlot
            
            ShortDesc = Item & " Plot"
        
        Case aprRumor
        
            If Subitem = "0" Then
                ShortDesc = Item & " Rumor"
            Else
                If Right(Item, 10) = " Influence" Then
                    ShortDesc = Left(Item, Len(Item) - 9) & Subitem & " Rumor"
                Else
                    ShortDesc = Item & " " & Subitem & " Rumor"
                End If
            End If
            
    End Select

    If When <> CompDate Then ShortDesc = Format(When, "m/d ") & ShortDesc

End Function

Public Function TargetName() As String
'
' Name:         TargetName
' Description:  Return the Name field of the associated item.
'
    
    Select Case Target
        Case aprAction, aprRumor
            TargetName = Format(When, "Short Date") & " " & Item
        Case aprPlot
            TargetName = Item
    End Select

End Function

