VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderedMenuClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum MenuType
    mtDefault
    mtPowerLevel
    mtVariableCost
    mtTopLevel
End Enum

Private Menus As LinkedList

Public Sub FillMenu(lvwList As ListView, MenuName As String, Style As MenuType)
'
' Name:         FillMenu
' Description:  Fills a ListView with items from the specified menu.
' Arguments:    A reference to the ListView, the name of the menu to use, and a
'               the type of menu to create.

    Dim SourceMenu As LinkedList
    Dim NewItem As ListItem
    
    Set SourceMenu = Nothing
    
    If Style = mtTopLevel Then
        Set SourceMenu = Menus
    Else
        Menus.MoveTo MenuName
        If Not Menus.Off Then Set SourceMenu = Menus.Item
    End If
    
    If Not SourceMenu Is Nothing Then
    
        SourceMenu.First
        Do Until SourceMenu.Off
        
            Set NewItem = lvwList.ListItems.Add(, , SourceMenu.Item.Name)
            
            Select Case Style
                Case mtPowerLevel
                    NewItem.Text = NewItem.Text & " (" & SourceMenu.Item.PowerLevel & ")"
                Case mtVariableCost
                    NewItem.Text = NewItem.Text & " (" & SourceMenu.Item.Cost & ")"
            End Select
    
            SourceMenu.MoveNext
            
        Loop
    
    End If

End Sub

Public Sub
