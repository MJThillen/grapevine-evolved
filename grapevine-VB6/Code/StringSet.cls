VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' Name:         StringSet.cls
' Description:  A small utility class that adds and subtracts from a set of strings,
'               using string operations.
'

Private MetaString As String
Private StringCount As Long
Private Cursor As Long

Private Const D As String = vbCr

Public Sub Add(ByRef S As String)
'
' Name:         Add
' Parameter:    S        a string to add
' Description:  Add a string to the set
'

    If Not Has(S) Then
        MetaString = MetaString & D & S & D
        StringCount = StringCount + 1
    End If
    
End Sub

Public Sub Remove(ByRef S As String)
'
' Name:         Remove
' Parameter:    S       a string to remove
' Description:  Remove a string from the set

    Dim I As Integer
    
    I = InStr(MetaString, (D & S & D))
    If I > 0 Then
        MetaString = Left(MetaString, I - 1) & Mid(MetaString, I + Len(D & S & D))
        StringCount = StringCount - 1
    End If

End Sub

Public Sub First()
'
' Name:         First
' Description:  Move to the first string in the set

    Cursor = 1

End Sub

Public Function Off() As Boolean
'
' Name:         Off
' Description:  Return TRUE if the cursor has left the set.

    Off = Cursor > Len(MetaString)

End Function

Public Sub MoveNext()
'
' Name:         MoveNext
' Description:  Move the cursor to the next name.

    Cursor = InStr(Cursor + 1, MetaString, D) + Len(D)

End Sub

Public Function StrItem() As String
'
' Name:         StrItem
' Description:  Return the string under the cursor.

    If Not Off Then _
        StrItem = Mid(MetaString, Cursor + 1, InStr(Cursor + 1, MetaString, D) - (Cursor + 1))

End Function

Public Sub Clear()
'
' Name:         Clear
' Description:  Clear the set.

    MetaString = ""
    Cursor = 1
    StringCount = 0

End Sub

Public Function Count() As Long
'
' Name:         Count
' Description:  Return the number of strings in the set.

    Count = StringCount

End Function

Public Function Has(ByRef S As String) As Boolean
'
' Name:         Has
' Parameter:    S       string to look for
' Description:  return TRUE if the string is in the set

    Has = (InStr(MetaString, (D & S & D)) > 0)

End Function

Public Function ViewString() As String
'
' Name:         ViewString
' Description:  View the set's internal string
'

    ViewString = MetaString

End Function

Public Sub StoreListBox(ByRef Box As ListBox, Optional All As Boolean = False)
'
' Name:         StoreListView
' Parameter:    Box        the ListBox whose selections to store
'               All        whether to select all entries, or only the selected ones
' Description:  Store the listbox's selections in the set.
'

    Dim I As Integer
    
    Clear
    For I = 0 To Box.ListCount - 1
        If Box.Selected(I) Or All Then Add Box.List(I)
    Next I

End Sub

Public Sub SelectListBox(ByRef Box As ListBox, SelFound As Boolean, SelNotFound As Boolean)
'
' Name:         SelectListBox
' Parameters:   Box         ListBox to make selections from
'               SelFound    whether to select items found in the set
'               SelNotFound whether to select items not found in the set
' Description:  Select items in a multiple-selection list box based on whether or not they're
'               in the string set.

    Dim I As Integer

    For I = (Box.ListCount - 1) To 0 Step -1
        Box.Selected(I) = IIf(Has(Box.List(I)), SelFound, SelNotFound)
    Next I

End Sub

Public Sub StoreListView(ByRef View As ListView, Optional All As Boolean = False)
'
' Name:         StoreListView
' Parameter:    View        the ListView whose selections to store
'               All         whether to store All entries
' Description:  Store the listview's selections in the set.
'

    Dim I As Integer
    
    Clear
    For I = 1 To View.ListItems.Count
        If View.ListItems(I).Selected Or All Then Add View.ListItems(I).Text
    Next I

End Sub

Public Sub SelectListView(ByRef View As ListView, SelFound As Boolean, SelNotFound As Boolean)
'
' Name:         SelectListView
' Parameters:   View         ListView to make selections from
'               SelFound    whether to select items found in the set
'               SelNotFound whether to select items not found in the set
' Description:  Select items in a multiple-selection list view based on whether or not they're
'               in the string set.

    Dim I As Integer
    
    For I = 1 To View.ListItems.Count
        View.ListItems(I).Selected = IIf(Has(View.ListItems(I).Text), SelFound, SelNotFound)
    Next I

End Sub

Private Sub Class_Initialize()
'
' Name:         Class_Initialize
' Description:  Initialize the cursor.

    Cursor = 1

End Sub
