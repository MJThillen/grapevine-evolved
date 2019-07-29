VERSION 5.00
Begin VB.Form frmSelectFromList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkNot 
      Alignment       =   1  'Right Justify
      Caption         =   "Not"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.ComboBox cboSearch 
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2595
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.ListBox lstList 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmSelectFromList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Inventory As QueryInventoryType

Public Function ShowSelect(Inv As QueryInventoryType, SearchName As String, Title As String) As String
'
' Name:         ShowSelect
' Parameters:   Inv             The inventory to show queries for
'               SearchName      The default search name to start with
'               Title           Title for the window
' Description:  Modally show the results of a merge operation or exchange load.
'

    Dim I As Integer

    I = -1
    Me.Caption = Title
    Inventory = Inv
    
    With Game.QueryEngine.QueryList
        .First
        Do Until .Off
            If .Item.Inventory = Inventory Then
                cboSearch.AddItem .Item.Name
                If .Item.Name = SearchName Then I = cboSearch.NewIndex
            End If
            .MoveNext
        Loop
    End With
        
    cboSearch.Visible = (cboSearch.ListCount > 0)
    chkNot.Visible = (cboSearch.ListCount > 0)
    
    If I > -1 Then
        cboSearch.ListIndex = I
    Else
        Call cboSearch_Click
    End If

    Me.Show vbModal, mdiMain

    ShowSelect = lstList.Text

    Unload Me

End Function

Private Sub cboSearch_Click()
'
' Name:         cboSearch_Click
' Description:  Flls the list box from the list of those matching the chosen search.
'

    Dim Search As QueryClass
    
    Screen.MousePointer = vbHourglass
    
    lstList.Clear
    
    With Game.QueryEngine.QueryList
        .MoveTo cboSearch.Text
        If Not .Off Then
            Set Search = .Item
        Else
            Set Search = New QueryClass
            Search.Inventory = Inventory
        End If
    End With
    
    With Game.QueryEngine
        .MakeQuery Search, , (chkNot.Value = vbChecked)
    
        .Results.First
        Do Until .Results.Off
            lstList.AddItem .Results.Item.Name
            .Results.MoveNext
        Loop
        
        If Not .Results.IsEmpty Then lstList.ListIndex = 0
    End With
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub chkNot_Click()
'
' Name:         chkNot_Click
' Description:  Repopulate the list.
'
    Call cboSearch_Click
    
End Sub

Private Sub cmdOK_Click()
'
' Name:         cmdOK_Click
' Description:  Hide this form.
'

    Me.Hide

End Sub

Private Sub cmdCancel_Click()
'
' Name:         cmdCancel_Click
' Description:  Deselect the list and hide this form.
'
    lstList.ListIndex = -1
    Me.Hide

End Sub

Private Sub lstList_DblClick()
'
' Name:         lstList_DblClick
' Description:  Affirm the choice.
'
    Call cmdOK_Click

End Sub
