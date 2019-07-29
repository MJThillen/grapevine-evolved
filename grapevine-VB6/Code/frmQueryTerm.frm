VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmQueryTerm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Query Term"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   1200
      Width           =   2055
   End
   Begin VB.ComboBox cboKey 
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.ComboBox cboCompare 
      Height          =   315
      ItemData        =   "frmQueryTerm.frx":0000
      Left            =   2160
      List            =   "frmQueryTerm.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.CheckBox chkNot 
      Caption         =   "NOT"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Left            =   4320
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox txtNumber 
      Height          =   315
      Left            =   6600
      TabIndex        =   5
      Text            =   "1"
      Top             =   240
      Width           =   615
   End
   Begin MSComCtl2.UpDown updNumber 
      Height          =   315
      Left            =   7215
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   240
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      Value           =   999
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtNumber"
      BuddyDispid     =   196615
      OrigLeft        =   7215
      OrigTop         =   240
      OrigRight       =   7710
      OrigBottom      =   555
      Max             =   999
      Min             =   -999
      Orientation     =   1
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label lblX 
      Caption         =   "x"
      Height          =   255
      Left            =   6450
      TabIndex        =   4
      Top             =   285
      Width           =   135
   End
End
Attribute VB_Name = "frmQueryTerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Query As QueryClass              'Query to edit
Private Editing As Boolean               'Whether to edit a clause or add a new one
Private EditKey As String                'Edit key to select when cboKey is populated
Private EditCompare As QueryCompareType  'Edit comparison to select when cboCompare is populated

Public Sub AddQueryTerm(MyQuery As QueryClass)
'
' Name:         AddQueryTerm
' Parameters:   MyQuery           Query to add a term to
' Description:  Add a new term to this query.
'

    Editing = False
    Set Query = MyQuery
    PopulateKeys
    txtFind.Text = ""
    txtNumber.Text = "1"
    chkNot.Value = vbUnchecked
    Me.Show vbModal, mdiMain

End Sub

Public Sub EditQueryTerm(MyQuery As QueryClass)
'
' Name:         EditQueryTerm
' Parameters:   MyQuery         Query whose CURRENT clause this form edits
' Description:  Edit the current clause of the given query.
'

    Editing = True
    Set Query = MyQuery
    
    With Query.Clause
        EditKey = .Key
        EditCompare = .Comparison
        PopulateKeys
        chkNot.Value = IIf(.CompNot, vbChecked, vbUnchecked)
        txtFind.Text = .Find
        txtNumber.Text = CStr(.Number)
    End With

    Me.Show vbModal, mdiMain

End Sub

Private Sub PopulateKeys()
'
' Name:         PopulateKeys
' Description:  Fill cboKey with character query keys and select one
'               (which in turn triggers cboKey_Click and cboCompare_Click).
'               If cboKey is already filled, just make the selection.
'

    Dim Key As Variant
    Dim I As Integer
    
    If Not Editing Then EditKey = qkGroup
    
    If cboKey.ListCount = 0 Then
        With Game.QueryEngine
            For Each Key In .TitlesToKeys
                If (.KeysToInventories(Key) And qiCharacters) Then
                    cboKey.AddItem CStr(.KeysToTitles(Key))
                    If Key = EditKey Then cboKey.ListIndex = cboKey.NewIndex
                End If
            Next Key
        End With
    Else
        For I = 0 To cboKey.ListCount - 1
            Key = Game.QueryEngine.TitlesToKeys(cboKey.List(I))
            If Key = EditKey Then cboKey.ListIndex = I
        Next I
    End If
    
End Sub

Private Sub ShowInputFields()
'
' Name:         ShowInputFields
' Description:  Format the visibility and alignment of txtFind, lblX,
'               txtNumber and updNumber according to the comparison
'               selected.
'
    
    Dim ShowFind As Boolean
    Dim ShowNumber As Boolean
    
    Select Case CLng(cboKey.Tag)
        
        Case qtField
            ShowFind = True
            ShowNumber = False
        Case qtNumber
            ShowFind = False
            ShowNumber = True
        Case qtTraitList
            Select Case cboCompare.ItemData(cboCompare.ListIndex)
                Case qcContains, qcContainsNote
                    ShowFind = True
                    ShowNumber = False
                Case qcTotals, qcTotalsLess, qcTotalsNoMore, qcTotalsMore, qcTotalsAtLeast
                    ShowFind = False
                    ShowNumber = True
                Case Else
                    ShowFind = True
                    ShowNumber = True
            End Select
        Case qtDate
            ShowFind = True
            ShowNumber = False
        Case qtBoolean
            ShowFind = False
            ShowNumber = False
    End Select
    
    txtFind.Visible = ShowFind
    lblX.Visible = ShowFind And ShowNumber
    txtNumber.Visible = ShowNumber
    updNumber.Visible = ShowNumber
        
    If ShowNumber And Not ShowFind Then
        txtNumber.Left = txtFind.Left
        updNumber.Left = txtNumber.Left + txtNumber.Width
    End If
    
End Sub

Private Sub cboCompare_Click()
'
' Name:         cboCompare_Click
' Description:  If necessary, rearrange the input fields.
'

    If cboKey.Tag = CStr(qtTraitList) Then ShowInputFields

End Sub

Private Sub cboKey_Click()
'
' Name:         cboKey_Click
' Description:  Populate cboCompare depending on the type of the key selected.
'

    Dim Key As String
    Dim KeyType As QueryKeyType
    Dim AutoSelect As Integer
    
    With Game.QueryEngine
        
        Key = .TitlesToKeys(cboKey.Text)
        KeyType = .KeysToTypes(Key)
        
        If cboKey.Tag <> CStr(KeyType) Then
            
            cboKey.Tag = CStr(KeyType)
            
            With cboCompare
            
                .Clear
                Select Case KeyType
                    Case qtField
                        .AddItem "contains", 0
                        .ItemData(0) = qcContains
                        .AddItem "equals", 1
                        .ItemData(1) = qcEquals
                        AutoSelect = 1
                    Case qtNumber
                        .AddItem "is less than", 0
                        .ItemData(0) = qcLess
                        .AddItem "is no more than", 1
                        .ItemData(1) = qcNoMore
                        .AddItem "equals", 2
                        .ItemData(2) = qcEquals
                        .AddItem "is at least", 3
                        .ItemData(3) = qcAtLeast
                        .AddItem "is greater than", 4
                        .ItemData(4) = qcGreater
                        AutoSelect = 2
                    Case qtTraitList
                        .AddItem "total less than", 0
                        .ItemData(0) = qcTotalsLess
                        .AddItem "total no more than", 1
                        .ItemData(1) = qcTotalsNoMore
                        .AddItem "total", 2
                        .ItemData(2) = qcTotals
                        .AddItem "total at least", 3
                        .ItemData(3) = qcTotalsAtLeast
                        .AddItem "total more than", 4
                        .ItemData(4) = qcTotalsMore
                        .AddItem "contain", 5
                        .ItemData(5) = qcContains
                        .AddItem "contain the note", 6
                        .ItemData(6) = qcContainsNote
                        .AddItem "contain less than", 7
                        .ItemData(7) = qcContainsLess
                        .AddItem "contain no more than", 8
                        .ItemData(8) = qcContainsNoMore
                        .AddItem "contain exactly", 9
                        .ItemData(9) = qcContainsExactly
                        .AddItem "contain at least", 10
                        .ItemData(10) = qcContainsAtLeast
                        .AddItem "contain more than", 11
                        .ItemData(11) = qcContainsMore
                        AutoSelect = 2
                    Case qtDate
                        .AddItem "is earlier than", 0
                        .ItemData(0) = qcLess
                        .AddItem "is no later than", 1
                        .ItemData(1) = qcNoMore
                        .AddItem "equals", 2
                        .ItemData(2) = qcEquals
                        .AddItem "is no earlier than", 3
                        .ItemData(3) = qcAtLeast
                        .AddItem "is later than", 4
                        .ItemData(4) = qcGreater
                        AutoSelect = 2
                        txtFind.Text = Format(Now, "Short Date")
                    Case qtBoolean
                        .AddItem "is true", 0
                        .ItemData(0) = qcIsTrue
                        .AddItem "is false", 1
                        .ItemData(1) = qcIsFalse
                        AutoSelect = 0
                End Select
                    
                If Editing Then
                    For AutoSelect = 0 To .ListCount - 1
                        If .ItemData(AutoSelect) = EditCompare Then
                            .ListIndex = AutoSelect
                        End If
                    Next AutoSelect
                Else
                    .ListIndex = AutoSelect
                End If
                
            End With
        
            ShowInputFields
        
        End If
        
    End With

End Sub

Private Sub cmdCancel_Click()
'
' Name:         cmdCancel_Click
' Description:  Do nothing but leave this window.
'
    Hide

End Sub

Private Sub cmdOK_Click()
'
' Name:         cmdOK_Click
' Description:  Apply the changes and disappear.
'

    Dim Key As String
    Dim Comp As Long
    
    Key = Game.QueryEngine.TitlesToKeys(cboKey.Text)
    Comp = cboCompare.ItemData(cboCompare.ListIndex)
    
    If Editing Then
        With Query.Clause
            .Key = Key
            .Find = txtFind.Text
            .Number = Val(txtNumber.Text)
            .Comparison = Comp
            .CompNot = (chkNot.Value = vbChecked)
        End With
    Else
        Query.AddClause Key, txtFind.Text, Val(txtNumber.Text), Comp, (chkNot.Value = vbChecked)
    End If
    
    Hide

End Sub

Private Sub txtFind_GotFocus()
'
' Name:         txtFind_GotFocus
' Description:  Select this box of text.
'
    SelectText txtFind
    
End Sub

Private Sub txtNumber_GotFocus()
'
' Name:         txtNumber_GotFocus
' Description:  Select this box of text.
'
    SelectText txtNumber

End Sub
