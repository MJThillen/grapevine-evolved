VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmInfluenceUse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Influence Use and Results"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   Icon            =   "frmInfluenceUse.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   9030
   Begin VB.CommandButton cmdESCClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdClearDate 
      Caption         =   "Clear all for this Date"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6360
      TabIndex        =   12
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CheckBox chkMarkUsed 
      Caption         =   "&Mark this Influence ""Used"" with regards to the Character's Influence Rumors for this Date"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   960
      Width           =   4695
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select this Date and Character"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   2880
      Width           =   2535
   End
   Begin VB.CommandButton cmdClearCharacter 
      Caption         =   "Clear all for this Character"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3840
      TabIndex        =   11
      Top             =   5400
      Width           =   2175
   End
   Begin VB.TextBox txtUse 
      Enabled         =   0   'False
      Height          =   2535
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1800
      Width           =   4695
   End
   Begin VB.CommandButton cmdClearUse 
      Caption         =   "C&lear this Use"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6360
      TabIndex        =   10
      Top             =   4440
      Width           =   2175
   End
   Begin VB.ComboBox cmbDate 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.ListBox lstCharacters 
      Height          =   1425
      IntegralHeight  =   0   'False
      ItemData        =   "frmInfluenceUse.frx":08CA
      Left            =   480
      List            =   "frmInfluenceUse.frx":08CC
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1320
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid grdInfluences 
      Height          =   2055
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3625
      _Version        =   393216
      BackColorBkg    =   -2147483633
      GridColor       =   -2147483633
      AllowBigSelection=   0   'False
      Enabled         =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      BorderStyle     =   0
      FormatString    =   " | ; | "
   End
   Begin VB.Label lblEmpty 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--No Character Selected--"
      Height          =   1815
      Left            =   360
      TabIndex        =   18
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label lblInfluence 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8460
      TabIndex        =   17
      Top             =   600
      Width           =   75
   End
   Begin VB.Label lblLabels 
      Caption         =   "Select an &Influence:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Use and Results:"
      Height          =   255
      Index           =   5
      Left            =   3840
      TabIndex        =   8
      Top             =   1560
      Width           =   4695
   End
   Begin VB.Label lblCharacterName 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      TabIndex        =   15
      Top             =   600
      Width           =   75
   End
   Begin VB.Label lblLabels 
      Caption         =   "Influence Use and/or Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   3720
      TabIndex        =   14
      Top             =   240
      Width           =   2520
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8595
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Date:"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Top             =   660
      Width           =   390
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Character:"
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   4815
      Index           =   4
      Left            =   3600
      TabIndex        =   16
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   2775
      Index           =   6
      Left            =   240
      TabIndex        =   19
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "frmInfluenceUse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ThisCharacter As String         'The selected character
Private ThisInfluence As String         'The selected influence
Private ThisLevel As Integer            'The selected influence level
Private UseChanged As Boolean           'whether or not the unfluence use has been edited

Private Const Checkmark = "×"

Private GridUpdateInProgress As Boolean 'whether VB is busy updating the grid
Private UseStorageInProgress As Boolean 'whether VB is busy updating the use
Private UseList As LinkedRumorList      'list of influence uses

Private Sub RefreshDateList()
'
' Name:         RefreshDateList
' Description:  Fill the list of dates with game dates in chronological order
'

    Dim LongDate As Long
    Dim Entry As Integer
    Dim Cursor As Integer

    Cursor = cmbDate.ListIndex
    cmbDate.Clear

    If Not InfluenceUseList.IsEmpty Then
        InfluenceUseList.First
        cmbDate.AddItem Format(InfluenceUseList.Item.DateStamp, "mmmm d, yyyy")
        cmbDate.ItemData(0) = CLng(InfluenceUseList.Item.DateStamp)
        InfluenceUseList.MoveNext
        
        Do Until InfluenceUseList.Off
            LongDate = CLng(InfluenceUseList.Item.DateStamp)
            Entry = 0
            Do Until Entry >= cmbDate.ListCount
                If cmbDate.ItemData(Entry) < LongDate Then
                    Entry = Entry + 1
                Else
                    Exit Do
                End If
            Loop
            cmbDate.AddItem Format(InfluenceUseList.Item.DateStamp, "mmmm d, yyyy"), Entry
            cmbDate.ItemData(Entry) = LongDate
            InfluenceUseList.MoveNext
        Loop
    End If
    
    If Cursor >= cmbDate.ListCount Then Cursor = cmbDate.ListCount - 1
    cmbDate.ListIndex = Cursor
    
End Sub

Private Sub RefreshNameList()
'
' Name:         RefreshNameList
' Description:  Fill the list of character names
'
        
    Dim Cursor As Integer

    Cursor = lstCharacters.ListIndex
    lstCharacters.Clear
        
    CharacterList.First
    Do Until CharacterList.Off
        lstCharacters.AddItem CharacterList.Item.Name
        CharacterList.MoveNext
    Loop

    If Cursor >= lstCharacters.ListCount Then Cursor = lstCharacters.ListCount - 1
    lstCharacters.ListIndex = Cursor

End Sub

Private Sub SetThisUse(Character As String, Influence As String, Level As Integer)
'
' Name:         SetThisUse
' Parameters:   Character       the character
'               Influence       the influence
'               Level           the level of influence
' Description:  Load the recorded influence use for the given character,
'               influence, and level
'

    If Not UseList Is Nothing Then
        
        If UseChanged Then StoreThisUse
        
        UseStorageInProgress = True
        
        ThisCharacter = Character
        ThisInfluence = Influence
        ThisLevel = Level
        
        UseList.MoveTo rtUse, ThisInfluence, ThisLevel, ThisCharacter
        If UseList.Off Then
            chkMarkUsed = False
            txtUse = ""
        Else
            chkMarkUsed = IIf(UseList.Item.Exclude, vbChecked, vbUnchecked)
            txtUse = UseList.Item.Text
        End If
    
        UseStorageInProgress = False
    End If
    
End Sub

Private Sub StoreThisUse()
'
' Name:         StoreThisUse
' Description:  Record an influence use for the selected character, influence,
'               and level
'

    Dim NewUse As RumorClass
    Dim NewText As String
    Dim Row As Integer

    If UseChanged And Not UseList Is Nothing Then
        NewText = TrimWhiteSpace(txtUse)
        UseList.MoveTo rtUse, ThisInfluence, ThisLevel, ThisCharacter
        
        For Row = 1 To grdInfluences.Rows - 1
            If grdInfluences.TextMatrix(Row, 0) = ThisInfluence Then _
                    Exit For
        Next Row
        
        If NewText = "" And chkMarkUsed = vbUnchecked Then
            grdInfluences.TextMatrix(Row, ThisLevel) = ""
            If Not UseList.Off Then UseList.Remove
        Else
            grdInfluences.TextMatrix(Row, ThisLevel) = Checkmark
            If UseList.Off Then
                Set NewUse = New RumorClass
                NewUse.Category = rtUse
                NewUse.Recipient = ThisInfluence
                NewUse.Level = ThisLevel
                NewUse.UsedBy = ThisCharacter
                NewUse.Exclude = (chkMarkUsed = vbChecked)
                NewUse.Text = NewText
                UseList.InsertSorted NewUse
            Else
                UseList.Item.Exclude = (chkMarkUsed = vbChecked)
                UseList.Item.Text = NewText
            End If
        End If
                
        UseChanged = False
    End If

End Sub

Private Sub EnableEntry(Maybe As Boolean)
'
' Name:         EnableEntry
' Parameters:   Maybe       whether to enable or disable entry
' Description:  Enable/Disable entry of influence uses
'

    grdInfluences.Visible = Maybe
    lblDate.Visible = Maybe
    lblCharacterName.Visible = Maybe
    lblInfluence.Visible = Maybe
    chkMarkUsed.Enabled = Maybe
    txtUse.Enabled = Maybe
    cmdClearUse.Enabled = Maybe
    cmdClearCharacter.Enabled = Maybe
    cmdClearDate.Enabled = Maybe

End Sub

Private Sub chkMarkUsed_Click()
'
' Name:         chkMarkUsed_Click
' Description:  Mark a rumor as "used"
'

    If Not UseStorageInProgress Then
        UseChanged = True
        Game.DataChanged = True
    End If

End Sub

Private Sub cmdClearCharacter_Click()
'
' Name:         cmdClearCharacter_Click
' Description:  Clear all influence uses for the selected character.
'
    
    Dim Row As Integer
    Dim Col As Integer
    
    If Not UseList Is Nothing Then
        If Not UseList.IsEmpty Then
            If MsgBox("This will PERMANENTLY clear all influence use for " & _
                    ThisCharacter & ".  Are you sure you want to do this?", _
                    vbYesNo + vbDefaultButton2, "Clear All for This Character") = vbYes Then
                    
                Game.DataChanged = False

                If Not (TrimWhiteSpace(txtUse) = "" And chkMarkUsed = vbUnchecked) Then
                        txtUse = ""
                        chkMarkUsed = vbUnchecked
                        UseChanged = True
                        StoreThisUse
                End If
                
                UseList.First
                Do Until UseList.Off
                    If UseList.Item.UsedBy = ThisCharacter Then
                        UseList.Remove
                    Else
                        UseList.MoveNext
                    End If
                Loop
                For Row = 1 To grdInfluences.Rows - 1
                    For Col = 1 To grdInfluences.Cols - 1
                        grdInfluences.TextMatrix(Row, Col) = ""
                    Next Col
                Next Row
                
            End If
        End If
    End If

End Sub

Private Sub cmdClearDate_Click()
'
' Name:         cmdClearDate_Click
' Description:  Clear all influence uses for the selected date.
'
    
    Dim Row As Integer
    Dim Col As Integer
    
    If Not UseList Is Nothing Then
        If Not UseList.IsEmpty Then
            If MsgBox("This will PERMANENTLY clear all influence use for all characters for " & _
                    Format(UseList.DateStamp, "mmmm d, yyyy") & ".  Are you sure you want to do this?", _
                    vbYesNo + vbDefaultButton2, "Clear All for This Date") = vbYes Then
                    
                If Not (TrimWhiteSpace(txtUse) = "" And chkMarkUsed = vbUnchecked) Then
                        txtUse = ""
                        chkMarkUsed = vbUnchecked
                        UseChanged = True
                        StoreThisUse
                End If
                
                Game.DataChanged = True
                UseList.Clear
                For Row = 1 To grdInfluences.Rows - 1
                    For Col = 1 To grdInfluences.Cols - 1
                        grdInfluences.TextMatrix(Row, Col) = ""
                    Next Col
                Next Row
                
            End If
        End If
    End If

End Sub

Private Sub cmdClearUse_Click()
'
' Name:         cmdClearUse_Click
' Description:  Clear a single influence use.
'

    If Not (TrimWhiteSpace(txtUse) = "" And chkMarkUsed = vbUnchecked) Then
        If MsgBox("This will permanently clear this influence use. " & _
                " Are you sure you want to do this?", vbYesNo + vbDefaultButton2, _
                "Clear This Use") = vbYes Then
    
            Game.DataChanged = True
            txtUse = ""
            chkMarkUsed = vbUnchecked
            UseChanged = True
            StoreThisUse
            
        End If
    End If
    
End Sub

Private Sub cmdESCClose_Click()
'
' Name:         cmdESCClose_Click
' Description:  Close the window when the user presses ESC.
'

    Unload Me

End Sub

Private Sub cmdSelect_Click()
'
' Name:         cmdSelect_Click
' Description:  Select a date and character, initializing the influence grid
'

    Dim InfluenceList As LinkedTraitList
    Dim Row As Integer
    Dim HighLevel As Integer
    Dim Col As Integer
    Dim GridFormatRows As String
    Dim GridFormatCols As String

    Screen.MousePointer = vbHourglass

    If lstCharacters.ListIndex <> -1 And cmbDate.Text <> "" Then
        
        CharacterList.MoveTo lstCharacters.Text
        InfluenceUseList.MoveTo CStr(CDate(cmbDate.Text))
        
        If Not (CharacterList.Off Or InfluenceUseList.Off) Then
            
            If Not GridUpdateInProgress Then StoreThisUse
            
            lblCharacterName = lstCharacters.Text
            lblDate = cmbDate.Text
            
            Set UseList = InfluenceUseList.Item
            Set InfluenceList = CharacterList.Item.InfluenceList
            
            If Not InfluenceList.IsEmpty Then
            
                GridUpdateInProgress = True
                
                '
                ' Supply the grid with rows, columns, and labels for each
                '
                grdInfluences.Rows = InfluenceList.TraitCount + 1
                grdInfluences.Cols = 2
                HighLevel = 1
                
                InfluenceList.First
                Do Until InfluenceList.Off
                
                    GridFormatRows = GridFormatRows & "|" & _
                             InfluenceList.Trait.Name & "  "
                    
                    '#' If InfluenceList.Item.Number > HighLevel Then
                    '#'         HighLevel = InfluenceList.Item.Number
                    '#'         grdInfluences.Cols = HighLevel + 1
                    '#' End If
                    
                    InfluenceList.MoveNext
                    
                Loop
                
                For Col = 1 To HighLevel
                    GridFormatCols = GridFormatCols & "| " & CStr(Col) & " "
                Next Col
                
                grdInfluences.FormatString = GridFormatCols & ";" & GridFormatRows
                
                '
                ' Color and fill the correct cells
                '
                
                InfluenceList.First
                Row = 1
                Do Until InfluenceList.Off
                    
                    grdInfluences.Row = Row
                    For Col = 1 To grdInfluences.Cols - 1
                        
                        grdInfluences.Col = Col
                        grdInfluences.CellFontBold = True
                        '#' If Col <= InfluenceList.Item.Number Then
                        '#'
                        '#'     grdInfluences.CellBackColor = grdInfluences.BackColor
                        '#'     grdInfluences.CellForeColor = grdInfluences.ForeColor
                        '#'
                        '#'     If UseList Is Nothing Then
                        '#'         grdInfluences.Text = ""
                        '#'     Else
                        '#'         UseList.MoveTo rtUse, InfluenceList.Item.Name, _
                        '#'                 Col, lblCharacterName
                        '#'         If UseList.Off Then
                        '#'             grdInfluences.Text = ""
                        '#'         Else
                        '#'             grdInfluences.Text = Checkmark
                        '#'         End If
                        '#'     End If
                        '#'
                        '#' Else
                        '#'     grdInfluences.CellBackColor = grdInfluences.BackColorBkg
                        '#'     grdInfluences.Text = ""
                        '#' End If

                    Next Col
                    InfluenceList.MoveNext
                    Row = Row + 1
                
                Loop
                
                EnableEntry True
                grdInfluences.Row = 1
                grdInfluences.Col = 1
                GridUpdateInProgress = False
                
                grdInfluences_EnterCell
                
            Else 'The influence list is empty
                
                lblEmpty = "--No Influence Exists for the Selected Character--"
                EnableEntry False
                cmdClearDate.Enabled = True

            End If
            
        Else 'Character not found
        
            If CharacterList.Off Then _
                MsgBox "Character name matching error!", vbExclamation, "Error!"
            If InfluenceUseList.Off Then _
                MsgBox "Influence Use List matching error!", vbExclamation, "Error!"
            Unload Me
        
        End If

    Else     'Misc. errors
    
        If lstCharacters.ListIndex = -1 Then
            lblEmpty = "--No Character Selected--"
        Else
            lblEmpty = "--No Dates Exist for Rumors--" & vbCrLf & _
                "Add dates from the Rumor Database screen."
        End If
        EnableEntry False
    
    End If

    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  Update the form data if it has changed in other windows
'

    If NamesChangedInfUse Then
        RefreshNameList
    End If
    
    If DatesChangedInfUse Then
        RefreshDateList
    End If

    If NamesChangedInfUse Or DatesChangedInfUse Or InfluenceChangedInfUse Then
        cmdSelect_Click
    End If

    NamesChangedInfUse = False
    DatesChangedInfUse = False
    InfluenceChangedInfUse = False
    
End Sub

Private Sub Form_Deactivate()
'
' Name:         Form_Deactivate
' Description:  Store the current influence use when the window loses
'               focus.
'

    StoreThisUse

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Initialize the date and character lists.
'

    RefreshDateList
    If cmbDate.ListCount > 0 Then
        cmbDate.ListIndex = cmbDate.ListCount - 1
    End If
    
    RefreshNameList
    If lstCharacters.ListCount > 0 Then
        lstCharacters.ListIndex = 0
    End If

    GridUpdateInProgress = True
    Call cmdSelect_Click
    GridUpdateInProgress = False
    
    DatesChangedInfUse = False
    NamesChangedInfUse = False
    InfluenceChangedInfUse = False
    mdiMain.OrientForm Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Unload
' Description:  Store the influence use and dismiss the window.
'

    StoreThisUse

End Sub

Private Sub grdInfluences_EnterCell()
'
' Name:         grdInfluences_EnterCell
' Description:  Store the current use and load into the entry area the use
'               associated with the selected influence and level.
'
    
    If Not GridUpdateInProgress Then
        With grdInfluences
            If .CellBackColor = .BackColorBkg And .Col > 1 Then
                .Col = .Col - 1
            Else
    
                StoreThisUse
                lblInfluence = .TextMatrix(.Row, 0) & " " & CStr(.Col)
                .CellBackColor = .BackColorSel
                .CellForeColor = .ForeColorSel
                SetThisUse lblCharacterName, .TextMatrix(.Row, 0), .Col

            End If
        End With
    End If
    
End Sub

Private Sub grdInfluences_LeaveCell()
'
' Name:         grdInfluences_LeaveCell
' Description:  Deselect this influence and level.
'

    With grdInfluences
        If Not GridUpdateInProgress And .CellBackColor = .BackColorSel Then
            .CellBackColor = .BackColor
            .CellForeColor = .ForeColor
        End If
    End With

End Sub

Private Sub lstCharacters_DblClick()
'
' Name:         lstCharacters_DblClick
' Description:  Shortcut to cmdSelect_Click.
'

    cmdSelect_Click

End Sub

Private Sub txtUse_Change()
'
' Name:         txtUse_Change
' Description:  Record a change to this influence use.
'
    
    If Not UseStorageInProgress Then
        UseChanged = True
        Game.DataChanged = True
    End If

End Sub

Private Sub txtUse_KeyPress(KeyAscii As Integer)
'
' Name:         txtUse_KeyPress
' Description:  Mark this influence "used".
'

    If Not (UseStorageInProgress Or UseChanged) Then
        If txtUse = "" Then chkMarkUsed = vbChecked
    End If

End Sub
